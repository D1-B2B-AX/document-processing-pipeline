from fastapi import APIRouter, UploadFile, File, HTTPException
from pydantic import BaseModel
from pptx import Presentation
from openai import OpenAI
import io
import re
import os

# =========================================================
# [설정] 환경 변수에서 API 키 로드 (Railway Env Var 사용 권장)
# =========================================================
# Railway 변수 설정: OPENAI_API_KEY
client = OpenAI(api_key=os.environ.get("OPENAI_API_KEY"))

router = APIRouter()

# =========================================================
# [키워드 설정] (제공 코드 그대로 사용)
# =========================================================
EXCLUDE_KEYWORDS = [
    "유사", "사례", "실적", "reference", "case", "history", "result",
    "강사프로필", "수행실적", "제안사", "회사소개"
]
OVERVIEW_KEYWORDS = [
    "과정 소개", "과정소개", "과정 개요", "과정개요", 
    "교육 소개", "교육소개", "교육 개요", "교육개요", 
    "개요", "소개", "overview", "summary", "요약", "제안 배경", "기획 의도",
    "목표", "대상" 
]
CURRICULUM_KEYWORDS = [
    "커리큘럼", "세부과정", "교육과정", "교육내용", "모듈구성", 
    "상세과정", "프로그램", "module", "schedule", "curriculum",
    "모듈", "구성", "일정", "방법", "contents", "agenda", "syllabus",
    "1일차", "2일차", "1h", "2h", "time" 
]

# =========================================================
# [헬퍼 함수] PPTX 파싱 로직 (제공 코드 이식)
# =========================================================
def normalize(text):
    return re.sub(r'\s+', '', str(text).lower())

def get_visual_title(slide):
    if slide.shapes.title and slide.shapes.title.text.strip():
        return slide.shapes.title.text.strip()
    
    candidates = []
    for shape in slide.shapes:
        if not hasattr(shape, "text") or not shape.text.strip():
            continue
        # 상단(top < 2000000 EMU)에 위치한 텍스트를 제목 후보로 간주
        if shape.top < 2000000: 
            candidates.append((shape.top, shape.left, shape.text.strip()))
    
    if candidates:
        candidates.sort(key=lambda x: (x[0], x[1])) 
        return candidates[0][2]
    return ""

def check_table_headers(slide):
    for shape in slide.shapes:
        if shape.has_table:
            header_text = ""
            try:
                for cell in shape.table.rows[0].cells:
                    header_text += cell.text + " "
            except:
                continue
            norm_header = normalize(header_text)
            for key in CURRICULUM_KEYWORDS:
                if normalize(key) in norm_header:
                    return True
    return False

def classify_slide_advanced(slide):
    title = get_visual_title(slide)
    norm_title = normalize(title)
    
    for key in EXCLUDE_KEYWORDS:
        if normalize(key) in norm_title: return "EXCLUDE"
    for key in CURRICULUM_KEYWORDS:
        if normalize(key) in norm_title: return "CURRICULUM"
    for key in OVERVIEW_KEYWORDS:
        if normalize(key) in norm_title: return "OVERVIEW"
    if check_table_headers(slide):
        return "CURRICULUM"
    return "OTHER"

def extract_text_from_slide(slide):
    lines = []
    visual_title = get_visual_title(slide)
    if visual_title:
        lines.append(f"### {visual_title}")
    
    for shape in slide.shapes:
        if hasattr(shape, "text") and shape.text.strip():
            if shape.text.strip() == visual_title:
                continue
            lines.append(shape.text.strip())
        
        if shape.has_table:
            for row in shape.table.rows:
                row_cells = [c.text.replace('\n', ' ').strip() for c in row.cells if c.text.strip()]
                if row_cells:
                    lines.append(f"| {' | '.join(row_cells)} |")
    return "\n".join(lines)

# =========================================================
# [LLM 호출] Markdown 변환 (제공 코드 이식)
# =========================================================
def generate_rag_markdown(filename, course_idx, overview_text, curriculum_text):
    if len(curriculum_text) < 50: return None

    prompt = f"""
    당신은 'B2B 교육 커리큘럼 정리 전문가'입니다.
    아래 제공된 Raw Text를 분석하여, RAG 검색에 최적화된 **Clean Markdown** 포맷으로 변환하십시오.

    [Input Source]
    - File: {filename}
    - Context: {overview_text[:3000]}
    - Content: {curriculum_text[:15000]}

    [Output Format Rules - Strict Markdown]
    1. **Metadata Block**: 문서 최상단에 아래 양식을 반드시 포함할 것.
       > **File**: {filename}
       > **Type**: B2B Corporate Training Curriculum
    
    2. **Section Structuring**:
       - 과정명/주제는 `# (H1)` 태그 사용
       - '교육 개요', '학습 목표' 등 대분류는 `## (H2)` 태그 사용
       - 세부 모듈/시간표는 `### (H3)` 태그 사용
    
    3. **Curriculum Table**:
       - 커리큘럼 상세 내용은 반드시 Markdown Table 혹은 계층형 List(`-`)로 정리할 것.
       - 시간(Time), 모듈명(Module), 세부내용(Detail)이 명확히 구분되어야 함.

    4. **Filtering**:
       - '강사 약력', '회사 홍보', '레퍼런스' 등 커리큘럼과 무관한 내용은 과감히 삭제할 것.
       - 정보가 없으면 없는 대로 놔둘 것 (지어내지 말 것).
       
    5. **No Chit-chat**: 서론/본론 없이 오직 Markdown 내용만 출력할 것. 유효 정보 없으면 `NO_DATA`.
    """

    try:
        response = client.chat.completions.create(
            model="gpt-4o",
            messages=[{"role": "user", "content": prompt}],
            temperature=0
        )
        result = response.choices[0].message.content.strip()
        if "NO_DATA" in result or len(result) < 50: return None
        return result
    except Exception as e:
        print(f"LLM Error: {e}")
        return None

# =========================================================
# [FastAPI 엔드포인트]
# =========================================================
@router.post("/parse")
async def parse_curriculum(file: UploadFile = File(...)):
    # 1. 파일 읽기
    content = await file.read()
    filename = file.filename
    
    # 2. PPTX 로드
    try:
        prs = Presentation(io.BytesIO(content))
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Invalid PPTX file: {str(e)}")

    courses = [] 
    current_course = {'overview': [], 'curriculum': []}
    
    # 3. 슬라이드 순회 및 분류
    for slide in prs.slides:
        slide_type = classify_slide_advanced(slide)
        if slide_type == "EXCLUDE": continue

        text = extract_text_from_slide(slide)

        if slide_type == "OVERVIEW":
            if current_course['curriculum']: 
                courses.append(current_course)
                current_course = {'overview': [], 'curriculum': []}
            current_course['overview'].append(text)

        elif slide_type == "CURRICULUM":
            current_course['curriculum'].append(text)

    if current_course['curriculum']:
        courses.append(current_course)

    # 4. 결과 생성 (LLM 호출 포함)
    results = []
    for idx, course in enumerate(courses):
        full_overview = "\n\n".join(course['overview'])
        full_curriculum = "\n\n".join(course['curriculum'])
        
        md_content = generate_rag_markdown(filename, idx+1, full_overview, full_curriculum)
        
        if md_content:
            # 파일명 정제 (clean_name 로직 일부 적용)
            safe_name = re.sub(r'[^a-zA-Z0-9가-힣\s]', '', filename.replace('_', ' ').replace('.pptx', ''))
            safe_name = re.sub(r'\s+', ' ', safe_name).strip()
            
            results.append({
                "course_index": idx + 1,
                "suggested_filename": f"{safe_name}_Course_{idx+1}.md",
                "markdown": md_content
            })

    return {
        "domain": "curriculum",
        "original_filename": filename,
        "parsed_courses": results, # 하나의 PPTX에서 여러 과정이 나올 수 있으므로 리스트 반환
        "count": len(results)
    }