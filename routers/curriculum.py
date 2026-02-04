from fastapi import APIRouter, UploadFile, File, HTTPException
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE  # [추가됨] 그룹 타입을 확인하기 위해 필요
from openai import OpenAI
import io
import re
import os

# =========================================================
# [설정] 환경 변수 및 클라이언트
# =========================================================
client = OpenAI(api_key=os.environ.get("OPENAI_API_KEY"))
router = APIRouter()

# =========================================================
# [키워드 설정] (원본 그대로 유지)
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
# [기능 1] PPTX 파싱 헬퍼 함수
# =========================================================
def normalize(text):
    return re.sub(r'\s+', '', str(text).lower())

def get_visual_title(slide):
    # PPTX 표준 제목
    if slide.shapes.title and slide.shapes.title.text.strip():
        return slide.shapes.title.text.strip()
    
    # 시각적 위치 기반 제목 추정 (Top < 2000000 EMU)
    candidates = []
    for shape in slide.shapes:
        if not hasattr(shape, "text") or not shape.text.strip():
            continue
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

# ---------------------------------------------------------
# [핵심 수정] 재귀 함수 추가 (그룹 내부 탐색용)
# ---------------------------------------------------------
def get_text_from_shape_recursive(shape):
    """도형이 그룹이면 재귀적으로 파고들어 모든 텍스트 추출"""
    text_parts = []

    # 1. 일반 텍스트 박스
    if hasattr(shape, "text") and shape.text.strip():
        text_parts.append(shape.text.strip())

    # 2. 표 (Table)
    if shape.shape_type == MSO_SHAPE_TYPE.TABLE:
        for row in shape.table.rows:
            row_cells = [c.text.replace('\n', ' ').strip() for c in row.cells if c.text.strip()]
            if row_cells:
                text_parts.append(f"| {' | '.join(row_cells)} |")

    # 3. 그룹 (Group) - 여기가 0건 문제 해결의 열쇠!
    if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
        for child in shape.shapes:
            text_parts.extend(get_text_from_shape_recursive(child))

    return text_parts

def extract_text_from_slide(slide):
    lines = []
    visual_title = get_visual_title(slide)
    if visual_title:
        lines.append(f"### {visual_title}")
    
    # [수정됨] 단순 반복문 -> 재귀 함수 호출로 변경
    for shape in slide.shapes:
        # 제목과 완전히 동일한 객체면 중복 추출 방지
        if slide.shapes.title and shape == slide.shapes.title:
            continue
            
        # 재귀적으로 텍스트 추출 (그룹 포함)
        extracted_texts = get_text_from_shape_recursive(shape)
        
        # 제목과 텍스트 내용이 중복되는 경우 제거 (visual_title 기준)
        for text in extracted_texts:
            if visual_title and text.strip() == visual_title:
                continue
            lines.append(text)

    return "\n".join(lines)

# =========================================================
# [기능 2] LLM Markdown 변환 (원본 그대로)
# =========================================================
def generate_rag_markdown(filename, course_idx, overview_text, curriculum_text):
    if len(curriculum_text) < 50: return None

    prompt = f"""
    당신은 '기업 교육 제안서 분석 전문가'입니다.
    제공된 Raw Text를 분석하여, RAG 모델 학습에 최적화된 **Clean Markdown** 포맷으로 변환하십시오.

    [Input Source]
    - File: {filename}
    - Context (개요): {overview_text[:3000]}
    - Content (커리큘럼): {curriculum_text[:15000]}

    [Output Rules - Strict]
    1. **Metadata Block**: 문서 최상단에 아래 양식을 반드시 포함할 것.
       > **File**: {filename}
       > **Type**: B2B Corporate Training Curriculum
       > **Keywords**: (문서의 핵심 태그: 대상, 주제, 툴, 시간 등을 나열)

    2. **Section Structuring**:
       - 과정명/주제는 `# (H1)` 태그 사용
       - '교육 개요', '학습 목표' 등 대분류는 `## (H2)` 태그 사용
       - 세부 모듈/시간표는 `### (H3)` 태그 사용
    
    3. **Curriculum Table & Time Info (매우 중요)**:
       - 커리큘럼 상세 내용은 반드시 Markdown Table 혹은 계층형 List(`-`)로 정리할 것.
       - **[중요] 소요 시간 보존:** '1H', '4H', '2시간', '09:00~12:00' 등의 시간 정보는 실제 운영에 필수적이므로 **절대 삭제하지 말 것.**
       - 시간 정보는 주제 옆에 괄호로 명시하거나(예: `### 모듈명 (2H)`), 테이블의 별도 컬럼으로 유지할 것.

    4. **Filtering**:
       - '강사 약력', '회사 홍보', '레퍼런스' 등 커리큘럼과 무관한 내용은 과감히 삭제할 것.
       - 정보가 없으면 없는 대로 놔둘 것 (지어내지 말 것).
       
    5. **No Chit-chat**: 서론/본론 없이 오직 Markdown 내용만 출력할 것. 만약 유효한 정보가 없다면 `NO_DATA` 출력.
    """

    try:
        response = client.chat.completions.create(
            model="gpt-4o",
            messages=[{"role": "user", "content": prompt}],
            temperature=0
        )
        result = response.choices[0].message.content.strip()
        
        if "NO_DATA" in result: return None
        if len(result) < 50: return None
        
        return result

    except Exception as e:
        print(f"LLM Error: {e}")
        return None

# =========================================================
# [FastAPI 엔드포인트] 메인 핸들러
# =========================================================
@router.post("/parse")
async def parse_curriculum(file: UploadFile = File(...)):
    # 1. 파일 읽기
    content = await file.read()
    original_filename = file.filename
    
    # 2. PPTX 로드
    try:
        prs = Presentation(io.BytesIO(content))
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Invalid PPTX file: {str(e)}")

    courses = [] 
    current_course = {'overview': [], 'curriculum': []}
    
    # 3. 슬라이드 순회 및 과정 분류
    for slide in prs.slides:
        slide_type = classify_slide_advanced(slide)
        
        if slide_type == "EXCLUDE": continue

        # [수정됨] 여기서 그룹 내부 텍스트까지 포함된 내용을 가져옵니다.
        text = extract_text_from_slide(slide)

        if slide_type == "OVERVIEW":
            if current_course['curriculum']: 
                courses.append(current_course)
                current_course = {'overview': [], 'curriculum': []}
            current_course['overview'].append(text)

        elif slide_type == "CURRICULUM":
            current_course['curriculum'].append(text)
        
        # [안전장치 추가] 분류가 'OTHER'라도 텍스트 양이 많으면(100자 이상) 커리큘럼으로 포함
        # (그룹 안에 숨어있던 텍스트가 이제 보이니까, 놓치지 않기 위해 추가)
        elif slide_type == "OTHER" and len(text) > 100:
            current_course['curriculum'].append(text)

    if current_course['curriculum']:
        courses.append(current_course)

    # 4. LLM 변환 및 결과 생성
    results = []
    for idx, course in enumerate(courses):
        full_overview = "\n\n".join(course['overview'])
        full_curriculum = "\n\n".join(course['curriculum'])
        
        md_content = generate_rag_markdown(original_filename, idx+1, full_overview, full_curriculum)
        
        if md_content:
            # 원본 파일명 유지 로직
            base_name = os.path.splitext(original_filename)[0]
            suggested_filename = f"{base_name}_Course_{idx+1}.md"
            
            results.append({
                "course_index": idx + 1,
                "suggested_filename": suggested_filename,
                "markdown": md_content
            })

    # 5. 최종 응답
    return {
        "domain": "curriculum",
        "original_filename": original_filename,
        "parsed_courses": results,
        "count": len(results)
    }