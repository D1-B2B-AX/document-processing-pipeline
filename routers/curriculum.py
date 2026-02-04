from fastapi import APIRouter, UploadFile, File, HTTPException
from pptx import Presentation
from openai import OpenAI
import io
import os

client = OpenAI(api_key=os.environ.get("OPENAI_API_KEY"))
router = APIRouter()

# =========================================================
# [설정] 제외 키워드
# =========================================================
EXCLUDE_KEYWORDS = [
    "유사", "사례", "실적", "reference", "case", "history", "result",
    "강사프로필", "수행실적", "제안사", "회사소개", "appendix", "별첨"
]

# =========================================================
# [핵심] 버전을 타지 않는 '무조건 재귀 탐색' (Duck Typing)
# =========================================================
def get_text_from_shape_recursive(shape):
    """
    모양(Type)을 따지지 않고, 텍스트나 하위 도형이 있으면 무조건 추출합니다.
    (라이브러리 버전이 달라도 100% 동작함)
    """
    text_parts = []

    try:
        # 1. 텍스트가 있는가? (TextFrame)
        if hasattr(shape, "text") and shape.text and shape.text.strip():
            text_parts.append(shape.text.strip())

        # 2. 표인가? (Table)
        if hasattr(shape, "table") and shape.table:
            for row in shape.table.rows:
                row_cells = [c.text.replace('\n', ' ').strip() for c in row.cells if c.text.strip()]
                if row_cells:
                    text_parts.append(f"| {' | '.join(row_cells)} |")

        # 3. 자식을 가진 컨테이너(그룹)인가? 
        # (MSO_SHAPE_TYPE 확인 대신, shapes 속성이 있는지로 판단 -> 버전 호환성 해결)
        if hasattr(shape, "shapes"):
            for child in shape.shapes:
                text_parts.extend(get_text_from_shape_recursive(child))
                
    except Exception as e:
        # 특정 도형에서 에러가 나도 멈추지 않고 무시
        print(f"⚠️ 도형 처리 중 스킵: {e}")
        pass

    return text_parts

def extract_all_text(slide):
    all_texts = []
    
    # 제목 처리
    try:
        if slide.shapes.title and slide.shapes.title.text.strip():
            all_texts.append(f"### {slide.shapes.title.text.strip()}")
    except:
        pass # 제목 없으면 패스

    # 본문 처리
    for shape in slide.shapes:
        # 제목 객체는 중복 방지를 위해 건너뜀
        try:
            if slide.shapes.title and shape == slide.shapes.title:
                continue
        except:
            pass
            
        # 재귀 추출 실행
        all_texts.extend(get_text_from_shape_recursive(shape))
        
    return "\n".join(all_texts)

# =========================================================
# [LLM] 마크다운 변환
# =========================================================
def generate_markdown(filename, text_content):
    if len(text_content) < 50: return None

    # 너무 길면 자르기 (토큰 비용 절약 및 에러 방지)
    safe_text = text_content[:25000]

    prompt = f"""
    당신은 '기업 교육 제안서 분석 전문가'입니다.
    아래 텍스트는 PPT에서 추출한 커리큘럼 내용입니다.
    
    [지시사항]
    1. 내용을 분석하여 **RAG용 Markdown**으로 정리해줘.
    2. 문서 상단에 `> **Keywords**: ...` 필수 포함.
    3. **시간 정보(1H, 2H, 09:00~) 절대 삭제 금지.**
    4. 표 형식은 Markdown Table로 변환.
    5. 잡담 없이 결과만 출력.
    
    [파일명] {filename}
    [내용]
    {safe_text}
    """

    try:
        response = client.chat.completions.create(
            model="gpt-4o",
            messages=[{"role": "user", "content": prompt}],
            temperature=0
        )
        return response.choices[0].message.content.strip()
    except Exception as e:
        print(f"❌ LLM Error: {e}")
        return None

# =========================================================
# [Endpoint]
# =========================================================
@router.post("/parse")
async def parse_curriculum(file: UploadFile = File(...)):
    print(f"\n🚀 [Duck Typing Fix] 파일 처리 시작: {file.filename}")
    
    content = await file.read()
    
    try:
        prs = Presentation(io.BytesIO(content))
    except Exception as e:
        raise HTTPException(status_code=400, detail="Invalid PPTX file")

    # 텍스트 추출 (필터링 없이 전체 수집)
    full_text_list = []
    
    for i, slide in enumerate(prs.slides):
        text = extract_all_text(slide)
        
        # 간단한 제외 키워드 체크
        is_exclude = False
        for key in EXCLUDE_KEYWORDS:
            if key in text: 
                is_exclude = True
                break
        
        if not is_exclude and len(text.strip()) > 5:
            full_text_list.append(f"\n--- [Slide {i+1}] ---\n{text}")

    combined_text = "\n".join(full_text_list)
    print(f"📝 추출된 텍스트 길이: {len(combined_text)}자")

    # 결과 생성
    results = []
    
    # 텍스트가 있으면 무조건 변환 시도
    if len(combined_text) > 30:
        md_content = generate_markdown(file.filename, combined_text)
        
        if md_content:
            base_name = os.path.splitext(file.filename)[0]
            results.append({
                "course_index": 1,
                "suggested_filename": f"{base_name}_Parsed.md",
                "markdown": md_content
            })
    else:
        print("🚨 여전히 텍스트가 0입니다. 이미지 파일이거나 암호화된 파일일 수 있습니다.")

    return {
        "domain": "curriculum",
        "original_filename": file.filename,
        "parsed_courses": results,
        "count": len(results)
    }