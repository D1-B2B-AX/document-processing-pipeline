from fastapi import APIRouter, UploadFile, File

# main.py가 찾고 있는 'router' 변수가 바로 이겁니다!
router = APIRouter()

@router.post("/parse")
async def parse_reference(file: UploadFile = File(...)):
    return {
        "domain": "reference",
        "filename": file.filename,
        "message": "Reference parser is ready (Logic not implemented yet)"
    }