from fastapi import APIRouter, UploadFile, File

router = APIRouter()

@router.post("/parse")
async def parse_instructor(file: UploadFile = File(...)):
    return {
        "domain": "instructor",
        "filename": file.filename,
        "message": "Instructor parser is ready (Logic not implemented yet)"
    }