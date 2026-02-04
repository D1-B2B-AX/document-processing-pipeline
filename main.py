from fastapi import FastAPI
from routers import curriculum, reference, instructor

app = FastAPI(title="Document Processing Pipeline")

# 3개의 도메인 창구를 등록합니다.
app.include_router(curriculum.router, prefix="/curriculum", tags=["Curriculum"])
app.include_router(reference.router, prefix="/reference", tags=["Reference"])
app.include_router(instructor.router, prefix="/instructor", tags=["Instructor"])

@app.get("/")
def health_check():
    return {"status": "online"}