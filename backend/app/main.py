import logging

from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware

from .routers import workbooks

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(name)s: %(message)s",
)

app = FastAPI(
    title="Excel Color Extractor",
    description="Extracts raw cell values and cell colors from Excel workbooks. "
    "Never assumes what a color means; reports only what is actually present in the file.",
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["http://localhost:5173", "http://127.0.0.1:5173"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

app.include_router(workbooks.router)


@app.get("/health")
async def health():
    return {"status": "ok"}
