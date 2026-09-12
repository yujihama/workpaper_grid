"""試作アプリのサーバ。調書グリッドの API（workpaper.py）を /api 配下に載せるだけ。"""
from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware

from .workpaper import router as workpaper_router

app = FastAPI(title="workpaper-grid-proto")
app.add_middleware(CORSMiddleware, allow_origins=["*"], allow_methods=["*"], allow_headers=["*"])
app.include_router(workpaper_router, prefix="/api")
