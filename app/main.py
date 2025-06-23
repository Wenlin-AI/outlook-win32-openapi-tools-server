"""FastAPI application factory."""

from __future__ import annotations

import os
from functools import lru_cache

from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from dotenv import load_dotenv

from .routers import tasks


@lru_cache(maxsize=1)
def create_app() -> FastAPI:
    load_dotenv()
    app = FastAPI(title="Outlook Tasks API", version="0.1.0")
    app.add_middleware(
        CORSMiddleware,
        allow_origins=["*"],
        allow_methods=["*"],
        allow_headers=["*"],
    )
    app.include_router(tasks.router)
    return app


def run() -> None:
    import uvicorn

    app = create_app()
    host = os.getenv("HOST", "0.0.0.0")
    port = int(os.getenv("PORT", "8124"))
    uvicorn.run(app, host=host, port=port)


if __name__ == "__main__":  # pragma: no cover
    run()
