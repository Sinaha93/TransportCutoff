from pathlib import Path

from fastapi import FastAPI
from fastapi.staticfiles import StaticFiles

from app.config import RuntimePaths
from app.db import Database
from app.web.routes import register_web


def health() -> dict[str, str]:
    return {"status": "ok"}


def create_app(*, database: Database | None = None) -> FastAPI:
    application = FastAPI()
    # Startup is lazy for the module-level app so merely importing /health never
    # creates a database in an unexpected working directory.
    application.state.database = database or Database(
        RuntimePaths.from_root(Path(__file__).resolve().parents[1]).database
    )
    application.state.database_ready = False
    application.state.import_reviews = {}
    application.add_api_route("/health", health, methods=["GET"])
    application.mount("/static", StaticFiles(directory=Path(__file__).parent / "web" / "static"), name="static")
    register_web(application)
    return application


app = create_app()
