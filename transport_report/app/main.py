from pathlib import Path

from fastapi import FastAPI
from fastapi.staticfiles import StaticFiles

from app.config import RuntimePaths
from app.db import Database
from app.web.routes import register_web


def health() -> dict[str, str]:
    return {"status": "ok"}


def create_app(*, database: Database | None = None, paths: RuntimePaths | None = None) -> FastAPI:
    application = FastAPI()
    application.state.runtime_paths = paths or RuntimePaths.from_root(Path(__file__).resolve().parents[1])
    # Startup is lazy for the module-level app so merely importing /health never
    # creates a database in an unexpected working directory.
    application.state.database = database or Database(
        application.state.runtime_paths.database
    )
    application.state.database_ready = False
    application.state.import_reviews = {}
    application.add_api_route("/health", health, methods=["GET"])
    application.mount("/static", StaticFiles(directory=Path(__file__).parent / "web" / "static"), name="static")
    register_web(application)
    return application


app = create_app()
