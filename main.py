from __future__ import annotations

import os
from typing import Any

from fastapi import FastAPI

from line_user_linker import DEFAULT_WORKBOOK_PATH, register_line_user_routes
from pcgs_history_bot import register_pcgs_history_routes


app = FastAPI(title="LINE Excel Services")


@app.get("/health")
def health() -> dict[str, Any]:
    return {
        "ok": True,
        "service": "line-excel-services",
        "workbook_path": str(DEFAULT_WORKBOOK_PATH),
    }


register_line_user_routes(app)
register_pcgs_history_routes(app)


if __name__ == "__main__":
    import uvicorn

    port = int(os.getenv("PORT", "8080"))
    uvicorn.run("main:app", host="0.0.0.0", port=port)
