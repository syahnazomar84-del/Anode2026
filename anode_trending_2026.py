"""Production entrypoint for publishing Anode Trending dashboard.

Run locally:
    python3 anode_trending_2026.py

Deploy (recommended):
    gunicorn anode_trending_2026:server --bind 0.0.0.0:8050 --workers 2
"""

import os
from typing import Any

from dash import dcc

from m1pq_a_plotly_dashboard import app as base_app

app = base_app
server = app.server

# View-only graph behavior (UI hardening).
VIEW_ONLY_GRAPH_CONFIG = {
    "editable": False,
    "displaylogo": False,
    "showTips": False,
    "modeBarButtonsToRemove": [
        "select2d",
        "lasso2d",
        "autoScale2d",
        "toggleSpikelines",
        "hoverClosestCartesian",
        "hoverCompareCartesian",
        "resetScale2d",
    ],
}


def _apply_graph_config(node: Any) -> None:
    if isinstance(node, dcc.Graph):
        current = dict(getattr(node, "config", {}) or {})
        current.update(VIEW_ONLY_GRAPH_CONFIG)
        node.config = current
        return

    children = getattr(node, "children", None)
    if children is None:
        return

    if isinstance(children, (list, tuple)):
        for child in children:
            _apply_graph_config(child)
    else:
        _apply_graph_config(children)


_apply_graph_config(app.layout)

# Disable Dash dev tools in published mode.
app.enable_dev_tools(
    debug=False,
    dev_tools_ui=False,
    dev_tools_props_check=False,
    dev_tools_serve_dev_bundles=False,
    dev_tools_hot_reload=False,
)


@server.after_request
def set_security_headers(resp):
    # Basic browser-side hardening for public viewing.
    resp.headers["X-Content-Type-Options"] = "nosniff"
    resp.headers["X-Frame-Options"] = "DENY"
    resp.headers["Referrer-Policy"] = "no-referrer"
    resp.headers["Permissions-Policy"] = "geolocation=(), microphone=(), camera=()"
    resp.headers["Cache-Control"] = "no-store"
    resp.headers["Content-Security-Policy"] = (
        "default-src 'self'; "
        "img-src 'self' data: blob:; "
        "style-src 'self' 'unsafe-inline'; "
        "script-src 'self' 'unsafe-inline'; "
        "connect-src 'self'; "
        "frame-ancestors 'none'"
    )
    return resp


@server.get("/healthz")
def healthz():
    return {"status": "ok"}, 200


if __name__ == "__main__":
    host = os.getenv("HOST", "0.0.0.0")
    port = int(os.getenv("PORT", "8050"))
    app.run(host=host, port=port, debug=False)
