"""Flask server for the room booking tablet PWA.

Serves the static app shell and a small /api backed by Microsoft Graph
(graph_client.py). Meant to run directly on the tablet:

    python server.py
"""

import os

import requests as requests_lib
from flask import Flask, jsonify, request, send_from_directory

import graph_client

app = Flask(__name__, static_folder=None)
ROOT_DIR = os.path.dirname(__file__)


def _graph_error_response(exc):
    if isinstance(exc, requests_lib.exceptions.HTTPError):
        status = exc.response.status_code if exc.response is not None else 502
        return jsonify({"error": "graph_error", "detail": str(exc)}), (401 if status == 401 else 502)
    return jsonify({"error": "graph_error", "detail": str(exc)}), 502


@app.route("/")
def index():
    return send_from_directory(ROOT_DIR, "index.html")


@app.route("/<path:filename>")
def static_files(filename):
    # Serves manifest.json, sw.js, icons, coredat-mark.svg alongside index.html.
    return send_from_directory(ROOT_DIR, filename)


@app.route("/api/agenda")
def api_agenda():
    try:
        agenda = graph_client.get_today_agenda()
    except Exception as exc:  # noqa: BLE001 - surfaced to the client as-is
        return _graph_error_response(exc)
    return jsonify([{k: v for k, v in item.items() if not k.startswith("_")} for item in agenda])


@app.route("/api/status")
def api_status():
    try:
        agenda = graph_client.get_today_agenda()
        status = graph_client.get_current_status(agenda)
    except Exception as exc:  # noqa: BLE001
        return _graph_error_response(exc)
    return jsonify(status)


@app.route("/api/employees")
def api_employees():
    return jsonify(graph_client.CONFIG.get("employees", []))


@app.route("/api/book", methods=["POST"])
def api_book():
    payload = request.get_json(silent=True) or {}
    employee = payload.get("employee")
    duration = payload.get("durationMinutes")

    if not employee or not isinstance(duration, int) or duration <= 0:
        return jsonify({"error": "invalid_request"}), 400

    try:
        event = graph_client.create_booking(employee, duration)
    except Exception as exc:  # noqa: BLE001
        return _graph_error_response(exc)
    return jsonify({"ok": True, "eventId": event.get("id")})


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
