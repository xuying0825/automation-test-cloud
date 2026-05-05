"""
Flask web server providing a chat interface for the eTeams Passport test agent.
Streams agent responses (text deltas + tool events) to the browser via SSE.
Conversation history is maintained per browser session so follow-up messages
like "能再跑一次吗" correctly refer to previous context.
"""

import asyncio
import json
import os
import queue
import secrets
import threading
import traceback
from flask import Flask, render_template, request, Response, stream_with_context, session

from agents import Runner
from agents.stream_events import RawResponsesStreamEvent, RunItemStreamEvent
from agent_config import create_agent

app = Flask(__name__)
app.secret_key = secrets.token_hex(32)

# Server-side history store: session_id -> list[message]
# (Flask session stores only the ID; actual history stays server-side to avoid
#  cookie size limits with long tool outputs)
_history_lock = threading.Lock()
_history_store: dict[str, list] = {}

_SENSITIVE_ARG_NAMES = {"password", "passwd", "pwd", "pass", "密码"}


def _sanitize_tool_args(value):
    """Mask sensitive tool-call arguments before streaming them to the UI."""
    if isinstance(value, dict):
        sanitized = {}
        for key, item in value.items():
            key_text = str(key).lower()
            if key_text in _SENSITIVE_ARG_NAMES or any(
                token in key_text for token in ("password", "passwd", "pwd")
            ):
                sanitized[key] = "******"
            else:
                sanitized[key] = _sanitize_tool_args(item)
        return sanitized
    if isinstance(value, list):
        return [_sanitize_tool_args(item) for item in value]
    return value


def _get_history(sid: str) -> list:
    with _history_lock:
        return list(_history_store.get(sid, []))


def _set_history(sid: str, history: list) -> None:
    with _history_lock:
        _history_store[sid] = history


# ---------------------------------------------------------------------------
# Agent runner (executes in a background thread with its own event loop)
# ---------------------------------------------------------------------------

def _run_agent(
    message: str,
    history: list,
    out_q: "queue.Queue[dict | None]",
) -> None:
    """Run the agent in a dedicated thread and push SSE-payload dicts to out_q."""

    async def _async_run() -> None:
        try:
            agent = create_agent()
            # Build input: prior history + new user message
            input_messages = history + [{"role": "user", "content": message}]
            result = Runner.run_streamed(agent, input=input_messages, max_turns=50)
            async for event in result.stream_events():

                # ── Text delta from the model ──────────────────────────
                if isinstance(event, RawResponsesStreamEvent):
                    data = event.data
                    # Chat Completions chunks (used by OpenAIChatCompletionsModel)
                    if hasattr(data, "choices"):
                        for choice in data.choices:
                            delta = getattr(choice, "delta", None)
                            if delta and getattr(delta, "content", None):
                                out_q.put({"type": "text_delta", "content": delta.content})
                    # Responses API format (fallback)
                    elif hasattr(data, "delta") and hasattr(data, "type"):
                        if "text" in str(data.type) and data.delta:
                            out_q.put({"type": "text_delta", "content": data.delta})

                # ── Tool call / tool output events ─────────────────────
                elif isinstance(event, RunItemStreamEvent):
                    name = getattr(event, "name", "")
                    item = getattr(event, "item", None)

                    if name == "tool_called" and item:
                        raw = getattr(item, "raw_item", None)
                        tool_name = getattr(raw, "name", "unknown") if raw else "unknown"
                        try:
                            args = json.loads(getattr(raw, "arguments", "{}")) if raw else {}
                            args = _sanitize_tool_args(args)
                            args_str = json.dumps(args, ensure_ascii=False)
                        except Exception:
                            args_str = "（参数解析失败，已隐藏）"
                        out_q.put({
                            "type": "tool_call",
                            "tool": tool_name,
                            "args": args_str,
                        })

                    elif name == "tool_output" and item:
                        output = str(getattr(item, "output", ""))[:2000]
                        out_q.put({"type": "tool_output", "content": output})

            # After stream completes, persist updated history
            new_history = result.to_input_list()
            out_q.put({"type": "history", "history": new_history})

        except Exception:
            out_q.put({"type": "error", "content": traceback.format_exc()})
        finally:
            out_q.put(None)  # sentinel: stream finished

    asyncio.run(_async_run())


# ---------------------------------------------------------------------------
# Routes
# ---------------------------------------------------------------------------

@app.route("/")
def index():
    return render_template("index.html")


@app.route("/chat", methods=["POST"])
def chat():
    body = request.get_json(silent=True) or {}
    message = (body.get("message") or "").strip()
    if not message:
        return {"error": "message 不能为空"}, 400

    # Assign a session ID if this is a new browser session
    if "sid" not in session:
        session["sid"] = secrets.token_hex(16)
    sid = session["sid"]

    history = _get_history(sid)

    out_q: queue.Queue[dict | None] = queue.Queue()
    thread = threading.Thread(target=_run_agent, args=(message, history, out_q), daemon=True)
    thread.start()

    def generate():
        while True:
            try:
                item = out_q.get(timeout=180)  # 3-minute hard timeout
            except queue.Empty:
                yield f"data: {json.dumps({'type': 'error', 'content': '等待超时，请重试。'})}\n\n"
                break

            if item is None:
                yield f"data: {json.dumps({'type': 'done'})}\n\n"
                break

            # Save history update (not sent to browser)
            if item.get("type") == "history":
                _set_history(sid, item["history"])
                continue

            yield f"data: {json.dumps(item, ensure_ascii=False)}\n\n"

    return Response(
        stream_with_context(generate()),
        mimetype="text/event-stream",
        headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"},
    )


@app.route("/reset", methods=["POST"])
def reset():
    """Clear conversation history for the current session."""
    sid = session.get("sid")
    if sid:
        _set_history(sid, [])
    return {"ok": True}


if __name__ == "__main__":
    host = os.environ.get("FLASK_HOST", "0.0.0.0")
    port = int(os.environ.get("FLASK_PORT", "5003"))
    app.run(debug=False, host=host, port=port, threaded=True)
