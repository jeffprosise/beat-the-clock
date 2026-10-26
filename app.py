import uuid
from flask import Flask, render_template, request, Response, stream_with_context, send_file, abort
from agno.db.in_memory import InMemoryDb
from agents import create_agent
from pathlib import Path
import os

app = Flask(__name__)
_memory = InMemoryDb()

ITINERARIES_DIR = (Path(__file__).parent / "itineraries").resolve()

# Home page
@app.route("/", methods=["GET"])
def index():
    return render_template("index.html")

# REST method for chatting with an agent and generating a streaming response
@app.route("/streaming_chat", methods=["POST"])
def streaming_chat():
    try:
        # Get the user input
        user_input = request.form.get('input')

        if not user_input or not user_input.strip():
            return Response('Input cannot be empty', status=400)

        # Create an agent bound to this session and run it
        session_id = request.headers.get('X-Session-ID') or str(uuid.uuid4())
        agent = create_agent(session_id, _memory)
        streaming_output = agent.run(user_input, stream=True)

        # Inline generator for streaming output
        def generate():
            try:
                for chunk in streaming_output:
                    if chunk.event == 'RunContent':
                        yield chunk.content
                    elif chunk.event == 'RunError':
                        yield f'\n\n**Error**: {chunk.content}'

            except Exception as e:
                yield f'\n\n**Error**: {get_error_message(e)}'

        response = Response(
            stream_with_context(generate()),
            mimetype='text/plain'
        )
        response.headers['X-Session-ID'] = session_id
        return response

    except Exception as e:
        response = Response(
            f'\n\n**Error**: {get_error_message(e)}',
            status=getattr(e, 'status_code', 500),
            mimetype='text/plain; charset=utf-8'
        )
        response.headers['X-Session-ID'] = session_id
        return response

# REST endpoint for downloading generated itineraries
@app.route("/itineraries/<filename>", methods=["GET"])
def download_itinerary(filename):
    safe_name = os.path.basename(filename)
    if not safe_name.lower().endswith(".docx"):
        abort(404)

    path = os.path.join(ITINERARIES_DIR, safe_name)
    if not os.path.isfile(path):
        abort(404)

    return send_file(
        path,
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        as_attachment=True,
        download_name=safe_name,
    )

# Helper function to generate user-friendly error messages from exceptions
def get_error_message(e):
    body = getattr(e, 'body', None)

    if isinstance(body, dict):
        error = body.get('error', body)

        if isinstance(error, dict):
            message = error.get('message')
            if message:
                return message

    return str(e)
