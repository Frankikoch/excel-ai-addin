"""
MCP Excel Server - Flask + OpenRouter/OpenAI/Anthropic/Google/OpenCode
Sirve a Excel AI Add-in via localtunnel

用法 (cómo ejecutarlo):
1. Google Colab: pon este código en una celda y ejecuta
2. Local: python mcp-excel-server.py
"""

import os
import time
import threading
import subprocess
import urllib.request
import requests
from flask import Flask, request, jsonify

# === CONFIGURACIÓN - PON AQUÍ TUS KEYS ===
OPENAI_API_KEY = os.environ.get("OPENAI_API_KEY", "")
ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY", "")
OPENROUTER_API_KEY = os.environ.get("OPENROUTER_API_KEY", "")
OPENCODE_API_KEY = os.environ.get("OPENCODE_API_KEY", "")

app = Flask(__name__)

# === MAPA DE MODELOS POR PROVEEDOR ===
MODELS = {
    "opencode": {
        "name": "OpenCode",
        "models": [
            {"id": "minimax-m2.5-free", "name": "MiniMax M2.5 Free"},
        ],
    },
    "openai": {
        "name": "OpenAI",
        "models": [
            {"id": "openai/gpt-4o-mini", "name": "GPT-4o Mini"},
            {"id": "openai/gpt-4o", "name": "GPT-4o"},
            {"id": "openai/gpt-4-turbo", "name": "GPT-4 Turbo"},
            {"id": "openai/gpt-3.5-turbo", "name": "GPT-3.5 Turbo"},
            {"id": "openai/o1-mini", "name": "o1 Mini"},
            {"id": "openai/o1", "name": "o1"},
        ],
    },
    "openrouter": {
        "name": "OpenRouter (Free)",
        "models": [
            {
                "id": "meta-llama/llama-3.1-8b-instruct:free",
                "name": "Meta Llama 3.1 8B",
            },
            {"id": "qwen/qwen2.5-7b-instruct:free", "name": "Qwen 2.5 7B"},
            {"id": "mistralai/mistral-7b-instruct:free", "name": "Mistral 7B"},
            {"id": "google/gemma-3-1b-it:free", "name": "Google Gemma 3 1B"},
            {
                "id": "nvidia/nemotron-3-nano-omni-30b-a3b-reasoning:free",
                "name": "Nvidia Nemotron 3 Nano",
            },
            {"id": "google/gemma-4-26b-a4b-it:free", "name": "Google Gemma 4 26B"},
        ],
    },
    "anthropic": {
        "name": "Anthropic",
        "models": [
            {"id": "anthropic/claude-3.5-haiku", "name": "Claude 3.5 Haiku"},
            {"id": "anthropic/claude-3.5-sonnet", "name": "Claude 3.5 Sonnet"},
            {"id": "anthropic/claude-3-opus", "name": "Claude 3 Opus"},
            {"id": "anthropic/claude-3-sonnet", "name": "Claude 3 Sonnet"},
            {"id": "anthropic/claude-3-haiku", "name": "Claude 3 Haiku"},
            {"id": "anthropic/claude-2.1", "name": "Claude 2.1"},
        ],
    },
    "google": {
        "name": "Google",
        "models": [
            {"id": "google/gemini-flash-1.5-8b", "name": "Gemini Flash 1.5"},
            {"id": "google/gemini-pro-1.5", "name": "Gemini Pro 1.5"},
            {"id": "google/gemini-1.5-flash-8b", "name": "Gemini 1.5 Flash 8B"},
            {"id": "google/gemini-1.0-pro", "name": "Gemini 1.0 Pro"},
            {"id": "google/gemini-2.0-flash-exp", "name": "Gemini 2.0 Flash"},
            {"id": "google/gemini-2.0-pro-exp", "name": "Gemini 2.0 Pro"},
        ],
    },
}


# === ENDPOINTS ===
@app.route("/health", methods=["GET"])
def health():
    return jsonify({"status": "ok", "providers": list(MODELS.keys())})


@app.route("/models", methods=["GET"])
def list_models():
    return jsonify(MODELS)


@app.route("/chat", methods=["POST"])
def chat():
    data = request.get_json(force=True) or {}
    message = data.get("message", "")
    context = data.get("context", [])
    provider = data.get("provider", "opencode")
    model = data.get("model", "minimax-m2.5-free")

    # Compactar contexto de Excel
    ctx_lines = []
    for c in (context or [])[:20]:
        addr = c.get("address", "")
        val = c.get("value", "")
        ctx_lines.append(f"{addr}: {val}")
    ctx = "\n".join(ctx_lines)

    user_content = message + (f"\n\n📊 Contexto de Excel:\n{ctx}" if ctx else "")

    system_prompt = "Eres un asistente para Excel. Responde siempre en español, de forma concisa y accionable. Cuando analices datos, proporciona insights específicos."

    try:
        if provider == "opencode":
            response = call_opencode(model, system_prompt, user_content)
        elif provider == "openai":
            response = call_openai(model, system_prompt, user_content)
        elif provider == "openrouter":
            response = call_openrouter(model, system_prompt, user_content)
        elif provider == "anthropic":
            response = call_anthropic(model, system_prompt, user_content)
        elif provider == "google":
            response = call_google(model, system_prompt, user_content)
        else:
            response = f"Proveedor no reconocido: {provider}"
    except Exception as e:
        response = f"❌ Error: {str(e)[:200]}"

    return jsonify({"response": response})


# === HELPERS ===
def call_opencode(model, system, user):
    if not OPENCODE_API_KEY:
        return "⚠️ OpenCode API Key no configurada"
    r = requests.post(
        "https://api.opencode.ai/v1/chat/completions",
        headers={
            "Authorization": f"Bearer {OPENCODE_API_KEY}",
            "Content-Type": "application/json",
        },
        json={
            "model": model,
            "messages": [
                {"role": "system", "content": system},
                {"role": "user", "content": user},
            ],
            "max_tokens": 800,
        },
        timeout=40,
    )
    if r.status_code != 200:
        return f"Error OpenCode: {r.status_code}"
    return (
        r.json().get("choices", [{}])[0].get("message", {}).get("content", "")
        or "Sin respuesta"
    )


def call_openai(model, system, user):
    if not OPENAI_API_KEY:
        return "⚠️ OpenAI API Key no configurada"
    r = requests.post(
        "https://api.openai.com/v1/chat/completions",
        headers={
            "Authorization": f"Bearer {OPENAI_API_KEY}",
            "Content-Type": "application/json",
        },
        json={
            "model": model,
            "messages": [
                {"role": "system", "content": system},
                {"role": "user", "content": user},
            ],
            "max_tokens": 800,
        },
        timeout=40,
    )
    if r.status_code != 200:
        return f"Error OpenAI: {r.status_code}"
    return (
        r.json().get("choices", [{}])[0].get("message", {}).get("content", "")
        or "Sin respuesta"
    )


def call_openrouter(model, system, user):
    if not OPENROUTER_API_KEY:
        return "⚠️ OpenRouter API Key no configurada"
    r = requests.post(
        "https://openrouter.ai/api/v1/chat/completions",
        headers={
            "Authorization": f"Bearer {OPENROUTER_API_KEY}",
            "Content-Type": "application/json",
            "HTTP-Referer": "https://frankikoch.github.io/excel-ai-addin",
            "X-Title": "Excel AI",
        },
        json={
            "model": model,
            "messages": [
                {"role": "system", "content": system},
                {"role": "user", "content": user},
            ],
            "max_tokens": 800,
        },
        timeout=40,
    )
    if r.status_code != 200:
        return f"Error OpenRouter: {r.status_code}"
    return (
        r.json().get("choices", [{}])[0].get("message", {}).get("content", "")
        or "Sin respuesta"
    )


def call_anthropic(model, system, user):
    if not ANTHROPIC_API_KEY:
        return "⚠️ Anthropic API Key no configurada"
    r = requests.post(
        "https://api.anthropic.com/v1/messages",
        headers={
            "x-api-key": ANTHROPIC_API_KEY,
            "anthropic-version": "2023-06-01",
            "content-type": "application/json",
        },
        json={
            "model": model,
            "max_tokens": 800,
            "system": system,
            "messages": [{"role": "user", "content": user}],
        },
        timeout=40,
    )
    if r.status_code != 200:
        return f"Error Anthropic: {r.status_code}"
    return r.json().get("content", [{}])[0].get("text", "") or "Sin respuesta"


def call_google(model, system, user):
    # Usa OpenRouter como fallback para Google si no tienes key directa
    if not OPENROUTER_API_KEY:
        return "⚠️ Configura OpenRouter API Key"
    return call_openrouter(model, system, user)


# === INICIAR ===
if __name__ == "__main__":
    PORT = 5002
    threading.Thread(
        target=lambda: app.run(
            host="0.0.0.0", port=PORT, debug=False, use_reloader=False
        ),
        daemon=True,
    ).start()
    print(f"🚀 Servidor en puerto {PORT}")

    def run_tunnel():
        p = subprocess.Popen(
            ["lt", "--port", str(PORT)],
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
        )
        for line in p.stdout:
            if "loca.lt" in line:
                url = line.strip().replace("your url is: ", "")
                print(f"\n✅ URL PÚBLICA: {url}")
                print(f"   GET  {url}/health")
                print(f"   POST {url}/chat")
                break

    threading.Thread(target=run_tunnel, daemon=True).start()
    time.sleep(3)
    print("⏳ Tunnel iniciando...")
