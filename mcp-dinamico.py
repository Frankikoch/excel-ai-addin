!pip install flask flask-cors openai -q

import threading, time, subprocess, requests, re, json
from flask import Flask, request, jsonify
from flask_cors import CORS
from openai import OpenAI

app = Flask(__name__)
CORS(app)

# API key por defecto (fallback)
DEFAULT_API_KEY = "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"

# Modelo por defecto (fallback)
DEFAULT_MODEL = "minimaxai/minimax-m2.7"

# Cache de clientes por API key
clients = {}

def get_client(api_key):
    global clients
    if api_key not in clients:
        clients[api_key] = OpenAI(
            base_url="https://integrate.api.nvidia.com/v1",
            api_key=api_key
        )
    return clients[api_key]

@app.route('/health', methods=['GET'])
def health():
    # Devuelve los parámetros actuales (si se han enviado)
    return jsonify({
        "status": "ok",
        "default_model": DEFAULT_MODEL,
        "default_provider": "nvidia"
    })

@app.route('/chat', methods=['POST'])
def chat():
    raw = request.get_data(as_text=True)
    print(f">>> RECIBIDO: {raw[:500]}", flush=True)

    try:
        data = json.loads(raw) if raw else {}
    except:
        data = {}

    # Obtener API key del request o usar default
    api_key = data.get("apiKey") or request.headers.get("Authorization", "").replace("Bearer ", "") or DEFAULT_API_KEY
    print(f">>> API KEY: {api_key[:20]}...", flush=True)

    # Obtener modelo del request o usar default
    model = data.get("model") or DEFAULT_MODEL
    print(f">>> MODELO: {model}", flush=True)

    # Obtener provider del request
    provider = data.get("provider", "nvidia")
    print(f">>> PROVIDER: {provider}", flush=True)

    texto = str(data.get('message', '') or '')
    texto = re.sub(r'\[object \w+\]', '', texto).strip()
    print(f">>> MENSAJE: '{texto}'", flush=True)

    if not texto:
        return jsonify({"response": "❌ No se recibió texto válido"})

    messages = [{'role': 'user', 'content': texto}]
    
    # Añadir contexto de Excel si existe
    cell_context = data.get('context', [])
    if cell_context:
        messages[0]['content'] += f"\n📊 Contexto de Excel:\n{json.dumps(cell_context, ensure_ascii=False)}"

    # Añadir contexto de la celda activa si existe
    active_cell = data.get('activeCell', {})
    if active_cell.get('formula'):
        messages[0]['content'] += f"\n📝 Celda activa: {active_cell.get('address')} = {active_cell.get('formula')}"

    try:
        if provider == "nvidia":
            client = get_client(api_key)
            completion = client.chat.completions.create(
                model=model,
                messages=messages,
                temperature=0.7,
                top_p=0.95,
                max_tokens=1000
            )
            reply = completion.choices[0].message.content
            print(f">>> NVIDIA: OK, reply: {reply[:100]}...", flush=True)
        elif provider == "openrouter":
            # OpenRouter
            resp = requests.post(
                'https://openrouter.ai/api/v1/chat/completions',
                headers={
                    'Authorization': f'Bearer {api_key}',
                    'Content-Type': 'application/json',
                    'HTTP-Referer': 'https://frankikoch.github.io/excel-ai-addin/',
                    'X-Title': 'Excel AI'
                },
                json={'model': model, 'messages': messages, 'max_tokens': 1000},
                timeout=30
            )
            result = resp.json()
            reply = result['choices'][0]['message']['content']
            print(f">>> OPENROUTER: OK", flush=True)
        else:
            reply = f"Proveedor '{provider}' no implementado. Usa 'nvidia' o 'openrouter'."
    except Exception as e:
        print(f">>> ERROR: {e}", flush=True)
        reply = f"❌ Error: {str(e)}"

    return jsonify({"response": reply})

# Arrancar servidor
threading.Thread(target=lambda: app.run(port=5000, use_reloader=False), daemon=True).start()
time.sleep(2)
print("✅ Flask corriendo con soporte dinámico para API Key y Modelo", flush=True)
print("📌 Endpoints: /health, /chat", flush=True)