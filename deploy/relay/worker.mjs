/**
 * Ретранслятор для invoice_bot.
 *
 * VPS Aeza (193.233.218.127) режется по IP: Groq / OpenRouter / OpenAI отдают 403
 * ещё до проверки ключа, api.telegram.org по IPv4 недоступен, по IPv6 нестабилен.
 * Воркер ходит туда за бота с чистого IP Cloudflare.
 *
 * Маршруты (secret — в пути, потому что telebot не умеет свои заголовки):
 *   /r/<secret>/bot<token>/<method>        -> api.telegram.org/bot<token>/<method>
 *   /r/<secret>/file/bot<token>/<path>     -> api.telegram.org/file/bot<token>/<path>
 *   /r/<secret>/ai/openai/v1/<endpoint>    -> openrouter.ai/api/v1/<endpoint>
 *
 * Для /ai/ ключ OpenRouter подставляет воркер: на VPS его нет вообще.
 * Оттуда же вычищаются groq-only параметры, которых OpenRouter не понимает.
 */

const TG_BASE = "https://api.telegram.org";
const AI_BASE_DEFAULT = "https://openrouter.ai/api/v1";

const DROP_REQ_HEADERS = new Set([
  "host", "content-length", "connection", "keep-alive", "transfer-encoding",
  "cf-connecting-ip", "cf-connecting-ipv6", "cf-ipcountry", "cf-ray", "cf-visitor",
  "cf-worker", "x-forwarded-for", "x-forwarded-proto", "x-real-ip", "accept-encoding",
]);

// groq-only: OpenRouter на них отвечает 400
const DROP_BODY_FIELDS = ["reasoning_effort", "include_reasoning", "service_tier"];

function forwardHeaders(src) {
  const out = new Headers();
  for (const [k, v] of src.entries()) {
    if (!DROP_REQ_HEADERS.has(k.toLowerCase())) out.set(k, v);
  }
  return out;
}

async function sanitizeAiBody(request) {
  const raw = await request.text();
  if (!raw) return raw;
  try {
    const body = JSON.parse(raw);
    for (const f of DROP_BODY_FIELDS) delete body[f];
    return JSON.stringify(body);
  } catch {
    return raw;
  }
}

export default {
  async fetch(request, env) {
    const url = new URL(request.url);
    const seg = url.pathname.split("/").filter((s) => s.length > 0);

    if (seg[0] !== "r" || seg[1] !== env.RELAY_SECRET) {
      return new Response("not found\n", { status: 404 });
    }

    const rest = seg.slice(2);
    if (rest.length === 0) return new Response("relay ok\n", { status: 200 });

    // диагностика: чей IP видят снаружи и доходит ли воркер до провайдера
    if (rest[0] === "diag") {
      const probe = url.searchParams.get("url");
      if (!probe) return new Response("diag: нужен ?url=\n", { status: 400 });
      try {
        const r = await fetch(probe, {
          headers: {
            "User-Agent": "invoice-relay/1.0",
            ...(url.searchParams.get("auth") ? { Authorization: "Bearer " + env.AI_KEY } : {}),
          },
        });
        const body = await r.text();
        return new Response(`status=${r.status}\n${body.slice(0, 500)}\n`, { status: 200 });
      } catch (e) {
        return new Response("diag error: " + String(e) + "\n", { status: 502 });
      }
    }

    let target;
    const headers = forwardHeaders(request.headers);
    let body;

    if (rest[0] === "ai") {
      // groq SDK шлёт /openai/v1/chat/completions — срезаем этот префикс
      let path = rest.slice(1);
      if (path[0] === "openai") path = path.slice(1);
      if (path[0] === "v1") path = path.slice(1);
      const aiBase = (env.AI_BASE || AI_BASE_DEFAULT).replace(/\/$/, "");
      target = `${aiBase}/${path.join("/")}${url.search}`;
      headers.set("Authorization", `Bearer ${env.AI_KEY}`);
      headers.set("HTTP-Referer", "https://termodinamika.workers.dev");
      headers.set("X-Title", "invoice-bot");
      if (request.method !== "GET" && request.method !== "HEAD") {
        body = await sanitizeAiBody(request);
        headers.set("Content-Type", "application/json");
      }
    } else {
      target = `${TG_BASE}/${rest.join("/")}${url.search}`;
      if (request.method !== "GET" && request.method !== "HEAD") body = request.body;
    }

    let upstream;
    try {
      upstream = await fetch(target, { method: request.method, headers, body });
    } catch (e) {
      return new Response(JSON.stringify({ ok: false, relay_error: String(e) }), {
        status: 502,
        headers: { "Content-Type": "application/json" },
      });
    }

    const outHeaders = new Headers(upstream.headers);
    outHeaders.delete("content-encoding");
    outHeaders.delete("content-length");
    outHeaders.delete("transfer-encoding");

    return new Response(upstream.body, { status: upstream.status, headers: outHeaders });
  },
};
