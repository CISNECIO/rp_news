const ALLOWED_IPS = new Set([
  "179.49.70.241"
]);

export async function onRequest(context) {
  const ip = context.request.headers.get("CF-Connecting-IP") || "";

  if (!ALLOWED_IPS.has(ip)) {
    return new Response("Forbidden", { status: 403 });
  }

  return context.next();
}