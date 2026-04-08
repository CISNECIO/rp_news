const ALLOWED_IPS = new Set([
  "179.49.70.241"
]);

function unauthorized() {
  return new Response("Authentication required", {
    status: 401,
    headers: {
      "WWW-Authenticate": 'Basic realm="Restricted Area"'
    }
  });
}

export async function onRequest(context) {
  const ip = context.request.headers.get("CF-Connecting-IP") || "";

  // Office/network IP enters directly
  if (ALLOWED_IPS.has(ip)) {
    return context.next();
  }

  // Everyone else needs Basic Auth
  const auth = context.request.headers.get("Authorization");
  if (!auth || !auth.startsWith("Basic ")) {
    return unauthorized();
  }

  try {
    const encoded = auth.slice(6);
    const decoded = atob(encoded);
    const sep = decoded.indexOf(":");

    if (sep === -1) {
      return unauthorized();
    }

    const user = decoded.slice(0, sep);
    const pass = decoded.slice(sep + 1);

    if (
      user === context.env.Julio &&
      pass === context.env.Profe
    ) {
      return context.next();
    }
  } catch (e) {
    return unauthorized();
  }

  return unauthorized();
}