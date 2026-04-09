function json(data, init = {}) {
  return new Response(JSON.stringify(data, null, 2), {
    status: init.status || 200,
    headers: {
      "content-type": "application/json; charset=UTF-8",
      "cache-control": "no-store",
      ...init.headers
    }
  });
}

export async function onRequestGet(context) {
  try {
    const result = await context.env.DB.prepare("SELECT 1 AS ok").first();

    return json({
      ok: true,
      runtime: "cloudflare-pages-functions",
      database: {
        binding: "DB",
        status: result?.ok === 1 ? "connected" : "unexpected"
      },
      now: new Date().toISOString()
    });
  } catch (error) {
    return json(
      {
        ok: false,
        runtime: "cloudflare-pages-functions",
        database: {
          binding: "DB",
          status: "error"
        },
        error: error.message
      },
      { status: 500 }
    );
  }
}
