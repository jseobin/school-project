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

function normalizeText(value, maxLength) {
  return String(value || "")
    .trim()
    .replace(/\s+/g, " ")
    .slice(0, maxLength);
}

export async function onRequestGet(context) {
  try {
    const { results } = await context.env.DB.prepare(
      "SELECT id, author, content, created_at FROM guestbook_entries ORDER BY id DESC LIMIT 20"
    ).all();

    return json({
      ok: true,
      entries: results || []
    });
  } catch (error) {
    return json(
      {
        ok: false,
        error: "Database query failed. Apply the D1 migration if this is a new environment.",
        details: error.message
      },
      { status: 500 }
    );
  }
}

export async function onRequestPost(context) {
  try {
    const body = await context.request.json();
    const author = normalizeText(body.author, 40);
    const content = normalizeText(body.content, 280);

    if (!author || !content) {
      return json(
        {
          ok: false,
          error: "Both author and content are required."
        },
        { status: 400 }
      );
    }

    const createdAt = new Date().toISOString();
    const insert = await context.env.DB.prepare(
      "INSERT INTO guestbook_entries (author, content, created_at) VALUES (?, ?, ?)"
    )
      .bind(author, content, createdAt)
      .run();

    const entry = await context.env.DB.prepare(
      "SELECT id, author, content, created_at FROM guestbook_entries WHERE id = ?"
    )
      .bind(insert.meta.last_row_id)
      .first();

    return json(
      {
        ok: true,
        entry
      },
      { status: 201 }
    );
  } catch (error) {
    return json(
      {
        ok: false,
        error: "Could not write the entry to D1.",
        details: error.message
      },
      { status: 500 }
    );
  }
}
