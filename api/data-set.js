export default async function handler(req, res) {
  if (req.method !== "POST") {
    return res.status(405).json({
      error: "Method not allowed",
    });
  }

  const supabaseUrl = process.env.SUPABASE_URL;
  const supabaseKey = process.env.SUPABASE_SERVICE_KEY;

  if (!supabaseUrl || !supabaseKey) {
    console.error(
      "Missing environment variables:",
      "SUPABASE_URL:",
      !!supabaseUrl,
      "SUPABASE_SERVICE_KEY:",
      !!supabaseKey
    );

    return res.status(500).json({
      error: "Supabase not configured",
      hasUrl: !!supabaseUrl,
      hasKey: !!supabaseKey,
    });
  }

  const { table, data } = req.body || {};

  const allowedTables = [
    "concrete_data",
    "certs_data",
  ];

  if (!table || !allowedTables.includes(table)) {
    return res.status(400).json({
      error: "Invalid table",
    });
  }

  if (!data || typeof data !== "object") {
    return res.status(400).json({
      error: "Invalid data",
    });
  }

  const payload = {};

  if (table === "concrete_data") {
    if (data.tickets !== undefined) {
      payload.tickets = data.tickets;
    }

    if (data.invoices !== undefined) {
      payload.invoices = data.invoices;
    }

    if (data.tests !== undefined) {
      payload.tests = data.tests;
    }
  }

  if (table === "certs_data") {
    if (data.certs !== undefined) {
      payload.certs = data.certs;
    }
  }

  if (Object.keys(payload).length === 0) {
    return res.status(400).json({
      error: "No valid data was provided",
    });
  }

  try {
    const url = `${supabaseUrl}/rest/v1/${table}?id=eq.1`;

    const response = await fetch(url, {
      method: "PATCH",
      headers: {
        apikey: supabaseKey,
        Authorization: `Bearer ${supabaseKey}`,
        "Content-Type": "application/json",
        Prefer: "return=representation",
      },
      body: JSON.stringify(payload),
    });

    const text = await response.text();

    if (!response.ok) {
      console.error(
        "Supabase data-set error:",
        response.status,
        text
      );

      return res.status(response.status).json({
        error: text || "Could not save data",
      });
    }

    const rows = text ? JSON.parse(text) : [];

    if (!Array.isArray(rows) || rows.length !== 1) {
      console.error(
        `Supabase update did not find id=1 in ${table}`
      );

      return res.status(409).json({
        error: `No id=1 row exists in ${table}`,
      });
    }

    return res.status(200).json({
      success: true,
      updated: 1,
    });
  } catch (error) {
    console.error("data-set error:", error);

    return res.status(500).json({
      error: error.message || "Unexpected server error",
    });
  }
}
