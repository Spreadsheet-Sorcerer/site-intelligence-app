export default async function handler(req, res) {
  // Never allow an older cached copy of the records to be returned.
  res.setHeader(
    "Cache-Control",
    "no-store, no-cache, must-revalidate, proxy-revalidate"
  );
  res.setHeader("Pragma", "no-cache");
  res.setHeader("Expires", "0");

  if (req.method !== "GET") {
    return res.status(405).json({
      error: "Method not allowed",
    });
  }

  const supabaseUrl = process.env.SUPABASE_URL;
  const supabaseKey = process.env.SUPABASE_SERVICE_KEY;

  if (!supabaseUrl || !supabaseKey) {
    console.error(
      "Missing Supabase variables:",
      !!supabaseUrl,
      !!supabaseKey
    );

    return res.status(500).json({
      error: "Supabase not configured",
      hasUrl: !!supabaseUrl,
      hasKey: !!supabaseKey,
    });
  }

  const { table } = req.query;

  if (
    !table ||
    !["concrete_data", "certs_data"].includes(table)
  ) {
    return res.status(400).json({
      error: "Invalid table",
    });
  }

  try {
    const url =
      `${supabaseUrl}/rest/v1/${table}?id=eq.1`;

    const response = await fetch(url, {
      method: "GET",
      headers: {
        apikey: supabaseKey,
        Authorization: `Bearer ${supabaseKey}`,
        "Content-Type": "application/json",
        "Cache-Control": "no-cache",
      },
    });

    const text = await response.text();

    if (!response.ok) {
      console.error(
        "Supabase data-get error:",
        response.status,
        text
      );

      return res.status(response.status).json({
        error: text || "Could not retrieve data",
      });
    }

    const rows = text ? JSON.parse(text) : [];

    return res.status(200).json(rows[0] || {});
  } catch (error) {
    console.error("data-get error:", error);

    return res.status(500).json({
      error:
        error.message || "Unexpected server error",
    });
  }
}
