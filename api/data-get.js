export default async function handler(req, res) {
  if (req.method !== "GET") return res.status(405).json({ error: "Method not allowed" });

  const supabaseUrl = process.env.SUPABASE_URL;
  const supabaseKey = process.env.SUPABASE_SERVICE_KEY;

  if (!supabaseUrl || !supabaseKey) {
    console.error("Missing env vars - URL:", !!supabaseUrl, "KEY:", !!supabaseKey);
    return res.status(500).json({ error: "Supabase not configured", hasUrl: !!supabaseUrl, hasKey: !!supabaseKey });
  }

  const { table } = req.query;
  if (!table || !["concrete_data", "certs_data"].includes(table)) {
    return res.status(400).json({ error: "Invalid table" });
  }

  try {
    const url = `${supabaseUrl}/rest/v1/${table}?id=eq.1`;
    console.log("Fetching:", url);
    const response = await fetch(url, {
      headers: {
        "apikey": supabaseKey,
        "Authorization": `Bearer ${supabaseKey}`,
        "Content-Type": "application/json",
      },
    });
    const text = await response.text();
    console.log("Supabase response status:", response.status, "body:", text);
    if (!response.ok) return res.status(500).json({ error: text });
    const data = JSON.parse(text);
    return res.status(200).json(data[0] || {});
  } catch (err) {
    console.error("data-get error:", err);
    return res.status(500).json({ error: err.message });
  }
}
