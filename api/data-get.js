export default async function handler(req, res) {
  if (req.method !== "GET") return res.status(405).json({ error: "Method not allowed" });

  const supabaseUrl = process.env.SUPABASE_URL;
  const supabaseKey = process.env.SUPABASE_SERVICE_KEY;

  if (!supabaseUrl || !supabaseKey) return res.status(500).json({ error: "Supabase not configured" });

  const { table } = req.query;
  if (!table || !["concrete_data", "certs_data"].includes(table)) {
    return res.status(400).json({ error: "Invalid table" });
  }

  try {
    const response = await fetch(`${supabaseUrl}/rest/v1/${table}?id=eq.1`, {
      headers: {
        "apikey": supabaseKey,
        "Authorization": `Bearer ${supabaseKey}`,
        "Content-Type": "application/json",
      },
    });
    const data = await response.json();
    return res.status(200).json(data[0] || {});
  } catch (err) {
    return res.status(500).json({ error: err.message });
  }
}
