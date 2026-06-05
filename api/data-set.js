export default async function handler(req, res) {
  if (req.method !== "POST") return res.status(405).json({ error: "Method not allowed" });

  const supabaseUrl = process.env.SUPABASE_URL;
  const supabaseKey = process.env.SUPABASE_SERVICE_KEY;

  if (!supabaseUrl || !supabaseKey) return res.status(500).json({ error: "Supabase not configured" });

  const { table, data } = req.body;
  if (!table || !["concrete_data", "certs_data"].includes(table)) {
    return res.status(400).json({ error: "Invalid table" });
  }

  // Explicitly cast JSON fields for Supabase
  const payload = {};
  if (table === "concrete_data") {
    if (data.tickets  !== undefined) payload.tickets  = data.tickets;
    if (data.invoices !== undefined) payload.invoices = data.invoices;
    if (data.tests    !== undefined) payload.tests    = data.tests;
  }
  if (table === "certs_data") {
    if (data.certs !== undefined) payload.certs = data.certs;
  }

  try {
    const response = await fetch(`${supabaseUrl}/rest/v1/${table}?id=eq.1`, {
      method: "PATCH",
      headers: {
        "apikey": supabaseKey,
        "Authorization": `Bearer ${supabaseKey}`,
        "Content-Type": "application/json",
        "Prefer": "return=minimal",
      },
      body: JSON.stringify(payload),
    });

    if (!response.ok) {
      const err = await response.text();
      console.error("Supabase error:", err);
      return res.status(500).json({ error: err });
    }
    return res.status(200).json({ success: true });
  } catch (err) {
    console.error("data-set error:", err);
    return res.status(500).json({ error: err.message });
  }
}
