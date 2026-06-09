export default async function handler(req, res) {
  if (req.method !== "POST") return res.status(405).json({ error: "Method not allowed" });

  const supabaseUrl = process.env.SUPABASE_URL;
  const supabaseKey = process.env.SUPABASE_SERVICE_KEY;
  if (!supabaseUrl || !supabaseKey) return res.status(500).json({ error: "Supabase not configured" });

  try {
    const { base64, fileName, mimeType, folder } = req.body;
    if (!base64 || !fileName) return res.status(400).json({ error: "Missing file data" });

    // Convert base64 to buffer
    const buffer = Buffer.from(base64, "base64");
    const safeName = fileName.replace(/[^a-zA-Z0-9._-]/g, "_");
    const path = `${folder || "misc"}/${Date.now()}-${safeName}`;

    const uploadRes = await fetch(
      `${supabaseUrl}/storage/v1/object/site-documents/${path}`,
      {
        method: "POST",
        headers: {
          "apikey": supabaseKey,
          "Authorization": `Bearer ${supabaseKey}`,
          "Content-Type": mimeType || "application/octet-stream",
          "x-upsert": "true",
        },
        body: buffer,
      }
    );

    if (!uploadRes.ok) {
      const err = await uploadRes.text();
      console.error("Supabase storage error:", err);
      return res.status(500).json({ error: err });
    }

    const publicUrl = `${supabaseUrl}/storage/v1/object/public/site-documents/${path}`;
    return res.status(200).json({ url: publicUrl });

  } catch (err) {
    console.error("file-upload error:", err);
    return res.status(500).json({ error: err.message });
  }
}
