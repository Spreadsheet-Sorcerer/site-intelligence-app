import { IncomingForm } from 'formidable';
import fs from 'fs';

export const config = { api: { bodyParser: false } };

export default async function handler(req, res) {
  if (req.method !== "POST") return res.status(405).json({ error: "Method not allowed" });

  const supabaseUrl = process.env.SUPABASE_URL;
  const supabaseKey = process.env.SUPABASE_SERVICE_KEY;
  if (!supabaseUrl || !supabaseKey) return res.status(500).json({ error: "Supabase not configured" });

  try {
    const form = new IncomingForm();
    const [fields, files] = await new Promise((resolve, reject) => {
      form.parse(req, (err, fields, files) => {
        if (err) reject(err);
        else resolve([fields, files]);
      });
    });

    const file = Array.isArray(files.file) ? files.file[0] : files.file;
    const folder = Array.isArray(fields.folder) ? fields.folder[0] : (fields.folder || "misc");
    const fileName = Array.isArray(fields.fileName) ? fields.fileName[0] : fields.fileName;

    const fileBuffer = fs.readFileSync(file.filepath);
    const mimeType = file.mimetype || "application/octet-stream";
    const safeName = fileName || file.originalFilename || `file-${Date.now()}`;
    const path = `${folder}/${Date.now()}-${safeName.replace(/[^a-zA-Z0-9._-]/g, "_")}`;

    const uploadRes = await fetch(
      `${supabaseUrl}/storage/v1/object/site-documents/${path}`,
      {
        method: "POST",
        headers: {
          "apikey": supabaseKey,
          "Authorization": `Bearer ${supabaseKey}`,
          "Content-Type": mimeType,
          "x-upsert": "true",
        },
        body: fileBuffer,
      }
    );

    if (!uploadRes.ok) {
      const err = await uploadRes.text();
      console.error("Supabase storage error:", err);
      return res.status(500).json({ error: err });
    }

    // Return the public URL
    const publicUrl = `${supabaseUrl}/storage/v1/object/public/site-documents/${path}`;
    return res.status(200).json({ url: publicUrl, path });

  } catch (err) {
    console.error("file-upload error:", err);
    return res.status(500).json({ error: err.message });
  }
}
