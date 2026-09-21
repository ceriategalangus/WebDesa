// api/debug.js — Cek apakah GROQ_API_KEY sudah terbaca (hapus setelah selesai debug)
export default async function handler(req, res) {
  const key = process.env.GROQ_API_KEY;
  if (!key) {
    return res.status(200).json({ 
      status: "MISSING",
      message: "GROQ_API_KEY belum ada di environment Vercel. Harap tambahkan lalu Redeploy."
    });
  }
  return res.status(200).json({ 
    status: "OK", 
    message: "GROQ_API_KEY ditemukan!",
    prefix: key.substring(0, 8) + "..." // tampilkan sebagian saja (aman)
  });
}
