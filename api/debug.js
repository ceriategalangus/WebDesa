// api/debug.js — Cek GROQ_API_KEY dan list models
export default async function handler(req, res) {
  const key = process.env.GROQ_API_KEY;
  if (!key) {
    return res.status(200).json({ status: "MISSING", message: "GROQ_API_KEY belum ada." });
  }
  
  try {
    const modelsRes = await fetch("https://api.groq.com/openai/v1/models", {
      headers: { "Authorization": "Bearer " + key }
    });
    const modelsData = await modelsRes.json();
    return res.status(200).json({
      status: "OK",
      models: modelsData.data ? modelsData.data.map(m => m.id) : modelsData
    });
  } catch(e) {
    return res.status(500).json({ error: e.message });
  }
}

