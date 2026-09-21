// api/chatbot.js — Vercel Serverless Function
// Menerima { message, history, system_context } dan meneruskan ke Groq

export default async function handler(req, res) {
  // Hanya menerima POST
  if (req.method !== "POST") {
    return res.status(405).json({ error: "Method Not Allowed" });
  }

  try {
    const { message, history, system_context } = req.body;

    if (!message) {
      return res.status(400).json({ error: "Message is required" });
    }

    const GROQ_API_KEY = process.env.GROQ_API_KEY;
    if (!GROQ_API_KEY) {
      return res.status(500).json({ error: "GROQ_API_KEY belum diset di Vercel Environment Variables" });
    }

    // Bangun system prompt
    const basePrompt =
      "Kamu adalah asisten pintar bernama Tanya Desa untuk Website Desa. " +
      "Jawablah dengan ramah, informatif, singkat, dan gunakan bahasa Indonesia yang sopan. " +
      "Jawablah seputar layanan desa, profil desa, atau hal-hal yang wajar ditanyakan warga.";

    const finalSystemPrompt = system_context
      ? basePrompt + "\n\nBERIKUT ADALAH DATA DESA SAAT INI (Gunakan data ini untuk menjawab jika relevan):\n" + system_context
      : basePrompt;

    // Bangun daftar pesan
    const messages = [
      { role: "system", content: finalSystemPrompt },
      ...(Array.isArray(history) ? history : []),
      { role: "user", content: message }
    ];

    // Panggil Groq API
    const groqRes = await fetch("https://api.groq.com/openai/v1/chat/completions", {
      method: "POST",
      headers: {
        "Authorization": "Bearer " + GROQ_API_KEY,
        "Content-Type": "application/json"
      },
      body: JSON.stringify({
        model: "llama-3.3-70b-versatile",
        messages: messages,
        temperature: 0.6,
        max_tokens: 500
      })
    });

    if (!groqRes.ok) {
      const errText = await groqRes.text();
      console.error("Groq HTTP Error", groqRes.status, errText);
      return res.status(500).json({ error: "Groq Error " + groqRes.status + ": " + errText });
    }

    const data = await groqRes.json();
    const reply = (data.choices && data.choices[0] && data.choices[0].message && data.choices[0].message.content)
      ? data.choices[0].message.content
      : "Maaf, saya tidak bisa merespons saat ini.";

    return res.status(200).json({ reply });

  } catch (error) {
    console.error("Chatbot Error:", error.message);
    return res.status(500).json({ error: "Terjadi kesalahan saat menghubungi server AI: " + error.message });
  }
}
