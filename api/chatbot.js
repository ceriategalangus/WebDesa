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

    // 1. Fetch daftar model yang tersedia untuk API Key ini
    let selectedModel = "llama3-8b-8192"; // default fallback
    try {
      const modelsRes = await fetch("https://api.groq.com/openai/v1/models", {
        headers: { "Authorization": "Bearer " + GROQ_API_KEY }
      });
      if (modelsRes.ok) {
        const modelsData = await modelsRes.json();
        const availableModels = modelsData.data ? modelsData.data.map(m => m.id) : [];
        
        // Cari model chat yang valid (hindari whisper/audio/guard)
        const chatModels = availableModels.filter(m => 
          !m.includes("whisper") && !m.includes("guard") && !m.includes("audio")
        );
        
        // Prioritaskan Llama 3.3/3.1, lalu Mixtral, lalu Gemma, lalu model chat apa saja
        const bestModel = 
          chatModels.find(m => m.includes("llama-3.3")) ||
          chatModels.find(m => m.includes("llama-3.1")) ||
          chatModels.find(m => m.includes("llama")) ||
          chatModels.find(m => m.includes("mixtral")) ||
          chatModels.find(m => m.includes("gemma")) ||
          chatModels[0]; // ambil apapun yang tersisa
          
        if (bestModel) {
          selectedModel = bestModel;
        }
      }
    } catch (e) {
      console.warn("Gagal fetch models list, menggunakan fallback.", e);
    }

    // Panggil Groq API
    const groqRes = await fetch("https://api.groq.com/openai/v1/chat/completions", {
      method: "POST",
      headers: {
        "Authorization": "Bearer " + GROQ_API_KEY,
        "Content-Type": "application/json"
      },
      body: JSON.stringify({
        model: selectedModel,
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
