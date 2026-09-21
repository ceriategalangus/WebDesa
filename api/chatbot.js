
export default async function handler(req, res) {
  // Hanya menerima metode POST
  if (req.method !== \'POST\') {
    return res.status(405).json({ error: \'Method Not Allowed\' });
  }

  try {
    const { message, history } = req.body;
    if (!message) {
      return res.status(400).json({ error: \'Message is required\' });
    }

    const GROQ_API_KEY = process.env.GROQ_API_KEY;
    if (!GROQ_API_KEY) {
      return res.status(500).json({ error: \'API Key belum disetting di Vercel (GROQ_API_KEY)\' });
    }

    // Bangun payload pesan
    const messages = [
      {
        role: \'system\',
        content: \'Kamu adalah asisten pintar bernama Tanya Desa untuk Website Desa. Jawablah dengan ramah, informatif, singkat, dan gunakan bahasa Indonesia yang sopan. Jawablah seputar layanan desa, profil desa, atau hal-hal yang wajar ditanyakan warga.\'
      },
      ...(history || []),
      { role: \'user\', content: message }
    ];

    const response = await fetch(\'https://api.groq.com/openai/v1/chat/completions\', {
      method: \'POST\',
      headers: {
        \'Authorization\': \'Bearer \' + GROQ_API_KEY,
        \'Content-Type\': \'application/json\'
      },
      body: JSON.stringify({
        model: \'llama-3.1-8b-instant\',
        messages: messages,
        temperature: 0.6,
        max_tokens: 400
      })
    });

    if (!response.ok) {
      const err = await response.text();
      throw new Error(\'Groq Error: \' + err);
    }

    const data = await response.json();
    const reply = data.choices[0]?.message?.content || \'Maaf, saya tidak bisa merespons saat ini.\';

    return res.status(200).json({ reply });

  } catch (error) {
    console.error(\'Chatbot Error:\', error);
    return res.status(500).json({ error: \'Terjadi kesalahan saat menghubungi server AI.\' });
  }
}
