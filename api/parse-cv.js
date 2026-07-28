const DEEPSEEK_URL = "https://api.deepseek.com/chat/completions";
const MAX_TEXT_LENGTH = 120_000;
const MAX_PROMPT_LENGTH = 20_000;

export default async function handler(req, res) {
  if (req.method !== "POST") {
    res.setHeader("Allow", "POST");
    return res.status(405).json({ error: "Method not allowed" });
  }

  const apiKey = process.env.DEEPSEEK_API_KEY;
  if (!apiKey) {
    return res.status(500).json({ error: "DEEPSEEK_API_KEY belum dikonfigurasi di Vercel." });
  }

  const { documentText, extractionPrompt } = req.body || {};
  if (typeof documentText !== "string" || !documentText.trim()) {
    return res.status(400).json({ error: "Teks dokumen kosong." });
  }
  if (typeof extractionPrompt !== "string" || !extractionPrompt.trim()) {
    return res.status(400).json({ error: "Instruksi ekstraksi kosong." });
  }
  if (documentText.length > MAX_TEXT_LENGTH) {
    return res.status(413).json({ error: "Dokumen terlalu panjang untuk diproses." });
  }
  if (extractionPrompt.length > MAX_PROMPT_LENGTH) {
    return res.status(413).json({ error: "Instruksi ekstraksi terlalu panjang." });
  }

  try {
    const upstream = await fetch(DEEPSEEK_URL, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Authorization: `Bearer ${apiKey}`,
      },
      body: JSON.stringify({
        model: "deepseek-chat",
        messages: [
          { role: "system", content: "You are a precise maritime crew data extraction API. Return valid JSON only." },
          {
            role: "user",
            content: `${extractionPrompt}\n\nHere is the raw text extracted from the CV:\n\n${documentText}`,
          },
        ],
        response_format: { type: "json_object" },
        temperature: 0.1,
      }),
    });

    const data = await upstream.json().catch(() => null);
    if (!upstream.ok) {
      return res.status(upstream.status).json({
        error: data?.error?.message || `DeepSeek API error (${upstream.status})`,
      });
    }

    const content = data?.choices?.[0]?.message?.content;
    if (!content) {
      return res.status(502).json({ error: "DeepSeek tidak mengembalikan hasil ekstraksi." });
    }

    return res.status(200).json({ content });
  } catch (error) {
    return res.status(502).json({ error: error.message || "Gagal menghubungi DeepSeek." });
  }
}
