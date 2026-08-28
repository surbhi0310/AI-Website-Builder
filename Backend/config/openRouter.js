const groqUrl = "https://api.groq.com/openai/v1/chat/completions";

export const generateResponse = async (prompt, maxTokens = 8000) => {
  try {
    const res = await fetch(groqUrl, {
      method: "POST",

      headers: {
        Authorization: `Bearer ${process.env.GROQ_API_KEY}`,
        "Content-Type": "application/json",
      },

      body: JSON.stringify({
        model: "openai/gpt-oss-120b",

        messages: [
          {
            role: "system",
            content:
              "You are a website generation AI. Return ONLY valid JSON. Do not use markdown or code fences.",
          },
          {
            role: "user",
            content: prompt,
          },
        ],

        temperature: 0.3,

        max_tokens: maxTokens,

        response_format: {
          type: "json_object",
        },
      }),
    });

    if (!res.ok) {
      const err = await res.text();

      throw new Error(`Groq API error: ${err}`);
    }

    const data = await res.json();

    if (!data.choices || !data.choices[0]) {
      throw new Error("Invalid response from Groq");
    }

    const content = data.choices[0].message?.content;

    if (!content) {
      throw new Error("Groq returned empty content");
    }

    return content;
  } catch (error) {
    console.error("❌ Error from generateResponse:", error);
    throw error;
  }
};