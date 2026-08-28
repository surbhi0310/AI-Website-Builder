const Extractjson = async (text) => {
  try {
    if (!text) {
      console.error("❌ No AI response received");
      return null;
    }

    const cleaned = text
      .replace(/```json/gi, "")
      .replace(/```/g, "")
      .trim();

    const firstBraces = cleaned.indexOf("{");
    const lastBraces = cleaned.lastIndexOf("}");

    if (firstBraces === -1 || lastBraces === -1) {
      console.error("❌ No JSON found in AI response");
      console.log(cleaned);
      return null;
    }

    const jsonString = cleaned.slice(
      firstBraces,
      lastBraces + 1
    );

    const parsed = JSON.parse(jsonString);

    if (!parsed.code) {
      console.error("❌ AI response does not contain code");
      return null;
    }

    return parsed;

  } catch (error) {
    console.error("❌ JSON parsing failed:", error);
    console.log("Raw AI response:", text);

    return null;
  }
};

export default Extractjson;