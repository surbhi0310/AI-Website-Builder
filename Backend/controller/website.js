import { generateResponse } from "../config/openRouter.js";
import UserDatabase from "../model/user.js";
import website from "../model/website.js";
import Extractjson from "../utils/extract.js";
import { nanoid } from "nanoid";

const masterPrompt = `
You are a senior frontend architect and UI/UX engineer.

Create a premium, modern, responsive website based on the user's request.

USER REQUIREMENT:
{USER_PROMPT}

RULES:
- Use ONLY HTML, CSS and vanilla JavaScript.
- Return ONE complete HTML document.
- Exactly one <style> tag.
- Exactly one <script> tag.
- No frameworks or external libraries.
- Use system fonts only.
- Use responsive CSS with Flexbox/Grid and media queries.
- Must work on mobile, tablet and desktop.
- No horizontal scrolling.
- Images may ONLY use https://images.unsplash.com/
- Every image URL must contain:
  ?auto=format&fit=crop&w=1200&q=80
- Images must be responsive.
- Create professional real content, not lorem ipsum.
- Buttons and navigation must actually work.
- Navigation should work as an SPA without page reload.
- Include Home, About, Services/Features and Contact sections/pages.
- Forms should have basic JavaScript validation.
- Add hover and active states.
- At least one page must be visible on initial load.
- Make the design visually polished and production-ready.

OUTPUT:
Return ONLY valid JSON.

{
  "message": "Short professional confirmation",
  "code": "<FULL VALID HTML DOCUMENT>"
}
`;

// Generate a new website
export const GenerateData = async (req, res) => {
  try {
    const { prompt } = req.body;

    if (!prompt || !prompt.trim()) {
      return res.status(400).json({
        message: "Prompt required for generating a website",
      });
    }

    const user = await UserDatabase.findById(req.user._id);

    if (!user) {
      return res.status(401).json({
        message: "Please login for website generation",
      });
    }

    if (user.credit < 50) {
      return res.status(400).json({
        message: "You don't have enough credits to generate a website",
      });
    }

    const finalprompt = masterPrompt.replace(
      "{USER_PROMPT}",
      prompt.trim()
    );

    let row = "";
    let parsed = null;

    for (let i = 0; i < 2 && !parsed; i++) {
      console.log(`🤖 AI generation attempt ${i + 1}`);

      try {
        row = await generateResponse(finalprompt);
        console.log("✅ AI response received");

        parsed = await Extractjson(row);
      } catch (error) {
        console.error("❌ AI generation error:", error);

        const errorMessage = error?.message || "";

        if (
          errorMessage.includes("rate_limit_exceeded") ||
          errorMessage.includes("tokens per minute") ||
          errorMessage.includes("Request too large")
        ) {
          return res.status(429).json({
            message:
              "AI usage limit reached. Please wait a little and try again.",
          });
        }

        if (i === 1) {
          return res.status(500).json({
            message: "AI could not generate the website",
          });
        }
      }

      if (!parsed && i === 0) {
        console.log("⚠️ First AI response was invalid. Retrying...");

        try {
          row = await generateResponse(
            finalprompt + "\n\nReturn ONLY valid JSON."
          );

          parsed = await Extractjson(row);
        } catch (error) {
          console.error("❌ Retry failed:", error);
        }
      }
    }

    if (!parsed || !parsed.code) {
      console.log("❌ AI returned invalid response:", row);

      return res.status(400).json({
        message: "AI returned invalid website data",
      });
    }

    const finalwebsite = await website.create({
      user: user._id,
      title: prompt.trim().slice(0, 60),
      latestcode: parsed.code,
      slug: nanoid(10),
      converasation: [
        {
          role: "ai",
          content: parsed.message || "Website generated successfully",
        },
        {
          role: "user",
          content: prompt.trim(),
        },
      ],
    });

    const updatedUser = await UserDatabase.findByIdAndUpdate(
      user._id,
      { $inc: { credit: -50 } },
      { new: true }
    );

    return res.status(201).json({
      websiteId: finalwebsite._id,
      remainingCredit: updatedUser.credit,
    });
  } catch (error) {
    console.error("❌ Generate website error:", error);

    return res.status(500).json({
      message: "Error while generating website",
    });
  }
};


// Get a single website
export const GetWebsiteById = async (req, res) => {
  try {
    const singlewebsite = await website.findOne({
      _id: req.params.id,
      user: req.user._id,
    });

    if (!singlewebsite) {
      return res.status(404).json({
        message: "Website is not found",
        success: false,
      });
    }

    return res.status(200).json(singlewebsite);
  } catch (error) {
    console.error("❌ Get website error:", error);

    return res.status(500).json({
      message: "Error while getting website",
    });
  }
};


// Update an existing website
export const changewebsite = async (req, res) => {
  try {
    const { prompt } = req.body;

    if (!prompt || !prompt.trim()) {
      return res.status(400).json({
        message: "Prompt required for updating website",
      });
    }

    const singlewebsite = await website.findOne({
      _id: req.params.id,
      user: req.user._id,
    });

    if (!singlewebsite) {
      return res.status(404).json({
        message: "Website is not found",
        success: false,
      });
    }

    const user = await UserDatabase.findById(req.user._id);

    if (!user) {
      return res.status(401).json({
        message: "Please login for response generation",
      });
    }

    if (user.credit < 25) {
      return res.status(400).json({
        message: "You don't have enough credits to update the website",
      });
    }

    const updateprompt = `
YOU ARE A SENIOR FRONTEND ENGINEER.

MODIFY THE EXISTING WEBSITE BASED ON USER REQUEST.

IMPORTANT:
- KEEP ALL EXISTING FEATURES
- ONLY APPLY THE REQUESTED CHANGE
- DO NOT REMOVE WORKING FEATURES
- RETURN THE FULL UPDATED HTML FILE
- Use ONLY HTML, CSS and vanilla JavaScript
- Keep exactly one <style> tag
- Keep exactly one <script> tag

CURRENT HTML:
${singlewebsite.latestcode}

USER REQUEST:
${prompt.trim()}

RETURN ONLY VALID JSON:

{
  "message": "Short confirmation describing the change",
  "code": "<COMPLETE UPDATED HTML DOCUMENT>"
}
`;

    console.log("🤖 Updating website with AI...");
    console.log("📝 User request:", prompt.trim());

    let row;

    try {
      row = await generateResponse(updateprompt);
    } catch (error) {
      console.error("❌ AI update failed:", error);

      const errorMessage = error?.message || "";

      // Handle Groq token / rate limits
      if (
        errorMessage.includes("rate_limit_exceeded") ||
        errorMessage.includes("tokens per minute") ||
        errorMessage.includes("Request too large") ||
        errorMessage.includes("Requested") &&
        errorMessage.includes("Limit")
      ) {
        return res.status(429).json({
          message:
            "AI usage limit reached. Please wait a little and try again.",
          code: "AI_RATE_LIMIT",
        });
      }

      return res.status(500).json({
        message: "AI could not update the website",
      });
    }

    const parsed = await Extractjson(row);

    if (!parsed || !parsed.code) {
      console.error("❌ AI returned invalid JSON");
      console.log("AI response:", row);

      return res.status(500).json({
        message: "AI returned invalid website data",
      });
    }

    singlewebsite.converasation.push(
      {
        role: "user",
        content: prompt.trim(),
      },
      {
        role: "ai",
        content: parsed.message || "Website updated successfully",
      }
    );

    singlewebsite.latestcode = parsed.code;

    await singlewebsite.save();

    const updatedUser = await UserDatabase.findByIdAndUpdate(
      user._id,
      { $inc: { credit: -25 } },
      { new: true }
    );

    return res.status(200).json({
      message: parsed.message || "Website updated successfully",
      code: parsed.code,
      remainingCredit: updatedUser.credit,
    });
  } catch (error) {
    console.error("❌ Update website error:", error);

    return res.status(500).json({
      message: "Unable to update website",
    });
  }
};


// Get all websites of logged-in user
export const getingallwebsite = async (req, res) => {
  try {
    const allwebsite = await website.find({
      user: req.user._id,
    });

    return res.status(200).json(allwebsite);
  } catch (error) {
    console.error("❌ Get all websites error:", error);

    return res.status(500).json({
      message: "Error while getting all websites",
    });
  }
};


// Deploy website
export const deploywebsite = async (req, res) => {
  try {
    const singlewebsite = await website.findOne({
      _id: req.params.id,
      user: req.user._id,
    });

    if (!singlewebsite) {
      return res.status(404).json({
        message: "Website is not found",
        success: false,
      });
    }

    if (!singlewebsite.slug) {
      singlewebsite.slug =
        singlewebsite.title
          .toLowerCase()
          .replace(/[^a-z0-9]/g, "")
          .slice(0, 60) +
        singlewebsite._id.toString().slice(-5);
    }

    singlewebsite.deployed = true;

    singlewebsite.deployUrl = `${process.env.FRONTEND_URL}/site/${singlewebsite.slug}`;

    await singlewebsite.save();

    return res.status(200).json({
      url: singlewebsite.deployUrl,
    });
  } catch (error) {
    console.error("❌ Deploy website error:", error);

    return res.status(500).json({
      message: "Error while deploying website",
    });
  }
};


// Get website by slug
export const getbyslug = async (req, res) => {
  try {
    const singlewebsite = await website.findOne({
      slug: req.params.slug,
      user: req.user._id,
    });

    if (!singlewebsite) {
      return res.status(404).json({
        message: "Website is not found",
        success: false,
      });
    }

    return res.status(200).json(singlewebsite);
  } catch (error) {
    console.error("❌ Get website by slug error:", error);

    return res.status(500).json({
      message: "Error while getting website",
    });
  }
};