import React, { useState } from "react";
import { BiSolidLeftArrowSquare } from "react-icons/bi";
import { motion } from "framer-motion";
import { useSelector } from "react-redux";
import { useNavigate } from "react-router-dom";
import axios from "axios";

function Generate() {
  const navigate = useNavigate();

  const sabdata = useSelector((state) => state.user.userData);

  const [prompt, setPrompt] = useState("");
  const [progress, setProgress] = useState(0);
  const [statusMessage, setStatusMessage] = useState("");
  const [isGenerating, setIsGenerating] = useState(false);

  const totalTime = 60;

  const Handlegeneratewebsite = async () => {
    if (!prompt.trim()) {
      setStatusMessage("Please enter your website idea first.");
      return;
    }

    setIsGenerating(true);
    setProgress(5);
    setStatusMessage("AI is thinking...");

    let progressTimer;

    try {
      setProgress(15);
      setStatusMessage("Understanding your idea...");

      // Fake progress while backend/AI is working
      progressTimer = setInterval(() => {
        setProgress((prev) => {
          if (prev >= 90) {
            return prev;
          }

          return prev + 5;
        });
      }, 1500);

      const result = await axios.post(
        "http://localhost:8000/website/generate",
        {
          prompt: prompt.trim(),
        },
        {
          withCredentials: true,
        },
      );

      clearInterval(progressTimer);

      console.log("✅ Website generated:", result.data);

      const websiteId = result.data.websiteId;

      console.log("Website ID:", websiteId);
      console.log(
        "Remaining credits:",
        result.data.remainingCredit,
      );

      setProgress(100);
      setStatusMessage("Website generated successfully!");

      /*
        IMPORTANT:

        Backend has already saved the generated HTML
        in MongoDB.

        Backend returned the website ID.

        Now navigate to the editor using that ID.
      */

      setTimeout(() => {
        setIsGenerating(false);
        setProgress(0);
        setStatusMessage("");

        // Open generated website in editor
        navigate(`/editor/${websiteId}`);
      }, 1200);
    } catch (error) {
      if (progressTimer) {
        clearInterval(progressTimer);
      }

      console.error("❌ Website generation failed:", error);

      if (error.response) {
        console.error(
          "Backend response:",
          error.response.data,
        );

        setStatusMessage(
          error.response.data?.message ||
            "Website generation failed",
        );
      } else {
        setStatusMessage(
          "Cannot connect to backend",
        );
      }

      setProgress(0);

      setTimeout(() => {
        setIsGenerating(false);
        setStatusMessage("");
      }, 2500);
    }
  };

  return (
    <div className="min-h-screen bg-black w-full text-white">
      {/* NAVBAR */}
      <motion.div
        initial={{ y: -40, opacity: 0 }}
        animate={{ y: 0, opacity: 1 }}
        transition={{ duration: 0.5 }}
        className="w-full border-b border-white/10"
      >
        <div className="max-w-6xl mx-auto flex justify-between items-center px-6 py-4">
          {/* BACK */}
          <div
            className="flex items-center gap-3 cursor-pointer"
            onClick={() => navigate("/")}
          >
            <BiSolidLeftArrowSquare
              size={28}
              className="hover:text-gray-400 transition"
            />

            <span className="text-lg font-semibold tracking-wide">
              Generate.AI
            </span>
          </div>

          {/* USER */}
          <motion.div
            whileHover={{ scale: 1.05 }}
            className="flex items-center gap-3 bg-white/5 border border-white/10 px-4 py-2 rounded-xl"
          >
            {sabdata?.avatar && (
              <motion.img
                src={sabdata.avatar}
                alt="profile"
                className="h-9 w-9 rounded-full object-cover"
                animate={{ scale: [1, 1.08, 1] }}
                transition={{
                  duration: 2,
                  repeat: Infinity,
                }}
              />
            )}

            <span className="text-sm font-medium">
              {sabdata?.name || "User"}
            </span>
          </motion.div>
        </div>
      </motion.div>

      {/* MAIN */}
      <div className="max-w-4xl mx-auto px-6 py-20 text-center">
        {/* TITLE */}
        <motion.h1
          initial={{ opacity: 0, y: 40 }}
          animate={{ opacity: 1, y: 0 }}
          transition={{ duration: 0.6 }}
          className="text-4xl font-bold mb-4"
        >
          Build Website with
          <span className="text-yellow-400">
            {" "}
            Real AI Power
          </span>
        </motion.h1>

        {/* DESCRIPTION */}
        <motion.p
          initial={{ opacity: 0 }}
          animate={{ opacity: 1 }}
          transition={{ delay: 0.3 }}
          className="text-gray-400 mb-10"
        >
          Describe your website idea and let AI
          generate it instantly with
          BasirTechnosoft.Ai
        </motion.p>

        {/* GENERATOR CARD */}
        <motion.div
          initial={{ opacity: 0, y: 50 }}
          animate={{ opacity: 1, y: 0 }}
          transition={{ duration: 0.6 }}
          className="bg-white/5 border border-white/10 rounded-xl p-6"
        >
          {/* PROMPT */}
          <textarea
            value={prompt}
            onChange={(e) =>
              setPrompt(e.target.value)
            }
            disabled={isGenerating}
            placeholder="Enter Your Prompt or Idea To Get Your Website..."
            className="w-full h-[140px] bg-black border border-white/10 rounded-lg p-4 outline-none focus:border-yellow-400 resize-none disabled:opacity-50"
          />

          {/* GENERATE BUTTON */}
          <motion.button
            whileHover={{
              scale: 1.05,
              boxShadow:
                "0px 0px 20px rgba(255,215,0,0.4)",
            }}
            whileTap={{ scale: 0.95 }}
            className="mt-6 px-8 py-3 bg-gray-900 text-white rounded-lg cursor-pointer font-bold relative overflow-hidden disabled:opacity-50 disabled:cursor-not-allowed"
            onClick={Handlegeneratewebsite}
            disabled={isGenerating}
          >
            {isGenerating
              ? "Generating..."
              : "Generate Website"}
          </motion.button>

          {/* STATUS */}
          {isGenerating && (
            <div className="mt-4 mb-2 text-center text-white font-medium">
              {statusMessage}
            </div>
          )}

          {/* PROGRESS */}
          {isGenerating && (
            <div className="mt-2 w-full h-3 bg-white/20 rounded-full overflow-hidden">
              <motion.div
                className="h-3 rounded-full bg-gradient-to-r from-blue-400 via-cyan-400 to-blue-600"
                initial={{ width: 0 }}
                animate={{
                  width: `${progress}%`,
                }}
                transition={{
                  ease: "easeInOut",
                  duration: 0.5,
                }}
              />
            </div>
          )}

          {/* ESTIMATED TIME */}
          {isGenerating && (
            <div className="mt-1 text-xs text-white/60">
              Estimated time:{" "}
              {Math.ceil(
                ((100 - progress) / 100) *
                  totalTime,
              )}{" "}
              sec
            </div>
          )}
        </motion.div>
      </div>
    </div>
  );
}

export default Generate;