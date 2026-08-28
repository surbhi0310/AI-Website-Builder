import axios from "axios";
import React, {
  useEffect,
  useState,
} from "react";
import {
  useNavigate,
  useParams,
} from "react-router-dom";
import { motion } from "framer-motion";

import {
  FaRocket,
  FaCode,
  FaEye,
  FaTimes,
  FaHome,
} from "react-icons/fa";

import Editor from "@monaco-editor/react";

function WebEditor() {
  const { id } = useParams();

  const navigate = useNavigate();

  const [website, setWebsite] = useState(null);
  const [err, setError] = useState("");

  const [code, setCode] = useState("");

  const [message, setMessage] = useState([]);

  const [prompt, setPrompt] = useState("");

  const [loading, setLoading] = useState(false);

  const [show, setShow] = useState(false);

  const [showPreview, setShowPreview] =
    useState(false);

  /*
  ==================================================
  FETCH GENERATED WEBSITE
  ==================================================
  */

  useEffect(() => {
    const fetchWebsite = async () => {
      try {
        setError("");

        console.log(
          "🔎 Fetching website:",
          id,
        );

        const res = await axios.get(
          `https://ai-website-builder-ahl4.onrender.com/website/getwebsite/${id}`,
          {
            withCredentials: true,
          },
        );

        console.log(
          "✅ Website received:",
          res.data,
        );

        setWebsite(res.data);

        // Generated HTML
        setCode(res.data.latestcode || "");

        // Previous conversation
        setMessage(
          res.data.converasation || [],
        );
      } catch (error) {
        console.error(
          "❌ Error loading website:",
          error,
        );

        setError(
          error?.response?.data?.message ||
            "Error loading website",
        );
      }
    };

    if (id) {
      fetchWebsite();
    }
  }, [id]);

  /*
  ==================================================
  UPDATE WEBSITE USING AI
  ==================================================
  */

  const handleUpdate = async () => {
    if (!prompt.trim()) {
      return;
    }

    setLoading(true);

    const currentPrompt = prompt.trim();

    const userMessage = {
      role: "user",
      content: currentPrompt,
    };

    setMessage((prev) => [
      ...prev,
      userMessage,
    ]);

    try {
      console.log(
        "🤖 Sending update request:",
        currentPrompt,
      );

      const result = await axios.post(
        `https://ai-website-builder-ahl4.onrender.com/website/update/${id}`,
        {
          prompt: currentPrompt,
        },
        {
          withCredentials: true,
        },
      );

      console.log(
        "✅ Website updated:",
        result.data,
      );

      const aiMessage = {
        role: "ai",
        content:
          result.data.message ||
          "Website updated successfully.",
      };

      setMessage((prev) => [
        ...prev,
        aiMessage,
      ]);

      /*
      Replace old HTML with new HTML
      */

      if (result.data.code) {
        setCode(result.data.code);
      }

      setPrompt("");
    } catch (error) {
      console.error(
        "❌ Website update failed:",
        error,
      );

      const errorMessage = {
        role: "ai",
        content:
          error?.response?.data?.message ||
          "Unable to update website.",
      };

      setMessage((prev) => [
        ...prev,
        errorMessage,
      ]);
    } finally {
      setLoading(false);
    }
  };

  /*
  ==================================================
  HOME
  ==================================================
  */

  const goHome = () => {
    navigate("/");
  };

  /*
  ==================================================
  LOADING
  ==================================================
  */

  if (!website && !err) {
    return (
      <div className="h-screen flex flex-col items-center justify-center text-white bg-black">
        <div className="text-xl font-semibold">
          Loading Editor...
        </div>

        <div className="text-gray-500 text-sm mt-2">
          Loading your generated website...
        </div>
      </div>
    );
  }

  /*
  ==================================================
  ERROR
  ==================================================
  */

  if (err) {
    return (
      <div className="h-screen flex flex-col items-center justify-center bg-black text-white">
        <h2 className="text-2xl font-bold text-red-500">
          Unable to load website
        </h2>

        <p className="text-gray-400 mt-3">
          {err}
        </p>

        <button
          onClick={goHome}
          className="mt-6 bg-white/10 px-5 py-3 rounded-lg hover:bg-white/20"
        >
          <FaHome className="inline mr-2" />
          Go Home
        </button>
      </div>
    );
  }

  /*
  ==================================================
  MAIN EDITOR
  ==================================================
  */

  return (
    <div className="h-screen flex flex-col bg-black text-white overflow-hidden">
      {/* ============================================
          TOP BAR
      ============================================ */}

      <div className="flex justify-between items-center px-6 py-3 border-b border-white/10 bg-black">
        {/* TITLE */}

        <div className="flex items-center gap-3">
          <h2 className="font-semibold text-lg">
            {website.title}
          </h2>

          <span className="text-xs text-gray-500">
            ID: {id}
          </span>
        </div>

        {/* ACTIONS */}

        <div className="flex gap-3">
          {/* LIVE PREVIEW */}

          <button
            className="flex items-center gap-2 bg-white/10 px-4 py-2 rounded hover:bg-white/20 transition"
            onClick={() =>
              setShowPreview(true)
            }
          >
            <FaEye />
            Live Preview
          </button>

          {/* FULL CODE */}

          <button
            className="flex items-center gap-2 bg-white/10 px-4 py-2 rounded hover:bg-white/20 transition"
            onClick={() =>
              setShow(true)
            }
          >
            <FaCode />
            Full Code
          </button>

          {/* DEPLOY */}

          <button
            className="flex items-center gap-2 bg-yellow-500 text-black px-4 py-2 rounded hover:bg-yellow-400 transition"
          >
            <FaRocket />
            Deploy
          </button>

          {/* HOME */}

          <button
            className="flex items-center gap-2 bg-white/10 px-4 py-2 rounded hover:bg-white/20 transition"
            onClick={goHome}
          >
            <FaHome />
            Home
          </button>
        </div>
      </div>

      {/* ============================================
          MAIN AREA
      ============================================ */}

      <div className="flex flex-1 overflow-hidden">
        {/* ==========================================
            LEFT AI CHAT
        ========================================== */}

        <div className="w-[35%] border-r border-white/10 flex flex-col">
          {/* CHAT */}

          <div className="flex-1 overflow-y-auto p-4 space-y-4">
            {message.length === 0 && (
              <div className="text-gray-500 text-sm">
                Ask AI to modify your website.
              </div>
            )}

            {message.map(
              (msg, index) => (
                <motion.div
                  key={index}
                  initial={{
                    opacity: 0,
                    y: 10,
                  }}
                  animate={{
                    opacity: 1,
                    y: 0,
                  }}
                  transition={{
                    duration: 0.3,
                  }}
                  className={`p-3 rounded-lg max-w-[80%] ${
                    msg.role === "user"
                      ? "bg-blue-600 ml-auto"
                      : "bg-gray-800"
                  }`}
                >
                  <p className="text-sm whitespace-pre-wrap">
                    {msg.content}
                  </p>
                </motion.div>
              ),
            )}
          </div>

          {/* INPUT */}

          <div className="p-4 border-t border-white/10 flex gap-2">
            <input
              value={prompt}
              onChange={(e) =>
                setPrompt(e.target.value)
              }
              onKeyDown={(e) => {
                if (
                  e.key === "Enter" &&
                  !loading
                ) {
                  handleUpdate();
                }
              }}
              type="text"
              placeholder="Ask AI to modify website..."
              className="flex-1 bg-black border border-white/20 px-4 py-2 rounded outline-none focus:border-blue-500"
            />

            <button
              onClick={handleUpdate}
              disabled={loading}
              className="bg-blue-600 px-4 py-2 rounded hover:bg-blue-500 disabled:opacity-50"
            >
              {loading
                ? "Thinking..."
                : "Ask"}
            </button>
          </div>
        </div>

        {/* ==========================================
            WEBSITE PREVIEW
        ========================================== */}

        <div className="flex-1 flex flex-col">
          <div className="flex-1 bg-white">
            <iframe
              srcDoc={code}
              title="website-preview"
              className="w-full h-full border-none"
              sandbox="allow-scripts allow-forms allow-modals"
            />
          </div>
        </div>
      </div>

      {/* ============================================
          FULL CODE PANEL
      ============================================ */}

      {show && (
        <motion.div
          initial={{
            x: "100%",
          }}
          animate={{
            x: 0,
          }}
          exit={{
            x: "100%",
          }}
          transition={{
            duration: 0.3,
          }}
          className="fixed top-0 right-0 w-[50%] h-full bg-black border-l border-white/10 z-50 flex flex-col"
        >
          {/* HEADER */}

          <div className="flex justify-between items-center px-4 py-3 border-b border-white/10">
            <h3 className="font-semibold">
              Full Code
            </h3>

            <button
              onClick={() =>
                setShow(false)
              }
              className="text-white hover:text-red-500"
            >
              <FaTimes />
            </button>
          </div>

          {/* MONACO */}

          <div className="flex-1">
            <Editor
              height="100%"
              language="html"
              theme="vs-dark"
              value={code}
              onChange={(value) =>
                setCode(value || "")
              }
              onMount={(editor) => {
                setTimeout(() => {
                  const action =
                    editor.getAction(
                      "editor.action.formatDocument",
                    );

                  if (action) {
                    action.run();
                  }
                }, 200);
              }}
              options={{
                wordWrap: "on",
                minimap: {
                  enabled: false,
                },
                fontSize: 14,
                automaticLayout: true,
              }}
            />
          </div>
        </motion.div>
      )}

      {/* ============================================
          FULL PREVIEW
      ============================================ */}

      {showPreview && (
        <div className="fixed top-0 left-0 w-full h-full bg-black z-50 flex flex-col">
          {/* TOP BAR */}

          <div className="flex justify-between items-center px-6 py-3 border-b border-white/10 bg-black text-white">
            <h2 className="text-lg font-semibold">
              Full Preview
            </h2>

            <button
              onClick={() =>
                setShowPreview(false)
              }
              className="bg-red-500 px-4 py-2 rounded hover:bg-red-400"
            >
              Close
            </button>
          </div>

          {/* WEBSITE */}

          <div className="flex-1 bg-white">
            <iframe
              srcDoc={code}
              title="full-preview"
              className="w-full h-full border-none"
              sandbox="allow-scripts allow-forms allow-modals"
            />
          </div>
        </div>
      )}
    </div>
  );
}

export default WebEditor;