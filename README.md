# 🚀 AI Website Builder

An AI-powered website builder built with the MERN stack that allows users to generate complete websites from natural-language prompts, edit them using AI, preview them instantly, and deploy them with a unique URL.

## ✨ Overview

AI Website Builder simplifies website creation by allowing users to describe what they want instead of writing the entire website manually.

For example:

> "Create a modern portfolio website for a software developer with a dark theme, projects section, skills, and contact form."

The application uses an AI API to generate the website code and provides an interactive editor where users can continue modifying the website using natural-language instructions.

---

## 🌟 Features

### 🤖 AI Website Generation

- Generate websites from natural-language prompts
- AI-generated HTML, CSS and JavaScript
- Responsive website generation
- Modern UI and structured layouts
- Automatic website content generation

### ✏️ AI Website Editor

- Modify generated websites using natural-language prompts
- Ask AI to change sections, styling, content, etc.
- Full generated source code available through Monaco Editor
- Changes are saved to the database

### 👀 Live Preview

- Instantly preview generated websites
- Full-screen preview mode
- Preview generated HTML directly inside the application

### 👤 Authentication

- Google authentication using Firebase
- Secure backend user handling
- User-specific websites and projects

### 💳 Credit System

- Users receive generation credits
- Website generation consumes credits
- AI modifications consume credits
- Remaining credits are tracked for each user

### 🗂️ Website Management

- Store generated websites in MongoDB
- Retrieve individual projects
- Maintain AI conversation history
- Manage multiple generated websites

### 🌐 Deployment

- Deploy generated websites using a unique slug
- Generate a public website URL
- Store deployment information in the database

---

## 🛠️ Tech Stack

### Frontend

- React.js
- JavaScript
- Tailwind CSS
- Framer Motion
- React Icons
- Axios
- Monaco Editor
- React Router
- Redux

### Backend

- Node.js
- Express.js
- MongoDB
- Mongoose
- JWT
- Nanoid

### AI

- Groq API
- Llama models
- Prompt engineering
- AI-generated HTML/CSS/JavaScript

### Authentication

- Firebase Authentication
- Google Sign-In
- JWT-based backend authentication

---

## 🏗️ Project Structure

```text
AI-website-Builder-final-project/
│
├── Backend/
│   ├── config/
│   ├── controller/
│   ├── middleware/
│   ├── model/
│   ├── routes/
│   ├── utils/
│   ├── index.js
│   ├── package.json
│   └── package-lock.json
│
├── Frontend/
│    ├── public/
│    ├── src/
│    ├── package.json
│    ├── package-lock.json
│    └── vite.config.js
│
└── README.md