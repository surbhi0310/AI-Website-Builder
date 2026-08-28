# 🚀 AI Website Builder

An AI-powered website builder built with the MERN stack.

Users can describe a website using a simple text prompt, and the AI generates a complete website. Users can then edit the generated website using natural-language instructions, preview it, save it, and deploy it with a unique URL.

---

## ✨ Features

### 🤖 AI Website Generation
- Generate complete websites from text prompts
- AI-generated HTML, CSS and JavaScript
- Responsive website layouts
- Professional content and UI generation

### ✏️ AI Website Editor
- Edit generated websites using natural-language prompts
- Change colors, sections, content, buttons, layouts, etc.
- View and edit generated code using Monaco Editor
- Changes are saved to MongoDB

### 👀 Live Preview
- Preview generated websites instantly
- Full-screen website preview
- View generated HTML directly inside the application

### 👤 Authentication
- Google Login using Firebase Authentication
- Secure user authentication
- User-specific projects and websites
- JWT-based backend authentication

### 💳 Credit System
- Users receive generation credits
- Website generation consumes credits
- AI website modifications consume credits
- Remaining credits are tracked for each user

### 🗂️ Website Management
- Create multiple websites
- Save websites in MongoDB
- Edit and update existing websites
- Store AI conversation history

### 🌐 Deployment
- Deploy generated websites
- Generate a unique URL for each website
- Access deployed websites through their unique URL

---

# 🛠️ Tech Stack

## Frontend

- **React.js** – Building the user interface
- **JavaScript** – Application logic
- **Tailwind CSS** – Styling
- **React Router** – Page navigation
- **Redux** – State management
- **Axios** – API requests
- **Framer Motion** – Animations
- **React Icons** – Icons
- **Monaco Editor** – Code editor
- **Vite** – Frontend development and build tool

## Backend

- **Node.js** – JavaScript runtime
- **Express.js** – Backend server and REST APIs
- **MongoDB** – Database
- **Mongoose** – MongoDB object modeling
- **JWT** – Authentication and authorization
- **Nanoid** – Generating unique website slugs

## AI

- **Groq API** – AI model API
- **Llama / supported Groq models** – Website generation and editing
- **Prompt Engineering** – Controlling AI website generation

## Authentication

- **Firebase Authentication**
- **Google Sign-In**
- **JWT**

---

# 🏗️ Project Structure

```text
AI-website-Builder-final-project/

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
│   ├── public/
│   ├── src/
│   ├── package.json
│   ├── package-lock.json
│   └── vite.config.js
│
└── README.md


<!-- ## ⚡ Setup & Installation

### 1. Clone the Repository

```bash
git clone YOUR_GITHUB_REPOSITORY_URL
cd AI-website-Builder-final-project

### 2. Setup Backend

cd Backend
npm install

Create a .env file inside the Backend folder:

MONGO_URI=your_mongodb_connection_string
GROQ_API_KEY=your_groq_api_key
JWT_SECRET=your_jwt_secret
FRONTEND_URL=http://localhost:5173

Start the backend:

npm run dev
### 3. Setup Frontend

Open a new terminal and run:

cd AI-website-Builder-final-project/Frontend
npm install
npm run dev

Then open:

http://localhost:5173
⚠️ Important

Do not upload your .env file or API keys to GitHub.

Make sure your .gitignore contains:

.env
node_modules -->