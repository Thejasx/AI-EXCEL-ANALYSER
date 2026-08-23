# Real-Time AI Excel Analyzer

A web-based interactive tool powered by **Google Gemini API** that parses, analyzes, and dynamically sorts Excel (`.xlsx`) datasets using natural language prompts.

🔗 **Live Demo:** [aianalyserexcel.vercel.app](https://aianalyserexcel.vercel.app)

---

## Features

- 📁 **Excel File Processing:** Upload any `.xlsx` dataset and automatically convert it to structured JSON/CSV for AI processing.
- 🤖 **Gemini AI Integration:** Powered by `gemini-2.5-flash` to process natural language sorting, filtering, and summary queries.
- ⚡ **Real-Time Analysis:** Instantly view data summaries and sorted preview tables directly on the dashboard.
- 📥 **Export to Excel:** Download the transformed, sorted, or filtered dataset back into an Excel sheet with a single click.
- 🎨 **Modern UI:** Built with Tailwind CSS for a fast, responsive user experience.

---

## Tech Stack

- **Frontend:** HTML, Tailwind CSS, JavaScript
- **Backend:** Node.js, Express.js
- **AI Model:** Google Gemini API (`@google/genai` SDK)
- **File Handling:** `xlsx` (SheetJS), `multer`
- **Hosting:** Vercel

---

## Getting Started Locally

### Prerequisites

- [Node.js](https://nodejs.org/) (v18 or higher)
- A [Google Gemini API Key](https://aistudio.google.com/)

### Installation

1. **Clone the repository:**
   ```bash
   git clone [https://github.com/your-username/ai-excel-analyzer.git](https://github.com/your-username/ai-excel-analyzer.git)
   cd ai-excel-analyzer
