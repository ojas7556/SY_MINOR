# 🎓 Virtual AI Tutor 1.0

An intelligent, AI-powered learning assistant built with **Streamlit** and **OpenAI**. Generate comprehensive study notes, interactive assessments, educational slideshows, and contextual illustrations from a single topic prompt.

---

## ✨ Features

- **📝 Comprehensive Study Notes**: Auto-generates exhaustive learning modules covering summaries, prerequisites, learning objectives, core concepts, exercises, projects, and glossaries.
- **🎨 AI-Powered Illustrations**: Contextual images are automatically created for different aspects of your topic using DALL-E.
- **❓ Interactive Quizzes**: Generates a 10-question multiple-choice quiz grouped by difficulty (Easy, Medium, Hard) with interactive submission and instant scoring.
- **📚 Reference Finder**: Recommends academic publications (with DOIs), YouTube video series, books (with ISBNs), and professional certification resources.
- **📥 Dynamic Export Suite**:
  - **PDF Export**: Download structured study notes, quizzes, and answer keys as printable PDFs.
  - **PowerPoint Presentation (PPTX)**: Instantly export presentation decks matching a premium dark theme.
- **📚 Local Learning History**: Keeps track of previously generated topics in a sidebar history log for quick loading.

---

## 🛠️ Technology Stack

- **Frontend/Application Framework**: [Streamlit](https://streamlit.io/)
- **Core LLM & Image Gen**: [OpenAI API](https://openai.com/) (GPT-4o, GPT-image-1)
- **Document Generation**:
  - [python-pptx](https://python-pptx.readthedocs.io/) (PowerPoint generation)
  - [fpdf](https://pyfpdf.github.io/fpdf2/) (PDF rendering)
- **Text Processing**: [Unidecode](https://pypi.org/project/Unidecode/) (Unicode clean rendering)

---

## 🚀 Getting Started

### 📋 Prerequisites

- **Python 3.8+**
- **OpenAI API Key**

### 📦 Local Installation

1. **Clone the repository** (or navigate to the project directory):
   ```bash
   cd minor_project
   ```

2. **Set up a virtual environment**:
   ```bash
   python -m venv venv
   
   # Activate on Windows:
   venv\Scripts\activate
   
   # Activate on macOS/Linux:
   source venv/bin/activate
   ```

3. **Install dependencies**:
   ```bash
   pip install -r requirements.txt
   ```

4. **Configure environment variables**:
   Create a `.env` file in the root of the `minor_project` folder:
   ```env
   OPENAI_API_KEY="your_openai_api_key"
   MODEL_NAME="gpt-4o"
   ```

5. **Run the application**:
   ```bash
   streamlit run app.py
   ```

---

## ☁️ Deployment

### Pushing to GitHub

1. Initialize git and commit:
   ```bash
   git init
   git add .
   git commit -m "Deploy: Virtual AI Tutor"
   ```
2. Link to your GitHub remote and push:
   ```bash
   git branch -M main
   git remote add origin https://github.com/YOUR_USERNAME/YOUR_REPO_NAME.git
   git push -u origin main
   ```

### Deploying to Streamlit Community Cloud

1. Log in to [share.streamlit.io](https://share.streamlit.io/) using your GitHub account.
2. Click **New app** and select your repository, branch (`main`), and set the Main file path to `app.py`.
3. In **Advanced Settings**, paste your OpenAI Credentials in the **Secrets** section:
   ```toml
   OPENAI_API_KEY = "your_openai_api_key_here"
   MODEL_NAME = "gpt-4o"
   ```
4. Click **Deploy**.

---

## 🔒 Security Note

Your `.env` and `.streamlit/secrets.toml` files contain sensitive API credentials. The included `.gitignore` is pre-configured to ensure these credentials are never pushed to public version control repositories.
