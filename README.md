# 🤖 PowerPoint AI Chatbot

An intelligent PowerPoint presentation generator and editor powered by Google's Gemini AI. Create, edit, and enhance presentations through natural language conversations.

## ✨ Features

### 🎯 **Core Functionality**
- **Create Presentations**: Generate professional presentations from simple prompts
- **Upload & Edit**: Upload existing PowerPoint files and edit them using AI
- **Smart Editing**: Natural language commands for slide modifications
- **Professional Design**: Beautiful blue-themed templates with consistent styling

### 💬 **Chat Interface**
- **Natural Language Processing**: Communicate with AI using everyday language
- **File Upload**: Drag-and-drop PowerPoint file upload via plus button
- **Real-time Editing**: Instant presentation updates and previews
- **Download Integration**: Seamless download of created/edited presentations

### 🎨 **Presentation Features**
- Professional slide layouts and designs
- Automatic content structuring and formatting
- Chart and visualization support
- Consistent branding and styling

## 🚀 Quick Start

### Prerequisites
- Python 3.8+
- Gemini API Key from Google AI Studio

### Installation

1. **Clone the repository**
   ```bash
   git clone https://github.com/yourusername/ppt-chatbot.git
   cd ppt-chatbot
   ```

2. **Install dependencies**
   ```bash
   pip install streamlit google-generativeai python-pptx python-dotenv matplotlib pandas plotly
   ```

3. **Set up environment variables**
   Create a `.env` file in the project root:
   ```
   GEMINI_API_KEY=your_gemini_api_key_here
   ```

4. **Run the application**
   ```bash
   streamlit run app.py
   ```

5. **Open in browser**
   Navigate to `http://localhost:8501`

## 🎯 How to Use

### Creating New Presentations
```
"Create a presentation about digital marketing"
"Make a 5-slide presentation on artificial intelligence"
"Generate slides about climate change with 6 slides"
```

### Editing Existing Presentations
1. Click the **➕** button to upload a PowerPoint file
2. Use natural language commands:
   ```
   "Edit slide 2 title to New Marketing Strategy"
   "Add a new slide about market analysis"
   "Modify slide 3 content about social media"
   "Show me slide 4"
   ```

### Download Results
- Download buttons appear automatically after creation/editing
- Files are saved with descriptive names
- Compatible with Microsoft PowerPoint

## 🛠️ Technical Stack

- **Frontend**: Streamlit
- **AI Model**: Google Gemini 2.5 Flash
- **Presentation Engine**: python-pptx
- **Charts**: Matplotlib, Plotly
- **Environment**: Python 3.8+

## 📁 Project Structure

```
ppt-chatbot/
├── app.py              # Main application file
├── .env               # Environment variables (not in repo)
├── .gitignore         # Git ignore file
├── README.md          # Project documentation
└── requirements.txt   # Python dependencies
```

## 🔧 Configuration

### Environment Variables
- `GEMINI_API_KEY`: Your Google Gemini API key

### Customization
- Modify slide templates in the `PowerPointChatbot` class
- Adjust AI prompts for different content styles
- Customize color schemes and layouts

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

## 📝 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🙏 Acknowledgments

- Google Gemini AI for natural language processing
- Streamlit for the web interface
- python-pptx library for PowerPoint manipulation

## 📞 Support

If you encounter any issues or have questions:
1. Check the [Issues](https://github.com/yourusername/ppt-chatbot/issues) page
2. Create a new issue with detailed description
3. Include error messages and steps to reproduce

---

**Made with ❤️ and AI** 🤖
