# 📧 Email Agent - Streamlit App

An intelligent email processing system that uses LLM analysis to prioritize emails and generate draft responses in real-time.

## 🚀 Features

- **Real-time Email Processing**: Live streaming output as emails are analyzed
- **LLM-powered Analysis**: Uses Ollama/Mistral for intelligent email prioritization
- **Microsoft Graph Integration**: Connects to Outlook/Office 365 email accounts
- **Smart Drafting**: Generates contextual email responses
- **Clean UI**: Streamlit-based interface with live updates

## 📋 Prerequisites

### 1. Ollama Setup
Install and run Ollama with Mistral model:
```bash
# Install Ollama (macOS/Linux)
curl -fsSL https://ollama.ai/install.sh | sh

# Pull Mistral model
ollama pull mistral

# Start Ollama server
ollama serve
```

### 2. Azure AD App Registration
Create an Azure AD app with Microsoft Graph permissions:
- Go to [Azure Portal](https://portal.azure.com)
- Navigate to "Azure Active Directory" > "App registrations"
- Create new registration
- Add API permissions:
  - `Mail.Read`
  - `Mail.Send` 
  - `Mail.ReadWrite`
  - `User.Read`
- Generate client secret

## 🛠️ Installation

1. **Clone and navigate to the project**:
```bash
cd EmailAgentStreamlit
```

2. **Install dependencies**:
```bash
pip install -r requirements.txt
```

3. **Set up environment variables**:
Create a `.env` file in the root directory:
```env
AZURE_CLIENT_ID=your_azure_client_id
AZURE_CLIENT_SECRET=your_azure_client_secret
AZURE_TENANT_ID=your_tenant_id_or_common
```

## 🚀 Usage

1. **Start the Streamlit app**:
```bash
streamlit run src/streamlit_app.py
```

2. **Open your browser** to `http://localhost:8501`

3. **Initialize the system** by clicking "🚀 Initialize System"

4. **Process emails** by clicking "🤖 Process Emails"

5. **Watch real-time analysis** as emails are processed and prioritized

## 📁 Project Structure

```
EmailAgentStreamlit/
├── src/
│   ├── streamlit_app.py          # Main Streamlit application
│   ├── llm_enhanced_system.py    # LLM email analysis system
│   └── outlook_agent.py          # Microsoft Graph integration
├── requirements.txt              # Python dependencies
├── README.md                    # This file
└── .env                         # Environment variables (create this)
```

## 🔧 Configuration

### Environment Variables
- `AZURE_CLIENT_ID`: Your Azure AD app client ID
- `AZURE_CLIENT_SECRET`: Your Azure AD app client secret  
- `AZURE_TENANT_ID`: Your tenant ID (or 'common' for multi-tenant)

### Ollama Configuration
- Default model: `mistral`
- Default host: `http://localhost:11434`
- Configurable in `llm_enhanced_system.py`

## 🎯 How It Works

1. **Authentication**: Connects to Microsoft Graph using Azure AD
2. **Email Fetching**: Retrieves recent emails from your inbox
3. **LLM Analysis**: Each email is analyzed for:
   - Priority score (0-100)
   - Urgency level
   - Email type and required actions
   - Key points and context
4. **Real-time Display**: Results stream live to the web interface
5. **Draft Generation**: Creates contextual response drafts

## 🛡️ Security Notes

- All credentials are stored in environment variables
- OAuth2 flow used for Microsoft Graph authentication
- No email content is sent to external services (LLM runs locally)
- Drafts are saved to your Outlook Drafts folder

## 🐛 Troubleshooting

### Common Issues

**"Cannot connect to Ollama"**
- Ensure Ollama is running: `ollama serve`
- Check if Mistral model is installed: `ollama list`

**"Azure credentials missing"**
- Verify `.env` file exists with correct credentials
- Check Azure AD app permissions are granted

**"Import errors"**
- Ensure all dependencies are installed: `pip install -r requirements.txt`
- Try installing google-adk separately if needed

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Test thoroughly
5. Submit a pull request

## 📄 License

This project is licensed under the MIT License.