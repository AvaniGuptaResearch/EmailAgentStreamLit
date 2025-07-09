import altair as alt
import numpy as np
import pandas as pd
import streamlit as st
#!/usr/bin/env python3
"""
Direct Streamlit Email Agent
Just shows the CLI output directly - no fancy parsing
"""

import sys
import os
from datetime import datetime
from contextlib import redirect_stdout
from dotenv import load_dotenv


# Add current directory to path for imports
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

try:
    from llm_enhanced_system import LLMEnhancedEmailSystem
    LLM_AVAILABLE = True
except ImportError as e:
    st.error(f"⚠️ LLM system not available: {e}")
    LLM_AVAILABLE = False

load_dotenv()

# Page config
st.set_page_config(
    page_title="Email Agent",
    page_icon="📧",
    layout="wide"
)

# Initialize session state
if 'llm_system' not in st.session_state:
    st.session_state.llm_system = None
if 'output' not in st.session_state:
    st.session_state.output = ""
if 'ollama_url' not in st.session_state:
    st.session_state.ollama_url = "http://localhost:11434"

def initialize_system(force_fresh=False):
    """Initialize the system and authenticate"""
    try:
        if not os.getenv('AZURE_CLIENT_ID'):
            st.error("❌ Azure credentials missing")
            return False
        
        # Add browser detection and warning for Safari
        st.markdown("""
        <script>
        if (navigator.userAgent.indexOf('Safari') !== -1 && navigator.userAgent.indexOf('Chrome') === -1) {
            console.warn('Safari detected - authentication may require additional steps');
        }
        </script>
        """, unsafe_allow_html=True)
        
        with st.spinner("🔧 Initializing and authenticating..."):
            st.session_state.llm_system = LLMEnhancedEmailSystem(ollama_model="mistral", ollama_url=st.session_state.ollama_url)
            # Trigger authentication during initialization
            st.session_state.llm_system.outlook.authenticate(force_fresh=force_fresh)
        
        st.success("✅ Ready and authenticated!")
        return True
    except Exception as e:
        st.error(f"❌ Failed: {str(e)}")
        
        # Enhanced error handling for common issues
        error_msg = str(e).lower()
        if "safari" in error_msg or "timeout" in error_msg or "signal" in error_msg:
            st.info("🦷 **Browser Compatibility Issue Detected**")
            st.markdown("""
            **Try these solutions:**
            1. **Use Chrome or Firefox** for better compatibility
            2. **Safari users**: Disable "Prevent Cross-Site Tracking" in Safari → Preferences → Privacy
            3. **Ensure cookies are enabled** for this site
            4. **Try the manual authentication** option in the authentication screen
            """)
        elif "authentication" in error_msg or "oauth" in error_msg:
            st.info("""
            **Authentication Issue**: 
            - Check the manual authentication option in the auth screen
            - Ensure your Azure app registration is properly configured
            - Try refreshing the page and authenticating again
            """)
        else:
            st.info("💡 **Troubleshooting**: Try refreshing the page and initializing again")
                    
        return False

def process_emails():
    """Process emails and show real-time output"""
    if not st.session_state.llm_system:
        st.error("Initialize first")
        return
    
    # Create containers for live output
    status_placeholder = st.empty()
    output_container = st.container()
    
    try:
        status_placeholder.info("🤖 Processing emails...")
        
        # Show output in real-time using st.write_stream simulation
        with output_container:
            st.subheader("📧 Live Email Processing")
            output_placeholder = st.empty()
            
            # Capture output with periodic updates
            
            class StreamingBuffer:
                def __init__(self, placeholder):
                    self.placeholder = placeholder
                    self.content = ""
                
                def write(self, text):
                    self.content += text
                    # Update display in real-time
                    self.placeholder.text(self.content)
                    return len(text)
                
                def flush(self):
                    pass
            
            streaming_buffer = StreamingBuffer(output_placeholder)
            
            with redirect_stdout(streaming_buffer):
                st.session_state.llm_system.process_emails_with_llm(max_emails=10, priority_threshold=60.0)
            
            st.session_state.output = streaming_buffer.content
        
        # Clear status
        status_placeholder.success("✅ Processing complete!")
        
    except Exception as e:
        status_placeholder.error(f"❌ Error: {str(e)}")

def main():
    st.title("📧Email Agent Demo")
    
    # Single column layout to fix the dual-text issue
    st.subheader("Controls")
    
    # Ollama URL configuration
    st.subheader("🤖 Ollama Configuration")
    
    with st.expander("ℹ️ Ollama Setup Instructions", expanded=False):
        st.markdown("""
        **Setup your local Ollama:**
        1. Install Ollama: `curl -fsSL https://ollama.com/install.sh | sh`
        2. Start Ollama server: `ollama serve`
        3. Pull a model: `ollama pull mistral`
        4. Your Ollama will be available at `http://localhost:11434`
        
        **For remote/cloud Ollama:**
        - Make sure your Ollama server is accessible from the internet
        - Use the full URL with protocol (e.g., `http://your-server:11434`)
        - Ensure firewall allows connections to port 11434
        """)
    
    ollama_url = st.text_input(
        "Ollama URL",
        value=st.session_state.ollama_url,
        help="Enter your local Ollama server URL (e.g., http://localhost:11434)",
        key="ollama_url_input"
    )
    
    # Update session state when URL changes
    if ollama_url != st.session_state.ollama_url:
        st.session_state.ollama_url = ollama_url
        # Reset system if URL changed
        if st.session_state.llm_system:
            st.session_state.llm_system = None
            st.info("🔄 URL changed. Please reinitialize the system.")
    
    # Test connection button
    if st.button("🔗 Test Ollama Connection"):
        try:
            import requests
            response = requests.get(f"{st.session_state.ollama_url}/api/tags", timeout=5)
            if response.status_code == 200:
                models = response.json().get('models', [])
                if models:
                    st.success(f"✅ Connected! Available models: {', '.join([m['name'] for m in models])}")
                else:
                    st.warning("⚠️ Connected but no models found. Run 'ollama pull mistral' first.")
            else:
                st.error(f"❌ Connection failed: {response.status_code}")
        except Exception as e:
            st.error(f"❌ Cannot connect: {str(e)}")
            st.info("💡 Make sure Ollama is running at the specified URL")
    
    # Authentication option
    force_fresh_login = st.checkbox("🔄 Force fresh login (for new users)", value=False, 
                                   help="Check this if you want to login with a different Microsoft account")
    
    # Mode selection
    st.subheader("⚙️ Processing Mode")
    mode_col1, mode_col2 = st.columns([2, 1])
    with mode_col1:
        analysis_mode = st.selectbox(
            "Choose analysis mode:",
            options=[
                "🔬 Deep Mode (Full Analysis)", 
                "⚡ Lite Mode (Fast LLM)", 
                "🏃 Ultra-Lite Mode (Keywords + Drafting)"
            ],
            index=0,
            help="Deep: All features (20 emails). Lite: LLM analysis only (10 emails). Ultra-Lite: Keywords + LLM drafting (5 emails)."
        )
    # Update mode if system is initialized
    if st.session_state.llm_system:
        # Map UI selection to internal mode
        if "🏃 Ultra-Lite Mode" in analysis_mode:
            requested_mode = "ultra_lite"
        elif "⚡ Lite Mode" in analysis_mode:
            requested_mode = "lite"
        else:
            requested_mode = "deep"
        
        # Update mode if changed
        if st.session_state.llm_system.processing_mode != requested_mode:
            st.session_state.llm_system.set_processing_mode(requested_mode)
    
    with mode_col2:
        if st.session_state.llm_system:
            current_mode = st.session_state.llm_system.get_current_mode()
            st.metric("Current Mode", current_mode)
    
    col1, col2 = st.columns([1, 1])
    
    with col1:
        if st.button("🚀 Initialize System"):
            initialize_system(force_fresh=force_fresh_login)
    
    with col2:
        if st.session_state.llm_system:
            st.success("✅ System Ready")
            if st.button("🤖 Process Emails"):
                process_emails()
        else:
            st.warning("⚠️ Initialize system first")
    
    # Output section (no duplicate display)
    if not st.session_state.output:
        st.info("💡 Click 'Process Emails' to start email analysis")

if __name__ == "__main__":
    main()