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
import io
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

def initialize_system():
    """Initialize the system"""
    try:
        if not os.getenv('AZURE_CLIENT_ID'):
            st.error("❌ Azure credentials missing")
            return False
        
        with st.spinner("🔧 Initializing..."):
            st.session_state.llm_system = LLMEnhancedEmailSystem(ollama_model="mistral")
        
        st.success("✅ Ready!")
        return True
    except Exception as e:
        st.error(f"❌ Failed: {str(e)}")
        return False

def process_emails():
    """Process emails and show output"""
    if not st.session_state.llm_system:
        st.error("Initialize first")
        return
    
    # Create containers for live output
    status_placeholder = st.empty()
    output_placeholder = st.empty()
    
    try:
        status_placeholder.info("🤖 Processing emails...")
        
        # Capture ALL output
        output_buffer = io.StringIO()
        
        with redirect_stdout(output_buffer):
            st.session_state.llm_system.process_emails_with_llm(max_emails=10, priority_threshold=60.0)
        
        # Get the complete output
        complete_output = output_buffer.getvalue()
        st.session_state.output = complete_output
        
        # Clear status and show results
        status_placeholder.empty()
        
        # Show the COMPLETE CLI output
        st.success("✅ Processing complete!")
        st.subheader("📧 Email Analysis & Priority Results")
        st.text(complete_output)  # Show ALL the CLI output as-is
        
    except Exception as e:
        status_placeholder.error(f"❌ Error: {str(e)}")

def main():
    st.title("📧 Email Agent - CLI Output")
    
    col1, col2 = st.columns([1, 4])
    
    with col1:
        st.subheader("Controls")
        
        if st.button("🚀 Initialize"):
            initialize_system()
        
        if st.session_state.llm_system:
            st.success("✅ Ready")
            
            if st.button("🤖 Process Emails"):
                process_emails()
        else:
            st.warning("⚠️ Not ready")
    
    with col2:
        if st.session_state.output:
            st.subheader("📝 Complete Results")
            st.text(st.session_state.output)
        else:
            st.info("Click 'Process Emails' to see your priority-sorted emails here")

if __name__ == "__main__":
    main()