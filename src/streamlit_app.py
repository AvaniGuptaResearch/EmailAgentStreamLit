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
current_dir = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, current_dir)

# Also add parent directory to path for deployment environments
parent_dir = os.path.dirname(current_dir)
if parent_dir not in sys.path:
    sys.path.insert(0, parent_dir)

try:
    from llm_enhanced_system import LLMEnhancedEmailSystem
    LLM_AVAILABLE = True
except ImportError as e:
    try:
        # Try importing from src directory explicitly
        import importlib.util
        spec = importlib.util.spec_from_file_location("llm_enhanced_system", 
                                                      os.path.join(current_dir, "llm_enhanced_system.py"))
        llm_module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(llm_module)
        LLMEnhancedEmailSystem = llm_module.LLMEnhancedEmailSystem
        LLM_AVAILABLE = True
    except Exception as e2:
        st.error(f"⚠️ LLM system not available: {e}")
        st.error(f"⚠️ Alternative import failed: {e2}")
        
        # Debug information
        st.write("**Debug Information:**")
        st.write(f"Current directory: {current_dir}")
        st.write(f"Parent directory: {parent_dir}")
        st.write(f"Python path: {sys.path[:5]}")  # Show first 5 entries
        
        # Check if files exist
        llm_file = os.path.join(current_dir, "llm_enhanced_system.py")
        outlook_file = os.path.join(current_dir, "outlook_agent.py")
        st.write(f"LLM file exists: {os.path.exists(llm_file)}")
        st.write(f"Outlook file exists: {os.path.exists(outlook_file)}")
        
        if os.path.exists(current_dir):
            st.write(f"Files in src directory: {os.listdir(current_dir)}")
        
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

def initialize_system(force_fresh=False):
    """Initialize the system and authenticate"""
    if not LLM_AVAILABLE:
        st.error("❌ LLM system not available - cannot initialize")
        return False
        
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
            st.session_state.llm_system = LLMEnhancedEmailSystem()
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
    if not LLM_AVAILABLE:
        st.error("❌ LLM system not available - cannot process emails")
        return
        
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