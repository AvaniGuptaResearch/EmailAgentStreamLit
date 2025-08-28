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
        
        LLM_AVAILABLE = False  # Import failed, so system not available

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
if 'auth_completed' not in st.session_state:
    st.session_state.auth_completed = False
if 'auth_in_progress' not in st.session_state:
    st.session_state.auth_in_progress = False

def initialize_system(force_fresh=False):
    """Initialize the system and authenticate"""
    if not LLM_AVAILABLE:
        st.error("❌ LLM system not available - cannot initialize")
        return False
    
    # Check if auth is in progress to avoid duplicate initialization
    if st.session_state.auth_in_progress:
        st.warning("🔄 Authentication already in progress...")
        return False
        
    try:
        if not os.getenv('AZURE_CLIENT_ID'):
            st.error("❌ Azure credentials missing")
            return False
        
        # Mark auth as in progress
        st.session_state.auth_in_progress = True
        
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
        
        # Mark auth as completed
        st.session_state.auth_completed = True
        st.session_state.auth_in_progress = False
        st.success("✅ System Ready! You can now process emails.")
        st.rerun()  # Refresh to show the Process Emails button
        return True
    except Exception as e:
        st.error(f"❌ Failed: {str(e)}")
        
        # Enhanced error handling for common issues
        error_msg = str(e).lower()
        if "safari" in error_msg or "timeout" in error_msg or "signal" in error_msg:
            st.error("🦷 **Browser Compatibility Issue Detected**")
            st.markdown("""
            **Try these solutions:**
            1. **Use Chrome or Firefox** for better compatibility
            2. **Safari users**: Disable "Prevent Cross-Site Tracking" in Safari → Preferences → Privacy
            3. **Ensure cookies are enabled** for this site
            4. **Follow the manual authentication steps** shown above
            """)
        elif "authentication" in error_msg or "oauth" in error_msg or "msal" in error_msg:
            st.error("🔐 **Authentication Required - Follow Steps Above**")
            st.info("📺 Check your **terminal/console window** for the authentication URL, then follow the 3 simple steps above.")
        else:
            st.warning("⚠️ **Need Manual Authentication**")
            st.info("🔗 **Look for authentication URL in the console/terminal** → Copy & paste in browser → Complete login → Try Initialize again")
        
        # Reset auth state on failure        
        st.session_state.auth_in_progress = False
        st.session_state.auth_completed = False
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
    
    # Authentication Instructions
    with st.expander("🔐 Authentication Instructions (READ FIRST)", expanded=True):
        st.info("**⚠️ Manual Authentication Required**")
        st.markdown("""
        **Simple 3-step process:**
        1. Click "🚀 Initialize System" below
        2. **Copy the authentication URL** that appears in console/terminal
        3. **Paste URL in browser and complete Microsoft login**
        
        That's it! Authentication completes automatically after login.
        
        💡 **Tip**: Keep console/terminal window visible!
        """)
    
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
            index=1,
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
        # Show different button states based on authentication status
        if st.session_state.auth_in_progress:
            st.info("🔄 Authentication in progress...")
        elif st.session_state.llm_system and st.session_state.auth_completed:
            st.success("✅ System Ready")
        else:
            if st.button("🚀 Initialize System"):
                initialize_system(force_fresh=force_fresh_login)
    
    with col2:
        if st.session_state.llm_system and st.session_state.auth_completed:
            if st.button("🤖 Process Emails"):
                process_emails()
        elif st.session_state.auth_in_progress:
            st.info("⏳ Complete authentication first")
        else:
            st.warning("⚠️ Initialize system first")
    
    # Persistent Priority Summary
    if st.session_state.llm_system:
        try:
            priority_summary = st.session_state.llm_system.get_persistent_priority_summary()
            if priority_summary['total_emails'] > 0:
                st.subheader("📊 Session Email Summary")
                col1, col2, col3, col4 = st.columns(4)
                
                with col1:
                    st.metric("Total Analyzed", priority_summary['total_emails'])
                with col2:
                    st.metric("Avg Priority", f"{priority_summary['avg_priority']:.1f}")
                with col3:
                    st.metric("High Priority", priority_summary['high_priority_count'])
                with col4:
                    critical_count = priority_summary['priorities_by_range']['critical_90_plus']
                    st.metric("Critical", critical_count)
                
                # Priority distribution
                if priority_summary['total_emails'] > 1:
                    st.subheader("📈 Priority Distribution")
                    ranges = priority_summary['priorities_by_range']
                    priority_data = {
                        'Priority Range': ['Critical (90+)', 'High (70-89)', 'Medium (50-69)', 'Low (<50)'],
                        'Count': [ranges['critical_90_plus'], ranges['high_70_89'], 
                                ranges['medium_50_69'], ranges['low_below_50']]
                    }
                    priority_df = pd.DataFrame(priority_data)
                    st.bar_chart(priority_df.set_index('Priority Range'))
        except Exception as e:
            pass  # Silently handle any errors in priority summary display
    
    # Show calendar confirmations in sidebar
    if hasattr(st.session_state, 'llm_system') and st.session_state.llm_system is not None:
        try:
            st.session_state.llm_system.show_pending_calendar_confirmations()
        except Exception as e:
            pass  # Silently handle calendar confirmation errors
    
    # Output section (no duplicate display)
    if not st.session_state.output:
        st.info("💡 Click 'Process Emails' to start email analysis")
    
    # Persistent authentication reminder (only show if not authenticated)
    if not st.session_state.auth_completed and not st.session_state.auth_in_progress:
        st.markdown("---")
        st.info("💡 **Need help?** Expand the **'Authentication Instructions'** section above for step-by-step guidance.")

if __name__ == "__main__":
    main()