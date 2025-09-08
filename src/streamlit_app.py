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
if 'prioritized_emails' not in st.session_state:
    st.session_state.prioritized_emails = []
if 'prioritized_emails_timestamp' not in st.session_state:
    st.session_state.prioritized_emails_timestamp = None

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
            st.error("🔐 **Authentication Required - Follow These Steps:**")
            st.markdown("""
            **Manual authentication needed:**
            
            **Step 1:** Look for a URL that starts with `https://login.microsoftonline.com/` in the console output below
            
            **Step 2:** Copy that entire URL
            
            **Step 3:** Open a new browser tab and paste the URL there
            
            **Step 4:** Complete Microsoft login in the browser
            
            **Step 5:** After login, you'll be redirected to a URL starting with `http://localhost:8503/?code=...`
            
            **Step 6:** Copy that entire redirect URL from your browser's address bar
            
            **Step 7:** Come back here and paste the redirect URL when prompted in the console
            
            **Step 8:** Click "🚀 Initialize System" again
            
            💡 **Tip:** Keep this browser tab open while you do the authentication!
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
        if st.button("🚀 Initialize System"):
            initialize_system(force_fresh=force_fresh_login)
    
    with col2:
        if st.session_state.llm_system:
            st.success("✅ System Ready")
            if st.button("🤖 Process Emails"):
                process_emails()
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
    
    # Show Prioritized Email List
    if st.session_state.llm_system:
        try:
            prioritized_emails = st.session_state.llm_system.get_prioritized_emails_from_session()
            if prioritized_emails:
                col1, col2 = st.columns([3, 1])
                with col1:
                    st.subheader("📋 Prioritized Email List")
                with col2:
                    if st.button("🗑️ Clear List", help="Clear the prioritized email list"):
                        st.session_state.prioritized_emails = []
                        st.session_state.prioritized_emails_timestamp = None
                        st.rerun()
                
                # Show timestamp of when list was created
                if hasattr(st.session_state, 'prioritized_emails_timestamp'):
                    st.caption(f"Last updated: {st.session_state.prioritized_emails_timestamp}")
                
                # Create expandable sections for each email
                for i, email in enumerate(prioritized_emails[:5]):  # Show top 5 emails
                    priority_color = "🔴" if email['priority_score'] >= 90 else "🟡" if email['priority_score'] >= 70 else "🟢"
                    
                    with st.expander(f"{priority_color} {email['priority_score']:.1f} - {email['subject'][:50]}{'...' if len(email['subject']) > 50 else ''}", expanded=i==0):
                        col1, col2 = st.columns([2, 1])
                        
                        with col1:
                            st.write(f"**From:** {email['sender']}")
                            st.write(f"**Type:** {email['email_type']}")
                            st.write(f"**Action:** {email['action_required']}")
                            st.write(f"**Preview:** {email['body_preview']}")
                        
                        with col2:
                            st.metric("Priority", f"{email['priority_score']:.1f}")
                            st.write(f"**Received:** {email['received_time']}")
                            if email['has_attachments']:
                                st.write("📎 Has attachments")
                            if not email['is_read']:
                                st.write("🔴 Unread")
                
                if len(prioritized_emails) > 5:
                    st.info(f"Showing top 5 of {len(prioritized_emails)} prioritized emails")
        except Exception as e:
            pass  # Silently handle any errors in prioritized email display
    
    # Show calendar confirmations in sidebar
    if hasattr(st.session_state, 'llm_system') and st.session_state.llm_system is not None:
        try:
            # Debug: Check if calendar confirmations exist
            if hasattr(st.session_state, 'pending_calendar_confirmations'):
                count = len(st.session_state.pending_calendar_confirmations)
                if count > 0:
                    st.sidebar.info(f"📅 {count} calendar events pending confirmation")
                else:
                    st.sidebar.info("📅 No calendar events pending")
            else:
                st.sidebar.info("📅 No calendar confirmation state found")
            
            st.session_state.llm_system.show_pending_calendar_confirmations()
        except Exception as e:
            st.sidebar.error(f"Calendar confirmation error: {e}")  # Show error instead of hiding it
    
    # Output section (no duplicate display)
    if not st.session_state.output:
        st.info("💡 Click 'Process Emails' to start AI-powered email intelligence analysis")
    
    # Footer
    st.markdown("---")
    st.markdown("""
    <div style='text-align: center; color: gray; font-size: 0.8em;'>
    🧠 Powered by State-of-the-Art AI Email Intelligence | 
    🎯 Business Impact Analysis | 
    🤖 Advanced Pattern Recognition | 
    ⚡ Real-time Priority Scoring
    </div>
    """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()