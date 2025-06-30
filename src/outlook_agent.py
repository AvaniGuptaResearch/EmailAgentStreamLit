#!/usr/bin/env python3
"""
Outlook Agent with Priority Sorting and Response Drafting
Uses Microsoft Graph API and local LLMs via Ollama and ADK framework

Features:
- Reads Outlook emails using Microsoft Graph API
- Prioritizes emails with smart scoring
- Drafts contextual responses
- Shows emails sorted by priority
- Manages send queue
- Works with Office 365 and institutional Outlook

Prerequisites:
- pip install google-adk msal requests python-dotenv
- ollama pull mistral (or your preferred model)
- ollama serve
- Azure AD app registration with Microsoft Graph permissions
"""

import asyncio
import json
import os
import sys
import re
from contextlib import AsyncExitStack
from datetime import datetime, timedelta
from typing import List, Dict, Any, Optional, Tuple
import pickle
from dataclasses import dataclass, asdict
import requests

from dotenv import load_dotenv
from google.adk.agents.llm_agent import LlmAgent
from google.adk.agents.sequential_agent import SequentialAgent
from google.adk.sessions import InMemorySessionService
from google.adk.runners import Runner
from google.genai import types

import msal

load_dotenv()

# Microsoft Graph API configuration
GRAPH_API_ENDPOINT = 'https://graph.microsoft.com/v1.0'
SCOPES = [
    'https://graph.microsoft.com/Mail.Read',
    'https://graph.microsoft.com/Mail.Send',
    'https://graph.microsoft.com/Mail.ReadWrite',  # Required for creating drafts
    'https://graph.microsoft.com/User.Read'
    # Optional: Uncomment if you want calendar integration
    # 'https://graph.microsoft.com/Calendars.Read'
]

@dataclass
class OutlookEmailData:
    id: str
    subject: str
    sender: str
    sender_email: str
    recipient: str
    body: str
    body_preview: str
    date: datetime
    importance: str  # Low, Normal, High
    is_read: bool
    has_attachments: bool
    categories: List[str]
    priority_score: float = 0.0
    urgency_level: str = "normal"  # urgent, normal, low
    needs_reply: bool = False
    draft_response: str = ""
    context_tags: List[str] = None
    detected_deadline: Optional[datetime] = None
    deadline_status: str = "none"  # none, approaching, overdue, met
    deadline_confidence: float = 0.0
    email_type: str = "normal"  # normal, meeting_invite, calendar_event, reminder, task
    action_required: str = "none"  # none, reply, attend, review, approve, complete
    event_key: Optional[str] = None  # For duplicate detection
    is_duplicate: bool = False

    def __post_init__(self):
        if self.context_tags is None:
            self.context_tags = []

# ─────────────────────────── Local LLM Integration ──

class OllamaModel:
    """Local LLM wrapper for Ollama compatible with ADK"""
    
    def __init__(self, model_name: str = "mistral", host: str = "http://localhost:11434"):
        self.model_name = model_name
        self.model = model_name  # For ADK compatibility
        self.host = host
        self.url = f"{host}/api/generate"
        
    def generate_content(self, prompt: str, **kwargs) -> str:
        """Generate content for ADK compatibility"""
        payload = {
            "model": self.model_name,
            "prompt": prompt,
            "stream": False,
            "options": {
                "temperature": kwargs.get("temperature", 0.7),
                "num_predict": kwargs.get("max_tokens", 500)
            }
        }
        
        try:
            response = requests.post(self.url, json=payload, timeout=30)
            response.raise_for_status()
            return response.json()["response"].strip()
        except Exception as e:
            return f"Error: {str(e)}"
    
    async def call(self, messages: List[Dict], **kwargs) -> str:
        """Async call method for compatibility"""
        # Convert messages to a single prompt for Ollama
        prompt = self._messages_to_prompt(messages)
        return self.generate_content(prompt, **kwargs)
    
    def _messages_to_prompt(self, messages: List[Dict]) -> str:
        """Convert chat messages to a single prompt"""
        prompt_parts = []
        for msg in messages:
            role = msg.get("role", "user")
            content = msg.get("content", "")
            if role == "system":
                prompt_parts.append(f"Instructions: {content}")
            elif role == "user":
                prompt_parts.append(f"User: {content}")
            elif role == "assistant":
                prompt_parts.append(f"Assistant: {content}")
        
        return "\n\n".join(prompt_parts) + "\n\nAssistant:"

# ─────────────────────────── Microsoft Graph Integration ──

class OutlookService:
    """Microsoft Graph API wrapper for Outlook"""
    
    def __init__(self, client_id: str, client_secret: str = None, tenant_id: str = "common"):
        self.client_id = client_id
        self.client_secret = client_secret
        self.tenant_id = tenant_id
        self.access_token = None
        self.app = None
        
    def authenticate_interactive(self):
        """Interactive authentication for desktop apps"""
        self.app = msal.PublicClientApplication(
            client_id=self.client_id,
            authority=f"https://login.microsoftonline.com/{self.tenant_id}"
        )
        
        # Try to get token from cache first
        accounts = self.app.get_accounts()
        if accounts:
            result = self.app.acquire_token_silent(SCOPES, account=accounts[0])
            if result and "access_token" in result:
                self.access_token = result["access_token"]
                print("✅ Using cached authentication token")
                return
        
        # Interactive login
        print("🔐 Opening browser for Microsoft login...")
        result = self.app.acquire_token_interactive(scopes=SCOPES)
        if "access_token" in result:
            self.access_token = result["access_token"]
            print("✅ Authentication successful!")
        else:
            error_msg = result.get('error_description', result.get('error', 'Unknown error'))
            raise Exception(f"Authentication failed: {error_msg}")
    
    def authenticate_client_credentials(self):
        """Client credentials flow for service apps"""
        if not self.client_secret:
            raise Exception("Client secret required for client credentials flow")
            
        self.app = msal.ConfidentialClientApplication(
            client_id=self.client_id,
            client_credential=self.client_secret,
            authority=f"https://login.microsoftonline.com/{self.tenant_id}"
        )
        
        result = self.app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
        if "access_token" in result:
            self.access_token = result["access_token"]
            print("✅ Service account authentication successful!")
        else:
            error_msg = result.get('error_description', result.get('error', 'Unknown error'))
            raise Exception(f"Service authentication failed: {error_msg}")
    
    def authenticate_web_oauth(self):
        """Web-based OAuth for serverless environments like Hugging Face Spaces"""
        import streamlit as st
        import urllib.parse
        import secrets
        import base64
        import hashlib
        
        # Initialize session state for OAuth
        if 'oauth_state' not in st.session_state:
            st.session_state.oauth_state = None
        if 'access_token' not in st.session_state:
            st.session_state.access_token = None
        
        # Check if we already have a token
        if st.session_state.access_token:
            self.access_token = st.session_state.access_token
            print("✅ Using cached OAuth token")
            return
        
        # Check for authorization code in URL parameters
        query_params = st.query_params
        
        if 'code' in query_params and 'state' in query_params:
            # We got the authorization code back from Microsoft
            auth_code = query_params['code']
            state = query_params['state']
            
            if state == st.session_state.oauth_state:
                try:
                    # Exchange authorization code for access token
                    self.access_token = self._exchange_code_for_token(auth_code)
                    st.session_state.access_token = self.access_token
                    print("✅ OAuth authentication successful!")
                    st.success("✅ Successfully authenticated with Microsoft!")
                    st.rerun()
                    return
                except Exception as e:
                    st.error(f"❌ Failed to exchange authorization code: {e}")
                    return
            else:
                st.error("❌ Invalid OAuth state parameter")
                return
        
        # Generate OAuth authorization URL
        st.session_state.oauth_state = secrets.token_urlsafe(32)
        
        # Get and display redirect URI info
        redirect_uri = self._get_redirect_uri()
        
        # Create authorization URL
        auth_url = self._generate_auth_url(st.session_state.oauth_state)
        
        st.warning("🔐 Authentication Required")
        
        # Show redirect URI configuration info
        with st.expander("⚙️ Configuration Info", expanded=False):
            st.info(f"**Redirect URI:** `{redirect_uri}`")
            st.markdown("""
            **Important:** Make sure this redirect URI is added to your Azure AD app registration:
            1. Go to [Azure Portal](https://portal.azure.com)
            2. Navigate to "Azure Active Directory" → "App registrations" → Your app
            3. Go to "Authentication" section
            4. Add the redirect URI shown above
            5. Set platform type to "Web"
            """)
        
        st.markdown("""
        To access your Outlook emails, you need to authenticate with Microsoft.
        
        **Steps:**
        1. Click the button below to open Microsoft login
        2. Sign in with your Microsoft account
        3. Grant permissions to the app
        4. You'll be redirected back here automatically
        """)
        
        if st.button("🔐 Authenticate with Microsoft", type="primary"):
            st.markdown(f'<meta http-equiv="refresh" content="0; url={auth_url}">', unsafe_allow_html=True)
            st.markdown(f"If not redirected automatically, [click here]({auth_url})")
        
        # Show manual option
        with st.expander("Manual Authentication (if button doesn't work)"):
            st.code(auth_url)
            st.markdown("Copy this URL and open it in your browser")
    
    def _generate_auth_url(self, state: str) -> str:
        """Generate Microsoft OAuth authorization URL"""
        import urllib.parse
        import streamlit as st
        import os
        
        # Auto-detect redirect URI based on environment
        redirect_uri = self._get_redirect_uri()
        
        params = {
            'client_id': self.client_id,
            'response_type': 'code',
            'redirect_uri': redirect_uri,
            'scope': ' '.join(SCOPES),
            'state': state,
            'response_mode': 'query'
        }
        
        auth_url = f"https://login.microsoftonline.com/{self.tenant_id}/oauth2/v2.0/authorize"
        return f"{auth_url}?{urllib.parse.urlencode(params)}"
    
    def _get_redirect_uri(self) -> str:
        """Auto-detect or configure redirect URI based on environment"""
        import streamlit as st
        import os
        
        # Check for environment variable first (most reliable)
        if os.getenv('REDIRECT_URI'):
            return os.getenv('REDIRECT_URI')
        
        # Try to auto-detect from Hugging Face environment
        space_id = os.getenv('SPACE_ID') or os.getenv('HF_SPACE_ID')
        if space_id:
            # Extract username and space name from space_id (format: username/space-name)
            if '/' in space_id:
                return f"https://{space_id.replace('/', '-')}.hf.space"
            else:
                # Fallback - user needs to configure manually
                st.error(f"❌ Cannot auto-detect redirect URI. Please set REDIRECT_URI environment variable to your Hugging Face Space URL")
                return "https://your-space-url.hf.space"
        
        # Local development fallback
        return "http://localhost:8501"
    
    def _exchange_code_for_token(self, auth_code: str) -> str:
        """Exchange authorization code for access token"""
        import requests
        
        redirect_uri = self._get_redirect_uri()
        
        token_data = {
            'client_id': self.client_id,
            'scope': ' '.join(SCOPES),
            'code': auth_code,
            'redirect_uri': redirect_uri,
            'grant_type': 'authorization_code'
        }
        
        token_url = f"https://login.microsoftonline.com/{self.tenant_id}/oauth2/v2.0/token"
        
        response = requests.post(token_url, data=token_data)
        
        if response.status_code == 200:
            token_response = response.json()
            return token_response['access_token']
        else:
            raise Exception(f"Token exchange failed: {response.text}")
    
    def authenticate(self):
        """Smart authentication - detect environment and use appropriate method"""
        # Check if running in serverless environment (like Hugging Face Spaces)
        import os
        is_serverless = os.getenv('SPACE_ID') or os.getenv('HF_SPACE_ID') or os.getenv('STREAMLIT_SHARING')
        
        if is_serverless:
            print("🌐 Detected serverless environment - using web-based OAuth")
            return self.authenticate_web_oauth()
        else:
            try:
                # Local environment - use interactive auth
                self.authenticate_interactive()
            except Exception as e:
                print(f"⚠️ Interactive auth failed: {e}")
                if self.client_secret:
                    print("🔄 Trying service account authentication...")
                    try:
                        self.authenticate_client_credentials()
                        print("⚠️ Using service account - some features may be limited")
                    except Exception as e2:
                        raise Exception(f"All authentication methods failed. Interactive: {e}, Service: {e2}")
                else:
                    raise e
    
    def _make_graph_request(self, endpoint: str, method: str = "GET", data: Dict = None) -> Dict:
        """Make authenticated request to Microsoft Graph"""
        if not self.access_token:
            raise Exception("Not authenticated. Call authenticate_* method first.")
        
        headers = {
            'Authorization': f'Bearer {self.access_token}',
            'Content-Type': 'application/json'
        }
        
        url = f"{GRAPH_API_ENDPOINT}{endpoint}"
        
        try:
            if method == "GET":
                response = requests.get(url, headers=headers)
            elif method == "POST":
                response = requests.post(url, headers=headers, json=data)
            else:
                raise Exception(f"Unsupported method: {method}")
            
            response.raise_for_status()
            return response.json()
        except requests.exceptions.RequestException as e:
            raise Exception(f"Graph API request failed: {e}")
    
    def get_user_info(self) -> Dict:
        """Get current user information"""
        try:
            result = self._make_graph_request("/me")
            return {
                'email': result.get('mail', result.get('userPrincipalName', '')),
                'name': result.get('displayName', ''),
                'id': result.get('id', '')
            }
        except Exception as e:
            print(f"Warning: Could not get user info: {e}")
            return {'email': '', 'name': '', 'id': ''}
    
    def get_recent_emails(self, max_results: int = 20, hours_back: int = 24) -> List[OutlookEmailData]:
        """Fetch recent emails from Outlook"""
        # Calculate time filter
        after_date = datetime.now() - timedelta(hours=hours_back)
        after_timestamp = after_date.strftime('%Y-%m-%dT%H:%M:%S.000Z')
        
        # Build query with filters
        filter_query = f"receivedDateTime ge {after_timestamp}"
        select_fields = "id,subject,sender,toRecipients,body,bodyPreview,receivedDateTime,importance,isRead,hasAttachments,categories"
        
        endpoint = f"/me/messages?$filter={filter_query}&$select={select_fields}&$top={max_results}&$orderby=receivedDateTime desc"
        
        try:
            result = self._make_graph_request(endpoint)
            emails = []
            
            for msg in result.get('value', []):
                email_data = self._parse_outlook_email(msg)
                if email_data:
                    emails.append(email_data)
            
            return emails
        except Exception as e:
            print(f"Error fetching emails: {e}")
            return []
    
    def _parse_outlook_email(self, message: Dict) -> Optional[OutlookEmailData]:
        """Parse Outlook email from Graph API response"""
        try:
            # Extract sender info
            sender_info = message.get('sender', {}).get('emailAddress', {})
            sender_name = sender_info.get('name', '')
            sender_email = sender_info.get('address', '')
            
            # Extract recipient info
            recipients = message.get('toRecipients', [])
            recipient_emails = [r.get('emailAddress', {}).get('address', '') for r in recipients]
            
            # Extract body content
            body_content = message.get('body', {})
            body_text = body_content.get('content', '')
            
            # Clean HTML if present
            if body_content.get('contentType') == 'html':
                body_text = self._clean_html(body_text)
            
            # Parse date
            received_date = datetime.fromisoformat(
                message.get('receivedDateTime', '').replace('Z', '+00:00')
            )
            
            return OutlookEmailData(
                id=message.get('id', ''),
                subject=message.get('subject', ''),
                sender=sender_name,
                sender_email=sender_email,
                recipient=', '.join(recipient_emails),
                body=body_text[:2000],  # Limit body length
                body_preview=message.get('bodyPreview', ''),
                date=received_date,
                importance=message.get('importance', 'normal'),
                is_read=message.get('isRead', False),
                has_attachments=message.get('hasAttachments', False),
                categories=message.get('categories', [])
            )
        except Exception as e:
            print(f"Error parsing email: {e}")
            return None
    
    def _clean_html(self, html_content: str) -> str:
        """Basic HTML tag removal"""
        import re
        # Remove HTML tags
        clean = re.sub('<.*?>', '', html_content)
        # Replace common HTML entities
        clean = clean.replace('&nbsp;', ' ').replace('&lt;', '<').replace('&gt;', '>').replace('&amp;', '&')
        return clean.strip()
    
    def _clean_email_body(self, body: str) -> str:
        """Clean email body to prevent signature duplication - MINIMAL cleaning only"""
        
        import re
        
        # Only remove obvious signature blocks that interfere with draft creation
        # Don't be too aggressive since this breaks LLM-generated content
        signature_patterns = [
            # Only remove the specific problematic signature format
            r'Avani Gupta\s*Client AI Engineer\s*Research Office.*?<https://www\.youtube\.com.*?>',
        ]
        
        cleaned_body = body.strip()
        
        # Apply minimal pattern removal
        for pattern in signature_patterns:
            cleaned_body = re.sub(pattern, '', cleaned_body, flags=re.DOTALL | re.IGNORECASE)
        
        # Clean up excessive whitespace only
        cleaned_body = re.sub(r'\n{3,}', '\n\n', cleaned_body)
        
        return cleaned_body.strip()
    
    def _format_body_for_outlook(self, body: str) -> str:
        """Format email body as proper HTML to ensure Outlook signature placement"""
        
        # Convert plain text to HTML with proper structure
        # This ensures Outlook can properly append the signature at the bottom
        html_body = body.replace('\n', '<br>\n')
        
        # Wrap in proper HTML structure that Outlook expects
        formatted_html = f"""<div style="font-family: Calibri, Arial, sans-serif; font-size: 11pt;">
{html_body}
</div>"""
        
        return formatted_html
    
    def create_draft(self, to_email: str, subject: str, body: str) -> Dict:
        """Create draft email in Outlook via Graph API with proper formatting"""
        
        # Clean the body content to remove any signature elements that might interfere
        clean_body = self._clean_email_body(body)
        
        # Use HTML content type to ensure proper signature placement
        # This helps Outlook maintain proper email structure
        draft_data = {
            "subject": subject,
            "body": {
                "contentType": "HTML",
                "content": self._format_body_for_outlook(clean_body)
            },
            "toRecipients": [
                {
                    "emailAddress": {
                        "address": to_email
                    }
                }
            ]
        }
        
        try:
            response = self._make_graph_request("/me/messages", method="POST", data=draft_data)
            print(f"✅ Draft created successfully: {subject}")
            return {"success": True, "draft_id": response.get("id"), "subject": subject}
        except Exception as e:
            print(f"❌ Error creating draft: {e}")
            return {"success": False, "error": str(e)}
    
    def send_email(self, to_email: str, subject: str, body: str) -> bool:
        """Send email via Graph API"""
        email_data = {
            "message": {
                "subject": subject,
                "body": {
                    "contentType": "Text",
                    "content": body
                },
                "toRecipients": [
                    {
                        "emailAddress": {
                            "address": to_email
                        }
                    }
                ]
            }
        }
        
        try:
            self._make_graph_request("/me/sendMail", method="POST", data=email_data)
            return True
        except Exception as e:
            print(f"Error sending email: {e}")
            return False
    
    def mark_email_as_read(self, email_id: str) -> bool:
        """Mark email as read (requires Mail.ReadWrite permission)"""
        try:
            data = {"isRead": True}
            self._make_graph_request(f"/me/messages/{email_id}", method="PATCH", data=data)
            return True
        except Exception as e:
            print(f"Error marking email as read: {e}")
            return False
    
    def get_calendar_events(self, hours_ahead: int = 48) -> List[Dict]:
        """Get upcoming calendar events (requires Calendars.Read permission)"""
        try:
            start_time = datetime.now().strftime('%Y-%m-%dT%H:%M:%S.000Z')
            end_time = (datetime.now() + timedelta(hours=hours_ahead)).strftime('%Y-%m-%dT%H:%M:%S.000Z')
            
            endpoint = f"/me/calendar/events?$filter=start/dateTime ge '{start_time}' and start/dateTime le '{end_time}'"
            result = self._make_graph_request(endpoint)
            return result.get('value', [])
        except Exception as e:
            print(f"Error fetching calendar events: {e}")
            return []

# ─────────────────────────── Priority Scoring ──

class OutlookEmailPrioritizer:
    """Advanced deadline-aware email prioritization for Outlook"""
    
    # Comprehensive urgency and deadline keywords from research
    URGENT_KEYWORDS = [
        'urgent', 'asap', 'emergency', 'critical', 'deadline', 'immediately',
        'action required', 'time sensitive', 'please respond', 'by end of day',
        'eod', 'follow up', 'meeting request', 'approval needed', 'priority',
        'important', 'high priority', 'attention', 'alert', 'immediate attention',
        'time-sensitive', 'urgent attention', 'needs immediate response'
    ]
    
    # Deadline-specific keywords and phrases
    DEADLINE_KEYWORDS = [
        'deadline', 'due date', 'expires', 'expiry', 'cutoff', 'final date',
        'submission deadline', 'application deadline', 'registration closes',
        'last day', 'final day', 'closing date', 'end date', 'must be completed',
        'overdue', 'past due', 'late', 'missed deadline', 'extension needed'
    ]
    
    # Time-sensitive action keywords
    ACTION_KEYWORDS = [
        'respond by', 'reply by', 'submit by', 'complete by', 'send by',
        'deliver by', 'finish by', 'approve by', 'review by', 'sign by',
        'confirm by', 'decide by', 'pay by', 'register by', 'apply by'
    ]
    
    # Calendar and meeting keywords
    CALENDAR_KEYWORDS = [
        'meeting', 'appointment', 'call', 'conference', 'webinar', 'session',
        'event', 'gathering', 'ceremony', 'celebration', 'party', 'lunch',
        'dinner', 'breakfast', 'interview', 'presentation', 'workshop',
        'training', 'seminar', 'demo', 'review meeting', 'standup', 'sync'
    ]
    
    # Calendar-specific phrases
    CALENDAR_PHRASES = [
        'accepted:', 'declined:', 'tentative:', 'calendar invite',
        'meeting request', 'scheduled for', 'rsvp', 'save the date',
        'mark your calendar', 'block your calendar', 'calendar reminder',
        'anniversary', 'birthday', 'holiday', 'conference room',
        'zoom meeting', 'teams meeting', 'google meet'
    ]
    
    # Date patterns for deadline extraction
    DATE_PATTERNS = [
        r'(?:by|before|due|deadline|expires?)\s+(?:on\s+)?(\d{1,2}[/-]\d{1,2}[/-]\d{2,4})',
        r'(?:by|before|due|deadline|expires?)\s+(?:on\s+)?((?:jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)\w*\s+\d{1,2}(?:st|nd|rd|th)?(?:,?\s+\d{2,4})?)',
        r'(?:by|before|due|deadline|expires?)\s+(?:on\s+)?((?:monday|tuesday|wednesday|thursday|friday|saturday|sunday)\w*)',
        r'(?:by|before|due|deadline|expires?)\s+(?:on\s+)?(today|tomorrow|tonight)',
        r'(?:by|before|due|deadline|expires?)\s+(?:on\s+)?(end of (?:day|week|month))',
        r'(?:by|before|due|deadline|expires?)\s+(?:on\s+)?(\d{1,2}:\d{2}\s*(?:am|pm)?)',
        r'(within\s+\d+\s+(?:hours?|days?|weeks?))',
        r'(?:next|this)\s+(monday|tuesday|wednesday|thursday|friday|saturday|sunday|week|month)'
    ]
    
    IMPORTANT_SENDERS = [
        'ceo', 'cto', 'manager', 'director', 'vp', 'president', 
        'client', 'customer', 'support', 'hr', 'admin', 'dean',
        'professor', 'instructor', 'coordinator', 'supervisor',
        'boss', 'lead', 'head', 'chief', 'senior', 'principal'
    ]
    
    # Common deadline scenarios with context
    DEADLINE_SCENARIOS = [
        "Project deliverables with approaching deadlines",
        "Client proposals requiring immediate response",
        "Compliance documents with regulatory deadlines", 
        "Meeting confirmations with time constraints",
        "Application deadlines for opportunities",
        "Payment reminders with due dates",
        "Contract renewals with expiry dates",
        "Academic assignment submissions",
        "Event registration closures",
        "Legal document filing deadlines"
    ]
    
    def detect_deadline(self, email: OutlookEmailData) -> Tuple[Optional[datetime], float]:
        """Extract deadline from email content with confidence score"""
        text = (email.subject + " " + email.body + " " + email.body_preview).lower()
        
        # Check for deadline keywords first
        has_deadline_keywords = any(keyword in text for keyword in self.DEADLINE_KEYWORDS)
        has_action_keywords = any(keyword in text for keyword in self.ACTION_KEYWORDS)
        
        if not (has_deadline_keywords or has_action_keywords):
            return None, 0.0
        
        confidence = 0.0
        best_deadline = None
        
        # Try to extract specific dates
        for pattern in self.DATE_PATTERNS:
            matches = re.findall(pattern, text, re.IGNORECASE)
            for match in matches:
                try:
                    deadline_date = self._parse_date_string(match)
                    if deadline_date:
                        confidence = max(confidence, 0.8)
                        if not best_deadline or deadline_date < best_deadline:
                            best_deadline = deadline_date
                except:
                    continue
        
        # If no specific date found, assign relative confidence based on keywords
        if not best_deadline:
            if any(word in text for word in ['today', 'tonight', 'asap', 'immediately']):
                best_deadline = datetime.now() + timedelta(hours=2)
                confidence = 0.9
            elif any(word in text for word in ['tomorrow', 'next day']):
                best_deadline = datetime.now() + timedelta(days=1)
                confidence = 0.8
            elif any(word in text for word in ['this week', 'end of week', 'eow']):
                days_until_friday = (4 - datetime.now().weekday()) % 7
                best_deadline = datetime.now() + timedelta(days=days_until_friday)
                confidence = 0.6
            elif has_deadline_keywords or has_action_keywords:
                confidence = 0.4  # Some deadline context but unclear timing
        
        return best_deadline, confidence
    
    def classify_email_type(self, email: OutlookEmailData, current_user_email: str = "") -> Tuple[str, str, str]:
        """Classify email type, required action, and event key"""
        text = (email.subject + " " + email.body + " " + email.body_preview).lower()
        subject_lower = email.subject.lower()
        
        # Check if this email is from the current user
        is_from_self = current_user_email and email.sender_email.lower() == current_user_email.lower()
        
        # Extract event key for duplicate detection
        event_key = self._extract_event_key(email.subject)
        
        # Calendar events and meeting invites - more specific classification
        if any(phrase in subject_lower for phrase in ['accepted:', 'declined:', 'tentative:']):
            if is_from_self:
                return "self_calendar_response", "none", event_key  # Your own response - very low priority
            else:
                return "other_calendar_response", "review", event_key  # Someone else's response
        elif subject_lower.startswith('fw:') or subject_lower.startswith('fwd:'):
            if any(keyword in text for keyword in self.CALENDAR_KEYWORDS):
                return "calendar_forward", "attend", event_key  # Forwarded invite
            else:
                return "forwarded", "review", event_key
        elif any(phrase in text for phrase in self.CALENDAR_PHRASES):
            if is_from_self:
                return "self_calendar_event", "none", event_key  # Your own calendar event
            else:
                return "meeting_invite", "attend", event_key
        elif any(keyword in text for keyword in self.CALENDAR_KEYWORDS):
            if any(word in text for word in ['request', 'invite', 'schedule']):
                if is_from_self:
                    return "self_calendar_event", "none", event_key  # Your own calendar event
                else:
                    return "meeting_invite", "attend", event_key
            else:
                if is_from_self:
                    return "self_calendar_event", "none", event_key  # Your own calendar event
                else:
                    return "calendar_event", "attend", event_key
        
        # Task or approval emails
        elif any(word in text for word in ['approve', 'approval needed', 'please approve']):
            return "approval", "approve", None
        elif any(word in text for word in ['review', 'please review', 'feedback needed']):
            return "review", "review", None
        elif any(word in text for word in ['complete', 'task', 'assignment', 'deliverable']):
            return "task", "complete", None
        
        # Question emails (need reply)
        elif '?' in email.body or '?' in email.subject:
            return "question", "reply", None
        elif any(word in text for word in ['respond', 'reply', 'answer', 'let me know']):
            return "question", "reply", None
        
        # Informational emails
        elif any(word in subject_lower for word in ['fwd:', 'fw:', 'forward']):
            return "forwarded", "review", None
        elif any(word in text for word in ['fyi', 'for your information', 'heads up']):
            return "informational", "review", None
        
        return "normal", "none", None
    
    def _extract_event_key(self, subject: str) -> Optional[str]:
        """Extract a key to identify the same event across multiple emails"""
        import re
        
        # Remove common prefixes and normalize
        subject_clean = subject.lower()
        subject_clean = re.sub(r'^(re:|fw:|fwd:|accepted:|declined:|tentative:)\s*', '', subject_clean)
        subject_clean = re.sub(r'\s+', ' ', subject_clean).strip()
        
        # Remove time/date patterns that might make events look different
        subject_clean = re.sub(r'\b\d{1,2}[:/]\d{1,2}(?:[:/]\d{2,4})?\b', '', subject_clean)
        subject_clean = re.sub(r'\b(?:jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)\w*\s+\d{1,2}(?:st|nd|rd|th)?\b', '', subject_clean)
        subject_clean = re.sub(r'\b(?:monday|tuesday|wednesday|thursday|friday|saturday|sunday)\w*\b', '', subject_clean)
        subject_clean = re.sub(r'\b\d{4}\b', '', subject_clean)  # Remove years
        
        # Extract meaningful parts (event name, etc.)
        # Remove very common words that don't help identify unique events
        stop_words = {'the', 'a', 'an', 'and', 'or', 'but', 'in', 'on', 'at', 'to', 'for', 'of', 'with', 'by', 'year', 'anniversary'}
        words = [word for word in subject_clean.split() if word not in stop_words and len(word) > 1]
        
        if words:
            # Create a key from meaningful words, sorted for consistency
            key_words = sorted(words[:5])  # Use up to 5 meaningful words, sorted
            return ' '.join(key_words)
        
        return subject_clean[:20]  # Fallback to first 20 chars
    
    def _parse_date_string(self, date_str: str) -> Optional[datetime]:
        """Parse various date string formats"""
        if not date_str:
            return None
            
        now = datetime.now()
        
        # Handle relative terms
        if 'today' in date_str.lower():
            return now.replace(hour=17, minute=0, second=0, microsecond=0)
        elif 'tomorrow' in date_str.lower():
            return (now + timedelta(days=1)).replace(hour=17, minute=0, second=0, microsecond=0)
        elif 'tonight' in date_str.lower():
            return now.replace(hour=23, minute=59, second=59, microsecond=0)
        
        # Handle weekdays
        weekdays = ['monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday', 'sunday']
        for i, day in enumerate(weekdays):
            if day in date_str.lower():
                days_ahead = (i - now.weekday()) % 7
                if days_ahead == 0:
                    days_ahead = 7  # Next week if same day
                return (now + timedelta(days=days_ahead)).replace(hour=17, minute=0, second=0, microsecond=0)
        
        # Try to parse common date formats
        date_formats = [
            '%m/%d/%Y', '%m-%d-%Y', '%m/%d/%y', '%m-%d-%y',
            '%d/%m/%Y', '%d-%m-%Y', '%d/%m/%y', '%d-%m-%y',
            '%Y-%m-%d', '%Y/%m/%d'
        ]
        
        for fmt in date_formats:
            try:
                return datetime.strptime(date_str.strip(), fmt)
            except:
                continue
        
        return None
    
    def get_deadline_status(self, deadline: Optional[datetime]) -> str:
        """Determine deadline status: none, approaching, overdue, met"""
        if not deadline:
            return "none"
        
        now = datetime.now()
        time_diff = (deadline - now).total_seconds()
        
        if time_diff < 0:
            return "overdue"
        elif time_diff < 3600:  # Less than 1 hour
            return "critical"
        elif time_diff < 24 * 3600:  # Less than 24 hours
            return "approaching"
        elif time_diff < 7 * 24 * 3600:  # Less than 7 days
            return "upcoming"
        else:
            return "distant"
    
    def score_email(self, email: OutlookEmailData, current_user_email: str = "") -> float:
        """Calculate deadline-aware priority score (0-100) for Outlook emails"""
        score = 50.0  # Base score
        
        # Classify email type and action
        email_type, action, event_key = self.classify_email_type(email, current_user_email)
        email.email_type = email_type
        email.action_required = action
        email.event_key = event_key
        
        # Detect deadlines first
        deadline, confidence = self.detect_deadline(email)
        email.detected_deadline = deadline
        email.deadline_confidence = confidence
        email.deadline_status = self.get_deadline_status(deadline)
        
        # DEADLINE SCORING - This is the key improvement
        if deadline and confidence > 0.3:
            now = datetime.now()
            time_to_deadline = (deadline - now).total_seconds() / 3600  # hours
            
            if time_to_deadline < 0:  # Overdue
                score += 40 + min(20, abs(time_to_deadline) / 24)  # More overdue = higher priority
            elif time_to_deadline < 2:  # Critical (< 2 hours)
                score += 35
            elif time_to_deadline < 24:  # Due today
                score += 30
            elif time_to_deadline < 72:  # Due within 3 days
                score += 25
            elif time_to_deadline < 168:  # Due within 1 week
                score += 15
            
            # Boost score based on confidence
            score += confidence * 10
        
        # Outlook importance flag
        if email.importance == 'high':
            score += 30
        elif email.importance == 'low':
            score -= 10
            
        # Unread emails get priority
        if not email.is_read:
            score += 15
            
        # MODIFIED TIME DECAY - Don't penalize old emails with deadlines
        hours_ago = (datetime.now() - email.date.replace(tzinfo=None)).total_seconds() / 3600
        if deadline and confidence > 0.5:
            # For emails with clear deadlines, age matters less
            if hours_ago < 1:
                score += 10  # Reduced boost for recent emails with deadlines
            # No penalty for old emails with deadlines
        else:
            # Original time decay for emails without deadlines
            if hours_ago < 1:
                score += 20
            elif hours_ago < 6:
                score += 10
            elif hours_ago > 24:
                score -= 10
            
        # Enhanced keyword analysis
        text = (email.subject + " " + email.body + " " + email.body_preview).lower()
        
        # Urgent keywords
        urgent_count = sum(1 for keyword in self.URGENT_KEYWORDS if keyword in text)
        score += urgent_count * 12
        
        # Deadline-specific keywords
        deadline_count = sum(1 for keyword in self.DEADLINE_KEYWORDS if keyword in text)
        score += deadline_count * 15
        
        # Action keywords
        action_count = sum(1 for keyword in self.ACTION_KEYWORDS if keyword in text)
        score += action_count * 10
        
        # Sender importance
        sender_lower = email.sender.lower()
        if any(title in sender_lower for title in self.IMPORTANT_SENDERS):
            score += 25
            
        # Email type specific scoring
        if email_type == "calendar_event":
            score += 8  # Calendar events are informational, lower priority
            email.needs_reply = False
        elif email_type == "calendar_forward":
            score += 25  # Forwarded invites need attention/response
            email.needs_reply = False  # Action is to attend, not reply
        elif email_type == "self_calendar_response":
            score += 1   # Your own response confirmation - very low priority
            email.needs_reply = False
        elif email_type == "self_calendar_event":
            score += 1   # Your own calendar event - very low priority
            email.needs_reply = False
        elif email_type == "other_calendar_response":
            score += 5   # Someone else's response - might need review
            email.needs_reply = False
        elif email_type == "meeting_invite":
            score += 20  # Meeting invites need response
            email.needs_reply = True
        elif email_type == "question":
            score += 15  # Questions need replies
            email.needs_reply = True
        elif email_type == "approval":
            score += 25  # Approvals are high priority
            email.needs_reply = True
        elif email_type == "review":
            score += 18  # Reviews need attention
            email.needs_reply = True
        elif email_type == "task":
            score += 22  # Tasks need completion
            email.needs_reply = False  # Action needed, not necessarily reply
        elif email_type == "informational":
            score += 5  # Just FYI, lower priority
            email.needs_reply = False
        elif email_type == "forwarded":
            score += 5  # Forwarded emails often just FYI
            email.needs_reply = False
        
        # Direct mention patterns (override for some types)
        if any(word in text for word in ['you', 'your', '@']) and email_type not in ["calendar_event", "informational"]:
            score += 10
            email.needs_reply = True
            
        # Categories boost
        if email.categories:
            if any('important' in cat.lower() for cat in email.categories):
                score += 20
                
        # Attachments might indicate important documents
        if email.has_attachments:
            score += 8
        
        # Final deadline override - critical deadlines always get high priority
        if email.deadline_status in ['overdue', 'critical'] and confidence > 0.6:
            score = max(score, 85)
        elif email.deadline_status == 'approaching' and confidence > 0.7:
            score = max(score, 75)
            
        return min(100.0, max(0.0, score))
    
    def detect_duplicates(self, emails: List[OutlookEmailData]) -> None:
        """Mark duplicate emails based on event keys and types"""
        event_groups = {}
        
        # Group emails by event key
        for email in emails:
            if email.event_key:
                if email.event_key not in event_groups:
                    event_groups[email.event_key] = []
                event_groups[email.event_key].append(email)
        
        # Mark duplicates within each group
        for event_key, group_emails in event_groups.items():
            if len(group_emails) > 1:
                # Sort by priority: calendar_forward > meeting_invite > calendar_event > other_calendar_response > self_calendar_response
                priority_order = {
                    'calendar_forward': 1,
                    'meeting_invite': 2, 
                    'calendar_event': 3,
                    'other_calendar_response': 4,
                    'self_calendar_response': 5
                }
                
                # Sort by type priority, then by score
                group_emails.sort(key=lambda e: (
                    priority_order.get(e.email_type, 5),
                    -e.priority_score
                ))
                
                # Keep the first (highest priority) email, mark others as duplicates
                primary_email = group_emails[0]
                for email in group_emails[1:]:
                    email.is_duplicate = True
                    # Reduce score for duplicates but don't eliminate completely
                    email.priority_score *= 0.3
                
                # Add context to primary email
                duplicate_types = [e.email_type for e in group_emails[1:]]
                if duplicate_types:
                    primary_email.context_tags.append(f"Has {len(duplicate_types)} related emails")
    
    def categorize_urgency(self, score: float, email: OutlookEmailData) -> str:
        """Convert score to urgency level with deadline context"""
        # Override urgency based on deadline status
        if email.deadline_status == 'overdue' and email.deadline_confidence > 0.6:
            return "critical"
        elif email.deadline_status == 'critical' and email.deadline_confidence > 0.5:
            return "urgent"
        elif score >= 85:
            return "urgent" 
        elif score >= 65:
            return "normal"
        else:
            return "low"

# ─────────────────────────── ADK Agents ──

async def build_outlook_reader_agent() -> LlmAgent:
    """Agent that reads and processes Outlook emails"""
    
    async def outlook_tool(query: str) -> str:
        """Custom Outlook reading tool"""
        # Get credentials from environment
        client_id = os.getenv('AZURE_CLIENT_ID')
        client_secret = os.getenv('AZURE_CLIENT_SECRET')  # Optional for public apps
        tenant_id = os.getenv('AZURE_TENANT_ID', 'common')
        
        if not client_id:
            return "Error: AZURE_CLIENT_ID not found in environment variables"
        
        outlook = OutlookService(client_id, client_secret, tenant_id)
        
        try:
            # Use smart authentication
            outlook.authenticate()
            
            # Get current user info
            user_info = outlook.get_user_info()
            current_user_email = user_info.get('email', '')
                
            emails = outlook.get_recent_emails(max_results=15)
            prioritizer = OutlookEmailPrioritizer()
            
            # Score emails
            for email in emails:
                email.priority_score = prioritizer.score_email(email, current_user_email)
                email.urgency_level = prioritizer.categorize_urgency(email.priority_score, email)
            
            # Detect and mark duplicates
            prioritizer.detect_duplicates(emails)
                
            # Sort by priority
            emails.sort(key=lambda e: e.priority_score, reverse=True)
            
            # Format for LLM
            email_summaries = []
            for i, email in enumerate(emails[:10]):  # Top 10
                read_status = "📖 READ" if email.is_read else "📧 UNREAD"
                importance = f"📌 {email.importance.upper()}" if email.importance != 'normal' else ""
                attachments = "📎" if email.has_attachments else ""
                
                # Deadline information
                deadline_info = ""
                if email.detected_deadline and email.deadline_confidence > 0.3:
                    deadline_str = email.detected_deadline.strftime('%Y-%m-%d %H:%M')
                    confidence_str = f"({email.deadline_confidence:.1f})"
                    status_emoji = {
                        'overdue': '🔴 OVERDUE',
                        'critical': '🟠 CRITICAL', 
                        'approaching': '🟡 DUE SOON',
                        'upcoming': '🟢 UPCOMING',
                        'distant': '🔵 FUTURE'
                    }.get(email.deadline_status, email.deadline_status.upper())
                    deadline_info = f"\nDeadline: {deadline_str} {confidence_str} - {status_emoji}"
                
                summary = f"""
Email {i+1} (Priority: {email.priority_score:.1f} - {email.urgency_level.upper()})
From: {email.sender} <{email.sender_email}>
Subject: {email.subject}
Date: {email.date.strftime('%Y-%m-%d %H:%M')}
Status: {read_status} {importance} {attachments}
Needs Reply: {email.needs_reply}{deadline_info}
Preview: {email.body_preview[:150]}...
---"""
                email_summaries.append(summary)
                
            return "\n".join(email_summaries)
            
        except Exception as e:
            return f"Error reading Outlook emails: {str(e)}"
    
    return LlmAgent(
        name="OutlookReaderAgent",
        model="gemini-1.5-flash",  # Will be overridden by custom tool
        description="Reads and prioritizes Outlook emails via Microsoft Graph",
        instruction=(
            "You are an email assistant for Outlook/Office 365 with advanced deadline detection. "
            "Analyze the provided emails with priority scores and deadline information. Focus on:\n"
            "1. OVERDUE emails (🔴) - highest priority regardless of age\n"
            "2. CRITICAL emails (🟠) - due within hours\n" 
            "3. APPROACHING deadlines (🟡) - due within 24 hours\n"
            "4. High-priority emails marked as 'urgent' or 'critical'\n"
            "5. Unread emails from important senders\n\n"
            "Common deadline scenarios to watch for:\n"
            "- Project deliverables with approaching deadlines\n"
            "- Client proposals requiring immediate response\n"
            "- Compliance documents with regulatory deadlines\n"
            "- Meeting confirmations with time constraints\n"
            "- Application deadlines for opportunities\n"
            "- Payment reminders with due dates\n"
            "- Contract renewals with expiry dates\n"
            "Pay attention to deadline confidence scores - higher confidence means more reliable deadline detection."
        ),
        tools=[outlook_tool],
        output_key="prioritized_emails"
    )

async def build_outlook_response_drafter_agent() -> LlmAgent:
    """Agent that drafts email responses for Outlook"""
    
    async def draft_responses_tool(emails_data: str) -> str:
        """Custom tool to draft responses for emails that need replies"""
        try:
            # This would contain the prioritized emails from the reader agent
            # For now, return a placeholder - in full implementation, this would use local LLM
            
            draft_responses = []
            
            # Parse email data and create drafts for emails that need replies
            lines = emails_data.split('\n')
            current_email = {}
            
            for line in lines:
                if line.startswith('Email ') and '(Priority:' in line:
                    if current_email and current_email.get('needs_reply'):
                        # Generate draft for previous email
                        draft = generate_email_draft(current_email, len(draft_responses) + 1)
                        draft_responses.append(draft)
                    
                    # Start new email
                    current_email = {'priority': line}
                elif line.startswith('From:'):
                    current_email['from'] = line.replace('From: ', '')
                elif line.startswith('Subject:'):
                    current_email['subject'] = line.replace('Subject: ', '')
                elif line.startswith('Needs Reply: True'):
                    current_email['needs_reply'] = True
                elif line.startswith('Preview:'):
                    current_email['preview'] = line.replace('Preview: ', '')
            
            # Handle last email
            if current_email and current_email.get('needs_reply'):
                draft = generate_email_draft(current_email, len(draft_responses) + 1)
                draft_responses.append(draft)
            
            return '\n\n'.join(draft_responses) if draft_responses else "No emails requiring responses found."
            
        except Exception as e:
            return f"Error drafting responses: {str(e)}"
    
    def generate_email_draft(email_info: dict, draft_num: int) -> str:
        """Generate a draft response for an email"""
        subject = email_info.get('subject', 'No Subject')
        sender = email_info.get('from', 'Unknown Sender')
        preview = email_info.get('preview', '')
        
        # Simple template-based drafting
        if 'meeting' in subject.lower() or 'appointment' in subject.lower():
            draft_body = f"Thank you for the meeting invitation. I will review my calendar and get back to you shortly with my availability."
        elif '?' in preview:
            draft_body = f"Thank you for your email. I have received your inquiry and will provide a detailed response shortly."
        elif 'approval' in preview.lower():
            draft_body = f"Thank you for submitting this for approval. I will review the details and provide my feedback by [DATE]."
        elif 'deadline' in preview.lower() or 'urgent' in preview.lower():
            draft_body = f"I acknowledge the urgency of this matter. I will prioritize this and respond by the deadline mentioned."
        else:
            draft_body = f"Thank you for your email. I have received your message and will respond appropriately."
        
        return f"""
DRAFT RESPONSE #{draft_num}
TO: {sender}
SUBJECT: Re: {subject}
BODY:
{draft_body}

Best regards,
[Your Name]
---"""
    
    return LlmAgent(
        name="OutlookResponseDrafterAgent", 
        model="gemini-1.5-flash",
        description="Drafts professional email responses for Outlook",
        instruction=(
            "You are a professional email response writer for corporate Outlook environments. "
            "For each high-priority email that needs a reply, draft a polite, concise, and professional response. "
            "Consider the email type (question, meeting invite, approval request, etc.) and context. "
            "Use appropriate business language and format. Address all questions asked. "
            "Format each draft clearly with 'TO:', 'SUBJECT: Re: [original subject]', and 'BODY:' sections. "
            "For calendar events, suggest confirming attendance. For questions, acknowledge and promise detailed response. "
            "For urgent items, acknowledge the urgency and commit to timeline."
        ),
        tools=[draft_responses_tool],
        output_key="draft_responses"
    )

async def build_outlook_manager_agent() -> LlmAgent:
    """Agent that manages the Outlook email workflow"""
    
    return LlmAgent(
        name="OutlookManagerAgent",
        model="gemini-1.5-flash",
        description="Manages Outlook email workflow and priorities",
        instruction=(
            "You are an email workflow manager for institutional Outlook environments. "
            "Based on the prioritized emails and draft responses, create a clear action plan. "
            "Consider business context, institutional hierarchy, and professional communication standards. "
            "List emails in order of priority with: "
            "1. Which emails to respond to first "
            "2. Suggested response for each "
            "3. Any follow-up actions needed "
            "4. Meeting requests or calendar items to address "
            "Format as a clear, actionable daily plan suitable for professional environments."
        )
    )

# ─────────────────────────── Main Pipeline ──

async def build_outlook_pipeline() -> SequentialAgent:
    """Main Outlook processing pipeline"""
    
    reader = await build_outlook_reader_agent()
    drafter = await build_outlook_response_drafter_agent() 
    manager = await build_outlook_manager_agent()
    
    return SequentialAgent(
        name="OutlookAssistant",
        description="Complete Outlook management with priority sorting and response drafting",
        sub_agents=[reader, drafter, manager]
    )

# ─────────────────────────── Runner ──

async def run_outlook_agent():
    """Run the Outlook agent pipeline"""
    
    async with AsyncExitStack() as stack:
        pipeline = await build_outlook_pipeline()
        
        svc = InMemorySessionService()
        sess = svc.create_session(app_name="outlook_agent", user_id="user", state={})
        
        runner = Runner(app_name="outlook_agent", agent=pipeline, session_service=svc)
        content = types.Content(
            role="user", 
            parts=[types.Part(text="Check my Outlook and help me prioritize responses for today.")]
        )
        
        print("🔄 Processing Outlook emails...")
        
        async for ev in runner.run_async(
            user_id=sess.user_id, 
            session_id=sess.id, 
            new_message=content
        ):
            if ev.is_final_response():
                response = "".join(p.text for p in ev.content.parts if p.text)
                print("\n📧 OUTLOOK PRIORITY REPORT\n")
                print("=" * 50)
                print(response)
                print("=" * 50)
                break

# ─────────────────────────── Interactive Mode ──

def interactive_outlook_mode():
    """Interactive Outlook email management"""
    
    client_id = os.getenv('AZURE_CLIENT_ID')
    client_secret = os.getenv('AZURE_CLIENT_SECRET')
    tenant_id = os.getenv('AZURE_TENANT_ID', 'common')
    
    if not client_id:
        print("❌ AZURE_CLIENT_ID not found in environment variables")
        return
    
    outlook = OutlookService(client_id, client_secret, tenant_id)
    prioritizer = OutlookEmailPrioritizer()
    
    try:
        # Authenticate using smart authentication
        outlook.authenticate()
        
        print("\n📧 Outlook Agent - Interactive Mode")
        print("Commands: 'check' (check emails), 'quit' (exit)")
        
        while True:
            cmd = input("\n> ").strip().lower()
            
            if cmd == 'quit':
                break
            elif cmd == 'check':
                print("🔄 Fetching Outlook emails...")
                emails = outlook.get_recent_emails(max_results=10)
                
                for email in emails:
                    email.priority_score = prioritizer.score_email(email)
                    email.urgency_level = prioritizer.categorize_urgency(email.priority_score, email)
                
                emails.sort(key=lambda e: e.priority_score, reverse=True)
                
                print(f"\n📋 Top {len(emails)} Emails by Priority:")
                print("-" * 60)
                
                for i, email in enumerate(emails):
                    urgency_emoji = "🔴" if email.urgency_level in ["urgent", "critical"] else "🟡" if email.urgency_level == "normal" else "🟢"
                    reply_needed = "✉️ Reply needed" if email.needs_reply else ""
                    read_status = "📖" if email.is_read else "📧"
                    importance = f"📌{email.importance.upper()}" if email.importance != 'normal' else ""
                    
                    # Deadline display
                    deadline_display = ""
                    if email.detected_deadline and email.deadline_confidence > 0.3:
                        deadline_str = email.detected_deadline.strftime('%m/%d %H:%M')
                        status_emoji = {
                            'overdue': '🔴',
                            'critical': '🟠', 
                            'approaching': '🟡',
                            'upcoming': '🟢',
                            'distant': '🔵'
                        }.get(email.deadline_status, '')
                        deadline_display = f" | Deadline: {status_emoji}{deadline_str}"
                    
                    print(f"{i+1}. {urgency_emoji} {email.subject[:40]}...")
                    print(f"    From: {email.sender}")
                    print(f"    Priority: {email.priority_score:.1f} ({email.urgency_level}) {reply_needed}{deadline_display}")
                    print(f"    Status: {read_status} {importance}")
                    print(f"    Date: {email.date.strftime('%Y-%m-%d %H:%M')}")
                    print()
            else:
                print("Unknown command. Try 'check' or 'quit'")
                
    except Exception as e:
        print(f"❌ Error: {e}")

# ─────────────────────────── CLI Entry Point ──

if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "interactive":
        interactive_outlook_mode()
    else:
        print("📧 Outlook Agent with Priority Sorting")
        print("Starting automated email processing...")
        asyncio.run(run_outlook_agent())