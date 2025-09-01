#!/usr/bin/env python3
"""
LLM-Enhanced Email System
Uses Ollama for sophisticated email analysis and personalized draft generation
"""

import requests
import json
import os
from typing import List, Dict, Any, Optional
from datetime import datetime, timedelta
from dataclasses import dataclass, field
from collections import Counter

from outlook_agent import OutlookService, OutlookEmailData, OutlookEmailPrioritizer
from dotenv import load_dotenv

load_dotenv()

@dataclass
class LLMAnalysisResult:
    """Result from LLM email analysis"""
    priority_score: float
    urgency_level: str
    email_type: str
    action_required: str
    should_reply: bool
    key_points: List[str]
    suggested_response_tone: str
    deadline_info: Optional[str]
    sender_relationship: str
    business_context: str
    confidence: float
    task_breakdown: List[str]  # New: Specific tasks to complete

@dataclass
class LLMDraftResult:
    """Result from LLM draft generation"""
    subject: str
    body: str
    tone: str
    confidence: float
    reasoning: str
    alternative_versions: List[str]

@dataclass
class EmailTemplate:
    """Email template for quick responses"""
    name: str
    category: str
    subject_template: str
    body_template: str
    tone: str
    use_cases: List[str]
    variables: List[str] = field(default_factory=list)  # Variables like {name}, {deadline}, etc.

@dataclass
class FollowUpItem:
    """Item requiring follow-up"""
    email_id: str
    subject: str
    sender: str
    follow_up_date: datetime
    reason: str
    priority: str
    status: str = "pending"  # pending, completed, overdue
    notes: str = ""

@dataclass
class EmailSummary:
    """Summary of long email content"""
    key_points: List[str]
    action_items: List[str]
    mentioned_people: List[str]
    mentioned_dates: List[str]
    estimated_read_time: str
    summary_text: str

@dataclass
class WritingStyle:
    """User's writing style analysis"""
    tone: str  # formal, casual, friendly, professional
    formality_level: float  # 0-1 (casual to formal)
    greeting_style: str  # Hi, Hello, Dear, etc.
    closing_style: str  # Best regards, Thanks, etc.
    average_length: int  # Average email length in words
    use_contractions: bool  # Don't vs Do not
    punctuation_style: str  # minimal, standard, expressive
    signature: str  # Common signature pattern
    common_phrases: List[str]  # Frequently used phrases
    response_time_preference: str  # immediate, same_day, within_24h

@dataclass
class EmailContext:
    """Multi-source context for draft generation"""
    current_email: OutlookEmailData
    email_history: List[str]  # Previous emails with this sender
    knowledge_base_entries: List[str]  # Relevant knowledge entries
    writing_style: WritingStyle
    conversation_thread: List[OutlookEmailData]  # Email thread
    user_preferences: Dict[str, Any]  # User settings and preferences

@dataclass
class SecurityAnalysis:
    """Email security analysis result"""
    is_suspicious: bool
    risk_level: str  # low, medium, high, critical
    threats_detected: List[str]
    suspicious_indicators: List[str]
    recommendations: List[str]
    confidence: float

@dataclass
class SentimentAnalysis:
    """Email sentiment analysis result"""
    sentiment: str  # positive, negative, neutral, urgent
    tone: str  # formal, casual, frustrated, excited, concerned
    urgency_indicators: List[str]
    emotion_score: float  # -1 to 1 (negative to positive)
    requires_immediate_attention: bool

class UnifiedLLMService:
    """Service for interacting with both OpenAI and Ollama LLMs"""
    
    def __init__(self, model_type: str = None, model: str = None, host: str = None):
        import os
        import logging
        from dotenv import load_dotenv
        
        load_dotenv()
        
        # Function to get configuration from Streamlit secrets or environment
        def get_config(key, default=None):
            # Check if we're running in Streamlit environment
            is_streamlit = False
            try:
                import streamlit as st
                if hasattr(st, 'secrets'):
                    is_streamlit = True
                    if key in st.secrets:
                        value = st.secrets[key]
                        # Skip placeholder values
                        if value and not str(value).startswith("your_") and "your_" not in str(value).lower():
                            return value
            except Exception:
                pass
            
            # For local development or fallback, try environment variables
            env_value = os.getenv(key)
            if env_value and env_value != default and not env_value.startswith("your_"):
                return env_value
            
            # If running locally and no env var, try to read from secrets.toml directly
            if not is_streamlit:
                try:
                    import toml
                    secrets_file = os.path.join(os.path.dirname(__file__), '..', '.streamlit', 'secrets.toml')
                    if os.path.exists(secrets_file):
                        with open(secrets_file, 'r') as f:
                            secrets = toml.load(f)
                            if key in secrets:
                                value = secrets[key]
                                if value and not str(value).startswith("your_"):
                                    return value
                except Exception:
                    pass
            
            return default
        
        # Configure logging
        log_level = get_config('LOG_LEVEL', 'INFO')
        logging.basicConfig(
            level=getattr(logging, log_level.upper()),
            format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
        )
        self.logger = logging.getLogger(__name__)
        
        self.model_type = model_type or get_config('MODEL_TYPE', 'openai')
        self._response_cache = {}  # Cache for LLM responses
        
        self.logger.info(f"🤖 Initializing LLM Service - Type: {self.model_type}")
        
        if self.model_type == 'ollama':
            self.model = model or get_config('OLLAMA_MODEL', 'mistral')
            self.host = host or get_config('OLLAMA_HOST', 'http://localhost:11434')
            self.url = f"{self.host}/api/generate"
            self.logger.info(f"🦙 Using Ollama - Model: {self.model}, Host: {self.host}")
            self._test_ollama_connection()
        elif self.model_type == 'openai':
            self.model = model or get_config('OPENAI_MODEL', 'gpt-4o-mini')
            self.api_key = get_config('OPENAI_API_KEY')
            
            # Debug logging for API key
            if self.api_key:
                masked_key = self.api_key[:10] + "***" + self.api_key[-4:] if len(self.api_key) > 14 else "***"
                self.logger.info(f"🔑 API Key loaded: {masked_key}")
            else:
                self.logger.error("❌ No OpenAI API key found!")
                raise ValueError("OPENAI_API_KEY not found in Streamlit secrets or environment variables")
            
            self.logger.info(f"🔥 Using OpenAI - Model: {self.model}")
            self._test_openai_connection()
        else:
            raise ValueError(f"Unsupported model_type: {self.model_type}. Use 'openai' or 'ollama'.")
    
    def _test_ollama_connection(self):
        """Test connection to Ollama"""
        try:
            response = requests.get(f"{self.host}/api/tags", timeout=5)
            if response.status_code == 200:
                models = response.json().get('models', [])
                available_models = [m['name'] for m in models]
                
                if any(self.model in model for model in available_models):
                    self.logger.info(f"✅ Connected to Ollama - Model '{self.model}' available")
                else:
                    self.logger.warning(f"⚠️ Model '{self.model}' not found. Available models: {available_models}")
                    if available_models:
                        self.model = available_models[0].split(':')[0]
                        self.logger.info(f"🔄 Switching to available model: {self.model}")
            else:
                self.logger.error(f"❌ Ollama connection failed: {response.status_code}")
        except Exception as e:
            self.logger.error(f"❌ Cannot connect to Ollama: {e}")
            self.logger.info("📝 Make sure Ollama is running: 'ollama serve'")
    
    def _test_openai_connection(self):
        """Test connection to OpenAI"""
        try:
            from openai import OpenAI
            client = OpenAI(api_key=self.api_key)
            # Test with a simple completion
            response = client.chat.completions.create(
                model=self.model,
                messages=[{"role": "user", "content": "test"}],
                max_tokens=1
            )
            self.logger.info(f"✅ Connected to OpenAI - Model '{self.model}' available")
        except Exception as e:
            self.logger.error(f"❌ Cannot connect to OpenAI: {e}")
    
    def generate_response(self, prompt: str, max_tokens: int = 1000, temperature: float = 0.7) -> str:
        """Generate response using either OpenAI or Ollama with caching"""
        # Create cache key from prompt and parameters
        import hashlib
        cache_key = hashlib.md5(f"{prompt}_{max_tokens}_{temperature}_{self.model_type}".encode()).hexdigest()
        
        # Check cache first
        if cache_key in self._response_cache:
            self.logger.debug(f"🔄 Cache hit for {self.model_type} - {self.model}")
            return self._response_cache[cache_key]
        
        try:
            self.logger.debug(f"🚀 Generating response with {self.model_type} - {self.model}")
            if self.model_type == 'ollama':
                response_text = self._generate_ollama_response(prompt, max_tokens, temperature)
            elif self.model_type == 'openai':
                response_text = self._generate_openai_response(prompt, max_tokens, temperature)
            else:
                return "Error: Unsupported model type"
            
            # Cache the response
            self._response_cache[cache_key] = response_text
            self.logger.debug(f"✅ Response generated successfully with {self.model_type}")
            return response_text
            
        except Exception as e:
            self.logger.error(f"❌ Error generating response with {self.model_type}: {str(e)}")
            return f"Error: {str(e)}"
    
    def _generate_ollama_response(self, prompt: str, max_tokens: int, temperature: float) -> str:
        """Generate response using Ollama"""
        payload = {
            "model": self.model,
            "prompt": prompt,
            "stream": False,
            "options": {
                "temperature": temperature,
                "num_predict": max_tokens,
                "stop": ["Human:", "Assistant:", "###", "\n\n\n", "---"],
                "top_k": 40,
                "top_p": 0.9,
                "repeat_penalty": 1.1
            }
        }
        
        try:
            response = requests.post(self.url, json=payload, timeout=60)
            response.raise_for_status()
            result = response.json()
            
            if 'response' in result:
                return result['response'].strip()
            else:
                print(f"❌ Unexpected response format: {result}")
                return "Error: Invalid response format"
                
        except requests.exceptions.Timeout:
            return "Error: LLM request timed out"
        except requests.exceptions.RequestException as e:
            return f"Error: LLM request failed - {str(e)}"
        except json.JSONDecodeError:
            return "Error: Invalid JSON response from LLM"
    
    def _generate_openai_response(self, prompt: str, max_tokens: int, temperature: float) -> str:
        """Generate response using OpenAI"""
        from openai import OpenAI
        
        client = OpenAI(api_key=self.api_key)
        
        response = client.chat.completions.create(
            model=self.model,
            messages=[{"role": "user", "content": prompt}],
            max_tokens=max_tokens,
            temperature=temperature
        )
        
        return response.choices[0].message.content.strip()
    
    def call_with_json_parsing(self, prompt: str) -> Dict[str, Any]:
        """Call LLM and parse JSON response"""
        response = self.generate_response(prompt, max_tokens=1000, temperature=0.3)
        try:
            # Extract JSON from response
            json_start = response.find('{')
            json_end = response.rfind('}') + 1
            if json_start != -1 and json_end > json_start:
                json_str = response[json_start:json_end]
                return json.loads(json_str)
            return {}
        except:
            return {}

class LLMEmailAnalyzer:
    """Advanced email analyzer using LLM"""
    
    def __init__(self, llm_service: UnifiedLLMService):
        self.llm = llm_service
    
    def analyze_email(self, email: OutlookEmailData, user_context: str = "", email_thread_context: str = "") -> LLMAnalysisResult:
        """Analyze email using LLM for sophisticated understanding"""
        
        analysis_prompt = f"""
You are an expert email analyst for a busy professional. Analyze this email and provide a detailed assessment.

EMAIL TO ANALYZE:
From: {email.sender} <{email.sender_email}>
Subject: {email.subject}
Date: {email.date}
Importance: {email.importance}
Read Status: {'Read' if email.is_read else 'Unread'}

Content:
{email.body[:1000]}

Preview: {email.body_preview}

USER CONTEXT:
{user_context or "Professional manager at a technology company"}

EMAIL THREAD CONTEXT:
{email_thread_context or "This appears to be a standalone email or start of new thread"}

PRIORITY ANALYSIS CRITERIA:
1. UNREAD emails get +25 priority points (this is UNREAD: {not email.is_read})
2. DEADLINE/URGENCY keywords: urgent, asap, critical, deadline, due, expires, today, tomorrow, approaching, timely delivery, submit, delivery (+35 points)
3. BUSINESS IMPORTANCE: meeting, approval, decision, client, proposal, contract, budget (+20 points)
4. REQUESTS/QUESTIONS: Contains "?", "please", "can you", "need", "request" (+15 points)
5. SENDER IMPORTANCE: CEO, Director, Manager, Client, Lead roles (+20 points)
6. AGE: Older unread emails get higher priority (+15-25 points)

ANALYSIS REQUIRED:
1. Priority Score (0-100): Calculate based on above criteria
2. Urgency Level: critical(85+)/urgent(70+)/normal(50+)/low(<50)
3. Email Type: meeting/question/request/deadline/appreciation/information/complaint/announcement
4. Action Required: reply/attend/approve/review/complete/none
5. Key Points: Main topics/requests in the email (bullet points)
6. Response Tone: formal/professional/friendly/casual
7. Deadline Info: Any deadlines mentioned (be specific, extract date/time)
8. Sender Relationship: manager/client/colleague/vendor/external
9. Business Context: How this affects work/projects
10. Confidence: How certain you are of this analysis (0-1)
11. Task Breakdown: Specific actionable tasks with extracted details (include links, dates, times, contact info, locations, file names)

PROFESSIONAL EMAIL PRIORITIZATION:
- Emails with specific deadlines (dates mentioned) = CRITICAL priority (95+)
- Unread emails with deadlines = CRITICAL priority (90+)
- Project delivery/submission requests = CRITICAL priority (85+)
- Client/manager/CEO requests = CRITICAL priority (85+)
- Meeting invitations from important people = HIGH priority (75+)
- Project updates with deadlines = HIGH priority (70+)
- Questions requiring answers = MEDIUM priority (60+)
- Internal updates = MEDIUM priority (50+)
- Newsletters/marketing = LOW priority (20-)

SMART CATEGORIZATION:
- URGENT: Contains deadline words, marked urgent, from VIP sender
- ACTIONABLE: Contains questions, requests, meeting invites
- INFORMATIONAL: Updates, announcements, newsletters
- SOCIAL: Personal emails, congratulations, thank you notes
- AUTOMATED: System emails, notifications, confirmations
- MARKETING: Promotional emails, newsletters, marketing campaigns
- NOREPLY: Automated emails that should not be replied to

EMAIL REPLY DECISION:
- NEVER reply to: marketing emails, no-reply addresses, automated notifications, promotional content
- ALWAYS reply to: direct questions, meeting requests, project discussions, client communications
- Email addresses containing: "marketing", "no-reply", "noreply", "automated", "system" = NO REPLY
- Marketing content indicators: "offer", "promotion", "sign up", "subscribe", "unsubscribe" = NO REPLY

TASK BREAKDOWN REQUIREMENTS:
For each task, extract and include specific actionable details:
- LINKS: Include portal URLs, document links, meeting links, survey links
- DATES/TIMES: Extract specific dates, times, deadlines, meeting schedules
- CALENDAR EVENTS: Include meeting details (date, time, location, participants)
- CONTACT INFO: Phone numbers, email addresses for follow-up
- DOCUMENT DETAILS: File names, document types, attachments mentioned
- LOCATIONS: Physical addresses, room numbers, virtual meeting rooms

ENHANCED TASK BREAKDOWN EXAMPLES:
- For IT survey emails: ["Complete IT survey at: [survey_link]", "Rate helpdesk service quality (1-5 scale)", "Submit feedback by: [specific_deadline_date]", "Save confirmation receipt for records"]
- For meeting requests: ["Check calendar availability for: [specific_date] at [specific_time]", "Confirm attendance to: [organizer_email]", "Prepare agenda items for [meeting_topic]", "Set calendar reminder for: [date] [time] in [location/zoom_link]"]
- For project updates: ["Review attached documents: [file_names]", "Update project status in: [project_portal_link]", "Coordinate with team members: [contact_list]", "Schedule follow-up meeting by: [date]"]
- For client inquiries: ["Research client requirements for: [specific_topic]", "Prepare detailed response by: [deadline]", "Gather supporting documentation from: [source_locations]", "Schedule call using: [calendar_link] for [proposed_times]"]
- For portal access: ["Access portal at: [portal_url]", "Login with credentials: [username_format]", "Complete required sections by: [deadline]", "Download forms from: [download_link]"]
- For deadline submissions: ["Prepare submission materials for: [topic]", "Submit via: [submission_portal/email]", "Deadline: [exact_date_time]", "Required format: [specifications]"]

Respond in this EXACT JSON format:
{{
    "priority_score": 75.5,
    "urgency_level": "normal", 
    "email_type": "request",
    "action_required": "reply",
    "should_reply": true,
    "key_points": ["Point 1", "Point 2", "Point 3"],
    "suggested_response_tone": "professional",
    "deadline_info": "By Friday 5 PM" or null,
    "sender_relationship": "colleague",
    "business_context": "Project coordination requiring immediate attention",
    "confidence": 0.85,
    "task_breakdown": ["Specific task 1", "Specific task 2", "Specific task 3"]
}}
"""

        try:
            response = self.llm.generate_response(analysis_prompt, max_tokens=500, temperature=0.3)
            
            # Extract JSON from response
            json_start = response.find('{')
            json_end = response.rfind('}') + 1
            
            if json_start != -1 and json_end > json_start:
                json_str = response[json_start:json_end]
                analysis_data = json.loads(json_str)
                
                return LLMAnalysisResult(
                    priority_score=float(analysis_data.get('priority_score', 50)),
                    urgency_level=analysis_data.get('urgency_level', 'normal'),
                    email_type=analysis_data.get('email_type', 'normal'),
                    action_required=analysis_data.get('action_required', 'none'),
                    should_reply=bool(analysis_data.get('should_reply', True)),
                    key_points=analysis_data.get('key_points', []),
                    suggested_response_tone=analysis_data.get('suggested_response_tone', 'professional'),
                    deadline_info=analysis_data.get('deadline_info'),
                    sender_relationship=analysis_data.get('sender_relationship', 'colleague'),
                    business_context=analysis_data.get('business_context', ''),
                    confidence=float(analysis_data.get('confidence', 0.5)),
                    task_breakdown=analysis_data.get('task_breakdown', [])
                )
            else:
                print(f"❌ Could not parse JSON from LLM response: {response}")
                return self._fallback_analysis(email)
                
        except json.JSONDecodeError as e:
            print(f"❌ JSON decode error: {e}")
            print(f"LLM Response: {response}")
            return self._fallback_analysis(email)
        except Exception as e:
            print(f"❌ LLM analysis error: {e}")
            return self._fallback_analysis(email)
    
    def _fallback_analysis(self, email: OutlookEmailData) -> LLMAnalysisResult:
        """Enhanced fallback analysis with professional prioritization"""
        
        text = (email.subject + " " + email.body).lower()
        
        # Start with base priority
        priority_score = 40.0
        
        # UNREAD EMAILS get higher priority
        if not email.is_read:
            priority_score += 25
        
        # DEADLINE/URGENCY keywords (enhanced)
        urgent_keywords = ['urgent', 'asap', 'critical', 'immediate', 'deadline', 'due', 'expires', 'today', 'tomorrow', 'approaching', 'timely delivery', 'submit', 'delivery']
        deadline_boost = 0
        for word in urgent_keywords:
            if word in text:
                deadline_boost += 35
                break  # Don't double-count
        
        # Extra boost for specific date mentions (like "july 2nd")
        import re
        from datetime import datetime, date
        
        date_patterns = [
            r'(january|february|march|april|may|june|july|august|september|october|november|december)\s+\d{1,2}',
            r'\d{1,2}/(1[0-2]|0?[1-9])/\d{2,4}',  # MM/DD/YYYY
            r'\d{1,2}-(1[0-2]|0?[1-9])-\d{2,4}',  # DD-MM-YYYY
            r'(today|tomorrow|this week|next week)',
        ]
        
        for pattern in date_patterns:
            if re.search(pattern, text, re.IGNORECASE):
                deadline_boost += 40  # Major boost for specific dates
                
                # CRITICAL BOOST: If deadline is today or tomorrow, make it CRITICAL priority
                current_date = datetime.now().date()
                if ('july 2' in text and current_date.month == 7 and current_date.day in [1, 2]) or \
                   'today' in text or 'tomorrow' in text:
                    deadline_boost += 60  # MASSIVE boost for same-day/next-day deadlines
                    print(f"   🚨 CRITICAL DEADLINE DETECTED: Adding maximum priority boost!")
                
                break
        
        priority_score += deadline_boost
        
        # BUSINESS IMPORTANCE keywords (enhanced for project delivery)
        important_keywords = ['meeting', 'approval', 'decision', 'budget', 'contract', 'client', 'proposal', 'review', 'modules', 'deliverable', 'submission', 'project team']
        if any(word in text for word in important_keywords):
            priority_score += 25
        
        # QUESTION/REQUEST indicators
        if any(indicator in text for indicator in ['?', 'please', 'can you', 'could you', 'need', 'request']):
            priority_score += 15
        
        # ENHANCED SENDER importance detection
        sender_lower = email.sender.lower()
        sender_email_lower = email.sender_email.lower()
        
        # VIP emails (C-Suite, University Leadership)
        vip_keywords = ['ceo', 'cto', 'cfo', 'president', 'vice president', 'director', 'head of', 'dean', 'provost', 'chancellor']
        vip_domains = ['presidentoffice@mbzuai.ac.ae', 'dean@mbzuai.ac.ae']
        
        if any(title in sender_lower for title in vip_keywords) or sender_email_lower in vip_domains:
            priority_score += 45  # Increased VIP boost
            print(f"   👑 VIP SENDER DETECTED: {email.sender}")
        
        # Management and senior roles
        mgmt_titles = ['manager', 'lead', 'supervisor', 'team lead', 'project manager', 'coordinator', 'chair']
        if any(title in sender_lower for title in mgmt_titles):
            priority_score += 30
        
        # Academic faculty and research
        academic_titles = ['professor', 'dr.', 'phd', 'researcher', 'faculty']
        if any(title in sender_lower for title in academic_titles):
            priority_score += 25
        
        # External clients and collaborators
        if 'client' in sender_lower or (any(domain in sender_email_lower for domain in ['.com', '.org', '.net']) and 'mbzuai.ac.ae' not in sender_email_lower):
            priority_score += 20
        
        # Internal colleagues
        if 'mbzuai.ac.ae' in sender_email_lower:
            priority_score += 15
        
        # AGE of email (older unread emails are more important)
        from datetime import datetime, timezone
        try:
            email_age_hours = (datetime.now(timezone.utc) - email.date).total_seconds() / 3600
            if email_age_hours > 24 and not email.is_read:
                priority_score += 15
            elif email_age_hours > 72 and not email.is_read:
                priority_score += 25
        except:
            pass
        
        # Cap at 100
        priority_score = min(priority_score, 100.0)
        
        # Determine urgency level
        if priority_score >= 85:
            urgency_level = 'critical'
        elif priority_score >= 70:
            urgency_level = 'urgent'
        elif priority_score >= 50:
            urgency_level = 'normal'
        else:
            urgency_level = 'low'
        
        # Enhanced deadline detection
        deadline_info = None
        deadline_patterns = ['due ', 'deadline ', 'by ', 'expires ', 'until ', 'approaching', 'july 2nd', 'july 2', 'submit by']
        for pattern in deadline_patterns:
            if pattern in text:
                # Try to extract deadline context
                start_idx = text.find(pattern)
                deadline_context = text[start_idx:start_idx+50]
                deadline_info = deadline_context.strip()
                break
        
        # Special handling for specific dates like "july 2nd"
        import re
        date_match = re.search(r'(january|february|march|april|may|june|july|august|september|october|november|december)\s+\d{1,2}(?:st|nd|rd|th)?', text, re.IGNORECASE)
        if date_match:
            deadline_info = f"Due: {date_match.group(0)}"
        
        # Smart reply detection in fallback
        should_reply = True
        sender_email_lower = email.sender_email.lower()
        if any(pattern in sender_email_lower for pattern in ['noreply', 'marketing', 'automated', 'notification']):
            should_reply = False
        if any(word in text for word in ['unsubscribe', 'promotional', 'marketing']):
            should_reply = False

        # Generate enhanced task breakdown with extracted details
        task_breakdown = []
        
        # Extract links using enhanced regex patterns
        import re
        # Enhanced URL pattern to catch more link formats
        url_patterns = [
            r'http[s]?://(?:[a-zA-Z]|[0-9]|[$-_@.&+]|[!*\\(\\),]|(?:%[0-9a-fA-F][0-9a-fA-F]))+',
            r'www\.(?:[a-zA-Z]|[0-9]|[$-_@.&+]|[!*\\(\\),]|(?:%[0-9a-fA-F][0-9a-fA-F]))+',
            r'(?:[a-zA-Z0-9-]+\.)+[a-zA-Z]{2,}(?:/[^\s]*)?'  # domain.com/path format
        ]
        urls = []
        for pattern in url_patterns:
            urls.extend(re.findall(pattern, email.body))
        # Remove duplicates and clean URLs
        urls = list(set([url.strip('.,;') for url in urls if len(url) > 10]))
        
        # Extract email addresses
        emails_mentioned = re.findall(r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b', email.body)
        
        # Extract phone numbers (basic pattern)
        phones = re.findall(r'\b(?:\+?1[-.\s]?)?\(?[0-9]{3}\)?[-.\s]?[0-9]{3}[-.\s]?[0-9]{4}\b', email.body)
        
        # Extract times (basic pattern)
        times = re.findall(r'\b(?:[01]?[0-9]|2[0-3]):[0-5][0-9]\s*(?:AM|PM|am|pm)?\b', email.body)
        
        # Deadline-related tasks with extracted details
        if any(word in text for word in ['deadline', 'due', 'submit', 'delivery']):
            task_breakdown.append("⏰ Review deadline requirements carefully")
            if deadline_info:
                task_breakdown.append(f"📅 Key deadline: {deadline_info}")
            if urls:
                task_breakdown.append(f"🔗 Access submission portal: {urls[0]}")
            task_breakdown.append("📋 Prepare materials/documents for submission")
            task_breakdown.append("✅ Submit before deadline")
            
        # Meeting-related tasks with extracted details
        elif any(word in text for word in ['meeting', 'schedule', 'calendar', 'appointment']):
            task_breakdown.append("📅 Check calendar availability")
            if times:
                task_breakdown.append(f"🕐 Meeting time: {times[0]}")
            if urls:
                task_breakdown.append(f"🔗 Meeting link: {urls[0]}")
            
            # Extract date information for calendar event creation
            date_patterns = [
                r'(?:january|february|march|april|may|june|july|august|september|october|november|december)\s+\d{1,2}(?:st|nd|rd|th)?(?:,?\s+\d{4})?',
                r'\d{1,2}[-/]\d{1,2}[-/]\d{2,4}',
                r'(?:monday|tuesday|wednesday|thursday|friday|saturday|sunday)'
            ]
            dates_found = []
            for pattern in date_patterns:
                dates_found.extend(re.findall(pattern, email.body, re.IGNORECASE))
            
            if dates_found:
                task_breakdown.append(f"📆 Meeting date: {dates_found[0]}")
                task_breakdown.append("🗓️ Add to calendar automatically")
            
            if emails_mentioned and emails_mentioned[0] != email.sender_email:
                task_breakdown.append(f"✉️ Confirm attendance to: {emails_mentioned[0]}")
            else:
                task_breakdown.append(f"✉️ Confirm attendance to: {email.sender_email}")
            task_breakdown.append("📝 Prepare agenda items if needed")
            
        # Question/request tasks with extracted details
        elif '?' in text or 'please' in text:
            task_breakdown.append("📖 Read email carefully")
            if urls:
                task_breakdown.append(f"🔗 Review linked resources: {urls[0]}")
            task_breakdown.append("💭 Prepare thoughtful response")
            if should_reply:
                task_breakdown.append(f"📤 Send reply to: {email.sender_email}")
                
        # Portal/survey tasks
        elif any(word in text for word in ['survey', 'portal', 'form', 'questionnaire']):
            if urls:
                task_breakdown.append(f"🔗 Access survey/portal: {urls[0]}")
            task_breakdown.append("📝 Complete required sections")
            if deadline_info:
                task_breakdown.append(f"⏰ Submit by: {deadline_info}")
            task_breakdown.append("💾 Save confirmation receipt")
                
        # Information/announcement tasks with extracted details
        else:
            task_breakdown.append("📖 Review email content")
            if urls:
                task_breakdown.append(f"🔗 Review linked documents: {urls[0]}")
            if phones:
                task_breakdown.append(f"📞 Contact info available: {phones[0]}")
            task_breakdown.append("📝 Take note of important information")
            if 'action' in text or 'required' in text:
                task_breakdown.append("⚡ Determine if action is needed")

        return LLMAnalysisResult(
            priority_score=priority_score,
            urgency_level=urgency_level,
            email_type='request' if '?' in text or 'please' in text else 'information',
            action_required='reply' if priority_score > 60 and should_reply else 'review',
            should_reply=should_reply,
            key_points=['Professional email requiring attention'],
            suggested_response_tone='professional',
            deadline_info=deadline_info,
            sender_relationship='professional',
            business_context='Business communication requiring timely response',
            confidence=0.6,
            task_breakdown=task_breakdown
        )

class EmailTemplateManager:
    """Manages email templates and quick responses"""
    
    def __init__(self):
        self.templates = self._load_default_templates()
    
    def _load_default_templates(self) -> List[EmailTemplate]:
        """Load default email templates"""
        return [
            EmailTemplate(
                name="Meeting Accept",
                category="meetings",
                subject_template="Re: {original_subject}",
                body_template="Hi {sender_name},\n\nI can attend the meeting on {meeting_date} at {meeting_time}. I'll add it to my calendar.\n\nLooking forward to it.",
                tone="professional",
                use_cases=["meeting invitation", "calendar request"],
                variables=["sender_name", "meeting_date", "meeting_time"]
            ),
            EmailTemplate(
                name="Meeting Decline",
                category="meetings", 
                subject_template="Re: {original_subject}",
                body_template="Hi {sender_name},\n\nUnfortunately, I have a conflict and won't be able to attend the meeting on {meeting_date}. Could we reschedule to another time?\n\nPlease let me know what works for you.",
                tone="professional",
                use_cases=["meeting invitation", "calendar conflict"],
                variables=["sender_name", "meeting_date"]
            ),
            EmailTemplate(
                name="Quick Acknowledgment",
                category="general",
                subject_template="Re: {original_subject}",
                body_template="Hi {sender_name},\n\nThank you for your email. I've received it and will review the details. I'll get back to you by {response_deadline}.",
                tone="professional",
                use_cases=["acknowledgment", "buying time"],
                variables=["sender_name", "response_deadline"]
            ),
            EmailTemplate(
                name="Request More Info",
                category="general",
                subject_template="Re: {original_subject} - Need Additional Information",
                body_template="Hi {sender_name},\n\nThank you for reaching out. To better assist you, could you please provide:\n\n• {info_needed_1}\n• {info_needed_2}\n\nOnce I have this information, I'll be able to help you more effectively.",
                tone="professional",
                use_cases=["clarification", "information request"],
                variables=["sender_name", "info_needed_1", "info_needed_2"]
            ),
            EmailTemplate(
                name="Out of Office Reply",
                category="automated",
                subject_template="Out of Office: Re: {original_subject}",
                body_template="Hi {sender_name},\n\nI'm currently out of the office and will return on {return_date}. I'll respond to your email when I'm back.\n\nFor urgent matters, please contact {backup_contact}.",
                tone="professional",
                use_cases=["vacation", "out of office"],
                variables=["sender_name", "return_date", "backup_contact"]
            ),
            EmailTemplate(
                name="Task Completion",
                category="updates",
                subject_template="Completed: {task_name}",
                body_template="Hi {sender_name},\n\nI've completed {task_name} as requested. {completion_details}\n\nPlease let me know if you need anything else.",
                tone="professional",
                use_cases=["task update", "completion notification"],
                variables=["sender_name", "task_name", "completion_details"]
            )
        ]
    
    def get_template_suggestions(self, email: OutlookEmailData, analysis: LLMAnalysisResult) -> List[EmailTemplate]:
        """Suggest appropriate templates based on email analysis"""
        suggestions = []
        
        email_text = (email.subject + " " + email.body).lower()
        
        # Meeting-related templates
        if analysis.email_type == "meeting" or "meeting" in email_text or "calendar" in email_text:
            suggestions.extend([t for t in self.templates if t.category == "meetings"])
        
        # Quick acknowledgment for urgent emails
        if analysis.urgency_level in ["urgent", "critical"]:
            suggestions.extend([t for t in self.templates if t.name == "Quick Acknowledgment"])
        
        # Request for more info if email is unclear
        if analysis.confidence < 0.7:
            suggestions.extend([t for t in self.templates if t.name == "Request More Info"])
        
        # General templates
        suggestions.extend([t for t in self.templates if t.category == "general"][:2])
        
        return suggestions[:3]  # Return top 3 suggestions
    
    def generate_from_template(self, template: EmailTemplate, variables: Dict[str, str]) -> str:
        """Generate email content from template with variable substitution"""
        subject = template.subject_template
        body = template.body_template
        
        for var, value in variables.items():
            subject = subject.replace(f"{{{var}}}", value)
            body = body.replace(f"{{{var}}}", value)
        
        return f"Subject: {subject}\\n\\nBody:\\n{body}"

class FollowUpTracker:
    """Tracks emails that need follow-up"""
    
    def __init__(self):
        self.follow_ups: List[FollowUpItem] = []
    
    def add_follow_up(self, email: OutlookEmailData, analysis: LLMAnalysisResult, days_ahead: int = 3) -> FollowUpItem:
        """Add email to follow-up list"""
        follow_up_date = datetime.now() + timedelta(days=days_ahead)
        
        # Determine reason for follow-up
        reason = "General follow-up"
        if hasattr(analysis, 'deadline_info') and analysis.deadline_info:
            reason = f"Deadline approaching: {analysis.deadline_info}"
        elif analysis.action_required == "reply":
            reason = "Response expected"
        elif analysis.email_type == "meeting":
            reason = "Meeting follow-up"
        
        follow_up = FollowUpItem(
            email_id=email.id,
            subject=email.subject,
            sender=email.sender,
            follow_up_date=follow_up_date,
            reason=reason,
            priority=analysis.urgency_level
        )
        
        self.follow_ups.append(follow_up)
        return follow_up
    
    def get_due_follow_ups(self) -> List[FollowUpItem]:
        """Get follow-ups that are due"""
        now = datetime.now()
        due_items = []
        
        for item in self.follow_ups:
            if item.status == "pending" and item.follow_up_date <= now:
                item.status = "overdue" if item.follow_up_date < now else "due"
                due_items.append(item)
        
        return sorted(due_items, key=lambda x: x.follow_up_date)
    
    def mark_completed(self, email_id: str, notes: str = ""):
        """Mark follow-up as completed"""
        for item in self.follow_ups:
            if item.email_id == email_id:
                item.status = "completed"
                item.notes = notes
                break

class EmailSummarizer:
    """Generates summaries of long emails"""
    
    def __init__(self, llm_service: UnifiedLLMService):
        self.llm = llm_service
    
    def should_summarize(self, email: OutlookEmailData) -> bool:
        """Determine if email needs summarization"""
        word_count = len(email.body.split())
        return word_count > 200  # Summarize emails longer than 200 words
    
    def generate_summary(self, email: OutlookEmailData) -> EmailSummary:
        """Generate comprehensive email summary"""
        
        summary_prompt = f"""
Analyze this email and provide a comprehensive summary:

FROM: {email.sender} <{email.sender_email}>
SUBJECT: {email.subject}
CONTENT: {email.body}

Please provide:
1. Key points (main topics discussed)
2. Action items (specific tasks or requests)
3. People mentioned (names and roles if identifiable) 
4. Important dates mentioned
5. Estimated reading time
6. Concise summary paragraph

Format as JSON:
{{
    "key_points": ["point1", "point2"],
    "action_items": ["action1", "action2"],
    "mentioned_people": ["person1", "person2"], 
    "mentioned_dates": ["date1", "date2"],
    "estimated_read_time": "2 minutes",
    "summary_text": "Brief summary paragraph"
}}
"""
        
        try:
            response = self.llm.generate_response(summary_prompt, max_tokens=400, temperature=0.3)
            
            # Extract JSON from response
            json_start = response.find('{')
            json_end = response.rfind('}') + 1
            
            if json_start != -1 and json_end > json_start:
                json_str = response[json_start:json_end]
                summary_data = json.loads(json_str)
                
                return EmailSummary(
                    key_points=summary_data.get('key_points', []),
                    action_items=summary_data.get('action_items', []),
                    mentioned_people=summary_data.get('mentioned_people', []),
                    mentioned_dates=summary_data.get('mentioned_dates', []),
                    estimated_read_time=summary_data.get('estimated_read_time', 'Unknown'),
                    summary_text=summary_data.get('summary_text', 'Summary not available')
                )
            else:
                return self._fallback_summary(email)
                
        except Exception as e:
            print(f"❌ Summary generation error: {e}")
            return self._fallback_summary(email)
    
    def _fallback_summary(self, email: OutlookEmailData) -> EmailSummary:
        """Generate basic summary when LLM fails"""
        word_count = len(email.body.split())
        read_time = f"{max(1, word_count // 200)} minute{'s' if word_count > 200 else ''}"
        
        return EmailSummary(
            key_points=[f"Email from {email.sender}", f"Subject: {email.subject}"],
            action_items=["Review email content"],
            mentioned_people=[email.sender],
            mentioned_dates=[],
            estimated_read_time=read_time,
            summary_text=f"Email from {email.sender} regarding {email.subject}. Contains {word_count} words."
        )

class EmailSecurityAnalyzer:
    """Analyzes emails for security threats"""
    
    def __init__(self, llm_service: UnifiedLLMService):
        self.llm = llm_service
    
    def analyze_security(self, email: OutlookEmailData) -> SecurityAnalysis:
        """Perform comprehensive security analysis"""
        
        # Quick rule-based checks
        suspicious_indicators = []
        risk_level = "low"
        
        # Check sender domain
        sender_domain = email.sender_email.split('@')[-1].lower()
        
        # Trusted institutional domains (reduce false positives)
        trusted_domains = [
            'mbzuai.ac.ae', 'sharepointonline.com', 'outlook.com', 'microsoft.com',
            'edu', 'ac.ae', 'gov.ae', 'ac.uk', 'edu.gov', 'microsoft365.com'
        ]
        
        is_trusted_domain = any(trusted in sender_domain for trusted in trusted_domains)
        
        if sender_domain in ['suspicious-domain.com', 'phishing-site.net']:  # Add known bad domains
            suspicious_indicators.append("Sender from suspicious domain")
            risk_level = "high"
        
        # Check for phishing indicators
        email_text = (email.subject + " " + email.body).lower()
        phishing_keywords = [
            'urgent action required', 'verify account', 'click here immediately',
            'suspended account', 'confirm identity', 'update payment info',
            'limited time offer', 'act now', 'winner selected', 'claim prize',
            'tax refund', 'inheritance', 'lottery', 'prince', 'million dollars'
        ]
        
        for keyword in phishing_keywords:
            if keyword in email_text:
                # Reduce risk level for trusted domains
                if is_trusted_domain:
                    if keyword in ['urgent action required', 'verify account']:
                        # These might be legitimate from trusted institutional sources
                        suspicious_indicators.append(f"Contains keyword '{keyword}' (from trusted domain)")
                        risk_level = "low" if risk_level == "low" else "medium"
                    else:
                        suspicious_indicators.append(f"Contains phishing keyword: '{keyword}'")
                        risk_level = "medium" if risk_level == "low" else "high"
                else:
                    suspicious_indicators.append(f"Contains phishing keyword: '{keyword}'")
                    risk_level = "medium" if risk_level == "low" else "high"
        
        # Check for suspicious links (but be more lenient with trusted domains)
        if 'http://' in email.body or any(domain in email.body for domain in ['.tk', '.ml', '.ga']):
            if is_trusted_domain:
                # SharePoint and institutional emails often have http links
                suspicious_indicators.append("Contains http links (from trusted domain)")
                # Don't increase risk level for trusted domains with http links
            else:
                suspicious_indicators.append("Contains potentially unsafe links")
                risk_level = "medium" if risk_level == "low" else risk_level
        
        # LLM-based advanced analysis
        security_prompt = f"""
Analyze this email for security threats and phishing attempts:

FROM: {email.sender} <{email.sender_email}>
SUBJECT: {email.subject}
CONTENT: {email.body[:1000]}

CONTEXT:
- Sender domain: {sender_domain}
- Trusted institutional domain: {is_trusted_domain}

IMPORTANT: If this is from a trusted institutional domain ({sender_domain}), be more lenient in threat assessment.

Look for:
1. Phishing attempts (fake login pages, credential theft)
2. Social engineering tactics 
3. Suspicious urgency or pressure tactics
4. Impersonation attempts
5. Malicious links or attachments
6. Spelling/grammar errors suggesting fake emails
7. Mismatched sender information

Rate risk level: low/medium/high/critical
- For trusted institutional emails, default to "low" unless clear threats exist
- Only escalate to "medium" if genuinely suspicious content is present

Format as JSON:
{{
    "additional_threats": ["threat1", "threat2"],
    "llm_risk_assessment": "low",
    "confidence": 0.85,
    "detailed_analysis": "Explanation of findings"
}}
"""
        
        try:
            response = self.llm.generate_response(security_prompt, max_tokens=300, temperature=0.1)
            
            # Extract JSON from response
            json_start = response.find('{')
            json_end = response.rfind('}') + 1
            
            if json_start != -1 and json_end > json_start:
                json_str = response[json_start:json_end]
                llm_data = json.loads(json_str)
                
                # Combine rule-based and LLM findings
                all_threats = suspicious_indicators + llm_data.get('additional_threats', [])
                llm_risk = llm_data.get('llm_risk_assessment', 'low')
                
                # Take higher risk level
                final_risk = max([risk_level, llm_risk], key=lambda x: ['low', 'medium', 'high', 'critical'].index(x))
                
                return SecurityAnalysis(
                    is_suspicious=len(all_threats) > 0 or final_risk != 'low',
                    risk_level=final_risk,
                    threats_detected=all_threats,
                    suspicious_indicators=suspicious_indicators,
                    recommendations=self._get_security_recommendations(final_risk, all_threats),
                    confidence=llm_data.get('confidence', 0.7)
                )
        
        except Exception as e:
            print(f"❌ Security analysis error: {e}")
        
        # Fallback to rule-based analysis only
        return SecurityAnalysis(
            is_suspicious=len(suspicious_indicators) > 0,
            risk_level=risk_level,
            threats_detected=suspicious_indicators,
            suspicious_indicators=suspicious_indicators,
            recommendations=self._get_security_recommendations(risk_level, suspicious_indicators),
            confidence=0.6
        )
    
    def _get_security_recommendations(self, risk_level: str, threats: List[str]) -> List[str]:
        """Generate security recommendations based on risk level"""
        recommendations = []
        
        if risk_level == "critical":
            recommendations.extend([
                "🚨 DO NOT CLICK any links or download attachments",
                "🚨 DO NOT provide any personal information",
                "Report this email to IT security team immediately",
                "Delete this email after reporting"
            ])
        elif risk_level == "high":
            recommendations.extend([
                "⚠️ Be extremely cautious with this email",
                "Verify sender identity through alternative means",
                "Do not click links or download attachments",
                "Consider reporting to security team"
            ])
        elif risk_level == "medium":
            recommendations.extend([
                "⚠️ Exercise caution with this email",
                "Verify any requests independently",
                "Be suspicious of urgent demands"
            ])
        else:
            recommendations.append("✅ Email appears safe, but always stay vigilant")
        
        return recommendations

@dataclass
class EmailCategory:
    """Email category classification result"""
    category: str  # project, department, personal, vendor, etc.
    subcategory: str
    confidence: float
    tags: List[str]

class SmartCategorizer:
    """Categorizes emails by project, department, and other criteria"""
    
    def __init__(self, llm_service: UnifiedLLMService):
        self.llm = llm_service
        self.categories = {
            'academic': ['probation', 'evaluation', 'grade', 'course', 'thesis', 'research'],
            'administrative': ['evaluation', 'approval', 'form', 'procedure', 'policy'],
            'event': ['invitation', 'seminar', 'conference', 'workshop', 'lecture', 'keynote'],
            'meeting': ['meeting', 'calendar', 'appointment', 'schedule', 'agenda'],
            'hr': ['hr', 'human resources', 'benefits', 'payroll', 'vacation', 'leave'],
            'finance': ['invoice', 'payment', 'budget', 'expense', 'purchase', 'billing'],
            'technical': ['bug', 'development', 'code', 'deployment', 'technical', 'api'],
            'vendor': ['vendor', 'supplier', 'contract', 'procurement', 'service provider'],
            'marketing': ['marketing', 'campaign', 'promotion', 'newsletter', 'announcement'],
            'general': ['update', 'news', 'information', 'notice', 'newsletter']
        }
    
    def categorize_email(self, email: OutlookEmailData, analysis: LLMAnalysisResult) -> EmailCategory:
        """Categorize email using rule-based and LLM analysis"""
        
        # Rule-based categorization
        email_text = (email.subject + " " + email.body).lower()
        rule_category = self._rule_based_categorization(email_text)
        
        # LLM-based advanced categorization
        category_prompt = f"""
Categorize this email into appropriate business categories based ONLY on the actual content:

FROM: {email.sender} <{email.sender_email}>
SUBJECT: {email.subject}
CONTENT: {email.body[:500]}

ANALYSIS CONTEXT:
- Email Type: {analysis.email_type}
- Sender Relationship: {analysis.sender_relationship}

IMPORTANT GUIDELINES:
- Only assign specific project names if explicitly mentioned in the email content
- If no specific project is mentioned, use general categories like "general", "administrative", "announcement"
- Do NOT invent or assume project names that aren't in the email
- For academic/university emails, use categories like "academic", "administrative", "event", "evaluation"

Classify into:
1. Primary Category: academic/administrative/meeting/hr/finance/technical/vendor/marketing/customer/personal/general/event
2. Subcategory: specific project name ONLY if explicitly mentioned, otherwise use "general" or leave empty
3. Tags: relevant keywords from the actual email content

Examples:
- Probation evaluation email → category: "administrative", subcategory: "evaluation"  
- IEC event invitation → category: "event", subcategory: "general"
- Newsletter → category: "general", subcategory: "newsletter"

Format as JSON:
{{
    "category": "administrative",
    "subcategory": "evaluation",
    "confidence": 0.85,
    "tags": ["probation", "evaluation", "academic"]
}}
"""
        
        try:
            response = self.llm.generate_response(category_prompt, max_tokens=200, temperature=0.3)
            
            # Extract JSON from response
            json_start = response.find('{')
            json_end = response.rfind('}') + 1
            
            if json_start != -1 and json_end > json_start:
                json_str = response[json_start:json_end]
                category_data = json.loads(json_str)
                
                return EmailCategory(
                    category=category_data.get('category', rule_category),
                    subcategory=category_data.get('subcategory', ''),
                    confidence=float(category_data.get('confidence', 0.7)),
                    tags=category_data.get('tags', [])
                )
        
        except Exception as e:
            print(f"❌ Categorization error: {e}")
        
        # Fallback to rule-based only
        return EmailCategory(
            category=rule_category,
            subcategory='',
            confidence=0.6,
            tags=[]
        )
    
    def _rule_based_categorization(self, email_text: str) -> str:
        """Simple rule-based categorization"""
        category_scores = {}
        
        for category, keywords in self.categories.items():
            score = sum(1 for keyword in keywords if keyword in email_text)
            if score > 0:
                category_scores[category] = score
        
        if category_scores:
            return max(category_scores, key=category_scores.get)
        
        return 'general'

class LLMResponseDrafter:
    """Advanced response drafter using LLM"""
    
    def __init__(self, llm_service: UnifiedLLMService):
        self.llm = llm_service
    
    def generate_draft(self, email: OutlookEmailData, analysis: LLMAnalysisResult, 
                      user_writing_style: str = "", user_name: str = "", user_email: str = "") -> LLMDraftResult:
        """Generate personalized draft response using LLM"""
        
        from datetime import datetime
        
        draft_prompt = f"""
CONTEXT: {user_name} received an email and needs to write a response.

EMAIL THAT {user_name} RECEIVED:
From: {email.sender} <{email.sender_email}>
To: {user_name}
Subject: {email.subject}
Content: {email.body[:800]}

YOUR TASK: Write {user_name}'s reply email TO {email.sender}.
IMPORTANT: You are writing AS {user_name}, responding TO {email.sender}.

ANALYSIS CONTEXT:
- Email Type: {analysis.email_type}
- Action Required: {analysis.action_required}
- Key Points: {', '.join(analysis.key_points)}
- Suggested Tone: {analysis.suggested_response_tone}
- Deadline: {analysis.deadline_info or 'None mentioned'}
- Sender Relationship: {analysis.sender_relationship}
- Business Context: {analysis.business_context}

{user_name}'S WRITING STYLE:
{user_writing_style or "Professional, friendly, concise. Uses 'Hi [Name]' greetings and 'Best regards' closings. Direct but polite communication style."}

CURRENT DATE: {datetime.now().strftime('%B %d, %Y (%A)')}

RESPONSE REQUIREMENTS FOR {user_name}:
1. Address the sender by first name (extract from sender name)
2. Acknowledge the specific content/request in the original email
3. Provide a helpful, appropriate response from {user_name}'s perspective
4. Match the suggested tone and relationship level
5. Include specific actions or next steps if needed
6. Use natural, human-like language
7. Keep professional but personalized
8. Address any deadlines or urgency
9. CRITICAL: Do NOT include ANY closing like "Best regards", "Sincerely", etc.
10. Do NOT include ANY signature, contact details, or organizational information
11. End the response with the actual email content only

DEADLINE HANDLING:
- CRITICAL: Always check if any mentioned deadline has already passed
- If a deadline has passed, acknowledge this and use phrases like "as soon as possible", "immediately", "urgently"
- NEVER promise to complete something by a date that has already passed
- If today is September 1st and email mentions "by July 15th", acknowledge the delay and commit to urgent completion

SPECIAL HANDLING:
- For IT Helpdesk forms: Check if this is a service request form. If so, acknowledge receipt and provide timeline for action.
- For marketing emails: Politely acknowledge if interested, or briefly decline if not relevant.
- For meeting requests: Accept/decline with calendar consideration.
- For questions: Provide helpful answers or timeline for detailed response.

You MUST respond with ONLY valid JSON. Generate the email content that will go after "Hi [SenderName]," 

SENDER INFO: The email is from {email.sender} <{email.sender_email}>
EXTRACT FIRST NAME: {email.sender.split()[0] if email.sender else 'there'}

EXAMPLE: If the sender name is "John Smith", your response should start with "Hi John,"

{{
    "subject": "Re: [original subject]", 
    "body": "Hi {email.sender.split()[0] if email.sender else 'there'},\
\
Thank you for the reminder about the July 15th deadline. I apologize for the delay - I will update the task progress on Notion as soon as possible and prioritize this urgently.",
    "tone": "professional",
    "confidence": 0.9,
    "reasoning": "Acknowledging passed deadline with appropriate urgency and commitment",
    "alternative_versions": []
}}

CRITICAL REQUIREMENTS:
- Use the actual sender's first name, NOT placeholder text
- If sender is "John Smith", write "Hi John," NOT "Hi [SenderFirstName],"
- Write {user_name}'s response TO this specific sender
- Address their specific request/content
- Be helpful and professional
- 2-3 sentences minimum
- NO signatures or closings (Outlook will add these automatically)
"""

        try:
            # Advanced prompting with prefiling technique (research-based)
            response = self.llm.generate_response(draft_prompt, max_tokens=1000, temperature=0.3)
            
            # Check if response is empty or too short
            if not response or len(response.strip()) < 20:
                print(f"⚠️ LLM response too short or empty, using fallback")
                return self._fallback_draft(email, analysis, user_name, user_email)
            
            
            # Enhanced JSON extraction with multiple fallback strategies
            json_str = self._extract_json_robust(response)
            
            if json_str:
                
                # Apply advanced JSON fixing from research
                json_str = self._fix_json_advanced(json_str)
                
                # Validate JSON before parsing
                if self._validate_json_structure(json_str):
                    draft_data = json.loads(json_str)
                else:
                    print("❌ JSON structure validation failed")
                    return self._fallback_draft(email, analysis, user_name, user_email)
                
                # Calculate response quality score
                response_body = draft_data.get('body', '')
                quality_score = self._calculate_response_quality(response_body, email, analysis)
                
                return LLMDraftResult(
                    subject=draft_data.get('subject', f"Re: {email.subject}"),
                    body=response_body,
                    tone=draft_data.get('tone', 'professional'),
                    confidence=min(float(draft_data.get('confidence', 0.7)), quality_score),
                    reasoning=draft_data.get('reasoning', ''),
                    alternative_versions=draft_data.get('alternative_versions', [])
                )
            else:
                print(f"❌ Could not extract JSON from LLM response")
                return self._fallback_draft(email, analysis, user_name, user_email)
                
        except json.JSONDecodeError as e:
            print(f"❌ JSON decode error in draft generation: {e}")
            return self._fallback_draft(email, analysis, user_name, user_email)
        except Exception as e:
            print(f"❌ LLM draft generation error: {e}")
            return self._fallback_draft(email, analysis, user_name, user_email)
    
    def _calculate_response_quality(self, response_body: str, email: OutlookEmailData, analysis: LLMAnalysisResult) -> float:
        """Calculate quality score for the generated response (open-source technique)"""
        
        quality_score = 1.0
        
        # Check if response addresses the original email content
        original_keywords = set(email.subject.lower().split() + email.body[:500].lower().split())
        response_keywords = set(response_body.lower().split())
        keyword_overlap = len(original_keywords & response_keywords) / max(len(original_keywords), 1)
        
        if keyword_overlap < 0.1:
            quality_score -= 0.3  # Poor content relevance
        elif keyword_overlap > 0.3:
            quality_score += 0.1  # Good content relevance
        
        # Check response length appropriateness
        response_words = len(response_body.split())
        if response_words < 10:
            quality_score -= 0.2  # Too short
        elif response_words > 200:
            quality_score -= 0.1  # Too long
        elif 20 <= response_words <= 100:
            quality_score += 0.1  # Good length
        
        # Check if it addresses the action required
        if analysis.action_required == "reply" and len(response_body) < 20:
            quality_score -= 0.2
        
        # Check for professional tone indicators
        professional_indicators = ['thank you', 'please', 'appreciate', 'understand', 'assist']
        if any(indicator in response_body.lower() for indicator in professional_indicators):
            quality_score += 0.1
        
        # Penalty for remaining signatures (shouldn't happen but check)
        if any(sig in response_body.lower() for sig in ['best regards', 'sincerely', 'client ai engineer']):
            quality_score -= 0.3
        
        return max(0.3, min(1.0, quality_score))
    
    def _extract_json_robust(self, response: str) -> str:
        """Advanced JSON extraction with multiple strategies (research-based)"""
        import re
        
        # Strategy 1: Look for JSON between braces with proper brace matching
        json_start = response.find('{')
        if json_start != -1:
            # Find matching closing brace by counting
            brace_count = 0
            json_end = json_start
            
            for i, char in enumerate(response[json_start:], json_start):
                if char == '{':
                    brace_count += 1
                elif char == '}':
                    brace_count -= 1
                    if brace_count == 0:
                        json_end = i + 1
                        break
            
            if json_end > json_start:
                potential_json = response[json_start:json_end]
                # Basic validation - must have required structure
                if '"subject"' in potential_json and '"body"' in potential_json:
                    return potential_json
        
        # Strategy 2: Look for code blocks (```json...```)
        json_block = re.search(r'```json\s*(.*?)\s*```', response, re.DOTALL)
        if json_block:
            return json_block.group(1).strip()
        
        # Strategy 3: Look for any code block
        code_block = re.search(r'```\s*(.*?)\s*```', response, re.DOTALL)
        if code_block:
            content = code_block.group(1).strip()
            if content.startswith('{') and content.endswith('}'):
                return content
        
        # Strategy 4: Try to reconstruct from partial JSON
        lines = response.split('\\n')
        json_lines = []
        in_json = False
        
        for line in lines:
            if '{' in line and not in_json:
                in_json = True
                json_lines.append(line[line.find('{'):])
            elif in_json:
                json_lines.append(line)
                if '}' in line:
                    break
        
        if json_lines:
            return '\\n'.join(json_lines)
        
        # Strategy 5: Simple fallback - try to build minimal JSON from response content
        if "Re:" in response and any(word in response.lower() for word in ["thank", "hi", "hello"]):
            # Extract subject
            subject_match = re.search(r'"subject":\s*"([^"]*)"', response)
            subject = subject_match.group(1) if subject_match else f"Re: {email.subject if hasattr(email, 'subject') else 'Your Email'}"
            
            # Find greeting and content
            greeting_match = re.search(r'(Hi\s+\w+|Hello\s+\w+)', response, re.IGNORECASE)
            greeting = greeting_match.group(1) if greeting_match else "Hi there"
            
            # Create minimal valid JSON
            return f'''{{
    "subject": "{subject}",
    "body": "{greeting},\
\
Thank you for your email. I will review this and respond accordingly.\
\
Best regards,\
Avani",
    "tone": "professional",
    "confidence": 0.7,
    "reasoning": "Reconstructed from partial response",
    "alternative_versions": []
}}'''
        
        return ""
    
    def _fix_json_advanced(self, json_str: str) -> str:
        """Advanced JSON fixing based on 2024 research best practices"""
        import re
        
        
        # Remove any trailing commas before closing braces/brackets
        json_str = re.sub(r',\s*}', '}', json_str)
        json_str = re.sub(r',\s*]', ']', json_str)
        
        # Fix unescaped quotes in string values
        json_str = re.sub(r"'([^']*)':", r'"\1":', json_str)  # Fix keys
        json_str = re.sub(r":\s*'([^']*)'", r': "\1"', json_str)  # Fix string values
        
        # Handle newlines in JSON string values more carefully
        # This is a critical fix - properly escape real newlines within strings
        lines = json_str.split('\\n')
        fixed_lines = []
        in_string = False
        
        for line in lines:
            # Count unescaped quotes to track if we're inside a string
            quote_count = 0
            i = 0
            while i < len(line):
                if line[i] == '"' and (i == 0 or line[i-1] != '\\'):
                    quote_count += 1
                i += 1
            
            if quote_count % 2 == 1:  # Odd number of quotes - we're entering/exiting a string
                in_string = not in_string
            
            if in_string and len(fixed_lines) > 0:
                # We're inside a JSON string that spans lines - escape the newline
                fixed_lines[-1] += '\
' + line
            else:
                fixed_lines.append(line)
        
        json_str = '\\n'.join(fixed_lines)
        
        # Fix incomplete strings (research finding: common truncation issue)
        if json_str.count('"') % 2 != 0:
            # Odd number of quotes - likely truncated
            json_str += '"'
        
        # Ensure closing brace if missing (common truncation)
        open_braces = json_str.count('{')
        close_braces = json_str.count('}')
        
        if open_braces > close_braces:
            json_str += '}' * (open_braces - close_braces)
        
        # Fix double-escaped backslashes (common LLM issue)
        json_str = json_str.replace('\\\
', '\
')
        json_str = json_str.replace('\\\"', '"')
        
        
        return json_str
    
    def _validate_json_structure(self, json_str: str) -> bool:
        """Validate JSON structure before parsing (prevents failures)"""
        try:
            import json
            temp_data = json.loads(json_str)
            
            # Validate critical fields exist (tone is optional)
            required_fields = ['subject', 'body']
            for field in required_fields:
                if field not in temp_data:
                    print(f"❌ Missing required field: {field}")
                    return False
            
            # Set defaults for optional fields
            if 'tone' not in temp_data:
                temp_data['tone'] = 'professional'
            if 'confidence' not in temp_data:
                temp_data['confidence'] = 0.7
            if 'reasoning' not in temp_data:
                temp_data['reasoning'] = 'Generated by LLM'
            
            # Validate data types
            if not isinstance(temp_data.get('confidence'), (int, float)):
                print(f"❌ Invalid confidence type: {type(temp_data.get('confidence'))}")
                return False
                
            return True
            
        except json.JSONDecodeError as e:
            print(f"❌ JSON validation failed: {e}")
            return False
    
    def _fallback_draft(self, email: OutlookEmailData, analysis: LLMAnalysisResult, user_name: str, user_email: str = "") -> LLMDraftResult:
        """Fallback draft if LLM fails"""
        
        sender_first_name = email.sender.split()[0] if email.sender else "there"
        first_name_only = user_name.split()[0] if user_name else "User"
        
        # Generate more specific responses based on content
        original_content = email.body.lower()
        
        if analysis.action_required == "attend":
            body = f"Hi {sender_first_name},\n\nThank you for the meeting invitation. I'll check my calendar and confirm my attendance shortly."
        elif 'deadline' in original_content:
            # Check if deadline mentions specific dates that might have passed
            from datetime import datetime
            import re
            
            # Look for date patterns that might indicate passed deadlines
            date_patterns = [
                r'(?:january|february|march|april|may|june|july|august|september|october|november|december)\s+\d{1,2}(?:st|nd|rd|th)?',
                r'\d{1,2}[-/]\d{1,2}[-/]\d{2,4}'
            ]
            
            found_past_date = False
            current_month = datetime.now().month
            
            # Simple check for obviously past dates (like July when it's September)
            if 'july' in original_content and current_month >= 9:
                found_past_date = True
            elif 'june' in original_content and current_month >= 9:
                found_past_date = True
            elif 'may' in original_content and current_month >= 9:
                found_past_date = True
            
            if found_past_date:
                body = f"Hi {sender_first_name},\n\nThank you for the reminder. I apologize for the delay - I will prioritize this and complete it as soon as possible."
            else:
                body = f"Hi {sender_first_name},\n\nThank you for reminding me about the deadline. I'm working on this and will ensure timely completion as requested."
        elif 'question' in original_content or '?' in email.body:
            body = f"Hi {sender_first_name},\n\nThank you for your question. I'll review the details and provide you with a comprehensive response shortly."
        elif analysis.action_required == "reply":
            body = f"Hi {sender_first_name},\n\nThank you for your email. I've received your message and will respond with the necessary details shortly."
        else:
            body = f"Hi {sender_first_name},\n\nThank you for your email. I'll review this and get back to you soon with the relevant information."
        
        return LLMDraftResult(
            subject=f"Re: {email.subject}",
            body=body,
            tone="professional",
            confidence=0.5,
            reasoning="Fallback template used due to LLM unavailability",
            alternative_versions=[]
        )

class LLMEnhancedEmailSystem:
    """Complete email system using LLM for analysis and drafting"""
    
    def __init__(self, model_type: str = None, model: str = None, host: str = None):
        # Initialize services
        self.llm = UnifiedLLMService(model_type=model_type, model=model, host=host)
        self.analyzer = LLMEmailAnalyzer(self.llm)
        self.drafter = LLMResponseDrafter(self.llm)
        
        # Initialize new feature managers
        self.template_manager = EmailTemplateManager()
        self.follow_up_tracker = FollowUpTracker()
        self.summarizer = EmailSummarizer(self.llm)
        self.security_analyzer = EmailSecurityAnalyzer(self.llm)
        self.categorizer = SmartCategorizer(self.llm)
        
        # Initialize enhanced context-aware components (inspired by inbox-zero)
        self.contextual_drafter = None  # Will be initialized after outlook authentication
        self.writing_analyzer = WritingStyleAnalyzer(self.llm)
        self.history_extractor = None  # Will be initialized after outlook authentication
        self.knowledge_base = KnowledgeBase()
        self.user_writing_style = None  # Will be analyzed from sent emails
        
        # Initialize new AI-powered systems (inbox-zero 2024 features)
        self.smart_categorizer = SmartEmailCategorizer(self.llm)
        self.cold_email_detector = None  # Will be initialized after outlook authentication
        self.automation_engine = None  # Will be initialized after outlook authentication
        
        # Function to get configuration from Streamlit secrets or environment
        def get_config(key, default=None):
            # Check if we're running in Streamlit environment
            is_streamlit = False
            try:
                import streamlit as st
                if hasattr(st, 'secrets'):
                    is_streamlit = True
                    if key in st.secrets:
                        value = st.secrets[key]
                        # Skip placeholder values
                        if value and not str(value).startswith("your_") and "your_" not in str(value).lower():
                            return value
            except Exception:
                pass
            
            # For local development or fallback, try environment variables
            env_value = os.getenv(key)
            if env_value and env_value != default and not env_value.startswith("your_"):
                return env_value
            
            # If running locally and no env var, try to read from secrets.toml directly
            if not is_streamlit:
                try:
                    import toml
                    secrets_file = os.path.join(os.path.dirname(__file__), '..', '.streamlit', 'secrets.toml')
                    if os.path.exists(secrets_file):
                        with open(secrets_file, 'r') as f:
                            secrets = toml.load(f)
                            if key in secrets:
                                value = secrets[key]
                                if value and not str(value).startswith("your_"):
                                    return value
                except Exception:
                    pass
            
            return default
        
        # Initialize Outlook service
        self.outlook = OutlookService(
            client_id=get_config('AZURE_CLIENT_ID'),
            client_secret=get_config('AZURE_CLIENT_SECRET'),
            tenant_id=get_config('AZURE_TENANT_ID', 'common')
        )
        
        # Initialize session state persistence
        self._initialize_session_persistence()
        
        # Stats
        self.emails_analyzed = 0
        self.drafts_created = 0
        self.llm_calls = 0
        
        # Processing mode (ultra_lite, lite, or deep)
        self.processing_mode = "deep"  # Options: "ultra_lite", "lite", "deep"
    
    def set_processing_mode(self, mode: str):
        """Set processing mode: ultra_lite, lite, or deep"""
        if mode not in ["ultra_lite", "lite", "deep"]:
            raise ValueError("Mode must be 'ultra_lite', 'lite', or 'deep'")
        
        self.processing_mode = mode
        mode_names = {
            "ultra_lite": "🏃 ULTRA-LITE MODE",
            "lite": "⚡ LITE MODE", 
            "deep": "🔬 DEEP MODE"
        }
        mode_descriptions = {
            "ultra_lite": "   🏃 Keyword analysis + LLM drafting only",
            "lite": "   ⚡ LLM analysis + drafting (fast)",
            "deep": "   🔬 Full analysis - All features enabled"
        }
        
        print(f"📊 Switched to {mode_names[mode]}")
        print(mode_descriptions[mode])
    
    def get_current_mode(self) -> str:
        """Get current processing mode display name"""
        mode_names = {
            "ultra_lite": "🏃 ULTRA-LITE MODE",
            "lite": "⚡ LITE MODE", 
            "deep": "🔬 DEEP MODE"
        }
        return mode_names[self.processing_mode]
    
    def _initialize_session_persistence(self):
        """Initialize session state persistence for email analysis results"""
        import streamlit as st
        
        # Initialize session state keys for persistence
        if 'email_analysis_results' not in st.session_state:
            st.session_state.email_analysis_results = {}
        if 'email_priorities' not in st.session_state:
            st.session_state.email_priorities = {}
        if 'email_drafts' not in st.session_state:
            st.session_state.email_drafts = {}
        if 'calendar_confirmations' not in st.session_state:
            st.session_state.calendar_confirmations = {}
    
    def _save_email_analysis(self, email: 'OutlookEmailData', analysis: LLMAnalysisResult, draft: LLMDraftResult = None):
        """Save email analysis results to session state"""
        import streamlit as st
        
        email_key = f"{email.id}_{email.sender_email}"
        
        # Save analysis results
        st.session_state.email_analysis_results[email_key] = {
            'subject': email.subject,
            'sender': email.sender,
            'sender_email': email.sender_email,
            'date': email.date,
            'priority_score': analysis.priority_score,
            'urgency_level': analysis.urgency_level,
            'email_type': analysis.email_type,
            'action_required': analysis.action_required,
            'should_reply': analysis.should_reply,
            'key_points': analysis.key_points,
            'deadline_info': analysis.deadline_info,
            'timestamp': datetime.now().isoformat()
        }
        
        # Save priority
        st.session_state.email_priorities[email_key] = analysis.priority_score
        
        # Save draft if available
        if draft:
            st.session_state.email_drafts[email_key] = {
                'subject': draft.subject,
                'body': draft.body,
                'tone': draft.tone,
                'confidence': draft.confidence,
                'reasoning': draft.reasoning,
                'timestamp': datetime.now().isoformat()
            }
    
    def _load_saved_analysis(self, email: 'OutlookEmailData') -> Optional[Dict]:
        """Load saved analysis results from session state"""
        import streamlit as st
        
        email_key = f"{email.id}_{email.sender_email}"
        return st.session_state.email_analysis_results.get(email_key)
    
    def _load_saved_draft(self, email: 'OutlookEmailData') -> Optional[Dict]:
        """Load saved draft from session state"""
        import streamlit as st
        
        email_key = f"{email.id}_{email.sender_email}"
        return st.session_state.email_drafts.get(email_key)
    
    def get_persistent_priority_summary(self) -> Dict[str, Any]:
        """Get summary of all persistent priorities for display"""
        import streamlit as st
        
        if not st.session_state.email_priorities:
            return {'total_emails': 0, 'avg_priority': 0, 'high_priority_count': 0}
        
        priorities = list(st.session_state.email_priorities.values())
        high_priority_count = sum(1 for p in priorities if p >= 70)
        
        return {
            'total_emails': len(priorities),
            'avg_priority': sum(priorities) / len(priorities),
            'high_priority_count': high_priority_count,
            'priorities_by_range': {
                'critical_90_plus': sum(1 for p in priorities if p >= 90),
                'high_70_89': sum(1 for p in priorities if 70 <= p < 90),
                'medium_50_69': sum(1 for p in priorities if 50 <= p < 70),
                'low_below_50': sum(1 for p in priorities if p < 50)
            }
        }
    
    # Legacy method for backward compatibility
    def set_lite_mode(self, enabled: bool = True):
        """Legacy method - use set_processing_mode instead"""
        self.set_processing_mode("lite" if enabled else "deep")
    
    def analyze_writing_style(self, sent_emails: List[Dict]) -> str:
        """Analyze user's writing style using LLM"""
        
        if not sent_emails:
            return "Professional, friendly, concise communication style."
        
        # Combine sent email content for analysis
        email_content = "\\n\\n---\\n\\n".join([
            f"Subject: {email.get('subject', '')}\nBody: {email.get('body', '')[:500]}"
            for email in sent_emails[:5]  # Analyze up to 5 recent emails
        ])
        
        style_prompt = f"""
Analyze these email examples to determine the user's writing style and communication patterns.

EMAIL EXAMPLES:
{email_content}

Identify and describe:
1. Formality level (formal/professional/casual)
2. Greeting style (Dear/Hi/Hello)
3. Closing style (Best regards/Thanks/Best)
4. Sentence structure (short/medium/long sentences)
5. Tone (warm/neutral/direct)
6. Common phrases and expressions
7. Level of detail in responses
8. Politeness indicators

Provide a concise writing style profile that can be used to generate similar responses.
Keep it under 200 words and focus on actionable style elements.
"""

        try:
            style_analysis = self.llm.generate_response(style_prompt, max_tokens=300, temperature=0.3)
            self.llm_calls += 1
            return style_analysis
        except Exception as e:
            print(f"❌ Writing style analysis failed: {e}")
            return "Professional, friendly, concise communication style with polite greetings and closings."
    
    def initialize_enhanced_features(self):
        """Initialize enhanced features after authentication (inspired by inbox-zero)"""
        try:
            # Initialize contextual drafter after outlook authentication
            if not self.contextual_drafter:
                self.contextual_drafter = ContextualDraftGenerator(self.llm, self.outlook)
            
            # Initialize history extractor after outlook authentication
            if not self.history_extractor:
                self.history_extractor = EmailHistoryExtractor(self.outlook)
            
            # Initialize cold email detector after outlook authentication
            if not self.cold_email_detector:
                self.cold_email_detector = ColdEmailDetector(self.llm, self.outlook)
            
            # Initialize automation engine after outlook authentication
            if not self.automation_engine:
                self.automation_engine = EmailAutomationEngine(self.llm, self.outlook)
            
            # Analyze user's writing style from sent emails if not already done
            if not self.user_writing_style:
                print("📝 Analyzing your writing style from sent emails...")
                self.user_writing_style = self.writing_analyzer.analyze_sent_emails(self.outlook)
                print(f"✅ Writing style analyzed: {self.user_writing_style.tone} tone, {self.user_writing_style.greeting_style} greetings")
            
        except Exception as e:
            print(f"⚠️ Could not initialize enhanced features: {e}")
            # Use default writing style
            self.user_writing_style = self.writing_analyzer._default_writing_style()
    
    def generate_contextual_draft(self, email: OutlookEmailData, analysis: LLMAnalysisResult) -> LLMDraftResult:
        """Generate enhanced draft with multi-source context (inspired by inbox-zero)"""
        try:
            # Ensure enhanced features are initialized
            if not self.user_writing_style or not self.contextual_drafter:
                self.initialize_enhanced_features()
            
            # Check if contextual drafter is still None after initialization
            if not self.contextual_drafter:
                print("⚠️ Contextual drafter not available, falling back to basic draft")
                return self.drafter.generate_response_draft(email, analysis, "", "")
            
            # Gather context from multiple sources
            email_history = []
            conversation_thread = [email]
            
            if self.history_extractor:
                # Get conversation history with this sender
                email_history = self.history_extractor.get_conversation_history(email)
                
                # Get thread emails if available
                conversation_thread = self.history_extractor.get_thread_emails(email)
            
            # Find relevant knowledge base entries
            knowledge_entries = self.knowledge_base.find_relevant_entries(email)
            
            # Build context object
            context = EmailContext(
                current_email=email,
                email_history=email_history,
                knowledge_base_entries=knowledge_entries,
                writing_style=self.user_writing_style,
                conversation_thread=conversation_thread,
                user_preferences={}  # Could be expanded with user settings
            )
            
            print(f"🔄 Generating contextual draft with:")
            print(f"   📧 Email history: {len(email_history)} emails")
            print(f"   💬 Thread: {len(conversation_thread)} emails")
            print(f"   📚 Knowledge entries: {len(knowledge_entries)} entries")
            print(f"   ✏️ Writing style: {self.user_writing_style.tone}")
            
            # Generate draft with full context
            draft = self.contextual_drafter.generate_contextual_draft(context)
            
            self.llm_calls += 1
            return draft
            
        except Exception as e:
            print(f"❌ Error in contextual draft generation: {e}")
            # Fallback to original draft method
            return self.drafter.generate_response_draft(email, analysis, "", "")
    
    def _search_deadline_emails(self):
        """Search for emails containing deadline keywords using Graph API filtering"""
        deadline_keywords = [
            'deadline', 'due', 'approaching', 'urgent', 'critical', 
            'submit', 'delivery', 'july 2', 'timely delivery', 'modules'
        ]
        
        found_emails = []
        
        for keyword in deadline_keywords[:3]:  # Limit to first 3 keywords to avoid too many API calls
            try:
                # Use the existing Outlook service's Graph API with custom filtering
                from datetime import datetime, timedelta, timezone
                
                # Search emails from past 7 days
                after_time = (datetime.now(timezone.utc) - timedelta(hours=168)).isoformat()
                
                # Build search filter using Graph API OData query (simplified for compatibility)
                # Note: Graph API has limitations with contains() function, so we'll use simpler filtering
                filter_query = f"receivedDateTime ge {after_time}"
                
                # Use the Graph API endpoint directly - get more emails since we're filtering client-side
                select_fields = "id,subject,sender,body,bodyPreview,receivedDateTime,importance,isRead,hasAttachments,categories,toRecipients"
                endpoint = f"/me/messages?$filter={filter_query}&$select={select_fields}&$top=20&$orderby=receivedDateTime desc"
                
                try:
                    response = self.outlook._make_graph_request(endpoint)
                    
                    if 'value' in response:
                        for email_data in response['value']:
                            # Client-side filtering for keywords since Graph API contains() is unreliable
                            subject = email_data.get('subject', '').lower()
                            body_preview = email_data.get('bodyPreview', '').lower()
                            body_content = email_data.get('body', {}).get('content', '').lower()
                            
                            # Check if keyword appears in subject, body preview, or body content
                            if (keyword.lower() in subject or 
                                keyword.lower() in body_preview or 
                                keyword.lower() in body_content):
                                
                                # Convert to OutlookEmailData
                                from outlook_agent import OutlookEmailData
                                from datetime import datetime
                                
                                email = OutlookEmailData(
                                    id=email_data.get('id', ''),
                                    subject=email_data.get('subject', ''),
                                    sender=email_data.get('sender', {}).get('emailAddress', {}).get('name', 'Unknown'),
                                    sender_email=email_data.get('sender', {}).get('emailAddress', {}).get('address', ''),
                                    recipient=', '.join([r.get('emailAddress', {}).get('name', '') for r in email_data.get('toRecipients', [])]),
                                    body=email_data.get('body', {}).get('content', ''),
                                    body_preview=email_data.get('bodyPreview', ''),
                                    date=datetime.fromisoformat(email_data.get('receivedDateTime', '').replace('Z', '+00:00')),
                                    importance=email_data.get('importance', 'normal'),
                                    is_read=email_data.get('isRead', False),
                                    has_attachments=email_data.get('hasAttachments', False),
                                    categories=email_data.get('categories', [])
                                )
                                
                                found_emails.append(email)
                                print(f"      ✅ Found '{keyword}' in: {email.subject[:40]}...")
                            
                except Exception as api_error:
                    print(f"   ⚠️ Search failed for '{keyword}': {api_error}")
                    continue
                    
            except Exception as e:
                print(f"   ⚠️ Error searching for '{keyword}': {e}")
                continue
        
        return found_emails
    
    def process_emails_with_llm(self, max_emails: int = 20, priority_threshold: float = 50.0):
        """Main workflow using LLM for analysis and drafting"""
        
        print("🤖 LLM-Enhanced Email Agent")
        print("=" * 40)
        print(f"🔧 Current mode: {self.get_current_mode()}")
        
        # SAFETY CAP: Maximum absolute limit to prevent system overload
        def get_safety_config(key, default):
            try:
                import streamlit as st
                if hasattr(st, 'secrets') and key in st.secrets:
                    return int(st.secrets[key])
            except:
                pass
            return default
        
        ABSOLUTE_MAX_EMAILS = get_safety_config('ABSOLUTE_MAX_EMAILS', 500)
        if max_emails > ABSOLUTE_MAX_EMAILS:
            print(f"⚠️ SAFETY CAP: Limiting processing from {max_emails} to {ABSOLUTE_MAX_EMAILS} emails to prevent timeout")
            max_emails = ABSOLUTE_MAX_EMAILS
        
        try:
            # Get user info (authentication should already be done in initialization)
            user_info = self.outlook.get_user_info()
            current_user_email = user_info.get('email', '')
            current_user_name = user_info.get('name', '')
            
            print(f"✅ Connected as: {current_user_name} ({current_user_email})")
            
            # Step 2: Initialize enhanced features (based on mode)
            mode = self.processing_mode
            
            if mode == "deep":
                print("🚀 Initializing enhanced context-aware features...")
                self.initialize_enhanced_features()
                print(f"✅ Enhanced features ready")
            elif mode == "lite":
                print("⚡ LITE MODE: Minimal feature initialization for speed")
                # Initialize only essential features for LLM analysis
            else:  # ultra_lite
                print("🏃 ULTRA-LITE MODE: Skipping enhanced features for maximum speed")
            
            # Step 3: Fetch emails (enhanced to catch critical emails)
            unread_only = False
            try:
                # Check Streamlit secrets first
                import streamlit as st
                if hasattr(st, 'secrets') and 'PROCESS_UNREAD_ONLY' in st.secrets:
                    unread_value = st.secrets['PROCESS_UNREAD_ONLY']
                    # Handle both boolean and string values
                    if isinstance(unread_value, bool):
                        unread_only = unread_value
                    else:
                        unread_only = str(unread_value).lower() in ['true', '1', 'yes', 'on']
                else:
                    # Fallback to environment variable
                    unread_only = os.getenv('PROCESS_UNREAD_ONLY', 'false').lower() in ['true', '1', 'yes', 'on']
            except:
                # Fallback to environment variable
                unread_only = os.getenv('PROCESS_UNREAD_ONLY', 'false').lower() in ['true', '1', 'yes', 'on']
            if unread_only:
                print(f"📥 Fetching all unread emails...")
            else:
                print(f"📥 Fetching recent emails...")
            
            # Adjust email limits and time windows based on mode
            mode = self.processing_mode
            
            # Get mode-specific limits from configuration
            def get_config(key, default=None):
                # Check if we're running in Streamlit environment
                is_streamlit = False
                try:
                    import streamlit as st
                    if hasattr(st, 'secrets'):
                        is_streamlit = True
                        if key in st.secrets:
                            value = st.secrets[key]
                            # Skip placeholder values
                            if value and not str(value).startswith("your_") and "your_" not in str(value).lower():
                                return value
                except Exception:
                    pass
                
                # For local development or fallback, try environment variables
                env_value = os.getenv(key)
                if env_value and env_value != default and not env_value.startswith("your_"):
                    return env_value
                
                # If running locally and no env var, try to read from secrets.toml directly
                if not is_streamlit:
                    try:
                        import toml
                        secrets_file = os.path.join(os.path.dirname(__file__), '..', '.streamlit', 'secrets.toml')
                        if os.path.exists(secrets_file):
                            with open(secrets_file, 'r') as f:
                                secrets = toml.load(f)
                                if key in secrets:
                                    value = secrets[key]
                                    if value and not str(value).startswith("your_"):
                                        return value
                    except Exception:
                        pass
                
                return default
            
            if mode == "ultra_lite":
                # Ultra-lite: Minimal emails for instant processing
                initial_max = min(max_emails, get_config('ULTRA_LITE_INITIAL_MAX', 5))
                extended_max = min(max_emails, get_config('ULTRA_LITE_EXTENDED_MAX', 8))
                if not unread_only:
                    print(f"🏃 ULTRA-LITE MODE: Processing up to {initial_max} emails for instant results")
            elif mode == "lite":
                # Lite: Moderate email count for fast LLM processing  
                initial_max = min(max_emails, get_config('LITE_INITIAL_MAX', 10))
                extended_max = min(max_emails, get_config('LITE_EXTENDED_MAX', 15))
                if not unread_only:
                    print(f"⚡ LITE MODE: Processing up to {initial_max} emails for fast LLM analysis")
            else:  # deep mode
                # Deep mode: Full email count
                deep_initial = get_config('DEEP_INITIAL_MAX', 20)
                deep_multiplier = get_config('DEEP_EXTENDED_MULTIPLIER', 2)
                initial_max = min(max_emails, deep_initial)
                extended_max = max_emails * deep_multiplier
                if not unread_only:
                    print(f"🔬 DEEP MODE: Processing up to {initial_max} emails with full analysis")
            
            # Fetch emails based on mode
            if unread_only:
                # Get unread-specific limit (0 = no limit) and ensure it's an integer
                unread_limit_raw = get_config('UNREAD_MAX_LIMIT', 0)
                try:
                    unread_limit = int(unread_limit_raw) if unread_limit_raw else 0
                except (ValueError, TypeError):
                    unread_limit = 0
                
                if unread_limit > 0:
                    # Apply safety cap to unread limit
                    safe_unread_limit = min(unread_limit, ABSOLUTE_MAX_EMAILS)
                    if safe_unread_limit < unread_limit:
                        print(f"⚠️ SAFETY CAP: Limiting unread processing from {unread_limit} to {safe_unread_limit} emails")
                    print(f"📬 Processing up to {safe_unread_limit} unread emails...")
                    emails = self.outlook.get_recent_emails(max_results=safe_unread_limit, hours_back=72)  # hours_back ignored in unread mode
                else:
                    print(f"📬 Processing ALL unread emails (capped at {ABSOLUTE_MAX_EMAILS} for safety)...")
                    emails = self.outlook.get_recent_emails(max_results=ABSOLUTE_MAX_EMAILS, hours_back=72)  # Safety cap instead of 1000
            else:
                # First, get recent emails from past 3 days (72 hours) to catch deadline emails
                print(f"🔍 Scanning past 72 hours for critical emails...")
                emails = self.outlook.get_recent_emails(max_results=initial_max, hours_back=72)
                
                # Also try to get more emails if we're not getting enough (but respect mode limits and safety cap)
                if len(emails) < 5:  # Reduced threshold for expansion
                    safe_extended_max = min(extended_max, ABSOLUTE_MAX_EMAILS)
                    if safe_extended_max < extended_max:
                        print(f"⚠️ SAFETY CAP: Limiting extended search from {extended_max} to {safe_extended_max} emails")
                    print(f"📧 Only found {len(emails)} emails, expanding search to past 7 days (max {safe_extended_max})...")
                    emails = self.outlook.get_recent_emails(max_results=safe_extended_max, hours_back=168)  # 7 days
            
            # For now, let's focus on the improved time window which should catch your deadline email
            # The Graph API search has compatibility issues that need further investigation
            print(f"📧 Analyzing {len(emails)} emails from extended time window...")
            
            # Warn user about large processing times
            if len(emails) > 100:
                print(f"⚠️ LARGE INBOX DETECTED: {len(emails)} emails found")
                print(f"   This may take several minutes to process. Consider using Lite or Ultra-Lite mode for faster processing.")
                if len(emails) > 300:
                    print(f"   🕒 ESTIMATED TIME: {len(emails) * 2 // 60} minutes in current mode")
            
            # Debug: Check if we have the deadline email in our list
            print(f"🔍 Checking for deadline keywords in {len(emails)} emails...")
            deadline_keywords = ['deadline', 'approaching', 'july', 'submit', 'delivery', 'modules']
            for email in emails:
                email_text = (email.subject + " " + email.body).lower()
                found_keywords = [kw for kw in deadline_keywords if kw in email_text]
                if found_keywords:
                    print(f"   📌 Found deadline keywords {found_keywords} in: {email.subject[:50]}...")
            
            self.emails_analyzed = len(emails)
            
            if not emails:
                print("📪 No emails found")
                return
            
            print(f"📧 Found {len(emails)} emails")
            
            # Step 4: Analysis with 3-Tier Mode Selection
            mode = self.processing_mode
            
            if mode == "ultra_lite":
                print("🤖 Analyzing emails (🏃 ULTRA-LITE MODE - Keywords + Drafting)...")
                analyzed_emails = self._analyze_emails_ultra_lite_mode(emails, current_user_name, current_user_email)
            elif mode == "lite":
                print("🤖 Analyzing emails with LLM (⚡ LITE MODE - Fast Processing)...")
                analyzed_emails = self._analyze_emails_lite_mode(emails, current_user_name, current_user_email)
            else:  # deep mode
                print("🤖 Analyzing emails with LLM (🔬 DEEP MODE - Full Analysis)...")
                analyzed_emails = self._analyze_emails_deep_mode(emails, current_user_name, current_user_email)
            
            # Sort by LLM-determined priority
            analyzed_emails.sort(key=lambda x: x[1]['core_analysis'].priority_score, reverse=True)
            
            # Step 5: Display LLM Analysis Results with professional prioritization
            print(f"\n🎯 PROFESSIONAL EMAIL PRIORITIZATION:")
            print("-" * 60)
            
            # Separate emails by priority categories
            critical_emails = [(email, analysis) for email, analysis in analyzed_emails if analysis['core_analysis'].priority_score >= 85]
            urgent_emails = [(email, analysis) for email, analysis in analyzed_emails if 70 <= analysis['core_analysis'].priority_score < 85]
            normal_emails = [(email, analysis) for email, analysis in analyzed_emails if 50 <= analysis['core_analysis'].priority_score < 70]
            low_emails = [(email, analysis) for email, analysis in analyzed_emails if analysis['core_analysis'].priority_score < 50]
            
            # Display by priority categories
            if critical_emails:
                print("🔴 CRITICAL PRIORITY (85+):")
                for i, (email, enhanced_analysis) in enumerate(critical_emails):
                    analysis = enhanced_analysis['core_analysis']
                    security = enhanced_analysis['security_analysis']
                    summary = enhanced_analysis['email_summary']
                    templates = enhanced_analysis['template_suggestions']
                    category = enhanced_analysis['email_category']
                    
                    unread_indicator = "📬 UNREAD" if not email.is_read else "📭"
                    print(f"   {i+1}. {unread_indicator} {email.subject[:40]}...")
                    print(f"      From: {email.sender} | Score: {analysis.priority_score:.1f}")
                    # Handle both string categories (lite/ultra-lite) and object categories (deep)
                    category_display = category if isinstance(category, str) else category.category
                    print(f"      Action: {analysis.action_required} | Type: {analysis.email_type} | Category: {category_display}")
                    if hasattr(category, 'subcategory') and category.subcategory and category.subcategory not in ['general', 'specific project name', '']:
                        print(f"      📂 Subcategory: {category.subcategory}")
                    
                    # Security Analysis
                    if security and security.is_suspicious:
                        print(f"      🚨 SECURITY: {getattr(security, 'risk_level', 'UNKNOWN').upper()} RISK")
                        for threat in getattr(security, 'threats_detected', [])[:2]:
                            print(f"         ⚠️ {threat}")
                    
                    if hasattr(analysis, 'deadline_info') and analysis.deadline_info:
                        print(f"      ⏰ DEADLINE: {analysis.deadline_info}")
                    
                    if hasattr(analysis, 'task_breakdown') and analysis.task_breakdown:
                        print(f"      📋 Tasks:")
                        for task in analysis.task_breakdown:
                            print(f"         • {task}")
                        
                        # Add calendar event as a task if meeting detected
                        if self._should_create_calendar_event(email, analysis):
                            try:
                                meeting_suggestion = self._get_llm_meeting_suggestion(email)
                                if meeting_suggestion and meeting_suggestion.get('should_create_meeting'):
                                    event_title = meeting_suggestion.get('title', f'Meeting: {email.subject[:30]}...')
                                    event_date = meeting_suggestion.get('date', 'TBD')
                                    event_time = f"{meeting_suggestion.get('start_time', 'TBD')}-{meeting_suggestion.get('end_time', 'TBD')}"
                                    location = meeting_suggestion.get('location', '')
                                    participants = meeting_suggestion.get('participants', [])
                                    
                                    print(f"         • 🗓️  {event_title}")
                                    print(f"            📅 {event_date} ⏰ {event_time}")
                                    if location:
                                        print(f"            📍 {location}")
                                    if participants:
                                        attendees_str = ', '.join(participants[:2])
                                        if len(participants) > 2:
                                            attendees_str += f" +{len(participants) - 2} more"
                                        print(f"            👥 {attendees_str}")
                                    print(f"            💡 Review & create in sidebar →")
                                    
                                    # Store meeting info for interactive creation
                                    self._queue_calendar_event_for_streamlit(email, analysis, meeting_suggestion)
                            except:
                                # Don't show generic message - LLM already provided specific guidance
                                pass
                    
                    # Email Summary (for long emails)
                    if summary:
                        print(f"      📄 Summary ({summary.estimated_read_time}): {summary.summary_text[:80]}...")
                    
                    # Template Suggestions
                    if templates:
                        print(f"      📝 Quick Templates: {', '.join([t.name for t in templates[:2]])}")
                    
                    print()
            
            if urgent_emails:
                print("🟡 URGENT PRIORITY (70-84):")
                for i, (email, enhanced_analysis) in enumerate(urgent_emails):
                    analysis = enhanced_analysis['core_analysis']
                    security = enhanced_analysis['security_analysis']
                    summary = enhanced_analysis['email_summary']
                    templates = enhanced_analysis['template_suggestions']
                    category = enhanced_analysis['email_category']
                    smart_category = enhanced_analysis['smart_category']  # New AI categorization
                    cold_analysis = enhanced_analysis['cold_analysis']   # New cold email detection
                    
                    unread_indicator = "📬 UNREAD" if not email.is_read else "📭"
                    print(f"   {i+1}. {unread_indicator} {email.subject[:40]}...")
                    print(f"      From: {email.sender} | Score: {analysis.priority_score:.1f}")
                    # Handle None smart_category for Lite mode
                    category_display = smart_category.primary_category if smart_category else enhanced_analysis['email_category']
                    print(f"      Action: {analysis.action_required} | Type: {analysis.email_type} | Category: {category_display}")
                    if smart_category and smart_category.subcategory and smart_category.subcategory != 'general':
                        print(f"      📂 Project: {smart_category.subcategory}")
                    
                    # Cold Email Detection Results
                    if cold_analysis and cold_analysis.is_cold_email:
                        print(f"      ❄️ Cold Email: {cold_analysis.confidence:.1f} confidence - {cold_analysis.recommended_action}")
                    
                    # Smart Category Confidence (only in Deep mode)
                    if smart_category and hasattr(smart_category, 'confidence') and smart_category.confidence < 0.7:
                        print(f"      🤔 Category confidence: {smart_category.confidence:.1f} - {smart_category.reasoning}")
                    
                    # Security Analysis
                    if security and security.is_suspicious:
                        print(f"      ⚠️ Security: {getattr(security, 'risk_level', 'unknown')} risk")
                    
                    if hasattr(analysis, 'deadline_info') and analysis.deadline_info:
                        print(f"      ⏰ Deadline: {analysis.deadline_info}")
                    
                    if hasattr(analysis, 'task_breakdown') and analysis.task_breakdown:
                        print(f"      📋 Tasks:")
                        for task in analysis.task_breakdown:
                            print(f"         • {task}")
                        
                        # Add calendar event as a task if meeting detected
                        if self._should_create_calendar_event(email, analysis):
                            try:
                                meeting_suggestion = self._get_llm_meeting_suggestion(email)
                                if meeting_suggestion and meeting_suggestion.get('should_create_meeting'):
                                    event_title = meeting_suggestion.get('title', f'Meeting: {email.subject[:30]}...')
                                    event_date = meeting_suggestion.get('date', 'TBD')
                                    event_time = f"{meeting_suggestion.get('start_time', 'TBD')}-{meeting_suggestion.get('end_time', 'TBD')}"
                                    location = meeting_suggestion.get('location', '')
                                    participants = meeting_suggestion.get('participants', [])
                                    
                                    print(f"         • 🗓️  {event_title}")
                                    print(f"            📅 {event_date} ⏰ {event_time}")
                                    if location:
                                        print(f"            📍 {location}")
                                    if participants:
                                        attendees_str = ', '.join(participants[:2])
                                        if len(participants) > 2:
                                            attendees_str += f" +{len(participants) - 2} more"
                                        print(f"            👥 {attendees_str}")
                                    print(f"            💡 Review & create in sidebar →")
                                    
                                    # Store meeting info for interactive creation
                                    self._queue_calendar_event_for_streamlit(email, analysis, meeting_suggestion)
                            except:
                                # Don't show generic message - LLM already provided specific guidance
                                pass
                    
                    # Template Suggestions
                    if templates:
                        print(f"      📝 Templates: {', '.join([t.name for t in templates[:2]])}")
                    
                    print()
            
            if normal_emails:
                print("🟢 NORMAL PRIORITY (50-69):")
                for i, (email, enhanced_analysis) in enumerate(normal_emails):
                    analysis = enhanced_analysis['core_analysis']
                    security = enhanced_analysis['security_analysis']
                    templates = enhanced_analysis['template_suggestions']
                    category = enhanced_analysis['email_category']
                    
                    unread_indicator = "📬 UNREAD" if not email.is_read else "📭"
                    print(f"   {i+1}. {unread_indicator} {email.subject[:40]}...")
                    print(f"      From: {email.sender} | Score: {analysis.priority_score:.1f}")
                    # Handle both string categories (lite/ultra-lite) and object categories (deep)
                    category_display = category if isinstance(category, str) else category.category
                    print(f"      Action: {analysis.action_required} | Category: {category_display}")
                    if hasattr(category, 'subcategory') and category.subcategory:
                        print(f"      📂 Project: {category.subcategory}")
                    
                    # Security warnings for normal priority emails too
                    if (security and hasattr(security, 'is_suspicious') and security.is_suspicious and 
                        hasattr(security, 'risk_level') and security.risk_level in ["medium", "high", "critical"]):
                        print(f"      ⚠️ Security: {getattr(security, 'risk_level', 'unknown')} risk")
                    
                    if hasattr(analysis, 'task_breakdown') and analysis.task_breakdown:
                        print(f"      📋 Tasks:")
                        for task in analysis.task_breakdown:
                            print(f"         • {task}")
                        
                        # Add calendar event as a task if meeting detected
                        if self._should_create_calendar_event(email, analysis):
                            try:
                                meeting_suggestion = self._get_llm_meeting_suggestion(email)
                                if meeting_suggestion and meeting_suggestion.get('should_create_meeting'):
                                    event_title = meeting_suggestion.get('title', f'Meeting: {email.subject[:30]}...')
                                    event_date = meeting_suggestion.get('date', 'TBD')
                                    event_time = f"{meeting_suggestion.get('start_time', 'TBD')}-{meeting_suggestion.get('end_time', 'TBD')}"
                                    location = meeting_suggestion.get('location', '')
                                    participants = meeting_suggestion.get('participants', [])
                                    
                                    print(f"         • 🗓️  {event_title}")
                                    print(f"            📅 {event_date} ⏰ {event_time}")
                                    if location:
                                        print(f"            📍 {location}")
                                    if participants:
                                        attendees_str = ', '.join(participants[:2])
                                        if len(participants) > 2:
                                            attendees_str += f" +{len(participants) - 2} more"
                                        print(f"            👥 {attendees_str}")
                                    print(f"            💡 Review & create in sidebar →")
                                    
                                    # Store meeting info for interactive creation
                                    self._queue_calendar_event_for_streamlit(email, analysis, meeting_suggestion)
                            except:
                                # Don't show generic message - LLM already provided specific guidance
                                pass
                    
                    print()
            
            if low_emails:
                print("⚪ LOW PRIORITY (<50) - Consider batch processing:")
                for i, (email, enhanced_analysis) in enumerate(low_emails[:3]):  # Show only first 3
                    analysis = enhanced_analysis['core_analysis']
                    security = enhanced_analysis['security_analysis']
                    
                    security_warning = ""
                    if security and security.is_suspicious:
                        security_warning = f" 🚨 {getattr(security, 'risk_level', 'unknown')} risk"
                    
                    print(f"   {i+1}. {email.subject[:40]}... (Score: {analysis.priority_score:.1f}){security_warning}")
                if len(low_emails) > 3:
                    print(f"   ... and {len(low_emails) - 3} more low priority emails")
                print()
            
            # Step 5.5: Enhanced Feature Summary
            # Follow-up Tracking Summary
            due_follow_ups = self.follow_up_tracker.get_due_follow_ups()
            if due_follow_ups:
                print(f"📅 FOLLOW-UP REMINDERS ({len(due_follow_ups)}):")
                for follow_up in due_follow_ups[:3]:
                    status_emoji = "🔴" if follow_up.status == "overdue" else "🟡"
                    print(f"   {status_emoji} {follow_up.subject[:35]}... from {follow_up.sender}")
                    print(f"      Due: {follow_up.follow_up_date.strftime('%Y-%m-%d %H:%M')} | {follow_up.reason}")
                if len(due_follow_ups) > 3:
                    print(f"   ... and {len(due_follow_ups) - 3} more follow-ups")
                print()
            
            # Security Summary
            suspicious_emails = [(email, analysis) for email, analysis in analyzed_emails 
                               if analysis['security_analysis'] and analysis['security_analysis'].is_suspicious]
            if suspicious_emails:
                print(f"🚨 SECURITY ALERTS ({len(suspicious_emails)}):")
                for email, enhanced_analysis in suspicious_emails[:3]:
                    security = enhanced_analysis['security_analysis']
                    print(f"   🚨 {email.subject[:35]}... | {getattr(security, 'risk_level', 'UNKNOWN').upper()} RISK")
                    for rec in getattr(security, 'recommendations', [])[:2]:
                        print(f"      {rec}")
                if len(suspicious_emails) > 3:
                    print(f"   ... and {len(suspicious_emails) - 3} more security alerts")
                print()
            
            # Step 6: Generate LLM Drafts
            actionable_emails = []
            for email, enhanced_analysis in analyzed_emails:
                analysis = enhanced_analysis['core_analysis']
                
                # Skip automated/noreply emails
                sender_email = email.sender_email.lower()
                sender_name = email.sender.lower()
                
                # Enhanced no-reply detection patterns
                skip_patterns = [
                    'noreply', 'no-reply', 'donotreply', 'do-not-reply',
                    'automated', 'notification', 'system', 'security',
                    'marketing', 'marketing-email', 'marketing-replies',
                    'website+security@huggingface.co', 'website@huggingface.co',
                    'notifications@', 'support@', 'alerts@', 'admin@',
                    'aws-marketing-email-replies@amazon.com',
                    'confirm your email', 'click this link', 'verify your account',
                    'huggingface.co', 'github.com', 'gitlab.com',
                    'presidentoffice@mbzuai.ac.ae',  # Often announcement-only
                    'iec@mbzuai.ac.ae'  # IEC center announcements
                ]
                
                # Institutional exceptions (emails that might be actionable despite sender)
                institutional_actionable_patterns = [
                    'probation evaluation', 'evaluation required', 'action required',
                    'response needed', 'please respond', 'rsvp', 'confirm attendance'
                ]
                
                # Also check subject line for automated patterns
                subject_lower = email.subject.lower()
                skip_subject_patterns = [
                    'confirm your email', 'verify your account', 'click this link',
                    'account verification', 'email confirmation', 'click here to confirm'
                ]
                
                # Check for feedback forms/surveys (don't reply to service request forms)
                email_body_lower = email.body.lower()
                form_indicators = [
                    'please tell us about your experience',
                    'rate your experience', 
                    'service evaluation form',
                    'feedback form',
                    'satisfaction survey',
                    'rate our service',
                    'please rate',
                    'how would you rate'
                ]
                
                # Exclude "evaluation" by itself as it could be academic evaluations  
                is_form = any(indicator in email_body_lower for indicator in form_indicators)
                is_form = is_form or any(indicator in subject_lower for indicator in form_indicators)
                
                # However, academic/probation evaluations should NOT be skipped
                academic_evaluation_patterns = [
                    'probation evaluation',
                    'academic evaluation', 
                    'performance evaluation',
                    'evaluation required',
                    'evaluation request'
                ]
                
                is_academic_evaluation = any(pattern in subject_lower or pattern in email_body_lower 
                                           for pattern in academic_evaluation_patterns)
                
                # Don't treat academic evaluations as forms to skip
                if is_academic_evaluation:
                    is_form = False
                
                should_skip_subject = any(pattern in subject_lower for pattern in skip_subject_patterns)
                
                should_skip = any(pattern in sender_email or pattern in sender_name for pattern in skip_patterns)
                
                # Check for institutional actionable exceptions
                is_actionable_institutional = any(pattern in subject_lower or pattern in email_body_lower 
                                                for pattern in institutional_actionable_patterns)
                
                # Override skip decision for actionable institutional emails
                if should_skip and is_actionable_institutional:
                    print(f"   ✅ Institutional email requires action: {email.subject[:40]}...")
                    should_skip = False
                
                if should_skip or should_skip_subject:
                    print(f"   ⏭️ Skipping automated email from: {email.sender} <{email.sender_email}>")
                    continue
                
                if is_form:
                    print(f"   📋 Skipping form/survey email: {email.subject[:40]}...")
                    continue
                
                # Check if email needs a response (let LLM decide)
                should_reply = getattr(analysis, 'should_reply', True)  # Default to True if not set
                if (should_reply and 
                    analysis.priority_score >= priority_threshold and
                    analysis.action_required in ['reply', 'attend', 'approve', 'review'] and
                    analysis.email_type not in ['self_calendar_response', 'self_calendar_event']):
                    actionable_emails.append((email, analysis))
                elif not should_reply:
                    sender_email = getattr(email, 'sender_email', email.sender)
                    print(f"   🚫 LLM detected no-reply email: {email.subject[:40]}... (from: {sender_email})")
            
            if not actionable_emails:
                print("🎉 No emails need responses right now!")
                return
            
            print(f"🤖 Generating LLM drafts for {len(actionable_emails)} emails...")
            print("=" * 60)
            
            for i, (email, analysis) in enumerate(actionable_emails[:5]):
                print(f"\n✨ Generating LLM draft {i+1}: {email.subject[:40]}...")
                
                # Generate enhanced contextual draft (inspired by inbox-zero)
                draft = self.generate_contextual_draft(email, analysis)
                # Note: LLM calls are tracked within generate_contextual_draft
                
                # Update session state with draft information
                self._save_email_analysis(email, analysis, draft)
                
                print(f"   🤖 LLM Draft Generated:")
                print(f"   📧 Subject: {draft.subject}")
                print(f"   🎯 Tone: {draft.tone}")
                print(f"   🎪 Confidence: {draft.confidence:.2f}")
                print(f"   💭 Reasoning: {draft.reasoning}")
                
                # Create reply draft in Outlook with proper threading
                try:
                    draft_result = self.outlook.create_draft_reply(
                        original_email=email,
                        reply_body=draft.body
                    )
                    
                    # Debug output to check if To/From fields are being set
                    print(f"   📧 To: {email.sender} <{email.sender_email}>")
                    print(f"   👤 From: {current_user_name} <{current_user_email}>")
                    
                    if draft_result.get('success'):
                        self.drafts_created += 1
                        print(f"   ✅ Draft saved to Outlook Drafts folder")
                        
                        # Show preview
                        preview = draft.body[:150] + "..." if len(draft.body) > 150 else draft.body
                        print(f"   📖 Preview: {preview}")
                    else:
                        print(f"   ❌ Failed to create draft: {draft_result.get('error')}")
                        print(f"   📝 Draft content:\n{draft.body}")
                
                except Exception as e:
                    print(f"   ❌ Error creating draft: {e}")
                    print(f"   📝 Draft content:\\n{draft.body}")
            
            # Step 7: Professional Summary
            print(f"\n🎯 PROFESSIONAL EMAIL ASSISTANT SUMMARY")
            print("=" * 45)
            
            # Count by priority
            critical_count = len([e for e, a in analyzed_emails if a['core_analysis'].priority_score >= 85])
            urgent_count = len([e for e, a in analyzed_emails if 70 <= a['core_analysis'].priority_score < 85])
            unread_count = len([e for e, a in analyzed_emails if not e.is_read])
            deadline_count = len([e for e, a in analyzed_emails if hasattr(a['core_analysis'], 'deadline_info') and a['core_analysis'].deadline_info])
            
            # Enhanced feature statistics
            security_alerts = len([e for e, a in analyzed_emails if a['security_analysis'] and a['security_analysis'].is_suspicious])
            summarized_emails = len([e for e, a in analyzed_emails if a['email_summary'] is not None])
            total_follow_ups = len(self.follow_up_tracker.follow_ups)
            due_follow_ups_count = len(self.follow_up_tracker.get_due_follow_ups())
            
            # Category distribution
            categories = {}
            for email, analysis in analyzed_emails:
                category_obj = analysis['email_category']
                category = category_obj if isinstance(category_obj, str) else category_obj.category
                categories[category] = categories.get(category, 0) + 1
            
            print(f"📊 EMAIL ANALYSIS:")
            print(f"   📧 Total emails analyzed: {self.emails_analyzed}")
            print(f"   📬 Unread emails: {unread_count}")
            print(f"   🔴 Critical priority: {critical_count}")
            print(f"   🟡 Urgent priority: {urgent_count}")
            print(f"   ⏰ With deadlines: {deadline_count}")
            
            print(f"\n🛡️ SECURITY ANALYSIS:")
            print(f"   🚨 Security alerts: {security_alerts}")
            if security_alerts > 0:
                risk_levels = {}
                for email, analysis in analyzed_emails:
                    if analysis['security_analysis'] and analysis['security_analysis'].is_suspicious:
                        risk = getattr(analysis['security_analysis'], 'risk_level', 'unknown')
                        risk_levels[risk] = risk_levels.get(risk, 0) + 1
                for risk, count in risk_levels.items():
                    print(f"      {risk.capitalize()}: {count}")
            else:
                print(f"   ✅ All emails appear safe")
            
            print(f"\n📄 EMAIL SUMMARIZATION:")
            print(f"   📝 Long emails summarized: {summarized_emails}")
            if summarized_emails > 0:
                avg_read_time = "2-3 minutes"  # Could calculate actual average
                print(f"   ⏱️ Average reading time saved: {avg_read_time} per email")
            
            print(f"\n📅 FOLLOW-UP TRACKING:")
            print(f"   📋 Total follow-ups scheduled: {total_follow_ups}")
            print(f"   🔔 Due/overdue follow-ups: {due_follow_ups_count}")
            
            print(f"\n📂 SMART CATEGORIZATION:")
            top_categories = sorted(categories.items(), key=lambda x: x[1], reverse=True)[:3]
            for category, count in top_categories:
                print(f"   📁 {category.capitalize()}: {count} emails")
            
            print(f"\n📝 TEMPLATE SUGGESTIONS:")
            total_templates = sum(len(a['template_suggestions']) for e, a in analyzed_emails)
            print(f"   📋 Template suggestions generated: {total_templates}")
            most_suggested = {}
            for email, analysis in analyzed_emails:
                for template in analysis['template_suggestions']:
                    most_suggested[template.name] = most_suggested.get(template.name, 0) + 1
            if most_suggested:
                top_template = max(most_suggested, key=most_suggested.get)
                print(f"   🏆 Most suggested template: {top_template}")
            
            # Calendar event statistics
            print(f"\n📅 CALENDAR EVENT CREATION (TESTING MODE):")
            calendar_created = len([e for e, a in analyzed_emails if e.calendar_event_status == "created"])
            calendar_duplicates = len([e for e, a in analyzed_emails if e.calendar_event_status == "duplicate"])
            calendar_failed = len([e for e, a in analyzed_emails if e.calendar_event_status == "failed"])
            calendar_total = calendar_created + calendar_duplicates + calendar_failed
            
            print(f"   📅 Total meeting emails processed: {calendar_total}")
            print(f"   ✅ Personal calendar events created: {calendar_created}")
            print(f"   ⚠️ Duplicate events skipped: {calendar_duplicates}")
            if calendar_failed > 0:
                print(f"   ❌ Failed to create: {calendar_failed}")
            if calendar_total == 0:
                print(f"   📭 No meeting invitations found in emails")
            if calendar_created > 0:
                print(f"   👤 Note: All events created for personal calendar only")
                print(f"   📝 No invitations sent to other attendees (testing mode)")
            
            print(f"\n🤖 AI PROCESSING STATS:")
            print(f"   🧠 LLM calls made: {self.llm_calls}")
            print(f"   📝 Drafts created: {self.drafts_created}")
            print(f"   ⏱️ Estimated time saved: ~{self.drafts_created * 12} minutes")
            
            print(f"\n🎯 NEW FEATURES ACTIVE:")
            print(f"   ✅ Email Templates & Quick Responses")
            print(f"   ✅ Follow-up Tracking & Reminders")
            print(f"   ✅ Email Summarization (for long emails)")
            print(f"   ✅ Security Analysis & Threat Detection")
            print(f"   ✅ Smart Email Categorization")
            print(f"   ✅ Task Breakdown Generation")
            
            print(f"\n💡 NEXT STEPS:")
            print(f"   1. Review critical and urgent emails first")
            if security_alerts > 0:
                print(f"   2. 🚨 Address {security_alerts} security alert(s) immediately")
            if due_follow_ups_count > 0:
                print(f"   3. 📅 Handle {due_follow_ups_count} due follow-up(s)")
            print(f"   4. Check Outlook drafts folder for generated responses")
            print(f"   5. Use suggested templates for quick replies")
            print(f"   6. Review email summaries to save reading time")
            print(f"📁 All drafts saved to: Outlook > Drafts folder")
            
            # Create prioritized email list draft in Outlook AND store in session state
            try:
                self._create_priority_email_list_draft(analyzed_emails, current_user_email)
                # Store prioritized emails in session state for persistent display
                self._store_prioritized_emails_in_session(analyzed_emails)
            except Exception as e:
                print(f"⚠️ Could not create priority email list draft: {e}")
            
        except Exception as e:
            print(f"❌ Error in LLM processing: {e}")
            import traceback
            traceback.print_exc()
    
    def _analyze_email_thread(self, email: OutlookEmailData) -> str:
        """Analyze email thread context (open-source technique)"""
        
        thread_indicators = []
        
        # Check if it's a reply
        if email.subject.startswith(('Re:', 'RE:', 'Fwd:', 'FW:')):
            thread_indicators.append("This is part of an ongoing email thread")
        
        # Check for forwarded emails
        if 'forwarded' in email.body.lower() or 'from:' in email.body[:200].lower():
            thread_indicators.append("This email contains forwarded content")
        
        # Check for urgency escalation
        urgency_words = ['urgent', 'asap', 'immediate', 'deadline', 'overdue']
        if any(word in email.subject.lower() for word in urgency_words):
            thread_indicators.append("This email has urgency indicators")
        
        # Check for follow-up patterns
        followup_patterns = ['follow up', 'following up', 'reminder', 'checking in', 'update on']
        if any(pattern in email.body.lower()[:300] for pattern in followup_patterns):
            thread_indicators.append("This appears to be a follow-up email")
        
        # Check for meeting-related thread
        meeting_words = ['meeting', 'schedule', 'calendar', 'appointment', 'call']
        if any(word in email.subject.lower() for word in meeting_words):
            thread_indicators.append("This email is related to scheduling/meetings")
        
        return "; ".join(thread_indicators) if thread_indicators else "Standalone email without clear thread context"
    
    def _should_create_calendar_event(self, email: OutlookEmailData, analysis: LLMAnalysisResult) -> bool:
        """Determine if a calendar event should be created for this email"""
        # Check if calendar invites are enabled via environment variable
        import streamlit as st
        
        # Check environment variable for calendar invite functionality
        enable_calendar = True
        if hasattr(st, 'secrets') and 'ENABLE_CALENDAR_INVITES' in st.secrets:
            enable_calendar = st.secrets['ENABLE_CALENDAR_INVITES']
        
        if not enable_calendar:
            print(f"   📅 Calendar invites disabled by ENABLE_CALENDAR_INVITES setting")
            return False
        
        # Create calendar events for meeting invitations and events
        meeting_indicators = [
            'meeting', 'schedule', 'appointment', 'calendar', 'invite', 'invitation',
            'event', 'seminar', 'conference', 'workshop', 'lecture', 'talk'
        ]
        
        email_text = (email.subject + " " + email.body).lower()
        has_meeting_keywords = any(keyword in email_text for keyword in meeting_indicators)
        
        # Also check if it's categorized as a meeting or event
        is_meeting_type = analysis.email_type in ['meeting', 'event'] or analysis.action_required == 'attend'
        
        # Debug: Show what we found
        # Simplified calendar detection logging
        
        # Extract time/date information to confirm it's a scheduled event
        import re
        
        # Enhanced time detection patterns
        time_patterns = [
            r'\b(?:[01]?[0-9]|2[0-3]):[0-5][0-9]\s*(?:AM|PM|am|pm)?\b',  # 2:30 PM
            r'\b(?:[01]?[0-9]|2[0-3])-(?:[01]?[0-9]|2[0-3])\s*(?:AM|PM|am|pm)?\b',  # 2-4 PM
            r'\bbetween\s+(?:[01]?[0-9]|2[0-3])-(?:[01]?[0-9]|2[0-3])\s*(?:AM|PM|am|pm)?\b',  # between 2-4 PM
            r'\bfrom\s+(?:[01]?[0-9]|2[0-3])\s*(?:AM|PM|am|pm)?\s*to\s*(?:[01]?[0-9]|2[0-3])\s*(?:AM|PM|am|pm)?\b'  # from 2 PM to 4 PM
        ]
        
        # Enhanced date detection patterns
        date_patterns = [
            r'(?:january|february|march|april|may|june|july|august|september|october|november|december)\s+\d{1,2}(?:st|nd|rd|th)?',  # July 15th
            r'\d{1,2}[-/]\d{1,2}[-/]\d{2,4}',  # 7/15/2025
            r'(?:monday|tuesday|wednesday|thursday|friday|saturday|sunday)',  # Monday
            r'\b(?:mon|tue|wed|thu|fri|sat|sun)\b',  # Mon, Tue, etc.
            r'\btomorrow\b|\btoday\b|\bnext\s+week\b|\bthis\s+week\b',  # relative dates
            r'\d{1,2}(?:st|nd|rd|th)?\s+(?:january|february|march|april|may|june|july|august|september|october|november|december)'  # 15th July
        ]
        
        has_time = any(bool(re.search(pattern, email.body, re.IGNORECASE)) for pattern in time_patterns)
        has_date = any(bool(re.search(pattern, email.body, re.IGNORECASE)) for pattern in date_patterns)
        
        # Check if email already contains a calendar invitation
        calendar_indicators = [
            'icalendar', 'calendar.ics', 'meeting.ics', 'invite.ics',
            'outlook-calendar', 'vcalendar', 'vevent'
        ]
        email_text = (email.subject + " " + email.body).lower()
        has_calendar_attachment = any(indicator in email_text for indicator in calendar_indicators)
        
        # Don't create calendar events if:
        # 1. Email already has calendar invitation attachment
        # 2. Subject suggests it's an automatic calendar notification
        if has_calendar_attachment or 'invitation:' in email.subject.lower():
            print(f"   📅 Skipping calendar creation - email already contains calendar invitation")
            return False
        
        # Generate unique event key for this email
        # More flexible criteria: meeting keywords + (meeting type OR time/date info)
        should_create = (has_meeting_keywords and is_meeting_type) or (has_meeting_keywords and (has_time or has_date))
        
        if should_create:
            email.event_key = self._generate_event_key(email, analysis)
            # Don't print misleading message here - LLM will make final decision
            return True
        else:
            # No verbose logging for non-meetings
            pass
        
        return False
    
    def _should_ask_calendar_confirmation(self) -> bool:
        """Check if user confirmation is required before creating calendar events"""
        import streamlit as st
        import os
        
        # Check configuration setting for calendar confirmation requirement
        def get_config(key, default=None):
            # Check if we're running in Streamlit environment
            is_streamlit = False
            try:
                if hasattr(st, 'secrets'):
                    is_streamlit = True
                    if key in st.secrets:
                        value = st.secrets[key]
                        if value is not None:
                            return value
            except Exception:
                pass
            
            # Fallback to environment variables
            env_value = os.getenv(key)
            if env_value is not None:
                return env_value.lower() == 'true' if isinstance(env_value, str) else bool(env_value)
            
            return default
        
        # Return the actual setting value (default to True - require confirmation)
        return get_config('REQUIRE_CALENDAR_CONFIRMATION', True)
    
    def _ask_calendar_confirmation(self, email: OutlookEmailData, analysis: LLMAnalysisResult) -> bool:
        """Ask user for confirmation before creating a calendar event"""
        import streamlit as st
        
        # Create a unique key for this email's confirmation
        confirmation_key = f"calendar_confirm_{email.id}"
        
        # Check if decision was made previously
        if confirmation_key in st.session_state:
            return st.session_state[confirmation_key]
        
        # Queue this email for calendar confirmation (don't interrupt main flow)
        if 'pending_calendar_confirmations' not in st.session_state:
            st.session_state.pending_calendar_confirmations = {}
        
        st.session_state.pending_calendar_confirmations[confirmation_key] = {
            'email': email,
            'analysis': analysis,
            'email_id': email.id,
            'subject': email.subject,
            'sender': email.sender,
            'sender_email': email.sender_email,
            'priority_score': analysis.priority_score,
            'email_type': analysis.email_type,
            'body_preview': email.body[:200] + "..." if len(email.body) > 200 else email.body
        }
        
        # Return None to indicate we need confirmation later (don't block main processing)
        return None
    
    def _generate_event_key(self, email: OutlookEmailData, analysis: LLMAnalysisResult) -> str:
        """Generate a unique key for event deduplication based on email content"""
        import hashlib
        
        # Create a key based on:
        # - Subject (cleaned)
        # - Sender email
        # - Key words from body
        # - Date/time if found
        
        subject_clean = email.subject.lower().replace('re:', '').replace('fwd:', '').strip()
        
        # Extract key meeting details
        import re
        time_match = re.search(r'\b(?:[01]?[0-9]|2[0-3]):[0-5][0-9]\s*(?:AM|PM|am|pm)?\b', email.body)
        date_match = re.search(r'(?:january|february|march|april|may|june|july|august|september|october|november|december)\s+\d{1,2}(?:st|nd|rd|th)?', email.body, re.IGNORECASE)
        
        key_components = [
            subject_clean,
            email.sender_email.lower(),
            time_match.group(0).lower() if time_match else "",
            date_match.group(0).lower() if date_match else ""
        ]
        
        # Create hash of key components
        key_string = "|".join(key_components)
        return hashlib.md5(key_string.encode()).hexdigest()[:12]  # Short unique key
    
    def _parse_meeting_datetime(self, date_str: str, time_str: str) -> datetime:
        """Parse extracted date and time strings into datetime object"""
        from datetime import datetime, timedelta
        import re
        
        if not date_str:
            raise ValueError("No date string provided")
        
        # Parse time
        hour, minute = 10, 0  # Default time
        if time_str:
            hour, minute = self._parse_time_string(time_str)
        
        # Parse date based on different formats
        date_str_lower = date_str.lower().strip()
        
        # Handle "today" and "tomorrow"
        if 'today' in date_str_lower:
            return datetime.now().replace(hour=hour, minute=minute, second=0, microsecond=0)
        elif 'tomorrow' in date_str_lower:
            return (datetime.now() + timedelta(days=1)).replace(hour=hour, minute=minute, second=0, microsecond=0)
        
        # Handle day of week (this/next Monday, Tuesday, etc.)
        weekdays = ['monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday', 'sunday']
        for i, day in enumerate(weekdays):
            if day in date_str_lower:
                today = datetime.now()
                days_ahead = i - today.weekday()
                if days_ahead <= 0 or 'next' in date_str_lower:  # Target day already happened this week or explicitly "next"
                    days_ahead += 7
                target_date = today + timedelta(days=days_ahead)
                return target_date.replace(hour=hour, minute=minute, second=0, microsecond=0)
        
        # Handle "Thursday, Aug 21, 2025" format (your specific case)
        match = re.match(r'(?:monday|tuesday|wednesday|thursday|friday|saturday|sunday),?\s+(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*\s+(\d{1,2}),?\s+(\d{4})', date_str_lower)
        if match:
            month_abbr, day, year = match.groups()
            month_map = {'jan': 1, 'feb': 2, 'mar': 3, 'apr': 4, 'may': 5, 'jun': 6,
                        'jul': 7, 'aug': 8, 'sep': 9, 'oct': 10, 'nov': 11, 'dec': 12}
            month = month_map.get(month_abbr[:3], 1)
            print(f"   🎯 Parsed specific format: {date_str} → {year}-{month:02d}-{day} {hour:02d}:{minute:02d}")
            return datetime(int(year), month, int(day), hour, minute)

        # Handle "Thu 21 Aug 2025" format
        match = re.match(r'(?:mon|tue|wed|thu|fri|sat|sun)[a-z]*\s+(\d{1,2})\s+(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*\s+(\d{4})', date_str_lower)
        if match:
            day, month_abbr, year = match.groups()
            month_map = {'jan': 1, 'feb': 2, 'mar': 3, 'apr': 4, 'may': 5, 'jun': 6,
                        'jul': 7, 'aug': 8, 'sep': 9, 'oct': 10, 'nov': 11, 'dec': 12}
            month = month_map.get(month_abbr[:3], 1)
            print(f"   🎯 Parsed short format: {date_str} → {year}-{month:02d}-{day} {hour:02d}:{minute:02d}")
            return datetime(int(year), month, int(day), hour, minute)
        
        # Handle full month names like "August 21, 2025"
        month_names = ['january', 'february', 'march', 'april', 'may', 'june',
                      'july', 'august', 'september', 'october', 'november', 'december']
        for i, month_name in enumerate(month_names):
            if month_name in date_str_lower:
                # Extract day and year
                day_match = re.search(r'\b(\d{1,2})(?:st|nd|rd|th)?\b', date_str_lower)
                year_match = re.search(r'\b(\d{4})\b', date_str_lower)
                
                day = int(day_match.group(1)) if day_match else 1
                year = int(year_match.group(1)) if year_match else datetime.now().year
                
                return datetime(year, i + 1, day, hour, minute)
        
        # Handle numeric formats like "21/08/2025"
        date_match = re.match(r'(\d{1,2})[-/](\d{1,2})[-/](\d{2,4})', date_str_lower)
        if date_match:
            day, month, year = date_match.groups()
            if len(year) == 2:
                year = int(year) + 2000
            return datetime(int(year), int(month), int(day), hour, minute)
        
        raise ValueError(f"Could not parse date format: {date_str}")
    
    def _parse_time_string(self, time_str: str) -> tuple:
        """Parse time string and return (hour, minute) tuple"""
        import re
        
        time_str = time_str.strip().upper()
        
        # Handle "11:00 AM" or "2:30 PM" format
        match = re.match(r'(\d{1,2}):(\d{2})\s*(AM|PM)?', time_str)
        if match:
            hour, minute, period = match.groups()
            hour, minute = int(hour), int(minute)
            
            if period == 'PM' and hour != 12:
                hour += 12
            elif period == 'AM' and hour == 12:
                hour = 0
            
            return hour, minute
        
        # Handle "11 AM" or "2 PM" format
        match = re.match(r'(\d{1,2})\s*(AM|PM)', time_str)
        if match:
            hour, period = match.groups()
            hour = int(hour)
            
            if period == 'PM' and hour != 12:
                hour += 12
            elif period == 'AM' and hour == 12:
                hour = 0
            
            return hour, 0
        
        # Handle 24-hour format like "14:30"
        match = re.match(r'(\d{1,2}):(\d{2})', time_str)
        if match:
            hour, minute = match.groups()
            return int(hour), int(minute)
        
        # Default fallback
        return 10, 0
    
    def _analyze_emails_ultra_lite_mode(self, emails, current_user_name: str, current_user_email: str):
        """ULTRA FAST analysis - keyword-based, LLM for drafting only"""
        analyzed_emails = []
        
        import time
        start_time = time.time()
        
        # Get configurable timeout
        def get_safety_config(key, default):
            try:
                import streamlit as st
                if hasattr(st, 'secrets') and key in st.secrets:
                    return int(st.secrets[key])
            except:
                pass
            return default
        
        TIMEOUT_SECONDS = get_safety_config('TIMEOUT_ULTRA_LITE_SECONDS', 180)  # 3 minutes default
        
        print(f"🏃 ULTRA LITE: Keyword analysis only...")
        
        for i, email in enumerate(emails):
            # Check timeout every 20 emails (since ultra-lite is fastest)
            if i > 0 and i % 20 == 0:
                elapsed = time.time() - start_time
                if elapsed > TIMEOUT_SECONDS:
                    print(f"⏰ TIMEOUT: Ultra-lite analysis stopped after {elapsed:.1f}s to prevent system overload")
                    print(f"✅ Successfully processed {len(analyzed_emails)} emails before timeout")
                    break
            try:
                print(f"   🏃 Quick scan: {email.subject[:30]}...")
                
                # Keyword-based analysis only
                analysis = self._create_minimal_analysis(email)
                
                # Calendar Event Creation (even in ultra-lite mode for meeting emails)
                calendar_event = None
                if self._should_create_calendar_event(email, analysis) and analysis.priority_score > 40:
                    # Check if user confirmation is required
                    if self._should_ask_calendar_confirmation():
                        confirmation = self._ask_calendar_confirmation(email, analysis)
                        if confirmation is True:
                            calendar_event = self._create_calendar_event_from_email(email, analysis)
                        elif confirmation is False:
                            print(f"   📅 User declined calendar event creation for: {email.subject}")
                        else:
                            # Confirmation is None - still waiting for user input
                            print(f"   📅 Calendar event queued for confirmation - check sidebar to approve/decline")
                    else:
                        # Create calendar event without confirmation
                        calendar_event = self._create_calendar_event_from_email(email, analysis)
                    
                    # Update email with calendar event status
                    if calendar_event:
                        if calendar_event.get('success'):
                            email.created_calendar_event_id = calendar_event.get('event_id')
                            email.calendar_event_status = "created"
                        elif calendar_event.get('duplicate'):
                            email.created_calendar_event_id = calendar_event.get('existing_event_id')
                            email.calendar_event_status = "duplicate"
                        else:
                            email.calendar_event_status = "failed"
                
                # Update email with minimal analysis
                email.priority_score = analysis.priority_score
                email.urgency_level = analysis.urgency_level
                email.email_type = analysis.email_type
                email.action_required = analysis.action_required
                
                # Ultra minimal analysis result
                enhanced_analysis = {
                    'core_analysis': analysis,
                    'security_analysis': None,
                    'email_summary': None,
                    'template_suggestions': [],
                    'email_category': "general",
                    'smart_category': None,
                    'cold_analysis': None,
                    'automation_actions': [],
                    'calendar_event': calendar_event
                }
                
                analyzed_emails.append((email, enhanced_analysis))
                # NO LLM calls for analysis - only for drafting later
                
            except Exception as e:
                print(f"❌ Error analyzing email '{email.subject[:40]}': {str(e)}")
                continue
        
        return analyzed_emails
    
    def _analyze_emails_lite_mode(self, emails, current_user_name: str, current_user_email: str):
        """FAST LLM analysis - core features only"""
        analyzed_emails = []
        
        import time
        start_time = time.time()
        
        # Get configurable timeout
        def get_safety_config(key, default):
            try:
                import streamlit as st
                if hasattr(st, 'secrets') and key in st.secrets:
                    return int(st.secrets[key])
            except:
                pass
            return default
        
        TIMEOUT_SECONDS = get_safety_config('TIMEOUT_LITE_MODE_SECONDS', 240)  # 4 minutes default
        
        print(f"⚡ LITE: Fast LLM analysis...")
        
        for i, email in enumerate(emails):
            # Check timeout every 10 emails (since lite is faster)
            if i > 0 and i % 10 == 0:
                elapsed = time.time() - start_time
                if elapsed > TIMEOUT_SECONDS:
                    print(f"⏰ TIMEOUT: Lite analysis stopped after {elapsed:.1f}s to prevent system overload")
                    print(f"✅ Successfully processed {len(analyzed_emails)} emails before timeout")
                    break
            try:
                print(f"   ⚡ Analyzing: {email.subject[:40]}...")
                
                # Minimal context for speed
                thread_context = ""
                user_context = f"User: {current_user_name}, Email: {current_user_email}, Role: Professional at MBZUAI"
                
                # Core LLM Analysis only
                analysis = self.analyzer.analyze_email(email, user_context, thread_context)
                
                # Save analysis to session state for persistence
                self._save_email_analysis(email, analysis)
                
                # Quick categorization
                email_category = self.categorizer.categorize_email(email, analysis)
                
                # Calendar Event Creation (for meeting emails with time info)
                calendar_event = None
                if self._should_create_calendar_event(email, analysis) and analysis.priority_score > 40:
                    # Check if user confirmation is required
                    if self._should_ask_calendar_confirmation():
                        confirmation = self._ask_calendar_confirmation(email, analysis)
                        if confirmation is True:
                            calendar_event = self._create_calendar_event_from_email(email, analysis)
                        elif confirmation is False:
                            print(f"   📅 User declined calendar event creation for: {email.subject}")
                        else:
                            # Confirmation is None - still waiting for user input
                            print(f"   📅 Calendar event queued for confirmation - check sidebar to approve/decline")
                    else:
                        # Create calendar event without confirmation
                        calendar_event = self._create_calendar_event_from_email(email, analysis)
                    
                    # Update email with calendar event status
                    if calendar_event:
                        if calendar_event.get('success'):
                            email.created_calendar_event_id = calendar_event.get('event_id')
                            email.calendar_event_status = "created"
                        elif calendar_event.get('duplicate'):
                            email.created_calendar_event_id = calendar_event.get('existing_event_id')
                            email.calendar_event_status = "duplicate"
                        else:
                            email.calendar_event_status = "failed"
                
                # Update email with LLM analysis
                email.priority_score = analysis.priority_score
                email.urgency_level = analysis.urgency_level
                email.email_type = analysis.email_type
                email.action_required = analysis.action_required
                
                # Lite analysis result
                enhanced_analysis = {
                    'core_analysis': analysis,
                    'security_analysis': None,
                    'email_summary': None,
                    'template_suggestions': [],
                    'email_category': email_category,
                    'smart_category': None,
                    'cold_analysis': None,
                    'automation_actions': [],
                    'calendar_event': calendar_event
                }
                
                analyzed_emails.append((email, enhanced_analysis))
                self.llm_calls += 1
                
            except Exception as e:
                print(f"❌ Error analyzing email '{email.subject[:40]}': {str(e)}")
                continue
        
        return analyzed_emails
    
    def _create_minimal_analysis(self, email):
        """Create analysis based on keywords only - NO LLM calls"""
        subject_lower = email.subject.lower()
        body_lower = email.body.lower()
        sender_lower = email.sender.lower()
        
        # Quick keyword-based priority scoring
        priority_score = 50.0  # Default
        urgency_level = "medium"
        email_type = "general"
        action_required = "monitor"
        
        # High priority keywords
        urgent_keywords = ['urgent', 'asap', 'deadline', 'emergency', 'critical', 'important']
        meeting_keywords = ['meeting', 'call', 'zoom', 'teams', 'conference', 'appointment']
        action_keywords = ['reply', 'respond', 'answer', 'feedback', 'review', 'approve', 'confirm']
        
        # Check for urgent keywords
        if any(keyword in subject_lower or keyword in body_lower for keyword in urgent_keywords):
            priority_score += 20
            urgency_level = "high"
        
        # Check for meeting keywords
        if any(keyword in subject_lower or keyword in body_lower for keyword in meeting_keywords):
            priority_score += 15
            email_type = "meeting"
            action_required = "attend"
        
        # Check for action keywords
        if any(keyword in subject_lower or keyword in body_lower for keyword in action_keywords):
            priority_score += 10
            action_required = "reply"
        
        # Check sender importance (simple heuristic)
        if any(domain in sender_lower for domain in ['mbzuai.ac.ae', 'boss', 'ceo', 'director']):
            priority_score += 15
        
        # Ensure bounds
        priority_score = min(max(priority_score, 10), 100)
        
        # Create minimal analysis object
        from types import SimpleNamespace
        analysis = SimpleNamespace()
        analysis.priority_score = priority_score
        analysis.urgency_level = urgency_level
        analysis.email_type = email_type
        analysis.action_required = action_required
        analysis.reasoning = "Quick keyword analysis"
        analysis.should_reply = action_required in ["reply", "attend", "approve"]
        analysis.deadline_info = None
        analysis.task_breakdown = None
        analysis.confidence = 1.0  # High confidence for keyword-based analysis
        analysis.sender_relationship = "unknown"
        
        return analysis
    
    def _analyze_emails_deep_mode(self, emails, current_user_name: str, current_user_email: str):
        """Comprehensive analysis with all features"""
        analyzed_emails = []
        
        import time
        start_time = time.time()
        
        # Get configurable timeout
        def get_safety_config(key, default):
            try:
                import streamlit as st
                if hasattr(st, 'secrets') and key in st.secrets:
                    return int(st.secrets[key])
            except:
                pass
            return default
        
        TIMEOUT_SECONDS = get_safety_config('TIMEOUT_DEEP_MODE_SECONDS', 300)  # 5 minutes default
        
        for i, email in enumerate(emails):
            # Show progress every 10 emails
            if i > 0 and i % 10 == 0:
                progress_pct = (i / len(emails)) * 100
                print(f"📈 Progress: {i}/{len(emails)} emails ({progress_pct:.1f}%)")
            
            # Check timeout every few emails
            if i > 0 and i % 5 == 0:
                elapsed = time.time() - start_time
                if elapsed > TIMEOUT_SECONDS:
                    print(f"⏰ TIMEOUT: Analysis stopped after {elapsed:.1f}s to prevent system overload")
                    print(f"✅ Successfully processed {len(analyzed_emails)} emails before timeout")
                    break
            try:
                print(f"   📧 {i+1}/{len(emails)} Analyzing: {email.subject[:40]}...")
                
                # Enhanced context with thread analysis
                thread_context = self._analyze_email_thread(email)
                user_context = f"User: {current_user_name}, Email: {current_user_email}, Role: Professional at MBZUAI"
                
                # Step 4a: Core LLM Analysis
                analysis = self.analyzer.analyze_email(email, user_context, thread_context)
                
                # Save analysis to session state for persistence
                self._save_email_analysis(email, analysis)
                
                # Step 4b: Security Analysis (lightweight)
                security_analysis = self.security_analyzer.analyze_security(email)
                
                # Step 4c: Email Summarization (only for very long emails)
                email_summary = None
                if len(email.body) > 2000 and self.summarizer.should_summarize(email):
                    email_summary = self.summarizer.generate_summary(email)
                    self.llm_calls += 1
                
                # Step 4d: Template Suggestions (lightweight)
                template_suggestions = self.template_manager.get_template_suggestions(email, analysis)
                
                # Step 4e: Smart Categorization (use cached results when possible)
                smart_category = self.smart_categorizer.categorize_email(email)
                
                # Step 4f: Cold Email Detection (skip if high priority)
                cold_analysis = None
                if self.cold_email_detector and analysis.priority_score < 70:
                    cold_analysis = self.cold_email_detector.detect_cold_email(email)
                
                # Step 4g: Legacy Categorization (for compatibility)
                email_category = self.categorizer.categorize_email(email, analysis)
                
                # Step 4h: Email Automation Rules (only for high priority emails)
                automation_actions = []
                if self.automation_engine and analysis.priority_score > 60:
                    automation_actions = self.automation_engine.process_email_with_rules(email, {
                        'core_analysis': analysis,
                        'smart_category': smart_category,
                        'cold_analysis': cold_analysis
                    })
                
                # Step 4i: Follow-up Tracking (only if really action required)
                if analysis.action_required in ["reply", "attend", "approve", "review"] and analysis.priority_score > 50:
                    follow_up_item = self.follow_up_tracker.add_follow_up(email, analysis)
                
                # Step 4j: Calendar Event Creation (only for meeting emails with time info)
                calendar_event = None
                if self._should_create_calendar_event(email, analysis) and analysis.priority_score > 40:
                    # Check if user confirmation is required
                    if self._should_ask_calendar_confirmation():
                        confirmation = self._ask_calendar_confirmation(email, analysis)
                        if confirmation is True:
                            calendar_event = self._create_calendar_event_from_email(email, analysis)
                        elif confirmation is False:
                            print(f"   📅 User declined calendar event creation for: {email.subject}")
                        else:
                            # Confirmation is None - still waiting for user input
                            print(f"   📅 Calendar event queued for confirmation - check sidebar to approve/decline")
                    else:
                        # Create calendar event without confirmation
                        calendar_event = self._create_calendar_event_from_email(email, analysis)
                    
                    # Update email with calendar event status
                    if calendar_event:
                        if calendar_event.get('success'):
                            email.created_calendar_event_id = calendar_event.get('event_id')
                            email.calendar_event_status = "created"
                        elif calendar_event.get('duplicate'):
                            email.created_calendar_event_id = calendar_event.get('existing_event_id')
                            email.calendar_event_status = "duplicate"
                        else:
                            email.calendar_event_status = "failed"
                
                # Update email with LLM analysis
                email.priority_score = analysis.priority_score
                email.urgency_level = analysis.urgency_level
                email.email_type = analysis.email_type
                email.action_required = analysis.action_required
                
                # Store all analysis results together
                enhanced_analysis = {
                    'core_analysis': analysis,
                    'security_analysis': security_analysis,
                    'email_summary': email_summary,
                    'template_suggestions': template_suggestions,
                    'email_category': email_category,
                    'smart_category': smart_category,  # New: AI-powered categorization
                    'cold_analysis': cold_analysis,   # New: Cold email detection
                    'automation_actions': automation_actions,  # New: Rule-based automation
                    'calendar_event': calendar_event
                }
                
                analyzed_emails.append((email, enhanced_analysis))
                self.llm_calls += 1
                
            except Exception as e:
                print(f"❌ Error analyzing email '{email.subject[:40]}': {str(e)}")
                import traceback
                traceback.print_exc()
                continue
        
        return analyzed_emails
    
    def _create_calendar_event_from_email(self, email: OutlookEmailData, analysis: LLMAnalysisResult) -> Dict:
        """Create a calendar event from meeting email details using LLM"""
        try:
            from datetime import datetime, timedelta
            
            email_full_text = email.subject + " " + email.body
            print(f"   📧 Using LLM to extract meeting details: {email.subject[:50]}...")
            
            # Use LLM to extract meeting details with structured output
            meeting_suggestion = self._get_llm_meeting_suggestion(email)
            
            if not meeting_suggestion or not meeting_suggestion.get('should_create_meeting'):
                print(f"   🚫 LLM determined no meeting should be created")
                return {"success": False, "message": "LLM determined no meeting needed"}
            
            print(f"   🤖 LLM suggested meeting: {meeting_suggestion['title']}")
            print(f"   📅 Proposed time: {meeting_suggestion['date']} at {meeting_suggestion['start_time']}-{meeting_suggestion['end_time']}")
            if meeting_suggestion.get('participants'):
                print(f"   👥 Participants: {', '.join(meeting_suggestion['participants'])}")
            
            # Convert LLM suggestion to calendar event format
            try:
                print(f"   🔍 Parsing datetime: {meeting_suggestion['date']} at {meeting_suggestion['start_time']}")
                event_datetime = self._parse_llm_datetime(
                    meeting_suggestion['date'], 
                    meeting_suggestion['start_time']
                )
                print(f"   📅 Initial datetime: {event_datetime}")
                
                # Check office hours and adjust if needed
                original_datetime = event_datetime
                event_datetime = self._adjust_for_office_hours(event_datetime, meeting_suggestion)
                if event_datetime != original_datetime:
                    print(f"   ⏰ Adjusted for office hours: {original_datetime} → {event_datetime}")
                
                duration_hours = self._calculate_meeting_duration(
                    meeting_suggestion['start_time'], 
                    meeting_suggestion['end_time']
                )
                print(f"   ⏱️ Duration: {duration_hours} hours")
                
                # Create times for Asia/Dubai timezone (UTC+4) 
                start_time = event_datetime.isoformat()
                end_time = (event_datetime + timedelta(hours=duration_hours)).isoformat()
                print(f"   📅 Final times (Asia/Dubai): {start_time} to {end_time}")
                
            except Exception as e:
                print(f"   ⚠️ Could not parse LLM date/time: {e}")
                # Fallback to next business day 2 PM
                event_datetime = self._get_next_business_day_time(14, 0)  # 2 PM
                start_time = event_datetime.isoformat()
                end_time = (event_datetime + timedelta(hours=2)).isoformat()
                print(f"   📅 Fallback times (Asia/Dubai): {start_time} to {end_time}")
            
            # Create enhanced description with participants and context
            participants = meeting_suggestion.get('participants', [])
            location = meeting_suggestion.get('location', '')
            
            description = f"""📧 Meeting created from email: {email.subject}
            
👤 Original organizer: {email.sender} ({email.sender_email})
📅 Auto-created for personal calendar tracking

👥 Meeting participants (DO NOT SEND INVITES - PERSONAL TRACKING ONLY):
{chr(10).join([f"   • {participant}" for participant in participants]) if participants else "   • No specific participants mentioned"}

🎯 Meeting purpose: {meeting_suggestion.get('purpose', 'General discussion')}
📝 Notes: {meeting_suggestion.get('notes', 'Meeting details extracted from email')}

📧 Original email content:
{email.body[:500]}{'...' if len(email.body) > 500 else ''}

⚠️ IMPORTANT: This is a PERSONAL calendar entry only. 
No invitations have been sent to other participants.
You may need to send meeting invitations manually if required."""
            
            event_details = {
                'subject': meeting_suggestion.get('title', f"Meeting: {email.subject}"),
                'start_time': start_time,
                'end_time': end_time,
                'description': description,
                'location': location,
                'attendees': []  # TESTING MODE: No attendees, personal calendar only
            }
            
            # Create the calendar event
            result = self.outlook.create_meeting_from_email(email, event_details)
            
            if result.get('success'):
                print(f"📅 Personal calendar event created: [PERSONAL] {event_details['subject']}")
                print(f"   🕐 Time: {start_time}")
                if location:
                    print(f"   📍 Location: {location}")
                print(f"   👤 Original organizer: {email.sender}")
                print(f"   ⚠️ Note: Personal calendar entry only, no invitations sent")
            elif result.get('duplicate'):
                print(f"⚠️ Calendar event already exists: {result.get('existing_event_subject')}")
                print(f"   🕐 Existing time: {result.get('existing_event_start')}")
                print(f"   📝 Skipping duplicate creation")
            else:
                print(f"❌ Failed to create calendar event: {result.get('error', 'Unknown error')}")
            
            return result
            
        except Exception as e:
            print(f"❌ Error creating calendar event from email: {e}")
            return {
                "success": False,
                "error": str(e),
                "message": f"Failed to create calendar event: {str(e)}"
            }
    
    def _create_calendar_event_with_custom_details(self, email: OutlookEmailData, edited_details: dict) -> Dict:
        """Create a calendar event using user-edited details from the UI"""
        try:
            from datetime import datetime, timedelta
            import pytz
            
            print(f"   📧 Creating calendar event with custom details: {edited_details['title']}")
            
            # Parse the date and time from edited details
            event_date = edited_details['date']  # YYYY-MM-DD format
            start_time_str = edited_details['start_time']  # HH:MM format
            end_time_str = edited_details['end_time']    # HH:MM format
            
            # Create datetime objects in Dubai timezone
            dubai_tz = pytz.timezone('Asia/Dubai')
            
            # Parse start datetime
            start_datetime_str = f"{event_date}T{start_time_str}:00"
            start_datetime = datetime.fromisoformat(start_datetime_str)
            start_datetime = dubai_tz.localize(start_datetime)
            
            # Parse end datetime  
            end_datetime_str = f"{event_date}T{end_time_str}:00"
            end_datetime = datetime.fromisoformat(end_datetime_str)
            end_datetime = dubai_tz.localize(end_datetime)
            
            print(f"   📅 Scheduled time (Asia/Dubai): {start_datetime} to {end_datetime}")
            
            # Create enhanced description with original email context
            description = f"""📧 Meeting created from email: {edited_details['original_email_subject']}
            
👤 Original organizer: {email.sender} ({email.sender_email})
📅 Auto-created for personal calendar tracking

🎯 Meeting purpose: {edited_details.get('purpose', 'Meeting discussion')}
📝 Notes: {edited_details.get('description', 'No additional notes')}

👥 Meeting participants (DO NOT SEND INVITES - PERSONAL TRACKING ONLY):
{chr(10).join([f"   • {participant}" for participant in edited_details.get('participants', [])]) if edited_details.get('participants') else "   • No specific participants mentioned"}

📧 Original email content:
{edited_details.get('original_email_body', email.body)[:500]}{'...' if len(edited_details.get('original_email_body', email.body)) > 500 else ''}

⚠️ IMPORTANT: This is a PERSONAL calendar entry only. 
No invitations have been sent to other participants.
You may need to send meeting invitations manually if required."""
            
            # Prepare event details for Outlook
            event_details = {
                'subject': edited_details['title'],
                'start_time': start_datetime.isoformat(),
                'end_time': end_datetime.isoformat(),
                'description': description,
                'location': edited_details.get('location', ''),
                'attendees': []  # Personal calendar only - no attendees
            }
            
            # Create the calendar event
            result = self.outlook.create_meeting_from_email(email, event_details)
            
            if result.get('success'):
                print(f"📅 Personal calendar event created with custom details: [PERSONAL] {event_details['subject']}")
                print(f"   🕐 Time: {start_datetime} to {end_datetime}")
                if edited_details.get('location'):
                    print(f"   📍 Location: {edited_details['location']}")
                print(f"   👤 Original organizer: {email.sender}")
                print(f"   ⚠️ Note: Personal calendar entry only, no invitations sent")
                
                return {
                    "success": True,
                    "message": f"Calendar event created: {event_details['subject']}",
                    "event_details": event_details
                }
            elif result.get('duplicate'):
                print(f"⚠️ Calendar event already exists: {result.get('existing_event_subject')}")
                print(f"   🕐 Existing time: {result.get('existing_event_start')}")
                print(f"   📝 Skipping duplicate creation")
                return {
                    "success": False,
                    "duplicate": True,
                    "message": f"Duplicate event exists: {result.get('existing_event_subject')}",
                    "existing_event": result
                }
            else:
                error_msg = result.get('error', 'Unknown error occurred')
                print(f"❌ Failed to create calendar event: {error_msg}")
                return {
                    "success": False,
                    "error": error_msg,
                    "message": f"Failed to create calendar event: {error_msg}"
                }
                
        except Exception as e:
            print(f"❌ Error creating calendar event with custom details: {e}")
            return {
                "success": False,
                "error": str(e),
                "message": f"Failed to create calendar event with custom details: {str(e)}"
            }
    
    def _get_llm_meeting_suggestion(self, email: OutlookEmailData) -> dict:
        """Use LLM to extract meeting details with structured output"""
        
        current_date = datetime.now().strftime("%A, %B %d, %Y")
        current_time = datetime.now().strftime("%I:%M %p")
        
        prompt = f"""
You are a meeting scheduling assistant. Analyze the following email and determine if a calendar event should be created.

CURRENT CONTEXT:
- Today: {current_date}
- Current time: {current_time}
- Office hours: 9:00 AM - 6:00 PM, Monday-Friday

EMAIL TO ANALYZE:
Subject: {email.subject}
From: {email.sender} ({email.sender_email})
Body: {email.body}

CORE PRINCIPLE: Only create calendar events when YOU are expected to send invitations to others.

SMART DECISION LOGIC:
❌ DO NOT create events when:
- Someone else is organizing and will send invites (HR, professors, departments, etc.)
- Email already contains meeting links/IDs (organizer handles invites)
- You're being invited to something already organized
- It's an announcement about an existing event

✅ DO create events ONLY when:
- You're asked to schedule a meeting and send invites
- Someone requests YOU to organize a meeting
- You need to coordinate schedules and send calendar invites
- The email asks you to "set up a meeting" or "schedule time"

INSTRUCTIONS:
1. Ask yourself: "Am I being asked to ORGANIZE and SEND INVITES for this meeting?"
2. If YES → create the meeting. If NO → don't create it.
3. Extract meeting details if applicable - PRESERVE ORIGINAL TIMES EXACTLY
4. If meeting is outside office hours or on weekend, suggest next business day during office hours
5. Identify all participants mentioned in the email

CRITICAL: Use 24-hour format for times. "2-4 PM" = start_time: "14:00", end_time: "16:00"

Respond in this EXACT JSON format:
{{
    "should_create_meeting": true/false,
    "title": "Meeting title/subject",
    "date": "YYYY-MM-DD",
    "start_time": "HH:MM",
    "end_time": "HH:MM", 
    "participants": ["Dr. Ramzi", "Avani Gupta", "etc"],
    "location": "location if mentioned or suggested",
    "purpose": "Brief description of meeting purpose",
    "notes": "Additional notes or context",
    "office_hours_adjusted": true/false,
    "confidence": 0.0-1.0,
    "reasoning": "Why this meeting should/shouldn't be created"
}}

TIME CONVERSION EXAMPLES:
- "2-4 PM" → start_time: "14:00", end_time: "16:00"
- "9:30 AM" → start_time: "09:30", end_time: "10:30"
- "11am-12pm" → start_time: "11:00", end_time: "12:00"
- "1:15-2:45 PM" → start_time: "13:15", end_time: "14:45"

EXAMPLES:
✅ CREATE EVENTS FOR:
- "Can we meet Monday 2-4 PM to discuss the project?" → should_create_meeting: true
- "Let's schedule a call to review the proposal" → should_create_meeting: true

EXAMPLES:

❌ DON'T CREATE (someone else organizes):
- "Fuel Up Fridays session tomorrow 11-11:30 AM" → false (HR will send invite)
- "Research talk Monday 11 AM-12 PM" → false (professor/department organized)
- "Join our team meeting: Teams ID 123..." → false (already has meeting link)
- "You're invited to..." → false (already organized)

✅ CREATE (you need to organize):
- "Can we meet Monday 2-4 PM to discuss?" → true (you schedule & invite)
- "Please set up a call with the team" → true (you organize)
- "Let's schedule time to review" → true (you coordinate)
"""
        
        try:
            response = self.llm.call_with_json_parsing(prompt)
            if response and isinstance(response, dict):
                print(f"   🔍 LLM Response: {response}")  # Debug output
                return response
            else:
                print(f"   ⚠️ LLM returned non-dict response for meeting suggestion")
                return {"should_create_meeting": False, "reasoning": "Invalid LLM response"}
        except Exception as e:
            print(f"   ❌ Error getting LLM meeting suggestion: {e}")
            return {"should_create_meeting": False, "reasoning": f"Error: {str(e)}"}
    
    def _parse_llm_datetime(self, date_str: str, time_str: str) -> datetime:
        """Parse LLM-provided date and time into datetime object"""
        from datetime import datetime
        import re
        
        # Parse date (YYYY-MM-DD format expected from LLM)
        try:
            date_parts = date_str.split('-')
            year, month, day = int(date_parts[0]), int(date_parts[1]), int(date_parts[2])
        except:
            # Fallback to next business day
            today = datetime.now()
            days_ahead = 1 if today.weekday() < 4 else (7 - today.weekday())  # Next weekday
            target_date = today + timedelta(days=days_ahead)
            year, month, day = target_date.year, target_date.month, target_date.day
        
        # Parse time (HH:MM format expected from LLM) 
        try:
            time_parts = time_str.split(':')
            hour, minute = int(time_parts[0]), int(time_parts[1])
        except:
            hour, minute = 14, 0  # Default to 2 PM
        
        return datetime(year, month, day, hour, minute)
    
    def _adjust_for_office_hours(self, event_datetime: datetime, meeting_suggestion: dict) -> datetime:
        """Adjust meeting time if outside office hours"""
        
        # Office hours: 9 AM - 6 PM, Monday-Friday
        if event_datetime.weekday() >= 5:  # Weekend
            # Move to next Monday
            days_ahead = 7 - event_datetime.weekday() + 1
            event_datetime = event_datetime + timedelta(days=days_ahead)
            event_datetime = event_datetime.replace(hour=14, minute=0)  # 2 PM default
            print(f"   ⏰ Moved weekend meeting to next Monday 2:00 PM")
        
        elif event_datetime.hour < 9:  # Before office hours
            event_datetime = event_datetime.replace(hour=9, minute=0)
            print(f"   ⏰ Moved early meeting to 9:00 AM")
        
        elif event_datetime.hour >= 18:  # After office hours
            # Move to next business day
            event_datetime = event_datetime + timedelta(days=1)
            if event_datetime.weekday() >= 5:  # If that's weekend, move to Monday
                days_ahead = 7 - event_datetime.weekday() + 1
                event_datetime = event_datetime + timedelta(days=days_ahead)
            event_datetime = event_datetime.replace(hour=9, minute=0)
            print(f"   ⏰ Moved after-hours meeting to next business day 9:00 AM")
        
        return event_datetime
    
    def _calculate_meeting_duration(self, start_time: str, end_time: str) -> float:
        """Calculate meeting duration in hours"""
        try:
            start_parts = start_time.split(':')
            end_parts = end_time.split(':')
            
            start_hour = int(start_parts[0]) + int(start_parts[1]) / 60
            end_hour = int(end_parts[0]) + int(end_parts[1]) / 60
            
            duration = end_hour - start_hour
            return max(duration, 1.0)  # Minimum 1 hour
        except:
            return 2.0  # Default 2 hours
    
    def _get_next_business_day_time(self, hour: int, minute: int) -> datetime:
        """Get next business day at specified time"""
        today = datetime.now()
        days_ahead = 1
        
        # If today is Friday, go to Monday
        if today.weekday() == 4:  # Friday
            days_ahead = 3
        elif today.weekday() == 5:  # Saturday  
            days_ahead = 2
        elif today.weekday() == 6:  # Sunday
            days_ahead = 1
        
        next_day = today + timedelta(days=days_ahead)
        return next_day.replace(hour=hour, minute=minute, second=0, microsecond=0)

    def _create_priority_email_list_draft(self, analyzed_emails, current_user_email):
        """Create a draft email with prioritized email list"""
        from datetime import datetime
        import pytz
        
        try:
            # Get current date and time in Dubai timezone
            dubai_tz = pytz.timezone('Asia/Dubai')
            now = datetime.now(dubai_tz)
            date_str = now.strftime('%Y-%m-%d')
            time_str = now.strftime('%H:%M')
            
            # Sort emails by priority score
            sorted_emails = sorted(analyzed_emails, key=lambda x: x[1]['core_analysis'].priority_score, reverse=True)
            
            # Create email body
            subject = f"Emails with priority [{date_str}, {time_str}]"
            
            body_parts = []
            body_parts.append(f"Email Priority Summary - {date_str} at {time_str}")
            body_parts.append("=" * 50)
            body_parts.append("")
            
            # Group by priority levels
            critical_emails = [(e, a) for e, a in sorted_emails if a['core_analysis'].priority_score >= 85]
            urgent_emails = [(e, a) for e, a in sorted_emails if 70 <= a['core_analysis'].priority_score < 85]
            
            # Add critical emails
            if critical_emails:
                body_parts.append("🔴 CRITICAL PRIORITY (85+ Score)")
                body_parts.append("-" * 30)
                for i, (email, analysis) in enumerate(critical_emails, 1):
                    core = analysis['core_analysis']
                    body_parts.append(f"{i}. Subject: {email.subject}")
                    body_parts.append(f"   From: {email.sender} <{email.sender_email}>")
                    body_parts.append(f"   Date: {email.date.strftime('%Y-%m-%d %H:%M')}")
                    body_parts.append(f"   Priority Score: {core.priority_score:.1f}/100")
                    body_parts.append(f"   Type: {core.email_type}")
                    body_parts.append(f"   Action Required: {core.action_required}")
                    
                    # Add deadline info if available
                    if hasattr(core, 'deadline_info') and core.deadline_info:
                        body_parts.append(f"   ⏰ Deadline: {core.deadline_info}")
                    
                    # Add task breakdown
                    if hasattr(core, 'task_breakdown') and core.task_breakdown:
                        body_parts.append(f"   📋 Tasks:")
                        for task in core.task_breakdown:
                            body_parts.append(f"      • {task}")
                    body_parts.append("")
            
            # Add urgent emails
            if urgent_emails:
                body_parts.append("🟡 URGENT PRIORITY (70-84 Score)")
                body_parts.append("-" * 30)
                for i, (email, analysis) in enumerate(urgent_emails[:5], 1):
                    core = analysis['core_analysis']
                    body_parts.append(f"{i}. {email.subject} - Score: {core.priority_score:.1f}")
                    body_parts.append(f"   From: {email.sender} <{email.sender_email}>")
                    body_parts.append(f"   Date: {email.date.strftime('%Y-%m-%d %H:%M')}")
                    body_parts.append(f"   Action Required: {core.action_required}")
                    
                    # Add deadline info if available
                    if hasattr(core, 'deadline_info') and core.deadline_info:
                        body_parts.append(f"   ⏰ Deadline: {core.deadline_info}")
                    
                    # Add task breakdown for urgent emails
                    if hasattr(core, 'task_breakdown') and core.task_breakdown:
                        body_parts.append(f"   📋 Tasks:")
                        for task in core.task_breakdown[:3]:  # Limit to first 3 tasks for urgent
                            body_parts.append(f"      • {task}")
                        if len(core.task_breakdown) > 3:
                            body_parts.append(f"      • ... and {len(core.task_breakdown) - 3} more tasks")
                    body_parts.append("")
            
            # Add summary statistics
            body_parts.append("")
            body_parts.append("📊 SUMMARY")
            body_parts.append(f"Total emails analyzed: {len(analyzed_emails)}")
            body_parts.append(f"Critical priority: {len(critical_emails)}")
            body_parts.append(f"Urgent priority: {len(urgent_emails)}")
            body_parts.append("")
            body_parts.append("📋 TASK BREAKDOWN INCLUDED")
            body_parts.append("Each high-priority email includes specific action items and tasks")
            body_parts.append("to help you process them efficiently.")
            body_parts.append("")
            body_parts.append("Generated by LLM-Enhanced Email Agent")
            
            # Join all parts
            body = "\n".join(body_parts)
            
            # Create the draft
            draft_result = self.outlook.create_draft(
                to_email=current_user_email,  # Send to self
                subject=subject,
                body=body
            )
            
            if draft_result.get('success'):
                print(f"📋 Priority email list draft created successfully")
                print(f"   Subject: {subject}")
            else:
                print(f"❌ Failed to create priority list draft: {draft_result.get('error', 'Unknown error')}")
                
            return draft_result
            
        except Exception as e:
            print(f"❌ Error creating priority email list draft: {e}")
            return {'success': False, 'error': str(e)}
    
    def _store_prioritized_emails_in_session(self, analyzed_emails):
        """Store prioritized emails in session state for persistent display"""
        import streamlit as st
        from datetime import datetime
        
        try:
            # Initialize session state for prioritized emails
            if 'prioritized_emails' not in st.session_state:
                st.session_state.prioritized_emails = []
            
            # Sort emails by priority score (highest first)
            sorted_emails = sorted(analyzed_emails, key=lambda x: x.priority_score, reverse=True)
            
            # Store top emails with essential info for display
            prioritized_data = []
            for email in sorted_emails[:10]:  # Store top 10 emails
                email_data = {
                    'id': email.id,
                    'subject': email.subject,
                    'sender': email.sender,
                    'sender_email': email.sender_email,
                    'priority_score': email.priority_score,
                    'urgency_level': getattr(email, 'urgency_level', 'Medium'),
                    'email_type': getattr(email, 'email_type', 'Unknown'),
                    'action_required': getattr(email, 'action_required', 'Review'),
                    'received_time': email.received_datetime.strftime('%Y-%m-%d %H:%M') if hasattr(email, 'received_datetime') and email.received_datetime else 'Unknown',
                    'body_preview': email.body[:100] + '...' if len(email.body) > 100 else email.body,
                    'is_read': getattr(email, 'is_read', True),
                    'has_attachments': getattr(email, 'has_attachments', False)
                }
                prioritized_data.append(email_data)
            
            # Store in session state with timestamp
            st.session_state.prioritized_emails = prioritized_data
            st.session_state.prioritized_emails_timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            
            print(f"✅ Stored {len(prioritized_data)} prioritized emails in session state")
            
        except Exception as e:
            print(f"❌ Error storing prioritized emails in session: {e}")
    
    def get_prioritized_emails_from_session(self):
        """Get prioritized emails from session state for display"""
        import streamlit as st
        
        if 'prioritized_emails' not in st.session_state:
            return []
        
        return st.session_state.prioritized_emails
    
    def show_pending_calendar_confirmations(self):
        """Show pending calendar confirmations in sidebar after main processing with editable preview fields"""
        import streamlit as st
        from datetime import datetime, timedelta
        
        if 'pending_calendar_confirmations' not in st.session_state or not st.session_state.pending_calendar_confirmations:
            return
        
        st.sidebar.write("---")
        st.sidebar.write("📅 **Calendar Events to Review**")
        st.sidebar.info("💡 Edit details below, then click Create Event")
        
        confirmations_to_remove = []
        
        for confirmation_key, confirmation_data in st.session_state.pending_calendar_confirmations.items():
            if confirmation_key in st.session_state:
                # Decision already made, process it
                decision = st.session_state[confirmation_key]
                if decision:
                    # Create calendar event with user-edited details
                    try:
                        email = confirmation_data['email']
                        analysis = confirmation_data['analysis']
                        
                        # Check if user has edited the event details
                        edited_details_key = f"{confirmation_key}_edited_details"
                        if edited_details_key in st.session_state:
                            # Use edited details to create custom calendar event
                            edited_details = st.session_state[edited_details_key]
                            calendar_event = self._create_calendar_event_with_custom_details(email, edited_details)
                        else:
                            # Use original method
                            calendar_event = self._create_calendar_event_from_email(email, analysis)
                        
                        if calendar_event and calendar_event.get('success'):
                            st.sidebar.success(f"✅ Event created: {confirmation_data['subject'][:25]}...")
                            # Show created event details briefly
                            if 'event_id' in calendar_event:
                                st.sidebar.caption(f"🔗 Event ID: {calendar_event['event_id'][:8]}...")
                        else:
                            st.sidebar.error(f"❌ Failed to create: {confirmation_data['subject'][:25]}...")
                            st.sidebar.caption("💡 Check permissions & try again")
                    except Exception as e:
                        st.sidebar.error(f"❌ Error creating calendar event: {str(e)}")
                else:
                    st.sidebar.info(f"⏭️ Skipped calendar event for: {confirmation_data['subject'][:30]}...")
                
                confirmations_to_remove.append(confirmation_key)
            else:
                # Show enhanced confirmation dialog with editable preview fields
                with st.sidebar.expander(f"📧 {confirmation_data['subject'][:30]}...", expanded=True):
                    st.write(f"**From:** {confirmation_data['sender']}")
                    st.write(f"**Priority:** {confirmation_data['priority_score']:.1f}/100")
                    st.write(f"**Type:** {confirmation_data['email_type']}")
                    st.write(f"**Preview:** {confirmation_data['body_preview']}")
                    
                    # Get LLM meeting suggestion for preview
                    try:
                        email = confirmation_data['email']
                        meeting_suggestion = self._get_llm_meeting_suggestion(email)
                        
                        if meeting_suggestion and meeting_suggestion.get('should_create_meeting'):
                            st.write("---")
                            st.write("**📝 Edit Event Details:**")
                            
                            # Event Title
                            default_title = meeting_suggestion.get('title', f"Meeting: {email.subject}")
                            event_title = st.text_input(
                                "Event Title:",
                                value=default_title,
                                key=f"{confirmation_key}_title"
                            )
                            
                            # Date and Time
                            col1, col2 = st.columns(2)
                            with col1:
                                try:
                                    default_date = datetime.strptime(meeting_suggestion.get('date', datetime.now().strftime('%Y-%m-%d')), '%Y-%m-%d').date()
                                except:
                                    default_date = datetime.now().date()
                                
                                event_date = st.date_input(
                                    "Date:",
                                    value=default_date,
                                    key=f"{confirmation_key}_date"
                                )
                            
                            with col2:
                                # Time inputs
                                default_start = meeting_suggestion.get('start_time', '14:00')
                                default_end = meeting_suggestion.get('end_time', '15:00')
                                
                                try:
                                    start_hour, start_min = map(int, default_start.split(':'))
                                    end_hour, end_min = map(int, default_end.split(':'))
                                except:
                                    start_hour, start_min = 14, 0
                                    end_hour, end_min = 15, 0
                                
                                start_time = st.time_input(
                                    "Start Time:",
                                    value=datetime.now().replace(hour=start_hour, minute=start_min).time(),
                                    key=f"{confirmation_key}_start_time"
                                )
                                
                                end_time = st.time_input(
                                    "End Time:",
                                    value=datetime.now().replace(hour=end_hour, minute=end_min).time(),
                                    key=f"{confirmation_key}_end_time"
                                )
                            
                            # Location
                            default_location = meeting_suggestion.get('location', '')
                            event_location = st.text_input(
                                "Location:",
                                value=default_location,
                                key=f"{confirmation_key}_location"
                            )
                            
                            # Description/Notes
                            default_notes = meeting_suggestion.get('notes', '')
                            event_description = st.text_area(
                                "Description/Notes:",
                                value=default_notes,
                                height=100,
                                key=f"{confirmation_key}_description"
                            )
                            
                            # Participants (read-only for now, extracted from LLM)
                            participants = meeting_suggestion.get('participants', [])
                            if participants:
                                st.write("**👥 Detected Participants:**")
                                for participant in participants:
                                    st.write(f"   • {participant}")
                            
                            # Store edited details in session state
                            edited_details = {
                                'title': event_title,
                                'date': event_date.strftime('%Y-%m-%d'),
                                'start_time': start_time.strftime('%H:%M'),
                                'end_time': end_time.strftime('%H:%M'),
                                'location': event_location,
                                'description': event_description,
                                'participants': participants,
                                'purpose': meeting_suggestion.get('purpose', ''),
                                'original_email_subject': email.subject,
                                'original_email_body': email.body
                            }
                            st.session_state[f"{confirmation_key}_edited_details"] = edited_details
                            
                        else:
                            st.warning("⚠️ Could not extract meeting details from email")
                            
                    except Exception as e:
                        st.error(f"❌ Error extracting meeting details: {str(e)}")
                    
                    st.write("---")
                    col1, col2 = st.columns([1.2, 0.8])
                    with col1:
                        if st.button("🗓️ Create Calendar Event", key=f"{confirmation_key}_yes", use_container_width=True):
                            st.session_state[confirmation_key] = True
                            st.rerun()
                    with col2:
                        if st.button("⏭️ Skip", key=f"{confirmation_key}_no", use_container_width=True):
                            st.session_state[confirmation_key] = False
                            st.rerun()
        
        # Clean up processed confirmations
        for key in confirmations_to_remove:
            del st.session_state.pending_calendar_confirmations[key]
            # Also clean up edited details
            edited_details_key = f"{key}_edited_details"
            if edited_details_key in st.session_state:
                del st.session_state[edited_details_key]
    
    def _queue_calendar_event_for_streamlit(self, email: OutlookEmailData, analysis: Dict, meeting_suggestion: Dict):
        """Queue calendar event for interactive confirmation in Streamlit"""
        import streamlit as st
        
        try:
            # Initialize pending confirmations if not exists
            if 'pending_calendar_confirmations' not in st.session_state:
                st.session_state.pending_calendar_confirmations = {}
            
            # Create unique key for this confirmation
            confirmation_key = f"calendar_event_{email.id}_{hash(email.subject)}"
            
            # Store the confirmation data
            confirmation_data = {
                'email': email,
                'analysis': analysis,
                'meeting_suggestion': meeting_suggestion,
                'subject': email.subject,
                'sender': email.sender,
                'priority_score': analysis.priority_score if hasattr(analysis, 'priority_score') else email.priority_score,
                'email_type': analysis.email_type if hasattr(analysis, 'email_type') else email.email_type,
                'body_preview': email.body_preview or email.body[:100]
            }
            
            # Queue for confirmation
            st.session_state.pending_calendar_confirmations[confirmation_key] = confirmation_data
            
        except Exception as e:
            print(f"   ❌ Error queuing calendar event for confirmation: {e}")

class WritingStyleAnalyzer:
    """Analyzes user's writing style from sent emails (inspired by inbox-zero)"""
    
    def __init__(self, llm_service: UnifiedLLMService):
        self.llm_service = llm_service
    
    def analyze_sent_emails(self, outlook: OutlookService, max_emails: int = 20) -> WritingStyle:
        """Analyze user's writing style from sent emails"""
        try:
            # Get sent emails for analysis
            sent_emails = self._get_sent_emails(outlook, max_emails)
            
            if not sent_emails:
                return self._default_writing_style()
            
            # Analyze patterns
            tone_analysis = self._analyze_tone(sent_emails)
            formality_analysis = self._analyze_formality(sent_emails)
            structure_analysis = self._analyze_structure(sent_emails)
            
            return WritingStyle(
                tone=tone_analysis['primary_tone'],
                formality_level=formality_analysis['formality_score'],
                greeting_style=structure_analysis['common_greeting'],
                closing_style=structure_analysis['common_closing'],
                average_length=structure_analysis['avg_length'],
                use_contractions=structure_analysis['uses_contractions'],
                punctuation_style=structure_analysis['punctuation_style'],
                signature=structure_analysis['signature_pattern'],
                common_phrases=structure_analysis['common_phrases'],
                response_time_preference="same_day"  # Default
            )
            
        except Exception as e:
            print(f"Error analyzing writing style: {e}")
            return self._default_writing_style()
    
    def _get_sent_emails(self, outlook: OutlookService, max_emails: int) -> List[str]:
        """Get user's sent emails for analysis"""
        try:
            # Get sent items folder
            endpoint = f"/me/mailFolders/sentitems/messages?$select=subject,body,toRecipients&$top={max_emails}&$orderby=sentDateTime desc"
            result = outlook._make_graph_request(endpoint)
            
            emails = []
            for msg in result.get('value', []):
                body_content = msg.get('body', {}).get('content', '')
                if len(body_content.strip()) > 50:  # Skip very short emails
                    emails.append(self._clean_email_body(body_content))
            
            return emails
            
        except Exception as e:
            print(f"Error fetching sent emails: {e}")
            return []
    
    def _clean_email_body(self, body: str) -> str:
        """Clean email body for analysis"""
        import re
        # Remove HTML tags
        clean = re.sub('<.*?>', '', body)
        # Remove excessive whitespace
        clean = re.sub(r'\s+', ' ', clean)
        # Remove signature blocks (common patterns)
        clean = re.sub(r'--\s*.*?$', '', clean, flags=re.MULTILINE | re.DOTALL)
        return clean.strip()
    
    def _analyze_tone(self, emails: List[str]) -> Dict[str, Any]:
        """Analyze tone patterns"""
        prompt = f"""
        Analyze the tone of these email samples and determine the primary writing tone:
        
        EMAIL SAMPLES:
        {chr(10).join(emails[:5])}  # Limit to first 5 emails
        
        Return JSON with:
        {{
            "primary_tone": "professional|friendly|formal|casual",
            "tone_confidence": 0.8,
            "tone_indicators": ["uses please and thank you", "direct but polite"]
        }}
        """
        
        try:
            response = self.llm_service.call_with_json_parsing(prompt)
            return response if response else {"primary_tone": "professional", "tone_confidence": 0.5}
        except:
            return {"primary_tone": "professional", "tone_confidence": 0.5}
    
    def _analyze_formality(self, emails: List[str]) -> Dict[str, Any]:
        """Analyze formality level"""
        import re
        
        total_words = 0
        formal_indicators = 0
        casual_indicators = 0
        
        for email in emails:
            words = email.split()
            total_words += len(words)
            
            # Formal indicators
            formal_patterns = ['Dear', 'Sincerely', 'Best regards', 'Please find', 'I would like to', 'Thank you for your']
            casual_patterns = ["Hi", "Hey", "Thanks", "Let me know", "No problem", "Sure thing"]
            
            for pattern in formal_patterns:
                if pattern.lower() in email.lower():
                    formal_indicators += 1
            
            for pattern in casual_patterns:
                if pattern.lower() in email.lower():
                    casual_indicators += 1
        
        # Calculate formality score (0-1)
        if formal_indicators + casual_indicators == 0:
            formality_score = 0.6  # Default professional
        else:
            formality_score = formal_indicators / (formal_indicators + casual_indicators)
        
        return {
            "formality_score": formality_score,
            "formal_indicators": formal_indicators,
            "casual_indicators": casual_indicators
        }
    
    def _analyze_structure(self, emails: List[str]) -> Dict[str, Any]:
        """Analyze structural patterns"""
        import re
        from collections import Counter
        
        greetings = []
        closings = []
        word_counts = []
        contractions_count = 0
        total_sentences = 0
        
        # Common greeting patterns
        greeting_patterns = [
            r'^(Hi|Hello|Dear|Hey)\s+(\w+)',
            r'^(Good morning|Good afternoon)',
            r'^(\w+),?\s*$'  # Name only
        ]
        
        # Common closing patterns  
        closing_patterns = [
            r'(Best regards?|Sincerely|Thanks?|Cheers|Best),?\s*$',
            r'(Looking forward|Thank you|Speak soon)',
            r'(Best|Regards|Thanks again)'
        ]
        
        for email in emails:
            lines = email.split('\\n')
            words = email.split()
            word_counts.append(len(words))
            
            # Count sentences for punctuation analysis
            sentences = re.split(r'[.!?]+', email)
            total_sentences += len([s for s in sentences if s.strip()])
            
            # Find greetings (first few lines)
            for line in lines[:3]:
                line = line.strip()
                if line:
                    for pattern in greeting_patterns:
                        match = re.search(pattern, line, re.IGNORECASE)
                        if match:
                            greetings.append(match.group(1).title())
                            break
                    break
            
            # Find closings (last few lines)
            for line in reversed(lines[-3:]):
                line = line.strip()
                if line:
                    for pattern in closing_patterns:
                        match = re.search(pattern, line, re.IGNORECASE)
                        if match:
                            closings.append(match.group(1).title())
                            break
                    break
            
            # Count contractions
            contractions = ["don't", "won't", "can't", "I'll", "we'll", "it's", "that's"]
            for contraction in contractions:
                contractions_count += email.lower().count(contraction)
        
        # Find most common patterns
        common_greeting = Counter(greetings).most_common(1)
        common_greeting = common_greeting[0][0] if common_greeting else "Hi"
        
        common_closing = Counter(closings).most_common(1)
        common_closing = common_closing[0][0] if common_closing else "Best regards"
        
        avg_length = int(sum(word_counts) / len(word_counts)) if word_counts else 100
        uses_contractions = contractions_count > len(emails) * 0.3  # 30% threshold
        
        # Extract common phrases (simplified)
        common_phrases = ["Thank you", "Please let me know", "Looking forward", "Best regards"]
        
        return {
            "common_greeting": common_greeting,
            "common_closing": common_closing,
            "avg_length": avg_length,
            "uses_contractions": uses_contractions,
            "punctuation_style": "standard",  # Simplified
            "signature_pattern": "",  # To be implemented
            "common_phrases": common_phrases
        }
    
    def _default_writing_style(self) -> WritingStyle:
        """Return default professional writing style"""
        return WritingStyle(
            tone="professional",
            formality_level=0.7,
            greeting_style="Hi",
            closing_style="Best regards",
            average_length=120,
            use_contractions=False,
            punctuation_style="standard",
            signature="",
            common_phrases=["Thank you", "Please let me know", "Best regards"],
            response_time_preference="same_day"
        )

@dataclass
class AdvancedEmailCategory:
    """Advanced email category classification"""
    primary_category: str  # urgent, important, interesting, promotional, social, etc.
    subcategory: str  # meeting, task, newsletter, receipt, personal, etc.
    confidence: float  # 0-1 confidence score
    reasoning: str  # Why this categorization was chosen
    suggested_action: str  # reply, archive, delete, follow_up, etc.
    priority_score: float  # 0-100 priority score
    auto_process: bool  # Whether this can be auto-processed

@dataclass
class ColdEmailAnalysis:
    """Cold email detection result"""
    is_cold_email: bool
    confidence: float  # 0-1 confidence score
    indicators: List[str]  # Why it's considered cold/not cold
    sender_reputation: str  # unknown, verified, trusted, suspicious
    recommended_action: str  # block, archive, review, respond
    risk_level: str  # low, medium, high

class SmartEmailCategorizer:
    """Advanced AI-powered email categorization (inspired by inbox-zero 2024)"""
    
    def __init__(self, llm_service: UnifiedLLMService):
        self.llm_service = llm_service
        self.category_patterns = self._load_category_patterns()
        self.sender_database = {}  # Cache for sender classifications
    
    def _load_category_patterns(self) -> Dict[str, Dict]:
        """Load smart categorization patterns based on latest inbox-zero research"""
        return {
            "urgent": {
                "keywords": ["urgent", "asap", "emergency", "critical", "deadline", "overdue"],
                "sender_patterns": ["ceo", "manager", "director", "urgent", "emergency"],
                "subject_patterns": ["urgent:", "asap:", "critical:", "deadline", "action required"],
                "priority_boost": 30
            },
            "important": {
                "keywords": ["important", "meeting", "project", "review", "approval", "decision"],
                "sender_patterns": ["boss", "client", "customer", "team", "project"],
                "subject_patterns": ["meeting:", "project:", "review:", "approval needed"],
                "priority_boost": 20
            },
            "interesting": {
                "keywords": ["update", "news", "article", "research", "insights", "announcement"],
                "sender_patterns": ["newsletter", "blog", "news", "updates", "insights"],
                "subject_patterns": ["update:", "news:", "weekly", "monthly", "newsletter"],
                "priority_boost": 5
            },
            "promotional": {
                "keywords": ["sale", "discount", "offer", "deal", "promotion", "marketing"],
                "sender_patterns": ["marketing", "sales", "offers", "deals", "promo"],
                "subject_patterns": ["sale", "discount", "offer", "deal", "% off"],
                "priority_boost": -10
            },
            "social": {
                "keywords": ["linkedin", "facebook", "twitter", "social", "network", "connect"],
                "sender_patterns": ["linkedin", "facebook", "twitter", "social", "network"],
                "subject_patterns": ["linkedin", "connection", "follow", "like", "comment"],
                "priority_boost": 0
            },
            "receipts": {
                "keywords": ["receipt", "invoice", "payment", "order", "confirmation", "billing"],
                "sender_patterns": ["billing", "payments", "receipts", "orders", "finance"],
                "subject_patterns": ["receipt", "invoice", "payment", "order #", "confirmation"],
                "priority_boost": 5
            }
        }
    
    def categorize_email(self, email: OutlookEmailData) -> AdvancedEmailCategory:
        """Advanced AI-powered email categorization"""
        try:
            # Prepare email content for analysis
            email_text = f"{email.subject} {email.body} {email.sender}".lower()
            
            # Pattern-based pre-classification (fast)
            pattern_category = self._pattern_based_classification(email, email_text)
            
            # AI-enhanced classification (more accurate)
            ai_category = self._ai_enhanced_classification(email, pattern_category)
            
            return ai_category
            
        except Exception as e:
            print(f"Error in email categorization: {e}")
            return self._fallback_categorization(email)
    
    def _pattern_based_classification(self, email: OutlookEmailData, email_text: str) -> str:
        """Fast pattern-based classification"""
        category_scores = {}
        
        for category, patterns in self.category_patterns.items():
            score = 0
            
            # Check keywords
            for keyword in patterns["keywords"]:
                if keyword in email_text:
                    score += 2
            
            # Check sender patterns
            for pattern in patterns["sender_patterns"]:
                if pattern in email.sender.lower() or pattern in email.sender_email.lower():
                    score += 3
            
            # Check subject patterns
            for pattern in patterns["subject_patterns"]:
                if pattern in email.subject.lower():
                    score += 4
            
            category_scores[category] = score
        
        # Return category with highest score
        return max(category_scores, key=category_scores.get) if max(category_scores.values()) > 0 else "general"
    
    def _ai_enhanced_classification(self, email: OutlookEmailData, pattern_category: str) -> EmailCategory:
        """AI-enhanced classification for better accuracy"""
        
        prompt = f"""
        Analyze this email and provide advanced categorization following modern inbox-zero principles:
        
        EMAIL DETAILS:
        From: {email.sender} <{email.sender_email}>
        Subject: {email.subject}
        Body: {email.body[:500]}...
        
        Pattern-based suggestion: {pattern_category}
        
        Categorize using these modern email categories:
        - urgent: Requires immediate attention (deadline, emergency, critical)
        - important: Significant but not urgent (meetings, projects, decisions)
        - interesting: Informational, educational (newsletters, articles, updates)
        - promotional: Marketing, sales, offers
        - social: Social media notifications, networking
        - receipts: Invoices, confirmations, billing
        - personal: Personal communications
        - automated: System notifications, reports
        
        Return JSON:
        {{
            "primary_category": "urgent|important|interesting|promotional|social|receipts|personal|automated",
            "subcategory": "meeting|task|newsletter|receipt|notification|etc",
            "confidence": 0.85,
            "reasoning": "Brief explanation of categorization",
            "suggested_action": "reply|archive|delete|follow_up|calendar|review",
            "priority_score": 75.0,
            "auto_process": true
        }}
        """
        
        try:
            response = self.llm_service.call_with_json_parsing(prompt)
            
            if response and 'primary_category' in response:
                return AdvancedEmailCategory(
                    primary_category=response['primary_category'],
                    subcategory=response.get('subcategory', 'general'),
                    confidence=response.get('confidence', 0.7),
                    reasoning=response.get('reasoning', 'AI classification'),
                    suggested_action=response.get('suggested_action', 'review'),
                    priority_score=response.get('priority_score', 50.0),
                    auto_process=response.get('auto_process', False)
                )
            else:
                return self._fallback_categorization(email)
                
        except Exception as e:
            print(f"AI categorization failed: {e}")
            return self._fallback_categorization(email)
    
    def _fallback_categorization(self, email: OutlookEmailData) -> AdvancedEmailCategory:
        """Fallback categorization when AI fails"""
        email_text = f"{email.subject} {email.body}".lower()
        
        # Simple heuristics
        if any(word in email_text for word in ["urgent", "asap", "critical", "deadline"]):
            category = "urgent"
            priority = 80.0
        elif any(word in email_text for word in ["meeting", "project", "review"]):
            category = "important"
            priority = 65.0
        elif any(word in email_text for word in ["newsletter", "update", "article"]):
            category = "interesting"
            priority = 40.0
        elif any(word in email_text for word in ["sale", "offer", "discount"]):
            category = "promotional"
            priority = 20.0
        else:
            category = "general"
            priority = 50.0
        
        return AdvancedEmailCategory(
            primary_category=category,
            subcategory="general",
            confidence=0.6,
            reasoning="Pattern-based fallback classification",
            suggested_action="review",
            priority_score=priority,
            auto_process=False
        )

class ColdEmailDetector:
    """Advanced cold email detection system (inspired by inbox-zero 2024)"""
    
    def __init__(self, llm_service: UnifiedLLMService, outlook_service: OutlookService):
        self.llm_service = llm_service
        self.outlook_service = outlook_service
        self.known_senders = set()  # Cache of known legitimate senders
        self.cold_indicators = self._load_cold_email_indicators()
    
    def _load_cold_email_indicators(self) -> Dict[str, List[str]]:
        """Load cold email detection patterns based on latest research"""
        return {
            "subject_patterns": [
                "re:", "fwd:", "quick question", "checking in", "following up",
                "partnership opportunity", "business proposal", "collaboration",
                "guest post", "link building", "seo services", "marketing services"
            ],
            "content_patterns": [
                "i hope this email finds you well", "i came across your",
                "i noticed your website", "quick question for you",
                "would love to connect", "business opportunity",
                "increase your revenue", "boost your sales",
                "improve your rankings", "guest posting opportunity"
            ],
            "sender_patterns": [
                "marketing", "sales", "business", "seo", "agency",
                "consultant", "outreach", "partnerships", "growth"
            ],
            "suspicious_patterns": [
                "urgent response needed", "act now", "limited time",
                "free trial", "no obligation", "risk free",
                "congratulations", "you've been selected"
            ]
        }
    
    def detect_cold_email(self, email: OutlookEmailData) -> ColdEmailAnalysis:
        """Comprehensive cold email detection"""
        try:
            # Step 1: Check if sender is known
            sender_reputation = self._check_sender_reputation(email)
            
            # Step 2: Pattern-based detection
            pattern_analysis = self._pattern_based_detection(email)
            
            # Step 3: AI-enhanced detection
            ai_analysis = self._ai_enhanced_detection(email, pattern_analysis)
            
            # Step 4: Combine results
            final_analysis = self._combine_detection_results(
                sender_reputation, pattern_analysis, ai_analysis
            )
            
            return final_analysis
            
        except Exception as e:
            print(f"Error in cold email detection: {e}")
            return self._fallback_cold_detection(email)
    
    def _check_sender_reputation(self, email: OutlookEmailData) -> Dict[str, Any]:
        """Check sender reputation and history"""
        sender_email = email.sender_email.lower()
        
        # Check if we've communicated before
        has_history = self._check_communication_history(sender_email)
        
        # Check domain reputation
        domain = sender_email.split('@')[-1] if '@' in sender_email else ""
        domain_reputation = self._check_domain_reputation(domain)
        
        return {
            "has_history": has_history,
            "domain_reputation": domain_reputation,
            "is_internal": domain in ["mbzuai.ac.ae", "gmail.com"],  # Add your trusted domains
            "sender_email": sender_email
        }
    
    def _check_communication_history(self, sender_email: str) -> bool:
        """Check if we've communicated with this sender before"""
        try:
            # Search recent emails for this sender
            endpoint = f"/me/messages?$select=sender&$top=100&$orderby=receivedDateTime desc"
            result = self.outlook_service._make_graph_request(endpoint)
            
            for msg in result.get('value', []):
                msg_sender = msg.get('sender', {}).get('emailAddress', {}).get('address', '').lower()
                if msg_sender == sender_email:
                    return True
            
            return False
            
        except Exception:
            return False
    
    def _check_domain_reputation(self, domain: str) -> str:
        """Check domain reputation"""
        trusted_domains = [
            "gmail.com", "outlook.com", "yahoo.com", "mbzuai.ac.ae",
            "microsoft.com", "google.com", "apple.com", "amazon.com"
        ]
        
        suspicious_domains = [
            "tempmail", "10minutemail", "guerrillamail", "mailinator"
        ]
        
        if domain in trusted_domains:
            return "trusted"
        elif any(suspicious in domain for suspicious in suspicious_domains):
            return "suspicious"
        else:
            return "unknown"
    
    def _pattern_based_detection(self, email: OutlookEmailData) -> Dict[str, Any]:
        """Pattern-based cold email detection"""
        email_text = f"{email.subject} {email.body}".lower()
        
        indicators = []
        confidence = 0.0
        
        # Check subject patterns
        subject_matches = 0
        for pattern in self.cold_indicators["subject_patterns"]:
            if pattern in email.subject.lower():
                subject_matches += 1
                indicators.append(f"Subject contains: '{pattern}'")
        
        # Check content patterns
        content_matches = 0
        for pattern in self.cold_indicators["content_patterns"]:
            if pattern in email_text:
                content_matches += 1
                indicators.append(f"Content contains: '{pattern}'")
        
        # Check sender patterns
        sender_matches = 0
        for pattern in self.cold_indicators["sender_patterns"]:
            if pattern in email.sender.lower() or pattern in email.sender_email.lower():
                sender_matches += 1
                indicators.append(f"Sender contains: '{pattern}'")
        
        # Check suspicious patterns
        suspicious_matches = 0
        for pattern in self.cold_indicators["suspicious_patterns"]:
            if pattern in email_text:
                suspicious_matches += 1
                indicators.append(f"Suspicious pattern: '{pattern}'")
        
        # Calculate confidence
        total_matches = subject_matches + content_matches + sender_matches + suspicious_matches
        confidence = min(total_matches * 0.2, 1.0)  # Each match adds 20% confidence, max 100%
        
        return {
            "indicators": indicators,
            "confidence": confidence,
            "subject_matches": subject_matches,
            "content_matches": content_matches,
            "sender_matches": sender_matches,
            "suspicious_matches": suspicious_matches
        }
    
    def _ai_enhanced_detection(self, email: OutlookEmailData, pattern_analysis: Dict) -> Dict[str, Any]:
        """AI-enhanced cold email detection"""
        
        prompt = f"""
        Analyze this email to detect if it's a cold email (unsolicited outreach):
        
        EMAIL DETAILS:
        From: {email.sender} <{email.sender_email}>
        Subject: {email.subject}
        Body: {email.body[:800]}...
        
        Pattern analysis found {len(pattern_analysis['indicators'])} cold email indicators.
        
        Consider these factors:
        1. Is this unsolicited outreach?
        2. Is the sender trying to sell something?
        3. Does it feel personalized or mass-sent?
        4. Is there clear value for the recipient?
        5. Does it follow cold email templates?
        
        Return JSON:
        {{
            "is_cold_email": true,
            "confidence": 0.85,
            "reasoning": "Clear cold outreach with sales intent",
            "recommended_action": "archive|block|review|respond",
            "personalization_level": "none|low|medium|high",
            "sales_intent": true
        }}
        """
        
        try:
            response = self.llm_service.call_with_json_parsing(prompt)
            return response if response else {}
        except Exception:
            return {}
    
    def _combine_detection_results(self, sender_rep: Dict, pattern_analysis: Dict, ai_analysis: Dict) -> ColdEmailAnalysis:
        """Combine all detection results into final analysis"""
        
        # Base decision on sender reputation
        if sender_rep["has_history"]:
            is_cold = False
            confidence = 0.1
        elif sender_rep["is_internal"]:
            is_cold = False
            confidence = 0.2
        else:
            # Use pattern + AI analysis
            pattern_confidence = pattern_analysis.get("confidence", 0.0)
            ai_confidence = ai_analysis.get("confidence", 0.0) if ai_analysis else 0.0
            ai_is_cold = ai_analysis.get("is_cold_email", False) if ai_analysis else False
            
            # Weight AI analysis more heavily
            combined_confidence = (pattern_confidence * 0.3) + (ai_confidence * 0.7)
            is_cold = ai_is_cold or combined_confidence > 0.6
            confidence = combined_confidence
        
        # Determine recommended action
        if is_cold and confidence > 0.8:
            action = "block"
            risk = "high"
        elif is_cold and confidence > 0.5:
            action = "archive"
            risk = "medium"
        elif is_cold:
            action = "review"
            risk = "low"
        else:
            action = "respond"
            risk = "low"
        
        return ColdEmailAnalysis(
            is_cold_email=is_cold,
            confidence=confidence,
            indicators=pattern_analysis.get("indicators", []),
            sender_reputation=sender_rep.get("domain_reputation", "unknown"),
            recommended_action=action,
            risk_level=risk
        )
    
    def _fallback_cold_detection(self, email: OutlookEmailData) -> ColdEmailAnalysis:
        """Fallback when detection fails"""
        return ColdEmailAnalysis(
            is_cold_email=False,
            confidence=0.5,
            indicators=["Detection system error"],
            sender_reputation="unknown",
            recommended_action="review",
            risk_level="low"
        )

@dataclass 
class EmailRule:
    """Email automation rule"""
    name: str
    description: str
    conditions: Dict[str, Any]  # Conditions to match
    actions: List[str]  # Actions to take
    enabled: bool = True
    priority: int = 0  # Higher priority rules run first

class EmailAutomationEngine:
    """Rule-based email automation system (inspired by inbox-zero 2024)"""
    
    def __init__(self, llm_service: UnifiedLLMService, outlook_service: OutlookService):
        self.llm_service = llm_service
        self.outlook = outlook_service  # Use consistent naming
        self.rules = self._load_default_rules()
        self.action_history = []  # Track actions for rollback capability
    
    def _load_default_rules(self) -> List[EmailRule]:
        """Load default automation rules"""
        return [
            EmailRule(
                name="Auto-archive newsletters",
                description="Automatically archive newsletters and promotional emails",
                conditions={
                    "category": ["promotional", "interesting"],
                    "keywords": ["newsletter", "unsubscribe", "weekly update"],
                    "sender_patterns": ["newsletter", "marketing", "promo"]
                },
                actions=["archive", "mark_read"],
                priority=1
            ),
            EmailRule(
                name="Block cold emails",
                description="Archive cold emails with high confidence",
                conditions={
                    "cold_email": True,
                    "cold_confidence": 0.8
                },
                actions=["archive", "label:Cold Email"],
                priority=2
            ),
            EmailRule(
                name="High priority notifications",
                description="Send notifications for urgent emails",
                conditions={
                    "category": ["urgent"],
                    "priority_score": 80
                },
                actions=["notify", "flag"],
                priority=3
            ),
            EmailRule(
                name="Auto-create calendar events",
                description="Create calendar events for meeting requests",
                conditions={
                    "keywords": ["meeting", "appointment", "call", "zoom"],
                    "category": ["important"],
                    "has_date": True
                },
                actions=["create_calendar_event", "reply_confirmation"],
                priority=2
            )
        ]
    
    def process_email_with_rules(self, email: OutlookEmailData, analysis_results: Dict) -> List[str]:
        """Process email through automation rules"""
        applied_actions = []
        
        try:
            # Sort rules by priority
            sorted_rules = sorted(self.rules, key=lambda r: r.priority, reverse=True)
            
            for rule in sorted_rules:
                if not rule.enabled:
                    continue
                
                # Check if rule conditions are met
                if self._rule_matches(rule, email, analysis_results):
                    print(f"📋 Applying rule: {rule.name}")
                    
                    # Execute rule actions
                    for action in rule.actions:
                        success = self._execute_action(action, email, analysis_results)
                        if success:
                            applied_actions.append(f"{rule.name}: {action}")
            
            return applied_actions
            
        except Exception as e:
            print(f"Error processing email rules: {e}")
            return []
    
    def _rule_matches(self, rule: EmailRule, email: OutlookEmailData, analysis: Dict) -> bool:
        """Check if rule conditions match the email"""
        conditions = rule.conditions
        
        # Check category conditions
        if "category" in conditions:
            smart_category = analysis.get('smart_category')
            if smart_category and smart_category.primary_category not in conditions["category"]:
                return False
        
        # Check cold email conditions
        if "cold_email" in conditions:
            cold_analysis = analysis.get('cold_analysis')
            if not cold_analysis or cold_analysis.is_cold_email != conditions["cold_email"]:
                return False
        
        # Check cold email confidence
        if "cold_confidence" in conditions:
            cold_analysis = analysis.get('cold_analysis')
            if not cold_analysis or not hasattr(cold_analysis, 'confidence') or cold_analysis.confidence < conditions["cold_confidence"]:
                return False
        
        # Check priority score
        if "priority_score" in conditions:
            core_analysis = analysis.get('core_analysis')
            if not core_analysis or core_analysis.priority_score < conditions["priority_score"]:
                return False
        
        # Check keywords
        if "keywords" in conditions:
            email_text = f"{email.subject} {email.body}".lower()
            if not any(keyword in email_text for keyword in conditions["keywords"]):
                return False
        
        # Check sender patterns
        if "sender_patterns" in conditions:
            sender_text = f"{email.sender} {email.sender_email}".lower()
            if not any(pattern in sender_text for pattern in conditions["sender_patterns"]):
                return False
        
        # Check if email has date/time information
        if "has_date" in conditions and conditions["has_date"]:
            email_text = f"{email.subject} {email.body}".lower()
            date_keywords = ["monday", "tuesday", "wednesday", "thursday", "friday", "saturday", "sunday",
                           "today", "tomorrow", "next week", "this week", "am", "pm"]
            if not any(keyword in email_text for keyword in date_keywords):
                return False
        
        return True
    
    def _execute_action(self, action: str, email: OutlookEmailData, analysis: Dict) -> bool:
        """Execute a rule action with real Microsoft Graph API calls"""
        try:
            from datetime import datetime
            
            # Track action for rollback capability
            action_record = {
                'action': action,
                'email_id': email.id,
                'email_subject': email.subject,
                'timestamp': datetime.now().isoformat(),
                'original_state': {},
                'success': False
            }
            success = False
            
            if action == "archive":
                # Archive email using Graph API
                try:
                    # Move to Archive folder
                    archive_folder_id = self.outlook.get_folder_id("Archive")
                    if archive_folder_id:
                        self.outlook.move_email(email.id, archive_folder_id)
                        print(f"   📁 Archived: {email.subject[:30]}...")
                        success = True
                    else:
                        print(f"   ❌ Archive folder not found, skipping archive for: {email.subject[:30]}...")
                        success = False
                except Exception as e:
                    print(f"   ❌ Failed to archive email: {e}")
                    success = False
            
            elif action == "mark_read":
                # Mark as read using Graph API
                try:
                    self.outlook.mark_email_as_read(email.id)
                    print(f"   📖 Marked as read: {email.subject[:30]}...")
                    success = True
                except Exception as e:
                    print(f"   ❌ Failed to mark as read: {e}")
                    success = False
            
            elif action.startswith("label:"):
                # Add category using Graph API
                label = action.split(":", 1)[1]
                try:
                    # Add category to email
                    self.outlook.add_category_to_email(email.id, label)
                    print(f"   🏷️ Added category '{label}': {email.subject[:30]}...")
                    success = True
                except Exception as e:
                    print(f"   ❌ Failed to add category: {e}")
                    success = False
            
            elif action == "notify":
                # Send notification (implemented as console output + potential system notification)
                print(f"   🔔 HIGH PRIORITY notification: {email.subject[:30]}...")
                # Could be enhanced with system notifications or webhooks
                success = True
            
            elif action == "flag":
                # Flag email using Graph API
                try:
                    self.outlook.flag_email(email.id)
                    print(f"   🚩 Flagged: {email.subject[:30]}...")
                    success = True
                except Exception as e:
                    print(f"   ❌ Failed to flag email: {e}")
                    success = False
            
            elif action == "create_calendar_event":
                # Create calendar event using Graph API
                try:
                    # Extract meeting details from email
                    event_details = self._extract_event_details(email, analysis)
                    if event_details and event_details.get('start'):
                        # Convert datetime to ISO format
                        start_time = event_details['start'].isoformat()
                        end_time = event_details['end'].isoformat()
                        
                        result = self.outlook.create_calendar_event(
                            subject=event_details['subject'],
                            start_time=start_time,
                            end_time=end_time,
                            description=event_details['body'],
                            location=event_details.get('location', ''),
                            attendees=event_details.get('attendees', [])
                        )
                        
                        if result.get('success'):
                            print(f"   📅 Created calendar event from: {email.subject[:30]}...")
                            success = True
                        else:
                            print(f"   ❌ Failed to create calendar event: {result.get('message', 'Unknown error')}")
                            success = False
                    else:
                        print(f"   ❌ Could not extract event details from: {email.subject[:30]}...")
                        success = False
                except Exception as e:
                    print(f"   ❌ Failed to create calendar event: {e}")
                    success = False
            
            elif action == "reply_confirmation":
                # Send auto-reply confirmation using Graph API
                try:
                    confirmation_message = f"Thank you for your email. I have received your message regarding '{email.subject}' and will respond accordingly."
                    self.outlook.send_reply(email.id, confirmation_message)
                    print(f"   ✉️ Sent confirmation reply to: {email.sender}")
                    success = True
                except Exception as e:
                    print(f"   ❌ Failed to send confirmation: {e}")
                    success = False
            
            else:
                print(f"   ❓ Unknown action: {action}")
                success = False
                
        except Exception as e:
            print(f"   ❌ Error executing action {action}: {e}")
            action_record['success'] = False
            action_record['error'] = str(e)
            self.action_history.append(action_record)
            return False
        
        # Mark action result and add to history
        action_record['success'] = success
        self.action_history.append(action_record)
        
        # Keep only last 50 actions to avoid memory issues
        if len(self.action_history) > 50:
            self.action_history = self.action_history[-50:]
        
        return success
    
    def get_recent_actions(self, limit: int = 10) -> List[Dict]:
        """Get recent automation actions"""
        return self.action_history[-limit:]
    
    def rollback_action(self, action_id: str) -> bool:
        """Rollback a specific action (placeholder for future implementation)"""
        # This would need to implement reverse operations
        # For now, just log the request
        print(f"   ⚠️ Rollback requested for action {action_id} - not implemented yet")
        return False
    
    def _extract_event_details(self, email: OutlookEmailData, analysis: Dict = None) -> Dict:
        """Extract calendar event details from email content"""
        try:
            import re
            from datetime import datetime, timedelta
            
            # Basic event details
            event_details = {
                'subject': email.subject,
                'body': email.body[:500],  # Truncate body
                'attendees': [email.sender_email],
                'start': None,
                'end': None,
                'location': None
            }
            
            # Extract date/time information
            text = email.body.lower()
            
            # Could use analysis for better extraction if provided
            if analysis and 'calendar_info' in analysis:
                calendar_info = analysis['calendar_info']
                if 'start_time' in calendar_info:
                    event_details['start'] = calendar_info['start_time']
                if 'end_time' in calendar_info:
                    event_details['end'] = calendar_info['end_time']
            
            # Try to extract time
            time_match = re.search(r'(\d{1,2}):(\d{2})\s*(am|pm)?', text, re.IGNORECASE)
            if time_match:
                hour = int(time_match.group(1))
                minute = int(time_match.group(2))
                ampm = time_match.group(3)
                
                if ampm and ampm.lower() == 'pm' and hour != 12:
                    hour += 12
                elif ampm and ampm.lower() == 'am' and hour == 12:
                    hour = 0
                
                # Default to tomorrow if no date specified
                start_date = datetime.now() + timedelta(days=1)
                event_details['start'] = start_date.replace(hour=hour, minute=minute, second=0, microsecond=0)
                event_details['end'] = event_details['start'] + timedelta(hours=1)  # Default 1 hour duration
            
            # Extract location
            location_patterns = [
                r'at\s+([^,\n]+)',  # at Conference Room A
                r'in\s+([^,\n]+)',  # in Building B
                r'location:\s*([^,\n]+)',  # location: Main Office
                r'room\s+([^,\n]+)',  # room 201
            ]
            
            for pattern in location_patterns:
                location_match = re.search(pattern, text, re.IGNORECASE)
                if location_match:
                    event_details['location'] = location_match.group(1).strip()
                    break
            
            # Only return if we found at least a time
            if event_details['start']:
                return event_details
            else:
                return None
                
        except Exception as e:
            print(f"   ❌ Error extracting event details: {e}")
            return None

class EmailHistoryExtractor:
    """Extracts relevant email history for context (inspired by inbox-zero)"""
    
    def __init__(self, outlook: OutlookService):
        self.outlook = outlook
    
    def get_conversation_history(self, email: OutlookEmailData, max_emails: int = 5) -> List[str]:
        """Get conversation history with the sender (simplified approach)"""
        try:
            # Get recent emails and filter client-side (more reliable than complex OData filters)
            endpoint = f"/me/messages?$select=subject,body,receivedDateTime,sender&$top=50&$orderby=receivedDateTime desc"
            
            result = self.outlook._make_graph_request(endpoint)
            
            history = []
            sender_email = email.sender_email.lower()
            
            for msg in result.get('value', []):
                # Check if email is from the same sender
                msg_sender = msg.get('sender', {}).get('emailAddress', {}).get('address', '').lower()
                if msg_sender == sender_email:
                    body_content = msg.get('body', {}).get('content', '')
                    subject = msg.get('subject', '')
                    
                    if body_content and len(body_content.strip()) > 20:
                        clean_body = self._clean_email_content(body_content)
                        history.append(f"Subject: {subject}\n{clean_body[:300]}...")
                        
                        if len(history) >= max_emails:
                            break
            
            return history
            
        except Exception as e:
            print(f"Error getting conversation history: {e}")
            return []
    
    def get_thread_emails(self, email: OutlookEmailData) -> List[OutlookEmailData]:
        """Get emails in the same conversation thread (simplified approach)"""
        try:
            # For now, just return the current email to avoid API complexity
            # This could be enhanced later with a more robust conversation threading approach
            return [email]
            
        except Exception as e:
            print(f"Error getting thread emails: {e}")
            return [email]
    
    def _clean_email_content(self, content: str) -> str:
        """Clean email content for context"""
        import re
        # Remove HTML tags
        clean = re.sub('<.*?>', '', content)
        # Remove excessive whitespace
        clean = re.sub(r'\s+', ' ', clean)
        return clean.strip()

class KnowledgeBase:
    """Simple knowledge base for email context (inspired by inbox-zero)"""
    
    def __init__(self):
        self.entries = self._load_default_knowledge()
    
    def _load_default_knowledge(self) -> List[Dict[str, str]]:
        """Load default knowledge entries"""
        return [
            {
                "topic": "meeting_scheduling",
                "content": "I typically prefer meetings in the afternoon between 2-5 PM. I use Outlook calendar for scheduling.",
                "keywords": ["meeting", "schedule", "appointment", "calendar"]
            },
            {
                "topic": "response_time",
                "content": "I usually respond to emails within 24 hours on business days.",
                "keywords": ["response", "reply", "answer", "get back"]
            },
            {
                "topic": "availability",
                "content": "I'm typically available Monday-Friday, 9 AM - 6 PM UAE time.",
                "keywords": ["available", "availability", "time", "schedule"]
            },
            {
                "topic": "contact_info",
                "content": "For urgent matters, please call or send a follow-up email with 'URGENT' in the subject.",
                "keywords": ["urgent", "emergency", "contact", "phone"]
            }
        ]
    
    def find_relevant_entries(self, email: OutlookEmailData, max_entries: int = 3) -> List[str]:
        """Find relevant knowledge base entries for the email"""
        email_text = (email.subject + " " + email.body).lower()
        
        relevant_entries = []
        for entry in self.entries:
            # Check if any keywords match
            if any(keyword in email_text for keyword in entry["keywords"]):
                relevant_entries.append(entry["content"])
        
        return relevant_entries[:max_entries]

class ContextualDraftGenerator:
    """Enhanced draft generator with multi-source context (inspired by inbox-zero)"""
    
    def __init__(self, llm_service: UnifiedLLMService, outlook_service=None):
        self.llm_service = llm_service
        self.outlook_service = outlook_service
        self.writing_analyzer = WritingStyleAnalyzer(llm_service)
        self.knowledge_base = KnowledgeBase()
    
    def generate_contextual_draft(self, context: EmailContext) -> LLMDraftResult:
        """Generate draft with full context awareness"""
        try:
            # Build comprehensive context prompt
            prompt = self._build_contextual_prompt(context)
            
            # Generate draft with LLM
            response = self.llm_service.call_with_json_parsing(prompt)
            
            if response and 'subject' in response and 'body' in response:
                return LLMDraftResult(
                    subject=response['subject'],
                    body=response['body'],
                    tone=response.get('tone', context.writing_style.tone),
                    confidence=response.get('confidence', 0.8),
                    reasoning=response.get('reasoning', 'Generated with contextual awareness'),
                    alternative_versions=response.get('alternatives', [])
                )
            else:
                # Fallback to template-based generation
                return self._generate_fallback_draft(context)
                
        except Exception as e:
            print(f"Error generating contextual draft: {e}")
            return self._generate_fallback_draft(context)
    
    def _build_contextual_prompt(self, context: EmailContext) -> str:
        """Build comprehensive prompt with all context sources"""
        
        # Email history context
        history_context = ""
        if context.email_history:
            history_context = f"""
PREVIOUS EMAILS WITH THIS SENDER:
{chr(10).join(context.email_history[:3])}
"""
        
        # Knowledge base context
        knowledge_context = ""
        if context.knowledge_base_entries:
            knowledge_context = f"""
RELEVANT KNOWLEDGE:
{chr(10).join(context.knowledge_base_entries)}
"""
        
        # Writing style context
        style_context = f"""
YOUR WRITING STYLE:
- Tone: {context.writing_style.tone}
- Greeting: {context.writing_style.greeting_style}
- Closing: {context.writing_style.closing_style}
- Average length: {context.writing_style.average_length} words
- Use contractions: {context.writing_style.use_contractions}
- Common phrases: {', '.join(context.writing_style.common_phrases[:3])}
"""
        
        # Get current user info
        try:
            user_info = self.outlook_service.get_user_info() if self.outlook_service else {}
            current_user_name = user_info.get('name', 'User')
            current_user_email = user_info.get('email', '')
        except:
            current_user_name = 'User'
            current_user_email = ''
        
        # Main prompt
        prompt = f"""
You are an AI email assistant helping {current_user_name} write a professional reply. Use the provided context to generate a personalized, contextually appropriate response.

IMPORTANT: You are writing AS {current_user_name} <{current_user_email}>, responding TO {context.current_email.sender} <{context.current_email.sender_email}>.

CURRENT EMAIL TO RESPOND TO:
From: {context.current_email.sender} <{context.current_email.sender_email}>
To: {current_user_name} <{current_user_email}>
Subject: {context.current_email.subject}
Body: {context.current_email.body[:500]}...

{history_context}

{knowledge_context}

{style_context}

CRITICAL USER IDENTIFICATION:
- Current user: {current_user_name} ({current_user_email})
- If the email mentions "{current_user_name}" or similar, that refers to YOU (the current user)
- You are writing as {current_user_name}, not about {current_user_name} as a third person
- Any references to "{current_user_name}" in the email should be interpreted as referring to yourself

INSTRUCTIONS:
1. Write a response FROM {current_user_name} TO {context.current_email.sender}
2. Use relevant information from previous emails and knowledge base
3. Be professional but match the established communication pattern
4. Address the main points in the current email
5. Use the user's preferred greeting and closing styles
6. Remember: You ARE {current_user_name}, not someone else writing about {current_user_name}

Return JSON with:
{{
    "subject": "Re: [original subject with any modifications]",
    "body": "Complete email body with greeting, content, and closing",
    "tone": "professional|friendly|formal",
    "confidence": 0.85,
    "reasoning": "Brief explanation of response approach",
    "alternatives": ["Alternative version 1", "Alternative version 2"]
}}
"""
        
        return prompt
    
    def _generate_fallback_draft(self, context: EmailContext) -> LLMDraftResult:
        """Generate fallback draft when AI fails"""
        email = context.current_email
        style = context.writing_style
        
        # Simple template-based response
        greeting = f"{style.greeting_style} {email.sender.split()[0] if email.sender else 'there'}"
        
        body_content = "Thank you for your email. I have received your message and will review it carefully."
        
        closing = f"\\n\\n{style.closing_style}"
        if style.signature:
            closing += f"\n{style.signature}"
        
        full_body = f"{greeting},\\n\\n{body_content}{closing}"
        
        return LLMDraftResult(
            subject=f"Re: {email.subject}",
            body=full_body,
            tone=style.tone,
            confidence=0.6,
            reasoning="Fallback template-based response",
            alternative_versions=[]
        )

# Example usage and CLI
def main():
    """Main function for LLM-enhanced email processing"""
    
    print("🤖 LLM-Enhanced Email Agent")
    print("Using Ollama for Advanced Email Analysis & Drafting")
    print("=" * 50)
    
    # Check requirements
    if not os.getenv('AZURE_CLIENT_ID'):
        print("❌ AZURE_CLIENT_ID not found in environment variables")
        return
    
    # Initialize LLM system
    try:
        email_system = LLMEnhancedEmailSystem()
        
        # Process emails
        max_emails = 10
        priority_threshold = 60.0
        
        email_system.process_emails_with_llm(max_emails, priority_threshold)
        
    except Exception as e:
        print(f"❌ Failed to initialize LLM system: {e}")
        print("📝 Make sure Ollama is running: 'ollama serve'")
        print("📝 And mistral model is installed: 'ollama pull mistral'")

if __name__ == "__main__":
    main()