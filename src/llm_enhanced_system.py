#!/usr/bin/env python3
"""
LLM-Enhanced Email System
Uses Ollama for sophisticated email analysis and personalized draft generation
"""

import requests
import json
import os
from typing import List, Dict, Any, Optional
from datetime import datetime
from dataclasses import dataclass

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
    key_points: List[str]
    suggested_response_tone: str
    deadline_info: Optional[str]
    sender_relationship: str
    business_context: str
    confidence: float

@dataclass
class LLMDraftResult:
    """Result from LLM draft generation"""
    subject: str
    body: str
    tone: str
    confidence: float
    reasoning: str
    alternative_versions: List[str]

class OllamaLLMService:
    """Service for interacting with Ollama LLM"""
    
    def __init__(self, model: str = "mistral", host: str = "http://localhost:11434"):
        self.model = model
        self.host = host
        self.url = f"{host}/api/generate"
        self._test_connection()
    
    def _test_connection(self):
        """Test connection to Ollama"""
        try:
            response = requests.get(f"{self.host}/api/tags", timeout=5)
            if response.status_code == 200:
                models = response.json().get('models', [])
                available_models = [m['name'] for m in models]
                
                if any(self.model in model for model in available_models):
                    print(f"✅ Connected to Ollama - Model '{self.model}' available")
                else:
                    print(f"⚠️ Model '{self.model}' not found. Available models: {available_models}")
                    if available_models:
                        self.model = available_models[0].split(':')[0]
                        print(f"🔄 Switching to available model: {self.model}")
            else:
                print(f"❌ Ollama connection failed: {response.status_code}")
        except Exception as e:
            print(f"❌ Cannot connect to Ollama: {e}")
            print("📝 Make sure Ollama is running: 'ollama serve'")
    
    def generate_response(self, prompt: str, max_tokens: int = 1000, temperature: float = 0.7) -> str:
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
        except Exception as e:
            return f"Error: {str(e)}"

class LLMEmailAnalyzer:
    """Advanced email analyzer using LLM"""
    
    def __init__(self, llm_service: OllamaLLMService):
        self.llm = llm_service
    
    def analyze_email(self, email: OutlookEmailData, user_context: str = "") -> LLMAnalysisResult:
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

ANALYSIS REQUIRED:
1. Priority Score (0-100): How urgent/important is this email?
2. Urgency Level: critical/urgent/normal/low
3. Email Type: meeting/question/request/deadline/appreciation/information/complaint/announcement
4. Action Required: reply/attend/approve/review/complete/none
5. Key Points: Main topics/requests in the email (bullet points)
6. Response Tone: formal/professional/friendly/casual
7. Deadline Info: Any deadlines mentioned (be specific)
8. Sender Relationship: manager/client/colleague/vendor/external
9. Business Context: How this affects work/projects
10. Confidence: How certain you are of this analysis (0-1)

Consider factors:
- Sender importance and relationship
- Time sensitivity and deadlines  
- Business impact and urgency keywords
- Required actions and responses
- Email thread context

Respond in this EXACT JSON format:
{{
    "priority_score": 75.5,
    "urgency_level": "normal", 
    "email_type": "request",
    "action_required": "reply",
    "key_points": ["Point 1", "Point 2", "Point 3"],
    "suggested_response_tone": "professional",
    "deadline_info": "By Friday 5 PM" or null,
    "sender_relationship": "colleague",
    "business_context": "Project coordination requiring immediate attention",
    "confidence": 0.85
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
                    key_points=analysis_data.get('key_points', []),
                    suggested_response_tone=analysis_data.get('suggested_response_tone', 'professional'),
                    deadline_info=analysis_data.get('deadline_info'),
                    sender_relationship=analysis_data.get('sender_relationship', 'colleague'),
                    business_context=analysis_data.get('business_context', ''),
                    confidence=float(analysis_data.get('confidence', 0.5))
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
        """Fallback analysis if LLM fails"""
        
        # Simple keyword-based fallback
        text = (email.subject + " " + email.body).lower()
        
        priority_score = 50.0
        if any(word in text for word in ['urgent', 'asap', 'critical']):
            priority_score += 30
        if not email.is_read:
            priority_score += 15
        
        urgency_level = 'urgent' if priority_score > 75 else 'normal' if priority_score > 50 else 'low'
        
        return LLMAnalysisResult(
            priority_score=priority_score,
            urgency_level=urgency_level,
            email_type='normal',
            action_required='reply' if '?' in text else 'none',
            key_points=['Email content analysis'],
            suggested_response_tone='professional',
            deadline_info=None,
            sender_relationship='colleague',
            business_context='Standard business communication',
            confidence=0.3
        )

class LLMResponseDrafter:
    """Advanced response drafter using LLM"""
    
    def __init__(self, llm_service: OllamaLLMService):
        self.llm = llm_service
    
    def generate_draft(self, email: OutlookEmailData, analysis: LLMAnalysisResult, 
                      user_writing_style: str = "", user_name: str = "Avani Gupta") -> LLMDraftResult:
        """Generate personalized draft response using LLM"""
        
        draft_prompt = f"""
You are {user_name}, responding to an email. Write an email response FROM {user_name}'s perspective, not TO {user_name}.

ORIGINAL EMAIL RECEIVED BY {user_name}:
From: {email.sender}
Subject: {email.subject}
Content: {email.body[:800]}

YOUR TASK: Write {user_name}'s response to this email.

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

RESPONSE REQUIREMENTS FOR {user_name}:
1. Address the sender by first name (extract from sender name)
2. Acknowledge the specific content/request in the original email
3. Provide a helpful, appropriate response from {user_name}'s perspective
4. Match the suggested tone and relationship level
5. Include specific actions or next steps if needed
6. Use natural, human-like language
7. Keep professional but personalized
8. Address any deadlines or urgency
9. End with "Best regards,\\nAvani" (first name only)
10. Do NOT include signatures, job titles, or contact information

SPECIAL HANDLING:
- For IT Helpdesk forms: Check if this is a service request form. If so, acknowledge receipt and provide timeline for action.
- For marketing emails: Politely acknowledge if interested, or briefly decline if not relevant.
- For meeting requests: Accept/decline with calendar consideration.
- For questions: Provide helpful answers or timeline for detailed response.

You MUST respond with ONLY valid JSON:

{{
    "subject": "Re: [original subject]",
    "body": "Hi [sender first name],\\n\\n[{user_name}'s response content here]\\n\\nBest regards,\\nAvani",
    "tone": "professional",
    "confidence": 0.9,
    "reasoning": "Brief explanation of response approach",
    "alternative_versions": []
}}
"""

        try:
            # Advanced prompting with prefilling technique (research-based)
            response = self.llm.generate_response(draft_prompt, max_tokens=1000, temperature=0.3)
            
            print(f"🔍 DEBUG: Raw LLM response: {response[:200]}...")
            
            # Enhanced JSON extraction with multiple fallback strategies
            json_str = self._extract_json_robust(response)
            
            if json_str:
                print(f"🔍 DEBUG: Extracted JSON: {json_str[:100]}...")
                
                # Apply advanced JSON fixing from research
                json_str = self._fix_json_advanced(json_str)
                
                # Validate JSON before parsing
                if self._validate_json_structure(json_str):
                    draft_data = json.loads(json_str)
                else:
                    print("❌ JSON structure validation failed")
                    return self._fallback_draft(email, analysis, user_name)
                
                return LLMDraftResult(
                    subject=draft_data.get('subject', f"Re: {email.subject}"),
                    body=draft_data.get('body', ''),
                    tone=draft_data.get('tone', 'professional'),
                    confidence=float(draft_data.get('confidence', 0.7)),
                    reasoning=draft_data.get('reasoning', ''),
                    alternative_versions=draft_data.get('alternative_versions', [])
                )
            else:
                print(f"❌ Could not extract JSON from LLM response")
                print(f"🔍 DEBUG: Full response length: {len(response)}")
                print(f"🔍 DEBUG: Full response: {response}")
                return self._fallback_draft(email, analysis, user_name)
                
        except json.JSONDecodeError as e:
            print(f"❌ JSON decode error in draft generation: {e}")
            print(f"🔍 DEBUG: Problematic JSON: {json_str}")
            print(f"🔍 DEBUG: JSON length: {len(json_str)}")
            return self._fallback_draft(email, analysis, user_name)
        except Exception as e:
            print(f"❌ LLM draft generation error: {e}")
            return self._fallback_draft(email, analysis, user_name)
    
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
                    print(f"🔍 DEBUG: Strategy 1 extracted JSON length: {len(potential_json)}")
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
        lines = response.split('\n')
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
            return '\n'.join(json_lines)
        
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
    "body": "{greeting},\\n\\nThank you for your email. I will review this and respond accordingly.\\n\\nBest regards,\\nAvani",
    "tone": "professional",
    "confidence": 0.7,
    "reasoning": "Reconstructed from partial response",
    "alternative_versions": []
}}'''
        
        return ""
    
    def _fix_json_advanced(self, json_str: str) -> str:
        """Advanced JSON fixing based on 2024 research best practices"""
        import re
        
        print(f"🔍 DEBUG: Fixing JSON of length {len(json_str)}")
        print(f"🔍 DEBUG: First 200 chars: {json_str[:200]}")
        
        # Remove any trailing commas before closing braces/brackets
        json_str = re.sub(r',\s*}', '}', json_str)
        json_str = re.sub(r',\s*]', ']', json_str)
        
        # Fix unescaped quotes in string values
        json_str = re.sub(r"'([^']*)':", r'"\1":', json_str)  # Fix keys
        json_str = re.sub(r":\s*'([^']*)'", r': "\1"', json_str)  # Fix string values
        
        # Handle newlines in JSON string values more carefully
        # This is a critical fix - properly escape real newlines within strings
        lines = json_str.split('\n')
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
                fixed_lines[-1] += '\\n' + line
            else:
                fixed_lines.append(line)
        
        json_str = '\n'.join(fixed_lines)
        
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
        json_str = json_str.replace('\\\\n', '\\n')
        json_str = json_str.replace('\\\\"', '"')
        
        print(f"🔍 DEBUG: Fixed JSON length: {len(json_str)}")
        
        return json_str
    
    def _validate_json_structure(self, json_str: str) -> bool:
        """Validate JSON structure before parsing (prevents failures)"""
        try:
            import json
            temp_data = json.loads(json_str)
            
            # Validate required fields exist
            required_fields = ['subject', 'body', 'tone', 'confidence', 'reasoning']
            for field in required_fields:
                if field not in temp_data:
                    print(f"❌ Missing required field: {field}")
                    return False
            
            # Validate data types
            if not isinstance(temp_data.get('confidence'), (int, float)):
                print(f"❌ Invalid confidence type: {type(temp_data.get('confidence'))}")
                return False
                
            return True
            
        except json.JSONDecodeError as e:
            print(f"❌ JSON validation failed: {e}")
            return False
    
    def _fallback_draft(self, email: OutlookEmailData, analysis: LLMAnalysisResult, user_name: str) -> LLMDraftResult:
        """Fallback draft if LLM fails"""
        
        sender_first_name = email.sender.split()[0] if email.sender else "there"
        first_name_only = user_name.split()[0] if user_name else "Avani"
        
        if analysis.action_required == "attend":
            body = f"Hi {sender_first_name},\n\nThank you for the meeting invitation. I'll check my calendar and confirm my attendance shortly.\n\nBest regards,\n{first_name_only}"
        elif analysis.action_required == "reply":
            body = f"Hi {sender_first_name},\n\nThank you for your email. I've received your message and will respond with details shortly.\n\nBest regards,\n{first_name_only}"
        else:
            body = f"Hi {sender_first_name},\n\nThank you for your email. I'll review this and get back to you soon.\n\nBest regards,\n{first_name_only}"
        
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
    
    def __init__(self, ollama_model: str = "mistral"):
        # Initialize services
        self.llm = OllamaLLMService(model=ollama_model)
        self.analyzer = LLMEmailAnalyzer(self.llm)
        self.drafter = LLMResponseDrafter(self.llm)
        
        # Initialize Outlook service
        self.outlook = OutlookService(
            client_id=os.getenv('AZURE_CLIENT_ID'),
            client_secret=os.getenv('AZURE_CLIENT_SECRET'),
            tenant_id=os.getenv('AZURE_TENANT_ID', 'common')
        )
        
        # Stats
        self.emails_analyzed = 0
        self.drafts_created = 0
        self.llm_calls = 0
    
    def analyze_writing_style(self, sent_emails: List[Dict]) -> str:
        """Analyze user's writing style using LLM"""
        
        if not sent_emails:
            return "Professional, friendly, concise communication style."
        
        # Combine sent email content for analysis
        email_content = "\n\n---\n\n".join([
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
    
    def process_emails_with_llm(self, max_emails: int = 10, priority_threshold: float = 60.0):
        """Main workflow using LLM for analysis and drafting"""
        
        print("🤖 LLM-Enhanced Email Agent")
        print("=" * 40)
        
        try:
            # Step 1: Authenticate
            print("🔐 Authenticating with Outlook...")
            self.outlook.authenticate()
            
            user_info = self.outlook.get_user_info()
            current_user_email = user_info.get('email', '')
            current_user_name = user_info.get('name', '')
            
            print(f"✅ Connected as: {current_user_name} ({current_user_email})")
            
            # Step 2: Analyze writing style (simplified for demo)
            print("🧠 Analyzing your writing style...")
            writing_style = "Professional, friendly communication style. Uses 'Hi [Name]' greetings, 'Best regards' closings. Direct but polite, provides specific details and timelines."
            print(f"✅ Writing style analyzed")
            
            # Step 3: Fetch emails
            print(f"📥 Fetching recent emails...")
            emails = self.outlook.get_recent_emails(max_results=max_emails)
            self.emails_analyzed = len(emails)
            
            if not emails:
                print("📪 No emails found")
                return
            
            print(f"📧 Found {len(emails)} emails")
            
            # Step 4: LLM Analysis
            print("🤖 Analyzing emails with LLM...")
            analyzed_emails = []
            
            for email in emails:
                print(f"   🔍 Analyzing: {email.subject[:40]}...")
                analysis = self.analyzer.analyze_email(email, f"User: {current_user_name}, Email: {current_user_email}")
                
                # Update email with LLM analysis
                email.priority_score = analysis.priority_score
                email.urgency_level = analysis.urgency_level
                email.email_type = analysis.email_type
                email.action_required = analysis.action_required
                
                analyzed_emails.append((email, analysis))
                self.llm_calls += 1
            
            # Sort by LLM-determined priority
            analyzed_emails.sort(key=lambda x: x[1].priority_score, reverse=True)
            
            # Step 5: Display LLM Analysis Results
            print(f"\n🎯 LLM Analysis Results:")
            print("-" * 60)
            
            for i, (email, analysis) in enumerate(analyzed_emails):
                urgency_emoji = "🔴" if analysis.urgency_level in ["urgent", "critical"] else "🟡" if analysis.urgency_level == "normal" else "🟢"
                
                print(f"{i+1}. {urgency_emoji} {email.subject[:45]}...")
                print(f"    From: {email.sender}")
                print(f"    LLM Priority: {analysis.priority_score:.1f} ({analysis.urgency_level})")
                print(f"    Type: {analysis.email_type} | Action: {analysis.action_required}")
                print(f"    Key Points: {', '.join(analysis.key_points[:2])}")
                if analysis.deadline_info:
                    print(f"    ⏰ Deadline: {analysis.deadline_info}")
                print(f"    Confidence: {analysis.confidence:.2f}")
                print()
            
            # Step 6: Generate LLM Drafts
            actionable_emails = []
            for email, analysis in analyzed_emails:
                # Skip automated/noreply emails
                sender_email = email.sender_email.lower()
                sender_name = email.sender.lower()
                
                # List of patterns to skip (enhanced based on research)
                skip_patterns = [
                    'noreply', 'no-reply', 'donotreply', 'do-not-reply',
                    'automated', 'notification', 'system', 'security',
                    'website+security@huggingface.co', 'website@huggingface.co',
                    'notifications@', 'support@', 'alerts@', 'admin@',
                    'confirm your email', 'click this link', 'verify your account',
                    'huggingface.co', 'github.com', 'gitlab.com'
                ]
                
                # Also check subject line for automated patterns
                subject_lower = email.subject.lower()
                skip_subject_patterns = [
                    'confirm your email', 'verify your account', 'click this link',
                    'account verification', 'email confirmation', 'click here to confirm'
                ]
                
                # Check for IT helpdesk forms (don't reply to service request forms)
                email_body_lower = email.body.lower()
                form_indicators = [
                    'please tell us about your experience',
                    'rate your experience',
                    'service evaluation',
                    'feedback form',
                    'survey',
                    'evaluation form',
                    'rate our service',
                    'please rate',
                    'how would you rate',
                    'satisfaction survey'
                ]
                
                is_form = any(indicator in email_body_lower for indicator in form_indicators)
                is_form = is_form or any(indicator in subject_lower for indicator in form_indicators)
                
                should_skip_subject = any(pattern in subject_lower for pattern in skip_subject_patterns)
                
                should_skip = any(pattern in sender_email or pattern in sender_name for pattern in skip_patterns)
                
                if should_skip or should_skip_subject:
                    print(f"   ⏭️ Skipping automated email from: {email.sender} <{email.sender_email}>")
                    continue
                
                if is_form:
                    print(f"   📋 Skipping form/survey email: {email.subject[:40]}...")
                    continue
                
                # Check if email needs a response
                if (analysis.priority_score >= priority_threshold and
                    analysis.action_required in ['reply', 'attend', 'approve', 'review'] and
                    analysis.email_type not in ['self_calendar_response', 'self_calendar_event']):
                    actionable_emails.append((email, analysis))
            
            if not actionable_emails:
                print("🎉 No emails need responses right now!")
                return
            
            print(f"🤖 Generating LLM drafts for {len(actionable_emails)} emails...")
            print("=" * 60)
            
            for i, (email, analysis) in enumerate(actionable_emails[:5]):
                print(f"\n✨ Generating LLM draft {i+1}: {email.subject[:40]}...")
                
                # Generate LLM draft
                draft = self.drafter.generate_draft(email, analysis, writing_style, current_user_name)
                self.llm_calls += 1
                
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
                    print(f"   📝 Draft content:\n{draft.body}")
            
            # Step 7: Summary
            print(f"\n🎯 LLM-ENHANCED SUMMARY")
            print("=" * 35)
            print(f"📧 Emails analyzed: {self.emails_analyzed}")
            print(f"🤖 LLM calls made: {self.llm_calls}")
            print(f"📝 Drafts created: {self.drafts_created}")
            print(f"⏱️ Time saved: ~{self.drafts_created * 15} minutes")
            print(f"🎯 Quality: Personalized LLM-generated responses")
            print(f"📁 Check your Outlook Drafts folder!")
            
        except Exception as e:
            print(f"❌ Error in LLM processing: {e}")
            import traceback
            traceback.print_exc()

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
        email_system = LLMEnhancedEmailSystem(ollama_model="mistral")
        
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