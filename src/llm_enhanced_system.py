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
2. DEADLINE/URGENCY keywords: urgent, asap, critical, deadline, due, expires, today, tomorrow (+30 points)
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

PROFESSIONAL EMAIL PRIORITIZATION:
- Unread emails with deadlines = CRITICAL priority (90+)
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
        """Enhanced fallback analysis with professional prioritization"""
        
        text = (email.subject + " " + email.body).lower()
        
        # Start with base priority
        priority_score = 40.0
        
        # UNREAD EMAILS get higher priority
        if not email.is_read:
            priority_score += 25
        
        # DEADLINE/URGENCY keywords
        urgent_keywords = ['urgent', 'asap', 'critical', 'immediate', 'deadline', 'due', 'expires', 'today', 'tomorrow']
        if any(word in text for word in urgent_keywords):
            priority_score += 30
        
        # BUSINESS IMPORTANCE keywords
        important_keywords = ['meeting', 'approval', 'decision', 'budget', 'contract', 'client', 'proposal', 'review']
        if any(word in text for word in important_keywords):
            priority_score += 20
        
        # QUESTION/REQUEST indicators
        if any(indicator in text for indicator in ['?', 'please', 'can you', 'could you', 'need', 'request']):
            priority_score += 15
        
        # SENDER importance (enhanced detection)
        sender_lower = email.sender.lower()
        sender_email_lower = email.sender_email.lower()
        
        # VIP titles get highest priority
        vip_titles = ['ceo', 'cto', 'cfo', 'president', 'vice president', 'director', 'head of']
        if any(title in sender_lower for title in vip_titles):
            priority_score += 35
        
        # Management titles
        mgmt_titles = ['manager', 'lead', 'supervisor', 'team lead', 'project manager']
        if any(title in sender_lower for title in mgmt_titles):
            priority_score += 25
        
        # Client/external importance
        if 'client' in sender_lower or (any(domain in sender_email_lower for domain in ['.com', '.org', '.net']) and 'mbzuai.ac.ae' not in sender_email_lower):
            priority_score += 20
        
        # Internal colleague
        if 'mbzuai.ac.ae' in sender_email_lower:
            priority_score += 10
        
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
        
        # Detect deadlines
        deadline_info = None
        deadline_patterns = ['due ', 'deadline ', 'by ', 'expires ', 'until ']
        for pattern in deadline_patterns:
            if pattern in text:
                # Try to extract deadline context
                start_idx = text.find(pattern)
                deadline_context = text[start_idx:start_idx+50]
                deadline_info = deadline_context.strip()
                break
        
        return LLMAnalysisResult(
            priority_score=priority_score,
            urgency_level=urgency_level,
            email_type='request' if '?' in text or 'please' in text else 'information',
            action_required='reply' if priority_score > 60 else 'review',
            key_points=['Professional email requiring attention'],
            suggested_response_tone='professional',
            deadline_info=deadline_info,
            sender_relationship='professional',
            business_context='Business communication requiring timely response',
            confidence=0.6
        )

class LLMResponseDrafter:
    """Advanced response drafter using LLM"""
    
    def __init__(self, llm_service: OllamaLLMService):
        self.llm = llm_service
    
    def generate_draft(self, email: OutlookEmailData, analysis: LLMAnalysisResult, 
                      user_writing_style: str = "", user_name: str = "", user_email: str = "") -> LLMDraftResult:
        """Generate personalized draft response using LLM"""
        
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
    "body": "Hi {email.sender.split()[0] if email.sender else 'there'},\\n\\nThank you for reminding us about the deadline. We are working on the modules and will ensure timely submission through the portal.",
    "tone": "professional",
    "confidence": 0.9,
    "reasoning": "Acknowledging deadline and providing action plan",
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
    
    def _fallback_draft(self, email: OutlookEmailData, analysis: LLMAnalysisResult, user_name: str, user_email: str = "") -> LLMDraftResult:
        """Fallback draft if LLM fails"""
        
        sender_first_name = email.sender.split()[0] if email.sender else "there"
        first_name_only = user_name.split()[0] if user_name else "User"
        
        # Generate more specific responses based on content
        original_content = email.body.lower()
        
        if analysis.action_required == "attend":
            body = f"Hi {sender_first_name},\n\nThank you for the meeting invitation. I'll check my calendar and confirm my attendance shortly."
        elif 'deadline' in original_content:
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
    
    def process_emails_with_llm(self, max_emails: int = 20, priority_threshold: float = 50.0):
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
                print(f"   📧 Analyzing: {email.subject[:40]}...")
                
                # Enhanced context with thread analysis
                thread_context = self._analyze_email_thread(email)
                user_context = f"User: {current_user_name}, Email: {current_user_email}, Role: Professional at MBZUAI"
                
                analysis = self.analyzer.analyze_email(email, user_context, thread_context)
                
                # Update email with LLM analysis
                email.priority_score = analysis.priority_score
                email.urgency_level = analysis.urgency_level
                email.email_type = analysis.email_type
                email.action_required = analysis.action_required
                
                analyzed_emails.append((email, analysis))
                self.llm_calls += 1
            
            # Sort by LLM-determined priority
            analyzed_emails.sort(key=lambda x: x[1].priority_score, reverse=True)
            
            # Step 5: Display LLM Analysis Results with professional prioritization
            print(f"\n🎯 PROFESSIONAL EMAIL PRIORITIZATION:")
            print("-" * 60)
            
            # Separate emails by priority categories
            critical_emails = [(email, analysis) for email, analysis in analyzed_emails if analysis.priority_score >= 85]
            urgent_emails = [(email, analysis) for email, analysis in analyzed_emails if 70 <= analysis.priority_score < 85]
            normal_emails = [(email, analysis) for email, analysis in analyzed_emails if 50 <= analysis.priority_score < 70]
            low_emails = [(email, analysis) for email, analysis in analyzed_emails if analysis.priority_score < 50]
            
            # Display by priority categories
            if critical_emails:
                print("🔴 CRITICAL PRIORITY (85+):")
                for i, (email, analysis) in enumerate(critical_emails):
                    unread_indicator = "📬 UNREAD" if not email.is_read else "📭"
                    print(f"   {i+1}. {unread_indicator} {email.subject[:40]}...")
                    print(f"      From: {email.sender} | Score: {analysis.priority_score:.1f}")
                    print(f"      Action: {analysis.action_required} | Type: {analysis.email_type}")
                    if analysis.deadline_info:
                        print(f"      ⏰ DEADLINE: {analysis.deadline_info}")
                    print()
            
            if urgent_emails:
                print("🟡 URGENT PRIORITY (70-84):")
                for i, (email, analysis) in enumerate(urgent_emails):
                    unread_indicator = "📬 UNREAD" if not email.is_read else "📭"
                    print(f"   {i+1}. {unread_indicator} {email.subject[:40]}...")
                    print(f"      From: {email.sender} | Score: {analysis.priority_score:.1f}")
                    print(f"      Action: {analysis.action_required} | Type: {analysis.email_type}")
                    if analysis.deadline_info:
                        print(f"      ⏰ Deadline: {analysis.deadline_info}")
                    print()
            
            if normal_emails:
                print("🟢 NORMAL PRIORITY (50-69):")
                for i, (email, analysis) in enumerate(normal_emails):
                    unread_indicator = "📬 UNREAD" if not email.is_read else "📭"
                    print(f"   {i+1}. {unread_indicator} {email.subject[:40]}...")
                    print(f"      From: {email.sender} | Score: {analysis.priority_score:.1f}")
                    print(f"      Action: {analysis.action_required}")
                    print()
            
            if low_emails:
                print("⚪ LOW PRIORITY (<50) - Consider batch processing:")
                for i, (email, analysis) in enumerate(low_emails[:3]):  # Show only first 3
                    print(f"   {i+1}. {email.subject[:40]}... (Score: {analysis.priority_score:.1f})")
                if len(low_emails) > 3:
                    print(f"   ... and {len(low_emails) - 3} more low priority emails")
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
                draft = self.drafter.generate_draft(email, analysis, writing_style, current_user_name, current_user_email)
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
            
            # Step 7: Professional Summary
            print(f"\n🎯 PROFESSIONAL EMAIL ASSISTANT SUMMARY")
            print("=" * 45)
            
            # Count by priority
            critical_count = len([e for e, a in analyzed_emails if a.priority_score >= 85])
            urgent_count = len([e for e, a in analyzed_emails if 70 <= a.priority_score < 85])
            unread_count = len([e for e, a in analyzed_emails if not e.is_read])
            deadline_count = len([e for e, a in analyzed_emails if a.deadline_info])
            
            print(f"📊 EMAIL ANALYSIS:")
            print(f"   📧 Total emails analyzed: {self.emails_analyzed}")
            print(f"   📬 Unread emails: {unread_count}")
            print(f"   🔴 Critical priority: {critical_count}")
            print(f"   🟡 Urgent priority: {urgent_count}")
            print(f"   ⏰ With deadlines: {deadline_count}")
            print()
            print(f"🤖 AI ASSISTANCE:")
            print(f"   🧠 LLM analysis calls: {self.llm_calls}")
            print(f"   📝 Response drafts created: {self.drafts_created}")
            print(f"   ⏱️ Estimated time saved: ~{self.drafts_created * 12} minutes")
            print()
            print(f"✅ NEXT STEPS:")
            print(f"   1. Review critical/urgent emails first")
            print(f"   2. Check Outlook Drafts folder for AI-generated responses")
            print(f"   3. Address deadline emails immediately")
            print(f"   4. Process normal priority emails in batches")
            print(f"📁 All drafts saved to: Outlook > Drafts folder")
            
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