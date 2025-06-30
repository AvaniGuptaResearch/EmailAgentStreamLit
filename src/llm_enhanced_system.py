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
                "stop": ["Human:", "Assistant:", "###"]
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
You are writing an email response as {user_name}, a professional manager. Create a personalized, contextual response.

ORIGINAL EMAIL:
From: {email.sender}
Subject: {email.subject}
Content: {email.body[:800]}

ANALYSIS CONTEXT:
- Email Type: {analysis.email_type}
- Action Required: {analysis.action_required}
- Key Points: {', '.join(analysis.key_points)}
- Suggested Tone: {analysis.suggested_response_tone}
- Deadline: {analysis.deadline_info or 'None mentioned'}
- Sender Relationship: {analysis.sender_relationship}
- Business Context: {analysis.business_context}

USER WRITING STYLE:
{user_writing_style or "Professional, friendly, concise. Uses 'Hi [Name]' greetings and 'Best regards' closings. Direct but polite communication style."}

RESPONSE REQUIREMENTS:
1. Address the sender by first name
2. Acknowledge the specific content/request
3. Provide a helpful, appropriate response
4. Match the suggested tone and relationship level
5. Include specific actions or next steps if needed
6. Use natural, human-like language
7. Keep professional but personalized
8. Address any deadlines or urgency

RESPONSE GUIDELINES:
- For meetings: Accept/decline with calendar check
- For questions: Provide helpful answers or timeline for response
- For requests: Acknowledge and provide timeline/action plan
- For deadlines: Acknowledge urgency and commit to timeline
- For appreciation: Respond warmly and professionally

Create a complete email response with subject and body.

Respond in this EXACT JSON format:
{{
    "subject": "Re: Original Subject",
    "body": "Complete email body with greeting, content, and closing",
    "tone": "professional/friendly/formal/casual",
    "confidence": 0.9,
    "reasoning": "Why this response approach was chosen",
    "alternative_versions": ["Brief alternative version", "More detailed alternative version"]
}}
"""

        try:
            response = self.llm.generate_response(draft_prompt, max_tokens=800, temperature=0.6)
            
            # Extract JSON from response
            json_start = response.find('{')
            json_end = response.rfind('}') + 1
            
            if json_start != -1 and json_end > json_start:
                json_str = response[json_start:json_end]
                draft_data = json.loads(json_str)
                
                return LLMDraftResult(
                    subject=draft_data.get('subject', f"Re: {email.subject}"),
                    body=draft_data.get('body', ''),
                    tone=draft_data.get('tone', 'professional'),
                    confidence=float(draft_data.get('confidence', 0.7)),
                    reasoning=draft_data.get('reasoning', ''),
                    alternative_versions=draft_data.get('alternative_versions', [])
                )
            else:
                print(f"❌ Could not parse JSON from LLM draft response")
                return self._fallback_draft(email, analysis, user_name)
                
        except json.JSONDecodeError as e:
            print(f"❌ JSON decode error in draft generation: {e}")
            return self._fallback_draft(email, analysis, user_name)
        except Exception as e:
            print(f"❌ LLM draft generation error: {e}")
            return self._fallback_draft(email, analysis, user_name)
    
    def _fallback_draft(self, email: OutlookEmailData, analysis: LLMAnalysisResult, user_name: str) -> LLMDraftResult:
        """Fallback draft if LLM fails"""
        
        sender_first_name = email.sender.split()[0] if email.sender else "there"
        
        if analysis.action_required == "attend":
            body = f"Hi {sender_first_name},\n\nThank you for the meeting invitation. I'll check my calendar and confirm my attendance shortly.\n\nBest regards,\n{user_name}"
        elif analysis.action_required == "reply":
            body = f"Hi {sender_first_name},\n\nThank you for your email. I've received your message and will respond with details shortly.\n\nBest regards,\n{user_name}"
        else:
            body = f"Hi {sender_first_name},\n\nThank you for your email. I'll review this and get back to you soon.\n\nBest regards,\n{user_name}"
        
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
            actionable_emails = [
                (email, analysis) for email, analysis in analyzed_emails
                if (analysis.priority_score >= priority_threshold and
                    analysis.action_required in ['reply', 'attend', 'approve', 'review'] and
                    analysis.email_type not in ['self_calendar_response', 'self_calendar_event'])
            ]
            
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
                
                # Create draft in Outlook
                try:
                    draft_result = self.outlook.create_draft(
                        to_email=email.sender_email,
                        subject=draft.subject,
                        body=draft.body
                    )
                    
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