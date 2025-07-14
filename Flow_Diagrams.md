# Email Assistant Flow Diagrams

## 1. OVERVIEW FLOW DIAGRAM

```
📧 EMAIL ASSISTANT - SYSTEM OVERVIEW
════════════════════════════════════════════════════════════════

📮 OUTLOOK INBOX                    🤖 AI PROCESSING                    📤 SMART OUTPUTS
┌─────────────────┐                ┌─────────────────┐                ┌─────────────────┐
│                 │                │                 │                │                 │
│  📧 New Emails  │──────────────▶ │  🧠 LLM Analysis│──────────────▶ │  📊 Prioritized │
│                 │                │                 │                │     Dashboard   │
│  📬 Unread: 4   │                │  🔍 Priority    │                │                 │
│  📭 Read: 5     │                │     Scoring     │                │  🔴 Critical: 2 │
│                 │                │                 │                │  🟡 Urgent: 1   │
│  ⏰ 72hr scan   │                │  📂 Smart       │                │  🟢 Normal: 3   │
│                 │                │     Categories  │                │  ⚪ Low: 3      │
└─────────────────┘                │                 │                │                 │
                                   │  📝 Draft       │                │  📝 3 Drafts    │
                                   │     Generation  │                │     Generated   │
                                   │                 │                │                 │
                                   │  🛡️ Security    │                │  ✅ 0 Threats  │
                                   │     Analysis    │                │     Detected    │
                                   └─────────────────┘                └─────────────────┘

🔄 PROCESSING MODES:
🔬 Deep (20 emails) | ⚡ Lite (10 emails) | 🏃 Ultra-Lite (5 emails)

⏱️ TIME SAVED: ~36 minutes per batch
```

## 2. END-USER FLOW DIAGRAM

```
👤 END-USER JOURNEY
══════════════════════════════════════════════════════════════════════════════════════

START
  │
  ▼
┌─────────────────────────────────────────────────────────────────────────────────────┐
│ 1️⃣ LAUNCH APPLICATION                                                               │
│ streamlit run src/streamlit_app.py                                                  │
│ 🌐 Open http://localhost:8501                                                      │
└─────────────────────────────────────────────────────────────────────────────────────┘
  │
  ▼
┌─────────────────────────────────────────────────────────────────────────────────────┐
│ 2️⃣ SELECT PROCESSING MODE                                                          │
│ 🔬 Deep Mode (Full Analysis)     ⚡ Lite Mode (Fast)     🏃 Ultra-Lite (Keywords) │
│ • 20 emails max                  • 10 emails max         • 5 emails max           │
│ • All features                   • Core LLM only         • Keywords + drafting    │
│ • ~5 min processing              • ~2 min processing     • ~30 sec processing     │
└─────────────────────────────────────────────────────────────────────────────────────┘
  │
  ▼
┌─────────────────────────────────────────────────────────────────────────────────────┐
│ 3️⃣ AUTHENTICATE                                                                    │
│ 🚀 Click "Initialize System"                                                       │
│ 🔐 Microsoft Azure AD Login                                                        │
│ ✅ Connect to: Avani.Gupta@mbzuai.ac.ae                                           │
└─────────────────────────────────────────────────────────────────────────────────────┘
  │
  ▼
┌─────────────────────────────────────────────────────────────────────────────────────┐
│ 4️⃣ PROCESS EMAILS                                                                  │
│ 🤖 Click "Process Emails"                                                          │
│ 📺 Watch Live Processing Stream:                                                   │
│   📥 Fetching emails...                                                            │
│   🔍 Analyzing with LLM...                                                         │
│   📊 Generating priority scores...                                                 │
│   📝 Creating drafts...                                                            │
└─────────────────────────────────────────────────────────────────────────────────────┘
  │
  ▼
┌─────────────────────────────────────────────────────────────────────────────────────┐
│ 5️⃣ REVIEW RESULTS                                                                  │
│ 📊 PRIORITY DASHBOARD:                                                             │
│   🔴 Critical (85+): Bank details (90.0), Newsletter deadline (85.0)             │
│   🟡 Urgent (70-84): Email agent tasks (80.0)                                     │
│   🟢 Normal (50-69): Jira invites, Grant updates                                  │
│                                                                                     │
│ 📂 SMART CATEGORIES:                                                               │
│   📁 HR: 2 emails  📁 Administrative: 2 emails  📁 Academic: 1 email             │
│                                                                                     │
│ 🛡️ SECURITY: 0 threats detected                                                   │
└─────────────────────────────────────────────────────────────────────────────────────┘
  │
  ▼
┌─────────────────────────────────────────────────────────────────────────────────────┐
│ 6️⃣ USE SMART DRAFTS                                                                │
│ 📝 3 Drafts Generated (0.85 confidence)                                           │
│ 📧 Saved to: Outlook > Drafts folder                                              │
│ 🎯 Professional tone matched to your writing style                                │
│ ✏️ Edit/Send directly from Outlook                                                │
└─────────────────────────────────────────────────────────────────────────────────────┘
  │
  ▼
┌─────────────────────────────────────────────────────────────────────────────────────┐
│ 7️⃣ TRACK PROGRESS                                                                  │
│ ⏱️ Time Saved: ~36 minutes                                                        │
│ 📈 Processed: 9 emails (4 unread, 5 read)                                        │
│ 🎯 Focus on: 2 Critical + 1 Urgent = 3 priority emails                           │
│ 🔄 Batch process: 6 normal/low priority emails                                    │
└─────────────────────────────────────────────────────────────────────────────────────┘
  │
  ▼
END - EMAIL MANAGEMENT COMPLETE ✅

💡 USER BENEFITS:
• 🎯 Focus on what matters most (priority-based)
• ⏱️ Save 60% time on email processing
• 📝 Professional responses ready instantly
• 🛡️ Security threats automatically detected
• 🔄 Reduce context switching with batch processing
```

## 3. TECHNICAL ARCHITECTURE DIAGRAM

```
🔧 TECHNICAL ARCHITECTURE - DETAILED SYSTEM FLOW
════════════════════════════════════════════════════════════════════════════════════════

┌─────────────────────────────────────────────────────────────────────────────────────┐
│ 🎨 FRONTEND LAYER (Streamlit)                                                      │
│ ┌─────────────────┐  ┌─────────────────┐  ┌─────────────────┐                     │
│ │ streamlit_app.py│  │ Mode Selection  │  │ Real-time UI    │                     │
│ │ • Page config   │  │ • Deep/Lite/    │  │ • Live streaming│                     │
│ │ • Session state │  │   Ultra-Lite    │  │ • Progress bars │                     │
│ │ • UI components │  │ • Email limits  │  │ • Status updates│                     │
│ └─────────────────┘  └─────────────────┘  └─────────────────┘                     │
└─────────────────────────────────────────────────────────────────────────────────────┘
                                    │
                                    ▼
┌─────────────────────────────────────────────────────────────────────────────────────┐
│ 🔐 AUTHENTICATION LAYER (Azure AD)                                                 │
│ ┌─────────────────┐  ┌─────────────────┐  ┌─────────────────┐                     │
│ │ OAuth2 Flow     │  │ Token Management│  │ User Session    │                     │
│ │ • Client ID     │  │ • Access tokens │  │ • Avani.Gupta@  │                     │
│ │ • Client Secret │  │ • Refresh tokens│  │   mbzuai.ac.ae  │                     │
│ │ • Tenant ID     │  │ • Expiration    │  │ • Permission    │                     │
│ └─────────────────┘  └─────────────────┘  └─────────────────┘                     │
└─────────────────────────────────────────────────────────────────────────────────────┘
                                    │
                                    ▼
┌─────────────────────────────────────────────────────────────────────────────────────┐
│ 📧 EMAIL INTEGRATION LAYER (Microsoft Graph)                                       │
│ ┌─────────────────┐  ┌─────────────────┐  ┌─────────────────┐                     │
│ │ outlook_agent.py│  │ Email Fetching  │  │ Draft Creation  │                     │
│ │ • Graph API     │  │ • Inbox scan    │  │ • Reply drafts  │                     │
│ │ • Email parsing │  │ • 72hr window   │  │ • Thread context│                     │
│ │ • Metadata extr │  │ • Unread filter │  │ • Save to drafts│                     │
│ └─────────────────┘  └─────────────────┘  └─────────────────┘                     │
└─────────────────────────────────────────────────────────────────────────────────────┘
                                    │
                                    ▼
┌─────────────────────────────────────────────────────────────────────────────────────┐
│ 🧠 CORE PROCESSING LAYER (LLM Enhanced System)                                     │
│ ┌─────────────────┐  ┌─────────────────┐  ┌─────────────────┐                     │
│ │llm_enhanced_    │  │ Mode Controller │  │ Email Processor │                     │
│ │system.py        │  │ • Deep: 20 max  │  │ • Batch process │                     │
│ │ • Main controller│  │ • Lite: 10 max  │  │ • Priority sort │                     │
│ │ • Processing mgr│  │ • Ultra: 5 max  │  │ • Thread analysis│                     │
│ └─────────────────┘  └─────────────────┘  └─────────────────┘                     │
└─────────────────────────────────────────────────────────────────────────────────────┘
                                    │
                                    ▼
┌─────────────────────────────────────────────────────────────────────────────────────┐
│ 🤖 AI PROCESSING LAYER (LLM Analysis)                                              │
│ ┌─────────────────┐  ┌─────────────────┐  ┌─────────────────┐                     │
│ │ LLM Client      │  │ Analysis Engine │  │ Response Cache  │                     │
│ │ • Ollama/OpenAI │  │ • Priority score│  │ • Hash-based    │                     │
│ │ • Model: Mistral│  │ • Urgency detect│  │ • Performance   │                     │
│ │ • Local host    │  │ • Action extract│  │ • Optimization  │                     │
│ └─────────────────┘  └─────────────────┘  └─────────────────┘                     │
└─────────────────────────────────────────────────────────────────────────────────────┘
                                    │
                                    ▼
┌─────────────────────────────────────────────────────────────────────────────────────┐
│ 📊 FEATURE PROCESSING LAYER (Mode-Specific)                                        │
│ ┌─────────────────┐  ┌─────────────────┐  ┌─────────────────┐                     │
│ │ Deep Mode       │  │ Lite Mode       │  │ Ultra-Lite Mode │                     │
│ │ ✅ LLM Analysis │  │ ✅ LLM Analysis │  │ ❌ Keyword Only │                     │
│ │ ✅ Security     │  │ ❌ No Security  │  │ ❌ No Security  │                     │
│ │ ✅ Summarization│  │ ❌ No Summary   │  │ ❌ No Summary   │                     │
│ │ ✅ Templates    │  │ ❌ No Templates │  │ ❌ No Templates │                     │
│ │ ✅ Smart Cat    │  │ ❌ Basic Cat    │  │ ❌ Static Cat   │                     │
│ │ ✅ Calendar     │  │ ❌ No Calendar  │  │ ❌ No Calendar  │                     │
│ │ ✅ Automation   │  │ ❌ No Automation│  │ ❌ No Automation│                     │
│ └─────────────────┘  └─────────────────┘  └─────────────────┘                     │
└─────────────────────────────────────────────────────────────────────────────────────┘
                                    │
                                    ▼
┌─────────────────────────────────────────────────────────────────────────────────────┐
│ 🔍 ANALYSIS MODULES (Deep Mode Only)                                               │
│ ┌─────────────────┐  ┌─────────────────┐  ┌─────────────────┐                     │
│ │ Security        │  │ Smart Category  │  │ Email Summary   │                     │
│ │ • Threat detect │  │ • HR/Admin/Tech │  │ • Key points    │                     │
│ │ • Risk scoring  │  │ • Subcategories │  │ • Action items  │                     │
│ │ • Safety flags  │  │ • Confidence    │  │ • People/dates  │                     │
│ └─────────────────┘  └─────────────────┘  └─────────────────┘                     │
│                                                                                     │
│ ┌─────────────────┐  ┌─────────────────┐  ┌─────────────────┐                     │
│ │ Template Engine │  │ Follow-up Track │  │ Calendar Integ  │                     │
│ │ • Quick replies │  │ • Deadline mgmt │  │ • Meeting detect│                     │
│ │ • Smart suggest │  │ • Priority queue│  │ • Event creation│                     │
│ │ • Personalized  │  │ • Status track  │  │ • Reminder set  │                     │
│ └─────────────────┘  └─────────────────┘  └─────────────────┘                     │
└─────────────────────────────────────────────────────────────────────────────────────┘
                                    │
                                    ▼
┌─────────────────────────────────────────────────────────────────────────────────────┐
│ 📝 DRAFT GENERATION LAYER                                                          │
│ ┌─────────────────┐  ┌─────────────────┐  ┌─────────────────┐                     │
│ │ Contextual      │  │ Writing Style   │  │ Response Drafter│                     │
│ │ Draft Generator │  │ Analysis        │  │ • Subject gen   │                     │
│ │ • Multi-context │  │ • Tone matching │  │ • Body content  │                     │
│ │ • History aware │  │ • Formal/casual │  │ • Confidence    │                     │
│ │ • Knowledge base│  │ • Signature     │  │ • Alternatives  │                     │
│ └─────────────────┘  └─────────────────┘  └─────────────────┘                     │
└─────────────────────────────────────────────────────────────────────────────────────┘
                                    │
                                    ▼
┌─────────────────────────────────────────────────────────────────────────────────────┐
│ 💾 OUTPUT LAYER (Results & Analytics)                                              │
│ ┌─────────────────┐  ┌─────────────────┐  ┌─────────────────┐                     │
│ │ Priority        │  │ Draft Storage   │  │ Analytics       │                     │
│ │ Dashboard       │  │ • Outlook drafts│  │ • Time saved    │                     │
│ │ • Color coded   │  │ • Thread history│  │ • Emails proc   │                     │
│ │ • Score ranking │  │ • Auto-save     │  │ • Success rates │                     │
│ │ • Categories    │  │ • Preview       │  │ • Performance   │                     │
│ └─────────────────┘  └─────────────────┘  └─────────────────┘                     │
└─────────────────────────────────────────────────────────────────────────────────────┘

📊 DATA FLOW:
Outlook → Graph API → Email Parser → LLM Analysis → Feature Processing → Draft Generation → Output

🔐 SECURITY:
• Local LLM processing (no data leaves network)
• OAuth2 token management
• Encrypted API communications
• No email content stored permanently

⚡ PERFORMANCE:
• Response caching (hash-based)
• Batch processing
• Mode-specific optimization
• Parallel analysis where possible

🧠 AI MODELS:
• Primary: Ollama/Mistral (local)
• Fallback: OpenAI (configurable)
• Context window: 4K tokens
• Response optimization: Stop words, temperature control
```

## 4. PROCESSING MODE COMPARISON DIAGRAM

```
🔄 PROCESSING MODE COMPARISON
════════════════════════════════════════════════════════════════════════════════════════

                    🔬 DEEP MODE          ⚡ LITE MODE         🏃 ULTRA-LITE MODE
                   ┌─────────────┐       ┌─────────────┐       ┌─────────────┐
Email Limit       │     20      │       │     10      │       │      5      │
Processing Time   │   ~5 min    │       │   ~2 min    │       │   ~30 sec   │
                   └─────────────┘       └─────────────┘       └─────────────┘

┌─────────────────────────────────────────────────────────────────────────────────────┐
│ CORE FEATURES                                                                       │
├─────────────────────────────────────────────────────────────────────────────────────┤
│ Priority Scoring    │      ✅ LLM        │      ✅ LLM        │   ⚠️ Keywords   │
│ Urgency Analysis    │      ✅ Full       │      ✅ Full       │   ⚠️ Basic      │
│ Action Detection    │      ✅ Full       │      ✅ Full       │   ⚠️ Basic      │
│ Email Classification│      ✅ Full       │      ✅ Full       │   ⚠️ Basic      │
│ Draft Generation    │      ✅ Enhanced   │      ✅ Basic      │   ✅ Basic      │
├─────────────────────────────────────────────────────────────────────────────────────┤
│ ADVANCED FEATURES                                                                   │
├─────────────────────────────────────────────────────────────────────────────────────┤
│ Security Analysis   │      ✅ Full       │      ❌ None       │   ❌ None       │
│ Email Summarization │      ✅ Yes        │      ❌ No         │   ❌ No         │
│ Template Suggestions│      ✅ Yes        │      ❌ No         │   ❌ No         │
│ Smart Categorization│      ✅ AI-based   │      ❌ Basic      │   ❌ Static     │
│ Calendar Integration│      ✅ Yes        │      ❌ No         │   ❌ No         │
│ Follow-up Tracking  │      ✅ Yes        │      ❌ No         │   ❌ No         │
│ Writing Style Learn │      ✅ Yes        │      ❌ No         │   ❌ No         │
│ Thread Analysis     │      ✅ Full       │      ❌ Minimal    │   ❌ None       │
└─────────────────────────────────────────────────────────────────────────────────────┘

🎯 USE CASES:
┌─────────────────────────────────────────────────────────────────────────────────────┐
│ 🔬 DEEP MODE:                                                                       │
│ • Executive/Manager email processing                                               │
│ • High-stakes business communications                                              │
│ • Security-sensitive environments                                                  │
│ • Complex project management                                                       │
│ • When accuracy > speed                                                            │
├─────────────────────────────────────────────────────────────────────────────────────┤
│ ⚡ LITE MODE:                                                                       │
│ • Daily routine email processing                                                   │
│ • Mid-level professionals                                                          │
│ • Balance of speed and accuracy                                                    │
│ • Regular business communications                                                  │
│ • Moderate email volumes                                                           │
├─────────────────────────────────────────────────────────────────────────────────────┤
│ 🏃 ULTRA-LITE MODE:                                                                │
│ • Quick triage and response                                                        │
│ • High-volume, low-complexity emails                                              │
│ • Speed-critical scenarios                                                         │
│ • Initial email screening                                                          │
│ • When speed >> accuracy                                                           │
└─────────────────────────────────────────────────────────────────────────────────────┘
```

## 5. INTEGRATION POINTS FOR POWERPOINT

**Slide Integration Notes:**

1. **Overview Diagram → Slide 4**: System overview with input/processing/output
2. **End-User Flow → Slide 12**: Demo flow with step-by-step user journey
3. **Technical Architecture → Slide 5**: Technical implementation details
4. **Mode Comparison → Slide 6**: Processing modes with feature breakdown

**Visual Styling Tips:**
- Use consistent color coding: 🔴 Critical, 🟡 Urgent, 🟢 Normal, ⚪ Low
- Maintain emoji consistency for visual appeal
- Use boxes and arrows for flow representation
- Include actual demo metrics (36 min saved, 0.85 confidence, etc.)
- Show real user identity (Avani.Gupta@mbzuai.ac.ae) for authenticity

**Animation Suggestions:**
- Flow diagrams: Animate arrows left-to-right
- Mode comparison: Reveal features row by row
- Technical architecture: Layer-by-layer reveal
- End-user flow: Step-by-step progression with highlighting