import docx
import copy
import os

TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), "template_BSOL_Senior-FS-.Net-ReactJS-Azure_Trinh-Van-Len.docx")
OUTPUT_PATH = os.path.join(os.path.dirname(__file__), "BSOL_Senior-AI-Engineer_Nguyen-Xuan-Cuong.docx")

doc = docx.Document(TEMPLATE_PATH)

# ==============================================================================
# 1. Candidate Title (P0)
# ==============================================================================
p0 = doc.paragraphs[0]
p0.runs[0].text = "NGUYEN XUAN CUONG"

# ==============================================================================
# 2. Personal Details (P4, P5, P6, P7)
# ==============================================================================
p4 = doc.paragraphs[4]  # Name
p4.runs[2].text = "Nguyen Xuan Cuong"

p7 = doc.paragraphs[7]  # was '\t\t'
p7_elem = p7._element
p7_parent = p7_elem.getparent()
p7_parent.remove(p7_elem)

personal_details = [
    ("Location", "\t", "Da Nang, Vietnam"),
    ("Email\t", "\t", "nxcforwork@gmail.com"),
    ("Phone\t", "\t", "+84 327 755 183"),
    ("Profiles", "\t", "GitHub: nxc1802  |  LinkedIn: cuong-nguyen-2715a136b  |  Portfolio: nxc-s-portfolio.vercel.app")
]

curr_elem = doc.paragraphs[6]._element
for label, tab_str, val in personal_details:
    new_p_elem = copy.deepcopy(p4._element)
    curr_elem.addnext(new_p_elem)
    curr_elem = new_p_elem
    new_p = docx.text.paragraph.Paragraph(new_p_elem, doc)
    new_p.runs[0].text = label
    new_p.runs[1].text = tab_str
    new_p.runs[2].text = val

# ==============================================================================
# 3. Educational Background & Additional Sections
# ==============================================================================
p_edu_idx = None
for i, p in enumerate(doc.paragraphs):
    if p.text.strip() == "EDUCATIONAL BACKGROUND":
        p_edu_idx = i
        break

p_edu_h = doc.paragraphs[p_edu_idx]
p_edu_body = doc.paragraphs[p_edu_idx + 2]
p_edu_body.runs[0].text = "FPT University, Da Nang, Vietnam – Bachelor of Science in Information Technology (AI Major)."

p_edu_sub = copy.deepcopy(p_edu_body._element)
p_edu_body._element.addnext(p_edu_sub)
p_edu_sub_p = docx.text.paragraph.Paragraph(p_edu_sub, doc)
p_edu_sub_p.runs[0].text = "Recipient of 100% Full-tuition FPT University Scholarship.  |  GPA: 8.28 / 10.00"

# Insert CERTIFICATE AND AWARD & RESEARCH & PUBLICATIONS before PROFESSIONAL SUMMARY
p_sum_idx = None
for i, p in enumerate(doc.paragraphs):
    if p.text.strip() == "PROFESSIONAL SUMMARY":
        p_sum_idx = i
        break

p_sum_elem = doc.paragraphs[p_sum_idx]._element

def insert_section_before(target_elem, heading_title, items):
    sp1 = copy.deepcopy(p_edu_h._element)
    docx.text.paragraph.Paragraph(sp1, doc).text = ""
    target_elem.addprevious(sp1)
    
    h = copy.deepcopy(p_edu_h._element)
    p_h = docx.text.paragraph.Paragraph(h, doc)
    p_h.runs[0].text = heading_title
    target_elem.addprevious(h)
    
    sp2 = copy.deepcopy(p_edu_h._element)
    docx.text.paragraph.Paragraph(sp2, doc).text = ""
    target_elem.addprevious(sp2)
    
    for item in items:
        it = copy.deepcopy(p_edu_body._element)
        p_it = docx.text.paragraph.Paragraph(it, doc)
        p_it.runs[0].text = item
        target_elem.addprevious(it)

cert_items = [
    "English: IELTS 6.5 Equivalent – Professional working proficiency in technical writing and communication.",
    "Recipient of 100% Full-tuition Scholarship – FPT University (Merit-based full 4-year tuition).",
    "2nd Runner-up: Hackathon GDG DevFest MienTrung 2025 (PiXity Multi-modal Tourism AI Agent).",
    "Second Prize: AI4SE 2025 Competition (Artificial Intelligence for Software Engineering).",
    "4th Place: Hackathon Nosana Builders Challenge: Agents 102 (MagicCV Autonomous Agent).",
    "Superteam Honorable Mention: Timelocked Autonomous Multi-signature Wallet (2025).",
    "Finalist: AI4LIFE 2024, Research Connection 2025, Research Festival 2025, Grab The Beyond 2026."
]

research_items = [
    "MAFIS: Multi-Agent Imitation by learning and sampling from FactorIzed Soft Q-function.",
    "Identification of Plant Leaf Diseases: Using CNN and SVM for precision agriculture.",
    "The GYM Exercise Classification Dataset: 2D/3D pose analysis using CNN-LSTM and MediaPipe.",
    "Towards Sparse Adversarial Perturbations: Gradient-Guided Approach for Efficient Image Attacks."
]

insert_section_before(p_sum_elem, "CERTIFICATE AND AWARD", cert_items)
insert_section_before(p_sum_elem, "RESEARCH & PUBLICATIONS", research_items)

# ==============================================================================
# 4. Professional Summary (In-Place Edit)
# ==============================================================================
summary_replacements = [
    ("Highly skilled IT professional with 9 years of software development experience.",
     "Highly skilled Senior AI & Software Engineer with 5+ years of software development and AI engineering experience. Proven track record in researching, designing, and delivering production-grade Generative AI agents, Advanced RAG architectures, and real-time Computer Vision systems."),
    (".NET web applications (9+ years)",
     "Generative AI & Autonomous Agent Orchestration: LangChain, LangGraph, Mastra, Model Context Protocol (MCP) (3+ years)"),
    ("Front-end frameworks: Angular and React JS (8+ years)",
     "Advanced RAG & Knowledge Graphs: Graph-RAG, Neo4j, Qdrant, Pinecone, Supabase pgvector (3+ years)"),
    ("Proficient in cloud environments, particularly Azure, and experienced with Azure DevOps.",
     "Computer Vision & Edge Deep Learning: YOLOv8/v10, MediaPipe, ByteTrack, Vision Transformers, FPGA & firmware integration (4+ years)"),
    ("Knowledgeable in diverse database systems, including:",
     "High-performance Asynchronous Backend & Distributed Systems: FastAPI, Python, Node.js, Celery, Redis, WebSockets, WebRTC (4+ years)"),
    ("Microsoft SQL Server",
     "Proficient in cloud environments and scalable deployment: AWS, GCP, Azure, Hugging Face Docker Spaces, Vercel, Supabase, Nosana"),
    ("PostgreSQL",
     "Knowledgeable in diverse database systems: PostgreSQL (pgvector), Supabase, Neo4j, Qdrant, Pinecone, Redis (AOF), Microsoft SQL Server"),
    ("MongoDB",
     "Demonstrated technical leadership abilities in designing end-to-end architectures, preparing technical proposals, and effort estimations."),
    ("Demonstrated leadership abilities in team management, fostering collaboration to achieve project objectives efficiently.",
     "Strong communicator with professional English skills (IELTS 6.5 equivalent), effectively engaging with clients, stakeholders, and cross-functional teams."),
    ("Strong communicator with intermediate English skills, effectively engaging with clients and stakeholders.",
     "Involved in all project lifecycle phases from requirement analysis and proof of concept to deployment and optimization."),
    ("Involved in project phases from proof of concept to estimation and planning.",
     "Committed to clean architecture, test-driven development, and industry best practices."),
    ("Committed to coding standards and best practices.",
     "A proactive, research-oriented team player with published papers in multi-agent RL and multiple hackathon awards (GDG DevFest, AI4SE, Nosana)."),
    ("A dedicated team player focused on delivering high-quality solutions.",
     "A dedicated team player focused on delivering high-impact, production-ready AI solutions."),
    ("Hands-on experience leveraging AI-powered IDEs (Cursor) for rapid prototyping, code review, refactoring, and improving overall development velocity",
     "Hands-on experience leveraging AI-powered IDEs (Cursor, Copilot) for rapid prototyping, code review, refactoring, and improving overall development velocity.")
]

for p in doc.paragraphs:
    txt = p.text.strip()
    for orig, rep in summary_replacements:
        if orig in txt:
            p.runs[0].text = rep
            if len(p.runs) > 1:
                for r in p.runs[1:]:
                    r.text = ""

# ==============================================================================
# 5. Software (In-Place Edit)
# ==============================================================================
software_replacements = [
    ("Databases", ": PostgreSQL (pgvector), Supabase, Neo4j, Qdrant, Pinecone, Redis (AOF), Microsoft SQL Server, SQLite."),
    ("Programming Languages", ": Python (Core/Expert), C++, TypeScript, JavaScript, SQL, Shell Script."),
    ("Frameworks & Platforms", ": PyTorch, TensorFlow, LangChain, LangGraph, Mastra, FastAPI, Celery, Node.js, Next.js, React, vLLM, Ollama, YOLOv8/v10, MediaPipe."),
    ("Programming Tools", ": VS Code, Cursor AI IDE, Docker, Git, Postman, Janus WebRTC."),
    ("Source Control", ": GIT (GitHub, GitLab)."),
    ("Software Process", ": Agile/Scrum, Waterfall."),
    ("OOAD/OOP", ": Object Oriented Analysis (OOA), Object Oriented Design (OOD), Object Oriented Programming (OOP), Multi-Agent Architecture, Graph-RAG.")
]

for p in doc.paragraphs:
    for label, val in software_replacements:
        if p.runs and label in p.runs[0].text:
            p.runs[0].text = label
            if len(p.runs) > 1:
                p.runs[1].text = val
                for r in p.runs[2:]:
                    r.text = ""
            break

# ==============================================================================
# 6. Table 0: Technology Skill (Cloned Rows from Template)
# ==============================================================================
t0 = doc.tables[0]
cat_tr = copy.deepcopy(t0.rows[1]._tr)
data_tr = copy.deepcopy(t0.rows[2]._tr)

for r in list(t0.rows)[1:]:
    t0._tbl.remove(r._tr)

skills_data = [
    ("CAT", "Programming Language"),
    ("Python", "5-Recognized authority", "5+ years", "2026", "Core language for AI/ML, Agent systems & Backend"),
    ("C++", "4-Applied theory", "5 years", "2026", "Deep learning inference, edge integration & algorithms"),
    ("TypeScript / JavaScript", "4-Applied theory", "4+ years", "2026", "Full-stack Next.js, React, Node.js applications"),
    ("SQL", "4-Applied theory", "5 years", "2026", "Relational queries, pgvector indexing & data modeling"),
    ("CAT", "AI, Machine Learning & Deep Learning"),
    ("PyTorch, TensorFlow", "5-Recognized authority", "5 years", "2026", "Deep learning research, model training & fine-tuning"),
    ("Computer Vision (YOLOv8/10, MediaPipe, ByteTrack, ViT)", "4-Applied theory", "4+ years", "2026", "Real-time object detection, tracking & pose analysis"),
    ("CAT", "Generative AI, Agents & RAG"),
    ("LLM Agent Frameworks (LangChain, LangGraph, Mastra, MCP)", "5-Recognized authority", "3+ years", "2026", "ReAct agents, tool-calling, multi-agent coordination"),
    ("Advanced RAG & Vector DBs (pgvector, Qdrant, Pinecone, Neo4j)", "4-Applied theory", "3+ years", "2026", "Hybrid search, Graph-RAG, 768-dim embeddings"),
    ("LLM Orchestration & Prompting (Gemini, OpenRouter, vLLM)", "4-Applied theory", "3+ years", "2026", "Multi-LLM fallback, prompt engineering, structured output"),
    ("CAT", "Backend, Distributed Systems & Database"),
    ("FastAPI, Node.js, REST APIs", "4-Applied theory", "4+ years", "2026", "High-throughput, asynchronous microservices"),
    ("Redis (AOF), Celery, WebSockets, WebRTC", "4-Applied theory", "3+ years", "2026", "Distributed worker queues, real-time streaming"),
    ("PostgreSQL / Supabase", "4-Applied theory", "4+ years", "2026", "Relational data, authentication, vector storage"),
    ("CAT", "DevOps & Cloud Deployment"),
    ("Docker & Docker Compose", "4-Applied theory", "4+ years", "2026", "Containerization, multi-stage builds, CI/CD"),
    ("Cloud Platforms (AWS, GCP, Azure, Hugging Face, Vercel, Nosana)", "4-Applied theory", "4+ years", "2026", "Cloud deployment, GPU serverless runtime"),
    ("CAT", "Foreign Language"),
    ("English", "4-Applied theory", "5+ years", "2026", "IELTS 6.5 equivalent (Professional working proficiency)"),
    ("Vietnamese", "5-Recognized authority", "Native", "2026", "Native language"),
]

for item in skills_data:
    if item[0] == "CAT":
        new_cat = copy.deepcopy(cat_tr)
        t0._tbl.append(new_cat)
        row = t0.rows[-1]
        row.cells[0].paragraphs[0].runs[0].text = item[1]
    else:
        new_data = copy.deepcopy(data_tr)
        t0._tbl.append(new_data)
        row = t0.rows[-1]
        for col_idx in range(5):
            c = row.cells[col_idx]
            val = item[col_idx]
            if c.paragraphs[0].runs:
                c.paragraphs[0].runs[0].text = val
                for r in c.paragraphs[0].runs[1:]:
                    r.text = ""
            else:
                c.paragraphs[0].text = val

# ==============================================================================
# 7. Table 1: Projects (Cloned Rows from Template)
# ==============================================================================
t1 = doc.tables[1]
t2 = doc.tables[2]

project_row_templates = [copy.deepcopy(t1.rows[i]._tr) for i in range(8)]

for r in list(t1.rows):
    t1._tbl.remove(r._tr)

doc._element.body.remove(t2._tbl)

projects_data = [
    {
        "name": "PiXity / LocalMate – Multi-Modal Tourism AI Agent Ecosystem",
        "duration": "03/2025 – Present",
        "position": "Senior AI Developer / AI Architect",
        "description": "Real-time Multi-modal Tourism & Conversational AI Assistant.\nApplication Type: Web & Autonomous AI Agent System.\nProgramming language: Python, TypeScript, Node.js.\nOS: Ubuntu Linux, Docker.\nDatabase: Neo4j, Pinecone, PostgreSQL, Redis.\nTechnology: FastAPI, Gemini API, OpenRouter, Graph-RAG, Docker CI/CD.",
        "type": "Product / Innovation (Won 2nd Runner-up at GDG DevFest MienTrung 2025).",
        "size": "4 members.",
        "responsibilities": [
            "Analyzed business and functional requirements to design the end-to-end multi-agent AI system architecture.",
            "Architected Graph-RAG pipeline combining Neo4j graph database and Pinecone vector search for spatial-temporal itinerary planning.",
            "Implemented multi-turn conversational agents with autonomous tool-calling capabilities using Gemini and OpenRouter APIs.",
            "Developed scalable asynchronous backend microservices using FastAPI and Node.js.",
            "Established containerized deployment workflows and CI/CD pipelines using Docker.",
            "Conducted performance profiling, issue troubleshooting, and latency optimization for real-time user interactions."
        ],
        "scope": "Requirement Analysis, Architecture Design, Coding, UT, IT, Deployment."
    },
    {
        "name": "Arcronics Edge AI & Cybersecurity Defense System",
        "duration": "06/2024 – 02/2025",
        "position": "Senior AI Engineer / Embedded AI Specialist",
        "description": "Malware Classification & Edge Cyber Attack Detection System.\nApplication Type: Deep Learning & Edge Computing.\nProgramming language: Python, C/C++, Shell Script.\nOS: Linux (Ubuntu), Embedded OS.\nDatabase: SQLite, PostgreSQL.\nTechnology: PyTorch, Scikit-learn, FPGA/Firmware Toolchains.",
        "type": "Enterprise Security Solution / R&D.",
        "size": "5 members.",
        "responsibilities": [
            "Researched, trained, and benchmarked deep learning models for malware classification and cyber attack detection.",
            "Designed automated MLOps pipelines for data preprocessing, feature extraction, model retraining, and continuous evaluation.",
            "Quantized, pruned, and integrated lightweight neural networks into hardware systems and microchips (Firmware, FPGA).",
            "Collaborated with hardware engineers to validate memory footprint and real-time execution constraints.",
            "Authored technical specifications, unit tests, and integration test documentation."
        ],
        "scope": "Research, Model Design, Hardware Integration, Coding, UT, IT."
    },
    {
        "name": "Manim Agent – Distributed Educational AI Storyboard & Video Platform",
        "duration": "10/2023 – 05/2024",
        "position": "Lead AI & Fullstack Architect",
        "description": "Automated Educational Storyboard Generation & Mathematical Video Rendering.\nApplication Type: Distributed Cloud AI Platform.\nProgramming language: Python, TypeScript.\nOS: Linux (Docker), Amazon Linux.\nDatabase: PostgreSQL (Supabase), Redis (AOF).\nTechnology: FastAPI, Celery, Redis, React SPA, Supabase Auth/Storage, Hugging Face Protected Spaces, Cairo/Pango.",
        "type": "R&D / Product.",
        "size": "2 members.",
        "responsibilities": [
            "Architected an end-to-end multi-service AI system that transforms educational ideas into structured storyboards, generates executable Manim Python scripts, and renders HD mathematical videos.",
            "Implemented a Human-in-the-Loop (HITL) review flow via WebSockets and real-time state synchronization for scene-by-scene video previews and interactive code iterations.",
            "Isolated untrusted Manim Python execution using asynchronous Celery worker queues and Redis AOF persistence, deployed within a protected Hugging Face Docker Space runtime.",
            "Built fullstack web interfaces with React and Supabase for cloud asset storage and authentication.",
            "Conducted end-to-end system testing, profiling, and deployment automation."
        ],
        "scope": "System Architecture, Fullstack Coding, Celery Sandbox Design, UT, IT, Cloud Deployment."
    },
    {
        "name": "MathSolver – Multi-Agent Geometry Problem Solver & Motion Renderer",
        "duration": "02/2023 – 09/2023",
        "position": "Senior AI Engineer",
        "description": "Automated Geometric Reasoning & 2D/3D Mathematical Motion Graphics.\nApplication Type: Multi-Agent AI System.\nProgramming language: Python, TypeScript.\nOS: Ubuntu Linux, Docker.\nDatabase: Supabase (PostgreSQL).\nTechnology: FastAPI, Next.js, OCR Engine, Manim Engine, OpenRouter API.",
        "type": "Product / R&D.",
        "size": "3 members.",
        "responsibilities": [
            "Developed a multi-agent system capable of extracting geometric shapes and equations from image inputs via OCR and generating step-by-step mathematical reasoning paths.",
            "Designed an automated code generation pipeline that translates proof steps into 2D/3D Manim motion graphics with multi-LLM fallback orchestration over OpenRouter APIs.",
            "Containerized backend services with Docker and Cairo/Pango system bindings to deliver real-time video compilation and interactive step visualizers.",
            "Optimized inference latency and handled edge cases in geometric parsing."
        ],
        "scope": "Algorithm Design, Multi-Agent Engineering, Coding, UT, IT, Deployment."
    },
    {
        "name": "Medagen – AI Patient Triage & Diagnostic Assistant",
        "duration": "07/2022 – 01/2023",
        "position": "Fullstack AI Developer",
        "description": "Clinical Symptom Triage, Differential Diagnosis & Multimodal Verification.\nApplication Type: Healthcare Decision Support System.\nProgramming language: Python, TypeScript.\nOS: Ubuntu Linux.\nDatabase: Supabase (pgvector), Qdrant.\nTechnology: LangChain ReAct Agents, Gemini API, Vision Transformers (ViT), CNNs, FastAPI.",
        "type": "Healthcare AI Product / R&D.",
        "size": "6 members.",
        "responsibilities": [
            "Built a clinical diagnostic assistant utilizing LangChain ReAct Agent framework and Gemini for multi-turn symptom triage, differential diagnosis, and medical tool-calling.",
            "Architected an Advanced RAG system over medical knowledge bases using 768-dim vector embeddings (Supabase Vector / Qdrant), hybrid search, and semantic document chunking.",
            "Integrated Computer Vision models (Vision Transformers / custom CNNs) for dermatological lesion classification and multimodal symptom verification.",
            "Established rigorous testing protocols for clinical safety and diagnostic accuracy."
        ],
        "scope": "Requirement Analysis, Architecture Design, Model Fine-tuning, Coding, UT, IT."
    },
    {
        "name": "MagicCV – Autonomous Resume Tailoring AI Agent",
        "duration": "01/2022 – 06/2022",
        "position": "AI Engineer",
        "description": "Autonomous Profile Extraction & Dynamic Resume Synthesis.\nApplication Type: Autonomous AI Agent Platform.\nProgramming language: Python, TypeScript.\nOS: Linux (Docker).\nDatabase: Supabase (pgvector).\nTechnology: Mastra Agent Framework, Gemini LLM, pdflatex Compiler, Nosana Decentralized GPU Network.",
        "type": "Product / Innovation (Finalist at Nosana Builders Challenge: Agents 102).",
        "size": "2 members.",
        "responsibilities": [
            "Created an autonomous CV tailoring agent that extracts candidate profiles from GitHub, LinkedIn, and YouTube APIs.",
            "Implemented 768-dim vector similarity search with 3-level fallback strategies and Gemini LLM ranking algorithms to match career data with job descriptions in 3 seconds.",
            "Rendered publication-quality PDFs via Nunjucks LaTeX templating pipeline.",
            "Containerized the application with Docker and deployed to Nosana decentralized GPU compute network."
        ],
        "scope": "Agent Architecture, Prompt Engineering, LaTeX Pipeline, Coding, UT, IT, Deployment."
    },
    {
        "name": "Lab Management – Real-Time AI Computer Vision Dashboard",
        "duration": "05/2021 – 12/2021",
        "position": "Computer Vision & Software Engineer",
        "description": "Indoor Occupancy Tracking & Crowd Flow Analytics.\nApplication Type: Computer Vision & Edge Analytics Dashboard.\nProgramming language: Python, TypeScript, C++.\nOS: Ubuntu Linux.\nDatabase: PostgreSQL, Redis.\nTechnology: YOLOv8, ByteTrack, Janus WebRTC, WebSockets, Next.js.",
        "type": "Enterprise Solution / IoT.",
        "size": "5 members.",
        "responsibilities": [
            "Deployed YOLO object detection and ByteTrack multi-object tracking models for real-time indoor occupancy tracking and crowd flow analytics.",
            "Streamed live video telemetry with Janus WebRTC servers and WebSockets, delivering low-latency spatial tracking analytics to Next.js dashboards.",
            "Optimized edge computer vision inference pipelines to achieve high frame rates (30+ FPS) under heavy video streams.",
            "Built interactive spatial heatmaps and time-series occupancy charts on the frontend."
        ],
        "scope": "Computer Vision Pipeline, Streaming Backend, UI Coding, UT, IT."
    }
]

for prj in projects_data:
    for r_tr in project_row_templates:
        t1._tbl.append(copy.deepcopy(r_tr))
    
    curr_rows = t1.rows[-8:]
    
    # Row 0: Name and Duration
    curr_rows[0].cells[1].paragraphs[0].runs[0].text = prj["name"]
    curr_rows[0].cells[2].paragraphs[0].runs[0].text = "Duration "
    
    # Row 1: Position and Duration
    curr_rows[1].cells[1].paragraphs[0].runs[0].text = prj["position"]
    curr_rows[1].cells[2].paragraphs[0].runs[0].text = prj["duration"]
    
    # Row 2: Separator (keep as is)
    
    # Row 3: Description
    c3 = curr_rows[3].cells[1]
    p_desc_tpl = copy.deepcopy(c3.paragraphs[0]._element)
    for p in list(c3.paragraphs):
        c3._tc.remove(p._element)
    for line in prj["description"].split('\n'):
        new_p_elem = copy.deepcopy(p_desc_tpl)
        c3._tc.append(new_p_elem)
        p_obj = docx.text.paragraph.Paragraph(new_p_elem, doc)
        p_obj.runs[0].text = line
        if len(p_obj.runs) > 1:
            for r in p_obj.runs[1:]:
                r.text = ""
                
    # Row 4: Type
    c4 = curr_rows[4].cells[1]
    c4.paragraphs[0].runs[0].text = prj["type"]
    for p in list(c4.paragraphs)[1:]:
        c4._tc.remove(p._element)
        
    # Row 5: Size
    c5 = curr_rows[5].cells[1]
    c5.paragraphs[0].runs[0].text = prj["size"]
    for p in list(c5.paragraphs)[1:]:
        c5._tc.remove(p._element)
        
    # Row 6: Responsibilities
    c6 = curr_rows[6].cells[1]
    p_resp_tpl = copy.deepcopy(c6.paragraphs[0]._element)
    for p in list(c6.paragraphs):
        c6._tc.remove(p._element)
    for resp in prj["responsibilities"]:
        new_p_elem = copy.deepcopy(p_resp_tpl)
        c6._tc.append(new_p_elem)
        p_obj = docx.text.paragraph.Paragraph(new_p_elem, doc)
        p_obj.runs[0].text = resp
        if len(p_obj.runs) > 1:
            for r in p_obj.runs[1:]:
                r.text = ""
                
    # Row 7: Scope
    c7 = curr_rows[7].cells[1]
    c7.paragraphs[0].runs[0].text = prj["scope"]
    for p in list(c7.paragraphs)[1:]:
        c7._tc.remove(p._element)

doc.save(OUTPUT_PATH)
print(f"Generated successfully: {OUTPUT_PATH}")
