import docx
from docx.shared import Pt, Inches, RGBColor
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn
from docx.enum.text import WD_ALIGN_PARAGRAPH

def set_cell_borders_and_margins(cell, top="808080", bottom="808080", left="808080", right="808080", sz="4", fill=None):
    tcPr = cell._tc.get_or_add_tcPr()
    borders_xml = f'''
    <w:tcBorders {nsdecls("w")}>
        <w:top w:val="{"single" if top else "none"}" w:sz="{sz}" w:space="0" w:color="{top if top else "auto"}"/>
        <w:left w:val="{"single" if left else "none"}" w:sz="{sz}" w:space="0" w:color="{left if left else "auto"}"/>
        <w:bottom w:val="{"single" if bottom else "none"}" w:sz="{sz}" w:space="0" w:color="{bottom if bottom else "auto"}"/>
        <w:right w:val="{"single" if right else "none"}" w:sz="{sz}" w:space="0" w:color="{right if right else "auto"}"/>
    </w:tcBorders>
    '''
    for existing in tcPr.findall(qn('w:tcBorders')):
        tcPr.remove(existing)
    tcPr.append(parse_xml(borders_xml))
    
    if fill:
        for existing in tcPr.findall(qn('w:shd')):
            tcPr.remove(existing)
        shd_xml = f'<w:shd {nsdecls("w")} w:val="clear" w:fill="{fill}"/>'
        tcPr.append(parse_xml(shd_xml))
        
    for existing in tcPr.findall(qn('w:tcMar')):
        tcPr.remove(existing)
    tcMar_xml = f'''
    <w:tcMar {nsdecls("w")}>
        <w:top w:w="80" w:type="dxa"/>
        <w:left w:w="120" w:type="dxa"/>
        <w:bottom w:w="80" w:type="dxa"/>
        <w:right w:w="120" w:type="dxa"/>
    </w:tcMar>
    '''
    tcPr.append(parse_xml(tcMar_xml))

def format_run(run, font_name="Calibri", size_pt=11, bold=False, italic=False, color_rgb=(0,0,0)):
    run.font.name = font_name
    if size_pt:
        run.font.size = Pt(size_pt)
    run.font.bold = bold
    run.font.italic = italic
    if color_rgb:
        run.font.color.rgb = RGBColor(*color_rgb)

def add_p(doc, text="", bold_prefix=None, bold=False, size_pt=11, space_before=0, space_after=60, line_spacing=1.15, align=WD_ALIGN_PARAGRAPH.LEFT):
    p = doc.add_paragraph()
    p.alignment = align
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing = line_spacing
    
    if bold_prefix:
        r1 = p.add_run(bold_prefix)
        format_run(r1, size_pt=size_pt, bold=True)
    if text:
        r2 = p.add_run(text)
        format_run(r2, size_pt=size_pt, bold=bold)
    return p

def add_heading_p(doc, text, size_pt=11, space_before=140, space_after=60):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing = 1.15
    p.paragraph_format.keep_with_next = True
    r = p.add_run(text)
    format_run(r, size_pt=size_pt, bold=True)
    return p

# Load template doc
doc = docx.Document('downloaded.docx')
body = doc._element.body

# Remove existing body children except sectPr
children = list(body)
for c in children:
    if not c.tag.endswith('sectPr'):
        body.remove(c)

# 1. Candidate Name (Header)
p_title = add_p(doc, "NGUYEN XUAN CUONG", bold=True, size_pt=16, space_before=100, space_after=140, align=WD_ALIGN_PARAGRAPH.LEFT)

# 2. PERSONAL DETAILS
add_heading_p(doc, "PERSONAL DETAILS")
add_p(doc, "Nguyen Xuan Cuong", bold_prefix="Name\t\t")
add_p(doc, "Vietnamese", bold_prefix="Nationality \t")
add_p(doc, "Male", bold_prefix="Sex\t\t")
add_p(doc, "Da Nang, Vietnam", bold_prefix="Location\t")
add_p(doc, "nxcforwork@gmail.com", bold_prefix="Email\t\t")
add_p(doc, "+84 327 755 183", bold_prefix="Phone\t\t")
add_p(doc, "GitHub: nxc1802  |  LinkedIn: cuong-nguyen-2715a136b  |  Portfolio: nxc-s-portfolio.vercel.app", bold_prefix="Profiles\t")

# 3. EDUCATIONAL BACKGROUND
add_heading_p(doc, "EDUCATIONAL BACKGROUND")
add_p(doc, "FPT University, Da Nang, Vietnam – Bachelor of Science in Information Technology (AI Major).")
add_p(doc, "Recipient of 100% Full-tuition FPT University Scholarship.  |  GPA: 8.28 / 10.00")

# 4. CERTIFICATE AND AWARD
add_heading_p(doc, "CERTIFICATE AND AWARD")
add_p(doc, "Professional working proficiency in technical research, architectural writing, and global collaboration.", bold_prefix="English: IELTS 6.5 Equivalent – ")
add_p(doc, "Four-year merit-based full academic scholarship covering 100% tuition fees.", bold_prefix="Recipient of 100% Full-tuition Scholarship – FPT University: ")
add_p(doc, "Hackathon GDG DevFest MienTrung 2025 (PiXity Multi-modal Tourism AI Agent).", bold_prefix="2nd Runner-up: ")
add_p(doc, "AI4SE 2025 Competition (Artificial Intelligence for Software Engineering).", bold_prefix="Second Prize: ")
add_p(doc, "Nosana Builders Challenge: Agents 102 (MagicCV Autonomous Resume Tailoring Agent).", bold_prefix="4th Place: ")
add_p(doc, "Timelocked Autonomous Multi-signature Smart Contract Wallet (2025).", bold_prefix="Superteam Honorable Mention: ")
add_p(doc, "AI4LIFE 2024, Research Connection 2025, Research Festival 2025, Grab The Beyond 2026.", bold_prefix="Finalist: ")

# 5. RESEARCH & PUBLICATIONS
add_heading_p(doc, "RESEARCH & PUBLICATIONS")
add_p(doc, "Multi-Agent Imitation by learning and sampling from FactorIzed Soft Q-function.", bold_prefix="• MAFIS: ")
add_p(doc, "Using CNN and SVM architectures for precision agriculture and disease diagnosis.", bold_prefix="• Identification of Plant Leaf Diseases: ")
add_p(doc, "2D/3D human kinematic pose analysis using CNN-LSTM and MediaPipe.", bold_prefix="• The GYM Exercise Classification Dataset: ")
add_p(doc, "Gradient-Guided Approach for Efficient Adversarial Image Attacks.", bold_prefix="• Towards Sparse Adversarial Perturbations: ")

# 6. PROFESSIONAL SUMMARY
add_heading_p(doc, "PROFESSIONAL SUMMARY")
add_p(doc, "Senior AI & Software Engineer with 5+ years of software development and AI engineering experience. Proven track record in researching, designing, and delivering production-grade Generative AI agents, Advanced RAG architectures, and real-time Computer Vision systems.")
add_p(doc, "Specializes in:")
add_p(doc, "Generative AI & Autonomous Agent Orchestration (5+ years in Software/AI, 3+ years in LangChain, LangGraph, Mastra, Model Context Protocol - MCP).", bold_prefix="• ")
add_p(doc, "Advanced RAG & Knowledge Graphs (Graph-RAG, Neo4j, Qdrant, Pinecone, Supabase pgvector, hybrid semantic search).", bold_prefix="• ")
add_p(doc, "Computer Vision & Edge Deep Learning (YOLOv8/v10, MediaPipe, ByteTrack, Vision Transformers, FPGA & firmware integration).", bold_prefix="• ")
add_p(doc, "High-performance Asynchronous Backend & Distributed Systems (FastAPI, Python, Node.js, Celery, Redis, WebSockets, WebRTC).", bold_prefix="• ")
add_p(doc, "Proficient in cloud environments and scalable deployment (AWS, GCP, Azure, Hugging Face Docker Spaces, Vercel, Supabase, Nosana Decentralized GPU Network).")
add_p(doc, "Demonstrated technical leadership abilities in designing end-to-end architectures, preparing technical proposals, and effort estimations.")
add_p(doc, "Hands-on experience leveraging AI-powered IDEs (Cursor, Copilot) for rapid prototyping, architecture refactoring, and maximizing development velocity.")
add_p(doc, "Strong communicator with professional English proficiency (IELTS 6.5 equivalent), engaging effectively with stakeholders, clients, and cross-functional teams.")
add_p(doc, "Committed to clean architecture, code quality, test-driven development, and industry best practices.")

# 7. SOFTWARE
add_heading_p(doc, "SOFTWARE")
add_p(doc, "PostgreSQL (pgvector), Supabase, Neo4j, Qdrant, Pinecone, Redis (AOF), Microsoft SQL Server, SQLite.", bold_prefix="Databases: ")
add_p(doc, "Python (Core/Expert), C++, TypeScript, JavaScript, SQL, Shell Script.", bold_prefix="Programming Languages: ")
add_p(doc, "PyTorch, TensorFlow, LangChain, LangGraph, Mastra, FastAPI, Celery, Node.js, Next.js, React, vLLM, Ollama, YOLOv8/v10, MediaPipe.", bold_prefix="Frameworks & Platforms: ")
add_p(doc, "AWS, Azure, Google Cloud Platform (GCP), Hugging Face Spaces, Vercel, Supabase, Docker, Docker Compose, Nosana GPU Network.", bold_prefix="Cloud Platforms & Deployment: ")
add_p(doc, "VS Code, Cursor AI IDE, Docker, Git, Postman, Janus WebRTC.", bold_prefix="Programming Tools: ")
add_p(doc, "GIT (GitHub, GitLab).", bold_prefix="Source Control: ")
add_p(doc, "Agile/Scrum, Waterfall.", bold_prefix="Software Process: ")
add_p(doc, "Object Oriented Analysis & Design (OOAD), Multi-Agent System Architecture, Advanced RAG & Graph-RAG, Microservices.", bold_prefix="OOAD/OOP & Architecture: ")

# 8. TECHNICAL EXPERTISE & SKILLS
add_heading_p(doc, "TECHNICAL EXPERTISE & SKILLS")
add_p(doc, "Proficiency Description", bold=True, space_after=30)
add_p(doc, "Basic knowledge, needs much guidance and support in using the competence", bold_prefix="1-Basic knowledge: ", space_after=20)
add_p(doc, "Able to work in the area, needs some guidance in using the competence", bold_prefix="2-Limited experience: ", space_after=20)
add_p(doc, "Good knowledge and skills, works independently, has sound experience in using the competence", bold_prefix="3-Practical application: ", space_after=20)
add_p(doc, "Advanced knowledge and skills, provide guidance and direction to others", bold_prefix="4-Applied theory: ", space_after=20)
add_p(doc, "Exceptional knowledge and skills, perceived as an authority in this competence area, is innovative in using the competence", bold_prefix="5-Recognized authority: ", space_after=60)
add_p(doc, "Technology skill", bold=True, space_after=60)

# Table 0: Technology skill table
skills_data = [
    # Category: Programming Language
    ("CAT", "Programming Language"),
    ("Python", "5-Recognized authority", "5+ years", "2026", "Core language for AI/ML, Agent systems & Backend"),
    ("C++", "4-Applied theory", "5 years", "2026", "Deep learning inference, edge integration & algorithms"),
    ("TypeScript / JavaScript", "4-Applied theory", "4+ years", "2026", "Full-stack Next.js, React, Node.js applications"),
    ("SQL", "4-Applied theory", "5 years", "2026", "Relational queries, pgvector indexing & data modeling"),
    # Category: AI, Machine Learning & Deep Learning
    ("CAT", "AI, Machine Learning & Deep Learning"),
    ("PyTorch, TensorFlow", "5-Recognized authority", "5 years", "2026", "Deep learning research, model training & fine-tuning"),
    ("Computer Vision (YOLOv8/10, MediaPipe, ByteTrack, ViT)", "4-Applied theory", "4+ years", "2026", "Real-time object detection, tracking & pose analysis"),
    # Category: Generative AI, Agents & RAG
    ("CAT", "Generative AI, Agents & RAG"),
    ("LLM Agent Frameworks (LangChain, LangGraph, Mastra, MCP)", "5-Recognized authority", "3+ years", "2026", "ReAct agents, tool-calling, multi-agent coordination"),
    ("Advanced RAG & Vector DBs (pgvector, Qdrant, Pinecone, Neo4j)", "4-Applied theory", "3+ years", "2026", "Hybrid search, Graph-RAG, 768-dim embeddings"),
    ("LLM Orchestration & Prompting (Gemini, OpenRouter, vLLM)", "4-Applied theory", "3+ years", "2026", "Multi-LLM fallback, prompt engineering, structured output"),
    # Category: Backend, Distributed Systems & Database
    ("CAT", "Backend, Distributed Systems & Database"),
    ("FastAPI, Node.js, REST APIs", "4-Applied theory", "4+ years", "2026", "High-throughput, asynchronous microservices"),
    ("Redis (AOF), Celery, WebSockets, WebRTC", "4-Applied theory", "3+ years", "2026", "Distributed worker queues, real-time streaming"),
    ("PostgreSQL / Supabase", "4-Applied theory", "4+ years", "2026", "Relational data, authentication, vector storage"),
    # Category: DevOps & Cloud Deployment
    ("CAT", "DevOps & Cloud Deployment"),
    ("Docker & Docker Compose", "4-Applied theory", "4+ years", "2026", "Containerization, multi-stage builds, CI/CD"),
    ("Cloud Platforms (AWS, GCP, Azure, Hugging Face, Vercel, Nosana)", "4-Applied theory", "4+ years", "2026", "Cloud deployment, GPU serverless runtime"),
    # Category: Foreign Language
    ("CAT", "Foreign Language"),
    ("English", "4-Applied theory", "5+ years", "2026", "IELTS 6.5 equivalent (Professional working proficiency)"),
    ("Vietnamese", "5-Recognized authority", "Native", "2026", "Native language"),
]

t0 = doc.add_table(rows=0, cols=5)
t0.autofit = False

col_widths_t0 = [Pt(128.35), Pt(115.5), Pt(65.85), Pt(60.8), Pt(101.65)]

# Header row
hdr_row = t0.add_row()
hdr_titles = ["Skill", "Proficiency", "Experience", "Last used", "Notes"]
for i, title in enumerate(hdr_titles):
    c = hdr_row.cells[i]
    c.width = col_widths_t0[i]
    set_cell_borders_and_margins(c, fill="bfbfbf")
    p = c.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER if i in [1, 2, 3] else WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.space_before = Pt(20)
    p.paragraph_format.space_after = Pt(20)
    r = p.add_run(title)
    format_run(r, bold=True, size_pt=10.5)

for item in skills_data:
    row = t0.add_row()
    if item[0] == "CAT":
        # Merge all 5 cells into cell 0
        c0 = row.cells[0]
        c0.merge(row.cells[4])
        set_cell_borders_and_margins(c0, fill=None)
        p = c0.paragraphs[0]
        p.paragraph_format.space_before = Pt(30)
        p.paragraph_format.space_after = Pt(30)
        r = p.add_run(item[1])
        format_run(r, bold=True, size_pt=10.5)
    else:
        for i, val in enumerate(item):
            c = row.cells[i]
            c.width = col_widths_t0[i]
            set_cell_borders_and_margins(c, fill=None)
            p = c.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if i in [2, 3] else WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.space_before = Pt(20)
            p.paragraph_format.space_after = Pt(20)
            r = p.add_run(val)
            format_run(r, size_pt=10, bold=(i==0))

# 9. PROFESSIONAL EXPERIENCE
add_heading_p(doc, "PROFESSIONAL EXPERIENCE", space_before=200, space_after=80)

# Projects list
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

t1 = doc.add_table(rows=0, cols=3)
t1.autofit = False

col_w_0 = Pt(100.3)
col_w_1 = Pt(245.05)
col_w_2 = Pt(127.15)
col_w_merged_12 = Pt(372.2)

for p_idx, prj in enumerate(projects_data):
    # Row 0: Project name (span 1) | name (span 1) | Duration (span 1)
    r0 = t1.add_row()
    c00 = r0.cells[0]; c00.width = col_w_0; set_cell_borders_and_margins(c00)
    p = c00.paragraphs[0]; p.paragraph_format.space_before = Pt(20); p.paragraph_format.space_after = Pt(20)
    format_run(p.add_run("Project name"), bold=True, size_pt=10.5)

    c01 = r0.cells[1]; c01.width = col_w_1; set_cell_borders_and_margins(c01)
    p = c01.paragraphs[0]; p.paragraph_format.space_before = Pt(20); p.paragraph_format.space_after = Pt(20)
    format_run(p.add_run(prj["name"]), bold=True, size_pt=10.5)

    c02 = r0.cells[2]; c02.width = col_w_2; set_cell_borders_and_margins(c02)
    p = c02.paragraphs[0]; p.paragraph_format.space_before = Pt(20); p.paragraph_format.space_after = Pt(20)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    format_run(p.add_run("Duration"), bold=True, size_pt=10.5)

    # Row 1: Position(s) (span 1) | position (span 1) | duration val (span 1)
    r1 = t1.add_row()
    c10 = r1.cells[0]; c10.width = col_w_0; set_cell_borders_and_margins(c10)
    p = c10.paragraphs[0]; p.paragraph_format.space_before = Pt(20); p.paragraph_format.space_after = Pt(20)
    format_run(p.add_run("Position(s)"), bold=True, size_pt=10.5)

    c11 = r1.cells[1]; c11.width = col_w_1; set_cell_borders_and_margins(c11)
    p = c11.paragraphs[0]; p.paragraph_format.space_before = Pt(20); p.paragraph_format.space_after = Pt(20)
    format_run(p.add_run(prj["position"]), bold=False, size_pt=10.5)

    c12 = r1.cells[2]; c12.width = col_w_2; set_cell_borders_and_margins(c12)
    p = c12.paragraphs[0]; p.paragraph_format.space_before = Pt(20); p.paragraph_format.space_after = Pt(20)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    format_run(p.add_run(prj["duration"]), bold=False, size_pt=10.5)

    # Row 2: Empty separator row (merge col 0 to col 2)
    r2 = t1.add_row()
    trPr = r2._tr.get_or_add_trPr()
    trPr.append(parse_xml(f'<w:trHeight {nsdecls("w")} w:val="143" w:hRule="atLeast"/>'))
    c20 = r2.cells[0]
    c20.merge(r2.cells[2])
    set_cell_borders_and_margins(c20, top="808080", bottom=None, left=None, right=None)
    p = c20.paragraphs[0]; p.paragraph_format.space_before = Pt(0); p.paragraph_format.space_after = Pt(0)
    format_run(p.add_run(""), size_pt=2)

    # Helper for span-2 content rows
    def add_span2_row(label, content, is_bullet_list=False):
        row = t1.add_row()
        cell_lbl = row.cells[0]; cell_lbl.width = col_w_0; set_cell_borders_and_margins(cell_lbl)
        p = cell_lbl.paragraphs[0]; p.paragraph_format.space_before = Pt(20); p.paragraph_format.space_after = Pt(20)
        format_run(p.add_run(label), bold=True, size_pt=10.5)

        cell_val = row.cells[1]
        cell_val.merge(row.cells[2])
        set_cell_borders_and_margins(cell_val)

        if is_bullet_list:
            p_first = cell_val.paragraphs[0]
            for b_idx, bullet in enumerate(content):
                p_b = p_first if b_idx == 0 else cell_val.add_paragraph()
                p_b.paragraph_format.space_before = Pt(10)
                p_b.paragraph_format.space_after = Pt(10)
                p_b.paragraph_format.line_spacing = 1.15
                format_run(p_b.add_run("•  "), bold=True, size_pt=10)
                format_run(p_b.add_run(bullet), bold=False, size_pt=10)
        else:
            lines = content.split('\n')
            p_first = cell_val.paragraphs[0]
            for l_idx, line in enumerate(lines):
                p_l = p_first if l_idx == 0 else cell_val.add_paragraph()
                p_l.paragraph_format.space_before = Pt(10)
                p_l.paragraph_format.space_after = Pt(10)
                p_l.paragraph_format.line_spacing = 1.15
                format_run(p_l.add_run(line), bold=False, size_pt=10)

    # Row 3: Project description
    add_span2_row("Project description", prj["description"])

    # Row 4: Project type
    add_span2_row("Project type", prj["type"])

    # Row 5: Team size
    add_span2_row("Team size", prj["size"])

    # Row 6: Responsibilities
    add_span2_row("Responsibilities", prj["responsibilities"], is_bullet_list=True)

    # Row 7: Project scope
    add_span2_row("Project scope", prj["scope"])

output_path = "/Volumes/WorkSpace/Project/Resume/BSOL_Senior-AI-Engineer_Nguyen-Xuan-Cuong.docx"
doc.save(output_path)
print(f"Generated successfully: {output_path}")
