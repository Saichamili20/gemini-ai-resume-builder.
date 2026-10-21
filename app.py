import json
import os
import tempfile

import gradio as gr
from google import genai
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

# Configure GEMINI_API_KEY in Render's Environment settings. Never hard-code it here.
API_KEY = os.environ.get("GEMINI_API_KEY")
client = genai.Client(api_key=API_KEY) if API_KEY else None

resume_instructions = """
You are an expert professional resume writer and ATS resume optimizer.
Convert the candidate's raw information into a professional one-page or two-page ATS-friendly resume.
Never invent companies, job titles, internships, projects, certifications, achievements, dates, technologies,
responsibilities, metrics, percentages, or statistics. Use only information explicitly provided by the candidate.
Improve grammar and professional wording while preserving all factual information. Include all relevant experience,
projects, certifications, and achievements. Use standard ATS-friendly section names. Do not use tables, columns,
emojis, or Markdown formatting inside JSON values. Do not duplicate information or create placeholders. Do not claim
experience with a technology unless it appears in the supplied information. Empty sections should be empty arrays.
Create a concise professional summary based on the candidate's actual education, skills, experience, and career goal.
"""

resume_schema = {
    "type": "object",
    "properties": {
        "professional_summary": {"type": "string"},
        "skills": {"type": "array", "items": {"type": "string"}},
        "education": {"type": "array", "items": {"type": "object", "properties": {
            "degree": {"type": "string"}, "institution": {"type": "string"},
            "dates": {"type": "string"}, "details": {"type": "array", "items": {"type": "string"}}
        }, "required": ["degree", "institution", "dates", "details"]}},
        "experience": {"type": "array", "items": {"type": "object", "properties": {
            "job_title": {"type": "string"}, "company": {"type": "string"},
            "dates": {"type": "string"}, "location": {"type": "string"},
            "bullets": {"type": "array", "items": {"type": "string"}}
        }, "required": ["job_title", "company", "dates", "location", "bullets"]}},
        "projects": {"type": "array", "items": {"type": "object", "properties": {
            "name": {"type": "string"}, "technologies": {"type": "string"},
            "bullets": {"type": "array", "items": {"type": "string"}}
        }, "required": ["name", "technologies", "bullets"]}},
        "certifications": {"type": "array", "items": {"type": "string"}},
        "achievements": {"type": "array", "items": {"type": "string"}},
    },
    "required": ["professional_summary", "skills", "education", "experience", "projects", "certifications", "achievements"]
}


def clean_text(text):
    if not text:
        return ""
    return str(text).replace("**", "").replace("__", "").replace("\\@", "@").replace("\\|", "|").strip()


def create_resume_docx(resume_data, candidate):
    doc = Document()
    section = doc.sections[0]
    section.top_margin = Inches(0.5)
    section.bottom_margin = Inches(0.5)
    section.left_margin = Inches(0.65)
    section.right_margin = Inches(0.65)

    def add_heading(text):
        if not text:
            return
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(6)
        p.paragraph_format.space_after = Pt(2)
        run = p.add_run(clean_text(text).upper())
        run.bold = True
        run.font.name = "Arial"
        run.font.size = Pt(11)

    def add_bullet(text):
        text = clean_text(text)
        if not text:
            return
        p = doc.add_paragraph(style="List Bullet")
        p.paragraph_format.left_indent = Inches(0.2)
        p.paragraph_format.first_line_indent = Inches(-0.1)
        p.paragraph_format.space_after = Pt(1)
        run = p.add_run(text)
        run.font.name = "Arial"
        run.font.size = Pt(9.5)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(clean_text(candidate.get("name", "")))
    run.bold = True
    run.font.name = "Arial"
    run.font.size = Pt(17)

    contact = [clean_text(candidate.get(k, "")) for k in ("email", "phone", "location", "linkedin", "github")]
    contact = [value for value in contact if value]
    if contact:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(" | ".join(contact))
        run.font.name = "Arial"
        run.font.size = Pt(8.5)

    summary = clean_text(resume_data.get("professional_summary", ""))
    if summary:
        add_heading("Professional Summary")
        p = doc.add_paragraph(summary)
        p.paragraph_format.space_after = Pt(2)
        for run in p.runs:
            run.font.name = "Arial"
            run.font.size = Pt(9.5)

    skills = [clean_text(s) for s in resume_data.get("skills", []) if clean_text(s)]
    if skills:
        add_heading("Skills")
        p = doc.add_paragraph(" | ".join(skills))
        for run in p.runs:
            run.font.name = "Arial"
            run.font.size = Pt(9.5)

    if resume_data.get("education"):
        add_heading("Education")
        for edu in resume_data["education"]:
            parts = [clean_text(edu.get(k, "")) for k in ("degree", "institution", "dates")]
            parts = [part for part in parts if part]
            if parts:
                p = doc.add_paragraph()
                p.paragraph_format.space_after = Pt(1)
                run = p.add_run(" | ".join(parts))
                run.bold = True
                run.font.name = "Arial"
                run.font.size = Pt(9.5)
            for detail in edu.get("details", []):
                add_bullet(detail)

    if resume_data.get("experience"):
        add_heading("Experience")
        for exp in resume_data["experience"]:
            parts = [clean_text(exp.get(k, "")) for k in ("job_title", "company", "dates", "location")]
            parts = [part for part in parts if part]
            if parts:
                p = doc.add_paragraph()
                p.paragraph_format.space_after = Pt(1)
                run = p.add_run(" | ".join(parts))
                run.bold = True
                run.font.name = "Arial"
                run.font.size = Pt(9.5)
            for bullet in exp.get("bullets", []):
                add_bullet(bullet)

    if resume_data.get("projects"):
        add_heading("Projects")
        for project in resume_data["projects"]:
            project_name = clean_text(project.get("name", ""))
            technologies = clean_text(project.get("technologies", ""))
            if project_name:
                p = doc.add_paragraph()
                p.paragraph_format.space_after = Pt(1)
                run = p.add_run(project_name)
                run.bold = True
                run.font.name = "Arial"
                run.font.size = Pt(9.5)
                if technologies:
                    run = p.add_run(" | " + technologies)
                    run.italic = True
                    run.font.name = "Arial"
                    run.font.size = Pt(9)
            for bullet in project.get("bullets", []):
                add_bullet(bullet)

    if resume_data.get("certifications"):
        add_heading("Certifications")
        for item in resume_data["certifications"]:
            add_bullet(item)
    if resume_data.get("achievements"):
        add_heading("Achievements")
        for item in resume_data["achievements"]:
            add_bullet(item)
    return doc


def generate_resume_app(name, email, phone, location, linkedin, github,
                        career_goal, education, skills, experience,
                        projects, certifications, achievements):
    if not name.strip() or not education.strip() or not skills.strip():
        raise gr.Error("Please enter your name, education, and skills.")
    if client is None:
        raise gr.Error("The Gemini API key is not configured. Add GEMINI_API_KEY in Render's Environment settings.")

    candidate = {
        "name": name.strip(), "email": email.strip(), "phone": phone.strip(),
        "location": location.strip(), "linkedin": linkedin.strip(), "github": github.strip(),
        "career_goal": career_goal.strip(), "education": education.strip(), "skills": skills.strip(),
        "experience": experience.strip(), "projects": projects.strip(),
        "certifications": certifications.strip(), "achievements": achievements.strip()
    }
    prompt = f"{resume_instructions}\n\nCANDIDATE INFORMATION:\n{json.dumps(candidate, ensure_ascii=False, indent=2)}\n\nReturn JSON matching the supplied schema. Use only the supplied candidate information."

    try:
        response = client.models.generate_content(
            model="gemini-2.5-flash",
            contents=prompt,
            config={"response_mime_type": "application/json", "response_schema": resume_schema, "temperature": 0.1},
        )
        resume_data = json.loads(response.text)
        doc = create_resume_docx(resume_data, candidate)
        with tempfile.NamedTemporaryFile(suffix=".docx", prefix="ATS_Resume_", delete=False) as temp_file:
            docx_path = temp_file.name
        doc.save(docx_path)
        return "Resume generated successfully!", docx_path
    except gr.Error:
        raise
    except Exception as exc:
        print(f"Resume generation error: {exc}")
        raise gr.Error("Could not generate the resume. Check the Gemini API key and quota, then try again.")


with gr.Blocks(title="AI Resume Builder") as demo:
    gr.Markdown("# AI Resume Builder")
    gr.Markdown("Enter your details to generate an ATS-friendly resume. Only provide information you want included.")
    with gr.Row():
        name = gr.Textbox(label="Full Name")
        email = gr.Textbox(label="Email")
        phone = gr.Textbox(label="Phone")
    with gr.Row():
        location = gr.Textbox(label="Location")
        linkedin = gr.Textbox(label="LinkedIn URL")
        github = gr.Textbox(label="GitHub URL")
    career_goal = gr.Textbox(label="Career Goal", lines=2)
    education = gr.Textbox(label="Education", lines=3)
    skills = gr.Textbox(label="Skills", lines=3)
    experience = gr.Textbox(label="Internships / Experience", lines=4)
    projects = gr.Textbox(label="Projects", lines=4)
    certifications = gr.Textbox(label="Certifications", lines=3)
    achievements = gr.Textbox(label="Achievements", lines=3)
    generate_button = gr.Button("Generate My Resume", variant="primary")
    status = gr.Textbox(label="Status", interactive=False)
    download = gr.File(label="Download Your Resume")
    generate_button.click(
        fn=generate_resume_app,
        inputs=[name, email, phone, location, linkedin, github, career_goal, education, skills,
                experience, projects, certifications, achievements],
        outputs=[status, download],
    )

if __name__ == "__main__":
    port = int(os.environ.get("PORT", "10000"))
    demo.launch(server_name="0.0.0.0", server_port=port, share=False)
