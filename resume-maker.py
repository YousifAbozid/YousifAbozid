from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT, WD_LINE_SPACING

# Start a new doc with corrected content
doc = Document()

style = doc.styles['Normal']
font = style.font
font.name = 'Arial'
font.size = Pt(10.5)

# Set compact spacing for Normal style
paragraph_format = style.paragraph_format
paragraph_format.space_before = Pt(0)
paragraph_format.space_after = Pt(3)
paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
paragraph_format.line_spacing = Pt(12)

# Header
p = doc.add_paragraph()
p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
p.paragraph_format.space_after = Pt(6)
run = p.add_run("Yousif Abozid\nSoftware Engineer")
run.bold = True
run.font.size = Pt(14)

contact_info = "yousif.abozid@yahoo.com | +201024022092 | Cairo, Egypt |\n" \
               "linkedin.com/in/yousifabozid | github.com/YousifAbozid | yousifabozid.github.io | t.me/YousifAbozid | He/Him"
p = doc.add_paragraph(contact_info)
p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
p.paragraph_format.space_after = Pt(12)

# Section Header Helper
def add_section_header(title):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(4)
    run = p.add_run(title)
    run.bold = True
    run.font.size = Pt(12)
    p.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT

# Paragraph Helper
def add_paragraph(text, bold=False, italic=False, size=10.5):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(3)
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    p.paragraph_format.line_spacing = Pt(12)
    run = p.add_run(text)
    run.bold = bold
    run.italic = italic
    run.font.size = Pt(size)
    return p

# Summary (no personal pronouns)
add_section_header("Summary")
add_paragraph("Self-taught software engineer trained through rigorous programs including Udacity, Helsinki University, "
              "Holberton School, and ALX Africa. Passionate about solving complex problems and crafting efficient, elegant software solutions. "
              "A lifelong learner thriving in dynamic environments and contributing to high-impact engineering projects.")

# Experience
add_section_header("Experience")

add_paragraph("Full Stack Software Engineer — Salem Ventures & TradeSocio                                    07/2024 – 07/2025", bold=True)
add_paragraph("Zahraa Al Maadi, Cairo, Egypt - On-Site")
add_paragraph(
    "• Developed Bruno, a trading platform used by 10K+ users, using React, Tailwind CSS, and Redux.\n"
    "• Increased platform performance by 60% through modularization and advanced hooks.\n"
    "• Integrated biometric authentication (Face ID/Fingerprint) across iOS & Android builds.\n"
    "• Delivered ERP system managing 500+ client accounts and permissions with React Query and REST APIs.\n"
    "• Reduced frontend defects by 40% after refactoring legacy code into reusable components.\n"
    "• Coordinated with cross-functional teams to launch 20+ features across 3 fintech products.\n"
    "• Provided mentorship to junior developers and onboarding support to new hires.\n"
    "• Partnered directly with CEO on strategic feature roadmap and market fit goals."
)

add_paragraph("Software Engineer — Mohandes Life Insurance                                                           03/2024 – 06/2024", bold=True)
add_paragraph("Dokki, Giza, Egypt - On-Site")
add_paragraph(
    "• Built insurance issuance system increasing processing efficiency by 80%.\n"
    "• Customized 10+ products to meet specific client requirements.\n"
    "• Achieved 99.9% uptime by maintaining and improving SISos 11 internal system.\n"
    "• Produced actionable reports adopted by executive leadership for strategic planning.\n"
    "• Led modernization of legacy stack, cutting support tickets by 45%."
)

add_paragraph("Software Engineer — ALX Africa                                                                                02/2023 – 06/2024", bold=True)
add_paragraph("Nairobi, Kenya - Remote")
add_paragraph(
    "• Completed 50+ projects, including full-stack web apps using React, Node.js, and PostgreSQL.\n"
    "• Mastered 12+ tools across front and backend: TypeScript, GraphQL, AWS, CI/CD, and testing frameworks.\n"
    "• Reviewed peer submissions and supported community learning groups weekly.\n"
    "• Contributed to production-grade apps featured in ALX engineering showcase."
)

add_paragraph("Software Engineer — Gig Bud                                                                                    02/2022 – 02/2023", bold=True)
add_paragraph("Zagazig, Egypt - Remote")
add_paragraph(
    "• Built scalable SaaS web app using Next.js and Express.\n"
    "• Decreased feature delivery time by 30% through improved time management.\n"
    "• Owned end-to-end implementation of core modules and database connections.\n"
    "• Participated in Agile sprints and contributed to daily standups and code reviews."
)

# Education
add_section_header("Education")

add_paragraph("Software Engineering Program — Holberton School & ALX Africa                           02/2023 – 06/2024", bold=True)
add_paragraph("Cairo, Egypt - Remote")
add_paragraph("Graduated with Front End Specialization from a highly selective and intensive 12-month program.")

add_paragraph("Advanced Full-Stack Web Development Nanodegree — Udacity                                01/2023 – 02/2023", bold=True)
add_paragraph("Cairo, Egypt - Remote")
add_paragraph("MCIT-sponsored program focusing on advanced topics for job readiness and tech industry immersion.")

add_paragraph("Full Stack Open MOOC — University of Helsinki                                                        09/2020 – 01/2021", bold=True)
add_paragraph("Helsinki, Finland - Remote")
add_paragraph("Completed MOOC with distinction; developed several apps and contributed to peer learning community.")

add_paragraph("Bachelor of Law — Faculty of Law - Zagazig University                                              09/2018 – 06/2022", bold=True)
add_paragraph("Zagazig, Egypt – On-Site")
add_paragraph("Completed a 4-year law degree.")

# Skills
add_section_header("Skills")
# add_paragraph("Languages: HTML, CSS, JavaScript, TypeScript, SQL, Bash, C, Python")

p = doc.add_paragraph()
p.paragraph_format.space_before = Pt(0)
p.paragraph_format.space_after = Pt(3)
p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
p.paragraph_format.line_spacing = Pt(12)
run1 = p.add_run("Languages: ")
run1.bold = True
run1.font.size = Pt(10.5)
run2 = p.add_run("HTML, CSS, JavaScript, TypeScript, SQL, Bash, C, Python")
run2.font.size = Pt(10.5)

p = doc.add_paragraph()
p.paragraph_format.space_before = Pt(0)
p.paragraph_format.space_after = Pt(3)
p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
p.paragraph_format.line_spacing = Pt(12)
run1 = p.add_run("Frameworks & Libraries: ")
run1.bold = True
run1.font.size = Pt(10.5)
run2 = p.add_run("React.js, Next.js, Vite.js, Redux, Context API, Express.js, Mongoose")
run2.font.size = Pt(10.5)

p = doc.add_paragraph()
p.paragraph_format.space_before = Pt(0)
p.paragraph_format.space_after = Pt(3)
p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
p.paragraph_format.line_spacing = Pt(12)
run1 = p.add_run("UI Libraries: ")
run1.bold = True
run1.font.size = Pt(10.5)
run2 = p.add_run("Material-UI, Tailwind CSS, Bootstrap, Shadcn‑UI, Styled Components")
run2.font.size = Pt(10.5)

p = doc.add_paragraph()
p.paragraph_format.space_before = Pt(0)
p.paragraph_format.space_after = Pt(3)
p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
p.paragraph_format.line_spacing = Pt(12)
run1 = p.add_run("State Management & Data Fetching: ")
run1.bold = True
run1.font.size = Pt(10.5)
run2 = p.add_run("React Query, Redux, Context API, Apollo Client")
run2.font.size = Pt(10.5)

p = doc.add_paragraph()
p.paragraph_format.space_before = Pt(0)
p.paragraph_format.space_after = Pt(3)
p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
p.paragraph_format.line_spacing = Pt(12)
run1 = p.add_run("Mobile / Cross‑Platform: ")
run1.bold = True
run1.font.size = Pt(10.5)
run2 = p.add_run("Capacitor (iOS/Android, biometric auth)")
run2.font.size = Pt(10.5)

p = doc.add_paragraph()
p.paragraph_format.space_before = Pt(0)
p.paragraph_format.space_after = Pt(3)
p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
p.paragraph_format.line_spacing = Pt(12)
run1 = p.add_run("API & Auth: ")
run1.bold = True
run1.font.size = Pt(10.5)
run2 = p.add_run("REST, GraphQL, JWT, Clerk, bcrypt, dotenv, cors, Webhooks, Axios, Fetch API")
run2.font.size = Pt(10.5)

p = doc.add_paragraph()
p.paragraph_format.space_before = Pt(0)
p.paragraph_format.space_after = Pt(3)
p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
p.paragraph_format.line_spacing = Pt(12)
run1 = p.add_run("Databases: ")
run1.bold = True
run1.font.size = Pt(10.5)
run2 = p.add_run("MongoDB, PostgreSQL")
run2.font.size = Pt(10.5)

p = doc.add_paragraph()
p.paragraph_format.space_before = Pt(0)
p.paragraph_format.space_after = Pt(3)
p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
p.paragraph_format.line_spacing = Pt(12)
run1 = p.add_run("Testing: ")
run1.bold = True
run1.font.size = Pt(10.5)
run2 = p.add_run("Jest, Jasmine, Cypress, Playwright")
run2.font.size = Pt(10.5)

p = doc.add_paragraph()
p.paragraph_format.space_before = Pt(0)
p.paragraph_format.space_after = Pt(3)
p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
p.paragraph_format.line_spacing = Pt(12)
run1 = p.add_run("CI/CD & DevOps: ")
run1.bold = True
run1.font.size = Pt(10.5)
run2 = p.add_run("Git, GitHub, GitHub Actions, CircleCI, Netlify, Heroku, Vercel, AWS")
run2.font.size = Pt(10.5)

p = doc.add_paragraph()
p.paragraph_format.space_before = Pt(0)
p.paragraph_format.space_after = Pt(3)
p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
p.paragraph_format.line_spacing = Pt(12)
run1 = p.add_run("Other Services & Tools: ")
run1.bold = True
run1.font.size = Pt(10.5)
run2 = p.add_run("Stripe, Cloudinary, Scroll‑Lock")
run2.font.size = Pt(10.5)

# Languages
add_section_header("Languages")
add_paragraph("Arabic — Native")
add_paragraph("English — C1")

# Save updated document
final_doc_path_2 = "./Yousif_Abozid_Resume_ATS_Optimized.docx"
doc.save(final_doc_path_2)
final_doc_path_2
