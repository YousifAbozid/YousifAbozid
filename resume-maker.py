from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT, WD_LINE_SPACING
from docx.oxml.shared import qn
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import nsdecls
from docx.oxml import parse_xml
from resume_data import *

# Section Header Helper with enhanced styling
def add_section_header(title):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.left_indent = Pt(0)
    p.paragraph_format.right_indent = Pt(0)
    
    run = p.add_run(title)
    run.bold = True
    run.font.size = Pt(13)
    run.font.color.rgb = RGBColor(0x1f, 0x4e, 0x79)  # Professional blue color
    p.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
    
    # Add underline effect
    run.font.underline = True
    
    # Add a divider line after section header
    divider = doc.add_paragraph()
    divider.paragraph_format.space_before = Pt(0)
    divider.paragraph_format.space_after = Pt(8)
    divider_run = divider.add_run("━" * 60)  # Unicode line character
    divider_run.font.color.rgb = RGBColor(0xe0, 0xe0, 0xe0)  # Light gray
    divider_run.font.size = Pt(8)

# Enhanced Paragraph Helper with better formatting
def add_paragraph(text, bold=False, italic=False, size=10.5, indent=0):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    p.paragraph_format.line_spacing = Pt(14)  # Slightly more line spacing
    p.paragraph_format.left_indent = Inches(indent)
    
    run = p.add_run(text)
    run.bold = bold
    run.italic = italic
    run.font.size = Pt(size)
    return p

# Hyperlink Helper Function
def add_hyperlink(paragraph, text, url):
    """Add a hyperlink to a paragraph"""
    # Get the paragraph's part (document part)
    part = paragraph.part
    r_id = part.relate_to(url, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink", is_external=True)
    
    # Create the w:hyperlink tag and add needed values
    hyperlink = paragraph._element.makeelement(qn('w:hyperlink'))
    hyperlink.set(qn('r:id'), r_id)
    
    # Create a new run object and add the text
    new_run = hyperlink.makeelement(qn('w:r'))
    rPr = new_run.makeelement(qn('w:rPr'))
    
    # Add hyperlink styling (blue color and underline)
    c = rPr.makeelement(qn('w:color'))
    c.set(qn('w:val'), "0563C1")  # Blue color
    rPr.append(c)
    
    u = rPr.makeelement(qn('w:u'))
    u.set(qn('w:val'), "single")  # Underline
    rPr.append(u)
    
    new_run.append(rPr)
    new_run.text = text
    hyperlink.append(new_run)
    
    # Append the hyperlink to the paragraph
    paragraph._element.append(hyperlink)

# Enhanced Paragraph Helper for Links
def add_paragraph_with_links(text_parts, bold=False, italic=False, size=10.5):
    """Add paragraph with multiple text parts, some can be hyperlinks"""
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(3)
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    p.paragraph_format.line_spacing = Pt(12)
    
    for part in text_parts:
        if isinstance(part, dict) and 'url' in part:
            # This is a hyperlink
            add_hyperlink(p, part['text'], part['url'])
        else:
            # Regular text
            run = p.add_run(part)
            run.bold = bold
            run.italic = italic
            run.font.size = Pt(size)
    
    return p

# Enhanced bullet point helper
def add_bullet_points(bullets, indent=0.2):
    for bullet in bullets:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(3)
        p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
        p.paragraph_format.line_spacing = Pt(14)
        p.paragraph_format.left_indent = Inches(indent)
        p.paragraph_format.first_line_indent = Inches(-0.15)
        
        # Add custom bullet
        bullet_run = p.add_run("")  # Custom bullet character
        bullet_run.font.color.rgb = RGBColor(0x1f, 0x4e, 0x79)
        bullet_run.font.size = Pt(10)
        
        text_run = p.add_run(bullet)
        text_run.font.size = Pt(10.5)

# Start a new doc with enhanced styling
doc = Document()

# Set document margins
sections = doc.sections
for section in sections:
    section.top_margin = Inches(0.7)
    section.bottom_margin = Inches(0.7)
    section.left_margin = Inches(0.8)
    section.right_margin = Inches(0.8)

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

# Header with enhanced styling
header_table = doc.add_table(rows=1, cols=1)
header_table.alignment = WD_TABLE_ALIGNMENT.CENTER
header_cell = header_table.cell(0, 0)

# Add subtle background to header
shading_elm = parse_xml(r'<w:shd {} w:fill="f8f9fa"/>'.format(nsdecls('w')))
header_cell._tc.get_or_add_tcPr().append(shading_elm)

header_p = header_cell.paragraphs[0]
header_p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
header_p.paragraph_format.space_before = Pt(8)
header_p.paragraph_format.space_after = Pt(8)

name_run = header_p.add_run(PERSONAL_INFO['name'])
name_run.bold = True
name_run.font.size = Pt(18)
name_run.font.color.rgb = RGBColor(0x1f, 0x4e, 0x79)

title_run = header_p.add_run(f"\n\n{PERSONAL_INFO['title']}")
title_run.font.size = Pt(13)
title_run.font.color.rgb = RGBColor(0x2c, 0x3e, 0x50)

# Contact info with enhanced formatting
contact_p = doc.add_paragraph()
contact_p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
contact_p.paragraph_format.space_after = Pt(16)
contact_p.paragraph_format.space_before = Pt(8)

# Add contact info with hyperlinks and separators
contact_parts = [
    {"text": PERSONAL_INFO['email'], "url": f"mailto:{PERSONAL_INFO['email']}"},
    " ● ",  # Bullet separator
    {"text": PERSONAL_INFO['phone'], "url": f"tel:{PERSONAL_INFO['phone']}"},
    " ● ",
    PERSONAL_INFO['location'],
    "\n",
    {"text": PERSONAL_INFO['linkedin'], "url": PERSONAL_INFO['linkedin']},
    " ● ",
    {"text": "GitHub", "url": PERSONAL_INFO['github']},
    " ● ",
    {"text": "Portfolio", "url": PERSONAL_INFO['portfolio']},
    " ● ",
    {"text": "Telegram", "url": PERSONAL_INFO['telegram']},
    " ● ",
    PERSONAL_INFO['pronouns']
]

for part in contact_parts:
    if isinstance(part, dict) and 'url' in part:
        add_hyperlink(contact_p, part['text'], part['url'])
    else:
        run = contact_p.add_run(part)
        run.font.size = Pt(10.5)
        if part in [" ● ", " • "]:
            run.font.color.rgb = RGBColor(0x1f, 0x4e, 0x79)

# Summary with box styling
add_section_header("Summary")
# Add subtle background to summary - remove the duplicate paragraph
summary_table = doc.add_table(rows=1, cols=1)
summary_cell = summary_table.cell(0, 0)
shading_elm = parse_xml(r'<w:shd {} w:fill="f8f9fa"/>'.format(nsdecls('w')))
summary_cell._tc.get_or_add_tcPr().append(shading_elm)
summary_cell.paragraphs[0].clear()
summary_cell.paragraphs[0].add_run(SUMMARY).font.size = Pt(11)

# Experience with enhanced formatting
add_section_header("Experience")

for i, exp in enumerate(EXPERIENCE):
    # Job title with enhanced styling
    title_p = add_paragraph(exp["title"], bold=True, size=11.5)
    for run in title_p.runs:
        if run.bold:
            run.font.color.rgb = RGBColor(0x2c, 0x3e, 0x50)
    
    # Company, location and duration on same line
    location_text = exp["location"]
    if "duration" in exp:
        location_text += f" | {exp['duration']}"
    
    location_p = add_paragraph(location_text, italic=True, size=10)
    for run in location_p.runs:
        run.font.color.rgb = RGBColor(0x6c, 0x75, 0x7d)
    
    # Bullet points with custom styling - use helper function
    add_bullet_points(exp["bullets"])
    
    # Add spacing between experience entries
    if i < len(EXPERIENCE) - 1:
        spacer = doc.add_paragraph()
        spacer.paragraph_format.space_after = Pt(6)

# Education with enhanced formatting
add_section_header("Education")

for edu in EDUCATION:
    title_p = add_paragraph(edu["title"], bold=True, size=11.5)
    for run in title_p.runs:
        if run.bold:
            run.font.color.rgb = RGBColor(0x2c, 0x3e, 0x50)
    
    # Location and duration on same line
    location_text = edu["location"]
    if "duration" in edu:
        location_text += f" | {edu['duration']}"
    
    location_p = add_paragraph(location_text, italic=True, size=10)
    for run in location_p.runs:
        run.font.color.rgb = RGBColor(0x6c, 0x75, 0x7d)
    
    add_paragraph(edu["description"], size=10.5)

# Projects with enhanced formatting
add_section_header("Projects")

for i, project in enumerate(PROJECTS):
    # Project title
    title_p = add_paragraph(project["title"], bold=True, size=11.5)
    for run in title_p.runs:
        if run.bold:
            run.font.color.rgb = RGBColor(0x2c, 0x3e, 0x50)
    
    # Project links
    if "demo_urls" in project:
        link_parts = []
        for name, url in project["demo_urls"].items():
            if link_parts:
                link_parts.extend([" │ "])  # Vertical bar separator
            link_parts.extend([
                f"{name}: ",
                {"text": "Demo", "url": url}
            ])
        add_paragraph_with_links(link_parts)
    else:
        add_paragraph_with_links([
            {"text": "Demo", "url": project["demo_url"]},
            " │ ",
            {"text": "Code", "url": project["code_url"]}
        ])
    
    # Project description with bullets - use helper function
    add_bullet_points(project["bullets"])
    
    # Add spacing between projects
    if i < len(PROJECTS) - 1:
        spacer = doc.add_paragraph()
        spacer.paragraph_format.space_after = Pt(6)

# Skills with enhanced paragraph layout
add_section_header("Skills")

for category, skills in SKILLS.items():
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    p.paragraph_format.line_spacing = Pt(14)
    
    # Category name with styling
    category_run = p.add_run(f"{category}: ")
    category_run.bold = True
    category_run.font.size = Pt(10.5)
    category_run.font.color.rgb = RGBColor(0x1f, 0x4e, 0x79)  # Blue color for categories
    
    # Skills list with styling
    skills_run = p.add_run(skills)
    skills_run.font.size = Pt(10.5)
    skills_run.font.color.rgb = RGBColor(0x2c, 0x3e, 0x50)  # Dark color for skills

# Languages with enhanced formatting
add_section_header("Languages")
for language in LANGUAGES:
    lang_p = add_paragraph(language, size=10.5)
    # Add subtle icon before language
    icon_run = lang_p.runs[0]
    icon_run.text = f"🌍 {language}"

# Save updated document
final_doc_path_2 = "./Yousif_Abozid_Resume.docx"
doc.save(final_doc_path_2)
print(f"Resume saved to {final_doc_path_2}")
