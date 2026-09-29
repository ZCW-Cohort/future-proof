#!/usr/bin/env python3
"""
Generate a PowerPoint presentation for the Future-Proof project
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor

# ZipCode Wilmington Brand Colors
PRIMARY = RGBColor(0, 40, 58)        # #00283a (Dark Navy Blue)
SECONDARY = RGBColor(59, 190, 140)   # #3bbe8c (Teal Green)
TERTIARY = RGBColor(71, 231, 170)    # #47E7AA (Light Green/Mint)
ACCENT = RGBColor(252, 207, 118)     # #fccf76 (Pale Yellow/Gold)
DARK = RGBColor(10, 30, 37)          # #0A1E25 (Very Dark Blue)
MEDIUM = RGBColor(16, 77, 99)        # #104d63 (Deep Teal)

def create_presentation():
    prs = Presentation()
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(7.5)

    # Slide 1: Title
    slide1 = prs.slides.add_slide(prs.slide_layouts[6])  # Blank layout
    set_background(slide1, PRIMARY)

    # Logo
    add_text_box(slide1, "ZIP CODE WILMINGTON", 0.5, 0.5, 3, 0.5, SECONDARY, 16, bold=True)

    # Slide number
    add_text_box(slide1, "1/10", 9, 0.5, 0.5, 0.3, SECONDARY, 14, bold=True, align=PP_ALIGN.RIGHT)

    # Title
    title_box = add_text_box(slide1, "Future-Proof", 1, 2, 8, 1, RGBColor(255, 255, 255), 60, bold=True)

    # Subtitle
    subtitle_box = add_text_box(slide1, "Personal Notes Manager", 1, 3.2, 8, 0.8, TERTIARY, 42, bold=True)

    # Tagline
    add_text_box(slide1, "Building Software That Lasts", 1, 4.3, 8, 0.6, ACCENT, 28, bold=True)

    # Description
    add_text_box(slide1, "A multi-phase educational project for adult beginners", 1, 5.2, 8, 0.5, RGBColor(255, 255, 255), 18)

    # Slide 2: What is Future-Proof Software?
    slide2 = prs.slides.add_slide(prs.slide_layouts[6])
    set_background(slide2, PRIMARY)

    add_text_box(slide2, "ZIP CODE WILMINGTON", 0.5, 0.5, 3, 0.5, SECONDARY, 16, bold=True)
    add_text_box(slide2, "2/10", 9, 0.5, 0.5, 0.3, SECONDARY, 14, bold=True, align=PP_ALIGN.RIGHT)

    add_text_box(slide2, "What is Future-Proof Software?", 0.5, 1, 9, 0.7, SECONDARY, 36, bold=True)
    add_text_box(slide2, "Code designed to remain useful, maintainable, and adaptable as technology evolves.", 0.5, 2, 9, 0.6, RGBColor(255, 255, 255), 20)

    bullets = [
        "Like building a house with a strong foundation",
        "Can support renovations and additions years later",
        "User needs evolve, new technologies emerge",
        "Future-proof code makes change easier and less expensive"
    ]
    add_bullet_list(slide2, bullets, 0.5, 3, 9, 3.5, RGBColor(255, 255, 255), 18)

    # Slide 3: The Problem
    slide3 = prs.slides.add_slide(prs.slide_layouts[6])
    set_background(slide3, DARK)

    add_text_box(slide3, "ZIP CODE WILMINGTON", 0.5, 0.5, 3, 0.5, SECONDARY, 16, bold=True)
    add_text_box(slide3, "3/10", 9, 0.5, 0.5, 0.3, SECONDARY, 14, bold=True, align=PP_ALIGN.RIGHT)

    add_text_box(slide3, "The Problem", 0.5, 1, 9, 0.7, SECONDARY, 36, bold=True)
    add_text_box(slide3, "Most note-taking apps today:", 0.5, 2, 9, 0.5, RGBColor(255, 255, 255), 20)

    bullets = [
        "Lock your data in proprietary formats",
        "Require specific software that may not exist in 10 years",
        "Store data in the cloud (what if the company shuts down?)",
        "Don't give you full control of your own notes"
    ]
    add_bullet_list(slide3, bullets, 0.5, 3, 9, 2.5, RGBColor(255, 255, 255), 18)

    add_text_box(slide3, "We're building something better.", 0.5, 6, 9, 0.5, TERTIARY, 22, bold=True)

    # Slide 4: Project Overview
    slide4 = prs.slides.add_slide(prs.slide_layouts[6])
    set_background(slide4, PRIMARY)

    add_text_box(slide4, "ZIP CODE WILMINGTON", 0.5, 0.5, 3, 0.5, SECONDARY, 16, bold=True)
    add_text_box(slide4, "4/10", 9, 0.5, 0.5, 0.3, SECONDARY, 14, bold=True, align=PP_ALIGN.RIGHT)

    add_text_box(slide4, "Personal Notes Manager", 0.5, 1, 9, 0.7, SECONDARY, 36, bold=True)
    add_text_box(slide4, "A text-based note-taking system that stores notes as plain UTF-8 text files with YAML metadata.", 0.5, 2, 9, 0.6, RGBColor(255, 255, 255), 18)

    # Code block
    code_text = """---
title: My Example Note
author: Zip Coder
tags: [example, documentation]
---

This is the content of my note.
It's just plain text!"""

    code_box = add_text_box(slide4, code_text, 1, 3.2, 8, 2.2, TERTIARY, 14)
    shape = code_box
    fill = shape.fill
    fill.solid()
    fill.fore_color.rgb = DARK

    add_text_box(slide4, "All notes kept in a hidden folder in your HOME directory", 0.5, 5.8, 9, 0.5, ACCENT, 16)

    # Slide 5: Why Text Files?
    slide5 = prs.slides.add_slide(prs.slide_layouts[6])
    set_background(slide5, DARK)

    add_text_box(slide5, "ZIP CODE WILMINGTON", 0.5, 0.5, 3, 0.5, SECONDARY, 16, bold=True)
    add_text_box(slide5, "5/10", 9, 0.5, 0.5, 0.3, SECONDARY, 14, bold=True, align=PP_ALIGN.RIGHT)

    add_text_box(slide5, "Why Text Files?", 0.5, 1, 9, 0.7, SECONDARY, 36, bold=True)

    bullets = [
        "Universal Compatibility - Readable on any OS, any editor",
        "Longevity - Text files have worked for decades, will work for decades more",
        "Version Control Friendly - Works seamlessly with Git",
        "Human Readable - No specialized software needed",
        "You Own Your Data - No cloud dependency",
        "Open Standards - Markdown + YAML"
    ]
    add_bullet_list(slide5, bullets, 0.5, 2, 9, 4.5, RGBColor(255, 255, 255), 16)

    # Slide 6: Phase 1 - CLI
    slide6 = prs.slides.add_slide(prs.slide_layouts[6])
    set_background(slide6, PRIMARY)

    add_text_box(slide6, "ZIP CODE WILMINGTON", 0.5, 0.5, 3, 0.5, SECONDARY, 16, bold=True)
    add_text_box(slide6, "6/10", 9, 0.5, 0.5, 0.3, SECONDARY, 14, bold=True, align=PP_ALIGN.RIGHT)

    add_text_box(slide6, "Phase 1: Command Line Interface", 0.5, 1, 9, 0.7, SECONDARY, 32, bold=True)
    add_text_box(slide6, "Build a CLI tool to manage your notes from the terminal:", 0.5, 1.9, 9, 0.5, RGBColor(255, 255, 255), 16)

    # Code block
    code_text = """notes create          # Create a new note
notes list            # List all notes
notes read <id>       # Display a note
notes edit <id>       # Edit a note
notes delete <id>     # Delete a note
notes search "query"  # Search notes
notes stats           # Show statistics"""

    code_box = add_text_box(slide6, code_text, 1, 2.7, 8, 2.8, TERTIARY, 13)
    shape = code_box
    fill = shape.fill
    fill.solid()
    fill.fore_color.rgb = DARK

    add_text_box(slide6, "Learn: File I/O, YAML parsing, command-line arguments, testing", 0.5, 5.8, 9, 0.5, ACCENT, 14)

    # Slide 7: Phase 2 - REST Server
    slide7 = prs.slides.add_slide(prs.slide_layouts[6])
    set_background(slide7, DARK)

    add_text_box(slide7, "ZIP CODE WILMINGTON", 0.5, 0.5, 3, 0.5, SECONDARY, 16, bold=True)
    add_text_box(slide7, "7/10", 9, 0.5, 0.5, 0.3, SECONDARY, 14, bold=True, align=PP_ALIGN.RIGHT)

    add_text_box(slide7, "Phase 2: REST Server + Web Frontend", 0.5, 1, 9, 0.7, SECONDARY, 32, bold=True)

    # Three boxes
    box1_bullets = ["GET /api/notes", "POST /api/notes", "PUT /api/notes/:id", "DELETE /api/notes/:id"]
    box2_bullets = ["Access from browser", "Mobile-responsive", "Vanilla JavaScript"]
    box3_bullets = ["Dataset support (CSV/JSON)", "S3 backup", "Authentication"]

    add_box_with_bullets(slide7, "REST API", box1_bullets, 0.5, 2.2, 2.8, 3.2, MEDIUM, ACCENT, RGBColor(255, 255, 255), SECONDARY)
    add_box_with_bullets(slide7, "Web Interface", box2_bullets, 3.6, 2.2, 2.8, 3.2, MEDIUM, ACCENT, RGBColor(255, 255, 255), SECONDARY)
    add_box_with_bullets(slide7, "Bonus", box3_bullets, 6.7, 2.2, 2.8, 3.2, MEDIUM, ACCENT, RGBColor(255, 255, 255), SECONDARY)

    add_text_box(slide7, "Learn: Client-server architecture, REST APIs, web development", 0.5, 5.8, 9, 0.5, TERTIARY, 14)

    # Slide 8: Phase 3 - GUI
    slide8 = prs.slides.add_slide(prs.slide_layouts[6])
    set_background(slide8, PRIMARY)

    add_text_box(slide8, "ZIP CODE WILMINGTON", 0.5, 0.5, 3, 0.5, SECONDARY, 16, bold=True)
    add_text_box(slide8, "8/10", 9, 0.5, 0.5, 0.3, SECONDARY, 14, bold=True, align=PP_ALIGN.RIGHT)

    add_text_box(slide8, "Phase 3: Graphical User Interface", 0.5, 1, 9, 0.7, SECONDARY, 32, bold=True)

    bullets = [
        "Rich text editor with markdown support",
        "Real-time preview of your notes",
        "Visual organization (folders/tags)",
        "Enhanced search with highlighting",
        "Dark/light theme options",
        "Export to PDF, HTML, and more"
    ]
    add_bullet_list(slide8, bullets, 0.5, 2, 9, 3.5, RGBColor(255, 255, 255), 18)

    add_text_box(slide8, "Same text files, better interface!", 0.5, 6, 9, 0.5, ACCENT, 20, bold=True)

    # Slide 9: Learning Objectives
    slide9 = prs.slides.add_slide(prs.slide_layouts[6])
    set_background(slide9, DARK)

    add_text_box(slide9, "ZIP CODE WILMINGTON", 0.5, 0.5, 3, 0.5, SECONDARY, 16, bold=True)
    add_text_box(slide9, "9/10", 9, 0.5, 0.5, 0.3, SECONDARY, 14, bold=True, align=PP_ALIGN.RIGHT)

    add_text_box(slide9, "What You'll Learn", 0.5, 1, 9, 0.7, SECONDARY, 36, bold=True)

    # Two columns
    core_bullets = ["File I/O operations", "Data parsing (YAML, JSON)", "Command-line processing", "Test-Driven Development"]
    advanced_bullets = ["REST API development", "Frontend-backend integration", "User interface design", "Progressive enhancement"]

    add_box_with_bullets(slide9, "Core Skills", core_bullets, 0.5, 2, 4, 3.5, MEDIUM, ACCENT, RGBColor(255, 255, 255), SECONDARY)
    add_box_with_bullets(slide9, "Advanced Topics", advanced_bullets, 5.5, 2, 4, 3.5, MEDIUM, ACCENT, RGBColor(255, 255, 255), SECONDARY)

    add_text_box(slide9, "Build skills that transfer to ANY programming project!", 0.5, 5.8, 9, 0.5, TERTIARY, 16, bold=True)

    # Slide 10: Getting Started
    slide10 = prs.slides.add_slide(prs.slide_layouts[6])
    set_background(slide10, PRIMARY)

    add_text_box(slide10, "ZIP CODE WILMINGTON", 0.5, 0.5, 3, 0.5, SECONDARY, 16, bold=True)
    add_text_box(slide10, "10/10", 9, 0.5, 0.5, 0.3, SECONDARY, 14, bold=True, align=PP_ALIGN.RIGHT)

    add_text_box(slide10, "Ready to Get Started?", 0.5, 1, 9, 0.7, SECONDARY, 36, bold=True)

    prereq_bullets = [
        "Java 11+ OR Python 3.11+",
        "Text editor (VS Code recommended)",
        "Git for version control",
        "Enthusiasm to learn!"
    ]
    add_box_with_bullets(slide10, "Prerequisites", prereq_bullets, 1, 2, 8, 2.2, MEDIUM, ACCENT, RGBColor(255, 255, 255), SECONDARY)

    add_text_box(slide10, "Choose your own project name and fork the repo", 0.5, 4.5, 9, 0.4, TERTIARY, 18, bold=True)
    add_text_box(slide10, "Suggested names: NoteNexus, MindScribe, ThoughtVault, NoteCraft", 0.5, 5.1, 9, 0.4, RGBColor(255, 255, 255), 14)
    add_text_box(slide10, "Let's build software that lasts! 🚀", 0.5, 6, 9, 0.8, SECONDARY, 28, bold=True, align=PP_ALIGN.CENTER)

    return prs

def set_background(slide, color):
    """Set solid background color for a slide"""
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = color

def add_text_box(slide, text, left, top, width, height, color, font_size, bold=False, align=PP_ALIGN.LEFT):
    """Add a text box to a slide"""
    left = Inches(left)
    top = Inches(top)
    width = Inches(width)
    height = Inches(height)

    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.text = text
    tf.word_wrap = True

    p = tf.paragraphs[0]
    p.font.size = Pt(font_size)
    p.font.color.rgb = color
    p.font.bold = bold
    p.alignment = align

    return txBox

def add_bullet_list(slide, bullets, left, top, width, height, color, font_size):
    """Add a bulleted list to a slide"""
    left = Inches(left)
    top = Inches(top)
    width = Inches(width)
    height = Inches(height)

    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = True

    for i, bullet in enumerate(bullets):
        if i == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
        p.text = bullet
        p.level = 0
        p.font.size = Pt(font_size)
        p.font.color.rgb = color
        p.space_before = Pt(6)

    return txBox

def add_box_with_bullets(slide, title, bullets, left, top, width, height, bg_color, title_color, text_color, border_color):
    """Add a colored box with title and bullets"""
    left_in = Inches(left)
    top_in = Inches(top)
    width_in = Inches(width)
    height_in = Inches(height)

    # Add rectangle
    shape = slide.shapes.add_shape(
        1,  # Rectangle
        left_in, top_in, width_in, height_in
    )

    fill = shape.fill
    fill.solid()
    fill.fore_color.rgb = bg_color

    line = shape.line
    line.color.rgb = border_color
    line.width = Pt(3)

    # Add title
    title_box = slide.shapes.add_textbox(
        left_in + Inches(0.2),
        top_in + Inches(0.15),
        width_in - Inches(0.4),
        Inches(0.4)
    )
    tf_title = title_box.text_frame
    tf_title.text = title
    p_title = tf_title.paragraphs[0]
    p_title.font.size = Pt(18)
    p_title.font.color.rgb = title_color
    p_title.font.bold = True

    # Add bullets
    bullet_box = slide.shapes.add_textbox(
        left_in + Inches(0.2),
        top_in + Inches(0.65),
        width_in - Inches(0.4),
        height_in - Inches(0.85)
    )
    tf_bullets = bullet_box.text_frame
    tf_bullets.word_wrap = True

    for i, bullet in enumerate(bullets):
        if i == 0:
            p = tf_bullets.paragraphs[0]
        else:
            p = tf_bullets.add_paragraph()
        p.text = bullet
        p.level = 0
        p.font.size = Pt(13)
        p.font.color.rgb = text_color
        p.space_before = Pt(4)

if __name__ == "__main__":
    print("Creating Future-Proof presentation...")
    prs = create_presentation()
    prs.save('future-proof-presentation.pptx')
    print("✓ Presentation saved as 'future-proof-presentation.pptx'")
