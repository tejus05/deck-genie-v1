from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from typing import Dict, Any
from image_fetcher import fetch_image_for_slide
from utils import FONTS, COLORS, FONT_SIZES, MARGINS, CONTENT_AREA, IMAGE_AREA

# Define apply_text_formatting function here to avoid circular import
def apply_text_formatting(text_frame, font_name=FONTS["body"], size=FONT_SIZES["body"], 
                         bold=False, color=COLORS["black"], alignment=PP_ALIGN.LEFT):
    """Apply consistent text formatting to a text frame."""
    text_frame.paragraphs[0].alignment = alignment
    
    for paragraph in text_frame.paragraphs:
        for run in paragraph.runs:
            run.font.name = font_name
            run.font.size = Pt(size)
            run.font.bold = bold
            run.font.color.rgb = RGBColor.from_string(color)

def create_market_slide(prs: Presentation, content: Dict[str, Any], presentation_context: Dict[str, Any] = None):
    """Create the market opportunity slide with bullet points."""
    # Use blank layout to avoid unwanted placeholder elements
    slide_layout = prs.slide_layouts[6]  # Blank layout
    slide = prs.slides.add_slide(slide_layout)
    
    # Add title manually
    left = Inches(MARGINS["left"])
    top = Inches(MARGINS["top"])
    width = Inches(CONTENT_AREA["width"])
    height = Inches(0.8)
    
    title_box = slide.shapes.add_textbox(left, top, width, height)
    title_text = title_box.text_frame
    title_text.text = content.get("title", "Market Opportunity")
    
    apply_text_formatting(
        title_text, 
        font_name=FONTS["title"], 
        size=FONT_SIZES["title"], 
        bold=True
    )
      # Add bullet points
    left = Inches(MARGINS["left"])
    top = Inches(MARGINS["top"] + 1.0)
    width = Inches(CONTENT_AREA["width"])
    height = Inches(4.5)
    
    # Adjust content area if an image is used
    content_width = CONTENT_AREA["width"]
    image_path = None
    
    # Check persona from presentation context to adjust layout
    persona = presentation_context.get("metadata", {}).get("persona", "Generic") if presentation_context else "Generic"
    
    # For Marketing persona, use slightly narrower content area for better readability
    if persona == "Marketing":
        content_width = CONTENT_AREA["width"] * 0.9
        width = Inches(content_width)
    
    try:
        # Try to get a related image
        if presentation_context:
            image_path = fetch_image_for_slide(
                "market analysis", 
                presentation_context.get("metadata", {}).get("product_name", ""),
                presentation_context.get("metadata", {}).get("company_name", "")
            )
        
        if image_path:
            # Make content area narrower to accommodate image
            content_width = CONTENT_AREA["width"] * 0.6
            width = Inches(content_width)
            
            # Add image
            img_left = Inches(MARGINS["left"] + content_width + 0.5)
            img_top = Inches(MARGINS["top"] + 1.0)
            img_width = Inches(IMAGE_AREA["width"] * 0.9)
            img_height = Inches(IMAGE_AREA["height"] * 0.9)
            
            slide.shapes.add_picture(image_path, img_left, img_top, width=img_width, height=img_height)
    except Exception:
        # Continue without an image if there's an error
        pass
        
    # Create text box for bullets
    bullet_box = slide.shapes.add_textbox(left, top, width, height)
    bullet_tf = bullet_box.text_frame
    bullet_tf.word_wrap = True
    
    # Add bullets from content
    bullets = content.get("bullets", ["No market data available."])
    
    for i, bullet_text in enumerate(bullets):
        if i == 0:
            p = bullet_tf.paragraphs[0]
        else:
            p = bullet_tf.add_paragraph()
        
        p.text = bullet_text
        p.level = 0
        
        # Add bullet and spacing
        p.space_before = Pt(12)
        p.space_after = Pt(6)
        
        # Apply font formatting
        for run in p.runs:
            run.font.name = FONTS["body"]
            run.font.size = Pt(FONT_SIZES["body"])
            run.font.color.rgb = RGBColor.from_string(COLORS["black"])

def create_roadmap_slide(prs: Presentation, content: Dict[str, Any], presentation_context: Dict[str, Any] = None):
    """Create the product roadmap slide with phases."""
    # Use blank layout to avoid unwanted placeholder elements
    slide_layout = prs.slide_layouts[6]  # Blank layout
    slide = prs.slides.add_slide(slide_layout)
    
    # Add title manually
    left = Inches(MARGINS["left"])
    top = Inches(MARGINS["top"])
    width = Inches(CONTENT_AREA["width"])
    height = Inches(0.8)
    
    title_box = slide.shapes.add_textbox(left, top, width, height)
    title_text = title_box.text_frame
    title_text.text = content.get("title", "Product Roadmap")
    
    apply_text_formatting(
        title_text, 
        font_name=FONTS["title"], 
        size=FONT_SIZES["title"], 
        bold=True
    )
    
    # Get phases from content
    phases = content.get("phases", [
        {"phase_name": "Phase 1", "items": ["Feature 1", "Feature 2"]},
        {"phase_name": "Phase 2", "items": ["Feature 3", "Feature 4"]}
    ])
      # Calculate layout parameters - use full slide width for better spacing
    num_phases = len(phases)
    available_width = 11.33 - (MARGINS["left"] * 2)  # Width of 16:9 slide minus margins
    phase_width = min(2.8, available_width / max(1, min(num_phases, 4)))
    
    # Draw timeline as a horizontal line
    line_left = Inches(MARGINS["left"] + 0.25)
    line_top = Inches(MARGINS["top"] + 1.5)
    line_width = Inches(available_width - 0.5)
    line_height = Inches(0.05)
    
    timeline = slide.shapes.add_textbox(line_left, line_top, line_width, line_height)
    timeline.fill.solid()
    timeline.fill.fore_color.rgb = RGBColor.from_string(COLORS["accent"])
    
    # Add each phase
    for i, phase in enumerate(phases):
        phase_left = Inches(MARGINS["left"] + 0.25 + (i * phase_width))
        
        # Phase name (above timeline)
        phase_name_top = Inches(MARGINS["top"] + 1.0)
        name_box = slide.shapes.add_textbox(
            phase_left, phase_name_top, 
            Inches(phase_width), Inches(0.4)
        )
        name_tf = name_box.text_frame
        name_tf.text = phase["phase_name"]
        apply_text_formatting(name_tf, size=FONT_SIZES["subtitle"] - 2, bold=True, alignment=PP_ALIGN.CENTER)
        
        # Add circle marker on timeline
        marker_size = 0.2
        marker_left = phase_left + Inches(phase_width/2 - marker_size/2)
        marker_top = line_top - Inches(marker_size/4)
        marker = slide.shapes.add_textbox(
            marker_left, marker_top,
            Inches(marker_size), Inches(marker_size)
        )
        marker.fill.solid()
        marker.fill.fore_color.rgb = RGBColor.from_string(COLORS["accent_dark"])          # Phase items (below timeline)
        items_top = Inches(MARGINS["top"] + 1.8)
        items_box = slide.shapes.add_textbox(
            phase_left, items_top,
            Inches(phase_width - 0.2), Inches(3.0)  # Reduce width more to prevent overlap
        )
        items_tf = items_box.text_frame
        items_tf.word_wrap = True
        
        # Check persona to adjust text size
        persona = presentation_context.get("metadata", {}).get("persona", "Generic") if presentation_context else "Generic"
        
        # Add items - make sure text is short enough to fit
        items = phase.get("items", [])
        for j, item in enumerate(items):
            if j == 0:
                p = items_tf.paragraphs[0]
            else:
                p = items_tf.add_paragraph()
            
            # Truncate item text if too long - use shorter text for Marketing persona
            max_length = 25 if persona == "Marketing" else 30
            short_item = item if len(item) < max_length else item[:max_length-3] + "..."
            p.text = f"• {short_item}"
            p.space_before = Pt(6)
            p.space_after = Pt(3)  # Add space after paragraph
        
        # Use smaller font for the items to fit better
        font_size_adjust = FONT_SIZES["body"] - 4
        if persona == "Marketing":
            font_size_adjust = FONT_SIZES["body"] - 5  # Even smaller font for Marketing
        apply_text_formatting(items_tf, size=font_size_adjust, alignment=PP_ALIGN.LEFT)

def create_team_slide(prs: Presentation, content: Dict[str, Any], presentation_context: Dict[str, Any] = None):
    """Create the team slide with team member highlights."""
    # Use blank layout
    slide_layout = prs.slide_layouts[6]  # Blank layout
    slide = prs.slides.add_slide(slide_layout)
    
    # Add title manually
    left = Inches(MARGINS["left"])
    top = Inches(MARGINS["top"])
    width = Inches(CONTENT_AREA["width"])
    height = Inches(0.8)
    
    title_box = slide.shapes.add_textbox(left, top, width, height)
    title_text = title_box.text_frame
    title_text.text = content.get("title", "Our Team")
    
    apply_text_formatting(
        title_text, 
        font_name=FONTS["title"], 
        size=FONT_SIZES["title"], 
        bold=True
    )
    
    # Get team members from content
    members = content.get("members", [
        {"name_role": "Team Member", "highlight": "Team highlight"}
    ])
      # Calculate layout parameters - use full slide width for better spacing
    num_members = len(members)
    available_width = 11.33 - (MARGINS["left"] * 2)  # Width of 16:9 slide minus margins
    member_width = min(2.5, available_width / max(1, min(num_members, 4)))
      # Add each team member with appropriate spacing
    for i, member in enumerate(members):
        # Only display up to 4 members per slide
        if i >= 4:
            break
            
        # Check persona to adjust layout
        persona = presentation_context.get("metadata", {}).get("persona", "Generic") if presentation_context else "Generic"
        
        # Calculate position with proper spacing - add more space for Marketing persona
        spacing_adjust = 0.35 if persona == "Marketing" else 0.25
        member_left = Inches(MARGINS["left"] + 0.5 + (i * (member_width + spacing_adjust)))
        member_top = Inches(MARGINS["top"] + 1.5)
        
        # Member box with subtle background - slightly smaller width for Marketing persona
        width_adjust = 0.4 if persona == "Marketing" else 0.3
        member_box = slide.shapes.add_textbox(
            member_left, member_top,
            Inches(member_width - width_adjust), Inches(3.5)
        )
        member_box.fill.solid()
        member_box.fill.fore_color.rgb = RGBColor.from_string(COLORS["light_gray"])
        
        # Name and role
        name_box = slide.shapes.add_textbox(
            member_left + Inches(0.15), member_top + Inches(0.2),
            Inches(member_width - 0.6), Inches(0.6)
        )
        name_tf = name_box.text_frame
        name_tf.text = member.get("name_role", "Team Member")
        apply_text_formatting(name_tf, size=FONT_SIZES["subtitle"] - 2, bold=True, alignment=PP_ALIGN.CENTER)
        
        # Member highlight/description
        highlight_box = slide.shapes.add_textbox(
            member_left + Inches(0.15), member_top + Inches(1.0),
            Inches(member_width - 0.6), Inches(2.0)
        )
        highlight_tf = highlight_box.text_frame
        highlight_tf.word_wrap = True
        highlight_tf.text = member.get("highlight", "")
        apply_text_formatting(highlight_tf, size=FONT_SIZES["body"] - 2, alignment=PP_ALIGN.CENTER)