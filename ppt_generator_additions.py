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
    """Create the market opportunity slide with proper layout."""
    slide_layout = prs.slide_layouts[6]  # Blank layout
    slide = prs.slides.add_slide(slide_layout)
    
    # Set background to white for consistency
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor.from_string(COLORS["white"])
    
    # Add title with consistent positioning
    title_left = Inches(0.5)
    title_top = Inches(0.5)
    title_width = Inches(12.33)
    title_height = Inches(0.8)
    
    title_box = slide.shapes.add_textbox(title_left, title_top, title_width, title_height)
    title_tf = title_box.text_frame
    title_tf.word_wrap = True
    title_tf.auto_size = False
    title_p = title_tf.paragraphs[0]
    
    # Get title from content, with product name substitution if needed
    product_name = ""
    if presentation_context and "metadata" in presentation_context:
        product_name = presentation_context["metadata"].get("product_name", "")
    
    title_text = content.get("title", "Market Opportunity")
    if product_name:
        title_text = title_text.replace("[Product Name]", product_name)
    
    title_p.text = title_text
    title_p.alignment = PP_ALIGN.LEFT
    
    # Apply title formatting
    for run in title_p.runs:
        run.font.name = FONTS["title"]
        run.font.size = Pt(32)
        run.font.bold = True
        run.font.color.rgb = RGBColor.from_string(COLORS["black"])
    
    # Create large market size callout (left side)
    if "market_size" in content:
        market_left = Inches(1.0)
        market_top = Inches(1.7)
        market_width = Inches(5.0)
        market_height = Inches(1.8)
        
        # Create heading
        label_box = slide.shapes.add_textbox(
            market_left, market_top, market_width, Inches(0.5))
        label_p = label_box.text_frame.paragraphs[0]
        label_p.text = "Market Size"
        label_p.alignment = PP_ALIGN.LEFT
        
        for run in label_p.runs:
            run.font.name = FONTS["body"]
            run.font.size = Pt(16)
            run.font.bold = True
            run.font.color.rgb = RGBColor.from_string(COLORS["black"])
        
        # Create value with larger font
        value_box = slide.shapes.add_textbox(
            market_left, Inches(market_top.inches + 0.6), market_width, Inches(1.0))
        value_p = value_box.text_frame.paragraphs[0]
        value_p.text = content["market_size"]
        value_p.alignment = PP_ALIGN.LEFT
        
        for run in value_p.runs:
            run.font.name = FONTS["body"]
            run.font.size = Pt(28)
            run.font.bold = True
            run.font.color.rgb = RGBColor(0, 100, 150)  # Blue for emphasis
    
    # Create growth rate callout (right side)
    if "growth_rate" in content:
        growth_left = Inches(7.0)
        growth_top = Inches(1.7)
        growth_width = Inches(5.0)
        growth_height = Inches(1.8)
        
        # Create heading
        label_box = slide.shapes.add_textbox(
            growth_left, growth_top, growth_width, Inches(0.5))
        label_p = label_box.text_frame.paragraphs[0]
        label_p.text = "Growth Rate"
        label_p.alignment = PP_ALIGN.LEFT
        
        for run in label_p.runs:
            run.font.name = FONTS["body"]
            run.font.size = Pt(16)
            run.font.bold = True
            run.font.color.rgb = RGBColor.from_string(COLORS["black"])
        
        # Create value with larger font
        value_box = slide.shapes.add_textbox(
            growth_left, Inches(growth_top.inches + 0.6), growth_width, Inches(1.0))
        value_p = value_box.text_frame.paragraphs[0]
        value_p.text = content["growth_rate"]
        value_p.alignment = PP_ALIGN.LEFT
        
        for run in value_p.runs:
            run.font.name = FONTS["body"]
            run.font.size = Pt(28)
            run.font.bold = True
            run.font.color.rgb = RGBColor(0, 100, 150)  # Blue for emphasis
    
    # Add description paragraph
    if "description" in content or "paragraph" in content:
        description_text = content.get("description", content.get("paragraph", ""))
        if product_name:
            description_text = description_text.replace("[Product Name]", product_name)
        
        desc_left = Inches(1.0)
        desc_top = Inches(3.8)  # Below the stats
        desc_width = Inches(11.33)
        desc_height = Inches(2.5)
        
        desc_box = slide.shapes.add_textbox(
            desc_left, desc_top, desc_width, desc_height)
        desc_tf = desc_box.text_frame
        desc_tf.word_wrap = True
        desc_tf.auto_size = False
        
        desc_p = desc_tf.paragraphs[0]
        # Truncate if needed
        max_chars = 500
        if len(description_text) > max_chars:
            description_text = description_text[:max_chars-3] + "..."
        desc_p.text = description_text
        desc_p.alignment = PP_ALIGN.LEFT
        
        for run in desc_p.runs:
            run.font.name = FONTS["body"]
            run.font.size = Pt(18)
            run.font.color.rgb = RGBColor.from_string(COLORS["black"])
    
    # Add source citation if available
    if "source" in content:
        source_left = Inches(1.0)
        source_top = Inches(6.5)
        source_width = Inches(11.33)
        source_height = Inches(0.5)
        
        source_box = slide.shapes.add_textbox(
            source_left, source_top, source_width, source_height)
        source_p = source_box.text_frame.paragraphs[0]
        source_p.text = content["source"]
        source_p.alignment = PP_ALIGN.LEFT
        
        for run in source_p.runs:
            run.font.name = FONTS["body"]
            run.font.size = Pt(12)
            run.font.italic = True
            run.font.color.rgb = RGBColor(100, 100, 100)  # Gray

def create_roadmap_slide(prs: Presentation, content: Dict[str, Any], presentation_context: Dict[str, Any] = None):
    """Create the product roadmap slide with proper layout."""
    slide_layout = prs.slide_layouts[6]  # Blank layout
    slide = prs.slides.add_slide(slide_layout)
    
    # Set background to white for consistency
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor.from_string(COLORS["white"])
    
    # Add title with consistent positioning
    title_left = Inches(0.5)
    title_top = Inches(0.5)
    title_width = Inches(12.33)
    title_height = Inches(0.8)
    
    title_box = slide.shapes.add_textbox(title_left, title_top, title_width, title_height)
    title_tf = title_box.text_frame
    title_tf.word_wrap = True
    title_tf.auto_size = False
    title_p = title_tf.paragraphs[0]
    
    # Get product name from presentation context
    product_name = ""
    if presentation_context and "metadata" in presentation_context:
        product_name = presentation_context["metadata"].get("product_name", "")
    
    # Set title with product name substitution
    title_text = content.get("title", "Product Roadmap")
    if product_name:
        title_text = title_text.replace("[Product Name]", product_name)
    
    title_p.text = title_text
    title_p.alignment = PP_ALIGN.LEFT
    
    # Apply title formatting
    for run in title_p.runs:
        run.font.name = FONTS["title"]
        run.font.size = Pt(32)
        run.font.bold = True
        run.font.color.rgb = RGBColor.from_string(COLORS["black"])
    
    # Get phases from content
    phases = []
    if "phases" in content and isinstance(content["phases"], list):
        phases = content["phases"]
    elif "milestones" in content and isinstance(content["milestones"], list):
        phases = content["milestones"]
        
    # Ensure we have exactly 3 phases for consistent layout
    if len(phases) != 3:
        if len(phases) < 3:
            # Add default phases if we have too few
            default_phases = [
                {"name": "Phase 1", "items": ["Initial launch", "Core features"]},
                {"name": "Phase 2", "items": ["Advanced capabilities", "Expanded integrations"]},
                {"name": "Phase 3", "items": ["Enterprise features", "Global expansion"]}
            ]
            while len(phases) < 3:
                phases.append(default_phases[len(phases)])
        else:
            # Truncate if too many phases
            phases = phases[:3]
    
    # Draw horizontal timeline
    timeline_left = Inches(1.0)
    timeline_top = Inches(2.5)  # Below title
    timeline_width = Inches(11.33)
    timeline_height = Inches(0.05)  # Thin line
    
    line_box = slide.shapes.add_textbox(
        timeline_left, timeline_top, timeline_width, timeline_height)
    line_box.fill.solid()
    line_box.fill.fore_color.rgb = RGBColor(0, 120, 215)  # Blue
    
    # Position the 3 phases along the timeline
    phase_positions = [Inches(1.0), Inches(5.5), Inches(10.0)]  # Left, center, right
    
    for i, phase in enumerate(phases):
        if i < 3:  # Limit to 3 phases
            phase_left = phase_positions[i]
            
            # Add circle marker on timeline
            marker_width = Inches(0.3)
            marker_height = Inches(0.3)
            marker_left = phase_left + Inches(0.85)  # Center it
            marker_top = timeline_top - Inches(0.125)  # Center vertically on timeline
            
            marker_box = slide.shapes.add_textbox(
                marker_left, marker_top, marker_width, marker_height)
            marker_box.fill.solid()
            marker_box.fill.fore_color.rgb = RGBColor(0, 120, 215)  # Blue
            
            # Add phase name above timeline
            phase_name = phase.get("name", f"Phase {i+1}")
            name_left = phase_left
            name_top = timeline_top - Inches(0.8)  # Above timeline
            name_width = Inches(2.0)
            name_height = Inches(0.5)
            
            name_box = slide.shapes.add_textbox(
                name_left, name_top, name_width, name_height)
            name_p = name_box.text_frame.paragraphs[0]
            name_p.text = phase_name
            name_p.alignment = PP_ALIGN.CENTER
            
            for run in name_p.runs:
                run.font.name = FONTS["body"]
                run.font.size = Pt(18)
                run.font.bold = True
                run.font.color.rgb = RGBColor.from_string(COLORS["black"])
            
            # Add phase items below timeline
            items = phase.get("items", [])
            items_left = phase_left
            items_top = timeline_top + Inches(0.5)  # Below timeline
            items_width = Inches(2.0)
            items_height = Inches(2.5)
            
            if items:
                items_box = slide.shapes.add_textbox(
                    items_left, items_top, items_width, items_height)
                items_tf = items_box.text_frame
                items_tf.word_wrap = True
                items_tf.auto_size = False
                
                # Add up to 3 items
                max_items = min(3, len(items))
                for j in range(max_items):
                    item_text = items[j]
                    # Truncate if needed
                    if len(item_text) > 25:
                        item_text = item_text[:22] + "..."
                    
                    if j == 0:
                        item_p = items_tf.paragraphs[0]
                    else:
                        item_p = items_tf.add_paragraph()
                        
                    item_p.text = f"• {item_text}"
                    item_p.alignment = PP_ALIGN.LEFT
                    
                    for run in item_p.runs:
                        run.font.name = FONTS["body"]
                        run.font.size = Pt(14)
                        run.font.color.rgb = RGBColor.from_string(COLORS["black"])
                    
                    # Add spacing
                    item_p.space_before = Pt(6)
                    item_p.space_after = Pt(6)

def create_team_slide(prs: Presentation, content: Dict[str, Any], presentation_context: Dict[str, Any] = None):
    """Create the team slide with proper layout."""
    slide_layout = prs.slide_layouts[6]  # Blank layout
    slide = prs.slides.add_slide(slide_layout)
    
    # Set background to white for consistency
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor.from_string(COLORS["white"])
    
    # Add title with consistent positioning
    title_left = Inches(0.5)
    title_top = Inches(0.5)
    title_width = Inches(12.33)
    title_height = Inches(0.8)
    
    title_box = slide.shapes.add_textbox(title_left, title_top, title_width, title_height)
    title_tf = title_box.text_frame
    title_tf.word_wrap = True
    title_tf.auto_size = False
    title_p = title_tf.paragraphs[0]
    
    # Get company name from content or presentation context
    company_name = ""
    if presentation_context and "metadata" in presentation_context:
        company_name = presentation_context["metadata"].get("company_name", "")
    
    # Add title text - use content title if available, otherwise generate from company name
    if "title" in content and content["title"]:
        title_text = content["title"]
    else:
        title_text = f"{company_name} Leadership Team" if company_name else "Leadership Team"
    
    title_p.text = title_text
    title_p.alignment = PP_ALIGN.CENTER  # Center title for team slide
    
    # Apply title formatting
    for run in title_p.runs:
        run.font.name = FONTS["title"]
        run.font.size = Pt(32)
        run.font.bold = True
        run.font.color.rgb = RGBColor.from_string(COLORS["black"])
    
    # Get team members
    team_members = content.get("team_members", content.get("members", []))
    
    # If no team members specified, create placeholders
    if not team_members:
        team_members = [
            {"name": "CEO", "role": "Chief Executive Officer"},
            {"name": "CTO", "role": "Chief Technology Officer"},
            {"name": "CFO", "role": "Chief Financial Officer"},
            {"name": "COO", "role": "Chief Operating Officer"}
        ]
    
    # Limit to 4 members for layout reasons
    team_members = team_members[:4]
    
    # Calculate grid layout
    if len(team_members) == 1:
        cols, rows = 1, 1
        x_positions = [6.67]  # Center
        y_positions = [2.0]
    elif len(team_members) == 2:
        cols, rows = 2, 1
        x_positions = [3.33, 10.0]
        y_positions = [2.0]
    elif len(team_members) == 3:
        cols, rows = 3, 1
        x_positions = [2.0, 6.67, 11.33]
        y_positions = [2.0]
    else:  # 4 members
        cols, rows = 2, 2
        x_positions = [3.33, 10.0]
        y_positions = [2.0, 4.5]
    
    # Add team members
    for i, member in enumerate(team_members):
        if i >= len(x_positions) * len(y_positions):
            break  # Safety check
            
        row = i // cols
        col = i % cols
        
        # Find correct position
        if row < len(y_positions) and col < len(x_positions):
            x_pos = x_positions[col] - 1.5  # Center adjust for text box
            y_pos = y_positions[row]
            
            # Team member icon (using placeholder icon)
            icon_box = slide.shapes.add_textbox(
                Inches(x_pos), Inches(y_pos), Inches(3.0), Inches(1.0))
            icon_p = icon_box.text_frame.paragraphs[0]
            icon_p.text = "👤"  # Person icon
            icon_p.alignment = PP_ALIGN.CENTER
            
            for run in icon_p.runs:
                run.font.size = Pt(72)
                
            # Name below icon
            name = member.get("name", f"Team Member {i+1}")
            name_box = slide.shapes.add_textbox(
                Inches(x_pos), Inches(y_pos + 1.2), Inches(3.0), Inches(0.5))
            name_p = name_box.text_frame.paragraphs[0]
            name_p.text = name
            name_p.alignment = PP_ALIGN.CENTER
            
            for run in name_p.runs:
                run.font.name = FONTS["body"]
                run.font.size = Pt(20)
                run.font.bold = True
                run.font.color.rgb = RGBColor.from_string(COLORS["black"])
            
            # Role below name
            role = member.get("role", "")
            if role:
                role_box = slide.shapes.add_textbox(
                    Inches(x_pos), Inches(y_pos + 1.7), Inches(3.0), Inches(0.5))
                role_p = role_box.text_frame.paragraphs[0]
                # Truncate role if too long
                if len(role) > 40:
                    role = role[:37] + "..."
                role_p.text = role
                role_p.alignment = PP_ALIGN.CENTER
                
                for run in role_p.runs:
                    run.font.name = FONTS["body"]
                    run.font.size = Pt(16)
                    run.font.color.rgb = RGBColor.from_string(COLORS["black"])
    
    # Add tagline at bottom if exists
    if "tagline" in content and content["tagline"]:
        tagline = content["tagline"]
        tagline_box = slide.shapes.add_textbox(
            Inches(1.0), Inches(6.5), Inches(11.33), Inches(0.5))
        tagline_p = tagline_box.text_frame.paragraphs[0]
        tagline_p.text = tagline
        tagline_p.alignment = PP_ALIGN.CENTER
        
        for run in tagline_p.runs:
            run.font.name = FONTS["body"]
            run.font.size = Pt(18)
            run.font.italic = True
            run.font.color.rgb = RGBColor.from_string(COLORS["black"])