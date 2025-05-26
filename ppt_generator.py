import io
import streamlit as st
from typing import Dict, Any, List, Optional
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from image_fetcher import fetch_image_for_slide, get_slide_icon
from utils import FONTS, COLORS, FONT_SIZES, MARGINS, CONTENT_AREA, IMAGE_AREA, match_icon_to_feature
# Import additional slide creation functions
from ppt_generator_additions import create_market_slide, create_roadmap_slide, create_team_slide

def create_presentation(content: Dict[str, Any], filename: str) -> io.BytesIO:
    """
    Create a PowerPoint presentation from generated content.
    
    Args:
        content: Dictionary containing content for all slides
        filename: Name to give the file
        
    Returns:
        BytesIO object containing the presentation
    """
    import streamlit as st
    
    prs = Presentation()
    
    # Set slide size to widescreen (16:9) - 13.33 x 7.5 inches
    prs.slide_width = Inches(13.33)
    prs.slide_height = Inches(7.5)
    
    # Initialize image cache for future customizations
    if 'original_images_cache' not in st.session_state:
        st.session_state.original_images_cache = {}
    
    # Create required slides first
    create_title_slide(prs, content['title_slide'])
    
    # Create other slides if they exist in the content
    if 'problem_slide' in content:
        create_problem_slide(prs, content['problem_slide'], content)
        
    if 'solution_slide' in content:
        create_solution_slide(prs, content['solution_slide'], content)
    
    if 'features_slide' in content:
        create_features_slide(prs, content['features_slide'])
    
    if 'advantage_slide' in content:
        create_advantage_slide(prs, content['advantage_slide'], content)
    
    if 'audience_slide' in content:
        create_audience_slide(prs, content['audience_slide'], content)
        
    # Add optional slides if they exist
    if 'market_slide' in content:
        create_market_slide(prs, content['market_slide'], content)
        
    if 'roadmap_slide' in content:
        create_roadmap_slide(prs, content['roadmap_slide'], content)
        
    if 'team_slide' in content:
        create_team_slide(prs, content['team_slide'], content)
    
    # CTA slide should always be last
    create_cta_slide(prs, content['cta_slide'])
    
    # Save to BytesIO
    output = io.BytesIO()
    prs.save(output)
    output.seek(0)
    
    return output

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

def create_title_slide(prs: Presentation, content: Dict[str, str]):
    """Create the title slide."""
    # Use title slide layout
    slide_layout = prs.slide_layouts[0]  # Title Slide layout
    slide = prs.slides.add_slide(slide_layout)
    
    # Set background to white
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor.from_string(COLORS["white"])
    
    # Title
    title = slide.shapes.title
    title.text = content["title"]
    apply_text_formatting(
        title.text_frame, 
        font_name=FONTS["title"], 
        size=FONT_SIZES["title_slide"], 
        bold=True, 
        alignment=PP_ALIGN.CENTER
    )
    
    # Company name (subtitle)
    subtitle = slide.placeholders[1]
    subtitle.text = content["subtitle"]
    apply_text_formatting(
        subtitle.text_frame, 
        font_name=FONTS["body"], 
        size=FONT_SIZES["subtitle"], 
        bold=False, 
        alignment=PP_ALIGN.CENTER
    )

def create_problem_slide(prs: Presentation, content: Dict[str, Any], presentation_context: Dict[str, Any] = None):
    """Create the problem statement slide."""
    # Use blank layout to avoid unwanted placeholder elements
    slide_layout = prs.slide_layouts[6]  # Blank layout
    slide = prs.slides.add_slide(slide_layout)
    
    # Add title manually
    title_left = Inches(1.0)
    title_top = Inches(0.5)
    title_width = Inches(11.33)
    title_height = Inches(1.0)
    
    title_box = slide.shapes.add_textbox(title_left, title_top, title_width, title_height)
    title_tf = title_box.text_frame
    title_p = title_tf.paragraphs[0]
    title_p.text = content["title"]
    title_p.alignment = PP_ALIGN.LEFT
    
    # Apply title formatting and ensure it stays on one line
    for run in title_p.runs:
        run.font.name = FONTS["title"]
        run.font.size = Pt(FONT_SIZES["title"])
        run.font.bold = True
        run.font.color.rgb = RGBColor.from_string(COLORS["black"])
    # Ensure title wraps naturally if it's too long
    title_tf.word_wrap = True
      # Create custom text box for bullet points with improved spacing
    left = Inches(CONTENT_AREA["left"])
    top = Inches(CONTENT_AREA["top"])
    width = Inches(CONTENT_AREA["width"])  # Use proper content width to prevent overlap
    height = Inches(CONTENT_AREA["height"])
    
    text_box = slide.shapes.add_textbox(left, top, width, height)
    tf = text_box.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.1)  # Small left margin for bullet points
    tf.margin_right = Inches(0.1)  # Small right margin to prevent text overflow
    tf.margin_top = Inches(0.1)
    tf.margin_bottom = Inches(0.1)
    
    # Clear any existing text and add bullet points
    if tf.paragraphs:
        p = tf.paragraphs[0]
        p.text = ""
    else:
        p = tf.add_paragraph()
        
    # Add bullet points with proper formatting and bullet characters
    # Handle different field names from slide content generators
    bullets = []
    if "pain_points" in content:
        bullets = content["pain_points"]
    elif "bullets" in content:
        bullets = content["bullets"]
    elif "differentiators" in content:
        # Handle differentiators format (objects with point and description)
        bullets = [item["point"] if isinstance(item, dict) else str(item) for item in content["differentiators"]]
    
    for i, bullet in enumerate(bullets):
        # Truncate bullet points if they're too long (prevents text overflow)
        bullet_text = bullet
        if len(bullet) > 120:  # Reasonable character limit for bullet points
            bullet_text = bullet[:117] + "..."
            
        if i == 0:
            p.text = f"• {bullet_text}"  # Add bullet character
        else:
            p = tf.add_paragraph()
            p.text = f"• {bullet_text}"  # Add bullet character
        
        # Apply bullet formatting
        p.level = 0
        p.alignment = PP_ALIGN.LEFT
        
        # Apply font formatting
        p.font.name = FONTS["body"]
        p.font.size = Pt(FONT_SIZES["body"])
        
        # Apply spacing for better readability
        p.space_before = Pt(6)  # Add space before each bullet
        p.space_after = Pt(6)   # Add space after each bullet
    
    apply_text_formatting(tf)
    
    # Add image directly to slide without using placeholder
    image_data = fetch_image_for_slide("problem", presentation_context, use_placeholders=False)
    
    # Cache the image for future customizations
    import streamlit as st
    if image_data and 'original_images_cache' in st.session_state:
        # Save a copy of the image data to cache
        image_data.seek(0)
        cached_data = image_data.read()
        st.session_state.original_images_cache['problem_slide'] = cached_data
        # Reset the image data for use
        image_data.seek(0)
    
    if image_data:
        # Position the image on the right side of the 16:9 slide
        pic = slide.shapes.add_picture(
            image_data,
            Inches(IMAGE_AREA["left"]),     # Positioned further right for 16:9 format
            Inches(IMAGE_AREA["top"]),
            Inches(IMAGE_AREA["width"]),    # Adjusted width for 16:9
            Inches(IMAGE_AREA["height"])
        )
    else:
        # Add fallback icon as a separate shape
        icon_left = Inches(IMAGE_AREA["left"] + IMAGE_AREA["width"]/2 - 0.5)
        icon_top = Inches(IMAGE_AREA["top"] + IMAGE_AREA["height"]/2 - 0.5)
        icon_box = slide.shapes.add_textbox(icon_left, icon_top, Inches(1), Inches(1))
        icon_tf = icon_box.text_frame
        icon_p = icon_tf.paragraphs[0]
        icon_info = get_slide_icon("problem")
        icon_p.text = icon_info["icon"]
        icon_p.alignment = PP_ALIGN.CENTER
        apply_text_formatting(icon_tf, size=72, alignment=PP_ALIGN.CENTER)

def create_solution_slide(prs: Presentation, content: Dict[str, Any], presentation_context: Dict[str, Any] = None):
    """Create the solution overview slide."""
    # Use blank layout to avoid unwanted placeholder elements
    slide_layout = prs.slide_layouts[6]  # Blank layout
    slide = prs.slides.add_slide(slide_layout)
    
    # Add title manually
    title_left = Inches(1.0)
    title_top = Inches(0.5)
    title_width = Inches(11.33)
    title_height = Inches(1.0)
    
    title_box = slide.shapes.add_textbox(title_left, title_top, title_width, title_height)
    title_tf = title_box.text_frame
    title_p = title_tf.paragraphs[0]
    title_p.text = content["title"]
    title_p.alignment = PP_ALIGN.LEFT
    
    # Apply title formatting and ensure it stays on one line
    for run in title_p.runs:
        run.font.name = FONTS["title"]
        run.font.size = Pt(FONT_SIZES["title"])
        run.font.bold = True
        run.font.color.rgb = RGBColor.from_string(COLORS["black"])
    # Ensure title wraps naturally if it's too long
    title_tf.word_wrap = True
    
    # Create custom text box for paragraph content with better spacing
    left = Inches(CONTENT_AREA["left"])
    top = Inches(CONTENT_AREA["top"])
    width = Inches(CONTENT_AREA["width"])  # Use defined width without expansion for proper spacing
    height = Inches(CONTENT_AREA["height"])
    text_box = slide.shapes.add_textbox(left, top, width, height)
    tf = text_box.text_frame
    tf.word_wrap = True
    tf.margin_right = Inches(0.1)  # Add right margin to prevent text overflow

    p = tf.paragraphs[0]
    # Truncate paragraph if it's too long to prevent text overflow
    # Handle different field names from solution slide content generator
    paragraph_text = ""
    if "paragraph" in content:
        paragraph_text = content["paragraph"]
    elif "description" in content:
        paragraph_text = content["description"]
    elif "value_proposition" in content:
        paragraph_text = content["value_proposition"]
    
    if len(paragraph_text) > 550:  # Reasonable limit for paragraphs
        paragraph_text = paragraph_text[:547] + "..."
    p.text = paragraph_text
    
    # Format paragraph with proper spacing
    p.space_before = Pt(6)
    p.space_after = Pt(6)
    p.line_spacing = 1.2  # Better line spacing
    
    apply_text_formatting(tf)
    
    # Add image directly to slide without using placeholder
    image_data = fetch_image_for_slide("solution", presentation_context, use_placeholders=False)
    
    # Cache the image for future customizations
    import streamlit as st
    if image_data and 'original_images_cache' in st.session_state:
        # Save a copy of the image data to cache
        image_data.seek(0)
        cached_data = image_data.read()
        st.session_state.original_images_cache['solution_slide'] = cached_data
        # Reset the image data for use
        image_data.seek(0)
    
    if image_data:
        # Position the image on the right side of the 16:9 slide
        pic = slide.shapes.add_picture(
            image_data,
            Inches(IMAGE_AREA["left"]),     # Positioned further right for 16:9 format
            Inches(IMAGE_AREA["top"]),
            Inches(IMAGE_AREA["width"]),    # Adjusted width for 16:9
            Inches(IMAGE_AREA["height"])
        )
    else:
        # Add fallback icon as a separate shape
        icon_left = Inches(IMAGE_AREA["left"] + IMAGE_AREA["width"]/2 - 0.5)
        icon_top = Inches(IMAGE_AREA["top"] + IMAGE_AREA["height"]/2 - 0.5)
        icon_box = slide.shapes.add_textbox(icon_left, icon_top, Inches(1), Inches(1))
        icon_tf = icon_box.text_frame
        icon_p = icon_tf.paragraphs[0]
        icon_p.text = get_slide_icon("solution")["icon"]
        icon_p.alignment = PP_ALIGN.CENTER
        apply_text_formatting(icon_tf, size=72, alignment=PP_ALIGN.CENTER)

def create_features_slide(prs: Presentation, content: Dict[str, Any]):
    """Create the key features slide."""
    # Use title only layout and add a table
    slide_layout = prs.slide_layouts[5]  # Title Only layout
    slide = prs.slides.add_slide(slide_layout)
      # Title
    title = slide.shapes.title
    title.text = content["title"]
    apply_text_formatting(
        title.text_frame, 
        font_name=FONTS["title"], 
        size=FONT_SIZES["title"], 
        bold=True
    )
    
    # Create table for features
    features = content["features"]
    rows, cols = len(features), 2
    
    left = Inches(CONTENT_AREA["left"])
    top = Inches(CONTENT_AREA["top"])
    width = Inches(CONTENT_AREA["width"] + IMAGE_AREA["width"])
    height = Inches(CONTENT_AREA["height"])
    table = slide.shapes.add_table(rows, cols, left, top, width, height).table
    # Set column widths - adjusted for 16:9 format
    table.columns[0].width = Inches(0.5)  # Icon column
    table.columns[1].width = Inches(11.83)  # Feature text column - wider for 16:9
    
    # Add features to table
    for i, feature in enumerate(features):
        # Handle both string and dictionary formats for features
        if isinstance(feature, dict):
            # Extract feature text from dictionary format
            feature_text = feature.get('feature', feature.get('name', feature.get('title', str(feature))))
            icon_text = feature_text
        else:
            # Handle string format
            feature_text = str(feature)
            icon_text = feature_text
        
        # Icon cell
        icon_cell = table.cell(i, 0)
        icon = match_icon_to_feature(icon_text)
        icon_cell.text = icon
        apply_text_formatting(icon_cell.text_frame, size=16, alignment=PP_ALIGN.CENTER)
        
        # Feature cell
        feature_cell = table.cell(i, 1)
        feature_cell.text = feature_text
        apply_text_formatting(feature_cell.text_frame)
        
        # Set row height
        table.rows[i].height = Inches(CONTENT_AREA["height"] / len(features))
    
    # Style table with no fill - keep it minimal without borders
    # since border styling is causing issues
    for cell in table.iter_cells():
        cell.fill.background()  # No fill
        # We're not setting any borders to avoid API compatibility issues

def create_advantage_slide(prs: Presentation, content: Dict[str, Any], presentation_context: Dict[str, Any] = None):
    """Create the competitive advantage slide."""
    # Use blank layout to avoid unwanted placeholder elements
    slide_layout = prs.slide_layouts[6]  # Blank layout
    slide = prs.slides.add_slide(slide_layout)
    
    # Add title manually
    title_left = Inches(1.0)
    title_top = Inches(0.5)
    title_width = Inches(11.33)
    title_height = Inches(1.0)
    
    title_box = slide.shapes.add_textbox(title_left, title_top, title_width, title_height)
    title_tf = title_box.text_frame
    title_p = title_tf.paragraphs[0]
    title_p.text = content["title"]
    title_p.alignment = PP_ALIGN.LEFT
    
    # Apply title formatting and ensure it stays on one line
    for run in title_p.runs:
        run.font.name = FONTS["title"]
        run.font.size = Pt(FONT_SIZES["title"])
        run.font.bold = True
        run.font.color.rgb = RGBColor.from_string(COLORS["black"])
      
    # Ensure title wraps naturally if it's too long
    title_tf.word_wrap = True
      
    # Create custom text box for bullet points with better spacing
    left = Inches(CONTENT_AREA["left"])
    top = Inches(CONTENT_AREA["top"])
    width = Inches(CONTENT_AREA["width"])  # Use defined width without expansion for proper spacing
    height = Inches(CONTENT_AREA["height"])
    
    text_box = slide.shapes.add_textbox(left, top, width, height)
    tf = text_box.text_frame
    tf.word_wrap = True
    tf.margin_right = Inches(0.1)  # Add right margin to prevent text overflow
    
    # Clear any existing text and add bullet points
    if tf.paragraphs:
        p = tf.paragraphs[0]
        p.text = ""
    else:
        p = tf.add_paragraph()
          # Add bullet points with proper formatting and bullet characters
    # Handle different field names from slide content generators
    bullets = []
    if "differentiators" in content:
        # Handle differentiators format (objects with point and description)
        bullets = [item["point"] if isinstance(item, dict) else str(item) for item in content["differentiators"]]
    elif "bullets" in content:
        bullets = content["bullets"]
    elif "pain_points" in content:
        bullets = content["pain_points"]
    
    for i, bullet in enumerate(bullets):
        # Truncate bullet points if they're too long (prevents text overflow)
        bullet_text = bullet
        if len(bullet) > 120:  # Reasonable character limit for bullet points
            bullet_text = bullet[:117] + "..."
            
        if i == 0:
            p.text = f"• {bullet_text}"  # Add bullet character
        else:
            p = tf.add_paragraph()
            p.text = f"• {bullet_text}"  # Add bullet character
        
        # Apply bullet formatting
        p.level = 0
        p.alignment = PP_ALIGN.LEFT
        
        # Apply font formatting
        p.font.name = FONTS["body"]
        p.font.size = Pt(FONT_SIZES["body"])
        
        # Apply spacing for better readability
        p.space_before = Pt(6)  # Add space before each bullet
        p.space_after = Pt(6)   # Add space after each bullet
    
    apply_text_formatting(tf)
    
    # Add image directly to slide without using placeholder
    image_data = fetch_image_for_slide("advantage", presentation_context, use_placeholders=False)
    
    # Cache the image for future customizations
    import streamlit as st
    if image_data and 'original_images_cache' in st.session_state:
        # Save a copy of the image data to cache
        image_data.seek(0)
        cached_data = image_data.read()
        st.session_state.original_images_cache['advantage_slide'] = cached_data
        # Reset the image data for use
        image_data.seek(0)
    
    if image_data:
        # Position the image on the right side of the 16:9 slide
        pic = slide.shapes.add_picture(
            image_data,
            Inches(IMAGE_AREA["left"]),     # Positioned further right for 16:9 format
            Inches(IMAGE_AREA["top"]),
            Inches(IMAGE_AREA["width"]),    # Adjusted width for 16:9
            Inches(IMAGE_AREA["height"])
        )
    else:
        # Add fallback icon as a separate shape
        icon_left = Inches(IMAGE_AREA["left"] + IMAGE_AREA["width"]/2 - 0.5)
        icon_top = Inches(IMAGE_AREA["top"] + IMAGE_AREA["height"]/2 - 0.5)
        icon_box = slide.shapes.add_textbox(icon_left, icon_top, Inches(1), Inches(1))
        icon_tf = icon_box.text_frame
        icon_p = icon_tf.paragraphs[0]
        icon_p.text = get_slide_icon("advantage")["icon"]
        icon_p.alignment = PP_ALIGN.CENTER
        apply_text_formatting(icon_tf, size=72, alignment=PP_ALIGN.CENTER)

def create_audience_slide(prs: Presentation, content: Dict[str, Any], presentation_context: Dict[str, Any] = None):
    """Create the target audience slide."""
    # Use blank layout to avoid unwanted placeholder elements
    slide_layout = prs.slide_layouts[6]  # Blank layout
    slide = prs.slides.add_slide(slide_layout)
    
    # Add title manually
    title_left = Inches(1.0)
    title_top = Inches(0.5)
    title_width = Inches(11.33)
    title_height = Inches(1.0)
    
    title_box = slide.shapes.add_textbox(title_left, title_top, title_width, title_height)
    title_tf = title_box.text_frame
    title_p = title_tf.paragraphs[0]
    title_p.text = content["title"]
    title_p.alignment = PP_ALIGN.LEFT
    
    # Apply title formatting and ensure it stays on one line
    for run in title_p.runs:
        run.font.name = FONTS["title"]
        run.font.size = Pt(FONT_SIZES["title"])
        run.font.bold = True
        run.font.color.rgb = RGBColor.from_string(COLORS["black"])
    # Ensure title wraps naturally if it's too long
    title_tf.word_wrap = True
    
    # Create custom text box for paragraph content with better spacing
    left = Inches(CONTENT_AREA["left"])
    top = Inches(CONTENT_AREA["top"])
    width = Inches(CONTENT_AREA["width"])  # Use defined width without expansion for proper spacing
    height = Inches(CONTENT_AREA["height"])
      text_box = slide.shapes.add_textbox(left, top, width, height)
    tf = text_box.text_frame
    tf.word_wrap = True
    tf.margin_right = Inches(0.1)  # Add right margin to prevent text overflow
    
    p = tf.paragraphs[0]
    # Truncate paragraph if it's too long to prevent text overflow
    # Handle different field names from audience slide content generator
    paragraph_text = ""
    if "paragraph" in content:
        paragraph_text = content["paragraph"]
    elif "description" in content:
        paragraph_text = content["description"]
    elif "content" in content:
        paragraph_text = content["content"]
    
    if len(paragraph_text) > 550:  # Reasonable limit for paragraphs
        paragraph_text = paragraph_text[:547] + "..."
    p.text = paragraph_text
    
    # Format paragraph with proper spacing
    p.space_before = Pt(6)
    p.space_after = Pt(6)
    p.line_spacing = 1.2  # Better line spacing
    
    apply_text_formatting(tf)
      # Add image directly to slide without using placeholder
    image_data = fetch_image_for_slide("audience", presentation_context, use_placeholders=False)
    
    # Cache the image for future customizations
    if image_data and 'original_images_cache' in st.session_state:
        # Save a copy of the image data to cache
        image_data.seek(0)
        cached_data = image_data.read()
        st.session_state.original_images_cache['audience_slide'] = cached_data
        # Reset the image data for use
        image_data.seek(0)
    
    if image_data:
        # Position the image on the right side of the 16:9 slide
        pic = slide.shapes.add_picture(
            image_data,
            Inches(IMAGE_AREA["left"]),     # Positioned further right for 16:9 format
            Inches(IMAGE_AREA["top"]),
            Inches(IMAGE_AREA["width"]),    # Adjusted width for 16:9
            Inches(IMAGE_AREA["height"])
        )
    else:
        # Add fallback icon as a separate shape
        icon_left = Inches(IMAGE_AREA["left"] + IMAGE_AREA["width"]/2 - 0.5)
        icon_top = Inches(IMAGE_AREA["top"] + IMAGE_AREA["height"]/2 - 0.5)
        icon_box = slide.shapes.add_textbox(icon_left, icon_top, Inches(1), Inches(1))
        icon_tf = icon_box.text_frame
        icon_p = icon_tf.paragraphs[0]
        icon_p.text = get_slide_icon("audience")["icon"]
        icon_p.alignment = PP_ALIGN.CENTER
        apply_text_formatting(icon_tf, size=72, alignment=PP_ALIGN.CENTER)

def create_cta_slide(prs: Presentation, content: Dict[str, Any]):
    """Create the call to action slide."""
    # Use title only layout
    slide_layout = prs.slide_layouts[5]  # Title Only layout
    slide = prs.slides.add_slide(slide_layout)
    
    # Title
    title = slide.shapes.title
    title.text = content.get("title", "Get Started")
    apply_text_formatting(
        title.text_frame, 
        font_name=FONTS["title"], 
        size=FONT_SIZES["title"],
        bold=True
    )
    # CTA text (center of slide)
    left = Inches(1.0)  # Increased margin for better spacing    top = Inches(2.5)
    width = Inches(11.33)  # Wider for 16:9 format
    height = Inches(1.5)
    
    text_box = slide.shapes.add_textbox(left, top, width, height)
    tf = text_box.text_frame
    tf.word_wrap = True
    
    p = tf.add_paragraph()
    p.text = content.get("cta_text", content.get("call_to_action", "Contact us today!"))
    p.alignment = PP_ALIGN.CENTER
    
    apply_text_formatting(
        tf, 
        font_name=FONTS["title"], 
        size=FONT_SIZES["call_to_action"], 
        bold=True,
        alignment=PP_ALIGN.CENTER
    )
    
    # Add any supporting bullet points
    if "bullets" in content and content["bullets"]:
        bullet_box = slide.shapes.add_textbox(
            Inches(2.5),  # More indented
            Inches(4.5),  # More space below the main CTA
            Inches(8.33), # Narrower for better centering
            Inches(1.5)   # Enough for a few bullets
        )
        bullet_tf = bullet_box.text_frame
        bullet_tf.word_wrap = True
        
        for i, bullet in enumerate(content["bullets"]):
            if i == 0:
                p = bullet_tf.paragraphs[0]
            else:
                p = bullet_tf.add_paragraph()
            
            # Truncate bullet if too long
            short_bullet = bullet if len(bullet) < 50 else bullet[:47] + "..."
            p.text = f"• {short_bullet}"
            p.alignment = PP_ALIGN.CENTER
            p.space_before = Pt(12)  # More space before each bullet
            p.space_after = Pt(12)   # More space after each bullet
        
        apply_text_formatting(bullet_tf, size=FONT_SIZES["body"] - 2, bold=False, alignment=PP_ALIGN.CENTER)
    
    # Optional grey stripe at bottom
    # Use a textbox with gray background instead of rectangle since add_rectangle is not available
    left = Inches(0)
    top = Inches(6.5)
    width = Inches(13.33)  # Full width of 16:9 slide
    height = Inches(0.5)
    
    stripe_box = slide.shapes.add_textbox(left, top, width, height)
    stripe_box.fill.solid()
    stripe_box.fill.fore_color.rgb = RGBColor.from_string(COLORS["gray"])
    
    # Make the textbox border invisible
    stripe_box.line.fill.background()