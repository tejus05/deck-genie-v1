import io
from typing import Dict, Any, List, Optional
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from image_fetcher import fetch_image_for_slide, get_slide_icon
from utils import FONTS, COLORS, FONT_SIZES, MARGINS, CONTENT_AREA, IMAGE_AREA, match_icon_to_feature

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
    
    # Set slide size to standard (4:3) - 10 x 7.5 inches
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(7.5)
    
    # Initialize image cache for future customizations
    if 'original_images_cache' not in st.session_state:
        st.session_state.original_images_cache = {}
    
    # Create slides and cache images during creation
    create_title_slide(prs, content['title_slide'])
    create_problem_slide(prs, content['problem_slide'], content)
    create_solution_slide(prs, content['solution_slide'], content) 
    create_features_slide(prs, content['features_slide'])
    create_advantage_slide(prs, content['advantage_slide'], content)
    create_audience_slide(prs, content['audience_slide'], content)
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
    # Use two content layout
    slide_layout = prs.slide_layouts[3]  # Two Content layout
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
    
    # Left content - bullet points
    left_content = slide.placeholders[1]
    tf = left_content.text_frame
    
    # Clear any existing text
    if tf.paragraphs:
        p = tf.paragraphs[0]
        p.text = ""
    else:
        p = tf.add_paragraph()
    
    # Add bullet points
    for i, bullet in enumerate(content["bullets"]):
        if i == 0:
            p.text = bullet
        else:
            p = tf.add_paragraph()
            p.text = bullet
        p.level = 0
        
    apply_text_formatting(tf)
    
    # Right content - image
    right_content = slide.placeholders[2]
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
        # Get placeholder dimensions
        placeholder_width = right_content.width
        placeholder_height = right_content.height
        placeholder_left = right_content.left
        placeholder_top = right_content.top

        # Add picture directly to slide (not to placeholder)
        pic = slide.shapes.add_picture(
            image_data,
            placeholder_left,
            placeholder_top,
            placeholder_width,
            placeholder_height
        )
    else:
        # Fallback to icon if image fetch fails
        icon_info = get_slide_icon("problem")
        tf = right_content.text_frame
        p = tf.add_paragraph()
        p.text = icon_info["icon"]
        p.alignment = PP_ALIGN.CENTER
        apply_text_formatting(tf, size=72, alignment=PP_ALIGN.CENTER)

def create_solution_slide(prs: Presentation, content: Dict[str, Any], presentation_context: Dict[str, Any] = None):
    """Create the solution overview slide."""
    # Use two content layout
    slide_layout = prs.slide_layouts[3]  # Two Content layout
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
    
    # Left content - paragraph
    left_content = slide.placeholders[1]
    tf = left_content.text_frame
    p = tf.paragraphs[0]
    p.text = content["paragraph"]
    
    apply_text_formatting(tf)
    
    # Right content - icon or image
    right_content = slide.placeholders[2]
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
        # Get placeholder dimensions
        placeholder_width = right_content.width
        placeholder_height = right_content.height
        placeholder_left = right_content.left
        placeholder_top = right_content.top

        # Add picture directly to slide (not to placeholder)
        pic = slide.shapes.add_picture(
            image_data,
            placeholder_left,
            placeholder_top,
            placeholder_width,
            placeholder_height
        )
    else:
        # Use lightbulb icon as fallback
        tf = right_content.text_frame
        p = tf.add_paragraph()
        p.text = get_slide_icon("solution")["icon"]
        p.alignment = PP_ALIGN.CENTER
        apply_text_formatting(tf, size=72, alignment=PP_ALIGN.CENTER)

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
    
    # Set column widths
    table.columns[0].width = Inches(0.5)  # Icon column
    table.columns[1].width = Inches(9.0)  # Feature text column
    
    # Add features to table
    for i, feature in enumerate(features):
        # Icon cell
        icon_cell = table.cell(i, 0)
        icon = match_icon_to_feature(feature)
        icon_cell.text = icon
        apply_text_formatting(icon_cell.text_frame, size=16, alignment=PP_ALIGN.CENTER)
        
        # Feature cell
        feature_cell = table.cell(i, 1)
        feature_cell.text = feature
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
    # Use two content layout
    slide_layout = prs.slide_layouts[3]  # Two Content layout
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
    
    # Left content - bullet points
    left_content = slide.placeholders[1]
    tf = left_content.text_frame
    
    # Clear any existing text
    if tf.paragraphs:
        p = tf.paragraphs[0]
        p.text = ""
    else:
        p = tf.add_paragraph()
    
    # Add bullet points
    for i, bullet in enumerate(content["bullets"]):
        if i == 0:
            p.text = bullet
        else:
            p = tf.add_paragraph()
            p.text = bullet
        p.level = 0
        
    apply_text_formatting(tf)
    
    # Right content - image
    right_content = slide.placeholders[2]
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
        # Get placeholder dimensions
        placeholder_width = right_content.width
        placeholder_height = right_content.height
        placeholder_left = right_content.left
        placeholder_top = right_content.top

        # Add picture directly to slide (not to placeholder)
        pic = slide.shapes.add_picture(
            image_data,
            placeholder_left,
            placeholder_top,
            placeholder_width,
            placeholder_height
        )
    else:
        # Fallback to trophy icon
        tf = right_content.text_frame
        p = tf.add_paragraph()
        p.text = get_slide_icon("advantage")["icon"]
        p.alignment = PP_ALIGN.CENTER
        apply_text_formatting(tf, size=72, alignment=PP_ALIGN.CENTER)

def create_audience_slide(prs: Presentation, content: Dict[str, Any], presentation_context: Dict[str, Any] = None):
    """Create the target audience slide."""
    # Use two content layout
    slide_layout = prs.slide_layouts[3]  # Two Content layout
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
    
    # Left content - paragraph
    left_content = slide.placeholders[1]
    tf = left_content.text_frame
    p = tf.paragraphs[0]
    p.text = content["paragraph"]
    
    apply_text_formatting(tf)
    
    # Right content - image
    right_content = slide.placeholders[2]
    image_data = fetch_image_for_slide("audience", presentation_context, use_placeholders=False)
    
    # Cache the image for future customizations
    import streamlit as st
    if image_data and 'original_images_cache' in st.session_state:
        # Save a copy of the image data to cache
        image_data.seek(0)
        cached_data = image_data.read()
        st.session_state.original_images_cache['audience_slide'] = cached_data
        # Reset the image data for use
        image_data.seek(0)
    
    if image_data:
        # Get placeholder dimensions
        placeholder_width = right_content.width
        placeholder_height = right_content.height
        placeholder_left = right_content.left
        placeholder_top = right_content.top

        # Add picture directly to slide (not to placeholder)
        pic = slide.shapes.add_picture(
            image_data,
            placeholder_left,
            placeholder_top,
            placeholder_width,
            placeholder_height
        )
    else:
        # Fallback to people icon
        tf = right_content.text_frame
        p = tf.add_paragraph()
        p.text = get_slide_icon("audience")["icon"]
        p.alignment = PP_ALIGN.CENTER
        apply_text_formatting(tf, size=72, alignment=PP_ALIGN.CENTER)

def create_cta_slide(prs: Presentation, content: Dict[str, Any]):
    """Create the call to action slide."""
    # Use title only layout
    slide_layout = prs.slide_layouts[5]  # Title Only layout
    slide = prs.slides.add_slide(slide_layout)
    
    # Title
    title = slide.shapes.title
    title.text = "Get Started"
    apply_text_formatting(
        title.text_frame, 
        font_name=FONTS["title"], 
        size=FONT_SIZES["title"], 
        bold=True
    )
    
    # CTA text (center of slide)
    left = Inches(0.5)
    top = Inches(2.5)
    width = Inches(9.0)
    height = Inches(1.5)
    
    text_box = slide.shapes.add_textbox(left, top, width, height)
    tf = text_box.text_frame
    tf.word_wrap = True
    
    p = tf.add_paragraph()
    p.text = content["call_to_action"]
    p.alignment = PP_ALIGN.CENTER
    
    apply_text_formatting(
        tf, 
        font_name=FONTS["title"], 
        size=FONT_SIZES["call_to_action"], 
        bold=True,
        alignment=PP_ALIGN.CENTER
    )
    
    # Optional grey stripe at bottom
    # Use a textbox with gray background instead of rectangle since add_rectangle is not available
    left = Inches(0)
    top = Inches(6.5)
    width = Inches(10)
    height = Inches(0.5)
    
    stripe_box = slide.shapes.add_textbox(left, top, width, height)
    stripe_box.fill.solid()
    stripe_box.fill.fore_color.rgb = RGBColor.from_string(COLORS["gray"])
    
    # Make the textbox border invisible
    stripe_box.line.fill.background()