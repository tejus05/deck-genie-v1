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
    text_frame.word_wrap = True
    text_frame.auto_size = False  # Prevent auto-sizing to control overflow
    
    for paragraph in text_frame.paragraphs:
        paragraph.alignment = alignment
        paragraph.line_spacing = 1.2  # Better line spacing for readability
        
        for run in paragraph.runs:
            run.font.name = font_name
            run.font.size = Pt(size)
            run.font.bold = bold
            run.font.color.rgb = RGBColor.from_string(color)

def create_title_slide(prs: Presentation, content: Dict[str, str]):
    """Create the title slide."""
    # Use blank layout for consistent control across all slides
    slide_layout = prs.slide_layouts[6]  # Blank layout
    slide = prs.slides.add_slide(slide_layout)
    
    # Set background to white
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor.from_string(COLORS["white"])
    
    # Title - centered and positioned at top third
    title_left = Inches(0.5)
    title_top = Inches(2.0)  # Move down from top
    title_width = Inches(12.33)
    title_height = Inches(2.0)  # Taller to handle multi-line titles
    
    title_box = slide.shapes.add_textbox(title_left, title_top, title_width, title_height)
    title_tf = title_box.text_frame
    title_tf.word_wrap = True
    title_tf.auto_size = False
    title_p = title_tf.paragraphs[0]
    
    # Replace product name placeholder
    product_name = content.get("product_name", "")
    clean_title = content["title"].replace("[Product Name]", product_name)
    title_p.text = clean_title
    title_p.alignment = PP_ALIGN.CENTER
    
    # Apply bold, large title formatting with proper font
    for run in title_p.runs:
        run.font.name = FONTS["title"]
        run.font.size = Pt(40)  # Larger size for title slide
        run.font.bold = True
        run.font.color.rgb = RGBColor.from_string(COLORS["black"])
    
    # Subtitle - centered and positioned below title
    subtitle_left = Inches(0.5)
    subtitle_top = Inches(4.0)  # Position below title
    subtitle_width = Inches(12.33)
    subtitle_height = Inches(1.0)
    
    subtitle_box = slide.shapes.add_textbox(subtitle_left, subtitle_top, subtitle_width, subtitle_height)
    subtitle_tf = subtitle_box.text_frame
    subtitle_tf.word_wrap = True
    subtitle_tf.auto_size = False
    subtitle_p = subtitle_tf.paragraphs[0]
    subtitle_p.text = content["subtitle"]
    subtitle_p.alignment = PP_ALIGN.CENTER
    
    # Apply subtitle formatting
    for run in subtitle_p.runs:
        run.font.name = FONTS["body"]
        run.font.size = Pt(28)  # Larger for better visibility
        run.font.color.rgb = RGBColor.from_string(COLORS["black"])

def create_problem_slide(prs: Presentation, content: Dict[str, Any], presentation_context: Dict[str, Any] = None):
    """Create the problem statement slide."""
    slide_layout = prs.slide_layouts[6]  # Blank layout
    slide = prs.slides.add_slide(slide_layout)
    
    product_name = presentation_context.get("metadata", {}).get("product_name", "")
    
    # Add title with clear positioning
    title_left = Inches(0.5)  # Consistent margin
    title_top = Inches(0.5)   # Consistent position from top
    title_width = Inches(12.33)
    title_height = Inches(0.8)
    
    title_box = slide.shapes.add_textbox(title_left, title_top, title_width, title_height)
    title_tf = title_box.text_frame
    title_tf.word_wrap = True
    title_tf.auto_size = False
    title_p = title_tf.paragraphs[0]
    
    # Replace any placeholder text
    clean_title = content["title"].replace("[Product Name]", product_name)
    title_p.text = clean_title
    title_p.alignment = PP_ALIGN.LEFT  # Left aligned for content slides
    
    # Apply title formatting
    for run in title_p.runs:
        run.font.name = FONTS["title"]
        run.font.size = Pt(32)  # Consistent title size
        run.font.bold = True
        run.font.color.rgb = RGBColor.from_string(COLORS["black"])
    
    # Create text box for bullet points
    left = Inches(0.5)
    top = Inches(1.7)  # Give space below title
    width = Inches(7.5)  # Width for text, leaving room for image
    height = Inches(5.0)
    
    text_box = slide.shapes.add_textbox(left, top, width, height)
    tf = text_box.text_frame
    tf.word_wrap = True
    tf.auto_size = False
    tf.margin_left = Inches(0.1)
    tf.margin_right = Inches(0.1)
    
    # Clear any existing text
    if tf.paragraphs:
        p = tf.paragraphs[0]
        p.text = ""
    else:
        p = tf.add_paragraph()
        
    # Handle different field names from slide content generators
    bullets = []
    if "pain_points" in content:
        bullets = content["pain_points"]
    elif "bullets" in content:
        bullets = content["bullets"]
    elif "differentiators" in content:
        bullets = [item["point"] if isinstance(item, dict) else str(item) for item in content["differentiators"]]
    
    # Keep max 3-4 bullets to prevent crowding
    max_bullets = 4
    bullets = bullets[:max_bullets]
    
    for i, bullet in enumerate(bullets):
        # Replace product name placeholders
        clean_bullet = bullet.replace("[Product Name]", product_name) if product_name else bullet
        
        # Truncate bullet points if needed
        max_chars = 85
        bullet_text = clean_bullet
        if len(clean_bullet) > max_chars:
            bullet_text = clean_bullet[:max_chars-3] + "..."
            
        if i == 0:
            p.text = f"• {bullet_text}"
        else:
            p = tf.add_paragraph()
            p.text = f"• {bullet_text}"
        
        # Apply bullet formatting
        p.level = 0
        p.alignment = PP_ALIGN.LEFT
        p.font.name = FONTS["body"]
        p.font.size = Pt(20)  # Larger for better readability
        p.space_before = Pt(10)
        p.space_after = Pt(10)
        p.line_spacing = 1.2
    
    # Add image with proper positioning
    image_data = fetch_image_for_slide("problem", presentation_context, use_placeholders=False)
    
    # Cache the image
    import streamlit as st
    if image_data and 'original_images_cache' in st.session_state:
        image_data.seek(0)
        cached_data = image_data.read()
        st.session_state.original_images_cache['problem_slide'] = cached_data
        image_data.seek(0)
    
    if image_data:
        # Position image on right side
        pic = slide.shapes.add_picture(
            image_data,
            Inches(8.3),   # Positioned for proper alignment
            Inches(2.0),   # Centered vertically
            Inches(4.5),   # Consistent width
            Inches(4.0)    # Consistent height
        )
    else:
        # Add fallback icon with better positioning
        icon_left = Inches(9.5)
        icon_top = Inches(3.0)
        icon_box = slide.shapes.add_textbox(icon_left, icon_top, Inches(2.0), Inches(2.0))
        icon_tf = icon_box.text_frame
        icon_tf.auto_size = False
        icon_p = icon_tf.paragraphs[0]
        icon_info = get_slide_icon("problem")
        icon_p.text = icon_info["icon"]
        icon_p.alignment = PP_ALIGN.CENTER
        apply_text_formatting(icon_tf, size=64, alignment=PP_ALIGN.CENTER)

def create_solution_slide(prs: Presentation, content: Dict[str, Any], presentation_context: Dict[str, Any] = None):
    """Create the solution overview slide."""
    slide_layout = prs.slide_layouts[6]  # Blank layout
    slide = prs.slides.add_slide(slide_layout)
    
    product_name = presentation_context.get("metadata", {}).get("product_name", "")
    
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
    
    # Replace any placeholder text
    clean_title = content["title"].replace("[Product Name]", product_name)
    title_p.text = clean_title
    title_p.alignment = PP_ALIGN.LEFT
    
    # Apply title formatting
    for run in title_p.runs:
        run.font.name = FONTS["title"]
        run.font.size = Pt(32)
        run.font.bold = True
        run.font.color.rgb = RGBColor.from_string(COLORS["black"])
    
    # Create custom text box for paragraph content
    left = Inches(0.5)
    top = Inches(1.7)
    width = Inches(7.5)
    height = Inches(5.0)
    
    text_box = slide.shapes.add_textbox(left, top, width, height)
    tf = text_box.text_frame
    tf.word_wrap = True
    tf.auto_size = False
    tf.margin_left = Inches(0.1)
    tf.margin_right = Inches(0.1)

    p = tf.paragraphs[0]
    
    # Handle different field names for content
    paragraph_text = ""
    if "paragraph" in content:
        paragraph_text = content["paragraph"]
    elif "description" in content:
        paragraph_text = content["description"]
    elif "value_proposition" in content:
        paragraph_text = content["value_proposition"]
    
    # Replace product name placeholders
    clean_paragraph = paragraph_text.replace("[Product Name]", product_name) if product_name else paragraph_text
    
    # Truncate paragraph based on available space
    max_chars = 400
    if len(clean_paragraph) > max_chars:
        clean_paragraph = clean_paragraph[:max_chars-3] + "..."

    p.text = clean_paragraph
    
    # Format paragraph with proper spacing
    p.space_before = Pt(6)
    p.space_after = Pt(6)
    p.line_spacing = 1.2
    p.font.name = FONTS["body"]
    p.font.size = Pt(20)  # Larger for better readability
    
    # Add image with consistent positioning
    image_data = fetch_image_for_slide("solution", presentation_context, use_placeholders=False)
    
    # Cache the image
    import streamlit as st
    if image_data and 'original_images_cache' in st.session_state:
        image_data.seek(0)
        cached_data = image_data.read()
        st.session_state.original_images_cache['solution_slide'] = cached_data
        image_data.seek(0)
    
    if image_data:
        # Position image on right side with consistent sizing
        pic = slide.shapes.add_picture(
            image_data,
            Inches(8.3),
            Inches(2.0),
            Inches(4.5),
            Inches(4.0)
        )
    else:
        # Add fallback icon
        icon_left = Inches(9.5)
        icon_top = Inches(3.0)
        icon_box = slide.shapes.add_textbox(icon_left, icon_top, Inches(2.0), Inches(2.0))
        icon_tf = icon_box.text_frame
        icon_tf.auto_size = False
        icon_p = icon_tf.paragraphs[0]
        icon_p.text = get_slide_icon("solution")["icon"]
        icon_p.alignment = PP_ALIGN.CENTER
        apply_text_formatting(icon_tf, size=64, alignment=PP_ALIGN.CENTER)

def create_features_slide(prs: Presentation, content: Dict[str, Any]):
    """Create the key features slide."""
    slide_layout = prs.slide_layouts[6]  # Blank layout for consistency
    slide = prs.slides.add_slide(slide_layout)
    
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
    title_p.text = content["title"].replace("[Product Name]", "")
    title_p.alignment = PP_ALIGN.LEFT
    
    # Apply title formatting
    for run in title_p.runs:
        run.font.name = FONTS["title"]
        run.font.size = Pt(32)
        run.font.bold = True
        run.font.color.rgb = RGBColor.from_string(COLORS["black"])
    
    # Get features list
    features = content["features"]
    max_features = 5  # Limit to 5 features for better spacing
    features = features[:max_features]
    
    # Create a more visually appealing list
    left_margin = Inches(1.0)
    top_start = Inches(1.7)
    feature_height = Inches(0.9)  # More space between features
    icon_width = Inches(0.8)
    text_width = Inches(11.0)
    
    for i, feature in enumerate(features):
        # Handle different feature formats
        if isinstance(feature, dict):
            feature_text = feature.get('feature', feature.get('name', feature.get('title', str(feature))))
            icon_text = feature_text
        else:
            feature_text = str(feature)
            icon_text = feature_text
        
        # Truncate feature text if needed
        max_chars = 100
        if len(feature_text) > max_chars:
            feature_text = feature_text[:max_chars-3] + "..."

        # Calculate vertical position
        top_position = top_start + (i * feature_height)
        
        # Add icon
        icon_box = slide.shapes.add_textbox(left_margin, top_position, icon_width, feature_height)
        icon_tf = icon_box.text_frame
        icon_tf.auto_size = False
        icon_p = icon_tf.paragraphs[0]
        icon_p.text = match_icon_to_feature(icon_text)
        icon_p.alignment = PP_ALIGN.CENTER
        
        # Apply icon formatting
        for run in icon_p.runs:
            run.font.name = FONTS["body"]
            run.font.size = Pt(24)
            run.font.bold = True
        
        # Add feature text
        text_left = left_margin + icon_width + Inches(0.3)
        text_box = slide.shapes.add_textbox(text_left, top_position, text_width, feature_height)
        text_tf = text_box.text_frame
        text_tf.word_wrap = True
        text_tf.auto_size = False
        text_p = text_tf.paragraphs[0]
        text_p.text = feature_text
        text_p.alignment = PP_ALIGN.LEFT
        
        # Apply text formatting
        for run in text_p.runs:
            run.font.name = FONTS["body"]
            run.font.size = Pt(20)
            run.font.color.rgb = RGBColor.from_string(COLORS["black"])

def create_advantage_slide(prs: Presentation, content: Dict[str, Any], presentation_context: Dict[str, Any] = None):
    """Create the competitive advantage slide."""
    slide_layout = prs.slide_layouts[6]  # Blank layout
    slide = prs.slides.add_slide(slide_layout)
    
    product_name = presentation_context.get("metadata", {}).get("product_name", "")
    
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
    title_p.text = content["title"].replace("[Product Name]", product_name)
    title_p.alignment = PP_ALIGN.LEFT
    
    # Apply title formatting
    for run in title_p.runs:
        run.font.name = FONTS["title"]
        run.font.size = Pt(32)
        run.font.bold = True
        run.font.color.rgb = RGBColor.from_string(COLORS["black"])
    
    # Create text box for bullets with consistent positioning
    left = Inches(0.5)
    top = Inches(1.7)
    width = Inches(7.5)
    height = Inches(5.0)
    
    text_box = slide.shapes.add_textbox(left, top, width, height)
    tf = text_box.text_frame
    tf.word_wrap = True
    tf.auto_size = False
    tf.margin_left = Inches(0.1)
    tf.margin_right = Inches(0.1)
    
    # Clear any existing text
    if tf.paragraphs:
        p = tf.paragraphs[0]
        p.text = ""
    else:
        p = tf.add_paragraph()
        
    # Handle different field names for bulleted content
    bullets = []
    if "differentiators" in content:
        bullets = [item["point"] if isinstance(item, dict) else str(item) for item in content["differentiators"]]
    elif "bullets" in content:
        bullets = content["bullets"]
    elif "pain_points" in content:
        bullets = content["pain_points"]
    
    # Limit bullets for better readability
    max_bullets = 4
    bullets = bullets[:max_bullets]
    
    for i, bullet in enumerate(bullets):
        # Replace any product name placeholders
        clean_bullet = bullet.replace("[Product Name]", product_name) if product_name else bullet
        
        # Truncate if needed
        max_chars = 85
        bullet_text = clean_bullet
        if len(clean_bullet) > max_chars:
            bullet_text = clean_bullet[:max_chars-3] + "..."

        if i == 0:
            p.text = f"• {bullet_text}"
        else:
            p = tf.add_paragraph()
            p.text = f"• {bullet_text}"
        
        # Apply bullet formatting
        p.level = 0
        p.alignment = PP_ALIGN.LEFT
        p.font.name = FONTS["body"]
        p.font.size = Pt(20)  # Larger for better readability
        p.space_before = Pt(10)
        p.space_after = Pt(10)
        p.line_spacing = 1.2
    
    # Add image with consistent positioning
    image_data = fetch_image_for_slide("advantage", presentation_context, use_placeholders=False)
    
    # Cache the image
    import streamlit as st
    if image_data and 'original_images_cache' in st.session_state:
        image_data.seek(0)
        cached_data = image_data.read()
        st.session_state.original_images_cache['advantage_slide'] = cached_data
        image_data.seek(0)
    
    if image_data:
        # Position image with consistent dimensions
        pic = slide.shapes.add_picture(
            image_data,
            Inches(8.3),
            Inches(2.0),
            Inches(4.5),
            Inches(4.0)
        )
    else:
        # Add fallback icon
        icon_left = Inches(9.5)
        icon_top = Inches(3.0)
        icon_box = slide.shapes.add_textbox(icon_left, icon_top, Inches(2.0), Inches(2.0))
        icon_tf = icon_box.text_frame
        icon_tf.auto_size = False
        icon_p = icon_tf.paragraphs[0]
        icon_p.text = get_slide_icon("advantage")["icon"]
        icon_p.alignment = PP_ALIGN.CENTER
        apply_text_formatting(icon_tf, size=64, alignment=PP_ALIGN.CENTER)

def create_audience_slide(prs: Presentation, content: Dict[str, Any], presentation_context: Dict[str, Any] = None):
    """Create the target audience slide."""
    slide_layout = prs.slide_layouts[6]  # Blank layout
    slide = prs.slides.add_slide(slide_layout)
    
    product_name = presentation_context.get("metadata", {}).get("product_name", "")
    
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
    title_p.text = content["title"].replace("[Product Name]", product_name)
    title_p.alignment = PP_ALIGN.LEFT
    
    # Apply title formatting
    for run in title_p.runs:
        run.font.name = FONTS["title"]
        run.font.size = Pt(32)
        run.font.bold = True
        run.font.color.rgb = RGBColor.from_string(COLORS["black"])
    
    # Create text box for content with consistent positioning
    left = Inches(0.5)
    top = Inches(1.7)
    width = Inches(7.5)
    height = Inches(5.0)
    
    text_box = slide.shapes.add_textbox(left, top, width, height)
    tf = text_box.text_frame
    tf.word_wrap = True
    tf.auto_size = False
    tf.margin_left = Inches(0.1)
    tf.margin_right = Inches(0.1)
    
    p = tf.paragraphs[0]
    
    # Handle different field names for text content
    paragraph_text = ""
    if "paragraph" in content:
        paragraph_text = content["paragraph"]
    elif "description" in content:
        paragraph_text = content["description"]
    elif "content" in content:
        paragraph_text = content["content"]
    
    # Replace product name placeholders
    clean_paragraph = paragraph_text.replace("[Product Name]", product_name) if product_name else paragraph_text
    
    # Truncate if needed
    max_chars = 400
    if len(clean_paragraph) > max_chars:
        clean_paragraph = clean_paragraph[:max_chars-3] + "..."

    p.text = clean_paragraph
    
    # Format paragraph with proper spacing
    p.space_before = Pt(6)
    p.space_after = Pt(6)
    p.line_spacing = 1.2
    p.font.name = FONTS["body"]
    p.font.size = Pt(20)  # Larger for better readability
    
    # Add image with consistent positioning
    image_data = fetch_image_for_slide("audience", presentation_context, use_placeholders=False)
    
    # Cache the image
    import streamlit as st
    if image_data and 'original_images_cache' in st.session_state:
        image_data.seek(0)
        cached_data = image_data.read()
        st.session_state.original_images_cache['audience_slide'] = cached_data
        image_data.seek(0)
    
    if image_data:
        # Position image with consistent dimensions
        pic = slide.shapes.add_picture(
            image_data,
            Inches(8.3),
            Inches(2.0),
            Inches(4.5),
            Inches(4.0)
        )
    else:
        # Add fallback icon
        icon_left = Inches(9.5)
        icon_top = Inches(3.0)
        icon_box = slide.shapes.add_textbox(icon_left, icon_top, Inches(2.0), Inches(2.0))
        icon_tf = icon_box.text_frame
        icon_tf.auto_size = False
        icon_p = icon_tf.paragraphs[0]
        icon_p.text = get_slide_icon("audience")["icon"]
        icon_p.alignment = PP_ALIGN.CENTER
        apply_text_formatting(icon_tf, size=64, alignment=PP_ALIGN.CENTER)

def create_cta_slide(prs: Presentation, content: Dict[str, Any]):
    """Create the call to action slide."""
    slide_layout = prs.slide_layouts[6]  # Blank layout
    slide = prs.slides.add_slide(slide_layout)
    
    # Main centered title
    title_left = Inches(0.5)
    title_top = Inches(1.5)  # Higher placement
    title_width = Inches(12.33)
    title_height = Inches(1.5)
    
    title_box = slide.shapes.add_textbox(title_left, title_top, title_width, title_height)
    title_tf = title_box.text_frame
    title_tf.word_wrap = True
    title_tf.auto_size = False
    title_p = title_tf.paragraphs[0]
    title_p.text = content.get("title", "Get Started Today")
    title_p.alignment = PP_ALIGN.CENTER
    
    # Apply title formatting - larger and bolder
    for run in title_p.runs:
        run.font.name = FONTS["title"]
        run.font.size = Pt(36)  # Larger for emphasis
        run.font.bold = True
        run.font.color.rgb = RGBColor.from_string(COLORS["black"])
    
    # CTA text - centered and prominent
    cta_left = Inches(1.0)
    cta_top = Inches(3.0)  # Positioned below title
    cta_width = Inches(11.33)
    cta_height = Inches(2.0)  # Taller for multi-line CTAs
    
    cta_box = slide.shapes.add_textbox(cta_left, cta_top, cta_width, cta_height)
    cta_tf = cta_box.text_frame
    cta_tf.word_wrap = True
    cta_tf.auto_size = False
    
    cta_p = cta_tf.paragraphs[0]
    cta_text = content.get("cta_text", content.get("call_to_action", "Contact us today to learn more!"))
    
    # Don't truncate CTAs - they're often short and important
    cta_p.text = cta_text
    cta_p.alignment = PP_ALIGN.CENTER
    
    # Apply CTA formatting - prominent and attention-grabbing
    for run in cta_p.runs:
        run.font.name = FONTS["title"]
        run.font.size = Pt(28)
        run.font.bold = True
        run.font.color.rgb = RGBColor.from_string(COLORS["black"])
    
    # Add any supporting bullets/text if available
    if "bullets" in content and content["bullets"]:
        bullet_top = Inches(5.0)  # Position below CTA
        bullet_height = Inches(1.5)
        
        bullet_box = slide.shapes.add_textbox(cta_left, bullet_top, cta_width, bullet_height)
        bullet_tf = bullet_box.text_frame
        bullet_tf.word_wrap = True
        
        # Add up to 2 bullets only
        max_bullets = min(2, len(content["bullets"]))
        for i in range(max_bullets):
            if i == 0:
                bullet_p = bullet_tf.paragraphs[0]
            else:
                bullet_p = bullet_tf.add_paragraph()
                
            bullet_p.text = f"• {content['bullets'][i]}"
            bullet_p.alignment = PP_ALIGN.CENTER
            
            # Format bullet text
            for run in bullet_p.runs:
                run.font.name = FONTS["body"]
                run.font.size = Pt(20)
                run.font.color.rgb = RGBColor.from_string(COLORS["black"])
    
    # Optional footer bar
    footer_left = Inches(0)
    footer_top = Inches(6.5)
    footer_width = Inches(13.33)
    footer_height = Inches(1.0)
    
    footer_box = slide.shapes.add_textbox(footer_left, footer_top, footer_width, footer_height)
    footer_box.fill.solid()
    footer_box.fill.fore_color.rgb = RGBColor.from_string(COLORS["gray"])
    footer_box.line.fill.background()