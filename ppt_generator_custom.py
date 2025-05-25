import io
from typing import Dict, Any, List, Optional
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from image_fetcher import fetch_image_for_slide, get_slide_icon
from utils import FONTS, COLORS, FONT_SIZES, MARGINS, CONTENT_AREA, IMAGE_AREA, match_icon_to_feature
from image_manager import ImageManager

def create_custom_presentation(content: Dict[str, Any], filename: str, slide_order: List[str], image_manager: ImageManager) -> io.BytesIO:
    """
    Create a customized PowerPoint presentation from modified content.
    
    Args:
        content: Dictionary containing modified content for slides
        filename: Name to give the file
        slide_order: Custom order of slides
        image_manager: Manager for custom images
        
    Returns:
        BytesIO object containing the presentation
    """
    prs = Presentation()
    
    # Set slide size to standard (4:3) - 10 x 7.5 inches
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(7.5)
    
    # Create slides in custom order
    slide_creators = {
        'title_slide': lambda: create_custom_title_slide(prs, content['title_slide'], image_manager),
        'problem_slide': lambda: create_custom_problem_slide(prs, content['problem_slide'], content, image_manager),
        'solution_slide': lambda: create_custom_solution_slide(prs, content['solution_slide'], content, image_manager),
        'features_slide': lambda: create_custom_features_slide(prs, content['features_slide'], image_manager),
        'advantage_slide': lambda: create_custom_advantage_slide(prs, content['advantage_slide'], content, image_manager),
        'audience_slide': lambda: create_custom_audience_slide(prs, content['audience_slide'], content, image_manager),
        'cta_slide': lambda: create_custom_cta_slide(prs, content['cta_slide'], image_manager)
    }
    
    # Create slides in the specified order
    for slide_key in slide_order:
        if slide_key in content and slide_key in slide_creators:
            slide_creators[slide_key]()
    
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

def create_custom_title_slide(prs: Presentation, content: Dict[str, str], image_manager: ImageManager):
    """Create the customized title slide."""
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
    
    # Add custom image if available
    custom_image = image_manager.get_image_for_slide('title_slide')
    if custom_image:
        # Add custom image as background or decorative element
        # Position it in the lower right corner
        pic = slide.shapes.add_picture(
            custom_image,
            Inches(7),
            Inches(5),
            Inches(2.5),
            Inches(2)
        )

def create_custom_problem_slide(prs: Presentation, content: Dict[str, Any], presentation_context: Dict[str, Any], image_manager: ImageManager):
    """Create the customized problem statement slide."""
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
    
    # Right content - custom image or default
    right_content = slide.placeholders[2]
    custom_image = image_manager.get_image_for_slide('problem_slide')
    
    if custom_image:
        placeholder_width = right_content.width
        placeholder_height = right_content.height
        placeholder_left = right_content.left
        placeholder_top = right_content.top

        pic = slide.shapes.add_picture(
            custom_image,
            placeholder_left,
            placeholder_top,
            placeholder_width,
            placeholder_height
        )
    else:
        # Use default image or icon
        image_data = fetch_image_for_slide("problem", presentation_context, use_placeholders=False)
        
        if image_data:
            placeholder_width = right_content.width
            placeholder_height = right_content.height
            placeholder_left = right_content.left
            placeholder_top = right_content.top

            pic = slide.shapes.add_picture(
                image_data,
                placeholder_left,
                placeholder_top,
                placeholder_width,
                placeholder_height
            )
        else:
            icon_info = get_slide_icon("problem")
            tf = right_content.text_frame
            p = tf.add_paragraph()
            p.text = icon_info["icon"]
            p.alignment = PP_ALIGN.CENTER
            apply_text_formatting(tf, size=72, alignment=PP_ALIGN.CENTER)

def create_custom_solution_slide(prs: Presentation, content: Dict[str, Any], presentation_context: Dict[str, Any], image_manager: ImageManager):
    """Create the customized solution overview slide."""
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
    
    # Right content - custom image or default
    right_content = slide.placeholders[2]
    custom_image = image_manager.get_image_for_slide('solution_slide')
    
    if custom_image:
        placeholder_width = right_content.width
        placeholder_height = right_content.height
        placeholder_left = right_content.left
        placeholder_top = right_content.top

        pic = slide.shapes.add_picture(
            custom_image,
            placeholder_left,
            placeholder_top,
            placeholder_width,
            placeholder_height
        )
    else:
        # Use default image or icon
        image_data = fetch_image_for_slide("solution", presentation_context, use_placeholders=False)
        
        if image_data:
            placeholder_width = right_content.width
            placeholder_height = right_content.height
            placeholder_left = right_content.left
            placeholder_top = right_content.top

            pic = slide.shapes.add_picture(
                image_data,
                placeholder_left,
                placeholder_top,
                placeholder_width,
                placeholder_height
            )
        else:
            tf = right_content.text_frame
            p = tf.add_paragraph()
            p.text = get_slide_icon("solution")["icon"]
            p.alignment = PP_ALIGN.CENTER
            apply_text_formatting(tf, size=72, alignment=PP_ALIGN.CENTER)

def create_custom_features_slide(prs: Presentation, content: Dict[str, Any], image_manager: ImageManager):
    """Create the customized key features slide."""
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
    
    # Style table with no fill
    for cell in table.iter_cells():
        cell.fill.background()
    
    # Add custom image if available (smaller, positioned below table)
    custom_image = image_manager.get_image_for_slide('features_slide')
    if custom_image:
        pic = slide.shapes.add_picture(
            custom_image,
            Inches(7),
            Inches(5.5),
            Inches(2.5),
            Inches(1.5)
        )

def create_custom_advantage_slide(prs: Presentation, content: Dict[str, Any], presentation_context: Dict[str, Any], image_manager: ImageManager):
    """Create the customized competitive advantage slide."""
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
    
    # Right content - custom image or default
    right_content = slide.placeholders[2]
    custom_image = image_manager.get_image_for_slide('advantage_slide')
    
    if custom_image:
        placeholder_width = right_content.width
        placeholder_height = right_content.height
        placeholder_left = right_content.left
        placeholder_top = right_content.top

        pic = slide.shapes.add_picture(
            custom_image,
            placeholder_left,
            placeholder_top,
            placeholder_width,
            placeholder_height
        )
    else:
        # Use default image or icon
        image_data = fetch_image_for_slide("advantage", presentation_context, use_placeholders=False)
        
        if image_data:
            placeholder_width = right_content.width
            placeholder_height = right_content.height
            placeholder_left = right_content.left
            placeholder_top = right_content.top

            pic = slide.shapes.add_picture(
                image_data,
                placeholder_left,
                placeholder_top,
                placeholder_width,
                placeholder_height
            )
        else:
            tf = right_content.text_frame
            p = tf.add_paragraph()
            p.text = get_slide_icon("advantage")["icon"]
            p.alignment = PP_ALIGN.CENTER
            apply_text_formatting(tf, size=72, alignment=PP_ALIGN.CENTER)

def create_custom_audience_slide(prs: Presentation, content: Dict[str, Any], presentation_context: Dict[str, Any], image_manager: ImageManager):
    """Create the customized target audience slide."""
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
    
    # Right content - custom image or default
    right_content = slide.placeholders[2]
    custom_image = image_manager.get_image_for_slide('audience_slide')
    
    if custom_image:
        placeholder_width = right_content.width
        placeholder_height = right_content.height
        placeholder_left = right_content.left
        placeholder_top = right_content.top

        pic = slide.shapes.add_picture(
            custom_image,
            placeholder_left,
            placeholder_top,
            placeholder_width,
            placeholder_height
        )
    else:
        # Use default image or icon
        image_data = fetch_image_for_slide("audience", presentation_context, use_placeholders=False)
        
        if image_data:
            placeholder_width = right_content.width
            placeholder_height = right_content.height
            placeholder_left = right_content.left
            placeholder_top = right_content.top

            pic = slide.shapes.add_picture(
                image_data,
                placeholder_left,
                placeholder_top,
                placeholder_width,
                placeholder_height
            )
        else:
            tf = right_content.text_frame
            p = tf.add_paragraph()
            p.text = get_slide_icon("audience")["icon"]
            p.alignment = PP_ALIGN.CENTER
            apply_text_formatting(tf, size=72, alignment=PP_ALIGN.CENTER)

def create_custom_cta_slide(prs: Presentation, content: Dict[str, Any], image_manager: ImageManager):
    """Create the customized call to action slide."""
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
    
    # Add custom image if available (positioned at bottom)
    custom_image = image_manager.get_image_for_slide('cta_slide')
    if custom_image:
        pic = slide.shapes.add_picture(
            custom_image,
            Inches(3.5),
            Inches(4.5),
            Inches(3),
            Inches(2)
        )
    else:
        # Optional grey stripe at bottom
        left = Inches(0)
        top = Inches(6.5)
        width = Inches(10)
        height = Inches(0.5)
        
        stripe_box = slide.shapes.add_textbox(left, top, width, height)
        stripe_box.fill.solid()
        stripe_box.fill.fore_color.rgb = RGBColor.from_string(COLORS["gray"])
        stripe_box.line.fill.background()
