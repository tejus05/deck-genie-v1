import requests
import io
import json
import google.generativeai as genai
from PIL import Image, ImageDraw, ImageFont
from typing import Optional, Dict, Any, Tuple, List
from utils import get_api_key, IMAGE_KEYWORDS
from slide_content_generators import MODEL_NAME
import random
import os
import streamlit as st

def fetch_image_for_slide(slide_type: str, context: Dict = None, use_placeholders: bool = False, search_terms: List[str] = None):
    """Fetch an appropriate image for a slide type."""
    # If using placeholders, return those directly
    if use_placeholders:
        return get_placeholder_image(slide_type)
        
    # Use search terms if provided, otherwise use default search terms for the slide type
    query = None
    if search_terms and len(search_terms) > 0:
        query = " ".join(search_terms[:3])  # Use top 3 search terms
    else:
        # Default queries based on slide type
        if slide_type == "problem":
            query = "business problem challenge"
        elif slide_type == "solution":
            query = "business solution innovation"
        elif slide_type == "advantage":
            query = "competitive advantage business"
        elif slide_type == "audience":
            query = "target audience business" 
        elif slide_type == "market":
            query = "market growth business chart"
        elif slide_type == "roadmap":
            query = "product roadmap timeline"
        elif slide_type == "team":
            query = "business team professional"
        else:
            query = "business professional"
    
    # Try to fetch from Unsplash
    try:
        image_data = fetch_image_from_unsplash(query)
        if image_data:
            return image_data
    except Exception as e:
        print(f"Error in fetch_image_from_unsplash: {e}")
    
    # If we get here, Unsplash failed, return a placeholder
    print(f"Using placeholder image for {slide_type} slide")
    return get_placeholder_image(slide_type)

def generate_image_query_with_gemini(slide_type: str, context: Dict[str, Any]) -> str:
    """
    Use Gemini API to generate a relevant image search query based on slide content.
    
    Args:
        slide_type: Type of slide
        context: Context containing slide content and presentation information
        
    Returns:
        A search query string optimized for image relevance
    """
    try:
        # Configure the Gemini API
        genai.configure(api_key=get_api_key("GEMINI_API_KEY"))
        
        # Create a prompt based on slide type and content
        if slide_type == "problem":
            prompt = f"""
            Generate a single, specific Unsplash search query (3-5 words) for a business presentation slide about this problem:
            
            Title: {context[slide_type].get('title', 'Problem Statement')}
            
            Problem points:
            {' '.join(context[slide_type].get('bullets', ['Business challenges']))}
            
            IMPORTANT: Keep your query focused specifically on B2B SaaS, enterprise technology, or business contexts.
            Avoid generic imagery and focus on professional/corporate tech environments.
            The image should be visually relevant to this content while staying in the B2B tech domain.
            
            Format your response as a JSON object with a single key "query" and its string value.
            Example: {{"query": "enterprise data challenges"}}
            """
        
        elif slide_type == "solution":
            prompt = f"""
            Generate a single, specific Unsplash search query (3-5 words) for a business presentation slide about this solution:
            
            Title: {context[slide_type].get('title', 'Our Solution')}
            
            Solution: 
            {context[slide_type].get('paragraph', 'Business solution')}
            
            Product name: {context.get('metadata', {}).get('product_name', '')}
            
            IMPORTANT: Keep your query focused specifically on B2B SaaS, enterprise technology, or business contexts.
            Avoid generic imagery and focus on professional/corporate tech environments.
            The image should be visually relevant to this content while staying in the B2B tech domain.
            
            Format your response as a JSON object with a single key "query" and its string value.
            Example: {{"query": "enterprise software solution"}}
            """
            
        elif slide_type == "advantage":
            prompt = f"""
            Generate a single, specific Unsplash search query (3-5 words) for a business presentation slide about competitive advantages:
            
            Title: {context[slide_type].get('title', 'Our Advantage')}
            
            Advantages:
            {' '.join(context[slide_type].get('bullets', ['Business advantage']))}
            
            IMPORTANT: Keep your query focused specifically on B2B SaaS, enterprise technology, or business contexts.
            Avoid generic imagery and focus on professional/corporate tech environments.
            The image should be visually relevant to this content while staying in the B2B tech domain.
            
            Format your response as a JSON object with a single key "query" and its string value.
            Example: {{"query": "business competitive advantage technology"}}
            """
            
        elif slide_type == "audience":
            prompt = f"""
            Generate a single, specific Unsplash search query (3-5 words) for a business presentation slide about this target audience:
            
            Title: {context[slide_type].get('title', 'Our Audience')}
            
            Target audience: 
            {context[slide_type].get('paragraph', 'Business professionals')}
            
            IMPORTANT: Keep your query focused specifically on B2B SaaS, enterprise technology, or business contexts.
            Avoid generic imagery and focus on professional/corporate audiences.
            The image should be visually relevant to this specific B2B audience, not consumer or general public.
            
            Format your response as a JSON object with a single key "query" and its string value.
            Example: {{"query": "enterprise IT professionals meeting"}}
            """
        
        else:
            # For unknown slide types, create a generic business query
            company_name = context.get('metadata', {}).get('company_name', '')
            product_name = context.get('metadata', {}).get('product_name', '')
            prompt = f"""
            Generate a single, specific Unsplash search query (3-5 words) for a business presentation slide.
            
            Company: {company_name}
            Product: {product_name}
            
            IMPORTANT: Keep your query focused specifically on B2B SaaS, enterprise technology, or business contexts.
            Avoid generic imagery and focus on professional/corporate tech environments.
            
            Format your response as a JSON object with a single key "query" and its string value.
            Example: {{"query": "enterprise technology solution"}}
            """
        
        # Call Gemini API
        model = genai.GenerativeModel(
            model_name=MODEL_NAME,
            generation_config={"temperature": 0.2, "max_output_tokens": 100}
        )
        
        # Include system instruction in the prompt instead of as a parameter
        prompt = "You create specific, accurate search queries for B2B SaaS presentation images, focusing on enterprise technology and business contexts.\n\n" + prompt
        
        response = model.generate_content(
            [
                {"role": "user", "parts": [prompt]},
            ]
        )
        
        # Parse the JSON response
        try:
            response_text = response.text
            content_start = response_text.find('{')
            content_end = response_text.rfind('}') + 1
            
            if content_start >= 0 and content_end > content_start:
                json_content = response_text[content_start:content_end]
                result = json.loads(json_content)
                
                # Return the query
                if "query" in result:
                    return result["query"]
                else:
                    raise ValueError("Response missing 'query' key")
            else:
                raise ValueError("No valid JSON found in response")
                
        except Exception as e:
            # Fallback if parsing fails
            print(f"Error parsing Gemini response: {str(e)}")
            fallback_queries = {
                "problem": "business challenge",
                "solution": "business solution",
                "advantage": "business advantage",
                "audience": "business meeting"
            }
            return fallback_queries.get(slide_type, "business")
            
    except Exception as e:
        print(f"Error generating image query with Gemini: {str(e)}")
        # Fallback to predefined keywords
        keywords = IMAGE_KEYWORDS.get(slide_type, ["business"])
        return random.choice(keywords)

def get_slide_icon(slide_type: str) -> Dict[str, Any]:
    """
    Get a default icon for a slide type.
    This is used as a fallback when images can't be fetched.
    
    Args:
        slide_type: Type of slide
        
    Returns:
        Dictionary with icon information
    """
    from utils import SLIDE_ICONS
    
    icon = SLIDE_ICONS.get(slide_type, "📊")
    
    return {
        "icon": icon,
        "size": (32, 32)
    }

def create_placeholder_image(slide_type: str) -> io.BytesIO:
    """
    Create a placeholder image when Unsplash fetch fails.
    
    Args:
        slide_type: Type of slide to create placeholder for
        
    Returns:
        BytesIO object containing a simple placeholder image
    """
    try:
        # Create a simple colored rectangle
        width, height = 800, 600
        colors = {
            "problem": "#FF6B6B",  # Red
            "solution": "#4ECDC4", # Teal  
            "advantage": "#45B7D1", # Blue
            "audience": "#96CEB4",  # Green
            "features": "#FFEAA7",  # Yellow
            "call_to_action": "#DDA0DD"  # Plum
        }
        
        color = colors.get(slide_type, "#95A5A6")  # Default gray
        
        # Create image
        image = Image.new('RGB', (width, height), color)
        draw = ImageDraw.Draw(image)
        
        # Add text
        text = f"{slide_type.upper()}\nIMAGE"
        
        try:
            # Try to use a default font
            font = ImageFont.truetype("Arial.ttf", 48)
        except:
            # Fallback to default font
            font = ImageFont.load_default()
        
        # Get text size and center it
        bbox = draw.textbbox((0, 0), text, font=font)
        text_width = bbox[2] - bbox[0]
        text_height = bbox[3] - bbox[1]
        
        x = (width - text_width) // 2
        y = (height - text_height) // 2
        
        # Draw text with outline for better visibility
        outline_color = "#FFFFFF" if slide_type != "features" else "#000000"
        for adj in range(-2, 3):
            for adj2 in range(-2, 3):
                draw.text((x+adj, y+adj2), text, font=font, fill=outline_color)
        
        # Draw main text
        main_color = "#000000" if slide_type == "features" else "#FFFFFF"
        draw.text((x, y), text, font=font, fill=main_color)
        
        # Save to BytesIO
        img_buffer = io.BytesIO()
        image.save(img_buffer, format='PNG')
        img_buffer.seek(0)
        
        return img_buffer
        
    except Exception as e:
        print(f"Error creating placeholder image: {str(e)}")
        # Return a minimal BytesIO with empty content as last resort
        return io.BytesIO()

def get_placeholder_image(slide_type: str):
    """
    Return a placeholder image for a given slide type when no image can be fetched.
    
    Args:
        slide_type: The type of slide that needs a placeholder image
        
    Returns:
        BytesIO object containing a basic placeholder image
    """
    # Create a basic image with slide type as text
    width, height = 800, 600
    background_colors = {
        "problem": (230, 230, 250),  # Lavender
        "solution": (240, 255, 240),  # Honeydew
        "advantage": (255, 240, 245),  # Lavender blush
        "audience": (240, 248, 255),  # Alice blue
        "market": (255, 250, 240),    # Floral white
        "roadmap": (245, 255, 250),   # Mint cream
        "team": (255, 245, 238),      # Seashell
        "features": (240, 255, 255),  # Azure
        "cta": (255, 255, 240)        # Ivory
    }
    
    # Default background color
    bg_color = background_colors.get(slide_type.lower(), (245, 245, 245))
    
    # Create image
    img = Image.new('RGB', (width, height), color=bg_color)
    draw = ImageDraw.Draw(img)
    
    # Try to use a system font
    try:
        # Try common system fonts
        font_options = ['Arial', 'Verdana', 'Tahoma', 'Calibri', 'Georgia']
        font = None
        for font_name in font_options:
            try:
                font = ImageFont.truetype(font_name, 40)
                break
            except IOError:
                continue
                
        if font is None:
            # Fallback to default font
            font = ImageFont.load_default()
    except Exception:
        font = ImageFont.load_default()
    
    # Draw placeholder text and design elements
    title = f"{slide_type.capitalize()} Slide"
    
    # Add a border
    border_margin = 20
    draw.rectangle([border_margin, border_margin, width-border_margin, height-border_margin], 
                  outline=(100, 100, 100), width=2)
    
    # Add centered text
    text_width, text_height = getattr(draw, 'textsize', lambda text, font: (200, 40))(title, font=font)
    position = ((width - text_width) // 2, (height - text_height) // 2)
    draw.text(position, title, fill=(80, 80, 80), font=font)
    
    # Add some design elements based on slide type
    for i in range(10):
        x = random.randint(50, width - 50)
        y = random.randint(50, height - 50)
        size = random.randint(10, 40)
        opacity = random.randint(30, 100)
        draw.ellipse([x, y, x+size, y+size], 
                    fill=(opacity, opacity, opacity, opacity))
    
    # Convert to BytesIO
    img_bytes = io.BytesIO()
    img.save(img_bytes, format='PNG')
    img_bytes.seek(0)
    
    return img_bytes

def fetch_image_from_unsplash(query: str) -> Optional[io.BytesIO]:
    """
    Fetch an image from Unsplash API based on a query.
    
    Args:
        query: Search term for the image
        
    Returns:
        BytesIO object containing the image or None if fetch fails
    """
    try:
        # Get the API key from environment
        api_key = os.environ.get("UNSPLASH_API_KEY")
        if not api_key:
            api_key = st.secrets.get("UNSPLASH_API_KEY")
        
        if not api_key:
            print("No Unsplash API key found. Using placeholder image.")
            return None
        
        # Prepare the request
        headers = {
            "Authorization": f"Client-ID {api_key}"
        }
        
        # Clean up the query to make it more likely to succeed
        safe_query = query.replace("[Product Name]", "product").strip()
        if not safe_query or len(safe_query) < 3:
            safe_query = "business professional"
            
        params = {
            "query": safe_query,
            "orientation": "landscape",
            "per_page": 5  # Fetch multiple options
        }
        
        # Make the API call with a timeout
        response = requests.get(
            "https://api.unsplash.com/search/photos",
            headers=headers,
            params=params,
            timeout=5  # 5 second timeout
        )
        
        # Check if the request was successful
        if response.status_code == 200:
            data = response.json()
            results = data.get("results", [])
            
            # If we have results, pick a random one
            if results:
                chosen_image = random.choice(results)
                image_url = chosen_image["urls"]["regular"]
                
                # Fetch the actual image
                image_response = requests.get(image_url, timeout=5)
                if image_response.status_code == 200:
                    # Return as BytesIO
                    image_data = io.BytesIO(image_response.content)
                    return image_data
        
        print(f"Unsplash API error: {response.status_code}, {response.text}")
        return None
        
    except Exception as e:
        print(f"Error fetching image from Unsplash: {e}")
        return None