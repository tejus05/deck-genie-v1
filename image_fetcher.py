import requests
import io
import json
import google.generativeai as genai
from PIL import Image, ImageDraw, ImageFont
from typing import Optional, Dict, Any, Tuple, List
from utils import get_api_key, IMAGE_KEYWORDS
import random

def fetch_image_for_slide(slide_type: str, context: Dict[str, Any] = None, use_placeholders: bool = False) -> Optional[io.BytesIO]:
    """
    Fetch a relevant image from Unsplash for a slide type.
    
    Args:
        slide_type: Type of slide (problem, solution, advantage, audience)
        context: Additional context about the presentation to improve image relevance
        use_placeholders: If True, skip Unsplash and use placeholders directly
        
    Returns:
        BytesIO object containing the processed image or None if failed
    """
    # If placeholders are requested or Unsplash key is missing, use placeholders
    if use_placeholders:
        print(f"Using placeholder image for {slide_type} (placeholders mode enabled)")
        return create_placeholder_image(slide_type)
    
    try:
        # Check if Unsplash API key is available
        try:
            access_key = get_api_key("UNSPLASH_API_KEY")
        except ValueError:
            print(f"Unsplash API key not found. Using placeholder image for {slide_type}")
            return create_placeholder_image(slide_type)
        
        # Generate a focused search query using Gemini if context is provided
        if context and slide_type in context:
            search_query = generate_image_query_with_gemini(slide_type, context)
        else:
            # Fallback to predefined keywords
            keywords = IMAGE_KEYWORDS.get(slide_type, ["business"])
            search_query = random.choice(keywords)
        
        # Build the request URL with proper encoding
        import urllib.parse
        encoded_query = urllib.parse.quote(search_query)
        url = f"https://api.unsplash.com/photos/random?query={encoded_query}&orientation=landscape"
        
        # Make the request with proper headers
        headers = {
            "Authorization": f"Client-ID {access_key}",
            "Accept-Version": "v1"
        }
        
        print(f"Fetching image for '{search_query}' from Unsplash...")
        response = requests.get(url, headers=headers, timeout=10)
        
        # Check response status
        if response.status_code == 403:
            print(f"Unsplash API access forbidden (403). Using placeholder image.")
            return create_placeholder_image(slide_type)
        elif response.status_code == 429:
            print(f"Unsplash API rate limit exceeded (429). Using placeholder image.")
            return create_placeholder_image(slide_type)
        elif response.status_code != 200:
            print(f"Unsplash API returned status {response.status_code}. Using placeholder image.")
            return create_placeholder_image(slide_type)
            print(f"Unsplash API rate limit exceeded (429). Falling back to placeholder.")
            return create_placeholder_image(slide_type)
        
        response.raise_for_status()
        
        # Parse the response
        data = response.json()
        
        # Get the image URL (regular size is a good balance)
        image_url = data["urls"]["regular"]
        
        # Download the image
        image_response = requests.get(image_url, timeout=10)
        image_response.raise_for_status()
        
        print(f"Successfully fetched image for '{search_query}'")
        
        # Return the original image without processing
        image_data = io.BytesIO(image_response.content)
        image_data.seek(0)
        
        return image_data
    except Exception as e:
        print(f"Error fetching image: {str(e)}")
        print(f"Falling back to placeholder image for slide type: {slide_type}")
        return create_placeholder_image(slide_type)

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
            model_name="gemini-1.5-flash",
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