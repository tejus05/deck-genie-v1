import requests
import io
from PIL import Image, ImageOps
from typing import Optional, Dict, Any, Tuple
from utils import get_api_key, IMAGE_KEYWORDS
import random

def fetch_image_for_slide(slide_type: str) -> Optional[io.BytesIO]:
    """
    Fetch a relevant image from Unsplash for a slide type.
    
    Args:
        slide_type: Type of slide (problem, solution, advantage, audience)
        
    Returns:
        BytesIO object containing the processed image or None if failed
    """
    try:
        # Get access key from environment
        access_key = get_api_key("UNSPLASH_API_KEY")
        
        # Get keywords for the slide type
        keywords = IMAGE_KEYWORDS.get(slide_type, ["business"])
        
        # Select a random keyword from the list
        keyword = random.choice(keywords)
        
        # Build the request URL
        url = f"https://api.unsplash.com/photos/random?query={keyword}&orientation=landscape"
        
        # Make the request
        headers = {"Authorization": f"Client-ID {access_key}"}
        response = requests.get(url, headers=headers)
        response.raise_for_status()
        
        # Parse the response
        data = response.json()
        
        # Get the image URL (regular size is a good balance)
        image_url = data["urls"]["regular"]
        
        # Download the image
        image_response = requests.get(image_url)
        image_response.raise_for_status()
        
        # Process the image (convert to grayscale)
        image_data = process_image(io.BytesIO(image_response.content))
        
        return image_data
    except Exception as e:
        print(f"Error fetching image: {str(e)}")
        return None

def process_image(image_data: io.BytesIO) -> io.BytesIO:
    """
    Process an image for use in the presentation.
    - Convert to grayscale
    - Apply subtle contrast enhancement
    
    Args:
        image_data: BytesIO object with the original image
        
    Returns:
        BytesIO object with the processed image
    """
    try:
        # Open the image using PIL
        image = Image.open(image_data)
        
        # Convert to grayscale
        gray_image = ImageOps.grayscale(image)
        
        # Enhance contrast slightly
        gray_image = ImageOps.autocontrast(gray_image, cutoff=0.5)
        
        # Save to a BytesIO object
        output = io.BytesIO()
        gray_image.save(output, format='PNG')
        output.seek(0)
        
        return output
    except Exception as e:
        print(f"Error processing image: {str(e)}")
        # Return original if processing fails
        image_data.seek(0)
        return image_data

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