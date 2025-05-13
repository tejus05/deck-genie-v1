import requests
import io
import json
import cohere
from PIL import Image
from typing import Optional, Dict, Any, Tuple, List
from utils import get_api_key, IMAGE_KEYWORDS
import random

def fetch_image_for_slide(slide_type: str, context: Dict[str, Any] = None) -> Optional[io.BytesIO]:
    """
    Fetch a relevant image from Pixabay for a slide type.
    
    Args:
        slide_type: Type of slide (problem, solution, advantage, audience)
        context: Additional context about the presentation to improve image relevance
        
    Returns:
        BytesIO object containing the processed image or None if failed
    """
    try:
        # Get access key from environment
        access_key = get_api_key("PIXABAY_API_KEY")
        
        # Generate a focused search query using Cohere if context is provided
        if context and slide_type in context:
            search_query = generate_image_query_with_cohere(slide_type, context)
        else:
            # Fallback to predefined keywords
            keywords = IMAGE_KEYWORDS.get(slide_type, ["business"])
            search_query = random.choice(keywords)
        
        # Build the request URL
        url = f"https://pixabay.com/api/?key={access_key}&q={search_query}&image_type=photo&orientation=horizontal"
        
        # Make the request
        response = requests.get(url)
        response.raise_for_status()
        
        # Parse the response
        data = response.json()
        
        # Get the image URL (regular size is a good balance)
        image_url = data["hits"][0]["webformatURL"]
        
        # Download the image
        image_response = requests.get(image_url)
        image_response.raise_for_status()
        
        # Return the original image without processing
        image_data = io.BytesIO(image_response.content)
        image_data.seek(0)
        
        return image_data
    except Exception as e:
        print(f"Error fetching image: {str(e)}")
        return None

def generate_image_query_with_cohere(slide_type: str, context: Dict[str, Any]) -> str:
    """
    Use Cohere API to generate a relevant image search query based on slide content.
    
    Args:
        slide_type: Type of slide
        context: Context containing slide content and presentation information
        
    Returns:
        A search query string optimized for image relevance
    """
    try:
        # Initialize the Cohere client
        cohere_client = cohere.Client('8RCuJ6TE6fjsiojseWEn5Mc6v31fuapFcxKoa0nO')  # Your Cohere API key
        
        # Create a prompt based on slide type and content
        if slide_type == "problem":
            prompt = f"""
            Generate a single, specific Pixabay search query (3-5 words) for a business presentation slide about this problem:
            
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
            Generate a single, specific Pixabay search query (3-5 words) for a business presentation slide about this solution:
            
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
            Generate a single, specific Pixabay search query (3-5 words) for a business presentation slide about competitive advantages:
            
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
            Generate a single, specific Pixabay search query (3-5 words) for a business presentation slide about this target audience:
            
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
            company_name = context.get('metadata', {}).get('company_name', 'Business')
            product_name = context.get('metadata', {}).get('product_name', 'Product')
            prompt = f"""
            Generate a single, specific Pixabay search query (3-5 words) for a business presentation slide.
            
            Company: {company_name}
            Product: {product_name}
            
            IMPORTANT: Keep your query focused specifically on B2B SaaS, enterprise technology, or business contexts.
            Avoid generic imagery and focus on professional/corporate tech environments.
            
            Format your response as a JSON object with a single key "query" and its string value.
            Example: {{"query": "enterprise technology solution"}}
            """
        
        # Call Cohere API
        response = cohere_client.generate(
            model='command-xlarge-20220302', 
            prompt=prompt, 
            max_tokens=100, 
            temperature=0.3
        )
        
        # Parse the response
        result = response.generations[0].text.strip()
        
        # Return the query
        return result if result else "business"
            
    except Exception as e:
        print(f"Error generating image query with Cohere: {str(e)}")
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
