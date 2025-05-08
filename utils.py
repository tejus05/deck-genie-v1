import re
import os
from dotenv import load_dotenv
from io import BytesIO
from typing import Dict, List, Tuple

# Load environment variables
load_dotenv()

# Constants for presentation design
FONTS = {
    "title": "Helvetica Neue",
    "body": "Helvetica Neue",
    "fallback1": "Segoe UI",
    "fallback2": "Arial"
}

COLORS = {
    "black": "000000",
    "white": "FFFFFF",
    "gray": "666666"
}

FONT_SIZES = {
    "title": 44,
    "title_slide": 48,
    "subtitle": 24,
    "body": 24,
    "call_to_action": 48
}

# Icons for features and slides
FEATURE_ICONS = [
    "✓",  # Checkmark
    "⚙️",  # Gear
    "📊",  # Chart
    "👤",  # User
    "🔒",  # Lock
    "⚡",  # Lightning
    "🔄"   # Sync
]

SLIDE_ICONS = {
    "problem": "⚠️",  # Warning
    "solution": "💡",  # Light bulb
    "features": "⚙️",  # Gear
    "advantage": "🏆",  # Trophy
    "audience": "👥",  # People
    "call_to_action": "🚀"  # Rocket
}

# Image search keywords for each slide type
IMAGE_KEYWORDS = {
    "problem": ["business challenge", "frustration", "obstacle", "problem", "difficulty"],
    "solution": ["innovation", "technology solution", "business solution", "digital transformation"],
    "features": ["software features", "technology features", "saas interface"],
    "advantage": ["competitive advantage", "business success", "leadership", "winning"],
    "audience": ["business team", "professional meeting", "corporate team", "business client"]
}

# Layout measurements (in inches)
MARGINS = {
    "left": 0.5,
    "right": 0.5,
    "top": 0.5,
    "bottom": 0.5
}

CONTENT_AREA = {
    "left": MARGINS["left"],
    "top": 1.5,  # Below title
    "width": 4.5,
    "height": 5.5
}

IMAGE_AREA = {
    "left": 5.5,  # Right side
    "top": 1.5,
    "width": 4.0,
    "height": 5.5
}


def sanitize_filename(filename: str) -> str:
    """
    Sanitize a string to be used as a filename.
    
    Args:
        filename: The original string
        
    Returns:
        A sanitized version safe for use as a filename
    """
    # Replace spaces with underscores and remove special characters
    sanitized = re.sub(r'[^\w\s-]', '', filename)
    sanitized = re.sub(r'[\s]+', '_', sanitized)
    return sanitized


def truncate_text(text: str, max_length: int, suffix: str = "...") -> str:
    """
    Truncate text if it exceeds the maximum length.
    
    Args:
        text: The text to truncate
        max_length: Maximum allowed length
        suffix: String to append if truncated
        
    Returns:
        Truncated text with suffix if needed
    """
    if len(text) <= max_length:
        return text
    return text[:max_length-len(suffix)] + suffix


def get_api_key(key_name: str) -> str:
    """
    Get API key from environment variables.
    
    Args:
        key_name: Name of the API key in .env file
        
    Returns:
        API key as string
        
    Raises:
        ValueError: If API key is not found
    """
    api_key = os.getenv(key_name)
    if not api_key:
        raise ValueError(f"{key_name} not found in environment variables")
    return api_key


def match_icon_to_feature(feature: str) -> str:
    """
    Match an appropriate icon to a feature based on content.
    
    Args:
        feature: The feature text
        
    Returns:
        An icon character from FEATURE_ICONS
    """
    feature_lower = feature.lower()
    
    # Map features to appropriate icons based on keywords
    if any(keyword in feature_lower for keyword in ["secure", "privacy", "compliance", "protect"]):
        return FEATURE_ICONS[4]  # Lock
    elif any(keyword in feature_lower for keyword in ["fast", "speed", "quick", "rapid"]):
        return FEATURE_ICONS[5]  # Lightning
    elif any(keyword in feature_lower for keyword in ["sync", "integrate", "connect"]):
        return FEATURE_ICONS[6]  # Sync
    elif any(keyword in feature_lower for keyword in ["data", "analytics", "report", "insight"]):
        return FEATURE_ICONS[2]  # Chart
    elif any(keyword in feature_lower for keyword in ["user", "customer", "experience", "interface"]):
        return FEATURE_ICONS[3]  # User
    elif any(keyword in feature_lower for keyword in ["feature", "function", "capability", "tool"]):
        return FEATURE_ICONS[1]  # Gear
    else:
        return FEATURE_ICONS[0]  # Default to checkmark