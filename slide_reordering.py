import streamlit as st
from typing import Dict, Any, List, Optional
import io
from ppt_generator import create_presentation

def get_standard_slide_order():
    """Get the standard order of slides in a presentation."""
    return [
        'title_slide', 
        'problem_slide', 
        'solution_slide', 
        'features_slide', 
        'advantage_slide', 
        'audience_slide',
        'market_slide',
        'roadmap_slide',
        'team_slide',
        'cta_slide'
    ]

def get_slide_title(slide_type: str) -> str:
    """Get a human-readable title for a slide type."""
    titles = {
        "title_slide": "Title Slide",
        "problem_slide": "Problem Statement",
        "solution_slide": "Solution Overview",
        "features_slide": "Key Features",
        "advantage_slide": "Competitive Advantage",
        "audience_slide": "Target Audience",
        "cta_slide": "Call to Action",
        "market_slide": "Market Analysis",
        "roadmap_slide": "Product Roadmap",
        "team_slide": "Team Overview"
    }
    return titles.get(slide_type, slide_type.replace('_', ' ').title())

def initialize_slide_order(content: Dict[str, Any]) -> List[str]:
    """
    Initialize the slide order from presentation content, 
    preserving the original presentation order.
    
    Args:
        content: Dictionary containing presentation content
        
    Returns:
        List of slide types in proper order
    """
    # Get standard order as a reference
    standard_order = get_standard_slide_order()
    
    # Extract slide types from content (keys ending with _slide)
    available_slides = [k for k in content.keys() if k.endswith('_slide')]
    
    # Organize slides in the standard order if they exist in available_slides
    ordered_slides = [slide for slide in standard_order if slide in available_slides]
    
    # Add any additional slides not in standard_order but in available_slides
    for slide in available_slides:
        if slide not in ordered_slides:
            ordered_slides.append(slide)
    
    return ordered_slides

def generate_reordered_presentation(
    original_content: Dict[str, Any], 
    custom_slide_order: List[str],
    filename: str,
    image_manager=None
) -> bytes:
    """
    Generate a reordered presentation based on the original content and a custom slide order.
    
    Args:
        original_content: Dictionary containing the original presentation content
        custom_slide_order: List specifying the new order of slides
        filename: Name for the output file
        image_manager: Optional image manager for custom images
        
    Returns:
        bytes: Binary data for the PowerPoint file
    """
    try:
        # Create a copy of the content to avoid modifying the original
        content_copy = {k: v for k, v in original_content.items()}
        
        # Generate the presentation with the custom slide order
        presentation_buffer = create_presentation(
            content_copy,
            filename,
            image_manager=image_manager,
            custom_slide_order=custom_slide_order
        )
        
        return presentation_buffer
    except Exception as e:
        st.error(f"Error generating reordered presentation: {str(e)}")
        # Return None or a fallback presentation
        return None

def render_slide_reordering_ui(content: Dict[str, Any]) -> None:
    """
    Render UI elements for reordering slides, ensuring the initial order
    matches the original presentation.
    
    Args:
        content: Dictionary containing presentation content
    """
    # Initialize slide order if not already in session state
    if 'slide_order' not in st.session_state:
        st.session_state.slide_order = initialize_slide_order(content)
    
    # Initialize modifications flag if it doesn't exist
    if 'has_modifications' not in st.session_state:
        st.session_state.has_modifications = False
    
    # Display current slide order with up/down buttons
    for i, slide_type in enumerate(st.session_state.slide_order):
        col1, col2, col3 = st.columns([5, 1, 1])
        
        slide_title = get_slide_title(slide_type)
        with col1:
            st.markdown(f"**{i+1}. {slide_title}**")
        
        # Move Up button
        with col2:
            if i > 0:  # Can't move up if it's already the first slide
                if st.button("↑", key=f"up_{slide_type}_{i}"):
                    # Swap this slide with the one above it
                    st.session_state.slide_order[i], st.session_state.slide_order[i-1] = \
                        st.session_state.slide_order[i-1], st.session_state.slide_order[i]
                    st.session_state.has_modifications = True
                    # Need to rerun to show updated order
                    st.rerun()
            else:
                st.write("")  # Empty space for alignment
        
        # Move Down button
        with col3:
            if i < len(st.session_state.slide_order) - 1:  # Can't move down if it's already the last slide
                if st.button("↓", key=f"down_{slide_type}_{i}"):
                    # Swap this slide with the one below it
                    st.session_state.slide_order[i], st.session_state.slide_order[i+1] = \
                        st.session_state.slide_order[i+1], st.session_state.slide_order[i]
                    st.session_state.has_modifications = True
                    # Need to rerun to show updated order
                    st.rerun()
            else:
                st.write("")  # Empty space for alignment
