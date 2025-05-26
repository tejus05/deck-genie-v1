import streamlit as st
from typing import Dict, Any, List
import time

class SlidePreviewGenerator:
    """Generates live preview cards for slides as they are created."""
    
    def __init__(self):
        """Initialize the preview generator."""
        self.preview_container = None
        
    def create_preview_container(self):
        """Create the container for live slide previews that persists throughout the session."""
        st.markdown("### 🔍 Live Preview")
        
        # Create persistent container that can be cleared and rewritten
        self.preview_container = st.empty()
        
        # Initialize session state for preview slides
        if 'preview_slides' not in st.session_state:
            st.session_state.preview_slides = []
    
    def reset_preview(self):
        """Reset the preview slides."""
        st.session_state.preview_slides = []
        self.render_all_preview_slides()
    
    def add_slide_preview(self, slide_data: Dict[str, Any], slide_type: str):
        """Add a slide to the preview."""
        # Create a preview slide object
        icon = self._get_slide_icon(slide_type)
        title = self._get_slide_title(slide_type)
        
        slide = {
            'type': slide_type,
            'icon': icon,
            'title': title,
            'data': slide_data
        }
        
        # Append to session state list
        st.session_state.preview_slides.append(slide)
        
        # Render all slides
        self.render_all_preview_slides()
    
    def render_all_preview_slides(self):
        """Render all preview slides with equal dimensions."""
        if not self.preview_container:
            self.preview_container = st.empty()
        
        # Clear and rebuild the entire preview using st.empty()
        with self.preview_container.container():
            slides = st.session_state.preview_slides
            if not slides:
                return
            
            num_slides = len(slides)
            # Determine optimal number of columns based on slide count
            # For fewer slides, use fewer columns for better visibility
            if num_slides <= 4:
                cols_per_row = 3  # Use 3 columns for 1-4 slides
            elif num_slides <= 8:
                cols_per_row = 4  # Use 4 columns for 5-8 slides
            else:
                cols_per_row = 5  # Use 5 columns for 9-10 slides
            
            num_rows = (num_slides + cols_per_row - 1) // cols_per_row  # Ceiling division
            
            # Use equal width for all columns
            for row in range(num_rows):
                cols = st.columns([1] * cols_per_row)  # Equal width distribution
                for col_idx in range(cols_per_row):
                    slide_idx = row * cols_per_row + col_idx
                    if slide_idx < num_slides:
                        slide = slides[slide_idx]
                        with cols[col_idx]:
                            self.render_single_slide_preview(slide)
    
    def render_single_slide_preview(self, slide: Dict[str, Any]):
        """Render a single slide preview card with consistent dimensions."""
        # Get content preview based on slide data
        content_preview = self.get_slide_content_preview(slide['data'])
        
        st.markdown(f"""
        <div style="
            border: 2px solid #0078D7; 
            border-radius: 12px; 
            padding: 10px; 
            text-align: center; 
            background: var(--background-color);
            margin: 5px 2px;
            height: 150px;
            width: 100%;
            display: flex;
            flex-direction: column;
            justify-content: flex-start;
            align-items: center;
            box-shadow: 0 2px 4px rgba(0,0,0,0.2);
            transition: transform 0.2s ease;
            overflow: hidden;
        ">
            <div style="font-size: 1.8em; margin-bottom: 4px;">{slide['icon']}</div>
            <div style="font-weight: bold; font-size: 0.85em; margin-bottom: 2px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; width: 100%;">{slide['title']}</div>
            <div style="font-size: 0.7em; line-height: 1.1; overflow: hidden; text-overflow: ellipsis; max-height: 65px; padding: 0 4px;">
                {content_preview}
            </div>
        </div>
        """, unsafe_allow_html=True)
    
    def get_slide_content_preview(self, slide_data: Dict[str, Any]) -> str:
        """Extract a meaningful preview from slide data."""
        if not slide_data:
            return "Generated content"
        
        # Try to get the most relevant content for preview (with shorter text limits)
        if 'title' in slide_data and slide_data['title']:
            return slide_data['title'][:40] + ('...' if len(slide_data['title']) > 40 else '')
        elif 'subtitle' in slide_data and slide_data['subtitle']:
            return slide_data['subtitle'][:40] + ('...' if len(slide_data['subtitle']) > 40 else '')
        elif 'content' in slide_data and slide_data['content']:
            return slide_data['content'][:40] + ('...' if len(slide_data['content']) > 40 else '')
        elif 'bullet_points' in slide_data and slide_data['bullet_points']:
            first_bullet = slide_data['bullet_points'][0] if slide_data['bullet_points'] else ""
            return first_bullet[:40] + ('...' if len(first_bullet) > 40 else '')
        elif 'bullets' in slide_data and slide_data['bullets']:
            first_bullet = slide_data['bullets'][0] if slide_data['bullets'] else ""
            return first_bullet[:40] + ('...' if len(first_bullet) > 40 else '')
        else:
            return "Content pending..."
    
    def _get_slide_icon(self, slide_type: str) -> str:
        """Get the icon for a given slide type."""
        icons = {
            "title_slide": "📊",
            "problem_slide": "⚠️",
            "solution_slide": "💡",
            "features_slide": "⚙️",
            "advantage_slide": "🏆",
            "audience_slide": "👥",
            "cta_slide": "🚀",
            "market_slide": "💰",
            "roadmap_slide": "🗓️",
            "team_slide": "🧑‍🤝‍🧑"
        }
        return icons.get(slide_type, "📄")
    
    def _get_slide_title(self, slide_type: str) -> str:
        """Get a human-readable title for a slide type."""
        titles = {
            "title_slide": "Title",
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
        return titles.get(slide_type, "Slide")

# Function to use outside the class for compatibility with older code
def simulate_slide_generation_with_preview(
    content_generator_func, 
    args, 
    progress_bar, 
    progress_status, 
    progress_detail,
    preview_generator: SlidePreviewGenerator,
    slide_placeholders: List = None
):
    """Simulate the generation of slides with a live preview.
    
    Args:
        content_generator_func: The function that generates presentation content
        args: Arguments to pass to the content generator function
        progress_bar: Streamlit progress bar component
        progress_status: Streamlit component for status text
        progress_detail: Streamlit component for detail text
        preview_generator: SlidePreviewGenerator instance
        slide_placeholders: Optional list of placeholders for each slide
    
    Returns:
        Generated content
    """
    # Simulate step 1: Planning (already done before this function)
    progress_bar.progress(20)
    
    # Simulate step 2: Content generation
    progress_status.markdown("### 🧠 Crafting the title slide...")
    progress_detail.markdown("_Creating an impactful introduction_")
    time.sleep(1)  # Simulate processing time
    progress_bar.progress(30)
    
    # Simulate step 3: Problem slide
    progress_status.markdown("### 🧠 Articulating the problem...")
    progress_detail.markdown("_Defining the challenge your product solves_")
    time.sleep(1)  # Simulate processing time
    progress_bar.progress(40)
    
    # Simulate step 4: Solution slide
    progress_status.markdown("### 🧠 Designing the solution narrative...")
    progress_detail.markdown("_Explaining how your product addresses the challenge_")
    time.sleep(1)  # Simulate processing time
    progress_bar.progress(50)
    
    # Simulate step 5: Features slide
    progress_status.markdown("### 🧠 Highlighting key features...")
    progress_detail.markdown("_Emphasizing what makes your product valuable_")
    time.sleep(1)  # Simulate processing time
    progress_bar.progress(60)
    
    # Generate content for real
    content = content_generator_func(*args)
    
    # Update preview with generated content
    for slide_type, slide_data in content.items():
        if slide_type != "metadata":
            # Add a very small delay to make the updates visible
            preview_generator.add_slide_preview(slide_data, slide_type)
    
    return content
