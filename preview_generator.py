import streamlit as st
from typing import Dict, Any, List
import time

class SlidePreviewGenerator:
    """Generates live preview cards for slides as they are created."""
    
    def __init__(self):
        self.slide_titles = [
            "Title Slide",
            "Problem Statement", 
            "Solution Overview",
            "Key Features",
            "Competitive Advantage",
            "Target Audience", 
            "Call to Action"
        ]
        self.slide_icons = [
            "📊", "⚠️", "💡", "⚙️", "🏆", "👥", "🚀"
        ]
        self.preview_container = None
    
    def create_preview_container(self):
        """Create the container for live slide previews that persists throughout the session."""
        st.markdown("### 🔍 Live Preview")
        
        # Create persistent container that can be cleared and rewritten
        self.preview_container = st.empty()
        
        # Initialize session state for preview slides
        if 'preview_slides' not in st.session_state:
            st.session_state.preview_slides = []
        
        # If we have slides, render them immediately
        if st.session_state.preview_slides:
            self.render_all_preview_slides()
        
        return self.preview_container, None
    
    def add_slide_to_preview(self, slide_index: int, slide_data: Dict[str, Any]):
        """Add a new slide to the live preview, showing progressive generation."""
        if slide_index < len(self.slide_titles):
            title = self.slide_titles[slide_index]
            icon = self.slide_icons[slide_index]
            
            # Extract actual title from slide data if available
            if slide_data and 'title' in slide_data:
                actual_title = slide_data['title']
            else:
                actual_title = title
            
            # Add to session state
            slide_preview = {
                'index': slide_index,
                'title': title,
                'icon': icon,
                'actual_title': actual_title,
                'data': slide_data
            }
            
            st.session_state.preview_slides.append(slide_preview)
            
            # Force immediate re-render
            self.render_all_preview_slides()
            
            # Force streamlit to update the display
            time.sleep(0.1)
    
    def render_all_preview_slides(self):
        """Render all preview slides."""
        if not self.preview_container:
            self.preview_container = st.empty()
        
        # Clear and rebuild the entire preview using st.empty()
        with self.preview_container.container():
            slides = st.session_state.preview_slides
            if not slides:
                return
            
            num_slides = len(slides)
            num_rows = (num_slides + 3) // 4  # Ceiling division
            
            for row in range(num_rows):
                cols = st.columns(4)
                for col_idx in range(4):
                    slide_idx = row * 4 + col_idx
                    if slide_idx < num_slides:
                        slide = slides[slide_idx]
                        with cols[col_idx]:
                            self.render_single_slide_preview(slide)
    
    def render_single_slide_preview(self, slide: Dict[str, Any]):
        """Render a single slide preview card with better styling."""
        # Get content preview based on slide data
        content_preview = self.get_slide_content_preview(slide['data'])
        
        st.markdown(f"""
        <div style="
            border: 2px solid #4CAF50; 
            border-radius: 12px; 
            padding: 16px; 
            text-align: center; 
            background: linear-gradient(135deg, #f8f9fa 0%, #e9ecef 100%);
            margin: 8px 0;
            min-height: 140px;
            display: flex;
            flex-direction: column;
            justify-content: center;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
            transition: transform 0.2s ease;
        ">
            <div style="font-size: 2.2em; margin-bottom: 8px;">{slide['icon']}</div>
            <div style="font-weight: bold; font-size: 1em; margin-bottom: 4px; color: #2c3e50;">{slide['title']}</div>
            <div style="font-size: 0.85em; color: #495057; line-height: 1.3; overflow: hidden; text-overflow: ellipsis;">
                {content_preview}
            </div>
        </div>
        """, unsafe_allow_html=True)
    
    def get_slide_content_preview(self, slide_data: Dict[str, Any]) -> str:
        """Extract a meaningful preview from slide data."""
        if not slide_data:
            return "Generated content"
        
        # Try to get the most relevant content for preview
        if 'title' in slide_data and slide_data['title']:
            return slide_data['title'][:60] + ('...' if len(slide_data['title']) > 60 else '')
        elif 'subtitle' in slide_data and slide_data['subtitle']:
            return slide_data['subtitle'][:60] + ('...' if len(slide_data['subtitle']) > 60 else '')
        elif 'content' in slide_data and slide_data['content']:
            return slide_data['content'][:60] + ('...' if len(slide_data['content']) > 60 else '')
        elif 'bullet_points' in slide_data and slide_data['bullet_points']:
            first_bullet = slide_data['bullet_points'][0] if slide_data['bullet_points'] else ""
            return first_bullet[:60] + ('...' if len(first_bullet) > 60 else '')
        else:
            return "Content generated"
    
    def update_preview_from_session_state(self, editor_content: Dict[str, Any], slide_order: List[str], deleted_slides: set):
        """Update preview based on current editor state - for real-time updates."""
        # Clear current preview
        st.session_state.preview_slides = []
        
        # Rebuild preview based on current slide order and content
        slide_type_to_index = {
            'title_slide': 0, 'problem_slide': 1, 'solution_slide': 2, 'features_slide': 3,
            'advantage_slide': 4, 'audience_slide': 5, 'cta_slide': 6
        }
        
        for slide_key in slide_order:
            if slide_key not in deleted_slides and slide_key in editor_content:
                original_index = slide_type_to_index.get(slide_key, 0)
                slide_data = editor_content[slide_key]
                
                # Add to preview with current content
                slide_preview = {
                    'index': original_index,
                    'title': self.slide_titles[original_index],
                    'icon': self.slide_icons[original_index],
                    'actual_title': slide_data.get('title', self.slide_titles[original_index]),
                    'data': slide_data
                }
                st.session_state.preview_slides.append(slide_preview)
        
        # Re-render all slides
        self.render_all_preview_slides()
    
    def clear_preview(self):
        """Clear the preview completely."""
        st.session_state.preview_slides = []
        if hasattr(self, 'preview_container') and self.preview_container:
            self.preview_container.empty()

    def update_preview_with_content(self, content: Dict[str, Any]):
        """Update preview with full content after generation."""
        # Clear current preview
        st.session_state.preview_slides = []
        
        # Rebuild preview from content
        slide_types = ['title_slide', 'problem_slide', 'solution_slide', 'features_slide', 
                      'advantage_slide', 'audience_slide', 'cta_slide']
        
        for i, slide_type in enumerate(slide_types):
            if slide_type in content:
                slide_preview = {
                    'index': i,
                    'title': self.slide_titles[i],
                    'icon': self.slide_icons[i],
                    'actual_title': content[slide_type].get('title', self.slide_titles[i]),
                    'data': content[slide_type]
                }
                st.session_state.preview_slides.append(slide_preview)
        
        # Re-render all slides
        self.render_all_preview_slides()

def simulate_slide_generation_with_preview(
    content_generator_func, 
    args, 
    progress_bar, 
    progress_status, 
    progress_detail,
    preview_generator: SlidePreviewGenerator,
    slide_placeholders: List = None
):
    """
    Simulate slide generation with true progressive live preview updates.
    Shows slides one by one as they are generated.
    """
    # Clear any existing preview slides
    st.session_state.preview_slides = []
    preview_generator.clear_preview()
    
    # Call the actual content generation function first
    try:
        progress_status.markdown("### 🧠 Generating content...")
        progress_detail.markdown("_Creating compelling slide content..._")
        progress_bar.progress(20)
        
        content = content_generator_func(*args)
        
        progress_status.markdown("### 🎨 Building slides progressively...")
        progress_detail.markdown("_Watch your presentation come to life!_")
        
        # Now simulate progressive slide generation for visual effect
        slide_types = ['title_slide', 'problem_slide', 'solution_slide', 'features_slide', 
                      'advantage_slide', 'audience_slide', 'cta_slide']
        
        slide_names = ["Title", "Problem", "Solution", "Features", "Advantage", "Audience", "CTA"]
        
        for i, slide_type in enumerate(slide_types):
            # Update progress status
            progress_detail.markdown(f"_Building {slide_names[i]} slide..._")
            
            # Progressive delay for visual effect (shorter for better UX)
            time.sleep(0.4)
            
            # Update progress
            progress_percent = 25 + (i + 1) * (40 / len(slide_types))  # 25-65% for slide generation
            progress_bar.progress(int(progress_percent))
            
            # Add slide to preview if it exists
            if slide_type in content:
                preview_generator.add_slide_to_preview(i, content[slide_type])
                
                # Small additional delay to show the slide being added
                time.sleep(0.2)
        
        return content
        
    except Exception as e:
        # Clear preview on error
        st.session_state.preview_slides = []
        preview_generator.clear_preview()
        raise e
