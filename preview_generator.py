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
    
    def create_preview_container(self):
        """Create the container for live slide previews."""
        st.markdown("### 🔍 Live Preview")
        preview_container = st.container()
        
        # Create placeholder for slides
        cols = st.columns(4)  # 4 slides per row
        slide_placeholders = []
        
        for i in range(len(self.slide_titles)):
            col_index = i % 4
            with cols[col_index]:
                placeholder = st.empty()
                slide_placeholders.append(placeholder)
        
        return preview_container, slide_placeholders
    
    def update_slide_preview(self, slide_index: int, slide_data: Dict[str, Any], placeholder):
        """Update a specific slide preview."""
        if slide_index < len(self.slide_titles):
            title = self.slide_titles[slide_index]
            icon = self.slide_icons[slide_index]
            
            # Extract actual title from slide data if available
            if slide_data and 'title' in slide_data:
                actual_title = slide_data['title']
            else:
                actual_title = title
            
            # Create a preview card
            with placeholder.container():
                st.markdown(f"""
                <div style="
                    border: 2px solid #e0e0e0; 
                    border-radius: 10px; 
                    padding: 20px; 
                    text-align: center; 
                    background-color: #f9f9f9;
                    margin: 10px 0;
                    min-height: 120px;
                    display: flex;
                    flex-direction: column;
                    justify-content: center;
                ">
                    <div style="font-size: 2em; margin-bottom: 10px;">{icon}</div>
                    <div style="font-weight: bold; font-size: 1.1em; margin-bottom: 5px;">{title}</div>
                    <div style="font-size: 0.9em; color: #666; overflow: hidden; text-overflow: ellipsis;">
                        {actual_title[:50]}{'...' if len(actual_title) > 50 else ''}
                    </div>
                </div>
                """, unsafe_allow_html=True)
    
    def show_placeholder_slide(self, slide_index: int, placeholder):
        """Show a placeholder for a slide that hasn't been generated yet."""
        if slide_index < len(self.slide_titles):
            title = self.slide_titles[slide_index]
            
            with placeholder.container():
                st.markdown(f"""
                <div style="
                    border: 2px dashed #d0d0d0; 
                    border-radius: 10px; 
                    padding: 20px; 
                    text-align: center; 
                    background-color: #f5f5f5;
                    margin: 10px 0;
                    min-height: 120px;
                    display: flex;
                    flex-direction: column;
                    justify-content: center;
                    opacity: 0.5;
                ">
                    <div style="font-size: 1.5em; margin-bottom: 10px;">⏳</div>
                    <div style="font-weight: bold; font-size: 1.1em; color: #999;">{title}</div>
                    <div style="font-size: 0.8em; color: #999;">Generating...</div>
                </div>
                """, unsafe_allow_html=True)

def simulate_slide_generation_with_preview(
    content_generator_func, 
    args, 
    progress_bar, 
    progress_status, 
    progress_detail,
    preview_generator: SlidePreviewGenerator,
    slide_placeholders: List
):
    """
    Simulate slide generation with live preview updates.
    This function wraps the actual content generation and updates previews.
    """
    # Show all placeholders initially
    for i, placeholder in enumerate(slide_placeholders):
        preview_generator.show_placeholder_slide(i, placeholder)
    
    # Call the actual content generation function
    try:
        content = content_generator_func(*args)
        
        # Update previews as each slide is "generated"
        slide_types = ['title_slide', 'problem_slide', 'solution_slide', 'features_slide', 
                      'advantage_slide', 'audience_slide', 'cta_slide']
        
        for i, slide_type in enumerate(slide_types):
            if i < len(slide_placeholders):
                # Simulate progressive generation
                time.sleep(0.3)  # Small delay for visual effect
                
                # Update progress
                progress_percent = 30 + (i + 1) * (30 / len(slide_types))  # 30-60% for slide generation
                progress_bar.progress(int(progress_percent))
                
                if slide_type in content:
                    preview_generator.update_slide_preview(i, content[slide_type], slide_placeholders[i])
                
                # Update progress status for each slide
                slide_names = ["Title", "Problem", "Solution", "Features", "Advantage", "Audience", "CTA"]
                if i < len(slide_names):
                    progress_detail.markdown(f"_Generated {slide_names[i]} slide..._")
        
        return content
        
    except Exception as e:
        # Clear all placeholders on error
        for placeholder in slide_placeholders:
            placeholder.empty()
        raise e
