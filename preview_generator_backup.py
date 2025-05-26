import streamlit as st
from typing import Dict, Any, List
import time

class SlidePreviewGenerator:
    """Generates live preview cards for slides as they are created."""
    
    def __init__(self):
        # Default slide titles and icons - these will be updated based on persona and slide count
        self.slide_titles = [
            "Title Slide",
            "Problem Statement", 
            "Solution Overview",
            "Key Features",
            "Competitive Advantage",
            "Target Audience", 
            "Market Opportunity",
            "Roadmap", 
            "Team", 
            "Call to Action"
        ]
        self.slide_icons = [
            "📊", "⚠️", "💡", "⚙️", "🏆", "👥", "💰", "🗓️", "👨‍💼", "🚀"
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
                    slide_idx = row * cols_per_row + col_idx                    if slide_idx < num_slides:
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
        elif 'paragraph' in slide_data and slide_data['paragraph']:
            return slide_data['paragraph'][:40] + ('...' if len(slide_data['paragraph']) > 40 else '')
        else:
            return "Content generated"
    
    def update_preview_from_session_state(self, editor_content: Dict[str, Any], slide_order: List[str], deleted_slides: set):
        """Update preview based on current editor state - for real-time updates."""
        # Clear current preview
        st.session_state.preview_slides = []
        
        # Rebuild preview based on current slide order and content
        slide_type_to_index = {
            'title_slide': 0, 'problem_slide': 1, 'solution_slide': 2, 'features_slide': 3,
            'advantage_slide': 4, 'audience_slide': 5, 'market_slide': 6, 'roadmap_slide': 7,
            'team_slide': 8, 'cta_slide': 9
        }
        
        for slide_key in slide_order:
            if slide_key not in deleted_slides and slide_key in editor_content:
                original_index = slide_type_to_index.get(slide_key, 0)
                slide_data = editor_content[slide_key]
                
                # Add to preview with current content
                slide_preview = {
                    'index': original_index,
                    'title': self.slide_titles[original_index] if original_index < len(self.slide_titles) else slide_key.replace("_", " ").title(),
                    'icon': self.slide_icons[original_index] if original_index < len(self.slide_icons) else "📝",
                    'actual_title': slide_data.get('title', self.slide_titles[original_index] if original_index < len(self.slide_titles) else ""),
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
        
        # Map slide keys to indices for ordering
        slide_type_to_index = {
            'title_slide': 0, 
            'problem_slide': 1, 
            'solution_slide': 2, 
            'features_slide': 3,
            'advantage_slide': 4, 
            'audience_slide': 5, 
            'market_slide': 6, 
            'roadmap_slide': 7,
            'team_slide': 8, 
            'cta_slide': 9
        }
        
        # Build a list of slides and sort by their index
        preview_slides = []
        for slide_key, slide_data in content.items():
            if slide_key == 'metadata':
                continue
                
            slide_index = slide_type_to_index.get(slide_key, 99)  # Default high index for unknown slides
            
            slide_preview = {
                'index': slide_index,
                'title': self.slide_titles[slide_index] if slide_index < len(self.slide_titles) else slide_key.replace("_", " ").title(),
                'icon': self.slide_icons[slide_index] if slide_index < len(self.slide_icons) else "📝",
                'actual_title': slide_data.get('title', self.slide_titles[slide_index] if slide_index < len(self.slide_titles) else ""),
                'data': slide_data
            }
            preview_slides.append(slide_preview)
        
        # Sort by index and add to session state
        preview_slides.sort(key=lambda x: x['index'])
        st.session_state.preview_slides = preview_slides
        
        # Re-render all slides
        self.render_all_preview_slides()

    def simulate_slide_generation_with_preview(
        self,
        content_generator_func, 
        args, 
        progress_bar, 
        progress_status, 
        progress_detail,
        slide_placeholders: List = None
    ):
        """
        Simulate slide generation with true progressive live preview updates.
        Shows slides one by one as they are generated.
        """
        # Clear any existing preview slides
        st.session_state.preview_slides = []
        self.clear_preview()
        
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
                          'advantage_slide', 'audience_slide', 'market_slide', 'roadmap_slide', 
                          'team_slide', 'cta_slide']
            
            slide_names = ["Title", "Problem", "Solution", "Features", "Advantage", "Audience", 
                          "Market", "Roadmap", "Team", "CTA"]
            
            # Get the selected slides from content, ensuring we respect persona
            persona = content.get("metadata", {}).get("persona", "Generic")
            slide_count = content.get("metadata", {}).get("slide_count", 8)
            
            # Define slide priorities based on persona - should match content_generator.py
            slide_priorities = {
                "Generic": [
                    'title_slide', 'problem_slide', 'solution_slide',  
                    'features_slide', 'advantage_slide', 'audience_slide',  
                    'market_slide', 'roadmap_slide', 'team_slide', 'cta_slide'
                ],
                "Technical": [
                    'title_slide', 'problem_slide', 'solution_slide',  
                    'features_slide', 'audience_slide', 'advantage_slide', 
                    'roadmap_slide', 'team_slide', 'market_slide', 'cta_slide'
                ],
                "Marketing": [
                    'title_slide', 'solution_slide', 'problem_slide',   
                    'advantage_slide', 'audience_slide', 'features_slide',  
                    'market_slide', 'team_slide', 'roadmap_slide', 'cta_slide'
                ],
                "Executive": [
                    'title_slide', 'market_slide', 'problem_slide',   
                    'solution_slide', 'advantage_slide', 'features_slide',  
                    'roadmap_slide', 'audience_slide', 'team_slide', 'cta_slide'
                ],
                "Investor": [
                    'title_slide', 'market_slide', 'problem_slide',   
                    'solution_slide', 'team_slide', 'advantage_slide', 
                    'roadmap_slide', 'features_slide', 'audience_slide', 'cta_slide'
                ]
            }
            
            # Get ordered slide types for this persona
            selected_priorities = slide_priorities.get(persona, slide_priorities["Generic"])
            
            # Only process slides that are in the content
            available_slides = [slide_type for slide_type in selected_priorities if slide_type in content]
            
            # Limit to the actual slide count
            available_slides = available_slides[:slide_count]
            
            for i, slide_type in enumerate(available_slides):
                # Map the slide type to the correct name and icon index
                slide_type_index = slide_types.index(slide_type) if slide_type in slide_types else i
                slide_name = slide_names[slide_type_index] if slide_type_index < len(slide_names) else slide_type.replace("_", " ").title()
                
                # Update progress status
                progress_detail.markdown(f"_Building {slide_name} slide..._")
                
                # Progressive delay for visual effect (shorter for better UX)
                time.sleep(0.4)
                
                # Update progress
                progress_percent = 25 + (i + 1) * (40 / len(available_slides))  # 25-65% for slide generation
                progress_bar.progress(int(progress_percent))
                
                # Add slide to preview
                self.add_slide_to_preview(slide_type_index, content[slide_type])
                
                # Small additional delay to show the slide being added
                time.sleep(0.2)
            
            return content
            
        except Exception as e:
            # Clear preview on error
            st.session_state.preview_slides = []
            self.clear_preview()
            raise e


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
    """
    Wrapper function that delegates to the class method.
    For backward compatibility with existing code.
    """
    return preview_generator.simulate_slide_generation_with_preview(
        content_generator_func, 
        args, 
        progress_bar, 
        progress_status, 
        progress_detail,
        slide_placeholders
    )