import streamlit as st
from typing import Dict, Any, List, Optional
import copy
import io
from PIL import Image

class SlideEditor:
    """Handles slide editing, reordering, and deletion functionality."""
    
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
        self.slide_keys = [
            "title_slide",
            "problem_slide", 
            "solution_slide",
            "features_slide",
            "advantage_slide",
            "audience_slide", 
            "cta_slide"
        ]
    
    def initialize_editor_state(self, original_content: Dict[str, Any]):
        """Initialize the editor state in session state."""
        if 'editor_content' not in st.session_state:
            st.session_state.editor_content = copy.deepcopy(original_content)
        
        if 'slide_order' not in st.session_state:
            st.session_state.slide_order = self.slide_keys.copy()
        
        # Ensure deleted_slides is always a set of strings
        if 'deleted_slides' not in st.session_state:
            st.session_state.deleted_slides = set()
        elif not isinstance(st.session_state.deleted_slides, set):
            # Fix corrupted data - convert to set
            st.session_state.deleted_slides = set()
        else:
            # Ensure all items in the set are strings
            cleaned_deleted = set()
            for item in st.session_state.deleted_slides:
                if isinstance(item, str):
                    cleaned_deleted.add(item)
            st.session_state.deleted_slides = cleaned_deleted
        
        if 'has_modifications' not in st.session_state:
            st.session_state.has_modifications = False
        
        if 'uploaded_images' not in st.session_state:
            st.session_state.uploaded_images = {}
    
    def render_slide_editor(self, original_content: Dict[str, Any]):
        """Render the slide editor interface."""
        self.initialize_editor_state(original_content)
        
        st.markdown("---")
        st.markdown("## 📝 Customize Your Presentation")
        
        # Slide reordering section
        st.markdown("### 🔄 Reorder Slides")
        self._render_slide_reordering()
        
        # Individual slide editors
        st.markdown("### ✏️ Edit Slide Content")
        self._render_individual_slide_editors()
        
        # Download buttons
        self._render_download_buttons(original_content)
    
    def _ensure_deleted_slides_safe(self):
        """Ensure deleted_slides is always a safe set of strings."""
        if not isinstance(st.session_state.deleted_slides, set):
            st.session_state.deleted_slides = set()
        else:
            # Clean any non-string items
            try:
                st.session_state.deleted_slides = {
                    item for item in st.session_state.deleted_slides 
                    if isinstance(item, str)
                }
            except (TypeError, AttributeError):
                st.session_state.deleted_slides = set()

    def _render_slide_reordering(self):
        """Render the slide reordering interface."""
        # Ensure deleted_slides is safe to use
        self._ensure_deleted_slides_safe()
        
        # Create list of active slides
        active_slides = [
            (key, title) for key, title in zip(self.slide_keys, self.slide_titles)
            if key not in st.session_state.deleted_slides
        ]
        
        if active_slides:
            # Ensure safety again before using deleted_slides
            self._ensure_deleted_slides_safe()
            
            # Create list of slide names for current order
            current_order = [
                self.slide_titles[self.slide_keys.index(key)] 
                for key in st.session_state.slide_order 
                if key not in st.session_state.deleted_slides
            ]
            
            st.info("📝 Use the ⬆️ ⬇️ buttons below to reorder slides:")
            
            # Simple up/down button interface
            for i, slide_title in enumerate(current_order):
                col1, col2, col3 = st.columns([2, 1, 1])
                with col1:
                    st.write(f"{i+1}. {slide_title}")
                with col2:
                    if i > 0 and st.button("⬆️", key=f"up_{i}"):
                        # Move slide up
                        slide_key = None
                        for key, title in zip(self.slide_keys, self.slide_titles):
                            if title == slide_title:
                                slide_key = key
                                break
                        if slide_key:
                            current_index = st.session_state.slide_order.index(slide_key)
                            st.session_state.slide_order[current_index], st.session_state.slide_order[current_index-1] = \
                                st.session_state.slide_order[current_index-1], st.session_state.slide_order[current_index]
                            st.session_state.has_modifications = True
                            st.rerun()
                with col3:
                    if i < len(current_order) - 1 and st.button("⬇️", key=f"down_{i}"):
                        # Move slide down
                        slide_key = None
                        for key, title in zip(self.slide_keys, self.slide_titles):
                            if title == slide_title:
                                slide_key = key
                                break
                        if slide_key:
                            current_index = st.session_state.slide_order.index(slide_key)
                            st.session_state.slide_order[current_index], st.session_state.slide_order[current_index+1] = \
                                st.session_state.slide_order[current_index+1], st.session_state.slide_order[current_index]
                            st.session_state.has_modifications = True
                            st.rerun()
    
    def _render_individual_slide_editors(self):
        """Render editors for individual slides."""
        # Ensure deleted_slides is safe to use
        self._ensure_deleted_slides_safe()
        
        for slide_key in st.session_state.slide_order:
            if slide_key not in st.session_state.deleted_slides:
                slide_index = self.slide_keys.index(slide_key)
                slide_title = self.slide_titles[slide_index]
                
                with st.expander(f"📄 {slide_title}", expanded=False):
                    self._render_single_slide_editor(slide_key, slide_title)
    
    def _render_single_slide_editor(self, slide_key: str, slide_title: str):
        """Render editor for a single slide."""
        if slide_key not in st.session_state.editor_content:
            st.warning(f"Content for {slide_title} not found.")
            return
        
        slide_content = st.session_state.editor_content[slide_key]
        
        # Delete slide button
        col1, col2 = st.columns([3, 1])
        with col2:
            if st.button(f"🗑️ Delete Slide", key=f"delete_{slide_key}"):
                st.session_state.deleted_slides.add(slide_key)
                st.session_state.has_modifications = True
                st.rerun()
        
        with col1:
            st.markdown(f"**Editing: {slide_title}**")
        
        # Title editing (for most slides)
        if 'title' in slide_content:
            new_title = st.text_input(
                "Slide Title",
                value=slide_content['title'],
                key=f"title_{slide_key}"
            )
            if new_title != slide_content['title']:
                st.session_state.editor_content[slide_key]['title'] = new_title
                st.session_state.has_modifications = True
        
        # Content editing based on slide type
        if slide_key == 'title_slide':
            self._edit_title_slide(slide_content, slide_key)
        elif slide_key in ['problem_slide', 'advantage_slide']:
            self._edit_bullet_slide(slide_content, slide_key)
        elif slide_key in ['solution_slide', 'audience_slide']:
            self._edit_paragraph_slide(slide_content, slide_key)
        elif slide_key == 'features_slide':
            self._edit_features_slide(slide_content, slide_key)
        elif slide_key == 'cta_slide':
            self._edit_cta_slide(slide_content, slide_key)
        
        # Image upload section
        self._render_image_upload(slide_key, slide_title)
    
    def _edit_title_slide(self, slide_content: Dict[str, Any], slide_key: str):
        """Edit title slide content."""
        if 'subtitle' in slide_content:
            new_subtitle = st.text_input(
                "Subtitle",
                value=slide_content['subtitle'],
                key=f"subtitle_{slide_key}"
            )
            if new_subtitle != slide_content['subtitle']:
                st.session_state.editor_content[slide_key]['subtitle'] = new_subtitle
                st.session_state.has_modifications = True
    
    def _edit_bullet_slide(self, slide_content: Dict[str, Any], slide_key: str):
        """Edit slides with bullet points."""
        if 'bullets' in slide_content:
            st.markdown("**Bullet Points:**")
            bullets = slide_content['bullets'].copy()
            
            for i, bullet in enumerate(bullets):
                col1, col2 = st.columns([4, 1])
                with col1:
                    new_bullet = st.text_area(
                        f"Bullet {i+1}",
                        value=bullet,
                        height=80,
                        key=f"bullet_{slide_key}_{i}"
                    )
                    if new_bullet != bullet:
                        st.session_state.editor_content[slide_key]['bullets'][i] = new_bullet
                        st.session_state.has_modifications = True
                
                with col2:
                    if st.button(f"🗑️", key=f"delete_bullet_{slide_key}_{i}"):
                        st.session_state.editor_content[slide_key]['bullets'].pop(i)
                        st.session_state.has_modifications = True
                        st.rerun()
            
            # Add new bullet point
            if st.button(f"➕ Add Bullet Point", key=f"add_bullet_{slide_key}"):
                st.session_state.editor_content[slide_key]['bullets'].append("New bullet point")
                st.session_state.has_modifications = True
                st.rerun()
    
    def _edit_paragraph_slide(self, slide_content: Dict[str, Any], slide_key: str):
        """Edit slides with paragraph content."""
        if 'paragraph' in slide_content:
            new_paragraph = st.text_area(
                "Content",
                value=slide_content['paragraph'],
                height=100,
                key=f"paragraph_{slide_key}"
            )
            if new_paragraph != slide_content['paragraph']:
                st.session_state.editor_content[slide_key]['paragraph'] = new_paragraph
                st.session_state.has_modifications = True
    
    def _edit_features_slide(self, slide_content: Dict[str, Any], slide_key: str):
        """Edit features slide content."""
        if 'features' in slide_content:
            st.markdown("**Features:**")
            features = slide_content['features'].copy()
            
            for i, feature in enumerate(features):
                col1, col2 = st.columns([4, 1])
                with col1:
                    new_feature = st.text_area(
                        f"Feature {i+1}",
                        value=feature,
                        height=80,
                        key=f"feature_{slide_key}_{i}"
                    )
                    if new_feature != feature:
                        st.session_state.editor_content[slide_key]['features'][i] = new_feature
                        st.session_state.has_modifications = True
                
                with col2:
                    if st.button(f"🗑️", key=f"delete_feature_{slide_key}_{i}"):
                        st.session_state.editor_content[slide_key]['features'].pop(i)
                        st.session_state.has_modifications = True
                        st.rerun()
            
            # Add new feature
            if st.button(f"➕ Add Feature", key=f"add_feature_{slide_key}"):
                st.session_state.editor_content[slide_key]['features'].append("New feature: Description")
                st.session_state.has_modifications = True
                st.rerun()
    
    def _edit_cta_slide(self, slide_content: Dict[str, Any], slide_key: str):
        """Edit call to action slide content."""
        if 'call_to_action' in slide_content:
            new_cta = st.text_input(
                "Call to Action",
                value=slide_content['call_to_action'],
                key=f"cta_{slide_key}"
            )
            if new_cta != slide_content['call_to_action']:
                st.session_state.editor_content[slide_key]['call_to_action'] = new_cta
                st.session_state.has_modifications = True
    
    def _render_image_upload(self, slide_key: str, slide_title: str):
        """Render image upload section for a slide."""
        st.markdown("**Custom Image:**")
        
        uploaded_file = st.file_uploader(
            f"Upload custom image for {slide_title}",
            type=['png', 'jpg', 'jpeg'],
            key=f"image_{slide_key}"
        )
        
        if uploaded_file is not None:
            # Store the uploaded image
            st.session_state.uploaded_images[slide_key] = uploaded_file.getvalue()
            st.session_state.has_modifications = True
            
            # Show preview
            image = Image.open(uploaded_file)
            st.image(image, caption="Uploaded Image Preview", width=200)
        
        # Show current custom image if exists
        elif slide_key in st.session_state.uploaded_images:
            image_data = st.session_state.uploaded_images[slide_key]
            image = Image.open(io.BytesIO(image_data))
            st.image(image, caption="Current Custom Image", width=200)
            
            if st.button(f"🗑️ Remove Custom Image", key=f"remove_image_{slide_key}"):
                del st.session_state.uploaded_images[slide_key]
                st.session_state.has_modifications = True
                st.rerun()
    
    def _render_download_buttons(self, original_content: Dict[str, Any]):
        """Render download buttons for original and modified presentations."""
        st.markdown("---")
        st.markdown("## 📥 Download Options")
        
        col1, col2 = st.columns(2)
        
        # Original presentation download
        with col1:
            st.markdown("### Original Presentation")
            st.markdown("Download the original generated presentation")
            
            from ppt_generator import create_presentation
            from utils import sanitize_filename
            
            try:
                original_filename = f"{sanitize_filename(original_content['metadata']['company_name'])}_{sanitize_filename(original_content['metadata']['product_name'])}_Overview.pptx"
                original_buffer = create_presentation(original_content, original_filename)
                
                st.download_button(
                    label="📥 Download Original",
                    data=original_buffer,
                    file_name=original_filename,
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    use_container_width=True,
                )
            except Exception as e:
                st.error(f"Error creating original presentation: {str(e)}")
        
        # Modified presentation download
        with col2:
            st.markdown("### Customized Presentation")
            if st.session_state.has_modifications:
                st.markdown("✅ **Changes detected** - Download your customized version")
                
                try:
                    modified_content = self._prepare_modified_content(original_content)
                    modified_filename = f"{sanitize_filename(original_content['metadata']['company_name'])}_{sanitize_filename(original_content['metadata']['product_name'])}_Custom.pptx"
                    
                    # Create modified presentation
                    from image_manager import ImageManager
                    image_manager = ImageManager(st.session_state.uploaded_images)
                    modified_buffer = self._create_modified_presentation(modified_content, modified_filename, image_manager)
                    
                    st.download_button(
                        label="📥 Download Customized",
                        data=modified_buffer,
                        file_name=modified_filename,
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                        use_container_width=True,
                    )
                except Exception as e:
                    st.error(f"Error creating customized presentation: {str(e)}")
            else:
                st.markdown("ℹ️ Make changes above to enable customized download")
                st.button(
                    "📥 Download Customized",
                    disabled=True,
                    use_container_width=True,
                )
    
    def _prepare_modified_content(self, original_content: Dict[str, Any]) -> Dict[str, Any]:
        """Prepare modified content based on user changes."""
        modified_content = copy.deepcopy(st.session_state.editor_content)
        
        # Remove deleted slides
        for slide_key in list(modified_content.keys()):
            if slide_key != 'metadata' and slide_key in st.session_state.deleted_slides:
                del modified_content[slide_key]
        
        # Preserve metadata
        modified_content['metadata'] = original_content['metadata']
        
        return modified_content
    
    def _create_modified_presentation(self, content: Dict[str, Any], filename: str, image_manager) -> io.BytesIO:
        """Create a modified presentation with custom images and content."""
        from ppt_generator_custom import create_custom_presentation
        return create_custom_presentation(content, filename, st.session_state.slide_order, image_manager)

    def get_slide_order(self) -> List[str]:
        """Get the current slide order."""
        return [key for key in st.session_state.slide_order if key not in st.session_state.deleted_slides]
