import streamlit as st
import os
import io
from datetime import datetime
from content_generator import generate_presentation_content
from ppt_generator import create_presentation
from utils import sanitize_filename
from preview_generator import SlidePreviewGenerator, simulate_slide_generation_with_preview
from slide_editor import SlideEditor

def load_demo_data():
    """Load demo data for presentation fields."""
    return {
        "company_name": "SecureLayer",
        "product_name": "ThreatGuard AI",
        "target_audience": "IT security teams at mid-sized enterprises and managed security service providers (MSSPs)",
        "problem_statement": "Mid-sized enterprises face a 3x higher risk of undetected cyber threats due to limited resources and fragmented security tools. Security teams spend 40% of their time on manual log analysis, leading to delayed responses and increased breach costs. Compliance requirements add complexity, with 62% of organizations struggling to keep up with evolving standards.",
        "key_features": "AI-Powered Threat Detection: Identify zero-day attacks and anomalies in real time\nAutomated Incident Response: Orchestrate and remediate threats with one click\nUnified Security Dashboard: Centralize alerts, logs, and compliance status\nContinuous Compliance Monitoring: Map controls to SOC2, ISO 27001, and GDPR\nSeamless Integration: Connect with 50+ security and IT tools\n24/7 Expert Support: Access cybersecurity experts anytime",
        "competitive_advantage": "ThreatGuard AI reduces mean time to detect (MTTD) by 80% and automates 90% of incident response actions. Unlike legacy SIEMs, our platform deploys in under 30 minutes with no coding required. Customers report a 60% reduction in compliance audit preparation time and 2x faster breach containment.",
        "call_to_action": "Book a free security assessment today"
    }

def render_ui():
    """Render the Streamlit UI for the presentation generator."""
    st.title("Deck Genie: B2B SaaS Presentation Generator")
    st.subheader("Create a minimalist, executive-ready product overview in seconds")
    
    # Initialize session state for demo toggle
    if "use_demo" not in st.session_state:
        st.session_state.use_demo = False
    
    # Initialize session state for presentation generation
    if "presentation_generated" not in st.session_state:
        st.session_state.presentation_generated = False
    
    if "original_content" not in st.session_state:
        st.session_state.original_content = None
    
    # Ensure deleted_slides is always properly initialized as a set of strings
    if 'deleted_slides' not in st.session_state:
        st.session_state.deleted_slides = set()
    elif not isinstance(st.session_state.deleted_slides, set):
        st.session_state.deleted_slides = set()
    else:
        # Additional cleanup - ensure all items are strings
        try:
            st.session_state.deleted_slides = {
                item for item in st.session_state.deleted_slides 
                if isinstance(item, str)
            }
        except (TypeError, AttributeError):
            st.session_state.deleted_slides = set()
    
    # Callback function to handle checkbox changes
    def toggle_demo():
        st.session_state.use_demo = not st.session_state.use_demo
    
    # Function to reset presentation state
    def reset_presentation():
        st.session_state.presentation_generated = False
        st.session_state.original_content = None
        # Clear editor state explicitly
        editor_keys = ['editor_content', 'slide_order', 'deleted_slides', 'has_modifications', 'uploaded_images']
        for key in editor_keys:
            if key in st.session_state:
                del st.session_state[key]
        # Clear any other editor-related keys
        for key in list(st.session_state.keys()):
            if key.startswith(('title_', 'bullet_', 'feature_', 'paragraph_', 'cta_', 'subtitle_', 'delete_', 'add_', 'image_', 'remove_')):
                del st.session_state[key]
    
    # Show reset button if presentation is generated
    if st.session_state.get('presentation_generated', False):
        col1, col2 = st.columns([3, 1])
        with col2:
            if st.button("🔄 Generate New Presentation", key="reset_btn"):
                reset_presentation()
                st.rerun()
    
    # Only show the form if no presentation is generated
    if not st.session_state.get('presentation_generated', False):
        # Checkbox outside the form to toggle demo data
        st.checkbox("Use Demo Data", value=st.session_state.use_demo, on_change=toggle_demo)
        
        with st.form("presentation_form"):
            col1, col2 = st.columns(2)
            
            with col1:
                demo_data = load_demo_data()
                
                if st.session_state.use_demo:
                    company_name = st.text_input("Company Name", value=demo_data["company_name"])
                    product_name = st.text_input("Product Name", value=demo_data["product_name"])
                    target_audience = st.text_input("Target Audience", value=demo_data["target_audience"])
                    problem_statement = st.text_area("Problem Statement (max 200 words)", 
                                                 value=demo_data["problem_statement"],
                                                 height=150)
                    key_features = st.text_area("Key Features (one per line, max 6)", 
                                            value=demo_data["key_features"],
                                            height=150)
                    competitive_advantage = st.text_area("Competitive Advantage (max 150 words)",
                                                     value=demo_data["competitive_advantage"],
                                                     height=150)
                    call_to_action = st.text_input("Call to Action", value=demo_data["call_to_action"])
                else:
                    company_name = st.text_input("Company Name")
                    product_name = st.text_input("Product Name")
                    target_audience = st.text_input("Target Audience")
                    problem_statement = st.text_area("Problem Statement (max 200 words)", height=150)
                    key_features = st.text_area("Key Features (one per line, max 6)", height=150)
                    competitive_advantage = st.text_area("Competitive Advantage (max 150 words)", height=150)
                    call_to_action = st.text_input("Call to Action")
            
            with col2:
                st.markdown("### Instructions")
                st.markdown("""
                1. Fill out all fields with your product information
                2. Click 'Generate Presentation' to create your deck
                3. Download the finished PowerPoint file
                
                **Tips for best results:**
                - Be specific in your problem statement
                - Limit key features to the most important 4-6 items
                - Highlight measurable advantages over competitors
                - Keep your call to action simple and direct
                """)
                
                st.markdown("### Preview")
                st.markdown("Your presentation will include these slides:")
                st.markdown("1. Title Slide")
                st.markdown("2. Problem Statement")
                st.markdown("3. Solution Overview")
                st.markdown("4. Key Features")
                st.markdown("5. Competitive Advantage") 
                st.markdown("6. Target Audience")
                st.markdown("7. Call to Action")
            
            submitted = st.form_submit_button("Generate Presentation")
        
        if submitted:
            # Validate inputs
            errors = []
            if not company_name or not product_name:
                errors.append("Company and product names are required")
            
            if not target_audience:
                errors.append("Target audience is required")
                
            if not problem_statement:
                errors.append("Problem statement is required")
            elif len(problem_statement.split()) > 200:
                errors.append("Problem statement exceeds 200 words")
                
            if not key_features:
                errors.append("Key features are required")
            else:
                features_list = [f for f in key_features.strip().split('\n') if f]
                if len(features_list) > 6:
                    errors.append("Maximum of 6 key features allowed")
                    
            if not competitive_advantage:
                errors.append("Competitive advantage is required")
            elif len(competitive_advantage.split()) > 150:
                errors.append("Competitive advantage exceeds 150 words")
                
            if not call_to_action:
                errors.append("Call to action is required")
                
            if errors:
                for error in errors:
                    st.error(error)
            else:
                try:
                    # Initialize preview generator and slide editor
                    preview_generator = SlidePreviewGenerator()
                    slide_editor = SlideEditor()
                    
                    # Set up a modern progress tracking system
                    progress_container = st.container()
                    with progress_container:
                        # Create a progress bar
                        progress_bar = st.progress(0)
                        progress_status = st.empty()
                        progress_detail = st.empty()
                        
                        # Step 1: Initialize
                        progress_status.markdown("### 🚀 Starting the creation process...")
                        progress_detail.markdown("_Analyzing your inputs and preparing the magic_")
                        progress_bar.progress(10)
                        st.balloons()  # Add a fun element at the start
                        
                        # Create live preview container
                        preview_container_ui, slide_placeholders = preview_generator.create_preview_container()
                        
                        # Process input data
                        features_list = [f for f in key_features.strip().split('\n') if f]
                        
                        # Step 2: Content generation with live preview
                        progress_status.markdown("### 🧠 Crafting your presentation content...")
                        progress_detail.markdown("_Our AI is creating compelling narratives for your slides_")
                        
                        # Generate presentation content with live preview updates
                        content = simulate_slide_generation_with_preview(
                            generate_presentation_content,
                            (company_name, product_name, target_audience, problem_statement, 
                             features_list, competitive_advantage, call_to_action),
                            progress_bar,
                            progress_status,
                            progress_detail,
                            preview_generator,
                            slide_placeholders
                        )
                        
                        # Step 3: Designing slides
                        progress_status.markdown("### 🎨 Designing your slides...")
                        progress_detail.markdown("_Creating a visually appealing presentation that stands out_")
                        progress_bar.progress(70)
                        
                        # Step 4: Finalizing
                        progress_status.markdown("### 📊 Assembling your presentation...")
                        progress_detail.markdown("_Putting everything together in a cohesive package_")
                        progress_bar.progress(90)
                        
                        # Create the PowerPoint presentation
                        filename = f"{sanitize_filename(company_name)}_{sanitize_filename(product_name)}_Overview.pptx"
                        presentation_buffer = create_presentation(content, filename)
                        
                        # Final step
                        progress_status.markdown("### ✨ Polishing final details...")
                        progress_detail.markdown("_Adding those final touches of excellence_")
                        progress_bar.progress(100)
                        
                        # Clear the progress elements
                        progress_container.empty()
                        
                        # Show success message
                        st.success("✅ Your presentation is ready!")
                        
                        # Store the original content in session state for editing
                        st.session_state.original_content = content
                        st.session_state.presentation_generated = True
                    
                except Exception as e:
                    st.error(f"An error occurred: {str(e)}")
                    st.exception(e)
    
    # Render the slide editor after generation (outside the form)
    if st.session_state.get('presentation_generated', False) and st.session_state.get('original_content'):
        slide_editor = SlideEditor()
        slide_editor.render_slide_editor(st.session_state.original_content)