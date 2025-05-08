import streamlit as st
import os
import io
from datetime import datetime
from content_generator import generate_presentation_content
from ppt_generator import create_presentation
from utils import sanitize_filename

def load_demo_data():
    """Load demo data for presentation fields."""
    return {
        "company_name": "CloudFlow Solutions",
        "product_name": "DataSync Pro",
        "target_audience": "Mid-size enterprise IT departments and data analysts",
        "problem_statement": "Organizations struggle with fragmented data across multiple systems, leading to inconsistent reporting, duplicate efforts, and missed insights. IT teams spend 30% of their time on manual data transfers, while analysts waste 15 hours weekly reconciling data discrepancies. Decision-makers often work with outdated information, causing costly mistakes and missed opportunities.",
        "key_features": "Real-time Data Integration: Connect any data source in minutes\nAutomated Workflows: Schedule and automate data transfers\nData Validation Engine: Ensure accuracy with 99.9% precision\nInsight Dashboard: Visualize KPIs and metrics instantly\nCollaborative Interface: Share insights across departments\nEnterprise-grade Security: SOC2 and GDPR compliant",
        "competitive_advantage": "Unlike competitors who require extensive setup time and coding knowledge, DataSync Pro offers a codeless interface that reduces implementation time by 70%. Our proprietary validation algorithms deliver 3x higher accuracy rates than the industry standard, and our tiered pricing model ensures companies only pay for what they use - averaging 40% cost savings compared to fixed-license competitors.",
        "call_to_action": "Schedule your personalized demo today"
    }

def render_ui():
    """Render the Streamlit UI for the presentation generator."""
    st.title("Deck Genie: B2B SaaS Presentation Generator")
    st.subheader("Create a minimalist, executive-ready product overview in seconds")
    
    with st.form("presentation_form"):
        col1, col2 = st.columns(2)
        
        with col1:
            use_demo = st.checkbox("Use Demo Data", value=False)
            
            if use_demo:
                demo_data = load_demo_data()
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
            with st.spinner("Generating your presentation..."):
                try:
                    # Process input data
                    features_list = [f for f in key_features.strip().split('\n') if f]
                    
                    # Generate presentation content using Claude
                    content = generate_presentation_content(
                        company_name=company_name,
                        product_name=product_name,
                        target_audience=target_audience,
                        problem_statement=problem_statement,
                        key_features=features_list,
                        competitive_advantage=competitive_advantage,
                        call_to_action=call_to_action
                    )
                    
                    # Create the PowerPoint presentation
                    filename = f"{sanitize_filename(company_name)}_{sanitize_filename(product_name)}_Overview.pptx"
                    
                    presentation_buffer = create_presentation(content, filename)
                    
                    # Provide download button
                    st.success("✅ Your presentation is ready!")
                    st.download_button(
                        label="Download Presentation",
                        data=presentation_buffer,
                        file_name=filename,
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                    )
                except Exception as e:
                    st.error(f"An error occurred: {str(e)}")
                    st.exception(e)