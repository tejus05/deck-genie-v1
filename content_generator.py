import os
import json
import google.generativeai as genai
from typing import Dict, List, Any
from utils import get_api_key, truncate_text
from slide_content_generators import (
    generate_title_slide_content, generate_problem_slide_content,
    generate_solution_slide_content, generate_features_slide_content,
    generate_advantage_slide_content, generate_audience_slide_content,
    generate_cta_slide_content, generate_market_slide_content,
    generate_roadmap_slide_content, generate_team_slide_content
)

def generate_presentation_content(
    company_name: str,
    product_name: str,
    target_audience: str,
    problem_statement: str,
    key_features: List[str],
    competitive_advantage: str,
    call_to_action: str,
    persona: str = "Generic",
    slide_count: int = 7
) -> Dict[str, Any]:
    """
    Generate content for all slides using Gemini API with persona customization.
    
    Args:
        company_name: The company name
        product_name: The product name
        target_audience: The target audience description
        problem_statement: The problem statement text
        key_features: List of key features
        competitive_advantage: The competitive advantage text
        call_to_action: The call to action text
        persona: Target persona ("Generic", "Technical", "Marketing", "Executive", "Investor")
        slide_count: Number of slides to generate (5-10)
        
    Returns:
        Dictionary containing structured content for all slides
    """
    try:
        # Configure the Gemini API
        genai.configure(api_key=get_api_key("GEMINI_API_KEY"))
        
        # Generate content for each slide with persona context
        title_content = generate_title_slide_content(product_name, company_name)
        problem_content = generate_problem_slide_content(problem_statement, persona)
        solution_content = generate_solution_slide_content(product_name, problem_statement, persona)
        features_content = generate_features_slide_content(key_features, persona)
        advantage_content = generate_advantage_slide_content(competitive_advantage, persona)
        audience_content = generate_audience_slide_content(target_audience, persona)
        cta_content = generate_cta_slide_content(call_to_action, product_name, persona)
        
        # Define all possible slides
        all_slides = {
            'title_slide': title_content,
            'problem_slide': problem_content,
            'solution_slide': solution_content,
            'features_slide': features_content,
            'advantage_slide': advantage_content,
            'audience_slide': audience_content,
            'cta_slide': cta_content
        }
        
        # Generate additional slides if needed for higher slide counts
        # These are candidates for inclusion based on slide_count
        candidate_additional_slides = {}
        
        # We'll generate these using lambda functions to defer execution until needed
        candidate_additional_slides['market_slide'] = lambda: generate_market_slide_content(target_audience, persona)
        candidate_additional_slides['roadmap_slide'] = lambda: generate_roadmap_slide_content(product_name, persona)
        candidate_additional_slides['team_slide'] = lambda: generate_team_slide_content(company_name, persona)

        # Select slides based on slide_count and persona
        selected_slides = select_slides_for_presentation(all_slides, candidate_additional_slides, slide_count, persona)
        
        # Add metadata
        selected_slides['metadata'] = {
            'company_name': company_name,
            'product_name': product_name,
            'persona': persona,
            'slide_count': slide_count
        }        
        return selected_slides
    except Exception as e:
        raise Exception(f"Error generating presentation content: {str(e)}")

def select_slides_for_presentation(all_slides: Dict, candidate_additional_slides: Dict, slide_count: int, persona: str = "Generic") -> Dict[str, Any]:
    """Select appropriate slides based on user-specified slide count and persona."""
    # Set the minimum number of slides (title and CTA are required)
    min_slides = 2
    
    # Set a slide priority order based on persona
    slide_priorities = {
        "Generic": [
            'title_slide',     # Always first
            'problem_slide',   
            'solution_slide',  
            'features_slide',  
            'advantage_slide', 
            'audience_slide',  
            'market_slide',    
            'roadmap_slide',   
            'team_slide',      
            'cta_slide'        # Always last
        ],
        "Technical": [
            'title_slide',     # Always first
            'problem_slide',   
            'solution_slide',  
            'features_slide',  
            'audience_slide',  
            'advantage_slide', 
            'roadmap_slide',   
            'team_slide',      
            'market_slide',    
            'cta_slide'        # Always last
        ],
        "Marketing": [
            'title_slide',     # Always first
            'solution_slide',  
            'problem_slide',   
            'advantage_slide', 
            'audience_slide',  
            'features_slide',  
            'market_slide',    
            'team_slide',      
            'roadmap_slide',   
            'cta_slide'        # Always last
        ],
        "Executive": [
            'title_slide',     # Always first
            'market_slide',    
            'problem_slide',   
            'solution_slide',  
            'advantage_slide', 
            'features_slide',  
            'roadmap_slide',   
            'audience_slide',  
            'team_slide',      
            'cta_slide'        # Always last
        ],
        "Investor": [
            'title_slide',     # Always first
            'market_slide',    
            'problem_slide',   
            'solution_slide',  
            'team_slide',      
            'advantage_slide', 
            'roadmap_slide',   
            'features_slide',  
            'audience_slide',  
            'cta_slide'        # Always last
        ]
    }
    
    # Use Generic priority if persona not found
    selected_priorities = slide_priorities.get(persona, slide_priorities["Generic"])
    
    # Ensure slide_count is within bounds (min 5, max 10)
    slide_count = max(5, min(slide_count, 10))
    
    # Start with empty selection
    selected = {}
    
    # Always include title_slide and cta_slide
    if 'title_slide' in all_slides:
        selected['title_slide'] = all_slides['title_slide']
    if 'cta_slide' in all_slides:
        selected['cta_slide'] = all_slides['cta_slide']
    
    # Fill in the rest based on priority until we hit the slide count
    remaining_slots = slide_count - len(selected)
    
    # First add slides from the all_slides dictionary based on persona priority
    for slide_key in selected_priorities:
        if slide_key in ['title_slide', 'cta_slide']:
            continue  # Already added
            
        if remaining_slots <= 0:
            break
            
        if slide_key in all_slides:
            selected[slide_key] = all_slides[slide_key]
            remaining_slots -= 1
    
    # If we still need more slides, use the candidate_additional_slides based on persona priority
    if remaining_slots > 0:
        for slide_key in selected_priorities:
            if remaining_slots <= 0:
                break
                
            if slide_key not in selected and slide_key in candidate_additional_slides:
                # Generate the content using the lambda function
                selected[slide_key] = candidate_additional_slides[slide_key]()
                remaining_slots -= 1
    
    return selected

def get_domain_context(domain: str) -> str:
    """Get domain-specific context for prompts."""
    contexts = {
        'investor': 'investor pitch focusing on market opportunity, growth potential, competitive advantage, and ROI',
        'marketing': 'marketing presentation emphasizing customer benefits, market positioning, and compelling value propositions',
        'generic': 'general business presentation with balanced technical and business content'
    }
    return contexts.get(domain, contexts['generic'])

def get_persona_context(persona: str) -> str:
    """Get persona-specific context for prompts."""
    contexts = {
        'Technical': 'technical audience who appreciates detailed specifications, architecture insights, and implementation details',
        'Marketing': 'business audience focused on outcomes, benefits, market impact, and customer success stories',
        'Generic': 'mixed audience requiring both technical credibility and business value focus',
        'Executive': 'executive audience focused on business impact, ROI, strategic alignment, and high-level benefits',
        'Investor': 'investor audience interested in market opportunity, growth potential, competitive differentiation, and financial metrics'
    }
    return contexts.get(persona, contexts['Generic'])
