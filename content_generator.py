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
        
        # Validate inputs to ensure quality content generation
        if not company_name or not product_name:
            raise ValueError("Company name and product name are required")
        
        # Sanitize and standardize persona value
        persona = persona.strip().title() if persona else "Generic"
        if persona not in ["Generic", "Technical", "Marketing", "Executive", "Investor"]:
            persona = "Generic"
        
        print(f"Generating content using {persona} persona with focused targeting")
        
        # Generate required slides with enhanced error handling
        try:
            title_content = generate_title_slide_content(product_name, company_name)
            if not validate_slide_content(title_content, "title_slide"):
                title_content = create_fallback_title_slide(product_name, company_name)
        except Exception as e:
            print(f"Error generating title slide: {str(e)}")
            title_content = create_fallback_title_slide(product_name, company_name)
            
        try:
            problem_content = generate_problem_slide_content(problem_statement, persona)
            if not validate_slide_content(problem_content, "problem_slide"):
                problem_content = create_fallback_problem_slide(problem_statement)
        except Exception as e:
            print(f"Error generating problem slide: {str(e)}")
            problem_content = create_fallback_problem_slide(problem_statement)
        
        try:
            solution_content = generate_solution_slide_content(product_name, problem_statement, persona)
            if not validate_slide_content(solution_content, "solution_slide"):
                solution_content = create_fallback_solution_slide(product_name, problem_statement)
        except Exception as e:
            print(f"Error generating solution slide: {str(e)}")
            solution_content = create_fallback_solution_slide(product_name, problem_statement)
        
        try:
            features_content = generate_features_slide_content(key_features, persona)
            if not validate_slide_content(features_content, "features_slide"):
                features_content = create_fallback_features_slide(key_features)
        except Exception as e:
            print(f"Error generating features slide: {str(e)}")
            features_content = create_fallback_features_slide(key_features)
        
        try:
            advantage_content = generate_advantage_slide_content(competitive_advantage, persona)
            if not validate_slide_content(advantage_content, "advantage_slide"):
                advantage_content = create_fallback_advantage_slide(competitive_advantage)
        except Exception as e:
            print(f"Error generating advantage slide: {str(e)}")
            advantage_content = create_fallback_advantage_slide(competitive_advantage)
        
        try:
            audience_content = generate_audience_slide_content(target_audience, persona)
            if not validate_slide_content(audience_content, "audience_slide"):
                audience_content = create_fallback_audience_slide(target_audience)
        except Exception as e:
            print(f"Error generating audience slide: {str(e)}")
            audience_content = create_fallback_audience_slide(target_audience)
        
        try:
            cta_content = generate_cta_slide_content(call_to_action, product_name, persona)
            if not validate_slide_content(cta_content, "cta_slide"):
                cta_content = create_fallback_cta_slide(call_to_action, product_name)
        except Exception as e:
            print(f"Error generating CTA slide: {str(e)}")
            cta_content = create_fallback_cta_slide(call_to_action, product_name)
        
        # Define all basic slides
        all_slides = {
            'title_slide': title_content,
            'problem_slide': problem_content,
            'solution_slide': solution_content,
            'features_slide': features_content,
            'advantage_slide': advantage_content,
            'audience_slide': audience_content,
            'cta_slide': cta_content
        }
        
        # Generate additional slides with strong validation and fallbacks
        candidate_additional_slides = {}
        
        # Immediate execution with error handling for additional slides
        try:
            market_content = generate_market_slide_content(target_audience, persona)
            if validate_slide_content(market_content, "market_slide"):
                candidate_additional_slides['market_slide'] = market_content
            else:
                candidate_additional_slides['market_slide'] = create_fallback_market_slide(target_audience)
        except Exception as e:
            print(f"Error generating market slide: {str(e)}")
            candidate_additional_slides['market_slide'] = create_fallback_market_slide(target_audience)
            
        try:
            roadmap_content = generate_roadmap_slide_content(product_name, persona)
            if validate_slide_content(roadmap_content, "roadmap_slide"):
                candidate_additional_slides['roadmap_slide'] = roadmap_content
            else:
                candidate_additional_slides['roadmap_slide'] = create_fallback_roadmap_slide(product_name)
        except Exception as e:
            print(f"Error generating roadmap slide: {str(e)}")
            candidate_additional_slides['roadmap_slide'] = create_fallback_roadmap_slide(product_name)
            
        try:
            team_content = generate_team_slide_content(company_name, persona)
            if validate_slide_content(team_content, "team_slide"):
                candidate_additional_slides['team_slide'] = team_content
            else:
                candidate_additional_slides['team_slide'] = create_fallback_team_slide(company_name)
        except Exception as e:
            print(f"Error generating team slide: {str(e)}")
            candidate_additional_slides['team_slide'] = create_fallback_team_slide(company_name)

        # Select slides based on slide_count and persona
        selected_slides = select_slides_for_presentation(all_slides, candidate_additional_slides, slide_count, persona)
        
        # Apply persona-specific enhancements to the selected content
        selected_slides = enhance_content_for_persona(selected_slides, persona)
        
        # Add metadata with persona focus indicator
        selected_slides['metadata'] = {
            'company_name': company_name,
            'product_name': product_name,
            'persona': persona,
            'slide_count': len(selected_slides) - 1,  # Exclude metadata from count
            'persona_focused': persona != "Generic"
        }
        
        return selected_slides
    except Exception as e:
        raise Exception(f"Error generating presentation content: {str(e)}")

def validate_slide_content(content: Dict, slide_type: str) -> bool:
    """Validate that slide content meets quality standards."""
    if not content:
        return False
    
    if slide_type == "title_slide":
        return "title" in content and "subtitle" in content and len(content["title"]) > 3
    
    elif slide_type == "problem_slide":
        return "title" in content and ("bullets" in content or "pain_points" in content)
    
    elif slide_type == "solution_slide":
        return "title" in content and ("paragraph" in content or "description" in content)
    
    elif slide_type == "features_slide":
        return "title" in content and "features" in content and len(content.get("features", [])) > 0
    
    elif slide_type == "advantage_slide":
        return "title" in content and ("differentiators" in content or "bullets" in content)
    
    elif slide_type == "audience_slide":
        return "title" in content and ("paragraph" in content or "description" in content or "content" in content)
    
    elif slide_type == "cta_slide":
        return "title" in content and ("cta_text" in content or "call_to_action" in content)
    
    elif slide_type == "market_slide":
        return "title" in content and ("market_size" in content or "description" in content)
    
    elif slide_type == "roadmap_slide":
        return "title" in content and ("phases" in content or "milestones" in content)
    
    elif slide_type == "team_slide":
        return "title" in content
    
    return True

# Fallback content creators for each slide type
def create_fallback_title_slide(product_name: str, company_name: str) -> Dict[str, Any]:
    return {
        "title": f"{product_name}",
        "subtitle": f"Presented by {company_name}",
        "product_name": product_name
    }

def create_fallback_problem_slide(problem_statement: str) -> Dict[str, Any]:
    lines = problem_statement.split('. ')
    return {
        "title": "The Problem",
        "pain_points": [line.strip() + "." for line in lines[:3] if line.strip()]
    }

def create_fallback_solution_slide(product_name: str, problem_statement: str) -> Dict[str, Any]:
    return {
        "title": f"Introducing {product_name}",
        "paragraph": f"{product_name} addresses these challenges with a comprehensive solution designed specifically for your needs."
    }

def create_fallback_features_slide(key_features: List[str]) -> Dict[str, Any]:
    return {
        "title": "Key Features",
        "features": key_features[:5]  # Limit to 5 features
    }

def create_fallback_advantage_slide(competitive_advantage: str) -> Dict[str, Any]:
    lines = competitive_advantage.split('. ')
    return {
        "title": "Our Competitive Advantage",
        "differentiators": [line.strip() + "." for line in lines[:3] if line.strip()]
    }

def create_fallback_audience_slide(target_audience: str) -> Dict[str, Any]:
    return {
        "title": "Target Audience",
        "paragraph": target_audience
    }

def create_fallback_cta_slide(call_to_action: str, product_name: str) -> Dict[str, Any]:
    return {
        "title": "Take the Next Step",
        "cta_text": call_to_action or f"Contact us today to learn more about {product_name}"
    }

def create_fallback_market_slide(target_audience: str) -> Dict[str, Any]:
    return {
        "title": "Market Opportunity",
        "market_size": "$25B by 2025",
        "growth_rate": "15% CAGR",
        "description": "The market for this solution is growing rapidly as more organizations face similar challenges."
    }

def create_fallback_roadmap_slide(product_name: str) -> Dict[str, Any]:
    return {
        "title": f"{product_name}: Product Roadmap",
        "phases": [
            {"name": "Phase 1", "items": ["Initial launch", "Core features", "First customers"]},
            {"name": "Phase 2", "items": ["Advanced analytics", "Integration APIs", "Expanded support"]},
            {"name": "Phase 3", "items": ["Enterprise features", "Global expansion", "Industry partnerships"]}
        ]
    }

def create_fallback_team_slide(company_name: str) -> Dict[str, Any]:
    return {
        "title": f"{company_name} Leadership Team",
        "team_members": [
            {"name": "Leadership Team", "role": "Our experienced team brings decades of industry expertise"}
        ],
        "tagline": "Building innovative solutions for tomorrow's challenges"
    }

def select_slides_for_presentation(all_slides: Dict, candidate_additional_slides: Dict, slide_count: int, persona: str = "Generic") -> Dict[str, Any]:
    """Select appropriate slides based on user-specified slide count and persona with focused targeting."""
    
    # Enhanced persona-specific slide priorities with focused targeting
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
            'title_slide',     # Always first - Technical product introduction
            'problem_slide',   # Technical pain points and challenges
            'solution_slide',  # Technical architecture and approach
            'features_slide',  # Detailed technical specifications
            'audience_slide',  # Technical user profiles and use cases
            'roadmap_slide',   # Technical development roadmap
            'advantage_slide', # Technical superiority and innovations
            'team_slide',      # Technical team expertise
            'market_slide',    # Technical market landscape
            'cta_slide'        # Always last - Technical next steps
        ],
        "Marketing": [
            'title_slide',     # Always first - Brand-focused introduction
            'audience_slide',  # Target customer segments and personas
            'problem_slide',   # Customer pain points and market gaps
            'solution_slide',  # Value proposition and customer benefits
            'advantage_slide', # Competitive differentiation and USPs
            'features_slide',  # Customer-facing features and benefits
            'market_slide',    # Market opportunity and customer demand
            'roadmap_slide',   # Go-to-market strategy and milestones
            'team_slide',      # Marketing and customer success team
            'cta_slide'        # Always last - Customer acquisition focused
        ],
        "Executive": [
            'title_slide',     # Always first - Strategic overview
            'market_slide',    # Business opportunity and market size
            'problem_slide',   # Business challenges and market needs
            'solution_slide',  # Strategic solution and business impact
            'advantage_slide', # Competitive positioning and moats
            'roadmap_slide',   # Strategic milestones and business goals
            'team_slide',      # Leadership and execution capability
            'features_slide',  # Key capabilities that drive business value
            'audience_slide',  # Target market and customer strategy
            'cta_slide'        # Always last - Strategic next steps
        ],
        "Investor": [
            'title_slide',     # Always first - Investment opportunity
            'market_slide',    # Total addressable market and opportunity
            'problem_slide',   # Market inefficiencies and unmet needs
            'solution_slide',  # Unique solution and scalability
            'team_slide',      # Management team and execution track record
            'advantage_slide', # Competitive moats and barriers to entry
            'roadmap_slide',   # Growth strategy and scaling milestones
            'audience_slide',  # Customer acquisition and retention
            'features_slide',  # Product differentiation and IP
            'cta_slide'        # Always last - Investment ask and terms
        ]
    }
    
    # Use Generic priority if persona not found
    selected_priorities = slide_priorities.get(persona, slide_priorities["Generic"])
    
    # For focused personas, ensure we generate ALL relevant slides regardless of slide_count
    # This allows users to see the complete persona-focused presentation
    if persona in ["Technical", "Marketing", "Executive", "Investor"]:
        # Generate all slides that are relevant to this persona
        selected = {}
        
        # Add all slides in persona-specific priority order
        for slide_key in selected_priorities:
            if slide_key in all_slides:
                selected[slide_key] = all_slides[slide_key]
            elif slide_key in candidate_additional_slides:
                selected[slide_key] = candidate_additional_slides[slide_key]
    else:
        # For Generic persona, use the original slide_count logic
        slide_count = max(5, min(slide_count, 10))
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
        
        # If we still need more slides, use the candidate_additional_slides
        if remaining_slots > 0:
            for slide_key in selected_priorities:
                if remaining_slots <= 0:
                    break
                    
                if slide_key not in selected and slide_key in candidate_additional_slides:
                    selected[slide_key] = candidate_additional_slides[slide_key]
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
    """Get enhanced persona-specific context for prompts with focused targeting."""
    contexts = {
        'Technical': '''technical audience including developers, engineers, architects, and IT professionals who need:
        - Detailed technical specifications and architecture insights
        - Implementation details, APIs, and integration capabilities
        - Performance metrics, scalability, and security considerations
        - Technical roadmap with development milestones
        - Code examples, technical documentation, and developer resources''',
        
        'Marketing': '''marketing professionals, brand managers, and customer-facing teams who need:
        - Customer-centric value propositions and benefit statements
        - Market positioning, competitive differentiation, and USPs
        - Customer personas, segments, and target audience insights
        - Go-to-market strategy, campaign messaging, and brand story
        - Customer success stories, testimonials, and market validation''',
        
        'Executive': '''C-suite executives, senior management, and strategic decision-makers who need:
        - High-level business impact, ROI, and strategic value
        - Market opportunity, competitive landscape, and business model
        - Strategic roadmap, business objectives, and key milestones
        - Risk assessment, resource requirements, and investment rationale
        - Leadership team capabilities and execution track record''',
        
        'Investor': '''investors, VCs, angels, and financial stakeholders who need:
        - Investment opportunity, market size, and growth potential
        - Business model, revenue streams, and path to profitability
        - Competitive advantages, barriers to entry, and market positioning
        - Financial projections, funding requirements, and use of capital
        - Management team experience and execution capabilities''',
        
        'Generic': '''mixed audience requiring balanced technical and business content with:
        - Clear value proposition that appeals to multiple stakeholders
        - Technical credibility balanced with business benefits
        - Comprehensive overview suitable for diverse decision-makers
        - Balanced detail level appropriate for general business audiences
        - Broad appeal covering technical, business, and strategic aspects'''
    }
    return contexts.get(persona, contexts['Generic'])

def enhance_content_for_persona(content: Dict[str, Any], persona: str) -> Dict[str, Any]:
    """
    Enhance and tailor content specifically for the target persona.
    This ensures each slide type is optimized for the persona's specific needs.
    """
    enhanced_content = content.copy()
    
    if persona == "Marketing":
        # Marketing-focused enhancements
        if 'audience_slide' in enhanced_content:
            # Add marketing-specific audience insights
            audience_content = enhanced_content['audience_slide']
            if 'paragraph' in audience_content:
                marketing_angle = " Our target customers are actively seeking solutions that deliver measurable ROI and competitive advantage in their market."
                audience_content['paragraph'] = audience_content['paragraph'] + marketing_angle
        
        if 'advantage_slide' in enhanced_content:
            # Emphasize competitive differentiation for marketing
            advantage_content = enhanced_content['advantage_slide']
            if 'title' not in advantage_content or 'competitive' not in advantage_content['title'].lower():
                advantage_content['title'] = "Competitive Differentiation & Unique Value"
        
        if 'cta_slide' in enhanced_content:
            # Marketing-focused CTA
            cta_content = enhanced_content['cta_slide']
            cta_content['title'] = "Ready to Accelerate Your Growth?"
            if 'bullets' not in cta_content:
                cta_content['bullets'] = []
            cta_content['bullets'] = [
                "Schedule a demo to see customer impact",
                "Get a customized ROI analysis for your market"
            ]
    
    elif persona == "Technical":
        # Technical-focused enhancements
        if 'features_slide' in enhanced_content:
            # Add technical depth to features
            features_content = enhanced_content['features_slide']
            features_content['title'] = "Technical Capabilities & Architecture"
        
        if 'solution_slide' in enhanced_content:
            # Technical solution focus
            solution_content = enhanced_content['solution_slide']
            solution_content['title'] = "Technical Architecture & Implementation"
        
        if 'cta_slide' in enhanced_content:
            # Technical-focused CTA
            cta_content = enhanced_content['cta_slide']
            cta_content['title'] = "Ready for Technical Deep Dive?"
            if 'bullets' not in cta_content:
                cta_content['bullets'] = []
            cta_content['bullets'] = [
                "Access technical documentation and APIs",
                "Schedule architecture review session"
            ]
    
    elif persona == "Executive":
        # Executive-focused enhancements
        if 'market_slide' in enhanced_content:
            # Executive market focus
            market_content = enhanced_content['market_slide']
            market_content['title'] = "Strategic Market Opportunity"
        
        if 'roadmap_slide' in enhanced_content:
            # Strategic roadmap focus
            roadmap_content = enhanced_content['roadmap_slide']
            roadmap_content['title'] = "Strategic Roadmap & Business Milestones"
        
        if 'cta_slide' in enhanced_content:
            # Executive-focused CTA
            cta_content = enhanced_content['cta_slide']
            cta_content['title'] = "Strategic Partnership Opportunity"
            if 'bullets' not in cta_content:
                cta_content['bullets'] = []
            cta_content['bullets'] = [
                "Executive briefing and strategic alignment",
                "ROI analysis and implementation roadmap"
            ]
    
    elif persona == "Investor":
        # Investor-focused enhancements
        if 'market_slide' in enhanced_content:
            # Investment opportunity focus
            market_content = enhanced_content['market_slide']
            market_content['title'] = "Investment Opportunity & Market Size"
        
        if 'team_slide' in enhanced_content:
            # Leadership and execution focus
            team_content = enhanced_content['team_slide']
            team_content['title'] = "Leadership Team & Execution Track Record"
        
        if 'cta_slide' in enhanced_content:
            # Investment-focused CTA
            cta_content = enhanced_content['cta_slide']
            cta_content['title'] = "Investment Opportunity"
            if 'bullets' not in cta_content:
                cta_content['bullets'] = []
            cta_content['bullets'] = [
                "Review detailed financial projections",
                "Discuss investment terms and partnership"
            ]
    
    return enhanced_content
