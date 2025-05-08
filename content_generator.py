import os
import json
import anthropic
from typing import Dict, List, Any
from utils import get_api_key, truncate_text

def generate_presentation_content(
    company_name: str,
    product_name: str,
    target_audience: str,
    problem_statement: str,
    key_features: List[str],
    competitive_advantage: str,
    call_to_action: str
) -> Dict[str, Any]:
    """
    Generate content for all slides using Claude API.
    
    Args:
        company_name: The company name
        product_name: The product name
        target_audience: The target audience description
        problem_statement: The problem statement text
        key_features: List of key features
        competitive_advantage: The competitive advantage text
        call_to_action: The call to action text
        
    Returns:
        Dictionary containing structured content for all slides
    """
    try:
        # Initialize the Anthropic client
        client = anthropic.Anthropic(api_key=get_api_key("ANTHROPIC_API_KEY"))
        
        # Generate content for each slide
        title_content = generate_title_slide_content(product_name, company_name)
        problem_content = generate_problem_slide_content(client, problem_statement)
        solution_content = generate_solution_slide_content(client, product_name, problem_statement)
        features_content = generate_features_slide_content(client, key_features)
        advantage_content = generate_advantage_slide_content(client, competitive_advantage)
        audience_content = generate_audience_slide_content(client, target_audience)
        cta_content = {'call_to_action': call_to_action}
        
        # Assemble all content
        presentation_content = {
            'title_slide': title_content,
            'problem_slide': problem_content,
            'solution_slide': solution_content,
            'features_slide': features_content,
            'advantage_slide': advantage_content,
            'audience_slide': audience_content,
            'cta_slide': cta_content,
            'metadata': {
                'company_name': company_name,
                'product_name': product_name
            }
        }
        
        return presentation_content
    except Exception as e:
        raise Exception(f"Error generating presentation content: {str(e)}")

def generate_title_slide_content(product_name: str, company_name: str) -> Dict[str, str]:
    """Generate content for the title slide."""
    return {
        'title': product_name,
        'subtitle': f"by {company_name}"
    }

def generate_problem_slide_content(client: anthropic.Anthropic, problem_statement: str) -> Dict[str, Any]:
    """Generate content for the problem statement slide."""
    prompt = f"""
    Generate 4-5 concise, detailed bullets for a 'Problem Statement' slide in a B2B SaaS product presentation.
    Each bullet should be executive-ready, data-driven, and highlight a specific pain point from this problem statement:
    
    "{problem_statement}"
    
    Guidelines:
    - Each bullet must be 1-2 sentences, maximum 20 words
    - Use outcome-oriented language (costs, time wasted, inefficiencies)
    - Include specific metrics or percentages where possible
    - Focus on organizational impact, not just individual pain
    - Use Title Case for the slide title
    
    Format your response as a JSON object with:
    1. A "title" key containing a compelling slide title (max 10 words)
    2. A "bullets" key containing an array of bullet points
    
    Example format:
    {{
      "title": "Industry Pain Points",
      "bullets": [
        "67% of managers waste 5+ hours weekly on manual reporting.",
        "Data silos lead to 23% duplication of work across teams."
      ]
    }}
    """
    
    response = client.messages.create(
        model="claude-3-opus-20240229",
        max_tokens=500,
        temperature=0.3,
        system="You are an expert B2B SaaS copywriter who creates executive-ready, minimalist presentation content.",
        messages=[
            {"role": "user", "content": prompt}
        ]
    )
    
    # Parse the JSON response
    try:
        content_start = response.content[0].text.find('{')
        content_end = response.content[0].text.rfind('}') + 1
        json_content = response.content[0].text[content_start:content_end]
        
        result = json.loads(json_content)
        
        # Ensure we have the expected keys
        if "title" not in result or "bullets" not in result:
            raise ValueError("Response missing required keys")
            
        # Limit to max 5 bullets
        result["bullets"] = result["bullets"][:5]
        
        return result
    except Exception as e:
        # Fallback if parsing fails
        return {
            "title": "The Problem",
            "bullets": [
                "Organizations struggle with fragmented data systems.",
                "Manual processes waste valuable time and resources.",
                "Decision-makers often work with outdated information.",
                "Teams face communication barriers across departments."
            ]
        }

def generate_solution_slide_content(client: anthropic.Anthropic, product_name: str, problem_statement: str) -> Dict[str, Any]:
    """Generate content for the solution overview slide."""
    prompt = f"""
    Generate a concise, impactful solution overview paragraph for a B2B SaaS product presentation slide.
    The product name is "{product_name}" and it addresses this problem:
    
    "{problem_statement}"
    
    Guidelines:
    - Write 2-3 sentences (maximum 60 words total)
    - Focus on outcomes and benefits, not features
    - Use authoritative, confident tone
    - Emphasize transformation from problem to solution
    - Use Title Case for the slide title
    
    Format your response as a JSON object with:
    1. A "title" key containing a compelling slide title (max 10 words)
    2. A "paragraph" key containing your solution paragraph
    
    Example format:
    {{
      "title": "Our Solution",
      "paragraph": "ProductX delivers seamless data integration across platforms, eliminating silos and reducing reporting time by 78%. Teams gain instant access to unified insights, enabling faster decisions and measurable ROI within 30 days of implementation."
    }}
    """
    
    response = client.messages.create(
        model="claude-3-opus-20240229",
        max_tokens=400,
        temperature=0.3,
        system="You are an expert B2B SaaS copywriter who creates executive-ready, minimalist presentation content.",
        messages=[
            {"role": "user", "content": prompt}
        ]
    )
    
    # Parse the JSON response
    try:
        content_start = response.content[0].text.find('{')
        content_end = response.content[0].text.rfind('}') + 1
        json_content = response.content[0].text[content_start:content_end]
        
        result = json.loads(json_content)
        
        # Ensure we have the expected keys
        if "title" not in result or "paragraph" not in result:
            raise ValueError("Response missing required keys")
            
        # Truncate paragraph if too long
        result["paragraph"] = truncate_text(result["paragraph"], 300)
        
        return result
    except Exception as e:
        # Fallback if parsing fails
        return {
            "title": "Our Solution",
            "paragraph": f"{product_name} streamlines your data workflow with an intuitive platform that connects all systems. Teams gain instant access to accurate information, reducing reporting time by 70% and enabling data-driven decisions that drive business growth."
        }

def generate_features_slide_content(client: anthropic.Anthropic, features: List[str]) -> Dict[str, Any]:
    """Generate content for the key features slide."""
    features_text = "\n".join([f"- {feature}" for feature in features])
    
    prompt = f"""
    Create concise, benefit-focused descriptions for each of these key features for a B2B SaaS product presentation.
    
    Here are the raw features:
    {features_text}
    
    Guidelines:
    - Keep the original feature name, but enhance its description
    - Each feature should be in format "Feature Name: Benefit statement"
    - Benefit statement should be 8-12 words maximum
    - Focus on outcomes, not technical details
    - Use active voice and present tense
    - Use Title Case for the slide title
    
    Format your response as a JSON object with:
    1. A "title" key containing a compelling slide title (max 10 words)
    2. A "features" key containing an array of enhanced feature descriptions
    
    Example format:
    {{
      "title": "Key Features",
      "features": [
        "Real-time Analytics: Make decisions faster with instant data insights",
        "Seamless Integration: Connect to existing systems without coding"
      ]
    }}
    """
    
    response = client.messages.create(
        model="claude-3-opus-20240229",
        max_tokens=600,
        temperature=0.3,
        system="You are an expert B2B SaaS copywriter who creates executive-ready, minimalist presentation content.",
        messages=[
            {"role": "user", "content": prompt}
        ]
    )
    
    # Parse the JSON response
    try:
        content_start = response.content[0].text.find('{')
        content_end = response.content[0].text.rfind('}') + 1
        json_content = response.content[0].text[content_start:content_end]
        
        result = json.loads(json_content)
        
        # Ensure we have the expected keys
        if "title" not in result or "features" not in result:
            raise ValueError("Response missing required keys")
            
        # Ensure we have the right number of features
        if len(result["features"]) != len(features):
            result["features"] = [f"{feature}: Enhanced functionality for better results" for feature in features]
        
        return result
    except Exception as e:
        # Fallback if parsing fails
        return {
            "title": "Key Features",
            "features": [f"{feature}: Enhanced functionality for better results" for feature in features]
        }

def generate_advantage_slide_content(client: anthropic.Anthropic, competitive_advantage: str) -> Dict[str, Any]:
    """Generate content for the competitive advantage slide."""
    prompt = f"""
    Generate 3-4 compelling bullet points highlighting competitive advantages for a B2B SaaS product presentation.
    Base the content on this competitive advantage statement:
    
    "{competitive_advantage}"
    
    Guidelines:
    - Each bullet should be 1-2 sentences, maximum 20 words
    - Focus on quantifiable advantages (faster, cheaper, more reliable)
    - Include specific metrics or percentages where possible
    - Compare to industry standards or competitors without naming them
    - Use Title Case for the slide title
    
    Format your response as a JSON object with:
    1. A "title" key containing a compelling slide title (max 10 words)
    2. A "bullets" key containing an array of advantage bullet points
    
    Example format:
    {{
      "title": "Why We Lead The Market",
      "bullets": [
        "70% faster implementation than industry average.",
        "ROI within 3 months versus typical 12-month payback."
      ]
    }}
    """
    
    response = client.messages.create(
        model="claude-3-opus-20240229",
        max_tokens=500,
        temperature=0.3,
        system="You are an expert B2B SaaS copywriter who creates executive-ready, minimalist presentation content.",
        messages=[
            {"role": "user", "content": prompt}
        ]
    )
    
    # Parse the JSON response
    try:
        content_start = response.content[0].text.find('{')
        content_end = response.content[0].text.rfind('}') + 1
        json_content = response.content[0].text[content_start:content_end]
        
        result = json.loads(json_content)
        
        # Ensure we have the expected keys
        if "title" not in result or "bullets" not in result:
            raise ValueError("Response missing required keys")
            
        # Limit to max 4 bullets
        result["bullets"] = result["bullets"][:4]
        
        return result
    except Exception as e:
        # Fallback if parsing fails
        return {
            "title": "Why We Win",
            "bullets": [
                "70% faster implementation than industry average.",
                "3x higher data accuracy rates than competitors.",
                "Flexible pricing saves companies 40% on licensing costs.",
                "24/7 dedicated support with 15-minute response times."
            ]
        }

def generate_audience_slide_content(client: anthropic.Anthropic, target_audience: str) -> Dict[str, Any]:
    """Generate content for the target audience slide."""
    prompt = f"""
    Generate a concise paragraph describing the ideal customer for a B2B SaaS product presentation.
    Base the content on this target audience description:
    
    "{target_audience}"
    
    Guidelines:
    - Write 2-3 sentences (maximum 60 words total)
    - Describe ideal customer profile in specific terms
    - Include company size, role titles, industry verticals if applicable
    - Focus on pain points and desired outcomes
    - Use Title Case for the slide title
    
    Format your response as a JSON object with:
    1. A "title" key containing a compelling slide title (max 10 words)
    2. A "paragraph" key containing your target audience paragraph
    
    Example format:
    {{
      "title": "Who We Serve",
      "paragraph": "Mid-size financial services organizations with 100-1000 employees and complex reporting requirements. Our solution serves CFOs, controllers, and finance teams seeking to reduce monthly close time and improve financial visibility across departments."
    }}
    """
    
    response = client.messages.create(
        model="claude-3-opus-20240229",
        max_tokens=400,
        temperature=0.3,
        system="You are an expert B2B SaaS copywriter who creates executive-ready, minimalist presentation content.",
        messages=[
            {"role": "user", "content": prompt}
        ]
    )
    
    # Parse the JSON response
    try:
        content_start = response.content[0].text.find('{')
        content_end = response.content[0].text.rfind('}') + 1
        json_content = response.content[0].text[content_start:content_end]
        
        result = json.loads(json_content)
        
        # Ensure we have the expected keys
        if "title" not in result or "paragraph" not in result:
            raise ValueError("Response missing required keys")
            
        # Truncate paragraph if too long
        result["paragraph"] = truncate_text(result["paragraph"], 300)
        
        return result
    except Exception as e:
        # Fallback if parsing fails
        return {
            "title": "Who We Serve",
            "paragraph": f"Mid-size enterprises with complex data needs and cross-departmental reporting requirements. Our platform serves CTOs, IT directors, and data teams seeking to streamline workflows, eliminate manual processes, and enable data-driven decision making."
        }