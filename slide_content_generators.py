import google.generativeai as genai
from typing import List, Dict, Any

def generate_title_slide_content(product_name: str, company_name: str) -> Dict[str, Any]:
    """Generate content for the title slide."""
    model = genai.GenerativeModel('gemini-1.5-flash')
    prompt = f"""
    Create a compelling and concise title slide content for a B2B SaaS product presentation.
    
    Product: {product_name}
    Company: {company_name}
    
    Generate a short subtitle (1 line) that highlights the product's key value proposition.
    Keep it professional and business-focused.
    
    Return the result as JSON with the following fields:
    - title: The product name
    - subtitle: A concise subtitle that captures the product's value proposition
    - company: The company name
    """
    
    response = model.generate_content(prompt)
    
    # Process the response to extract structured content
    try:
        content_text = response.text
        # Check if the response contains JSON-like structure and extract it
        if '{' in content_text and '}' in content_text:
            start_idx = content_text.find('{')
            end_idx = content_text.rfind('}') + 1
            json_str = content_text[start_idx:end_idx]
            import json
            content = json.loads(json_str)
        else:
            # Fall back to default structure if JSON parsing fails
            content = {
                "title": product_name,
                "subtitle": "Innovative Software Solution",
                "company": company_name
            }
    except Exception:
        # Fallback content if parsing fails
        content = {
            "title": product_name,
            "subtitle": "Innovative Software Solution",
            "company": company_name
        }
    
    return content

def generate_problem_slide_content(problem_statement: str, persona: str) -> Dict[str, Any]:
    """Generate content for the problem slide."""
    model = genai.GenerativeModel('gemini-1.5-flash')
    
    persona_context = get_persona_context(persona)
    
    prompt = f"""
    Create concise problem slide content for a B2B SaaS product presentation.
    Adapt the content for a {persona_context}.
    
    Problem Statement: {problem_statement}
    
    Extract 3 key pain points from this statement and format them as short, impactful bullet points.
    
    Return the result as JSON with the following fields:
    - title: A compelling problem slide title (e.g., "The Challenge" or similar)
    - pain_points: An array of 3 concise bullet points highlighting key problems
    - description: A short (2-3 sentence) summary of the overall problem
    """
    
    response = model.generate_content(prompt)
    
    # Process the response to extract structured content
    try:
        content_text = response.text
        # Check if the response contains JSON-like structure and extract it
        if '{' in content_text and '}' in content_text:
            start_idx = content_text.find('{')
            end_idx = content_text.rfind('}') + 1
            json_str = content_text[start_idx:end_idx]
            import json
            content = json.loads(json_str)
        else:
            # Fall back to default structure if JSON parsing fails
            content = {
                "title": "The Challenge",
                "pain_points": [
                    "Businesses struggle with increasing complexity",
                    "Manual processes lead to inefficiency",
                    "Legacy systems create security vulnerabilities"
                ],
                "description": problem_statement[:100] + "..."
            }
    except Exception:
        # Fallback content if parsing fails
        content = {
            "title": "The Challenge",
            "pain_points": [
                "Businesses struggle with increasing complexity",
                "Manual processes lead to inefficiency",
                "Legacy systems create security vulnerabilities"
            ],
            "description": problem_statement[:100] + "..."
        }
    
    return content

def generate_solution_slide_content(product_name: str, problem_statement: str, persona: str) -> Dict[str, Any]:
    """Generate content for the solution slide."""
    model = genai.GenerativeModel('gemini-1.5-flash')
    
    persona_context = get_persona_context(persona)
    
    prompt = f"""
    Create solution slide content for a B2B SaaS product presentation.
    Adapt the content for a {persona_context}.
    
    Product Name: {product_name}
    Problem Statement: {problem_statement}
    
    Explain how this product solves the described problem.
    
    Return the result as JSON with the following fields:
    - title: A solution-focused slide title (e.g., "Our Solution" or incorporating the product name)
    - value_proposition: A single-sentence core value proposition  
    - description: A concise (2-3 sentence) explanation of how the product solves the problem
    """
    
    response = model.generate_content(prompt)
    
    # Process the response to extract structured content
    try:
        content_text = response.text
        # Check if the response contains JSON-like structure and extract it
        if '{' in content_text and '}' in content_text:
            start_idx = content_text.find('{')
            end_idx = content_text.rfind('}') + 1
            json_str = content_text[start_idx:end_idx]
            import json
            content = json.loads(json_str)
        else:
            # Fall back to default structure if JSON parsing fails
            content = {
                "title": f"Introducing {product_name}",
                "value_proposition": f"{product_name} streamlines operations and boosts efficiency",
                "description": f"{product_name} offers a comprehensive solution to address the challenges faced by organizations today."
            }
    except Exception:
        # Fallback content if parsing fails
        content = {
            "title": f"Introducing {product_name}",
            "value_proposition": f"{product_name} streamlines operations and boosts efficiency",
            "description": f"{product_name} offers a comprehensive solution to address the challenges faced by organizations today."
        }
    
    return content

def generate_features_slide_content(key_features: List[str], persona: str) -> Dict[str, Any]:
    """Generate content for the features slide."""
    model = genai.GenerativeModel('gemini-1.5-flash')
    
    persona_context = get_persona_context(persona)
    features_text = "\n".join([f"- {feature}" for feature in key_features])
    
    prompt = f"""
    Create key features slide content for a B2B SaaS product presentation.
    Adapt the content for a {persona_context}.
    
    Features:
    {features_text}
    
    Return the result as JSON with the following fields:
    - title: A features-focused slide title
    - features: An array of feature objects, each containing:
      - name: The feature name/title
      - description: A short (1 sentence) benefit-focused description
    """
    
    response = model.generate_content(prompt)
    
    # Process the response to extract structured content
    try:
        content_text = response.text
        # Check if the response contains JSON-like structure and extract it
        if '{' in content_text and '}' in content_text:
            start_idx = content_text.find('{')
            end_idx = content_text.rfind('}') + 1
            json_str = content_text[start_idx:end_idx]
            import json
            content = json.loads(json_str)
        else:
            # Fall back to default structure if JSON parsing fails
            simple_features = []
            for i, feature in enumerate(key_features[:6]):  # Limit to first 6 features
                if ":" in feature:
                    name, desc = feature.split(":", 1)
                    simple_features.append({"name": name.strip(), "description": desc.strip()})
                else:
                    simple_features.append({"name": feature, "description": "Provides significant business value"})
            
            content = {
                "title": "Key Features",
                "features": simple_features
            }
    except Exception:
        # Fallback content if parsing fails
        simple_features = []
        for i, feature in enumerate(key_features[:6]):  # Limit to first 6 features
            if ":" in feature:
                name, desc = feature.split(":", 1)
                simple_features.append({"name": name.strip(), "description": desc.strip()})
            else:
                simple_features.append({"name": feature, "description": "Provides significant business value"})
        
        content = {
            "title": "Key Features",
            "features": simple_features
        }
    
    return content

def generate_advantage_slide_content(competitive_advantage: str, persona: str) -> Dict[str, Any]:
    """Generate content for the competitive advantage slide."""
    model = genai.GenerativeModel('gemini-1.5-flash')
    
    persona_context = get_persona_context(persona)
    
    prompt = f"""
    Create competitive advantage slide content for a B2B SaaS product presentation.
    Adapt the content for a {persona_context}.
    
    Competitive Advantage details:
    {competitive_advantage}
    
    Extract 3-4 key differentiators from this information.
    
    Return the result as JSON with the following fields:
    - title: A competitive advantage slide title (e.g., "Why Choose Us" or similar)
    - differentiators: An array of differentiation points (3-4 items), each containing:
      - point: The competitive advantage point
      - description: A brief explanation of why it matters
    """
    
    response = model.generate_content(prompt)
    
    # Process the response to extract structured content
    try:
        content_text = response.text
        # Check if the response contains JSON-like structure and extract it
        if '{' in content_text and '}' in content_text:
            start_idx = content_text.find('{')
            end_idx = content_text.rfind('}') + 1
            json_str = content_text[start_idx:end_idx]
            import json
            content = json.loads(json_str)
        else:
            # Fall back to default structure if JSON parsing fails
            content = {
                "title": "Why Choose Us",
                "differentiators": [
                    {"point": "Superior Performance", "description": "Our solution delivers faster results with greater accuracy"},
                    {"point": "Easy Integration", "description": "Seamlessly fits into your existing technology stack"},
                    {"point": "Exceptional Support", "description": "24/7 access to expert assistance"}
                ]
            }
    except Exception:
        # Fallback content if parsing fails
        content = {
            "title": "Why Choose Us",
            "differentiators": [
                {"point": "Superior Performance", "description": "Our solution delivers faster results with greater accuracy"},
                {"point": "Easy Integration", "description": "Seamlessly fits into your existing technology stack"},
                {"point": "Exceptional Support", "description": "24/7 access to expert assistance"}
            ]
        }
    
    return content

def generate_audience_slide_content(target_audience: str, persona: str) -> Dict[str, Any]:
    """Generate content for the target audience slide."""
    model = genai.GenerativeModel('gemini-1.5-flash')
    
    persona_context = get_persona_context(persona)
    
    prompt = f"""
    Create target audience slide content for a B2B SaaS product presentation.
    Adapt the content for a {persona_context}.
    
    Target Audience:
    {target_audience}
    
    Return the result as JSON with the following fields:
    - title: A target audience focused slide title
    - audience_segments: An array of 2-3 key audience segments or personas
    - use_cases: An array of 2-3 primary use cases for these segments
    """
    
    response = model.generate_content(prompt)
    
    # Process the response to extract structured content
    try:
        content_text = response.text
        # Check if the response contains JSON-like structure and extract it
        if '{' in content_text and '}' in content_text:
            start_idx = content_text.find('{')
            end_idx = content_text.rfind('}') + 1
            json_str = content_text[start_idx:end_idx]
            import json
            content = json.loads(json_str)
        else:
            # Fall back to default structure if JSON parsing fails
            content = {
                "title": "Who We Serve",
                "audience_segments": ["Enterprise Teams", "Mid-Market Companies", "Service Providers"],
                "use_cases": [
                    "Streamlining Operations",
                    "Enhancing Security Measures",
                    "Improving Customer Experiences"
                ]
            }
    except Exception:
        # Fallback content if parsing fails
        content = {
            "title": "Who We Serve",
            "audience_segments": ["Enterprise Teams", "Mid-Market Companies", "Service Providers"],
            "use_cases": [
                "Streamlining Operations",
                "Enhancing Security Measures",
                "Improving Customer Experiences"
            ]
        }
    
    return content

def generate_cta_slide_content(call_to_action: str, product_name: str, persona: str) -> Dict[str, Any]:
    """Generate content for the call to action slide."""
    model = genai.GenerativeModel('gemini-1.5-flash')
    
    persona_context = get_persona_context(persona)
    
    prompt = f"""
    Create call-to-action slide content for a B2B SaaS product presentation.
    Adapt the content for a {persona_context}.
    
    Product Name: {product_name}
    Call to Action: {call_to_action}
    
    Return the result as JSON with the following fields:
    - title: An action-oriented slide title
    - cta_main: The primary call-to-action statement (should incorporate the provided CTA)
    - next_steps: An array of 2-3 specific next steps the audience can take
    - contact_info: Placeholder for contact information (generic)
    """
    
    response = model.generate_content(prompt)
    
    # Process the response to extract structured content
    try:
        content_text = response.text
        # Check if the response contains JSON-like structure and extract it
        if '{' in content_text and '}' in content_text:
            start_idx = content_text.find('{')
            end_idx = content_text.rfind('}') + 1
            json_str = content_text[start_idx:end_idx]
            import json
            content = json.loads(json_str)
        else:
            # Fall back to default structure if JSON parsing fails
            content = {
                "title": "Next Steps",
                "cta_main": call_to_action,
                "next_steps": [
                    "Schedule a demo",
                    "Speak with our specialists",
                    "Request a custom quote"
                ],
                "contact_info": "example@company.com | (555) 123-4567"
            }
    except Exception:
        # Fallback content if parsing fails
        content = {
            "title": "Next Steps",
            "cta_main": call_to_action,
            "next_steps": [
                "Schedule a demo",
                "Speak with our specialists",
                "Request a custom quote"
            ],
            "contact_info": "example@company.com | (555) 123-4567"
        }
    
    return content

def generate_market_slide_content(target_audience: str, persona: str) -> Dict[str, Any]:
    """Generate content for the market analysis slide."""
    model = genai.GenerativeModel('gemini-1.5-flash')
    
    persona_context = get_persona_context(persona)
    
    prompt = f"""
    Create market analysis slide content for a B2B SaaS product presentation.
    Adapt the content for a {persona_context}.
    
    Target Audience: {target_audience}
    
    Return the result as JSON with the following fields:
    - title: A market-focused slide title
    - market_size: A statement about market size/opportunity
    - trends: An array of 2-3 relevant industry trends
    - growth_points: An array of 2-3 growth indicators or statistics (use realistic but general numbers)
    """
    
    response = model.generate_content(prompt)
    
    # Process the response to extract structured content
    try:
        content_text = response.text
        # Check if the response contains JSON-like structure and extract it
        if '{' in content_text and '}' in content_text:
            start_idx = content_text.find('{')
            end_idx = content_text.rfind('}') + 1
            json_str = content_text[start_idx:end_idx]
            import json
            content = json.loads(json_str)
        else:
            # Fall back to default structure if JSON parsing fails
            content = {
                "title": "Market Opportunity",
                "market_size": "The global market is projected to reach $25B by 2025, growing at 15% CAGR",
                "trends": [
                    "Increasing adoption of cloud-based solutions",
                    "Growing demand for automation and AI integration",
                    "Rising focus on data security and compliance"
                ],
                "growth_points": [
                    "Enterprise segment growing at 18% annually",
                    "35% of companies planning increased investment",
                    "ROI realized within 6 months for 70% of customers"
                ]
            }
    except Exception:
        # Fallback content if parsing fails
        content = {
            "title": "Market Opportunity",
            "market_size": "The global market is projected to reach $25B by 2025, growing at 15% CAGR",
            "trends": [
                "Increasing adoption of cloud-based solutions",
                "Growing demand for automation and AI integration",
                "Rising focus on data security and compliance"
            ],
            "growth_points": [
                "Enterprise segment growing at 18% annually",
                "35% of companies planning increased investment",
                "ROI realized within 6 months for 70% of customers"
            ]
        }
    
    return content

def generate_roadmap_slide_content(product_name: str, persona: str) -> Dict[str, Any]:
    """Generate content for the product roadmap slide."""
    model = genai.GenerativeModel('gemini-1.5-flash')
    
    persona_context = get_persona_context(persona)
    
    prompt = f"""
    Create product roadmap slide content for a B2B SaaS product presentation.
    Adapt the content for a {persona_context}.
    
    Product Name: {product_name}
    
    Return the result as JSON with the following fields:
    - title: A roadmap-focused slide title
    - current_quarter: Features available now (array of 2-3 items)
    - next_quarter: Features coming soon (array of 2-3 items)
    - future: Strategic direction (array of 1-2 items)
    """
    
    response = model.generate_content(prompt)
    
    # Process the response to extract structured content
    try:
        content_text = response.text
        # Check if the response contains JSON-like structure and extract it
        if '{' in content_text and '}' in content_text:
            start_idx = content_text.find('{')
            end_idx = content_text.rfind('}') + 1
            json_str = content_text[start_idx:end_idx]
            import json
            content = json.loads(json_str)
        else:
            # Fall back to default structure if JSON parsing fails
            content = {
                "title": f"{product_name} Roadmap",
                "current_quarter": [
                    "Core functionality and integrations",
                    "Advanced analytics dashboard",
                    "Mobile application support"
                ],
                "next_quarter": [
                    "AI-powered recommendations",
                    "Extended API capabilities",
                    "Enhanced security features"
                ],
                "future": [
                    "International expansion",
                    "Enterprise-scale solutions"
                ]
            }
    except Exception:
        # Fallback content if parsing fails
        content = {
            "title": f"{product_name} Roadmap",
            "current_quarter": [
                "Core functionality and integrations",
                "Advanced analytics dashboard",
                "Mobile application support"
            ],
            "next_quarter": [
                "AI-powered recommendations",
                "Extended API capabilities",
                "Enhanced security features"
            ],
            "future": [
                "International expansion",
                "Enterprise-scale solutions"
            ]
        }
    
    return content

def generate_team_slide_content(company_name: str, persona: str) -> Dict[str, Any]:
    """Generate content for the team overview slide."""
    model = genai.GenerativeModel('gemini-1.5-flash')
    
    persona_context = get_persona_context(persona)
    
    prompt = f"""
    Create team overview slide content for a B2B SaaS product presentation.
    Adapt the content for a {persona_context}.
    
    Company Name: {company_name}
    
    Return the result as JSON with the following fields:
    - title: A team-focused slide title
    - team_highlight: A brief statement about the team's expertise/experience
    - key_roles: An array of 3-4 generic leadership positions with brief descriptions
    """
    
    response = model.generate_content(prompt)
    
    # Process the response to extract structured content
    try:
        content_text = response.text
        # Check if the response contains JSON-like structure and extract it
        if '{' in content_text and '}' in content_text:
            start_idx = content_text.find('{')
            end_idx = content_text.rfind('}') + 1
            json_str = content_text[start_idx:end_idx]
            import json
            content = json.loads(json_str)
        else:
            # Fall back to default structure if JSON parsing fails
            content = {
                "title": f"Meet the {company_name} Team",
                "team_highlight": f"Our experienced team brings 50+ years of combined expertise in industry solutions",
                "key_roles": [
                    {"role": "Chief Executive Officer", "description": "Strategic vision and leadership"},
                    {"role": "Chief Technology Officer", "description": "Product architecture and innovation"},
                    {"role": "VP of Customer Success", "description": "Ensuring client outcomes"},
                    {"role": "Head of Development", "description": "Platform engineering excellence"}
                ]
            }
    except Exception:
        # Fallback content if parsing fails
        content = {
            "title": f"Meet the {company_name} Team",
            "team_highlight": f"Our experienced team brings 50+ years of combined expertise in industry solutions",
            "key_roles": [
                {"role": "Chief Executive Officer", "description": "Strategic vision and leadership"},
                {"role": "Chief Technology Officer", "description": "Product architecture and innovation"},
                {"role": "VP of Customer Success", "description": "Ensuring client outcomes"},
                {"role": "Head of Development", "description": "Platform engineering excellence"}
            ]
        }
    
    return content

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
