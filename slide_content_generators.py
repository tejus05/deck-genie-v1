import google.generativeai as genai
from typing import Dict, List, Any
import json

def generate_title_slide_content(product_name: str, company_name: str) -> Dict[str, Any]:
    """Generate title slide content."""
    model = genai.GenerativeModel('gemini-pro')
    
    prompt = f"""
    Create a professional title slide for a business presentation with these details:
    - Product name: {product_name}
    - Company name: {company_name}
    
    Return a JSON object with these fields:
    - title: The main title for the slide (should include the product name)
    - subtitle: A subtitle (should include the company name)
    - product_name: The product name for reference
    
    Make the title concise but impactful. Format as valid JSON only.
    """
    
    response = model.generate_content(prompt)
    content = extract_json_from_response(response.text)
    
    # Ensure correct fields exist
    if not content.get("title"):
        content["title"] = f"{product_name}"
    if not content.get("subtitle"):
        content["subtitle"] = f"Presented by {company_name}"
    content["product_name"] = product_name
    
    return content

def generate_problem_slide_content(problem_statement: str, persona: str) -> Dict[str, Any]:
    """Generate problem slide content with persona-specific focus."""
    model = genai.GenerativeModel('gemini-pro')
    
    # Enhanced persona-specific prompts
    persona_prompts = {
        "Technical": """Focus on technical pain points, system limitations, performance issues, and engineering challenges. 
        Emphasize scalability problems, integration difficulties, maintenance overhead, and technical debt.""",
        
        "Marketing": """Focus on customer pain points, market gaps, user experience issues, and customer acquisition challenges. 
        Emphasize customer dissatisfaction, market demand, competitive weaknesses, and opportunities for differentiation.""",
        
        "Executive": """Focus on business impact, strategic challenges, operational inefficiencies, and competitive threats. 
        Emphasize revenue impact, market positioning, resource allocation, and strategic risks.""",
        
        "Investor": """Focus on market inefficiencies, unmet demand, scalable problems, and monetizable pain points. 
        Emphasize market size of problems, competitive gaps, and opportunities for disruption and returns.""",
        
        "Generic": """Focus on balanced mix of technical and business challenges that affect multiple stakeholders."""
    }
    
    persona_specific = persona_prompts.get(persona, persona_prompts["Generic"])
    
    prompt = f"""
    Create a problem statement slide for {persona} audience based on this description:
    "{problem_statement}"
    
    {persona_specific}
    
    Make the content highly relevant and actionable for {persona} professionals.
    
    Return a JSON object with these fields:
    - title: A compelling slide title about the problem (tailored for {persona} audience)
    - pain_points: An array of 3-4 specific pain points extracted/derived from the problem statement
    
    Format as valid JSON only. Each bullet point should be specific to {persona} concerns.
    """
    
    response = model.generate_content(prompt)
    content = extract_json_from_response(response.text)
    
    # Ensure correct fields exist
    if not content.get("title"):
        content["title"] = "The Problem"
    if not content.get("pain_points") or not isinstance(content.get("pain_points"), list):
        # Extract simple bullet points from the problem statement
        content["pain_points"] = [s.strip() + "." for s in problem_statement.split(". ") if s.strip()][:3]
    
    return content

def generate_solution_slide_content(product_name: str, problem_statement: str, persona: str) -> Dict[str, Any]:
    """Generate solution slide content with persona-specific value proposition."""
    model = genai.GenerativeModel('gemini-pro')
    
    # Enhanced persona-specific solution focus
    persona_prompts = {
        "Technical": """Focus on technical architecture, implementation approach, engineering solutions, and technical benefits. 
        Emphasize scalability, performance, integration capabilities, and technical innovation.""",
        
        "Marketing": """Focus on customer value proposition, user benefits, market differentiation, and customer success. 
        Emphasize customer outcomes, user experience improvements, and competitive advantages.""",
        
        "Executive": """Focus on business impact, strategic value, operational improvements, and competitive positioning. 
        Emphasize ROI, efficiency gains, market opportunities, and strategic advantages.""",
        
        "Investor": """Focus on market opportunity, scalable solution, revenue potential, and competitive moats. 
        Emphasize addressable market, growth potential, monetization strategy, and sustainable advantages.""",
        
        "Generic": """Focus on balanced technical and business value that appeals to multiple stakeholders."""
    }
    
    persona_specific = persona_prompts.get(persona, persona_prompts["Generic"])
    
    prompt = f"""
    Create a solution overview slide for {persona} audience for "{product_name}" that addresses:
    "{problem_statement}"
    
    {persona_specific}
    
    Make the solution compelling and specific to {persona} decision-making criteria.
    
    Return a JSON object with these fields:
    - title: A compelling slide title introducing the solution (include product name, tailored for {persona})
    - paragraph: A focused paragraph (60-80 words) explaining how {product_name} solves the problem for {persona} audience
    
    The paragraph should clearly articulate value specific to {persona} professionals.
    Format as valid JSON only.
    """
    
    response = model.generate_content(prompt)
    content = extract_json_from_response(response.text)
    
    # Ensure correct fields exist
    if not content.get("title"):
        content["title"] = f"Introducing {product_name}"
    if not content.get("paragraph"):
        content["paragraph"] = f"{product_name} is designed to address these challenges with an innovative approach that helps organizations achieve better outcomes."
    
    return content

def generate_features_slide_content(key_features: List[str], persona: str) -> Dict[str, Any]:
    """Generate features slide content."""
    model = genai.GenerativeModel('gemini-pro')
    
    # Limit to 5 features
    features_text = "\n".join(f"- {f}" for f in key_features[:5])
    
    persona_specific = ""
    if persona == "Technical":
        persona_specific = "Focus on technical capabilities and specifications."
    elif persona == "Marketing":
        persona_specific = "Focus on benefits and value to customers."
    elif persona == "Executive":
        persona_specific = "Focus on business impact and strategic advantages."
    elif persona == "Investor":
        persona_specific = "Focus on market differentiation and competitive advantages."
    
    prompt = f"""
    Create a key features slide based on these features:
    {features_text}
    
    {persona_specific}
    
    Return a JSON object with these fields:
    - title: A slide title for key features
    - features: A refined list of the key features (keep the same features but enhance their descriptions)
    
    Each feature should be concise but clear. Format as valid JSON only.
    """
    
    response = model.generate_content(prompt)
    content = extract_json_from_response(response.text)
    
    # Ensure correct fields exist
    if not content.get("title"):
        content["title"] = "Key Features"
    if not content.get("features") or not isinstance(content.get("features"), list):
        content["features"] = key_features[:5]
    
    return content

def generate_advantage_slide_content(competitive_advantage: str, persona: str) -> Dict[str, Any]:
    """Generate competitive advantage slide content."""
    model = genai.GenerativeModel('gemini-pro')
    
    persona_specific = ""
    if persona == "Technical":
        persona_specific = "Focus on technical superiority and unique capabilities."
    elif persona == "Marketing":
        persona_specific = "Focus on market differentiation and customer benefits."
    elif persona == "Executive":
        persona_specific = "Focus on business impact, ROI, and strategic advantages."
    elif persona == "Investor":
        persona_specific = "Focus on market positioning, barriers to entry, and sustainable competitive edges."
    
    prompt = f"""
    Create a competitive advantage slide based on this statement:
    "{competitive_advantage}"
    
    {persona_specific}
    
    Return a JSON object with these fields:
    - title: A compelling slide title about competitive advantages
    - differentiators: An array of 3-4 key differentiators or advantages
    
    Make each differentiator concise but specific. Format as valid JSON only.
    """
    
    response = model.generate_content(prompt)
    content = extract_json_from_response(response.text)
    
    # Ensure correct fields exist
    if not content.get("title"):
        content["title"] = "Our Competitive Advantage"
    if not content.get("differentiators") or not isinstance(content.get("differentiators"), list):
        # Extract simple bullet points from the competitive advantage
        content["differentiators"] = [s.strip() + "." for s in competitive_advantage.split(". ") if s.strip()][:3]
    
    return content

def generate_audience_slide_content(target_audience: str, persona: str) -> Dict[str, Any]:
    """Generate target audience slide content."""
    model = genai.GenerativeModel('gemini-pro')
    
    persona_specific = ""
    if persona == "Technical":
        persona_specific = "Focus on technical roles, use cases, and technical needs."
    elif persona == "Marketing":
        persona_specific = "Focus on demographics, psychographics, and audience pain points."
    elif persona == "Executive":
        persona_specific = "Focus on key decision-makers, organizational impact, and business needs."
    elif persona == "Investor":
        persona_specific = "Focus on market segments, target market size, and customer acquisition strategy."
    
    prompt = f"""
    Create a target audience slide based on this description:
    "{target_audience}"
    
    {persona_specific}
    
    Return a JSON object with these fields:
    - title: A compelling slide title about the target audience
    - paragraph: A concise paragraph (60-80 words) that provides a rich description of the target audience
    
    Format as valid JSON only.
    """
    
    response = model.generate_content(prompt)
    content = extract_json_from_response(response.text)
    
    # Ensure correct fields exist
    if not content.get("title"):
        content["title"] = "Target Audience"
    if not content.get("paragraph"):
        content["paragraph"] = target_audience
    
    return content

def generate_cta_slide_content(call_to_action: str, product_name: str, persona: str) -> Dict[str, Any]:
    """Generate call to action slide content."""
    model = genai.GenerativeModel('gemini-pro')
    
    persona_specific = ""
    if persona == "Technical":
        persona_specific = "Focus on technical next steps like demos, trials, or technical discussions."
    elif persona == "Marketing":
        persona_specific = "Focus on engagement, lead generation, and customer journey."
    elif persona == "Executive":
        persona_specific = "Focus on high-level business discussions, partnerships, and strategic decisions."
    elif persona == "Investor":
        persona_specific = "Focus on investment opportunities, funding rounds, or partnership discussions."
    
    prompt = f"""
    Create a call-to-action slide based on this CTA:
    "{call_to_action}"
    
    Product name: {product_name}
    {persona_specific}
    
    Return a JSON object with these fields:
    - title: An engaging slide title for the call to action
    - cta_text: A clear, compelling call-to-action statement
    - bullets: (optional) An array of 2-3 supporting points or next steps
    
    Format as valid JSON only.
    """
    
    response = model.generate_content(prompt)
    content = extract_json_from_response(response.text)
    
    # Ensure correct fields exist
    if not content.get("title"):
        content["title"] = "Get Started Today"
    if not content.get("cta_text"):
        content["cta_text"] = call_to_action or f"Contact us to learn more about {product_name}"
    
    return content

def generate_market_slide_content(target_audience: str, persona: str) -> Dict[str, Any]:
    """Generate market opportunity slide content."""
    model = genai.GenerativeModel('gemini-pro')
    
    persona_specific = ""
    if persona == "Investor" or persona == "Executive":
        persona_specific = "Include specific market size numbers, growth rates, and market opportunity details."
    
    prompt = f"""
    Create a market opportunity slide for a product targeting this audience:
    "{target_audience}"
    
    {persona_specific}
    
    Return a JSON object with these fields:
    - title: A compelling slide title about the market opportunity
    - market_size: A specific market size figure with dollar amount and timeframe (e.g. "$25B by 2025")
    - growth_rate: Annual growth rate with percentage (e.g. "15% CAGR")
    - description: A brief paragraph about the market opportunity and trends
    
    Be specific and data-driven with realistic but impressive figures. Format as valid JSON only.
    """
    
    response = model.generate_content(prompt)
    content = extract_json_from_response(response.text)
    
    # Ensure correct fields exist
    if not content.get("title"):
        content["title"] = "Market Opportunity"
    if not content.get("market_size"):
        content["market_size"] = "$25B by 2025"
    if not content.get("growth_rate"):
        content["growth_rate"] = "15% CAGR"
    if not content.get("description"):
        content["description"] = "The market is experiencing significant growth as organizations increasingly recognize the need for innovative solutions in this space."
    
    return content

def generate_roadmap_slide_content(product_name: str, persona: str) -> Dict[str, Any]:
    """Generate product roadmap slide content."""
    model = genai.GenerativeModel('gemini-pro')
    
    persona_specific = ""
    if persona == "Technical":
        persona_specific = "Focus on technical features, integrations, and architecture evolution."
    elif persona == "Marketing":
        persona_specific = "Focus on customer-facing features and market expansion."
    elif persona == "Executive":
        persona_specific = "Focus on strategic milestones, market expansion, and business objectives."
    elif persona == "Investor":
        persona_specific = "Focus on growth milestones, market expansion, and revenue phases."
    
    prompt = f"""
    Create a product roadmap slide for "{product_name}" with three clear phases.
    
    {persona_specific}
    
    Return a JSON object with these fields:
    - title: A slide title for the product roadmap (include product name)
    - phases: An array of 3 phase objects, each containing:
      - name: The phase name (e.g., "Phase 1: Foundation")
      - items: An array of 3-4 key deliverables or milestones for that phase
    
    Make the roadmap realistic and strategic. Format as valid JSON only.
    """
    
    response = model.generate_content(prompt)
    content = extract_json_from_response(response.text)
    
    # Ensure correct fields exist
    if not content.get("title"):
        content["title"] = f"{product_name}: Product Roadmap"
    
    # Ensure phases exist and have the right structure
    if not content.get("phases") or not isinstance(content.get("phases"), list):
        content["phases"] = [
            {
                "name": "Phase 1",
                "items": ["Initial launch", "Core features", "First customers"]
            },
            {
                "name": "Phase 2",
                "items": ["Advanced analytics", "Integration APIs", "Expanded support"]
            },
            {
                "name": "Phase 3",
                "items": ["Enterprise features", "Global expansion", "Industry partnerships"]
            }
        ]
    else:
        # Fix any phases that don't have the right structure
        for i, phase in enumerate(content["phases"]):
            if not isinstance(phase, dict):
                content["phases"][i] = {
                    "name": f"Phase {i+1}",
                    "items": ["Milestone 1", "Milestone 2", "Milestone 3"]
                }
            elif "name" not in phase:
                phase["name"] = f"Phase {i+1}"
            elif "items" not in phase or not isinstance(phase["items"], list):
                phase["items"] = ["Milestone 1", "Milestone 2", "Milestone 3"]
    
    return content

def generate_team_slide_content(company_name: str, persona: str) -> Dict[str, Any]:
    """Generate team slide content."""
    model = genai.GenerativeModel('gemini-pro')
    
    prompt = f"""
    Create a team slide for "{company_name}" with fictional leadership team members.
    
    Return a JSON object with these fields:
    - title: A slide title for the team/leadership slide
    - team_members: An array of 4 team member objects, each containing:
      - name: The team member's name
      - role: Their role or title
    - tagline: A brief company tagline or mission statement
    
    Be professional and diverse. Format as valid JSON only.
    """
    
    response = model.generate_content(prompt)
    content = extract_json_from_response(response.text)
    
    # Ensure correct fields exist
    if not content.get("title"):
        content["title"] = f"{company_name} Leadership Team"
    
    # Ensure team_members exist and have the right structure
    if not content.get("team_members") or not isinstance(content.get("team_members"), list):
        content["team_members"] = [
            {"name": "Alex Johnson", "role": "Chief Executive Officer"},
            {"name": "Sam Washington", "role": "Chief Technology Officer"},
            {"name": "Jordan Smith", "role": "VP of Product"},
            {"name": "Taylor Rivera", "role": "VP of Customer Success"}
        ]
    
    # Ensure tagline exists
    if not content.get("tagline"):
        content["tagline"] = f"Building innovative solutions for tomorrow's challenges"
    
    return content

def extract_json_from_response(response_text: str) -> Dict[str, Any]:
    """Extract JSON from model response, handling various formats."""
    # If the response is already a dict, return it
    if isinstance(response_text, dict):
        return response_text
    
    # Handle string response
    try:
        # Look for JSON block in markdown code blocks
        if "```json" in response_text:
            json_text = response_text.split("```json")[1].split("```")[0].strip()
            return json.loads(json_text)
        
        # Look for JSON block with just code blocks
        elif "```" in response_text:
            json_text = response_text.split("```")[1].split("```")[0].strip()
            return json.loads(json_text)
            
        # Try to parse the entire response as JSON
        else:
            return json.loads(response_text.strip())
    except:
        try:
            # Try to extract just the JSON part with braces
            if "{" in response_text and "}" in response_text:
                json_text = response_text[response_text.find("{"):response_text.rfind("}")+1]
                return json.loads(json_text)
        except:
            pass
    
    # Return empty dict if all parsing fails
    print(f"Failed to parse JSON from response: {response_text[:100]}...")
    return {}
