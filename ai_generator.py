"""
ai_generator.py - Generate assessment questions and benchmark answers using AI.
Supports: Ollama (free, local) or Claude API (paid, better quality).
"""

import json
import requests
from typing import Optional


def check_ollama_available() -> bool:
    """Check if Ollama is running locally."""
    try:
        resp = requests.get("http://localhost:11434/api/tags", timeout=3)
        return resp.status_code == 200
    except Exception:
        return False


def get_ollama_models() -> list:
    """Get list of installed Ollama models."""
    try:
        resp = requests.get("http://localhost:11434/api/tags", timeout=3)
        if resp.status_code == 200:
            data = resp.json()
            return [m["name"] for m in data.get("models", [])]
    except Exception:
        pass
    return []


def check_claude_api(api_key: str) -> bool:
    """Verify a Claude API key works."""
    try:
        resp = requests.post(
            "https://api.anthropic.com/v1/messages",
            headers={
                "x-api-key": api_key,
                "anthropic-version": "2023-06-01",
                "content-type": "application/json"
            },
            json={
                "model": "claude-sonnet-4-20250514",
                "max_tokens": 10,
                "messages": [{"role": "user", "content": "Hi"}]
            },
            timeout=10
        )
        return resp.status_code == 200
    except Exception:
        return False


def _build_question_prompt(unit_code: str, unit_title: str, knowledge_evidence: list, 
                           performance_evidence: list, assessment_conditions: str) -> str:
    """Build the prompt for generating assessment questions."""
    ke_text = "\n".join([f"- {ke}" for ke in knowledge_evidence])
    pe_text = "\n".join([f"- {pe}" for pe in performance_evidence[:10]])
    
    return f"""You are a VET (Vocational Education and Training) assessment developer in Australia.

Generate assessment questions for the unit: {unit_code} — {unit_title}

KNOWLEDGE EVIDENCE (each item must be covered by at least one question):
{ke_text}

PERFORMANCE EVIDENCE (for context):
{pe_text}

RULES:
1. Create one question per knowledge evidence item (or group closely related items into one question).
2. Questions should be SPECIFIC and PRACTICAL — not generic "describe/explain" prompts.
3. Include scenario-based questions where appropriate (e.g., "A customer brings in a vehicle with...")
4. Include questions that require calculations, measurements, or specific values where relevant.
5. Each question must have a BENCHMARK ANSWER — the answer an assessor would use to judge the response.
6. Benchmark answers should include specific technical details, not vague descriptions.
7. Group related knowledge evidence items into single questions where it makes sense.

OUTPUT FORMAT — respond with ONLY a JSON array, no other text:
[
  {{
    "number": 1,
    "question": "The specific question text",
    "benchmark": "The benchmark answer with specific technical details",
    "ke_items": ["knowledge evidence item 1 this covers", "knowledge evidence item 2"]
  }},
  ...
]

Generate approximately {min(len(knowledge_evidence), 30)} questions. Respond with ONLY the JSON array."""


def _build_mapping_prompt(unit_code: str, questions: list, elements: list, 
                          knowledge_evidence: list) -> str:
    """Build prompt for mapping questions to performance criteria and knowledge evidence."""
    
    q_text = "\n".join([f"Q{q['number']}: {q['question'][:100]}" for q in questions])
    
    el_text = ""
    for el in elements:
        el_text += f"\nElement {el.number}: {el.title}\n"
        for pc in el.performance_criteria:
            el_text += f"  {pc.number} {pc.text}\n"
    
    ke_text = "\n".join([f"- {ke}" for ke in knowledge_evidence])
    
    return f"""Map these assessment questions to the unit's Performance Criteria and Knowledge Evidence.

QUESTIONS:
{q_text}

ELEMENTS AND PERFORMANCE CRITERIA:
{el_text}

KNOWLEDGE EVIDENCE:
{ke_text}

For each question, identify which Knowledge Evidence items it covers.
Respond with ONLY a JSON array:
[
  {{
    "question_number": 1,
    "ke_items_covered": ["exact text of KE item 1", "exact text of KE item 2"]
  }},
  ...
]

Respond with ONLY the JSON array."""


def generate_with_ollama(prompt: str, model: str = "llama3.1:8b", 
                         progress_callback=None) -> Optional[str]:
    """Generate text using Ollama."""
    try:
        if progress_callback:
            progress_callback(f"Generating with Ollama ({model})... this may take a few minutes")
        
        resp = requests.post(
            "http://localhost:11434/api/generate",
            json={
                "model": model,
                "prompt": prompt,
                "stream": False,
                "options": {
                    "temperature": 0.3,
                    "num_predict": 8000,
                }
            },
            timeout=300  # 5 min timeout for slow models
        )
        
        if resp.status_code == 200:
            return resp.json().get("response", "")
        else:
            return None
    except Exception as e:
        print(f"Ollama error: {e}")
        return None


def generate_with_claude(prompt: str, api_key: str, 
                         progress_callback=None) -> Optional[str]:
    """Generate text using Claude API."""
    try:
        if progress_callback:
            progress_callback("Generating with Claude API...")
        
        resp = requests.post(
            "https://api.anthropic.com/v1/messages",
            headers={
                "x-api-key": api_key,
                "anthropic-version": "2023-06-01",
                "content-type": "application/json"
            },
            json={
                "model": "claude-sonnet-4-20250514",
                "max_tokens": 8000,
                "messages": [{"role": "user", "content": prompt}]
            },
            timeout=300
        )
        
        if resp.status_code == 200:
            data = resp.json()
            return data["content"][0]["text"]
        else:
            print(f"Claude API error: {resp.status_code} {resp.text[:200]}")
            return None
    except Exception as e:
        print(f"Claude error: {e}")
        return None


def _parse_json_response(text: str) -> list:
    """Extract JSON array from AI response, handling markdown fences etc."""
    text = text.strip()
    
    # Remove markdown code fences
    if text.startswith("```"):
        lines = text.split("\n")
        lines = [l for l in lines if not l.strip().startswith("```")]
        text = "\n".join(lines)
    
    # Find the JSON array
    start = text.find("[")
    end = text.rfind("]")
    if start >= 0 and end > start:
        try:
            return json.loads(text[start:end+1])
        except json.JSONDecodeError:
            pass
    
    # Try the whole thing
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        pass
    
    return []


def generate_questions(unit_data, provider: str = "ollama", api_key: str = "",
                       model: str = "llama3.1:8b", progress_callback=None) -> list:
    """
    Generate assessment questions with benchmark answers.
    
    Args:
        unit_data: UnitData object
        provider: "ollama" or "claude"
        api_key: Claude API key (only needed for claude provider)
        model: Ollama model name (only for ollama provider)
        progress_callback: function to report progress
    
    Returns:
        List of dicts with keys: number, question, benchmark, ke_items
    """
    # Build knowledge evidence as plain strings
    ke_strings = []
    for ke in unit_data.knowledge_evidence:
        text = ke.text if hasattr(ke, 'text') else str(ke)
        if text and len(text) > 3:
            ke_strings.append(text)
    
    pe_strings = unit_data.performance_evidence or []
    
    if not ke_strings:
        return []
    
    prompt = _build_question_prompt(
        unit_data.code, unit_data.title, 
        ke_strings, pe_strings,
        unit_data.assessment_conditions
    )
    
    if progress_callback:
        progress_callback(f"Generating {len(ke_strings)} questions using {provider}...")
    
    if provider == "claude":
        raw = generate_with_claude(prompt, api_key, progress_callback)
    else:
        raw = generate_with_ollama(prompt, model, progress_callback)
    
    if not raw:
        if progress_callback:
            progress_callback("AI generation failed — check your connection/API key")
        return []
    
    questions = _parse_json_response(raw)
    
    if progress_callback:
        progress_callback(f"Generated {len(questions)} questions")
    
    # Ensure numbering
    for i, q in enumerate(questions):
        q["number"] = i + 1
    
    return questions
