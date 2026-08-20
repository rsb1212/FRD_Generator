"""
Free AI Module for SOP + RGT → FRD
Supports:
  • Ollama (local, completely free)        →  http://localhost:11434
  • Google Gemini (free tier 1,500 req/day) →  GEMINI_API_KEY
  • OpenAI-compatible (OpenRouter, etc.)    →  OPENAI_API_KEY + OPENAI_BASE_URL
  • Local NLP fallback (zero API, zero GPU) →  sklearn + difflib
"""

import os
import re
import json
import difflib
import time
import requests
from typing import List, Dict, Optional
from dataclasses import dataclass

def get_gemini_key() -> str:
    return os.getenv("GEMINI_API_KEY", "")

def get_openai_key() -> str:
    return os.getenv("OPENAI_API_KEY", "")

def get_ollama_url() -> str:
    return os.getenv("OLLAMA_URL", "http://localhost:11434")

@dataclass
class GapItem:
    category: str
    severity: str
    description: str
    suggestion: str

@dataclass
class AIInsight:
    summary: str
    gaps: List[GapItem]
    enhanced_rules: List[str]
    risk_flags: List[str]

_ollama_cache = {'result': None, 'ts': 0}

def _ollama_available() -> bool:
    """Check if Ollama is reachable — cached for 60s to avoid blocking every page load."""
    now = time.monotonic()
    if now - _ollama_cache['ts'] < 60:
        return _ollama_cache['result']
    try:
        requests.get(f"{get_ollama_url()}/api/tags", timeout=1)
        _ollama_cache.update(result=True, ts=now)
        return True
    except Exception:
        _ollama_cache.update(result=False, ts=now)
        return False

def _call_ollama(prompt: str, system: str = "") -> str:
    try:
        r = requests.post(
            f"{get_ollama_url()}/api/generate",
            json={
                "model": os.getenv("OLLAMA_MODEL", "llama3.2"),
                "prompt": prompt,
                "system": system,
                "stream": False,
                "options": {"temperature": 0.2, "num_predict": 2048}
            },
            timeout=120
        )
        r.raise_for_status()
        return r.json().get("response", "")
    except Exception as e:
        return f"[Ollama Error: {e}]"

def _call_gemini(prompt: str) -> str:
    key = get_gemini_key()
    if not key:
        return "[Gemini API key not set]"
    models = ["gemini-2.0-flash", "gemini-2.5-flash", "gemini-flash-latest"]
    last_err = ""
    for model in models:
        try:
            url = f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent?key={key}"
            r = requests.post(
                url,
                json={
                    "contents": [{"parts": [{"text": prompt}]}],
                    "generationConfig": {"temperature": 0.2, "maxOutputTokens": 2048}
                },
                timeout=60
            )
            data = r.json()
            if r.status_code == 200 and "candidates" in data:
                return data["candidates"][0]["content"]["parts"][0]["text"]
            elif "error" in data:
                msg = data["error"].get("message", r.text)
                last_err = f"Gemini Error ({r.status_code}): {msg}"
                if r.status_code == 429:
                    # Rate/quota limit hit
                    break
        except Exception as e:
            last_err = str(e)
    return f"[{last_err}]"

def _call_openai(prompt: str, system: str = "") -> str:
    key = get_openai_key()
    if not key:
        return "[OpenAI-compatible key not set]"
    
    base = os.getenv("OPENAI_BASE_URL", "https://api.openai.com/v1")
    is_azure = "openai.azure.com" in base.lower()
    
    url = base
    if not is_azure and not url.endswith("/chat/completions"):
        url = f"{base}/chat/completions"
        
    headers = {"Content-Type": "application/json"}
    if is_azure:
        headers["api-key"] = key
    else:
        headers["Authorization"] = f"Bearer {key}"
        
    payload = {
        "messages": [
            {"role": "system", "content": system},
            {"role": "user", "content": prompt}
        ],
        "temperature": 0.2,
        "max_tokens": 2048
    }
    
    if not is_azure:
        payload["model"] = os.getenv("OPENAI_MODEL", "gpt-3.5-turbo")

    try:
        r = requests.post(
            url,
            headers=headers,
            json=payload,
            timeout=60
        )
        r.raise_for_status()
        return r.json()["choices"][0]["message"]["content"]
    except Exception as e:
        return f"[OpenAI Error: {e}]"

def _llm(prompt: str, system: str = "You are a meticulous business analyst.", provider: str = "gemini") -> str:
    """Route to the chosen LLM provider."""
    if provider == "gemini":
        if get_gemini_key():
            res = _call_gemini(prompt)
            if not res.startswith("["):
                return res
            return res
        return "[Gemini API key not set]"
        
    elif provider == "openai":
        if get_openai_key():
            res = _call_openai(prompt, system)
            if not res.startswith("["):
                return res
            return res
        return "[OpenAI API key not set]"

    return "[Invalid AI provider configured.]"


def _local_gap_analysis(sop_text: str, rgt_text: str) -> AIInsight:
    def sentences(txt):
        return [s.strip() for s in re.split(r'[.\n]+', txt) if len(s.strip()) > 15]
    sop_sents = sentences(sop_text)
    rgt_sents = sentences(rgt_text)
    gaps = []
    rgt_keywords = ["invoice", "eft", "auto-cancel", "dc activation", "account type",
                    "frequency calculator", "decimal round off", "nach", "ecs",
                    "auto-debit", "lapsed", "discontinue", "grace period"]
    for kw in rgt_keywords:
        in_rgt = kw in rgt_text.lower()
        in_sop = kw in sop_text.lower()
        if in_rgt and not in_sop:
            gaps.append(GapItem(
                category="Missing in SOP",
                severity="High" if kw in ["invoice", "eft", "auto-cancel"] else "Medium",
                description=f"Keyword '{kw}' appears in RGT but is absent from SOP.",
                suggestion=f"Add explicit SOP procedure covering '{kw}'."
            ))
        elif in_sop and not in_rgt:
            gaps.append(GapItem(
                category="Missing in RGT",
                severity="Low",
                description=f"SOP mentions '{kw}' but RGT does not formalize it as a requirement.",
                suggestion=f"Consider adding '{kw}' to RGT Delta / Business Requirements."
            ))
    for rs in rgt_sents[:30]:
        best = max(difflib.SequenceMatcher(None, rs, ss).ratio() for ss in sop_sents[:60])
        if 0.4 < best < 0.75:
            gaps.append(GapItem(
                category="Mismatch / Ambiguity",
                severity="Medium",
                description=f"RGT sentence partially matches SOP but may conflict: \"{rs[:100]}...\"",
                suggestion="Review alignment between RGT requirement and SOP procedure."
            ))
    summary = (
        f"Local NLP analysis found {len(gaps)} potential gaps between SOP and RGT. "
        f"{sum(1 for g in gaps if g.severity=='High')} high-severity items require attention."
    )
    return AIInsight(
        summary=summary,
        gaps=gaps,
        enhanced_rules=[],
        risk_flags=["Analysis performed without LLM — consider enabling Ollama or Gemini for deeper insights."]
    )

def compare_documents(sop_text: str, rgt_text: str, use_ai: bool = True) -> AIInsight:
    if not use_ai or not (_ollama_available() or get_gemini_key() or get_openai_key()):
        return _local_gap_analysis(sop_text, rgt_text)
    prompt = f"""You are a senior business analyst performing a gap analysis between an SOP and an RGT.

### SOP CONTENT:
{sop_text[:6000]}

### RGT CONTENT:
{rgt_text[:6000]}

### TASK:
1. Identify requirements present in RGT but missing or unclear in SOP.
2. Identify procedures in SOP not backed by RGT requirements.
3. Flag any contradictions or ambiguous alignments.
4. Suggest specific improvements.

Return STRICT JSON with this exact structure:
{{
  "summary": "short executive summary",
  "gaps": [
    {{
      "category": "Missing in SOP" | "Missing in RGT" | "Mismatch",
      "severity": "High" | "Medium" | "Low",
      "description": "detailed description",
      "suggestion": "actionable fix"
    }}
  ],
  "enhanced_rules": ["improved business rule 1", "improved business rule 2"],
  "risk_flags": ["risk 1", "risk 2"]
}}
Return ONLY the JSON. No markdown fences."""
    raw = _llm(prompt, system="You output only valid JSON.", provider="gemini")
    raw = raw.strip()
    if raw.startswith("```"):
        raw = raw.split("```", 2)[-1]
        if raw.startswith("json"):
            raw = raw[4:]
    raw = raw.strip()
    try:
        data = json.loads(raw)
        gaps = [GapItem(**g) for g in data.get("gaps", [])]
        return AIInsight(
            summary=data.get("summary", "AI analysis complete."),
            gaps=gaps,
            enhanced_rules=data.get("enhanced_rules", []),
            risk_flags=data.get("risk_flags", [])
        )
    except Exception:
        fallback = _local_gap_analysis(sop_text, rgt_text)
        fallback.risk_flags.append(f"LLM parsing failed. Raw preview: {raw[:200]}")
        return fallback

def enhance_frd(frd_data: dict, sop_text: str, rgt_text: str) -> dict:
    if not (_ollama_available() or get_gemini_key() or get_openai_key()):
        frd_data["ai_note"] = "AI enhancement skipped — no provider configured."
        return frd_data
    prompt = f"""You are a technical writer specializing in insurance IT FRDs.

### CURRENT FRD:
{json.dumps({k: v for k, v in frd_data.items() if k not in ('test_scenarios', 'delta_features')}, indent=2)[:5000]}

### SOURCE SOP SNIPPET:
{sop_text[:2000]}

### SOURCE RGT SNIPPET:
{rgt_text[:2000]}

### TASK:
Rewrite the following FRD fields to be more precise, professional, and comprehensive. Return STRICT JSON:
{{
  "objective": "...",
  "current_process": "...",
  "future_process": "...",
  "enhanced_business_rules": ["rule 1", "rule 2", ...],
  "out_of_scope_additions": ["item 1", "item 2"],
  "assumptions": ["assumption 1", "assumption 2"],
  "compliance_notes": "..."
}}
Return ONLY JSON."""
    raw = _llm(prompt, system="You output only valid JSON.", provider="gemini")
    raw = raw.strip()
    if raw.startswith("```"):
        raw = raw.split("```", 2)[-1]
        if raw.startswith("json"):
            raw = raw[4:]
    raw = raw.strip()
    try:
        data = json.loads(raw)
        frd_data["objective"] = data.get("objective", frd_data.get("objective", ""))
        frd_data["current_process"] = data.get("current_process", frd_data.get("current_process", ""))
        frd_data["future_process"] = data.get("future_process", frd_data.get("future_process", ""))
        if data.get("enhanced_business_rules"):
            frd_data["business_rules"] = data["enhanced_business_rules"]
        if data.get("out_of_scope_additions"):
            frd_data["out_of_scope"].extend(data["out_of_scope_additions"])
        frd_data["assumptions"] = data.get("assumptions", [])
        frd_data["compliance_notes"] = data.get("compliance_notes", "")
        frd_data["ai_enhanced"] = True
    except Exception as e:
        frd_data["ai_note"] = f"AI enhancement attempted but parsing failed: {e}"
    return frd_data

def generate_ai_frd(sop_text: str, rgt_text: str, provider: str = "gemini") -> dict:
    if provider == "gemini" and not get_gemini_key():
        return {"error": "Gemini API key not configured."}
    if provider == "openai" and not get_openai_key():
        return {"error": "OpenAI API key not configured."}
    prompt = f"""You are an expert business analyst in the insurance domain (Bajaj Allianz Life).

### SOP:
{sop_text[:5000]}

### RGT:
{rgt_text[:5000]}

### TASK:
Generate a complete FRD as STRICT JSON with these exact keys:
- call_id (string, format CR-FREQ-YYYY-MMDD)
- call_summary (string)
- objective (string)
- business_requirement (list of strings)
- applicable_platform (string)
- priority (string)
- stakeholders (list of strings)
- current_process (string)
- future_process (string)
- dependencies (list of strings)
- out_of_scope (list of strings)
- business_rules (list of strings)
- process_json (object representing a flowchart, containing: "process_name" (string), "nodes" (list of objects: id, actor, type (start/process/decision/end), label), and "connections" (list of objects: from, to, label (optional)))
- assumptions (list of strings)
- compliance_notes (string)
- risk_flags (list of strings)

Return ONLY JSON. No markdown."""
    raw = _llm(prompt, system="You output only valid JSON.", provider=provider)
    raw = raw.strip()
    if raw.startswith("```"):
        raw = raw.split("```", 2)[-1]
        if raw.startswith("json"):
            raw = raw[4:]
    raw = raw.strip()
    try:
        return json.loads(raw)
    except Exception as e:
        return {"error": f"AI generation failed: {e}", "raw_preview": raw[:500]}

_ai_status_cache = {'result': None, 'ts': 0}

def ai_status() -> dict:
    """Return AI provider availability — cached for 30s."""
    now = time.monotonic()
    if _ai_status_cache['result'] is not None and now - _ai_status_cache['ts'] < 30:
        return _ai_status_cache['result']
    gemini_key = get_gemini_key()
    openai_key = get_openai_key()
    result = {
        "ollama": {"available": _ollama_available(), "url": get_ollama_url(), "model": os.getenv("OLLAMA_MODEL", "llama3.2")},
        "gemini": {"available": bool(gemini_key), "key_set": bool(gemini_key)},
        "openai_compatible": {"available": bool(openai_key), "base_url": os.getenv("OPENAI_BASE_URL", "https://api.openai.com/v1")},
        "local_nlp": {"available": True, "note": "Always available as fallback"}
    }
    _ai_status_cache.update(result=result, ts=now)
    return result

