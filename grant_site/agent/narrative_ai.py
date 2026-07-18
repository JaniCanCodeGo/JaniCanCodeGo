"""Optional Claude-powered prose tailoring.

When the `anthropic` package is installed and credentials are available
(ANTHROPIC_API_KEY or an `ant auth login` profile), each generated document
is rewritten by Claude into tailored, competitive grant prose — instead of
template merge — while preserving every fact and figure. Without
credentials the pipeline runs unchanged and notes that tailoring was
skipped.
"""

import os
import re

try:
    import anthropic
    HAS_ANTHROPIC = True
except ImportError:
    HAS_ANTHROPIC = False

MODEL = "claude-opus-4-8"

SYSTEM = (
    "You are an expert grant writer. Rewrite the draft document you are "
    "given into polished, persuasive, funder-ready prose.\n"
    "Rules:\n"
    "- Preserve every numeric figure, dollar amount, source citation, EIN, "
    "URL, and legal statement exactly; never invent statistics, partners, "
    "or outcomes.\n"
    "- Keep the same section structure: output Markdown with a single '# ' "
    "title line followed by '## ' section headings matching the draft's "
    "sections (you may improve heading wording slightly).\n"
    "- Write in confident, specific, active prose; remove template "
    "boilerplate; weave the applicant's local data into the needs "
    "narrative.\n"
    "- If a draft section contains a placeholder or '(complete before "
    "filing)', keep the placeholder visible — do not fill it with invented "
    "content.\n"
    "- Respect any page/word limits stated in the context.")


def available():
    if not HAS_ANTHROPIC:
        return False
    if os.environ.get("ANTHROPIC_API_KEY") or os.environ.get("ANTHROPIC_AUTH_TOKEN"):
        return True
    # An `ant auth login` profile also works; cheapest reliable check is to
    # try constructing the client and let the first call decide.
    return bool(os.environ.get("ANTHROPIC_PROFILE"))


def tailor_document(doc, context_notes=""):
    """Rewrite a {title, sections} document with Claude. Returns the new
    document, or None if tailoring is unavailable/failed (callers keep the
    template version)."""
    if not available():
        return None
    draft = f"# {doc['title']}\n\n" + "\n\n".join(
        f"## {h}\n\n{b}" for h, b in doc["sections"])
    prompt = ""
    if context_notes:
        prompt += f"Context about the applicant and funder:\n{context_notes}\n\n"
    prompt += f"Rewrite this draft:\n\n{draft}"
    try:
        client = anthropic.Anthropic()
        with client.messages.stream(
            model=MODEL,
            max_tokens=32000,
            thinking={"type": "adaptive"},
            system=SYSTEM,
            messages=[{"role": "user", "content": prompt}],
        ) as stream:
            message = stream.get_final_message()
        if message.stop_reason == "refusal":
            return None
        text = "".join(b.text for b in message.content if b.type == "text")
        return _parse_markdown(text, doc)
    except Exception:
        return None


def _parse_markdown(text, fallback_doc):
    title_match = re.search(r"^#\s+(.+)$", text, re.M)
    title = title_match.group(1).strip() if title_match else fallback_doc["title"]
    parts = re.split(r"^##\s+", text, flags=re.M)
    sections = []
    for part in parts[1:]:
        lines = part.split("\n", 1)
        heading = lines[0].strip()
        body = (lines[1] if len(lines) > 1 else "").strip()
        if heading and body:
            sections.append((heading, body))
    if not sections:
        return None
    return {"title": title, "sections": sections}
