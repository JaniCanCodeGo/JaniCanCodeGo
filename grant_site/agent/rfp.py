"""RFP/NOFO ingestion — parse a pasted funding announcement and write the
application to the funder's own questions.

extract_rfp(text) pulls out: application questions, page/word limits,
deadlines, and scoring criteria. build_rfp_response(...) then answers each
extracted question by routing it to the pipeline's content (need, program
design, budget, evaluation, capacity, sustainability) so every answer is
responsive to the funder's actual prompt. When the Claude API is available
(see narrative_ai.py) the answers are additionally rewritten to the specific
question wording and any stated limits.
"""

import re

QUESTION_LINE = re.compile(
    r"^\s*(?:\d+[\.\)]\s+|[A-Z][\.\)]\s+|[•\-\*]\s+|Q\d+[:.]?\s+)?"
    r"(.{15,400}?\?)\s*$", re.M)
IMPERATIVE = re.compile(
    r"^\s*(?:\d+[\.\)]\s+|[A-Z][\.\)]\s+|[•\-\*]\s+)?"
    r"((?:Describe|Explain|Provide|Identify|Summarize|List|Demonstrate|"
    r"Detail|Discuss|Outline|Include|State|Address)\b.{10,400}?[\.\:])\s*$",
    re.M | re.I)
LIMIT_RE = re.compile(
    r"(?:no more than|not to exceed|maximum(?: of)?|limited? to|up to)\s+"
    r"(\d{1,4})\s*(pages?|words?|characters?)", re.I)
DEADLINE_RE = re.compile(
    r"(?:due|deadline|submitted?|received)\D{0,40}?"
    r"((?:January|February|March|April|May|June|July|August|September|"
    r"October|November|December)\s+\d{1,2},?\s+\d{4}|\d{1,2}/\d{1,2}/\d{2,4})",
    re.I)
POINTS_RE = re.compile(r"^\s*(.{5,120}?)[\s\.\-–—]*\(?(\d{1,3})\s*"
                       r"(?:points?|pts\.?|%)\)?\s*$", re.M | re.I)

TOPIC_MAP = [
    (("need", "problem", "gap", "data", "population", "community",
      "statement of need", "target"), "need"),
    (("budget", "cost", "funds", "expenditure", "financial", "allocat"),
     "budget"),
    (("evaluat", "measure", "outcome", "metric", "assess", "success",
      "data collection", "report"), "evaluation"),
    (("sustain", "continu", "beyond the grant", "future funding"),
     "sustainability"),
    (("capacity", "experience", "qualification", "history", "staff",
      "organization", "track record", "board"), "capacity"),
    (("partner", "collaborat", "letter of support", "mou"), "partners"),
    (("goal", "objective", "activit", "project", "program", "plan",
      "implement", "timeline", "approach", "describe the"), "project"),
]


def _classify(question):
    q = question.lower()
    for keywords, topic in TOPIC_MAP:
        if any(k in q for k in keywords):
            return topic
    return "project"


def extract_rfp(text):
    """Parse an RFP/NOFO into questions, limits, deadlines, and scoring."""
    text = text or ""
    questions, seen = [], set()
    for match in list(QUESTION_LINE.finditer(text)) + \
            list(IMPERATIVE.finditer(text)):
        q = re.sub(r"\s+", " ", match.group(1)).strip()
        key = q.lower()
        if key not in seen:
            seen.add(key)
            questions.append({"question": q, "topic": _classify(q)})
    limits = [{"limit": int(n), "unit": u.lower().rstrip("s") + "s"}
              for n, u in LIMIT_RE.findall(text)][:10]
    deadlines = sorted(set(DEADLINE_RE.findall(text)))[:5]
    scoring = [{"criterion": c.strip(), "points": int(p)}
               for c, p in POINTS_RE.findall(text)
               if int(p) <= 100][:15]
    return {"questions": questions[:30], "limits": limits,
            "deadlines": deadlines, "scoring": scoring,
            "parsed_chars": len(text)}


def _answer_for_topic(topic, ctx):
    """Route a question topic to content assembled by earlier stages.

    ctx keys: org, need_text, project_text, budget_text, evaluation_text,
    capacity_text, sustainability_text, partners_text.
    """
    return {
        "need": ctx["need_text"],
        "budget": ctx["budget_text"],
        "evaluation": ctx["evaluation_text"],
        "sustainability": ctx["sustainability_text"],
        "capacity": ctx["capacity_text"],
        "partners": ctx.get("partners_text") or
        "Partnerships and letters of support are being assembled; see the "
        "readiness checklist for status.",
        "project": ctx["project_text"],
    }.get(topic, ctx["project_text"])


def build_rfp_response(rfp, ctx, funder_name=""):
    """Build a document answering each RFP question in order."""
    sections = []
    meta_lines = []
    if rfp["deadlines"]:
        meta_lines.append("Deadline(s) found in RFP: " +
                          "; ".join(rfp["deadlines"]))
    if rfp["limits"]:
        meta_lines.append("Stated limits: " + "; ".join(
            f"{l['limit']} {l['unit']}" for l in rfp["limits"]))
    if rfp["scoring"]:
        meta_lines.append("Scoring rubric found: " + "; ".join(
            f"{s['criterion']} ({s['points']} pts)" for s in rfp["scoring"]))
    sections.append((
        "About this response",
        ("Answers below are mapped to the funder's own questions extracted "
         "from the RFP text. Verify each answer against the RFP before "
         "submission and trim to the stated limits.\n\n" +
         "\n".join(meta_lines)) if meta_lines else
        "Answers below are mapped to the funder's own questions extracted "
        "from the RFP text. Verify each against the RFP before submission."))
    if not rfp["questions"]:
        sections.append((
            "No questions detected",
            "The parser found no application questions in the pasted RFP "
            "text. Paste the section of the RFP that lists the narrative "
            "questions (often called 'Application Questions', 'Narrative', "
            "or 'Selection Criteria') and re-run."))
    for i, q in enumerate(rfp["questions"], 1):
        sections.append((f"{i}. {q['question']}",
                         _answer_for_topic(q["topic"], ctx)))
    title = (f"{ctx['org'].get('name', 'Organization')} — Response to "
             f"{funder_name or 'Funder'} RFP")
    return {"title": title, "sections": sections}
