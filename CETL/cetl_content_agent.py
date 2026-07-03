#!/usr/bin/env python3
"""
Clear Enough To Lead (CETL) — Content Generation Agent

Public-facing project: a workgroup teaching perimenopausal women with ADHD
to build web projects and small Claude agents, paired with a podcast on
leadership, life, and stress management.

No API key needed. Tools/CLI build structured prompts; Claude generates
the content natively.

--- MCP (Claude Desktop or claude.ai) ---
Use cetl_mcp_server.py.

--- CLI ---
python3 cetl_content_agent.py curriculum --type lesson-plan --module html-basics
python3 cetl_content_agent.py curriculum --type exercise --module localhost-projects
python3 cetl_content_agent.py podcast --type episode-outline --topic "the burnout math"
python3 cetl_content_agent.py podcast --type show-notes --episode "The Bar Is Clear Enough"
python3 cetl_content_agent.py audience --type welcome-message
python3 cetl_content_agent.py host --type teaching-script --module mini-claude-agents
python3 cetl_content_agent.py promo --type social-post --platform instagram --about "workgroup enrollment open"
"""

import argparse
import sys
from datetime import datetime
from pathlib import Path

SCRIPT_DIR = Path(__file__).parent
OUTPUT_DIR = SCRIPT_DIR / "outputs"

# ---------------------------------------------------------------------------
# Persona — the public CETL voice
# ---------------------------------------------------------------------------

PERSONA = """You are the writing voice for "Clear Enough To Lead" (CETL) — a public workgroup and podcast founded by Jani, a manager navigating ADHD and perimenopause who teaches other women to build web projects and small Claude agents, and hosts a podcast on leadership, life, and stress management.

AUDIENCE: Perimenopausal women with ADHD who lead teams, run households, and are ready to build something for themselves. Assume intelligence and life experience. Assume zero prior coding background unless stated otherwise. Never assume low capability — these are competent adults who happen to have a brain and body doing two hard things at once.

VOICE RULES:
- Warm, direct, funny — never a wellness brochure, never a LinkedIn inspirational post
- No corporate buzzwords, no "honor your journey," no "lean in," no toxic positivity
- Confidence without condescension — never talk down to beginners
- Name the real obstacles (brain fog, hot flashes, imposter syndrome, "I don't have a tech background") without dwelling on them as excuses — they're just facts to plan around
- Humor is welcome and expected, but this is public-facing: keep it professional enough to put in front of an audience she doesn't know personally yet
- Short sentences. Real talk. No fluff paragraphs before getting to the point.

CONTENT MODES:
- CURRICULUM content: clear, step-by-step, ADHD-friendly (short sections, one concept at a time, explicit "what you'll have built" at the end)
- PODCAST content: conversational, structured but not stiff, built for spoken delivery
- AUDIENCE-FACING content: warm, encouraging, respects their time and intelligence
- HOST-FACING content: written for Jani herself to use while teaching or recording — can be more casual/sarcastic since it's her own prep notes, not public copy
- PROMO content: clear value proposition, no hype-speak, sounds like a real person wrote it"""


# ---------------------------------------------------------------------------
# Prompt templates
# ---------------------------------------------------------------------------

CURRICULUM_PROMPTS = {
    "lesson-plan": """Write a lesson plan for a CETL "Build With Claude" module.

Module: {module}
Additional context: {context}

Requirements:
- One-sentence "what you'll have built by the end" at the very top
- Broken into short sections (5-10 minutes each) — no wall-of-text sections
- Each section: one concept, one small win
- Explicitly name at least one thing that commonly goes wrong at this stage, and how to recover from it (not just "here's how to avoid it")
- Ends with a recap and a one-line bridge to the next module
- Assume zero prior coding background unless the module context says otherwise""",

    "exercise": """Write a hands-on exercise for a CETL "Build With Claude" module.

Module: {module}
Additional context: {context}

Requirements:
- One clear deliverable — something the learner can point to and say "I built this"
- Numbered steps, one action per step
- Call out the exact moment where learners commonly get stuck, with a specific fix
- A "if you get stuck, ask Claude this" prompt suggestion at each tricky step
- A stretch goal at the end for anyone who finishes early (never required)""",

    "troubleshooting-faq": """Write a troubleshooting FAQ for common issues in a CETL "Build With Claude" module.

Module: {module}
Additional context: {context}

Requirements:
- Format: Q&A pairs, most common issue first
- Plain-language explanation of WHY the error happens, not just the fix
- No jargon without a one-line translation
- Include at least one "this isn't you doing something wrong, this happens to everyone" reassurance
- 5-8 Q&A pairs""",
}

PODCAST_PROMPTS = {
    "episode-outline": """Write an episode outline for the "Clear Enough To Lead" podcast.

Topic: {topic}
Additional context: {context}

Requirements:
- Cold open hook — one or two sentences that pull the listener in immediately
- 3-5 segment breakdown with a one-line description of what's covered in each
- At least one moment flagged for a personal story or example from Jani's own experience
- A close that ties back to the show's core theme (leadership + life + stress management)
- Suggested episode length
- 3-5 discussion questions if this episode includes a guest""",

    "show-notes": """Write show notes for a "Clear Enough To Lead" podcast episode.

Episode: {episode}
Additional context: {context}

Requirements:
- One-paragraph episode summary (2-3 sentences, hook-forward, written to make someone want to click play)
- 3-5 bullet "what you'll hear" points
- Any resources/links mentioned (placeholder brackets if unspecified)
- A pull-quote — one line from the episode that works as a standalone quote for social media""",

    "interview-questions": """Write interview questions for a guest on the "Clear Enough To Lead" podcast.

Guest/topic context: {context}

Requirements:
- 8-10 questions moving from warm-up to substantive to reflective
- At least 2 questions that connect the guest's experience to ADHD, perimenopause, or leadership under real-life constraints
- One closing question that gives the guest room to plug their own work""",
}

AUDIENCE_PROMPTS = {
    "welcome-message": """Write a welcome message for a new member joining the CETL "Build With Claude" workgroup.

Additional context: {context}

Requirements:
- Warm, direct, no hand-holding tone that implies she can't handle this
- Sets expectations: what the workgroup is, what it isn't, how much time it takes
- Names that some sessions will be harder than others and that's normal, not a sign she's behind
- Ends with a clear "here's your very first step" instruction""",

    "encouragement-note": """Write a short encouragement note for CETL workgroup members.

Situation/context: {context}

Requirements:
- Acknowledges a specific real struggle (stuck on an exercise, missed a session, feeling behind) without being saccharine
- No toxic positivity — validate that it's genuinely hard sometimes
- Ends with one small, concrete next action, not vague encouragement
- Under 150 words""",

    "assignment-recap": """Write a recap message for CETL workgroup members after a session or exercise.

Session/module: {context}

Requirements:
- What was covered, in plain language
- What they should have working/built by now
- One optional stretch task for anyone with energy left
- A note on what's coming next session""",
}

HOST_PROMPTS = {
    "teaching-script": """Write teaching talking points for Jani to use while leading a CETL "Build With Claude" session.

Module: {module}
Additional context: {context}

Requirements:
- Written in first person, for Jani to speak from — casual, can include her own sarcasm/humor
- Structured as talking points, not a word-for-word script — bullet phrases she can glance at
- Flags natural pause points for questions
- Includes at least one honest personal aside she can share (framed as an example she can adapt, not a fixed script)
- A note on timing for each section""",

    "podcast-intro-script": """Write an opening monologue script for Jani to record for a "Clear Enough To Lead" podcast episode.

Episode topic: {topic}
Additional context: {context}

Requirements:
- First person, in Jani's voice — direct, funny, real
- Under 200 words
- Sets up the episode topic without giving away the best material
- Ends with a natural transition into the first segment""",
}

PROMO_PROMPTS = {
    "social-post": """Write a social media post promoting CETL (the workgroup and/or podcast).

Platform: {platform}
What this post is about: {about}
Additional context: {context}

Requirements:
- Matches the platform's normal tone and length conventions
- Leads with the real value, not hype language
- Sounds like a real person wrote it, not a marketing team
- Includes a clear, low-pressure call to action
- No more than 2 emoji, and only if genuinely fitting the platform""",

    "newsletter": """Write a newsletter update for CETL subscribers.

What this update is about: {about}
Additional context: {context}

Requirements:
- Direct subject line, no clickbait
- Opens with the actual news/update in the first sentence
- Personal, first-person voice from Jani
- Ends with one clear next step for the reader""",
}


# ---------------------------------------------------------------------------
# Prompt builders
# ---------------------------------------------------------------------------

def _build(section: str, prompts: dict, content_type: str, **kwargs) -> str:
    if content_type not in prompts:
        return f"Unknown type '{content_type}' for '{section}'. Choose from: {', '.join(prompts)}"
    filled = {k: (v if v else "none provided") for k, v in kwargs.items()}
    request = prompts[content_type].format(**filled)
    return f"{PERSONA}\n\n---\n\n{request}"


def build_curriculum_prompt(content_type: str, module: str = "", context: str = "") -> str:
    return _build("curriculum", CURRICULUM_PROMPTS, content_type, module=module, context=context)


def build_podcast_prompt(content_type: str, topic: str = "", episode: str = "", context: str = "") -> str:
    return _build("podcast", PODCAST_PROMPTS, content_type, topic=topic, episode=episode, context=context)


def build_audience_prompt(content_type: str, context: str = "") -> str:
    return _build("audience", AUDIENCE_PROMPTS, content_type, context=context)


def build_host_prompt(content_type: str, module: str = "", topic: str = "", context: str = "") -> str:
    return _build("host", HOST_PROMPTS, content_type, module=module, topic=topic, context=context)


def build_promo_prompt(content_type: str, platform: str = "", about: str = "", context: str = "") -> str:
    return _build("promo", PROMO_PROMPTS, content_type, platform=platform, about=about, context=context)


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def save_prompt(prompt: str, section: str, content_type: str) -> Path:
    folder = OUTPUT_DIR / section
    folder.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = folder / f"{content_type.replace('-', '_')}_{timestamp}.txt"
    filename.write_text(prompt, encoding="utf-8")
    return filename


def main():
    parser = argparse.ArgumentParser(
        prog="cetl",
        description="Clear Enough To Lead — builds prompts for Claude to generate CETL content",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    sub = parser.add_subparsers(dest="section", required=True)

    cur = sub.add_parser("curriculum", help="Teaching curriculum content")
    cur.add_argument("--type", "-t", required=True, choices=list(CURRICULUM_PROMPTS.keys()))
    cur.add_argument("--module", "-m", default="")
    cur.add_argument("--context", "-c", default="")
    cur.add_argument("--no-save", action="store_true")

    pod = sub.add_parser("podcast", help="Podcast content")
    pod.add_argument("--type", "-t", required=True, choices=list(PODCAST_PROMPTS.keys()))
    pod.add_argument("--topic", default="")
    pod.add_argument("--episode", default="")
    pod.add_argument("--context", "-c", default="")
    pod.add_argument("--no-save", action="store_true")

    aud = sub.add_parser("audience", help="Content directed at workgroup members")
    aud.add_argument("--type", "-t", required=True, choices=list(AUDIENCE_PROMPTS.keys()))
    aud.add_argument("--context", "-c", default="")
    aud.add_argument("--no-save", action="store_true")

    host = sub.add_parser("host", help="Jani's own teaching/hosting prep content")
    host.add_argument("--type", "-t", required=True, choices=list(HOST_PROMPTS.keys()))
    host.add_argument("--module", "-m", default="")
    host.add_argument("--topic", default="")
    host.add_argument("--context", "-c", default="")
    host.add_argument("--no-save", action="store_true")

    promo = sub.add_parser("promo", help="Marketing and outreach content")
    promo.add_argument("--type", "-t", required=True, choices=list(PROMO_PROMPTS.keys()))
    promo.add_argument("--platform", "-p", default="")
    promo.add_argument("--about", "-a", default="")
    promo.add_argument("--context", "-c", default="")
    promo.add_argument("--no-save", action="store_true")

    args = parser.parse_args()

    builders = {
        "curriculum": lambda: build_curriculum_prompt(args.type, module=args.module, context=args.context),
        "podcast": lambda: build_podcast_prompt(args.type, topic=args.topic, episode=args.episode, context=args.context),
        "audience": lambda: build_audience_prompt(args.type, context=args.context),
        "host": lambda: build_host_prompt(args.type, module=args.module, topic=args.topic, context=args.context),
        "promo": lambda: build_promo_prompt(args.type, platform=args.platform, about=args.about, context=args.context),
    }
    prompt = builders[args.section]()

    print()
    print("=" * 60)
    print("  Clear Enough To Lead — Prompt Ready")
    print("=" * 60)
    print()
    print(prompt)
    print()

    if not args.no_save:
        out_path = save_prompt(prompt, args.section, args.type)
        print(f"  Saved to: {out_path}")
        print(f"  -> Paste this prompt into Claude to generate your content.")

    print("=" * 60)
    print()


if __name__ == "__main__":
    main()
