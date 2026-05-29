#!/usr/bin/env python3
"""
Clear Enough To Lead (CETL) — Content Generation Agent

A writing tool for a middle-aged Black American woman managing a team
while navigating ADHD and perimenopause. Voice: humorous, sarcastic, real.

No API key needed. Prompts are built here; Claude generates the content.

--- MCP (Claude Desktop) ---
Use cetl_mcp_server.py. Claude calls the tools, gets the prompt, and responds.

--- CLI ---
The CLI builds a ready-to-use prompt and saves it to outputs/.
Copy the prompt into Claude.ai (or any Claude interface) to generate the content.

Usage:
  # Content directed at others (team, boss, HR)
  python3 cetl_agent.py for-others --type email --situation "team missed the deadline"
  python3 cetl_agent.py for-others --type feedback --situation "employee keeps talking over people"
  python3 cetl_agent.py for-others --type meeting-agenda --topic "Q3 retro" --duration 45
  python3 cetl_agent.py for-others --type boundary --situation "manager keeps scheduling over my focus blocks"

  # Content for yourself
  python3 cetl_agent.py for-self --type brain-dump --input "I have 47 things to do and I can't remember any of them"
  python3 cetl_agent.py for-self --type fog-recovery
  python3 cetl_agent.py for-self --type morning-check --level low --symptoms "brain fog, hot flash at 3am"
  python3 cetl_agent.py for-self --type self-advocacy --situation "I need a standing desk and flexible start time"
  python3 cetl_agent.py for-self --type energy-plan --level medium --symptoms "mild fog" --commitments "10am standup, 2pm 1:1s"
"""

import argparse
import sys
from datetime import datetime
from pathlib import Path

SCRIPT_DIR = Path(__file__).parent
OUTPUT_DIR = SCRIPT_DIR / "outputs"

# ---------------------------------------------------------------------------
# Persona
# ---------------------------------------------------------------------------

PERSONA = """You are the writing voice for "Clear Enough To Lead" — a content tool built for a middle-aged Black American woman who manages a team while navigating ADHD and perimenopause simultaneously.

Her voice is: direct, funny, sarcastic, smart, and completely done with corporate theater. She is not mean — she is real. She calls things what they are. She handles her business with excellence even on the days when she has walked into the kitchen three times and still doesn't know why, when a hot flash hit mid-presentation, or when she hyperfocused on one email for 45 minutes and forgot about the meeting she scheduled.

TONE RULES — never break these:
- No corporate buzzwords (unless you are actively making fun of them)
- No wellness-coach language ("honor your journey," "hold space," "lean in")
- No passive voice when active voice works fine
- No LinkedIn inspirational post energy
- No sugarcoating — but also no catastrophizing
- DO use humor, irony, and self-aware sarcasm
- DO write like a smart, experienced Black woman who has figured out most things but not all things, and is fine with that
- DO keep it tight — she doesn't have the attention span or time for a wall of text

FOR OTHERS content: Professional enough for the workplace. Her personality shows — direct, clear, respectful — but it doesn't read like a robot or a people-pleaser wrote it.

FOR SELF content: More personal, more sarcastic, fully acknowledges the reality of symptoms. Treats her like the intelligent adult she is."""


# ---------------------------------------------------------------------------
# Prompt templates
# ---------------------------------------------------------------------------

FOR_OTHERS_PROMPTS = {
    "email": """Write a workplace email for the following situation:

Situation: {situation}
Additional context: {context}

Requirements:
- Professional but with personality — not robotic, not a form letter
- Subject line that actually says what the email is about
- Gets to the point in the first sentence (no "Hope this finds you well")
- No filler phrases unless being used for strategic effect
- Length matches the situation: short if simple, longer only when necessary
- Sign-off that fits the tone

Format:
Subject: [subject line]

[email body]""",

    "feedback": """Write a feedback conversation script for a manager to have with an employee.

Situation: {situation}
Context: {context}

Requirements:
- Opening that doesn't lead with hollow praise before the real point
- Clear, direct statement of the behavior or accomplishment
- Show where the employee responds with: [Employee response]
- Concrete next steps or expectations — nothing vague
- Closing that maintains the person's dignity
- Stage directions in brackets where useful: [Pause], [Wait for response]
- Under 400 words — she doesn't have bandwidth for a monologue""",

    "meeting-agenda": """Create a meeting agenda that actually respects everyone's time and attention span.

Topic: {topic}
Duration: {duration} minutes
Context (attendees, purpose, background): {context}

Requirements:
- A one-sentence "why are we here" at the top — no one should have to guess
- Every agenda item has: owner, time limit, and expected outcome (discussion vs. decision vs. update)
- Hard time boxes — build in 5-minute buffer because something always runs long
- A "Parking Lot" section for things that come up but aren't on the agenda
- Action items section at the end: item | owner | deadline
- A firm but non-mean note that late arrivals will not get a recap
- No item over 15 minutes without a break or transition""",

    "boundary": """Write a boundary-setting script for the following situation.

Situation: {situation}
Context: {context}

Requirements:
- Direct — no apologizing for having the boundary
- States the what without over-explaining the why
- Offers a clear alternative or path forward where appropriate
- Firm but not hostile

Provide TWO versions:
1. Version 1 — says it once, clearly
2. Version 2 — says it again when they didn't hear it the first time (because they won't)

Include both a verbal script and an email version where appropriate.""",
}

FOR_SELF_PROMPTS = {
    "brain-dump": """Take this brain dump and turn it into something that can actually be acted on.

Brain dump: {input_text}

Output format:
1. One-sentence acknowledgment of the situation (honest, a little funny)
2. URGENT — do today or something actually explodes
3. IMPORTANT — needs to happen this week
4. LOW STAKES — whenever
5. WAITING ON SOMEONE ELSE — not your problem right now, document it and move on
6. DELETE THIS THOUGHT — things that don't actually need doing (be ruthless)
7. One closing line: encouraging without being obnoxious about it""",

    "fog-recovery": """She's lost the thread. Generate a fog recovery protocol.

Context: {context}

Requirements:
- Physical steps to stop and reorient — real and specific, not "take a deep breath"
- The "retrace your last 3 steps" technique with actual instructions
- A short list of likely suspects (things she was probably doing)
- A reminder that this is a documented neurological symptom, not a sign she's losing it
- Tone: patient, mildly sarcastic best friend — not a neurologist""",

    "morning-check": """Create a morning check-in routine for a day when symptoms are real and she still has to show up.

Energy/symptom level: {level}
Specific symptoms or context: {symptoms}

Requirements:
- Under 10 items total
- Start with physical stabilization (water, food, temperature — perimenopause-aware)
- One "anchor task" — the single thing that must happen today if nothing else does
- One "grace item" — one thing she can explicitly let slide today without consequence
- A sentence she can say to herself before opening Slack or email
- Completable in 5 minutes or less
- Realistic, not aspirational — survival checklist, not a productivity system""",

    "self-advocacy": """Write a script for requesting a workplace accommodation or adjustment.

What she needs: {situation}
Context (boss, HR, workplace culture): {context}

Requirements:
- She does not owe anyone a detailed medical explanation
- Clear, specific statement of what she's requesting
- Brief professional framing — enough to be taken seriously, not so much she's oversharing
- Mentions her track record and competence (write it without making her cringe)
- Response prepared for: "But everyone else manages fine"
- Firm and collaborative, not apologetic and not confrontational

Provide:
1. Verbal script (conversation with her manager)
2. Written version (email to HR or manager, with subject line)""",

    "energy-plan": """Build a realistic work plan for today based on where she actually is right now.

Energy level: {level} (low/medium/high)
Current symptoms or physical state: {symptoms}
Today's known commitments: {commitments}

Requirements:
- Prioritize by both importance AND energy cost
- Low energy: protect the brain for ONE hard thing; fill the rest with low-lift tasks
- Medium energy: normal load with actual scheduled breaks
- High energy: good day — don't waste it all on Slack
- One "protect this block" — a time she should defend from being scheduled over
- Specific food/water timing notes (brain and hormone support, not diet culture)
- Honest about what's not happening today if it's a low day
- Permission slip at the end: one sentence releasing her from guilt about the adjusted plan""",
}


# ---------------------------------------------------------------------------
# Prompt builders (used by both CLI and MCP server)
# ---------------------------------------------------------------------------

def build_for_others_prompt(content_type: str, situation: str = "", context: str = "",
                             topic: str = "", duration: int = 60) -> str:
    if content_type not in FOR_OTHERS_PROMPTS:
        return f"Unknown type '{content_type}'. Choose from: {', '.join(FOR_OTHERS_PROMPTS)}"

    request = FOR_OTHERS_PROMPTS[content_type].format(
        situation=situation or "(not specified)",
        context=context or "none provided",
        topic=topic or situation or "(not specified)",
        duration=duration,
    )
    return f"{PERSONA}\n\n---\n\n{request}"


def build_for_self_prompt(content_type: str, input_text: str = "", context: str = "",
                           situation: str = "", level: str = "medium",
                           symptoms: str = "", commitments: str = "") -> str:
    if content_type not in FOR_SELF_PROMPTS:
        return f"Unknown type '{content_type}'. Choose from: {', '.join(FOR_SELF_PROMPTS)}"

    request = FOR_SELF_PROMPTS[content_type].format(
        input_text=input_text or context or "(not provided — generate a general example)",
        context=context or "none provided",
        situation=situation or context or "(not specified)",
        level=level,
        symptoms=symptoms or "none specified",
        commitments=commitments or "none specified",
    )
    return f"{PERSONA}\n\n---\n\n{request}"


# ---------------------------------------------------------------------------
# CLI — builds the prompt and saves it; paste into Claude.ai to generate
# ---------------------------------------------------------------------------

def save_prompt(prompt: str, direction: str, content_type: str) -> Path:
    folder = OUTPUT_DIR / direction.replace("-", "_")
    folder.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = folder / f"{content_type.replace('-', '_')}_{timestamp}.txt"
    filename.write_text(prompt, encoding="utf-8")
    return filename


def main():
    parser = argparse.ArgumentParser(
        prog="cetl",
        description="Clear Enough To Lead — builds prompts for Claude to generate your content",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    sub = parser.add_subparsers(dest="direction", required=True)

    fo = sub.add_parser("for-others", help="Content directed at your team, boss, or HR")
    fo.add_argument("--type", "-t", required=True, choices=list(FOR_OTHERS_PROMPTS.keys()))
    fo.add_argument("--situation", "-s", default="")
    fo.add_argument("--context", "-c", default="")
    fo.add_argument("--topic", default="")
    fo.add_argument("--duration", "-d", type=int, default=60)
    fo.add_argument("--no-save", action="store_true")

    fs = sub.add_parser("for-self", help="Content for managing your own brain and day")
    fs.add_argument("--type", "-t", required=True, choices=list(FOR_SELF_PROMPTS.keys()))
    fs.add_argument("--input", "-i", default="")
    fs.add_argument("--situation", "-s", default="")
    fs.add_argument("--context", "-c", default="")
    fs.add_argument("--level", "-l", choices=["low", "medium", "high"], default="medium")
    fs.add_argument("--symptoms", default="")
    fs.add_argument("--commitments", default="")
    fs.add_argument("--no-save", action="store_true")

    args = parser.parse_args()

    if args.direction == "for-others":
        prompt = build_for_others_prompt(
            content_type=args.type,
            situation=args.situation,
            context=args.context,
            topic=args.topic,
            duration=args.duration,
        )
    else:
        prompt = build_for_self_prompt(
            content_type=args.type,
            input_text=getattr(args, "input", ""),
            situation=args.situation,
            context=args.context,
            level=args.level,
            symptoms=args.symptoms,
            commitments=args.commitments,
        )

    print()
    print("=" * 60)
    print("  Clear Enough To Lead — Prompt Ready")
    print("=" * 60)
    print()
    print(prompt)
    print()

    if not args.no_save:
        out_path = save_prompt(prompt, args.direction, args.type)
        print(f"  Saved to: {out_path}")
        print(f"  → Paste this prompt into Claude.ai to generate your content.")

    print("=" * 60)
    print()


if __name__ == "__main__":
    main()
