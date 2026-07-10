# Clear Enough To Lead (CETL) — Business Plan

**Founder:** Jani
**Date:** July 2026
**Tagline:** *Leadership, life, and stress management for women who are managing everything anyway.*

> **How this plan was researched.** Market and evidence claims below were gathered from
> primary sources (peer-reviewed studies, platform data, industry reports) and each key
> claim was independently fact-checked before inclusion. Claims that failed verification
> (including some widely repeated ones, like the "$600 billion menopause economy" figure
> and "MOOCs complete at 3% vs. 75%+ for cohorts") were excluded or corrected. Sources are
> cited in numbered footnotes at the end.

---

## 1. Executive Summary

Clear Enough To Lead (CETL) is a two-track business serving one sharply defined customer:
**the perimenopausal woman with ADHD — typically 40–55, a professional who leads teams at
work and runs a household at home — who wants to build real technical skills for herself.**

- **Track 1 — "Build With Claude" (revenue engine):** a paid, cohort-based workgroup that
  takes members from zero coding background to building web pages, local projects, small
  interactive projects, and finally small working Claude agents. Four modules, live
  sessions, structured accountability, no gatekeeping.
- **Track 2 — the "Clear Enough To Lead" podcast (audience engine):** a show on
  leadership, life, and stress management for the same audience. In the early years the
  podcast's job is trust-building and cohort enrollment, not ad revenue.

**Why this business, and why now:**

1. **The pain is real, measurable, and expensive.** A Mayo Clinic study of 4,440 employed
   US women aged 45–60 found 13.4% reported an adverse work outcome due to menopause
   symptoms, with an estimated **$1.8 billion/year in lost US work time** ($26.6 billion
   including medical costs).[^1] For women with ADHD it is dramatically worse: in an
   ADDitude survey of 1,500+ women, **94% said their ADHD symptoms grew more severe during
   perimenopause and menopause**, and roughly 70% described brain fog, memory issues, and
   overwhelm as having a "life-altering impact" in their 40s and 50s.[^3]
2. **Nobody serves this intersection.** Women-focused coding schools exist (Skillcrush,
   SheCodes), ADHD coaching exists, menopause femtech exists — but no offering combines
   beginner-friendly technical education with ADHD-informed structure for midlife women.
   Only about 7% of femtech startups even focus on menopause.[^15]
3. **The delivery model is evidence-matched to the audience.** Cohort-based courses
   complete at meaningfully higher rates than self-paced ones (64.2% vs. 48.2% in one
   32,000-course platform analysis),[^16] and the accountability mechanisms that make
   cohorts work — live structure, peers, "body doubling" — are precisely the supports the
   ADHD literature recommends for adult learners.[^10][^11][^12]
4. **AI has collapsed the learning curve.** Teaching "build with Claude" rather than
   "memorize syntax" means beginners ship working projects in weeks, which is the
   confidence loop ADHD learners need.

**The business model:** premium cohort fees ($299 beta → $500–$950 per seat as the
program matures, in line with cohort-course pricing benchmarks[^18][^19]) plus a
lower-priced alumni community membership, with the podcast as the top of funnel. A
realistic Year 1 target is 3 cohorts and roughly **$20K–$35K revenue**, growing toward
**$100K+ in Year 3** as cohort size, price, and the membership base grow. CETL is
designed to be founder-run, low-overhead, and profitable at small scale — not
venture-scale, and deliberately so.

---

## 2. Company Overview

### Mission
Give midlife women with ADHD a judgment-free path to real technical capability — and a
voice (the podcast) that talks about leadership and stress without pretending the fog
doesn't happen.

### What CETL is not
- Not a coding bootcamp promising job placement.
- Not medical advice, ADHD treatment, or menopause care. CETL is education and community;
  the plan cites health research to size the problem, not to make clinical claims.
- Not a content mill. Small cohorts, taught live, by the founder.

### Founder
Jani — practitioner-founder who has lived the exact journey the curriculum teaches
(zero background → shipped web projects → built working Claude agents, including the
agent tooling in this repository), and who leads teams professionally. Founder-audience
fit is the moat: the teacher *is* the customer, five steps ahead.

### Legal & structure (to be completed)
Single-member LLC recommended at launch for liability separation and simple pass-through
taxes. Standard needs: business bank account, liability waiver + terms of service for
the workgroup, podcast release forms for guests, and a clear "education, not medical
advice" disclaimer given the health-adjacent audience.

---

## 3. The Problem — Researched Pain Points

### 3.1 Menopause measurably damages careers

- In the Mayo Clinic Proceedings study of 4,440 employed US women aged 45–60, **13.4%
  reported at least one adverse work outcome** due to menopause symptoms and 10.8% missed
  work in the prior year (median 3 days).[^1]
- The same study estimated **$1.8 billion/year lost in the US from missed work alone,
  $26.6 billion including medical expenses**.[^1][^2]
- Impact scales with severity: women in the highest symptom-severity quartile were
  **15.6× more likely** to report an adverse work outcome than those in the lowest.[^1]
- Career attrition is documented directly: in a Fawcett Society survey of UK women aged
  45–55, 10% had left a job because of menopause symptoms, 14% had cut hours, 8% had not
  applied for promotion, and 73% reported brain fog.[^6]

### 3.2 ADHD makes midlife dramatically harder — and is often diagnosed late

- ADHD in girls and women is under-recognized and under-researched, producing a pattern
  of **late diagnosis across the lifespan**.[^7] ADDitude's survey of 2,600+ women with
  ADHD found 43% were first diagnosed between ages 41–50, and 61% said ADHD had its
  greatest life impact between 40–59.[^4]
- Hormonal transitions — including perimenopause — **exacerbate ADHD symptoms and mood
  disturbance**, while tailored treatments remain lacking; the mechanism runs through
  estrogen's interaction with dopaminergic pathways that support executive function.[^7]
- In ADDitude's survey of 1,500+ women, **94% reported ADHD symptoms grew more severe in
  perimenopause/menopause**, ~70% said brain fog/memory/overwhelm had a "life-altering
  impact" in their 40s–50s (vs. 11% calling memory problems life-altering in their
  20s–30s), and 83% reported experiencing some ADHD symptoms for the first time during
  this transition.[^3]
- A 2025 Journal of Attention Disorders study of 656 women aged 45–60 found **higher ADHD
  symptom severity correlated with worse menopausal difficulties and poorer quality of
  life** on nearly all measures — and notably, it was symptom severity, not diagnosis
  status, that drove the effect, underscoring how many struggling women are undiagnosed
  or under-supported.[^8]

### 3.3 Traditional online learning fails this learner

- Self-paced online courses lose roughly half their students: one analysis of 32,000+
  courses found **48.2% completion for self-paced/open-access courses vs. 64.2% for
  cohort-based** ones,[^16] and Class Central puts the median MOOC completion rate at
  just 12.6%.[^16] (Widely repeated claims of "3–6% MOOC completion vs. 75%+ for
  cohorts" did not survive fact-checking and are deliberately not used here.)
- Adults with ADHD specifically do worse with unguided formats: RCT evidence on
  self-guided internet-delivered ADHD interventions shows weaker engagement and outcomes
  than guided ones.[^13]
- Coding education has a further problem for this audience: it is optimized for
  career-switchers in their 20s–30s, priced accordingly (Skillcrush's beginner package is
  $1,599–$2,499[^20]), and structured as long self-paced grinds — the exact format the
  completion data says fails.

**The synthesis:** a large, professionally senior, financially capable population is
hitting a wall of brain fog and executive dysfunction at the peak of their careers, being
diagnosed (or not) late, and being offered either generic self-paced courses that they
statistically won't finish or clinical services that don't teach them anything to build.
That's the gap CETL occupies.

---

## 4. Market Analysis

### 4.1 Market size (top-down)

| Layer | Size | Source |
|---|---|---|
| US workforce in some phase of menopause transition | ~¼ of the US workforce; 76.8% of women 45–54 are in the labor force | AARP[^5] |
| Femtech market (broader women's-health economy) | ~$52–63B in 2025, growing 11–15%/yr | Grand View Research[^14] |
| Menopause share of femtech innovation | only ~7% of femtech startups | PreScouter[^15] |
| ADHD-support software market (closest proxy for ADHD-support spend) | $1.9B (2025) → projected $6.7B by 2033 | Grand View Research[^17] |

A note on honesty: the popular "**$600 billion menopause economy**" figure failed
independent verification and is not used in this plan. The verified, defensible story is
narrower and better: a quarter of the workforce is affected,[^5] the cost of inaction is
quantified in the billions,[^1] and almost no one is building for it.[^15]

### 4.2 Serviceable market (bottom-up)

CETL doesn't need a percentage of a billion-dollar market; it needs cohorts of 10–20
women a few times a year. The bottom-up funnel:

- **Reachable audience:** women 40–55 with diagnosed or suspected ADHD who follow
  ADHD/menopause media. Single data point on scale: ADDitude alone surveyed 1,500–2,600+
  such women for the studies cited above[^3][^4] — the engaged community around these
  publications, podcasts, and social channels numbers in the hundreds of thousands.
- **Conversion math for viability:** 2,000 podcast listeners/email subscribers × 2%
  cohort conversion = 40 seats/year — full capacity for a founder-taught program.
  Viability does not depend on capturing a market; it depends on building a
  1,000–5,000-person audience of exactly the right women.

### 4.3 Ideal customer profile

- Woman, 40–55, in perimenopause or early menopause.
- ADHD — diagnosed (possibly recently) or strongly suspected.
- Leads people: manager, director, team lead, business owner, or the de-facto leader of
  her household's logistics.
- Household income sufficient for $500–$1,500 discretionary professional development;
  many can expense it (professional-topic courses around $500 are commonly employer-
  reimbursable[^19]).
- Emotional job-to-be-done: *"Prove to myself my brain still works — and build something
  that's mine."*

---

## 5. The Solution — and Why It's Evidence-Based

### 5.1 Product design principles, mapped to evidence

| CETL design choice | Evidence it rests on |
|---|---|
| **Live cohorts, not self-paced content** | Cohort courses complete at 64.2% vs. 48.2% self-paced across 32,000+ courses;[^16] Harvard's move of case-method courses online with peer collaboration reached 85% completion vs. single-digit MOOC rates of that era[^18] |
| **Co-working sessions ("body doubling") built into every module** | A controlled experiment found adults with ADHD completed tasks faster with greater perceived accuracy and sustained attention with a body double (human *or* AI) than alone, and preferred it;[^10] Cleveland Clinic describes body doubling as external structure that compensates for executive-function deficits[^11]. Honest caveat: the evidence base is young — prior work was inconclusive and some neurodivergent people find social presence uncomfortable,[^10] so co-working is offered, never mandatory |
| **Structured accountability, goal-setting, and check-ins** | A review of 19 studies (16 peer-reviewed) found ADHD coaching built on goal-setting, psychoeducation, and accountability improves outcomes and is increasingly recommended in multimodal adult ADHD care[^12] |
| **Guided, not self-guided** | Adults with ADHD show weaker outcomes and engagement in self-guided internet-delivered interventions vs. guided formats[^13] |
| **Claude as the always-available tutor** | Removes the #1 beginner failure mode (getting stuck alone at 9pm); normalizes AI-assisted building as the *method*, not a crutch — and AI body-doubling evidence suggests the "working with Claude" framing itself supports attention[^10] |
| **Small wins in week one** | Curriculum ships a visible artifact per module (a live page → a running localhost project → an interactive project → a working agent), designed for the ADHD reward cycle rather than delayed gratification |

### 5.2 The offer stack

1. **"Build With Claude" cohort (core, $299 beta → $500–$950).** 6–8 weeks, 10–20 seats,
   2 live sessions/week (one teach, one co-working), lifetime access to materials. Four
   modules (already defined in `data/curriculum.json`): HTML Basics → Localhost Projects
   → Before Apps → Building Mini Claude Agents.
2. **CETL Circle — alumni community (recurring, ~$19–29/mo).** Weekly body-doubling
   co-working calls, show-and-tell, next-project support. The hybrid "premium cohort +
   lower-priced ongoing membership" structure is a standard, recommended cohort-business
   model.[^19]
3. **Podcast (free).** Leadership, life, and stress management. Builds the audience the
   funnel math in §4.2 requires; monetizes lightly and late (see §8).
4. **Later (Year 2+):** a self-paced "starter" tier for the waitlist, corporate/ERG
   workshops ("AI literacy for women's employee resource groups"), 1:1 intensives.

---

## 6. Competitive Landscape

| Competitor / category | What they do | Price point | Why CETL wins the niche |
|---|---|---|---|
| **Skillcrush** | Woman-owned online coding school, self-paced "Break Into Tech" | $1,599–$2,499[^20] | Self-paced format the completion data indicts;[^16] career-switch framing; no ADHD or midlife design |
| **SheCodes, Grace Hopper, Hackbright & other women's bootcamps** | Coding bootcamps for women[^21] | Hundreds to $10K+ | Job-placement oriented, intense schedules hostile to executive dysfunction; none targets the perimenopause+ADHD niche[^21] |
| **Maven & cohort-course platforms** | Marketplace/infrastructure for cohort courses (500k+ learners; a16z-backed)[^18] | ~10% rev share | Not a competitor but a **launch channel** — CETL can run early cohorts on Maven for distribution and social proof |
| **ADHD coaching** | 1:1 accountability and skills coaching[^12] | $150–$300+/session (market rates) | Coaching builds systems but doesn't teach you to *build things*; CETL delivers accountability *and* a portfolio |
| **Menopause femtech** | Symptom tracking, telehealth, HRT access | Subscription apps | Clinical/symptom focus; none offers skill-building or identity-rebuilding — and only ~7% of femtech even addresses menopause[^15] |
| **Free content (YouTube, MOOCs, ChatGPT)** | Infinite unstructured material | $0 | Median MOOC completion 12.6%;[^16] free content is the *reason* the audience believes they "can't learn tech" — CETL sells the structure, not the information |

**Positioning statement:** *The only build-something program designed for how a
perimenopausal ADHD brain actually works — live, structured, funny, and finished in
weeks, with Claude as your co-pilot and a room full of women who get it.*

**Defensibility:** the moat is not the curriculum (copyable) but (a) founder-audience
authenticity, (b) the podcast's trust asset, (c) a community that compounds, and
(d) speed — the niche is visibly empty today.[^15][^21]

---

## 7. Marketing & Customer Acquisition

**Strategy: one channel done well — the podcast — surrounded by borrowed audiences.**

1. **Podcast as trust engine.** Weekly. Every episode maps to a pain point in §3
   (brain fog at work, leading while exhausted, the late-diagnosis story, "am I still
   smart?"). Niche professional-audience podcasts command premium sponsor rates precisely
   because the audience is concentrated and high-intent[^24] — the same concentration
   that makes it a superb enrollment channel.
2. **Borrowed audiences.** Guest on ADHD and menopause podcasts; pitch ADDitude-style
   publications (their own surveys prove the audience's hunger[^3][^4]); partner with
   menopause and ADHD coaches who don't teach tech.
3. **Build-in-public.** Jani's own agent projects (this repo) are the demo. "I built my
   own chief-of-staff agent, and I'll teach you to build yours" is the hook.
4. **Email list as the conversion asset.** Podcast → free "First Page in a Weekend"
   mini-guide → nurture sequence → cohort waitlist. Cohort launches to the list 3–4×/year.
5. **Beta cohort as marketing.** Run the first cohort at $299 explicitly as a beta —
   discounted first cohorts with prices raised in later runs is standard, recommended
   practice[^19] — and harvest testimonials, before/afters, and completion stats.

**CAC posture:** $0 paid acquisition in Year 1. All growth is content + partnerships.
Paid tests only after two cohorts prove conversion and testimonial supply.

---

## 8. Revenue Model & Financial Projections

*All projections are founder planning assumptions, not researched facts; the benchmarks
they lean on are cited.*

### 8.1 Pricing rationale

- Cohort programs with live components typically price at **$800–$2,500** (full range
  $500–$5,000+),[^19] and Maven's own guidance is that full cohort courses run
  **$950–$3,500+**, with ~$500 recommended for a first-time instructor's short course.[^18]
- Cohorts can charge a premium over self-paced precisely because of accountability and
  community[^18] — CETL's entire design.
- CETL prices below career-switch competitors ($1,599+ at Skillcrush[^20]) because the
  promise is transformation and capability, not job placement.

### 8.2 Three-year sketch

| | Year 1 | Year 2 | Year 3 |
|---|---|---|---|
| Cohorts run | 3 (1 beta) | 4 | 5 |
| Avg seats | 10 → 14 | 16 | 18 |
| Avg price | $299 beta, then $500 | $650 | $800 |
| Cohort revenue | ~$17K | ~$42K | ~$72K |
| Community members (avg, $24/mo) | 15 | 60 | 130 |
| Community revenue | ~$4K | ~$17K | ~$37K |
| Podcast revenue | ~$0 | $2–5K | $5–15K |
| **Total revenue** | **~$21K** | **~$60–65K** | **~$115–125K** |

**Podcast monetization reality check:** 2025 host-read CPMs average $18–26,[^22] which at
early-stage download numbers is immaterial; shows under ~5,000 downloads/episode are
advised to sell flat-fee host-read integrations instead of CPM,[^23] and niche
professional shows can command roughly 60% higher rates once established.[^24] Hence the
plan treats podcast dollars as a Year 2–3 bonus and its Year 1 value as enrollment.

### 8.3 Cost structure

Lean by design: podcast hosting + recording gear (~$1–2K one-time), website/email/
community tooling (~$150–300/mo), cohort platform (Maven's ~10% rev share[^18] or a
flat-fee alternative), insurance/LLC/accounting (~$1–2K/yr), Claude subscriptions for
teaching demos. **Break-even lands inside the first two cohorts.** Founder time is the
real cost — see risks.

---

## 9. Operations & Roadmap

**Phase 0 (now–month 2):** finish the four curriculum modules (repo content pipeline
already exists — `cetl_content_agent.py` generates lesson plans, scripts, and promo);
launch podcast with 3-episode backlog; landing page + email capture; LLC + waiver/ToS.

**Phase 1 (months 2–4):** enroll 8–12 beta students at $299 from warm network + first
listeners; run the beta; instrument everything (attendance, completion, artifacts
shipped); collect testimonials.

**Phase 2 (months 4–9):** revise curriculum from beta learnings; raise price to ~$500;
run cohorts 2–3; open CETL Circle to alumni; begin guest-podcast circuit.

**Phase 3 (months 9–18):** price toward $650–$950 as social proof accumulates;[^18][^19]
evaluate Maven as an additional distribution channel;[^18] pilot one corporate/ERG
workshop; first podcast sponsorships as flat-fee host-reads.[^23]

**KPIs:** cohort completion rate (target ≥70% — above the 64.2% cohort benchmark[^16]),
artifacts shipped per student, NPS, email list growth, waitlist-to-enrollment conversion,
community churn.

---

## 10. Risks & Mitigations

| Risk | Mitigation |
|---|---|
| **Founder is the single point of failure** (and is the demographic she serves — energy/fog days are real) | Small cohorts; 2 sessions/week max; recorded sessions; async-first materials; community peer-support absorbs load; the curriculum's honesty about fog *is* the brand |
| **Body-doubling/coaching evidence is promising but young** — overclaiming invites backlash | Market outcomes ("you'll ship 4 projects with a group"), not therapy claims; cite evidence with its caveats, as this plan does[^10][^12] |
| **AI tooling shifts under the curriculum** | Teach durable concepts (structure, servers, agents-as-loops) with Claude as the current instrument; modules are versioned in-repo and regenerable via the content agent |
| **Niche too small / conversion assumptions wrong** | Bottom-up model needs only ~40 seats/yr; if ADHD+perimenopause proves too narrow, the concentric circle (all midlife women in leadership who want tech skills) is one message-tweak away |
| **Copycats** | Move fast, own the podcast trust asset, let community compound; the niche is defensible by authenticity more than IP |
| **Health-adjacent positioning invites regulatory/clinical confusion** | Prominent education-not-medical-advice disclaimers; no symptom-treatment claims; podcast guests with clinical credentials speak for themselves |
| **Platform dependence (Maven, Substack, etc.)** | Own the email list; treat platforms as channels, never the home |

---

## Footnotes

[^1]: Faubion SS, et al. "Impact of Menopause Symptoms on Women in the Workplace." *Mayo Clinic Proceedings* (2023). Survey of 4,440 employed US women aged 45–60: 13.4% adverse work outcome; 10.8% missed work (median 3 days); est. $1.8B/yr US lost work time, $26.6B incl. medical costs; highest-severity quartile 15.6× more likely to report adverse work outcomes. https://pubmed.ncbi.nlm.nih.gov/37115119/

[^2]: Mayo Clinic News Network. "Mayo Clinic study puts price tag on cost of menopause symptoms for women in the workplace" (2023). https://newsnetwork.mayoclinic.org/discussion/mayo-clinic-study-puts-price-tag-on-cost-of-menopause-symptoms-for-women-in-the-workplace/

[^3]: ADDitude Magazine. "Menopause Symptoms Exacerbate ADHD in Women" — survey of 1,500+ women with diagnosed/undiagnosed ADHD: 94% reported ADHD symptoms grew more severe in peri/menopause; ~70% called brain fog/memory/overwhelm "life-altering" in their 40s–50s (vs. 11% for memory in their 20s–30s); 83% experienced some ADHD symptoms for the first time during the transition. https://www.additudemag.com/menopause-symptoms-adhd-survey/

[^4]: ADDitude Magazine. "Perimenopausal Symptoms Are More Severe, Begin Earlier in Women with ADHD" — survey of 2,600+ women: 43% first diagnosed at ages 41–50; 61% said ADHD had its greatest life impact at ages 40–59. https://www.additudemag.com/perimenopausal-symptoms-women-adhd/

[^5]: AARP Public Policy Institute. "Menopause in the Workplace Has an Economic Impact" — ~¼ of the US workforce is in some phase of the menopause transition; 76.8% of women aged 45–54 are in the labor force. https://www.aarp.org/pri/topics/work-finances-retirement/employers-workforce/menopause-workplace/

[^6]: Fawcett Society survey of 4,000 UK women aged 45–55 (reported by People Management): 10% left a job due to menopause symptoms; 14% cut hours; 8% did not apply for promotion; 44% said ability to work was affected; 73% reported brain fog. https://www.peoplemanagement.co.uk/article/1754967/one-10-women-quit-job-menopause-symptoms-survey-reveals

[^7]: "Research advances and future directions in female ADHD: hormonal fluctuations, mood, and cognition." *Frontiers* review (PMC12277363): ADHD in girls/women is under-recognized and under-researched, contributing to late diagnosis; hormonal transitions incl. perimenopause exacerbate ADHD symptoms and mood disturbance via estrogen–dopamine interactions affecting executive function. https://pmc.ncbi.nlm.nih.gov/articles/PMC12277363/

[^8]: "Examining the Link Between ADHD Symptoms and Menopausal Experiences." *Journal of Attention Disorders* (2025). Cross-sectional study, 656 women aged 45–60 (245 with ADHD diagnosis): higher ADHD symptom severity correlated with worse menopausal difficulties and poorer quality of life on nearly all WHQ/MENQoL measures; symptom severity, not diagnosis status, drove the effect. https://journals.sagepub.com/doi/10.1177/10870547251355006

[^9]: Attention Deficit Disorder Association (ADDA). "ADHD and Perimenopause/Menopause — How Symptoms Overlap." https://add.org/adhd-and-perimenopause-menopause/

[^10]: "You Are Not Alone: Designing Body Doubling for ADHD in Virtual Reality" (arXiv, 2025). Controlled experiment, 12 adults with ADHD: faster task completion, greater perceived accuracy and sustained attention with a human or AI body double vs. alone; participants preferred body doubling. Authors note prior research had not conclusively demonstrated effectiveness and some neurodivergent people report discomfort with social presence. https://arxiv.org/pdf/2509.12153

[^11]: Cleveland Clinic. "How Body Doubling Helps With ADHD" — body doubling as external structure and accountability compensating for executive-function deficits. https://health.clevelandclinic.org/body-doubling-for-adhd

[^12]: CHADD. "Emerging Evidence for the Effectiveness of ADHD Coaching" — review of 19 studies (16 peer-reviewed): coaching built on goal-setting, psychoeducation, and accountability improves outcomes; increasingly recommended in multimodal adult ADHD treatment; authors note methodological limits (small samples, few controlled designs). https://chadd.org/wp-content/uploads/2018/10/2B_Emerging_Evidence_for_the_Effectiveness_of_ADHD_Coaching_.pdf

[^13]: "A self-guided internet-delivered intervention for adults with ADHD" — RCT (PMC10033990): weaker engagement/outcomes for self-guided vs. guided formats in adults with ADHD. https://pmc.ncbi.nlm.nih.gov/articles/PMC10033990/

[^14]: Grand View Research. "FemTech Market Size, Share & Growth Analysis Report, 2030" — femtech market ~$52–63B (2025), 11–15% CAGR. https://www.grandviewresearch.com/industry-analysis/femtech-market-report

[^15]: PreScouter. "Menopause: the untapped opportunity in femtech" — only ~7% of femtech startups focus on menopause; femtech grew from ~11 companies (2006) to 200+ (2021). (Note: this source's "$600B by 2030" market projection failed independent verification and is not relied on in this plan.) https://www.prescouter.com/inquiry/untapped-menopause-market-opportunity/

[^16]: Ruzuku. "The Completion Gap: What 32,000 Courses Reveal" — cohort-based courses 64.2% completion vs. 48.2% open-access/self-paced across 32,000+ courses; cites Class Central's median MOOC completion of 12.6%. https://www.ruzuku.com/learn/articles/completion-gap

[^17]: Grand View Research. "ADHD Apps Market Size & Trends Report" — $1.9B (2025) projected to $6.7B by 2033 (17.5% CAGR); used here as the closest proxy for consumer ADHD-support spending. https://www.grandviewresearch.com/industry-analysis/adhd-apps-market-report

[^18]: Maven. "In online ed, content is no longer king — cohorts are" (orig. a16z essay by Wes Kao, Maven co-founder) — Harvard's 2014 online case-method courses with peer collaboration reached 85% completion vs. single-digit MOOC rates of that era; Maven raised a $20M a16z-led Series A; and Maven, "How to decide the price & length of your course" — recommends first-time instructors run a 1–2 week course at ~$500; full cohort courses typically $950–$3,500+; cohorts can charge a premium over self-paced due to accountability and community; discounted beta first cohorts with later price increases are standard. https://maven.com/resources/a16z-cohorts-are-king ; https://maven.com/resources/course-price-and-length

[^19]: Disco. "Pricing Models for Cohort-Based Courses" — live-component cohort programs typically $800–$2,500 (range $500–$5,000+); recommends hybrid premium-cohort + lower-priced ongoing community membership. https://www.disco.co/blog/pricing-models-for-cohort-based-courses-guide

[^20]: Skillcrush. "Skillcrush Is an Alternative to Coding Bootcamps" — Break Into Tech at $1,599, or $2,499 with job-guarantee track. https://skillcrush.com/blog/skillcrush-alternative-to-coding-bootcamps/

[^21]: Course Report. "The 10 Best Coding Bootcamps for Women in 2025" — landscape of women-focused coding education (SheCodes, Skillcrush, Grace Hopper, Hackbright, etc.); none targets the perimenopause+ADHD niche. https://www.coursereport.com/blog/the-10-best-coding-bootcamps-for-women-in-2025

[^22]: Podscan. "Podcast Advertising Rates 2025: CPM Benchmarks" — host-read ads average $18–22 CPM (30s) and $24–26 CPM (60s); health/wellness among the fastest-growing categories. https://podscan.fm/blog/podcast-advertising-rates-cpm-benchmarks-2025

[^23]: Pod Partnerships. "Podcast advertising CPM benchmarks: what brands actually pay in 2026" — shows under ~5,000 downloads/episode should sell flat-fee host-read integrations rather than CPM; niche professional-audience mid-rolls command premium rates. https://podpartnerships.com/blogs/podcast-advertising-cpm-benchmarks

[^24]: InfluenceFlow. "Podcast Sponsorship Rate Cards: The Complete 2026 Guide" — niche professional-audience podcasts command ~60% higher CPMs; micro podcasts typically charge $18–35 CPM or $100–500 flat per episode. https://influenceflow.io/resources/podcast-sponsorship-rate-cards-the-complete-2026-guide-for-creators-brands/
