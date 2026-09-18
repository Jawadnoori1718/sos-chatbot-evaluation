# Who Judges the Judges?

**Stakeholder-defined evaluation of candidate base models for a university student wellbeing signposting chatbot.**

Summer 2026 AI Research Internship, School of Computer Science, University of Nottingham.
Part of the [Responsible AI UK](https://rai.ac.uk) Cornerstone 2 AI Assurance programme.

---

## The question

Small charities and university student-support services are deploying LLM chatbots in wellbeing and signposting roles. These are sensitive deployments, yet the organisations making them are rarely able to fund an external audit.

**Can a small organisation meaningfully check an AI system it is about to deploy — and trust what the check tells it?**

## What was done

| Stage | |
|---|---|
| **31** | LLM evaluation tools inventoried across 29 dimensions |
| **6** | accessibility criteria applied as a screen — no code required; no install or browser-based; free at our scale; plain-language outputs; open methodology; configurable to a specific organisation |
| **1** | tool met all six: [Weval](https://weval.org), self-hosted locally |
| **11** | scenarios and **74** criteria, written from UK university wellbeing practitioner requirements and reviewed by mental-health academics |
| **4** | candidate base models × **2** cross-provider LLM judges = **44** evaluations |

Models evaluated: Claude Haiku 4.5 · Gemini 2.5 Flash · GPT-4.1-mini · Llama 3.1 8B
Judges: GPT-4o-mini (via OpenRouter) and Claude Haiku 4.5

## Findings

### 1. Models refuse well. They notice badly.

Coverage is highest where the correct response is to decline, and lowest where risk is implicit. The two weakest scenarios are among the most consequential: somatic presentation (62.0%) and direct suicidal disclosure (68.3%, carrying the heaviest practitioner weight in the blueprint).

![Scenario × model coverage](scenario-model-coverage.png)

### 2. Over half the evaluations were statistically unreliable — and the dashboard did not say so.

Weval measures inter-judge agreement using Krippendorff's α. Of 44 evaluations it flagged **23 as unreliable** and 18 as reliable; three were mathematically unstable and five returned negative α — agreement worse than chance, the worst at −0.688. At criterion level, **110 of 292** scored criteria showed disagreement between the two judges.

![Judge agreement across 44 evaluations](judge-agreement.png)

*Each mark is one evaluation, sorted by α. Orange is unreliable.*

Gemini 2.5 Flash ranked **second** overall while producing the least reliable scores of any model (mean α = 0.338; 9 of 11 evaluations unreliable).

### 3. Aggregation and judge configuration were consequential, and largely invisible.

The headline leaderboard averages scenarios equally, discarding the practitioner-assigned weights; weighted, the gap between best and worst model widens from 16.6 to 20.1 points. Separately, one of the two judges was also one of the four candidate models, and the other shared a provider family with a candidate. This is a configuration risk surfaced, not a bias effect measured — but it is one a non-specialist user would be unlikely to notice.

### Results

| Model | Weighted | Unweighted | Mean α | Unreliable |
|---|---:|---:|---:|---:|
| Claude Haiku 4.5 | **85.8%** | 84.3% | 0.525 | 6 / 11 |
| Gemini 2.5 Flash | **81.2%** | 80.9% | 0.338 | 9 / 11 |
| GPT-4.1-mini | **72.1%** | 73.4% | 0.623 | 5 / 11 |
| Llama 3.1 8B | **65.7%** | 67.7% | 0.663 | 6 / 11 |

---

## Repository contents

| File | What it is |
|---|---|
| `sos-v03-stakeholder.yml` | The evaluation blueprint — 11 scenarios, 74 criteria, practitioner weights |
| `sos-v03-runA_comparison.json` | Full run output, including every judge's reasoning |
| `Weval-Criterion-Review.xlsx` | Per-criterion scores and both judges' reasoning, formatted for practitioner review |
| `build_review_workbook.py` | Builds the review workbook from a Weval run |
| `scenario-model-coverage.png` · `judge-agreement.png` | Figures used above |
| `poster.pdf` | Conference poster (A1) |

### Reproducing the run

```bash
git clone https://github.com/weval-org/app.git weval && cd weval
pnpm install
cp .env.template .env        # add OPENAI_API_KEY, ANTHROPIC_API_KEY, OPENROUTER_API_KEY

mkdir my-blueprints
cp /path/to/sos-v03-stakeholder.yml my-blueprints/

STORAGE_PROVIDER=local pnpm cli run-config local \
  --config my-blueprints/sos-v03-stakeholder.yml \
  --run-label "sos-v03-runA" --cache \
  --gen-retries 3 --gen-timeout-ms 90000

STORAGE_PROVIDER=local pnpm dev     # dashboard at localhost:3172
```

Route the two judges through different providers. Running both against the same provider's rate-limit budget causes one to fail silently partway through, leaving a share of criteria scored by a single judge while the dashboard still reports a complete result.

### Rebuilding the review workbook

```bash
pip install openpyxl pyyaml
python build_review_workbook.py \
  sos-v03-stakeholder.yml \
  sos-v03-runA_comparison.json
```

---

## Limitations

- Single run; no repeats and no confidence intervals
- Single-turn only, which may penalise a model for appropriately asking a clarifying question
- Both judges shared a provider family with a candidate model
- One use case; no student participants in criteria design
- LLM judgements not yet validated against human ratings

**Scope:** candidate base models were evaluated against practitioner-derived criteria. No deployed system was evaluated.

## Next steps

1. **Human evaluation of the judges** — practitioners rate a subset of the judge reasoning in `Weval-Criterion-Review.xlsx`, so LLM judgements can be validated against human ratings
2. **Core vs. context-specific criteria** — a mandatory core set for any student-facing support chatbot, plus optional sets selected by purpose
3. **Independent judge panel** — no candidate's provider judging; repeat runs with confidence intervals
4. **Multi-turn evaluation**
5. **Student co-production** of criteria alongside practitioners

## References

1. Krippendorff, K. (2004) *Content Analysis: An Introduction to Its Methodology*. Thousand Oaks, CA: Sage.
2. Collective Intelligence Project. Weval. https://weval.org
3. World Health Organization (2011) *Psychological First Aid: Guide for Field Workers*. Geneva: WHO.
4. MindEval (2025) Multi-turn mental-health evaluation. arXiv:2511.18491

## Supervision

**Prof Joel Fischer** — Mixed Reality Laboratory, School of Computer Science, University of Nottingham
**Dr Aislinn Gómez Bergin** — Responsible AI UK / MindTech, University of Nottingham

Criteria reviewed by mental-health academics in an expert session on 11 August 2026.

## Author

**Jawad Noori** — [jawadnoori.co.uk](https://jawadnoori.co.uk) · [LinkedIn](https://linkedin.com/in/jawadnoori1)

## Licence

Code and documentation: MIT (see [LICENSE](LICENSE)).
The evaluation blueprint encodes requirements gathered through supervised stakeholder engagement; please cite this repository and contact the supervisors listed above before reuse in published work.
