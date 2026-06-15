# Helix MedTech — Strategy Simulation

A finance-led business simulation built around **Helix MedTech**, a medical-technology
company that combines a mature **medical-devices** business (the *core* to exploit) with a
fast-growing **digital-health SaaS platform** (the *explore* / growth engine). Teams play
**three rounds**, making seven strategic decisions each round; the model converts those
decisions into financial results and **enterprise value (EV)**.

## What's in this folder

| File | What it is |
|---|---|
| `medtech-simulation.html` | The interactive **web shell**. Double-click to open in any browser. Teams dashboard + per-team workspace with one tab per decision area, a live financial readout, and a leaderboard. Progress saves to the browser. |
| `helix-medtech-simulation.xlsx` | The **Excel workbook**. Decision dropdowns, a transparent financial engine, results charts, and a teams dashboard. Plays all three rounds in Excel. |
| `build_simulation.py` | The generator. Single source of truth for the model — rebuilds both files (`python3 build_simulation.py`). |

Both deliverables share the **exact same financial model**, so a strategy gives the same
answer in the browser and in Excel.

## How to run a session

1. Open the **web shell** (`medtech-simulation.html`) or hand each team a copy of the **workbook**.
2. Each round, every team picks **one option** in each of the seven decision areas.
3. Lock the round and advance. Each round opens from the prior round's closing position.
4. After Round 3, compare teams on the **Teams Dashboard**. Highest enterprise value wins.

## Starting position (Round 1 opening)

| Item | Value |
|---|---|
| Core revenue (medical devices) | $1,800M |
| Explore revenue (SaaS / ventures) | $200M |
| Total revenue | $2,000M |
| Gross margin | 60% |
| Operating costs (recurring) | $700M |
| EBITDA | $500M |
| Cash | $300M |
| Risk index (0–100, lower better) | 30 |
| Recurring-revenue mix | 15% |
| EV multiple (base) | 9x |
| **Enterprise value** | **$4,950M** |

Each round also applies organic growth before decisions: core **+1.5%**, explore **+8.0%**.

## How the maths works

Each round, the seven chosen options' impacts are **summed**, then flow through a P&L:

```
core_rev      = prior_core    × (1 + (base_core_growth   + Σ coreG) / 100)
explore_rev   = prior_explore × (1 + (base_explore_growth + Σ expG) / 100)
gross_margin% = prior_margin% + Σ margin            (carries forward)
gross_profit  = total_revenue × gross_margin% / 100
op_costs      = prior_op_costs + Σ opex             (carries forward)
EBITDA        = gross_profit − op_costs
cash          = prior_cash + EBITDA − Σ spend       (one-time investment)
risk          = clamp(0..100, prior_risk + Σ risk)
recurring%    = clamp(0..90,  prior_recurring + Σ recur)
```

**Enterprise value = EBITDA × EV multiple**, where:

```
multiple = clamp(6..18,
    9
    + cumulative Σ mult                     (strategic re-rating, carries forward)
    + recurring% / 100 × 6              (investors pay up for recurring SaaS revenue)
    + revenue_growth% × 0.15             (growth premium)
    − (risk − 30) × 0.05)            (risk discount)
```

**Score** = EV indexed to the starting EV (Start = 100), with a penalty if cash runs negative.
The lesson: short-term cuts can flatter EBITDA but shrink the multiple (via lower growth,
less recurring revenue and higher risk), so they often *destroy* enterprise value.

## Decision areas, scales, options and impacts

Impacts per round (additive). Legend: **Spend** one-time investment $M · **Core g** / **SaaS g**
growth percentage points · **Margin** gross-margin pp · **Opex** recurring $M · **Risk** index pts
· **Mult** EV-multiple pts · **Recur** recurring-mix pp.

### ⚙️ Managing the Core (Exploit)

*How hard you run, defend and modernise the mature medical-devices business.*

| Scale | Option | Spend | Core g | SaaS g | Margin | Opex | Risk | Mult | Recur | Rationale |
|---|---|--:|--:|--:|--:|--:|--:|--:|--:|---|
| Defensive | **Harvest & cut** | 0 | -3 | 0 | 1.5 | -40 | 5 | -0.5 | 0 | Maximise short-term cash from devices; minimal reinvestment. |
| Steady | **Maintain** | 20 | 0 | 0 | 0 | 0 | 0 | 0 | 0 | Keep the core running as-is; protect share without big bets. |
| Efficiency | **Optimise & defend** | 60 | 2 | 0 | 2 | -10 | -3 | 0.3 | 0 | Lean + automation in plants; targeted upgrades to defend share. |
| Offensive | **Invest to grow core** | 120 | 5 | 0 | -0.5 | 20 | 3 | 0.4 | 0 | New indications & geographies for the device line. |

### 🧭 Explore Portfolio

*Investment in the digital-health SaaS platform and new growth ventures.*

| Scale | Option | Spend | Core g | SaaS g | Margin | Opex | Risk | Mult | Recur | Rationale |
|---|---|--:|--:|--:|--:|--:|--:|--:|--:|---|
| None | **Pause exploration** | 0 | 0 | -10 | 0.5 | 0 | 4 | -1.0 | 0 | Freeze venture spend; bank the savings, accept slower SaaS. |
| Focused | **Selective bets** | 50 | 0 | 10 | 0 | 10 | 0 | 0.5 | 2 | Fund the 1-2 strongest digital-health products only. |
| Balanced | **Balanced portfolio** | 110 | 0 | 25 | -0.5 | 20 | 2 | 1.2 | 3 | A managed portfolio of SaaS bets across the pipeline. |
| Aggressive | **Aggressive venture build** | 200 | 0 | 45 | -1.5 | 40 | 6 | 2.0 | 5 | Go big on the platform; chase category leadership in SaaS. |

### ⚡ Managing Disruption

*Response to AI diagnostics and connected-device entrants attacking the core.*

| Scale | Option | Spend | Core g | SaaS g | Margin | Opex | Risk | Mult | Recur | Rationale |
|---|---|--:|--:|--:|--:|--:|--:|--:|--:|---|
| Reactive | **Wait & see** | 0 | -2 | 0 | 0 | 0 | 12 | -1.0 | 0 | Hold position; assume the threat is over-hyped. |
| Collaborate | **Partner / license** | 40 | 1 | 0 | 0 | 0 | -4 | 0.5 | 2 | License AI / connectivity from a fast mover; integrate quickly. |
| M&A | **Acquire capability** | 180 | 3 | 10 | 0 | 15 | -6 | 1.0 | 3 | Buy a disruptor to own the new platform outright. |
| Build | **Build internal moonshot** | 130 | 0 | 15 | -1 | 0 | -2 | 0.8 | 2 | Stand up a skunkworks to leapfrog the entrants. |

### 🚚 Supply Chain

*Trading off landed cost against resilience for devices and components.*

| Scale | Option | Spend | Core g | SaaS g | Margin | Opex | Risk | Mult | Recur | Rationale |
|---|---|--:|--:|--:|--:|--:|--:|--:|--:|---|
| Cost | **Lowest cost (single-source)** | 0 | 0 | 0 | 2 | 0 | 10 | -0.5 | 0 | Offshore, single-source for the best unit cost. |
| Balanced | **Dual-source** | 40 | 0 | 0 | 0.5 | 5 | -3 | 0 | 0 | Second suppliers on critical parts to cut single points of failure. |
| Resilient | **Regionalise / nearshore** | 90 | 0 | 0 | -0.5 | 10 | -8 | 0.3 | 0 | Nearshore key flows; shorten and de-risk the network. |
| Integrate | **Vertical integration** | 160 | 1 | 0 | 1.5 | -5 | -6 | 0.4 | 0 | Automate & bring critical manufacturing in-house. |

### 💰 Revenue & Commercial Model

*Pricing and business model — from one-off device sales to recurring outcomes.*

| Scale | Option | Spend | Core g | SaaS g | Margin | Opex | Risk | Mult | Recur | Rationale |
|---|---|--:|--:|--:|--:|--:|--:|--:|--:|---|
| Volume | **Discount for volume** | 10 | 3 | 0 | -2.5 | 0 | 0 | -0.3 | 0 | Cut price to win units and defend share. |
| Premium | **Hold premium pricing** | 10 | -1 | 0 | 2 | 0 | 0 | 0.3 | 0 | Protect price and brand; accept slower volume. |
| Value | **Value-based selling** | 50 | 3 | 0 | 1 | 15 | 0 | 0.5 | 3 | Sell on clinical & economic outcomes to hospitals. |
| Recurring | **Device-as-a-Service (recurring)** | 90 | 2 | 5 | -1 | 20 | 0 | 1.5 | 10 | Bundle hardware + SaaS into subscriptions / outcomes contracts. |

### 👥 People & Talent

*Workforce strategy — the execution engine behind every other choice.*

| Scale | Option | Spend | Core g | SaaS g | Margin | Opex | Risk | Mult | Recur | Rationale |
|---|---|--:|--:|--:|--:|--:|--:|--:|--:|---|
| Cut | **Cost reduction / layoffs** | 0 | -2 | -5 | 1 | -60 | 8 | -0.8 | 0 | Reduce headcount to protect near-term margin. |
| Steady | **Maintain** | 20 | 0 | 0 | 0 | 0 | 0 | 0 | 0 | Hold the workforce steady; no major change. |
| Develop | **Upskill & reskill** | 60 | 1 | 5 | 1 | 10 | -3 | 0.5 | 0 | Reskill for software & data; raise productivity. |
| Transform | **Transform culture & talent** | 110 | 2 | 10 | 1.5 | 30 | -5 | 1.0 | 0 | Top-quartile talent + culture change for agility. |

### 🏛️ Corporate Functions

*Overhead and the digital backbone (finance, IT, shared services).*

| Scale | Option | Spend | Core g | SaaS g | Margin | Opex | Risk | Mult | Recur | Rationale |
|---|---|--:|--:|--:|--:|--:|--:|--:|--:|---|
| Cut | **Slash overhead** | 0 | 0 | -3 | 0 | -50 | 6 | -0.4 | 0 | Aggressively cut corporate cost; accept fragility. |
| Steady | **Maintain** | 15 | 0 | 0 | 0 | 0 | 0 | 0 | 0 | Run functions as-is. |
| Digitise | **Selective digitisation** | 70 | 0 | 0 | 0.5 | -20 | -2 | 0.3 | 0 | Automate the highest-cost processes first. |
| Transform | **Full digital core + shared services** | 140 | 1 | 0 | 1 | -45 | -4 | 0.6 | 0 | One digital backbone; shared services at scale. |

## The three rounds

**Round 1 — Pressure on the Core.** Reimbursement cuts squeeze device margins and an AI-diagnostics start-up lands its first hospital contracts. The board wants a plan that protects the core while standing up the digital-health platform.

**Round 2 — Disruption Accelerates.** A rival launches a connected-device + subscription bundle, and a component shock exposes single-source supply. Investors start rewarding recurring revenue and punishing fragility.

**Round 3 — The New Normal.** Value-based care mandates reshape buying. The market now prizes recurring SaaS revenue, proven innovation and low operating risk. Time to convert your build-up into enterprise value.

## Rebuilding

```bash
pip install openpyxl
python3 build_simulation.py
```

Edit the `MODEL` section near the top of `build_simulation.py` to retune impacts, add
options, or change the narrative — then rebuild to regenerate both the workbook and the web shell.

