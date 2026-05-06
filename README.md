# Startup Scout

Automatically discover pre-seed European startups matching your investor thesis — scored, verified, and exported to Excel in minutes.

**No hallucinations.** Every company comes from a real web source (Hacker News, university spinoff pages, EU VC portfolios, accelerator databases, RSS feeds). AI is only used to score thesis fit — never to invent companies.

---

## How it works

```
Your thesis → Keywords → 28 real sources → AI scoring → Verified → Excel report
```

1. **Scout** — scrapes 28 sources (HN, EPFL/Oxford/TU Delft spinoffs, Seedcamp/Creandum/Northzone portfolios, EU-Startups, BetaList, GitHub, and more)
2. **Analyst** — scores each company against your thesis across 5 dimensions (sector, geo, stage, tech, theme fit)
3. **Verifier** — checks for funding signals to filter out companies that already raised Series A+
4. **Export** — color-coded Excel file with clickable links and full AI reasoning

---

## Setup (5 minutes)

### 1. Get a free Groq API key
Go to [console.groq.com](https://console.groq.com) → API Keys → Create key.
Free tier: 100,000 tokens/day — enough for one full run.

### 2. Clone and install
```bash
git clone https://github.com/YOUR_USERNAME/startup-scout.git
cd startup-scout
pip install -r requirements.txt
```

### 3. Add your key
```bash
echo "GROQ_API_KEY=your_key_here" > .env
```

### 4. Run
```bash
python3 startup_scout.py
```

Paste your investor thesis when prompted, press Enter twice, and wait ~5 minutes.

---

## Output

An Excel file with:
- Company name, description, country, source
- Scores: overall / sector / geo / stage / tech / theme (1–5)
- AI reasoning, best signal, red flags
- Funding verification status
- Clickable links

Color coded: score 4-5 = green · score 3 = yellow · score 1-2 = red

---

## Sources (28 total)

| Category | Sources |
|---|---|
| Hacker News | Launch HN (20 pages) + keyword search |
| EU startup news | EU-Startups, Sifted, Tech.eu, Nordic9, Maddyness |
| VC portfolios | Seedcamp, LocalGlobe, Creandum, Speedinvest, Point Nine, Cherry, Earlybird, Balderton, Northzone + more |
| Accelerators | EIT InnoEnergy, Climate-KIC, Startup Wise Guys, EIT Digital, Station F |
| Universities | EPFL, ETH, TU Delft, KTH, Imperial, Oxford, Cambridge, TU Munich, KU Leuven, Aalto + more |
| Launch platforms | BetaList, Product Hunt |
| Databases | GitHub API, CORDIS EU grants, EIC Portfolio, Crunchbase, Dealroom, Wellfound |

---

## Requirements

- Python 3.8+
- Free [Groq API key](https://console.groq.com)

```
pip install -r requirements.txt
```
