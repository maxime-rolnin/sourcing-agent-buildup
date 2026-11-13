# Buy-and-build sourcing agent

A small Python agent that does the first pass of an analyst's target screening for a buy-and-build strategy. It finds companies, enriches them with public financial data and ranks them in an Excel long list that a human then reviews.

The use case shown here: an industrial buyer looking for independent companies in **water dispenser rental and maintenance** in France, with revenue between €1M and €10M.

## Why public tenders

To scout an athlete, you read competition results rather than visiting every stadium. This agent does the same with companies. In this trade, the "results" are public procurement awards: every time a city hall, hospital or ministry picks a supplier to rent and maintain its water dispensers, the award is published. A company that wins these contracts proves three things at once: recurring revenue, a covered territory and real service capacity.

## How it works

| Step | What the agent does | Source |
|---|---|---|
| 1. Signal | Pulls award notices for water dispenser rental and maintenance contracts (2022 to 2026) and the names of the winners | [BOAMP open data](https://boamp-datadila.opendatasoft.com/explore/dataset/boamp/) |
| 2. Filter | Removes large groups (Elis, Culligan and its brands, Rentokil Initial), adjacent trades and the buyer's own subsidiaries | Rules in the script |
| 3. Enrich | Looks up each remaining company: company ID, revenue, net income, headcount, management | [Recherche d'entreprises API](https://recherche-entreprises.api.gouv.fr/docs/) |
| 4. Score | Applies a weighted grid on six criteria; weights and scores are editable yellow cells and the ranking recalculates in Excel | Excel formulas |
| 5. Human | Every line is checked before any contact | |

No API key is needed: both sources are public and free.

## Output

The agent writes one Excel workbook with five tabs:

* **Lisez-moi**: method, sources, limits and colour legend
* **Long list**: ranked targets with financials, tender count and the analyst's reading
* **Signaux**: every award notice, one line per winner
* **Concurrence**: share of awards by player (large groups vs independents)
* **Exclus**: who was removed and why

An example is in [`examples/longlist_exemple.xlsx`](examples/longlist_exemple.xlsx). On the cached data, the agent reviews 54 award notices and 68 awards, isolates 12 independent winners and ranks 3 priority targets.

## Run it

```bash
pip install -r requirements.txt

python sourcing_agent.py --cache                 # replays the cached data in data/
python sourcing_agent.py --live                  # queries the public APIs (internet needed)
python sourcing_agent.py --live --out my_list.xlsx
```

To adapt it to another sector, edit the `THESE` dictionary (keywords, size range), the exclusion lists and the NAF codes at the top of the script.

## Limits

* Public tenders show only part of the market: companies serving only private clients do not appear. The next source to plug in would be trade directories and company websites.
* Registry revenue comes from filed accounts and can be missing when accounts are confidential. Net income is not EBITDA: the full accounts must be read.
* Matching a tender winner to a company ID is automatic and must be verified.
* The acquirer is anonymised in this public version. Scores are illustrative hypotheses, not investment recommendations, and no company listed here has been contacted.

## Built with

Python and openpyxl. The code was written with an AI coding agent; the investment thesis, scoring grid and review of the results are mine.

Author: Maxime Rolnin
