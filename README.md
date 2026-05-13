# Job Seeker Scraper

This project collects leads of people who appear to be actively looking for jobs across public forums and exports the results to `job_seekers.xlsx`.

## Sources

- Reddit via PRAW:
  - `r/foicontratado`
  - `r/brdev`
  - `r/devbrasil`
  - `r/programacao`
  - `r/cscareerquestions`
- Hacker News via Algolia public API
- TabNews via `https://www.tabnews.com.br/api/v1/contents`
- GUJ via HTML scraping

## Output

Running the scraper creates an Excel file named `job_seekers.xlsx` with:

- `Leads`: one consolidated sheet sorted by date descending
- One sheet per source: `Reddit`, `HackerNews`, `TabNews`, `GUJ`
- A second-stage matcher can generate `job_referral_matches.xlsx` with candidate-to-job bundles from Micro1 referral opportunities

Every row uses this schema:

`source | forum | author | title | content | url | date | location_country | technologies | experience | keywords_matched`

## Setup

1. Create and activate a virtual environment.

```powershell
python -m venv .venv
.venv\Scripts\Activate.ps1
```

2. Install the dependencies.

```powershell
python -m pip install -r requirements.txt
```

3. Configure Reddit credentials in environment variables.

```powershell
$env:REDDIT_CLIENT_ID="your_client_id"
$env:REDDIT_CLIENT_SECRET="your_client_secret"
```

If you want them to persist across new PowerShell sessions on Windows, use:

```powershell
setx REDDIT_CLIENT_ID "your_client_id"
setx REDDIT_CLIENT_SECRET "your_client_secret"
```

4. Run the scraper.

```powershell
python scraper.py
```

## Referral matching

After generating `job_seekers.xlsx`, you can match candidates against recent Micro1 referral jobs and export bundles of up to 5 links per candidate.

Preferred option: set your Micro1 authorization header from a logged-in referral session.

```powershell
$env:MICRO1_AUTH_TOKEN="your_token_or_bearer_value"
python match_referrals.py
```

If you want it to persist across PowerShell sessions:

```powershell
setx MICRO1_AUTH_TOKEN "your_token_or_bearer_value"
```

If you already saved the API response to a local JSON file, you can run without live API access:

```powershell
python match_referrals.py --jobs-json C:\path\to\eligible_jobs.json
```

The matcher:

- Paginates through the Micro1 `eligible-jobs` endpoint
- Filters jobs posted in the last 30 days
- Filters candidates to recent leads as well, using a 30-day lookback by default
- Focuses on structured Hacker News candidates by default
- Produces `job_referral_matches.xlsx` with `Candidates`, `Micro1Jobs`, `Matches`, and `Bundles` sheets

## How to get Reddit API keys

1. Sign in to Reddit.
2. Go to `https://www.reddit.com/prefs/apps`.
3. Click `create another app...`.
4. Choose `script`.
5. Fill in the name and redirect URI. A placeholder like `http://localhost:8080` is enough for read-only use.
6. After saving:
   - `client_id` is the short string shown under the app name.
   - `client_secret` is the secret field for the app.

## Notes

- The scraper deduplicates leads by URL before export.
- Each source runs independently. If one source fails, the others still run.
- The script sleeps between requests to reduce the chance of rate-limit issues.
- `location_country`, `technologies`, and `experience` are best-effort structured fields. They are most reliable for Hacker News "Who wants to be hired?" posts that follow a semi-structured format.
- `match_referrals.py` expects either `MICRO1_AUTH_TOKEN` / `MICRO1_AUTH_HEADER` or a saved `eligible-jobs` JSON payload.
- The GUJ module first tries `https://www.guj.com.br/c/empregos`. If that category is unavailable, it falls back to the GUJ homepage and scans recent topic links there.
