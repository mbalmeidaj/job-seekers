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

Every row uses this schema:

`source | forum | author | title | content | url | date | keywords_matched`

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
- The GUJ module first tries `https://www.guj.com.br/c/empregos`. If that category is unavailable, it falls back to the GUJ homepage and scans recent topic links there.
