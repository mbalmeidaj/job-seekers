"""Match scraped job seekers to recent Micro1 referral opportunities."""

from __future__ import annotations

import argparse
import json
import logging
import math
import os
import re
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any

import pandas as pd
import requests

from scraper import auto_adjust_column_widths, normalize_text, sanitize_excel_text


INPUT_FILE = "job_seekers.xlsx"
OUTPUT_FILE = "job_referral_matches.xlsx"
DEFAULT_SHEET_NAME = "HackerNews"
DEFAULT_JOB_LOOKBACK_DAYS = 30
DEFAULT_CANDIDATE_LOOKBACK_DAYS = 14
DEFAULT_PAGE_SIZE = 100
DEFAULT_BUNDLE_SIZE = 5
MICRO1_JOBS_ENDPOINT = "https://prod-api.micro1.ai/api/v1/referral/portal/eligible-jobs"

US_STATE_NAMES = {
    "alabama", "alaska", "arizona", "arkansas", "california", "colorado", "connecticut", "delaware",
    "florida", "georgia", "hawaii", "idaho", "illinois", "indiana", "iowa", "kansas", "kentucky",
    "louisiana", "maine", "maryland", "massachusetts", "michigan", "minnesota", "mississippi", "missouri",
    "montana", "nebraska", "nevada", "new hampshire", "new jersey", "new mexico", "new york",
    "north carolina", "north dakota", "ohio", "oklahoma", "oregon", "pennsylvania", "rhode island",
    "south carolina", "south dakota", "tennessee", "texas", "utah", "vermont", "virginia", "washington",
    "west virginia", "wisconsin", "wyoming", "district of columbia",
}

COUNTRY_NORMALIZATION_MAP = {
    "england": "United Kingdom",
    "scotland": "United Kingdom",
    "wales": "United Kingdom",
    "northern ireland": "United Kingdom",
}

TECH_DISPLAY_MAP = {
    "aws": "AWS",
    "css": "CSS",
    "docker": "Docker",
    "fastapi": "FastAPI",
    "figma": "Figma",
    "go": "Go",
    "html": "HTML",
    "ios": "iOS",
    "java": "Java",
    "javascript": "JavaScript",
    "kotlin": "Kotlin",
    "llm": "LLM",
    "node": "Node",
    "postgres": "Postgres",
    "pytorch": "PyTorch",
    "python": "Python",
    "react": "React",
    "sql": "SQL",
    "swift": "Swift",
    "typescript": "TypeScript",
    "ui": "UI",
    "ux": "UX",
}

LOW_SIGNAL_JOB_SKILLS = {
    "written communication",
    "verbal communication",
    "attention to detail",
    "problem-solving",
    "problem solving",
    "critical thinking",
    "teamwork",
    "collaboration",
    "remote collaboration",
    "remote work capability",
    "remote work proficiency",
    "organization & attention to detail",
    "strong communication",
}

EXPERIENCE_TO_JOB_KEYWORDS = {
    "Product Engineer": ["product engineer", "founding engineer", "software engineer", "full stack", "backend", "frontend", "developer"],
    "Data Engineer": ["data engineer", "analytics engineer", "data platform", "etl", "data infrastructure", "pipeline"],
    "Data Scientist": ["data scientist", "machine learning", "data analysis", "research", "analytics"],
    "ML Engineer": ["machine learning", "ml engineer", "ai engineer", "genai", "llm", "applied ai", "artificial intelligence"],
    "AI Researcher": ["research engineer", "research scientist", "ai researcher", "ml researcher", "ai"],
    "Backend Engineer": ["backend engineer", "software engineer", "platform engineer", "distributed systems", "services"],
    "Frontend Engineer": ["frontend engineer", "front-end engineer", "ui engineer", "web engineer", "react", "javascript", "typescript"],
    "Full-Stack Engineer": ["full stack", "software engineer", "product engineer", "backend", "frontend", "developer"],
    "Software Developer": ["software engineer", "developer", "software developer", "engineer"],
    "Mobile Engineer": ["mobile engineer", "ios", "android", "swift", "kotlin"],
    "DevOps / Platform Engineer": ["devops", "platform engineer", "cloud", "infrastructure", "sre", "site reliability"],
    "Quant / Finance": ["finance", "investment", "quant", "valuation", "financial modeling", "capital markets", "trading"],
}


def configure_logging() -> None:
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
    )


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Match job seekers to Micro1 referral jobs.")
    parser.add_argument("--input", default=INPUT_FILE, help="Input workbook with scraped job seekers.")
    parser.add_argument("--output", default=OUTPUT_FILE, help="Output workbook for referral bundles.")
    parser.add_argument("--sheet", default=DEFAULT_SHEET_NAME, help="Workbook sheet to use as candidate source.")
    parser.add_argument("--lookback-days", type=int, default=DEFAULT_JOB_LOOKBACK_DAYS, help="Only keep jobs posted within this many days.")
    parser.add_argument(
        "--candidate-lookback-days",
        type=int,
        default=DEFAULT_CANDIDATE_LOOKBACK_DAYS,
        help="Only keep candidates whose lead date is within this many days.",
    )
    parser.add_argument("--page-size", type=int, default=DEFAULT_PAGE_SIZE, help="Micro1 API page size.")
    parser.add_argument("--bundle-size", type=int, default=DEFAULT_BUNDLE_SIZE, help="Max referral links to keep per candidate.")
    parser.add_argument(
        "--jobs-json",
        default=os.getenv("MICRO1_JOBS_JSON", ""),
        help="Optional path to a saved Micro1 eligible-jobs JSON response or a JSON array of jobs.",
    )
    return parser.parse_args()


def build_session() -> requests.Session:
    session = requests.Session()
    session.headers.update({"User-Agent": "job-seeker-referral-matcher/1.0"})
    return session


def build_authorization_header() -> str:
    explicit_header = os.getenv("MICRO1_AUTH_HEADER", "").strip()
    if explicit_header:
        return explicit_header

    token = os.getenv("MICRO1_AUTH_TOKEN", "").strip()
    if not token:
        return ""
    if token.lower().startswith("bearer "):
        return token
    return f"Bearer {token}"


def parse_micro1_datetime(value: str | None) -> datetime | None:
    if not value:
        return None
    try:
        return datetime.strptime(value, "%Y-%m-%d %H:%M:%S")
    except ValueError:
        return None


def normalize_job_record(job: dict[str, Any]) -> dict[str, Any]:
    skills = [sanitize_excel_text(str(skill).strip()) for skill in (job.get("skills") or []) if str(skill).strip()]
    posted_at = parse_micro1_datetime(job.get("date_posted"))
    return {
        "job_id": sanitize_excel_text(str(job.get("job_id") or "")),
        "job_name": sanitize_excel_text(str(job.get("job_name") or "").strip()),
        "company_name": sanitize_excel_text(str(job.get("company_name") or "").strip()),
        "engagement_type": sanitize_excel_text(str(job.get("engagement_type") or "").strip()),
        "location_type": sanitize_excel_text(str(job.get("location_type") or "").strip()),
        "date_posted": sanitize_excel_text(str(job.get("date_posted") or "").strip()),
        "posted_at": posted_at,
        "no_of_openings": job.get("no_of_openings"),
        "is_high_demand_job": str(job.get("is_high_demand_job") or "") == "1",
        "skills": skills,
        "skills_text": ", ".join(skills),
        "referral_reward_amount": job.get("referral_reward_amount"),
        "apply_url": sanitize_excel_text(str(job.get("apply_url") or "").strip()),
        "ideal_hourly_rate_min": ((job.get("ideal_hourly_rate") or {}) if isinstance(job.get("ideal_hourly_rate"), dict) else {}).get("min"),
        "ideal_hourly_rate_max": ((job.get("ideal_hourly_rate") or {}) if isinstance(job.get("ideal_hourly_rate"), dict) else {}).get("max"),
    }


def load_jobs_from_json(path_str: str) -> list[dict[str, Any]]:
    path = Path(path_str)
    payload = json.loads(path.read_text(encoding="utf-8"))
    if isinstance(payload, dict):
        if isinstance(payload.get("data"), list):
            return [normalize_job_record(job) for job in payload["data"]]
        raise ValueError("JSON object does not contain a 'data' list.")
    if isinstance(payload, list):
        return [normalize_job_record(job) for job in payload]
    raise ValueError("Unsupported JSON payload format for jobs.")


def fetch_recent_micro1_jobs(session: requests.Session, *, lookback_days: int, page_size: int) -> list[dict[str, Any]]:
    auth_header = build_authorization_header()
    if not auth_header:
        raise RuntimeError(
            "Micro1 auth is missing. Set MICRO1_AUTH_TOKEN or MICRO1_AUTH_HEADER, or pass --jobs-json with a saved response."
        )

    session.headers["Authorization"] = auth_header
    cutoff = datetime.now() - timedelta(days=lookback_days)
    page = 1
    total_pages: int | None = None
    recent_jobs: list[dict[str, Any]] = []

    while True:
        params = {
            "page": page,
            "limit": page_size,
            "sort_by": "date_created",
            "sort_direction": "desc",
        }
        response = session.get(MICRO1_JOBS_ENDPOINT, params=params, timeout=30)
        response.raise_for_status()
        payload = response.json()

        if not payload.get("status"):
            raise RuntimeError(f"Micro1 API returned an unsuccessful response: {payload}")

        raw_jobs = payload.get("data") or []
        total = int(payload.get("total") or 0)
        total_pages = max(math.ceil(total / page_size), 1)
        logging.info("Fetched Micro1 page %s/%s with %s jobs", page, total_pages, len(raw_jobs))

        page_jobs = [normalize_job_record(job) for job in raw_jobs]
        fresh_jobs = [job for job in page_jobs if job["posted_at"] and job["posted_at"] >= cutoff]
        recent_jobs.extend(fresh_jobs)

        if not raw_jobs or page >= total_pages:
            break

        oldest_page_job = min((job["posted_at"] for job in page_jobs if job["posted_at"]), default=None)
        if oldest_page_job and oldest_page_job < cutoff:
            break

        page += 1

    deduped: dict[str, dict[str, Any]] = {}
    for job in recent_jobs:
        deduped[job["job_id"]] = job
    return list(deduped.values())


def load_candidates(input_path: str, sheet_name: str) -> pd.DataFrame:
    dataframe = pd.read_excel(input_path, sheet_name=sheet_name).fillna("")
    if "forum" in dataframe.columns:
        dataframe = dataframe[dataframe["forum"].astype(str).str.strip() == "Who wants to be hired?"]
    if {"experience", "technologies"}.issubset(dataframe.columns):
        mask = (dataframe["experience"].astype(str).str.strip() != "") | (dataframe["technologies"].astype(str).str.strip() != "")
        dataframe = dataframe[mask]
    dataframe = dataframe.copy()
    dataframe["candidate_date"] = pd.to_datetime(dataframe.get("date", ""), errors="coerce")
    dataframe["candidate_richness"] = (
        dataframe.get("experience", "").astype(str).str.len()
        + dataframe.get("technologies", "").astype(str).str.len()
        + dataframe.get("content", "").astype(str).str.len()
    )
    if "author" in dataframe.columns:
        dataframe = dataframe.sort_values(
            ["candidate_date", "candidate_richness"],
            ascending=[False, False],
            na_position="last",
        ).drop_duplicates(subset=["author"], keep="first")
    dataframe = dataframe.drop(columns=["candidate_date", "candidate_richness"])
    dataframe = dataframe.reset_index(drop=True)
    dataframe.insert(0, "candidate_id", [f"CAND-{index + 1:04d}" for index in dataframe.index])
    return dataframe


def filter_recent_candidates(candidates_df: pd.DataFrame, lookback_days: int) -> pd.DataFrame:
    if "date" not in candidates_df.columns:
        return candidates_df

    cutoff = datetime.now().date() - timedelta(days=lookback_days)
    candidate_dates = pd.to_datetime(candidates_df["date"], errors="coerce").dt.date
    mask = candidate_dates.notna() & (candidate_dates >= cutoff)
    filtered = candidates_df[mask].sort_values("date", ascending=False).reset_index(drop=True)
    if "candidate_id" in filtered.columns:
        filtered["candidate_id"] = [f"CAND-{index + 1:04d}" for index in filtered.index]
    return filtered


def normalize_candidate_country(value: str) -> str:
    cleaned = sanitize_excel_text(str(value).strip())
    normalized = normalize_text(cleaned)
    if not normalized:
        return ""
    if normalized in US_STATE_NAMES:
        return "United States"
    return COUNTRY_NORMALIZATION_MAP.get(normalized, cleaned)


def split_items(value: str) -> list[str]:
    if not value:
        return []
    parts: list[str] = []
    for chunk in re.split(r"[,\n;]+", value):
        cleaned = sanitize_excel_text(chunk.strip())
        if cleaned:
            parts.append(cleaned)
        for extra in re.split(r"[/|()]+", cleaned):
            extra_clean = sanitize_excel_text(extra.strip())
            if extra_clean and extra_clean not in parts:
                parts.append(extra_clean)
    return parts


def candidate_role_keywords(experience_value: str) -> dict[str, list[str]]:
    role_map: dict[str, list[str]] = {}
    for role in split_items(experience_value):
        keywords = EXPERIENCE_TO_JOB_KEYWORDS.get(role, [role])
        role_map[role] = keywords
    return role_map


def candidate_tech_keywords(technologies_value: str) -> list[str]:
    normalized_terms: list[str] = []
    for term in split_items(technologies_value):
        normalized = normalize_text(term)
        if normalized and normalized not in normalized_terms:
            normalized_terms.append(normalized)
    return normalized_terms


def filter_low_signal_skills(skills: list[str]) -> list[str]:
    filtered: list[str] = []
    for skill in skills:
        normalized = normalize_text(skill)
        if not normalized or normalized in LOW_SIGNAL_JOB_SKILLS:
            continue
        filtered.append(skill)
    return filtered


def phrase_in_text(phrase: str, text: str) -> bool:
    normalized_phrase = normalize_text(phrase)
    normalized_text = normalize_text(text)
    if not normalized_phrase or not normalized_text:
        return False
    return re.search(rf"(?<!\w){re.escape(normalized_phrase)}(?!\w)", normalized_text) is not None


def obfuscate_url(url: str) -> str:
    cleaned = sanitize_excel_text(url.strip())
    if not cleaned:
        return ""
    obfuscated = cleaned.replace("https://", "http s: //").replace("http://", "http: //")
    obfuscated = obfuscated.replace(".", " . ")
    obfuscated = obfuscated.replace("?", " ?")
    obfuscated = obfuscated.replace("&", " &")
    return re.sub(r"\s+", " ", obfuscated).strip()


def summarize_match_reason(candidate: pd.Series, job: dict[str, Any], reasons_text: str) -> str:
    normalized_reasons = normalize_text(reasons_text)
    experience_text = sanitize_excel_text(str(candidate.get("experience", "")).strip())

    if "role overlap" in normalized_reasons and "tech overlap" in normalized_reasons:
        if experience_text:
            return f"It looked like the closest match because both the role focus and the stack line up well with your {experience_text} background."
        return "It looked like the closest match because both the role focus and the stack line up well with what you described."
    if "tech overlap" in normalized_reasons:
        return "It stood out mainly because the stack lines up with the technologies mentioned in your post."
    if "role overlap" in normalized_reasons:
        return "It stood out mainly because the role focus lines up with the experience described in your post."
    if "skill overlap" in normalized_reasons:
        return "It looked relevant because several of the listed skills overlap with what you shared publicly."
    return f"It looked like the closest match to the role and skill signals in the {sanitize_excel_text(job['job_name'])} bundle."


def get_job_skill_values(job: dict[str, Any]) -> list[str]:
    raw_skills = job.get("skills")
    if isinstance(raw_skills, list):
        return [sanitize_excel_text(str(skill)) for skill in raw_skills if str(skill).strip()]
    return split_items(sanitize_excel_text(str(job.get("skills_text", ""))))


def extract_job_focus_label(candidate: pd.Series, job: dict[str, Any]) -> str:
    job_text = " ".join([job["job_name"], " ".join(get_job_skill_values(job))])
    parenthetical_match = re.search(r"\(([^)]+)\)", job["job_name"])
    if parenthetical_match:
        label = sanitize_excel_text(parenthetical_match.group(1))
        label = re.sub(r"\s*[\+/]\s*", " and ", label)
        return label

    focus_terms: list[str] = []
    for tech in candidate_tech_keywords(str(candidate.get("technologies", ""))):
        if tech in TECH_DISPLAY_MAP and phrase_in_text(tech, job_text):
            display_value = TECH_DISPLAY_MAP[tech]
            if display_value not in focus_terms:
                focus_terms.append(display_value)

    if len(focus_terms) >= 2:
        return " and ".join(focus_terms[:2])

    return sanitize_excel_text(job["job_name"])


def summarize_alternative_jobs(alternative_jobs: list[dict[str, Any]]) -> str:
    titles = [sanitize_excel_text(job["job_name"]) for job in alternative_jobs if job.get("job_name")]
    if not titles:
        return ""
    if len(titles) == 1:
        return f"There is also a related role on the same job board: {titles[0]}."
    return f"There are also a couple of related roles on the same job board, including {titles[0]} and {titles[1]}."


def score_job_for_candidate(candidate: pd.Series, job: dict[str, Any]) -> tuple[float, list[str]]:
    job_text = " ".join([job["job_name"], " ".join(job["skills"])])
    candidate_text = (
        " ".join(
            [
                str(candidate.get("title", "")),
                str(candidate.get("content", "")),
                str(candidate.get("experience", "")),
                str(candidate.get("technologies", "")),
            ]
        )
    )
    candidate_text_norm = normalize_text(candidate_text)

    score = 0.0
    reasons: list[str] = []

    role_hits: list[str] = []
    for role_name, keywords in candidate_role_keywords(str(candidate.get("experience", ""))).items():
        for keyword in keywords:
            if phrase_in_text(keyword, job_text):
                role_hits.append(role_name)
                break
    if role_hits:
        score += min(len(set(role_hits)) * 14, 42)
        reasons.append(f"role overlap: {', '.join(dict.fromkeys(role_hits))}")

    tech_hits: list[str] = []
    for tech in candidate_tech_keywords(str(candidate.get("technologies", ""))):
        if len(tech) < 3:
            continue
        if phrase_in_text(tech, job_text):
            tech_hits.append(tech)
    if tech_hits:
        score += min(len(set(tech_hits)) * 4, 28)
        reasons.append(f"tech overlap: {', '.join(dict.fromkeys(tech_hits[:7]))}")

    skill_hits: list[str] = []
    for skill in filter_low_signal_skills(job["skills"]):
        if phrase_in_text(skill, candidate_text):
            skill_hits.append(skill)
    if skill_hits:
        score += min(len(set(skill_hits)) * 3, 24)
        reasons.append(f"skill overlap: {', '.join(dict.fromkeys(skill_hits[:6]))}")

    if job["job_name"] and phrase_in_text(job["job_name"], candidate_text):
        score += 10
        reasons.append("job title appears in candidate text")

    if job["is_high_demand_job"] and score > 0:
        score += 1.5
        reasons.append("high-demand job")

    return score, reasons


def build_matches(candidates_df: pd.DataFrame, jobs: list[dict[str, Any]], bundle_size: int) -> tuple[pd.DataFrame, pd.DataFrame]:
    match_rows: list[dict[str, Any]] = []
    bundle_rows: list[dict[str, Any]] = []

    for _, candidate in candidates_df.iterrows():
        scored_jobs: list[dict[str, Any]] = []
        for job in jobs:
            score, reasons = score_job_for_candidate(candidate, job)
            if score <= 0:
                continue
            scored_jobs.append(
                {
                    "candidate_id": candidate["candidate_id"],
                    "author": sanitize_excel_text(str(candidate.get("author", ""))),
                    "candidate_title": sanitize_excel_text(str(candidate.get("title", ""))),
                    "candidate_experience": sanitize_excel_text(str(candidate.get("experience", ""))),
                    "candidate_technologies": sanitize_excel_text(str(candidate.get("technologies", ""))),
                    "job_id": job["job_id"],
                    "job_name": job["job_name"],
                    "company_name": job["company_name"],
                    "date_posted": job["date_posted"],
                    "engagement_type": job["engagement_type"],
                    "location_type": job["location_type"],
                    "skills_text": job["skills_text"],
                    "referral_reward_amount": job["referral_reward_amount"],
                    "apply_url": job["apply_url"],
                    "score": round(score, 2),
                    "match_reasons": sanitize_excel_text(" | ".join(reasons)),
                }
            )

        scored_jobs.sort(
            key=lambda row: (
                row["score"],
                row["referral_reward_amount"] if row["referral_reward_amount"] is not None else -1,
                row["date_posted"],
            ),
            reverse=True,
        )

        top_jobs = scored_jobs[:bundle_size]
        match_rows.extend(top_jobs)

        best_job = top_jobs[0] if top_jobs else None
        alternative_jobs = top_jobs[1:3]
        experience_text = sanitize_excel_text(str(candidate.get("experience", "")).strip())
        intro_line = (
            f"Based on your {experience_text} background, one role that seemed especially relevant was {best_job['job_name']}."
            if best_job and experience_text
            else f"Based on your background, one role that seemed especially relevant was {best_job['job_name']}."
            if best_job
            else "Based on your background, I found a few roles that looked relevant."
        )
        best_reason = (
            summarize_match_reason(candidate, best_job, best_job["match_reasons"])
            if best_job and best_job.get("match_reasons")
            else "It looked like the closest match to the role and stack signals in your bundle."
        )
        alternatives_summary = summarize_alternative_jobs(alternative_jobs)
        best_focus_label = extract_job_focus_label(candidate, best_job) if best_job else ""
        recommendation_line = (
            f"I would start with the {best_focus_label} one and tailor the application to that specific stack rather than applying broadly."
            if best_job and best_focus_label and normalize_text(best_focus_label) != normalize_text(best_job["job_name"])
            else f"I would start with the {best_job['job_name']} one and tailor the application to that specific role rather than applying broadly."
            if best_job
            else ""
        )

        message_lines = [
            intro_line,
            best_reason,
        ]
        if alternatives_summary:
            message_lines.extend(["", alternatives_summary])
        if recommendation_line:
            message_lines.extend(["", recommendation_line])
        if best_job:
            message_lines.extend(
                [
                    "",
                    f"You also could try this (remove the spaces) {obfuscate_url(best_job['apply_url'])}",
                ]
            )
        suggested_message = sanitize_excel_text("\n".join(message_lines))

        bundle_links = "\n".join(
            sanitize_excel_text(f"{index + 1}. {job['job_name']} - {obfuscate_url(job['apply_url'])}")
            for index, job in enumerate(top_jobs)
        )
        bundle_row: dict[str, Any] = {
            "Date": sanitize_excel_text(str(candidate.get("date", ""))),
            "Author": sanitize_excel_text(str(candidate.get("author", ""))),
            "Country": normalize_candidate_country(str(candidate.get("location_country", ""))),
            "Experience": sanitize_excel_text(str(candidate.get("experience", ""))),
            "URL da postagem": sanitize_excel_text(str(candidate.get("url", ""))),
            "Link dos bundles": bundle_links,
            "Mensagem sugerida": suggested_message,
        }

        if top_jobs:
            bundle_rows.append(bundle_row)

    matches_df = pd.DataFrame(match_rows)
    if not matches_df.empty:
        matches_df = matches_df.sort_values(["candidate_id", "score"], ascending=[True, False]).reset_index(drop=True)

    bundles_df = pd.DataFrame(bundle_rows)
    return matches_df, bundles_df


def jobs_to_dataframe(jobs: list[dict[str, Any]]) -> pd.DataFrame:
    records = []
    for job in jobs:
        records.append(
            {
                "job_id": job["job_id"],
                "job_name": job["job_name"],
                "company_name": job["company_name"],
                "date_posted": job["date_posted"],
                "engagement_type": job["engagement_type"],
                "location_type": job["location_type"],
                "skills_text": job["skills_text"],
                "referral_reward_amount": job["referral_reward_amount"],
                "ideal_hourly_rate_min": job["ideal_hourly_rate_min"],
                "ideal_hourly_rate_max": job["ideal_hourly_rate_max"],
                "is_high_demand_job": job["is_high_demand_job"],
                "apply_url": job["apply_url"],
            }
        )
    return pd.DataFrame(records)


def export_matches(
    *,
    output_path: str,
    candidates_df: pd.DataFrame,
    jobs_df: pd.DataFrame,
    matches_df: pd.DataFrame,
    bundles_df: pd.DataFrame,
) -> None:
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        candidates_df.to_excel(writer, sheet_name="Candidates", index=False)
        jobs_df.to_excel(writer, sheet_name="Micro1Jobs", index=False)
        matches_df.to_excel(writer, sheet_name="Matches", index=False)
        bundles_df.to_excel(writer, sheet_name="Bundles", index=False)
        auto_adjust_column_widths(writer)


def main() -> None:
    configure_logging()
    args = parse_args()

    candidates_df = load_candidates(args.input, args.sheet)
    candidates_df = filter_recent_candidates(candidates_df, args.candidate_lookback_days)
    logging.info(
        "Loaded %s recent candidate rows from sheet %s (lookback=%s days)",
        len(candidates_df),
        args.sheet,
        args.candidate_lookback_days,
    )

    if args.jobs_json:
        jobs = load_jobs_from_json(args.jobs_json)
        cutoff = datetime.now() - timedelta(days=args.lookback_days)
        jobs = [job for job in jobs if job["posted_at"] and job["posted_at"] >= cutoff]
        logging.info("Loaded %s recent jobs from local JSON", len(jobs))
    else:
        session = build_session()
        jobs = fetch_recent_micro1_jobs(session, lookback_days=args.lookback_days, page_size=args.page_size)
        logging.info("Fetched %s recent jobs from Micro1 API", len(jobs))

    matches_df, bundles_df = build_matches(candidates_df, jobs, args.bundle_size)
    jobs_df = jobs_to_dataframe(jobs)
    export_matches(
        output_path=args.output,
        candidates_df=candidates_df,
        jobs_df=jobs_df,
        matches_df=matches_df,
        bundles_df=bundles_df,
    )

    logging.info(
        "Exported %s jobs, %s candidate-job matches, and %s bundles to %s",
        len(jobs_df),
        len(matches_df),
        len(bundles_df),
        args.output,
    )


if __name__ == "__main__":
    main()
