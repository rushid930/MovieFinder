import os
from io import BytesIO

import pandas as pd
import requests
import streamlit as st
from dotenv import load_dotenv

load_dotenv()

# ============================================================
# CONFIGURATION
# ============================================================

TMDB_TOKEN = os.getenv("TMDB_TOKEN")
BASE_URL = "https://api.themoviedb.org/3"
LANGUAGE = "en-US"

HEADERS = {
    "Authorization": f"Bearer {TMDB_TOKEN}",
    "accept": "application/json",
}

st.set_page_config(
    page_title="Movie Finder",
    page_icon="🎬",
    layout="wide",
)

# ============================================================
# TMDB API
# ============================================================

def tmdb_get(endpoint, params=None):
    """Make an authenticated request to TMDB."""
    response = requests.get(
        f"{BASE_URL}/{endpoint.lstrip('/')}",
        headers=HEADERS,
        params=params or {},
        timeout=30,
    )
    response.raise_for_status()
    return response.json()


def search_movie(title, year=None):
    """Return TMDB search results for a movie."""
    params = {
        "query": title.strip(),
        "language": LANGUAGE,
        "include_adult": "false",
    }

    if year:
        params["year"] = year

    data = tmdb_get("search/movie", params)
    results = data.get("results", [])

    # Prefer an exact title match, and then an exact year.
    normalized_title = title.strip().casefold()

    def match_score(movie):
        movie_title = movie.get("title", "").strip().casefold()
        release_date = movie.get("release_date", "")
        title_match = movie_title == normalized_title
        year_match = bool(
            year and release_date.startswith(str(year))
        )
        return (title_match, year_match)

    return sorted(results, key=match_score, reverse=True)


def get_movie_details(movie_id):
    """Get movie information and director credits."""
    return tmdb_get(
        f"movie/{movie_id}",
        {
            "language": LANGUAGE,
            "append_to_response": "credits",
        },
    )

# ============================================================
# DATA EXTRACTION
# ============================================================

def get_director(details):
    """Extract director names, formatting them Last, First."""
    crew = details.get("credits", {}).get("crew", [])
    directors = []

    for person in crew:
        if person.get("job") != "Director":
            continue

        name = person.get("name", "").strip()
        if not name:
            continue

        parts = name.split()

        if len(parts) >= 2:
            name = f"{parts[-1]}, {' '.join(parts[:-1])}"

        directors.append(name)

    return ", ".join(directors) or "Unknown"


def get_countries(details):
    countries = details.get("production_countries", [])
    names = [
        country.get("name", "")
        for country in countries
        if country.get("name")
    ]
    return ", ".join(names) or "Unknown"


def get_genres(details):
    genres = details.get("genres", [])
    names = [
        genre.get("name", "").strip()
        for genre in genres
        if genre.get("name", "").strip()
    ]
    return ", ".join(names) or "Unknown"


def make_movie_row(details):
    """Convert TMDB details into the requested table columns."""
    release_date = details.get("release_date", "")

    return {
        "Director": get_director(details),
        "Title": details.get("title") or "Unknown",
        "Year": release_date[:4] if release_date else "Unknown",
        "Country": get_countries(details),
        "Genres": get_genres(details),
    }

# ============================================================
# EXPORT
# ============================================================

def make_excel(rows):
    """Create an Excel workbook in memory."""
    output = BytesIO()
    df = pd.DataFrame(rows)

    with pd.ExcelWriter(
        output,
        engine="xlsxwriter",
    ) as writer:
        df.to_excel(
            writer,
            index=False,
            sheet_name="Movies",
        )

        workbook = writer.book
        worksheet = writer.sheets["Movies"]

        header_format = workbook.add_format({
            "bold": True,
            "bg_color": "#D9EAF7",
            "border": 1,
            "text_wrap": True,
        })

        cell_format = workbook.add_format({
            "valign": "top",
            "text_wrap": True,
        })

        for col_num, column in enumerate(df.columns):
            worksheet.write(
                0, col_num, column, header_format
            )

            max_length = max(
                len(str(column)),
                *(len(str(value)) for value in df[column]),
            )

            worksheet.set_column(
                col_num,
                col_num,
                min(max(max_length + 2, 12), 45),
                cell_format,
            )

        worksheet.freeze_panes(1, 0)
        worksheet.autofilter(
            0, 0, len(df), len(df.columns) - 1
        )

    output.seek(0)
    return output.getvalue()
