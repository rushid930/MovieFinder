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
