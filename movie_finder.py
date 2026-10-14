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

# ============================================================
# USER INTERFACE
# ============================================================

st.title("🎬 Movie Finder")
st.caption(
    "Search The Movie Database (TMDB) and explore movie details."
)

if not TMDB_TOKEN:
    st.error(
        "TMDB_TOKEN is missing. Add your TMDB Read Access "
        "Token to the .env file and restart the app."
    )
    st.stop()


# Search form
with st.form("movie_search_form"):
    st.subheader("Search for a movie")

    col1, col2 = st.columns([3, 1])

    with col1:
        movie_title = st.text_input(
            "Movie title",
            placeholder="e.g. The Dark Knight",
        )

    with col2:
        movie_year = st.number_input(
            "Release year (optional)",
            min_value=0,
            max_value=2100,
            value=0,
            help="Leave as 0 if you don't know the year.",
        )

    submitted = st.form_submit_button(
        "🔎 Search Movies",
        type="primary",
        use_container_width=True,
    )


# Search TMDB when the form is submitted
if submitted:
    if not movie_title.strip():
        st.session_state["movie_candidates"] = []
        st.warning("Please enter a movie title.")
    else:
        try:
            with st.spinner("Searching TMDB..."):
                candidates = search_movie(
                    movie_title,
                    int(movie_year) if movie_year else None,
                )

            st.session_state["movie_candidates"] = candidates
            st.session_state["search_title"] = movie_title.strip()

            # Reset the selected result after a new search
            st.session_state.pop("selected_movie_id", None)
            st.session_state.pop("movie_details", None)

            if not candidates:
                st.info("No matching movies found. Try another title or year.")

        except requests.RequestException as exc:
            st.error(f"TMDB search failed: {exc}")


# Movie selection
candidates = st.session_state.get("movie_candidates", [])

if candidates:
    st.subheader("Matching movies")

    def movie_label(movie):
        title = movie.get("title") or "Untitled"
        date = movie.get("release_date") or ""
        year = date[:4] if date else "Year unknown"
        original = movie.get("original_title", "")

        label = f"{title} ({year})"

        if original and original != title:
            label += f" — Original title: {original}"

        return label

    selected_movie = st.selectbox(
        "Choose the movie you want",
        options=candidates,
        format_func=movie_label,
        key="movie_choice",
    )

    if st.button(
        "Get Movie Details",
        type="primary",
        use_container_width=False,
    ):
        try:
            with st.spinner("Loading movie details and credits..."):
                details = get_movie_details(selected_movie["id"])

            st.session_state["movie_details"] = details
            st.session_state["selected_movie_id"] = selected_movie["id"]

        except requests.RequestException as exc:
            st.error(f"Could not load movie details: {exc}")


# Results table
details = st.session_state.get("movie_details")

if details:
    st.divider()
    st.subheader("Movie Details")

    row = make_movie_row(details)
    df = pd.DataFrame([row])

    st.dataframe(
        df,
        use_container_width=True,
        hide_index=True,
        column_config={
            "Director": st.column_config.TextColumn(
                "Director",
                help="Director name(s), formatted Last, First",
            ),
            "Title": st.column_config.TextColumn("Title"),
            "Year": st.column_config.TextColumn("Year"),
            "Country": st.column_config.TextColumn(
                "Country",
                help="Production countries",
            ),
            "Genres": st.column_config.TextColumn(
                "Genres",
                help="All listed movie genres",
            ),
        },
    )

    # Optional overview
    overview = details.get("overview", "")
    if overview:
        with st.expander("Movie overview"):
            st.write(overview)

    # Export buttons
    st.subheader("Export results")

    export_col1, export_col2, _ = st.columns([1, 1, 3])

    with export_col1:
        st.download_button(
            "⬇️ Download Excel",
            data=make_excel([row]),
            file_name="movie_details.xlsx",
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
            use_container_width=True,
        )

    with export_col2:
        csv_data = pd.DataFrame([row]).to_csv(
            index=False
        ).encode("utf-8-sig")

        st.download_button(
            "⬇️ Download CSV",
            data=csv_data,
            file_name="movie_details.csv",
            mime="text/csv",
            use_container_width=True,
        )

st.divider()
st.caption("Movie metadata provided by TMDB.")