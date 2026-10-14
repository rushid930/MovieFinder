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
