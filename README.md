<div align="center">

<img src="https://img.shields.io/badge/Python-3.8+-3776AB?style=for-the-badge&logo=python&logoColor=white"/>
<img src="https://img.shields.io/badge/Pandas-150458?style=for-the-badge&logo=pandas&logoColor=white"/>
<img src="https://img.shields.io/badge/Matplotlib-11557C?style=for-the-badge&logo=python&logoColor=white"/>
<img src="https://img.shields.io/badge/Seaborn-4C72B0?style=for-the-badge&logo=python&logoColor=white"/>
<img src="https://img.shields.io/badge/Microsoft_Excel-217346?style=for-the-badge&logo=microsoft-excel&logoColor=white"/>
<img src="https://img.shields.io/badge/Kaggle-20BEFF?style=for-the-badge&logo=kaggle&logoColor=white"/>

<br/><br/>

# TMDB Movie Credits Analysis

### E-Commerce Sales Analysis · 4,803 Movies · Cast & Crew Insights

<p>A data analysis project that parses <strong>complex nested JSON data</strong> from the TMDB 5000 Movie Credits dataset<br/>
to extract insights about directors, actors, and production team sizes —<br/>
visualized with Matplotlib and exported to a multi-sheet Excel workbook.</p>

<br/>

| Movies Analyzed | Cast Records | Crew Records | Charts Generated |
|:---:|:---:|:---:|:---:|
| **4,803** | **~150,000+** | **~100,000+** | **5 PNG Charts** |

</div>

---

## Table of Contents

- [Project Overview](#-project-overview)
- [Files](#-files)
- [Quick Start](#-quick-start)
- [How It Works](#-how-it-works)
- [Visualizations](#-visualizations)
- [Key Insights](#-key-insights)
- [Excel Output](#-excel-output)
- [Project Structure](#-project-structure)
- [Tech Stack](#-tech-stack)
- [Future Improvements](#-future-improvements)

---

## Project Overview

This project analyzes the **TMDB 5000 Movie Credits dataset** to uncover patterns in the film industry — identifying the most prolific directors and lead actors, understanding typical production team sizes, and exploring the relationship between cast and crew.

**The core challenge:** The `cast` and `crew` columns don't contain simple text — each cell holds a full **JSON list embedded as a plain string**, with detailed info on every person involved in the film. This required custom parsing before any analysis could begin, making it a real-world data wrangling exercise.

**What this project delivers:**
- Raw JSON strings → clean, structured DataFrame columns
- 5 publication-quality charts saved as PNG files
- Multi-sheet Excel workbook ready to share with non-technical stakeholders

---

## Files

| File | Purpose |
|:---|:---|
| `movie_analysis.py` | Main analysis script — parses JSON, generates 5 charts + Excel |
| `tmdb_5000_credits.csv` | Source dataset (download from Kaggle) |
| `tmdb_credits_analysis.xlsx` | Excel output — 4 sheets with summaries and stats |
| `top_directors.png` | Bar chart — top 10 most frequent directors |
| `top_actors.png` | Bar chart — top 10 most frequent lead actors |
| `cast_size_dist.png` | Histogram — cast size distribution |
| `crew_size_dist.png` | Histogram — crew size distribution |
| `cast_vs_crew.png` | Scatter plot — cast vs crew size correlation |
| `requirements.txt` | Python dependencies |

---

## Quick Start

### · Run the Analysis

```bash
pip install pandas matplotlib seaborn openpyxl
python movie_analysis.py
```

**Output:**
```
outputs/
├── top_directors.png
├── top_actors.png
├── cast_size_dist.png
├── crew_size_dist.png
├── cast_vs_crew.png
└── tmdb_credits_analysis.xlsx
```

> Download `tmdb_5000_credits.csv` from [Kaggle — TMDB 5000 Movie Dataset](https://www.kaggle.com/datasets/tmdb/tmdb-movie-metadata) and place it in the project root before running.

---

## How It Works

### Step 1 — Load the data

```python
import pandas as pd
import ast

df = pd.read_csv('tmdb_5000_credits.csv')
# 4,803 rows × 4 columns: movie_id, title, cast, crew
```

---

### Step 2 — Parse JSON columns

Each cell in `cast` and `crew` looks like this — a string, not a real list:

```
"[{'name': 'Sam Worthington', 'order': 0, 'character': 'Jake Sully'}, ...]"
```

We convert it into a real Python list using `ast.literal_eval()`:

```python
def parse_json(cell):
    try:
        return ast.literal_eval(cell)   # string → real Python list
    except:
        return []                        # safe fallback for nulls/broken cells

df['cast_parsed'] = df['cast'].apply(parse_json)
df['crew_parsed'] = df['crew'].apply(parse_json)
```

> `ast.literal_eval` is used instead of `json.loads` because the dataset uses **Python-style single quotes** inside strings, which strict JSON doesn't support.

---

### Step 3 — Extract structured fields

```python
# Director = crew member whose job is 'Director'
def get_director(crew_list):
    for member in crew_list:
        if member.get('job') == 'Director':
            return member.get('name')
    return None

# Top 3 cast = sorted by 'order' (0 = lead actor)
def get_top_cast(cast_list, n=3):
    sorted_cast = sorted(cast_list, key=lambda x: x.get('order', 99))
    return [c['name'] for c in sorted_cast[:n]]

df['director']   = df['crew_parsed'].apply(get_director)
df['lead_actor'] = df['cast_parsed'].apply(lambda x: x[0]['name'] if x else None)
df['top_cast']   = df['cast_parsed'].apply(get_top_cast)
df['cast_size']  = df['cast_parsed'].apply(len)
df['crew_size']  = df['crew_parsed'].apply(len)
```

---

### Step 4 — Visualize

```python
import matplotlib.pyplot as plt

top_directors = df['director'].value_counts().head(10)

plt.figure(figsize=(12, 6))
top_directors.sort_values().plot(kind='barh', color='steelblue', edgecolor='black')
plt.title('Top 10 Most Frequent Directors')
plt.tight_layout()
plt.savefig('top_directors.png', dpi=150)
plt.show()
```

---

### Step 5 — Export to Excel

```python
with pd.ExcelWriter('tmdb_credits_analysis.xlsx', engine='openpyxl') as writer:
    df_clean.to_excel(writer, sheet_name='All Movies', index=False)
    genre_summary.to_excel(writer, sheet_name='Top Directors')
    top_actors.to_excel(writer, sheet_name='Top Lead Actors')
    summary.to_excel(writer, sheet_name='Summary Stats')
```

---

## Visualizations

| # | File | Chart Type | What It Shows |
|:--|:---|:---|:---|
| 1 | `top_directors.png` | Horizontal bar | Top 10 directors by total movie count |
| 2 | `top_actors.png` | Horizontal bar | Top 10 lead actors by appearances |
| 3 | `cast_size_dist.png` | Histogram | Distribution of cast sizes across all films |
| 4 | `crew_size_dist.png` | Histogram | Distribution of crew sizes across all films |
| 5 | `cast_vs_crew.png` | Scatter plot | Relationship between cast size and crew size |

---

## Key Insights

```
 A small group of directors dominate — Spielberg, Scott & Eastwood appear most frequently
 Most movies have 15–50 cast members; ensemble films can exceed 200
 Crew size is consistently 3–4x larger than cast size
 Strong positive correlation between cast and crew size — big productions scale on both fronts
 The lead actor (order=0) changes significantly across genres and budget levels
```

---

## Excel Output

The file `tmdb_credits_analysis.xlsx` contains **4 sheets**:

| Sheet | Contents |
|:---|:---|
| `All Movies` | Full cleaned dataset — director, lead actor, top 3 cast, cast size, crew size |
| `Top Directors` | Top 20 directors ranked by movie count |
| `Top Lead Actors` | Top 20 lead actors ranked by appearances |
| `Summary Stats` | Avg / Max / Min for cast size and crew size |

---

## Project Structure

```
tmdb-credits-analysis/
│
├──  movie_analysis.py                ← main script
├──  ecommerce_sales_dataset.csv      ← source data (download from Kaggle)
├──  requirements.txt
├──  README.md
│
└──  outputs/
    ├──  top_directors.png
    ├──  top_actors.png
    ├──  cast_size_dist.png
    ├──  crew_size_dist.png
    ├──  cast_vs_crew.png
    └──  tmdb_credits_analysis.xlsx
```

---

## Tech Stack

| Tool | Version | Used For |
|:---|:---|:---|
| `Python` | 3.8+ | All scripting and data processing |
| `pandas` | ≥1.3 | Data loading, cleaning, and transformation |
| `ast` | built-in | Safe JSON string → Python list parsing |
| `matplotlib` | ≥3.4 | All chart generation (bar, histogram, scatter) |
| `seaborn` | ≥0.11 | Statistical chart styling and heatmap |
| `openpyxl` | ≥3.0 | Multi-sheet Excel workbook export |

---

## Future Improvements

- [ ] **Merge with `tmdb_5000_movies.csv`** — join ratings, genres, budget & revenue with cast/crew data
- [ ] **Director ratings analysis** — which directors consistently produce the highest-rated films?
- [ ] **Gender representation** — male vs female breakdown across roles and decades using the `gender` field
- [ ] **Actor-director network graph** — visualize collaboration frequency using NetworkX
- [ ] **Interactive dashboard** — Plotly or Streamlit for filterable, zoomable charts
- [ ] **NLP on character names** — cluster recurring character archetypes across genres

---

## Acknowledgements

- Dataset: [TMDB 5000 Movie Dataset](https://www.kaggle.com/datasets/tmdb/tmdb-movie-metadata) via Kaggle
- Data provided by [The Movie Database (TMDB)](https://www.themoviedb.org/)

---

## License

This project is for **educational and portfolio purposes**.  
Dataset is subject to [TMDB's Terms of Use](https://www.themoviedb.org/documentation/api/terms-of-use).

---

<div align="center">


Made with 🐍 Python · 🐼 Pandas · 📊 Matplotlib

</div>
