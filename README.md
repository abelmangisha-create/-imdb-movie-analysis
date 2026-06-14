# IMDB Movie Insights Analysis (1933–2020)

An Excel-based data analysis project exploring box office performance, audience ratings, and movie trends across nearly 90 years of cinema.

---

## Dataset

- **Source:** IMDB Movies Dataset
- **Records:** 1,964 movies
- **Period:** 1933–2020
- **Fields:** Movie title, release year, age restriction, duration (minutes), IMDB rating, gross profit ($M), and votes

---

## Tools Used

- Microsoft Excel
  - Data cleaning and transformation
  - IFS, MAXIFS, XLOOKUP
  - Pivot Tables
  - Correlation Analysis

---

## What I Explored

### 1. Data Cleaning & Transformation
- Standardised raw IMDB data into a structured, analysis-ready table
- Classified movies into decades (1930s–2020s) using a custom `IFS` formula
- Mapped age restrictions to consistent categories (PG, PG-13, R, R18+)

### 2. Highest Grossing Movies
- Identified the highest grossing movie per decade using `MAXIFS`
- Retrieved corresponding movie titles using `XLOOKUP`
- Repeated the same analysis by age restriction category

### 3. Correlation Analysis
Tested whether gross profit could be predicted by:
- **IMDB Rating** → correlation of **0.19** (very weak)
- **Movie Duration** → correlation of **0.29** (weak)

### Key Finding
> There is no meaningful correlation between gross profit and ratings, nor between gross profit and movie duration. While weak positive correlations exist (0.19 and 0.29 respectively), they are too low to suggest any reliable predictive relationship.

Box office success is likely driven by factors not captured in this dataset — such as marketing spend, franchise recognition, and release timing.

---

## Project Structure

```
├── IMDB_Movie_Insights_Analysis.xlsx
│   ├── Messy IMDB_Movie_Insights     # Raw dataset
│   ├── Cleaned & Transformed         # Cleaned data with decade + age restriction columns
│   ├── Analysis                      # Main analysis sheet
│   ├── Pivot Tables                  # Breakdowns by decade and age restriction
│   ├── Highest Grossing Movies       # MAXIFS + XLOOKUP results
│   └── Correlation Analysis          # Correlation coefficients
```

---

## Key Formulas Used

**Decade classification:**
```excel
=IFS([@Year]<=1939,"1930s",[@Year]<=1949,"1940s",[@Year]<=1959,"1950s",
     [@Year]<=1969,"1960s",[@Year]<=1979,"1970s",[@Year]<=1989,"1980s",
     [@Year]<=1999,"1990s",[@Year]<=2009,"2000s",[@Year]<=2019,"2010s",
     [@Year]>=2020,"2020s")
```

**Highest grossing movie per decade:**
```excel
=XLOOKUP(MAXIFS([Gross Profit($M)],[Movies by Decade],"1930s"),
         [Gross Profit($M)],[Movie Name (Title)])
```
