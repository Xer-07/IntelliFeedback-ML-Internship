# Data Directory

This directory stores raw and processed datasets for the **IntelliFeedback** Customer Sentiment Analysis project.

## Dataset Overview
- **Primary Source:** Public E-Commerce / Consumer Reviews Benchmark (e.g., Amazon Product Reviews / Kaggle Consumer Reviews).
- **Format:** CSV / JSON
- **Target Attributes:**
  - `review_id`: Unique identifier
  - `review_text`: Unstructured text written by the customer
  - `rating`: Star rating (1 to 5)
  - `sentiment`: Target label mapped as:
    - `0`: Negative (1 - 2 Stars)
    - `1`: Neutral (3 Stars)
    - `2`: Positive (4 - 5 Stars)
  - `timestamp`: Date/time of review posting

> **Note:** Raw `.csv` files are excluded from Git version control via `.gitignore` to keep repository size lean. Data ingestion scripts in `src/` automate download and local storage.
