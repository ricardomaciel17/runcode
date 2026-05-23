# 🏃 Race Results Scraper

A **Python script** that collects the **latest race result for athletes** from the **OpenResults** website.

The program supports two input methods:

* 📊 **Excel file with athlete names**
* 🌐 **Event URL (RSF or OnSports / RaceZone)**

Results are processed and exported to an **Excel file**, with visual highlights for top performances.

---

# 📦 Features

* Automatically fetch the **latest race result for each athlete**
* Input via **Excel spreadsheet**
* Input via **event URL**
* Parallel processing using **ThreadPoolExecutor**
* Random delays between requests to avoid blocking
* Automatic **Excel export**
* Visual highlighting for top results

---

# 🗂 Project Structure

```
project/
│
├─ main.py
├─ scriptv1.py
├─ scriptv2.py
├─ utils.py
│
├─ corrida02.xlsx
└─ resultados.xlsx
```

| File        | Description                                                                  |
| ----------- | ---------------------------------------------------------------------------- |
| main.py     | Entry point that selects which version of the script to run                  |
| scriptv1.py | Processes athlete names from an Excel file                                   |
| scriptv2.py | Processes athletes from an event URL                                         |
| utils.py    | Utility functions (text normalization, slug creation, pace conversion, etc.) |

---

# ⚙️ Installation

Requires **Python 3.9 or higher**.

Install dependencies:

```bash
pip install requests pandas beautifulsoup4 openpyxl
```

---

# 🚀 Usage

Run the main script:

```bash
python main.py
```

The program will prompt you to choose a version:

```
1 = Excel
2 = URL
```

---

# 📊 Version 1 — Excel Input

Choosing **1** will:

1. Read the file `corrida02.xlsx`
2. Fetch the latest result for each athlete
3. Generate `resultados.xlsx`

### Excel Structure

The spreadsheet must contain a column named:

```
Nome
```

Example:

| Nome        |
| ----------- |
| João Silva  |
| Maria Souza |
| Pedro Costa |

---

# 🌐 Version 2 — Event URL

Choosing **2** will prompt for the event URL.

```
Enter event URL (RSF or OnSports):
```

---

## RSF Events

Flow:

1. Parse the event **XML**
2. List available **courses (percursos)**
3. Optionally filter by course

---

## OnSports / RaceZone Events

Flow:

1. Load event **JSON**
2. Group athletes by **gender + race code**
3. Optionally filter a specific group

---

# 📈 Collected Data

The script extracts the following fields:

| Field      | Description              |
| ---------- | ------------------------ |
| nome       | Athlete name             |
| modalidade | Race modality / distance |
| posicao    | Finishing position       |
| pace       | Average pace             |
| tempo      | Total race time          |

---

# 🎨 Excel Highlights

After generating `resultados.xlsx`, rows are colored automatically:

| Color    | Criteria                       |
| -------- | ------------------------------ |
| 🟦 Blue  | Top 10 finish                  |
| 🟩 Green | Pace faster than **4:00 / km** |

---

# ⚡ Performance

To avoid being blocked by the website:

* Random delay between requests
* Extra pause every **20 athletes**
* Up to **3 parallel threads**

---

# 📄 Example Output

```
Finished ✅
Names processed: 120
Failures: 7
```

---

# ⚠️ Notes

* The script depends on the current **OpenResults HTML structure**
* Changes in the website may require parser adjustments
* Athletes without available results may return `None`

---

# 📜 License

This project is intended for educational purposes and sports result analysis.
