import random
import time
import unicodedata
import requests
import xml.etree.ElementTree as ET
import pandas as pd
from concurrent.futures import ThreadPoolExecutor, as_completed
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from bs4 import BeautifulSoup

# -------------------
# Utilitários
# -------------------
def normalize_text(value):
    if value is None:
        return ""
    text = str(value)
    text = text.replace("\xa0", " ")
    text = unicodedata.normalize("NFD", text)
    text = text.encode("ascii", "ignore").decode("utf-8")
    text = " ".join(text.split())
    return text.strip()

def slugify(name):
    return normalize_text(name).lower().replace(" ", "-")

def pace_to_seconds(pace):
    try:
        m, s = pace.split(":")
        return int(m) * 60 + int(s)
    except:
        return 9999

def extract_number(value):
    try:
        return int("".join(filter(str.isdigit, str(value))))
    except:
        return 9999

# -------------------
# Sessão HTTP
# -------------------
session = requests.Session()
session.headers.update({"User-Agent": "Mozilla/5.0"})

# -------------------
# Parse Clax
# -------------------
def parse_clax_event(url):
    """Retorna lista de atletas (nome+percurso) e lista única de percursos"""
    try:
        r = session.get(url, timeout=10)
        r.encoding = "utf-8"
        root = ET.fromstring(r.text)

        athletes = []
        percursos_set = set()

        for athlete in root.findall(".//Engages/E"):
            nome = athlete.attrib.get("n")
            percurso = athlete.attrib.get("p")
            if nome and percurso:
                athletes.append({"nome": nome, "percurso": percurso})
                percursos_set.add(percurso)

        percursos = sorted(list(percursos_set))
        return athletes, percursos
    except Exception as e:
        print(f"Erro ao processar Clax: {e}")
        return [], []

# -------------------
# Consulta RunnersHub/OpenResults
# -------------------
def fetch_last_result(name):
    url = f"https://openresults.run/resultados/{slugify(name)}/"
    try:
        response = session.get(url, timeout=10)
        if response.status_code != 200:
            return None
        soup = BeautifulSoup(response.text, "html.parser")
        event = soup.find("h3", class_="h6")
        if not event:
            return None

        modality = None
        for a in event.find_all("a", href=True):
            href = a["href"]
            if "?modalidade=" in href and "categoria=" not in href:
                modality = normalize_text(a.text)
                break

        table = event.find_next("table")
        if not table:
            return None

        rows = table.find_all("tr")
        if len(rows) < 2:
            return None

        data = [normalize_text(td.text) for td in rows[1].find_all("td")]

        return {
            "nome": normalize_text(name),
            "modalidade": modality,
            "posicao": data[1],
            "pace": data[4],
            "tempo": data[5],
        }
    except:
        return None

# -------------------
# Worker Paralelo
# -------------------
def worker(name):
    time.sleep(random.uniform(1.1, 3.6))
    return fetch_last_result(name)

def process_names(names, max_workers=3):
    results = []
    total = len(names)
    failed = 0

    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        futures = {}

        for i, name in enumerate(names, 1):
            # pausa longa a cada 20 envios
            if i % 20 == 0:
                time.sleep(random.uniform(15, 34))

            future = executor.submit(worker, name)
            futures[future] = name

        for future in as_completed(futures):
            result = future.result()
            if result:
                print(f"{result['nome']} → {result.get('modalidade', '')}")
                results.append(result)
            else:
                failed += 1

    return results, total, failed

# -------------------
# Ordenação por cores
# -------------------
def sort_by_priority(results):
    """
    Ordena a lista de resultados:
    - verdes (pace < 4:00) → primeiro
    - azuis (posição <= 10) → depois
    - resto → por último
    """
    def get_priority(result):
        pos = extract_number(result.get("posicao", 9999))
        pace = pace_to_seconds(result.get("pace", "99:99"))

        if pace < 240:
            return 0  # verde
        elif pos <= 10:
            return 1  # azul
        else:
            return 2  # resto

    return sorted(results, key=get_priority)

# -------------------
# Colorir Excel
# -------------------
def color_results(filepath):
    wb = load_workbook(filepath)
    ws = wb.active

    colors = {
        "green": PatternFill("solid", fgColor="C6EFCE"),
        "blue": PatternFill("solid", fgColor="BDD7EE"),
    }

    for row in ws.iter_rows(min_row=2):
        pos = extract_number(row[2].value)
        pace = row[3].value
        fill = None

        if pos <= 10:
            fill = colors["blue"]
        elif pace_to_seconds(pace) < 240:
            fill = colors["green"]

        if fill:
            for cell in row:
                cell.fill = fill

    wb.save(filepath)

# -------------------
# Função Principal
# -------------------
def run_v2():
    clax_url = input("Digite a URL completa do .clax?t=...: ").strip()
    if not clax_url:
        print("URL inválida.")
        return []

    athletes, percursos = parse_clax_event(clax_url)
    if not athletes:
        print("Nenhum atleta encontrado.")
        return []

    print("Percursos disponíveis:")
    for i, p in enumerate(percursos, 1):
        print(f"{i} - {p}")

    choice = input("Escolha o número do percurso (Enter para todos): ").strip()
    if choice.isdigit() and 1 <= int(choice) <= len(percursos):
        chosen_percurso = percursos[int(choice) - 1]
    else:
        chosen_percurso = None

    if chosen_percurso:
        filtered_names = [a["nome"] for a in athletes if a["percurso"] == chosen_percurso]
    else:
        filtered_names = [a["nome"] for a in athletes]

    print(f"Atletas filtrados: {len(filtered_names)}")

    results, total, failed = process_names(filtered_names)

    # -------------------
    # Ordenar e exportar para Excel
    # -------------------
    results_sorted = sort_by_priority(results)  # aplica a ordem verde→azul→resto
    output_file = "resultados.xlsx"
    pd.DataFrame(results_sorted).to_excel(output_file, index=False)
    color_results(output_file)

    print("Finished ✅")
    print(f"Nomes processados: {total}")
    print(f"Falhas: {failed}")

    return results_sorted

# -------------------
# Entry Point
# -------------------
if __name__ == "__main__":
    run_v2()