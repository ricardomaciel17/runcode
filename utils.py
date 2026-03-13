import unicodedata

def normalize_text(value):
    """Remove acentos, espaços extras e normaliza texto."""
    if value is None:
        return ""
    text = str(value).replace("\xa0", " ")
    text = unicodedata.normalize("NFD", text)
    text = text.encode("ascii", "ignore").decode("utf-8")
    return " ".join(text.split()).strip()

def slugify(name):
    """Converte nome em slug para URLs."""
    return normalize_text(name).lower().replace(" ", "-")

def pace_to_seconds(pace):
    """Converte string MM:SS em segundos."""
    try:
        m, s = pace.split(":")
        return int(m) * 60 + int(s)
    except:
        return 9999

def extract_number(value):
    """Extrai números de uma string, retorna 9999 se falhar."""
    try:
        return int("".join(filter(str.isdigit, str(value))))
    except:
        return 9999