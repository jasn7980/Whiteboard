"""Genera public/progreso.json para la tarjeta 'Progreso' del Whiteboard.

Fuentes: los CSV de TrainingPeaks en ~/workspace/your_files/
(75 entrenos detallados + series de carrera Ottawa / Petit Train 2026).
"""
import csv, json, re
from datetime import date, datetime
from collections import defaultdict
from pathlib import Path

YF = Path.home() / "workspace" / "your_files"
OUT = Path.home() / "workspace" / "whiteboard" / "public" / "progreso.json"

def hms(s):
    s = str(s).strip()
    if not s or s == "-":
        return None
    parts = [float(p) for p in s.split(":")]
    if len(parts) == 3:
        return parts[0] * 3600 + parts[1] * 60 + parts[2]
    if len(parts) == 2:
        return parts[0] * 60 + parts[1]
    return parts[0]

def fmt(sec):
    sec = int(round(sec))
    return f"{sec // 3600}:{(sec % 3600) // 60:02d}:{sec % 60:02d}"

def pace(sec, km):
    if not sec or not km:
        return None
    p = sec / km
    return f"{int(p // 60)}:{int(p % 60):02d}"

# --- Series semanales de km y TSS (carrera) -------------------------------
weekly = defaultdict(lambda: {"km": 0.0, "tss": 0})
for f in ["Carrera_Ottawa_2026.csv", "Carrera_PetitTrain_2026.csv"]:
    for r in csv.DictReader(open(YF / f)):
        d = date.fromisoformat(r["Fecha"])
        y, w, _ = d.isocalendar()
        key = f"{y}-W{w:02d}"
        try:
            weekly[key]["km"] += float(r["Distancia (km)"])
            weekly[key]["tss"] += float(r["TSS"])
        except (ValueError, KeyError):
            pass
weeks = [{"week": k, "km": round(v["km"], 1), "tss": int(round(v["tss"]))}
         for k, v in sorted(weekly.items())]

# --- Sesiones de intervalos (detalle completo de Instrucciones) ------------
pat = re.compile(r"interval|×|x\d|serie|repet|fartlek|tempo|umbral|velocidad",
                 re.IGNORECASE)
intervals = []
for f in ["Detalle_Ottawa_2026.csv", "Detalle_PetitTrain_2026.csv"]:
    block = "Ottawa" if "Ottawa" in f else "Petit Train"
    for r in csv.DictReader(open(YF / f)):
        if pat.search(r["Nombre"]) or pat.search(r["Instrucciones"]):
            intervals.append({
                "fecha": r["Fecha"], "bloque": block, "nombre": r["Nombre"],
                "km": r["Distancia (km)"], "duracion": r["Duración"],
                "ritmo": r["Ritmo prom (min/km)"], "tss": r["TSS"],
                "instrucciones": r["Instrucciones"],
            })
intervals.sort(key=lambda x: x["fecha"])

data = {
    "updated": datetime.now().strftime("%Y-%m-%d"),
    "goal": {"time": "3:45:00", "seconds": 13500,
             "pace": "5:20",
             "races": [
                 {"id": "ottawa", "name": "Ottawa 2027", "date": "2027-05-30",
                  "note": "Tamarack Ottawa Race Weekend",
                  "status": "✓ Inscrito"},
                 {"id": "chicago", "name": "Chicago 2027", "date": "2027-10-10",
                  "note": "Inscrito al sorteo 26 oct 2026 · resultados 8 dic 2026",
                  "status": "En sorteo"},
             ]},
    "marathons": [
        {"race": "Ottawa", "date": "2026-05-24", "time": "4:00:23",
         "seconds": hms("4:00:23"), "pace": pace(hms("4:00:23"), 42.195)},
        {"race": "Petit Train du Nord", "date": "2026-10-04", "time": "3:55:00",
         "seconds": hms("3:55:00"), "pace": pace(hms("3:55:00"), 42.195)},
    ],
    "half": {"race": "Montreal", "date": "2026-04-19", "time": "1:51:34",
             "seconds": hms("1:51:34")},
    "personalBests": [
        {"race": "Maratón Lala", "date": "2019", "time": "3:48",
         "note": "Marca personal de maratón"},
        {"race": "Maratón Lala", "date": "2020", "time": "3:52",
         "note": "Segundo maratón más rápido"},
        {"race": "Medio maratón ESPN", "date": "2019-05-05", "time": "1:36",
         "note": "Marca personal de medio maratón"},
    ],
    "weekly": weeks,
    "intervalSessions": intervals,
    "totals": {
        "workouts": 75,
        "ottawa_km": 931, "petit_km": 889,
        "ottawa_tss": 5310, "petit_tss": 5560,
    },
}
# Progreso hacia 3:45: de 4:00:23 a 3:55:00 de un total necesario de 15:23
start, now, goal = hms("4:00:23"), hms("3:55:00"), 13500
data["progress"] = {
    "from": "4:00:23", "now": "3:55:00", "goal": "3:45:00",
    "gained": fmt(start - now), "remaining": fmt(now - goal),
    "pct": round((start - now) / (start - goal) * 100),
}
OUT.write_text(json.dumps(data, ensure_ascii=False, indent=1))
print(f"Wrote {OUT}")
print(f"weeks={len(weeks)} intervals={len(intervals)}")
print(f"marathon pace Ottawa={data['marathons'][0]['pace']} "
      f"Petit={data['marathons'][1]['pace']} goal=5:20")
