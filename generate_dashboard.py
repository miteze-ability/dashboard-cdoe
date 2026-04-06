from __future__ import annotations

import base64
import csv
import io
import json
import zipfile
from collections import Counter
from datetime import datetime, timedelta
from pathlib import Path
from typing import Dict, List
import xml.etree.ElementTree as ET
from PIL import Image


BASE_DIR = Path(__file__).resolve().parent
CSV_PATH = BASE_DIR / "RSP_Celulas_e_CDOEs_Vazias_000000000000.csv"
XLSX_PATH = BASE_DIR / "Consolidado Reparo por Planta.xlsx"
OUTPUT_PATH = BASE_DIR / "dashboard_cdoe.html"
INDEX_OUTPUT_PATH = BASE_DIR / "index.html"
LOGO_PATH = Path(r"C:\Users\Andrew Miteze\Pictures\Logo Ability.png")

XML_NS = {
    "a": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}


def normalize_text(value: str | None) -> str:
    if value is None:
        return ""
    return " ".join(str(value).replace("\xa0", " ").strip().split())


def normalize_status(value: str) -> str:
    cleaned = normalize_text(value).upper()
    replacements = {
        "EM SERVIÇO": "EM SERVICO",
        "EM SERVIÇo": "EM SERVICO",
        "EM SERVI?O": "EM SERVICO",
    }
    return replacements.get(cleaned, cleaned)


def sheet_rows(path: Path, sheet_name: str) -> List[Dict[str, str]]:
    with zipfile.ZipFile(path) as archive:
        shared_strings: List[str] = []
        if "xl/sharedStrings.xml" in archive.namelist():
            root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
            for item in root.findall("a:si", XML_NS):
                parts = [node.text or "" for node in item.iterfind(".//a:t", XML_NS)]
                shared_strings.append("".join(parts))

        workbook = ET.fromstring(archive.read("xl/workbook.xml"))
        rels = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
        rel_map = {rel.attrib["Id"]: rel.attrib["Target"] for rel in rels}

        sheet_target = None
        for sheet in workbook.find("a:sheets", XML_NS):
            if sheet.attrib["name"] == sheet_name:
                rel_id = sheet.attrib[
                    "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
                ]
                sheet_target = "xl/" + rel_map[rel_id]
                break

        if not sheet_target:
            raise ValueError(f"Aba '{sheet_name}' nao encontrada em {path.name}")

        root = ET.fromstring(archive.read(sheet_target))
        rows = root.findall(".//a:sheetData/a:row", XML_NS)

        header_map: Dict[str, str] = {}
        data: List[Dict[str, str]] = []

        for row_index, row in enumerate(rows):
            values: Dict[str, str] = {}
            for cell in row.findall("a:c", XML_NS):
                ref = cell.attrib.get("r", "")
                col = "".join(ch for ch in ref if ch.isalpha())
                cell_type = cell.attrib.get("t")
                value_node = cell.find("a:v", XML_NS)
                inline_node = cell.find("a:is", XML_NS)

                text = ""
                if cell_type == "s" and value_node is not None:
                    text = shared_strings[int(value_node.text)]
                elif cell_type == "inlineStr" and inline_node is not None:
                    text = "".join(node.text or "" for node in inline_node.iterfind(".//a:t", XML_NS))
                elif value_node is not None:
                    text = value_node.text or ""

                values[col] = normalize_text(text)

            if row_index == 0:
                header_map = {col: values[col] for col in values}
                continue

            record = {header_map[col]: values.get(col, "") for col in header_map}
            data.append(record)

        return data


def load_cdo_rows() -> List[Dict[str, str]]:
    rows: List[Dict[str, str]] = []
    with CSV_PATH.open("r", encoding="utf-8", errors="replace", newline="") as handle:
        reader = csv.DictReader(handle, delimiter=";")
        for row in reader:
            rows.append({key: normalize_text(value) for key, value in row.items()})
    return rows


def top_counter(counter: Counter, limit: int = 12) -> List[Dict[str, int | str]]:
    return [{"nome": name, "total": total} for name, total in counter.most_common(limit) if name]


def excel_serial_to_iso(value: str) -> str:
    raw = normalize_text(value)
    if not raw:
        return ""
    try:
        serial = float(raw)
    except ValueError:
        return ""
    base = datetime(1899, 12, 30)
    converted = base + timedelta(days=serial)
    return converted.date().isoformat()


def build_logo_data_uri() -> str:
    with Image.open(LOGO_PATH).convert("RGBA") as image:
        pixels = image.load()
        width, height = image.size
        for y in range(height):
            for x in range(width):
                r, g, b, a = pixels[x, y]
                if a and r > 245 and g > 245 and b > 245:
                    pixels[x, y] = (255, 255, 255, 0)

        bbox = image.getbbox()
        cropped = image.crop(bbox) if bbox else image
        buffer = io.BytesIO()
        cropped.save(buffer, format="PNG")
    encoded = base64.b64encode(buffer.getvalue()).decode("ascii")
    return f"data:image/png;base64,{encoded}"


def build_dashboard_data() -> Dict[str, object]:
    falha_rows = sheet_rows(XLSX_PATH, "Consolidado Reparo por Planta")
    falha_rows = [
        {
            "estacao": normalize_text(row.get("estacao")),
            "municipio": normalize_text(row.get("municipio")),
            "cdoe": normalize_text(row.get("cdo_name")),
            "celula": normalize_text(row.get("Celula")),
            "subcausa": normalize_text(row.get("Subcausa")),
            "causa_macro": normalize_text(row.get("Causa Macro")),
            "agrupador": normalize_text(row.get("Agrupador") or row.get("agrupador")),
            "data_abertura": excel_serial_to_iso(row.get("dat_abertura", "")),
            "data_fechamento": excel_serial_to_iso(row.get("dat_fechamento", "")),
        }
        for row in falha_rows
    ]

    cdo_rows = load_cdo_rows()
    cdo_rows = [
        {
            "estacao": normalize_text(row.get("cd_estacao_sigla_hc")),
            "municipio": normalize_text(row.get("ds_municipio_hp")),
            "cdoe": normalize_text(row.get("ds_nome_cdo_hp")),
            "celula": normalize_text(row.get("cd_celula_hp")),
            "status": normalize_status(row.get("ds_cdo_est_operacional", "")),
            "ptp": normalize_text(row.get("cd_cdo_ptp_name")),
        }
        for row in cdo_rows
        if normalize_text(row.get("ds_nome_cdo_hp"))
    ]

    estacoes = sorted({row["estacao"] for row in falha_rows if row["estacao"]} | {row["estacao"] for row in cdo_rows if row["estacao"]})
    municipios = sorted({row["municipio"] for row in falha_rows if row["municipio"]} | {row["municipio"] for row in cdo_rows if row["municipio"]})
    causas_macro = sorted({row["causa_macro"] for row in falha_rows if row["causa_macro"]})
    datas = sorted(
        {
            row["data_abertura"]
            for row in falha_rows
            if row["data_abertura"]
        }
    )

    cdoe_groups: Dict[str, Dict[str, object]] = {}
    for row in cdo_rows:
        item = cdoe_groups.setdefault(
            row["cdoe"],
            {
                "cdoe": row["cdoe"],
                "estacoes": set(),
                "municipios": set(),
                "celula_ids": set(),
                "status_set": set(),
                "ptp_count": 0,
                "celula_count": 0,
            },
        )
        if row["estacao"]:
            item["estacoes"].add(row["estacao"])
        if row["municipio"]:
            item["municipios"].add(row["municipio"])
        if row["celula"]:
            item["celula_ids"].add(row["celula"])
        if row["status"]:
            item["status_set"].add(row["status"])
        if row["ptp"]:
            item["ptp_count"] += 1
        if row["celula"]:
            item["celula_count"] += 1

    cdoes = []
    cdoe_cell_map = {}
    for item in cdoe_groups.values():
        statuses = sorted(item["status_set"])
        status_label = ", ".join(statuses) if statuses else "NAO ATIVA"
        cdoe_cell_map[item["cdoe"]] = sorted(item["celula_ids"])
        cdoes.append(
            {
                "cdoe": item["cdoe"],
                "estacoes": sorted(item["estacoes"]),
                "municipios": sorted(item["municipios"]),
                "celula_ids": sorted(item["celula_ids"]),
                "status_list": statuses,
                "status": status_label,
                "ptps": item["ptp_count"],
                "celulas": item["celula_count"],
                "is_ativa": "EM SERVICO" in statuses,
                "is_vazia": item["ptp_count"] == 0,
                "is_com_servico": item["ptp_count"] > 0,
                "is_nao_ativa": "EM SERVICO" not in statuses,
                "is_atencao": bool(statuses) and any(status not in {"EM SERVICO"} for status in statuses),
            }
        )
    cdoes.sort(key=lambda row: row["cdoe"])

    return {
        "estacoes": estacoes,
        "municipios": municipios,
        "causasMacro": causas_macro,
        "dataMin": datas[0] if datas else "",
        "dataMax": datas[-1] if datas else "",
        "falhasRows": falha_rows,
        "cdoesRows": cdoes,
        "cdoeCellMap": cdoe_cell_map,
        "notas": {
            "ativa": "CDOE ativa = possui status operacional 'Em Servico' na base CSV.",
            "vazia": "CDOE vazia = sem PTP vinculado no campo cd_cdo_ptp_name na base CSV.",
        },
    }


def dashboard_html(data: Dict[str, object]) -> str:
    json_data = json.dumps(data, ensure_ascii=False)
    logo_uri = build_logo_data_uri()
    return f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Dashboard CDOE e Células</title>
  <style>
    :root {{
      --bg: #f4f6fa;
      --surface: #ffffff;
      --surface-strong: #ffffff;
      --surface-soft: #eef2f8;
      --ink: #182230;
      --muted: #667085;
      --line: #d8dee8;
      --accent: #d92d20;
      --accent-soft: #fdecea;
      --warn: #175cd3;
      --warn-soft: #e9f2ff;
      --danger: #344054;
      --danger-soft: #f2f4f7;
      --shadow: 0 16px 40px rgba(15, 23, 42, 0.08);
    }}
    * {{ box-sizing: border-box; }}
    body {{
      margin: 0;
      font-family: "Segoe UI", Tahoma, Geneva, Verdana, sans-serif;
      background: linear-gradient(180deg, #f8fafc 0%, #f4f6fa 100%);
      color: var(--ink);
    }}
    .wrap {{
      max-width: 1440px;
      margin: 0 auto;
      padding: 28px 24px 56px;
    }}
    .hero {{
      display: grid;
      grid-template-columns: 320px 1fr;
      gap: 28px;
      align-items: center;
      background: linear-gradient(180deg, #ffffff 0%, #f8fafc 100%);
      color: var(--ink);
      border-radius: 24px;
      padding: 28px 30px;
      box-shadow: var(--shadow);
      border: 1px solid var(--line);
      margin-bottom: 18px;
    }}
    .brand-card {{
      display: flex;
      align-items: center;
      justify-content: center;
      min-height: 132px;
      background: linear-gradient(180deg, #ffffff, #f8fafc);
      border: 1px solid var(--line);
      border-radius: 20px;
      padding: 22px;
    }}
    .brand-card img {{
      width: 100%;
      max-width: 220px;
      height: auto;
      display: block;
    }}
    .hero h1 {{
      margin: 0 0 8px;
      font-size: clamp(30px, 4vw, 42px);
      line-height: 1.08;
      letter-spacing: -0.03em;
    }}
    .hero p {{
      margin: 0;
      color: var(--muted);
      max-width: 820px;
      line-height: 1.6;
      font-size: 15px;
    }}
    .filters {{
      display: grid;
      grid-template-columns: repeat(14, minmax(0, 1fr));
      gap: 16px;
      margin: 0 0 18px;
      padding: 20px;
      background: var(--surface);
      border: 1px solid var(--line);
      border-radius: 22px;
      box-shadow: var(--shadow);
    }}
    .card, .metric {{
      background: var(--surface-strong);
      border: 1px solid var(--line);
      border-radius: 20px;
      box-shadow: var(--shadow);
    }}
    .filter-card {{
      padding: 16px;
      min-width: 0;
      background: var(--surface-soft);
      box-shadow: none;
    }}
    .filter-card:nth-child(1) {{ grid-column: span 2; }}
    .filter-card:nth-child(2) {{ grid-column: span 2; }}
    .filter-card:nth-child(3) {{ grid-column: span 4; }}
    .filter-card:nth-child(4) {{ grid-column: span 2; }}
    .filter-card:nth-child(5) {{ grid-column: span 2; }}
    .filter-card:nth-child(6) {{ grid-column: span 2; }}
    .causa-card {{
      min-height: 0;
    }}
    label {{
      display: block;
      font-size: 12px;
      text-transform: uppercase;
      letter-spacing: 0.08em;
      color: var(--muted);
      margin-bottom: 8px;
      font-weight: 700;
    }}
    select, input {{
      width: 100%;
      border: 1px solid #c7d0dd;
      border-radius: 14px;
      padding: 12px 14px;
      font-size: 15px;
      background: #ffffff;
      color: var(--ink);
      outline: none;
    }}
    select:focus, input:focus {{
      border-color: #175cd3;
      box-shadow: 0 0 0 3px rgba(23, 92, 211, 0.12);
    }}
    .checkbox-panel {{
      position: relative;
    }}
    .multi-select-trigger {{
      width: 100%;
      min-height: 48px;
      border: 1px solid #c7d0dd;
      border-radius: 14px;
      padding: 12px 42px 12px 14px;
      font-size: 15px;
      background: #ffffff;
      color: var(--ink);
      text-align: left;
      cursor: pointer;
      position: relative;
    }}
    .multi-select-trigger::after {{
      content: "▾";
      position: absolute;
      right: 14px;
      top: 50%;
      transform: translateY(-50%);
      color: var(--muted);
      font-size: 14px;
    }}
    .multi-select-trigger.is-open::after {{
      content: "▴";
    }}
    .multi-select-menu {{
      display: none;
      position: absolute;
      top: calc(100% + 8px);
      left: 0;
      right: 0;
      z-index: 20;
      background: #ffffff;
      border: 1px solid var(--line);
      border-radius: 16px;
      box-shadow: 0 20px 40px rgba(15, 23, 42, 0.14);
      padding: 14px;
    }}
    .multi-select-menu.is-open {{
      display: grid;
      gap: 10px;
    }}
    .checkbox-actions {{
      display: flex;
      gap: 8px;
      flex-wrap: wrap;
    }}
    .chip-btn {{
      border: 1px solid var(--line);
      background: #ffffff;
      color: var(--ink);
      border-radius: 999px;
      padding: 8px 12px;
      font-size: 12px;
      cursor: pointer;
    }}
    .chip-btn:hover {{
      border-color: #175cd3;
      color: #175cd3;
    }}
    .checkbox-list {{
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
      gap: 8px 12px;
      max-height: 164px;
      overflow: auto;
      padding-right: 4px;
    }}
    .check-item {{
      display: flex;
      align-items: center;
      gap: 8px;
      font-size: 13px;
      color: var(--ink);
    }}
    .check-item input {{
      width: 16px;
      height: 16px;
      accent-color: #ef4444;
    }}
    .metrics {{
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
      gap: 16px;
      margin: 18px 0 26px;
    }}
    .active-selection {{
      display: none;
      align-items: center;
      justify-content: space-between;
      gap: 12px;
      padding: 14px 18px;
      margin: -6px 0 18px;
    }}
    .active-selection.is-visible {{
      display: flex;
    }}
    .active-selection-meta {{
      display: flex;
      flex-wrap: wrap;
      gap: 10px;
      align-items: center;
    }}
    .selection-chip {{
      display: inline-flex;
      align-items: center;
      padding: 8px 12px;
      border-radius: 999px;
      background: var(--warn-soft);
      color: var(--warn);
      font-size: 12px;
      font-weight: 700;
      border: 1px solid #cfe0ff;
    }}
    .metric {{
      padding: 18px;
      position: relative;
      overflow: hidden;
    }}
    .metric::after {{
      content: "";
      position: absolute;
      inset: auto -20px -28px auto;
      width: 84px;
      height: 84px;
      border-radius: 999px;
      background: var(--warn-soft);
      opacity: 1;
    }}
    .metric h2 {{
      margin: 0;
      font-size: 13px;
      color: var(--muted);
      text-transform: uppercase;
      letter-spacing: 0.08em;
    }}
    .metric strong {{
      display: block;
      margin-top: 10px;
      font-size: clamp(26px, 4vw, 38px);
      line-height: 1;
    }}
    .metric span {{
      display: block;
      margin-top: 8px;
      color: var(--muted);
      font-size: 13px;
      max-width: 22ch;
    }}
    .grid {{
      display: grid;
      grid-template-columns: repeat(12, 1fr);
      gap: 18px;
    }}
    .panel {{
      grid-column: span 6;
      padding: 20px;
    }}
    .panel.wide {{
      grid-column: span 12;
    }}
    .panel h3 {{
      margin: 0 0 4px;
      font-size: 20px;
    }}
    .panel p {{
      margin: 0 0 18px;
      color: var(--muted);
      line-height: 1.5;
    }}
    .bars {{
      display: grid;
      gap: 12px;
    }}
    .bar-row {{
      display: grid;
      gap: 6px;
    }}
    .bar-filter-btn {{
      appearance: none;
      background: transparent;
      border: 0;
      padding: 0;
      text-align: left;
      cursor: pointer;
      display: grid;
      gap: 6px;
    }}
    .bar-filter-btn:hover .track {{
      box-shadow: inset 0 0 0 1px #b7c7e5;
    }}
    .bar-filter-btn.is-active .track {{
      box-shadow: inset 0 0 0 2px #175cd3;
    }}
    .bar-filter-btn.is-active .bar-label span:first-child {{
      color: #175cd3;
      font-weight: 700;
    }}
    .bar-label {{
      display: flex;
      justify-content: space-between;
      gap: 10px;
      font-size: 14px;
      font-weight: 600;
    }}
    .track {{
      height: 12px;
      background: #e6ebf2;
      border-radius: 999px;
      overflow: hidden;
    }}
    .fill {{
      height: 100%;
      border-radius: 999px;
      background: linear-gradient(90deg, #175cd3, #528bff);
    }}
    .tables {{
      display: grid;
      grid-template-columns: repeat(3, minmax(0, 1fr));
      gap: 16px;
    }}
    .section-toolbar {{
      display: flex;
      flex-wrap: wrap;
      justify-content: space-between;
      align-items: center;
      gap: 14px;
      margin-bottom: 18px;
    }}
    .mini-filter {{
      min-width: 220px;
      max-width: 320px;
      padding: 14px;
      border-radius: 18px;
      background: rgba(255, 255, 255, 0.03);
      border: 1px solid rgba(255, 255, 255, 0.08);
    }}
    .section-actions {{
      display: flex;
      align-items: center;
      gap: 12px;
      flex-wrap: wrap;
    }}
    .section-grid {{
      display: grid;
      grid-template-columns: minmax(320px, 0.9fr) minmax(0, 2.1fr);
      gap: 18px;
      margin-bottom: 20px;
    }}
    table {{
      width: 100%;
      border-collapse: collapse;
      font-size: 14px;
    }}
    th, td {{
      text-align: left;
      padding: 10px 8px;
      border-bottom: 1px solid #edf1f7;
      vertical-align: top;
    }}
    th {{
      font-size: 12px;
      text-transform: uppercase;
      letter-spacing: 0.08em;
      color: var(--muted);
    }}
    .pill {{
      display: inline-flex;
      align-items: center;
      gap: 8px;
      border-radius: 999px;
      padding: 8px 12px;
      font-size: 13px;
      font-weight: 700;
      margin: 0 10px 10px 0;
    }}
    .pill.ok {{ background: var(--warn-soft); color: var(--warn); }}
    .pill.warn {{ background: var(--accent-soft); color: var(--accent); }}
    .pill.danger {{ background: var(--danger-soft); color: var(--danger); }}
    .notes {{
      margin-top: 22px;
      background: #fbfcfe;
      border: 1px dashed #cdd7e4;
      border-radius: 18px;
      padding: 16px 18px;
      color: var(--muted);
      line-height: 1.55;
    }}
    .empty {{
      padding: 18px;
      border-radius: 16px;
      background: #f8fafc;
      color: var(--muted);
      text-align: center;
    }}
    @media (max-width: 1080px) {{
      .hero {{ grid-template-columns: 1fr; }}
      .panel {{ grid-column: span 12; }}
      .tables {{ grid-template-columns: 1fr; }}
      .section-grid {{ grid-template-columns: 1fr; }}
      .filters {{ grid-template-columns: repeat(2, minmax(0, 1fr)); }}
      .filter-card:nth-child(1),
      .filter-card:nth-child(2),
      .filter-card:nth-child(3),
      .filter-card:nth-child(4),
      .filter-card:nth-child(5),
      .filter-card:nth-child(6) {{ grid-column: span 1; }}
    }}
    @media (max-width: 720px) {{
      .filters {{ grid-template-columns: 1fr; }}
    }}
  </style>
</head>
<body>
  <div class="wrap">
    <section class="hero">
      <div class="brand-card">
        <img src="{logo_uri}" alt="Ability">
      </div>
      <div>
        <h1>Painel Gerencial de CDOEs e Ofensores Operacionais</h1>
        <p>
          Visão executiva para acompanhamento de falhas, ofensores por célula, CDOE e município,
          além do monitoramento de CDOEs vazias, ativas, não ativas e com serviço vinculado.
        </p>
      </div>
    </section>

    <section class="filters">
      <div class="card filter-card">
        <label for="estacao">Estação</label>
        <select id="estacao"></select>
      </div>
      <div class="card filter-card">
        <label for="municipio">Município</label>
        <select id="municipio"></select>
      </div>
      <div class="card filter-card causa-card">
        <label>Causa macro</label>
        <div id="causaMacroPanel" class="checkbox-panel"></div>
      </div>
      <div class="card filter-card">
        <label for="statusCdoe">Situação da CDOE</label>
        <select id="statusCdoe"></select>
      </div>
      <div class="card filter-card">
        <label for="dataInicio">Data inicial</label>
        <input id="dataInicio" type="date">
      </div>
      <div class="card filter-card">
        <label for="dataFim">Data final</label>
        <input id="dataFim" type="date">
      </div>
    </section>

    <section class="metrics">
      <article class="metric">
        <h2>Total de falhas</h2>
        <strong id="totalFalhas">0</strong>
        <span>Ocorrências no consolidado de reparo por planta.</span>
      </article>
      <article class="metric">
        <h2>CDOEs com falha</h2>
        <strong id="cdoesComFalha">0</strong>
        <span>Quantidade distinta de CDOEs impactadas.</span>
      </article>
      <article class="metric">
        <h2>Células com falha</h2>
        <strong id="celulasComFalha">0</strong>
        <span>Quantidade distinta de células ofensores.</span>
      </article>
      <article class="metric">
        <h2>CDOEs ativas</h2>
        <strong id="ativasQtd">0</strong>
        <span>Status operacional ativo na base CSV.</span>
      </article>
      <article class="metric">
        <h2>CDOEs vazias</h2>
        <strong id="vaziasQtd">0</strong>
        <span>Sem PTP vinculado no cadastro da base CSV.</span>
      </article>
      <article class="metric">
        <h2>CDOEs com serviço</h2>
        <strong id="atencaoQtd">0</strong>
        <span>CDOEs que não estão vazias e possuem PTP vinculado.</span>
      </article>
    </section>

    <section class="card active-selection" id="activeSelectionPanel">
      <div class="active-selection-meta">
        <strong>Filtro interativo ativo:</strong>
        <div id="activeSelectionChips"></div>
      </div>
      <button type="button" class="chip-btn" id="clearInteractiveSelection">Limpar seleção</button>
    </section>

    <section class="grid">
      <article class="card panel">
        <h3>Quem mais apresentou falhas</h3>
        <p>Ranking dos principais ofensores no recorte selecionado.</p>
        <div id="topCdoesBars" class="bars"></div>
      </article>

      <article class="card panel">
        <h3>Falhas por subcausa</h3>
        <p>Mostra quais tipos de falha mais puxam o volume na visão filtrada.</p>
        <div id="topSubcausasBars" class="bars"></div>
      </article>

      <article class="card panel">
        <h3>Células ofensoras</h3>
        <p>Ranking das células com maior volume de falhas no recorte selecionado.</p>
        <div id="topCelulasBars" class="bars"></div>
      </article>

      <article class="card panel">
        <h3>Municípios ofensores</h3>
        <p>Ranking dos municípios com maior volume de falhas no recorte selecionado.</p>
        <div id="topMunicipiosBars" class="bars"></div>
      </article>

      <article class="card panel wide">
        <h3>Detalhamento dos ofensores</h3>
        <p>Separado entre CDOEs ofensores, células ofensores e municípios ofensores.</p>
        <div class="tables">
          <div>
            <table>
              <thead><tr><th>CDOE</th><th>Total</th></tr></thead>
              <tbody id="topCdoesTable"></tbody>
            </table>
          </div>
          <div>
            <table>
              <thead><tr><th>Célula</th><th>Total</th></tr></thead>
              <tbody id="topCelulasTable"></tbody>
            </table>
          </div>
          <div>
            <table>
              <thead><tr><th>Município</th><th>Total</th></tr></thead>
              <tbody id="topMunicipiosTable"></tbody>
            </table>
          </div>
        </div>
      </article>

      <article class="card panel wide">
        <h3>Situação das CDOEs</h3>
        <p>Consulta rápida para saber quais CDOEs estão ativas, vazias, não ativas ou já têm serviço vinculado.</p>
        <div class="section-toolbar">
          <div>
            <span class="pill ok" id="pillAtivas">Ativas: 0</span>
            <span class="pill warn" id="pillVazias">Vazias: 0</span>
            <span class="pill danger" id="pillComServico">Com serviço: 0</span>
          </div>
          <div class="section-actions">
            <div class="mini-filter">
              <label for="statusCdoeSecao">Filtro desta seção</label>
              <select id="statusCdoeSecao"></select>
            </div>
            <button type="button" class="chip-btn" id="exportCdoesExcel">Exportar Excel</button>
          </div>
        </div>
        <div class="section-grid">
          <div class="card filter-card">
            <h3>Top células com mais CDOEs vazias</h3>
            <p>Volume de CDOEs sem PTP por célula no recorte selecionado.</p>
            <div id="topCelulasVaziasBars" class="bars"></div>
          </div>
          <div>
          </div>
        </div>
        <div class="tables">
          <div>
            <table>
              <thead><tr><th>CDOE ativa</th><th>Status</th><th>PTPs</th></tr></thead>
              <tbody id="ativasTable"></tbody>
            </table>
          </div>
          <div>
            <table>
              <thead><tr><th>CDOE vazia</th><th>Status</th><th>PTPs</th></tr></thead>
              <tbody id="vaziasTable"></tbody>
            </table>
          </div>
          <div>
            <table>
              <thead><tr><th>CDOE com serviço</th><th>Status</th><th>PTPs</th></tr></thead>
              <tbody id="comServicoTable"></tbody>
            </table>
          </div>
        </div>
        <div class="notes">
          <strong>Regras usadas no dashboard:</strong><br>
          <span id="notaAtiva"></span><br>
          <span id="notaVazia"></span><br>
          <span id="notaComServico"></span>
        </div>
      </article>
    </section>
  </div>

  <script>
    const DATA = {json_data};

    const estacaoSelect = document.getElementById("estacao");
    const municipioSelect = document.getElementById("municipio");
    const causaMacroPanel = document.getElementById("causaMacroPanel");
    const statusCdoeSelect = document.getElementById("statusCdoe");
    const statusCdoeSecaoSelect = document.getElementById("statusCdoeSecao");
    const dataInicioInput = document.getElementById("dataInicio");
    const dataFimInput = document.getElementById("dataFim");
    const activeSelectionPanel = document.getElementById("activeSelectionPanel");
    const activeSelectionChips = document.getElementById("activeSelectionChips");
    const exportCdoesExcelButton = document.getElementById("exportCdoesExcel");
    const interactiveSelection = {{
      cdoe: null,
      celula: null,
      municipio: null,
      subcausa: null,
    }};
    let latestCdoesSecao = {{
      ativas: [],
      vazias: [],
      comServico: [],
    }};

    function formatNumber(value) {{
      return new Intl.NumberFormat("pt-BR").format(value || 0);
    }}

    function fillSelect(select, items, allLabel) {{
      select.innerHTML = "";
      const defaults = [{{ value: allLabel.value, label: allLabel.label }}];
      for (const item of defaults.concat(items.map((value) => ({{ value, label: value }})))) {{
        const option = document.createElement("option");
        option.value = item.value;
        option.textContent = item.label;
        select.appendChild(option);
      }}
    }}

    function renderCauseMacroOptions() {{
      causaMacroPanel.innerHTML = `
        <button type="button" class="multi-select-trigger" id="causaMacroTrigger">Todas as causas macro</button>
        <div class="multi-select-menu" id="causaMacroMenu">
          <div class="checkbox-actions">
            <button type="button" class="chip-btn" id="selectAllCauses">Selecionar todas</button>
            <button type="button" class="chip-btn" id="clearCauses">Limpar</button>
          </div>
          <div class="checkbox-list">
            ${{DATA.causasMacro.map((item) => `
              <label class="check-item">
                <input type="checkbox" class="cause-checkbox" value="${{item}}" checked>
                <span>${{item}}</span>
              </label>
            `).join("")}}
          </div>
        </div>
      `;

      const trigger = document.getElementById("causaMacroTrigger");
      const menu = document.getElementById("causaMacroMenu");
      trigger.addEventListener("click", (event) => {{
        event.stopPropagation();
        trigger.classList.toggle("is-open");
        menu.classList.toggle("is-open");
      }});

      document.getElementById("selectAllCauses").addEventListener("click", () => {{
        document.querySelectorAll(".cause-checkbox").forEach((el) => el.checked = true);
        updateCauseMacroTrigger();
        render();
      }});
      document.getElementById("clearCauses").addEventListener("click", () => {{
        document.querySelectorAll(".cause-checkbox").forEach((el) => el.checked = false);
        updateCauseMacroTrigger();
        render();
      }});
      document.querySelectorAll(".cause-checkbox").forEach((el) => el.addEventListener("change", () => {{
        updateCauseMacroTrigger();
        render();
      }}));

      document.addEventListener("click", (event) => {{
        if (!causaMacroPanel.contains(event.target)) {{
          trigger.classList.remove("is-open");
          menu.classList.remove("is-open");
        }}
      }});

      updateCauseMacroTrigger();
    }}

    function selectedCauses() {{
      return Array.from(document.querySelectorAll(".cause-checkbox:checked")).map((el) => el.value);
    }}

    function updateCauseMacroTrigger() {{
      const selected = selectedCauses();
      const trigger = document.getElementById("causaMacroTrigger");
      if (!trigger) return;
      if (selected.length === DATA.causasMacro.length) {{
        trigger.textContent = "Todas as causas macro";
        return;
      }}
      if (!selected.length) {{
        trigger.textContent = "Nenhuma causa selecionada";
        return;
      }}
      if (selected.length === 1) {{
        trigger.textContent = selected[0];
        return;
      }}
      trigger.textContent = `${{selected.length}} causas selecionadas`;
    }}

    function hasInteractiveSelection() {{
      return Object.values(interactiveSelection).some(Boolean);
    }}

    function interactiveSelectionEntries() {{
      const labels = {{
        cdoe: "CDOE",
        celula: "Célula",
        municipio: "Município",
        subcausa: "Subcausa",
      }};
      return Object.entries(interactiveSelection)
        .filter(([, value]) => Boolean(value))
        .map(([key, value]) => ({{ label: labels[key], value }}));
    }}

    function updateInteractiveSelectionPanel() {{
      const entries = interactiveSelectionEntries();
      if (!entries.length) {{
        activeSelectionPanel.classList.remove("is-visible");
        activeSelectionChips.innerHTML = "";
        return;
      }}
      activeSelectionPanel.classList.add("is-visible");
      activeSelectionChips.innerHTML = entries
        .map((item) => `<span class="selection-chip">${{item.label}}: ${{item.value}}</span>`)
        .join("");
    }}

    function toggleInteractiveSelection(dimension, value) {{
      interactiveSelection[dimension] = interactiveSelection[dimension] === value ? null : value;
      render();
    }}

    function escapeHtml(value) {{
      return String(value ?? "")
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;");
    }}

    function exportSectionToExcel() {{
      const sections = [
        ["CDOE Ativa", latestCdoesSecao.ativas],
        ["CDOE Vazia", latestCdoesSecao.vazias],
        ["CDOE Com Serviço", latestCdoesSecao.comServico],
      ];

      const tables = sections.map(([title, rows]) => `
        <h2>${{escapeHtml(title)}}</h2>
        <table border="1">
          <thead>
            <tr>
              <th>Categoria</th>
              <th>CDOE</th>
              <th>Status</th>
              <th>PTPs</th>
              <th>Células</th>
              <th>Estações</th>
              <th>Municípios</th>
            </tr>
          </thead>
          <tbody>
            ${{(rows.length ? rows : [{{ cdoe: "", status: "Sem dados para este filtro", ptps: "", celula_ids: [], estacoes: [], municipios: [] }}]).map((row) => `
              <tr>
                <td>${{escapeHtml(title)}}</td>
                <td>${{escapeHtml(row.cdoe)}}</td>
                <td>${{escapeHtml(row.status)}}</td>
                <td>${{escapeHtml(row.ptps)}}</td>
                <td>${{escapeHtml((row.celula_ids || []).join(", "))}}</td>
                <td>${{escapeHtml((row.estacoes || []).join(", "))}}</td>
                <td>${{escapeHtml((row.municipios || []).join(", "))}}</td>
              </tr>
            `).join("")}}
          </tbody>
        </table>
        <br>
      `).join("");

      const html = `
        <html xmlns:o="urn:schemas-microsoft-com:office:office"
              xmlns:x="urn:schemas-microsoft-com:office:excel"
              xmlns="http://www.w3.org/TR/REC-html40">
        <head>
          <meta charset="utf-8">
          <title>Exportação CDOEs</title>
        </head>
        <body>
          <h1>Detalhamento de CDOEs</h1>
          <p>Exportado do dashboard com os filtros atuais da seção.</p>
          ${{tables}}
        </body>
        </html>
      `;

      const blob = new Blob([html], {{ type: "application/vnd.ms-excel;charset=utf-8;" }});
      const link = document.createElement("a");
      const date = new Date().toISOString().slice(0, 10);
      link.href = URL.createObjectURL(blob);
      link.download = `cdoes_detalhadas_${{date}}.xls`;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      URL.revokeObjectURL(link.href);
    }}

    function inDateRange(value, start, end) {{
      if (!value) return true;
      if (start && value < start) return false;
      if (end && value > end) return false;
      return true;
    }}

    function statusMatches(item, statusFilter) {{
      if (statusFilter === "TODAS") return true;
      if (statusFilter === "ATIVAS") return item.is_ativa;
      if (statusFilter === "VAZIAS") return item.is_vazia;
      if (statusFilter === "COM_SERVICO") return item.is_com_servico;
      if (statusFilter === "NAO_ATIVAS") return item.is_nao_ativa;
      if (statusFilter === "ATENCAO") return item.is_atencao;
      return true;
    }}

    function filterFalhas() {{
      const estacao = estacaoSelect.value;
      const municipio = municipioSelect.value;
      const causasSelecionadas = selectedCauses();
      const dataInicio = dataInicioInput.value;
      const dataFim = dataFimInput.value;

      return DATA.falhasRows.filter((row) => {{
        if (estacao !== "TODAS" && row.estacao !== estacao) return false;
        if (municipio !== "TODOS" && row.municipio !== municipio) return false;
        if (!causasSelecionadas.length) return false;
        if (!causasSelecionadas.includes(row.causa_macro)) return false;
        if (interactiveSelection.cdoe && row.cdoe !== interactiveSelection.cdoe) return false;
        if (interactiveSelection.celula && row.celula !== interactiveSelection.celula) return false;
        if (interactiveSelection.municipio && row.municipio !== interactiveSelection.municipio) return false;
        if (interactiveSelection.subcausa && row.subcausa !== interactiveSelection.subcausa) return false;
        if (!inDateRange(row.data_abertura, dataInicio, dataFim)) return false;
        return true;
      }});
    }}

    function filterCdoes(statusOverride = null, allowedCdoes = null) {{
      const estacao = estacaoSelect.value;
      const municipio = municipioSelect.value;
      const statusFilter = statusOverride || statusCdoeSelect.value;

      return DATA.cdoesRows.filter((row) => {{
        if (estacao !== "TODAS" && !row.estacoes.includes(estacao)) return false;
        if (municipio !== "TODOS" && !row.municipios.includes(municipio)) return false;
        if (interactiveSelection.cdoe && row.cdoe !== interactiveSelection.cdoe) return false;
        if (interactiveSelection.celula && !row.celula_ids.includes(interactiveSelection.celula)) return false;
        if (interactiveSelection.municipio && !row.municipios.includes(interactiveSelection.municipio)) return false;
        if (allowedCdoes && !allowedCdoes.has(row.cdoe)) return false;
        if (!statusMatches(row, statusFilter)) return false;
        return true;
      }});
    }}

    function summarizeFalhas(rows) {{
      const cdoeCounter = new Map();
      const celulaCounter = new Map();
      const municipioCounter = new Map();
      const subcausaCounter = new Map();
      const causaCounter = new Map();

      for (const row of rows) {{
        if (row.cdoe) cdoeCounter.set(row.cdoe, (cdoeCounter.get(row.cdoe) || 0) + 1);
        if (row.celula) {{
          celulaCounter.set(row.celula, (celulaCounter.get(row.celula) || 0) + 1);
        }} else if (row.cdoe && DATA.cdoeCellMap[row.cdoe]?.length) {{
          DATA.cdoeCellMap[row.cdoe].forEach((cell) => {{
            celulaCounter.set(cell, (celulaCounter.get(cell) || 0) + 1);
          }});
        }}
        if (row.municipio) municipioCounter.set(row.municipio, (municipioCounter.get(row.municipio) || 0) + 1);
        if (row.subcausa) subcausaCounter.set(row.subcausa, (subcausaCounter.get(row.subcausa) || 0) + 1);
        if (row.causa_macro) causaCounter.set(row.causa_macro, (causaCounter.get(row.causa_macro) || 0) + 1);
      }}

      const rank = (counter, limit = 12) =>
        Array.from(counter.entries())
          .sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0]))
          .slice(0, limit)
          .map(([nome, total]) => ({{ nome, total }}));

      return {{
        totalFalhas: rows.length,
        cdoesComFalha: cdoeCounter.size,
        celulasComFalha: celulaCounter.size,
        topCdoes: rank(cdoeCounter),
        topCelulas: rank(celulaCounter),
        topMunicipios: rank(municipioCounter),
        topSubcausas: rank(subcausaCounter),
        topCausasMacro: rank(causaCounter, 8),
      }};
    }}

    function summarizeCelulasVazias(rows) {{
      const cellCounter = new Map();
      for (const row of rows) {{
        if (!row.is_vazia) continue;
        const cells = row.celula_ids && row.celula_ids.length ? row.celula_ids : ["SEM CÉLULA"];
        for (const cell of cells) {{
          cellCounter.set(cell, (cellCounter.get(cell) || 0) + 1);
        }}
      }}
      return Array.from(cellCounter.entries())
        .sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0]))
        .slice(0, 10)
        .map(([nome, total]) => ({{ nome, total }}));
    }}

    function summarizeCdoes(rows) {{
      const ativas = rows.filter((row) => row.is_ativa);
      const vazias = rows.filter((row) => row.is_vazia);
      const comServico = rows.filter((row) => row.is_com_servico);
      const atencao = rows.filter((row) => row.is_atencao);
      return {{
        ativasQtd: ativas.length,
        vaziasQtd: vazias.length,
        comServicoQtd: comServico.length,
        atencaoQtd: atencao.length,
        ativas: ativas.slice(0, 300),
        vazias: vazias.slice(0, 300),
        comServico: comServico.slice(0, 300),
        atencao: atencao.slice(0, 300),
      }};
    }}

    function tableRows(items, cols) {{
      if (!items || !items.length) {{
        return `<tr><td colspan="${{cols}}" class="empty">Nenhum dado para esse filtro.</td></tr>`;
      }}
      return items.map((item) => `
        <tr>
          <td>${{item.cdoe || item.nome || "-"}}</td>
          <td>${{item.status || formatNumber(item.total)}}</td>
          <td>${{item.ptps ?? ""}}</td>
        </tr>
      `).join("");
    }}

    function rankedTable(items, nameKey = "nome") {{
      if (!items || !items.length) {{
        return `<tr><td colspan="2" class="empty">Nenhum dado para esse filtro.</td></tr>`;
      }}
      return items.map((item) => `
        <tr>
          <td>${{item[nameKey] || "-"}}</td>
          <td>${{formatNumber(item.total)}}</td>
        </tr>
      `).join("");
    }}

    function barPalette(containerId) {{
      const palettes = {{
        topCdoesBars: ["#175cd3", "#5b8def"],
        topSubcausasBars: ["#2e90a6", "#6cb6c9"],
        topCelulasBars: ["#4c6fff", "#8aa2ff"],
        topMunicipiosBars: ["#6172f3", "#98a2ff"],
        topCelulasVaziasBars: ["#d92d20", "#f97066"],
      }};
      return palettes[containerId] || ["#175cd3", "#5b8def"];
    }}

    function barDimension(containerId) {{
      const dimensions = {{
        topCdoesBars: "cdoe",
        topSubcausasBars: "subcausa",
        topCelulasBars: "celula",
        topMunicipiosBars: "municipio",
        topCelulasVaziasBars: "celula",
      }};
      return dimensions[containerId] || null;
    }}

    function bars(containerId, items) {{
      const container = document.getElementById(containerId);
      if (!items || !items.length) {{
        container.innerHTML = `<div class="empty">Nenhum dado para esse filtro.</div>`;
        return;
      }}
      const max = Math.max(...items.map((item) => item.total), 1);
      const [startColor, endColor] = barPalette(containerId);
      const dimension = barDimension(containerId);
      const activeValue = dimension ? interactiveSelection[dimension] : null;
      container.innerHTML = items.map((item) => `
        <div class="bar-row">
          <button type="button" class="bar-filter-btn ${{activeValue === item.nome ? "is-active" : ""}}" data-dimension="${{dimension || ""}}" data-value="${{item.nome}}">
            <div class="bar-label">
              <span>${{item.nome}}</span>
              <span>${{formatNumber(item.total)}}</span>
            </div>
            <div class="track"><div class="fill" style="width: ${{Math.max(8, (item.total / max) * 100)}}%; background: linear-gradient(90deg, ${{startColor}}, ${{endColor}});"></div></div>
          </button>
        </div>
      `).join("");

      container.querySelectorAll(".bar-filter-btn").forEach((button) => {{
        button.addEventListener("click", () => {{
          const selectedDimension = button.dataset.dimension;
          const selectedValue = button.dataset.value;
          if (!selectedDimension) return;
          toggleInteractiveSelection(selectedDimension, selectedValue);
        }});
      }});
    }}

    function render() {{
      const falhasRowsFiltradas = filterFalhas();
      const allowedCdoes = hasInteractiveSelection()
        ? new Set(falhasRowsFiltradas.map((row) => row.cdoe).filter(Boolean))
        : null;
      const falhas = summarizeFalhas(falhasRowsFiltradas);
      const cdoes = summarizeCdoes(filterCdoes(null, allowedCdoes));
      const cdoesSecao = summarizeCdoes(filterCdoes(statusCdoeSecaoSelect.value, allowedCdoes));
      const topCelulasVazias = summarizeCelulasVazias(filterCdoes(statusCdoeSecaoSelect.value, allowedCdoes));
      latestCdoesSecao = cdoesSecao;

      document.getElementById("totalFalhas").textContent = formatNumber(falhas.totalFalhas);
      document.getElementById("cdoesComFalha").textContent = formatNumber(falhas.cdoesComFalha);
      document.getElementById("celulasComFalha").textContent = formatNumber(falhas.celulasComFalha);
      document.getElementById("ativasQtd").textContent = formatNumber(cdoes.ativasQtd);
      document.getElementById("vaziasQtd").textContent = formatNumber(cdoes.vaziasQtd);
      document.getElementById("atencaoQtd").textContent = formatNumber(cdoes.comServicoQtd);

      document.getElementById("pillAtivas").textContent = `Ativas: ${{formatNumber(cdoesSecao.ativasQtd)}}`;
      document.getElementById("pillVazias").textContent = `Vazias: ${{formatNumber(cdoesSecao.vaziasQtd)}}`;
      document.getElementById("pillComServico").textContent = `Com serviço: ${{formatNumber(cdoesSecao.comServicoQtd)}}`;
      document.getElementById("notaAtiva").textContent = DATA.notas.ativa;
      document.getElementById("notaVazia").textContent = DATA.notas.vazia;
      document.getElementById("notaComServico").textContent = "CDOE com serviço = possui PTP preenchido, então não está vazia.";
      updateInteractiveSelectionPanel();

      bars("topCdoesBars", falhas.topCdoes.slice(0, 8));
      bars("topSubcausasBars", falhas.topSubcausas.slice(0, 8));
      bars("topCelulasBars", falhas.topCelulas.slice(0, 8));
      bars("topMunicipiosBars", falhas.topMunicipios.slice(0, 8));
      bars("topCelulasVaziasBars", topCelulasVazias);

      document.getElementById("topCdoesTable").innerHTML = rankedTable(falhas.topCdoes);
      document.getElementById("topCelulasTable").innerHTML = rankedTable(falhas.topCelulas);
      document.getElementById("topMunicipiosTable").innerHTML = rankedTable(falhas.topMunicipios);

      document.getElementById("ativasTable").innerHTML = tableRows(cdoesSecao.ativas, 3);
      document.getElementById("vaziasTable").innerHTML = tableRows(cdoesSecao.vazias, 3);
      document.getElementById("comServicoTable").innerHTML = tableRows(cdoesSecao.comServico, 3);
    }}

    fillSelect(estacaoSelect, DATA.estacoes, {{ value: "TODAS", label: "Todas as estações" }});
    fillSelect(municipioSelect, DATA.municipios, {{ value: "TODOS", label: "Todos os municípios" }});
    fillSelect(statusCdoeSelect, [
      "ATIVAS",
      "VAZIAS",
      "COM_SERVICO",
      "NAO_ATIVAS",
      "ATENCAO"
    ], {{ value: "TODAS", label: "Todas as situações" }});
    fillSelect(statusCdoeSecaoSelect, [
      "ATIVAS",
      "COM_SERVICO",
      "NAO_ATIVAS",
      "VAZIAS",
      "ATENCAO"
    ], {{ value: "TODAS", label: "Todas nesta seção" }});

    renderCauseMacroOptions();

    statusCdoeSelect.querySelector('option[value="ATIVAS"]').textContent = "Somente ativas";
    statusCdoeSelect.querySelector('option[value="VAZIAS"]').textContent = "Somente vazias";
    statusCdoeSelect.querySelector('option[value="COM_SERVICO"]').textContent = "Somente com serviço";
    statusCdoeSelect.querySelector('option[value="NAO_ATIVAS"]').textContent = "Somente não ativas";
    statusCdoeSelect.querySelector('option[value="ATENCAO"]').textContent = "Somente em atenção";
    statusCdoeSecaoSelect.querySelector('option[value="ATIVAS"]').textContent = "Somente ativas";
    statusCdoeSecaoSelect.querySelector('option[value="COM_SERVICO"]').textContent = "Somente com serviço";
    statusCdoeSecaoSelect.querySelector('option[value="NAO_ATIVAS"]').textContent = "Somente não ativas";
    statusCdoeSecaoSelect.querySelector('option[value="VAZIAS"]').textContent = "Somente vazias";
    statusCdoeSecaoSelect.querySelector('option[value="ATENCAO"]').textContent = "Somente em atenção";

    dataInicioInput.value = DATA.dataMin || "";
    dataFimInput.value = DATA.dataMax || "";
    dataInicioInput.min = DATA.dataMin || "";
    dataInicioInput.max = DATA.dataMax || "";
    dataFimInput.min = DATA.dataMin || "";
    dataFimInput.max = DATA.dataMax || "";

    document.getElementById("clearInteractiveSelection").addEventListener("click", () => {{
      Object.keys(interactiveSelection).forEach((key) => interactiveSelection[key] = null);
      render();
    }});
    exportCdoesExcelButton.addEventListener("click", exportSectionToExcel);

    estacaoSelect.addEventListener("change", render);
    municipioSelect.addEventListener("change", render);
    statusCdoeSelect.addEventListener("change", render);
    statusCdoeSecaoSelect.addEventListener("change", render);
    dataInicioInput.addEventListener("change", render);
    dataFimInput.addEventListener("change", render);

    render();
  </script>
</body>
</html>
"""


def main() -> None:
    data = build_dashboard_data()
    html = dashboard_html(data)
    OUTPUT_PATH.write_text(html, encoding="utf-8")
    INDEX_OUTPUT_PATH.write_text(html, encoding="utf-8")
    print(f"Dashboard gerado em: {OUTPUT_PATH}")


if __name__ == "__main__":
    main()
