# Archimate-Excel-Importer

Een Python script dat een Excel-bestand omzet naar CSV-bestanden die je direct kunt importeren in **Archi** als ArchiMate-model, gebaseerd op het **ZiRA**-raamwerk.

---

## Functionaliteit

- Leest applicatienamen en bijbehorende Grouping-nodes uit een Excel-bestand
- Maakt **Application Components** aan in ArchiMate
- Maakt **Grouping-elementen** aan (uniek, geen duplicaten)
- Legt **Composition-relaties** van elke Grouping naar de bijbehorende applicaties
- Genereert drie CSV-bestanden die Archi direct kan importeren

---

## Vereisten

- Python 3.8 of hoger
- [Archi](https://www.archimatetool.com/) met CSV-import ondersteuning
- Python library: `openpyxl`

Installeer de library via:

```bash
pip install openpyxl
```

---

## Excel-formaat

Het script verwacht een Excel-bestand (`.xlsx`) met de volgende structuur:

| Applicatie   | Grouping-Node  |
|--------------|----------------|
| SAP ERP      | Finance        |
| Salesforce   | CRM            |
| ServiceNow   | IT Management  |
| SAP BW       | Finance        |

- **Kolom A – Applicatie**: naam van de applicatie
- **Kolom B – Grouping-Node**: naam van de Grouping waartoe de applicatie behoort

> Als kolom B leeg is, wordt de applicatie automatisch ingedeeld onder de Grouping **"Overig"**.

---

## Gebruik

**1. Configureer het script**

Pas de variabelen bovenin `excel_to_archi_csv.py` aan:

```python
EXCEL_BESTAND     = "applicaties.xlsx"  # pad naar je Excel-bestand
KOLOM_APPLICATIE  = "Applicatie"        # kolomnaam kolom A
KOLOM_GROUPING    = "Grouping-Node"     # kolomnaam kolom B
SHEET_NAAM        = None                # None = eerste sheet
UITVOER_MAP       = "archi_import"      # map voor de gegenereerde CSV-bestanden
```

**2. Voer het script uit**

```bash
python excel_to_archi_csv.py
```

**3. Importeer in Archi**

- Open Archi
- Ga naar **File → Import → CSV**
- Selecteer de map `archi_import/`
- Klik **Finish**

---

## Gegenereerde bestanden

Het script maakt een map `archi_import/` aan met drie bestanden:

| Bestand           | Inhoud                                      |
|-------------------|---------------------------------------------|
| `elements.csv`    | Grouping-nodes en Application Components    |
| `relations.csv`   | Composition-relaties                        |
| `properties.csv`  | Properties (leeg, uitbreidbaar)             |

---

## ArchiMate-structuur

```
Grouping (bijv. "Finance")
  └── [Composition] ApplicationComponent (bijv. "SAP ERP")
  └── [Composition] ApplicationComponent (bijv. "SAP BW")

Grouping (bijv. "CRM")
  └── [Composition] ApplicationComponent (bijv. "Salesforce")
```

---

## Uitbreidingsmogelijkheden

- Extra kolommen uit Excel toevoegen als **properties** (bijv. leverancier, status, eigenaar)
- Ondersteuning voor **meerdere sheets**
- Genereren van een **view/diagram** in Archi via jArchi scripting

---

## Licentie

MIT License
