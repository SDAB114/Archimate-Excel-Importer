# Archimate-Excel-Importer

Een Python script dat een Excel-bestand omzet en direct importeert in een bestaand **Archi** `.archimate` model, gebaseerd op het **ZiRA**-raamwerk.

---

## Functionaliteit

- Leest applicatienamen en bijbehorende Grouping-nodes uit een Excel-bestand
- Maakt **Application Components** aan in het ArchiMate-model
- Voegt **Composition-relaties** toe van elke Grouping naar de bijbehorende applicaties
- Plaatst de applicaties visueel in de juiste Grouping-blokken in een bestaande view
- Vergroot automatisch de hoogte van een Grouping-blok als er meer ruimte nodig is
- Ondersteunt zowel gecomprimeerde als ongecomprimeerde `.archimate` bestanden

---

## Vereisten

- Python 3.8 of hoger
- [Archi](https://www.archimatetool.com/)
- Python libraries: `openpyxl` en `lxml`

Installeer de libraries via:

```bash
pip install openpyxl lxml
```

---

## Excel-formaat

Het script verwacht een Excel-bestand (`.xlsx`) met de volgende structuur:

| Applicatie   | Grouping-Node         |
|--------------|-----------------------|
| HiX          | Behandeling           |
| Sectra       | Aanvullend onderzoek  |
| AFAS         | Bedrijfsondersteuning |

- **Kolom A – Applicatie**: naam van de applicatie
- **Kolom B – Grouping-Node**: naam van de Grouping waartoe de applicatie behoort (case-insensitief)

> Als kolom B leeg is, wordt de applicatie automatisch ingedeeld onder de Grouping **"Overig"**.

---

## Gebruik

**1. Configureer het script**

Pas de variabelen bovenin `excel_to_archi_xml.py` aan:

```python
EXCEL_BESTAND     = r"C:\pad\naar\applicaties.xlsx"
ARCHIMATE_BESTAND = r"C:\pad\naar\model.archimate"
UITVOER_BESTAND   = r"C:\pad\naar\model_updated.archimate"

KOLOM_APPLICATIE  = "Applicatie"       # kolomnaam kolom A
KOLOM_GROUPING    = "Grouping-Node"    # kolomnaam kolom B
SHEET_NAAM        = None               # None = eerste sheet

VIEW_NAAM         = "Applicatiefunctiemodel"  # naam van de bestaande view in Archi
```

**2. Voer het script uit**

```bash
python excel_to_archi_xml.py
```

**3. Open het resultaat in Archi**

```
File → Open → model_updated.archimate
```

Het originele `.archimate` bestand blijft ongewijzigd bewaard.

---

## Hoe het werkt

```
Excel (Applicatie + Grouping-Node)
        ↓
Python script
        ↓
1. Leest bestaand .archimate model (ook gecomprimeerd)
2. Zoekt de opgegeven Grouping-nodes in het model
3. Maakt Application Components aan in de Application → Elementen folder
4. Voegt Composition-relaties toe in de Relations folder
5. Plaatst de applicaties visueel in de juiste Grouping-blokken in de view
6. Vergroot Grouping-blokken automatisch indien nodig
7. Slaat op als geldig .archimate bestand
```

---

## ArchiMate-structuur

```
Grouping (bijv. "Behandeling")
  └── [Composition] ApplicationComponent (bijv. "HiX")

Grouping (bijv. "Aanvullend onderzoek")
  └── [Composition] ApplicationComponent (bijv. "Sectra")
```

---

## Uitbreidingsmogelijkheden

- Extra kolommen uit Excel toevoegen als **properties** (bijv. leverancier, status, eigenaar)
- Ondersteuning voor **meerdere sheets**
- Automatisch aanmaken van nieuwe Grouping-nodes als ze nog niet bestaan

---

## Licentie

MIT License
