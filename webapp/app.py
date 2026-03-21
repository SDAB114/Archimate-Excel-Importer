"""
Archi Applicatie Importer - Webapp
====================================
Flask webapp om applicaties toe te voegen aan een Archi model via een formulier.

Gebruik:
  pip install flask openpyxl lxml
  python app.py
  Open: http://localhost:5000
"""

from flask import Flask, render_template, request, jsonify
import uuid, zipfile, io, os
import openpyxl
from lxml import etree

app = Flask(__name__)

# ─────────────────────────────────────────────
# CONFIGURATIE — pas deze waarden aan
# ─────────────────────────────────────────────

ARCHIMATE_BESTAND = r"C:\Archi\Bravis.archimate"
VIEW_NAAM         = "Applicatiefunctiemodel"

# Mapstructuur:
# Archimate-Excel-Importer/
#   webapp/
#     app.py
#     templates/
#       index.html

XSI = "http://www.w3.org/2001/XMLSchema-instance"



# ─────────────────────────────────────────────
# ARCHI FUNCTIES (zelfde logica als bulk script)
# ─────────────────────────────────────────────

def nieuw_id():
    return "id-" + str(uuid.uuid4())

def lees_archimate():
    parser = etree.XMLParser(remove_blank_text=True)
    if zipfile.is_zipfile(ARCHIMATE_BESTAND):
        with zipfile.ZipFile(ARCHIMATE_BESTAND, "r") as z:
            xml_naam = [f for f in z.namelist() if f.endswith(".xml")][0]
            with z.open(xml_naam) as f:
                inhoud = f.read()
        return etree.parse(io.BytesIO(inhoud), parser)
    return etree.parse(ARCHIMATE_BESTAND, parser)

def sla_op(tree):
    xml_bytes = etree.tostring(tree, pretty_print=True, xml_declaration=True, encoding="UTF-8")
    with zipfile.ZipFile(ARCHIMATE_BESTAND, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("model.xml", xml_bytes)

def bouw_id_naam_map(root):
    return {el.get("id"): el.get("name")
            for el in root.iter()
            if el.get("id") and el.get("name")}

def haal_views_op():
    """Haalt alle viewnamen op uit het model."""
    tree = lees_archimate()
    root = tree.getroot()
    views = []
    for el in root.iter():
        typ  = el.get(f"{{{XSI}}}type") or ""
        naam = el.get("name")
        if naam and "DiagramModel" in typ:
            views.append(naam)
    return sorted(views)


def haal_targets_op(view_naam):
    """
    Haalt alle Groupings en ApplicationFunctions op uit een specifieke view.
    Geeft een lijst van dicts terug: [{naam, type}]
    """
    tree = lees_archimate()
    root = tree.getroot()
    id_naam = bouw_id_naam_map(root)

    # Bouw id -> xsi:type map
    id_type = {}
    for el in root.iter():
        eid  = el.get("id")
        typ  = el.get(f"{{{XSI}}}type") or ""
        if eid and typ:
            id_type[eid] = typ

    view = next((el for el in root.iter() if el.get("name") == view_naam), None)
    if view is None:
        return []

    targets = []
    namen   = set()
    for el in view.iter():
        ref = el.get("archimateElement")
        if not ref or ref not in id_naam:
            continue
        naam = id_naam[ref]
        typ  = id_type.get(ref, "")
        if naam in namen:
            continue
        if "ApplicationFunction" in typ:
            targets.append({"naam": naam, "type": "ApplicationFunction"})
            namen.add(naam)
        elif "BusinessService" in typ or "Grouping" in typ or not typ:
            targets.append({"naam": naam, "type": "Grouping"})
            namen.add(naam)

    return sorted(targets, key=lambda x: x["naam"])

def zoek_elementen_folder(root):
    for el in root:
        if el.tag == "folder" and el.get("type") == "application":
            for child in el:
                if child.tag == "folder" and child.get("name") == "Elementen":
                    return child
            return el
    return root

def zoek_relaties_folder(root):
    for el in root:
        if el.tag == "folder" and el.get("type") == "relations":
            return el
    return root

def voeg_app_toe(applicatie_naam, target_naam, target_type, view_naam):
    tree = lees_archimate()
    root = tree.getroot()
    id_naam = bouw_id_naam_map(root)

    # Zoek view
    view = next((el for el in root.iter() if el.get("name") == view_naam), None)
    if view is None:
        return False, f"View '{view_naam}' niet gevonden"

    # Zoek target element in model
    grouping_el = next(
        (el for el in root.iter()
         if (el.get("name") or "").lower() == target_naam.lower()),
        None
    )
    if grouping_el is None:
        return False, f"'{target_naam}' niet gevonden in model"
    grouping_id = grouping_el.get("id")
    grouping_naam = target_naam

    # Controleer of applicatie al bestaat
    bestaand = next(
        (el for el in root.iter()
         if (el.get("name") or "").lower() == applicatie_naam.lower()
         and "ApplicationComponent" in (el.get(f"{{{XSI}}}type") or "")),
        None
    )
    if bestaand is not None:
        app_id = bestaand.get("id")
        app_status = "bestaand"
    else:
        folder = zoek_elementen_folder(root)
        nieuw_el = etree.SubElement(folder, "element")
        nieuw_el.set(f"{{{XSI}}}type", "archimate:ApplicationComponent")
        nieuw_el.set("name", applicatie_naam)
        app_id = nieuw_id()
        nieuw_el.set("id", app_id)
        app_status = "nieuw"

    # Voeg relatie toe — Composition voor beide typen
    rel_bestaat = any(
        "CompositionRelationship" in (el.get(f"{{{XSI}}}type") or "")
        and el.get("source") == grouping_id
        and el.get("target") == app_id
        for el in root.iter()
    )
    if not rel_bestaat:
        rel_folder = zoek_relaties_folder(root)
        rel = etree.SubElement(rel_folder, "element")
        rel.set(f"{{{XSI}}}type", "archimate:CompositionRelationship")
        rel.set("id", nieuw_id())
        rel.set("source", grouping_id)
        rel.set("target", app_id)

    sla_op(tree)
    return True, f"'{applicatie_naam}' succesvol toegevoegd aan '{grouping_naam}' ({app_status})"

# ─────────────────────────────────────────────
# ROUTES
# ─────────────────────────────────────────────

@app.route("/")
def index():
    views = haal_views_op()
    return render_template("index.html", views=views, view_naam=VIEW_NAAM)

@app.route("/targets", methods=["POST"])
def targets():
    data      = request.get_json()
    view_naam = (data.get("view_naam") or "").strip()
    return jsonify(haal_targets_op(view_naam))


@app.route("/alle-functies")
def alle_functies():
    """Haalt alle ApplicationFunctions op uit het hele model."""
    tree = lees_archimate()
    root = tree.getroot()
    functies = []
    namen    = set()
    for el in root.iter():
        typ  = el.get(f"{{{XSI}}}type") or ""
        naam = el.get("name")
        if naam and "ApplicationFunction" in typ and naam not in namen:
            functies.append({"naam": naam, "type": "ApplicationFunction"})
            namen.add(naam)
    return jsonify(sorted(functies, key=lambda x: x["naam"]))


@app.route("/maak-view", methods=["POST"])
def maak_view():
    """Maakt een nieuwe view aan met een applicatiefunctie en alle gerelateerde applicaties."""
    data       = request.get_json()
    view_naam  = (data.get("view_naam") or "").strip()
    functie_naam = (data.get("functie_naam") or "").strip()

    if not view_naam or not functie_naam:
        return jsonify({"success": False, "bericht": "Vul alle velden in"})

    tree = lees_archimate()
    root = tree.getroot()
    id_naam = bouw_id_naam_map(root)

    # Controleer of view al bestaat
    bestaande_view = next(
        (el for el in root.iter() if el.get("name") == view_naam),
        None
    )
    if bestaande_view is not None:
        return jsonify({"success": False, "bericht": f"View '{view_naam}' bestaat al"})

    # Zoek applicatiefunctie
    fn_el = next(
        (el for el in root.iter()
         if (el.get("name") or "").lower() == functie_naam.lower()
         and "ApplicationFunction" in (el.get(f"{{{XSI}}}type") or "")),
        None
    )
    if fn_el is None:
        return jsonify({"success": False, "bericht": f"Applicatiefunctie '{functie_naam}' niet gevonden"})

    fn_id = fn_el.get("id")

    # Zoek alle gerelateerde ApplicationComponents via CompositionRelationship
    app_ids = [
        el.get("target") for el in root.iter()
        if "CompositionRelationship" in (el.get(f"{{{XSI}}}type") or "")
        and el.get("source") == fn_id
        and el.get("target")
    ]

    # Zoek de Views folder
    views_folder = next(
        (el for el in root.iter()
         if el.tag == "folder" and (el.get("type") == "diagrams" or el.get("name") in ("Views", "Diagrams"))),
        root
    )

    # Maak nieuwe view aan
    view_el = etree.SubElement(views_folder, "element")
    view_el.set(f"{{{XSI}}}type", "archimate:ArchimateDiagramModel")
    view_el.set("name", view_naam)
    view_el.set("id",   nieuw_id())

    # Afmetingen
    APP_B, APP_H, MARGE = 120, 55, 15
    FN_BREEDTE = max(400, len(app_ids) * (APP_B + MARGE) + MARGE)
    FN_HOOGTE  = 100 + (((len(app_ids) - 1) // 6) + 1) * (APP_H + MARGE) + MARGE

    # Voeg applicatiefunctie toe aan view
    fn_view = etree.SubElement(view_el, "child")
    fn_view.set(f"{{{XSI}}}type", "archimate:DiagramObject")
    fn_view.set("id",               nieuw_id())
    fn_view.set("archimateElement", fn_id)
    fn_bounds = etree.SubElement(fn_view, "bounds")
    fn_bounds.set("x",      "24")
    fn_bounds.set("y",      "24")
    fn_bounds.set("width",  str(FN_BREEDTE))
    fn_bounds.set("height", str(FN_HOOGTE))

    # Voeg applicaties toe als children binnen de applicatiefunctie
    for i, app_id in enumerate(app_ids):
        kolom  = i % 6
        rij_nr = i // 6
        x = MARGE + kolom * (APP_B + MARGE)
        y = 30 + rij_nr * (APP_H + MARGE)

        app_child = etree.SubElement(fn_view, "child")
        app_child.set(f"{{{XSI}}}type", "archimate:DiagramObject")
        app_child.set("id",               nieuw_id())
        app_child.set("archimateElement", app_id)
        ab = etree.SubElement(app_child, "bounds")
        ab.set("x",      str(x))
        ab.set("y",      str(y))
        ab.set("width",  str(APP_B))
        ab.set("height", str(APP_H))

    sla_op(tree)
    return jsonify({
        "success": True,
        "bericht": f"View '{view_naam}' aangemaakt met {len(app_ids)} applicaties"
    })

@app.route("/toevoegen", methods=["POST"])
def toevoegen():
    data        = request.get_json()
    applicatie  = (data.get("applicatie") or "").strip()
    target      = (data.get("grouping") or "").strip()
    target_type = (data.get("target_type") or "Grouping").strip()
    view_naam   = (data.get("view_naam") or VIEW_NAAM).strip()

    if not applicatie or not target or not view_naam:
        return jsonify({"success": False, "bericht": "Vul alle velden in"})

    success, bericht = voeg_app_toe(applicatie, target, target_type)
    return jsonify({"success": success, "bericht": bericht})

@app.route("/bulk", methods=["POST"])
def bulk():
    if "bestand" not in request.files:
        return jsonify({"success": False, "bericht": "Geen bestand ontvangen"})

    bestand = request.files["bestand"]
    if not bestand.filename.endswith((".xlsx", ".xls")):
        return jsonify({"success": False, "bericht": "Alleen .xlsx of .xls bestanden zijn toegestaan"})

    import tempfile, os
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    bestand.save(tmp.name)
    tmp.close()

    try:
        wb = openpyxl.load_workbook(tmp.name)
        ws = wb.active

        # Lees rijen vanaf rij 2 (rij 1 = headers)
        rijen = []
        for rij in ws.iter_rows(min_row=2, values_only=True):
            applicatie = rij[0]
            functie    = rij[1] if len(rij) > 1 else None
            if applicatie and str(applicatie).strip():
                rijen.append((
                    str(applicatie).strip(),
                    str(functie).strip() if functie and str(functie).strip() else None
                ))

        if not rijen:
            return jsonify({"success": False, "bericht": "Geen geldige rijen gevonden in het bestand"})

        # Verwerk het model
        tree = lees_archimate()
        root = tree.getroot()

        toegevoegd  = 0
        overgeslagen = 0
        resultaten  = []

        for applicatie_naam, functie_naam in rijen:

            # Zoek applicatiefunctie
            fn_el = next(
                (el for el in root.iter()
                 if (el.get("name") or "").lower() == functie_naam.lower()
                 and "ApplicationFunction" in (el.get(f"{{{XSI}}}type") or "")),
                None
            ) if functie_naam else None

            if fn_el is None:
                resultaten.append(f"⚠ '{functie_naam}' niet gevonden in model")
                overgeslagen += 1
                continue

            fn_id = fn_el.get("id")

            # Controleer of ApplicationComponent al bestaat
            app_el = next(
                (el for el in root.iter()
                 if (el.get("name") or "").lower() == applicatie_naam.lower()
                 and "ApplicationComponent" in (el.get(f"{{{XSI}}}type") or "")),
                None
            )
            if app_el is not None:
                app_id = app_el.get("id")
                # Applicatie bestaat al — check of relatie ook al bestaat
                rel_bestaat = any(
                    "CompositionRelationship" in (el.get(f"{{{XSI}}}type") or "")
                    and el.get("source") == fn_id
                    and el.get("target") == app_id
                    for el in root.iter()
                )
                if rel_bestaat:
                    resultaten.append(f"~ '{applicatie_naam}' → '{functie_naam}' (relatie bestaat al)")
                    overgeslagen += 1
                    continue
                else:
                    # Maak alleen de relatie aan
                    rel_folder = zoek_relaties_folder(root)
                    rel = etree.SubElement(rel_folder, "element")
                    rel.set(f"{{{XSI}}}type", "archimate:CompositionRelationship")
                    rel.set("id", nieuw_id())
                    rel.set("source", fn_id)
                    rel.set("target", app_id)
                    resultaten.append(f"✓ '{applicatie_naam}' → '{functie_naam}' (relatie toegevoegd)")
                    toegevoegd += 1
                    continue

            # Maak nieuw ApplicationComponent aan
            folder = zoek_elementen_folder(root)
            nieuw_el = etree.SubElement(folder, "element")
            nieuw_el.set(f"{{{XSI}}}type", "archimate:ApplicationComponent")
            nieuw_el.set("name", applicatie_naam)
            app_id = nieuw_id()
            nieuw_el.set("id", app_id)

            # Maak relatie aan
            rel_folder = zoek_relaties_folder(root)
            rel = etree.SubElement(rel_folder, "element")
            rel.set(f"{{{XSI}}}type", "archimate:CompositionRelationship")
            rel.set("id", nieuw_id())
            rel.set("source", fn_id)
            rel.set("target", app_id)

            resultaten.append(f"✓ '{applicatie_naam}' → '{functie_naam}' (nieuw)")
            toegevoegd += 1

        sla_op(tree)

        bericht = f"{toegevoegd} toegevoegd, {overgeslagen} overgeslagen"
        return jsonify({"success": True, "bericht": bericht, "resultaten": resultaten})

    except Exception as e:
        return jsonify({"success": False, "bericht": f"Fout: {str(e)}"})
    finally:
        os.unlink(tmp.name)


if __name__ == "__main__":
    print("=" * 45)
    print("  Archi Applicatie Importer")
    print("=" * 45)
    print(f"  Model : {ARCHIMATE_BESTAND}")
    print(f"  View  : {VIEW_NAAM}")
    print(f"  Open  : http://localhost:5000")
    print("=" * 45)
    app.run(debug=False)
