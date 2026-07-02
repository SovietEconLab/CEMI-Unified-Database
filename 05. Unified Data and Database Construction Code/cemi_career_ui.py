#!/usr/bin/env python3
"""
CEMI Career Database — single-file HTML UI generator
==============================================================================
Reads:  cemi_career.db (built by `cemi_career_db.py`)
Writes: cemi_career_ui.html

The UI is centred on the *career-search* requirement: every career field
recorded in the source spreadsheets — main CEMI position, dual external post,
pre-CEMI service, transfer, dismissal, party, education, ethnicity, specialty,
department, academic title — is independently filterable, and a person's
career timeline shows promotions in the correct order, including within-year
sequencing (Acting → substantive) and dual employment.

Run from anywhere — the DB is auto-discovered the same way as the builder:
explicit --db path → CWD → script's own directory.
"""

from __future__ import annotations

import argparse
import json
import sqlite3
import sys
from pathlib import Path


DEFAULTS = {"db": "cemi_career.db", "out": "cemi_career_ui.html"}


def script_dir() -> Path:
    try:
        return Path(__file__).resolve().parent
    except NameError:
        return Path.cwd().resolve()


def find_db(arg: str) -> Path:
    cwd, here = Path.cwd().resolve(), script_dir()
    candidates: list[Path] = []
    if arg:
        p = Path(arg).expanduser()
        candidates.append(p if p.is_absolute() else (cwd / p))
        if not p.is_absolute() and p.parent == Path("."):
            candidates.append(here / p.name)
    candidates.append(cwd / DEFAULTS["db"])
    candidates.append(here / DEFAULTS["db"])
    seen: set = set()
    ordered: list[Path] = []
    for c in candidates:
        r = c.resolve()
        if r in seen: continue
        seen.add(r); ordered.append(r)
    for c in ordered:
        if c.exists():
            return c
    print("[ERROR] cemi_career.db not found.  Searched:")
    for c in ordered:
        print("       ", c)
    sys.exit(1)


def resolve_out(arg: str) -> Path:
    p = Path(arg or DEFAULTS["out"]).expanduser()
    return (p if p.is_absolute() else (Path.cwd() / p)).resolve()


# ──────────────────────────────────────────────────────────────────────────────
#  §1  Pull career-shaped data out of the DB
# ──────────────────────────────────────────────────────────────────────────────

def query_all(conn: sqlite3.Connection, sql: str, *params) -> list[dict]:
    conn.row_factory = sqlite3.Row
    return [dict(r) for r in conn.execute(sql, params).fetchall()]


def build_payload(conn: sqlite3.Connection) -> dict:
    print("  · loading vocabulary tables …")
    # Per-year institute-phase map.  Years younger than 1963 came from the
    # predecessor lab; 1963+ are CEMI proper.  Anything older is a life-event
    # year (party membership / BA graduation, etc.) and is tagged Pre-Institute.
    year_phase = {}
    for r in conn.execute(
        "SELECT year, institute_phase, institute_label FROM calendar_year ORDER BY year"):
        year_phase[int(r[0])] = {"phase": r[1], "label": r[2]}

    vocab = {
        "position":    query_all(conn, "SELECT position_id AS id, label_en AS label, label_ru, is_acting, rank_order, tier FROM position_rank ORDER BY rank_order, label_en"),
        "department":  query_all(conn, "SELECT department_id AS id, label FROM department ORDER BY label"),
        "institution": query_all(conn, "SELECT institution_id AS id, label, kind, label_ru, classification_evidence FROM institution ORDER BY label"),
        "specialty":   query_all(conn, "SELECT specialty_id AS id, label FROM specialty ORDER BY label"),
        "phd_field":   query_all(conn, "SELECT field_id AS id, label FROM phd_field ORDER BY label"),
        "school":      query_all(conn, "SELECT school_id AS id, label FROM school ORDER BY label"),
        "degree":      query_all(conn, "SELECT degree_id AS id, label FROM degree ORDER BY label"),
        "ethnicity":   query_all(conn, "SELECT ethnicity_id AS id, label_en AS label FROM ethnicity ORDER BY label_en"),
        "academic_title": query_all(conn, "SELECT title_id AS id, label_en AS label, rank_order FROM academic_title ORDER BY rank_order DESC"),
        "party":       query_all(conn, "SELECT party_id AS id, label FROM party_org ORDER BY label"),
    }

    print("  · building per-person career index …")
    # ── Person summary, with computed first/last year and event counts ──
    persons = query_all(conn, """
        SELECT  p.person_id    AS id,
                p.canonical_name AS name,
                p.surname        AS surname,
                p.initials       AS initials,
                p.birth_year     AS birth_year,
                e.label_en       AS ethnicity,
                MIN(o.year)      AS first_year,
                MAX(o.year)      AS last_year,
                COUNT(DISTINCT o.year) AS years_observed
          FROM person p
          LEFT JOIN ethnicity e ON e.ethnicity_id = p.ethnicity_id
          LEFT JOIN person_year_observation o ON o.person_id = p.person_id
          GROUP BY p.person_id
          ORDER BY p.canonical_name
    """)

    # Build helper lookups for the search filter
    appt_by_person: dict[int, list[int]] = {}
    for r in conn.execute("SELECT DISTINCT person_id, position_id FROM cemi_appointment"):
        appt_by_person.setdefault(r[0], []).append(r[1])
    inst_by_person_pre: dict[int, set] = {}
    for r in conn.execute("SELECT person_id, institution_id FROM pre_cemi_role"):
        inst_by_person_pre.setdefault(r[0], set()).add(r[1])
    inst_by_person_dual: dict[int, set] = {}
    for r in conn.execute("SELECT person_id, institution_id FROM dual_position"):
        inst_by_person_dual.setdefault(r[0], set()).add(r[1])
    inst_by_person_tran: dict[int, set] = {}
    for r in conn.execute("SELECT person_id, institution_id FROM transfer_event"):
        inst_by_person_tran.setdefault(r[0], set()).add(r[1])
    party_by_person: dict[int, set] = {}
    for r in conn.execute("SELECT person_id, party_id FROM party_event"):
        party_by_person.setdefault(r[0], set()).add(r[1])
    field_by_person: dict[int, set] = {}
    for r in conn.execute("SELECT person_id, field_id FROM phd_defense WHERE field_id IS NOT NULL"):
        field_by_person.setdefault(r[0], set()).add(r[1])
    spec_by_person: dict[int, set] = {}
    for r in conn.execute("SELECT person_id, specialty_id FROM person_year_observation WHERE specialty_id IS NOT NULL"):
        spec_by_person.setdefault(r[0], set()).add(r[1])
    dept_by_person: dict[int, set] = {}
    for r in conn.execute("SELECT person_id, department_id FROM person_year_observation WHERE department_id IS NOT NULL"):
        dept_by_person.setdefault(r[0], set()).add(r[1])
    title_by_person: dict[int, set] = {}
    for r in conn.execute("SELECT person_id, title_id FROM academic_title_award"):
        title_by_person.setdefault(r[0], set()).add(r[1])
    deg_by_person: dict[int, set] = {}
    for r in conn.execute("SELECT person_id, degree_id FROM person_year_observation WHERE degree_id IS NOT NULL"):
        deg_by_person.setdefault(r[0], set()).add(r[1])
    school_by_person: dict[int, set] = {}
    for r in conn.execute("SELECT person_id, grad_school_id FROM person_year_observation WHERE grad_school_id IS NOT NULL"):
        school_by_person.setdefault(r[0], set()).add(r[1])
    for r in conn.execute("SELECT person_id, ba_school_id FROM person_year_observation WHERE ba_school_id IS NOT NULL"):
        school_by_person.setdefault(r[0], set()).add(r[1])
    inst_primary_by_person: dict[int, set] = {}
    for r in conn.execute("SELECT person_id, primary_institution_id FROM person_year_observation WHERE primary_institution_id IS NOT NULL"):
        inst_primary_by_person.setdefault(r[0], set()).add(r[1])

    # Per-institution count of distinct persons whose Primary Institution
    # (column "Primary Institution" in data.xlsx) equals that institution.
    #     • For CEMI itself:           full-time CEMI staff.
    #     • For any external institution: that institution is the person's
    #       PRIMARY employer and CEMI is therefore a part-time / dual post.
    # The two cases together exhaust the "primary" relationship axis.
    inst_primary_count: dict[int, int] = {}
    for inst_set in inst_primary_by_person.values():
        for iid in inst_set:
            inst_primary_count[iid] = inst_primary_count.get(iid, 0) + 1

    # Identify the CEMI institution row so the JS-side can distinguish
    # "Primary = CEMI" from "Primary = external".  The institution table
    # stores CEMI with label == 'CEMI'; the COALESCE guards against the
    # ru/en/abbrev variants the loader has used historically.
    cemi_row = conn.execute(
        "SELECT institution_id FROM institution "
        "WHERE label = 'CEMI' "
        "   OR label = 'ЦЭМИ' "
        "   OR label = 'Central Economic Mathematical Institute' "
        "   OR label = 'Central Economic-Mathematical Institute' "
        "ORDER BY institution_id LIMIT 1"
    ).fetchone()
    cemi_inst_id = cemi_row[0] if cemi_row else None

    # Compute, for every person, which institutional phases they were
    # observed in and the year-range of each phase.  Used by the profile
    # header line "Predecessor Lab 1961–1962 → CEMI 1963–1965".
    obs_by_person: dict = {}
    for r in conn.execute("SELECT person_id, year FROM person_year_observation WHERE year IS NOT NULL"):
        obs_by_person.setdefault(int(r[0]), []).append(int(r[1]))

    def _phase_for(year):
        if year >= 1963:  return "CEMI"
        if year >= 1958:  return "Predecessor Lab"
        return "Pre-Institute"

    for p in persons:
        pid = p["id"]
        # Phase spans: {phase: [first_year, last_year]}
        spans = {}
        for y in obs_by_person.get(pid, []):
            ph = _phase_for(y)
            if ph not in spans: spans[ph] = [y, y]
            else:
                if y < spans[ph][0]: spans[ph][0] = y
                if y > spans[ph][1]: spans[ph][1] = y
        p["phase_spans"] = spans

        # Compose “any institution this person is connected to”
        all_inst: set = set()
        for src in (inst_primary_by_person, inst_by_person_pre, inst_by_person_dual, inst_by_person_tran):
            all_inst |= src.get(pid, set())
        p["positions"]    = sorted(appt_by_person.get(pid, []))
        p["institutions"] = sorted(all_inst)
        p["parties"]      = sorted(party_by_person.get(pid, set()))
        p["fields"]       = sorted(field_by_person.get(pid, set()))
        p["specialties"]  = sorted(spec_by_person.get(pid, set()))
        p["departments"]  = sorted(dept_by_person.get(pid, set()))
        p["titles"]       = sorted(title_by_person.get(pid, set()))
        p["degrees"]      = sorted(deg_by_person.get(pid, set()))
        p["schools"]      = sorted(school_by_person.get(pid, set()))

    # ── Per-person events grouped by year ────────────────────────────────
    print("  · loading career events per person …")
    appts = query_all(conn, """
        SELECT a.person_id, a.year, a.within_year_order AS ord,
               p.label_en AS position, a.is_acting, a.rank_order AS rank,
               p.tier
          FROM cemi_appointment a JOIN position_rank p ON p.position_id = a.position_id
         ORDER BY a.person_id, a.year, a.within_year_order""")
    duals = query_all(conn, """
        SELECT d.person_id, d.year, i.label AS institution, d.position_label AS role,
               d.duration_text AS duration, d.institution_kind AS kind
          FROM dual_position d JOIN institution i ON i.institution_id = d.institution_id
         ORDER BY d.person_id, d.year""")
    pres  = query_all(conn, """
        SELECT pr.person_id, i.label AS institution, pr.title, pr.institution_kind AS kind
          FROM pre_cemi_role pr JOIN institution i ON i.institution_id = pr.institution_id
         ORDER BY pr.person_id""")
    trans = query_all(conn, """
        SELECT t.person_id, t.year, i.label AS institution, t.institution_kind AS kind
          FROM transfer_event t JOIN institution i ON i.institution_id = t.institution_id
         ORDER BY t.person_id, t.year""")
    diss  = query_all(conn, """
        SELECT d.person_id, d.year, d.dismissal_date AS date,
               p.label_en AS position
          FROM dismissal_event d
          LEFT JOIN position_rank p ON p.position_id = d.position_id
         ORDER BY d.person_id, d.year""")
    party_ev = query_all(conn, """
        SELECT pe.person_id, po.label AS party, pe.role, pe.join_year, pe.join_date
          FROM party_event pe JOIN party_org po ON po.party_id = pe.party_id
         ORDER BY pe.person_id, pe.join_year""")
    phd_ev = query_all(conn, """
        SELECT pd.person_id, f.label AS field, pd.defense_date AS date
          FROM phd_defense pd LEFT JOIN phd_field f ON f.field_id = pd.field_id
         ORDER BY pd.person_id, pd.defense_date""")
    # v5: enrich start_of_work with the position-span context so the
    # timeline can label rows precisely (Joined CEMI for span_order = 1
    # entries; Position change → <position> for subsequent rows).
    starts = query_all(conn, """
        SELECT s.person_id, s.start_date, s.is_initial_join,
               s.span_id, ps.span_order, pr.label_en AS position,
               ps.is_acting AS span_is_acting,
               ps.start_year AS span_start_year, ps.end_year AS span_end_year
          FROM start_of_work s
          LEFT JOIN cemi_position_span ps ON ps.span_id = s.span_id
          LEFT JOIN position_rank pr      ON pr.position_id = ps.position_id
         ORDER BY s.person_id, s.start_date""")
    titles = query_all(conn, """
        SELECT a.person_id, t.label_en AS title, a.first_year
          FROM academic_title_award a JOIN academic_title t ON t.title_id = a.title_id
         ORDER BY a.person_id""")
    obs = query_all(conn, """
        SELECT o.person_id, o.year,
               i.label AS primary_institution,
               d.label AS department,
               s.label AS specialty,
               g.label AS degree,
               t.label_en AS academic_title,
               gs.label AS grad_school,
               ba.label AS ba_school
          FROM person_year_observation o
          LEFT JOIN institution i ON i.institution_id = o.primary_institution_id
          LEFT JOIN department d ON d.department_id = o.department_id
          LEFT JOIN specialty s ON s.specialty_id = o.specialty_id
          LEFT JOIN degree g ON g.degree_id = o.degree_id
          LEFT JOIN academic_title t ON t.title_id = o.academic_title_id
          LEFT JOIN school gs ON gs.school_id = o.grad_school_id
          LEFT JOIN school ba ON ba.school_id = o.ba_school_id
         ORDER BY o.person_id, o.year""")

    def group_by_person(rows: list[dict]) -> dict[int, list[dict]]:
        out: dict[int, list[dict]] = {}
        for r in rows:
            pid = r.pop("person_id")
            out.setdefault(pid, []).append(r)
        return out

    # Derived spans: one row per continuous stretch a person held the same
    # (position, is_acting).  When the DB lacks the table (older build) the
    # UI's JS still has a fallback that recomputes on the fly.
    try:
        spans = query_all(conn, """
            SELECT cps.person_id, cps.position_id, cps.is_acting,
                   cps.start_year, cps.end_year, cps.n_years, cps.span_order,
                   p.label_en AS position, p.tier, p.rank_order AS rank,
                   (SELECT MIN(s.start_date) FROM start_of_work s
                     WHERE s.span_id = cps.span_id
                       AND s.start_date IS NOT NULL) AS span_start_date
              FROM cemi_position_span cps
              JOIN position_rank p ON p.position_id = cps.position_id
             ORDER BY cps.person_id, cps.span_order""")
    except Exception:
        spans = []

    # Derived per-person degree history.  One row per (person, degree).
    # Missing on older DBs → falls back to client-side derivation in JS.
    try:
        degree_hist = query_all(conn, """
            SELECT pd.person_id, pd.degree_id, pd.first_year, pd.last_year,
                   pd.n_years, d.label AS degree,
                   pd.school_id, s.label    AS school,
                   pd.defense_date,
                   pd.field_id,  pf.label   AS field
              FROM person_degree pd
              JOIN degree d ON d.degree_id = pd.degree_id
              LEFT JOIN school s    ON s.school_id  = pd.school_id
              LEFT JOIN phd_field pf ON pf.field_id = pd.field_id
             ORDER BY pd.person_id, pd.first_year""")
    except Exception:
        degree_hist = []

    events = {
        "appointments": group_by_person(appts),
        "duals":        group_by_person(duals),
        "pre_cemi":     group_by_person(pres),
        "transfers":    group_by_person(trans),
        "dismissals":   group_by_person(diss),
        "party":        group_by_person(party_ev),
        "phd":          group_by_person(phd_ev),
        "starts":       group_by_person(starts),
        "titles":       group_by_person(titles),
        "observations": group_by_person(obs),
        "position_spans": group_by_person(spans),
        "degree_history": group_by_person(degree_hist),
    }

    # ── Aggregate / demographic data for the dashboard tabs ──────────────
    print("  · loading aggregate dashboards …")
    agg = {
        "personnel_totals": query_all(conn, """
            SELECT year, total_scientists, total_all_staff, national_doctor, phd,
                   women, women_national_doctor, women_phd, source_material
              FROM demo_personnel_totals ORDER BY year"""),
        "academic_degrees": query_all(conn, """
            SELECT year, total_scientists, national_doctor, phd, pct_with_degree,
                   professors, docents, sns_title, academicians, source_material
              FROM demo_academic_degrees ORDER BY year"""),
        "positions": query_all(conn, """
            SELECT dp.year, p.label_en AS position, p.label_ru AS position_ru,
                   p.rank_order, p.is_acting, p.tier,
                   dp.total, dp.national_doctor, dp.phd,
                   dp.cpsu_members, dp.komsomol, dp.source_material
              FROM demo_position dp JOIN position_rank p ON p.position_id = dp.position_id
              ORDER BY dp.year, p.rank_order"""),
        "nationality": query_all(conn, """
            SELECT dn.year, n.label_ru AS ru, n.label_en AS en,
                   dn.total, dn.national_doctor, dn.phd
              FROM demo_nationality dn JOIN nationality n ON n.nationality_id = dn.nationality_id
              ORDER BY dn.year, dn.total DESC"""),
        "age": query_all(conn, """
            SELECT da.year, ab.label_ru AS ru, ab.label_en AS en, ab.lower_bound, ab.upper_bound,
                   da.total, da.national_doctor, da.phd,
                   da.academician_or_corr, da.professor, da.associate_professor,
                   da.senior_researcher, da.junior_researcher
              FROM demo_age da JOIN age_bracket ab ON ab.bracket_id = da.bracket_id
              ORDER BY da.year, ab.lower_bound"""),
        "fields": query_all(conn, """
            SELECT df.year, rf.code, rf.label_ru AS ru, rf.label_en AS en, rf.is_aggregate,
                   df.total, df.national_doctor, df.phd
              FROM demo_field df JOIN research_field rf ON rf.rfield_id = df.rfield_id
              ORDER BY df.year, rf.label_en"""),
        "party": query_all(conn, """
            SELECT year, total_scientists, cpsu_members, cpsu_pct, komsomol, komsomol_pct,
                   source_material
              FROM demo_party ORDER BY year"""),
        "trainees": query_all(conn, """
            SELECT year, total_trainees, from_other_institutes
              FROM demo_trainees ORDER BY year"""),
        # `subfields_by_period` is keyed by the sheet name from
        # research_field_subfield.xlsx.  In the current workbook every
        # sheet is named plain "YYYY", so period == str(year).  Older
        # YYYY.MM-labelled workbooks (1971.01 / 1971.06) are still
        # supported transparently — the key just holds whatever the
        # sheet was called.  `subfields_by_year` is the legacy alias.
        "subfields_by_period": {},
        "subfields_by_year":   {},
        "glossary": query_all(conn, """
            SELECT russian, english, category FROM glossary ORDER BY glossary_id"""),
        "sheets_index": query_all(conn, """
            SELECT year, institution, source_file, sheet_name,
                   first_cell_ru, first_cell_en, rows, cols,
                   role, target_table
              FROM sheets_index ORDER BY role, source_file, year, sheet_name"""),
    }
    for r in conn.execute("""
        SELECT ds.year, ds.period, rf.label_en AS field, rs.label_en AS subfield,
               ds.total_personnel, ds.national_doctor, ds.phd,
               ds.extras_json, ds.schema_year_signature
          FROM demo_subfield ds
          JOIN research_subfield rs ON rs.subfield_id = ds.subfield_id
          JOIN research_field rf ON rf.rfield_id = rs.rfield_id
         ORDER BY ds.year, ds.period, rf.label_en, rs.label_en"""):
        key = r["period"]
        bucket = agg["subfields_by_period"].setdefault(key, {
            "year": r["year"], "period": r["period"],
            "signature": None, "rows": [],
        })
        bucket["signature"] = r["schema_year_signature"]
        bucket["rows"].append({
            "field": r["field"], "subfield": r["subfield"],
            "total": r["total_personnel"], "national_doctor": r["national_doctor"], "phd": r["phd"],
            "extras": json.loads(r["extras_json"]) if r["extras_json"] else {},
        })
    # Mirror into subfields_by_year so the old JS code path still finds rows.
    # When two periods exist for the same year (e.g. 1971), the LATER snapshot
    # wins in the legacy view.  The new `subfields_by_period` payload is the
    # authoritative one.
    for key in sorted(agg["subfields_by_period"].keys()):
        b = agg["subfields_by_period"][key]
        agg["subfields_by_year"][str(b["year"])] = {
            "signature": b["signature"], "rows": b["rows"], "period": b["period"],
        }

    # ── Annual yearly counts of distinct researchers (individual side) ──
    yearly = query_all(conn, """
        SELECT year, COUNT(DISTINCT person_id) AS researchers
          FROM person_year_observation GROUP BY year ORDER BY year""")

    # ── Per-year × per-position counts derived from individual data ───────
    # The aggregate `demo_position` table only contains 4 broad buckets and
    # is missing trainees / engineers / leadership for most years.  The
    # individual-level `cemi_appointment` table records every appointment
    # (including Trainee, Graduate Student, Engineer, etc.), and
    # `person_year_observation` carries the per-year degree, so we can
    # compute an accurate yearly distribution — including trainee — plus
    # a per-position degree breakdown.
    print("  · computing per-year x per-position breakdown (incl. trainees) ...")
    pos_meta_rows = list(conn.execute(
        "SELECT position_id, label_en, label_ru, is_acting, rank_order, tier FROM position_rank"))
    pos_meta = {r[0]: {"position_id": r[0], "label_en": r[1], "label_ru": r[2],
                         "is_acting": r[3], "rank_order": r[4], "tier": r[5]}
                for r in pos_meta_rows}

    # 1) For each (person, year), pick the final position held that year — the
    # appointment with the highest within_year_order; ties broken by rank_order.
    # This avoids double-counting a person who was promoted within the same year.
    last_pos: dict = {}
    for pid, yr, pos_id, wyo, ro in conn.execute(
        "SELECT person_id, year, position_id, within_year_order, rank_order FROM cemi_appointment"):
        key = (pid, yr)
        cand = (wyo or 0, ro or 0, pos_id)
        prev = last_pos.get(key)
        if prev is None or cand > prev:
            last_pos[key] = cand

    # 2) Build per-(person, year) degree set
    deg_label = {did: lbl for did, lbl in conn.execute("SELECT degree_id, label FROM degree")}
    deg_by_py: dict = {}
    for pid, yr, did in conn.execute(
        "SELECT person_id, year, degree_id FROM person_year_observation"
        " WHERE year IS NOT NULL AND degree_id IS NOT NULL"):
        deg_by_py.setdefault((pid, yr), set()).add(deg_label.get(did, ""))

    # 3) Group and count → one row per (year, position)
    pos_buckets: dict = {}
    for (pid, yr), (_, _, pos_id) in last_pos.items():
        b = pos_buckets.setdefault((yr, pos_id),
                                   {"total": 0, "phd": 0, "national_doctor": 0, "no_degree": 0})
        b["total"] += 1
        degs = deg_by_py.get((pid, yr), set())
        if "National Doctor" in degs:
            b["national_doctor"] += 1
        elif "Ph.D." in degs:
            b["phd"] += 1
        else:
            b["no_degree"] += 1

    positions_individual = []
    for (yr, pos_id), cnt in sorted(
            pos_buckets.items(),
            key=lambda x: (x[0][0], pos_meta.get(x[0][1], {}).get("rank_order") or 0)):
        pm = pos_meta.get(pos_id, {})
        positions_individual.append({
            "year":        yr,
            "position":    pm.get("label_en"),
            "position_ru": pm.get("label_ru"),
            "tier":        pm.get("tier"),
            "rank_order":  pm.get("rank_order"),
            "is_acting":   pm.get("is_acting"),
            "total":           cnt["total"],
            "national_doctor": cnt["national_doctor"],
            "phd":             cnt["phd"],
            "no_degree":       cnt["no_degree"],
        })
    agg["positions_individual"] = positions_individual

    # ── Headline summary for the badges ─────────────────────────────────
    pt = agg["personnel_totals"]
    peak = max(pt, key=lambda r: (r["total_scientists"] or 0)) if pt else {}
    deg_with = [r for r in agg["academic_degrees"] if r["national_doctor"] is not None]
    summary = {
        "first_year": min((p["first_year"] for p in persons if p["first_year"]), default=1962),
        "last_year":  max((p["last_year"]  for p in persons if p["last_year"]),  default=1988),
        "total_persons": len(persons),
        "peak_year": peak.get("year"),
        "peak_scientists": peak.get("total_scientists"),
        "n_appointments": len(appts),
        "n_dual_posts":   len(duals),
        "n_pre_cemi":     len(pres),
        "n_transfers":    len(trans),
        "n_dismissals":   len(diss),
        "n_party":        len(party_ev),
        "n_phd":          len(phd_ev),
        "latest_natdoc":  deg_with[-1]["national_doctor"] if deg_with else None,
        "latest_phd":     deg_with[-1]["phd"] if deg_with else None,
        "latest_deg_year": deg_with[-1]["year"] if deg_with else None,
        "n_subfields_distinct": len({(b["year"], r["subfield"]) for b in agg["subfields_by_period"].values() for r in b["rows"]}) if agg.get("subfields_by_period") else len({(y, r["subfield"]) for y, b in agg["subfields_by_year"].items() for r in b["rows"]}),
    }

    return {"persons": persons, "events": events, "vocab": vocab,
            "yearly": yearly, "agg": agg, "summary": summary,
            "year_phase": year_phase,
            "inst_primary_count": inst_primary_count,
            "cemi_inst_id": cemi_inst_id}


# ──────────────────────────────────────────────────────────────────────────────
#  §2  HTML template
# ──────────────────────────────────────────────────────────────────────────────

HTML = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width,initial-scale=1.0"/>
<title>CEMI Career Database</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.3/dist/chart.umd.min.js"></script>
<style>
:root{
  --bg:#0e1219; --panel:#181c27; --panel2:#1f2434; --line:#2c3145;
  --ink:#dde1ee; --dim:#878fa8; --accent:#7a9fd4; --accent2:#e3a73e; --accent3:#65bf86; --warn:#cf7373;
  --r:10px;
}
*{box-sizing:border-box;margin:0;padding:0;}
body{background:var(--bg);color:var(--ink);font:14px/1.55 'Segoe UI','Malgun Gothic',system-ui,sans-serif;}
header{background:linear-gradient(120deg,#10141f,#1c2236);border-bottom:1px solid var(--line);padding:24px 36px;}
header h1{font-size:21px;font-weight:700;color:#eef0f8;}
header h1 small{color:var(--accent2);font-weight:600;font-size:13px;margin-left:10px;}
header .sub{color:var(--dim);font-size:13px;margin-top:4px;}
.badges{display:flex;gap:8px;margin-top:12px;flex-wrap:wrap;}
.badges .b{background:var(--panel2);border:1px solid var(--line);border-radius:18px;padding:4px 12px;font-size:12px;color:var(--dim);}
.badges .b b{color:var(--accent2);font-weight:600;}

nav{background:var(--panel);border-bottom:1px solid var(--line);padding:0 36px;display:flex;gap:0;overflow-x:auto;}
nav button{background:none;border:none;color:var(--dim);padding:14px 16px;font-size:13px;cursor:pointer;border-bottom:2px solid transparent;white-space:nowrap;font-family:inherit;}
nav button:hover{color:var(--ink);}
nav button.on{color:var(--accent);border-bottom-color:var(--accent);font-weight:600;}

main{max-width:1380px;margin:0 auto;padding:30px;}
.tab{display:none;}
.tab.on{display:block;}

.row{display:grid;grid-template-columns:1fr 1fr;gap:18px;}
.col-2{grid-column:1 / span 2;}
.card{background:var(--panel);border:1px solid var(--line);border-radius:var(--r);padding:20px;margin-bottom:18px;}
.card h3{font-size:14px;font-weight:600;color:#eef0f8;margin-bottom:14px;}

.k-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:14px;margin-bottom:24px;}
.k-card{background:var(--panel);border:1px solid var(--line);border-radius:var(--r);padding:18px;}
.k-card .lbl{font-size:11px;color:var(--dim);text-transform:uppercase;letter-spacing:.6px;}
.k-card .val{font-size:26px;font-weight:700;color:var(--accent2);margin-top:3px;}
.k-card .sub{font-size:11px;color:var(--dim);margin-top:2px;}

.chart-wrap canvas{max-height:340px;}
.chart-tall canvas{max-height:480px;}

table{width:100%;border-collapse:collapse;font-size:13px;}
.tbl-wrap{max-height:520px;overflow:auto;border-radius:8px;}
thead th{background:var(--panel2);color:var(--dim);font-size:11px;text-transform:uppercase;letter-spacing:.5px;padding:9px 12px;text-align:left;position:sticky;top:0;border-bottom:1px solid var(--line);}
tbody tr{border-bottom:1px solid var(--line);}
tbody tr:hover{background:var(--panel2);}
tbody td{padding:8px 12px;}
tbody td.num{text-align:right;color:var(--accent2);font-variant-numeric:tabular-nums;}
tbody tr.click{cursor:pointer;}
tbody tr.sel{background:rgba(227,167,62,.12)!important;}

.pill{display:inline-block;padding:1px 8px;border-radius:10px;font-size:11px;background:rgba(122,159,212,.18);color:var(--accent);}
.pill-a{background:rgba(227,167,62,.18);color:var(--accent2);}
.pill-act{background:rgba(207,115,115,.2);color:var(--warn);}

.filter{display:grid;grid-template-columns:repeat(auto-fit,minmax(200px,1fr));gap:10px;margin-bottom:14px;}
.filter input,.filter select{background:var(--panel2);border:1px solid var(--line);color:var(--ink);border-radius:6px;padding:7px 10px;font-size:13px;font-family:inherit;}
.filter button{background:var(--panel2);border:1px solid var(--line);color:var(--ink);border-radius:6px;padding:7px 10px;font-size:13px;cursor:pointer;font-family:inherit;}

.results-banner{background:rgba(227,167,62,.08);border:1px solid rgba(227,167,62,.25);border-radius:8px;padding:8px 14px;margin-bottom:12px;font-size:13px;color:var(--dim);}
.results-banner b{color:var(--accent2);font-weight:600;}

/* Grand-total row used in the Positions data table (reference layout) */
tbody tr.grand-total-row td{background:rgba(227,167,62,.09);color:var(--accent2);font-weight:600;border-top:2px solid rgba(227,167,62,.4);}
tbody tr.grand-total-row:hover td{background:rgba(227,167,62,.13);}

.profile-head{background:linear-gradient(120deg,rgba(122,159,212,.12),rgba(227,167,62,.06));border:1px solid var(--line);border-radius:var(--r);padding:20px 22px;margin-bottom:18px;}
.profile-head .nm{font-size:24px;font-weight:700;color:#eef0f8;}
.profile-head .meta{display:flex;flex-wrap:wrap;gap:12px;font-size:12px;color:var(--dim);margin-top:5px;}
.profile-head .meta b{color:var(--accent2);}
.facts{display:grid;grid-template-columns:repeat(auto-fit,minmax(190px,1fr));gap:10px;margin-top:14px;}
.facts .f{background:var(--panel2);border-radius:6px;padding:8px 12px;}
.facts .f .l{font-size:10px;color:var(--dim);text-transform:uppercase;letter-spacing:.5px;}
.facts .f .v{font-size:13px;color:var(--ink);margin-top:1px;word-break:break-word;}

.tl{position:relative;padding-left:24px;border-left:2px solid var(--line);}
.ev{position:relative;padding:8px 0 8px 16px;margin-bottom:4px;}
.ev::before{content:'';position:absolute;left:-31px;top:14px;width:10px;height:10px;border-radius:50%;background:var(--panel);border:2px solid var(--accent);}
.ev.appt::before{border-color:var(--accent);}
.ev.appt-acting::before{border-color:var(--accent);}
.ev.dual::before{border-color:#bf80d8;}
.ev.pre::before{border-color:var(--dim);background:var(--dim);}
.ev.start::before{border-color:var(--accent3);background:var(--accent3);}
.ev.position-change::before{border-color:var(--accent);}
.ev.transfer::before{border-color:var(--accent2);}
.ev.dismissal::before{border-color:var(--warn);background:var(--warn);}
.ev.party::before{border-color:#bf6c82;}
.ev.phd::before{border-color:#6cbfbf;background:#6cbfbf;}
.ev.title::before{border-color:#ffd700;background:#ffd700;}
.ev .yr{display:inline-block;font-weight:700;color:var(--accent2);font-size:13px;margin-right:8px;min-width:90px;}
.ev .ord{display:inline-block;font-size:10px;color:var(--dim);margin-right:6px;}
.ev .kind{display:inline-block;font-size:10px;text-transform:uppercase;letter-spacing:.5px;color:var(--dim);margin-right:6px;padding:1px 7px;background:var(--panel2);border-radius:8px;}
.ev .txt b{color:#eef0f8;}
.ev .txt span{color:var(--dim);font-size:12px;margin-left:4px;}

/* ── Career-timeline marker legend ── */
.timeline-legend{display:flex;flex-wrap:wrap;gap:14px;font-size:11px;color:var(--dim);
                  margin:2px 0 12px;padding:8px 12px;background:var(--panel2);
                  border:1px solid var(--line);border-radius:6px;}
.timeline-legend .lg{display:inline-flex;align-items:center;gap:6px;}
.timeline-legend .dot{display:inline-block;width:10px;height:10px;border-radius:50%;
                       background:var(--panel);border:2px solid var(--line);}
.timeline-legend .dot.start    {background:var(--accent3);border-color:var(--accent3);}
.timeline-legend .dot.appt     {background:var(--panel); border-color:var(--accent);}
.timeline-legend .dot.dismissal{background:var(--warn);  border-color:var(--warn);}
.timeline-legend .dot.transfer {background:var(--panel); border-color:var(--accent2);}

.section{margin-top:18px;}
.section h4{font-size:13px;font-weight:600;color:#eef0f8;margin-bottom:8px;padding-bottom:5px;border-bottom:1px solid var(--line);display:flex;align-items:center;gap:8px;}
.section h4 .n{background:var(--panel2);color:var(--accent);font-size:11px;font-weight:500;padding:1px 8px;border-radius:9px;}

.empty{padding:18px;text-align:center;color:var(--dim);font-size:13px;font-style:italic;}

footer{text-align:center;padding:20px;font-size:12px;color:var(--dim);border-top:1px solid var(--line);margin-top:18px;}

.tag-row{display:flex;flex-wrap:wrap;gap:5px;margin-top:6px;}
.tag-row .tag{background:var(--panel2);border:1px solid var(--line);border-radius:8px;padding:1px 7px;font-size:11px;color:var(--dim);}
.tag-row .tag b{color:var(--ink);font-weight:500;}

/* ── Institutional phase badges ── */
.phase-badge{display:inline-block;padding:1px 8px;border-radius:9px;font-size:10px;font-weight:600;text-transform:uppercase;letter-spacing:.4px;margin-left:6px;vertical-align:middle;}
.phase-cemi{background:rgba(122,159,212,.16);color:var(--accent);border:1px solid rgba(122,159,212,.4);}
.phase-pred{background:rgba(227,167,62,.16);color:var(--accent2);border:1px solid rgba(227,167,62,.45);}
.phase-pre{background:rgba(135,143,168,.16);color:var(--dim);border:1px solid rgba(135,143,168,.4);}
.phase-info-banner{background:rgba(227,167,62,.07);border:1px solid rgba(227,167,62,.3);border-radius:6px;padding:10px 14px;margin-bottom:14px;font-size:12px;color:var(--ink);line-height:1.55;}
.phase-info-banner b{color:var(--accent2);}
.phase-info-banner code{background:var(--panel2);padding:1px 5px;border-radius:3px;font-size:11px;}

/* ── Pagination controls (Career Search) ── */
.pagination{display:flex;align-items:center;justify-content:center;gap:14px;padding:12px 8px 4px;}
.pagination button{background:var(--panel2);border:1px solid var(--line);color:var(--ink);
                   border-radius:6px;padding:6px 14px;cursor:pointer;font-family:inherit;font-size:13px;}
.pagination button:hover:not(:disabled){background:var(--panel);border-color:var(--accent);}
.pagination button:disabled{opacity:.35;cursor:not-allowed;}
.pagination .pageinfo{font-size:12px;color:var(--dim);font-variant-numeric:tabular-nums;}
.pagination .pageinfo b{color:var(--accent2);}

/* ── New side-panel cards: Pre-CEMI, Party, Dual ── */
.info-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(290px,1fr));gap:10px;margin-top:6px;}
.info-item{background:var(--panel2);border:1px solid var(--line);border-radius:7px;padding:10px 12px;}
.info-item .info-year{font-weight:700;color:var(--accent2);font-size:13px;margin-bottom:4px;font-variant-numeric:tabular-nums;}
.info-item .info-title{font-weight:600;color:#eef0f8;font-size:13px;}
.info-item .info-sub{font-size:12px;color:var(--ink);margin-top:3px;}
.info-item .info-meta{font-size:11px;color:var(--dim);margin-top:5px;font-style:italic;}
</style>
</head>
<body>
<header>
  <h1>CEMI Career Database <small>v3 · career-search build</small></h1>
  <p class="sub">Personnel, demographic and career-event data for the Central Economic Mathematical Institute (ARAN F.1959 / Op.1)</p>
  <div class="badges">
    <div class="b">Years <b>__FY__–__LY__</b></div>
    <div class="b">Persons <b>__NP__</b></div>
    <div class="b">Peak scientists <b>__PEAK__</b> (<b>__PYR__</b>)</div>
    <div class="b">Appointments <b>__APP__</b></div>
    <div class="b">Dual posts <b>__DUAL__</b></div>
    <div class="b">Pre-CEMI roles <b>__PRE__</b></div>
    <div class="b">Transfers <b>__TR__</b></div>
    <div class="b">Dismissals <b>__DIS__</b></div>
    <div class="b">Party events <b>__PARTY__</b></div>
    <div class="b">PhD defences <b>__PHD__</b></div>
  </div>
</header>

<nav>
  <button class="on" onclick="show('overview',event)">Overview</button>
  <button onclick="show('search',event)">Career search</button>
  <button onclick="show('institutions',event)">Institutions</button>
  <button onclick="show('positions',event)">Positions</button>
  <button onclick="show('personnel',event)">Personnel growth</button>
  <button onclick="show('degrees',event)">Degrees &amp; titles</button>
  <button onclick="show('nationality',event)">Nationality</button>
  <button onclick="show('age',event)">Age distribution</button>
  <button onclick="show('research',event)">Research Fields</button>
  <button onclick="show('partytrainees',event)">Party &amp; trainees</button>
  <button onclick="show('provenance',event)">Provenance</button>
  <button onclick="show('glossary',event)">Glossary</button>
</nav>

<main>

<div id="t-overview" class="tab on">
  <div class="k-grid" id="kpis"></div>
  <div class="row">
    <div class="card"><h3>Active researchers per year (individual records)</h3><div class="chart-wrap"><canvas id="ov-active"></canvas></div></div>
    <div class="card"><h3>Total scientists &amp; all staff (aggregate)</h3><div class="chart-wrap"><canvas id="ov-total"></canvas></div></div>
    <div class="card"><h3>Career events by type</h3><div class="chart-wrap"><canvas id="ov-events"></canvas></div></div>
    <div class="card"><h3>Position tier distribution (aggregate, latest year)</h3><div class="chart-wrap"><canvas id="ov-tier"></canvas></div></div>
  </div>
</div>

<div id="t-search" class="tab">
  <div class="phase-info-banner">
    <b>Two independent filter axes.</b><br>
    <b>(1) Phase</b> &mdash; which CEMI-side institute existed at a given year:
    <span class="phase-badge phase-cemi">CEMI</span> for 1963 and later (the Central Economic Mathematical Institute proper) and
    <span class="phase-badge phase-pred">Predecessor Lab</span> for 1958&ndash;1962 (the <i>Laboratory of Mathematical Methods Applied to Economic Research and Planning</i> out of which CEMI was reorganised in 1963).
    Use the <code>Phase</code> filter to scope the search to one of these two phases.<br>
    <b>(2) Role / relation to CEMI</b> &mdash; what kind of institutional record the person carries:
    <span class="phase-badge phase-pre">Pre-CEMI</span> previous-service institution (column <code>pre_cemi_role</code>);
    <span class="phase-badge phase-pred">Dual</span> part-time / dual appointment (column <code>dual_position</code>);
    <span class="phase-badge phase-cemi">Transferred</span> transferred-to-institution record (column <code>transfer_event</code>).
    Every person is by definition a CEMI scientist, so &ldquo;CEMI service&rdquo; is universal and not exposed as a filter.
    Use the <code>Role</code> filter to restrict to persons who have at least one record of the selected type.<br>
    Pre-1958 years (e.g. party-membership year = 1940, BA graduation date) belong to neither axis &mdash; they are <i>life-event</i> dates attached to a person and are searchable via the <code>Active from / to year</code> inputs.
  </div>
  <div class="card">
    <h3>Career search · every field is filterable</h3>
    <div class="filter">
      <input id="f-name" placeholder="Name (partial match)" oninput="search()"/>
      <select id="f-pos" onchange="search()"><option value="">— Any CEMI position —</option></select>
      <select id="f-inst" onchange="search()"><option value="">— Any institution (pre / dual / transfer / primary) —</option></select>
      <select id="f-deg" onchange="search()"><option value="">— Any degree —</option></select>
      <select id="f-eth" onchange="search()"><option value="">— Any ethnicity —</option></select>
      <select id="f-spec" onchange="search()"><option value="">— Any specialty —</option></select>
      <select id="f-fld" onchange="search()"><option value="">— Any PhD field —</option></select>
      <select id="f-dept" onchange="search()"><option value="">— Any department —</option></select>
      <select id="f-title" onchange="search()"><option value="">— Any academic title —</option></select>
      <select id="f-party" onchange="search()"><option value="">— Any party affiliation —</option></select>
      <select id="f-school" onchange="search()"><option value="">— Any school (BA or grad) —</option></select>
      <select id="f-phase" onchange="search()">
        <option value="">— Any phase —</option>
        <option value="CEMI">CEMI (1963+)</option>
        <option value="Predecessor Lab">Predecessor Lab (1958–1962)</option>
      </select>
      <select id="f-role" onchange="search()">
        <option value="">— Any role / relation —</option>
        <option value="pre">Pre-CEMI service (previous-service institution)</option>
        <option value="dual">Dual / part-time appointment</option>
        <option value="transfer">Transferred-to institution (transfer record)</option>
      </select>
      <input id="f-yf" type="number" placeholder="Active from year" min="1900" max="1990" oninput="search()"/>
      <input id="f-yt" type="number" placeholder="Active to year"   min="1962" max="1990" oninput="search()"/>
      <button onclick="reset()">Clear filters</button>
    </div>
    <div id="banner" class="results-banner" style="display:none"></div>
    <div class="tbl-wrap" id="results" style="max-height:380px"></div>
  </div>

  <div class="card" id="profile-card" style="display:none">
    <div id="profile"></div>
  </div>
</div>

<div id="t-institutions" class="tab">
  <div class="card"><h3>Institutions ranked by aggregate person-years touching CEMI</h3>
    <div class="filter">
      <input id="i-search" placeholder="Filter institutions (name)" oninput="renderInstitutions()"/>
      <select id="i-kind" onchange="renderInstitutions()" title="How does this institution relate to CEMI?">
        <option value="">— Any CEMI relationship —</option>
        <option value="Pre-CEMI">Pre-CEMI (previous-service institution)</option>
        <option value="Dual">Dual / part-time appointment</option>
        <option value="Transfer">Transfer-out destination</option>
        <option value="Primary-CEMI">Primary = CEMI (full-time CEMI staff)</option>
        <option value="Primary-Ext">Primary = external institution (CEMI = part-time)</option>
      </select>
      <select id="i-kindtax" onchange="renderInstitutions()" title="Intrinsic institution kind (Academy / VUZ / Branch / Enterprise / Other)">
        <option value="">— Any institution kind —</option>
      </select>
    </div>
    <div class="tbl-wrap" id="i-table"></div>
  </div>
</div>

<!-- Positions — reference-style layout (two stacked horizontal-bar charts,
     grand-total banner, single bilingual data table with grand-total row). -->
<div id="t-positions" class="tab">
  <div class="card">
    <h3>Yearly position breakdown — every position incl. Trainee (individual-level data)</h3>
    <p style="font-size:12px;color:var(--dim);margin-bottom:10px;">
      Counts derive from <code>cemi_appointment</code>: for each (person, year) we
      keep the final substantive position held that year, then group by
      position rank.  This captures every observed position — Trainee, Graduate
      Student, Engineer, Senior Engineer, Junior/Senior Researcher, Heads,
      Leadership — and is more complete than the 4-bucket aggregate sheet.
      The degree chart maps each person's degree in
      <code>person_year_observation</code> (Ph.D. / National Doctor / not recorded)
      onto their position for the selected year.
    </p>
    <div class="filter">
      <select id="p-year" onchange="renderPositionsTab()"></select>
    </div>
    <div class="row">
      <div class="card"><h3>Position distribution — total personnel</h3><div class="chart-tall"><canvas id="pos-chart-total"></canvas></div></div>
      <div class="card"><h3>Degree ratio within each position (% of position total)</h3><div class="chart-tall"><canvas id="pos-chart-deg"></canvas></div></div>
    </div>
    <div id="pos-banner" class="results-banner" style="display:none"></div>
    <div class="tbl-wrap" id="pos-data-table"></div>
  </div>
  <div class="card"><h3>CEMI position rank ladder</h3>
    <p style="font-size:12px;color:var(--dim);margin-bottom:10px;">Rank values are dimensionless ordinals chosen so that an "Acting X" sits just below the substantive "X". This is the ladder that drives within-year promotion ordering.</p>
    <div class="tbl-wrap" id="pos-ladder-table"></div>
  </div>
</div>

<div id="t-personnel" class="tab">
  <div class="row">
    <div class="card"><h3>Total scientists &amp; all staff</h3><div class="chart-tall"><canvas id="pn-staff"></canvas></div></div>
    <div class="card"><h3>Women breakdown over time</h3><div class="chart-tall"><canvas id="pn-women"></canvas></div></div>
  </div>
  <div class="card"><h3>Personnel Totals — every column from cemi_translation_db.xlsx · Personnel Totals (9 columns)</h3>
    <div class="tbl-wrap" id="pn-table"></div>
  </div>
</div>

<div id="t-degrees" class="tab">
  <div class="row">
    <div class="card"><h3>National Doctor &amp; Ph.D. counts over time</h3><div class="chart-tall"><canvas id="dg-deg"></canvas></div></div>
    <div class="card"><h3>Academic titles (Professors / Docents / SNS / Academicians)</h3><div class="chart-tall"><canvas id="dg-titles"></canvas></div></div>
  </div>
  <div class="card"><h3>% with degree</h3><div class="chart-wrap"><canvas id="dg-pct"></canvas></div></div>
  <div class="card"><h3>Academic Degrees — every column (9 columns)</h3>
    <div class="tbl-wrap" id="dg-table"></div>
  </div>
</div>

<div id="t-nationality" class="tab">
  <div class="card">
    <div class="filter">
      <select id="nat-year" onchange="renderNationality()"></select>
    </div>
    <div class="row">
      <div class="card"><h3>Distribution (donut)</h3><div class="chart-wrap"><canvas id="nat-donut"></canvas></div></div>
      <div class="card"><h3>Trend over time — top 6 by total</h3><div class="chart-wrap"><canvas id="nat-trend"></canvas></div></div>
    </div>
    <h3 style="margin-top:14px;">Nationality table (Russian + English) for the selected year</h3>
    <div class="tbl-wrap" id="nat-table"></div>
  </div>
</div>

<div id="t-age" class="tab">
  <div class="card">
    <div class="filter">
      <select id="age-year" onchange="renderAge()"></select>
    </div>
    <div class="row">
      <div class="card"><h3>Total per bracket</h3><div class="chart-wrap"><canvas id="age-total"></canvas></div></div>
      <div class="card"><h3>National Doctor &amp; Ph.D. within brackets</h3><div class="chart-wrap"><canvas id="age-deg"></canvas></div></div>
    </div>
    <div class="card" id="age-titles-card" style="margin-top:14px;">
      <h3>Academic title / rank breakdown per age bracket (stacked)</h3>
      <p style="font-size:11px;color:var(--dim);margin-bottom:8px;">
        Counts of persons in each age bracket holding the listed academic
        title or research-staff rank, taken directly from
        <code>cemi_translation_db.xlsx · Age Distribution</code>.  Years that
        do not record these columns leave the chart blank.
      </p>
      <div class="chart-tall"><canvas id="age-titles"></canvas></div>
    </div>
    <h3 style="margin-top:14px;">Age bracket table (Russian + English) for the selected year</h3>
    <div class="tbl-wrap" id="age-table"></div>
  </div>
</div>

<!-- ── Research Fields (merged: top-level fields as parent, subfields as child) ── -->
<div id="t-research" class="tab">
  <div class="card">
    <p style="font-size:12px;color:var(--dim);margin-bottom:10px;">
      Pick a year — every Research Field row is rendered together with its
      child Subfield rows from the matching sheet of
      <code>research_field_subfield.xlsx</code>.  The parent's totals equal
      the sum of its own children in that year.  Sheet names are now plain
      <code>YYYY</code> (e.g. <b>1965</b>, <b>1971</b>) so each year has a
      single snapshot.
    </p>
    <div class="filter">
      <label style="font-size:11px;color:var(--dim);text-transform:uppercase;letter-spacing:.5px;align-self:center;margin-right:2px;">Year</label>
      <select id="rf-year" onchange="renderResearch()"></select>
      <!-- Hidden in current workbook (period == year); preserved purely as a
           defensive fallback if a future workbook re-introduces YYYY.MM. -->
      <select id="rf-snapshot" onchange="renderResearch()" style="display:none;"></select>
      <select id="rf-field"  onchange="renderResearch()"><option value="">— All Research Fields —</option></select>
    </div>
    <p id="rf-sig" style="font-size:11px;color:var(--dim);font-family:monospace;margin-bottom:10px;"></p>

    <!-- Chart 1: all Research Fields for the selected year -->
    <h4 style="font-size:13px;font-weight:600;color:#eef0f8;margin:6px 0 8px;">All Research Fields</h4>
    <div class="chart-tall"><canvas id="rf-chart"></canvas></div>

    <!-- Chart 2: all Subfields for the selected year — hidden when none -->
    <div id="rf-sub-chart-wrap" style="margin-top:14px;">
      <h4 style="font-size:13px;font-weight:600;color:#eef0f8;margin:6px 0 8px;">All Subfields</h4>
      <div class="chart-tall"><canvas id="rf-sub-chart"></canvas></div>
    </div>

    <div id="rf-blocks" style="margin-top:14px;"></div>
  </div>

  <div class="card">
    <h3>Top 10 Subfields — trend over time</h3>
    <p style="color:var(--dim);font-size:12px;margin-bottom:10px;">
      The ten Subfields with the highest cumulative <i>Total Personnel</i>
      across every year, plotted year by year.  Each year contributes a
      single value, taken from its <code>YYYY</code> sheet in
      <code>research_field_subfield.xlsx</code>.
    </p>
    <div class="chart-tall"><canvas id="rf-trend"></canvas></div>
  </div>
</div>

<div id="t-partytrainees" class="tab">
  <div class="row">
    <div class="card"><h3>CPSU and Komsomol counts</h3><div class="chart-wrap"><canvas id="pt-counts"></canvas></div></div>
    <div class="card"><h3>Membership rate (%)</h3><div class="chart-wrap"><canvas id="pt-pct"></canvas></div></div>
  </div>
  <div class="card"><h3>Party Membership table (6 columns)</h3>
    <div class="tbl-wrap" id="pt-party-table"></div>
  </div>
  <div class="card"><h3>Research Trainees table (3 columns)</h3>
    <div class="tbl-wrap" id="pt-trainee-table"></div>
  </div>
</div>

<div id="t-provenance" class="tab">
  <div class="card">
    <h3>Provenance — every sheet feeding the DB, archival + intermediate</h3>
    <p style="color:var(--dim);font-size:12px;margin-bottom:10px;">
      Each row records one Excel sheet that contributes data to <code>cemi_career.db</code>.
      The 147 <b>primary archival source</b> rows come from the ARAN F.1959 archival files
      (Russian state archive originals) listed in the <i>Source Files Index</i> sheet of
      <code>cemi_translation_db.xlsx</code>.  The remaining rows describe every sheet of the
      three input workbooks the DB builder consumes —
      <code>data.xlsx</code> (personnel), <code>cemi_translation_db.xlsx</code> (aggregates),
      and <code>research_field_subfield.xlsx</code> (yearly subfield sheets) — with the DB
      table each sheet feeds into shown in the <i>Target table</i> column.
    </p>
    <div class="filter">
      <input id="prv-search" placeholder="Filter (text in any column)" oninput="renderProvenance()"/>
      <select id="prv-role" onchange="renderProvenance()"><option value="">— Any role —</option></select>
      <select id="prv-file" onchange="renderProvenance()"><option value="">— Any source file —</option></select>
      <select id="prv-inst" onchange="renderProvenance()"><option value="">— Any institution —</option></select>
      <select id="prv-target" onchange="renderProvenance()"><option value="">— Any target table —</option></select>
    </div>
    <div class="tbl-wrap" id="prv-table"></div>
  </div>
</div>

<!-- (Subfields tab was merged into Research Fields; the JS bucket
     `subfields_by_period` remains available for code that still references it.) -->

<div id="t-glossary" class="tab">
  <div class="card">
    <div class="filter">
      <input id="g-search" placeholder="Search Russian or English term" oninput="renderGlossary()"/>
      <select id="g-cat" onchange="renderGlossary()"><option value="">— Any category —</option></select>
    </div>
    <div id="g-count" style="color:#878fa8;font-size:12px;padding:0 12px 6px"></div>
    <div class="tbl-wrap" id="g-table"></div>
  </div>
</div>

</main>
<footer>CEMI Career Database · ARAN F.1959 / Op.1 · v3 build (Acting → Substantive within-year ordering)</footer>

<script>
const D = __DATA__;
const VOCAB = D.vocab;
const INST_PRIMARY_COUNT = D.inst_primary_count || {};
const CEMI_INST_ID = D.cemi_inst_id;
const PERSONS = D.persons;
const EV = D.events;
const AGG = D.agg;
const SUM = D.summary;
const YEARLY = D.yearly;
const YEAR_PHASE = D.year_phase || {};

// ── Institute-phase helpers ──
function phaseFor(year){
  if (year == null) return null;
  const cached = YEAR_PHASE[String(year)];
  if (cached) return cached.phase;
  const y = +year;
  if (y >= 1963)  return 'CEMI';
  if (y >= 1958)  return 'Predecessor Lab';
  return 'Pre-Institute';
}
function phasePill(year){
  const p = phaseFor(year);
  if (p == null) return '';
  if (p === 'CEMI')             return '<span class="phase-badge phase-cemi">CEMI</span>';
  if (p === 'Predecessor Lab')  return '<span class="phase-badge phase-pred">Predecessor</span>';
  return '<span class="phase-badge phase-pre">Pre-Inst.</span>';
}
// <select> options can't carry HTML; append a plain-text phase suffix so users
// see "1962 · Predecessor Lab" right in the dropdown.
function yearOptionLabel(year){
  const p = phaseFor(year);
  if (p === 'Predecessor Lab')  return year + ' · Predecessor Lab';
  if (p === 'Pre-Institute')    return year + ' · Pre-Institute';
  return String(year);
}

// ── Lookups for fast O(1) label resolution ──
const VL = {};
for (const k in VOCAB) {
  VL[k] = {};
  for (const v of VOCAB[k]) VL[k][v.id] = v;
}

const PAL = ['#7a9fd4','#e3a73e','#65bf86','#cf7373','#bf80d8','#65b6c4','#c4a165','#a8c465','#c465a8','#65a3c4','#dab06f','#82c982','#6f8fc4','#c4828f','#82c4d4'];
const c = (i,a=1)=>{const h=PAL[i%PAL.length];const r=parseInt(h.slice(1,3),16),g=parseInt(h.slice(3,5),16),b=parseInt(h.slice(5,7),16);return`rgba(${r},${g},${b},${a})`;};
const BASE = {plugins:{legend:{labels:{color:'#878fa8',font:{size:11}}}},scales:{x:{ticks:{color:'#878fa8'},grid:{color:'#2c3145'}},y:{ticks:{color:'#878fa8'},grid:{color:'#2c3145'}}}};
function co(o={}){return md(JSON.parse(JSON.stringify(BASE)),o);}
function md(t,s){for(const k in s){if(s[k]&&typeof s[k]==='object'&&!Array.isArray(s[k])){if(!t[k])t[k]={};md(t[k],s[k]);}else t[k]=s[k];}return t;}

const CH={};
function mk(id,type,data,opts={}){const el=document.getElementById(id);if(!el)return;if(CH[id])CH[id].destroy();CH[id]=new Chart(el,{type,data,options:opts});}

const fmt = x => x==null ? '—' : (Number.isInteger(x) ? x.toLocaleString() : (+x).toLocaleString(undefined,{maximumFractionDigits:1}));
const esc = s => s==null ? '—' : String(s).replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));

// ── Tab switcher ───────────────────────────────────────────────
const INIT={}, DONE={};
function show(name, ev){
  document.querySelectorAll('.tab').forEach(e=>e.classList.remove('on'));
  document.querySelectorAll('nav button').forEach(b=>b.classList.remove('on'));
  document.getElementById('t-'+name).classList.add('on');
  ev.currentTarget.classList.add('on');
  if(!DONE[name]&&INIT[name]){INIT[name]();DONE[name]=true;}
}

// ── KPI badges + overview charts ───────────────────────────────
function fillKPIs(){
  document.getElementById('kpis').innerHTML = [
    {l:'Years covered',          v:`${SUM.first_year}–${SUM.last_year}`, s:'span across all individual records'},
    {l:'Unique persons',          v:fmt(SUM.total_persons), s:'with at least one observation'},
    {l:'Peak total scientists',   v:fmt(SUM.peak_scientists), s:'year ' + (SUM.peak_year ?? '—')},
    {l:'Appointments recorded',   v:fmt(SUM.n_appointments), s:'CEMI position events'},
    {l:'Dual posts',              v:fmt(SUM.n_dual_posts), s:'concurrent external roles'},
    {l:'Pre-CEMI roles',          v:fmt(SUM.n_pre_cemi), s:'previous service'},
    {l:'Transfers out',           v:fmt(SUM.n_transfers), s:'transferred to another institution'},
    {l:'Dismissals',              v:fmt(SUM.n_dismissals), s:'recorded leave-of-service events'},
    {l:'Party events',            v:fmt(SUM.n_party), s:'CPSU / Komsomol joinings'},
    {l:'PhD defences',            v:fmt(SUM.n_phd), s:'recorded with date or field'},
    {l:'Latest National Doctors', v:fmt(SUM.latest_natdoc), s:'year ' + (SUM.latest_deg_year ?? '—')},
    {l:'Latest Ph.D.s',           v:fmt(SUM.latest_phd), s:'year ' + (SUM.latest_deg_year ?? '—')},
  ].map(k=>`<div class="k-card"><div class="lbl">${k.l}</div><div class="val">${k.v}</div><div class="sub">${k.s}</div></div>`).join('');
}

INIT.overview = function(){
  fillKPIs();
  mk('ov-active','bar',{labels:YEARLY.map(r=>r.year), datasets:[{label:'Active researchers',data:YEARLY.map(r=>r.researchers),backgroundColor:c(2,.7)}]},
     co({plugins:{legend:{display:false}}}));
  const pt=AGG.personnel_totals;
  mk('ov-total','line',{labels:pt.map(r=>r.year),datasets:[
    {label:'Total scientists',  data:pt.map(r=>r.total_scientists),  borderColor:c(0), backgroundColor:c(0,.12), tension:.3, fill:true, spanGaps:true},
    {label:'Grand total — staff',data:pt.map(r=>r.total_all_staff),  borderColor:c(1), backgroundColor:c(1,.08), tension:.3, fill:true, spanGaps:true},
  ]}, co({plugins:{legend:{position:'top'}}}));

  const counts = {
    'CEMI appointments':   SUM.n_appointments,
    'Dual external posts': SUM.n_dual_posts,
    'Pre-CEMI roles':      SUM.n_pre_cemi,
    'Transfers out':       SUM.n_transfers,
    'Dismissals':          SUM.n_dismissals,
    'Party events':        SUM.n_party,
    'PhD defences':        SUM.n_phd,
  };
  mk('ov-events','bar',{labels:Object.keys(counts),datasets:[{label:'count',data:Object.values(counts),backgroundColor:Object.keys(counts).map((_,i)=>c(i,.75))}]},
     co({indexAxis:'y',plugins:{legend:{display:false}}}));

  // tier composition for latest year
  const lastYr = Math.max(...AGG.positions.map(r=>r.year));
  const lastPos = AGG.positions.filter(r=>r.year===lastYr);
  const tierTotals = {};
  for (const r of lastPos) tierTotals[r.tier||'unknown'] = (tierTotals[r.tier||'unknown']||0) + (r.total||0);
  mk('ov-tier','doughnut',{labels:Object.keys(tierTotals),datasets:[{data:Object.values(tierTotals),backgroundColor:Object.keys(tierTotals).map((_,i)=>c(i,.85))}]},
     {plugins:{legend:{position:'right',labels:{color:'#878fa8',font:{size:11}}}}});
};
INIT.overview(); DONE.overview = true;

// ── Career search ─────────────────────────────────────────────
function fillVocabSelect(id, items, lblKey='label'){
  const sel = document.getElementById(id);
  for (const it of items) {
    const o = document.createElement('option');
    o.value = it.id;
    o.textContent = it[lblKey];
    sel.appendChild(o);
  }
}

INIT.search = function(){
  fillVocabSelect('f-pos',   VOCAB.position);
  fillVocabSelect('f-inst',  VOCAB.institution);
  fillVocabSelect('f-deg',   VOCAB.degree);
  fillVocabSelect('f-eth',   VOCAB.ethnicity);
  fillVocabSelect('f-spec',  VOCAB.specialty);
  fillVocabSelect('f-fld',   VOCAB.phd_field);
  fillVocabSelect('f-dept',  VOCAB.department);
  fillVocabSelect('f-title', VOCAB.academic_title);
  // f-party is populated with role-aware options below.
  fillVocabSelect('f-school',VOCAB.school);

  // ── Role-aware Party Affiliation dropdown ───────────────────────
  // Inspect EV.party to discover which (party, role) pairs actually
  // occur in the data.  For each party we emit:
  //   • "<party> (any)"    — matches Members + Candidates
  //   • "<party> — <role>" — one entry per observed role
  // Values are encoded as "PartyLabel" or "PartyLabel:Role" so the
  // search() filter can parse them unambiguously.
  {
    const sel = document.getElementById('f-party');
    const partyRoles = {};   // partyLabel → Set<role>
    for (const pid in EV.party) {
      for (const ev of (EV.party[pid] || [])) {
        if (!ev.party) continue;
        (partyRoles[ev.party] = partyRoles[ev.party] || new Set()).add(ev.role || '(unspecified)');
      }
    }
    const parties = Object.keys(partyRoles).sort();
    for (const party of parties) {
      const optAny = document.createElement('option');
      optAny.value = party;
      optAny.textContent = party + ' (any)';
      sel.appendChild(optAny);
      const roles = [...partyRoles[party]].sort();
      for (const role of roles) {
        const o = document.createElement('option');
        o.value = party + ':' + role;
        o.textContent = party + ' — ' + role;
        sel.appendChild(o);
      }
    }
  }

  search();
};

let SELECTED = null;

// ── Career Search pagination state ─────────────────────────────────
const SEARCH_PAGE_SIZE = 100;
let SEARCH_PAGE = 1;

// `search()` is bound to every filter input/change; resetting to page 1
// is the correct behaviour any time a filter changes.  `gotoPage(delta)`
// just navigates without resetting filters.
function search(){
  SEARCH_PAGE = 1;
  renderSearchResults();
}
function gotoPage(delta){
  SEARCH_PAGE = Math.max(1, SEARCH_PAGE + delta);
  renderSearchResults();
}

function renderSearchResults(){
  const fName = (document.getElementById('f-name').value||'').trim().toLowerCase();
  const fPos  = +document.getElementById('f-pos').value || null;
  const fInst = +document.getElementById('f-inst').value || null;
  const fDeg  = +document.getElementById('f-deg').value || null;
  const fEth  = (document.getElementById('f-eth').selectedOptions[0]||{}).text;
  const fSpec = +document.getElementById('f-spec').value || null;
  const fFld  = +document.getElementById('f-fld').value || null;
  const fDept = +document.getElementById('f-dept').value || null;
  const fTitle= +document.getElementById('f-title').value || null;
  // f-party value is now a string: "" / "CPSU" / "CPSU:Candidate" / etc.
  const fPartyRaw = document.getElementById('f-party').value || '';
  let fPartyName = null, fPartyRole = null;
  if (fPartyRaw){
    const idx = fPartyRaw.indexOf(':');
    if (idx >= 0){ fPartyName = fPartyRaw.slice(0, idx); fPartyRole = fPartyRaw.slice(idx+1); }
    else         { fPartyName = fPartyRaw; }
  }
  const fSchool=+document.getElementById('f-school').value || null;
  const fPhase = (document.getElementById('f-phase')||{value:''}).value;
  const fRole  = (document.getElementById('f-role') ||{value:''}).value;
  const yf = parseInt(document.getElementById('f-yf').value)||null;
  const yt = parseInt(document.getElementById('f-yt').value)||null;

  let rows = PERSONS;
  if (fName)  rows = rows.filter(p => (p.name||'').toLowerCase().includes(fName));
  if (fPos)   rows = rows.filter(p => p.positions.includes(fPos));
  if (fInst)  rows = rows.filter(p => p.institutions.includes(fInst));
  if (fDeg)   rows = rows.filter(p => p.degrees.includes(fDeg));
  if (fEth && fEth !== '— Any ethnicity —') rows = rows.filter(p => p.ethnicity === fEth);
  if (fSpec)  rows = rows.filter(p => p.specialties.includes(fSpec));
  if (fFld)   rows = rows.filter(p => p.fields.includes(fFld));
  if (fDept)  rows = rows.filter(p => p.departments.includes(fDept));
  if (fTitle) rows = rows.filter(p => p.titles.includes(fTitle));
  if (fPartyName) {
    // Role-aware party filter — consult EV.party rather than p.parties.
    rows = rows.filter(p => {
      const evs = EV.party[p.id] || [];
      return evs.some(e =>
        e.party === fPartyName && (fPartyRole == null || e.role === fPartyRole));
    });
  }
  if (fSchool)rows = rows.filter(p => p.schools.includes(fSchool));
  if (yf!=null) rows = rows.filter(p => (p.last_year||0) >= yf);
  if (yt!=null) rows = rows.filter(p => (p.first_year||9999) <= yt);
  if (fPhase) {
    rows = rows.filter(p => p.phase_spans && p.phase_spans[fPhase]);
  }
  if (fRole) {
    if (fRole === 'pre') {
      rows = rows.filter(p => (EV.pre_cemi[p.id] || []).length > 0);
    } else if (fRole === 'dual') {
      rows = rows.filter(p => (EV.duals[p.id] || []).length > 0);
    } else if (fRole === 'transfer') {
      rows = rows.filter(p => (EV.transfers[p.id] || []).length > 0);
    }
  }

  // Banner — total match count + active party filter description.
  const partyDesc = fPartyName
    ? ` · party filter: <b>${esc(fPartyName + (fPartyRole ? ' — ' + fPartyRole : ' (any)'))}</b>`
    : '';
  document.getElementById('banner').style.display = '';
  document.getElementById('banner').innerHTML =
    `<b>${fmt(rows.length)}</b> persons match (out of ${fmt(PERSONS.length)} total)${partyDesc}`;

  const sorted = rows.slice().sort((a,b)=>(a.surname||'').localeCompare(b.surname||''));
  const totalPages = Math.max(1, Math.ceil(sorted.length / SEARCH_PAGE_SIZE));
  if (SEARCH_PAGE > totalPages) SEARCH_PAGE = totalPages;
  const startIdx = (SEARCH_PAGE - 1) * SEARCH_PAGE_SIZE;
  const display = sorted.slice(startIdx, startIdx + SEARCH_PAGE_SIZE);

  const top = (p) => {
    const ev = (EV.appointments[p.id]||[]).slice().sort((a,b)=>(b.rank-a.rank)||(b.year-a.year));
    return ev.length ? ev[0].position : '—';
  };

  const head = '<table><thead><tr>'
    + '<th>Name</th><th>Years</th><th>Top position reached</th><th>Birth</th><th>Ethnicity</th><th>Events</th>'
    + '</tr></thead><tbody>';
  const body = display.map(p => {
    const evCount = (EV.appointments[p.id]?.length||0)
                  + (EV.duals[p.id]?.length||0)
                  + (EV.pre_cemi[p.id]?.length||0)
                  + (EV.transfers[p.id]?.length||0)
                  + (EV.dismissals[p.id]?.length||0)
                  + (EV.party[p.id]?.length||0)
                  + (EV.phd[p.id]?.length||0);
    const ys = (p.first_year && p.last_year) ? (p.first_year===p.last_year ? p.first_year : `${p.first_year}–${p.last_year}`) : '—';
    return `<tr class="click ${p.id===SELECTED?'sel':''}" onclick="openPerson(${p.id})">
              <td><b>${esc(p.name)}</b></td>
              <td>${ys}</td>
              <td>${esc(top(p))}</td>
              <td>${p.birth_year||'—'}</td>
              <td>${esc(p.ethnicity)}</td>
              <td><span class="pill">${evCount}</span></td>
            </tr>`;
  }).join('');
  let html = head + body + '</tbody></table>';

  // Pagination controls — only shown when more than one page exists.
  if (sorted.length > SEARCH_PAGE_SIZE) {
    const firstRow = startIdx + 1;
    const lastRow  = Math.min(startIdx + SEARCH_PAGE_SIZE, sorted.length);
    html += `<div class="pagination">
        <button onclick="gotoPage(-1)" ${SEARCH_PAGE === 1 ? 'disabled' : ''}>← Prev</button>
        <span class="pageinfo">Page <b>${SEARCH_PAGE}</b> / ${totalPages}
          · rows <b>${fmt(firstRow)}–${fmt(lastRow)}</b> of <b>${fmt(sorted.length)}</b></span>
        <button onclick="gotoPage(1)" ${SEARCH_PAGE === totalPages ? 'disabled' : ''}>Next →</button>
      </div>`;
  }

  document.getElementById('results').innerHTML = html;
}
function reset(){
  ['f-name','f-yf','f-yt'].forEach(id=>document.getElementById(id).value='');
  ['f-pos','f-inst','f-deg','f-eth','f-spec','f-fld','f-dept','f-title','f-party','f-school','f-phase','f-role'].forEach(id=>{
    const el=document.getElementById(id); if (el) el.value='';
  });
  search();
}

function openPerson(pid){
  SELECTED = pid;
  const p = PERSONS.find(x=>x.id===pid);
  if (!p) return;
  const card = document.getElementById('profile-card');
  card.style.display = '';

  // ── 1) Consolidated appointment ranges ─────────────────────────────
  // Preferred source: the `cemi_position_span` table (derived at DB build
  // time).  Fallback: re-fold from EV.appointments on the fly so older DBs
  // without the derived table still render correctly.
  let apptRanges = (EV.position_spans && EV.position_spans[pid]) || [];
  if (apptRanges.length) {
    // Normalise the DB-side field names to what the renderer below expects.
    // span_start_date (when known) lets the timeline show the precise
    // transition date instead of only the year — this is what distinguishes
    // Altaev's 1965-02-02 (Joined + Senior Engineer) from his 1965-04-01
    // (transition to Acting Junior Researcher).
    apptRanges = apptRanges.slice()
      .sort((a,b) => a.span_order - b.span_order)
      .map(s => ({
        position:        s.position,
        is_acting:       s.is_acting,
        tier:            s.tier,
        rank:            s.rank,
        start_year:      s.start_year,
        end_year:        s.end_year,
        span_order:      s.span_order,
        span_start_date: s.span_start_date,
      }));
  } else {
    // Fallback: fold EV.appointments client-side.
    const appts = (EV.appointments[pid]||[]).slice().sort((a,b) =>
      (a.year - b.year) || ((a.ord||0) - (b.ord||0))
    );
    apptRanges = [];
    for (const a of appts) {
      const cur = apptRanges[apptRanges.length-1];
      if (cur && cur.position === a.position && cur.is_acting === a.is_acting) {
        cur.end_year = a.year;
      } else {
        apptRanges.push({
          position:  a.position,
          is_acting: a.is_acting,
          tier:      a.tier,
          rank:      a.rank,
          start_year: a.year,
          end_year:   a.year,
        });
      }
    }
  }

  // ── 2) Build the main career timeline ──────────────────────────────
  // The timeline carries: Joined-CEMI (start_of_work with is_initial_join=1),
  // Position changes (start_of_work with is_initial_join=0 — linked to a
  // later position span via span_id), PhD defence, Academic title,
  // Position ranges, Dismissal, Transfer.
  // Pre-CEMI / Party / Dual move into dedicated side cards (below).
  const events = [];

  // v5: render every start_of_work row.  Rows with is_initial_join=1
  // (i.e. opening the first position span) appear as "Joined …"; rows
  // with is_initial_join=0 appear as "Position change → <position>",
  // using the span_id-linked position label so the reader can see which
  // role the spell opened with.  Persons whose original-hire row was
  // not recorded carry is_initial_join=0 on every row and therefore
  // have no "Joined" event on the timeline — the underlying spell is
  // still visible as a position range.
  const startList = (EV.starts[pid]||[]).slice().sort((a,b) =>
    (a.start_date||'').localeCompare(b.start_date||'')
  );
  for (const s of startList){
    if (!s) continue;
    const sy = parseInt((s.start_date||'').slice(0,4)) || null;
    if (s.is_initial_join){
      const sphase = phaseFor(sy);
      const joined =
        sphase === 'Predecessor Lab'  ? 'Joined the <b>Laboratory of Mathematical Methods Applied to Economic Research and Planning</b> (CEMI predecessor)'
      : sphase === 'Pre-Institute'    ? 'Recorded start of work — predates the predecessor lab'
      : 'Joined <b>CEMI</b>';
      events.push({yr:sy, ord:0, kind:'start', label:'Joined',
                   text:`${joined}<span>${s.start_date||''}</span>`});
    } else {
      // Position change — use the span_id-linked label if available.
      const posLabel = s.position ? esc(s.position) : 'new position';
      const actingPrefix = s.span_is_acting ? '<i>Acting</i> ' : '';
      events.push({yr:sy, ord:0.05, kind:'position-change', label:'Position change',
                   text:`Position change → ${actingPrefix}<b>${posLabel}</b><span>${s.start_date||''}</span>`});
    }
  }

  // Position ranges become single timeline rows.  When the DB linked a
  // start_of_work row to this span (span_start_date != null), we surface
  // it in the year column so two same-year transitions are visually
  // distinct.  ord uses span_order so within-year spells sort reliably.
  for (const r of apptRanges){
    const eph = phaseFor(r.start_year);
    const atInst =
      eph === 'Predecessor Lab' ? ' at the <i>Predecessor Lab</i>'
    : eph === 'Pre-Institute'   ? ''
    : ' at <i>CEMI</i>';
    let range;
    if (r.span_start_date) {
      // YYYY-MM-DD start; for multi-year spans append " – endYear".
      range = (r.end_year && r.end_year !== r.start_year)
                ? `${r.span_start_date} – ${r.end_year}`
                : r.span_start_date;
    } else {
      range = (r.start_year === r.end_year)
                ? String(r.start_year)
                : `${r.start_year}–${r.end_year}`;
    }
    // Span_order disambiguates ord within a single year so Senior Engineer
    // (#1) renders before Acting Junior Researcher (#2) when both start
    // in 1965 — a Date-tied tie-break gives stable presentation order.
    const ord = 1 + (Number.isFinite(r.span_order) ? r.span_order * 0.001 : 0);
    events.push({
      yr: r.start_year, range, ord,
      kind: r.is_acting ? 'appt-acting' : 'appt',
      label: eph === 'Predecessor Lab' ? 'Lab post' : 'Position',
      text: `<b>${esc(r.position)}</b>${atInst}<span>tier=${r.tier}, rank=${r.rank}${r.is_acting?', acting':''}</span>`,
    });
  }

  // PhD defence events used to appear here as point events; they now
  // live in the Degrees side card (which renders degree name + school +
  // defense year + field).  EV.phd is no longer pushed to the timeline.
  // Academic title events used to live on the timeline; they now have a
  // dedicated side card (see degreeHTML / titleHTML below), so they are
  // intentionally omitted from the timeline event list.

  // Causal order within a single year: Dismissal happens FIRST, then Transfer.
  for (const e of (EV.dismissals[pid]||[]))
    events.push({yr:e.year, ord:9, kind:'dismissal', label:'Dismissal',
                 text:`Dismissed${e.position?' from <b>'+esc(e.position)+'</b>':''}${e.date?'<span>'+e.date+'</span>':''}`});
  for (const e of (EV.transfers[pid]||[]))
    events.push({yr:e.year, ord:9.5, kind:'transfer', label:'Transfer',
                 text:`Transferred to <b>${esc(e.institution)}</b>${e.kind?'<span>'+esc(e.kind)+'</span>':''}`});

  events.sort((a,b)=>{
    if (a.yr==null && b.yr!=null) return 1;
    if (b.yr==null && a.yr!=null) return -1;
    if (a.yr !== b.yr) return (a.yr||0) - (b.yr||0);
    return (a.ord||0) - (b.ord||0);
  });

  // ── 3) Side-panel content: Pre-CEMI, Party, Dual, Degrees, Titles ──
  const preList   = EV.pre_cemi[pid]||[];
  const partyList = EV.party[pid]||[];
  const dualList  = EV.duals[pid]||[];
  const titleList = EV.titles[pid]||[];

  // Degree history — prefer the DB-side person_degree derived table; fall
  // back to deriving from EV.observations on older DBs.
  let degList = (EV.degree_history && EV.degree_history[pid]) || [];
  if (!degList.length) {
    const obs = EV.observations[pid] || [];
    const byDeg = {};
    for (const o of obs) {
      if (!o.degree || o.year == null) continue;
      const b = byDeg[o.degree] = byDeg[o.degree] || {
        degree: o.degree, first_year: o.year, last_year: o.year, n_years: 1, years: new Set([o.year]),
      };
      if (o.year < b.first_year) b.first_year = o.year;
      if (o.year > b.last_year)  b.last_year  = o.year;
      b.years.add(o.year);
      b.n_years = b.years.size;
    }
    degList = Object.values(byDeg).sort((a,b) => a.first_year - b.first_year);
  }

  const preCemiHTML = preList.length ? `
    <div class="section"><h4>Pre-CEMI Career<span class="n">${preList.length}</span></h4>
      <div class="info-grid">
        ${preList.map(e => `
          <div class="info-item">
            <div class="info-title">${esc(e.title||'(role unspecified)')}</div>
            <div class="info-sub">${esc(e.institution)}</div>
            ${e.kind ? `<div class="info-meta">${esc(e.kind)}</div>` : ''}
          </div>`).join('')}
      </div>
    </div>` : '';

  // Surface CPSU Member / CPSU Candidate / Komsomol Member etc. directly
  // in the card title so Candidate status is immediately obvious.
  const partyHTML = partyList.length ? `
    <div class="section"><h4>Party Affiliation<span class="n">${partyList.length}</span></h4>
      <div class="info-grid">
        ${partyList.map(e => {
          const party = e.party || '(unspecified)';
          const role  = e.role  || '';
          const title = role ? `${party} ${role}` : party;
          // Highlight Candidate status with the same color used for "acting" position pills.
          const cls = role.toLowerCase() === 'candidate' ? ' style="color:var(--warn);"' : '';
          return `<div class="info-item">
            <div class="info-year">${e.join_year ?? '—'}</div>
            <div class="info-title"${cls}>${esc(title)}</div>
          </div>`;
        }).join('')}
      </div>
    </div>` : '';

  // Dual positions are grouped by (role, institution, kind) — when the same
  // dual post appears across multiple recording years we merge the duration_text
  // (or year) into a single "YYYY-YYYY" range.
  const dualKeyed = {};
  for (const e of dualList){
    const role = e.role || '(role unspecified)';
    const key = `${role}|${e.institution}|${e.kind||''}`;
    const dur = (e.duration && String(e.duration).trim()) || (e.year != null ? String(e.year) : '');
    const bucket = dualKeyed[key] = dualKeyed[key] || {
      role, institution: e.institution, kind: e.kind, durations: new Set(), years: new Set(),
    };
    if (dur) bucket.durations.add(dur);
    if (e.year != null) bucket.years.add(+e.year);
  }
  const dualGrouped = Object.values(dualKeyed).map(b => {
    let span = '';
    // Prefer explicit duration_text when present.  When several distinct
    // duration strings exist (e.g. "1965" + "1965-1966"), pick the one that
    // is itself already a range; otherwise compute min-max from years.
    const durs = [...b.durations];
    const ranged = durs.find(d => /\d{4}\s*[–-]\s*\d{4}/.test(d));
    if (ranged) span = ranged;
    else if (durs.length === 1) span = durs[0];
    else if (b.years.size){
      const ys = [...b.years].sort((a,b)=>a-b);
      span = (ys[0] === ys[ys.length-1]) ? String(ys[0]) : `${ys[0]}–${ys[ys.length-1]}`;
    }
    return { role:b.role, institution:b.institution, kind:b.kind, span };
  });
  const dualHTML = dualGrouped.length ? `
    <div class="section"><h4>Dual Positions<span class="n">${dualGrouped.length}</span></h4>
      <div class="info-grid">
        ${dualGrouped.map(e => `
          <div class="info-item">
            <div class="info-year">${esc(e.span || '—')}</div>
            <div class="info-title">${esc(e.role)}</div>
            <div class="info-sub">at ${esc(e.institution)}</div>
            ${e.kind ? `<div class="info-meta">${esc(e.kind)}</div>` : ''}
          </div>`).join('')}
      </div>
    </div>` : '';

  // Degrees — display year + degree + school (when known) + defense /
  // field (when known).  Missing pieces simply leave their line blank,
  // per user spec.  When a precise PhD defense date is available we use
  // its year as the headline year; otherwise we fall back to first_year
  // from observations.
  const degreeHTML = degList.length ? `
    <div class="section"><h4>Degrees<span class="n">${degList.length}</span></h4>
      <div class="info-grid">
        ${degList.slice().sort((a,b) => {
            const ya = e2yr(a), yb = e2yr(b);
            return (ya||0) - (yb||0);
        }).map(e => {
          const yr = e2yr(e);
          const lines = [];
          if (e.school)       lines.push(`<div class="info-sub">at ${esc(e.school)}</div>`);
          if (e.defense_date) lines.push(`<div class="info-meta">defended ${esc(e.defense_date)}</div>`);
          if (e.field)        lines.push(`<div class="info-meta">field: ${esc(e.field)}</div>`);
          // Year-span shown only when no more specific defense_date exists.
          if (!e.defense_date && e.last_year && e.last_year !== e.first_year)
            lines.push(`<div class="info-meta">observed ${e.first_year}–${e.last_year}</div>`);
          return `<div class="info-item">
            <div class="info-year">${yr ?? '—'}</div>
            <div class="info-title">${esc(e.degree)}</div>
            ${lines.join('')}
          </div>`;
        }).join('')}
      </div>
    </div>` : '';
  // Tiny helper kept local to openPerson — pick the most precise year.
  function e2yr(e){
    if (e.defense_date) {
      const y = parseInt(String(e.defense_date).slice(0,4));
      if (!Number.isNaN(y)) return y;
    }
    return e.first_year ?? null;
  }

  // Academic titles — uses the existing academic_title_award table directly.
  const titleHTML = titleList.length ? `
    <div class="section"><h4>Academic Titles<span class="n">${titleList.length}</span></h4>
      <div class="info-grid">
        ${titleList.slice().sort((a,b) => (a.first_year||0) - (b.first_year||0)).map(e => `
          <div class="info-item">
            <div class="info-year">${e.first_year ?? '—'}</div>
            <div class="info-title">${esc(e.title)}</div>
          </div>`).join('')}
      </div>
    </div>` : '';

  // ── 4) Profile head facts ──────────────────────────────────────────
  const facts = [
    ['Person ID', p.id],
    ['Surname',   p.surname],
    ['Initials',  p.initials],
    ['Birth year',p.birth_year],
    ['Ethnicity', p.ethnicity],
    ['First observed', p.first_year],
    ['Last observed',  p.last_year],
    ['Years observed', p.years_observed],
  ];
  const tagsHTML = (label, arr, vocabKey, key='label') => {
    if (!arr || !arr.length) return '';
    const tags = arr.map(id => `<span class="tag"><b>${esc(VL[vocabKey][id]?.[key] ?? id)}</b></span>`).join(' ');
    return `<div class="section"><h4>${label}<span class="n">${arr.length}</span></h4><div class="tag-row">${tags}</div></div>`;
  };
  const PHASE_ORDER = ['Pre-Institute', 'Predecessor Lab', 'CEMI'];
  const PHASE_PILL  = {
    'Pre-Institute':   'phase-pre',
    'Predecessor Lab': 'phase-pred',
    'CEMI':            'phase-cemi',
  };
  const phaseSegments = PHASE_ORDER
    .filter(ph => p.phase_spans && p.phase_spans[ph])
    .map(ph => {
      const [a, b] = p.phase_spans[ph];
      const range = a === b ? String(a) : `${a}–${b}`;
      return `<span class="phase-badge ${PHASE_PILL[ph]}">${ph}</span> <b>${range}</b>`;
    })
    .join(' <span style="color:var(--dim);margin:0 6px;">→</span> ');

  let html = `
    <div class="profile-head">
      <div class="nm">${esc(p.name)}</div>
      <div class="meta">
        <span><b>${p.years_observed||0}</b> years observed</span>
        <span>${p.first_year||'—'}–${p.last_year||'—'}</span>
        ${p.birth_year ? `<span>born ${p.birth_year}</span>` : ''}
        ${p.ethnicity ? `<span>ethnicity: <b>${esc(p.ethnicity)}</b></span>` : ''}
      </div>
      ${phaseSegments ? `<div class="meta" style="margin-top:8px;">${phaseSegments}</div>` : ''}
      <div class="facts">
        ${facts.map(([l,v])=>`<div class="f"><div class="l">${l}</div><div class="v">${esc(v)}</div></div>`).join('')}
      </div>
    </div>`;

  // Side-panel cards first (above the timeline)
  html += preCemiHTML + partyHTML + dualHTML + degreeHTML + titleHTML;

  // Career timeline — consolidated, with side info stripped out.
  // A small inline legend explains the marker colors used by .ev::before.
  const timelineLegend = `
      <div class="timeline-legend">
        <span class="lg"><span class="dot start"></span>Joined CEMI</span>
        <span class="lg"><span class="dot appt"></span>Position (incl. acting)</span>
        <span class="lg"><span class="dot dismissal"></span>Dismissal</span>
        <span class="lg"><span class="dot transfer"></span>Transfer</span>
      </div>`;
  if (events.length) {
    html += `<div class="section"><h4>Career timeline<span class="n">${events.length}</span></h4>
      ${timelineLegend}
      <div class="tl">${events.map(e => `
        <div class="ev ${e.kind}">
          <span class="yr">${e.range || e.yr || '—'}</span>${e.yr != null ? phasePill(e.yr) : ''}
          <span class="kind">${e.label}</span>
          <span class="txt">${e.text}</span>
        </div>`).join('')}
      </div></div>`;
  } else {
    html += '<div class="empty">No career events.</div>';
  }

  // Tag rows (positions/institutions/etc.)
  html += tagsHTML('Positions held',          p.positions,   'position');
  html += tagsHTML('Institutions touched',    p.institutions,'institution');
  html += tagsHTML('Departments',             p.departments, 'department');
  html += tagsHTML('Specialties',             p.specialties, 'specialty');
  html += tagsHTML('PhD field(s)',            p.fields,      'phd_field');
  html += tagsHTML('Schools (BA / graduate)', p.schools,     'school');
  // Note: Degrees and Academic Titles are now rendered as dedicated
  // side cards above (see degreeHTML / titleHTML).  Parties remain as a
  // tag row here so the search-vocab linkage stays visible.
  html += tagsHTML('Parties',                 p.parties,     'party');

  document.getElementById('profile').innerHTML = html;
  search();   // refresh selection styling
  card.scrollIntoView({behavior:'smooth', block:'start'});
}

// ── Institutions tab ───────────────────────────────────────────
// Institutions tab — pagination state (independent of the Career
// Search pager so the two tabs do not interfere).
const INST_PAGE_SIZE = 100;
let INST_PAGE = 1;
function gotoInstPage(delta){
  INST_PAGE = Math.max(1, INST_PAGE + delta);
  renderInstitutions();
}

INIT.institutions = function(){
  // Build a flat table summarising each institution's CEMI exposure.
  const rows = VOCAB.institution.map(inst => {
    let nPre=0, nDual=0, nTrans=0;
    for (const p of PERSONS){
      if ((EV.pre_cemi[p.id]||[]).some(e => VL.institution[inst.id] && (e.institution===VL.institution[inst.id].label))) nPre++;
      if ((EV.duals[p.id]||[]).some(e => e.institution===inst.label)) nDual++;
      if ((EV.transfers[p.id]||[]).some(e => e.institution===inst.label)) nTrans++;
    }
    // nPrimary = distinct persons whose Primary Institution (data.xlsx
    // "Primary Institution" column → person_year_observation
    // .primary_institution_id) equals this institution.  For the CEMI row
    // itself this counts full-time CEMI staff; for any external row it
    // counts persons whose primary employer is that institution and for
    // whom CEMI is a part-time / dual post.
    const nPrimary = INST_PRIMARY_COUNT[inst.id] || 0;
    return {id:inst.id, label:inst.label, label_ru:inst.label_ru,
            kind:inst.kind, evidence:inst.classification_evidence,
            nPre, nDual, nTrans, nPrimary,
            total:nPre+nDual+nTrans+nPrimary};
  });
  window.__INST_ROWS__ = rows.sort((a,b)=>b.total-a.total);

  // Populate the institution-kind taxonomy dropdown from whatever values
  // actually exist in VOCAB.institution.  This adapts automatically to
  // the new Academy / VUZ / Branch / Enterprise / Other scheme AND to
  // any future relabelling.
  const ksel = document.getElementById('i-kindtax');
  if (ksel) {
    const kinds = [...new Set(VOCAB.institution.map(i => i.kind).filter(Boolean))].sort();
    for (const k of kinds) {
      const o = document.createElement('option');
      o.value = k; o.textContent = k;
      ksel.appendChild(o);
    }
  }

  INST_PAGE = 1;
  renderInstitutions();
};

function renderInstitutions(){
  const q = (document.getElementById('i-search').value||'').trim().toLowerCase();
  const k = document.getElementById('i-kind').value;
  const kt = (document.getElementById('i-kindtax')||{value:''}).value;
  let rows = window.__INST_ROWS__||[];
  if (q)  rows = rows.filter(r=>(r.label||'').toLowerCase().includes(q));
  if (k==='Pre-CEMI')     rows = rows.filter(r=>r.nPre>0);
  if (k==='Dual')         rows = rows.filter(r=>r.nDual>0);
  if (k==='Transfer')     rows = rows.filter(r=>r.nTrans>0);
  // Primary axis split into two mutually-exclusive groups so the user can
  // see (a) the CEMI row alone — full-time CEMI staff — or (b) every
  // external institution that is anyone's Primary while CEMI is part-time.
  if (k==='Primary-CEMI') rows = rows.filter(r=>r.id === CEMI_INST_ID && r.nPrimary>0);
  if (k==='Primary-Ext')  rows = rows.filter(r=>r.id !== CEMI_INST_ID && r.nPrimary>0);
  if (kt) rows = rows.filter(r => (r.kind||'') === kt);

  const totalPages = Math.max(1, Math.ceil(rows.length / INST_PAGE_SIZE));
  if (INST_PAGE > totalPages) INST_PAGE = totalPages;
  const startIdx = (INST_PAGE - 1) * INST_PAGE_SIZE;
  const display  = rows.slice(startIdx, startIdx + INST_PAGE_SIZE);

  let html = '<table><thead><tr><th>Institution</th><th>Kind</th><th>Pre-CEMI</th><th>Dual</th><th>Transfer-out</th><th title="Persons whose Primary Institution equals this row.  For the CEMI row → full-time CEMI staff; for external rows → persons whose primary employer is here and CEMI is part-time.">Primary</th><th>Total</th></tr></thead><tbody>';
  display.forEach(r => {
    // Russian short form shown beneath the English label when known; the
    // classification rationale (e.g. "RU: 'НИИ'") is exposed as a hover
    // tooltip on the Kind cell.
    const nameCell = r.label_ru
      ? `<b>${esc(r.label)}</b><br><span style="color:var(--dim);font-size:11px;">${esc(r.label_ru)}</span>`
      : esc(r.label);
    const kindCell = r.evidence
      ? `<span title="${esc(r.evidence)}" style="border-bottom:1px dotted var(--dim);cursor:help;">${esc(r.kind)}</span>`
      : esc(r.kind);
    html += `<tr><td>${nameCell}</td><td>${kindCell}</td>`
          + `<td class="num">${r.nPre}</td><td class="num">${r.nDual}</td>`
          + `<td class="num">${r.nTrans}</td><td class="num">${r.nPrimary}</td>`
          + `<td class="num">${r.total}</td></tr>`;
  });
  html += '</tbody></table>';
  if (rows.length > INST_PAGE_SIZE) {
    const firstRow = startIdx + 1;
    const lastRow  = Math.min(startIdx + INST_PAGE_SIZE, rows.length);
    html += `<div class="pagination">
        <button onclick="gotoInstPage(-1)" ${INST_PAGE === 1 ? 'disabled' : ''}>← Prev</button>
        <span class="pageinfo">Page <b>${INST_PAGE}</b> / ${totalPages}
          · rows <b>${fmt(firstRow)}–${fmt(lastRow)}</b> of <b>${fmt(rows.length)}</b></span>
        <button onclick="gotoInstPage(1)" ${INST_PAGE === totalPages ? 'disabled' : ''}>Next →</button>
      </div>`;
  }
  document.getElementById('i-table').innerHTML = html;
}

// Reset Institutions pager to page 1 whenever filters change.
['i-search','i-kind','i-kindtax'].forEach(id => {
  document.addEventListener('DOMContentLoaded', () => {
    const el = document.getElementById(id);
    if (!el) return;
    const trigger = el.tagName === 'INPUT' ? 'input' : 'change';
    el.addEventListener(trigger, () => { INST_PAGE = 1; });
  });
});

// ── Positions tab ─────────────────────────────────────────────
// Layout follows the cemi_demographic_ui reference: per-year selector,
// two horizontal-bar charts side-by-side (Total Personnel · Degree
// Distribution), a grand-total banner, then the bilingual data table
// ending with a highlighted Grand Total row.  The rank-ladder card is
// retained beneath as supplementary metadata.
INIT.positions = function(){
  // Rank ladder ----------------------------------------------------
  const ladderHead = '<table><thead><tr><th>Rank</th><th>Position</th><th>Tier</th><th>Acting?</th><th>CEMI appointments</th></tr></thead><tbody>';
  const counts = {};
  for (const pid in EV.appointments) for (const a of EV.appointments[pid]) counts[a.position] = (counts[a.position]||0)+1;
  const ladderRows = VOCAB.position.slice().sort((a,b)=>a.rank_order-b.rank_order).map(p =>
    `<tr><td class="num">${p.rank_order}</td><td>${esc(p.label)}</td><td>${esc(p.tier)}</td><td>${p.is_acting? '<span class="pill pill-act">acting</span>':'<span class="pill">substantive</span>'}</td><td class="num">${counts[p.label]||0}</td></tr>`);
  document.getElementById('pos-ladder-table').innerHTML = ladderHead + ladderRows.join('') + '</tbody></table>';

  // Year selector --------------------------------------------------
  // Use the individual-data-derived breakdown which covers every position
  // (Trainee, Engineer, Researcher, Leadership, …) and every observed year.
  const SRC = (AGG.positions_individual && AGG.positions_individual.length)
              ? AGG.positions_individual : AGG.positions;
  const years = [...new Set(SRC.map(r=>r.year))].sort((a,b)=>a-b);
  const sel = document.getElementById('p-year'); sel.innerHTML='';
  for (const y of years){const o=document.createElement('option');o.value=y;o.textContent=yearOptionLabel(y);sel.appendChild(o);}
  if (years.length) sel.value = years[years.length-1];
  renderPositionsTab();
};
function renderPositionsTab(){
  const y = parseInt(document.getElementById('p-year').value);
  const SRC = (AGG.positions_individual && AGG.positions_individual.length)
              ? AGG.positions_individual : AGG.positions;

  // Sort positions by total descending so the longest bars sit at the top.
  const rows = SRC.filter(r=>r.year===y).slice()
                  .sort((a,b)=>(b.total||0)-(a.total||0));
  const labels = rows.map(r=>r.position);

  // Chart 1 — Total personnel per position (single-series horizontal bar)
  mk('pos-chart-total','bar',{labels,
      datasets:[{label:'Total Personnel', data:rows.map(r=>r.total||0),
                 backgroundColor:labels.map((_,i)=>c(i,.75))}]},
      co({indexAxis:'y', plugins:{legend:{display:false},
          tooltip:{callbacks:{label:(ctx)=>`${ctx.parsed.x} persons`}}}}));

  // Chart 2 — Degree distribution within each position
  // The individual-level source provides a "No degree / not recorded" bucket
  // alongside Ph.D. and National Doctor, so we render a 100% stacked bar
  // chart that shows the degree composition of every position for the year.
  const hasNoDegree = rows.some(r => r.no_degree != null);
  const pct = (n, t) => t ? Math.round((n/t)*1000)/10 : 0;
  const degDatasets = [
    {label:'National Doctor', data: rows.map(r => pct(r.national_doctor||0, r.total||0)),
      backgroundColor:c(1,.85), stack:'deg'},
    {label:'Ph.D.',           data: rows.map(r => pct(r.phd||0,             r.total||0)),
      backgroundColor:c(0,.85), stack:'deg'},
  ];
  if (hasNoDegree) {
    degDatasets.push({label:'No degree / not recorded',
      data: rows.map(r => pct(r.no_degree||0, r.total||0)),
      backgroundColor:c(8,.55), stack:'deg'});
  }
  mk('pos-chart-deg','bar',{labels, datasets: degDatasets},
      co({indexAxis:'y',
          plugins:{legend:{position:'top'},
            tooltip:{callbacks:{label:(ctx)=>{
              const r = rows[ctx.dataIndex];
              const key = ctx.dataset.label==='Ph.D.' ? 'phd'
                        : ctx.dataset.label==='National Doctor' ? 'national_doctor'
                        : 'no_degree';
              return `${ctx.dataset.label}: ${ctx.parsed.x}% (${r[key]||0} of ${r.total||0})`;
            }}}},
          scales:{x:{stacked:true, min:0, max:100,
                     ticks:{color:'#878fa8', callback:v=>v+'%'}, grid:{color:'#2c3145'}},
                  y:{stacked:true}}}));

  // Grand-total banner --------------------------------------------
  const tot = rows.reduce((a,r)=>a+(r.total||0),0);
  const nd  = rows.reduce((a,r)=>a+(r.national_doctor||0),0);
  const phd = rows.reduce((a,r)=>a+(r.phd||0),0);
  const nodeg = rows.reduce((a,r)=>a+(r.no_degree||0),0);
  const cpsu= rows.reduce((a,r)=>a+(r.cpsu_members||0),0);
  const kom = rows.reduce((a,r)=>a+(r.komsomol||0),0);
  const banner = document.getElementById('pos-banner');
  if (rows.length){
    banner.style.display='block';
    let txt = `Grand Total (${y}): <b>${fmt(tot)}</b> personnel across <b>${rows.length}</b> position types`
            + ` · National Doctors: <b>${fmt(nd)}</b> · Ph.D.s: <b>${fmt(phd)}</b>`;
    if (hasNoDegree) txt += ` · No degree recorded: <b>${fmt(nodeg)}</b>`;
    banner.innerHTML = txt;
  } else {
    banner.style.display='none';
  }

  // Data table with a highlighted Grand-Total row ------------------
  const hdrCols = hasNoDegree
    ? '<th>Position (Russian)</th><th>Position (English)</th><th>Tier</th><th>Rank</th>'
      + '<th>Total</th><th>National Doctor</th><th>Ph.D.</th><th>No degree</th><th>% with degree</th>'
    : '<th>Position (Russian)</th><th>Position (English)</th><th>Tier</th><th>Rank</th>'
      + '<th>Total</th><th>National Doctor</th><th>Ph.D.</th><th>CPSU Members</th><th>Komsomol</th>';
  const head = '<table><thead><tr>' + hdrCols + '</tr></thead><tbody>';
  const body = rows.map(r => {
    const totR = r.total || 0;
    const withDeg = (r.national_doctor||0) + (r.phd||0);
    const pctDeg = totR ? Math.round(withDeg/totR*1000)/10 : 0;
    if (hasNoDegree) {
      return `<tr>
        <td>${esc(r.position_ru)}</td>
        <td>${esc(r.position)}${r.is_acting?' <span class="pill pill-act">acting</span>':''}</td>
        <td>${esc(r.tier)}</td>
        <td class="num">${fmt(r.rank_order)}</td>
        <td class="num">${fmt(r.total)}</td>
        <td class="num">${fmt(r.national_doctor)}</td>
        <td class="num">${fmt(r.phd)}</td>
        <td class="num">${fmt(r.no_degree)}</td>
        <td class="num">${pctDeg}%</td>
      </tr>`;
    }
    return `<tr>
      <td>${esc(r.position_ru)}</td>
      <td>${esc(r.position)}${r.is_acting?' <span class="pill pill-act">acting</span>':''}</td>
      <td>${esc(r.tier)}</td>
      <td class="num">${fmt(r.rank_order)}</td>
      <td class="num">${fmt(r.total)}</td>
      <td class="num">${fmt(r.national_doctor)}</td>
      <td class="num">${fmt(r.phd)}</td>
      <td class="num">${fmt(r.cpsu_members)}</td>
      <td class="num">${fmt(r.komsomol)}</td>
    </tr>`;
  }).join('');
  let grand = '';
  if (rows.length){
    const withDegT = nd + phd;
    const pctDegT  = tot ? Math.round(withDegT/tot*1000)/10 : 0;
    if (hasNoDegree){
      grand = `<tr class="grand-total-row">
        <td>—</td>
        <td><b>Grand Total</b></td>
        <td>—</td>
        <td class="num">—</td>
        <td class="num">${fmt(tot)}</td>
        <td class="num">${fmt(nd)}</td>
        <td class="num">${fmt(phd)}</td>
        <td class="num">${fmt(nodeg)}</td>
        <td class="num">${pctDegT}%</td>
      </tr>`;
    } else {
      grand = `<tr class="grand-total-row">
        <td>—</td>
        <td><b>Grand Total</b></td>
        <td>—</td>
        <td class="num">—</td>
        <td class="num">${fmt(tot)}</td>
        <td class="num">${fmt(nd)}</td>
        <td class="num">${fmt(phd)}</td>
        <td class="num">${fmt(cpsu)}</td>
        <td class="num">${fmt(kom)}</td>
      </tr>`;
    }
  }
  document.getElementById('pos-data-table').innerHTML = head + body + grand + '</tbody></table>';
}

// ── Personnel growth tab (cemi_translation_db.xlsx · Personnel Totals) ───────
INIT.personnel = function(){
  const pt = AGG.personnel_totals;
  const yrs = pt.map(r=>r.year);
  mk('pn-staff','line',{labels:yrs,datasets:[
    {label:'Total scientists', data:pt.map(r=>r.total_scientists), borderColor:c(0), backgroundColor:c(0,.12), tension:.3, fill:true, spanGaps:true},
    {label:'Total all staff',  data:pt.map(r=>r.total_all_staff),  borderColor:c(1), backgroundColor:c(1,.10), tension:.3, fill:true, spanGaps:true},
    {label:'National Doctor',  data:pt.map(r=>r.national_doctor),  borderColor:c(2), tension:.3, spanGaps:true},
    {label:'Ph.D.',            data:pt.map(r=>r.phd),              borderColor:c(3), tension:.3, spanGaps:true},
  ]}, co({plugins:{legend:{position:'top'}}}));
  mk('pn-women','line',{labels:yrs,datasets:[
    {label:'Women',                 data:pt.map(r=>r.women),                  borderColor:c(4), tension:.3, spanGaps:true},
    {label:'Women National Doctor', data:pt.map(r=>r.women_national_doctor), borderColor:c(2), tension:.3, spanGaps:true},
    {label:'Women Ph.D.',           data:pt.map(r=>r.women_phd),             borderColor:c(3), tension:.3, spanGaps:true},
  ]}, co({plugins:{legend:{position:'top'}}}));

  const head = '<table><thead><tr>'
    + '<th>Year</th><th>Total Scientists</th><th>Total All Staff</th>'
    + '<th>National Doctor</th><th>Ph.D.</th>'
    + '<th>Women</th><th>Women National Doctor</th><th>Women Ph.D.</th>'
    + '<th>Source Material</th>'
    + '</tr></thead><tbody>';
  const body = pt.map(r => `<tr>
      <td>${r.year}</td>
      <td class="num">${fmt(r.total_scientists)}</td>
      <td class="num">${fmt(r.total_all_staff)}</td>
      <td class="num">${fmt(r.national_doctor)}</td>
      <td class="num">${fmt(r.phd)}</td>
      <td class="num">${fmt(r.women)}</td>
      <td class="num">${fmt(r.women_national_doctor)}</td>
      <td class="num">${fmt(r.women_phd)}</td>
      <td><small style="color:var(--dim);">${esc(r.source_material)}</small></td>
    </tr>`).join('');
  document.getElementById('pn-table').innerHTML = head + body + '</tbody></table>';
};

// ── Degrees & titles tab (cemi_translation_db.xlsx · Academic Degrees) ───────
INIT.degrees = function(){
  const d = AGG.academic_degrees;
  const yrs = d.map(r=>r.year);
  mk('dg-deg','line',{labels:yrs,datasets:[
    {label:'National Doctor', data:d.map(r=>r.national_doctor), borderColor:c(2), tension:.3, spanGaps:true},
    {label:'Ph.D.',           data:d.map(r=>r.phd),             borderColor:c(3), tension:.3, spanGaps:true},
    {label:'Total scientists',data:d.map(r=>r.total_scientists),borderColor:c(0), tension:.3, spanGaps:true},
  ]}, co({plugins:{legend:{position:'top'}}}));
  mk('dg-titles','line',{labels:yrs,datasets:[
    {label:'Professors',   data:d.map(r=>r.professors),    borderColor:c(0), tension:.3, spanGaps:true},
    {label:'Docents',      data:d.map(r=>r.docents),       borderColor:c(2), tension:.3, spanGaps:true},
    {label:'SNS Title',    data:d.map(r=>r.sns_title),     borderColor:c(3), tension:.3, spanGaps:true},
    {label:'Academicians', data:d.map(r=>r.academicians),  borderColor:c(4), tension:.3, spanGaps:true},
  ]}, co({plugins:{legend:{position:'top'}}}));
  mk('dg-pct','line',{labels:yrs,datasets:[
    {label:'% with degree', data:d.map(r=>r.pct_with_degree), borderColor:c(1), backgroundColor:c(1,.12), tension:.3, fill:true, spanGaps:true},
  ]}, co({plugins:{legend:{display:false}},scales:{y:{ticks:{color:'#878fa8',callback:v=>v+'%'},grid:{color:'#2c3145'}}}}));

  const head = '<table><thead><tr>'
    + '<th>Year</th><th>Total Scientists</th><th>National Doctor</th><th>Ph.D.</th>'
    + '<th>% with Degree</th><th>Professors</th><th>Docents</th><th>SNS Title</th><th>Academicians</th>'
    + '<th>Source Material</th>'
    + '</tr></thead><tbody>';
  const body = d.map(r => `<tr>
      <td>${r.year}</td>
      <td class="num">${fmt(r.total_scientists)}</td>
      <td class="num">${fmt(r.national_doctor)}</td>
      <td class="num">${fmt(r.phd)}</td>
      <td class="num">${fmt(r.pct_with_degree)}</td>
      <td class="num">${fmt(r.professors)}</td>
      <td class="num">${fmt(r.docents)}</td>
      <td class="num">${fmt(r.sns_title)}</td>
      <td class="num">${fmt(r.academicians)}</td>
      <td><small style="color:var(--dim);">${esc(r.source_material)}</small></td>
    </tr>`).join('');
  document.getElementById('dg-table').innerHTML = head + body + '</tbody></table>';
};

// ── Nationality tab (cemi_translation_db.xlsx · Nationality) ─────────────────
INIT.nationality = function(){
  const years = [...new Set(AGG.nationality.map(r=>r.year))].sort((a,b)=>a-b);
  const sel = document.getElementById('nat-year'); sel.innerHTML='';
  for (const y of years){const o=document.createElement('option');o.value=y;o.textContent=yearOptionLabel(y);sel.appendChild(o);}
  if (years.length) sel.value = years[years.length-1];
  // trend chart: top-6 nationalities by total across the corpus
  const totals = {};
  for (const r of AGG.nationality) totals[r.en] = (totals[r.en]||0) + (r.total||0);
  const top6 = Object.entries(totals).sort((a,b)=>b[1]-a[1]).slice(0,6).map(x=>x[0]);
  const ds = top6.map((n,i)=>({
    label:n, borderColor:c(i), tension:.3, spanGaps:true,
    data: years.map(y => {
      const m = AGG.nationality.find(r=>r.year===y && r.en===n);
      return m ? m.total : null;
    }),
  }));
  mk('nat-trend','line',{labels:years,datasets:ds}, co({plugins:{legend:{position:'top',labels:{color:'#878fa8',font:{size:10}}}}}));
  renderNationality();
};
function renderNationality(){
  const y = parseInt(document.getElementById('nat-year').value);
  const rows = AGG.nationality.filter(r=>r.year===y).sort((a,b)=>(b.total||0)-(a.total||0));
  // Donut
  const top8 = rows.slice(0,8);
  const other = rows.slice(8).reduce((s,r)=>s+(r.total||0),0);
  const labels = [...top8.map(r=>r.en),...(other?['Other']:[])];
  const vals   = [...top8.map(r=>r.total),...(other?[other]:[])];
  mk('nat-donut','doughnut',{labels,datasets:[{data:vals,backgroundColor:labels.map((_,i)=>c(i,.85))}]},
     {plugins:{legend:{position:'right',labels:{color:'#878fa8',font:{size:11}}}}});
  // Full bilingual table
  const head = '<table><thead><tr>'
    + '<th>Year</th><th>Nationality (Russian)</th><th>Nationality (English)</th>'
    + '<th>Total</th><th>National Doctor</th><th>Ph.D.</th>'
    + '</tr></thead><tbody>';
  const body = rows.map(r => `<tr>
      <td>${r.year}</td><td>${esc(r.ru)}</td><td>${esc(r.en)}</td>
      <td class="num">${fmt(r.total)}</td>
      <td class="num">${fmt(r.national_doctor)}</td>
      <td class="num">${fmt(r.phd)}</td>
    </tr>`).join('');
  document.getElementById('nat-table').innerHTML = head + body + '</tbody></table>';
}

// ── Age distribution tab (cemi_translation_db.xlsx · Age Distribution) ──────
INIT.age = function(){
  const years = [...new Set(AGG.age.map(r=>r.year))].sort((a,b)=>a-b);
  const sel = document.getElementById('age-year'); sel.innerHTML='';
  for (const y of years){const o=document.createElement('option');o.value=y;o.textContent=yearOptionLabel(y);sel.appendChild(o);}
  if (years.length) sel.value = years[years.length-1];
  renderAge();
};
function renderAge(){
  const y = parseInt(document.getElementById('age-year').value);
  const rows = AGG.age.filter(r=>r.year===y).sort((a,b)=>((a.lower_bound||0)-(b.lower_bound||0)));
  mk('age-total','bar',{labels:rows.map(r=>r.en),datasets:[{label:'Total',data:rows.map(r=>r.total),backgroundColor:c(2,.75)}]},
     co({plugins:{legend:{display:false}}}));
  mk('age-deg','bar',{labels:rows.map(r=>r.en),datasets:[
     {label:'National Doctor', data:rows.map(r=>r.national_doctor), backgroundColor:c(0,.8)},
     {label:'Ph.D.',           data:rows.map(r=>r.phd),             backgroundColor:c(3,.8)},
  ]}, co({plugins:{legend:{position:'top'}}}));

  // Chart 3 — academic title / rank breakdown per age bracket (stacked).
  // Years that do not record these columns produce a blank chart; show a
  // small "no data" notice in that case instead of an empty canvas.
  const titleCols = [
    {key:'academician_or_corr',  label:'Academician / Corresponding Member', color: c(0,.85)},
    {key:'professor',            label:'Professor',                          color: c(1,.85)},
    {key:'associate_professor',  label:'Associate Professor',                color: c(2,.85)},
    {key:'senior_researcher',    label:'Senior Researcher',                  color: c(3,.85)},
    {key:'junior_researcher',    label:'Junior Researcher',                  color: c(4,.85)},
  ];
  const anyTitleData = rows.some(r => titleCols.some(c => r[c.key] != null));
  const titlesCard = document.getElementById('age-titles-card');
  if (titlesCard) {
    if (anyTitleData) {
      titlesCard.style.display = '';
      mk('age-titles','bar',{
        labels: rows.map(r => r.en),
        datasets: titleCols.map(tc => ({
          label: tc.label,
          data:  rows.map(r => r[tc.key] ?? 0),
          backgroundColor: tc.color,
          stack: 'titles',
        })),
      }, co({plugins:{legend:{position:'top', labels:{color:'#878fa8', font:{size:11}}}},
              scales:{x:{stacked:true, ticks:{color:'#878fa8'}, grid:{color:'#2c3145'}},
                      y:{stacked:true, ticks:{color:'#878fa8'}, grid:{color:'#2c3145'}}}}));
    } else {
      titlesCard.style.display = 'none';
      if (CH['age-titles']) { CH['age-titles'].destroy(); delete CH['age-titles']; }
    }
  }

  // Extended data table — surfaces every column of the Age Distribution sheet.
  const head = '<table><thead><tr>'
    + '<th>Year</th><th>Age Bracket (Russian)</th><th>Age Bracket (English)</th>'
    + '<th>Total</th><th>National Doctor</th><th>Ph.D.</th>'
    + '<th>Acad./Corr.</th><th>Prof.</th><th>Assoc. Prof.</th>'
    + '<th>Senior Res.</th><th>Junior Res.</th>'
    + '</tr></thead><tbody>';
  const body = rows.map(r => `<tr>
      <td>${r.year}</td><td>${esc(r.ru)}</td><td>${esc(r.en)}</td>
      <td class="num">${fmt(r.total)}</td>
      <td class="num">${fmt(r.national_doctor)}</td>
      <td class="num">${fmt(r.phd)}</td>
      <td class="num">${fmt(r.academician_or_corr)}</td>
      <td class="num">${fmt(r.professor)}</td>
      <td class="num">${fmt(r.associate_professor)}</td>
      <td class="num">${fmt(r.senior_researcher)}</td>
      <td class="num">${fmt(r.junior_researcher)}</td>
    </tr>`).join('');
  document.getElementById('age-table').innerHTML = head + body + '</tbody></table>';
}

// ── Research Fields tab — merged view ─────────────────────────────────
//
// Top-level Research Field rows are rendered as parent cards.  Each card
// expands to its child Subfield rows from the matching YYYY sheet of
// research_field_subfield.xlsx.  Selectors:
//   • Year          — picks the year (single snapshot per year now)
//   • Field filter  — limits the cards to one Research Field
//
// The trend chart at the bottom plots top 10 Subfields year by year.
//
// Sheet names in the current workbook are plain "YYYY"; older YYYY.MM
// labels (1971.01 / 1971.06) are still tolerated by the snapshot logic
// below so that a downgrade of the workbook does not break the UI.

// Index of snapshot keys grouped by their year.  Built once.
let RF_SNAPSHOTS_BY_YEAR = null;
function rfSnapshotsByYear(){
  if (RF_SNAPSHOTS_BY_YEAR) return RF_SNAPSHOTS_BY_YEAR;
  RF_SNAPSHOTS_BY_YEAR = {};
  for (const k of Object.keys(AGG.subfields_by_period || {})) {
    const yr = parseInt(String(k).slice(0,4));
    (RF_SNAPSHOTS_BY_YEAR[yr] = RF_SNAPSHOTS_BY_YEAR[yr] || []).push(k);
  }
  for (const yr in RF_SNAPSHOTS_BY_YEAR) RF_SNAPSHOTS_BY_YEAR[yr].sort();
  return RF_SNAPSHOTS_BY_YEAR;
}

INIT.research = function(){
  // Union of years from both sources — even a year that only the yearly
  // Research Fields sheet covers is selectable.
  const snapshotIdx = rfSnapshotsByYear();
  const years = new Set();
  for (const r of AGG.fields) years.add(r.year);
  for (const y of Object.keys(snapshotIdx)) years.add(parseInt(y));
  const yearList = [...years].sort((a,b)=>a-b);

  const ys = document.getElementById('rf-year'); ys.innerHTML='';
  for (const y of yearList){
    const o = document.createElement('option');
    o.value = y;
    // Phase suffix only — Research Fields data now derives entirely from
    // research_field_subfield.xlsx, so every selectable year always has
    // both Field-level and Subfield-level data.
    o.textContent = (typeof yearOptionLabel === 'function') ? yearOptionLabel(y) : String(y);
    ys.appendChild(o);
  }
  if (yearList.length) ys.value = yearList[yearList.length-1];

  // ── Bottom trend chart — top 10 Subfields over time ────────────────
  // We pick one canonical snapshot per year (the latest), aggregate each
  // Subfield's Total Personnel across those years, take the top 10 by
  // cumulative total, then plot each as its own time series.
  {
    const periodSrc = AGG.subfields_by_period || {};
    const latestByYear = {};
    for (const k of Object.keys(periodSrc)) {
      const yr = parseInt(String(k).slice(0,4));
      if (!latestByYear[yr] || k > latestByYear[yr]) latestByYear[yr] = k;
    }
    const cumByName   = {};   // subfield → Σ across years
    const seriesByName = {};   // subfield → { year: total }
    const isMarker = s => !s || /^(total|итого)$/i.test(String(s).trim());
    for (const [yr, key] of Object.entries(latestByYear)) {
      const bucket = periodSrc[key];
      if (!bucket) continue;
      const yi = parseInt(yr);
      for (const r of bucket.rows) {
        if (isMarker(r.subfield)) continue;
        if (r.total == null || Number.isNaN(+r.total)) continue;
        const name = r.subfield;
        cumByName[name]    = (cumByName[name] || 0) + (+r.total);
        seriesByName[name] = seriesByName[name] || {};
        // If the same Subfield appears twice in one snapshot, keep the
        // larger value (defensive — should not happen in well-formed data).
        seriesByName[name][yi] = Math.max(seriesByName[name][yi] || 0, +r.total);
      }
    }
    const top10 = Object.entries(cumByName)
                        .sort((a,b) => b[1] - a[1])
                        .slice(0, 10)
                        .map(x => x[0]);
    const years = Object.keys(latestByYear).map(Number).sort((a,b) => a - b);
    const datasets = top10.map((name, i) => ({
      label: name,
      borderColor: c(i),
      tension: .3,
      spanGaps: true,
      data: years.map(y => seriesByName[name]?.[y] ?? null),
    }));
    mk('rf-trend','line',{labels:years,datasets},
       co({plugins:{legend:{position:'top',labels:{color:'#878fa8',font:{size:10}}}}}));
  }

  renderResearch();
};

function renderResearch(){
  const year         = parseInt(document.getElementById('rf-year').value);
  const fltSel       = document.getElementById('rf-field');
  const snapshotIdx  = rfSnapshotsByYear();
  const snapshotKeys = snapshotIdx[year] || [];

  // Snapshot picker is only visible when the year has more than one snapshot.
  const ss = document.getElementById('rf-snapshot');
  if (snapshotKeys.length > 1) {
    const prev = ss.value;
    ss.innerHTML = snapshotKeys.map(k => `<option ${prev===k?'selected':''}>${k}</option>`).join('');
    if (!snapshotKeys.includes(ss.value)) ss.value = snapshotKeys[snapshotKeys.length-1];
    ss.style.display = '';
  } else {
    ss.innerHTML = '';
    ss.style.display = 'none';
  }
  // Resolve which snapshot (if any) to use.
  const period   = (snapshotKeys.length > 1) ? ss.value : (snapshotKeys[0] || '');
  const snapshot = period ? (AGG.subfields_by_period || {})[period] : null;
  const allRows  = snapshot ? snapshot.rows : [];

  // ── Parent grouping ───────────────────────────────────────────────
  // When snapshot exists, derive parents from subfield groupings (source
  // of truth for the hierarchy).  Otherwise fall back to AGG.fields.
  let parents = [];
  let groupOrder = [];
  let groups     = {};
  if (snapshot) {
    for (const r of allRows) {
      const key = r.field || '(unspecified)';
      if (!(key in groups)) { groups[key] = []; groupOrder.push(key); }
      groups[key].push(r);
    }
    const isMarker = s => !s || /^(total|итого)$/i.test(String(s).trim());
    const sumCol = (rows, k) => {
      let any=false, t=0;
      for (const r of rows){ const v=r[k]; if(v!=null && !Number.isNaN(+v)){any=true; t+=+v;} }
      return any ? t : null;
    };
    parents = groupOrder.map(name => {
      const children = (groups[name]||[]).filter(r => !isMarker(r.subfield));
      return {
        name,
        child_count: children.length,
        total: sumCol(children, 'total'),
        national_doctor: sumCol(children, 'national_doctor'),
        phd: sumCol(children, 'phd'),
        children,
      };
    });
  } else {
    // Yearly-only fallback — every row of AGG.fields for this year.
    parents = AGG.fields.filter(r=>r.year===year).map(r => ({
      name: r.en, child_count: 0,
      total: r.total, national_doctor: r.national_doctor, phd: r.phd, children: [],
      code: r.code, ru: r.ru, is_aggregate: r.is_aggregate,
    }));
    groupOrder = parents.map(p=>p.name);
  }

  // ── Field-filter dropdown — derived from the active source ────────
  const previousFilter = fltSel.value;
  fltSel.innerHTML = '<option value="">— All Research Fields —</option>'
    + groupOrder.map(n => `<option ${previousFilter===n?'selected':''}>${esc(n)}</option>`).join('');
  if (!groupOrder.includes(previousFilter)) fltSel.value = '';
  if (fltSel.value) parents = parents.filter(p => p.name === fltSel.value);

  // ── Signature / source line ───────────────────────────────────────
  // period equals the sheet name; with the current workbook that is just
  // "YYYY".  Show a clean "Year YYYY" header so the user does not see
  // "Snapshot 1965" when the sheet is plainly named 1965.
  const periodLabel = snapshot && period && period !== String(year)
                        ? `Snapshot ${period}` : `Year ${year}`;
  document.getElementById('rf-sig').textContent = snapshot
    ? `${periodLabel} · schema_year_signature: ${snapshot.signature || '(none)'} · ${allRows.length} source rows · ${groupOrder.length} Research Fields`
    : `Year ${year} · no subfield sheet found in research_field_subfield.xlsx`;

  // ── Chart 1: all Research Fields ──────────────────────────────────
  mk('rf-chart','bar',
     {labels:parents.map(p=>p.name),
      datasets:[{label: snapshot ? 'Total personnel (Σ subfields)' : 'Total personnel (yearly source)',
                 data: parents.map(p=>p.total),
                 backgroundColor: parents.map((p,i)=>c(i, (p.is_aggregate || snapshot) ? .85 : .55))}]},
     co({indexAxis:'y',plugins:{legend:{display:false}}}));

  // ── Chart 2: all Subfields ────────────────────────────────────────
  // Only rendered when a snapshot exists; otherwise the wrapping <div>
  // is hidden so the canvas does not take up vertical space.
  const subWrap = document.getElementById('rf-sub-chart-wrap');
  if (snapshot) {
    subWrap.style.display = '';
    const subRows = [];
    for (const p of parents) for (const child of p.children) subRows.push(child);
    mk('rf-sub-chart','bar',
       {labels: subRows.map(r=>r.subfield),
        datasets:[{label:'Total personnel',
                   data: subRows.map(r=>r.total),
                   backgroundColor: subRows.map((_,i)=>c(i, .8))}]},
       co({indexAxis:'y',plugins:{legend:{display:false}}}));
  } else {
    // Hide the second chart entirely — no canvas, no header — when
    // there is no snapshot data for this year.
    subWrap.style.display = 'none';
    if (CH['rf-sub-chart']) { CH['rf-sub-chart'].destroy(); delete CH['rf-sub-chart']; }
  }

  // ── Hierarchical cards ────────────────────────────────────────────
  if (!parents.length) {
    document.getElementById('rf-blocks').innerHTML =
      '<div class="empty">No Research Field rows for ' + year + '.</div>';
    return;
  }
  const blocks = parents.map(p => {
    const hasChildren = (p.children || []).length > 0;
    const subTable = hasChildren
      ? '<table style="margin-top:8px;"><thead><tr>'
        + '<th>Subfield</th><th>Total</th><th>National Doctor</th><th>Ph.D.</th><th>Year-specific extras</th>'
        + '</tr></thead><tbody>'
        + p.children.map(r => {
            const extras = r.extras && Object.keys(r.extras).length
              ? Object.entries(r.extras).map(([k,v])=>`${esc(k)}: ${fmt(v)}`).join(' · ')
              : '';
            return `<tr>
                <td>${esc(r.subfield)}</td>
                <td class="num">${fmt(r.total)}</td>
                <td class="num">${fmt(r.national_doctor)}</td>
                <td class="num">${fmt(r.phd)}</td>
                <td>${extras}</td>
              </tr>`;
          }).join('')
        + '</tbody></table>'
      : (snapshot
          ? '<div class="empty" style="padding:10px;">No subfield rows under this Research Field.</div>'
          : '<div class="empty" style="padding:10px;">No subfield snapshot for this year — only yearly Field totals available.</div>');

    return `
      <div class="card" style="margin-bottom:12px;">
        <div style="display:flex;align-items:baseline;gap:12px;flex-wrap:wrap;">
          <h3 style="margin:0;">${esc(p.name)}</h3>
          ${p.code ? `<span class="pill">${esc(p.code)}</span>` : ''}
          ${p.ru   ? `<span style="color:var(--dim);font-size:12px;">${esc(p.ru)}</span>` : ''}
          ${snapshot
              ? `<span class="pill pill-a">${p.child_count} subfield${p.child_count===1?'':'s'}</span>`
              : (p.is_aggregate ? '<span class="pill pill-a">aggregate</span>' : '<span class="pill">leaf</span>')}
          <span style="margin-left:auto;font-size:13px;color:var(--dim);">
            ${snapshot ? 'Σ ' : ''}Total <b style="color:var(--accent2);">${fmt(p.total)}</b>
            · ${snapshot ? 'Σ ' : ''}National Doctor <b style="color:var(--accent2);">${fmt(p.national_doctor)}</b>
            · ${snapshot ? 'Σ ' : ''}Ph.D. <b style="color:var(--accent2);">${fmt(p.phd)}</b>
          </span>
        </div>
        <div style="margin-top:6px;font-size:11px;color:var(--dim);">
          ${snapshot
              ? `parent totals = sum of this Research Field's own subfields in <code>research_field_subfield.xlsx · ${esc(period || String(year))}</code>`
              : `no subfield sheet for ${year}; parent totals shown as-is`}
        </div>
        ${subTable}
      </div>`;
  }).join('');

  document.getElementById('rf-blocks').innerHTML = blocks;
}

// ── Party & trainees tab (Party Membership + Research Trainees) ─────────────
INIT.partytrainees = function(){
  const pm = AGG.party;
  mk('pt-counts','line',{labels:pm.map(r=>r.year),datasets:[
    {label:'CPSU members',     data:pm.map(r=>r.cpsu_members), borderColor:c(0), tension:.3, spanGaps:true},
    {label:'Komsomol members', data:pm.map(r=>r.komsomol),     borderColor:c(1), tension:.3, spanGaps:true},
    {label:'Total scientists', data:pm.map(r=>r.total_scientists), borderColor:c(4), tension:.3, spanGaps:true},
  ]}, co({plugins:{legend:{position:'top'}}}));
  mk('pt-pct','line',{labels:pm.map(r=>r.year),datasets:[
    {label:'CPSU %',     data:pm.map(r=>r.cpsu_pct),     borderColor:c(0), tension:.3, spanGaps:true},
    {label:'Komsomol %', data:pm.map(r=>r.komsomol_pct), borderColor:c(1), tension:.3, spanGaps:true},
  ]}, co({plugins:{legend:{position:'top'}},scales:{y:{ticks:{color:'#878fa8',callback:v=>v+'%'},grid:{color:'#2c3145'}}}}));

  const partyHead = '<table><thead><tr>'
    + '<th>Year</th><th>Total Scientists</th>'
    + '<th>CPSU Members</th><th>CPSU %</th>'
    + '<th>Komsomol</th><th>Komsomol %</th>'
    + '<th>Source Material</th>'
    + '</tr></thead><tbody>';
  const partyBody = pm.map(r => `<tr>
      <td>${r.year}</td>
      <td class="num">${fmt(r.total_scientists)}</td>
      <td class="num">${fmt(r.cpsu_members)}</td>
      <td class="num">${fmt(r.cpsu_pct)}</td>
      <td class="num">${fmt(r.komsomol)}</td>
      <td class="num">${fmt(r.komsomol_pct)}</td>
      <td><small style="color:var(--dim);">${esc(r.source_material)}</small></td>
    </tr>`).join('');
  document.getElementById('pt-party-table').innerHTML = partyHead + partyBody + '</tbody></table>';

  const tr = AGG.trainees;
  const trHead = '<table><thead><tr><th>Year</th><th>Total Trainees</th><th>From Other Institutes</th></tr></thead><tbody>';
  const trBody = tr.map(r => `<tr>
      <td>${r.year}</td>
      <td class="num">${fmt(r.total_trainees)}</td>
      <td class="num">${fmt(r.from_other_institutes)}</td>
    </tr>`).join('');
  document.getElementById('pt-trainee-table').innerHTML = trHead + trBody + '</tbody></table>';
};

// ── Provenance tab (Sheets Index) ────────────────────────────────────────────
INIT.provenance = function(){
  const rows0 = AGG.sheets_index || [];
  const fill = (id, vals) => {
    const sel = document.getElementById(id); if (!sel) return;
    [...new Set(vals.filter(Boolean))].sort().forEach(v => {
      const o = document.createElement('option'); o.value=v; o.textContent=v; sel.appendChild(o);
    });
  };
  fill('prv-file',   rows0.map(r=>r.source_file));
  fill('prv-inst',   rows0.map(r=>r.institution));
  fill('prv-role',   rows0.map(r=>r.role));
  fill('prv-target', rows0.map(r=>r.target_table));
  renderProvenance();
};
function roleBadge(role){
  if (!role) return '';
  const cls = role.indexOf('archival') >= 0 ? 'phase-pred'
            : role.indexOf('personnel') >= 0 ? 'phase-cemi'
            : 'phase-pre';
  return `<span class="phase-badge ${cls}" style="margin-left:0;">${esc(role)}</span>`;
}
function renderProvenance(){
  const all = AGG.sheets_index || [];
  if (!all.length){
    document.getElementById('prv-table').innerHTML =
      '<div class="empty" style="padding:24px;">No provenance rows found in <code>sheets_index</code>. '
      + 'Rebuild <code>cemi_career.db</code> with the updated DB builder.</div>';
    return;
  }
  const q  = (document.getElementById('prv-search').value||'').trim().toLowerCase();
  const ff = document.getElementById('prv-file').value;
  const fi = (document.getElementById('prv-inst')||{value:''}).value;
  const fr = (document.getElementById('prv-role')||{value:''}).value;
  const ft = (document.getElementById('prv-target')||{value:''}).value;
  let rows = all;
  if (ff) rows = rows.filter(r=>r.source_file===ff);
  if (fi) rows = rows.filter(r=>r.institution===fi);
  if (fr) rows = rows.filter(r=>r.role===fr);
  if (ft) rows = rows.filter(r=>r.target_table===ft);
  if (q) {
    rows = rows.filter(r =>
      (r.source_file||'').toLowerCase().includes(q) ||
      (r.sheet_name||'').toLowerCase().includes(q) ||
      (r.institution||'').toLowerCase().includes(q) ||
      (r.role||'').toLowerCase().includes(q) ||
      (r.target_table||'').toLowerCase().includes(q) ||
      (r.first_cell_ru||'').toLowerCase().includes(q) ||
      (r.first_cell_en||'').toLowerCase().includes(q));
  }

  // Group counts for the headline banner
  const totalsByRole = {};
  for (const r of all) totalsByRole[r.role||'(unknown)'] = (totalsByRole[r.role||'(unknown)']||0)+1;
  const roleSummary = Object.entries(totalsByRole)
        .map(([k,v]) => `${esc(k)}: <b>${fmt(v)}</b>`).join(' · ');
  const banner = `<div class="results-banner" style="margin-bottom:10px;">
      Showing <b>${fmt(rows.length)}</b> of ${fmt(all.length)} provenance rows
      ${fr?`· role <b>${esc(fr)}</b>`:''}
      ${ff?`· source file <b>${esc(ff)}</b>`:''}
      ${fi?`· institution <b>${esc(fi)}</b>`:''}
      ${ft?`· target table <b>${esc(ft)}</b>`:''}
      ${q?`· text "<b>${esc(q)}</b>"`:''}
      <div style="font-size:11px;margin-top:4px;color:var(--dim);">Overall: ${roleSummary}</div>
    </div>`;

  const head = '<table><thead><tr>'
    + '<th>Year</th><th>Role</th><th>Source File</th><th>Sheet</th>'
    + '<th>Target table</th><th>Institution</th>'
    + '<th>First cell (Russian)</th><th>First cell (English)</th>'
    + '<th>Rows</th><th>Cols</th>'
    + '</tr></thead><tbody>';
  const body = rows.map(r => `<tr>
      <td>${r.year ?? '—'}${r.year != null ? phasePill(r.year) : ''}</td>
      <td>${roleBadge(r.role)}</td>
      <td>${esc(r.source_file)}</td>
      <td>${esc(r.sheet_name)}</td>
      <td>${r.target_table ? `<code>${esc(r.target_table)}</code>` : '<span style="color:var(--dim);">—</span>'}</td>
      <td>${esc(r.institution)}</td>
      <td>${esc(r.first_cell_ru)}</td>
      <td>${esc(r.first_cell_en)}</td>
      <td class="num">${fmt(r.rows)}</td>
      <td class="num">${fmt(r.cols)}</td>
    </tr>`).join('');
  document.getElementById('prv-table').innerHTML = banner + head + body + '</tbody></table>';
}

// ── Subfields tab ─────────────────────────────────────────────
//
// As of the v3 research_field_subfield workbook, sheet names use a
// "YYYY.MM" form (e.g. "1971.01", "1971.06").  The selector lists every
// snapshot ("period") rather than just the year, so 1971's two captures
// remain individually inspectable.
// (INIT.subfields / renderSubfields removed — folded into INIT.research above.)

// ── Glossary tab ──────────────────────────────────────────────
// Populated from the SQLite `glossary` table, which is loaded by the DB
// builder from the "Translation Glossary" sheet of cemi_translation_db.xlsx
// (605 controlled-vocabulary entries spanning position / degree / nationality
// / field-of-science / form-label categories).  The header counter below
// reports the total count so an "empty" glossary tab is unambiguously
// distinguished from a successful load with a narrow filter applied.
INIT.glossary = function(){
  const all = AGG.glossary || [];
  const cats = [...new Set(all.map(r=>r.category).filter(Boolean))].sort();
  const sel = document.getElementById('g-cat');
  for (const c of cats) {
    const o = document.createElement('option'); o.value=c; o.textContent=c; sel.appendChild(o);
  }
  renderGlossary();
};
function renderGlossary(){
  const total = (AGG.glossary || []).length;
  const counterEl = document.getElementById('g-count');
  const q = (document.getElementById('g-search').value||'').trim().toLowerCase();
  const cat = document.getElementById('g-cat').value;
  let rows = AGG.glossary || [];
  if (q) rows = rows.filter(r => (r.russian||'').toLowerCase().includes(q) || (r.english||'').toLowerCase().includes(q));
  if (cat) rows = rows.filter(r => r.category === cat);
  if (counterEl) counterEl.textContent =
    total === 0
      ? 'No glossary entries are loaded — check Translation Glossary sheet in source workbook.'
      : `Showing ${rows.length.toLocaleString()} of ${total.toLocaleString()} controlled-vocabulary entries.`;
  if (total === 0) {
    document.getElementById('g-table').innerHTML =
      '<p style="color:#878fa8;padding:12px">'
      + 'The <code>glossary</code> table is empty.  Re-run the DB builder '
      + 'against a workbook that contains a <b>Translation Glossary</b> sheet '
      + 'with columns <code>#</code>, <code>Russian (original)</code>, '
      + '<code>English (translation)</code>, <code>Category</code>.</p>';
    return;
  }
  const head = '<table><thead><tr><th>Russian</th><th>English</th><th>Category</th></tr></thead><tbody>';
  document.getElementById('g-table').innerHTML = head
    + rows.map(r=>`<tr><td>${esc(r.russian)}</td><td>${esc(r.english)}</td><td><span class="pill">${esc(r.category)}</span></td></tr>`).join('')
    + '</tbody></table>';
}
</script>
</body>
</html>"""


# ──────────────────────────────────────────────────────────────────────────────
#  §3  Render
# ──────────────────────────────────────────────────────────────────────────────

def render(payload: dict, out: Path) -> None:
    s = payload["summary"]
    repls = {
        "__FY__":    str(s.get("first_year") or ""),
        "__LY__":    str(s.get("last_year")  or ""),
        "__NP__":    f"{s.get('total_persons') or 0:,}",
        "__PEAK__":  f"{s.get('peak_scientists') or 0:,}",
        "__PYR__":   str(s.get("peak_year") or ""),
        "__APP__":   f"{s.get('n_appointments') or 0:,}",
        "__DUAL__":  f"{s.get('n_dual_posts') or 0:,}",
        "__PRE__":   f"{s.get('n_pre_cemi') or 0:,}",
        "__TR__":    f"{s.get('n_transfers') or 0:,}",
        "__DIS__":   f"{s.get('n_dismissals') or 0:,}",
        "__PARTY__": f"{s.get('n_party') or 0:,}",
        "__PHD__":   f"{s.get('n_phd') or 0:,}",
    }
    html = HTML
    for k, v in repls.items():
        html = html.replace(k, v)
    payload_str = json.dumps(payload, ensure_ascii=False, default=str)
    html = html.replace("__DATA__", payload_str)
    out.write_text(html, encoding="utf-8")
    print(f"[OK] {out}  ({out.stat().st_size/1024:.1f} KB)")


# ──────────────────────────────────────────────────────────────────────────────
#  §4  CLI
# ──────────────────────────────────────────────────────────────────────────────

def main() -> None:
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("--db",  default=DEFAULTS["db"])
    p.add_argument("--out", default=DEFAULTS["out"])
    args = p.parse_args()

    db_path  = find_db(args.db)
    out_path = resolve_out(args.out)
    print(f"── DB:  {db_path}")
    print(f"── Out: {out_path}")

    conn = sqlite3.connect(db_path)
    payload = build_payload(conn)
    conn.close()
    render(payload, out_path)


if __name__ == "__main__":
    main()
