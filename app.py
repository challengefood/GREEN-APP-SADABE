from __future__ import annotations

import io
import json
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any

import pandas as pd
import streamlit as st

from greenapp import db
from greenapp.db import ACTIVITY_TYPES, PRIORITIES, STATUSES
from greenapp.exporters import activities_to_xlsx, make_budget_workbook, make_leave_docx
from greenapp.parser import analyze_file

APP_DIR = Path(__file__).resolve().parent
LOGO = APP_DIR / "assets" / "logo_sadabe.png"
HERO = APP_DIR / "assets" / "hero_sadabe.jpeg"

st.set_page_config(page_title="GREEN'APP - SADABE", page_icon="🌿", layout="wide")

CUSTOM_CSS = """
<style>
:root { --green:#2E7D32; --light:#E8F5E9; --dark:#1B5E20; --brown:#704023; }
.block-container { padding-top: 1.2rem; padding-bottom: 3rem; }
.main-title { font-size: 2.3rem; font-weight: 800; color: var(--dark); margin-bottom: 0; }
.subtitle { color:#456; font-size: 1rem; margin-top:0; }
.metric-card { background:#ffffff; border:1px solid #dfe8df; border-left:6px solid var(--green); padding:1rem; border-radius:14px; box-shadow:0 2px 10px rgba(0,0,0,0.04); }
.card { background:#ffffff; border:1px solid #e5e7eb; border-radius:16px; padding:1rem; margin-bottom:0.7rem; box-shadow:0 1px 8px rgba(0,0,0,0.03); }
.badge { padding:0.2rem 0.55rem; border-radius:999px; font-size:0.78rem; font-weight:700; }
.badge-red { background:#fee2e2; color:#991b1b; }
.badge-orange { background:#ffedd5; color:#9a3412; }
.badge-green { background:#dcfce7; color:#166534; }
.badge-gray { background:#f3f4f6; color:#374151; }
.small-note { color:#6b7280; font-size:0.86rem; }
hr { margin: 1rem 0; }
</style>
"""
st.markdown(CUSTOM_CSS, unsafe_allow_html=True)


def init() -> None:
    db.init_db()
    st.session_state.setdefault("user", None)


def today() -> date:
    return date.today()


def to_date(value: Any) -> date | None:
    if value in (None, "", pd.NaT):
        return None
    try:
        return pd.to_datetime(value).date()
    except Exception:
        return None


def date_str(value: Any) -> str:
    d = to_date(value)
    return d.isoformat() if d else ""


def money(x: Any) -> str:
    try:
        return f"{float(x):,.0f} Ar".replace(",", " ")
    except Exception:
        return "0 Ar"


def compute_urgency(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df
    out = df.copy()
    urgencies = []
    days_lefts = []
    for _, r in out.iterrows():
        status = str(r.get("status", ""))
        prio = str(r.get("priority", "Normale"))
        d = to_date(r.get("planned_date")) or to_date(r.get("end_date")) or to_date(r.get("start_date"))
        if not d:
            urgencies.append("Sans date")
            days_lefts.append(None)
            continue
        delta = (d - today()).days
        days_lefts.append(delta)
        if status == "Terminée":
            urgencies.append("OK")
        elif delta < 0:
            urgencies.append("En retard")
        elif delta <= 7 or prio in ["Haute", "Critique"]:
            urgencies.append("Urgent")
        else:
            urgencies.append("Normal")
    out["urgence"] = urgencies
    out["jours_restants"] = days_lefts
    return out


def status_badge(value: str) -> str:
    v = value or ""
    if v == "En retard":
        cls = "badge-red"
    elif v == "Urgent":
        cls = "badge-orange"
    elif v in ["OK", "Normal"]:
        cls = "badge-green"
    else:
        cls = "badge-gray"
    return f'<span class="badge {cls}">{v}</span>'


def sidebar_header() -> None:
    with st.sidebar:
        if LOGO.exists():
            st.image(str(LOGO), use_container_width=True)
        st.markdown("### GREEN'APP")
        st.caption("Outil central SADABE : planification, équipe, budget, congé, reporting.")
        if st.session_state.get("user"):
            u = st.session_state["user"]
            st.success(f"Connecté : {u['full_name']}")
            st.caption(f"Rôle : {u['role']}")
            if st.button("Se déconnecter"):
                st.session_state["user"] = None
                st.rerun()


def login_screen() -> None:
    col1, col2 = st.columns([1.2, 1])
    with col1:
        if HERO.exists():
            st.image(str(HERO), use_container_width=True)
        st.markdown("<div class='main-title'>GREEN'APP</div>", unsafe_allow_html=True)
        st.markdown("<p class='subtitle'>Plateforme centrale de planification et de gestion administrative de SADABE.</p>", unsafe_allow_html=True)
        st.info("Chaque utilisateur doit créer un compte. L’accès est possible après validation par l’administrateur.")
    with col2:
        tab1, tab2 = st.tabs(["Connexion", "Créer un compte"])
        with tab1:
            st.subheader("Connexion")
            email = st.text_input("Email", value="admin@sadabe.org")
            password = st.text_input("Mot de passe", type="password", value="")
            if st.button("Se connecter", type="primary", use_container_width=True):
                user, err = db.authenticate(email, password)
                if err:
                    st.error(err)
                else:
                    st.session_state["user"] = user
                    st.rerun()
            st.caption("Compte admin initial : admin@sadabe.org / admin123 — à changer après installation.")
        with tab2:
            st.subheader("Créer un compte")
            name = st.text_input("Nom complet", key="reg_name")
            poste = st.text_input("Poste", key="reg_poste")
            email2 = st.text_input("Email", key="reg_email")
            p1 = st.text_input("Mot de passe", type="password", key="reg_p1")
            p2 = st.text_input("Confirmer le mot de passe", type="password", key="reg_p2")
            if st.button("Envoyer la demande de compte", use_container_width=True):
                if not name or not email2 or len(p1) < 6:
                    st.error("Nom, email et mot de passe de 6 caractères minimum requis.")
                elif p1 != p2:
                    st.error("Les deux mots de passe ne correspondent pas.")
                else:
                    ok, msg = db.register_user(name, email2, poste, p1)
                    st.success(msg) if ok else st.error(msg)


def page_title(title: str, subtitle: str = "") -> None:
    st.markdown(f"<div class='main-title'>{title}</div>", unsafe_allow_html=True)
    if subtitle:
        st.markdown(f"<p class='subtitle'>{subtitle}</p>", unsafe_allow_html=True)
    st.divider()


def select_month(default: date | None = None, key: str = "month") -> tuple[int, int, str, str]:
    d = default or today()
    col1, col2 = st.columns(2)
    with col1:
        year = st.number_input("Année", min_value=2020, max_value=2035, value=d.year, step=1, key=f"{key}_year")
    with col2:
        month = st.selectbox("Mois", list(range(1, 13)), index=d.month - 1, format_func=lambda m: datetime(2000, m, 1).strftime("%B"), key=f"{key}_month")
    start, end = db.get_month_bounds(int(year), int(month))
    return int(year), int(month), start, end


def get_lookup() -> dict[str, Any]:
    return {
        "projects": db.option_map("projects"),
        "partners": db.option_map("partners"),
        "teams": db.option_map("teams"),
        "members": db.members_map(),
    }


def id_to_name(map_: dict[str, int], id_: int | None) -> str | None:
    if not id_:
        return None
    for k, v in map_.items():
        if int(v) == int(id_):
            return k
    return None


def dashboard() -> None:
    page_title("Tableau de bord", "Visualiser les activités du mois, les urgences, les responsables, les projets et les partenaires.")
    look = get_lookup()
    with st.container(border=True):
        st.markdown("#### Filtres")
        c1, c2, c3, c4, c5 = st.columns([1, 1, 1, 1, 1])
        with c1:
            year, month, start, end = select_month(key="dash")
        with c2:
            p_name = st.selectbox("Projet", ["Tous"] + list(look["projects"].keys()))
            p_id = look["projects"].get(p_name) if p_name != "Tous" else None
        with c3:
            partner_name = st.selectbox("Partenaire / bailleur", ["Tous"] + list(look["partners"].keys()))
            partner_id = look["partners"].get(partner_name) if partner_name != "Tous" else None
        with c4:
            member_name = st.selectbox("Personne", ["Tous"] + list(look["members"].keys()))
            member_id = look["members"].get(member_name) if member_name != "Tous" else None
        with c5:
            status = st.selectbox("Statut", ["Tous"] + STATUSES)
    df = db.get_activities_filtered(start, end, p_id, partner_id, member_id, status)
    df = compute_urgency(df)

    total = len(df)
    urgent = int(df["urgence"].isin(["Urgent", "En retard"]).sum()) if not df.empty else 0
    late = int((df["urgence"] == "En retard").sum()) if not df.empty else 0
    missing_desc = int(df["description"].fillna("").eq("").sum()) if not df.empty else 0
    budget = float(df.get("budget_estimated", pd.Series(dtype=float)).fillna(0).sum()) if not df.empty else 0.0
    k1, k2, k3, k4, k5 = st.columns(5)
    k1.metric("Activités du mois", total)
    k2.metric("Urgentes / retard", urgent)
    k3.metric("En retard", late)
    k4.metric("Descriptions manquantes", missing_desc)
    k5.metric("Budget estimé", money(budget))

    if df.empty:
        st.warning("Aucune activité trouvée pour ces filtres. Ajoutez une activité ou importez un document.")
        return

    c1, c2, c3 = st.columns(3)
    with c1:
        st.markdown("#### Par projet")
        st.bar_chart(df.groupby("project").size().rename("Activités"))
    with c2:
        st.markdown("#### Par statut")
        st.bar_chart(df.groupby("status").size().rename("Activités"))
    with c3:
        st.markdown("#### Par urgence")
        st.bar_chart(df.groupby("urgence").size().rename("Activités"))

    st.markdown("#### Planning filtré")
    view_cols = ["urgence", "title", "description", "project", "partner", "planned_date", "end_date", "responsible", "membres", "equipes", "status", "priority", "budget_estimated"]
    view = df[[c for c in view_cols if c in df.columns]].copy()
    st.dataframe(view, use_container_width=True, hide_index=True)

    export = activities_to_xlsx(view)
    st.download_button("⬇️ Exporter ce planning en Excel", export, file_name=f"GREENAPP_planning_{year}_{month:02d}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    st.markdown("#### Points à compléter")
    missing = df[(df["description"].fillna("").eq("")) | (df["responsible"].fillna("").eq("")) | (df["planned_date"].fillna("").eq(""))]
    if missing.empty:
        st.success("Toutes les activités filtrées ont au moins une description, une date et un responsable principal.")
    else:
        st.dataframe(missing[["id", "title", "smart_notes", "project", "planned_date", "responsible"]], use_container_width=True, hide_index=True)


def add_activity_form(prefix: str = "manual") -> None:
    look = get_lookup()
    with st.form(f"{prefix}_add_activity"):
        st.markdown("#### Ajouter une activité manuelle / non intégrée dans un document")
        title = st.text_input("Titre de l’activité *", key=f"{prefix}_title")
        desc = st.text_area("Description détaillée", key=f"{prefix}_desc", height=120)
        c1, c2, c3 = st.columns(3)
        with c1:
            activity_type = st.selectbox("Type", ACTIVITY_TYPES, key=f"{prefix}_type")
            project_names = st.multiselect("Projet(s) concerné(s)", list(look["projects"].keys()), key=f"{prefix}_projects")
            primary_project = st.selectbox("Projet principal", ["Aucun"] + project_names, key=f"{prefix}_primary_project") if project_names else "Aucun"
        with c2:
            partner_names = st.multiselect("Partenaires / bailleurs concernés", list(look["partners"].keys()), key=f"{prefix}_partners")
            location = st.text_input("Lieu / site", key=f"{prefix}_location")
        with c3:
            priority = st.selectbox("Priorité", PRIORITIES, index=1, key=f"{prefix}_priority")
            status = st.selectbox("Statut", STATUSES, key=f"{prefix}_status")
            budget = st.number_input("Budget estimé (Ar)", min_value=0.0, step=1000.0, key=f"{prefix}_budget")
        c4, c5, c6 = st.columns(3)
        with c4:
            start_d = st.date_input("Date début", value=None, key=f"{prefix}_start")
        with c5:
            end_d = st.date_input("Date fin / échéance", value=None, key=f"{prefix}_end")
        with c6:
            plan_d = st.date_input("Jour planifié", value=None, key=f"{prefix}_planned")
        c7, c8 = st.columns(2)
        with c7:
            resp = st.selectbox("Responsable principal", ["Aucun"] + list(look["members"].keys()), key=f"{prefix}_resp")
        with c8:
            team_names = st.multiselect("Équipe(s) responsable(s)", list(look["teams"].keys()), key=f"{prefix}_teams")
        member_names = st.multiselect("Membres SADABE affectés", list(look["members"].keys()), key=f"{prefix}_assigned")
        role_in_activity = st.text_input("Rôle commun des membres ajoutés", value="Membre de l’activité", key=f"{prefix}_role")
        submitted = st.form_submit_button("Enregistrer l’activité", type="primary")
    if submitted:
        if not title.strip():
            st.error("Le titre est obligatoire.")
            return
        primary_project_id = look["projects"].get(primary_project) if primary_project != "Aucun" else (look["projects"].get(project_names[0]) if project_names else None)
        partner_id = look["partners"].get(partner_names[0]) if partner_names else None
        activity_id = db.insert_activity(
            {
                "title": title,
                "description": desc,
                "activity_type": activity_type,
                "project_id": primary_project_id,
                "partner_id": partner_id,
                "location": location,
                "start_date": start_d.isoformat() if start_d else None,
                "end_date": end_d.isoformat() if end_d else None,
                "planned_date": plan_d.isoformat() if plan_d else None,
                "status": status,
                "priority": priority,
                "responsible_member_id": look["members"].get(resp) if resp != "Aucun" else None,
                "budget_estimated": budget,
                "smart_score": 100,
                "smart_notes": "Saisie manuelle complète" if desc and plan_d and resp != "Aucun" else "Saisie manuelle à compléter",
                "created_by": st.session_state["user"]["id"],
            },
            project_ids=[look["projects"][n] for n in project_names],
            partner_ids=[look["partners"][n] for n in partner_names],
        )
        with db.get_conn() as conn:
            for m in member_names:
                conn.execute("INSERT OR IGNORE INTO activity_assignments(activity_id, member_id, role_in_activity, is_lead) VALUES (?, ?, ?, 0)", (activity_id, look["members"][m], role_in_activity))
            for t in team_names:
                conn.execute("INSERT OR IGNORE INTO activity_teams(activity_id, team_id, role_in_activity) VALUES (?, ?, ?)", (activity_id, look["teams"][t], "Équipe responsable"))
        st.success("Activité ajoutée avec responsables, membres et équipes.")
        st.rerun()


def page_import_add() -> None:
    page_title("Ajouter / Importer des activités", "Import intelligent Excel, CSV ou Word : extraction titre, description, dates, projet, partenaire, responsable et budget.")
    tab1, tab2 = st.tabs(["Import intelligent", "Saisie manuelle"])
    with tab1:
        look = get_lookup()
        uploaded = st.file_uploader("Importer un fichier Excel, CSV ou Word", type=["xlsx", "xls", "csv", "docx"])
        if uploaded is not None:
            file_bytes = uploaded.getvalue()
            with st.spinner("Analyse intelligente du document en cours..."):
                rows, summary = analyze_file(file_bytes, uploaded.name, look["projects"], look["partners"], look["members"])
            st.markdown("#### Résumé automatique")
            c1, c2, c3, c4, c5 = st.columns(5)
            c1.metric("Activités détectées", summary["total_detected"])
            c2.metric("Score moyen", f"{summary['avg_smart_score']} %")
            c3.metric("Sans description", summary["missing_description"])
            c4.metric("Sans date", summary["missing_date"])
            c5.metric("Sans responsable", summary["missing_responsible"])
            if not rows:
                st.error("Aucune activité détectée. Vérifiez que le document contient un tableau ou des paragraphes d’activités.")
                return
            df = pd.DataFrame(rows)
            display_cols = ["title", "description", "project_detected", "partner_detected", "responsible_text", "planned_date", "start_date", "end_date", "priority", "status", "budget_estimated", "smart_score", "smart_notes"]
            edited = st.data_editor(
                df[display_cols],
                use_container_width=True,
                hide_index=True,
                num_rows="dynamic",
                column_config={
                    "description": st.column_config.TextColumn(width="large"),
                    "title": st.column_config.TextColumn(width="medium"),
                    "budget_estimated": st.column_config.NumberColumn(format="%d Ar"),
                },
            )
            st.warning("Après l’import, les responsables inconnus restent dans la colonne 'responsible_text'. Ajoutez-les ensuite dans Équipe & responsables ou affectez-les dans la planification.")
            if st.button("Enregistrer les activités analysées", type="primary"):
                saved = 0
                for i, row in edited.iterrows():
                    if not str(row.get("title", "")).strip():
                        continue
                    original = rows[i] if i < len(rows) else {}
                    payload = dict(original)
                    for c in display_cols:
                        payload[c] = row.get(c)
                    payload["created_by"] = st.session_state["user"]["id"]
                    db.insert_activity(payload, project_ids=[payload["project_id"]] if payload.get("project_id") else [], partner_ids=[payload["partner_id"]] if payload.get("partner_id") else [])
                    saved += 1
                db.execute(
                    "INSERT INTO import_batches(filename, imported_by, total_rows, saved_rows, analysis_summary) VALUES (?, ?, ?, ?, ?)",
                    (uploaded.name, st.session_state["user"]["id"], len(rows), saved, json.dumps(summary, ensure_ascii=False)),
                )
                st.success(f"{saved} activité(s) enregistrée(s).")
                st.rerun()
    with tab2:
        add_activity_form("manual_page")


def page_planification() -> None:
    page_title("Planification du mois", "Voir les activités à faire ce mois, ajouter des activités, affecter équipes, membres, tâches, jours et responsables.")
    look = get_lookup()
    year, month, start, end = select_month(key="plan")
    with st.expander("➕ Ajouter une activité directement dans la planification du mois", expanded=False):
        add_activity_form("plan")
    st.markdown("#### Activités du mois")
    df = compute_urgency(db.get_activities_filtered(start, end))
    if df.empty:
        st.info("Aucune activité ce mois. Ajoutez une activité ci-dessus ou importez un document.")
        return
    for _, row in df.iterrows():
        with st.container(border=True):
            c1, c2, c3 = st.columns([4, 1.2, 1.2])
            with c1:
                st.markdown(f"### #{int(row['id'])} — {row['title']}")
                st.markdown(status_badge(row.get("urgence", "")), unsafe_allow_html=True)
                st.write(row.get("description") or "_Aucune description. À compléter._")
                st.caption(f"Projet : {row.get('project') or '—'} | Partenaire : {row.get('partner') or row.get('partenaires_associes') or '—'} | Lieu : {row.get('location') or '—'}")
            with c2:
                st.metric("Statut", row.get("status") or "—")
                st.metric("Priorité", row.get("priority") or "—")
            with c3:
                st.metric("Jour planifié", row.get("planned_date") or "—")
                st.metric("Responsable", row.get("responsible") or "—")
            with st.expander("Planifier / affecter cette activité", expanded=False):
                edit_activity_block(int(row["id"]), look)
            tasks = db.query_df(
                """
                SELECT t.id, t.task_title, t.task_description, t.task_date, t.status, m.full_name AS responsable, t.notes
                FROM activity_tasks t LEFT JOIN members m ON m.id=t.responsible_member_id
                WHERE t.activity_id=? ORDER BY COALESCE(t.task_date, '9999-12-31'), t.id
                """,
                (int(row["id"]),),
            )
            if not tasks.empty:
                st.markdown("**Tâches détaillées**")
                st.dataframe(tasks, use_container_width=True, hide_index=True)


def edit_activity_block(activity_id: int, look: dict[str, Any]) -> None:
    act = db.get_activity_full(activity_id)
    if not act:
        st.error("Activité introuvable.")
        return
    current_member_ids = db.query_df("SELECT member_id FROM activity_assignments WHERE activity_id=?", (activity_id,))["member_id"].tolist()
    current_team_ids = db.query_df("SELECT team_id FROM activity_teams WHERE activity_id=?", (activity_id,))["team_id"].tolist()
    member_options = list(look["members"].keys())
    team_options = list(look["teams"].keys())
    current_members = [name for name, _id in look["members"].items() if _id in current_member_ids]
    current_teams = [name for name, _id in look["teams"].items() if _id in current_team_ids]
    resp_current = id_to_name(look["members"], act.get("responsible_member_id")) or "Aucun"
    with st.form(f"edit_activity_{activity_id}"):
        c1, c2, c3 = st.columns(3)
        with c1:
            plan_d = st.date_input("Jour planifié", value=to_date(act.get("planned_date")), key=f"pld_{activity_id}")
            status = st.selectbox("Statut", STATUSES, index=STATUSES.index(act.get("status")) if act.get("status") in STATUSES else 0, key=f"sts_{activity_id}")
        with c2:
            priority = st.selectbox("Priorité", PRIORITIES, index=PRIORITIES.index(act.get("priority")) if act.get("priority") in PRIORITIES else 1, key=f"prio_{activity_id}")
            responsible = st.selectbox("Responsable principal", ["Aucun"] + member_options, index=(["Aucun"] + member_options).index(resp_current) if resp_current in (["Aucun"] + member_options) else 0, key=f"resp_{activity_id}")
        with c3:
            budget = st.number_input("Budget estimé (Ar)", min_value=0.0, value=float(act.get("budget_estimated") or 0), step=1000.0, key=f"bud_{activity_id}")
            location = st.text_input("Lieu", value=act.get("location") or "", key=f"loc_{activity_id}")
        desc = st.text_area("Description détaillée", value=act.get("description") or "", height=120, key=f"desc_{activity_id}")
        assigned = st.multiselect("Membres SADABE affectés à l’activité", member_options, default=current_members, key=f"ass_{activity_id}")
        teams = st.multiselect("Équipes responsables", team_options, default=current_teams, key=f"teams_{activity_id}")
        role_common = st.text_input("Rôle des membres nouvellement ajoutés", value="Membre de l’activité", key=f"rolec_{activity_id}")
        saved = st.form_submit_button("Mettre à jour l’activité", type="primary")
    if saved:
        with db.get_conn() as conn:
            conn.execute(
                """
                UPDATE activities SET planned_date=?, status=?, priority=?, responsible_member_id=?, budget_estimated=?, location=?, description=?, updated_at=CURRENT_TIMESTAMP
                WHERE id=?
                """,
                (plan_d.isoformat() if plan_d else None, status, priority, look["members"].get(responsible) if responsible != "Aucun" else None, budget, location, desc, activity_id),
            )
            conn.execute("DELETE FROM activity_assignments WHERE activity_id=?", (activity_id,))
            for m in assigned:
                conn.execute("INSERT OR IGNORE INTO activity_assignments(activity_id, member_id, role_in_activity, is_lead) VALUES (?, ?, ?, ?)", (activity_id, look["members"][m], role_common, 1 if m == responsible else 0))
            conn.execute("DELETE FROM activity_teams WHERE activity_id=?", (activity_id,))
            for t in teams:
                conn.execute("INSERT OR IGNORE INTO activity_teams(activity_id, team_id, role_in_activity) VALUES (?, ?, ?)", (activity_id, look["teams"][t], "Équipe responsable"))
        st.success("Activité mise à jour.")
        st.rerun()

    with st.form(f"add_task_{activity_id}"):
        st.markdown("##### Ajouter une tâche détaillée")
        t1, t2, t3 = st.columns(3)
        with t1:
            task_title = st.text_input("Tâche", key=f"tasktitle_{activity_id}")
            task_date = st.date_input("Jour de la tâche", value=to_date(act.get("planned_date")) or today(), key=f"taskdate_{activity_id}")
        with t2:
            task_resp = st.selectbox("Responsable de la tâche", ["Aucun"] + member_options, key=f"taskresp_{activity_id}")
            task_status = st.selectbox("Statut tâche", ["À faire", "En cours", "Terminée", "Reportée"], key=f"taskstatus_{activity_id}")
        with t3:
            task_notes = st.text_area("Notes", key=f"tasknotes_{activity_id}")
        task_desc = st.text_area("Description de la tâche", key=f"taskdesc_{activity_id}")
        add_task = st.form_submit_button("Ajouter la tâche")
    if add_task:
        if task_title.strip():
            db.execute(
                "INSERT INTO activity_tasks(activity_id, task_title, task_description, task_date, responsible_member_id, status, notes) VALUES (?, ?, ?, ?, ?, ?, ?)",
                (activity_id, task_title, task_desc, task_date.isoformat(), look["members"].get(task_resp) if task_resp != "Aucun" else None, task_status, task_notes),
            )
            st.success("Tâche ajoutée.")
            st.rerun()
        else:
            st.error("Le titre de la tâche est obligatoire.")


def page_team() -> None:
    page_title("Équipe & responsables", "Ajouter les membres SADABE, leurs postes, et constituer des équipes responsables.")
    tab1, tab2, tab3 = st.tabs(["Membres SADABE", "Équipes", "Affectation membres → équipes"])
    with tab1:
        with st.form("add_member"):
            st.markdown("#### Ajouter un membre")
            c1, c2, c3 = st.columns(3)
            with c1:
                full_name = st.text_input("Nom complet")
                poste = st.text_input("Poste")
            with c2:
                email = st.text_input("Email")
                phone = st.text_input("Téléphone")
            with c3:
                team_default = st.text_input("Équipe par défaut")
                notes = st.text_area("Notes")
            submitted = st.form_submit_button("Ajouter le membre", type="primary")
        if submitted:
            if not full_name:
                st.error("Nom obligatoire.")
            else:
                try:
                    db.execute("INSERT INTO members(full_name, poste, email, phone, team_default, notes) VALUES (?, ?, ?, ?, ?, ?)", (full_name, poste, email, phone, team_default, notes))
                    st.success("Membre ajouté.")
                    st.rerun()
                except Exception as e:
                    st.error(f"Erreur : {e}")
        members = db.query_df("SELECT id, full_name, poste, email, phone, team_default, active, notes FROM members ORDER BY full_name")
        st.dataframe(members, use_container_width=True, hide_index=True)
    with tab2:
        with st.form("add_team"):
            name = st.text_input("Nom de l’équipe")
            desc = st.text_area("Description / rôle de l’équipe")
            submitted = st.form_submit_button("Ajouter l’équipe", type="primary")
        if submitted:
            if not name:
                st.error("Nom obligatoire.")
            else:
                db.execute("INSERT OR IGNORE INTO teams(name, description) VALUES (?, ?)", (name, desc))
                st.success("Équipe ajoutée.")
                st.rerun()
        teams = db.query_df("SELECT id, name, description, active FROM teams ORDER BY name")
        st.dataframe(teams, use_container_width=True, hide_index=True)
    with tab3:
        look = get_lookup()
        if not look["members"] or not look["teams"]:
            st.info("Ajoutez au moins un membre et une équipe.")
        else:
            with st.form("team_members"):
                team_name = st.selectbox("Équipe", list(look["teams"].keys()))
                member_names = st.multiselect("Membres", list(look["members"].keys()))
                role = st.text_input("Rôle dans l’équipe", value="Membre")
                submitted = st.form_submit_button("Ajouter à l’équipe", type="primary")
            if submitted:
                with db.get_conn() as conn:
                    for m in member_names:
                        conn.execute("INSERT OR REPLACE INTO team_members(team_id, member_id, role_in_team) VALUES (?, ?, ?)", (look["teams"][team_name], look["members"][m], role))
                st.success("Affectation enregistrée.")
                st.rerun()
            df = db.query_df(
                """
                SELECT t.name AS equipe, m.full_name AS membre, m.poste, tm.role_in_team
                FROM team_members tm
                JOIN teams t ON t.id=tm.team_id
                JOIN members m ON m.id=tm.member_id
                ORDER BY t.name, m.full_name
                """
            )
            st.dataframe(df, use_container_width=True, hide_index=True)


def page_projects_partners() -> None:
    page_title("Projets & partenaires", "Gérer les projets, bailleurs et partenaires utilisables dans la planification.")
    tab1, tab2 = st.tabs(["Projets", "Partenaires / bailleurs"])
    with tab1:
        with st.form("add_project"):
            name = st.text_input("Nom du projet")
            desc = st.text_area("Description")
            submitted = st.form_submit_button("Ajouter le projet", type="primary")
        if submitted and name:
            db.execute("INSERT OR IGNORE INTO projects(name, description) VALUES (?, ?)", (name, desc))
            st.success("Projet ajouté.")
            st.rerun()
        st.dataframe(db.query_df("SELECT id, name, description, active FROM projects ORDER BY name"), use_container_width=True, hide_index=True)
    with tab2:
        with st.form("add_partner"):
            name = st.text_input("Nom partenaire / bailleur")
            typ = st.selectbox("Type", ["Partenaire", "Bailleur", "Université", "Administration", "Communauté", "Autre"])
            desc = st.text_area("Description")
            submitted = st.form_submit_button("Ajouter partenaire / bailleur", type="primary")
        if submitted and name:
            db.execute("INSERT OR IGNORE INTO partners(name, description, type) VALUES (?, ?, ?)", (name, desc, typ))
            st.success("Partenaire ajouté.")
            st.rerun()
        st.dataframe(db.query_df("SELECT id, name, type, description, active FROM partners ORDER BY name"), use_container_width=True, hide_index=True)


def page_budget() -> None:
    page_title("Demande de budget", "Créer automatiquement un canevas financier SADABE avec calculs : total, avances et reste à percevoir.")
    look = get_lookup()
    activities = db.query_df("SELECT id, title FROM activities ORDER BY id DESC LIMIT 500")
    activity_opts = {f"#{r.id} — {r.title}": int(r.id) for _, r in activities.iterrows()} if not activities.empty else {}
    with st.form("budget_request_form"):
        st.markdown("#### Informations mission / activité")
        c0, c1, c2 = st.columns(3)
        with c0:
            linked = st.selectbox("Activité liée", ["Aucune"] + list(activity_opts.keys()))
            title_default = ""
            if linked != "Aucune":
                a = db.get_activity_full(activity_opts[linked])
                title_default = a.get("title", "") if a else ""
            title = st.text_input("Intitulé de la mission / demande", value=title_default)
        with c1:
            beneficiary = st.selectbox("Bénéficiaire", ["Aucun"] + list(look["members"].keys()))
            responsible = st.selectbox("Responsable / validateur", ["Aucun"] + list(look["members"].keys()))
        with c2:
            project = st.selectbox("Projet", ["Aucun"] + list(look["projects"].keys()))
            partner = st.selectbox("Bailleur / partenaire", ["Aucun"] + list(look["partners"].keys()))
        c3, c4, c5 = st.columns(3)
        with c3:
            location = st.text_input("Lieu(x)")
        with c4:
            date_depart = st.date_input("Date départ", value=today())
        with c5:
            date_retour = st.date_input("Date retour", value=today() + timedelta(days=1))
        justification = st.text_area("Justification / contexte de la demande")
        default_days = max((date_retour - date_depart).days + 1, 1)
        default_lines = pd.DataFrame([
            {"rubrique": "Per diem", "base_calcul": "Nb jours x taux journalier", "qty": default_days, "unit_rate": 40000, "total": 0, "received_before": 0, "received_during": 0, "comment": ""},
            {"rubrique": "Hébergement", "base_calcul": "Nb nuitées x coût unitaire", "qty": 0, "unit_rate": 0, "total": 0, "received_before": 0, "received_during": 0, "comment": ""},
            {"rubrique": "Déjeuner route", "base_calcul": "Nb repas x coût unitaire", "qty": 0, "unit_rate": 0, "total": 0, "received_before": 0, "received_during": 0, "comment": ""},
            {"rubrique": "Carburant", "base_calcul": "Distance / trajet / litres", "qty": 0, "unit_rate": 4900, "total": 0, "received_before": 0, "received_during": 0, "comment": ""},
            {"rubrique": "Frais de déplacement", "base_calcul": "Transport / billets / taxi / moto", "qty": 0, "unit_rate": 0, "total": 0, "received_before": 0, "received_during": 0, "comment": ""},
            {"rubrique": "Dépense imprévue", "base_calcul": "Montant exceptionnel justifié", "qty": 0, "unit_rate": 0, "total": 0, "received_before": 0, "received_during": 0, "comment": ""},
            {"rubrique": "Autres frais", "base_calcul": "À préciser", "qty": 0, "unit_rate": 0, "total": 0, "received_before": 0, "received_during": 0, "comment": ""},
        ])
        st.markdown("#### Lignes budgétaires")
        lines_df = st.data_editor(default_lines, num_rows="dynamic", use_container_width=True, hide_index=True)
        submitted = st.form_submit_button("Calculer et générer le canevas financier", type="primary")
    if submitted:
        if not title:
            st.error("L’intitulé est obligatoire.")
            return
        lines = lines_df.fillna(0).to_dict(orient="records")
        clean_lines = []
        total_needs = received_before = received_during = 0.0
        for line in lines:
            qty = float(line.get("qty") or 0)
            rate = float(line.get("unit_rate") or 0)
            custom_total = float(line.get("total") or 0)
            total = qty * rate if qty and rate else custom_total
            rb = float(line.get("received_before") or 0)
            rd = float(line.get("received_during") or 0)
            rem = total - rb - rd
            clean = dict(line)
            clean.update({"qty": qty, "unit_rate": rate, "total": total, "received_before": rb, "received_during": rd, "remaining": rem})
            clean_lines.append(clean)
            total_needs += total
            received_before += rb
            received_during += rd
        remaining = total_needs - received_before - received_during
        req = {
            "title": title,
            "activity_id": activity_opts.get(linked) if linked != "Aucune" else None,
            "beneficiary_member_id": look["members"].get(beneficiary) if beneficiary != "Aucun" else None,
            "beneficiary_name": beneficiary.split(" — ")[0] if beneficiary != "Aucun" else "",
            "responsible_member_id": look["members"].get(responsible) if responsible != "Aucun" else None,
            "responsible_name": responsible.split(" — ")[0] if responsible != "Aucun" else "",
            "project_id": look["projects"].get(project) if project != "Aucun" else None,
            "project_name": project if project != "Aucun" else "",
            "partner_id": look["partners"].get(partner) if partner != "Aucun" else None,
            "partner_name": partner if partner != "Aucun" else "",
            "location": location,
            "date_depart": date_depart.isoformat(),
            "date_retour": date_retour.isoformat(),
            "nb_days": max((date_retour - date_depart).days + 1, 1),
            "total_needs": total_needs,
            "received_before": received_before,
            "received_during": received_during,
            "remaining": remaining,
            "justification": justification,
            "created_by": st.session_state["user"]["id"],
        }
        with db.get_conn() as conn:
            cur = conn.execute(
                """
                INSERT INTO budget_requests(activity_id, title, beneficiary_member_id, responsible_member_id, project_id, partner_id, location, date_depart, date_retour, nb_days, total_needs, received_before, received_during, remaining, justification, created_by)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (req["activity_id"], req["title"], req["beneficiary_member_id"], req["responsible_member_id"], req["project_id"], req["partner_id"], req["location"], req["date_depart"], req["date_retour"], req["nb_days"], total_needs, received_before, received_during, remaining, justification, req["created_by"]),
            )
            req_id = cur.lastrowid
            for line in clean_lines:
                if not str(line.get("rubrique", "")).strip():
                    continue
                conn.execute(
                    """
                    INSERT INTO budget_lines(request_id, rubrique, base_calcul, qty, unit_rate, total, received_before, received_during, remaining, comment)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (req_id, line.get("rubrique"), line.get("base_calcul"), line.get("qty"), line.get("unit_rate"), line.get("total"), line.get("received_before"), line.get("received_during"), line.get("remaining"), line.get("comment")),
                )
        st.success(f"Demande de budget enregistrée. Total : {money(total_needs)} | Reste à percevoir : {money(remaining)}")
        xlsx = make_budget_workbook(req, clean_lines)
        st.download_button("⬇️ Télécharger le canevas financier rempli", xlsx, file_name=f"demande_budget_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    st.markdown("#### Historique des demandes")
    hist = db.query_df(
        """
        SELECT br.id, br.title, m.full_name AS beneficiaire, p.name AS projet, pa.name AS partenaire, br.date_depart, br.date_retour, br.total_needs, br.remaining, br.status, br.created_at
        FROM budget_requests br
        LEFT JOIN members m ON m.id=br.beneficiary_member_id
        LEFT JOIN projects p ON p.id=br.project_id
        LEFT JOIN partners pa ON pa.id=br.partner_id
        ORDER BY br.id DESC LIMIT 50
        """
    )
    st.dataframe(hist, use_container_width=True, hide_index=True)


def page_leave() -> None:
    page_title("Demande de congé", "Remplir un canevas RH prêt à envoyer : motif, date du... au..., intérim, contact et responsable.")
    look = get_lookup()
    with st.form("leave_form"):
        c1, c2, c3 = st.columns(3)
        with c1:
            member = st.selectbox("Membre SADABE", ["Saisie manuelle"] + list(look["members"].keys()))
            manual_name = st.text_input("Nom complet si saisie manuelle") if member == "Saisie manuelle" else ""
        with c2:
            leave_type = st.selectbox("Type de congé", ["Congé annuel", "Congé exceptionnel", "Congé maladie", "Congé sans solde", "Permission", "Autre"])
            motif = st.text_area("Motif")
        with c3:
            start_d = st.date_input("Date du", value=today())
            end_d = st.date_input("Au", value=today() + timedelta(days=1))
        c4, c5, c6 = st.columns(3)
        with c4:
            poste_manual = st.text_input("Poste si saisie manuelle") if member == "Saisie manuelle" else ""
            interim = st.text_input("Personne assurant l’intérim")
        with c5:
            contact = st.text_input("Contact pendant le congé")
        with c6:
            manager = st.text_input("Responsable hiérarchique / RH")
        submitted = st.form_submit_button("Générer la demande de congé", type="primary")
    if submitted:
        nb_days = max((end_d - start_d).days + 1, 1)
        if member == "Saisie manuelle":
            full_name = manual_name
            poste = poste_manual
            member_id = None
        else:
            full_name = member.split(" — ")[0]
            poste = member.split(" — ")[1] if " — " in member else ""
            member_id = look["members"].get(member)
        if not full_name:
            st.error("Le nom est obligatoire.")
            return
        payload = {
            "member_id": member_id,
            "full_name": full_name,
            "poste": poste,
            "leave_type": leave_type,
            "motif": motif,
            "start_date": start_d.isoformat(),
            "end_date": end_d.isoformat(),
            "nb_days": nb_days,
            "interim": interim,
            "contact": contact,
            "manager": manager,
            "created_by": st.session_state["user"]["id"],
        }
        db.execute(
            """
            INSERT INTO leave_requests(member_id, full_name, poste, leave_type, motif, start_date, end_date, nb_days, interim, contact, manager, created_by)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (payload["member_id"], payload["full_name"], payload["poste"], payload["leave_type"], payload["motif"], payload["start_date"], payload["end_date"], payload["nb_days"], payload["interim"], payload["contact"], payload["manager"], payload["created_by"]),
        )
        docx_bytes = make_leave_docx(payload)
        st.success(f"Demande de congé générée pour {full_name} — {nb_days} jour(s).")
        st.download_button("⬇️ Télécharger la demande de congé Word", docx_bytes, file_name=f"demande_conge_{full_name.replace(' ', '_')}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
    st.markdown("#### Historique")
    hist = db.query_df("SELECT id, full_name, poste, leave_type, motif, start_date, end_date, nb_days, manager, status, created_at FROM leave_requests ORDER BY id DESC LIMIT 50")
    st.dataframe(hist, use_container_width=True, hide_index=True)


def page_admin() -> None:
    page_title("Administration", "Validation des comptes et sauvegarde de la base de données.")
    user = st.session_state["user"]
    if user["role"] != "admin":
        st.error("Accès réservé à l’administrateur.")
        return
    tab1, tab2, tab3 = st.tabs(["Comptes", "Sauvegarde", "Journal imports"])
    with tab1:
        users = db.query_df("SELECT id, full_name, email, poste, role, status, created_at, last_login FROM users ORDER BY created_at DESC")
        st.dataframe(users, use_container_width=True, hide_index=True)
        st.markdown("#### Modifier un compte")
        if not users.empty:
            user_choice = st.selectbox("Compte", [f"#{r.id} — {r.full_name} — {r.email}" for _, r in users.iterrows()])
            uid = int(user_choice.split(" — ")[0].replace("#", ""))
            c1, c2 = st.columns(2)
            with c1:
                new_status = st.selectbox("Statut", ["pending", "approved", "blocked"])
            with c2:
                new_role = st.selectbox("Rôle", ["admin", "manager", "staff", "lecture"])
            if st.button("Mettre à jour le compte", type="primary"):
                db.execute("UPDATE users SET status=?, role=? WHERE id=?", (new_status, new_role, uid))
                st.success("Compte mis à jour.")
                st.rerun()
    with tab2:
        db_path = Path(db.DB_PATH)
        if db_path.exists():
            st.download_button("⬇️ Télécharger la base SQLite", db_path.read_bytes(), file_name="greenapp_backup.db", mime="application/octet-stream")
        st.info("Pour un vrai usage multi-utilisateur permanent sur Streamlit Cloud, utilisez une base externe comme PostgreSQL/Supabase. SQLite suffit pour un prototype interne.")
    with tab3:
        imports = db.query_df("SELECT * FROM import_batches ORDER BY id DESC LIMIT 100")
        st.dataframe(imports, use_container_width=True, hide_index=True)


def main() -> None:
    init()
    sidebar_header()
    if not st.session_state.get("user"):
        login_screen()
        return
    pages = [
        "Tableau de bord",
        "Planification du mois",
        "Ajouter / Importer activités",
        "Équipe & responsables",
        "Projets & partenaires",
        "Demande de budget",
        "Demande de congé",
    ]
    if st.session_state["user"]["role"] == "admin":
        pages.append("Administration")
    with st.sidebar:
        page = st.radio("Navigation", pages)
    if page == "Tableau de bord":
        dashboard()
    elif page == "Planification du mois":
        page_planification()
    elif page == "Ajouter / Importer activités":
        page_import_add()
    elif page == "Équipe & responsables":
        page_team()
    elif page == "Projets & partenaires":
        page_projects_partners()
    elif page == "Demande de budget":
        page_budget()
    elif page == "Demande de congé":
        page_leave()
    elif page == "Administration":
        page_admin()


if __name__ == "__main__":
    main()
