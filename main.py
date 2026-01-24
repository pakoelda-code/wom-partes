# -*- coding: utf-8 -*-
# WOM - Partes de Mantenimiento (WEB) - Supabase (Postgres)
#
# Requisitos (requirements.txt):
# fastapi
# uvicorn
# python-multipart
# itsdangerous
# reportlab
# psycopg2-binary
#
# Variables de entorno:
# DATABASE_URL   (Supabase Pooler, p.ej. ...pooler.supabase.com:6543/postgres)
# SESSION_SECRET (recomendado)

import os
import random
import string
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple
from zoneinfo import ZoneInfo

import psycopg2
from psycopg2.extras import RealDictCursor

from fastapi import FastAPI, Request, Form
from fastapi.responses import HTMLResponse, RedirectResponse, FileResponse, PlainTextResponse
from starlette.middleware.sessions import SessionMiddleware

TZ = ZoneInfo("Europe/Madrid")

TIPOS = ["ELECTRÓNICA", "MOBILIARIO", "ESTRUCTURA", "ELEMENTOS SUELTOS", "OTROS/AS"]

ESTADOS_ENCARGADO = [
    "SIN ESTADO",
    "TRABAJO PENDIENTE/EN COLA",
    "TRABAJO EN PROCESO",
    "TRABAJO TERMINADO/REPARADO",
    "TRABAJO DESESTIMADO",
]
ESTADOS_FINALIZADOS = {"TRABAJO TERMINADO/REPARADO", "TRABAJO DESESTIMADO"}

ALL_MARKER = "__TODAS__"  # valor especial en multiselect


# =========================
# DB (Supabase Postgres)
# =========================
DATABASE_URL = os.getenv("DATABASE_URL", "").strip()


def _ensure_db_url() -> str:
    if not DATABASE_URL:
        raise RuntimeError("Falta DATABASE_URL en variables de entorno")
    return DATABASE_URL


def db_conn():
    url = _ensure_db_url()
    if "sslmode=" not in url:
        return psycopg2.connect(url, cursor_factory=RealDictCursor, sslmode="require")
    return psycopg2.connect(url, cursor_factory=RealDictCursor)


def db_all(sql: str, params=()) -> List[Dict[str, Any]]:
    with db_conn() as conn:
        with conn.cursor() as cur:
            cur.execute(sql, params)
            rows = cur.fetchall()
            return list(rows or [])


def db_one(sql: str, params=()) -> Optional[Dict[str, Any]]:
    rows = db_all(sql, params)
    return rows[0] if rows else None


def db_exec(sql: str, params=()) -> None:
    with db_conn() as conn:
        with conn.cursor() as cur:
            cur.execute(sql, params)
        conn.commit()


def ensure_schema_and_seed() -> None:
    db_exec(
        """
    create table if not exists public.wom_users (
      code text primary key,
      name text not null,
      role text not null check (role in ('TRABAJADOR','ENCARGADO','JEFE')),
      created_at timestamptz not null default now()
    );
    """
    )

    db_exec(
        """
    create table if not exists public.wom_rooms (
      id bigserial primary key,
      name text not null unique,
      created_at timestamptz not null default now()
    );
    """
    )

    db_exec(
        """
    create table if not exists public.wom_tickets (
      id bigserial primary key,
      referencia char(6) not null unique,

      created_at timestamptz not null default now(),
      created_by_code text references public.wom_users(code) on delete set null,
      created_by_name text not null,

      room_id bigint references public.wom_rooms(id) on delete set null,
      room_name text not null,

      tipo text not null check (tipo in ('ELECTRÓNICA','MOBILIARIO','ESTRUCTURA','ELEMENTOS SUELTOS','OTROS/AS')),
      descripcion text not null,

      solucionado_por_usuario boolean not null default false,
      reparacion_usuario text not null default '',

      visto_por_encargado boolean not null default false,
      estado_encargado text not null default 'SIN ESTADO' check (
        estado_encargado in (
          'SIN ESTADO',
          'TRABAJO PENDIENTE/EN COLA',
          'TRABAJO EN PROCESO',
          'TRABAJO TERMINADO/REPARADO',
          'TRABAJO DESESTIMADO'
        )
      ),
      observaciones_encargado text not null default '',

      updated_at timestamptz not null default now()
    );
    """
    )

    db_exec(
        "create index if not exists wom_tickets_created_at_idx on public.wom_tickets(created_at desc);"
    )
    db_exec(
        "create index if not exists wom_tickets_estado_idx on public.wom_tickets(estado_encargado);"
    )
    db_exec(
        "create index if not exists wom_tickets_user_idx on public.wom_tickets(created_by_code);"
    )
    db_exec(
        "create index if not exists wom_tickets_room_idx on public.wom_tickets(room_name);"
    )

    count_users = db_one("select count(*)::int as n from public.wom_users;")
    if count_users and count_users["n"] == 0:
        db_exec(
            """
        insert into public.wom_users (code, name, role) values
        ('P000A','Pako','ENCARGADO'),
        ('I001A','Isa','TRABAJADOR'),
        ('J002R','Javi','TRABAJADOR'),
        ('A003N','Adrián','TRABAJADOR'),
        ('D004I','Dani','TRABAJADOR'),
        ('C005S','Carlos','TRABAJADOR'),
        ('P006O','Pacardo','TRABAJADOR'),
        ('R007A','Rebeca','TRABAJADOR'),
        ('M001X','Manu','JEFE'),
        ('L002X','Luis','JEFE')
        on conflict (code) do nothing;
        """
        )

    count_rooms = db_one("select count(*)::int as n from public.wom_rooms;")
    if count_rooms and count_rooms["n"] == 0:
        db_exec(
            """
        insert into public.wom_rooms (name) values
        ('SOTANO'),
        ('HAMMER KILLER'),
        ('RELIQUIAS DE JUDY'),
        ('PESADILLAS 2')
        on conflict (name) do nothing;
        """
        )


# =========================
# Helpers negocio
# =========================
def now_madrid() -> datetime:
    return datetime.now(TZ)


def formatear_fecha_hora(dt_value) -> Tuple[str, str]:
    try:
        if isinstance(dt_value, str):
            dt = datetime.fromisoformat(dt_value)
        else:
            dt = dt_value
        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=TZ)
        else:
            dt = dt.astimezone(TZ)
        return dt.strftime("%d/%m/%Y"), dt.strftime("%H:%M")
    except Exception:
        return "??/??/????", "??:??"


def get_user_by_code(code: str) -> Optional[Dict[str, str]]:
    c = (code or "").strip().upper()
    row = db_one(
        "select code, name, role from public.wom_users where upper(code)=upper(%s) limit 1;",
        (c,),
    )
    if not row:
        return None
    return {"codigo": row["code"].strip().upper(), "nombre": row["name"], "rol": row["role"]}


def get_salas() -> List[str]:
    rows = db_all("select name from public.wom_rooms order by name asc;")
    return [r["name"] for r in rows]


def generar_referencia() -> str:
    alfabeto = string.ascii_uppercase + string.digits
    while True:
        ref = "".join(random.choice(alfabeto) for _ in range(6))
        exists = db_one(
            "select 1 as x from public.wom_tickets where referencia=%s limit 1;", (ref,)
        )
        if not exists:
            return ref


def ticket_por_ref(ref: str) -> Optional[Dict[str, Any]]:
    r = (ref or "").strip().upper()
    return db_one("select * from public.wom_tickets where referencia=%s;", (r,))


def update_ticket(ref: str, set_sql: str, params: Tuple[Any, ...]) -> None:
    r = (ref or "").strip().upper()
    db_exec(
        f"update public.wom_tickets set {set_sql}, updated_at=now() where referencia=%s;",
        params + (r,),
    )


def sanitize_salas_selection(salas_selected: Optional[List[str]]) -> Optional[List[str]]:
    if not salas_selected:
        return None
    cleaned: List[str] = []
    for s in salas_selected:
        if not s:
            continue
        s = s.strip()
        if not s:
            continue
        cleaned.append(s)
    if not cleaned:
        return None
    if ALL_MARKER in cleaned:
        return None
    seen = set()
    out: List[str] = []
    for s in cleaned:
        if s not in seen:
            seen.add(s)
            out.append(s)
    return out or None


# =========================
# PDF (ReportLab Platypus)
# =========================
def _xml_escape(s: str) -> str:
    return (s or "").replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def _to_paragraph_text_multiline(s: str) -> str:
    return _xml_escape(s or "").replace("\n", "<br/>")


def _query_partes_en_proceso_filtrado(
    salas_filtro: Optional[List[str]],
) -> List[Dict[str, Any]]:
    if salas_filtro:
        return db_all(
            """
            select
              referencia,
              created_at,
              created_by_name,
              room_name,
              tipo,
              descripcion,
              solucionado_por_usuario,
              reparacion_usuario,
              visto_por_encargado,
              estado_encargado,
              observaciones_encargado
            from public.wom_tickets
            where estado_encargado not in ('TRABAJO TERMINADO/REPARADO','TRABAJO DESESTIMADO')
              and room_name = any(%s)
            order by created_at desc;
        """,
            (salas_filtro,),
        )
    return db_all(
        """
        select
          referencia,
          created_at,
          created_by_name,
          room_name,
          tipo,
          descripcion,
          solucionado_por_usuario,
          reparacion_usuario,
          visto_por_encargado,
          estado_encargado,
          observaciones_encargado
        from public.wom_tickets
        where estado_encargado not in ('TRABAJO TERMINADO/REPARADO','TRABAJO DESESTIMADO')
        order by created_at desc;
    """
    )


def generar_pdf_partes_en_proceso(salas_filtro: Optional[List[str]]) -> Path:
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.units import mm
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib import colors
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
    from reportlab.platypus.flowables import HRFlowable

    rows = _query_partes_en_proceso_filtrado(salas_filtro)

    out_dir = Path.cwd()
    ts = now_madrid().strftime("%Y%m%d_%H%M%S")
    out_path = out_dir / f"relacion_partes_en_proceso_{ts}.pdf"

    doc = SimpleDocTemplate(
        str(out_path),
        pagesize=A4,
        leftMargin=12 * mm,
        rightMargin=12 * mm,
        topMargin=12 * mm,
        bottomMargin=12 * mm,
        title="Relación de Partes en Proceso",
    )

    styles = getSampleStyleSheet()
    st_title = ParagraphStyle("title_small", parent=styles["Heading2"], fontSize=12, leading=14, spaceAfter=6)
    st_line = ParagraphStyle("line", parent=styles["Normal"], fontSize=8, leading=9, spaceAfter=1)
    st_label = ParagraphStyle("label", parent=styles["Normal"], fontSize=8, leading=9, spaceBefore=1, spaceAfter=0)
    st_mono = ParagraphStyle("mono", parent=styles["Normal"], fontName="Courier", fontSize=7.5, leading=9, spaceAfter=1)

    def e(s: str) -> str:
        return _xml_escape(s or "").replace("\n", "<br/>")

    filtro_txt = "TODAS" if not salas_filtro else ", ".join(salas_filtro)
    story = []
    story.append(Paragraph("Relación de Partes en Proceso", st_title))
    story.append(Paragraph(f"Salas: <b>{e(filtro_txt)}</b> — Generado: {now_madrid().strftime('%d/%m/%Y %H:%M')}", st_line))
    story.append(Spacer(1, 3))

    azul_sala = "#003366"

    for p in rows:
        fecha, hora = formatear_fecha_hora(p.get("created_at"))
        ref = (p.get("referencia") or "").strip()
        sala = p.get("room_name") or ""
        tipo = p.get("tipo") or ""
        prio = (p.get("priority") or "MEDIO").upper()
        autor = p.get("created_by_name") or ""

        desc = p.get("descripcion") or ""
        rep = p.get("reparacion_usuario") or ""
        com = p.get("observaciones_encargado") or ""

        # Línea 1: Ref / Fecha-Hora / Sala
        line1 = (
            f"<b>Ref:</b> {e(ref)}&nbsp;&nbsp;&nbsp;"
            f"<b>Fecha y hora:</b> {e(fecha)} {e(hora)}&nbsp;&nbsp;&nbsp;"
            f"<b>Sala:</b> <font color='{azul_sala}'><b>{e(sala)}</b></font>"
        )
        # Línea 2: Tipo / Prioridad / Usuario
        line2 = (
            f"<b>Tipo:</b> {e(tipo)}&nbsp;&nbsp;&nbsp;"
            f"<b>Nivel de prioridad:</b> {e(prio)}&nbsp;&nbsp;&nbsp;"
            f"<b>Usuario:</b> {e(autor)}"
        )

        story.append(Paragraph(line1, st_line))
        story.append(Paragraph(line2, st_line))

        story.append(Paragraph("<b>Descripción:</b>", st_label))
        story.append(Paragraph(e(desc) or "-", st_mono))

        if rep.strip():
            story.append(Paragraph("<b>Reparación / solución del usuario:</b>", st_label))
            story.append(Paragraph(e(rep), st_mono))

        story.append(Paragraph("<b>Comentario del Encargado:</b>", st_label))
        story.append(Paragraph(e(com) or "-", st_mono))

        story.append(Spacer(1, 1))
        story.append(HRFlowable(width="100%", thickness=0.3, color=colors.lightgrey))
        story.append(Spacer(1, 2))

    doc.build(story)
    return out_path


    for p in rows:
        fecha, hora = formatear_fecha_hora(p.get("created_at"))
        ref = (p.get("referencia") or "").strip()
        sala = p.get("room_name") or ""
        tipo = p.get("tipo") or ""
        autor = p.get("created_by_name") or ""
        visto = "Sí" if p.get("visto_por_encargado") else "No"
        estado = p.get("estado_encargado") or "SIN ESTADO"

        descripcion = p.get("descripcion") or "(Sin descripción)"
        observaciones = (p.get("observaciones_encargado") or "").strip() or "(Sin observaciones)"
        reparacion = (p.get("reparacion_usuario") or "").strip() or "(No aplica)"

        story.append(Paragraph(f"<b>Referencia:</b> {_xml_escape(ref)}", label))
        story.append(Paragraph(f"<b>Fecha/Hora:</b> {_xml_escape(fecha)} {_xml_escape(hora)}", label))
        story.append(Paragraph(f"<b>Sala:</b> {_xml_escape(sala)}", label))
        story.append(Paragraph(f"<b>Tipo:</b> {_xml_escape(tipo)}", label))
        story.append(Paragraph(f"<b>Creado por:</b> {_xml_escape(autor)}", label))
        story.append(
            Paragraph(
                f"<b>Visto:</b> {_xml_escape(visto)} &nbsp;&nbsp; <b>Estado:</b> {_xml_escape(estado)}",
                label,
            )
        )
        story.append(Spacer(1, 4))

        story.append(Paragraph("<b>Reparación realizada por el trabajador (si aplica):</b>", label))
        story.append(Paragraph(_to_paragraph_text_multiline(reparacion), block))

        story.append(Paragraph("<b>Observaciones del encargado:</b>", label))
        story.append(Paragraph(_to_paragraph_text_multiline(observaciones), block))

        story.append(Paragraph("<b>Descripción del parte:</b>", label))
        story.append(Paragraph(_to_paragraph_text_multiline(descripcion), block))

        story.append(HRFlowable(thickness=0.6, width="100%"))
        story.append(Spacer(1, 10))

    doc.build(story)
    return out_path


# =========================
# HTML helpers
# =========================
def h(s: Any) -> str:
    import html

    return html.escape("" if s is None else str(s))


def page(title: str, body: str) -> str:
    return f"""<!doctype html>
<html lang="es">
<head>
  <meta charset="utf-8"/>
  <meta name="viewport" content="width=device-width,initial-scale=1"/>
  <title>{h(title)}</title>
  <style>
    body{{font-family: -apple-system, system-ui, Arial; margin: 24px; max-width: 980px}}
    .card{{border:1px solid #ddd; border-radius:12px; padding:16px; margin:12px 0}}
    .row{{display:flex; gap:12px; flex-wrap:wrap}}
    .btn{{display:inline-block; padding:10px 14px; border-radius:10px; border:1px solid #333; text-decoration:none; color:#111; background:#fff}}
    .btn2{{display:inline-block; padding:10px 14px; border-radius:10px; border:1px solid #999; text-decoration:none; color:#111; background:#fff}}
    .danger{{border-color:#c00; color:#c00}}
    input, select, textarea{{width:100%; padding:10px; border-radius:10px; border:1px solid #ccc; box-sizing:border-box}}
    label{{font-weight:600; display:block; margin-top:10px}}
    textarea{{min-height:140px}}
    .muted{{color:#666}}
    .top{{display:flex; justify-content:space-between; align-items:center; gap:10px}}
    .pill{{display:inline-block; padding:4px 10px; border-radius:999px; border:1px solid #ddd; font-size:12px; margin-right:6px; margin-top:6px}}
    table{{width:100%; border-collapse:collapse}}
    th,td{{text-align:left; padding:8px; border-bottom:1px solid #eee; vertical-align:top}}
    code{{background:#f6f6f6; padding:2px 6px; border-radius:6px}}
    .ticket{{border:1px solid #eee; border-radius:12px; padding:12px; margin:12px 0}}
    .ticket h3{{margin:0 0 6px 0}}
    .hr{{border-top:1px solid #eee; margin:10px 0}}
.btn-attn{{font-weight:700; color:#8a0041; border-color:#8a0041; background:#ffe4f0}}
  </style>
</head>
<body>
{body}
</body></html>"""


def user_from_session(request: Request):
    return request.session.get("user")


def require_login(request: Request):
    u = user_from_session(request)
    if not u:
        return RedirectResponse("/", status_code=303)
    return None


def role_home_path(role: str) -> str:
    role = (role or "").upper()
    if role == "ENCARGADO":
        return "/encargado"
    if role == "JEFE":
        return "/jefe"
    return "/trabajador"


def salas_multiselect_html(salas: List[str], selected: Optional[List[str]], label: str) -> str:
    selected = selected or [ALL_MARKER]
    opts: List[str] = []
    sel_all = "selected" if (ALL_MARKER in selected) else ""
    opts.append(f"<option value='{ALL_MARKER}' {sel_all}>TODAS</option>")
    for s in salas:
        sel = "selected" if (s in selected) else ""
        opts.append(f"<option value='{h(s)}' {sel}>{h(s)}</option>")

    return f"""
      <label>{h(label)}</label>
      <select name="salas" multiple size="{min(max(len(salas)+1, 5), 10)}" id="salas_select" onchange="enforceAllRule()">
        {''.join(opts)}
      </select>
      <p class="muted" style="margin-top:8px">
        Consejo: selecciona <b>TODAS</b> o selecciona una o varias salas. Si eliges TODAS, se ignorarán otras selecciones.
      </p>
      <script>
        function enforceAllRule() {{
          var sel = document.getElementById('salas_select');
          var values = Array.from(sel.selectedOptions).map(o => o.value);
          if (values.includes('{ALL_MARKER}') && values.length > 1) {{
            for (var i=0; i<sel.options.length; i++) {{
              sel.options[i].selected = (sel.options[i].value === '{ALL_MARKER}');
            }}
          }}
        }}
      </script>
    """


def render_ticket_blocks(
    rows: List[Dict[str, Any]],
    back_href: str,
    title: str,
    subtitle: str,
    show_link: bool = True,
) -> str:
    blocks: List[str] = []
    for p in rows:
        fecha, hora = formatear_fecha_hora(p.get("created_at"))
        ref = (p.get("referencia") or "").strip()
        visto = "Sí" if p.get("visto_por_encargado") else "No"
        estado = p.get("estado_encargado") or "SIN ESTADO"
        sol = bool(p.get("solucionado_por_usuario", False))

        rep = (p.get("reparacion_usuario") or "").strip()
        if sol:
            rep_txt = rep if rep else "(No indicó reparación)"
        else:
            rep_txt = "(No aplica)"

        obs = (p.get("observaciones_encargado") or "").strip() or "(Sin observaciones)"
        desc = (p.get("descripcion") or "").strip() or "(Sin descripción)"

        header = h(ref)
        if show_link:
            header = f"<a href='/parte/{h(ref)}'>{h(ref)}</a>"

        blocks.append(
            f"""
          <div class="ticket">
            <h3>Referencia: {header}</h3>
            <div class="pill">Fecha/Hora: {h(fecha)} {h(hora)}</div>
            <div class="pill">Sala: {h(p.get('room_name',''))}</div>
            <div class="pill">Tipo: {h(p.get('tipo',''))}</div>
            <div class="pill">Creado por: {h(p.get('created_by_name',''))}</div>
            <div class="pill">Visto: {h(visto)}</div>
            <div class="pill">Estado: {h(estado)}</div>
            <div class="hr"></div>
            <p><b>Reparación realizada por el trabajador (si aplica):</b><br/>{h(rep_txt).replace(chr(10), "<br/>")}</p>
            <p><b>Observaciones del encargado:</b><br/>{h(obs).replace(chr(10), "<br/>")}</p>
            <p><b>Descripción del parte:</b><br/>{h(desc).replace(chr(10), "<br/>")}</p>
          </div>
        """
        )

    body = f"""
      <div class="top">
        <div>
          <h2>{h(title)}</h2>
          <p class="muted">{h(subtitle)}</p>
        </div>
        <div><a class="btn2" href="{h(back_href)}">Volver</a></div>
      </div>
      <div class="card">
        {''.join(blocks) if blocks else "<p>No hay partes para el filtro seleccionado.</p>"}
      </div>
    """
    return body


# =========================
# FASTAPI APP
# =========================
app = FastAPI()
app.add_middleware(
    SessionMiddleware,
    secret_key=os.getenv("SESSION_SECRET", "wom_local_secret_key_cambia_esto"),
)


@app.on_event("startup")
def _startup():
    ensure_schema_and_seed()


@app.get("/health")
def health():
    return PlainTextResponse("ok")


# =========================
# LOGIN / LOGOUT
# =========================
@app.get("/", response_class=HTMLResponse)
def login_page(request: Request):
    u = user_from_session(request)
    if u:
        return RedirectResponse("/home", status_code=303)

    body = '''
    <div class="card">
      <h2>PARTES DE MANTENIMIENTO DE WOM</h2>
      <p class="muted"><i>Versión 1.5 Enero 2026</i></p>
      <form method="post" action="/login">
        <label>Código personal</label>
        <input name="codigo" placeholder="Ej: A123B" autocomplete="off"/>
        <div style="margin-top:12px">
          <button class="btn" type="submit">Entrar</button>
        </div>
      </form>

      <p class="muted" style="margin-top:14px; font-style:italic; font-size:0.92em;">
        *** Novedades de la Versión 1.5 ***<br/><br/>
        - Corrección de errores de creación de formularios<br/>
        - Ahora los trabajadores pueden filtrar por mes y año su listado de partes Finalizados<br/>
        - Nuevas opciones de registro en el menú de encargado<br/>
        - Corrección de arreglos en el menú de Jefes
      </p>
    </div>
    '''
    return page("Login", body)



@app.post("/login")
def do_login(request: Request, codigo: str = Form(...)):
    info = get_user_by_code(codigo)
    if not info:
        return HTMLResponse(
            page(
                "Login",
                """
          <div class='card'>
            <h3>Código no reconocido</h3>
            <p><a class='btn2' href='/'>Volver</a></p>
          </div>
        """,
            ),
            status_code=400,
        )

    request.session["user"] = info
    return RedirectResponse("/home", status_code=303)


@app.get("/home")
def home(request: Request):
    r = require_login(request)
    if r:
        return r
    u = user_from_session(request)
    return RedirectResponse(role_home_path(u.get("rol", "")), status_code=303)


@app.get("/logout")
def logout(request: Request):
    request.session.clear()
    return RedirectResponse("/", status_code=303)


# =========================
# TRABAJADOR (menú + flujos)
# =========================
@app.get("/trabajador", response_class=HTMLResponse)
def worker_menu(request: Request):
    r = require_login(request)
    if r:
        return r
    u = user_from_session(request)
    if u["rol"] != "TRABAJADOR":
        return RedirectResponse(role_home_path(u["rol"]), status_code=303)

    body = f"""
    <div class="top">
      <div>
        <h2>PARTES DE MANTENIMIENTO DE WOM</h2>
        <p>Hola <b>{h(u["nombre"])}</b>! Comencemos a dar un parte...</p>
      </div>
      <div><a class="btn2" href="/logout">Salir</a></div>
    </div>

    <div class="card">
      <div class="row">
        <a class="btn" href="/trabajador/nuevo">Crear nuevo parte</a>
        <a class="btn" href="/trabajador/activos">Ver partes en proceso</a>
        <a class="btn" href="/trabajador/finalizados">Ver partes finalizados</a>
      </div>
    </div>
    """
    return page("Trabajador", body)


@app.get("/trabajador/nuevo", response_class=HTMLResponse)
def worker_new_form(request: Request):
    r = require_login(request)
    if r:
        return r
    u = user_from_session(request)
    if u["rol"] != "TRABAJADOR":
        return RedirectResponse(role_home_path(u["rol"]), status_code=303)

    ref = generar_referencia()
    salas = get_salas()
    salas_opts = "".join([f"<option value='{h(s)}'>{h(s)}</option>" for s in salas])
    tipos_opts = "".join([f"<option value='{h(t)}'>{h(t)}</option>" for t in TIPOS])

    body = f"""
    <div class="top">
      <div><h2>Nuevo parte</h2><p class="muted">Referencia generada: <code>{h(ref)}</code> (anótala)</p></div>
      <div><a class="btn2" href="/trabajador">Volver</a></div>
    </div>

    <div class="card">
      <form method="post" action="/trabajador/nuevo">
        <input type="hidden" name="referencia" value="{h(ref)}"/>

        <label>Sala</label>
        <select name="sala">{salas_opts}</select>

        <label>Tipo</label>
        <select name="tipo">{tipos_opts}</select>

        <label>Descripción</label>
        <textarea name="descripcion" placeholder="Describe en detalle..."></textarea>

        <label>¿Has podido solucionar tú el problema?</label>
        <select name="solucionado" id="solucionado" onchange="toggleReparacion()">
          <option value="NO">NO</option>
          <option value="SI">SI</option>
        </select>

        <div id="reparacion_wrap" style="display:none;">
          <label>¿Qué solución o reparación has hecho?</label>
          <textarea name="reparacion_usuario" id="reparacion_usuario" placeholder="Describe la reparación..."></textarea>
        </div>

        <div style="margin-top:12px">
          <button class="btn" type="submit">Guardar parte</button>
        </div>
      </form>
    </div>

    <script>
      function toggleReparacion() {{
        var v = document.getElementById("solucionado").value;
        var wrap = document.getElementById("reparacion_wrap");
        var txt = document.getElementById("reparacion_usuario");
        if (v === "SI") {{
          wrap.style.display = "block";
          txt.disabled = false;
        }} else {{
          wrap.style.display = "none";
          txt.value = "";
          txt.disabled = true;
        }}
      }}
      toggleReparacion();
    </script>
    """
    return page("Nuevo parte", body)


@app.post("/trabajador/nuevo")
def worker_new_submit(
    request: Request,
    referencia: str = Form(...),
    sala: str = Form(...),
    tipo: str = Form(...),
    descripcion: str = Form(""),
    solucionado: str = Form("NO"),
    reparacion_usuario: str = Form(""),
):
    r = require_login(request)
    if r:
        return r
    u = user_from_session(request)
    if u["rol"] != "TRABAJADOR":
        return RedirectResponse(role_home_path(u["rol"]), status_code=303)

    ref = (referencia or "").strip().upper()
    sala_name = (sala or "").strip()
    tipo_name = (tipo or "").strip()
    desc = (descripcion or "").strip() or "(Sin descripción)"

    sol = (solucionado or "").strip().upper() == "SI"
    rep = (reparacion_usuario or "").strip() if sol else ""

    room = db_one("select id, name from public.wom_rooms where name=%s;", (sala_name,))
    room_id = room["id"] if room else None

    db_exec(
        """
        insert into public.wom_tickets
        (referencia, created_by_code, created_by_name, room_id, room_name, tipo, descripcion,
         solucionado_por_usuario, reparacion_usuario, visto_por_encargado, estado_encargado, observaciones_encargado)
        values
        (%s, %s, %s, %s, %s, %s, %s, %s, %s, false, 'SIN ESTADO', '')
        on conflict (referencia) do nothing;
    """,
        (ref, u["codigo"], u["nombre"], room_id, sala_name, tipo_name, desc, sol, rep),
    )

    return RedirectResponse(f"/parte/{ref}", status_code=303)


@app.get("/trabajador/activos", response_class=HTMLResponse)
def worker_activos(request: Request):
    r = require_login(request)
    if r:
        return r
    u = user_from_session(request)
    if u["rol"] != "TRABAJADOR":
        return RedirectResponse(role_home_path(u["rol"]), status_code=303)

    rows = db_all(
        """
        select referencia, created_at, created_by_name, room_name, tipo, estado_encargado, visto_por_encargado
        from public.wom_tickets
        where estado_encargado not in ('TRABAJO TERMINADO/REPARADO','TRABAJO DESESTIMADO')
        order by created_at desc;
    """
    )

    trs = ""
    for p in rows:
        f, hh = formatear_fecha_hora(p.get("created_at"))
        visto = "Sí" if p.get("visto_por_encargado") else "No"
        ref = (p.get("referencia") or "").strip()
        trs += f"""
        <tr>
          <td><a href="/parte/{h(ref)}">{h(ref)}</a></td>
          <td>{h(f)} {h(hh)}</td>
          <td>{h(p.get("created_by_name",""))}</td>
          <td>{h(p.get("room_name",""))}</td>
          <td>{h(p.get("tipo",""))}</td>
          <td>{h(p.get("estado_encargado","SIN ESTADO"))}</td>
          <td>{h(visto)}</td>
        </tr>
        """

    body = f"""
    <div class="top">
      <div><h2>Partes en proceso</h2><p class="muted">Listado de todos los partes no finalizados.</p></div>
      <div><a class="btn2" href="/trabajador">Volver</a></div>
    </div>
    <div class="card">
      <table>
        <thead><tr><th>Ref</th><th>Fecha</th><th>Autor</th><th>Sala</th><th>Tipo</th><th>Estado</th><th>Visto</th></tr></thead>
        <tbody>{trs or "<tr><td colspan='7'>No hay partes.</td></tr>"}</tbody>
      </table>
    </div>
    """
    return page("En proceso", body)


@app.get("/trabajador/finalizados", response_class=HTMLResponse)
def worker_finalizados(request: Request):
    r = require_login(request)
    if r:
        return r
    u = user_from_session(request)
    if u["rol"] != "TRABAJADOR":
        return RedirectResponse(role_home_path(u["rol"]), status_code=303)

    now = now_madrid()
    body = f'''
    <div class="top">
      <div><h2>Partes finalizados</h2></div>
      <div><a class="btn2" href="/trabajador">Volver</a></div>
    </div>

    <div class="card">
      <h3>Filtrar por mes y año</h3>
      <form method="post" action="/trabajador/finalizados">
        <label>Mes</label>
        <select name="mes">
          {''.join([f"<option value='{m}' {'selected' if m==now.month else ''}>{m:02d}</option>" for m in range(1,13)])}
        </select>
        <label>Año</label>
        <input name="anio" type="number" value="{now.year}" min="2020" max="2100" required/>
        <div style="margin-top:12px">
          <button class="btn" type="submit">Ver finalizados</button>
        </div>
      </form>
    </div>
    '''
    return page("Finalizados", body)


@app.post("/trabajador/finalizados", response_class=HTMLResponse)
def worker_finalizados_post(request: Request, mes: int = Form(...), anio: int = Form(...)):
    r = require_login(request)
    if r:
        return r
    u = user_from_session(request)
    if u["rol"] != "TRABAJADOR":
        return RedirectResponse(role_home_path(u["rol"]), status_code=303)

    mval = int(mes)
    yval = int(anio)
    if mval < 1 or mval > 12:
        mval = now_madrid().month
    if yval < 2000 or yval > 2100:
        yval = now_madrid().year

    start, end = month_bounds(yval, mval)

    rows = db_all(
        '''
        select referencia, created_at, created_by_name, room_name, tipo, priority, estado_encargado, visto_por_encargado
        from public.wom_tickets
        where estado_encargado in ('TRABAJO TERMINADO/REPARADO','TRABAJO DESESTIMADO')
          and created_at >= %s and created_at < %s
        order by created_at desc;
    ''',
        (start, end),
    )

    trs = ""
    for p in rows:
        f, hh = formatear_fecha_hora(p.get("created_at"))
        visto = "Sí" if p.get("visto_por_encargado") else "No"
        ref = (p.get("referencia") or "").strip()
        estado = p.get("estado_encargado", "SIN ESTADO")
        prio = p.get("priority", "MEDIO")
        trs += f'''
        <tr>
          <td><a href="/parte/{h(ref)}">{h(ref)}</a></td>
          <td>{h(f)} {h(hh)}</td>
          <td>{h(p.get("created_by_name",""))}</td>
          <td>{h(p.get("room_name",""))}</td>
          <td>{h(p.get("tipo",""))}</td>
          {_estado_cell(estado, prio)}
          <td>{h(visto)}</td>
        </tr>
        '''

    body = f'''
    <div class="top">
      <div><h2>Partes finalizados</h2><p class="muted">Filtrado: {mval:02d}/{yval}</p></div>
      <div><a class="btn2" href="/trabajador/finalizados">Cambiar filtro</a></div>
    </div>
    <div class="card">
      <table>
        <thead><tr><th>Ref</th><th>Fecha</th><th>Autor</th><th>Sala</th><th>Tipo</th><th>Estado</th><th>Visto</th></tr></thead>
        <tbody>{trs or "<tr><td colspan='7'>No hay partes.</td></tr>"}</tbody>
      </table>
    </div>
    '''
    return page("Finalizados", body)



# =========================
# JEFE
# =========================
@app.get("/jefe", response_class=HTMLResponse)
def jefe_menu(request: Request):
    r = require_login(request)
    if r:
        return r
    u = user_from_session(request)
    if u["rol"] != "JEFE":
        return RedirectResponse(role_home_path(u["rol"]), status_code=303)

    body = f"""
    <div class="top">
      <div>
        <h2>VISTA DE JEFATURA - PARTES WOM</h2>
        <p>Bienvenido <b>{h(u["nombre"])}</b>.</p>
      </div>
      <div><a class="btn2" href="/logout">Salir</a></div>
    </div>

    <div class="card">
      <div class="row">
        <a class="btn" href="/jefe/en_proceso">Ver listado de partes en activo</a>
        <a class="btn" href="/jefe/finalizados">Ver listado de partes finalizados</a>
        <a class="btn" href="/jefe/consulta_en_proceso">Consulta de partes en proceso</a>
      </div>
    </div>
    """
    return page("Jefe", body)


@app.get("/jefe/en_proceso", response_class=HTMLResponse)
def jefe_en_proceso(request: Request):
    r = require_login(request)
    if r:
        return r
    u = user_from_session(request)
    if u["rol"] != "JEFE":
        return RedirectResponse(role_home_path(u["rol"]), status_code=303)

    rows = db_all(
        """
        select referencia, created_at, created_by_name, room_name, tipo, estado_encargado, visto_por_encargado
        from public.wom_tickets
        where estado_encargado not in ('TRABAJO TERMINADO/REPARADO','TRABAJO DESESTIMADO')
        order by created_at desc;
    """
    )

    trs = ""
    for p in rows:
        f, hh = formatear_fecha_hora(p.get("created_at"))
        visto = "Sí" if p.get("visto_por_encargado") else "No"
        ref = (p.get("referencia") or "").strip()
        trs += f"""
        <tr>
          <td><a href="/parte/{h(ref)}">{h(ref)}</a></td>
          <td>{h(f)} {h(hh)}</td>
          <td>{h(p.get("created_by_name",""))}</td>
          <td>{h(p.get("room_name",""))}</td>
          <td>{h(p.get("tipo",""))}</td>
          <td>{h(p.get("estado_encargado","SIN ESTADO"))}</td>
          <td>{h(visto)}</td>
        </tr>
        """

    body = f"""
    <div class="top">
      <div><h2>Partes en activo</h2></div>
      <div><a class="btn2" href="/jefe">Volver</a></div>
    </div>
    <div class="card">
      <table>
        <thead><tr><th>Ref</th><th>Fecha</th><th>Autor</th><th>Sala</th><th>Tipo</th><th>Estado</th><th>Visto</th></tr></thead>
        <tbody>{trs or "<tr><td colspan='7'>No hay partes.</td></tr>"}</tbody>
      </table>
    </div>
    """
    return page("Jefe - En activo", body)


@app.get("/jefe/finalizados", response_class=HTMLResponse)
def jefe_finalizados(request: Request):
    r = require_login(request)
    if r:
        return r
    u = user_from_session(request)
    if u["rol"] != "JEFE":
        return RedirectResponse(role_home_path(u["rol"]), status_code=303)

    rows = db_all(
        """
        select referencia, created_at, created_by_name, room_name, tipo, estado_encargado, visto_por_encargado
        from public.wom_tickets
        where estado_encargado in ('TRABAJO TERMINADO/REPARADO','TRABAJO DESESTIMADO')
        order by created_at desc;
    """
    )

    trs = ""
    for p in rows:
        f, hh = formatear_fecha_hora(p.get("created_at"))
        visto = "Sí" if p.get("visto_por_encargado") else "No"
        ref = (p.get("referencia") or "").strip()
        trs += f"""
        <tr>
          <td><a href="/parte/{h(ref)}">{h(ref)}</a></td>
          <td>{h(f)} {h(hh)}</td>
          <td>{h(p.get("created_by_name",""))}</td>
          <td>{h(p.get("room_name",""))}</td>
          <td>{h(p.get("tipo",""))}</td>
          <td>{h(p.get("estado_encargado","SIN ESTADO"))}</td>
          <td>{h(visto)}</td>
        </tr>
        """

    body = f"""
    <div class="top">
      <div><h2>Partes finalizados</h2></div>
      <div><a class="btn2" href="/jefe">Volver</a></div>
    </div>
    <div class="card">
      <table>
        <thead><tr><th>Ref</th><th>Fecha</th><th>Autor</th><th>Sala</th><th>Tipo</th><th>Estado</th><th>Visto</th></tr></thead>
        <tbody>{trs or "<tr><td colspan='7'>No hay partes.</td></tr>"}</tbody>
      </table>
    </div>
    """
    return page("Jefe - Finalizados", body)


@app.get("/jefe/consulta_en_proceso", response_class=HTMLResponse)
def jefe_consulta_en_proceso_form(request: Request):
    r = require_login(request)
    if r:
        return r
    u = user_from_session(request)
    if u["rol"] != "JEFE":
        return RedirectResponse(role_home_path(u["rol"]), status_code=303)

    salas = get_salas()
    selector = salas_multiselect_html(salas, None, "Selecciona sala(s) para filtrar (o TODAS)")

    body = f"""
    <div class="top">
      <div><h2>Consulta de partes en proceso</h2></div>
      <div><a class="btn2" href="/jefe">Volver</a></div>
    </div>

    <div class="card">
      <form method="post" action="/jefe/consulta_en_proceso">
        {selector}
        <div style="margin-top:12px">
          <button class="btn" type="submit">Ver partes</button>
        </div>
      </form>
    </div>
    """
    return page("Jefe - Consulta", body)


@app.post("/jefe/consulta_en_proceso", response_class=HTMLResponse)
def jefe_consulta_en_proceso_result(request: Request, salas: List[str] = Form([])):
    r = require_login(request)
    if r:
        return r
    u = user_from_session(request)
    if u["rol"] != "JEFE":
        return RedirectResponse(role_home_path(u["rol"]), status_code=303)

    salas_filtro = sanitize_salas_selection(salas)
    rows = _query_partes_en_proceso_filtrado(salas_filtro)

    filtro_txt = "TODAS LAS SALAS" if not salas_filtro else ", ".join(salas_filtro)
    body = render_ticket_blocks(
        rows=rows,
        back_href="/jefe",
        title="Consulta de partes en proceso",
        subtitle=f"Filtro de salas: {filtro_txt}",
        show_link=True,
    )
    return page("Jefe - Resultados", body)


# =========================
# DETALLE PARTE (común)
# =========================
@app.get("/parte/{ref}", response_class=HTMLResponse)
def parte_detalle(request: Request, ref: str):
    r = require_login(request)
    if r:
        return r
    u = user_from_session(request)

    p = ticket_por_ref(ref)
    if not p:
        return HTMLResponse(
            page(
                "No encontrado",
                f"<div class='card'><h3>No existe el parte {h(ref)}</h3></div>",
            ),
            status_code=404,
        )

    fecha, hora = formatear_fecha_hora(p.get("created_at"))
    visto = "Sí" if p.get("visto_por_encargado") else "No"
    estado = p.get("estado_encargado") or "SIN ESTADO"
    sol = bool(p.get("solucionado_por_usuario", False))
    rep = (p.get("reparacion_usuario") or "").strip()
    obs = (p.get("observaciones_encargado") or "").strip()

    rep_html = ""
    if sol:
        rep_html = f"""
        <div class="card">
          <h3>Reparación realizada por el trabajador</h3>
          <p>{h(rep if rep else "(No indicó reparación)").replace(chr(10),"<br/>")}</p>
        </div>
        """

    obs_html = ""
    if obs:
        obs_html = f"""
        <div class="card">
          <h3>Observaciones del encargado</h3>
          <p>{h(obs).replace(chr(10),"<br/>")}</p>
        </div>
        """

    back = role_home_path(u["rol"])
    body = f"""
    <div class="top">
      <div>
        <h2>Parte {h((p.get("referencia") or "").strip())}</h2>
        <div class="pill">Fecha: {h(fecha)} {h(hora)}</div>
        <div class="pill">Visto: {h(visto)}</div>
        <div class="pill">Estado: {h(estado)}</div>
      </div>
      <div><a class="btn2" href="{h(back)}">Volver</a></div>
    </div>

    <div class="card">
      <p><b>Sala:</b> {h(p.get("room_name",""))}</p>
      <p><b>Tipo:</b> {h(p.get("tipo",""))}</p>
      <p><b>Creado por:</b> {h(p.get("created_by_name",""))}</p>
      <p><b>¿Solucionado por el usuario?:</b> {"Sí" if sol else "No"}</p>
    </div>

    {rep_html}
    {obs_html}

    <div class="card">
      <h3>Descripción</h3>
      <p>{h(p.get("descripcion","")).replace(chr(10),"<br/>")}</p>
    </div>
    """

    if u["rol"] == "ENCARGADO":
        estados_opts = "".join(
            [
                f"<option value='{h(e)}' {'selected' if e==estado else ''}>{h(e)}</option>"
                for e in ESTADOS_ENCARGADO
            ]
        )
        body += f"""
        <div class="card">
          <h3>Acciones del encargado</h3>

          <form method="post" action="/encargado/mark_visto/{h((p.get("referencia") or "").strip())}">
            <button class="btn" type="submit">Marcar como leído/visto</button>
          </form>

          <form method="post" action="/encargado/set_estado/{h((p.get("referencia") or "").strip())}" style="margin-top:12px">
            <label>Cambiar estado</label>
            <select name="estado">{estados_opts}</select>
            <div style="margin-top:10px">
              <button class="btn" type="submit">Guardar estado</button>
            </div>
          </form>

          <form method="post" action="/encargado/set_obs/{h((p.get("referencia") or "").strip())}" style="margin-top:12px">
            <label>Observaciones del encargado (editable)</label>
            <textarea name="obs">{h(p.get("observaciones_encargado",""))}</textarea>
            <div style="margin-top:10px">
              <button class="btn" type="submit">Guardar observaciones</button>
            </div>
          </form>
        </div>
        """

    return page("Detalle", body)


# =========================
# ENCARGADO
# =========================
@app.get("/encargado", response_class=HTMLResponse)
def admin_menu(request: Request):
    r = require_login(request)
    if r:
        return r
    u = user_from_session(request)
    if u["rol"] != "ENCARGADO":
        return RedirectResponse(role_home_path(u["rol"]), status_code=303)

    row = db_one(
        '''
        select count(*)::int as n
        from public.wom_tickets
        where coalesce(estado_encargado,'') not in ('TRABAJO TERMINADO/REPARADO','TRABAJO DESESTIMADO')
          and visto_por_encargado = false;
    '''
    )
    unseen = int((row or {}).get("n") or 0)
    pend_class = "btn btn-attn" if unseen > 0 else "btn"

    body = f'''
    <div class="top">
      <div>
        <h2>CONTROL DE PARTES DE MANTENIMIENTO</h2>
        <p>¡Bienvenido <b>{h(u["nombre"]).upper()}</b>!</p>
      </div>
      <div><a class="btn2" href="/logout">Salir</a></div>
    </div>

    <div class="card">
      <div class="row">
        <a class="{pend_class}" href="/encargado/pendientes">Ver pendientes</a>
        <a class="btn" href="/encargado/finalizados">Ver finalizados</a>
        <a class="btn" href="/encargado/gestion_partes">Gestión de Partes</a>
        <a class="btn" href="/encargado/gestion_usuarios">Gestión de Usuarios</a>
      </div>
    </div>
    '''
    return page("Encargado", body)



@app.get("/encargado/pendientes", response_class=HTMLResponse)
def admin_pendientes(request: Request):
    r = require_login(request)
    if r:
        return r
    u = user_from_session(request)
    if u["rol"] != "ENCARGADO":
        return RedirectResponse(role_home_path(u["rol"]), status_code=303)

    rows = db_all(
        """
        select referencia, created_at, created_by_name, room_name, tipo, estado_encargado, visto_por_encargado
        from public.wom_tickets
        where estado_encargado not in ('TRABAJO TERMINADO/REPARADO','TRABAJO DESESTIMADO')
        order by created_at desc;
    """
    )

    trs = ""
    for p in rows:
        f, hh = formatear_fecha_hora(p.get("created_at"))
        visto = "Sí" if p.get("visto_por_encargado") else "No"
        ref = (p.get("referencia") or "").strip()
        trs += f"""
        <tr>
          <td><a href="/parte/{h(ref)}">{h(ref)}</a></td>
          <td>{h(f)} {h(hh)}</td>
          <td>{h(p.get("created_by_name",""))}</td>
          <td>{h(p.get("room_name",""))}</td>
          <td>{h(p.get("tipo",""))}</td>
          <td>{h(p.get("estado_encargado","SIN ESTADO"))}</td>
          <td>{h(visto)}</td>
        </tr>
        """

    body = f"""
    <div class="top">
      <div><h2>Pendientes / en curso</h2></div>
      <div><a class="btn2" href="/encargado">Volver</a></div>
    </div>
    <div class="card">
      <table>
        <thead><tr><th>Ref</th><th>Fecha</th><th>Autor</th><th>Sala</th><th>Tipo</th><th>Estado</th><th>Visto</th></tr></thead>
        <tbody>{trs or "<tr><td colspan='7'>No hay partes.</td></tr>"}</tbody>
      </table>
    </div>
    """
    return page("Pendientes", body)


@app.get("/encargado/finalizados", response_class=HTMLResponse)
def admin_finalizados(request: Request):
    r = require_login(request)
    if r:
        return r
    u = user_from_session(request)
    if u["rol"] != "ENCARGADO":
        return RedirectResponse(role_home_path(u["rol"]), status_code=303)

    rows = db_all(
        """
        select referencia, created_at, created_by_name, room_name, tipo, estado_encargado, visto_por_encargado
        from public.wom_tickets
        where estado_encargado in ('TRABAJO TERMINADO/REPARADO','TRABAJO DESESTIMADO')
        order by created_at desc;
    """
    )

    trs = ""
    for p in rows:
        f, hh = formatear_fecha_hora(p.get("created_at"))
        visto = "Sí" if p.get("visto_por_encargado") else "No"
        ref = (p.get("referencia") or "").strip()
        trs += f"""
        <tr>
          <td><a href="/parte/{h(ref)}">{h(ref)}</a></td>
          <td>{h(f)} {h(hh)}</td>
          <td>{h(p.get("created_by_name",""))}</td>
          <td>{h(p.get("room_name",""))}</td>
          <td>{h(p.get("tipo",""))}</td>
          <td>{h(p.get("estado_encargado","SIN ESTADO"))}</td>
          <td>{h(visto)}</td>
        </tr>
        """

    body = f"""
    <div class="top">
      <div><h2>Finalizados</h2></div>
      <div><a class="btn2" href="/encargado">Volver</a></div>
    </div>
    <div class="card">
      <table>
        <thead><tr><th>Ref</th><th>Fecha</th><th>Autor</th><th>Sala</th><th>Tipo</th><th>Estado</th><th>Visto</th></tr></thead>
        <tbody>{trs or "<tr><td colspan='7'>No hay partes.</td></tr>"}</tbody>
      </table>
    </div>
    """
    return page("Finalizados", body)


@app.post("/encargado/mark_visto/{ref}")
def admin_mark_visto(request: Request, ref: str):
    r = require_login(request)
    if r:
        return r
    u = user_from_session(request)
    if u["rol"] != "ENCARGADO":
        return RedirectResponse(role_home_path(u["rol"]), status_code=303)

    update_ticket(ref, "visto_por_encargado=true", ())
    return RedirectResponse(f"/parte/{ref}", status_code=303)


@app.post("/encargado/set_estado/{ref}")
def admin_set_estado(request: Request, ref: str, estado: str = Form(...)):
    r = require_login(request)
    if r:
        return r
    u = user_from_session(request)
    if u["rol"] != "ENCARGADO":
        return RedirectResponse(role_home_path(u["rol"]), status_code=303)

    est = (estado or "").strip()
    if est in ESTADOS_ENCARGADO:
        update_ticket(ref, "estado_encargado=%s, visto_por_encargado=true", (est,))
    return RedirectResponse(f"/parte/{ref}", status_code=303)


@app.post("/encargado/set_obs/{ref}")
def admin_set_obs(request: Request, ref: str, obs: str = Form("")):
    r = require_login(request)
    if r:
        return r
    u = user_from_session(request)
    if u["rol"] != "ENCARGADO":
        return RedirectResponse(role_home_path(u["rol"]), status_code=303)

    update_ticket(ref, "observaciones_encargado=%s, visto_por_encargado=true", ((obs or "").strip(),))
    return RedirectResponse(f"/parte/{ref}", status_code=303)


# =========================
# ENCARGADO - Gestión de Partes
# =========================
@app.get("/encargado/gestion_partes", response_class=HTMLResponse)
def admin_gestion_partes(request: Request):
    r = require_login(request)
    if r:
        return r
    u = user_from_session(request)
    if u["rol"] != "ENCARGADO":
        return RedirectResponse(role_home_path(u["rol"]), status_code=303)

    body = """
    <div class="top">
      <div><h2>Gestión de Partes</h2></div>
      <div><a class="btn2" href="/encargado">Volver</a></div>
    </div>

    <div class="card">
      <div class="row">
        <a class="btn" href="/encargado/pdf">Generar PDF de partes en proceso</a>
        <a class="btn" href="/encargado/visualizar_en_proceso">Visualizar partes en Proceso</a>
        <a class="btn danger" href="/encargado/eliminar_partes">Eliminar partes del sistema</a>
      </div>
      <p class="muted" style="margin-top:10px">Eliminar un parte lo borra para todos los roles.</p>
    </div>
    """
    return page("Gestión de Partes", body)


@app.get("/encargado/visualizar_en_proceso", response_class=HTMLResponse)
def admin_visualizar_en_proceso_form(request: Request):
    r = require_login(request)
    if r:
        return r
    u = user_from_session(request)
    if u["rol"] != "ENCARGADO":
        return RedirectResponse(role_home_path(u["rol"]), status_code=303)

    salas = get_salas()
    selector = salas_multiselect_html(salas, None, "Selecciona sala(s) para filtrar (o TODAS)")

    body = f"""
    <div class="top">
      <div><h2>Visualizar partes en proceso</h2></div>
      <div><a class="btn2" href="/encargado/gestion_partes">Volver</a></div>
    </div>

    <div class="card">
      <form method="post" action="/encargado/visualizar_en_proceso">
        {selector}
        <div style="margin-top:12px">
          <button class="btn" type="submit">Ver partes</button>
        </div>
      </form>
    </div>
    """
    return page("Encargado - Visualizar", body)


@app.post("/encargado/visualizar_en_proceso", response_class=HTMLResponse)
def admin_visualizar_en_proceso_result(request: Request, salas: List[str] = Form([])):
    r = require_login(request)
    if r:
        return r
    u = user_from_session(request)
    if u["rol"] != "ENCARGADO":
        return RedirectResponse(role_home_path(u["rol"]), status_code=303)

    salas_filtro = sanitize_salas_selection(salas)
    rows = _query_partes_en_proceso_filtrado(salas_filtro)

    filtro_txt = "TODAS LAS SALAS" if not salas_filtro else ", ".join(salas_filtro)
    body = render_ticket_blocks(
        rows=rows,
        back_href="/encargado/gestion_partes",
        title="Partes en proceso (visualización)",
        subtitle=f"Filtro de salas: {filtro_txt}",
        show_link=True,
    )
    return page("Encargado - Visualizar", body)


@app.get("/encargado/pdf", response_class=HTMLResponse)
def admin_pdf_form(request: Request):
    r = require_login(request)
    if r:
        return r
    u = user_from_session(request)
    if u["rol"] != "ENCARGADO":
        return RedirectResponse(role_home_path(u["rol"]), status_code=303)

    salas = get_salas()
    selector = salas_multiselect_html(salas, None, "Selecciona sala(s) para generar el PDF (o TODAS)")

    body = f"""
    <div class="top">
      <div><h2>Generar PDF - Partes en proceso</h2></div>
      <div><a class="btn2" href="/encargado/gestion_partes">Volver</a></div>
    </div>

    <div class="card">
      <form method="post" action="/encargado/pdf">
        {selector}
        <div style="margin-top:12px">
          <button class="btn" type="submit">Generar PDF</button>
        </div>
      </form>
    </div>
    """
    return page("PDF - Filtro", body)


@app.post("/encargado/pdf")
def admin_pdf_generate(request: Request, salas: List[str] = Form([])):
    r = require_login(request)
    if r:
        return r
    u = user_from_session(request)
    if u["rol"] != "ENCARGADO":
        return RedirectResponse(role_home_path(u["rol"]), status_code=303)

    salas_filtro = sanitize_salas_selection(salas)
    pdf_path = generar_pdf_partes_en_proceso(salas_filtro)
    return FileResponse(str(pdf_path), media_type="application/pdf", filename=pdf_path.name)


# =========================
# ENCARGADO - Eliminar partes
# =========================
@app.get("/encargado/eliminar_partes", response_class=HTMLResponse)
def admin_eliminar_partes_menu(request: Request):
    r = require_login(request)
    if r:
        return r
    u = user_from_session(request)
    if u["rol"] != "ENCARGADO":
        return RedirectResponse(role_home_path(u["rol"]), status_code=303)

    body = """
    <div class="top">
      <div><h2>Eliminar partes</h2></div>
      <div><a class="btn2" href="/encargado/gestion_partes">Volver</a></div>
    </div>

    <div class="card">
      <p>¿Qué lista quieres revisar?</p>
      <div class="row">
        <a class="btn danger" href="/encargado/eliminar_partes/lista?tipo=pendientes">Pendientes / en curso</a>
        <a class="btn danger" href="/encargado/eliminar_partes/lista?tipo=finalizados">Finalizados</a>
      </div>
    </div>
    """
    return page("Eliminar partes", body)


@app.get("/encargado/eliminar_partes/lista", response_class=HTMLResponse)
def admin_eliminar_partes_lista(request: Request, tipo: str = "pendientes"):
    r = require_login(request)
    if r:
        return r
    u = user_from_session(request)
    if u["rol"] != "ENCARGADO":
        return RedirectResponse(role_home_path(u["rol"]), status_code=303)

    finalizados = (tipo or "").lower() == "finalizados"
    if finalizados:
        rows = db_all(
            """
            select referencia, created_at, created_by_name, room_name, estado_encargado
            from public.wom_tickets
            where estado_encargado in ('TRABAJO TERMINADO/REPARADO','TRABAJO DESESTIMADO')
            order by created_at desc;
        """
        )
        titulo = "Finalizados"
    else:
        rows = db_all(
            """
            select referencia, created_at, created_by_name, room_name, estado_encargado
            from public.wom_tickets
            where estado_encargado not in ('TRABAJO TERMINADO/REPARADO','TRABAJO DESESTIMADO')
            order by created_at desc;
        """
        )
        titulo = "Pendientes / en curso"

    trs = ""
    for p in rows:
        f, hh = formatear_fecha_hora(p.get("created_at"))
        ref = (p.get("referencia") or "").strip()
        trs += f"""
        <tr>
          <td>{h(ref)}</td>
          <td>{h(f)} {h(hh)}</td>
          <td>{h(p.get("created_by_name",""))}</td>
          <td>{h(p.get("room_name",""))}</td>
          <td>{h(p.get("estado_encargado","SIN ESTADO"))}</td>
          <td><a class="btn danger" href="/encargado/eliminar_partes/confirmar/{h(ref)}">Eliminar</a></td>
        </tr>
        """

    body = f"""
    <div class="top">
      <div><h2>Eliminar partes - {h(titulo)}</h2></div>
      <div><a class="btn2" href="/encargado/eliminar_partes">Volver</a></div>
    </div>
    <div class="card">
      <table>
        <thead><tr><th>Ref</th><th>Fecha</th><th>Autor</th><th>Sala</th><th>Estado</th><th></th></tr></thead>
        <tbody>{trs or "<tr><td colspan='6'>No hay partes.</td></tr>"}</tbody>
      </table>
    </div>
    """
    return page("Eliminar partes", body)


@app.get("/encargado/eliminar_partes/confirmar/{ref}", response_class=HTMLResponse)
def admin_eliminar_partes_confirmar(request: Request, ref: str):
    r = require_login(request)
    if r:
        return r
    u = user_from_session(request)
    if u["rol"] != "ENCARGADO":
        return RedirectResponse(role_home_path(u["rol"]), status_code=303)

    body = f"""
    <div class="card">
      <h2>Confirmación</h2>
      <p>¿Realmente quiere eliminar el parte <b>{h(ref)}</b>?</p>
      <div class="row" style="margin-top:12px">
        <form method="post" action="/encargado/eliminar_partes/confirmar/{h(ref)}">
          <button class="btn danger" type="submit">Sí, eliminar</button>
        </form>
        <a class="btn2" href="/encargado/eliminar_partes">No, volver</a>
      </div>
      <p class="muted" style="margin-top:10px">Esta acción es irreversible.</p>
    </div>
    """
    return page("Confirmar eliminación", body)


@app.post("/encargado/eliminar_partes/confirmar/{ref}")
def admin_eliminar_partes_do(request: Request, ref: str):
    r = require_login(request)
    if r:
        return r
    u = user_from_session(request)
    if u["rol"] != "ENCARGADO":
        return RedirectResponse(role_home_path(u["rol"]), status_code=303)

    rref = (ref or "").strip().upper()
    db_exec("delete from public.wom_tickets where referencia=%s;", (rref,))
    return RedirectResponse("/encargado/gestion_partes", status_code=303)


# =========================
# ENCARGADO - Gestión de Usuarios
# =========================
@app.get("/encargado/gestion_usuarios", response_class=HTMLResponse)
def admin_gestion_usuarios(request: Request):
    r = require_login(request)
    if r:
        return r
    u = user_from_session(request)
    if u["rol"] != "ENCARGADO":
        return RedirectResponse(role_home_path(u["rol"]), status_code=303)

    body = """
    <div class="top">
      <div><h2>Gestión de Usuarios</h2></div>
      <div><a class="btn2" href="/encargado">Volver</a></div>
    </div>

    <div class="card">
      <div class="row">
        <a class="btn" href="/encargado/usuarios/listar">Listar Usuarios del sistema</a>
        <a class="btn danger" href="/encargado/usuarios/eliminar">Eliminar Usuario</a>
        <a class="btn" href="/encargado/usuarios/crear">Crear Usuario</a>
        <a class="btn" href="/encargado/salas">Gestionar las Salas de Escape</a>
      </div>
    </div>
    """
    return page("Gestión de Usuarios", body)


@app.get("/encargado/usuarios/listar", response_class=HTMLResponse)
def admin_listar_usuarios(request: Request):
    r = require_login(request)
    if r:
        return r
    u = user_from_session(request)
    if u["rol"] != "ENCARGADO":
        return RedirectResponse(role_home_path(u["rol"]), status_code=303)

    users = db_all("select code, name, role from public.wom_users order by role, name;")
    counts = db_all("select created_by_code as code, count(*)::int as n from public.wom_tickets group by created_by_code;")
    count_map = { (c.get("code") or "").upper(): int(c.get("n") or 0) for c in counts }

    rows = ""
    for us in users:
        code = (us.get("code") or "").strip()
        n = count_map.get(code.upper(), 0)
        rows += f'''
        <tr>
          <td>{h(code)}</td>
          <td>{h(us.get("name",""))}</td>
          <td>{h(us.get("role",""))}</td>
          <td style="text-align:right">{n}</td>
        </tr>
        '''

    body = f'''
    <div class="top">
      <div><h2>Usuarios del sistema</h2></div>
      <div><a class="btn2" href="/encargado/gestion_usuarios">Volver</a></div>
    </div>

    <div class="card">
      <table>
        <thead><tr><th>Código</th><th>Nombre</th><th>Rol</th><th>Partes emitidos</th></tr></thead>
        <tbody>{rows or "<tr><td colspan='4'>No hay usuarios.</td></tr>"}</tbody>
      </table>
    </div>
    '''
    return page("Listar Usuarios", body)



@app.get("/encargado/usuarios/crear", response_class=HTMLResponse)
def admin_crear_usuario_form(request: Request):
    r = require_login(request)
    if r:
        return r
    u = user_from_session(request)
    if u["rol"] != "ENCARGADO":
        return RedirectResponse(role_home_path(u["rol"]), status_code=303)

    body = """
    <div class="top">
      <div><h2>Crear Usuario</h2></div>
      <div><a class="btn2" href="/encargado/gestion_usuarios">Volver</a></div>
    </div>

    <div class="card">
      <form method="post" action="/encargado/usuarios/crear">
        <label>Código (ej: X123Y)</label>
        <input name="codigo" autocomplete="off"/>

        <label>Nombre</label>
        <input name="nombre" autocomplete="off"/>

        <label>Rol</label>
        <select name="rol">
          <option value="TRABAJADOR">TRABAJADOR</option>
          <option value="JEFE">JEFE</option>
          <option value="ENCARGADO">ENCARGADO</option>
        </select>

        <div style="margin-top:12px">
          <button class="btn" type="submit">Crear</button>
        </div>
      </form>
    </div>
    """
    return page("Crear Usuario", body)


@app.post("/encargado/usuarios/crear")
def admin_crear_usuario_do(
    request: Request,
    codigo: str = Form(...),
    nombre: str = Form(...),
    rol: str = Form(...),
):
    r = require_login(request)
    if r:
        return r
    u = user_from_session(request)
    if u["rol"] != "ENCARGADO":
        return RedirectResponse(role_home_path(u["rol"]), status_code=303)

    c = (codigo or "").strip().upper()
    n = (nombre or "").strip()
    rr = (rol or "").strip().upper()

    if not c or not n or rr not in {"TRABAJADOR", "JEFE", "ENCARGADO"}:
        return HTMLResponse(
            page(
                "Error",
                "<div class='card'><h3>Datos inválidos</h3><p><a class='btn2' href='/encargado/usuarios/crear'>Volver</a></p></div>",
            ),
            status_code=400,
        )

    exists = db_one("select 1 as x from public.wom_users where upper(code)=upper(%s);", (c,))
    if exists:
        return HTMLResponse(
            page(
                "Error",
                f"<div class='card'><h3>Ya existe un usuario con código {h(c)}</h3><p><a class='btn2' href='/encargado/usuarios/crear'>Volver</a></p></div>",
            ),
            status_code=400,
        )

    db_exec("insert into public.wom_users (code, name, role) values (%s,%s,%s);", (c, n, rr))
    return RedirectResponse("/encargado/usuarios/listar", status_code=303)


@app.get("/encargado/usuarios/eliminar", response_class=HTMLResponse)
def admin_eliminar_usuario_lista(request: Request):
    r = require_login(request)
    if r:
        return r
    u = user_from_session(request)
    if u["rol"] != "ENCARGADO":
        return RedirectResponse(role_home_path(u["rol"]), status_code=303)

    users = db_all("select code, name, role from public.wom_users order by role, name;")

    rows = ""
    for us in users:
        code = us["code"]
        disabled = code.upper() == u["codigo"].upper()
        btn = "(No puedes eliminarte)" if disabled else f"<a class='btn danger' href='/encargado/usuarios/eliminar/confirmar/{h(code)}'>Eliminar</a>"
        rows += f"""
        <tr>
          <td>{h(code)}</td>
          <td>{h(us["name"])}</td>
          <td>{h(us["role"])}</td>
          <td>{btn}</td>
        </tr>
        """

    body = f"""
    <div class="top">
      <div><h2>Eliminar Usuario</h2></div>
      <div><a class="btn2" href="/encargado/gestion_usuarios">Volver</a></div>
    </div>

    <div class="card">
      <table>
        <thead><tr><th>Código</th><th>Nombre</th><th>Rol</th><th></th></tr></thead>
        <tbody>{rows or "<tr><td colspan='4'>No hay usuarios.</td></tr>"}</tbody>
      </table>
      <p class="muted" style="margin-top:10px">Eliminar un usuario NO borra los partes existentes.</p>
    </div>
    """
    return page("Eliminar Usuario", body)

@app.get("/encargado/usuarios/eliminar/confirmar/{code}", response_class=HTMLResponse)
def admin_eliminar_usuario_confirmar(request: Request, code: str):
    r = require_login(request)
    if r:
        return r
    u = user_from_session(request)
    if u["rol"] != "ENCARGADO":
        return RedirectResponse(role_home_path(u["rol"]), status_code=303)

    code = (code or "").strip().upper()
    target = get_user_by_code(code)
    if not target:
        return RedirectResponse("/encargado/usuarios/eliminar", status_code=303)

    if code == (u.get("codigo") or "").strip().upper():
        msg = "No puedes eliminar tu propio usuario."
        body = f'''
        <div class="top"><div><h2>Eliminar usuario</h2></div><div><a class="btn2" href="/encargado/usuarios/eliminar">Volver</a></div></div>
        <div class="card"><p style="font-weight:700; color:#b00000"><b>{h(msg)}</b></p></div>
        '''
        return page("Eliminar usuario", body)

    body = f'''
    <div class="top">
      <div><h2>Eliminar usuario</h2></div>
      <div><a class="btn2" href="/encargado/usuarios/eliminar">Volver</a></div>
    </div>
    <div class="card">
      <p>Vas a eliminar al usuario: <b>{h(target.get("name",""))}</b> ({h(target.get("role",""))})</p>
      <p class="muted">Código: {h(code)}</p>
      <p><b>¿Realmente quieres eliminar este usuario?</b></p>
      <form method="post" action="/encargado/usuarios/eliminar/confirmar/{h(code)}">
        <button class="btn" type="submit">Sí, eliminar</button>
        <a class="btn2" href="/encargado/usuarios/eliminar" style="margin-left:8px">Cancelar</a>
      </form>
    </div>
    '''
    return page("Confirmar eliminación", body)


@app.post("/encargado/usuarios/eliminar/confirmar/{code}")
def admin_eliminar_usuario_confirmar_post(request: Request, code: str):
    r = require_login(request)
    if r:
        return r
    u = user_from_session(request)
    if u["rol"] != "ENCARGADO":
        return RedirectResponse(role_home_path(u["rol"]), status_code=303)

    code = (code or "").strip().upper()
    if code == (u.get("codigo") or "").strip().upper():
        return RedirectResponse("/encargado/usuarios/eliminar", status_code=303)

    # no permitir eliminar el encargado principal
    if code == "P000A":
        return RedirectResponse("/encargado/usuarios/eliminar", status_code=303)

    db_exec("delete from public.wom_users where code=%s;", (code,))
    return RedirectResponse("/encargado/usuarios/eliminar", status_code=303)



@app.get("/encargado/salas", response_class=HTMLResponse)
def admin_salas(request: Request):
    r = require_login(request)
    if r:
        return r
    u = user_from_session(request)
    if u["rol"] != "ENCARGADO":
        return RedirectResponse(role_home_path(u["rol"]), status_code=303)

    salas = get_salas()
    items = "".join([f"<li>{h(s)}</li>" for s in salas]) or "<li>No hay salas.</li>"

    body = f"""
    <div class="top">
      <div><h2>Gestionar Salas de Escape</h2></div>
      <div><a class="btn2" href="/encargado/gestion_usuarios">Volver</a></div>
    </div>

    <div class="card">
      <h3>Salas actuales</h3>
      <ul>{items}</ul>
    </div>

    <div class="card">
      <h3>Añadir sala</h3>
      <form method="post" action="/encargado/salas">
        <label>Nombre de la sala</label>
        <input name="sala" autocomplete="off" placeholder="Ej: NUEVA SALA"/>
        <div style="margin-top:12px">
          <button class="btn" type="submit">Añadir</button>
        </div>
      </form>
      <p class="muted" style="margin-top:10px">Estas salas aparecerán en el desplegable de “Nuevo parte”.</p>
    </div>
    """
    return page("Salas", body)


@app.post("/encargado/salas")
def admin_salas_add(request: Request, sala: str = Form(...)):
    r = require_login(request)
    if r:
        return r
    u = user_from_session(request)
    if u["rol"] != "ENCARGADO":
        return RedirectResponse(role_home_path(u["rol"]), status_code=303)

    s = (sala or "").strip()
    if not s:
        return RedirectResponse("/encargado/salas", status_code=303)

    db_exec("insert into public.wom_rooms (name) values (%s) on conflict (name) do nothing;", (s,))
    return RedirectResponse("/encargado/salas", status_code=303)
