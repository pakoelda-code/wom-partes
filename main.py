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
# Pillow
#
# Variables de entorno:
# DATABASE_URL   (Supabase Pooler, p.ej. ...pooler.supabase.com:6543/postgres)
# SESSION_SECRET (recomendado)

import os
import random
import string
import math
import urllib.request
import urllib.error
import urllib.parse
import mimetypes
import json
from io import BytesIO

# Pillow (compresión de imágenes en servidor). Si no está instalado, se mostrará un error claro al subir imágenes.
PIL_AVAILABLE = True
try:
    from PIL import Image, ImageOps  # type: ignore
except Exception:
    PIL_AVAILABLE = False

MAX_IMG_BYTES = 100 * 1024   # 100 KB por imagen final en Storage
MAX_IMG_DIM = 1280           # máximo ancho/alto

def compress_image_to_target(image_bytes: bytes, target_bytes: int = MAX_IMG_BYTES) -> bytes:
    """Convierte la imagen a WEBP y ajusta tamaño/calidad para intentar <= target_bytes."""
    if not PIL_AVAILABLE:
        raise RuntimeError("Falta la dependencia Pillow en el servidor. Añade 'Pillow' a requirements.txt y redeploy.")
    img = Image.open(BytesIO(image_bytes))
    img = ImageOps.exif_transpose(img)

    # Normaliza modo
    if img.mode in ("RGBA", "LA") or (img.mode == "P" and "transparency" in getattr(img, "info", {})):
        bg = Image.new("RGB", img.size, (255, 255, 255))
        rgba = img.convert("RGBA")
        bg.paste(rgba, mask=rgba.split()[-1])
        img = bg
    else:
        img = img.convert("RGB")

    # Reescalado
    img.thumbnail((MAX_IMG_DIM, MAX_IMG_DIM))

    last = b""
    for quality in [80, 70, 60, 50, 40, 30, 25, 20]:
        out = BytesIO()
        img.save(out, format="WEBP", quality=quality, method=6)
        data = out.getvalue()
        last = data
        if len(data) <= target_bytes:
            return data

    for dim in [1024, 900, 800, 700, 600]:
        tmp = img.copy()
        tmp.thumbnail((dim, dim))
        for quality in [60, 50, 40, 30, 25, 20]:
            out = BytesIO()
            tmp.save(out, format="WEBP", quality=quality, method=6)
            data = out.getvalue()
            last = data
            if len(data) <= target_bytes:
                return data

    return last
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple
from zoneinfo import ZoneInfo

import psycopg2
from psycopg2.extras import RealDictCursor

from fastapi import FastAPI, Request, Form, UploadFile, File
from fastapi.responses import HTMLResponse, RedirectResponse, FileResponse, PlainTextResponse
from starlette.middleware.sessions import SessionMiddleware

TZ = ZoneInfo("Europe/Madrid")

TIPOS = ["ELECTRÓNICA", "MOBILIARIO", "ESTRUCTURA", "ELEMENTOS SUELTOS", "OTROS/AS"]



# ------------ Prioridades ------------
PRIORIDADES = [
    ("URGENTE", "Urgente", "#b00000"),
    ("MEDIO", "Medio", "#d97706"),
    ("DEMORABLE", "Demorable", "#15803d"),
]
PRIORIDAD_COLOR = {k: color for (k, _label, color) in PRIORIDADES}
PRIORIDADES_VALIDAS = {p[0] for p in PRIORIDADES}

def prio_label(prio: str) -> str:
    p = (prio or "").strip().upper()
    if p == "URGENTE":
        return "Urgente"
    if p == "DEMORABLE":
        return "Demorable"
    return "Medio"

def prio_css_class(prio: str) -> str:
    p = (prio or "").strip().upper()
    if p == "URGENTE":
        return "prio-urg"
    if p == "DEMORABLE":
        return "prio-dem"
    return "prio-med"

def prio_span(prio: str, txt: str) -> str:
    return f"<span class='{prio_css_class(prio)}'>{h(txt or '')}</span>"
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



def prio_badge(prio: str) -> str:
    """Devuelve un span coloreado para el texto de prioridad."""
    p = (prio or "MEDIO").strip().upper()
    if p == "URGENTE":
        return "<span style='font-weight:800;color:#d00;'>Urgente</span>"
    if p == "DEMORABLE":
        return "<span style='font-weight:700;color:#1b7a1b;'>Demorable</span>"
    return "<span style='font-weight:700;color:#d57a00;'>Medio</span>"


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
        """
    create table if not exists public.wom_hours (
      id bigserial primary key,
      worker_code text not null references public.wom_users(code) on delete cascade,
      worker_name text not null,
      room_name text not null,

      entry_at timestamptz null,
      exit_at timestamptz null,

      recorded_by_code text null references public.wom_users(code) on delete set null,
      recorded_by_name text not null default '',

      created_at timestamptz not null default now()
    );
    """
    )
    db_exec("create index if not exists wom_hours_worker_idx on public.wom_hours(worker_code);")
    db_exec("create index if not exists wom_hours_entry_idx on public.wom_hours(entry_at desc);")
    db_exec("create index if not exists wom_hours_room_idx on public.wom_hours(room_name);")

    
    # Migración suave (si la tabla ya existía)
    db_exec("alter table public.wom_tickets add column if not exists priority text not null default 'MEDIO';")
    db_exec("alter table public.wom_tickets add column if not exists image_url text;")
    db_exec("alter table public.wom_tickets add column if not exists image_path text;")
    # Tabla de imágenes por parte (hasta 3)
    db_exec(
        """
    create table if not exists public.wom_ticket_images (
      id bigserial primary key,
      ticket_id bigint not null references public.wom_tickets(id) on delete cascade,
      position int not null check (position between 1 and 3),
      image_url text not null,
      image_path text not null,
      created_at timestamptz not null default now(),
      unique(ticket_id, position)
    );
    """
    )
    db_exec("create index if not exists wom_ticket_images_ticket_idx on public.wom_ticket_images(ticket_id);")
    db_exec("create index if not exists wom_tickets_priority_idx on public.wom_tickets(priority);")
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


def month_bounds(year: int, month: int):
    """Devuelve (inicio, fin) del mes en zona Europe/Madrid (tz-aware)."""
    y = int(year); m_ = int(month)
    if m_ < 1 or m_ > 12:
        raise ValueError("Mes inválido")
    start = datetime(y, m_, 1, 0, 0, 0, tzinfo=TZ)
    if m_ == 12:
        end = datetime(y + 1, 1, 1, 0, 0, 0, tzinfo=TZ)
    else:
        end = datetime(y, m_ + 1, 1, 0, 0, 0, tzinfo=TZ)
    return start, end


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


def _safe_ext(filename: str) -> str:
    name = (filename or "").strip().lower()
    if "." not in name:
        return ""
    ext = "." + name.rsplit(".", 1)[-1]
    if ext in (".jpg", ".jpeg", ".png", ".webp"):
        return ext
    return ""


def supabase_storage_upload(bucket: str, path: str, file_bytes: bytes, content_type: str) -> str:
    """
    Sube un objeto a Supabase Storage usando la API REST. Devuelve URL pública.
    Requiere bucket público, o bien que luego uses URLs firmadas (no implementado aquí).
    """
    supabase_url = (os.getenv("SUPABASE_URL", "") or "").strip().rstrip("/")
    key = (
        (os.getenv("SUPABASE_SERVICE_ROLE_KEY", "") or "").strip()
        or (os.getenv("SUPABASE_SERVICE_KEY", "") or "").strip()
        or (os.getenv("SUPABASE_KEY", "") or "").strip()
    )
    if not supabase_url or not key:
        raise RuntimeError("Falta SUPABASE_URL y/o SUPABASE_SERVICE_ROLE_KEY (o SUPABASE_KEY)")

    url = f"{supabase_url}/storage/v1/object/{bucket}/{path}"
    headers = {
        "Authorization": f"Bearer {key}",
        "apikey": key,
        "Content-Type": content_type or "application/octet-stream",
        "x-upsert": "true",
    }
    req = urllib.request.Request(url, data=file_bytes, headers=headers, method="PUT")
    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            resp.read()
    except urllib.error.HTTPError as e:
        body = (e.read() or b"").decode("utf-8", errors="ignore")
        raise RuntimeError(f"Error subiendo imagen: {e.code} {e.reason} {body}")

    return f"{supabase_url}/storage/v1/object/public/{bucket}/{path}"




def _supabase_creds() -> Tuple[str, str]:
    supabase_url = (os.getenv("SUPABASE_URL", "") or "").strip().rstrip("/")
    key = (
        (os.getenv("SUPABASE_SERVICE_ROLE_KEY", "") or "").strip()
        or (os.getenv("SUPABASE_SERVICE_KEY", "") or "").strip()
        or (os.getenv("SUPABASE_KEY", "") or "").strip()
    )
    if not supabase_url or not key:
        raise RuntimeError("Falta SUPABASE_URL y/o SUPABASE_SERVICE_ROLE_KEY (o SUPABASE_KEY)")
    return supabase_url, key


def supabase_storage_remove(bucket: str, paths: List[str]) -> None:
    """Elimina uno o varios objetos del bucket (por path) usando la API REST."""
    if not paths:
        return
    supabase_url, key = _supabase_creds()
    bucket = (bucket or "").strip() or "partes"
    url = f"{supabase_url}/storage/v1/object/remove/{bucket}"
    payload = json.dumps({"prefixes": paths}).encode("utf-8")
    headers = {
        "Authorization": f"Bearer {key}",
        "apikey": key,
        "Content-Type": "application/json",
    }
    req = urllib.request.Request(url, data=payload, headers=headers, method="POST")
    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            resp.read()
    except urllib.error.HTTPError as e:
        body = (e.read() or b"").decode("utf-8", errors="ignore")
        raise RuntimeError(f"Error borrando imagen: {e.code} {e.reason} {body}")


def cleanup_ticket_images(ticket_id: int) -> None:
    """Borra imágenes asociadas a un parte: BD + Storage. No lanza si falla Storage."""
    if not ticket_id:
        return
    bucket = (os.getenv("SUPABASE_STORAGE_BUCKET", "") or "").strip() or "partes"

    # Paths en tabla nueva
    rows = db_all("select image_path from public.wom_ticket_images where ticket_id=%s order by position asc;", (ticket_id,))
    paths = [(r.get("image_path") or "").strip() for r in rows if (r.get("image_path") or "").strip()]

    # Path legacy en wom_tickets
    legacy = db_one("select image_path from public.wom_tickets where id=%s;", (ticket_id,))
    if legacy and (legacy.get("image_path") or "").strip():
        paths.append((legacy.get("image_path") or "").strip())

    # Deduplicar
    paths = list(dict.fromkeys(paths))

    # Intentar borrar en Storage
    if paths:
        try:
            supabase_storage_remove(bucket, paths)
        except Exception:
            # No rompemos flujo si falla el borrado remoto
            pass

    # Limpiar BD
    db_exec("delete from public.wom_ticket_images where ticket_id=%s;", (ticket_id,))
    db_exec("update public.wom_tickets set image_url=null, image_path=null where id=%s;", (ticket_id,))

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
              priority,
              descripcion,
              solucionado_por_usuario,
              reparacion_usuario,
              visto_por_encargado,
              estado_encargado,
              observaciones_encargado
            from public.wom_tickets
            where (estado_encargado is null or estado_encargado not in ('TRABAJO TERMINADO/REPARADO','TRABAJO DESESTIMADO'))
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
          priority,
          descripcion,
          solucionado_por_usuario,
          reparacion_usuario,
          visto_por_encargado,
          estado_encargado,
          observaciones_encargado
        from public.wom_tickets
        where (estado_encargado is null or estado_encargado not in ('TRABAJO TERMINADO/REPARADO','TRABAJO DESESTIMADO'))
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
    st_mono = ParagraphStyle("mono", parent=styles["Normal"], fontName="Courier", fontSize=8.5, leading=10, spaceAfter=1)

    def e(s: str) -> str:
        return _xml_escape(s or "").replace("\n", "<br/>")

    filtro_txt = "TODAS" if not salas_filtro else ", ".join(salas_filtro)
    story = []
    story.append(Paragraph("Relación de Partes en Proceso", st_title))
    story.append(Paragraph(f"Salas: <b>{e(filtro_txt)}</b> — Generado: {now_madrid().strftime('%d/%m/%Y %H:%M')}", st_line))
    story.append(Spacer(1, 14))

    azul_sala = "#003366"

    for p in rows:
        fecha, hora = formatear_fecha_hora(p.get("created_at"))
        ref = (p.get("referencia") or "").strip()
        sala = p.get("room_name") or ""
        tipo = p.get("tipo") or ""
        prio = (p.get("priority") or "MEDIO").upper()
        prio_opts_sel = "\n".join([
            f"<option value='{h(k)}'" + (" selected" if k == prio else "") + f">{h(v)}</option>"
            for k, v, _ in PRIORIDADES
        ])
        autor = p.get("created_by_name") or ""
        estado = p.get("estado_encargado") or "SIN ESTADO"

        desc = p.get("descripcion") or ""
        rep = p.get("reparacion_usuario") or ""
        com = p.get("observaciones_encargado") or ""

        # Línea 1: Ref / Fecha-Hora / Sala
        line1 = (
            f"<b>Ref:</b> {e(ref)}&nbsp;&nbsp;&nbsp;"
            f"<b>Fecha y hora:</b> {e(fecha)} {e(hora)}&nbsp;&nbsp;&nbsp;"
            f"<b>Sala:</b> <font color='{azul_sala}'><b>{e(sala)}</b></font>"
        )
        # Línea 2: Tipo / Prioridad / Usuario / Estado
        line2 = (
            f"<b>Tipo:</b> {e(tipo)}&nbsp;&nbsp;&nbsp;"
            f"<b>Nivel de prioridad:</b> {e(prio_label(prio))}&nbsp;&nbsp;&nbsp;"
            f"<b>Usuario:</b> {e(autor)}&nbsp;&nbsp;&nbsp;"
            f"<b>Estado:</b> {e(estado)}"
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

        story.append(Spacer(1, 10))
        story.append(HRFlowable(width="100%", thickness=1.2, color=colors.black))
        story.append(Spacer(1, 10))

    doc.build(story)
    return out_path


    for p in rows:
        fecha, hora = formatear_fecha_hora(p.get("created_at"))
        ref = (p.get("referencia") or "").strip()
        sala = p.get("room_name") or ""
        tipo = p.get("tipo") or ""
        autor = p.get("created_by_name") or ""
        estado = p.get("estado_encargado") or "SIN ESTADO"
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
  .prio-urg{{color:#b00000;font-weight:800;}}
  .prio-med{{color:#d97706;font-weight:800;}}
  .prio-dem{{color:#15803d;font-weight:800;}}
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
    image_url = (p.get("image_url") or "").strip()
    img_block = ""
    if image_url:
        img_block = (
            f"<p><b>Imagen adjunta:</b> <a class='btn2' href='{h(image_url)}' target='_blank'>Ver imagen</a></p>"
            f"<div style='margin-top:10px'><img src='{h(image_url)}' style='max-width:100%; border:1px solid #ddd; border-radius:12px'/></div>"
        )
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
      <form method="post" action="/trabajador/nuevo" enctype="multipart/form-data">
        <input type="hidden" name="referencia" value="{h(ref)}"/>

        <label>Sala</label>
        <select name="sala">{salas_opts}</select>

        <label>Tipo</label>
        <select name="tipo">{tipos_opts}</select>

        
        <label>Nivel de prioridad</label>
        <select name="priority" required>
          <option value="URGENTE" style="color:#b00000;font-weight:800;">Urgente</option>
          <option value="MEDIO" selected style="color:#d97706;font-weight:800;">Medio</option>
          <option value="DEMORABLE" style="color:#15803d;font-weight:800;">Demorable</option>
        </select>

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

        <label>Imágenes (opcional, máx 3). Se comprimen automáticamente.</label>
        <input type="file" name="imagenes" accept="image/*" multiple/>

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
    priority: str = Form('MEDIO'),
    descripcion: str = Form(""),
    solucionado: str = Form("NO"),
    reparacion_usuario: str = Form(""),
    imagenes: List[UploadFile] = File([]),
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
    prio = (priority or "MEDIO").strip().upper()
    if prio not in ("URGENTE", "MEDIO", "DEMORABLE"):
        prio = "MEDIO"
    desc = (descripcion or "").strip() or "(Sin descripción)"

    sol = (solucionado or "").strip().upper() == "SI"
    rep = (reparacion_usuario or "").strip() if sol else ""

    image_url = None
    image_path = None

    # --- Manejo de hasta 3 imágenes (se comprimen a ~100KB y se convierten a WEBP) ---
    files: List[UploadFile] = []
    if imagenes:
        for f in imagenes:
            if f is not None and getattr(f, "filename", ""):
                files.append(f)

    if len(files) > 3:
        return HTMLResponse(
            page("Error", "<div class='card'><h3>Máximo 3 imágenes por parte</h3><p><a class='btn2' href='/trabajador/nuevo'>Volver</a></p></div>"),
            status_code=400,
        )

    # Inserta primero el ticket para obtener ticket_id
    room = db_one("select id, name from public.wom_rooms where name=%s;", (sala_name,))
    room_id = room["id"] if room else None

    db_exec(
        """
        insert into public.wom_tickets
        (referencia, created_by_code, created_by_name, room_id, room_name, tipo, priority, descripcion,
         solucionado_por_usuario, reparacion_usuario, image_url, image_path, visto_por_encargado, estado_encargado, observaciones_encargado)
        values
        (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, false, 'SIN ESTADO', '')
        on conflict (referencia) do nothing;
        """,
        (ref, u["codigo"], u["nombre"], room_id, sala_name, tipo_name, prio, desc, sol, rep, None, None),
    )

    ticket_row = db_one("select id from public.wom_tickets where referencia=%s;", (ref,))
    ticket_id = ticket_row["id"] if ticket_row else None

    if files and not ticket_id:
        return HTMLResponse(
            page("Error", "<div class='card'><h3>No se pudo obtener el ID del parte para guardar imágenes</h3><p><a class='btn2' href='/trabajador/nuevo'>Volver</a></p></div>"),
            status_code=500,
        )

    if files:
        bucket = os.getenv("SUPABASE_STORAGE_BUCKET", "partes")
        ts = now_madrid().strftime("%Y%m%d_%H%M%S")
        for pos, f in enumerate(files, start=1):
            raw = f.file.read()
            if not raw:
                continue

            # Límite de entrada (para no reventar memoria/tiempo)
            if len(raw) > 8 * 1024 * 1024:
                return HTMLResponse(
                    page("Error", "<div class='card'><h3>Una de las imágenes supera 8MB</h3><p><a class='btn2' href='/trabajador/nuevo'>Volver</a></p></div>"),
                    status_code=400,
                )

            try:
                compressed = compress_image_to_target(raw, MAX_IMG_BYTES)
            except Exception as ex:
                return HTMLResponse(
                    page("Error", f"<div class='card'><h3>Error procesando la imagen</h3><p class='muted'>{h(str(ex))}</p><p><a class='btn2' href='/trabajador/nuevo'>Volver</a></p></div>"),
                    status_code=500,
                )

            token = "".join(random.choice(string.ascii_lowercase + string.digits) for _ in range(6))
            image_path_i = f"tickets/{ref}_{ts}_{token}_{pos}.webp"
            try:
                image_url_i = supabase_storage_upload(bucket, image_path_i, compressed, "image/webp")
            except Exception as ex:
                return HTMLResponse(
                    page("Error", f"<div class='card'><h3>Error subiendo la imagen</h3><p class='muted'>{h(str(ex))}</p><p><a class='btn2' href='/trabajador/nuevo'>Volver</a></p></div>"),
                    status_code=500,
                )

            # Inserta en tabla de imágenes
            try:
                db_exec(
                    """
                    insert into public.wom_ticket_images (ticket_id, position, image_url, image_path)
                    values (%s, %s, %s, %s)
                    on conflict (ticket_id, position) do update set image_url=excluded.image_url, image_path=excluded.image_path;
                    """,
                    (ticket_id, pos, image_url_i, image_path_i),
                )
            except Exception:
                pass

            if pos == 1:
                image_url = image_url_i
                image_path = image_path_i

        # Guarda la primera imagen también en wom_tickets (compatibilidad)
        if image_url:
            db_exec(
                "update public.wom_tickets set image_url=%s, image_path=%s where id=%s;",
                (image_url, image_path, ticket_id),
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
        select referencia, created_at, created_by_name, room_name, tipo, priority, estado_encargado, visto_por_encargado
        from public.wom_tickets
        where (estado_encargado is null or estado_encargado not in ('TRABAJO TERMINADO/REPARADO','TRABAJO DESESTIMADO'))
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
          <td>{prio_span(p.get("priority"), p.get("estado_encargado","SIN ESTADO"))}</td>
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
          <td>{prio_span(prio, estado)}</td>
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
        select referencia, created_at, created_by_name, room_name, tipo, priority, estado_encargado, visto_por_encargado
        from public.wom_tickets
        where (estado_encargado is null or estado_encargado not in ('TRABAJO TERMINADO/REPARADO','TRABAJO DESESTIMADO'))
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
          <td>{prio_span(p.get("priority"), p.get("estado_encargado","SIN ESTADO"))}</td>
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

    now = now_madrid()
    mes = (request.query_params.get("mes") or str(now.month)).strip()
    anio = (request.query_params.get("anio") or str(now.year)).strip()

    rows = []
    error = ""
    try:
        mes_i = int(mes); anio_i = int(anio)
        ts_start, ts_end = month_bounds(anio_i, mes_i)
        rows = db_all(
            """
            select referencia, created_at, created_by_name, room_name, tipo, priority, estado_encargado
            from public.wom_tickets
            where estado_encargado in ('TRABAJO TERMINADO/REPARADO','TRABAJO DESESTIMADO')
              and created_at >= %s and created_at < %s
            order by created_at desc;
            """,
            (ts_start, ts_end),
        )
    except Exception as e:
        error = str(e)

    trs = ""
    for p in rows:
        f, hh = formatear_fecha_hora(p.get("created_at"))
        ref = (p.get("referencia") or "").strip()
        trs += f"""
        <tr>
          <td><a href="/parte/{h(ref)}">{h(ref)}</a></td>
          <td>{h(f)} {h(hh)}</td>
          <td>{h(p.get("created_by_name",""))}</td>
          <td>{h(p.get("room_name",""))}</td>
          <td>{h(p.get("tipo",""))}</td>
          <td>{prio_span(p.get("priority"), p.get("estado_encargado","SIN ESTADO"))}</td>
        </tr>
        """

    body = f"""
    <div class="top">
      <div><h2>Ver listado de partes finalizados</h2></div>
      <div><a class="btn2" href="/jefe">Volver</a></div>
    </div>

    <div class="card">
      <form method="get" action="/jefe/finalizados">
        <div class="grid2">
          <div>
            <label>Mes</label>
            <input name="mes" type="number" min="1" max="12" value="{h(mes)}" required>
          </div>
          <div>
            <label>Año</label>
            <input name="anio" type="number" min="2000" max="2100" value="{h(anio)}" required>
          </div>
        </div>
        <button class="btn" type="submit">Filtrar</button>
      </form>
      {f"<p class='warn'>Error en filtro: {h(error)}</p>" if error else ""}
    </div>

    <div class="card">
      <table>
        <thead><tr><th>Ref</th><th>Fecha</th><th>Autor</th><th>Sala</th><th>Tipo</th><th>Estado</th></tr></thead>
        <tbody>{trs or "<tr><td colspan='6'>No hay partes.</td></tr>"}</tbody>
      </table>
    </div>
    """
    return page("Finalizados", body)


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
    prio_current = (p.get("priority") or "MEDIO").upper()
    prio = prio_current
    prio_color = PRIORIDAD_COLOR.get(prio_current, "#f39c12")
    prio_badge_html = f"<b style='color:{prio_color}'>{h(prio_label(prio_current))}</b>"
    prio_options_html = "".join(
    [
    f"<option value='{h(k)}' {'selected' if k==prio_current else ''}>{h(v)}</option>"
    for k, v, _c in PRIORIDADES
    ]
    )
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

    # --- Imágenes adjuntas (hasta 3) ---
    imgs: List[str] = []
    try:
        if p.get("id"):
            rows = db_all(
                "select image_url from public.wom_ticket_images where ticket_id=%s order by position asc;",
                (p["id"],),
            )
            imgs = [(r.get("image_url") or "").strip() for r in rows if (r.get("image_url") or "").strip()]
    except Exception:
        imgs = []

    if not imgs:
        single = (p.get("image_url") or "").strip()
        if single:
            imgs = [single]

    img_block = ""
    if imgs:
        links = []
        for i, url in enumerate(imgs, start=1):
            links.append(f'<a href="{h(url)}" target="_blank">📷 Ver imagen {i}</a>')
        img_block = "<p><b>Imágenes:</b><br/>" + "<br/>".join(links) + "</p>"

    back = role_home_path(u["rol"])
    body = f"""
    <div class="top">
      <div>
        <h2>Parte {h((p.get("referencia") or "").strip())}</h2>
        <div class="pill">Fecha: {h(fecha)} {h(hora)}</div>
        <div class="pill">Visto: {h(visto)}</div>
        <div class="pill">Estado: {prio_span(prio, estado)}</div>
        <div class="pill">Prioridad: {prio_badge_html}</div>
      </div>
      <div><a class="btn2" href="{h(back)}">Volver</a></div>
    </div>

    <div class="card">
      <p><b>Sala:</b> {h(p.get("room_name",""))}</p>
      <p><b>Tipo:</b> {h(p.get("tipo",""))}</p>
      <p><b>Creado por:</b> {h(p.get("created_by_name",""))}</p>
      <p><b>¿Solucionado por el usuario?:</b> {"Sí" if sol else "No"}</p>
      {img_block}
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

          <form method="post" action="/encargado/set_priority/{ref}" style="margin-top:12px">
            <label class="small">Nivel de prioridad:</label>
            <select name="priority" class="input">
              {prio_options_html}
            </select>
            <button class="btn2" type="submit">Cambiar prioridad</button>
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

    urg_row = db_one(
        '''
        select count(*)::int as n
        from public.wom_tickets
        where coalesce(estado_encargado,'') not in ('TRABAJO TERMINADO/REPARADO','TRABAJO DESESTIMADO')
          and visto_por_encargado = false
          and upper(coalesce(priority,'')) = 'URGENTE';
        '''
    )
    urgentes_sin_ver = int((urg_row or {}).get('n') or 0)

    
    urgente_banner = ""
    if urgentes_sin_ver > 0:
        urgente_banner = f"<div style='margin-top:10px;font-weight:800;color:#d00;'>¡TIENES {urgentes_sin_ver} PARTE/S URGENTE/S!</div>"
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
        <a class="btn" href="/encargado/horas">Control de Horas</a>
      </div>
    </div>
    {urgente_banner}
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
        select referencia, created_at, created_by_name, room_name, tipo, priority, estado_encargado, visto_por_encargado
        from public.wom_tickets
        where (estado_encargado is null or estado_encargado not in ('TRABAJO TERMINADO/REPARADO','TRABAJO DESESTIMADO'))
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
          <td>{prio_span(p.get("priority"), p.get("estado_encargado","SIN ESTADO"))}</td>
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

    now = now_madrid()
    mes = (request.query_params.get("mes") or str(now.month)).strip()
    anio = (request.query_params.get("anio") or str(now.year)).strip()

    rows = []
    error = ""
    try:
        mes_i = int(mes); anio_i = int(anio)
        ts_start, ts_end = month_bounds(anio_i, mes_i)
        rows = db_all(
            """
            select referencia, created_at, created_by_name, room_name, tipo, priority, estado_encargado, visto_por_encargado
            from public.wom_tickets
            where estado_encargado in ('TRABAJO TERMINADO/REPARADO','TRABAJO DESESTIMADO')
              and created_at >= %s and created_at < %s
            order by created_at desc;
            """,
            (ts_start, ts_end),
        )
    except Exception as e:
        error = str(e)

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
          <td>{prio_span(p.get("priority"), p.get("estado_encargado","SIN ESTADO"))}</td>
          <td>{h(visto)}</td>
        </tr>
        """

    body = f"""
    <div class="top">
      <div><h2>Finalizados</h2></div>
      <div><a class="btn2" href="/encargado">Volver</a></div>
    </div>

    <div class="card">
      <form method="get" action="/encargado/finalizados">
        <div class="grid2">
          <div>
            <label>Mes</label>
            <input name="mes" type="number" min="1" max="12" value="{h(mes)}" required>
          </div>
          <div>
            <label>Año</label>
            <input name="anio" type="number" min="2000" max="2100" value="{h(anio)}" required>
          </div>
        </div>
        <button class="btn" type="submit">Filtrar</button>
      </form>
      {f"<p class='warn'>Error en filtro: {h(error)}</p>" if error else ""}
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
        if est in ('TRABAJO TERMINADO/REPARADO','TRABAJO DESESTIMADO'):
            t = ticket_por_ref(ref)
            if t and t.get('id'):
                cleanup_ticket_images(int(t['id']))
    return RedirectResponse(f"/parte/{ref}", status_code=303)



@app.post("/encargado/set_priority/{ref}")
def admin_set_priority(request: Request, ref: str, priority: str = Form("MEDIO")):
    r = require_login(request)
    if r:
        return r
    u = user_from_session(request)
    if u["rol"] != "ENCARGADO":
        return RedirectResponse(role_home_path(u["rol"]), status_code=303)

    pr = (priority or "MEDIO").strip().upper()
    if pr not in PRIORIDADES_VALIDAS:
        pr = "MEDIO"

    db_exec("update public.wom_tickets set priority=%s where referencia=%s;", (pr, ref))
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
            where (estado_encargado is null or estado_encargado not in ('TRABAJO TERMINADO/REPARADO','TRABAJO DESESTIMADO'))
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
          <td>{prio_span(p.get("priority"), p.get("estado_encargado","SIN ESTADO"))}</td>
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
    t = ticket_por_ref(rref)
    if t and t.get('id'):
        cleanup_ticket_images(int(t['id']))
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
# =========================
# ENCARGADO - Control de Horas
# =========================
def _workers_for_hours() -> List[Dict[str, str]]:
    rows = db_all(
        "select code, name, role from public.wom_users where role in ('TRABAJADOR','ENCARGADO') order by name asc;"
    )
    return [{"code": r["code"], "name": r["name"], "role": r["role"]} for r in rows]


def _round_to_half_hours(hours: float) -> float:
    if hours <= 0:
        return 0.0
    return math.floor(hours * 2 + 0.5) / 2.0


def _parse_dt_local(dt_str: str) -> Optional[datetime]:
    s = (dt_str or "").strip()
    if not s:
        return None
    try:
        dt = datetime.fromisoformat(s)
    except Exception:
        return None
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=TZ)
    else:
        dt = dt.astimezone(TZ)
    return dt


@app.get("/encargado/horas", response_class=HTMLResponse)
def horas_menu(request: Request):
    r = require_login(request)
    if r:
        return r
    u = user_from_session(request)
    if u.get("rol") != "ENCARGADO":
        return RedirectResponse(role_home_path(u.get("rol", "")), status_code=303)

    body = """
    <div class="top">
      <div><h2>Control de Horas</h2></div>
      <div><a class="btn2" href="/encargado">Volver</a></div>
    </div>

    <div class="card">
      <div class="row">
        <a class="btn" href="/encargado/horas/add">Añadir Entrada/Salida</a>
        <a class="btn" href="/encargado/horas/consultar">Consultar Horas</a>
        <a class="btn" href="/encargado/horas/pdf">Generar PDF de Horas</a>
      </div>
    </div>
    """
    return page("Control de Horas", body)


@app.get("/encargado/horas/add", response_class=HTMLResponse)
def horas_add_form(request: Request):
    r = require_login(request)
    if r:
        return r
    u = user_from_session(request)
    if u.get("rol") != "ENCARGADO":
        return RedirectResponse(role_home_path(u.get("rol", "")), status_code=303)

    workers = _workers_for_hours()
    salas = get_salas()
    w_opts = "".join([f"<option value='{h(w['code'])}'>{h(w['name'])}</option>" for w in workers])
    s_opts = "".join([f"<option value='{h(s)}'>{h(s)}</option>" for s in salas])

    msg = (request.query_params.get("msg") or "").strip()
    msg_html = f"<div class='card' style='border-color:#ddd;background:#fafafa'><b>{h(msg)}</b></div>" if msg else ""

    body = f"""
    <div class="top">
      <div><h2>Añadir Entrada/Salida</h2><p class="muted">Registra entrada/salida actual o manual.</p></div>
      <div><a class="btn2" href="/encargado/horas">Volver</a></div>
    </div>

    {msg_html}

    <div class="card">
      <form method="post" action="/encargado/horas/add">
        <label>Trabajador</label>
        <select name="worker_code" required>{w_opts}</select>

        <label>Sala</label>
        <select name="room_name" required>{s_opts}</select>

        <div class="row" style="margin-top:12px">
          <button class="btn" name="action" value="entrada_now" type="submit">Entrada (ahora)</button>
          <button class="btn" name="action" value="salida_now" type="submit">Salida (ahora)</button>
        </div>

        <div class="hr"></div>

        <h3 style="margin-top:10px">Registrar Entrada/Salida MANUAL</h3>
        <p class="muted" style="margin-top:-6px">Puedes poner solo Entrada (abre registro), solo Salida (cierra el registro abierto), o Entrada+Salida (registro cerrado).</p>

        <label>Entrada (manual)</label>
        <input type="datetime-local" name="entry_manual"/>

        <label>Salida (manual)</label>
        <input type="datetime-local" name="exit_manual"/>

        <div style="margin-top:12px">
          <button class="btn2" name="action" value="manual" type="submit">Registrar manual</button>
        </div>
      </form>
    </div>
    """
    return page("Añadir Entrada/Salida", body)


@app.post("/encargado/horas/add")
def horas_add_submit(
    request: Request,
    worker_code: str = Form(...),
    room_name: str = Form(...),
    action: str = Form(...),
    entry_manual: str = Form(""),
    exit_manual: str = Form(""),
):
    r = require_login(request)
    if r:
        return r
    u = user_from_session(request)
    if u.get("rol") != "ENCARGADO":
        return RedirectResponse(role_home_path(u.get("rol", "")), status_code=303)

    wcode = (worker_code or "").strip().upper()
    sala = (room_name or "").strip()

    w = db_one("select code, name from public.wom_users where upper(code)=upper(%s) limit 1;", (wcode,))
    if not w:
        return RedirectResponse("/encargado/horas/add?msg=" + urllib.parse.quote("Trabajador no válido"), status_code=303)

    now = now_madrid()

    open_row = db_one(
        "select id, entry_at from public.wom_hours where worker_code=%s and room_name=%s and exit_at is null order by entry_at desc nulls last limit 1;",
        (wcode, sala),
    )

    def go(msg: str):
        return RedirectResponse("/encargado/horas/add?msg=" + urllib.parse.quote(msg), status_code=303)

    if action == "entrada_now":
        if open_row:
            return go("Debe registrar la salida del trabajador primero.")
        db_exec(
            """
            insert into public.wom_hours (worker_code, worker_name, room_name, entry_at, exit_at, recorded_by_code, recorded_by_name)
            values (%s, %s, %s, %s, null, %s, %s);
            """,
            (wcode, w["name"], sala, now, u["codigo"], u["nombre"]),
        )
        return go("Entrada registrada correctamente.")

    if action == "salida_now":
        if not open_row:
            return go("Debe registrar la entrada del trabajador primero.")
        db_exec(
            "update public.wom_hours set exit_at=%s, recorded_by_code=%s, recorded_by_name=%s where id=%s;",
            (now, u["codigo"], u["nombre"], open_row["id"]),
        )
        return go("Salida registrada correctamente.")

    if action == "manual":
        en = _parse_dt_local(entry_manual)
        ex = _parse_dt_local(exit_manual)

        if en and ex and ex < en:
            return go("La salida no puede ser anterior a la entrada.")

        if en and ex:
            db_exec(
                """
                insert into public.wom_hours (worker_code, worker_name, room_name, entry_at, exit_at, recorded_by_code, recorded_by_name)
                values (%s, %s, %s, %s, %s, %s, %s);
                """,
                (wcode, w["name"], sala, en, ex, u["codigo"], u["nombre"]),
            )
            return go("Registro manual (entrada y salida) guardado.")

        if en and not ex:
            if open_row:
                return go("Debe registrar la salida del trabajador primero.")
            db_exec(
                """
                insert into public.wom_hours (worker_code, worker_name, room_name, entry_at, exit_at, recorded_by_code, recorded_by_name)
                values (%s, %s, %s, %s, null, %s, %s);
                """,
                (wcode, w["name"], sala, en, u["codigo"], u["nombre"]),
            )
            return go("Entrada manual registrada correctamente.")

        if ex and not en:
            if not open_row:
                return go("Debe registrar la entrada del trabajador primero.")
            entry_at = open_row.get("entry_at")
            if entry_at:
                entry_at = entry_at.astimezone(TZ) if entry_at.tzinfo else entry_at.replace(tzinfo=TZ)
                if ex < entry_at:
                    return go("La salida manual no puede ser anterior a la entrada registrada.")
            db_exec(
                "update public.wom_hours set exit_at=%s, recorded_by_code=%s, recorded_by_name=%s where id=%s;",
                (ex, u["codigo"], u["nombre"], open_row["id"]),
            )
            return go("Salida manual registrada correctamente.")

        return go("No se indicó entrada ni salida en el registro manual.")

    return go("Acción no reconocida.")


@app.get("/encargado/horas/consultar", response_class=HTMLResponse)
def horas_consultar_form(request: Request):
    r = require_login(request)
    if r:
        return r
    u = user_from_session(request)
    if u.get("rol") != "ENCARGADO":
        return RedirectResponse(role_home_path(u.get("rol", "")), status_code=303)

    workers = _workers_for_hours()
    now = now_madrid()
    mes = (request.query_params.get("mes") or str(now.month)).strip()
    anio = (request.query_params.get("anio") or str(now.year)).strip()
    worker_code = (request.query_params.get("worker_code") or (workers[0]["code"] if workers else "")).strip().upper()

    w_opts = "".join([f"<option value='{h(w['code'])}' {'selected' if w['code']==worker_code else ''}>{h(w['name'])}</option>" for w in workers])
    months_opts = "".join([f"<option value='{i}' {'selected' if str(i)==mes else ''}>{i:02d}</option>" for i in range(1, 13)])
    years = [now.year - 1, now.year, now.year + 1]
    years_opts = "".join([f"<option value='{y}' {'selected' if str(y)==anio else ''}>{y}</option>" for y in years])

    rows = []
    total = 0.0
    error = ""

    try:
        mes_i = int(mes); anio_i = int(anio)
        ts_start, ts_end = month_bounds(anio_i, mes_i)
        rows = db_all(
            """
            select id, room_name, entry_at, exit_at
            from public.wom_hours
            where worker_code=%s and entry_at >= %s and entry_at < %s
            order by entry_at asc nulls last;
            """,
            (worker_code, ts_start, ts_end),
        )
    except Exception as ex:
        error = str(ex)
        rows = []

    trs = ""
    for rr in rows:
        en_f, en_h = formatear_fecha_hora(rr.get("entry_at"))
        ex_f, ex_h = (("-", "-") if not rr.get("exit_at") else formatear_fecha_hora(rr.get("exit_at")))
        hrs_txt = "-"
        if rr.get("entry_at") and rr.get("exit_at"):
            dt_en = rr["entry_at"]; dt_ex = rr["exit_at"]
            dt_en = dt_en.astimezone(TZ) if dt_en.tzinfo else dt_en.replace(tzinfo=TZ)
            dt_ex = dt_ex.astimezone(TZ) if dt_ex.tzinfo else dt_ex.replace(tzinfo=TZ)
            hours = (dt_ex - dt_en).total_seconds() / 3600.0
            hrs = _round_to_half_hours(hours)
            total += hrs
            hrs_txt = f"{hrs:.1f}"
        del_url = f"/encargado/horas/delete/{rr['id']}?worker_code={urllib.parse.quote(worker_code)}&mes={urllib.parse.quote(str(mes))}&anio={urllib.parse.quote(str(anio))}"
        trs += f"""
        <tr>
          <td>{h(rr.get('room_name',''))}</td>
          <td>{h(en_f)} {h(en_h)}</td>
          <td>{h(ex_f)} {h(ex_h)}</td>
          <td>{h(hrs_txt)}</td>
          <td>
            <form method="post" action="{del_url}" onsubmit="return confirm('¿Eliminar este registro?');">
              <button class="btn2 danger" type="submit">Eliminar</button>
            </form>
          </td>
        </tr>
        """

    body = f"""
    <div class="top">
      <div><h2>Consultar Horas</h2></div>
      <div><a class="btn2" href="/encargado/horas">Volver</a></div>
    </div>

    <div class="card">
      <form method="get" action="/encargado/horas/consultar">
        <div class="row">
          <div style="flex:1">
            <label>Trabajador</label>
            <select name="worker_code">{w_opts}</select>
          </div>
          <div style="flex:1">
            <label>Mes</label>
            <select name="mes">{months_opts}</select>
          </div>
          <div style="flex:1">
            <label>Año</label>
            <select name="anio">{years_opts}</select>
          </div>
        </div>
        <div style="margin-top:12px">
          <button class="btn" type="submit">Filtrar</button>
        </div>
      </form>
    </div>

    {f"<div class='card'><b style='color:#c00'>{h(error)}</b></div>" if error else ""}

    <div class="card">
      <table>
        <thead><tr><th>Sala</th><th>Entrada</th><th>Salida</th><th>NºHoras</th><th></th></tr></thead>
        <tbody>
          {trs or "<tr><td colspan='5'>No hay registros para el filtro.</td></tr>"}
        </tbody>
      </table>
      <div class="hr"></div>
      <p><b>TOTAL:</b> {h(f"{total:.1f}")} horas</p>
    </div>
    """
    return page("Consultar Horas", body)


@app.post("/encargado/horas/delete/{hid}")
def horas_delete(request: Request, hid: int):
    r = require_login(request)
    if r:
        return r
    u = user_from_session(request)
    if u.get("rol") != "ENCARGADO":
        return RedirectResponse(role_home_path(u.get("rol", "")), status_code=303)

    db_exec("delete from public.wom_hours where id=%s;", (hid,))
    qs = str(request.url.query or "")
    back = "/encargado/horas/consultar"
    if qs:
        back += "?" + qs
    return RedirectResponse(back, status_code=303)


@app.get("/encargado/horas/pdf", response_class=HTMLResponse)
def horas_pdf_form(request: Request):
    r = require_login(request)
    if r:
        return r
    u = user_from_session(request)
    if u.get("rol") != "ENCARGADO":
        return RedirectResponse(role_home_path(u.get("rol", "")), status_code=303)

    workers = _workers_for_hours()
    now = now_madrid()
    w_opts = "".join([f"<option value='{h(w['code'])}'>{h(w['name'])}</option>" for w in workers])
    months_opts = "".join([f"<option value='{i}' {'selected' if i==now.month else ''}>{i:02d}</option>" for i in range(1, 13)])
    years = [now.year - 1, now.year, now.year + 1]
    years_opts = "".join([f"<option value='{y}' {'selected' if y==now.year else ''}>{y}</option>" for y in years])

    body = f"""
    <div class="top">
      <div><h2>Generar PDF de Horas</h2></div>
      <div><a class="btn2" href="/encargado/horas">Volver</a></div>
    </div>

    <div class="card">
      <form method="post" action="/encargado/horas/pdf">
        <label>Trabajador</label>
        <select name="worker_code" required>{w_opts}</select>

        <div class="row">
          <div style="flex:1">
            <label>Mes</label>
            <select name="mes">{months_opts}</select>
          </div>
          <div style="flex:1">
            <label>Año</label>
            <select name="anio">{years_opts}</select>
          </div>
        </div>

        <div style="margin-top:12px">
          <button class="btn" type="submit">Generar PDF</button>
        </div>
      </form>
    </div>
    """
    return page("PDF Horas", body)


def _query_horas(worker_code: str, year: int, month: int) -> List[Dict[str, Any]]:
    ts_start, ts_end = month_bounds(year, month)
    return db_all(
        """
        select id, room_name, entry_at, exit_at
        from public.wom_hours
        where worker_code=%s and entry_at >= %s and entry_at < %s
        order by entry_at asc nulls last;
        """,
        (worker_code, ts_start, ts_end),
    )


@app.post("/encargado/horas/pdf")
def horas_pdf_generate(
    request: Request,
    worker_code: str = Form(...),
    mes: str = Form(...),
    anio: str = Form(...),
):
    r = require_login(request)
    if r:
        return r
    u = user_from_session(request)
    if u.get("rol") != "ENCARGADO":
        return RedirectResponse(role_home_path(u.get("rol", "")), status_code=303)

    try:
        m_i = int(mes); y_i = int(anio)
    except Exception:
        return HTMLResponse(page("Error", "<div class='card'><h3>Mes/Año inválido</h3></div>"), status_code=400)

    wcode = (worker_code or "").strip().upper()
    w = db_one("select code, name from public.wom_users where upper(code)=upper(%s) limit 1;", (wcode,))
    if not w:
        return HTMLResponse(page("Error", "<div class='card'><h3>Trabajador no válido</h3></div>"), status_code=400)

    rows = _query_horas(wcode, y_i, m_i)

    from reportlab.lib.pagesizes import A4
    from reportlab.lib.units import mm
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib import colors

    out_dir = Path.cwd()
    ts = now_madrid().strftime("%Y%m%d_%H%M%S")
    out_path = out_dir / f"horas_{wcode}_{y_i}_{m_i:02d}_{ts}.pdf"

    doc = SimpleDocTemplate(
        str(out_path),
        pagesize=A4,
        leftMargin=12 * mm,
        rightMargin=12 * mm,
        topMargin=12 * mm,
        bottomMargin=12 * mm,
        title="Horas de trabajo de mantenimiento",
    )

    styles = getSampleStyleSheet()
    st_title = ParagraphStyle("t", parent=styles["Normal"], fontName="Helvetica-Bold", fontSize=16, leading=18)
    st_mid = ParagraphStyle("m", parent=styles["Normal"], fontName="Helvetica", fontSize=11, leading=13)

    story = []
    story.append(Paragraph("HORAS DE TRABAJO DE MANTENIMIENTO", st_title))
    story.append(Spacer(1, 10))
    story.append(Paragraph(f"Trabajador: <b>{_xml_escape(w['name'])}</b>", st_mid))
    story.append(Paragraph(f"Mes y año: <b>{m_i:02d}/{y_i}</b>", st_mid))
    story.append(Spacer(1, 6))
    story.append(Paragraph("<para><font color='#000000'>______________________________________________</font></para>", st_mid))
    story.append(Spacer(1, 10))

    data = [["Sala", "Entrada", "Salida", "NºHoras"]]
    total = 0.0
    for rr in rows:
        en_f, en_h = formatear_fecha_hora(rr.get("entry_at"))
        ex_f, ex_h = (("-", "-") if not rr.get("exit_at") else formatear_fecha_hora(rr.get("exit_at")))
        hrs_txt = "-"
        if rr.get("entry_at") and rr.get("exit_at"):
            dt_en = rr["entry_at"]; dt_ex = rr["exit_at"]
            dt_en = dt_en.astimezone(TZ) if dt_en.tzinfo else dt_en.replace(tzinfo=TZ)
            dt_ex = dt_ex.astimezone(TZ) if dt_ex.tzinfo else dt_ex.replace(tzinfo=TZ)
            hours = (dt_ex - dt_en).total_seconds() / 3600.0
            hrs = _round_to_half_hours(hours)
            total += hrs
            hrs_txt = f"{hrs:.1f}"
        data.append([rr.get("room_name", ""), f"{en_f} {en_h}", f"{ex_f} {ex_h}", hrs_txt])

    data.append(["", "", "TOTAL", f"{total:.1f}"])

    table = Table(data, colWidths=[55 * mm, 45 * mm, 45 * mm, 20 * mm])
    table.setStyle(
        TableStyle(
            [
                ("FONTNAME", (0, 0), (-1, -1), "Courier"),
                ("FONTSIZE", (0, 0), (-1, -1), 8),
                ("LINEBELOW", (0, 0), (-1, 0), 1, colors.black),
                ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
                ("BACKGROUND", (0, 0), (-1, 0), colors.whitesmoke),
                ("FONTNAME", (0, 0), (-1, 0), "Courier-Bold"),
                ("FONTNAME", (2, -1), (-1, -1), "Courier-Bold"),
            ]
        )
    )
    story.append(table)

    doc.build(story)
    return FileResponse(str(out_path), media_type="application/pdf", filename=out_path.name)
