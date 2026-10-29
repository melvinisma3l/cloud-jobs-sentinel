# ==============================================================================
# CLOUD AUTOMATION DÍA 18 (CIERRE DE FASE 3): CENTINELA MAESTRO EN PILOTO AUTOMÁTICO
# UNE LAS 4 ARMAS CLOUD:
# 1. Blindaje Nocturno (Exponential Backoff contra caídas y bloqueos 429)
# 2. Filtrado de Negocio y Deduplicación en Memoria (set)
# 3. Generación de Excel Corporativo Navy Blue (Artifact descargable)
# 4. Alerta Instantánea en Vivo (Discord Webhook + Dashboard GitHub Summary)
# ==============================================================================

import os
import sys
import time
from datetime import datetime
import requests
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

if sys.platform == "win32":
    sys.stdout.reconfigure(encoding="utf-8")


def descargar_con_blindaje(url, max_reintentos=3):
    """Arma 1: Blindaje Nocturno con espera exponencial."""
    headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) CloudMaster/3.0"}
    for intento in range(1, max_reintentos + 1):
        try:
            print(f"🛡️ [Intento #{intento}] Conectando a {url}...")
            resp = requests.get(url, headers=headers, timeout=12)
            if resp.status_code == 200:
                print("✅ Conexión 200 OK.")
                return resp.json()
            print(f"⚠️ Código recibido: {resp.status_code}")
        except (requests.exceptions.Timeout, requests.exceptions.ConnectionError) as e:
            print(f"⚠️ Fallo temporal de red: {type(e).__name__}")

        if intento < max_reintentos:
            espera = intento * 2
            print(f"⏳ Esperando {espera}s antes de reintentar...")
            time.sleep(espera)
    return None


def enviar_alerta_discord(webhook_url, empleos_top, timestamp):
    """Arma 4A: Dispara notificación en tiempo real al celular/PC vía Discord."""
    if not webhook_url:
        print("ℹ️ Sin DISCORD_WEBHOOK configurado (modo silencioso).")
        return

    lineas = []
    for emp in empleos_top[:5]:
        lineas.append(f"• **{emp['puesto']}** en `{emp['empresa']}` — [Aplicar ↗]({emp['url']})")

    mensaje = {
        "username": "Centinela Cloud 24/7",
        "embeds": [
            {
                "title": "🚀 Reporte Diario Cloud — Nuevas Vacantes Detectadas",
                "description": "\n".join(lineas),
                "color": 1981066,  # Azul corporativo #1E3A8A
                "footer": {"text": f"Ejecutado en GitHub Actions | {timestamp} UTC"}
            }
        ]
    }
    try:
        r = requests.post(webhook_url, json=mensaje, timeout=10)
        if r.status_code in [200, 204]:
            print("📲 ¡Alerta enviada en vivo a Discord exitosamente!")
        else:
            print(f"⚠️ Discord respondió código: {r.status_code}")
    except Exception as e:
        print(f"⚠️ No se pudo enviar alerta Discord: {e}")


def publicar_dashboard_github(empleos_limpios, timestamp):
    """Arma 4B: Dibuja un panel visual directo en la pantalla de GitHub Actions."""
    ruta_summary = os.environ.get("GITHUB_STEP_SUMMARY")
    if not ruta_summary:
        return

    md = [
        f"## 🤖 Dashboard Cloud 24/7 — Ejecución `{timestamp} UTC`",
        f"- **Total vacantes únicas procesadas:** `{len(empleos_limpios)}`",
        "- **Estado del Excel Corporativo:** `Listo en Artifacts abajo ⬇️`",
        "",
        "| Puesto | Empresa | Stack / Tags | Enlace |",
        "| :--- | :--- | :--- | :--- |"
    ]
    for e in empleos_limpios[:10]:
        md.append(f"| **{e['puesto']}** | {e['empresa']} | `{e['tags']}` | [Ver Vacante ↗]({e['url']}) |")

    with open(ruta_summary, "a", encoding="utf-8") as f:
        f.write("\n".join(md) + "\n")


def ejecutar_pipeline_maestro():
    ahora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print("=" * 72)
    print(f"🚀 CENTINELA MAESTRO V3 EN PILOTO AUTOMÁTICO | {ahora}")
    print("=" * 72)

    # 1. Descarga blindada
    datos = descargar_con_blindaje("https://remoteok.com/api")
    if not datos:
        print("🚨 Fin limpio: Servidor externo no disponible hoy.")
        return

    # 2. Saneamiento + Deduplicación con set()
    vistos = set()
    empleos_limpios = []
    for item in datos:
        if not isinstance(item, dict) or "position" not in item:
            continue
        puesto = item.get("position", "").strip()
        empresa = item.get("company", "N/A").strip()
        clave_unica = f"{puesto}|{empresa}".lower()

        if clave_unica in vistos:
            continue
        vistos.add(clave_unica)

        empleos_limpios.append({
            "puesto": puesto,
            "empresa": empresa,
            "tags": ", ".join(item.get("tags", [])[:4]),
            "ubicacion": item.get("location", "Worldwide") or "Worldwide",
            "url": item.get("url", "")
        })
        if len(empleos_limpios) == 20:
            break

    print(f"🧹 Saneamiento completado: {len(empleos_limpios)} vacantes únicas listas.")

    # 3. Exportación Excel Ejecutivo Navy Blue
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Cloud Jobs Intelligence"

    NAVY = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
    ZEBRA = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
    WHITE = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    THIN = Side(border_style="thin", color="CBD5E1")
    BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

    ws["A1"] = "AUTONOMOUS CLOUD PIPELINE — GLOBAL REMOTE TECH JOBS"
    ws["A1"].font = Font(name="Calibri", size=14, bold=True, color="1E3A8A")
    ws["A2"] = f"Verified Unique Listings: {len(empleos_limpios)} | Generated: {ahora} UTC"
    ws["A2"].font = Font(name="Calibri", size=10, color="64748B")

    headers = ["Job Title", "Company", "Tech Stack", "Location", "Direct Apply"]
    for col_idx, text in enumerate(headers, start=1):
        c = ws.cell(row=4, column=col_idx, value=text)
        c.fill = NAVY
        c.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        c.alignment = Alignment(horizontal="center" if col_idx > 3 else "left", vertical="center")
        c.border = BORDER
    ws.row_dimensions[4].height = 26

    for idx, e in enumerate(empleos_limpios, start=5):
        fill = ZEBRA if idx % 2 == 0 else WHITE
        formula = f'=HYPERLINK("{e["url"]}", "Apply Now ↗")' if e["url"] else "N/A"
        valores = [e["puesto"], e["empresa"], e["tags"], e["ubicacion"], formula]

        for c_idx, val in enumerate(valores, start=1):
            c = ws.cell(row=idx, column=c_idx, value=val)
            c.fill = fill
            c.border = BORDER
            if c_idx == 1:
                c.font = Font(name="Calibri", size=10, bold=True, color="0F172A")
            elif c_idx == 5:
                c.font = Font(name="Calibri", size=10, color="2563EB", underline="single")
            else:
                c.font = Font(name="Calibri", size=10, color="334155")
            c.alignment = Alignment(horizontal="center" if c_idx in [4, 5] else "left", vertical="center")
        ws.row_dimensions[idx].height = 20

    anchos = {"A": 40, "B": 22, "C": 34, "D": 18, "E": 18}
    for letra, ancho in anchos.items():
        ws.column_dimensions[letra].width = ancho

    archivo_salida = "Cloud_Master_Jobs_Report.xlsx"
    wb.save(archivo_salida)
    print(f"📊 Excel Corporativo guardado: {archivo_salida}")

    # 4. Dashboard visual en GitHub + Alerta en vivo a Discord (si hay secreto configurado)
    publicar_dashboard_github(empleos_limpios, ahora)
    webhook = os.environ.get("DISCORD_WEBHOOK", "").strip()
    enviar_alerta_discord(webhook, empleos_limpios, ahora)

    print("=" * 72)
    print("🏆 PIPELINE MAESTRO COMPLETADO AL 100%")
    print("=" * 72)


if __name__ == "__main__":
    ejecutar_pipeline_maestro()
