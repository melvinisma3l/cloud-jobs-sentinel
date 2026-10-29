# ==============================================================================
# PROYECTO PRÁCTICO DE INICIATIVA PROPIA: CAZADOR CLOUD DE OFERTAS IPHONE 17 / AIR
# TARGET: Gazelle USA Certified Store (https://buy.gazelle.com/collections/iphones/products.json)
# LÓGICA DE NEGOCIO:
#   1. Extrae el catálogo vivo de iPhones con Blindaje Nocturno.
#   2. Filtra solo la última generación (iPhone 17 / iPhone Air).
#   3. Busca entre todas las variantes (color/condición) SOLO las que están EN STOCK
#      y atrapa la variante MÁS BARATA disponible de cada modelo.
#   4. Ordena el ranking de la oferta más barata a la más cara.
#   5. Genera Excel Corporativo + Dispara el Top 5 de Gangas directo a Discord.
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


def descargar_catalogo_iphones(max_reintentos=3):
    url = "https://buy.gazelle.com/collections/iphones/products.json?limit=100"
    headers = {"User-Agent": "Mozilla/5.0"}

    for intento in range(1, max_reintentos + 1):
        try:
            print(f"🛡️ [Intento #{intento}] Escaneando tienda Gazelle USA...")
            resp = requests.get(url, headers=headers, timeout=12)
            if resp.status_code == 200:
                return resp.json().get("products", [])
            print(f"⚠️ Servidor respondió código {resp.status_code}")
        except (requests.exceptions.Timeout, requests.exceptions.ConnectionError) as e:
            print(f"⚠️ Error de red temporal: {type(e).__name__}")

        if intento < max_reintentos:
            time.sleep(intento * 2)
    return []


def cazar_mejores_ofertas(productos):
    ofertas = []
    vistos = set()

    for prod in productos:
        titulo = prod.get("title", "").strip()
        titulo_min = titulo.lower()

        # Filtro 1: Solo la última generación (iPhone 17 y nuevo iPhone Air)
        if "iphone 17" not in titulo_min and "iphone air" not in titulo_min:
            continue

        if titulo_min in vistos:
            continue
        vistos.add(titulo_min)

        # Filtro 2: Buscar únicamente variantes que tengan stock real (available == True)
        variantes_en_stock = [
            v for v in prod.get("variants", []) if v.get("available") is True
        ]
        if not variantes_en_stock:
            continue  # Si todas las tallas/colores están agotadas, ignoramos este modelo

        # Filtro 3: Quedarnos con la variante MÁS BARATA en stock de este modelo
        mejor_variante = min(variantes_en_stock, key=lambda v: float(v.get("price", 99999)))
        mejor_precio = float(mejor_variante.get("price", 0.0))
        detalle_variante = mejor_variante.get("title", "Standard")
        url_compra = f"https://buy.gazelle.com/products/{prod.get('handle', '')}"

        ofertas.append({
            "modelo": titulo,
            "variante": detalle_variante,
            "precio": mejor_precio,
            "stock_opciones": len(variantes_en_stock),
            "url": url_compra
        })

    # Ordenamos de MENOR a MAYOR precio para que las mejores gangas queden arriba:
    ofertas.sort(key=lambda x: x["precio"])
    return ofertas


def exportar_excel_ofertas(ofertas, timestamp):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "iPhone 17 Best Deals"

    NAVY = PatternFill(start_color="0F172A", end_color="0F172A", fill_type="solid")
    ZEBRA = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
    WHITE = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    GREEN_BG = PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid")
    THIN = Side(border_style="thin", color="CBD5E1")
    BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

    ws["A1"] = "APPLE IPHONE 17 & IPHONE AIR — LIVE USA DEAL TRACKER"
    ws["A1"].font = Font(name="Calibri", size=14, bold=True, color="0F172A")
    ws["A2"] = f"Source: Gazelle USA Certified | Verified In-Stock Deals: {len(ofertas)} | Scan: {timestamp} UTC"
    ws["A2"].font = Font(name="Calibri", size=10, color="64748B")

    headers = ["Rank", "iPhone Model", "Cheapest In-Stock Variant", "Best Price (USD)", "Available Variants", "Direct Deal Link"]
    for col_idx, text in enumerate(headers, start=1):
        c = ws.cell(row=4, column=col_idx, value=text)
        c.fill = NAVY
        c.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        c.alignment = Alignment(horizontal="center" if col_idx not in [2, 3] else "left", vertical="center")
        c.border = BORDER
    ws.row_dimensions[4].height = 26

    for rank, item in enumerate(ofertas, start=1):
        fila = rank + 4
        fill = ZEBRA if fila % 2 == 0 else WHITE
        formula = f'=HYPERLINK("{item["url"]}", "Grab Deal ↗")'
        datos_fila = [f"#{rank}", item["modelo"], item["variante"], item["precio"], item["stock_opciones"], formula]

        for c_idx, val in enumerate(datos_fila, start=1):
            c = ws.cell(row=fila, column=c_idx, value=val)
            c.fill = GREEN_BG if c_idx == 4 and rank <= 3 else fill
            c.border = BORDER
            if c_idx == 2:
                c.font = Font(name="Calibri", size=10, bold=True, color="0F172A")
            elif c_idx == 4:
                c.font = Font(name="Calibri", size=10, bold=True, color="166534" if rank <= 3 else "0F172A")
                c.number_format = "$#,##0.00"
            elif c_idx == 6:
                c.font = Font(name="Calibri", size=10, color="2563EB", underline="single")
            else:
                c.font = Font(name="Calibri", size=10, color="334155")
            c.alignment = Alignment(horizontal="right" if c_idx == 4 else ("center" if c_idx in [1, 5, 6] else "left"), vertical="center")
        ws.row_dimensions[fila].height = 21

    anchos = {"A": 9, "B": 38, "C": 28, "D": 18, "E": 19, "F": 18}
    for col, w in anchos.items():
        ws.column_dimensions[col].width = w

    nombre_excel = "iPhone_17_Best_Deals_USA.xlsx"
    wb.save(nombre_excel)
    return nombre_excel


def enviar_gangas_discord(webhook_url, ofertas, timestamp):
    if not webhook_url or not ofertas:
        print("ℹ️ Sin DISCORD_WEBHOOK activo en este entorno.")
        return

    medallas = ["🥇", "🥈", "🥉", "🔥", "⚡"]
    lineas = []
    for idx, o in enumerate(ofertas[:5]):
        icono = medallas[idx]
        lineas.append(
            f"{icono} **${o['precio']:,.2f} USD** — **{o['modelo']}**\n"
            f"   └ Variante: `{o['variante']}` ({o['stock_opciones']} en stock) • [Comprar Oferta ↗]({o['url']})"
        )

    payload = {
        "username": "Cazador iPhone 17 USA 🍎",
        "embeds": [
            {
                "title": "🍎 TOP 5 MEJORES OFERTAS EN STOCK — IPHONE 17 & AIR",
                "description": "\n\n".join(lineas),
                "color": 15844367,  # Dorado Apple #F1C40F
                "footer": {"text": f"Escaneo en vivo Gazelle USA | {timestamp} UTC"}
            }
        ]
    }
    r = requests.post(webhook_url, json=payload, timeout=10)
    if r.status_code in [200, 204]:
        print("📲 ¡Top 5 gangas de iPhone 17 enviadas a Discord!")


def main():
    ahora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print("=" * 74)
    print(f"🍎 CAZADOR DE OFERTAS IPHONE 17 / AIR EN VIVO | {ahora}")
    print("=" * 74)

    productos = descargar_catalogo_iphones()
    if not productos:
        print("🚨 No se pudo descargar el catálogo hoy.")
        return

    ofertas = cazar_mejores_ofertas(productos)
    print(f"🎯 ¡Se encontraron {len(ofertas)} modelos de última generación en stock!")
    print("-" * 74)
    for i, o in enumerate(ofertas[:6], start=1):
        print(f"#{i} | ${o['precio']:>7.2f} | {o['modelo']:<36} | {o['variante']}")
    print("-" * 74)

    archivo = exportar_excel_ofertas(ofertas, ahora)
    print(f"📊 Excel guardado: {archivo}")

    # Dashboard visual en GitHub Actions si corre en la nube:
    summary = os.environ.get("GITHUB_STEP_SUMMARY")
    if summary:
        filas_md = [
            f"## 🍎 Ranking en Vivo: Mejores Ofertas iPhone 17 & Air (`{ahora} UTC`)",
            "| Rank | Precio (USD) | Modelo | Mejor Variante en Stock | Link |",
            "| :---: | :---: | :--- | :--- | :---: |"
        ]
        for idx, o in enumerate(ofertas, start=1):
            filas_md.append(f"| **#{idx}** | **${o['precio']:,.2f}** | {o['modelo']} | `{o['variante']}` | [Ver Oferta ↗]({o['url']}) |")
        with open(summary, "a", encoding="utf-8") as f:
            f.write("\n".join(filas_md) + "\n")

    webhook = os.environ.get("DISCORD_WEBHOOK", "").strip()
    enviar_gangas_discord(webhook, ofertas, ahora)


if __name__ == "__main__":
    main()
