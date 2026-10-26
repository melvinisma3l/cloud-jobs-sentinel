# ==============================================================================
# PROYECTO CLOUD 24/7: SENTINELA DE EMPLEOS REMOTOS EN LA NUBE
# TARGET: RemoteOK Global Tech Jobs API (JSON Directo)
# OBJETIVO: Correr sin navegador en servidores de GitHub y alertar nuevos puestos
# ==============================================================================

import sys
import json
from datetime import datetime
import requests

if sys.platform == "win32":
    sys.stdout.reconfigure(encoding="utf-8")

def ejecutar_centinela_cloud():
    ahora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print("=" * 70)
    print(f"🤖 CENTINELA CLOUD ACTIVO EN LA NUBE | TIMESTAMP: {ahora}")
    print("=" * 70)

    url = "https://remoteok.com/api"
    headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) CloudWatcher/1.0"}

    try:
        resp = requests.get(url, headers=headers, timeout=15)
        if resp.status_code == 200:
            datos = resp.json()
            # El primer elemento de RemoteOK suele ser un aviso legal, los demás son empleos:
            empleos = [item for item in datos if isinstance(item, dict) and "position" in item][:10]
            
            print(f"✅ Conexión 200 OK. Monitoreando los últimos {len(empleos)} empleos remotos:")
            print("-" * 70)
            
            snapshot = []
            for j in empleos:
                puesto = j.get("position", "N/A")
                empresa = j.get("company", "N/A")
                tags = ", ".join(j.get("tags", [])[:3])
                url_postulacion = j.get("url", "")

                print(f"💼 {puesto[:30]:<32} | 🏢 {empresa[:18]:<20} | 🏷️ {tags}")
                
                snapshot.append({
                    "timestamp": ahora,
                    "puesto": puesto,
                    "empresa": empresa,
                    "tags": tags,
                    "url": url_postulacion
                })

            ruta_json = "C:\\Users\\yidkw\\.gemini\\antigravity\\scratch\\Cloud_Automation_Lab\\ultimos_empleos.json"
            with open(ruta_json, "w", encoding="utf-8") as f:
                json.dump(snapshot, f, indent=4, ensure_ascii=False)

            print("-" * 70)
            print("💾 Snapshot de datos actualizado en disco exitosamente.")
            print("=" * 70)
        else:
            print(f"⚠️ Servidor respondió con código: {resp.status_code}")
    except Exception as e:
        print(f"❌ Error durante el rastreo: {e}")

if __name__ == "__main__":
    ejecutar_centinela_cloud()
