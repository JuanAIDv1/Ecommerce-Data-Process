import subprocess
import datetime
PYTHON = r"C:\Users\juan_.DESKTOP-S8MDVBQ\AppData\Local\Programs\Python\Python313\python.exe"
RUTAS = [
    r"C:\Ingestador\scripts\dropipro\save_cookies.py",
    r"C:\Ingestador\data_raw\dropipro\scrape_incremental_ordenes_actuales.py",
    r"C:\Ingestador\data_raw\dropipro\scrape_incremental_ordenes_nuevas.py",
    r"C:\Ingestador\data_raw\dropea\scrape_dropea.py"
]
LOG_FILE = r"C:\Ingestador\errores_ejecucion.txt"
def log_error(script, error):
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write("\n" + "="*60 + "\n")
        f.write(f"📅 {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
        f.write(f"❌ Error en: {script}\n")
        f.write(f"Detalle:\n{error}\n")
def ejecutar_scripts():
    for script in RUTAS:
        print(f"\n🚀 Ejecutando {script}...\n")
        try:
            subprocess.run([PYTHON, script], check=True)
            print(f"✅ Finalizado: {script}")
        except subprocess.CalledProcessError as e:
            print(f"⚠️ Error en {script}, continuando con el siguiente...")
            log_error(script, e)
if __name__ == "__main__":
    ejecutar_scripts()
    print("\n🎯 Ejecución completa")
