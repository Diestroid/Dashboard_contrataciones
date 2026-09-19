"""Lanzador único para el Dashboard (uso normal y .exe de PyInstaller).

- Localiza app.py tanto en modo script como congelado (_MEIPASS).
- Elige un puerto libre, arranca Streamlit y abre el navegador.
- La carpeta de Excels se puede pasar como argumento o con DASHBOARD_DATA_DIR:
      DashboardAlfresco.exe "C:\\Contratos\\Excels"
      set DASHBOARD_DATA_DIR=C:\\Contratos\\Excels && DashboardAlfresco.exe
"""
import os
import socket
import sys
import threading
import webbrowser


def recurso(nombre):
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, nombre)


def puerto_libre():
    s = socket.socket()
    s.bind(("127.0.0.1", 0))
    _, p = s.getsockname()
    s.close()
    return p


def main():
    # 1) Carpeta de datos desde argumento o env (app.py la lee de env).
    if len(sys.argv) > 1 and os.path.isdir(sys.argv[1]):
        os.environ["DASHBOARD_DATA_DIR"] = os.path.abspath(sys.argv[1])

    app = recurso("app.py")
    if not os.path.isfile(app):
        # Fallback: app.py junto al .exe (modo --add-data con otra estructura)
        alt = os.path.join(os.path.dirname(os.path.abspath(sys.executable)), "app.py")
        if os.path.isfile(alt):
            app = alt
        else:
            print(f"No se encontró app.py. Buscado en: {app} y {alt}")
            input("Pulsa Enter para salir…")
            sys.exit(1)

    puerto = int(os.environ.get("DASHBOARD_PORT", puerto_libre()))
    url = f"http://localhost:{puerto}"

    threading.Timer(2.5, lambda: webbrowser.open(url)).start()
    print(f"Abriendo dashboard en {url}")
    print(f"Carpeta de datos: {os.environ.get('DASHBOARD_DATA_DIR', '(la elegirás en la app)')}")
    print("No cierres esta ventana mientras usas el dashboard.")

    from streamlit.web import cli as stcli

    sys.argv = [
        "streamlit", "run", app,
        "--server.port", str(puerto),
        "--server.headless", "true",
        "--server.enableCORS", "false",
        "--browser.gatherUsageStats", "false",
        "--global.developmentMode", "false",
    ]
    sys.exit(stcli.main())


if __name__ == "__main__":
    main()
