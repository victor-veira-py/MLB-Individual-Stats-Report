import requests
import pandas as pd


def reporte_estilo_ejecutivo_pro(player_id):
    """
    Genera un reporte estadístico individual con formato empresarial.
    Extrae datos de la StatsAPI de MLB, calcula totales de carrera y 
    exporta a Excel con estilos de alta visibilidad (Look & Feel de MLB).
    """
    try:
        # 1. Recuperación de metadatos desde la StatsAPI de MLB
        info = requests.get(f"https://statsapi.mlb.com/api/v1/people/{player_id}").json()
        nombre_completo = info["people"][0].get("fullName", "Jugador")

        # Consulta de estadísticas agregadas por temporada (Year-by-Year)
        url = f"https://statsapi.mlb.com/api/v1/people/{player_id}/stats?stats=yearByYear&group=hitting&sportIds=1"
        data = requests.get(url).json()

        if "stats" not in data or not data["stats"][0]["splits"]:
            print(f"No se encontraron registros para el ID: {player_id}")
            return

        splits = data["stats"][0]["splits"]
        lista_temporadas = []

        def fmt_mlb(valor):
            """Aplica el formato estándar de MLB para porcentajes (.XXX)"""
            try:
                return f"{float(valor):.3f}".replace("0.", ".")
            except:
                return ".000"

        # 2. Procesamiento de la data anual
        for s in splits:
            stat = s.get("stat", {})
            ab = int(stat.get("atBats", 0))
            if ab == 0: continue

            # Gestión de nombres de equipo (Manejo de filas 'TOTAL' de la API)
            nombre_equipo = s.get("team", {}).get("name", "---")
            if nombre_equipo == "---":
                nombre_equipo = "TOTAL"

            lista_temporadas.append({
                "TEMPORADA": int(s.get("season")),
                "EQUIPO": nombre_equipo,
                "J": int(stat.get("gamesPlayed", 0)), "AB": ab, "R": int(stat.get("runs", 0)),
                "H": int(stat.get("hits", 0)), "2B": int(stat.get("doubles", 0)),
                "3B": int(stat.get("triples", 0)), "HR": int(stat.get("homeRuns", 0)),
                "RBI": int(stat.get("rbi", 0)), "BB": int(stat.get("baseOnBalls", 0)),
                "HBP": int(stat.get("hitByPitch", 0)), "K": int(stat.get("strikeOuts", 0)),
                "SB": int(stat.get("stolenBases", 0)), "CS": int(stat.get("caughtStealing", 0)),
                "AVG": float(stat.get("avg", 0)), "OBP": float(stat.get("obp", 0)),
                "SLG": float(stat.get("slg", 0)), "OPS": float(stat.get("ops", 0))
            })

        df = pd.DataFrame(lista_temporadas)

        # 3. Cálculo de Estadísticas de Carrera (Career Totals)
        # Filtramos filas 'TOTAL' para que la sumatoria sea exacta
        df_solo_equipos = df[df["EQUIPO"] != "TOTAL"]

        fila_totales = {
            "TEMPORADA": "",
            "EQUIPO": "CARRERA",
            "J": df_solo_equipos["J"].sum(), "AB": df_solo_equipos["AB"].sum(),
            "R": df_solo_equipos["R"].sum(), "H": df_solo_equipos["H"].sum(),
            "2B": df_solo_equipos["2B"].sum(), "3B": df_solo_equipos["3B"].sum(),
            "HR": df_solo_equipos["HR"].sum(), "RBI": df_solo_equipos["RBI"].sum(),
            "BB": df_solo_equipos["BB"].sum(), "HBP": df_solo_equipos["HBP"].sum(),
            "K": df_solo_equipos["K"].sum(), "SB": df_solo_equipos["SB"].sum(),
            "CS": df_solo_equipos["CS"].sum(),
            "AVG": df_solo_equipos["H"].sum() / df_solo_equipos["AB"].sum() if df_solo_equipos["AB"].sum() > 0 else 0,
            "OBP": df_solo_equipos["OBP"].mean(), "SLG": df_solo_equipos["SLG"].mean(),
            "OPS": df_solo_equipos["OPS"].mean()
        }

        # Consolidación final y aplicación de formato visual .XXX
        df = pd.concat([df, pd.DataFrame([fila_totales])], ignore_index=True)
        for col in ["AVG", "OBP", "SLG", "OPS"]:
            df[col] = df[col].apply(fmt_mlb)

        # 4. Generación de Reporte Excel con XlsxWriter (Estilo Profesional)
        nombre_archivo = f"Reporte_Individual_{nombre_completo.replace(' ', '_')}.xlsx"

        with pd.ExcelWriter(nombre_archivo, engine='xlsxwriter') as writer:
            df.to_excel(writer, sheet_name='Estadisticas', index=False)
            workbook = writer.book
            worksheet = writer.sheets['Estadisticas']

            # Definición de formatos
            f_header = workbook.add_format(
                {'bold': True, 'bg_color': '#1F4E78', 'font_color': 'white', 'align': 'center', 'border': 1})
            f_data = workbook.add_format({'align': 'center', 'border': 1, 'border_color': '#D9D9D9'})
            f_total = workbook.add_format({'bold': True, 'bg_color': '#F2F2F2', 'align': 'center', 'border': 1})

            worksheet.hide_gridlines(2)
            worksheet.ignore_errors({'number_stored_as_text': f'A2:T{len(df) + 1}'})

            # Ajuste dinámico de columnas (Usando la lógica que nos funcionó hoy)
            for i, col in enumerate(df.columns):
                ancho = max(df[col].astype(str).map(len).max(), len(col)) + 3
                worksheet.set_column(i, i, ancho, f_data)
                worksheet.write(0, i, col, f_header)

            # Aplicación de estilo a la fila de Totales (CARRERA)
            idx_ultima = len(df)
            for i in range(len(df.columns)):
                worksheet.write(idx_ultima, i, df.iloc[-1, i], f_total)

        print(f"\n[SISTEMA] Reporte generado exitosamente: {nombre_archivo}")

    except Exception as e:
        print(f"❌ Error crítico: {e}")


if __name__ == "__main__":
    # Ejemplo con José Altuve (514888)
    reporte_estilo_ejecutivo_pro("514888")