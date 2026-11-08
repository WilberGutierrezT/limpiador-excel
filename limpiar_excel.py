#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
limpiar_excel.py — Limpiador automático de hojas de cálculo desordenadas.

Toma un archivo Excel o CSV con datos "sucios" (espacios de sobra, filas y
columnas vacías, duplicados, encabezados mal ubicados, fechas y números
guardados como texto) y devuelve una versión limpia, más un reporte de todo
lo que se corrigió.

Uso:
    python limpiar_excel.py entrada.xlsx
    python limpiar_excel.py entrada.xlsx -o salida.xlsx
    python limpiar_excel.py datos.csv --hoja "Ventas" --sin-duplicados

El resultado se guarda como "<nombre>_limpio.xlsx" si no se indica -o.
El reporte de cambios se imprime en pantalla y se guarda junto al archivo.

Autor: Wilber Gutierrez Taboada
Licencia: MIT
"""

import argparse
import sys
from pathlib import Path
from datetime import datetime

import pandas as pd


# --------------------------------------------------------------------------- #
#  Funciones de limpieza. Cada una recibe un DataFrame y devuelve (df, nota).  #
#  La "nota" es una línea de texto para el reporte final.                      #
# --------------------------------------------------------------------------- #

def detectar_fila_encabezado(df_crudo, max_filas=10):
    """
    Cuando el Excel trae títulos o celdas en blanco antes de la tabla real,
    el encabezado verdadero no está en la fila 0. Busca la primera fila que
    parece una cabecera: mayoría de celdas no vacías y de tipo texto.
    Devuelve el índice de esa fila (0 si no encuentra algo mejor).
    """
    mejor_fila = 0
    mejor_puntaje = -1
    limite = min(max_filas, len(df_crudo))
    for i in range(limite):
        fila = df_crudo.iloc[i]
        no_vacias = fila.notna().sum()
        texto = sum(isinstance(v, str) and v.strip() != "" for v in fila)
        # Una buena cabecera tiene muchas celdas llenas y en su mayoría texto.
        puntaje = no_vacias + texto
        if puntaje > mejor_puntaje:
            mejor_puntaje = puntaje
            mejor_fila = i
    return mejor_fila


def limpiar_nombres_columnas(df):
    """Quita espacios, saltos de línea y nombres duplicados o vacíos."""
    nuevos = []
    vistos = {}
    for col in df.columns:
        nombre = str(col).strip().replace("\n", " ")
        while "  " in nombre:
            nombre = nombre.replace("  ", " ")
        if nombre == "" or nombre.lower().startswith("unnamed"):
            nombre = "columna_sin_nombre"
        # Evita columnas con el mismo nombre (col, col_2, col_3...).
        if nombre in vistos:
            vistos[nombre] += 1
            nombre = f"{nombre}_{vistos[nombre]}"
        else:
            vistos[nombre] = 1
        nuevos.append(nombre)
    df.columns = nuevos
    return df, f"Encabezados normalizados: {list(df.columns)}"


def quitar_filas_columnas_vacias(df):
    """Elimina filas y columnas que están completamente vacías."""
    filas_antes, cols_antes = df.shape
    df = df.dropna(axis=0, how="all")
    df = df.dropna(axis=1, how="all")
    filas_q = filas_antes - df.shape[0]
    cols_q = cols_antes - df.shape[1]
    return df, f"Se eliminaron {filas_q} filas vacías y {cols_q} columnas vacías."


def _columnas_texto(df):
    """Columnas de texto, compatible con pandas 2 y 3 (object y string)."""
    return df.select_dtypes(include=["object", "string"]).columns


def recortar_espacios(df):
    """Quita espacios al inicio/fin de todas las celdas de texto."""
    columnas_texto = _columnas_texto(df)
    total = 0
    for col in columnas_texto:
        original = df[col].copy()
        df[col] = df[col].apply(
            lambda v: v.strip() if isinstance(v, str) else v
        )
        total += int((original != df[col]).sum())
    return df, f"Se recortaron espacios sobrantes en {total} celdas de texto."


def normalizar_vacios(df):
    """Convierte textos como '', 'NA', 'N/A', 'null', '-' en celdas vacías reales."""
    marcadores = {"", "na", "n/a", "null", "none", "-", "--", "s/d", "sin dato"}
    total = 0
    columnas_texto = _columnas_texto(df)
    for col in columnas_texto:
        mask = df[col].apply(
            lambda v: isinstance(v, str) and v.strip().lower() in marcadores
        )
        total += int(mask.sum())
        df.loc[mask, col] = pd.NA
    return df, f"Se marcaron como vacías {total} celdas con valores tipo 'N/A', '-', 'null'."


def convertir_numeros(df):
    """
    Intenta convertir a número las columnas de texto que en realidad son cifras
    (p. ej. '1,234.50', 'S/ 1200', '  45 '). Solo convierte si la mayoría de la
    columna se puede interpretar como número, para no romper textos legítimos.
    """
    convertidas = []
    columnas_texto = _columnas_texto(df)
    for col in columnas_texto:
        serie = df[col].dropna().astype(str)
        if serie.empty:
            continue
        limpia = (
            serie.str.replace(r"[^\d,.\-]", "", regex=True)  # quita S/, $, letras, espacios
                 .str.replace(",", "", regex=False)           # separador de miles
        )
        numerica = pd.to_numeric(limpia, errors="coerce")
        proporcion = numerica.notna().mean()
        if proporcion >= 0.8:  # 80% o más son números -> es una columna numérica
            df[col] = pd.to_numeric(
                df[col].astype(str)
                       .str.replace(r"[^\d,.\-]", "", regex=True)
                       .str.replace(",", "", regex=False),
                errors="coerce",
            )
            convertidas.append(col)
    nota = (
        f"Columnas convertidas a número: {convertidas}"
        if convertidas else "No se detectaron columnas de texto que fueran números."
    )
    return df, nota


def convertir_fechas(df):
    """
    Detecta columnas cuyo nombre o contenido sugiere fechas y las normaliza a
    formato de fecha. Conservador: solo si el 80%+ de los valores parsea.
    """
    convertidas = []
    pistas = ("fecha", "date", "dia", "día", "emision", "emisión", "vencim")

    def _parsear(serie):
        # format="mixed" permite distintos formatos por celda (15/01/2026,
        # 2026-01-16, 17-01-2026) sin forzar uno solo para toda la columna.
        return pd.to_datetime(
            serie, errors="coerce", dayfirst=True, format="mixed"
        )

    for col in df.columns:
        serie = df[col].dropna()
        if serie.empty:
            continue
        nombre_sugiere = any(p in str(col).lower() for p in pistas)
        es_texto = col in _columnas_texto(df)
        if es_texto or nombre_sugiere:
            try:
                parseada = _parsear(serie.astype(str))
            except (ValueError, TypeError):
                continue
            if parseada.notna().mean() >= 0.8:
                df[col] = _parsear(df[col].astype(str))
                convertidas.append(col)
    nota = (
        f"Columnas convertidas a fecha: {convertidas}"
        if convertidas else "No se detectaron columnas de fecha."
    )
    return df, nota


def estandarizar_texto(df):
    """
    Unifica variantes de una misma palabra que solo difieren en mayúsculas o
    espacios (p. ej. 'CUSCO', 'cusco', ' Cusco ' -> 'Cusco'). Para cada columna
    de texto, agrupa los valores por su forma en minúsculas y elige como forma
    correcta la escritura MÁS FRECUENTE que ya aparece en los datos. Así no
    rompe acrónimos legítimos (p. ej. 'Disco SSD' no se convierte en 'Disco Ssd').
    """
    from collections import Counter, defaultdict
    columnas_texto = _columnas_texto(df)
    total = 0
    for col in columnas_texto:
        serie = df[col]
        no_nulos = serie.dropna().astype(str)
        if no_nulos.empty:
            continue
        grupos = defaultdict(Counter)
        for original in no_nulos:
            clave = original.strip().lower()
            grupos[clave][original.strip()] += 1
        # Forma canónica = la escritura más común dentro de cada grupo.
        canon = {clave: cnt.most_common(1)[0][0] for clave, cnt in grupos.items()}

        def mapear(v):
            if isinstance(v, str):
                return canon.get(v.strip().lower(), v)
            return v

        nueva = serie.apply(mapear)
        total += int((serie.fillna("\x00") != nueva.fillna("\x00")).sum())
        df[col] = nueva
    return df, (
        f"Se unificaron {total} celdas con variantes de mayúsculas/espacios "
        f"(p. ej. 'CUSCO'/'cusco' → 'Cusco')."
    )


def quitar_duplicados(df):
    """Elimina filas completamente duplicadas."""
    antes = len(df)
    df = df.drop_duplicates()
    quitadas = antes - len(df)
    return df, f"Se eliminaron {quitadas} filas duplicadas."


# --------------------------------------------------------------------------- #
#  Orquestación                                                               #
# --------------------------------------------------------------------------- #

def cargar(ruta, hoja):
    """Carga Excel o CSV sin asumir dónde está el encabezado."""
    ruta = Path(ruta)
    if ruta.suffix.lower() in (".xlsx", ".xlsm", ".xls"):
        crudo = pd.read_excel(ruta, sheet_name=hoja, header=None)
    elif ruta.suffix.lower() in (".csv", ".tsv"):
        sep = "\t" if ruta.suffix.lower() == ".tsv" else None
        crudo = pd.read_csv(ruta, header=None, sep=sep, engine="python")
    else:
        raise ValueError(f"Formato no soportado: {ruta.suffix}")
    return crudo


def limpiar(ruta_entrada, ruta_salida=None, hoja=0, sin_duplicados=True):
    reporte = []
    reporte.append(f"Archivo de entrada: {ruta_entrada}")
    reporte.append(f"Procesado: {datetime.now():%Y-%m-%d %H:%M}")

    crudo = cargar(ruta_entrada, hoja)
    reporte.append(f"Dimensiones originales (crudas): {crudo.shape[0]} filas x {crudo.shape[1]} columnas")

    # 1. Encontrar la fila de encabezado real y re-encuadrar la tabla.
    fila_cab = detectar_fila_encabezado(crudo)
    df = crudo.iloc[fila_cab + 1:].copy()
    df.columns = crudo.iloc[fila_cab]
    reporte.append(f"Fila de encabezado detectada: fila {fila_cab + 1} del archivo.")

    # 2. Cadena de limpieza. Cada paso aporta una línea al reporte.
    pasos = [
        limpiar_nombres_columnas,
        quitar_filas_columnas_vacias,
        recortar_espacios,
        normalizar_vacios,
        convertir_numeros,
        convertir_fechas,
        estandarizar_texto,
    ]
    if sin_duplicados:
        pasos.append(quitar_duplicados)

    for paso in pasos:
        df, nota = paso(df)
        reporte.append(f"  - {nota}")

    df = df.reset_index(drop=True)
    reporte.append(f"Dimensiones finales (limpias): {df.shape[0]} filas x {df.shape[1]} columnas")

    # 3. Guardar.
    if ruta_salida is None:
        p = Path(ruta_entrada)
        ruta_salida = p.with_name(f"{p.stem}_limpio.xlsx")
    with pd.ExcelWriter(ruta_salida, engine="openpyxl") as xls:
        df.to_excel(xls, index=False, sheet_name="Datos limpios")
    reporte.append(f"Archivo limpio guardado en: {ruta_salida}")

    # 4. Guardar el reporte junto al archivo.
    ruta_reporte = Path(ruta_salida).with_name(f"{Path(ruta_salida).stem}_reporte.txt")
    ruta_reporte.write_text("\n".join(reporte), encoding="utf-8")

    return df, reporte, ruta_salida


def main():
    parser = argparse.ArgumentParser(
        description="Limpia un Excel o CSV desordenado y genera un reporte de cambios."
    )
    parser.add_argument("entrada", help="Ruta del archivo .xlsx, .xls o .csv a limpiar")
    parser.add_argument("-o", "--salida", default=None, help="Ruta del archivo de salida")
    parser.add_argument("--hoja", default=0, help="Nombre o número de hoja (Excel). Por defecto la primera.")
    parser.add_argument("--con-duplicados", action="store_true",
                        help="No eliminar filas duplicadas.")
    args = parser.parse_args()

    try:
        hoja = int(args.hoja)
    except (ValueError, TypeError):
        hoja = args.hoja

    try:
        _, reporte, salida = limpiar(
            args.entrada,
            args.salida,
            hoja=hoja,
            sin_duplicados=not args.con_duplicados,
        )
    except Exception as e:
        print(f"ERROR: {e}", file=sys.stderr)
        sys.exit(1)

    print("\n".join(reporte))
    print("\nListo. Revisa el archivo limpio y el reporte .txt que lo acompaña.")


if __name__ == "__main__":
    main()
