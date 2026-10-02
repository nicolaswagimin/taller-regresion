"""Verifica los modelos exportados en modelos/ (.joblib y .pkl).

Para cada ejercicio vuelve a entrenar el modelo original con los mismos datos y la
misma división (80/20, random_state=42) y comprueba que los modelos cargados desde
disco predicen exactamente lo mismo y que los metadatos guardados son correctos.

Uso:  venv/bin/python verificar_modelos.py
"""
import pickle
import sys
from pathlib import Path

import joblib
import numpy as np
import pandas as pd
from sklearn.linear_model import LinearRegression
from sklearn.metrics import mean_squared_error, r2_score
from sklearn.model_selection import train_test_split

BASE = Path(__file__).resolve().parent
EJERCICIOS = {
    "dolar": ("Precio_Dolar", ["Dia", "Inflacion", "Tasa_interes"]),
    "glucosa": ("Nivel_Glucosa", ["Edad", "IMC", "Actividad_Fisica"]),
    "energia": ("Consumo_Energia", ["Temperatura", "Hora", "Dia_Semana"]),
}


def cargar(ruta):
    if ruta.suffix == ".joblib":
        return joblib.load(ruta)
    with open(ruta, "rb") as f:
        return pickle.load(f)


def verificar(nombre, objetivo, variables):
    df = pd.read_csv(BASE / "data" / f"{nombre}_data.csv").dropna().drop_duplicates().reset_index(drop=True)
    X_train, X_test, y_train, y_test = train_test_split(
        df[variables], df[objetivo], test_size=0.2, random_state=42)
    original = LinearRegression().fit(X_train, y_train)
    pred_original = original.predict(X_test)
    mse = mean_squared_error(y_test, pred_original)
    r2 = r2_score(y_test, pred_original)

    print(f"\n=== {nombre.upper()} ({objetivo}) ===")
    todo_ok = True
    for ext in (".joblib", ".pkl"):
        ruta = BASE / "modelos" / f"modelo_{nombre}{ext}"
        paquete = cargar(ruta)
        modelo = paquete["modelo"]
        pred_cargado = modelo.predict(X_test)
        comprobaciones = {
            "predicciones iguales al modelo original": np.allclose(pred_cargado, pred_original),
            "coeficientes e intercepto iguales": np.allclose(modelo.coef_, original.coef_)
            and np.isclose(modelo.intercept_, original.intercept_),
            "nombres de variables correctos": paquete["variables"] == variables
            and paquete["objetivo"] == objetivo,
            "métricas guardadas correctas": np.isclose(paquete["metricas"]["prueba"]["MSE"], mse)
            and np.isclose(paquete["metricas"]["prueba"]["RMSE"], np.sqrt(mse))
            and np.isclose(paquete["metricas"]["prueba"]["R2"], r2),
            "rangos guardados correctos": all(
                np.isclose(paquete["rangos"][c]["min"], df[c].min())
                and np.isclose(paquete["rangos"][c]["max"], df[c].max()) for c in df.columns),
        }
        print(f"  {ruta.name}")
        for texto, ok in comprobaciones.items():
            print(f"    [{'OK' if ok else 'FALLO'}] {texto}")
            todo_ok = todo_ok and bool(ok)
        print(f"    diferencia máxima entre predicciones: {np.abs(pred_cargado - pred_original).max():.2e}")
    print(f"  Métricas en prueba -> MSE: {mse:,.4f} | RMSE: {np.sqrt(mse):,.4f} | R²: {r2:.4f}")
    return todo_ok


if __name__ == "__main__":
    resultados = [verificar(n, obj, var) for n, (obj, var) in EJERCICIOS.items()]
    if all(resultados):
        print("\nRESULTADO: los 3 modelos (.joblib y .pkl) cargan y predicen igual que los originales.")
    else:
        print("\nRESULTADO: hay comprobaciones que fallaron. Revisar la salida anterior.")
        sys.exit(1)
