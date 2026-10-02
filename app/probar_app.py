"""Prueba automática de la interfaz web con streamlit.testing.

Simula a un usuario en los 3 escenarios (Dólar, Glucosa, Energía) y comprueba que
la predicción mostrada en pantalla es la misma que da el modelo exportado. También
prueba la validación de entradas.

Uso:  venv/bin/python app/probar_app.py
"""
from pathlib import Path

import joblib
import pandas as pd
from streamlit.testing.v1 import AppTest

APP = Path(__file__).resolve().parent / "app.py"
MODELOS = APP.parent.parent / "modelos"

ESCENARIOS = {
    "Dólar": ("modelo_dolar.joblib", {"Dia": "300", "Inflacion": "0.021", "Tasa_interes": "5.2"}),
    "Glucosa": ("modelo_glucosa.joblib", {"Edad": "45", "IMC": "27.5", "Actividad_Fisica": "3"}),
    "Energía": ("modelo_energia.joblib", {"Temperatura": "30.5", "Hora": "18", "Dia_Semana": "6"}),
}


def abrir(ejercicio):
    at = AppTest.from_file(str(APP), default_timeout=30).run()
    at.radio(key="ejercicio").set_value(ejercicio).run()
    assert not at.exception, at.exception
    return at


def predecir(at, ejercicio, valores):
    for variable, texto in valores.items():
        at.text_input(key=f"{ejercicio}_{variable}").input(texto)
    at.button(key="predecir").click().run()
    assert not at.exception, at.exception
    return at


def probar_escenarios():
    for ejercicio, (archivo, valores) in ESCENARIOS.items():
        paquete = joblib.load(MODELOS / archivo)
        fila = pd.DataFrame([{v: float(t) for v, t in valores.items()}])[paquete["variables"]]
        esperado = f"{paquete['modelo'].predict(fila)[0]:,.2f} {paquete['unidad_objetivo']}"

        at = predecir(abrir(ejercicio), ejercicio, valores)
        mostrado = at.metric[0].value
        assert mostrado == esperado, f"{ejercicio}: se mostró {mostrado}, se esperaba {esperado}"
        assert not at.error and not at.warning, f"{ejercicio}: mensajes inesperados"
        assert paquete["ecuacion"] in at.code[0].value
        print(f"[OK] {ejercicio}: entradas {valores} -> {mostrado}")


def probar_validaciones():
    # Texto que no es un número
    at = predecir(abrir("Glucosa"), "Glucosa", {"Edad": "abc"})
    assert len(at.error) == 1 and "no es un número válido" in at.error[0].value and not at.metric
    print("[OK] Validación: texto no numérico -> error y no se predice")

    # Campo vacío
    at = predecir(abrir("Dólar"), "Dólar", {"Inflacion": ""})
    assert len(at.error) == 1 and "vacío" in at.error[0].value and not at.metric
    print("[OK] Validación: campo vacío -> error y no se predice")

    # Valor imposible (hora 25) y valor no entero
    at = predecir(abrir("Energía"), "Energía", {"Hora": "25", "Dia_Semana": "2.5"})
    assert len(at.error) == 2 and not at.metric
    assert "no puede ser mayor que 24" in at.error[0].value and "entero" in at.error[1].value
    print("[OK] Validación: hora 25 y día 2.5 -> dos errores y no se predice")

    # Valor válido pero fuera del rango de los datos: predice y avisa
    at = predecir(abrir("Dólar"), "Dólar", {"Dia": "600"})
    assert len(at.warning) == 1 and "fuera del rango" in at.warning[0].value and len(at.metric) == 1
    print(f"[OK] Validación: día 600 (fuera de rango) -> advertencia y predicción {at.metric[0].value}")

    # Coma decimal aceptada
    at = predecir(abrir("Glucosa"), "Glucosa", {"IMC": "27,5"})
    assert not at.error and len(at.metric) == 1
    print(f"[OK] Validación: coma decimal «27,5» aceptada -> {at.metric[0].value}")


if __name__ == "__main__":
    probar_escenarios()
    probar_validaciones()
    print("\nRESULTADO: la interfaz funciona en los 3 escenarios y valida las entradas.")
