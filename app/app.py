"""Interfaz web del Laboratorio 1 de Minería de Datos.

Permite elegir un ejercicio (Dólar, Glucosa o Energía), ingresar por teclado los
valores de las variables independientes y obtener la predicción del modelo de
regresión lineal múltiple exportado en modelos/.

Uso:  venv/bin/streamlit run app/app.py
"""
from pathlib import Path

import joblib
import pandas as pd
import streamlit as st

MODELOS = Path(__file__).resolve().parent.parent / "modelos"

# Textos de presentación de cada ejercicio (el resto sale del modelo exportado)
EJERCICIOS = {
    "Dólar": {
        "archivo": "modelo_dolar.joblib",
        "descripcion": "Predicción del precio del dólar a partir del día, la inflación y la tasa de interés.",
        "ejemplo": {"Dia": "250", "Inflacion": "0.02", "Tasa_interes": "5.0"},
        "ayuda": {
            "Dia": "Número de día (entero, desde 1).",
            "Inflacion": "Tasa de inflación diaria, como proporción (por ejemplo 0.02).",
            "Tasa_interes": "Tasa de interés diaria (por ejemplo 5.0).",
        },
    },
    "Glucosa": {
        "archivo": "modelo_glucosa.joblib",
        "descripcion": "Predicción del nivel de glucosa en sangre a partir de la edad, el IMC y la actividad física.",
        "ejemplo": {"Edad": "50", "IMC": "25.0", "Actividad_Fisica": "5"},
        "ayuda": {
            "Edad": "Edad del paciente en años.",
            "IMC": "Índice de masa corporal (kg/m²).",
            "Actividad_Fisica": "Horas de ejercicio por semana.",
        },
    },
    "Energía": {
        "archivo": "modelo_energia.joblib",
        "descripcion": "Predicción del consumo de energía a partir de la temperatura, la hora y el día de la semana.",
        "ejemplo": {"Temperatura": "25.0", "Hora": "12", "Dia_Semana": "3"},
        "ayuda": {
            "Temperatura": "Temperatura en grados Celsius.",
            "Hora": "Hora del día (entero de 1 a 24).",
            "Dia_Semana": "Día de la semana (entero: 1 = lunes ... 7 = domingo).",
        },
    },
}


@st.cache_resource
def cargar_paquete(archivo):
    """Carga el modelo exportado junto con sus variables, métricas y rangos."""
    return joblib.load(MODELOS / archivo)


def numero(valor):
    """Formatea un número sin ceros sobrantes."""
    return f"{valor:,.4f}".rstrip("0").rstrip(".")


def validar(texto, variable, paquete):
    """Convierte el texto ingresado en número y lo valida.

    Devuelve (valor, error, advertencia): `error` impide predecir; `advertencia`
    indica que el valor está fuera del rango de los datos de entrenamiento.
    """
    texto = texto.strip().replace(",", ".")
    if not texto:
        return None, f"**{variable}**: el campo está vacío.", None
    try:
        valor = float(texto)
    except ValueError:
        return None, f"**{variable}**: «{texto}» no es un número válido.", None
    if valor != valor or valor in (float("inf"), float("-inf")):
        return None, f"**{variable}**: «{texto}» no es un número válido.", None
    if variable in paquete["variables_enteras"] and valor != int(valor):
        return None, f"**{variable}**: debe ser un número entero.", None

    minimo, maximo = paquete["limites_validos"][variable]
    if minimo is not None and valor < minimo:
        return None, f"**{variable}**: no puede ser menor que {numero(minimo)}.", None
    if maximo is not None and valor > maximo:
        return None, f"**{variable}**: no puede ser mayor que {numero(maximo)}.", None

    rango = paquete["rangos"][variable]
    advertencia = None
    if valor < rango["min"] or valor > rango["max"]:
        advertencia = (
            f"**{variable}** = {numero(valor)} está fuera del rango de los datos de entrenamiento "
            f"({numero(rango['min'])} a {numero(rango['max'])}). La predicción es una extrapolación "
            "y puede ser poco confiable."
        )
    return valor, None, advertencia


st.set_page_config(page_title="Predicciones con regresión lineal", page_icon="📈", layout="centered")
st.title("Predicciones con regresión lineal múltiple")
st.caption("Laboratorio 1 – Minería de Datos · Modelos entrenados con scikit-learn siguiendo CRISP-DM")

ejercicio = st.radio("Seleccione el ejercicio", list(EJERCICIOS), horizontal=True, key="ejercicio")
config = EJERCICIOS[ejercicio]

try:
    paquete = cargar_paquete(config["archivo"])
except FileNotFoundError:
    st.error(
        f"No se encontró el modelo `modelos/{config['archivo']}`. "
        "Ejecute primero los notebooks de la carpeta `notebooks/` para generarlo."
    )
    st.stop()

variables = paquete["variables"]
objetivo = paquete["objetivo"]
unidad = paquete["unidad_objetivo"]
rangos = paquete["rangos"]

st.subheader(f"Ejercicio: {ejercicio}")
st.write(config["descripcion"])

st.markdown("**Ingrese los valores de las variables independientes**")
entradas = {}
columnas = st.columns(len(variables))
for columna, variable in zip(columnas, variables):
    with columna:
        entradas[variable] = st.text_input(
            paquete["etiquetas"][variable],
            value=config["ejemplo"][variable],
            key=f"{ejercicio}_{variable}",
            help=config["ayuda"][variable],
        )
        st.caption(
            f"Rango en los datos: {numero(rangos[variable]['min'])} a {numero(rangos[variable]['max'])}"
        )

if st.button("Calcular predicción", type="primary", key="predecir"):
    valores, errores, advertencias = {}, [], []
    for variable in variables:
        valor, error, advertencia = validar(entradas[variable], variable, paquete)
        valores[variable] = valor
        if error:
            errores.append(error)
        if advertencia:
            advertencias.append(advertencia)

    if errores:
        for error in errores:
            st.error(error)
        st.info("Corrija los valores indicados y vuelva a calcular.")
    else:
        fila = pd.DataFrame([valores])[variables]
        prediccion = float(paquete["modelo"].predict(fila)[0])
        st.metric(f"{objetivo} estimado", f"{prediccion:,.2f} {unidad}")
        st.caption(
            f"Error típico del modelo (RMSE en prueba): ± {paquete['metricas']['prueba']['RMSE']:,.2f} {unidad}"
        )
        for advertencia in advertencias:
            st.warning(advertencia)

        st.markdown("**¿Cómo se obtuvo este valor?**")
        aportes = pd.DataFrame(
            [{"Término": "Intercepto", "Valor ingresado": "", "Coeficiente": "",
              f"Aporte ({unidad})": f"{paquete['intercepto']:,.2f}"}]
            + [{"Término": v, "Valor ingresado": numero(valores[v]),
                "Coeficiente": f"{paquete['coeficientes'][v]:,.4f}",
                f"Aporte ({unidad})": f"{paquete['coeficientes'][v] * valores[v]:,.2f}"} for v in variables]
            + [{"Término": "Predicción (suma)", "Valor ingresado": "", "Coeficiente": "",
                f"Aporte ({unidad})": f"{prediccion:,.2f}"}]
        )
        st.table(aportes.set_index("Término"))

st.divider()
st.subheader("Información del modelo")

st.markdown("**Ecuación**")
st.code(paquete["ecuacion"], language=None)

st.markdown("**Métricas de desempeño**")
metricas = pd.DataFrame(paquete["metricas"]).T.rename(
    index={"entrenamiento": "Entrenamiento (80 %)", "prueba": "Prueba (20 %)"}, columns={"R2": "R²"})
st.table(metricas.style.format({"MSE": "{:,.2f}", "RMSE": "{:,.2f}", "R²": "{:.4f}"}))

st.markdown("**Coeficientes y rangos de los datos de entrenamiento**")
detalle = pd.DataFrame({
    "Coeficiente": [f"{paquete['coeficientes'][v]:,.4f}" for v in variables],
    "Coef. estandarizado": [f"{paquete['coef_estandarizados'][v]:+.3f}" for v in variables],
    "p-valor": ["< 0.001" if paquete["p_valores"][v] < 0.001 else f"{paquete['p_valores'][v]:.3f}"
                for v in variables],
    "Mínimo": [numero(rangos[v]["min"]) for v in variables],
    "Máximo": [numero(rangos[v]["max"]) for v in variables],
    "Media": [numero(rangos[v]["media"]) for v in variables],
}, index=variables)
st.table(detalle)
st.caption(
    f"{objetivo} en los datos: de {numero(rangos[objetivo]['min'])} a {numero(rangos[objetivo]['max'])} {unidad} "
    f"(media {numero(rangos[objetivo]['media'])}). Modelo entrenado con {paquete['n_entrenamiento']} registros "
    f"y evaluado con {paquete['n_prueba']}."
)
