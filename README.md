# Laboratorio 1 – Minería de Datos: regresión lineal múltiple

Tres modelos de regresión lineal múltiple desarrollados con la metodología CRISP-DM,
exportados a disco y usados desde una interfaz web:

| Ejercicio | Variable dependiente | Variables independientes |
|---|---|---|
| 1. Dólar | `Precio_Dolar` | `Dia`, `Inflacion`, `Tasa_interes` |
| 2. Glucosa | `Nivel_Glucosa` | `Edad`, `IMC`, `Actividad_Fisica` |
| 3. Energía | `Consumo_Energia` | `Temperatura`, `Hora`, `Dia_Semana` |

## Estructura

```
taller_regresion/
├── enunciado.pdf              Enunciado del laboratorio
├── requirements.txt           Dependencias
├── data/                      Los 3 datasets (CSV)
├── notebooks/                 Un notebook por ejercicio (fases CRISP-DM)
│   ├── 01_dolar.ipynb
│   ├── 02_glucosa.ipynb
│   └── 03_energia.ipynb
├── modelos/                   Modelos exportados (.joblib y .pkl) y metadatos (.json)
├── verificar_modelos.py       Comprueba que los modelos exportados predicen igual
├── app/
│   ├── app.py                 Interfaz web (Streamlit)
│   └── probar_app.py          Prueba automática de la interfaz
└── informe/
    ├── informe.docx           Informe escrito
    ├── generar_informe.py     Script que genera el informe
    └── figuras/               Gráficas (PNG) generadas por los notebooks
```

## 1. Preparar el entorno

Todos los comandos se ejecutan desde la carpeta `taller_regresion/`.

```bash
python3 -m venv venv                         # solo si no existe la carpeta venv/
venv/bin/pip install -r requirements.txt
```

En Windows se usa `venv\Scripts\` en lugar de `venv/bin/`.

## 2. Ejecutar los notebooks

Abrirlos con Jupyter o VS Code (seleccionando el intérprete de `venv/`) y usar
«Ejecutar todo», o ejecutarlos desde la terminal:

```bash
cd notebooks
../venv/bin/jupyter nbconvert --to notebook --execute --inplace 01_dolar.ipynb 02_glucosa.ipynb 03_energia.ipynb
cd ..
```

Cada notebook entrena su modelo, guarda las gráficas en `informe/figuras/` y exporta el
modelo en `modelos/` (`modelo_<ejercicio>.joblib`, `modelo_<ejercicio>.pkl` y
`metadatos_<ejercicio>.json`).

## 3. Verificar los modelos exportados

```bash
venv/bin/python verificar_modelos.py
```

Carga los seis archivos y comprueba que predicen igual que los modelos originales.

## 4. Abrir la interfaz web

```bash
venv/bin/streamlit run app/app.py
```

Se abre en el navegador (http://localhost:8501). Se elige el ejercicio, se escriben los
valores y se pulsa «Calcular predicción». Para cerrarla: `Ctrl + C` en la terminal.

Prueba automática de la interfaz en los 3 escenarios:

```bash
venv/bin/python app/probar_app.py
```

## 5. Regenerar el informe

```bash
venv/bin/python informe/generar_informe.py
```

Vuelve a crear `informe/informe.docx` con los resultados de `modelos/metadatos_*.json`.
**Atención:** sobrescribe el archivo, así que se pierden los cambios hechos a mano
(portada, capturas de pantalla). Si ya se editó el informe, guardar antes una copia.

## Resultados (conjunto de prueba, 20 %)

| Ejercicio | MSE | RMSE | R² | Variable de mayor impacto |
|---|---|---|---|---|
| Dólar | 2 376.97 | 48.75 pesos | 0.9963 | `Dia` |
| Glucosa | 233.69 | 15.29 mg/dL | 0.6814 | `Edad` |
| Energía | 429.52 | 20.72 kWh | 0.8968 | `Temperatura` |

## Despliegue en Streamlit Community Cloud

1. Entrar a https://share.streamlit.io con la cuenta de GitHub y elegir «Create app».
2. Repositorio: `nicolaswagimin/taller-regresion` · Rama: `main` · Archivo principal: `app/app.py`.
3. En «Advanced settings» elegir la misma versión de Python con la que se entrenaron
   los modelos (3.14) o la más reciente disponible.
4. Pulsar «Deploy». Las dependencias se instalan desde `requirements.txt`, que fija las
   mismas versiones usadas al entrenar (en especial `scikit-learn==1.9.1`).

La app carga los modelos con rutas relativas a `app/app.py` (`Path(__file__)`), así que
funciona igual en el servidor que en local.
