
# Guía de Compilación: Staffing Optimizer a .EXE

Esta guía detalla los pasos para convertir el script de Python (`main.py`) que utiliza **Flet**, **PuLP**, **HighSpy** y **OpenPyXL** en un único archivo ejecutable (`.exe`) para Windows.

## 📋 1. Requisitos Previos

Asegúrate de tener instalado Python (versión 3.9 o superior recomendada).

## 🛠️ 2. Preparación del Entorno (Virtual Environment)

Para evitar errores de dependencias y reducir el tamaño del archivo final, es **crucial** trabajar en un entorno limpio.

1.  Abre tu terminal (PowerShell o CMD) en la carpeta del proyecto.

2.  Crea el entorno virtual: ```
```
python -m venv .venv
```

3.  Activa el entorno:
 * **Windows:**
```
.\.venv\Scripts\activate
```
(Verás `(.venv)` al inicio de tu línea de comandos).

4.  Instala **SOLO** las librerías necesarias:
```
pip install flet pulp openpyxl highspy pyinstaller
```

---

## 🧹 3. Limpieza (Importante si falló antes)

Si has intentado compilar anteriormente y falló, debes borrar los archivos temporales para evitar conflictos de configuración:

1.  Borra la carpeta **`build`**.
2.  Borra la carpeta **`dist`**.
3.  Borra el archivo **`StaffingOptimizer.spec`** (si existe).

---

## 🚀 4. El Comando de Compilación

Este es el paso crítico. Usaremos `pyinstaller` directamente con flags específicos para asegurar que las librerías matemáticas (que suelen ocultarse) se incluyan correctamente.

Asegúrate de estar en la carpeta donde está `main.py` y ejecuta:

```
pyinstaller --name "StaffingOptimizer" --onefile --console --collect-all pulp --hidden-import=flet --hidden-import=highspy --hidden-import=openpyxl main.py
```
Hace falta poner "--collect-all pulp" porque la libreria PuLP llama a un .exe cuando ejecuta CBC. Si ponemos "--hidden-import=pulp" solo, no se importa el .exe que ejecuta CBC y entonces te dice siempre que el modelo es infeasible.