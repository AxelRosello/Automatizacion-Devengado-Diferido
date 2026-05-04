# 🖥️ Sistema de Procesamiento Contable con Interfaz Gráfica (GUI)

### 🚀 Descripción del Proyecto
Este software es una solución avanzada de automatización diseñada para gestionar cierres contables mensuales de forma intuitiva. A diferencia de scripts convencionales, esta herramienta incluye una **interfaz gráfica (GUI)** que permite a cualquier usuario seleccionar el mes de proceso sin necesidad de interactuar con el código.

---

### 🛠️ Problemas Solucionados

* **Facilidad de Uso**: Implementación de una ventana emergente para que el personal administrativo seleccione el periodo contable fácilmente.
* **Trazabilidad y Auditoría**: Generación automática de archivos **.log** que registran cada factura procesada, errores de formato y éxitos de ejecución.
* **Extracción de Datos No Estructurados**: Uso de **Expresiones Regulares (Regex)** para detectar periodos de servicio dentro de descripciones de texto complejas.
* **Cálculo Automático de Diferidos**: Determina con precisión los días pendientes y los montos a diferir para el mes siguiente, separando automáticamente divisas (PEN/USD).

---

### 💻 Características Técnicas

* **Librerías Clave**: `Tkinter` (interfaz), `Pandas` (motor de datos) y `Regex` (análisis de texto).
* **Manejo de Errores**: Sistema de excepciones robusto que evita que el programa se detenga ante datos inconsistentes.
* **Estructura de Datos**: Exportación normalizada a Excel con 33 columnas detalladas listas para ser importadas a sistemas ERP como Odoo.

---

### 📂 Estructura del Repositorio

* **devengado_diferido.py**: Código fuente principal con la lógica de negocio e interfaz.
* **ejecucion_proceso.log**: Archivo generado automáticamente para seguimiento de auditoría.
