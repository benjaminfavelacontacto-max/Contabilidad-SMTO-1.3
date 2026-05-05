# FinDash — Dashboard Financiero

Dashboard financiero 100% del lado del cliente. Sube tu Excel y obtén un análisis visual profesional al instante.

---

## 🚀 Inicio rápido

1. Abre `index.html` en cualquier navegador moderno (Chrome, Firefox, Edge, Safari).
2. Arrastra tu archivo `.xlsx` al área de carga o haz clic para seleccionarlo.
3. El dashboard se genera automáticamente.

> No se necesita servidor, instalación ni conexión a internet (solo para la carga inicial de fuentes y librerías CDN).

---

## 📋 Formato del Excel esperado

El sistema detecta automáticamente los encabezados. Se recomienda tener columnas con nombres similares a:

| Columna       | Nombres aceptados                                              |
|---------------|---------------------------------------------------------------|
| Fecha         | `Fecha`, `Date`, `Dia`, `Periodo`                             |
| Tipo          | `Tipo`, `Type`, `Movimiento`, `Naturaleza`                    |
| Categoría     | `Categoria`, `Category`, `Rubro`, `Descripcion`               |
| Subcategoría  | `Subcategoria`, `Subcat`, `Detalle`                           |
| Monto         | `Monto`, `Amount`, `Valor`, `Importe`, `Total`                |

### Valores para columna Tipo
- **Ingreso**: `ingreso`, `income`, `entrada`, `credito`, `in`, `+`
- **Egreso**: `egreso`, `gasto`, `expense`, `salida`, `debito`, `out`, `-`
- Si no hay columna de tipo, se infiere por el signo del monto (positivo = ingreso, negativo = egreso).

---

## 📊 Funcionalidades

- **KPIs automáticos**: Total ingresos, egresos, balance neto y tasa de ahorro.
- **Gráfica de barras**: Ingresos vs Egresos por mes.
- **Gráfica de dona**: Distribución de egresos por categoría.
- **Tabla de categorías**: Total, porcentaje y número de transacciones.
- **Tabla de transacciones**: Últimas 100 transacciones detalladas.
- **Filtro por mes**: En el header, menú desplegable para filtrar.
- **Exportar CSV**: Descarga los datos filtrados como archivo CSV.
- **Dark / Light mode**: Botón en el header para cambiar tema.

---

## 🗂 Estructura del proyecto

```
findash/
├── index.html    # Estructura y layout HTML
├── styles.css    # Estilos dark mode premium
├── script.js     # Lógica: lectura Excel, procesamiento, gráficas
└── README.md     # Este archivo
```

---

## 🛠 Tecnologías

- **HTML5 / CSS3 / JavaScript ES6+** — sin frameworks
- **SheetJS (xlsx)** — lectura de archivos Excel
- **Chart.js 4** — gráficas interactivas
- **Google Fonts (Inter)** — tipografía

---

## 🔒 Privacidad

Todo el procesamiento ocurre en el navegador. Los datos **nunca se envían a ningún servidor**.

---

## 📝 Licencia

MIT — libre para uso personal y comercial.
