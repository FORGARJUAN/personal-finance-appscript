# 📊 Personal Finance & Expense Tracker (Google Apps Script + Sheets)

> **Full-stack Serverless Single Page Application (SPA)** built with **Google Apps Script**, **Bootstrap 5**, **Chart.js**, and **Google Sheets** as a reactive database.

---

## 🇺🇸 English Version

### 📸 Screenshots & UI Showcase

<!-- Reemplaza los nombres de archivo con los nombres de tus capturas -->
| Expense Registration (SPA) | Financial Analytics Dashboard |
| :---: | :---: |
| ![Expense Register](screenshot_registro.png) | ![Analytics Dashboard](screenshot_dashboard.png) |

| Expense History & Filters | Security & Settings |
| :---: | :---: |
| ![History & Trend](screenshot_historial.png) | ![Settings & Setup](screenshot_ajustes.png) |

---

### 🚀 Overview
A lightweight, cloud-based financial management system designed to track personal and household income, expenses, and monthly budget targets. It eliminates manual spreadsheet data entry through an intuitive mobile-first web interface backed by Google Workspace infrastructure.

### ✨ Key Features
- **Mobile-First Responsive Interface:** Built with Bootstrap 5 and Bootstrap Icons, optimized for mobile devices.
- **Dynamic CRUD Operations:** Create, Read, Update, and permanently Delete transactions directly from the web client.
- **Analytics & Visual Dashboard:**
  - Doughnut distribution chart by category using **Chart.js**.
  - 7-day spending vs. income trend line chart.
  - Interactive monthly checklist with real-time tracking of pending and fulfilled obligations.
- **Dynamic Budget Projections:** Automated spreadsheet formula generation (`SUMIFS`) linking expenses to annual projection matrices.
- **Multi-Profile & Payment Methods:** Categorize expenses by profile (e.g., *Household*, *Personal*) and payment instrument (*Cash*, *Debit/Nequi*, *Credit Card*).
- **Attachment & Receipt Storage:** Uploads invoices and receipt images directly to designated monthly folders in **Google Drive**.
- **PIN-Based Access Control:** Secure actions (modifying categories, projections, and sensitive operations) protected by a hashed/stored PIN.
- **Data Portability:** One-click CSV and XLSX export/import pipelines.

### 🛠️ Tech Stack & Architecture
- **Frontend:** HTML5, CSS3, JavaScript (ES6+), Bootstrap 5, SweetAlert2, Chart.js.
- **Backend / Serverless:** Google Apps Script (V8 runtime Engine).
- **Database & Storage:** Google Sheets (Relational-like structure), Google Drive API (Blob & file management).
- **Communication:** Asynchronous client-server RPC via `google.script.run`.

---

## 🇪🇸 Versión en Español

### 🚀 Descripción General
Sistema web integral (Single Page Application) para la gestión y control de finanzas personales y del hogar. Permite registrar gastos e ingresos en tiempo real, visualizar analíticas interactivas y proyectar presupuestos mensuales, utilizando Google Sheets como base de datos en la nube y Google Drive como almacenamiento de comprobantes.

### ✨ Funcionalidades Principales
- **Interfaz Móvil Moderna:** Diseñada con Bootstrap 5, navegación tipo app nativa con barra inferior fija y transiciones fluidas.
- **Gestión Completa de Registros (CRUD):** Registro rápido, edición de montos/conceptos y eliminación segura de transacciones.
- **Dashboard Analítico Interactivo:**
  - Gráfico de dona con distribución de gastos por categoría (**Chart.js**).
  - Gráfico de tendencia de ingresos vs. egresos de los últimos 7 días.
  - Checklist interactivo de compromisos mensuales (pagados vs. pendientes).
- **Automatización de Fórmulas y Proyecciones:** Generación programática de hojas de cálculo con matrices de proyección anual basadas en fórmulas `SUMIFS`.
- **Carga de Evidencias a Google Drive:** Almacenamiento automático de fotos de facturas y recibos organizados por carpetas mensuales.
- **Seguridad por PIN:** Protección de configuraciones críticas, metas de ahorro y modificaciones de usuarios mediante clave PIN.
- **Importación y Exportación de Datos:** Descarga directa de reportes en formatos `.csv` y `.xlsx`, con soporte para restauración de datos.

### 🛠️ Tecnologías Utilizadas
- **Frontend:** HTML5, CSS3, JavaScript moderno (ES6), Bootstrap 5, SweetAlert2, Chart.js.
- **Backend:** Google Apps Script (Motor V8).
- **Base de Datos y Almacenamiento:** Google Sheets API, Google Drive API.

---

## 👨‍💻 Author
Developed by **FORGARJUAN** — *Full-Stack Micro-Apps & AI-Assisted Automation.*
