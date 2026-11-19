# MCP Excel Clínico — Bioanálisis y Bioestadística

Servidor MCP (*Model Context Protocol*) desarrollado sobre `mcp-starter-kit` que conecta un agente de IA local (`opencode`) con archivos de Microsoft Excel (`.xlsx`) para automatizar la inspección, filtrado clínico y análisis estadístico descriptivo de bases de datos de laboratorio.

---

## 1. Requisitos e Instalación Paso a Paso

### Requisitos previos
* **Node.js** v20 o superior (LTS).
* **Git** instalado en el sistema.
* **Microsoft Excel** (o visor compatible `.xlsx`) para revisar los resultados.
* Cliente compatible con MCP (`opencode` o Google Antigravity) configurado con un proveedor de IA (ej. OpenRouter).

### Instalación y compilación
1. Clonar este repositorio y entrar en la carpeta del proyecto:
   ```bash
   git clone <URL_DE_TU_REPOSITORIO_EN_GITHUB>
   cd mcp-excel-clinico