import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import * as path from "node:path";
import * as fs from "node:fs";
import * as XLSX from "xlsx";
import { error, texto } from "../lib/respuestas.js";

const CARPETA_SEGURA = path.resolve(process.cwd(), "datos_examen");

function resolverRutaSegura(nombreArchivo: string): string {
  const limpio = path.basename(nombreArchivo);
  return path.join(CARPETA_SEGURA, limpio);
}

export function register(server: McpServer): void {
  server.registerTool(
    "leer_excel_clinico",
    {
      title: "Leer Excel clínico o bioestadístico",
      description:
        "Lee un archivo Excel (.xlsx) ubicado en la carpeta autorizada 'datos_examen' y devuelve los nombres de sus hojas, las columnas detectadas y una vista previa de los datos. Úsala cuando el usuario pida revisar, explorar o inspeccionar una base de datos de laboratorio.",
      inputSchema: {
        archivo: z
          .string()
          .describe("Nombre del archivo Excel dentro de datos_examen (ej: hemogramas.xlsx)"),
        hoja: z
          .string()
          .optional()
          .describe("Nombre opcional de la pestaña a leer; si se omite, lee la primera hoja"),
      },
      annotations: { readOnlyHint: true },
    },
    async ({ archivo, hoja }) => {
      try {
        if (!archivo.endsWith(".xlsx")) {
          return error("Solo se permiten archivos con extensión .xlsx.");
        }
        const ruta = resolverRutaSegura(archivo);
        if (!fs.existsSync(ruta)) {
          return error(`El archivo '${path.basename(archivo)}' no existe en la carpeta datos_examen.`);
        }

        const workbook = XLSX.read(fs.readFileSync(ruta), { type: "buffer" });
        const nombreHoja = hoja && workbook.SheetNames.includes(hoja) ? hoja : workbook.SheetNames[0];
        const sheet = workbook.Sheets[nombreHoja];
        const filas = XLSX.utils.sheet_to_json<Record<string, unknown>>(sheet);

        const columnas = filas.length > 0 ? Object.keys(filas[0]) : [];
        const muestra = filas.slice(0, 10);

        return texto(
          JSON.stringify(
            {
              archivo: path.basename(archivo),
              hojasDisponibles: workbook.SheetNames,
              hojaLeida: nombreHoja,
              totalRegistros: filas.length,
              columnas,
              primeros10Registros: muestra,
            },
            null,
            2,
          ),
        );
      } catch (e) {
        return error(`Fallo al leer el Excel clínico: ${(e as Error).message}`);
      }
    },
  );
}