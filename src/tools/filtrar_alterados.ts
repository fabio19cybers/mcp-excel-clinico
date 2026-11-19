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
    "filtrar_alterados",
    {
      title: "Filtrar valores clínicos fuera de rango",
      description:
        "Filtra los registros de un archivo Excel donde una columna numérica (ej. Hemoglobina, Glicemia, Leucocitos) esté por debajo de un mínimo o por encima de un máximo de referencia, y crea una nueva hoja llamada 'Alterados' en el mismo Excel. Úsala para detectar casos críticos o fuera de rango.",
      inputSchema: {
        archivo: z
          .string()
          .describe("Nombre del archivo Excel en datos_examen (ej: hemogramas.xlsx)"),
        columna: z
          .string()
          .describe("Nombre exacto de la columna numérica a evaluar (ej: Hemoglobina)"),
        minimo: z
          .number()
          .describe("Valor mínimo normal de referencia (valores menores se consideran alterados)"),
        maximo: z
          .number()
          .describe("Valor máximo normal de referencia (valores mayores se consideran alterados)"),
        hojaOrigen: z
          .string()
          .optional()
          .describe("Nombre opcional de la hoja origen; si se omite, usa la primera hoja"),
      },
      annotations: { readOnlyHint: false, destructiveHint: false },
    },
    async ({ archivo, columna, minimo, maximo, hojaOrigen }) => {
      try {
        if (!archivo.endsWith(".xlsx")) {
          return error("Solo se permiten archivos con extensión .xlsx.");
        }
        if (minimo > maximo) {
          return error("El valor mínimo no puede ser mayor que el máximo.");
        }
        const ruta = resolverRutaSegura(archivo);
        if (!fs.existsSync(ruta)) {
          return error(`El archivo '${path.basename(archivo)}' no existe en datos_examen.`);
        }

        const workbook = XLSX.read(fs.readFileSync(ruta), { type: "buffer" });
        const nombreHoja =
          hojaOrigen && workbook.SheetNames.includes(hojaOrigen)
            ? hojaOrigen
            : workbook.SheetNames[0];
        const sheet = workbook.Sheets[nombreHoja];
        const filas = XLSX.utils.sheet_to_json<Record<string, unknown>>(sheet);

        if (filas.length === 0 || !(columna in filas[0])) {
          return error(`La columna '${columna}' no existe en la hoja '${nombreHoja}'.`);
        }

        const alterados = filas
          .filter((fila) => {
            const val = Number(fila[columna]);
            return !Number.isNaN(val) && (val < minimo || val > maximo);
          })
          .map((fila) => {
            const val = Number(fila[columna]);
            return {
              ...fila,
              Interpretacion: val < minimo ? "BAJO (< min)" : "ALTO (> max)",
            };
          });

        const nuevaHoja = XLSX.utils.json_to_sheet(alterados);
        const nombreDestino = "Alterados";
        const idxExistente = workbook.SheetNames.indexOf(nombreDestino);
        if (idxExistente !== -1) {
          workbook.Sheets[nombreDestino] = nuevaHoja;
        } else {
          XLSX.utils.book_append_sheet(workbook, nuevaHoja, nombreDestino);
        }

        XLSX.writeFile(workbook, ruta);

        return texto(
          `Filtrado completado en '${path.basename(archivo)}'. Se encontraron ${alterados.length} registros alterados de ${filas.length} totales para '${columna}' (rango ${minimo} - ${maximo}). Guardados en la pestaña '${nombreDestino}'.`,
        );
      } catch (e) {
        return error(`Fallo al filtrar alterados: ${(e as Error).message}`);
      }
    },
  );
}