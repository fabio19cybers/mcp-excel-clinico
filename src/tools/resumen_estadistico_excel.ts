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
    "resumen_estadistico_excel",
    {
      title: "Generar resumen estadístico en Excel",
      description:
        "Calcula estadísticos descriptivos (N, Media, Desviación Estándar, Mínimo y Máximo) de una columna numérica en un Excel de 'datos_examen' y crea una pestaña llamada 'Resumen_Estadistico' con los resultados.",
      inputSchema: {
        archivo: z
          .string()
          .describe("Nombre del archivo Excel en datos_examen (ej: hemogramas.xlsx)"),
        columna: z
          .string()
          .describe("Nombre exacto de la columna numérica a analizar (ej: Glicemia)"),
        hojaOrigen: z
          .string()
          .optional()
          .describe("Nombre opcional de la hoja origen; si se omite, usa la primera hoja"),
      },
      annotations: { readOnlyHint: false, destructiveHint: false },
    },
    async ({ archivo, columna, hojaOrigen }) => {
      try {
        if (!archivo.endsWith(".xlsx")) {
          return error("Solo se permiten archivos con extensión .xlsx.");
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

        const valores = filas
          .map((f) => Number(f[columna]))
          .filter((v) => !Number.isNaN(v));

        if (valores.length === 0) {
          return error(`La columna '${columna}' no contiene valores numéricos válidos.`);
        }

        const n = valores.length;
        const suma = valores.reduce((acc, v) => acc + v, 0);
        const media = suma / n;
        const varianza =
          n > 1 ? valores.reduce((acc, v) => acc + (v - media) ** 2, 0) / (n - 1) : 0;
        const desviacionEstandar = Math.sqrt(varianza);
        const minimo = Math.min(...valores);
        const maximo = Math.max(...valores);

        const tablaResumen = [
          {
            Variable: columna,
            Hoja_Origen: nombreHoja,
            N_Muestras: n,
            Promedio: Number(media.toFixed(2)),
            Desviacion_Estandar: Number(desviacionEstandar.toFixed(2)),
            Minimo: minimo,
            Maximo: maximo,
          },
        ];

        const nuevaHoja = XLSX.utils.json_to_sheet(tablaResumen);
        const nombreDestino = "Resumen_Estadistico";
        const idxExistente = workbook.SheetNames.indexOf(nombreDestino);
        if (idxExistente !== -1) {
          workbook.Sheets[nombreDestino] = nuevaHoja;
        } else {
          XLSX.utils.book_append_sheet(workbook, nuevaHoja, nombreDestino);
        }

        XLSX.writeFile(workbook, ruta);

        return texto(
          `Resumen estadístico generado en '${path.basename(archivo)}' (pestaña '${nombreDestino}'):\n` +
            JSON.stringify(tablaResumen[0], null, 2),
        );
      } catch (e) {
        return error(`Fallo al calcular estadísticas: ${(e as Error).message}`);
      }
    },
  );}