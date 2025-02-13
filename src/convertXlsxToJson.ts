import { parseExcelDate } from "./dateConvert";
import * as XLSX from "xlsx";
import * as fs from "fs";
import * as path from "path";

export function convertXlsxToJson(fileName: string) {
    const uploadsDir = "C:\\Users\\Kauê Carmo\\Desktop\\xlsxConvert\\src\\XLSX-IMPORT-TEST\\uploads";
    const filePath = path.join(uploadsDir, fileName);
    const outputJsonPath = path.join(uploadsDir, fileName.replace(".xlsx", ".json"));

    if (!fs.existsSync(filePath)) {
        console.error(`Arquivo não encontrado: ${filePath}`);
        return;
    }

    const workbook = XLSX.readFile(filePath);
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];

    let jsonData = XLSX.utils.sheet_to_json(sheet, { defval: null });

    // Captura todas as colunas existentes, ignorando colunas vazias ("__EMPTY", "__EMPTY_1", etc.)
    const allColumns = new Set<string>();
    jsonData.forEach(row => {
        Object.keys(row)
            .filter(col => !col.startsWith("__EMPTY")) // Remove colunas vazias
            .forEach(col => allColumns.add(col));
    });

    // Preenche colunas ausentes com null e converte Data Outorga
    jsonData = jsonData.map((row: any) => {
        const completeRow: any = {};

        allColumns.forEach(col => {
            completeRow[col] = row[col] ?? null;
        });

        // Verifica e converte a data, se necessário
        if (typeof completeRow["Data_Outorga"] === "number") {
            completeRow["Data_Outorga"] = parseExcelDate(completeRow["Data_Outorga"]).toISOString().split("T")[0];
        }

        return completeRow;
    });

    fs.writeFileSync(outputJsonPath, JSON.stringify(jsonData, null, 2), "utf-8");
    console.log(`Conversão concluída! JSON salvo em: ${outputJsonPath}`);
}