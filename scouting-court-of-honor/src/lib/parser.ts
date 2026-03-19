import * as XLSX from "xlsx";

export interface ScoutAwardInfo {
  scoutName: string;
  originalName: string;
  type: string;
  award: string;
  earnedDate?: string | number;
  // include anything else needed
}

/**
 * Parses a File object (CSV or XLSX) into an array of ScoutAwardInfo objects.
 */
export async function parseUploadFile(file: File): Promise<ScoutAwardInfo[]> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();

    reader.onload = (e) => {
      try {
        const data = e.target?.result;
        // Read file using sheetjs
        const workbook = XLSX.read(data, { type: "binary" });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];

        // Convert exactly to JSON
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 }) as any[][];

        if (!jsonData || jsonData.length === 0) {
          resolve([]);
          return;
        }

        // Find headers
        let headerRowIndex = 0;
        let headers: string[] = [];

        for (let i = 0; i < jsonData.length; i++) {
          const row = jsonData[i].filter((c) => c != null && c !== "");
          if (row.length > 2) {
            headerRowIndex = i;
            headers = jsonData[i].map((c) => (c ? String(c).trim().toLowerCase() : ""));
            break;
          }
        }

        const scoutIdx = headers.findIndex((h) => h.includes("scout") || h.includes("name"));
        const typeIdx = headers.findIndex((h) => h === "type");
        const awardIdx = headers.findIndex((h) => h === "award");
        const earnedIdx = headers.findIndex((h) => h.includes("earn"));

        if (scoutIdx === -1 || typeIdx === -1 || awardIdx === -1) {
          throw new Error("Missing required columns: Scout, Type, or Award.");
        }

        const parsedData: ScoutAwardInfo[] = [];

        for (let i = headerRowIndex + 1; i < jsonData.length; i++) {
          const row = jsonData[i];
          // skip blank lines completely
          if (!row || !row[scoutIdx] || String(row[scoutIdx]).trim() === "") {
            continue;
          }

          let rawName = String(row[scoutIdx]).trim();
          if (rawName.includes(",")) {
            const parts = rawName.split(",");
            if (parts.length === 2) {
              rawName = `${parts[1].trim()} ${parts[0].trim()}`;
            }
          }

          let rawAward = String(row[awardIdx]).trim();
          const rawType = String(row[typeIdx]).trim();

          // Clean up the dirty award strings like "Swimming (2022 rqmts)*"
          if (rawType.toLowerCase().includes("merit badge")) {
             // Remove any parenthetical grouping and asterisks from Merit Badges
             rawAward = rawAward.replace(/\s*\(.*?\)/g, "").replace(/\*/g, "").trim();
          } else {
             // Just remove asterisks for other awards, but preserve things like "(gold pin)"
             rawAward = rawAward.replace(/\*/g, "").trim();
          }

          parsedData.push({
            scoutName: rawName,
            originalName: String(row[scoutIdx]).trim(),
            type: rawType,
            award: rawAward,
            earnedDate: earnedIdx !== -1 ? row[earnedIdx] : undefined,
          });
        }

        resolve(parsedData);
      } catch (err) {
        reject(err);
      }
    };

    reader.onerror = (err) => reject(err);

    // Read as binary string so XLSX can parse it (handles both csv/xlsx)
    reader.readAsBinaryString(file);
  });
}
