import * as XLSX from "xlsx";
import { Timestamp } from "firebase-admin/firestore";

// Registrant status types
export const ResidentStatusEnum: { [key: string]: string } = {
  是: "resident",
  否: "non_resident",
};

export type ResidentStatusType = "resident" | "non_resident" | "other";

export interface Registrant {
  name: string;
  email: string;
  phone: string;
  gender: string;
  age: string;
  line_id: string;
  residient_type: ResidentStatusType;
  created_at: Timestamp;
  updated_at: Timestamp;
}

export interface Registration {
  registrant_id: string;
  surveycake_hash: string;
  activity_name: string;
  name: string;
  residient_type: ResidentStatusType;
  email: string;
  phone: string;
  gender: string | null;
  line_id: string | null;
  housing_location: string | null;
  age: string | null;
  children_count: string | null;
  sports_experience: string | null;
  injury_history: string | null;
  info_source: string | null;
  suggestions: string | null;
  submit_time: Timestamp;
  created_at: Timestamp;
}

interface ExcelRowData {
  [key: string]: string | number | undefined;
}

export interface ProcessResult {
  success: boolean;
  message: string;
  processedCount: number;
  skippedCount: number;
  duplicateCount: number;
}

/**
 * Parse Excel file from base64 string and return structured data
 */
export function parseExcelFromBase64(base64Data: string): ExcelRowData[] {
  // Decode base64 to buffer
  const buffer = Buffer.from(base64Data, "base64");

  // Read workbook from buffer
  const workbook = XLSX.read(buffer, { type: "buffer" });
  const sheetName = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[sheetName];

  // Convert to JSON
  return XLSX.utils.sheet_to_json(worksheet) as ExcelRowData[];
}

/**
 * Extract registrant data from Excel row
 */
export function extractRowData(row: ExcelRowData) {
  const activityName = String(row["以下活動請擇一"] || "");
  const name = String(row["姓名"] || "");
  const email = String(row["電子郵件"] || "");
  const surveycakeHash = String(row["Hash"] || "");
  const phone = String(row["聯絡電話"] || "");
  const gender = String(row["性別"] || "");
  const age = String(row["參與者年齡"] || "");
  const lineId = String(row["Line ID（意者可留）"] || "");
  const childrenCount = String(row["小孩人數"] || "");
  const residentStatus: ResidentStatusType =
    (ResidentStatusEnum[
      String(row["請問您是興隆社宅2區的住戶嗎？"])
    ] as ResidentStatusType) || "other";
  const housingLocation = String(row["您是來自哪個臺北市社會住宅？"] || "");
  const sportsExperience = String(row["運動經歷幾年？"] || "");
  const injuryHistory = String(row["是否有受傷病史？（沒有請填無）"] || "");
  const infoSource = String(row["請問您從何處得知本次活動資訊？"] || "");
  const suggestions = String(
    row[
      "針對活動，有什麼建議或想和主辦單位說的話嗎？請在這裡留言喔～謝謝您！"
    ] || ""
  );
  const submitTimeStr = String(row["填答時間"] || "");

  // Parse submit time
  let submitTime: Timestamp;
  if (submitTimeStr) {
    const parsedDate = new Date(submitTimeStr);
    if (!isNaN(parsedDate.getTime())) {
      submitTime = Timestamp.fromDate(parsedDate);
    } else {
      submitTime = Timestamp.fromDate(new Date());
    }
  } else {
    submitTime = Timestamp.fromDate(new Date());
  }

  return {
    activityName,
    name,
    email,
    surveycakeHash,
    phone,
    gender,
    age,
    lineId,
    childrenCount,
    residentStatus,
    housingLocation,
    sportsExperience,
    injuryHistory,
    infoSource,
    suggestions,
    submitTime,
  };
}

/**
 * Validate that ALL fields are present (non-empty)
 */
export function isValidRow(row: ReturnType<typeof extractRowData>): boolean {
  return !!(
    row.activityName &&
    row.name &&
    row.email &&
    row.surveycakeHash &&
    row.phone &&
    row.gender &&
    row.age &&
    row.lineId &&
    row.childrenCount &&
    row.residentStatus &&
    row.housingLocation &&
    row.sportsExperience &&
    row.injuryHistory &&
    row.infoSource &&
    row.suggestions
  );
}

/**
 * Build registrant document
 */
export function buildRegistrantDoc(
  rowData: ReturnType<typeof extractRowData>
): Registrant {
  return {
    name: rowData.name,
    email: rowData.email,
    phone: rowData.phone,
    gender: rowData.gender,
    age: rowData.age,
    line_id: rowData.lineId,
    residient_type: rowData.residentStatus,
    created_at: Timestamp.fromDate(new Date()),
    updated_at: Timestamp.fromDate(new Date()),
  };
}

/**
 * Build registration document
 */
export function buildRegistrationDoc(
  rowData: ReturnType<typeof extractRowData>,
  registrantId: string
): Registration {
  return {
    registrant_id: registrantId,
    surveycake_hash: rowData.surveycakeHash,
    activity_name: rowData.activityName,
    name: rowData.name,
    residient_type: rowData.residentStatus,
    email: rowData.email,
    phone: rowData.phone,
    gender: rowData.gender || null,
    line_id: rowData.lineId || null,
    housing_location: rowData.housingLocation || null,
    age: rowData.age || null,
    children_count: rowData.childrenCount || null,
    sports_experience: rowData.sportsExperience || null,
    injury_history: rowData.injuryHistory || null,
    info_source: rowData.infoSource || null,
    suggestions: rowData.suggestions || null,
    submit_time: rowData.submitTime,
    created_at: Timestamp.fromDate(new Date()),
  };
}
