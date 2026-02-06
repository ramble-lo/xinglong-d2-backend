import { setGlobalOptions } from "firebase-functions";
import { onRequest } from "firebase-functions/https";
import * as logger from "firebase-functions/logger";
import * as admin from "firebase-admin";
import cors from "cors";
import * as XLSX from "xlsx";

// Initialize Firebase Admin
admin.initializeApp();

// Set global options for cost control
setGlobalOptions({ maxInstances: 10 });

// Initialize CORS middleware
const corsHandler = cors({ origin: true });

// Types for Excel processing
interface ExcelRowData {
  [key: string]: string | number | undefined;
}

type ResidentStatusType =
  | "xinglongd2"
  | "wenshan"
  | "otherTaipeiSocialHousing"
  | "other";

const ResidentStatusEnum: Record<string, ResidentStatusType> = {
  是: "xinglongd2",
  "否，我是文山區鄰近居民": "wenshan",
  "否，我是其他臺北市社會住宅的住戶": "otherTaipeiSocialHousing",
  以上皆非: "other",
};

interface Registrant {
  name: string;
  email: string;
  phone: string;
  gender?: string;
  age?: string;
  line_id?: string;
  residient_type?: ResidentStatusType;
  created_at: admin.firestore.Timestamp;
  updated_at: admin.firestore.Timestamp;
}

interface Registration {
  name: string;
  email: string;
  surveycake_hash: string;
  phone: string;
  gender?: string | null;
  residient_type?: ResidentStatusType;
  registrant_id: string;
  activity_name: string;
  age?: string | null;
  children_count?: string | null;
  sports_experience?: string | null;
  injury_history?: string | null;
  info_source?: string | null;
  suggestions?: string | null;
  line_id?: string | null;
  housing_location?: string | null;
  submit_time: admin.firestore.Timestamp;
  created_at: admin.firestore.Timestamp;
}

type UserRole = "admin" | "general" | "guest";
type UserTeam = "admin" | "platform" | "finance" | "venue" | "supplies";

interface UserInfo {
  id?: string;
  name: string;
  email: string;
  community_code: string;
  created_at: admin.firestore.Timestamp;
  role: UserRole;
  team?: UserTeam;
  totp_secret?: string;
}

const verifyAuthToken = async (
  request: import("firebase-functions/https").Request,
): Promise<admin.auth.DecodedIdToken> => {
  const authHeader = request.headers.authorization;

  if (!authHeader || !authHeader.startsWith("Bearer ")) {
    throw new Error("Missing or invalid Authorization header");
  }

  const idToken = authHeader.split("Bearer ")[1];

  try {
    const decodedToken = await admin.auth().verifyIdToken(idToken);
    return decodedToken;
  } catch (error) {
    logger.error("Token verification failed:", error);
    throw new Error("Invalid or expired token");
  }
};

// Get all users from Firestore (requires authentication)
export const getFirestoreData = onRequest(async (request, response) => {
  return corsHandler(request, response, async () => {
    try {
      // Verify authentication token
      let decodedToken: admin.auth.DecodedIdToken;
      try {
        decodedToken = await verifyAuthToken(request);
      } catch (authError) {
        response.status(401).send({
          success: false,
          message:
            authError instanceof Error
              ? authError.message
              : "Authentication failed",
        });
        return;
      }

      const db = admin.firestore();

      // 取得 users collection 的所有文件
      const usersSnapshot = await db.collection("users").get();

      // 將所有文件轉換成陣列
      const users = usersSnapshot.docs.map((doc) => ({
        id: doc.id,
        ...doc.data(),
      }));

      logger.info(
        `Retrieved ${users.length} users from Firestore by user: ${decodedToken.email}`,
      );

      // 回傳資料
      response.send({
        success: true,
        count: users.length,
        users: users,
        message: "Successfully retrieved all users",
      });
    } catch (error) {
      logger.error("Error getting users:", error);
      response.status(500).send({
        success: false,
        message: "Internal server error",
      });
    }
  });
});

// Upload and process Excel application form (requires authentication)
export const uploadApplicationForm = onRequest(async (request, response) => {
  return corsHandler(request, response, async () => {
    // Only allow POST requests
    if (request.method !== "POST") {
      response.status(405).send({ error: "Method not allowed" });
      return;
    }

    // Verify authentication token
    let decodedToken: admin.auth.DecodedIdToken;
    try {
      decodedToken = await verifyAuthToken(request);
    } catch (authError) {
      response.status(401).send({
        success: false,
        message:
          authError instanceof Error
            ? authError.message
            : "Authentication failed",
      });
      return;
    }

    try {
      const { fileBase64, fileName } = request.body;

      if (!fileBase64) {
        response.status(400).send({
          success: false,
          message: "Missing fileBase64 in request body",
          processedCount: 0,
        });
        return;
      }

      logger.info(
        `Processing Excel file: ${fileName || "unknown"} by user: ${decodedToken.email}`,
      );

      // Decode Base64 to buffer
      const buffer = Buffer.from(fileBase64, "base64");

      // Read Excel file
      const workbook = XLSX.read(buffer, { type: "buffer" });
      const sheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[sheetName];
      const jsonData = XLSX.utils.sheet_to_json(worksheet) as ExcelRowData[];

      const db = admin.firestore();
      let processedCount = 0;
      let skippedCount = 0;
      let duplicateCount = 0;

      for (const row of jsonData) {
        try {
          // 提取欄位資料，使用中文欄位名稱
          const activityName = String(row["以下活動請擇一"] || "");
          const name = String(row["姓名"] || "");
          const email = String(row["電子郵件"] || "");
          const surveycakeHash = String(row["Hash"] || "");
          const phone = String(row["聯絡電話"] || "");
          const gender = String(row["性別"] || "");
          const age = String(row["參與者年齡"] || "");
          const lineId = String(row["Line ID（意者可留）"] || "");
          const childrenCount = String(row["小孩人數"] || "");
          const residentStatusKey = String(
            row["請問您是興隆社宅2區的住戶嗎？"] || "",
          );
          const residentStatus: ResidentStatusType =
            ResidentStatusEnum[residentStatusKey] || "other";
          const housingLocation = String(
            row["您是來自哪個臺北市社會住宅？"] || "",
          );
          const sportsExperience = String(row["運動經歷幾年？"] || "");
          const injuryHistory = String(
            row["是否有受傷病史？（沒有請填無）"] || "",
          );
          const infoSource = String(
            row["請問您從何處得知本次活動資訊？"] || "",
          );
          const suggestions = String(
            row[
              "針對活動，有什麼建議或想和主辦單位說的話嗎？請在這裡留言喔～謝謝您！"
            ] || "",
          );
          const submitTimeStr = String(row["填答時間"] || "");

          // 跳過空白資料行
          if (!name || !email || !phone || !activityName) {
            skippedCount++;
            continue;
          }

          // 處理提交時間
          let submitTime: admin.firestore.Timestamp;
          if (submitTimeStr) {
            const parsedDate = new Date(submitTimeStr);
            if (!isNaN(parsedDate.getTime())) {
              submitTime = admin.firestore.Timestamp.fromDate(parsedDate);
            } else {
              submitTime = admin.firestore.Timestamp.fromDate(new Date());
            }
          } else {
            submitTime = admin.firestore.Timestamp.fromDate(new Date());
          }

          // 檢查報名者是否已存在
          const registrantsQuery = await db
            .collection("registrants")
            .where("name", "==", name)
            .where("phone", "==", phone)
            .get();

          let registrantId: string;

          if (registrantsQuery.empty) {
            const registrant: Registrant = {
              name,
              email,
              phone,
              gender,
              age,
              line_id: lineId,
              residient_type: residentStatus,
              created_at: admin.firestore.Timestamp.fromDate(new Date()),
              updated_at: admin.firestore.Timestamp.fromDate(new Date()),
            };
            // 建立新的報名者
            const newRegistrantRef = await db
              .collection("registrants")
              .add(registrant);
            registrantId = newRegistrantRef.id;
          } else {
            registrantId = registrantsQuery.docs[0].id;
          }

          // 檢查是否已有相同的報名歷史記錄
          const registrationQuery = await db
            .collection("registration_history")
            .where("surveycake_hash", "==", surveycakeHash)
            .get();

          if (!registrationQuery.empty) {
            duplicateCount++;
            continue; // 如果已存在，則跳過此筆資料
          }

          const registration: Registration = {
            registrant_id: registrantId,
            surveycake_hash: surveycakeHash,
            activity_name: activityName,
            name: name,
            residient_type: residentStatus,
            email: email,
            phone: phone,
            gender: gender || null,
            line_id: lineId || null,
            housing_location: housingLocation || null,
            age: age || null,
            children_count: childrenCount || null,
            sports_experience: sportsExperience || null,
            injury_history: injuryHistory || null,
            info_source: infoSource || null,
            suggestions: suggestions || null,
            submit_time: submitTime,
            created_at: admin.firestore.Timestamp.fromDate(new Date()),
          };

          // 新增報名歷史記錄
          await db.collection("registration_history").add(registration);

          processedCount++;
        } catch (rowError) {
          logger.error("處理單行資料失敗:", rowError);
          skippedCount++;
        }
      }

      const messageParts = [`成功處理 ${processedCount} 筆資料`];
      if (skippedCount > 0) {
        messageParts.push(`跳過 ${skippedCount} 筆無效資料`);
      }
      if (duplicateCount > 0) {
        messageParts.push(`跳過 ${duplicateCount} 筆重複資料`);
      }

      logger.info(
        `Excel processing complete: processed=${processedCount}, skipped=${skippedCount}, duplicates=${duplicateCount}`,
      );

      response.send({
        success: true,
        message: messageParts.join("，"),
        processedCount,
        skippedCount,
        duplicateCount,
      });
    } catch (error) {
      logger.error("處理 Excel 檔案失敗:", error);
      response.status(500).send({
        success: false,
        message: "處理 Excel 檔案時發生錯誤，請檢查檔案格式",
        processedCount: 0,
      });
    }
  });
});

// Get user info by email (requires authentication)
export const getUserInfo = onRequest(async (request, response) => {
  return corsHandler(request, response, async () => {
    try {
      // Verify authentication token
      let decodedToken: admin.auth.DecodedIdToken;
      try {
        decodedToken = await verifyAuthToken(request);
      } catch (authError) {
        response.status(401).send({
          success: false,
          message:
            authError instanceof Error
              ? authError.message
              : "Authentication failed",
        });
        return;
      }

      const email = request.query.email as string;

      if (!email) {
        response.status(400).send({
          success: false,
          message: "Missing email parameter",
        });
        return;
      }

      // Optional: Only allow users to query their own info
      // Uncomment the following if you want to restrict access
      // if (decodedToken.email !== email) {
      //   response.status(403).send({
      //     success: false,
      //     message: "You can only access your own user info",
      //   });
      //   return;
      // }

      const db = admin.firestore();

      // 查詢 users collection 中符合 email 的使用者
      const usersQuery = await db
        .collection("users")
        .where("email", "==", email)
        .get();

      if (usersQuery.empty) {
        response.status(404).send({
          success: false,
          message: `User with email ${email} not found`,
        });
        return;
      }

      // 取得第一個符合的使用者資料
      const userDoc = usersQuery.docs[0];
      const userInfo: UserInfo = {
        id: userDoc.id,
        ...userDoc.data(),
      } as UserInfo;

      logger.info(
        `Retrieved user info for email: ${email} by user: ${decodedToken.email}`,
      );

      response.send({
        success: true,
        data: userInfo,
        message: "Successfully retrieved user info",
      });
    } catch (error) {
      logger.error("Error getting user info:", error);
      response.status(500).send({
        success: false,
        message: "Internal server error",
      });
    }
  });
});
