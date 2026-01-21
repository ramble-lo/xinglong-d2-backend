/**
 * Import function triggers from their respective submodules:
 *
 * import {onCall} from "firebase-functions/v2/https";
 * import {onDocumentWritten} from "firebase-functions/v2/firestore";
 *
 * See a full list of supported triggers at https://firebase.google.com/docs/functions
 */

import { setGlobalOptions } from "firebase-functions";
import { onRequest } from "firebase-functions/https";
import { onCall, HttpsError } from "firebase-functions/v2/https";
import * as logger from "firebase-functions/logger";
import * as admin from "firebase-admin";
import cors from "cors";
import {
  parseExcelFromBase64,
  extractRowData,
  isValidRow,
  buildRegistrantDoc,
  buildRegistrationDoc,
  ProcessResult,
} from "./processExcel";

// Initialize Firebase Admin
admin.initializeApp();

// Start writing functions
// https://firebase.google.com/docs/functions/typescript

// For cost control, you can set the maximum number of containers that can be
// running at the same time. This helps mitigate the impact of unexpected
// traffic spikes by instead downgrading performance. This limit is a
// per-function limit. You can override the limit for each function using the
// `maxInstances` option in the function's options, e.g.
// `onRequest({ maxInstances: 5 }, (req, res) => { ... })`.
// NOTE: setGlobalOptions does not apply to functions using the v1 API. V1
// functions should each use functions.runWith({ maxInstances: 10 }) instead.
// In the v1 API, each function can only serve one request per container, so
// this will be the maximum concurrent request count.
setGlobalOptions({ maxInstances: 10 });

export const helloWorld = onRequest((request, response) => {
  logger.info("Hello logs!", { structuredData: true });
  response.send("Hello from Firebase!");
});

// Example: Get data from Firestore
export const getFirestoreData = onRequest(async (request, response) => {
  try {
    const db = admin.firestore();

    // 方法 1: 取得單一文件 (document)
    // 將 'your-collection' 和 'document-id' 替換成你的實際 collection 和 document ID
    const docRef = db.collection("your-collection").doc("document-id");
    const doc = await docRef.get();

    if (!doc.exists) {
      response.status(404).send({ error: "Document not found" });
      return;
    }

    // 方法 2: 取得整個 collection 的所有文件
    const collectionRef = db.collection("your-collection");
    const snapshot = await collectionRef.get();
    const allDocs = snapshot.docs.map((doc) => ({
      id: doc.id,
      ...doc.data(),
    }));

    // 方法 3: 使用查詢條件
    const querySnapshot = await db
      .collection("your-collection")
      .where("fieldName", "==", "value") // 替換成你的欄位和值
      .limit(10)
      .get();

    const queryResults = querySnapshot.docs.map((doc) => ({
      id: doc.id,
      ...doc.data(),
    }));

    // 回傳資料
    response.send({
      singleDocument: doc.data(),
      allDocuments: allDocs,
      queryResults: queryResults,
    });
  } catch (error) {
    logger.error("Error getting Firestore data:", error);
    response.status(500).send({ error: "Internal server error" });
  }
});

// Initialize CORS middleware
const corsHandler = cors({ origin: true });

// Get all users from Firestore
export const getAllUsers = onRequest(async (request, response) => {
  // Handle CORS
  return corsHandler(request, response, async () => {
    try {
      const db = admin.firestore();

      // 取得 users collection 的所有文件
      const usersSnapshot = await db.collection("users").get();

      // 將所有文件轉換成陣列
      const users = usersSnapshot.docs.map((doc) => ({
        id: doc.id,
        ...doc.data(),
      }));

      logger.info(`Retrieved ${users.length} users from Firestore`);

      // 回傳資料
      response.send({
        count: users.length,
        users: users,
      });
    } catch (error) {
      logger.error("Error getting users:", error);
      response.status(500).send({ error: "Internal server error" });
    }
  });
});

// ============================================
// Excel Upload Processing Cloud Function
// ============================================

interface ProcessExcelRequest {
  fileData: string; // Base64 encoded Excel file
  fileName: string;
}

/**
 * Cloud Function to process Excel file uploads
 * Frontend sends base64 encoded file, this function parses and writes to Firestore
 */
export const processExcelUpload = onCall<ProcessExcelRequest>(
  {
    // Memory and timeout settings for processing large files
    memory: "256MiB",
    timeoutSeconds: 120,
    // Enforce App Check for security
    enforceAppCheck: true,
  },
  async (request): Promise<ProcessResult> => {
    // Verify authentication
    if (!request.auth) {
      throw new HttpsError("unauthenticated", "必須登入才能上傳檔案");
    }

    const { fileData, fileName } = request.data;

    // Validate input
    if (!fileData) {
      throw new HttpsError("invalid-argument", "缺少檔案資料");
    }

    if (!fileName) {
      throw new HttpsError("invalid-argument", "缺少檔案名稱");
    }

    logger.info(`Processing file: ${fileName} for user: ${request.auth.uid}`);

    try {
      const db = admin.firestore();

      // Parse Excel from base64
      const jsonData = parseExcelFromBase64(fileData);
      logger.info(`Parsed ${jsonData.length} rows from Excel`);

      let processedCount = 0;
      let skippedCount = 0;
      let duplicateCount = 0;

      // Process each row
      for (const row of jsonData) {
        try {
          const rowData = extractRowData(row);

          // Skip invalid rows
          if (!isValidRow(rowData)) {
            skippedCount++;
            continue;
          }

          // Check if registrant exists
          const registrantsQuery = await db
            .collection("registrants")
            .where("name", "==", rowData.name)
            .where("phone", "==", rowData.phone)
            .get();

          let registrantId: string;

          if (registrantsQuery.empty) {
            // Create new registrant
            const registrantDoc = buildRegistrantDoc(rowData);
            const newRegistrantRef = await db
              .collection("registrants")
              .add(registrantDoc);
            registrantId = newRegistrantRef.id;
          } else {
            registrantId = registrantsQuery.docs[0].id;
          }

          // Check for duplicate registration
          const registrationQuery = await db
            .collection("registration_history")
            .where("surveycake_hash", "==", rowData.surveycakeHash)
            .get();

          if (!registrationQuery.empty) {
            duplicateCount++;
            continue;
          }

          // Create registration
          const registrationDoc = buildRegistrationDoc(rowData, registrantId);
          await db.collection("registration_history").add(registrationDoc);

          processedCount++;
        } catch (error) {
          logger.error("Error processing row:", error);
          skippedCount++;
        }
      }

      // Build result message
      const messageParts = [`成功處理 ${processedCount} 筆資料`];
      if (skippedCount > 0) {
        messageParts.push(`跳過 ${skippedCount} 筆無效資料`);
      }
      if (duplicateCount > 0) {
        messageParts.push(`跳過 ${duplicateCount} 筆重複資料`);
      }

      logger.info(`Processing complete: ${messageParts.join(", ")}`);

      return {
        success: true,
        message: messageParts.join("，"),
        processedCount,
        skippedCount,
        duplicateCount,
      };
    } catch (error) {
      logger.error("Error processing Excel file:", error);
      throw new HttpsError(
        "internal",
        "處理 Excel 檔案時發生錯誤，請檢查檔案格式"
      );
    }
  }
);
