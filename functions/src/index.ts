import { setGlobalOptions } from "firebase-functions";
import { onRequest } from "firebase-functions/https";
import * as logger from "firebase-functions/logger";
import * as admin from "firebase-admin";
import cors from "cors";

// Initialize Firebase Admin
admin.initializeApp();

// Set global options for cost control
setGlobalOptions({ maxInstances: 10 });

// Initialize CORS middleware
const corsHandler = cors({ origin: true });

// Get all users from Firestore
export const getFirestoreData = onRequest(async (request, response) => {
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
        message: "Successfully retrieved all users",
      });
    } catch (error) {
      logger.error("Error getting users:", error);
      response.status(500).send({ error: "Internal server error" });
    }
  });
});
