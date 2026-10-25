import { cert, getApps, initializeApp } from "firebase-admin/app";
import { getFirestore } from "firebase-admin/firestore";

const headers = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "Content-Type, Authorization",
  "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
  "Content-Type": "application/json; charset=utf-8",
};

function reply(statusCode, payload) {
  return { statusCode, headers, body: JSON.stringify(payload) };
}

function adminDb() {
  if (getApps().length) return getFirestore();
  const raw = process.env.FIREBASE_SERVICE_ACCOUNT;
  if (!raw) throw new Error("FIREBASE_SERVICE_ACCOUNT belum disetel di Netlify.");
  const serviceAccount = JSON.parse(raw);
  initializeApp({ credential: cert(serviceAccount), projectId: serviceAccount.project_id || "indimediawifi" });
  return getFirestore();
}

function normalizeId(value) {
  return String(value || "").trim().toLocaleLowerCase("id-ID");
}

export async function handler(event) {
  if (event.httpMethod === "OPTIONS") return { statusCode: 204, headers, body: "" };
  if (!["GET", "POST"].includes(event.httpMethod)) return reply(405, { error: "Method tidak didukung." });

  try {
    const db = adminDb();

    if (event.httpMethod === "GET") {
      const action = String(event.queryStringParameters?.action || "customer");
      if (action === "status") {
        const snap = await db.collection("settings").doc("server").get();
        const config = snap.exists ? snap.data() : {};
        const target = String(config?.target || "").trim();
        if (!target) return reply(200, { status: "unknown", label: "Belum dikonfigurasi", target: "" });
        let url = target;
        if (!/^https?:\/\//i.test(url)) url = "https://" + url;
        const started = Date.now();
        try {
          const response = await fetch(url, { method: "GET", redirect: "follow", signal: AbortSignal.timeout(5000) });
          const status = response.status;
          return reply(200, {
            status: status >= 200 && status < 500 ? "online" : "offline",
            label: status >= 200 && status < 500 ? "Dapat dijangkau" : "Server bermasalah",
            target,
            httpStatus: status,
            checkedAt: new Date().toISOString(),
            latencyMs: Date.now() - started,
          });
        } catch {
          return reply(200, { status: "offline", label: "Tidak dapat dijangkau", target, checkedAt: new Date().toISOString(), latencyMs: Date.now() - started });
        }
      }

      const id = normalizeId(event.queryStringParameters?.id);
      if (!id) return reply(400, { error: "ID pelanggan wajib diisi." });
      const customersSnap = await db.collection("customers").get();
      const customerDoc = customersSnap.docs.find(d => normalizeId(d.data().customerId || d.data().loginId || "") === id);
      if (!customerDoc) return reply(404, { error: "ID pelanggan tidak ditemukan. Minta admin memeriksa ID login." });
      const customer = customerDoc.data();
      const legacyNo = Number(customer.legacyNo ?? customerDoc.id);
      const billsSnap = await db.collection("bills").where("legacyNo", "==", legacyNo).get();
      const bills = billsSnap.docs.map(d => {
        const b = d.data();
        const amount = Number(b.amount ?? b.amountRaw ?? 0) || 0;
        const discount = Math.max(0, Number(b.discount) || 0);
        return {
          id: d.id,
          billingMonth: b.billingMonth || "",
          year: Number(b.year || 2026),
          amount,
          discount,
          total: Math.max(0, amount - discount),
          status: String(b.status || "unpaid").toLowerCase(),
          paymentNote: b.paymentNote || "",
          paidAt: b.paidAt || "",
        };
      }).sort((a,b) => b.year - a.year || String(b.billingMonth).localeCompare(String(a.billingMonth), "id"));
      const unpaid = bills.filter(b => !["paid", "lunas"].includes(b.status));
      return reply(200, {
        customer: {
          legacyNo,
          name: customer.name || "",
          customerId: customer.customerId || customer.loginId || "",
          status: customer.status || "active",
          phone: customer.phone || "",
          packageName: customer.packageName || "",
          basePrice: Number(customer.basePrice || 0),
          dueDay: Number(customer.dueDay || 10),
        },
        unpaidBills: unpaid,
        totalOutstanding: unpaid.reduce((sum, b) => sum + b.total, 0),
        serverStatus: null,
      });
    }

    const authHeader = event.headers?.authorization || event.headers?.Authorization || "";
    const token = authHeader.replace(/^Bearer\s+/i, "");
    if (!token) return reply(401, { error: "Token admin wajib disertakan." });
    const { getAuth } = await import("firebase-admin/auth");
    const decoded = await getAuth().verifyIdToken(token);
    if (!decoded.email || !decoded.email_verified && !decoded.firebase?.sign_in_provider) {
      return reply(403, { error: "Akun admin tidak valid." });
    }
    const body = JSON.parse(event.body || "{}");
    const target = String(body.target || "").trim();
    if (target && !/^(https?:\/\/)?([a-z0-9-]+\.)+[a-z]{2,}(:\d{1,5})?(\/[^\s]*)?$|^(\d{1,3}\.){3}\d{1,3}(:\d{1,5})?$|^([a-f0-9:]+)$/i.test(target)) {
      return reply(400, { error: "Target harus berupa domain atau alamat IP yang valid." });
    }
    await db.collection("settings").doc("server").set({ target, updatedAt: new Date().toISOString(), updatedBy: decoded.email }, { merge: true });
    return reply(200, { ok: true, target });
  } catch (error) {
    console.error("customer-portal error", error);
    return reply(500, { error: error instanceof Error ? error.message : "Kesalahan server." });
  }
}
