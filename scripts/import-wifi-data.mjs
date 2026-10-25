/**
 * One-time migration for IndimediaWifi only.
 *
 * Usage:
 *   GOOGLE_APPLICATION_CREDENTIALS=/secure/path/service-account.json \
 *   FIREBASE_PROJECT_ID=your-project-id \
 *   node scripts/import-wifi-data.mjs /path/to/indimediawifi-migration-data.json
 *
 * Default is a dry run. Add --apply to write to Firestore.
 * Add --force only if you intentionally want to overwrite the same legacy IDs.
 */
import fs from "node:fs";
import { applicationDefault, initializeApp } from "firebase-admin/app";
import { FieldValue, getFirestore } from "firebase-admin/firestore";

const inputPath = process.argv.find((arg) => arg.endsWith(".json"));
const apply = process.argv.includes("--apply");
const force = process.argv.includes("--force");
const projectId = process.env.FIREBASE_PROJECT_ID;

if (!inputPath || !projectId) {
  console.error("Usage: FIREBASE_PROJECT_ID=... node scripts/import-wifi-data.mjs <data.json> [--apply] [--force]");
  process.exit(1);
}
if (!process.env.GOOGLE_APPLICATION_CREDENTIALS) {
  console.error("Set GOOGLE_APPLICATION_CREDENTIALS to a service-account JSON path. Never commit that file.");
  process.exit(1);
}

const data = JSON.parse(fs.readFileSync(inputPath, "utf8"));
if (!Array.isArray(data.customers) || !Array.isArray(data.bills)) {
  throw new Error("Input JSON must contain customers[] and bills[]");
}
initializeApp({ credential: applicationDefault(), projectId });
const db = getFirestore();
const markerRef = db.collection("wifi_settings").doc("migration_cat_wifi_2026");
const marker = await markerRef.get();
if (marker.exists && !force) {
  throw new Error("Migration marker already exists. No writes made. Use --force only after reviewing existing Firestore data.");
}

const customers = data.customers;
const bills = data.bills;
console.log(`Project: ${projectId}`);
console.log(`Target collections: wifi_customers (${customers.length}), wifi_bills (${bills.length}), wifi_settings (1 marker)`);
console.log(`Mode: ${apply ? "WRITE" : "DRY RUN"}`);
if (!apply) {
  console.log("Dry run only; no Firestore documents were changed. Add --apply to import.");
  process.exit(0);
}

const writes = [];
for (const customer of customers) {
  const id = `legacy-${String(customer.legacyNo).padStart(3, "0")}`;
  writes.push({
    ref: db.collection("wifi_customers").doc(id),
    data: {
      ...customer,
      source: "cat_wifi_2026_migration",
      migratedAt: FieldValue.serverTimestamp(),
    },
  });
}
for (const bill of bills) {
  const customerId = `legacy-${String(bill.legacyNo).padStart(3, "0")}`;
  const id = `${customerId}_${bill.billingMonth}`;
  writes.push({
    ref: db.collection("wifi_bills").doc(id),
    data: {
      ...bill,
      customerId,
      source: "cat_wifi_2026_migration",
      migratedAt: FieldValue.serverTimestamp(),
    },
  });
}

// Firestore batches allow at most 500 writes. Keep each batch below that limit.
for (let i = 0; i < writes.length; i += 400) {
  const batch = db.batch();
  for (const write of writes.slice(i, i + 400)) {
    batch.set(write.ref, write.data, { merge: force });
  }
  await batch.commit();
}
await markerRef.set({
  sourceFile: data.sourceFile ?? "unknown",
  sourceSheet: data.sourceSheet ?? "unknown",
  year: data.year ?? 2026,
  customerCount: customers.length,
  billCount: bills.length,
  completedAt: FieldValue.serverTimestamp(),
  collectionNames: ["wifi_customers", "wifi_bills", "wifi_settings"],
});
console.log(`Import complete: ${customers.length} customers and ${bills.length} monthly records.`);
