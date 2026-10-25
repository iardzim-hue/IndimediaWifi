
import { initializeApp, getApps, getApp } from "firebase/app";
import { getAuth } from "firebase/auth";
import { getFirestore } from "firebase/firestore";
import { getAnalytics, isSupported } from "firebase/analytics";

const firebaseConfig = {
  apiKey: "AIzaSyDFKfuOYeUXJOOprBEZEdneYuxmMN-Q9Cg",
  authDomain: "indimediawifi.firebaseapp.com",
  projectId: "indimediawifi",
  storageBucket: "indimediawifi.firebasestorage.app",
  messagingSenderId: "1058690502393",
  appId: "1:1058690502393:web:afb77f76aeadc998ea5c28",
  measurementId: "G-E7H8DP074K",
};

const app = getApps().length ? getApp() : initializeApp(firebaseConfig);

const auth = getAuth(app);
const db = getFirestore(app);

// Analytics opsional dan hanya berjalan jika didukung browser.
if (typeof window !== "undefined") {
  isSupported()
    .then((supported) => {
      if (supported) getAnalytics(app);
    })
    .catch(() => {});
}

export { app, auth, db };
