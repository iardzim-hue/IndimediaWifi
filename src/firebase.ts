import { initializeApp, getApps, type FirebaseApp } from "firebase/app";
import { getAuth, type Auth } from "firebase/auth";
import { getFirestore, type Firestore } from "firebase/firestore";

// Konfigurasi web Firebase untuk project IndimediaWifi.
// Firebase Web API key memang digunakan di client; keamanan data diatur melalui Authentication dan Firestore Rules.
const firebaseConfig = {
  apiKey: "AIzaSyDFKfuOYeUXJOOprBEZEdneYuxmMN-Q9Cg",
  authDomain: "indimediawifi.firebaseapp.com",
  projectId: "indimediawifi",
  storageBucket: "indimediawifi.firebasestorage.app",
  messagingSenderId: "1058690502393",
  appId: "1:1058690502393:web:afb77f76aeadc998ea5c28",
  measurementId: "G-E7H8DP074K",
};

const hasFirebaseConfig = Boolean(firebaseConfig.apiKey && firebaseConfig.authDomain && firebaseConfig.projectId && firebaseConfig.appId);
const app: FirebaseApp = getApps().length ? getApps()[0] : initializeApp(firebaseConfig);
const auth: Auth = getAuth(app);
const db: Firestore = getFirestore(app);

export { app, auth, db, hasFirebaseConfig };
