import { initializeApp, getApps, getApp } from 'firebase/app';
import { getFirestore } from 'firebase/firestore';
import { getAnalytics, isSupported } from 'firebase/analytics';

const firebaseConfig = {
  apiKey: "AIzaSyAZcq0tTEcIsi1R0izpBXTVHkG5YDiDZDc",
  authDomain: "elevate-qc-app.firebaseapp.com",
  projectId: "elevate-qc-app",
  storageBucket: "elevate-qc-app.firebasestorage.app",
  messagingSenderId: "429836461404",
  appId: "1:429836461404:web:4ec87dd6267e21a30f8d99",
  measurementId: "G-VNNG6G8M4E"
};

// Initialize Firebase
const app = !getApps().length ? initializeApp(firebaseConfig) : getApp();
const db = getFirestore(app);

// Analytics (hanya di browser)
let analytics: any = null;
if (typeof window !== 'undefined') {
  isSupported().then(yes => {
    if (yes) analytics = getAnalytics(app);
  });
}

export { app, db, analytics };