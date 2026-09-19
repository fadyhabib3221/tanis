// Firebase configuration
import { initializeApp, getApps } from "firebase/app";
import { getAuth, setPersistence, browserSessionPersistence } from "firebase/auth";
import { getFirestore } from "firebase/firestore";
import { getStorage } from "firebase/storage";

// Reads from NEXT_PUBLIC_FIREBASE_* env vars (see .env.local.example and
// README.md's "Configure Firebase" step) when they're set, falling back to
// this repo's own original project otherwise — so nothing changes for an
// existing deployment that hasn't set these, but copying this codebase for
// a different client now actually works the way the README already told
// people it did: set .env.local and get your own project, instead of
// silently staying pointed at this hardcoded one because the code never
// read those variables in the first place. scripts/seedAirlines.js follows
// the same fallback pattern for the same reason.
const firebaseConfig = {
  apiKey: process.env.NEXT_PUBLIC_FIREBASE_API_KEY || "AIzaSyCf5AdJ7pWy1DYsXYR4_MbyN0q4D-ja-aA",
  authDomain: process.env.NEXT_PUBLIC_FIREBASE_AUTH_DOMAIN || "tanis-9e15b.firebaseapp.com",
  projectId: process.env.NEXT_PUBLIC_FIREBASE_PROJECT_ID || "tanis-9e15b",
  storageBucket: process.env.NEXT_PUBLIC_FIREBASE_STORAGE_BUCKET || "tanis-9e15b.firebasestorage.app",
  messagingSenderId: process.env.NEXT_PUBLIC_FIREBASE_MESSAGING_SENDER_ID || "401226831173",
  appId: process.env.NEXT_PUBLIC_FIREBASE_APP_ID || "1:401226831173:web:e0e899140594974fd81a2c",
  measurementId: process.env.NEXT_PUBLIC_FIREBASE_MEASUREMENT_ID || "G-9NH56VSWQZ",
};

// Initialize Firebase only once
const app = getApps().length === 0 ? initializeApp(firebaseConfig) : getApps()[0];

export const auth = getAuth(app);
// Session-only persistence: the signed-in state lives only in this tab.
// Closing the tab (or the browser) signs the user out automatically —
// reopening the app always lands back on the login screen. Guarded for
// the server-render pass, where sessionStorage doesn't exist.
if (typeof window !== "undefined") {
  setPersistence(auth, browserSessionPersistence).catch(() => {});
}
export const db = getFirestore(app);
export const storage = getStorage(app);
// Exported so other modules (e.g. lib/auth.js's secondary-app trick for
// creating employees without signing the admin out) always initialize
// against the SAME project as everything else, instead of keeping their
// own separate copy that can silently go stale after a project switch.
export { firebaseConfig };

export default app;
