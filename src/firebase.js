import { initializeApp } from 'firebase/app';
import { getDatabase, ref, onValue, set, remove, get, runTransaction } from 'firebase/database';

// Only the database URL is needed (see .env.example)
const url = import.meta.env.VITE_FIREBASE_DATABASE_URL;
export const configured = Boolean(url);
export const db = configured ? getDatabase(initializeApp({ databaseURL: url })) : null;
export { ref, onValue, set, remove, get, runTransaction };
