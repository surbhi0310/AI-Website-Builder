import { initializeApp } from "firebase/app";
import {
  getAuth,
  GoogleAuthProvider,
  setPersistence,
  browserSessionPersistence,
} from "firebase/auth";

const firebaseConfig = {
  apiKey: "AIzaSyBaw2oJ9JWuRrukReZnid1YcMl9nWpKL48",
  authDomain: "genweb-a.firebaseapp.com",
  projectId: "genweb-a",
  storageBucket: "genweb-a.firebasestorage.app",
  messagingSenderId: "134551201138",
  appId: "1:134551201138:web:d271e4a33b56d65af369eb",
  measurementId: "G-6ZLWW1F6QR"
};

const app = initializeApp(firebaseConfig);

export const auth = getAuth(app);
export const provider = new GoogleAuthProvider();

// Don't keep Firebase login permanently after browser restart
setPersistence(auth, browserSessionPersistence);
