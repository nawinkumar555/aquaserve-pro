import { initializeApp } from "firebase/app";
import { getFirestore } from "firebase/firestore";
import { getAuth } from "firebase/auth"; 
import { getStorage } from "firebase/storage"; 

const firebaseConfig = {
  apiKey: "AIzaSyCbIDzq3G68Wc-MChEzAH5crlqqXCYR1pU",
  authDomain: "aquaserve-pro-system.firebaseapp.com",
  projectId: "aquaserve-pro-system",
  storageBucket: "aquaserve-pro-system.firebasestorage.app",
  messagingSenderId: "964146409816",
  appId: "1:964146409816:web:08a17529576f3181c4752d",
  measurementId: "G-M225FPM01X"
};

const app = initializeApp(firebaseConfig);

export const db = getFirestore(app);
export const auth = getAuth(app);
export const storage = getStorage(app);