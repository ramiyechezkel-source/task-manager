import { initializeApp } from "firebase/app";
import { getFirestore } from "firebase/firestore";
import { getAuth } from "firebase/auth";

const firebaseConfig = {
  apiKey: "AIzaSyBaPdwSAQSvM0wSrIlcJM1rMVt2TPoKFds",
  authDomain: "my-task-board-6d5ec.firebaseapp.com",
  projectId: "my-task-board-6d5ec",
  storageBucket: "my-task-board-6d5ec.firebasestorage.app",
  messagingSenderId: "764236575881",
  appId: "1:764236575881:web:3aee1b6076ebbdccfbc258"
};

const app = initializeApp(firebaseConfig);
export const db = getFirestore(app);
export const auth = getAuth(app);