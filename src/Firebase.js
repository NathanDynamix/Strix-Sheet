import { initializeApp } from "firebase/app";
import { 
  getAuth, 
  GoogleAuthProvider, 
  signInWithPopup, 
  signInWithEmailAndPassword,
  createUserWithEmailAndPassword,
  sendPasswordResetEmail,
  signOut,
  onAuthStateChanged,
  updateProfile
} from "firebase/auth";
import { getFirestore } from "firebase/firestore";

const firebaseConfig = {
  apiKey: "AIzaSyDlfkCiSdJqo_H742a4Y_l2fjKG7EmaI4A",
  authDomain: "cart-11fbe.firebaseapp.com",
  databaseURL: "https://cart-11fbe-default-rtdb.firebaseio.com",
  projectId: "cart-11fbe",
  storageBucket: "cart-11fbe.firebasestorage.app",
  messagingSenderId: "816862038718",
  appId: "1:816862038718:web:8fbf05d37cdfbdddc2ecef"
};

const app = initializeApp(firebaseConfig);
const auth = getAuth(app);
const db = getFirestore(app);
const provider = new GoogleAuthProvider();

// Configure Google provider
provider.setCustomParameters({
  prompt: 'select_account'
});

export { 
  auth, 
  db,
  provider, 
  signInWithPopup, 
  signInWithEmailAndPassword,
  createUserWithEmailAndPassword,
  sendPasswordResetEmail,
  signOut,
  onAuthStateChanged,
  updateProfile
};