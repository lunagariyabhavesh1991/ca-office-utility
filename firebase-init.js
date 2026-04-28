import { initializeApp } from "https://www.gstatic.com/firebasejs/11.0.1/firebase-app.js";
import { getAuth, createUserWithEmailAndPassword, signInWithEmailAndPassword, onAuthStateChanged, signOut, updateProfile, sendPasswordResetEmail } from "https://www.gstatic.com/firebasejs/11.0.1/firebase-auth.js";
import { getFirestore, doc, setDoc, addDoc, collection, serverTimestamp, getDoc, getDocs, query, orderBy, where, updateDoc, increment, limit, deleteDoc } from "https://www.gstatic.com/firebasejs/11.0.1/firebase-firestore.js";

const firebaseConfig = {
  apiKey: "AIzaSyDaInSM0602SYzJ9D7eg16Nkz4NjC0aULI",
  authDomain: "ca-office-utility.firebaseapp.com",
  projectId: "ca-office-utility",
  storageBucket: "ca-office-utility.firebasestorage.app",
  messagingSenderId: "279120698156",
  appId: "1:279120698156:web:a2f69c652b8d052ff52a89",
  measurementId: "G-ES3YSGF1WH"
};

const app = initializeApp(firebaseConfig);
const auth = getAuth(app);
const db = getFirestore(app);

export { 
  app, 
  auth, 
  db, 
  createUserWithEmailAndPassword, 
  signInWithEmailAndPassword, 
  onAuthStateChanged, 
  signOut, 
  updateProfile, 
  sendPasswordResetEmail,
  doc, 
  setDoc, 
  addDoc, 
  collection, 
  serverTimestamp, 
  getDoc,
  getDocs,
  query,
  orderBy,
  where,
  updateDoc,
  increment,
  limit,
  deleteDoc
};
