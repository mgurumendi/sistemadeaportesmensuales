import { initializeApp } from "firebase/app";
import { getFirestore } from "firebase/firestore"; // Agregamos la herramienta de base de datos

const firebaseConfig = {
  apiKey: "AIzaSyBEdomXHrjMxvvPTKzIA2wofwQT2MtP0hM",
  authDomain: "sistema-de-aportes.firebaseapp.com",
  projectId: "sistema-de-aportes",
  storageBucket: "sistema-de-aportes.firebasestorage.app",
  messagingSenderId: "716782879825",
  appId: "1:716782879825:web:20d17371bc9cea5036e470"
};

// Inicializamos Firebase
export const app = initializeApp(firebaseConfig);
// Exportamos la base de datos para poder usarla en toda la app
export const db = getFirestore(app);