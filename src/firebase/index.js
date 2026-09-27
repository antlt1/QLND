// Import the functions you need from the SDKs you need
import { initializeApp } from "firebase/app";
import { getAnalytics } from "firebase/analytics";
// TODO: Add SDKs for Firebase products that you want to use
// https://firebase.google.com/docs/web/setup#available-libraries

// Your web app's Firebase configuration
// For Firebase JS SDK v7.20.0 and later, measurementId is optional
const firebaseConfig = {
  apiKey: "AIzaSyA21puOOEVULDMocGzkeJwEpop3UXX9-cg",
  authDomain: "db-preson.firebaseapp.com",
  databaseURL: "https://db-preson-default-rtdb.asia-southeast1.firebasedatabase.app",
  projectId: "db-preson",
  storageBucket: "db-preson.firebasestorage.app",
  messagingSenderId: "439870683002",
  appId: "1:439870683002:web:8190e3db47702c4afd08c3",
  measurementId: "G-FEWCDDBBF1"
};

// Initialize Firebase
const app = initializeApp(firebaseConfig);
const analytics = getAnalytics(app);