// Ket noi Firebase (dung lai config tu src/firebase/index.js)
import { initializeApp } from 'firebase/app';
import { getDatabase } from 'firebase/database';

const firebaseConfig = {
  apiKey: 'AIzaSyA21puOOEVULDMocGzkeJwEpop3UXX9-cg',
  authDomain: 'db-preson.firebaseapp.com',
  databaseURL: 'https://db-preson-default-rtdb.asia-southeast1.firebasedatabase.app',
  projectId: 'db-preson',
  storageBucket: 'db-preson.firebasestorage.app',
  messagingSenderId: '439870683002',
  appId: '1:439870683002:web:8190e3db47702c4afd08c3',
  measurementId: 'G-FEWCDDBBF1',
};

const app = initializeApp(firebaseConfig);
export const db = getDatabase(app);
