import React from 'react';
import ReactDOM from 'react-dom/client';
import App from './App.jsx';
import './index.css';

ReactDOM.createRoot(document.getElementById('root')).render(
  <React.StrictMode>
    <App />
  </React.StrictMode>
);

// PR-1 — Registro del Service Worker (PWA).
// Network-first: un deploy nuevo se sirve en la próxima carga/reapertura de la
// app, sin dejar bundle viejo cacheado. En localhost/HTTP el registro puede
// fallar silenciosamente; en producción (Vercel, HTTPS) funciona.
if ('serviceWorker' in navigator) {
  window.addEventListener('load', () => {
    navigator.serviceWorker.register('/sw.js').catch((err) => {
      console.warn('SW registro falló:', err);
    });
  });
}
