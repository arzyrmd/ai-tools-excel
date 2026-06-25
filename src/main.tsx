import React from 'react';
import ReactDOM from 'react-dom/client';
import App from './App';
import './index.css';

/* Menginisialisasi React setelah SDK OfficeJS siap */
// @ts-ignore
Office.onReady((info) => {
  const isExcel = info.host === Office.HostType.Excel;
  console.log(`Office.js is ready. Host: ${info.host}. Platform: ${info.platform}`);
  
  ReactDOM.createRoot(document.getElementById('root')!).render(
    <React.StrictMode>
      <App isExcel={isExcel} />
    </React.StrictMode>
  );
});
