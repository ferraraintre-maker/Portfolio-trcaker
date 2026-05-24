import React from 'react'
import ReactDOM from 'react-dom/client'
import App from './App' // Senza l'estensione .jsx finale per evitare conflitti

ReactDOM.createRoot(document.getElementById('root')).render(
  <React.StrictMode>
    <App />
  </React.StrictMode>,
)
