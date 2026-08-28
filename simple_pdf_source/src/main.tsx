import React from 'react'
import ReactDOM from 'react-dom/client'
import App from './App'
import './styles.css'

// Stray drops outside a drop target would otherwise navigate the window away.
window.addEventListener('dragover', (event) => event.preventDefault())
window.addEventListener('drop', (event) => event.preventDefault())

ReactDOM.createRoot(document.getElementById('root')!).render(
  <React.StrictMode>
    <App />
  </React.StrictMode>,
)
