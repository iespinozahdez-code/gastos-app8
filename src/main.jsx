import React from 'react'
import ReactDOM from 'react-dom/client'
import GastosApp from './GastosApp.jsx'

// Polyfill de window.storage (API de artefactos de Claude) usando localStorage,
// para que la app funcione igual una vez hospedada fuera de claude.ai.
if (!window.storage) {
  window.storage = {
    async get(key) {
      const raw = localStorage.getItem(key)
      if (raw === null) throw new Error('key not found')
      return { key, value: raw, shared: false }
    },
    async set(key, value) {
      localStorage.setItem(key, value)
      return { key, value, shared: false }
    },
    async delete(key) {
      const existed = localStorage.getItem(key) !== null
      localStorage.removeItem(key)
      return { key, deleted: existed, shared: false }
    },
    async list(prefix = '') {
      const keys = Object.keys(localStorage).filter(k => k.startsWith(prefix))
      return { keys, prefix, shared: false }
    },
  }
}

ReactDOM.createRoot(document.getElementById('root')).render(
  <React.StrictMode>
    <GastosApp />
  </React.StrictMode>
)
