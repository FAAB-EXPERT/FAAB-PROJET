import { StrictMode, Component, type ReactNode } from 'react'
import { createRoot } from 'react-dom/client'
import './index.css'
import App from './App.tsx'

class ErrorBoundary extends Component<{ children: ReactNode }, { hasError: boolean }> {
  constructor(props: { children: ReactNode }) {
    super(props);
    this.state = { hasError: false };
  }
  static getDerivedStateFromError() { return { hasError: true }; }
  render() {
    if (this.state.hasError) {
      return (
        <div style={{ padding: '2rem', textAlign: 'center', color: '#fff', background: '#0f172a', height: '100vh', display: 'flex', flexDirection: 'column', justifyContent: 'center' }}>
          <h1>Oups ! Une erreur est survenue lors du chargement.</h1>
          <p>Essayez de vider le cache ou de redémarrer l'application.</p>
          <button onClick={() => { localStorage.clear(); window.location.reload(); }} style={{ marginTop: '1rem', padding: '0.5rem 1rem', background: '#6366f1', color: '#fff', border: 'none', borderRadius: '4px', cursor: 'pointer' }}>Réinitialiser les données (Attention: Perte des données locales)</button>
        </div>
      );
    }
    return this.props.children;
  }
}

// Polyfill et shims sécurisés pour environnement desktop natif Tauri v2
if (typeof window !== 'undefined') {
  // Indicateur d'environnement Tauri
  (window as any).__TAURI_APP__ = true;
}

const rootElement = document.getElementById('root');
if (!rootElement) {
  throw new Error("L'élément racine #root est introuvable.");
}

createRoot(rootElement).render(
  <StrictMode>
    <ErrorBoundary>
      <App />
    </ErrorBoundary>
  </StrictMode>,
)
