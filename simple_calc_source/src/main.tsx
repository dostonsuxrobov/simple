import React, { Component } from 'react'
import type { ReactNode } from 'react'
import ReactDOM from 'react-dom/client'
import App from './App'
import { runRecoverySave } from './lib/recovery'
import './styles.css'

interface RecoveryBoundaryState {
  error: Error | null
  status: string
}

class RecoveryBoundary extends Component<{ children: ReactNode }, RecoveryBoundaryState> {
  state: RecoveryBoundaryState = { error: null, status: '' }

  static getDerivedStateFromError(error: unknown): Partial<RecoveryBoundaryState> {
    return { error: error instanceof Error ? error : new Error(String(error)) }
  }

  saveRecoveryCopy = async () => {
    this.setState({ status: 'Saving recovery copy…' })
    try {
      const result = await runRecoverySave()
      this.setState({ status: result ? `Saved ${result.name}` : 'Save canceled' })
    } catch (error) {
      this.setState({ status: error instanceof Error ? error.message : 'The recovery copy could not be saved.' })
    }
  }

  render() {
    if (!this.state.error) return this.props.children
    return (
      <div className="crash-screen">
        <div className="crash-card">
          <h1>Something went wrong</h1>
          <p>simple_calc hit an unexpected error. Save a recovery copy of your workbook, then reload the app.</p>
          <pre>{this.state.error.message || String(this.state.error)}</pre>
          <div className="crash-actions">
            <button type="button" className="primary-action" onClick={() => { void this.saveRecoveryCopy() }}>Save recovery copy</button>
            <button type="button" className="secondary-action" onClick={() => window.location.reload()}>Reload</button>
          </div>
          {this.state.status && <span className="crash-status" aria-live="polite">{this.state.status}</span>}
        </div>
      </div>
    )
  }
}

ReactDOM.createRoot(document.getElementById('root')!).render(
  <React.StrictMode>
    <RecoveryBoundary>
      <App />
    </RecoveryBoundary>
  </React.StrictMode>,
)
