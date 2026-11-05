import React, { useState, useEffect } from 'react';
import AdaptiveCardViewer from './AdaptiveCardViewer';
import './App.css';

const API_BASE = 'http://localhost:5000/api';

export default function App() {
  const [requests, setRequests] = useState([]);
  const [selectedRequestId, setSelectedRequestId] = useState('REQ-1001');
  const [currentCardJson, setCurrentCardJson] = useState(null);
  const [currentRequest, setCurrentRequest] = useState(null);
  const [viewMode, setViewMode] = useState('teams'); // 'teams' | 'outlook'
  const [loading, setLoading] = useState(false);
  const [notification, setNotification] = useState('');
  const [actionLog, setActionLog] = useState([]);

  // Fetch initial request list
  const fetchRequests = async () => {
    try {
      const res = await fetch(`${API_BASE}/requests`);
      const data = await res.json();
      if (data.success) {
        setRequests(data.data);
      }
    } catch (err) {
      console.error('Error fetching requests:', err);
    }
  };

  // Fetch specific Adaptive Card JSON for selected request
  const fetchCardData = async (reqId) => {
    setLoading(true);
    try {
      const res = await fetch(`${API_BASE}/card/${reqId}`);
      const data = await res.json();
      if (data.success) {
        setCurrentRequest(data.data);
        setCurrentCardJson(data.adaptiveCard);
      }
    } catch (err) {
      console.error('Error fetching card:', err);
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    fetchRequests();
    fetchCardData(selectedRequestId);
  }, [selectedRequestId]);

  // Handle action triggered from card (e.g. Approve or Reject click)
  const handleCardAction = async (payload) => {
    const actionType = payload.action; // APPROVE or REJECT
    const comment = payload.managerComment || '';

    addLog(`Sent [Action.Submit] -> Action: ${actionType}, ReqId: ${payload.requestId}, Comment: "${comment}"`);

    try {
      const res = await fetch(`${API_BASE}/action`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          action: actionType,
          requestId: payload.requestId,
          managerComment: comment
        })
      });

      const result = await res.json();
      if (result.success) {
        setNotification(result.message);
        setCurrentRequest(result.data);
        setCurrentCardJson(result.adaptiveCard);
        fetchRequests(); // Refresh request list badges
        setTimeout(() => setNotification(''), 4000);
      }
    } catch (err) {
      console.error('Action error:', err);
      alert('Failed to send card action: ' + err.message);
    }
  };

  const handleReset = async () => {
    await fetch(`${API_BASE}/reset`, { method: 'POST' });
    fetchRequests();
    fetchCardData(selectedRequestId);
    addLog('Reset simulation back to Pending states.');
  };

  const addLog = (msg) => {
    const time = new Date().toLocaleTimeString();
    setActionLog(prev => [`[${time}] ${msg}`, ...prev.slice(0, 8)]);
  };

  return (
    <div className="app-container">
      {/* Top Navigation Bar */}
      <header className="app-header">
        <div className="header-brand">
          <span className="brand-logo">🎴</span>
          <div>
            <h1>Adaptive Cards Studio & Tester</h1>
            <p className="subtitle">Interactive Microsoft Teams & Outlook Actionable Messages Simulation</p>
          </div>
        </div>
        <div className="header-actions">
          <button className="btn-secondary" onClick={handleReset}>
            🔄 Reset Demo Requests
          </button>
        </div>
      </header>

      {/* Main Grid Content */}
      <main className="main-content">
        {/* Left Column: Requests & Architecture Overview */}
        <section className="sidebar-panel">
          <div className="panel-card">
            <h3>📋 Platform4x Leave Workflows</h3>
            <p className="panel-desc">Select a request below to generate its Adaptive Card JSON:</p>
            <div className="request-list">
              {requests.map(req => (
                <div
                  key={req.id}
                  className={`request-item ${selectedRequestId === req.id ? 'active' : ''}`}
                  onClick={() => setSelectedRequestId(req.id)}
                >
                  <div className="req-header">
                    <strong>{req.employeeName}</strong>
                    <span className={`status-badge badge-${req.status.toLowerCase()}`}>
                      {req.status}
                    </span>
                  </div>
                  <div className="req-sub">
                    <span>{req.leaveType}</span> • <span>{req.id}</span>
                  </div>
                </div>
              ))}
            </div>
          </div>

          <div className="panel-card info-box">
            <h4>💡 How Adaptive Cards Work</h4>
            <ol className="step-list">
              <li><strong>JSON Payload:</strong> Backend defines pure JSON schema (elements, columns, actions).</li>
              <li><strong>Native Rendering:</strong> Host app (Teams / Outlook / Web) renders native buttons & inputs.</li>
              <li><strong>Zero Context-Switch:</strong> User approves directly in chat/email without visiting the website!</li>
            </ol>
          </div>

          <div className="panel-card">
            <h4>📡 Action Event Logs (Teams / Bot Hook)</h4>
            <div className="log-console">
              {actionLog.length === 0 ? (
                <span className="log-empty">No actions submitted yet. Click Approve/Reject on card!</span>
              ) : (
                actionLog.map((log, i) => <div key={i} className="log-line">{log}</div>)
              )}
            </div>
          </div>
        </section>

        {/* Center / Right Column: Live Card Rendering & View Switcher */}
        <section className="preview-panel">
          {notification && (
            <div className="notification-banner">
              ✅ {notification}
            </div>
          )}

          <div className="view-mode-tabs">
            <span className="tabs-label">Simulation Host:</span>
            <button
              className={`tab-btn ${viewMode === 'teams' ? 'active-tab' : ''}`}
              onClick={() => setViewMode('teams')}
            >
              💬 Microsoft Teams View
            </button>
            <button
              className={`tab-btn ${viewMode === 'outlook' ? 'active-tab' : ''}`}
              onClick={() => setViewMode('outlook')}
            >
              ✉️ Outlook Actionable Message View
            </button>
          </div>

          {/* Host Simulated Window */}
          <div className={`host-frame ${viewMode}-frame`}>
            <div className="frame-header">
              {viewMode === 'teams' ? (
                <>
                  <div className="teams-avatar">🤖</div>
                  <div>
                    <div className="sender-name">Platform4x Bot <span className="bot-tag">BOT</span></div>
                    <div className="sender-sub">Teams Channel: HR & Operations • Just now</div>
                  </div>
                </>
              ) : (
                <>
                  <div className="outlook-avatar">📨</div>
                  <div>
                    <div className="sender-name">Platform4x Workflow System &lt;notifications@platform4x.com&gt;</div>
                    <div className="sender-sub">Subject: Action Required: Leave Request for Rahul Kumar</div>
                  </div>
                </>
              )}
            </div>

            <div className="frame-body">
              {loading ? (
                <div className="loading-state">Generating card...</div>
              ) : currentCardJson ? (
                <AdaptiveCardViewer
                  cardJson={currentCardJson}
                  onExecuteAction={handleCardAction}
                  theme={viewMode}
                />
              ) : (
                <div>No Card Data</div>
              )}
            </div>
          </div>

          {/* Collapsible / JSON Code Inspector */}
          <details className="json-details">
            <summary>🔍 View Generated Adaptive Card JSON Payload</summary>
            <pre className="json-box">
              {JSON.stringify(currentCardJson, null, 2)}
            </pre>
          </details>
        </section>
      </main>
    </div>
  );
}
