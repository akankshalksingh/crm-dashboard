import React, { useState, useEffect, useCallback } from "react";
import { PublicClientApplication } from "@azure/msal-browser";
import { msalConfig, loginRequest } from "./authConfig";
import { getEmails, getSentEmails, getInboxEmails, getTodayEvents, getUser, detectPOEmails, detectFollowUps } from "./graphService";
import { summarizePO, summarizeFollowUp, generateTodoList, chatWithGemini } from "./geminiService";
import "./App.css";

const msalInstance = new PublicClientApplication(msalConfig);

function timeAgo(dateStr) {
  const date = new Date(dateStr);
  const now = new Date();
  const diff = Math.floor((now - date) / (1000 * 60 * 60 * 24));
  if (diff === 0) return "Today";
  if (diff === 1) return "Yesterday";
  return `${diff} days ago`;
}

function formatTime(dateStr) {
  return new Date(dateStr).toLocaleTimeString("en-IN", { hour: "2-digit", minute: "2-digit" });
}

function formatDate(dateStr) {
  return new Date(dateStr).toLocaleDateString("en-IN", { weekday: "long", day: "numeric", month: "long", year: "numeric" });
}

function getInitials(name) {
  if (!name) return "??";
  return name.split(" ").map(n => n[0]).slice(0, 2).join("").toUpperCase();
}

function priorityDot(priority) {
  const colors = { high: "#E24B4A", medium: "#EF9F27", low: "#378ADD" };
  return <span className="dot" style={{ background: colors[priority] || colors.low }} />;
}

export default function App() {
  const [isAuthenticated, setIsAuthenticated] = useState(false);
  const [token, setToken] = useState(null);
  const [user, setUser] = useState(null);
  const [loading, setLoading] = useState(false);
  const [loadingStep, setLoadingStep] = useState("");
  const [followUps, setFollowUps] = useState([]);
  const [poEmails, setPoEmails] = useState([]);
  const [events, setEvents] = useState([]);
  const [todoList, setTodoList] = useState([]);
  const [summaries, setSummaries] = useState({});
  const [chatMessages, setChatMessages] = useState([]);
  const [chatInput, setChatInput] = useState("");
  const [chatLoading, setChatLoading] = useState(false);
  const [activeTab, setActiveTab] = useState("overview");
  const [initialized, setInitialized] = useState(false);

  useEffect(() => {
    msalInstance.initialize().then(() => {
      msalInstance.handleRedirectPromise().then(response => {
        if (response) handleAuthResponse(response);
        const accounts = msalInstance.getAllAccounts();
        if (accounts.length > 0) silentSignIn(accounts[0]);
      });
      setInitialized(true);
    });
  }, []);

  async function handleAuthResponse(response) {
    setToken(response.accessToken);
    setIsAuthenticated(true);
  }

  async function silentSignIn(account) {
    try {
      const response = await msalInstance.acquireTokenSilent({ ...loginRequest, account });
      setToken(response.accessToken);
      setIsAuthenticated(true);
    } catch {}
  }

  async function signIn() {
    try {
      const response = await msalInstance.loginPopup(loginRequest);
      setToken(response.accessToken);
      setIsAuthenticated(true);
    } catch (err) {
      console.error("Login failed", err);
    }
  }

  async function signOut() {
    await msalInstance.logoutPopup();
    setIsAuthenticated(false);
    setToken(null);
    setUser(null);
  }

  const loadDashboard = useCallback(async () => {
    if (!token) return;
    setLoading(true);
    try {
      setLoadingStep("Getting your profile...");
      const userData = await getUser(token);
      setUser(userData);

      setLoadingStep("Reading your inbox...");
      const [inboxData, sentData, calendarData] = await Promise.all([
        getInboxEmails(token),
        getSentEmails(token),
        getTodayEvents(token),
      ]);

      const inbox = inboxData.value || [];
      const sent = sentData.value || [];
      const cal = calendarData.value || [];

      setLoadingStep("Finding follow-ups...");
      const fu = detectFollowUps(sent, inbox);
      setFollowUps(fu.slice(0, 8));

      setLoadingStep("Detecting PO emails...");
      const po = detectPOEmails(inbox);
      setPoEmails(po.slice(0, 6));

      setEvents(cal);

      setLoadingStep("Generating your to-do list...");
      const todos = await generateTodoList(fu, po, cal);
      setTodoList(todos);

      setLoadingStep("Summarizing POs with AI...");
      const newSummaries = {};
      for (const email of po.slice(0, 4)) {
        try {
          newSummaries[email.id] = await summarizePO(email);
        } catch {}
      }
      for (const email of fu.slice(0, 4)) {
        try {
          newSummaries[email.id] = await summarizeFollowUp(email);
        } catch {}
      }
      setSummaries(newSummaries);

    } catch (err) {
      console.error("Dashboard load error", err);
    }
    setLoading(false);
    setLoadingStep("");
  }, [token]);

  useEffect(() => {
    if (isAuthenticated && token) loadDashboard();
  }, [isAuthenticated, token, loadDashboard]);

  async function sendChat() {
    if (!chatInput.trim()) return;
    const userMsg = chatInput.trim();
    setChatInput("");
    setChatMessages(prev => [...prev, { role: "user", text: userMsg }]);
    setChatLoading(true);
    try {
      const context = `${followUps.length} follow-ups pending, ${poEmails.length} PO approvals needed, ${events.length} meetings today.`;
      const reply = await chatWithGemini(userMsg, context);
      setChatMessages(prev => [...prev, { role: "assistant", text: reply }]);
    } catch {
      setChatMessages(prev => [...prev, { role: "assistant", text: "Sorry, couldn't get a response. Try again." }]);
    }
    setChatLoading(false);
  }

  if (!initialized) return <div className="splash"><div className="spinner" /></div>;

  if (!isAuthenticated) {
    return (
      <div className="login-screen">
        <div className="login-card">
          <div className="login-logo">LK</div>
          <h1>Morning Briefing</h1>
          <p>Your daily command center — follow-ups, POs, meetings, all in one place.</p>
          <button className="signin-btn" onClick={signIn}>
            <svg width="20" height="20" viewBox="0 0 23 23" fill="none">
              <path d="M1 1h10v10H1V1zm11 0h10v10H12V1zM1 12h10v10H1V12zm11 0h10v10H12V12z" fill="#fff" opacity=".9"/>
            </svg>
            Sign in with Microsoft
          </button>
          <p className="login-note">Uses your existing Outlook & Calendar. No data is stored.</p>
        </div>
      </div>
    );
  }

  if (loading) {
    return (
      <div className="splash">
        <div className="spinner" />
        <p className="loading-text">{loadingStep}</p>
      </div>
    );
  }

  const overdueCount = followUps.length;
  const poCount = poEmails.length;
  const meetingCount = events.length;
  const nextMeeting = events[0];

  return (
    <div className="app">
      <header className="header">
        <div className="header-left">
          <div className="logo-mark">LK</div>
          <div>
            <h1>Good morning, {user?.displayName?.split(" ")[0] || "LK"}</h1>
            <p className="header-date">{formatDate(new Date())}</p>
          </div>
        </div>
        <div className="header-right">
          <button className="refresh-btn" onClick={loadDashboard}>↻ Refresh</button>
          <button className="signout-btn" onClick={signOut}>Sign out</button>
        </div>
      </header>

      <div className="metric-row">
        <div className={`metric ${overdueCount > 0 ? "metric-alert" : ""}`}>
          <div className="metric-label">Follow-ups overdue</div>
          <div className="metric-value">{overdueCount}</div>
          <div className="metric-sub">{overdueCount > 0 ? `Oldest: ${timeAgo(followUps[followUps.length - 1]?.sentDateTime)}` : "All caught up"}</div>
        </div>
        <div className={`metric ${poCount > 0 ? "metric-warn" : ""}`}>
          <div className="metric-label">POs pending</div>
          <div className="metric-value">{poCount}</div>
          <div className="metric-sub">Needs your review</div>
        </div>
        <div className="metric">
          <div className="metric-label">Meetings today</div>
          <div className="metric-value">{meetingCount}</div>
          <div className="metric-sub">{nextMeeting ? `Next: ${formatTime(nextMeeting.start.dateTime)}` : "None scheduled"}</div>
        </div>
        <div className="metric">
          <div className="metric-label">To-do items</div>
          <div className="metric-value">{todoList.length}</div>
          <div className="metric-sub">AI generated</div>
        </div>
      </div>

      <div className="tabs">
        {["overview", "followups", "po", "calendar", "chat"].map(tab => (
          <button key={tab} className={`tab ${activeTab === tab ? "tab-active" : ""}`} onClick={() => setActiveTab(tab)}>
            {tab === "overview" ? "Overview" : tab === "followups" ? `Follow-ups (${overdueCount})` : tab === "po" ? `PO Approvals (${poCount})` : tab === "calendar" ? "Calendar" : "Ask AI"}
          </button>
        ))}
      </div>

      <div className="content">
        {activeTab === "overview" && (
          <div className="two-col">
            <div className="card">
              <div className="card-header">
                <h2>Today's to-do</h2>
                <span className="badge badge-blue">AI generated</span>
              </div>
              {todoList.length === 0 ? <p className="empty">No items generated yet.</p> : todoList.map((item, i) => (
                <div key={i} className="todo-row">
                  {priorityDot(item.priority)}
                  <span className="todo-text">{item.text}</span>
                  <span className="todo-tag">{item.type === "followup" ? "Follow-up" : item.type === "po" ? "PO" : "Meeting"}</span>
                </div>
              ))}
            </div>
            <div className="card">
              <div className="card-header">
                <h2>Today's meetings</h2>
                <span className="badge badge-blue">{meetingCount} scheduled</span>
              </div>
              {events.length === 0 ? <p className="empty">No meetings today.</p> : events.map(event => (
                <div key={event.id} className="event-row">
                  <div className="event-time">{formatTime(event.start.dateTime)}</div>
                  <div className="event-body">
                    <div className="event-title">{event.subject}</div>
                    {event.location?.displayName && <div className="event-loc">{event.location.displayName}</div>}
                  </div>
                  <a href={event.webLink} target="_blank" rel="noreferrer" className="open-btn">Open ↗</a>
                </div>
              ))}
            </div>
          </div>
        )}

        {activeTab === "followups" && (
          <div className="card">
            <div className="card-header">
              <h2>Follow-up needed</h2>
              <span className="badge badge-red">{overdueCount} no reply</span>
            </div>
            {followUps.length === 0 ? <p className="empty">No follow-ups needed. You're on top of everything!</p> : followUps.map(email => (
              <div key={email.id} className="email-row">
                <div className="avatar">{getInitials(email.toRecipients?.[0]?.emailAddress?.name)}</div>
                <div className="email-body">
                  <div className="email-meta">
                    <span className="email-from">{email.toRecipients?.[0]?.emailAddress?.name || email.toRecipients?.[0]?.emailAddress?.address}</span>
                    <span className="email-time">{timeAgo(email.sentDateTime)}</span>
                  </div>
                  <div className="email-subject">{email.subject}</div>
                  {summaries[email.id] && <div className="ai-summary">AI: {summaries[email.id]}</div>}
                  <div className="email-actions">
                    <a href={email.webLink} target="_blank" rel="noreferrer" className="btn-primary">Open in Outlook ↗</a>
                    <button className="btn-ghost" onClick={() => setActiveTab("chat")}>Ask AI about this</button>
                  </div>
                </div>
              </div>
            ))}
          </div>
        )}

        {activeTab === "po" && (
          <div className="card">
            <div className="card-header">
              <h2>PO approvals</h2>
              <span className="badge badge-amber">{poCount} pending</span>
            </div>
            {poEmails.length === 0 ? <p className="empty">No PO emails detected.</p> : poEmails.map(email => (
              <div key={email.id} className="email-row">
                <div className="avatar po-avatar">{getInitials(email.from?.emailAddress?.name)}</div>
                <div className="email-body">
                  <div className="email-meta">
                    <span className="email-from">{email.from?.emailAddress?.name || email.from?.emailAddress?.address}</span>
                    <span className="email-time">{timeAgo(email.receivedDateTime)}</span>
                  </div>
                  <div className="email-subject">{email.subject}</div>
                  {summaries[email.id] && <div className="ai-summary">AI: {summaries[email.id]}</div>}
                  <div className="email-actions">
                    <a href={email.webLink} target="_blank" rel="noreferrer" className="btn-primary">Open in Outlook ↗</a>
                    <button className="btn-ghost" onClick={() => setActiveTab("chat")}>Ask AI about this</button>
                  </div>
                </div>
              </div>
            ))}
          </div>
        )}

        {activeTab === "calendar" && (
          <div className="card">
            <div className="card-header">
              <h2>Today's calendar</h2>
              <span className="badge badge-blue">{formatDate(new Date())}</span>
            </div>
            {events.length === 0 ? <p className="empty">No meetings scheduled for today.</p> : events.map(event => (
              <div key={event.id} className="event-row-full">
                <div className="event-time-block">
                  <div className="event-start">{formatTime(event.start.dateTime)}</div>
                  <div className="event-end">{formatTime(event.end.dateTime)}</div>
                </div>
                <div className="event-details">
                  <div className="event-title">{event.subject}</div>
                  {event.location?.displayName && <div className="event-loc">{event.location.displayName}</div>}
                  {event.organizer?.emailAddress?.name && <div className="event-org">Organizer: {event.organizer.emailAddress.name}</div>}
                </div>
                <a href={event.webLink} target="_blank" rel="noreferrer" className="btn-primary">Open ↗</a>
              </div>
            ))}
          </div>
        )}

        {activeTab === "chat" && (
          <div className="card chat-card">
            <div className="card-header">
              <h2>Ask your AI assistant</h2>
              <span className="badge badge-blue">Gemini</span>
            </div>
            <div className="chat-messages">
              {chatMessages.length === 0 && (
                <div className="chat-hints">
                  <p className="hint-label">Try asking:</p>
                  {["Summarize my most urgent follow-up", "Draft a follow-up email for a vendor", "What should I prioritize today?", "Help me write a PO approval response"].map((hint, i) => (
                    <button key={i} className="hint-btn" onClick={() => { setChatInput(hint); }}>{hint}</button>
                  ))}
                </div>
              )}
              {chatMessages.map((msg, i) => (
                <div key={i} className={`chat-msg ${msg.role}`}>
                  <div className="msg-bubble">{msg.text}</div>
                </div>
              ))}
              {chatLoading && <div className="chat-msg assistant"><div className="msg-bubble typing">Thinking...</div></div>}
            </div>
            <div className="chat-input-row">
              <input
                className="chat-input"
                value={chatInput}
                onChange={e => setChatInput(e.target.value)}
                onKeyDown={e => e.key === "Enter" && sendChat()}
                placeholder="Ask anything about your day..."
              />
              <button className="send-btn" onClick={sendChat} disabled={chatLoading}>Send</button>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}
