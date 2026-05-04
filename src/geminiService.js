const GEMINI_API_KEY = process.env.REACT_APP_GEMINI_API_KEY;
const GEMINI_URL = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=${GEMINI_API_KEY}`;

async function callGemini(prompt) {
  const response = await fetch(GEMINI_URL, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      contents: [{ parts: [{ text: prompt }] }],
      generationConfig: { temperature: 0.3, maxOutputTokens: 1024 },
    }),
  });
  if (!response.ok) throw new Error(`Gemini error: ${response.status}`);
  const data = await response.json();
  return data.candidates?.[0]?.content?.parts?.[0]?.text || "";
}

export async function summarizePO(email) {
  const prompt = `You are an assistant for a real estate General Manager in India.
Summarize this purchase order / vendor email in 2 lines max. Be specific about: vendor name, what is being ordered, amount if mentioned, and urgency.
Subject: ${email.subject}
Preview: ${email.bodyPreview}
Reply in plain text, no bullets, no markdown.`;
  return callGemini(prompt);
}

export async function summarizeFollowUp(email) {
  const prompt = `You are an assistant for a real estate General Manager.
This is an email sent by the GM that got no reply in 3+ days. In one line, explain what action is needed.
Subject: ${email.subject}
Preview: ${email.bodyPreview}
Reply in plain text, one line only.`;
  return callGemini(prompt);
}

export async function chatWithGemini(userMessage, context) {
  const prompt = `You are a smart executive assistant for LK Singh, General Manager at Panchshil Realty, Pune. 
He manages large-scale real estate construction projects including MEP and interior fit-out works.
Context about his day: ${context}

User question: ${userMessage}

Be concise, practical, and professional. Reply in plain text.`;
  return callGemini(prompt);
}

export async function generateTodoList(followUps, poEmails, events) {
  const prompt = `You are an assistant for a real estate GM. Generate a prioritized to-do list for today based on:

FOLLOW-UPS NEEDED (no reply in 3+ days):
${followUps.slice(0, 5).map(e => `- ${e.subject} (sent to ${e.toRecipients?.[0]?.emailAddress?.name || 'unknown'})`).join('\n')}

PO APPROVALS PENDING:
${poEmails.slice(0, 5).map(e => `- ${e.subject}`).join('\n')}

MEETINGS TODAY:
${events.slice(0, 5).map(e => `- ${e.subject} at ${new Date(e.start.dateTime).toLocaleTimeString('en-IN', { hour: '2-digit', minute: '2-digit' })}`).join('\n')}

Return a JSON array of todo items. Each item: { "text": "action needed", "type": "followup|po|meeting", "priority": "high|medium|low" }
Return ONLY the JSON array, no other text.`;

  const result = await callGemini(prompt);
  try {
    const clean = result.replace(/```json|```/g, "").trim();
    return JSON.parse(clean);
  } catch {
    return [];
  }
}
