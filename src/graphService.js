export async function callGraph(endpoint, token) {
  const response = await fetch(endpoint, {
    headers: { Authorization: `Bearer ${token}` },
  });
  if (!response.ok) throw new Error(`Graph API error: ${response.status}`);
  return response.json();
}

export async function getEmails(token) {
  const url = `https://graph.microsoft.com/v1.0/me/messages?$top=50&$orderby=receivedDateTime desc&$select=id,subject,from,receivedDateTime,isRead,bodyPreview,webLink`;
  return callGraph(url, token);
}

export async function getSentEmails(token) {
  const url = `https://graph.microsoft.com/v1.0/me/mailFolders/sentItems/messages?$top=50&$orderby=sentDateTime desc&$select=id,subject,toRecipients,sentDateTime,bodyPreview,webLink,conversationId`;
  return callGraph(url, token);
}

export async function getInboxEmails(token) {
  const url = `https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages?$top=50&$orderby=receivedDateTime desc&$select=id,subject,from,receivedDateTime,isRead,bodyPreview,webLink,conversationId`;
  return callGraph(url, token);
}

export async function getTodayEvents(token) {
  const now = new Date();
  const startOfDay = new Date(now.getFullYear(), now.getMonth(), now.getDate()).toISOString();
  const endOfDay = new Date(now.getFullYear(), now.getMonth(), now.getDate(), 23, 59, 59).toISOString();
  const url = `https://graph.microsoft.com/v1.0/me/calendarView?startDateTime=${startOfDay}&endDateTime=${endOfDay}&$select=id,subject,start,end,location,webLink,organizer&$orderby=start/dateTime`;
  return callGraph(url, token);
}

export async function getUser(token) {
  return callGraph("https://graph.microsoft.com/v1.0/me?$select=displayName,mail,jobTitle", token);
}

export function detectPOEmails(emails) {
  const poKeywords = ["purchase order", "PO ", "P.O.", "vendor", "invoice", "approval", "quotation", "quote", "procurement", "supply"];
  return emails.filter(email => {
    const text = `${email.subject} ${email.bodyPreview}`.toLowerCase();
    return poKeywords.some(kw => text.includes(kw.toLowerCase()));
  });
}

export function detectFollowUps(sentEmails, inboxEmails) {
  const threeDaysAgo = new Date(Date.now() - 3 * 24 * 60 * 60 * 1000);
  const inboxConversationIds = new Set(inboxEmails.map(e => e.conversationId));
  
  return sentEmails.filter(sent => {
    const sentDate = new Date(sent.sentDateTime);
    const isOld = sentDate < threeDaysAgo;
    const hasNoReply = !inboxConversationIds.has(sent.conversationId);
    return isOld && hasNoReply;
  });
}
