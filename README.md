# LK Morning Dashboard

A daily briefing dashboard for LK Singh — follow-ups, PO approvals, calendar, and AI chat.

## Setup

### 1. Add your API keys

Edit `.env` file:
```
REACT_APP_AZURE_CLIENT_ID=e48fff16-3351-48b7-90d8-d1b055de9c7b
REACT_APP_GEMINI_API_KEY=YOUR_GEMINI_KEY_HERE
REACT_APP_REDIRECT_URI=https://your-app.vercel.app
```

### 2. Add API Permissions in Azure

Go to portal.azure.com → App Registrations → LK Dashboard → API Permissions → Add:
- Microsoft Graph → Delegated → `Mail.Read`
- Microsoft Graph → Delegated → `Mail.Send`  
- Microsoft Graph → Delegated → `Calendars.Read`
- Microsoft Graph → Delegated → `User.Read`

### 3. Deploy to Vercel

```bash
npm install -g vercel
vercel
```

Or connect your GitHub repo to Vercel and add env vars in Vercel dashboard.

### 4. Add Redirect URI in Azure

After deploy, go to Azure → App Registrations → LK Dashboard → Authentication → Add redirect URI:
```
https://your-app.vercel.app
```

### 5. Share with LK

Send him the Vercel URL. He clicks "Sign in with Microsoft" — done.

## Features

- **Follow-ups**: Detects sent emails with no reply in 3+ days
- **PO Approvals**: Detects vendor/procurement emails in inbox
- **Calendar**: Today's meetings with direct Outlook links
- **AI To-do**: Gemini generates prioritized daily task list
- **AI Chat**: Ask anything about your day
- All email links open directly in Outlook Web

## Tech Stack

- React (frontend)
- Microsoft Graph API (Outlook + Calendar)
- Gemini API (AI summarization + chat)
- MSAL.js (Microsoft authentication)
- Vercel (hosting)
