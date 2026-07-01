# BinaryGuard API

This API combines:

- Website contact form email via Microsoft Graph
- Portal OTP request and verification
- Supabase `otp_codes` integration
- Microsoft 365 email delivery

## File placement

Place these files inside your existing `binaryguard-api` component:

```text
binaryguard-website
└─ binaryguard-api
   ├─ server.ts
   ├─ package.json
   ├─ .env
   ├─ .env.example
   └─ .gitignore
```

## DigitalOcean App Platform settings

Component type:

```text
Web Service
```

Build command:

```bash
npm install
```

Run command:

```bash
npm start
```

HTTP port:

```text
8080
```

## Environment variables

Copy the values from `.env` into DigitalOcean App Platform.

Do not commit real secrets to GitHub:

```text
SUPABASE_SERVICE_ROLE_KEY
AZURE_CLIENT_SECRET
OTP_PEPPER
```

## API endpoints

```http
GET /api/health
GET /health
POST /api/contact
POST /api/otp/request
POST /api/otp/verify
```

## Portal frontend variable

Add this to the portal frontend:

```env
VITE_PORTAL_API_URL=https://api.binaryguard.ca
```
