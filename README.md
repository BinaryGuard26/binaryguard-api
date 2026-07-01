# BinaryGuard API Runtime Fix

This fixes the DigitalOcean runtime error:

```text
Node.js 20 detected without native WebSocket support.
```

## Replace these files in your `binaryguard-api` GitHub repository

```text
server.ts
package.json
.env.example
.gitignore
README.md
```

## Important

Do not upload your real `.env` file to GitHub. Add real values in DigitalOcean App Platform → Environment Variables.

## DigitalOcean

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

After pushing the files, use:

```text
Actions → Force Rebuild and Deploy
```

Then test:

```text
https://your-api-url/health
```
