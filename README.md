# BinaryGuard API Start Script Fix

This fixes the DigitalOcean deploy error:

```text
npm error Missing script: "start"
```

## Replace this file

Replace only this file in your `binaryguard-api` GitHub repository:

```text
package.json
```

## DigitalOcean settings

Build Command:

```bash
npm install
```

Run Command:

```bash
npm start
```

HTTP Port:

```text
8080
```

After replacing the file, push to GitHub and run:

```text
Actions → Force Rebuild and Deploy
```
