import express from "express";
import cors from "cors";
import dotenv from "dotenv";
import { ConfidentialClientApplication } from "@azure/msal-node";

dotenv.config();

// DigitalOcean Node 20 runtime may not expose native WebSocket for Supabase Realtime.
if (!globalThis.WebSocket) {
  globalThis.WebSocket = WebSocket as unknown as typeof globalThis.WebSocket;
}


const app = express();

app.use(cors());
app.use(express.json());

const cca = new ConfidentialClientApplication({
  auth: {
    clientId: process.env.AZURE_CLIENT_ID!,
    authority: `https://login.microsoftonline.com/${process.env.AZURE_TENANT_ID}`,
    clientSecret: process.env.AZURE_CLIENT_SECRET!,
  },
});

async function getAccessToken() {
  const result = await cca.acquireTokenByClientCredential({
    scopes: ["https://graph.microsoft.com/.default"],
  });

  if (!result?.accessToken) {
    throw new Error("Failed to get Microsoft Graph access token");
  }

  return result.accessToken;
}

app.get("/api/health", (_req, res) => {
  res.json({
    success: true,
    message: "SMTP API healthy",
  });
});

app.post("/api/contact", async (req, res) => {
  try {
    const { name, company, email, phone, subject, message } = req.body;

    if (!name || !email || !subject || !message) {
      return res.status(400).json({
        success: false,
        message: "Required fields are missing.",
      });
    }

    const accessToken = await getAccessToken();

    const sender = process.env.SMTP_USER!;
    const receiver = process.env.CONTACT_RECEIVER_EMAIL || sender;

    const graphResponse = await fetch(
      `https://graph.microsoft.com/v1.0/users/${sender}/sendMail`,
      {
        method: "POST",
        headers: {
          Authorization: `Bearer ${accessToken}`,
          "Content-Type": "application/json",
        },
        body: JSON.stringify({
          message: {
            subject: `Website Contact: ${subject}`,
            body: {
              contentType: "HTML",
              content: `
                <h2>New Contact Request</h2>

                <p><strong>Name:</strong> ${name}</p>
                <p><strong>Company:</strong> ${company || "N/A"}</p>
                <p><strong>Email:</strong> ${email}</p>
                <p><strong>Phone:</strong> ${phone || "N/A"}</p>
                <p><strong>Subject:</strong> ${subject}</p>

                <hr />

                <p><strong>Message:</strong></p>
                <p>${message}</p>
              `,
            },
            toRecipients: [
              {
                emailAddress: {
                  address: receiver,
                },
              },
            ],
            replyTo: [
              {
                emailAddress: {
                  address: email,
                },
              },
            ],
          },
          saveToSentItems: true,
        }),
      }
    );

    if (!graphResponse.ok) {
      const errorText = await graphResponse.text();
      console.error("Microsoft Graph error:", errorText);

      return res.status(500).json({
        success: false,
        message: "Microsoft Graph email failed.",
        error: errorText,
      });
    }

    return res.status(200).json({
      success: true,
      message: "Email sent successfully.",
    });
  } catch (error) {
    console.error("Contact API error:", error);

    return res.status(500).json({
      success: false,
      message: "Failed to send message.",
    });
  }
});

const PORT = process.env.PORT || 3001;

app.listen(PORT, () => {
  console.log(`SMTP API running on port ${PORT}`);
});
