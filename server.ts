import express, { Request, Response } from "express";
import cors from "cors";
import helmet from "helmet";
import dotenv from "dotenv";
import crypto from "crypto";
import { createClient } from "@supabase/supabase-js";
import { ConfidentialClientApplication } from "@azure/msal-node";

dotenv.config();

const requiredEnv = [
  "SUPABASE_URL",
  "SUPABASE_SERVICE_ROLE_KEY",
  "AZURE_TENANT_ID",
  "AZURE_CLIENT_ID",
  "AZURE_CLIENT_SECRET",
  "SMTP_USER",
  "SMTP_FROM",
  "OTP_PEPPER",
];

for (const key of requiredEnv) {
  if (!process.env[key]) {
    throw new Error(`Missing required environment variable: ${key}`);
  }
}

type OtpPurpose = "registration" | "login" | "recovery" | "cancel_order";

type TenantLookup = {
  tenantId: string;
  tenantName: string;
  domain: string;
};

type OtpRow = {
  id: string;
  tenant_id: string | null;
  email: string;
  purpose: OtpPurpose;
  code_hash: string;
  expires_at: string;
  attempt_count: number;
  max_attempts: number;
  used_at: string | null;
};

const app = express();
const PORT = Number(process.env.PORT || 8080);
const otpExpiryMinutes = Number(process.env.OTP_EXPIRY_MINUTES || 5);
const maxAttempts = Number(process.env.OTP_MAX_ATTEMPTS || 3);

const allowedOrigins = String(process.env.FRONTEND_ORIGIN || "https://portal.binaryguard.ca")
  .split(",")
  .map((origin) => origin.trim())
  .filter(Boolean);

app.use(helmet());
app.use(express.json({ limit: "100kb" }));
app.use(
  cors({
    origin(origin, callback) {
      if (!origin || allowedOrigins.includes(origin)) {
        callback(null, true);
        return;
      }

      callback(new Error(`CORS blocked origin: ${origin}`));
    },
    methods: ["GET", "POST", "OPTIONS"],
    allowedHeaders: ["Content-Type", "Authorization"],
  })
);

const supabase = createClient(
  process.env.SUPABASE_URL!,
  process.env.SUPABASE_SERVICE_ROLE_KEY!,
  {
    auth: {
      persistSession: false,
      autoRefreshToken: false,
    },
  }
);

const cca = new ConfidentialClientApplication({
  auth: {
    clientId: process.env.AZURE_CLIENT_ID!,
    authority: `https://login.microsoftonline.com/${process.env.AZURE_TENANT_ID}`,
    clientSecret: process.env.AZURE_CLIENT_SECRET!,
  },
});

function normalizeEmail(email: unknown): string {
  return String(email || "").trim().toLowerCase();
}

function normalizePurpose(purpose: unknown): OtpPurpose | null {
  const value = String(purpose || "").trim().toLowerCase();
  const allowed: OtpPurpose[] = ["registration", "login", "recovery", "cancel_order"];
  return allowed.includes(value as OtpPurpose) ? (value as OtpPurpose) : null;
}

function nowIso(): string {
  return new Date().toISOString();
}

function expiresAtIso(): string {
  return new Date(Date.now() + otpExpiryMinutes * 60 * 1000).toISOString();
}

function generateOtp(): string {
  return String(crypto.randomInt(100000, 1000000));
}

function hashOtp(email: string, purpose: OtpPurpose, code: string): string {
  return crypto
    .createHmac("sha256", process.env.OTP_PEPPER!)
    .update(`${normalizeEmail(email)}:${purpose}:${String(code).trim()}`)
    .digest("hex");
}

function escapeHtml(value: unknown): string {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

async function getAccessToken(): Promise<string> {
  const result = await cca.acquireTokenByClientCredential({
    scopes: ["https://graph.microsoft.com/.default"],
  });

  if (!result?.accessToken) {
    throw new Error("Failed to get Microsoft Graph access token");
  }

  return result.accessToken;
}

async function sendGraphMail({
  to,
  subject,
  html,
  replyTo,
}: {
  to: string;
  subject: string;
  html: string;
  replyTo?: string;
}) {
  const accessToken = await getAccessToken();
  const sender = process.env.SMTP_USER!;
  const graphResponse = await fetch(
    `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(sender)}/sendMail`,
    {
      method: "POST",
      headers: {
        Authorization: `Bearer ${accessToken}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        message: {
          subject,
          body: {
            contentType: "HTML",
            content: html,
          },
          toRecipients: [
            {
              emailAddress: {
                address: to,
              },
            },
          ],
          ...(replyTo
            ? {
                replyTo: [
                  {
                    emailAddress: {
                      address: replyTo,
                    },
                  },
                ],
              }
            : {}),
        },
        saveToSentItems: true,
      }),
    }
  );

  if (!graphResponse.ok) {
    const errorText = await graphResponse.text();
    throw new Error(`Microsoft Graph sendMail failed: ${errorText}`);
  }
}

async function findTenantByEmail(email: string): Promise<TenantLookup | null> {
  const domain = email.split("@")[1];
  if (!domain) return null;

  const { data, error } = await supabase
    .from("allowed_domains")
    .select("tenant_id, domain, is_active, tenants(id, name, status)")
    .eq("domain", domain)
    .eq("is_active", true)
    .maybeSingle();

  if (error) throw error;

  const tenantData = data as any;
  if (!tenantData?.tenants || tenantData.tenants.status !== "active") {
    return null;
  }

  return {
    tenantId: tenantData.tenant_id,
    tenantName: tenantData.tenants.name,
    domain: tenantData.domain,
  };
}

async function sendOtpEmail(email: string, code: string, purpose: OtpPurpose) {
  const purposeLabel = purpose.replace("_", " ");
  const html = `
    <div style="font-family:Arial,sans-serif;max-width:620px;margin:auto;color:#0f172a">
      <h2>BinaryGuard Secure Client Portal</h2>
      <p>Your one-time password is:</p>
      <div style="font-size:32px;font-weight:800;letter-spacing:6px;background:#f1f5f9;padding:18px 22px;border-radius:10px;text-align:center">
        ${escapeHtml(code)}
      </div>
      <p>This code expires in <strong>${otpExpiryMinutes} minutes</strong>.</p>
      <p style="color:#64748b;font-size:13px">Purpose: ${escapeHtml(purposeLabel)}</p>
      <p style="color:#64748b;font-size:13px">If you did not request this code, please ignore this email or contact BinaryGuard support.</p>
    </div>
  `;

  await sendGraphMail({
    to: email,
    subject: "BinaryGuard Portal OTP Code",
    html,
  });
}

app.get("/api/health", (_req: Request, res: Response) => {
  res.json({
    success: true,
    ok: true,
    service: "binaryguard-api",
    time: nowIso(),
  });
});

app.get("/health", (_req: Request, res: Response) => {
  res.json({
    success: true,
    ok: true,
    service: "binaryguard-api",
    time: nowIso(),
  });
});

app.post("/api/contact", async (req: Request, res: Response) => {
  try {
    const { name, company, email, phone, subject, message } = req.body;

    if (!name || !email || !subject || !message) {
      return res.status(400).json({
        success: false,
        message: "Required fields are missing.",
      });
    }

    const receiver = process.env.CONTACT_RECEIVER_EMAIL || process.env.SMTP_FROM!;

    await sendGraphMail({
      to: receiver,
      replyTo: normalizeEmail(email),
      subject: `Website Contact: ${escapeHtml(subject)}`,
      html: `
        <h2>New Contact Request</h2>
        <p><strong>Name:</strong> ${escapeHtml(name)}</p>
        <p><strong>Company:</strong> ${escapeHtml(company || "N/A")}</p>
        <p><strong>Email:</strong> ${escapeHtml(email)}</p>
        <p><strong>Phone:</strong> ${escapeHtml(phone || "N/A")}</p>
        <p><strong>Subject:</strong> ${escapeHtml(subject)}</p>
        <hr />
        <p><strong>Message:</strong></p>
        <p>${escapeHtml(message)}</p>
      `,
    });

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

app.post("/api/otp/request", async (req: Request, res: Response) => {
  try {
    const email = normalizeEmail(req.body.email);
    const purpose = normalizePurpose(req.body.purpose);

    if (!email || !email.includes("@")) {
      return res.status(400).json({ ok: false, message: "Valid email is required." });
    }

    if (!purpose) {
      return res.status(400).json({ ok: false, message: "Valid OTP purpose is required." });
    }

    const tenant = await findTenantByEmail(email);
    if (!tenant) {
      return res.status(403).json({ ok: false, message: "This corporate email domain is not approved." });
    }

    await supabase
      .from("otp_codes")
      .update({ used_at: nowIso() })
      .eq("email", email)
      .eq("purpose", purpose)
      .is("used_at", null);

    const code = generateOtp();
    const codeHash = hashOtp(email, purpose, code);

    const { error: insertError } = await supabase.from("otp_codes").insert({
      tenant_id: tenant.tenantId,
      email,
      purpose,
      code_hash: codeHash,
      expires_at: expiresAtIso(),
      attempt_count: 0,
      max_attempts: maxAttempts,
    });

    if (insertError) throw insertError;

    await sendOtpEmail(email, code, purpose);

    return res.json({
      ok: true,
      message: "OTP sent successfully.",
      expires_in_minutes: otpExpiryMinutes,
    });
  } catch (error) {
    console.error("OTP request error:", error);
    return res.status(500).json({
      ok: false,
      message: "Unable to send OTP. Please try again or contact support.",
    });
  }
});

app.post("/api/otp/verify", async (req: Request, res: Response) => {
  try {
    const email = normalizeEmail(req.body.email);
    const purpose = normalizePurpose(req.body.purpose);
    const code = String(req.body.code || "").trim();

    if (!email || !email.includes("@")) {
      return res.status(400).json({ ok: false, message: "Valid email is required." });
    }

    if (!purpose) {
      return res.status(400).json({ ok: false, message: "Valid OTP purpose is required." });
    }

    if (!/^\d{6}$/.test(code)) {
      return res.status(400).json({ ok: false, message: "Valid 6-digit OTP code is required." });
    }

    const { data, error: fetchError } = await supabase
      .from("otp_codes")
      .select("id, tenant_id, email, purpose, code_hash, expires_at, attempt_count, max_attempts, used_at")
      .eq("email", email)
      .eq("purpose", purpose)
      .is("used_at", null)
      .order("created_at", { ascending: false })
      .limit(1)
      .maybeSingle();

    if (fetchError) throw fetchError;

    const otpRow = data as OtpRow | null;

    if (!otpRow) {
      return res.status(404).json({ ok: false, message: "OTP not found or already used." });
    }

    if (new Date(otpRow.expires_at).getTime() < Date.now()) {
      await supabase.from("otp_codes").update({ used_at: nowIso() }).eq("id", otpRow.id);
      return res.status(410).json({ ok: false, message: "OTP has expired." });
    }

    if (otpRow.attempt_count >= otpRow.max_attempts) {
      await supabase.from("otp_codes").update({ used_at: nowIso() }).eq("id", otpRow.id);
      return res.status(429).json({ ok: false, message: "Maximum OTP attempts exceeded." });
    }

    const submittedHash = hashOtp(email, purpose, code);
    const storedHash = otpRow.code_hash;

    if (submittedHash.length !== storedHash.length) {
      return res.status(401).json({ ok: false, message: "Invalid OTP code." });
    }

    const valid = crypto.timingSafeEqual(
      Buffer.from(submittedHash, "hex"),
      Buffer.from(storedHash, "hex")
    );

    if (!valid) {
      const nextAttempts = otpRow.attempt_count + 1;
      const updatePayload: { attempt_count: number; used_at?: string } = {
        attempt_count: nextAttempts,
      };

      if (nextAttempts >= otpRow.max_attempts) {
        updatePayload.used_at = nowIso();
      }

      await supabase.from("otp_codes").update(updatePayload).eq("id", otpRow.id);

      return res.status(401).json({
        ok: false,
        message: nextAttempts >= otpRow.max_attempts ? "Maximum OTP attempts exceeded." : "Invalid OTP code.",
        attempts_remaining: Math.max(otpRow.max_attempts - nextAttempts, 0),
      });
    }

    await supabase.from("otp_codes").update({ used_at: nowIso() }).eq("id", otpRow.id);

    if (purpose === "registration") {
      await supabase
        .from("user_registration_requests")
        .update({
          status: "otp_verified",
          otp_verified_at: nowIso(),
        })
        .eq("corporate_email", email)
        .in("status", ["submitted", "otp_pending"]);
    }

    return res.json({
      ok: true,
      message: "OTP verified successfully.",
      tenant_id: otpRow.tenant_id,
    });
  } catch (error) {
    console.error("OTP verify error:", error);
    return res.status(500).json({
      ok: false,
      message: "Unable to verify OTP. Please try again or contact support.",
    });
  }
});

app.use((_req: Request, res: Response) => {
  res.status(404).json({ ok: false, message: "Endpoint not found." });
});

app.listen(PORT, "0.0.0.0", () => {
  console.log(`BinaryGuard API running on port ${PORT}`);
});
