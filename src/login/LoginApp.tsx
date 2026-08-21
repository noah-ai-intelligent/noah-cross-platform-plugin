import React, { useState, useEffect } from "react";
import { REDIRECT_URI } from "../config";
import { request } from "../api";
import { saveTokens } from "../tokenStorage";
import type { TokenPair } from "../tokenStorage";
import { isSignedIn } from "../auth";
import "../styles/tailwind.css";

type Step = "email" | "code" | "org";

interface CaptchaChallenge {
  captcha_id: string;
  image: string;
}

interface OTPRequestResponse {
  dev_code?: string;
  captcha_required: boolean;
}

type Provider = "google" | "microsoft";

export function LoginApp({
  onSuccess,
  onSsoClick
}: {
  onSuccess?: (tokens: TokenPair, orgId?: string) => void;
  onSsoClick?: () => Promise<void> | void;
} = {}) {
  const [step, setStep] = useState<Step>("email");
  const [email, setEmail] = useState("");
  const [code, setCode] = useState("");
  const [orgs, setOrgs] = useState<{ id: string; name: string }[]>([]);
  const [selectedOrg, setSelectedOrg] = useState<string>("");

  useEffect(() => {
    void isSignedIn().then(signedIn => {
      if (signedIn) {
        window.location.href = "/taskpane";
      }
    });
  }, []);

  const [captcha, setCaptcha] = useState<CaptchaChallenge | null>(null);
  const [captchaAnswer, setCaptchaAnswer] = useState("");
  const [captchaRequired, setCaptchaRequired] = useState(true);

  const [error, setError] = useState<string | null>(null);
  const [submitting, setSubmitting] = useState(false);
  const [resendIn, setResendIn] = useState(0);

  // Check for returning SSO redirect
  useEffect(() => {
    const params = new URLSearchParams(window.location.search);
    const authCode = params.get("code");
    const state = params.get("state");
    const authError = params.get("error");

    if (authError) {
      setError(`Sign-in was cancelled or failed: ${authError}`);
      return;
    }

    if (authCode && state) {
      void completeSso(authCode, state);
    }
  }, []);

  async function loadCaptcha() {
    setCaptchaAnswer("");
    try {
      const res = await request<CaptchaChallenge>("/auth/captcha");
      setCaptcha(res);
    } catch {
      setCaptcha(null);
    }
  }

  useEffect(() => {
    if (captchaRequired && !captcha) {
      void loadCaptcha();
    }
  }, [captchaRequired, captcha]);

  useEffect(() => {
    if (resendIn <= 0) return;
    const id = window.setInterval(() => setResendIn((s) => Math.max(0, s - 1)), 1000);
    return () => window.clearInterval(id);
  }, [resendIn]);

  function extractErr(err: any): string {
    return err?.message || "Something went wrong. Try again.";
  }

  async function requestOtp(e?: React.FormEvent) {
    if (e) e.preventDefault();
    if (captchaRequired && (!captcha || !captchaAnswer.trim())) {
      setError("Enter the characters shown in the captcha.");
      return;
    }
    setSubmitting(true);
    setError(null);
    try {
      const body: any = { email };
      if (captchaRequired && captcha) {
        body.captcha_id = captcha.captcha_id;
        body.captcha_answer = captchaAnswer.trim();
      }

      const data = await request<OTPRequestResponse>("/auth/request-otp", {
        method: "POST",
        body: JSON.stringify(body),
      });

      setStep("code");
      setCode("");
      setResendIn(30);
      setCaptchaRequired(data.captcha_required);
      if (data.captcha_required) {
        void loadCaptcha();
      }
    } catch (err: any) {
      if (err.status === 400 && String(err.message).toLowerCase().includes("captcha")) {
        setCaptchaRequired(true);
        void loadCaptcha();
        setError("Invalid or missing captcha.");
      } else if (err.status === 429) {
        setResendIn(30);
        setError("Too many requests. Please try again later.");
      } else {
        setError(extractErr(err));
      }
    } finally {
      setSubmitting(false);
    }
  }

  // Store pending tokens if we need to select an org
  const [pendingTokens, setPendingTokens] = useState<TokenPair | null>(null);

  async function handleSuccess(tokens: TokenPair, orgId?: string) {
    if (onSuccess) {
      await saveTokens(tokens);
      if (orgId) localStorage.setItem("noah_selected_org", orgId);
      onSuccess(tokens, orgId);
    } else if (typeof Office !== "undefined" && Office.context?.ui) {
      Office.context.ui.messageParent(JSON.stringify({ ok: true, tokens, orgId }));
    } else {
      // Fallback for direct browser testing
      await saveTokens(tokens);
      if (orgId) localStorage.setItem("noah_selected_org", orgId);
      window.location.href = "/taskpane";
    }
  }

  async function processTokens(tokens: TokenPair) {
    try {
      // Temporarily use the tokens to fetch /auth/me
      const res = await request<any>("/auth/me", {
        headers: { Authorization: `Bearer ${tokens.access_token}` }
      });
      const userOrgs = res.organizations || [];
      if (userOrgs.length > 1) {
        setOrgs(userOrgs.map((o: any) => ({ id: o.organization_id, name: o.organization_name })));
        setSelectedOrg(userOrgs[0].organization_id);
        setPendingTokens(tokens);
        setStep("org");
      } else {
        await handleSuccess(tokens, userOrgs[0]?.organization_id);
      }
    } catch {
      await handleSuccess(tokens);
    }
  }

  async function verifyOtp(e: React.FormEvent) {
    e.preventDefault();
    if (code.length < 6) return;
    setSubmitting(true);
    setError(null);
    try {
      const tokens = await request<TokenPair>("/auth/verify-otp", {
        method: "POST",
        body: JSON.stringify({ email, code }),
      });
      await processTokens(tokens);
    } catch (err: any) {
      setError(extractErr(err));
    } finally {
      setSubmitting(false);
    }
  }

  async function startSso(provider: Provider) {
    if (onSsoClick) {
      try {
        setSubmitting(true);
        await onSsoClick();
      } catch (err: any) {
        setError(extractErr(err));
      } finally {
        setSubmitting(false);
      }
      return;
    }
    setSubmitting(true);
    setError(null);
    try {
      sessionStorage.setItem("noah.sso.provider", provider);
      const data = await request<{ authorization_url: string }>(`/auth/sso/${provider}/start`, {
        method: "POST",
        body: JSON.stringify({ redirect_uri: REDIRECT_URI }),
      });
      window.location.href = data.authorization_url;
    } catch (err: any) {
      setError(extractErr(err));
      setSubmitting(false);
    }
  }

  async function completeSso(authCode: string, state: string) {
    const provider = sessionStorage.getItem("noah.sso.provider") as Provider | null;
    if (!provider) {
      setError("Sign-in session expired. Try again.");
      return;
    }
    setError(null);
    setSubmitting(true);
    try {
      const tokens = await request<TokenPair>(`/auth/sso/${provider}/callback`, {
        method: "POST",
        body: JSON.stringify({ code: authCode, state, redirect_uri: REDIRECT_URI }),
      });
      await processTokens(tokens);
    } catch (err: any) {
      setError(extractErr(err));
      setSubmitting(false);
    }
  }

  const GoogleIcon = () => (
    <svg width="18" height="18" viewBox="0 0 18 18" aria-hidden="true">
      <path fill="#4285F4" d="M17.64 9.2c0-.637-.057-1.251-.164-1.84H9v3.481h4.844a4.14 4.14 0 01-1.796 2.717v2.258h2.908c1.702-1.567 2.684-3.874 2.684-6.615z" />
      <path fill="#34A853" d="M9 18c2.43 0 4.467-.806 5.956-2.184l-2.908-2.258c-.806.54-1.837.86-3.048.86-2.344 0-4.328-1.584-5.036-3.711H.957v2.332A8.997 8.997 0 009 18z" />
      <path fill="#FBBC05" d="M3.964 10.707A5.41 5.41 0 013.682 9c0-.593.102-1.17.282-1.707V4.961H.957A8.997 8.997 0 000 9c0 1.452.348 2.827.957 4.04l3.007-2.333z" />
      <path fill="#EA4335" d="M9 3.58c1.321 0 2.508.454 3.44 1.345l2.582-2.58C13.463.891 11.426 0 9 0A8.997 8.997 0 00.957 4.961L3.964 7.293C4.672 5.166 6.656 3.58 9 3.58z" />
    </svg>
  );

  const MicrosoftIcon = () => (
    <svg width="18" height="18" viewBox="0 0 18 18" aria-hidden="true">
      <rect x="1" y="1" width="7.5" height="7.5" fill="#F25022" />
      <rect x="9.5" y="1" width="7.5" height="7.5" fill="#7FBA00" />
      <rect x="1" y="9.5" width="7.5" height="7.5" fill="#00A4EF" />
      <rect x="9.5" y="9.5" width="7.5" height="7.5" fill="#FFB900" />
    </svg>
  );

  const inputClass = "w-full box-border h-10 px-3 bg-transparent border border-border rounded-lg text-ink text-[14px] mb-3 outline-none";

  const btnBaseClass = "flex items-center justify-center gap-3 w-full h-11 rounded-lg border border-border bg-canvas text-ink text-[14px] font-medium cursor-pointer transition-colors hover:bg-surface hover:border-border-strong disabled:opacity-50 disabled:cursor-not-allowed";
  const primaryBtnClass = `${btnBaseClass} !bg-ink !text-canvas !border-none mt-2 hover:!bg-black`;

  return (
    <div className="bg-canvas border border-border rounded-xl w-full max-w-[400px] overflow-hidden">
      <div className="relative p-4 text-center">
        <div className="absolute top-0 left-1/2 -translate-x-1/2 h-0.5 w-12 bg-accent"></div>
        <h1 className="text-[20px] font-semibold m-0">Sign in</h1>
        <p className="text-ink-muted text-[14px] mt-1 mb-0">Just one question away from your data.</p>
      </div>

      <div className="p-4 pt-4 px-8 pb-8 flex flex-col gap-2">
        {step === "email" && (
          <form onSubmit={requestOtp} className="flex flex-col">
            <input
              type="email"
              placeholder="you@company.com"
              value={email}
              onChange={(e) => setEmail(e.target.value)}
              required
              disabled={submitting}
              className={inputClass}
            />

            {captchaRequired && (
              <div className="mb-3">
                <div className="text-[12px] text-ink-muted mb-1 uppercase">Verify you're human</div>
                <div className="flex flex-col gap-2">
                  <div className="flex gap-2 items-center">
                    <div className="h-10 w-[120px] bg-border rounded-lg overflow-hidden">
                      {captcha ? <img src={captcha.image} alt="captcha" className="h-full w-full object-contain" /> : null}
                    </div>
                    <button type="button" onClick={loadCaptcha} className="bg-transparent border border-zinc-200 rounded-lg px-3 h-10 cursor-pointer text-[12px] text-ink-secondary">
                      Refresh
                    </button>
                  </div>
                  <input
                    type="text"
                    placeholder="Type chars"
                    value={captchaAnswer}
                    onChange={(e) => setCaptchaAnswer(e.target.value)}
                    required
                    className={`${inputClass} !mb-0 w-full`}
                  />
                </div>
              </div>
            )}

            {error && <div className="text-danger text-[13px] mb-3">{error}</div>}

            <button type="submit" disabled={submitting} className={primaryBtnClass}>
              {submitting ? "Sending..." : "Send one-time code"}
            </button>

            <div className="flex items-center my-5">
              <div className="flex-1 h-px bg-border"></div>
              <span className="text-[12px] text-ink-muted px-2">or with SSO</span>
              <div className="flex-1 h-px bg-border"></div>
            </div>

            <button type="button" onClick={() => startSso("google")} disabled={submitting} className={`${btnBaseClass} px-4 text-[#3c4043] border-[#dadce0] bg-white`}>
              <GoogleIcon /> <span className="flex-1 text-center pr-[18px]">Continue with Google</span>
            </button>
            <div className="h-3"></div>
            <button type="button" onClick={() => startSso("microsoft")} disabled={submitting} className={`${btnBaseClass} px-4 text-[#5e5e5e] border-[#8c8c8c] bg-white`}>
              <MicrosoftIcon /> <span className="flex-1 text-center pr-[18px]">Continue with Microsoft</span>
            </button>
          </form>
        )}

        {step === "code" && (
          <form onSubmit={verifyOtp} className="flex flex-col">
            <div className="text-[14px] text-ink-muted mb-5 text-center leading-relaxed">
              Code sent to <strong className="text-ink">{email}</strong>
              <div
                className="text-emerald cursor-pointer mt-1 text-[13px] font-medium"
                onClick={() => setStep("email")}
              >
                Change email address
              </div>
            </div>

            <input
              type="text"
              placeholder="000000"
              value={code}
              onChange={(e) => setCode(e.target.value.replace(/[^0-9a-zA-Z]/g, ''))}
              required
              disabled={submitting}
              maxLength={6}
              className={`${inputClass} tracking-[16px] pl-[28px] text-center text-[24px] !h-14 font-semibold font-mono`}
            />

            {error && <div className="text-danger text-[13px] mb-3 text-center">{error}</div>}

            <div className="text-[13px] text-ink-muted mb-5 text-center">
              Check your inbox. <span className="ml-[2px]"></span>
              <button
                type="button"
                onClick={() => requestOtp()}
                disabled={resendIn > 0 || submitting}
                className={`border-none bg-transparent text-[13px] p-0 font-medium ${resendIn > 0 ? "text-ink-muted cursor-default" : "text-emerald cursor-pointer"}`}
              >
                {resendIn > 0 ? `Resend in ${resendIn}s` : "Resend code"}
              </button>
            </div>

            <button type="submit" disabled={submitting || code.length < 6} className={primaryBtnClass}>
              {submitting ? "Verifying..." : "Verify & sign in"}
            </button>
          </form>
        )}

        {step === "org" && (
          <div className="flex flex-col">
            <div className="text-[14px] text-ink-muted mb-4">
              Your account belongs to multiple organizations. Pick the one you want to work in.
            </div>

            <label className="text-[12px] text-ink-muted mb-1 font-semibold">Organization</label>
            <select
              value={selectedOrg}
              onChange={(e) => setSelectedOrg(e.target.value)}
              className={`${inputClass} p-2 appearance-auto`}
            >
              {orgs.map((o) => (
                <option key={o.id} value={o.id}>
                  {o.name}
                </option>
              ))}
            </select>

            <button
              onClick={() => {
                if (pendingTokens && selectedOrg) {
                  handleSuccess(pendingTokens, selectedOrg);
                }
              }}
              className={`${primaryBtnClass} mt-4`}
            >
              Continue
            </button>
          </div>
        )}
      </div>
    </div>
  );
}
