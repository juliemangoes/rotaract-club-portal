import { useEffect, useState } from "react";
import { supabase, supabaseConfigured } from "../lib/supabase";
import Portal from "../Portal.jsx";

const CRAN = "#D41367", CRAN_DK = "#A50D50", AZURE = "#0067C8", GOLD = "#F7A81B";
const INK = "#2A1420", PAPER = "#F6F3F5", LINE = "#E9E1E6";
const DISPLAY = "'Archivo', system-ui, sans-serif";
const globalStyle = `
@import url('https://fonts.googleapis.com/css2?family=Archivo:wght@700;800;900&family=Public+Sans:wght@400;500;600;700&display=swap');
@keyframes rc-spin { to { transform: rotate(360deg); } }
@keyframes rc-blob-a { 0%, 100% { transform: translate(0, 0) scale(1); } 50% { transform: translate(6%, -8%) scale(1.12); } }
@keyframes rc-blob-b { 0%, 100% { transform: translate(0, 0) scale(1); } 50% { transform: translate(-8%, 6%) scale(1.08); } }
@keyframes rc-rise { from { opacity: 0; transform: translateY(18px); } to { opacity: 1; transform: translateY(0); } }
.rc-input { transition: box-shadow .15s, border-color .15s; }
.rc-input:focus { outline: none; border-color: ${CRAN} !important; box-shadow: 0 0 0 4px ${CRAN}1F; }
.rc-btn { transition: transform .15s, box-shadow .15s, opacity .15s; }
.rc-btn:active { transform: scale(.98); }
.rc-btn:not(:disabled):hover { transform: translateY(-1px); }
`;

function Shell({ children, eyebrow = "Rotaract", title = "Club Portal" }) {
  return (
    <div className="min-h-screen flex flex-col items-center relative overflow-hidden" style={{ fontFamily: "'Public Sans', system-ui, sans-serif", color: INK, background: INK }}>
      <style>{globalStyle}</style>
      {/* Ambient gradient + glow blobs */}
      <div className="absolute inset-0" style={{ background: `linear-gradient(155deg, ${CRAN} 0%, ${CRAN_DK} 45%, ${INK} 100%)` }} />
      <div className="absolute rounded-full" style={{ width: 420, height: 420, top: -140, left: -120, background: AZURE, opacity: 0.55, filter: "blur(90px)", animation: "rc-blob-a 14s ease-in-out infinite" }} />
      <div className="absolute rounded-full" style={{ width: 360, height: 360, top: 60, right: -140, background: GOLD, opacity: 0.35, filter: "blur(90px)", animation: "rc-blob-b 16s ease-in-out infinite" }} />
      <div className="absolute inset-0" style={{ backgroundImage: "radial-gradient(rgba(255,255,255,.14) 1px, transparent 1.5px)", backgroundSize: "22px 22px", opacity: 0.4 }} />

      <div className="w-full max-w-md px-6 pt-14 pb-8 relative">
        <div className="flex flex-col items-center text-center">
          <div className="rounded-full mb-4" style={{ width: 76, height: 76, background: "#fff", boxShadow: "0 10px 30px rgba(0,0,0,.35)", display: "flex", alignItems: "center", justifyContent: "center", overflow: "hidden" }}>
            <img src="/icons/pwa-512.png" alt="" width={102} height={102} style={{ animation: "rc-spin 70s linear infinite" }} />
          </div>
          <div className="uppercase font-bold" style={{ fontFamily: DISPLAY, fontSize: 11, letterSpacing: ".28em", color: "rgba(255,255,255,.75)" }}>{eyebrow}</div>
          <h1 className="font-black text-white" style={{ fontFamily: DISPLAY, fontSize: 34, textShadow: "0 2px 20px rgba(0,0,0,.25)" }}>{title}</h1>
        </div>
      </div>

      <div className="w-full max-w-md px-5 pb-10 relative flex-1" style={{ animation: "rc-rise .5s ease-out" }}>
        <div className="rounded-3xl p-6" style={{ background: "#fff", boxShadow: "0 20px 50px rgba(0,0,0,.25)", border: `1px solid ${LINE}` }}>
          {children}
        </div>
        <div className="text-center mt-6" style={{ fontSize: 11.5, color: "rgba(255,255,255,.55)" }}>Fellowship through service</div>
      </div>
    </div>
  );
}
const inputStyle = { border: `1.5px solid ${LINE}`, borderRadius: 12, padding: "12px 14px", fontSize: 15, width: "100%", background: "#fff" };
function Btn({ children, onClick, disabled, quiet }) {
  return (
    <button onClick={onClick} disabled={disabled} className="rc-btn w-full rounded-xl font-bold py-3.5 disabled:opacity-40"
      style={{
        background: quiet ? "#fff" : `linear-gradient(135deg, ${CRAN}, ${CRAN_DK})`,
        color: quiet ? INK : "#fff",
        border: quiet ? `1.5px solid ${LINE}` : "none",
        boxShadow: quiet ? "none" : `0 8px 20px ${CRAN}4D`,
        fontFamily: DISPLAY, fontSize: 15,
      }}>
      {children}
    </button>
  );
}

function AuthScreen() {
  const [mode, setMode] = useState("signin");
  const [name, setName] = useState("");
  const [email, setEmail] = useState("");
  const [pw, setPw] = useState("");
  const [msg, setMsg] = useState("");
  const [busy, setBusy] = useState(false);
  const go = async () => {
    setBusy(true); setMsg("");
    try {
      if (mode === "signup") {
        const { error } = await supabase.auth.signUp({
          email, password: pw, options: { data: { full_name: name.trim() } },
        });
        if (error) throw error;
        setMsg("Account created. If email confirmation is on, check your inbox, then sign in.");
        setMode("signin");
      } else {
        const { error } = await supabase.auth.signInWithPassword({ email, password: pw });
        if (error) throw error;
      }
    } catch (e) { setMsg(e.message); }
    setBusy(false);
  };
  return (
    <Shell>
      <div className="flex gap-2 mb-5 rounded-xl p-1" style={{ background: PAPER }}>
        {[["signin", "Sign in"], ["signup", "Create account"]].map(([m, label]) => (
          <button key={m} onClick={() => setMode(m)} className="rc-btn flex-1 rounded-lg font-bold py-2.5"
            style={{ fontFamily: DISPLAY, fontSize: 13.5, background: mode === m ? INK : "transparent", color: mode === m ? "#fff" : "#6B5A64", border: "none", boxShadow: mode === m ? "0 4px 12px rgba(42,20,32,.25)" : "none" }}>{label}</button>
        ))}
      </div>
      {mode === "signup" && <input className="rc-input" style={{ ...inputStyle, marginBottom: 12 }} placeholder="Full name (as the club knows you)" value={name} onChange={(e) => setName(e.target.value)} />}
      <input className="rc-input" style={{ ...inputStyle, marginBottom: 12 }} type="email" placeholder="Email" value={email} onChange={(e) => setEmail(e.target.value)} autoComplete="email" />
      <input className="rc-input" style={{ ...inputStyle, marginBottom: 16 }} type="password" placeholder="Password" value={pw} onChange={(e) => setPw(e.target.value)} autoComplete={mode === "signup" ? "new-password" : "current-password"} />
      <Btn onClick={go} disabled={busy || !email || !pw || (mode === "signup" && !name.trim())}>{busy ? "…" : mode === "signup" ? "Create account" : "Sign in"}</Btn>
      {msg && <p className="mt-3 text-center" style={{ fontSize: 13, color: "#8A5A00" }}>{msg}</p>}
      <p className="mt-5 text-center" style={{ fontSize: 12.5, color: "#9A8B93" }}>Use the email address your club has on file so your member record links automatically.</p>
    </Shell>
  );
}

function ClubSetup({ onDone, onSignOut }) {
  const [code, setCode] = useState("");
  const [msg, setMsg] = useState("");
  const [busy, setBusy] = useState(false);
  const go = async () => {
    setBusy(true); setMsg("");
    try {
      const { error } = await supabase.rpc("join_club", { p_code: code.trim().toLowerCase() });
      if (error) throw error;
      onDone();
    } catch (e) { setMsg(e.message); }
    setBusy(false);
  };
  return (
    <Shell>
      <div className="mb-5 font-black text-center" style={{ fontFamily: DISPLAY, fontSize: 18 }}>Join your club</div>
      <input className="rc-input" style={{ ...inputStyle, marginBottom: 16, letterSpacing: ".12em", textTransform: "lowercase" }} placeholder="Club invite code" value={code} onChange={(e) => setCode(e.target.value)} />
      <Btn onClick={go} disabled={busy || !code.trim()}>{busy ? "…" : "Join club"}</Btn>
      <p className="mt-3 text-center" style={{ fontSize: 12.5, color: "#9A8B93" }}>Ask your club administrator for the invite code. The first person to join with a new club's code becomes its President; after that, the board approves each new member.</p>
      {msg && <p className="mt-3 text-center" style={{ fontSize: 13, color: "#C22F2F" }}>{msg}</p>}
      <div className="mt-6"><Btn quiet onClick={onSignOut}>Sign out</Btn></div>
    </Shell>
  );
}

export default function AuthGate() {
  const [session, setSession] = useState(null);
  const [membership, setMembership] = useState(undefined); // undefined = loading
  const [ready, setReady] = useState(false);

  useEffect(() => {
    if (!supabaseConfigured) { setReady(true); return; }
    supabase.auth.getSession().then(({ data }) => { setSession(data.session); setReady(true); });
    const { data: sub } = supabase.auth.onAuthStateChange((_e, s) => setSession(s));
    return () => sub.subscription.unsubscribe();
  }, []);

  const loadMembership = async () => {
    setMembership(undefined);
    const { data } = await supabase
      .from("club_members")
      .select("club_id, role, clubs(name, join_code)")
      .limit(1);
    setMembership(data && data.length ? data[0] : null);
  };
  useEffect(() => { if (session) loadMembership(); }, [session?.user?.id]);

  if (!supabaseConfigured) {
    // No backend configured: run in local demo mode (data stays on this device).
    return <Portal clubId="demo" demo />;
  }
  if (!ready) return null;
  if (!session) return <AuthScreen />;
  if (membership === undefined) return <Shell><p style={{ textAlign: "center", color: "#8A7580" }}>Loading your club…</p></Shell>;
  if (!membership) return <ClubSetup onDone={loadMembership} onSignOut={() => supabase.auth.signOut()} />;

  return (
    <Portal
      key={membership.club_id}
      clubId={membership.club_id}
      demo={false}
      userEmail={session.user.email}
      userName={session.user.user_metadata?.full_name || session.user.email}
      clubName={membership.clubs?.name || "Rotaract Club"}
      joinCode={membership.clubs?.join_code || ""}
      onSignOutAuth={() => supabase.auth.signOut()}
    />
  );
}
