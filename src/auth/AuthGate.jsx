import { useEffect, useState } from "react";
import { supabase, supabaseConfigured } from "../lib/supabase";
import Portal from "../Portal.jsx";

const CRAN = "#D41367", INK = "#2A1420", PAPER = "#F6F3F5", LINE = "#E9E1E6";
const DISPLAY = "'Archivo', system-ui, sans-serif";
const fontLink = `@import url('https://fonts.googleapis.com/css2?family=Archivo:wght@700;800;900&family=Public+Sans:wght@400;500;600;700&display=swap');`;

function Shell({ children }) {
  return (
    <div className="min-h-screen flex flex-col items-center" style={{ background: PAPER, fontFamily: "'Public Sans', system-ui, sans-serif", color: INK }}>
      <style>{fontLink}</style>
      <div className="w-full" style={{ background: `linear-gradient(150deg, ${CRAN}, #A50D50)` }}>
        <div className="max-w-md mx-auto px-6 pt-12 pb-9">
          <div className="uppercase font-bold" style={{ fontFamily: DISPLAY, fontSize: 11, letterSpacing: ".22em", color: "rgba(255,255,255,.8)" }}>Rotaract</div>
          <h1 className="font-black text-white" style={{ fontFamily: DISPLAY, fontSize: 32 }}>Club Portal</h1>
        </div>
      </div>
      <div className="w-full max-w-md px-5 py-6">{children}</div>
    </div>
  );
}
const inputStyle = { border: `1.5px solid ${LINE}`, borderRadius: 12, padding: "12px 14px", fontSize: 15, width: "100%", background: "#fff" };
function Btn({ children, onClick, disabled, quiet }) {
  return (
    <button onClick={onClick} disabled={disabled}
      className="w-full rounded-xl font-bold py-3.5 active:opacity-80 disabled:opacity-40"
      style={{ background: quiet ? "#fff" : CRAN, color: quiet ? INK : "#fff", border: quiet ? `1.5px solid ${LINE}` : "none", fontFamily: DISPLAY, fontSize: 15 }}>
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
      <div className="flex gap-2 mb-5">
        {[["signin", "Sign in"], ["signup", "Create account"]].map(([m, label]) => (
          <button key={m} onClick={() => setMode(m)} className="flex-1 rounded-xl font-bold py-2.5"
            style={{ fontFamily: DISPLAY, fontSize: 14, background: mode === m ? INK : "#fff", color: mode === m ? "#fff" : "#6B5A64", border: `1.5px solid ${mode === m ? INK : LINE}` }}>{label}</button>
        ))}
      </div>
      {mode === "signup" && <input style={{ ...inputStyle, marginBottom: 12 }} placeholder="Full name (as the club knows you)" value={name} onChange={(e) => setName(e.target.value)} />}
      <input style={{ ...inputStyle, marginBottom: 12 }} type="email" placeholder="Email" value={email} onChange={(e) => setEmail(e.target.value)} autoComplete="email" />
      <input style={{ ...inputStyle, marginBottom: 16 }} type="password" placeholder="Password" value={pw} onChange={(e) => setPw(e.target.value)} autoComplete={mode === "signup" ? "new-password" : "current-password"} />
      <Btn onClick={go} disabled={busy || !email || !pw || (mode === "signup" && !name.trim())}>{busy ? "…" : mode === "signup" ? "Create account" : "Sign in"}</Btn>
      {msg && <p className="mt-3 text-center" style={{ fontSize: 13, color: "#8A5A00" }}>{msg}</p>}
      <p className="mt-5 text-center" style={{ fontSize: 12.5, color: "#9A8B93" }}>Use the email address your club has on file so your member record links automatically.</p>
    </Shell>
  );
}

function ClubSetup({ onDone, onSignOut }) {
  const [mode, setMode] = useState("join");
  const [name, setName] = useState("");
  const [code, setCode] = useState("");
  const [msg, setMsg] = useState("");
  const [busy, setBusy] = useState(false);
  const go = async () => {
    setBusy(true); setMsg("");
    try {
      if (mode === "create") {
        const { error } = await supabase.rpc("create_club", { p_name: name.trim() });
        if (error) throw error;
      } else {
        const { error } = await supabase.rpc("join_club", { p_code: code.trim().toLowerCase() });
        if (error) throw error;
      }
      onDone();
    } catch (e) { setMsg(e.message); }
    setBusy(false);
  };
  return (
    <Shell>
      <div className="flex gap-2 mb-5">
        {[["join", "Join a club"], ["create", "Start a club"]].map(([m, label]) => (
          <button key={m} onClick={() => setMode(m)} className="flex-1 rounded-xl font-bold py-2.5"
            style={{ fontFamily: DISPLAY, fontSize: 14, background: mode === m ? INK : "#fff", color: mode === m ? "#fff" : "#6B5A64", border: `1.5px solid ${mode === m ? INK : LINE}` }}>{label}</button>
        ))}
      </div>
      {mode === "join" ? (
        <>
          <input style={{ ...inputStyle, marginBottom: 16, letterSpacing: ".12em", textTransform: "lowercase" }} placeholder="Club invite code" value={code} onChange={(e) => setCode(e.target.value)} />
          <Btn onClick={go} disabled={busy || !code.trim()}>{busy ? "…" : "Join club"}</Btn>
          <p className="mt-3 text-center" style={{ fontSize: 12.5, color: "#9A8B93" }}>Ask an officer for the invite code (Club settings → invite code). After joining, the board approves your membership.</p>
        </>
      ) : (
        <>
          <input style={{ ...inputStyle, marginBottom: 16 }} placeholder="Club name (e.g. Rotaract Club of Corozal)" value={name} onChange={(e) => setName(e.target.value)} />
          <Btn onClick={go} disabled={busy || !name.trim()}>{busy ? "…" : "Create club — you become President"}</Btn>
        </>
      )}
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
