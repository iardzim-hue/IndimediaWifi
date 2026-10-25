import { useState } from "react";
import type { FormEvent } from "react";
import { ArrowRight, CheckCircle2, CircleHelp, LockKeyhole, RadioTower, ShieldCheck, Signal, Wifi, X } from "lucide-react";

type LoginType = "customer" | "admin" | null;

function Brand({ compact = false }: { compact?: boolean }) {
  return <a className={`brand ${compact ? "brand-compact" : ""}`} href="#home" aria-label="IndimediaWiFi beranda">
    <img src="/indimediawifi-logo.svg" alt="" className="brand-logo" />
    <span className="brand-words"><strong>Indimedia<span>Wifi</span></strong>{!compact && <small>Internet murah untuk semua</small>}</span>
  </a>;
}

export default function App() {
  const [loginType, setLoginType] = useState<LoginType>(null);
  const [customerId, setCustomerId] = useState("");
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [notice, setNotice] = useState("");

  function openLogin(type: LoginType) {
    setLoginType(type);
    setNotice("");
  }

  function submitLogin(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();
    setNotice(loginType === "customer"
      ? "Form sudah siap. Verifikasi ID dan tampilan tagihan akan diaktifkan setelah autentikasi pelanggan terhubung ke Firebase."
      : "Login admin belum diaktifkan. Firebase Authentication dan hak akses admin harus dikonfigurasi sebelum panel dibuka.");
  }

  return <main className="site-shell" id="home">
    <header className="topbar">
      <Brand compact />
      <nav className="topnav" aria-label="Navigasi utama">
        <a href="#status">Status jaringan</a>
        <a href="#bantuan">Pusat bantuan</a>
      </nav>
      <button className="nav-login" onClick={() => openLogin("customer")}>Login <ArrowRight size={16} /></button>
    </header>

    <section className="hero">
      <div className="hero-copy">
        <div className="eyebrow"><i /> INTERNET UNTUK SEMUA</div>
        <h1>Koneksi yang<br />menghubungkan <span>semua.</span></h1>
        <p className="hero-description">Informasi jaringan, layanan pelanggan, dan bantuan koneksi IndimediaWiFi dalam satu portal yang mudah diakses.</p>
        <div className="hero-actions">
          <button className="button-primary" onClick={() => openLogin("customer")}>Login Pelanggan <ArrowRight size={18} /></button>
          <button className="button-secondary" onClick={() => openLogin("admin")}><ShieldCheck size={17} /> Login Admin</button>
        </div>
        <div className="trust-row"><span><CheckCircle2 size={15} /> Portal layanan terpadu</span><span><LockKeyhole size={15} /> Akses terkontrol</span></div>
      </div>
      <div className="hero-visual" aria-label="Ilustrasi jaringan WiFi">
        <div className="visual-grid" />
        <div className="orbit orbit-a" /><div className="orbit orbit-b" />
        <div className="network-core">
          <div className="signal-rings"><Wifi size={100} strokeWidth={2.4} /></div>
          <div className="router-device"><span /><span /><span /><span /></div>
          <div className="router-base" />
        </div>
        <div className="floating-card floating-top"><RadioTower size={19} /><span><b>Network center</b><small>Monitoring jaringan</small></span></div>
        <div className="floating-card floating-bottom"><i /><span><b>Terhubung lebih dekat</b><small>IndimediaWiFi service portal</small></span></div>
      </div>
    </section>

    <section className="status-section" id="status">
      <div className="section-heading"><div><span className="section-kicker">NETWORK OVERVIEW</span><h2>Status layanan pusat</h2><p>Informasi status akan mengikuti hasil monitoring server yang terhubung.</p></div><span className="demo-label"><Activity size={14} /> Monitoring belum aktif</span></div>
      <div className="status-card">
        <div className="status-symbol"><Signal size={25} /></div>
        <div className="status-main"><b>Server pusat IndimediaWiFi</b><span>Belum terhubung ke sistem monitoring</span></div>
        <div className="status-state"><i /> Belum diketahui</div>
        <div className="status-divider" />
        <div className="status-meta"><span>STATUS</span><b>Menunggu konfigurasi</b></div>
      </div>
      <p className="status-note"><CircleHelp size={15} /> Status belum real-time sampai monitoring server dikonfigurasi. Halaman ini tidak menganggap server aktif tanpa data nyata.</p>
    </section>

    <section className="help-section" id="bantuan">
      <div className="help-icon"><Wifi size={24} /></div>
      <div className="help-copy"><h3>Ada kendala koneksi?</h3><p>Periksa status server pusat terlebih dahulu. Jika status pusat normal tetapi koneksi rumah terganggu, tim teknis dapat melakukan pengecekan lanjutan di area pelanggan.</p></div>
      <button className="help-link" onClick={() => openLogin("customer")}>Buka portal pelanggan <ArrowRight size={16} /></button>
    </section>

    <footer className="footer"><Brand compact /><span>© {new Date().getFullYear()} IndimediaWiFi · Internet murah untuk semua</span><span><i /> Monitoring belum aktif</span></footer>

    {loginType && <div className="modal-backdrop" onMouseDown={event => { if (event.target === event.currentTarget) openLogin(null); }}>
      <section className="login-modal" role="dialog" aria-modal="true" aria-labelledby="login-title">
        <button className="modal-close" onClick={() => openLogin(null)} aria-label="Tutup dialog"><X size={19} /></button>
        <div className="modal-brand"><Brand /></div>
        <div className="modal-kicker">{loginType === "customer" ? "CUSTOMER PORTAL" : "ADMIN CONTROL CENTER"}</div>
        <h2 id="login-title">{loginType === "customer" ? "Login Pelanggan" : "Login Admin"}</h2>
        <p className="modal-description">{loginType === "customer" ? "Masukkan ID pelanggan untuk melanjutkan ke portal layanan Anda." : "Masuk ke pusat kontrol pengelolaan IndimediaWiFi."}</p>
        <form onSubmit={submitLogin}>
          {loginType === "customer" ? <><label htmlFor="customer-id">ID Pelanggan</label><input id="customer-id" autoFocus value={customerId} onChange={event => setCustomerId(event.target.value)} placeholder="Contoh: WIFI-00125" autoComplete="username" required /></> : <><label htmlFor="admin-email">Email admin</label><input id="admin-email" type="email" autoFocus value={email} onChange={event => setEmail(event.target.value)} placeholder="Email admin" autoComplete="username" required /><label htmlFor="admin-password">Password</label><input id="admin-password" type="password" value={password} onChange={event => setPassword(event.target.value)} placeholder="Password" autoComplete="current-password" required /></>}
          <button className="button-primary form-submit" type="submit">Login <ArrowRight size={17} /></button>
        </form>
        {notice && <div className="notice-box" role="status">{notice}</div>}
        <div className="modal-security"><LockKeyhole size={14} /> Data pelanggan dan panel admin belum dibuka sebelum verifikasi akses diaktifkan.</div>
      </section>
    </div>}
  </main>;
}
