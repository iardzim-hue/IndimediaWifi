import { useRef, useState } from "react";
import type { FormEvent, ChangeEvent } from "react";
import * as XLSX from "xlsx";
import { Activity, ArrowDownToLine, ArrowRight, CheckCircle2, CircleHelp, FileSpreadsheet, HardDriveUpload, LockKeyhole, LogOut, RadioTower, Search, ShieldCheck, Signal, Wifi, X, Users, ReceiptText, LayoutDashboard, Settings2 } from "lucide-react";

type LoginType = "customer" | "admin" | null;
type Customer = { legacyNo: number; name: string; status?: string };
type Bill = { legacyNo: number; customerName: string; billingMonth: string; year?: number; amount?: number | null; amountRaw?: string; paymentNote?: string; status?: string };
type BackupData = { sourceFile?: string; sourceSheet?: string; year?: number; customers: Customer[]; bills: Bill[]; exportedAt?: string };

const demoCustomers: Customer[] = [
  { legacyNo: 1, name: "Riski", status: "active" },
  { legacyNo: 2, name: "Bik Saroh", status: "off" },
  { legacyNo: 3, name: "Mas Dwi", status: "active" },
  { legacyNo: 4, name: "Danu", status: "active" },
  { legacyNo: 5, name: "Dek Fila", status: "active" },
];

function Brand({ compact = false }: { compact?: boolean }) {
  return <a className={`brand ${compact ? "brand-compact" : ""}`} href="#home" aria-label="IndimediaWiFi beranda">
    <img src="/indimediawifi-logo.svg" alt="" className="brand-logo" />
    <span className="brand-words"><strong>Indimedia<span>Wifi</span></strong>{!compact && <small>Internet murah untuk semua</small>}</span>
  </a>;
}

function downloadExcel(data: BackupData) {
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(data.customers.map(c => ({ No: c.legacyNo, "Nama Pelanggan": c.name, Status: c.status || "active" }))), "Data Pelanggan");
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(data.bills.map(b => ({ "No Pelanggan": b.legacyNo, "Nama Pelanggan": b.customerName, Bulan: b.billingMonth, Tahun: b.year || data.year || 2026, Nominal: b.amount ?? b.amountRaw ?? "", Keterangan: b.paymentNote || "", Status: b.status || "unpaid" }))), "Tagihan Bulanan");
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet([{ File: data.sourceFile || "IndimediaWifi", Sheet: data.sourceSheet || "Backup Admin", Tahun: data.year || 2026, "Waktu Backup": new Date().toLocaleString("id-ID") }]), "Info Backup");
  XLSX.writeFile(wb, "indimediawifi-backup-" + new Date().toISOString().slice(0,10) + ".xlsx");
}

export default function App() {
  const [loginType, setLoginType] = useState<LoginType>(null);
  const [customerId, setCustomerId] = useState("");
  const [username, setUsername] = useState("");
  const [password, setPassword] = useState("");
  const [notice, setNotice] = useState("");
  const [isAdmin, setIsAdmin] = useState(false);
  const [activeMenu, setActiveMenu] = useState("dashboard");
  const [search, setSearch] = useState("");
  const [customers, setCustomers] = useState<Customer[]>(demoCustomers);
  const [bills, setBills] = useState<Bill[]>([]);
  const [backupNotice, setBackupNotice] = useState("");
  const [billMonth, setBillMonth] = useState("Semua");
  const [billYear, setBillYear] = useState("2026");
  const fileRef = useRef<HTMLInputElement>(null);

  function openLogin(type: LoginType) {
    setLoginType(type); setNotice(""); setUsername(""); setPassword("");
  }

  function submitLogin(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();
    if (loginType === "admin") {
      // Temporary preview-only credentials. Replace with Firebase Authentication before production.
      if (username.trim() === "iamardi" && password === "iamardi123") {
        setIsAdmin(true); setLoginType(null); setActiveMenu("dashboard"); setNotice("");
      } else setNotice("Username atau password tidak cocok.");
      return;
    }
    setNotice("Login pelanggan belum diaktifkan. Akun pelanggan akan diverifikasi melalui Firebase setelah autentikasi dikonfigurasi.");
  }

  function logout() { setIsAdmin(false); setActiveMenu("dashboard"); }

  function handleBackupUpload(event: ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0]; if (!file) return;
    const reader = new FileReader();
    reader.onload = () => {
      try {
        const wb = XLSX.read(reader.result, { type: "array" });
        const cs = wb.Sheets["Data Pelanggan"], bs = wb.Sheets["Tagihan Bulanan"];
        if (!cs || !bs) throw new Error("Workbook harus memiliki sheet Data Pelanggan dan Tagihan Bulanan.");
        const cr = XLSX.utils.sheet_to_json<Record<string, unknown>>(cs);
        const br = XLSX.utils.sheet_to_json<Record<string, unknown>>(bs);
        const nc: Customer[] = cr.map(r => ({ legacyNo:Number(r["No"]), name:String(r["Nama Pelanggan"] || ""), status:String(r["Status"] || "active") })).filter(c => c.name && Number.isFinite(c.legacyNo));
        const nb: Bill[] = br.map(r => ({ legacyNo:Number(r["No Pelanggan"]), customerName:String(r["Nama Pelanggan"] || ""), billingMonth:String(r["Bulan"] || ""), year:Number(r["Tahun"] || 2026), amount:typeof r["Nominal"]==="number"?r["Nominal"]:null, amountRaw:String(r["Nominal"] ?? ""), paymentNote:String(r["Keterangan"] || ""), status:String(r["Status"] || "unpaid") })).filter(b => Number.isFinite(b.legacyNo) && b.billingMonth);
        if (!nc.length) throw new Error("Tidak ada data pelanggan valid.");
        setCustomers(nc); setBills(nb); setBackupNotice("Excel dimuat ke pratinjau: " + nc.length + " pelanggan, " + nb.length + " tagihan. Belum disimpan ke Firestore.");
      } catch (e) { setBackupNotice(e instanceof Error ? e.message : "File Excel gagal dibaca."); }
      event.target.value = "";
    };
    reader.readAsArrayBuffer(file);
  }

  const filteredCustomers = customers.filter(c => c.name.toLowerCase().includes(search.toLowerCase()) || String(c.legacyNo).includes(search));
  const paidCount = bills.filter(b => b.status === "paid").length;

  if (isAdmin) return <main className="admin-shell">
    <aside className="admin-sidebar">
      <Brand />
      <div className="admin-side-label">MENU UTAMA</div>
      {([
        { key: "dashboard", label: "Dashboard", Icon: LayoutDashboard },
        { key: "customers", label: "Data Pelanggan", Icon: Users },
        { key: "billing", label: "Tagihan Bulanan", Icon: ReceiptText },
        { key: "backup", label: "Backup & Restore", Icon: HardDriveUpload },
        { key: "settings", label: "Pengaturan", Icon: Settings2 },
      ]).map(({ key, label, Icon }) => <button key={key} className={activeMenu === key ? "admin-nav active" : "admin-nav"} onClick={() => { setActiveMenu(key); setBackupNotice(""); }}><Icon size={17} />{label}</button>)}
      <div className="admin-side-bottom"><span className="admin-avatar">IA</span><span><b>iamardi</b><small>Administrator demo</small></span><button aria-label="Keluar" onClick={logout}><LogOut size={16}/></button></div>
    </aside>
    <section className="admin-main">
      <header className="admin-topbar"><div><span className="admin-breadcrumb">IndimediaWifi /</span><b>{({dashboard:"Dashboard",customers:"Data Pelanggan",billing:"Tagihan Bulanan",backup:"Backup & Restore",settings:"Pengaturan"} as Record<string,string>)[activeMenu]}</b></div><button className="admin-top-action" onClick={logout}><LogOut size={15}/> Keluar</button></header>
      {activeMenu === "dashboard" && <div className="admin-content">
        <div className="admin-page-heading"><div><span className="section-kicker">ADMIN CONTROL CENTER</span><h1>Dashboard</h1><p>Ringkasan administrasi pelanggan dan tagihan IndimediaWifi.</p></div><button className="excel-button" onClick={() => { setActiveMenu("backup"); setBackupNotice(""); }}><HardDriveUpload size={16}/> Kelola Backup</button></div>
        <div className="admin-stat-grid"><div className="admin-stat"><span><Users size={17}/> Total Pelanggan</span><strong>{customers.length}</strong><small>Data yang tampil di panel ini</small></div><div className="admin-stat"><span><ReceiptText size={17}/> Catatan Tagihan</span><strong>{bills.length}</strong><small>Data tagihan pada sesi ini</small></div><div className="admin-stat"><span><CheckCircle2 size={17}/> Status Lunas</span><strong>{paidCount}</strong><small>Dari data backup yang dimuat</small></div><div className="admin-stat"><span><Activity size={17}/> Database</span><strong className="pending-text">Belum terhubung</strong><small>Firestore perlu dikonfigurasi</small></div></div>
        <section className="excel-panel"><div className="excel-panel-heading"><div><FileSpreadsheet size={19}/><span><b>Ringkasan Data Pelanggan</b><small>Pratinjau gaya lembar kerja Excel</small></span></div><button className="excel-button" onClick={() => downloadExcel({ sourceFile: "IndimediaWifi", sourceSheet: "Data Pelanggan", year: 2026, customers, bills })}><ArrowDownToLine size={15}/> Download Backup</button></div>
          <div className="excel-table-wrap"><table className="excel-table"><thead><tr><th>No</th><th>Nama Pelanggan</th><th>Status Layanan</th><th>Tagihan Tersimpan</th></tr></thead><tbody>{filteredCustomers.slice(0,8).map(c => <tr key={c.legacyNo}><td>{c.legacyNo}</td><td>{c.name}</td><td><span className={c.status === "off" ? "sheet-status off" : "sheet-status"}>{c.status === "off" ? "OFF" : "Aktif"}</span></td><td>{bills.filter(b=>b.legacyNo===c.legacyNo).length ? bills.filter(b=>b.legacyNo===c.legacyNo).length + " bulan" : "—"}</td></tr>)}</tbody></table></div>
          <p className="admin-hint">Panel ini masih mode pratinjau. Data contoh bukan database live dan belum ditulis ke Firebase.</p>
        </section>
      </div>}
      {activeMenu === "customers" && <div className="admin-content"><div className="admin-page-heading"><div><span className="section-kicker">CUSTOMER MANAGEMENT</span><h1>Data Pelanggan</h1><p>Daftar pelanggan dengan format tabel menyerupai lembar Excel.</p></div><button className="excel-button" onClick={() => downloadJson({sourceSheet:"Data Pelanggan",year:2026,customers,bills})}><ArrowDownToLine size={15}/> Backup</button></div><div className="excel-panel"><div className="table-toolbar"><div className="admin-search"><Search size={16}/><input value={search} onChange={e=>setSearch(e.target.value)} placeholder="Cari nomor atau nama pelanggan"/></div><span>{filteredCustomers.length} pelanggan</span></div><div className="excel-table-wrap"><table className="excel-table"><thead><tr><th>No</th><th>Nama Pelanggan</th><th>Status</th><th>ID Legacy</th></tr></thead><tbody>{filteredCustomers.map(c=><tr key={c.legacyNo}><td>{c.legacyNo}</td><td>{c.name}</td><td><span className={c.status==="off"?"sheet-status off":"sheet-status"}>{c.status==="off"?"OFF":"Aktif"}</span></td><td>legacy-{String(c.legacyNo).padStart(3,"0")}</td></tr>)}</tbody></table></div></div></div>}
      {activeMenu === "billing" && <div className="admin-content"><div className="admin-page-heading"><div><span className="section-kicker">BILLING MANAGEMENT</span><h1>Tagihan Bulanan</h1><p>Catatan tagihan per pelanggan dan per bulan.</p></div><button className="excel-button" onClick={() => downloadJson({sourceSheet:"Tagihan Bulanan",year:2026,customers,bills})}><ArrowDownToLine size={15}/> Backup</button></div><div className="excel-panel"><div className="excel-table-wrap"><table className="excel-table"><thead><tr><th>No Pelanggan</th><th>Nama Pelanggan</th><th>Bulan</th><th>Nominal</th><th>Keterangan</th><th>Status</th></tr></thead><tbody>{bills.slice(0,100).map((b,i)=><tr key={i}><td>{b.legacyNo}</td><td>{b.customerName}</td><td>{b.billingMonth}</td><td>{b.amount == null ? (b.amountRaw || "—") : new Intl.NumberFormat("id-ID",{style:"currency",currency:"IDR",maximumFractionDigits:0}).format(b.amount)}</td><td>{b.paymentNote || "—"}</td><td><span className={b.status==="paid"?"sheet-status":"sheet-status off"}>{b.status || "Belum ada status"}</span></td></tr>)}</tbody></table>{bills.length===0 && <div className="empty-sheet">Belum ada data tagihan di sesi ini. Muat backup JSON untuk melihat catatan.</div>}</div></div></div>}
      {activeMenu === "backup" && <div className="admin-content"><div className="admin-page-heading"><div><span className="section-kicker">DATA SAFETY</span><h1>Backup & Restore</h1><p>Simpan salinan data atau muat kembali file backup Excel.</p></div></div><div className="backup-grid"><section className="backup-card"><div className="backup-icon"><ArrowDownToLine size={22}/></div><h3>Download Backup</h3><p>Unduh salinan data pelanggan dan tagihan yang saat ini tersedia di panel.</p><div className="backup-counts"><span>{customers.length} pelanggan</span><span>{bills.length} tagihan</span></div><button className="button-primary" onClick={()=>downloadJson({sourceFile:"IndimediaWifi",sourceSheet:"Backup Admin",year:2026,customers,bills})}><ArrowDownToLine size={16}/> Download file backup</button></section><section className="backup-card"><div className="backup-icon upload"><HardDriveUpload size={22}/></div><h3>Upload / Restore Backup</h3><p>Pilih file backup JSON IndimediaWifi. File akan diperiksa sebelum dimuat ke tampilan.</p><div className="backup-file-note">Format: <b>.xlsx / .xls</b> · sheet Data Pelanggan dan Tagihan Bulanan</div><input ref={fileRef} className="hidden-file-input" type="file" accept=".xlsx,.xls" onChange={handleBackupUpload}/><button className="button-primary" onClick={()=>fileRef.current?.click()}><HardDriveUpload size={16}/> Pilih file backup</button></section></div>{backupNotice && <div className="backup-result" role="status">{backupNotice}</div>}<div className="backup-warning"><LockKeyhole size={18}/><div><b>Catatan penting</b><p>Saat ini download dan upload bekerja pada data sesi pratinjau di browser. Restore belum menulis ke Firestore. Setelah konfigurasi Firebase dan aturan akses aman aktif, fungsi ini perlu dihubungkan ke database utama.</p></div></div></div>}
      {activeMenu === "settings" && <div className="admin-content"><div className="admin-page-heading"><div><span className="section-kicker">SYSTEM CONFIGURATION</span><h1>Pengaturan</h1><p>Informasi sistem dan status koneksi database.</p></div></div><div className="excel-panel settings-panel"><div><span>Nama usaha</span><b>Indimedia Wifi</b></div><div><span>Slogan</span><b>Internet murah untuk semua</b></div><div><span>Sumber data</span><b>Mode pratinjau lokal</b></div><div><span>Firestore</span><b className="pending-text">Belum dikonfigurasi / belum diuji</b></div><div><span>Login admin</span><b>Demo lokal — bukan autentikasi produksi</b></div></div></div>}
    </section>
  </main>;

  return <main className="site-shell" id="home">
    <header className="topbar"><Brand compact/><nav className="topnav" aria-label="Navigasi utama"><a href="#status">Status jaringan</a><a href="#bantuan">Pusat bantuan</a></nav><button className="nav-login" onClick={()=>openLogin("customer")}>Login <ArrowRight size={16}/></button></header>
    <section className="hero"><div className="hero-copy"><div className="eyebrow"><i/> INTERNET UNTUK SEMUA</div><h1>Koneksi yang<br/>menghubungkan <span>semua.</span></h1><p className="hero-description">Informasi jaringan, layanan pelanggan, dan bantuan koneksi IndimediaWiFi dalam satu portal yang mudah diakses.</p><div className="hero-actions"><button className="button-primary" onClick={()=>openLogin("customer")}>Login Pelanggan <ArrowRight size={18}/></button><button className="button-secondary" onClick={()=>openLogin("admin")}><ShieldCheck size={17}/> Login Admin</button></div><div className="trust-row"><span><CheckCircle2 size={15}/> Portal layanan terpadu</span><span><LockKeyhole size={15}/> Akses terkontrol</span></div></div><div className="hero-visual" aria-label="Ilustrasi jaringan WiFi"><div className="visual-grid"/><div className="orbit orbit-a"/><div className="orbit orbit-b"/><div className="network-core"><div className="signal-rings"><Wifi size={100} strokeWidth={2.4}/></div><div className="router-device"><span/><span/><span/><span/></div><div className="router-base"/></div><div className="floating-card floating-top"><RadioTower size={19}/><span><b>Network center</b><small>Monitoring jaringan</small></span></div><div className="floating-card floating-bottom"><i/><span><b>Terhubung lebih dekat</b><small>IndimediaWiFi service portal</small></span></div></div></section>
    <section className="status-section" id="status"><div className="section-heading"><div><span className="section-kicker">NETWORK OVERVIEW</span><h2>Status layanan pusat</h2><p>Informasi status akan mengikuti hasil monitoring server yang terhubung.</p></div><span className="demo-label"><Activity size={14}/> Monitoring belum aktif</span></div><div className="status-card"><div className="status-symbol"><Signal size={25}/></div><div className="status-main"><b>Server pusat IndimediaWiFi</b><span>Belum terhubung ke sistem monitoring</span></div><div className="status-state"><i/> Belum diketahui</div><div className="status-divider"/><div className="status-meta"><span>STATUS</span><b>Menunggu konfigurasi</b></div></div><p className="status-note"><CircleHelp size={15}/> Status belum real-time sampai monitoring server dikonfigurasi. Halaman ini tidak menganggap server aktif tanpa data nyata.</p></section>
    <section className="help-section" id="bantuan"><div className="help-icon"><Wifi size={24}/></div><div className="help-copy"><h3>Ada kendala koneksi?</h3><p>Periksa status server pusat terlebih dahulu. Jika status pusat normal tetapi koneksi rumah terganggu, tim teknis dapat melakukan pengecekan lanjutan di area pelanggan.</p></div><button className="help-link" onClick={()=>openLogin("customer")}>Buka portal pelanggan <ArrowRight size={16}/></button></section>
    <footer className="footer"><Brand compact/><span>© {new Date().getFullYear()} IndimediaWiFi · Internet murah untuk semua</span><span><i/> Monitoring belum aktif</span></footer>
    {loginType && <div className="modal-backdrop" onMouseDown={event=>{if(event.target===event.currentTarget)openLogin(null);}}><section className="login-modal" role="dialog" aria-modal="true" aria-labelledby="login-title"><button className="modal-close" onClick={()=>openLogin(null)} aria-label="Tutup dialog"><X size={19}/></button><div className="modal-brand"><Brand/></div><div className="modal-kicker">{loginType==="customer"?"CUSTOMER PORTAL":"ADMIN CONTROL CENTER"}</div><h2 id="login-title">{loginType==="customer"?"Login Pelanggan":"Login Admin"}</h2><p className="modal-description">{loginType==="customer"?"Masukkan ID pelanggan untuk melanjutkan ke portal layanan Anda.":"Masuk ke pusat kontrol pengelolaan IndimediaWiFi."}</p><form onSubmit={submitLogin}>{loginType==="customer"?<><label htmlFor="customer-id">ID Pelanggan</label><input id="customer-id" autoFocus value={customerId} onChange={event=>setCustomerId(event.target.value)} placeholder="Contoh: WIFI-00125" autoComplete="username" required/></>:<><label htmlFor="admin-username">Username admin</label><input id="admin-username" autoFocus value={username} onChange={event=>setUsername(event.target.value)} placeholder="Username" autoComplete="username" required/><label htmlFor="admin-password">Password</label><input id="admin-password" type="password" value={password} onChange={event=>setPassword(event.target.value)} placeholder="Password" autoComplete="current-password" required/></>}<button className="button-primary form-submit" type="submit">Login <ArrowRight size={17}/></button></form>{notice && <div className="notice-box" role="status">{notice}</div>}{loginType==="admin" && <div className="demo-credential">Akses demo: <b>iamardi</b> / <b>iamardi123</b></div>}<div className="modal-security"><LockKeyhole size={14}/> Login admin ini hanya untuk pratinjau. Jangan gunakan kredensial demo untuk sistem produksi.</div></section></div>}
  </main>;
}
