import { useEffect, useRef, useState } from "react";
import type { FormEvent, ChangeEvent } from "react";
import * as XLSX from "xlsx";
import { onAuthStateChanged, signInWithEmailAndPassword, signOut } from "firebase/auth";
import { collection, onSnapshot, doc, setDoc, deleteDoc, writeBatch, serverTimestamp, getDocs, query, where } from "firebase/firestore";
import { auth, db, hasFirebaseConfig } from "./firebase";
import { Activity, ArrowDownToLine, ArrowRight, CheckCircle2, CircleHelp, FileSpreadsheet, HardDriveUpload, LockKeyhole, LogOut, RadioTower, Search, ShieldCheck, Signal, Wifi, X, Users, ReceiptText, LayoutDashboard, Settings2 } from "lucide-react";

type LoginType = "customer" | "admin" | null;
type Customer = { legacyNo: number; name: string; customerId?: string; status?: string; phone?: string; address?: string; packageName?: string; basePrice?: number; dueDay?: number };
type PortalData = { customer: Customer; unpaidBills: Array<{id:string;billingMonth:string;year:number;amount:number;discount:number;total:number;status:string;paymentNote?:string}>; totalOutstanding:number };
type ServerStatus = {status:string;label:string;target?:string;checkedAt?:string;latencyMs?:number};
type Bill = { id?: string; legacyNo: number; customerName: string; billingMonth: string; year?: number; amount?: number | null; amountRaw?: string; discount?: number; paidAt?: string; paymentNote?: string; status?: string };
type BackupData = { sourceFile?: string; sourceSheet?: string; year?: number; customers: Customer[]; bills: Bill[]; exportedAt?: string };

const demoCustomers: Customer[] = [
  { legacyNo: 1, name: "Riski", status: "active", phone: "081234560001", address: "Contoh alamat 1", packageName: "Reguler", basePrice: 100000, dueDay: 10 },
  { legacyNo: 2, name: "Bik Saroh", status: "off", phone: "081234560002", address: "Contoh alamat 2", packageName: "Hemat", basePrice: 75000, dueDay: 5 },
  { legacyNo: 3, name: "Mas Dwi", status: "active", phone: "081234560003", address: "Contoh alamat 3", packageName: "Premium", basePrice: 125000, dueDay: 15 },
  { legacyNo: 4, name: "Danu", status: "active", phone: "081234560004", address: "Contoh alamat 4", packageName: "Reguler", basePrice: 100000, dueDay: 10 },
  { legacyNo: 5, name: "Dek Fila", status: "active", phone: "081234560005", address: "Contoh alamat 5", packageName: "Hemat", basePrice: 85000, dueDay: 20 },
];

function Brand({ compact = false }: { compact?: boolean }) {
  return <a className={`brand ${compact ? "brand-compact" : ""}`} href="#home" aria-label="IndimediaWiFi beranda">
    <img src="/indimediawifi-logo.svg" alt="" className="brand-logo" />
    <span className="brand-words"><strong>Indimedia<span>Wifi</span></strong>{!compact && <small>Internet murah untuk semua</small>}</span>
  </a>;
}

function downloadExcel(data: BackupData) {
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(data.customers.map(c => ({ No: c.legacyNo, "Nama Pelanggan": c.name, "ID Login": c.customerId || "", Status: c.status || "active", "Nomor HP": c.phone || "", Alamat: c.address || "", Paket: c.packageName || "", "Harga Dasar": c.basePrice ?? 0, "Jatuh Tempo": c.dueDay ?? 10 }))), "Data Pelanggan");
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(data.bills.map(b => ({ "No Pelanggan": b.legacyNo, "Nama Pelanggan": b.customerName, Bulan: b.billingMonth, Tahun: b.year || data.year || 2026, Nominal: b.amount ?? b.amountRaw ?? "", Keterangan: b.paymentNote || "", Status: b.status || "unpaid", Diskon: b.discount ?? 0, "Tanggal Bayar": b.paidAt || "" }))), "Tagihan Bulanan");
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet([{ File: data.sourceFile || "IndimediaWifi", Sheet: data.sourceSheet || "Backup Admin", Tahun: data.year || 2026, "Waktu Backup": new Date().toLocaleString("id-ID") }]), "Info Backup");
  XLSX.writeFile(wb, "indimediawifi-backup-" + new Date().toISOString().slice(0,10) + ".xlsx");
}

export default function App() {
  const [loginType, setLoginType] = useState<LoginType>(null);
  const [customerId, setCustomerId] = useState("");
  const [portalData, setPortalData] = useState<PortalData | null>(null);
  const [serverStatus, setServerStatus] = useState<ServerStatus>({status:"unknown",label:"Belum diketahui"});
  const [serverTarget, setServerTarget] = useState("");
  const [serverSaving, setServerSaving] = useState(false);
  const [username, setUsername] = useState("");
  const [password, setPassword] = useState("");
  const [notice, setNotice] = useState("");
  const [isAdmin, setIsAdmin] = useState(false);
  const [dbStatus, setDbStatus] = useState("Belum login");
  const [customersLoaded, setCustomersLoaded] = useState(false);
  const [billsLoaded, setBillsLoaded] = useState(false);
  const [dbError, setDbError] = useState("");
  const [saving, setSaving] = useState(false);
  const [activeMenu, setActiveMenu] = useState("dashboard");
  const [search, setSearch] = useState("");
  const [customers, setCustomers] = useState<Customer[]>([]);
  const [bills, setBills] = useState<Bill[]>([]);
  const [backupNotice, setBackupNotice] = useState("");
  const [billMonth, setBillMonth] = useState("Semua");
  const [billYear, setBillYear] = useState("2026");
  const [editor, setEditor] = useState<{kind:"customer"|"bill"; index?:number; data:Record<string,string>} | null>(null);
  const fileRef = useRef<HTMLInputElement>(null);

  function openLogin(type: LoginType) {
    setLoginType(type); setNotice(""); setUsername(""); setPassword("");
  }

  async function submitLogin(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();
    if (loginType === "admin") {
      if (!auth || !hasFirebaseConfig) { setNotice("Firebase belum dikonfigurasi. Periksa src/firebase.ts."); return; }
      try {
        await signInWithEmailAndPassword(auth, username.trim(), password);
        setIsAdmin(true); setLoginType(null); setActiveMenu("dashboard"); setNotice(""); setDbError(""); setDbStatus("Terhubung");
      } catch {
        setNotice("Login Firebase gagal. Buat akun admin email/password di Firebase Authentication dan pastikan Email/Password aktif.");
      }
      return;
    }
    try {
      const response = await fetch(`/.netlify/functions/customer-portal?id=${encodeURIComponent(customerId.trim())}`);
      const result = await response.json();
      if (!response.ok) throw new Error(result.error || "Login pelanggan gagal.");
      setPortalData(result as PortalData);
      setLoginType(null); setNotice("");
    } catch (error) {
      setNotice(error instanceof Error ? error.message : "Tidak dapat menghubungi portal pelanggan.");
    }
  }

  async function logout() {
    try { if (auth) await signOut(auth); } finally { setIsAdmin(false); setPortalData(null); setLoginType(null); setActiveMenu("dashboard"); setCustomers([]); setBills([]); setDbStatus("Belum login"); }
  }

  useEffect(() => {
    if (!auth) return;
    const unsubscribe = onAuthStateChanged(auth, user => {
      setIsAdmin(Boolean(user));
      setDbStatus(user ? "Terhubung" : "Belum login");
    });
    return unsubscribe;
  }, []);

  useEffect(() => {
    if (!isAdmin || !db || !auth?.currentUser) return;
    setDbStatus("Terhubung"); setDbError("");
    const stopCustomers = onSnapshot(collection(db, "customers"), snapshot => {
      setCustomers(snapshot.docs.map(item => ({ ...item.data(), legacyNo: Number(item.data().legacyNo ?? item.id) } as Customer)).sort((a,b)=>a.legacyNo-b.legacyNo)); setCustomersLoaded(true);
    }, error => { setDbStatus("Gagal"); setDbError(error.message); });
    const stopBills = onSnapshot(collection(db, "bills"), snapshot => {
      setBills(snapshot.docs.map(item => ({ ...item.data(), id: item.id } as Bill)).sort((a,b)=>(b.year||0)-(a.year||0) || a.legacyNo-b.legacyNo)); setBillsLoaded(true);
    }, error => { setDbStatus("Gagal"); setDbError(error.message); });
    return () => { stopCustomers(); stopBills(); };
  }, [isAdmin]);

  useEffect(() => {
    let active = true;
    const checkServer = async () => {
      try {
        const response = await fetch("/.netlify/functions/customer-portal?action=status");
        const result = await response.json();
        if (active && response.ok) {
          setServerStatus(result as ServerStatus);
          setServerTarget(String(result.target || ""));
        }
      } catch {
        if (active) setServerStatus({status:"unknown",label:"Monitoring tidak tersedia"});
      }
    };
    void checkServer();
    const timer = window.setInterval(checkServer, 60000);
    return () => { active = false; window.clearInterval(timer); };
  }, []);

  async function saveServerTarget() {
    if (!auth?.currentUser) { setDbError("Sesi admin Firebase tidak ditemukan."); return; }
    setServerSaving(true);
    try {
      const token = await auth.currentUser.getIdToken();
      const response = await fetch("/.netlify/functions/customer-portal", {
        method:"POST",
        headers:{"Content-Type":"application/json", Authorization:`Bearer ${token}`},
        body:JSON.stringify({target:serverTarget})
      });
      const result = await response.json();
      if (!response.ok) throw new Error(result.error || "Gagal menyimpan target monitoring.");
      setDbError("Target monitoring berhasil disimpan.");
      const statusResponse = await fetch("/.netlify/functions/customer-portal?action=status");
      const statusResult = await statusResponse.json();
      if (statusResponse.ok) setServerStatus(statusResult as ServerStatus);
    } catch (error) {
      setDbError(error instanceof Error ? error.message : "Gagal menyimpan target monitoring.");
    } finally {
      setServerSaving(false);
    }
  }

  async function seedDemoData() {
    if (!db) { setDbError("Firestore belum dikonfigurasi."); return; }
    if (!customersLoaded || !billsLoaded) { setDbError("Tunggu sampai Firestore selesai memuat data."); return; }
    if (customers.length || bills.length) { setDbError("Data demo hanya bisa dimuat saat database masih kosong."); return; }
    const demoBills: Bill[] = [
      {id:"demo-1",legacyNo:1,customerName:"Riski",billingMonth:"Januari",year:2026,amount:100000,status:"paid",paidAt:"2026-01-10",paymentNote:"Data demo"},
      {id:"demo-2",legacyNo:1,customerName:"Riski",billingMonth:"Februari",year:2026,amount:100000,status:"paid",paidAt:"2026-02-10",paymentNote:"Data demo"},
      {id:"demo-3",legacyNo:2,customerName:"Bik Saroh",billingMonth:"Januari",year:2026,amount:75000,status:"unpaid",paymentNote:"Data demo"},
      {id:"demo-4",legacyNo:2,customerName:"Bik Saroh",billingMonth:"Februari",year:2026,amount:75000,status:"unpaid",paymentNote:"Data demo"},
      {id:"demo-5",legacyNo:3,customerName:"Mas Dwi",billingMonth:"Januari",year:2026,amount:125000,status:"paid",paidAt:"2026-01-15",paymentNote:"Data demo"},
      {id:"demo-6",legacyNo:3,customerName:"Mas Dwi",billingMonth:"Februari",year:2026,amount:125000,status:"unpaid",paymentNote:"Data demo"},
      {id:"demo-7",legacyNo:4,customerName:"Danu",billingMonth:"Januari",year:2026,amount:100000,status:"paid",paidAt:"2026-01-10",paymentNote:"Data demo"},
      {id:"demo-8",legacyNo:5,customerName:"Dek Fila",billingMonth:"Januari",year:2026,amount:85000,status:"unpaid",paymentNote:"Data demo"}
    ];
    try {
      const batch=writeBatch(db);
      demoCustomers.forEach(customer=>batch.set(doc(db!,"customers",String(customer.legacyNo)),{...customer,isDemo:true,updatedAt:serverTimestamp()}));
      demoBills.forEach(bill=>batch.set(doc(db!,"bills",bill.id!),{...bill,isDemo:true,updatedAt:serverTimestamp()}));
      await batch.commit(); setDbError("Data demo tersimpan di Firestore. Silakan sesuaikan dengan data asli.");
    } catch(error) { setDbError(error instanceof Error?error.message:"Gagal menyimpan data demo."); }
  }

  function handleBackupUpload(event: ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = async () => {
      try {
        const wb = XLSX.read(reader.result, { type: "array", cellDates: true });
        let nc: Customer[] = [];
        let nb: Bill[] = [];
        let sourceSheet = "";
        let sourceYear = 2026;
        const customerSheet = wb.Sheets["Data Pelanggan"];
        const billSheet = wb.Sheets["Tagihan Bulanan"];

        if (customerSheet && billSheet) {
          sourceSheet = "Data Pelanggan + Tagihan Bulanan";
          const cr = XLSX.utils.sheet_to_json<Record<string, unknown>>(customerSheet, { defval: "" });
          const br = XLSX.utils.sheet_to_json<Record<string, unknown>>(billSheet, { defval: "" });
          nc = cr.map(r => ({
            legacyNo: Number(r["No"]), name: String(r["Nama Pelanggan"] || "").trim(),
            status: ["off","OFF","nonaktif","Nonaktif"].includes(String(r["Status"] || "")) ? "off" : "active",
            phone: String(r["Nomor HP"] || ""), address: String(r["Alamat"] || ""),
            packageName: String(r["Paket"] || ""), basePrice: Number(r["Harga Dasar"] || 0),
            dueDay: Number(r["Jatuh Tempo"] || 10)
          })).filter(x => x.name && Number.isFinite(x.legacyNo));
          nb = br.map((r, i) => {
            const raw = r["Nominal"];
            const rawText = String(raw ?? "").trim();
            const amount = typeof raw === "number" ? raw : (/^\d+(?:\.\d+)?$/.test(rawText) ? Number(rawText) : null);
            const month = String(r["Bulan"] || "").trim();
            const year = Number(r["Tahun"] || 2026);
            return {
              legacyNo: Number(r["No Pelanggan"]), customerName: String(r["Nama Pelanggan"] || "").trim(),
              billingMonth: month, year, amount, amountRaw: amount == null ? rawText : "",
              discount: Number(r["Diskon"] || 0), paidAt: String(r["Tanggal Bayar"] || ""),
              paymentNote: String(r["Keterangan"] || ""),
              status: ["paid","lunas","Lunas","PAID"].includes(String(r["Status"] || "")) ? "paid" : "unpaid",
              sourceSheet, sourceRow: i + 2
            } as Bill;
          }).filter(x => Number.isFinite(x.legacyNo) && x.billingMonth);
          sourceYear = Number(br[0]?.["Tahun"] || 2026);
        } else {
          const sourceName = wb.SheetNames.find(name => /^CAT WIFI(?:\s+\d{4})?$/i.test(name.trim()));
          if (!sourceName) throw new Error("Format file tidak dikenal. Gunakan backup IndimediaWifi atau sheet CAT WIFI 2026.");
          sourceSheet = sourceName;
          const rows = XLSX.utils.sheet_to_json<unknown[]>(wb.Sheets[sourceName], { header: 1, raw: true, defval: null }) as unknown[][];
          const header = rows[0] || [];
          const months = ["Januari","Februari","Maret","April","Mei","Juni","Juli","Agustus","September","Oktober","November","Desember"];
          const monthColumns: Array<{ col: number; month: string; year: number }> = [];
          header.forEach((value, col) => {
            if (col < 2 || col > 25 || value == null) return;
            let date: Date | null = value instanceof Date ? value : null;
            if (!date && typeof value === "number") {
              const d = XLSX.SSF.parse_date_code(value);
              if (d) date = new Date(d.y, d.m - 1, d.d || 1);
            }
            if (!date && typeof value === "string") {
              const d = new Date(value);
              if (!Number.isNaN(d.getTime())) date = d;
            }
            if (date && !Number.isNaN(date.getTime())) {
              monthColumns.push({ col, month: months[date.getMonth()], year: date.getFullYear() });
              sourceYear = date.getFullYear();
            }
          });
          rows.slice(1).forEach((row, index) => {
            const legacyNo = Number(row[0]);
            const name = String(row[1] ?? "").trim();
            if (!name || !Number.isFinite(legacyNo) || /^total$/i.test(name)) return;
            let status = "active";
            let basePrice = 0;
            monthColumns.forEach(({ col, month, year }) => {
              const raw = row[col];
              const noteRaw = row[col + 1];
              const note = noteRaw instanceof Date
                ? `${String(noteRaw.getDate()).padStart(2,"0")}/${String(noteRaw.getMonth()+1).padStart(2,"0")}/${String(noteRaw.getFullYear()).slice(-2)}`
                : String(noteRaw ?? "").trim();
              if (typeof raw === "string" && /^OFF$/i.test(raw.trim())) { status = "off"; return; }
              if (raw == null || raw === "") return;
              let amount: number | null = null;
              let amountRaw = "";
              if (typeof raw === "number" && Number.isFinite(raw)) amount = raw;
              else {
                const txt = String(raw).trim().replace(/\s/g, "");
                if (/^\d+(?:\.\d+)?(?:\+\d+(?:\.\d+)?)+$/.test(txt)) {
                  amount = txt.split("+").reduce((sum, part) => sum + Number(part), 0);
                  amountRaw = String(raw).trim();
                } else amountRaw = String(raw).trim();
              }
              if (amount != null && amount > 0) basePrice = amount;
              if ((amount == null || amount <= 0) && !amountRaw) return;
              nb.push({
                legacyNo, customerName: name, billingMonth: month, year, amount, amountRaw,
                paymentNote: note, status: note ? "paid" : "unpaid", sourceSheet, sourceRow: index + 2
              } as Bill);
            });
            nc.push({ legacyNo, name, status, basePrice });
          });
        }

        if (!nc.length) throw new Error("Tidak ditemukan data pelanggan valid pada file ini.");
        if (!db || !auth?.currentUser) throw new Error("Login admin Firebase diperlukan sebelum restore.");
        const approved = window.confirm(
          `Restore file "${file.name}" ke Firestore?\n\nSumber: ${sourceSheet}\nPelanggan: ${nc.length}\nCatatan tagihan: ${nb.length}\n\nData dengan nomor pelanggan dan periode tagihan yang sama akan diperbarui. Data lain tidak dihapus.`
        );
        if (!approved) { setBackupNotice("Restore dibatalkan. Tidak ada data yang diubah."); return; }
        const batch = writeBatch(db);
        const demoCustomers = await getDocs(query(collection(db, "customers"), where("isDemo", "==", true)));
        const demoBills = await getDocs(query(collection(db, "bills"), where("isDemo", "==", true)));
        demoCustomers.docs.forEach(item => batch.delete(item.ref));
        demoBills.docs.forEach(item => batch.delete(item.ref));
        nc.forEach(customer => batch.set(doc(db!, "customers", String(customer.legacyNo)), {
          ...customer, sourceFile: file.name, sourceSheet, isDemo: false, updatedAt: serverTimestamp()
        }, { merge: true }));
        nb.forEach(bill => {
          const months = ["Januari","Februari","Maret","April","Mei","Juni","Juli","Agustus","September","Oktober","November","Desember"];
          const monthNumber = /^\d{4}-\d{2}$/.test(bill.billingMonth) ? bill.billingMonth.slice(5,7) :
            String(months.indexOf(bill.billingMonth) + 1).padStart(2,"0");
          const id = `import-${bill.legacyNo}-${bill.year || sourceYear}-${monthNumber}`;
          batch.set(doc(db!, "bills", id), {
            ...bill, id, year: bill.year || sourceYear, sourceFile: file.name, sourceSheet, isDemo: false, updatedAt: serverTimestamp()
          }, { merge: true });
        });
        if (nc.length + nb.length + demoCustomers.size + demoBills.size > 500) throw new Error("File terlalu besar untuk satu batch Firestore.");
        await batch.commit();
        setBackupNotice(`Restore berhasil: ${nc.length} pelanggan dan ${nb.length} tagihan dari "${sourceSheet}" tersimpan ke Firestore. Data demo lama dibersihkan jika ada.`);
      } catch (error) {
        setBackupNotice(error instanceof Error ? error.message : "File Excel gagal dibaca.");
      } finally {
        event.target.value = "";
      }
    };
    reader.readAsArrayBuffer(file);
  }

  const filteredCustomers = customers.filter(c => c.name.toLowerCase().includes(search.toLowerCase()) || String(c.legacyNo).includes(search) || String(c.customerId || "").toLowerCase().includes(search.toLowerCase()));
  const paidCount = bills.filter(b => b.status === "paid").length;

  if (portalData && !isAdmin) return <main className="site-shell customer-portal-shell">
    <header className="topbar"><Brand compact/><span className="portal-top-label">PORTAL PELANGGAN</span><button className="nav-login" onClick={logout}>Keluar <LogOut size={16}/></button></header>
    <section className="portal-content">
      <div className="admin-page-heading"><div><span className="section-kicker">CUSTOMER DASHBOARD</span><h1>Halo, {portalData.customer.name}</h1><p>Informasi layanan dan ringkasan tagihan Anda.</p></div></div>
      <div className="admin-stat-grid portal-stat-grid">
        <div className="admin-stat"><span>Nomor Pelanggan</span><strong>{portalData.customer.legacyNo}</strong><small>ID login: {portalData.customer.customerId || customerId}</small></div>
        <div className="admin-stat"><span>Status Layanan</span><strong>{portalData.customer.status === "off" ? "OFF" : "Aktif"}</strong><small>{portalData.customer.packageName || "Paket belum diatur"}</small></div>
        <div className="admin-stat"><span>Total Tagihan Belum Lunas</span><strong>{new Intl.NumberFormat("id-ID",{style:"currency",currency:"IDR",maximumFractionDigits:0}).format(portalData.totalOutstanding)}</strong><small>{portalData.unpaidBills.length} tagihan</small></div>
      </div>
      <section className="excel-panel server-check-panel"><div className="excel-panel-heading"><div><RadioTower size={19}/><span><b>Status server pusat</b><small>{serverStatus.target || "Target belum dikonfigurasi"}</small></span></div><span className={serverStatus.status==="online"?"sheet-status":"sheet-status off"}>{serverStatus.label}</span></div><div className="server-check-grid"><div><span>Latency</span><b>{serverStatus.latencyMs != null ? serverStatus.latencyMs+" ms" : "—"}</b></div><div><span>Pemeriksaan terakhir</span><b>{serverStatus.checkedAt ? new Date(serverStatus.checkedAt).toLocaleString("id-ID") : "—"}</b></div></div></section>
      <section className="excel-panel"><div className="excel-panel-heading"><div><ReceiptText size={19}/><span><b>Tagihan belum lunas</b><small>Nominal dikurangi diskon</small></span></div></div><div className="excel-table-wrap"><table className="excel-table"><thead><tr><th>Periode</th><th>Nominal</th><th>Diskon</th><th>Total</th><th>Status</th></tr></thead><tbody>{portalData.unpaidBills.map(b=><tr key={b.id}><td>{b.billingMonth} {b.year}</td><td>{new Intl.NumberFormat("id-ID",{style:"currency",currency:"IDR",maximumFractionDigits:0}).format(b.amount)}</td><td>{new Intl.NumberFormat("id-ID",{style:"currency",currency:"IDR",maximumFractionDigits:0}).format(b.discount)}</td><td>{new Intl.NumberFormat("id-ID",{style:"currency",currency:"IDR",maximumFractionDigits:0}).format(b.total)}</td><td><span className="sheet-status off">Belum lunas</span></td></tr>)}</tbody></table>{portalData.unpaidBills.length===0&&<div className="empty-sheet">Tidak ada tagihan yang belum lunas. Terima kasih!</div>}</div></section>
    </section>
    <footer className="footer"><Brand compact/><span>© {new Date().getFullYear()} IndimediaWiFi</span></footer>
  </main>;

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
        <div className="admin-page-heading"><div><span className="section-kicker">ADMIN CONTROL CENTER</span><h1>Dashboard</h1><p>Ringkasan administrasi pelanggan dan tagihan IndimediaWifi.</p></div>{customersLoaded&&billsLoaded&&customers.length===0&&bills.length===0&&<button className="excel-button" onClick={()=>void seedDemoData()}>Isi data demo ke Firestore</button>}<button className="excel-button" onClick={() => { setActiveMenu("backup"); setBackupNotice(""); }}><HardDriveUpload size={16}/> Kelola Backup</button></div>
        <div className="admin-stat-grid"><div className="admin-stat"><span><Users size={17}/> Total Pelanggan</span><strong>{customers.length}</strong><small>Data yang tampil di panel ini</small></div><div className="admin-stat"><span><ReceiptText size={17}/> Catatan Tagihan</span><strong>{bills.length}</strong><small>Data tagihan pada sesi ini</small></div><div className="admin-stat"><span><CheckCircle2 size={17}/> Status Lunas</span><strong>{paidCount}</strong><small>Dari data backup yang dimuat</small></div><div className="admin-stat"><span><Activity size={17}/> Database</span><strong className={dbStatus==="Terhubung"?"":"pending-text"}>{dbStatus}</strong><small>{dbError || "Sinkronisasi Firestore live"}</small></div></div>
        <section className="excel-panel server-check-panel"><div className="excel-panel-heading"><div><RadioTower size={19}/><span><b>Pemeriksaan Server / MikroTik</b><small>Status monitoring jaringan</small></span></div><span className={serverStatus.status==="online"?"sheet-status":"sheet-status off"}>{serverStatus.label}</span></div><p className="admin-hint">Masukkan domain atau IP yang bisa dijangkau dari server Netlify. Ini menguji keterjangkauan target, bukan kualitas Wi-Fi setiap pelanggan.</p><label className="server-target-label">Target monitoring (domain atau IP)<input value={serverTarget} onChange={e=>setServerTarget(e.target.value)} placeholder="contoh: status.indimediawifi.id atau 203.0.113.10"/></label><div className="admin-edit-actions"><button className="excel-button" onClick={()=>void saveServerTarget()} disabled={serverSaving}>{serverSaving?"Menyimpan…":"Simpan target & cek"}</button><span>{serverStatus.target || "Belum dikonfigurasi"} {serverStatus.checkedAt ? "· "+new Date(serverStatus.checkedAt).toLocaleString("id-ID") : ""}</span></div></section>
        <section className="excel-panel"><div className="excel-panel-heading"><div><FileSpreadsheet size={19}/><span><b>Ringkasan Data Pelanggan</b><small>Pratinjau gaya lembar kerja Excel</small></span></div><button className="excel-button" onClick={() => downloadExcel({ sourceFile: "IndimediaWifi", sourceSheet: "Data Pelanggan", year: 2026, customers, bills })}><ArrowDownToLine size={15}/> Download Backup</button></div>
          <div className="excel-table-wrap"><table className="excel-table"><thead><tr><th>No</th><th>Nama Pelanggan</th><th>Status Layanan</th><th>Tagihan Tersimpan</th></tr></thead><tbody>{filteredCustomers.slice(0,8).map(c => <tr key={c.legacyNo}><td>{c.legacyNo}</td><td>{c.name}</td><td><span className={c.status === "off" ? "sheet-status off" : "sheet-status"}>{c.status === "off" ? "OFF" : "Aktif"}</span></td><td>{bills.filter(b=>b.legacyNo===c.legacyNo).length ? bills.filter(b=>b.legacyNo===c.legacyNo).length + " bulan" : "—"}</td></tr>)}</tbody></table></div>
          <p className="admin-hint">Data pelanggan dan tagihan dimuat langsung dari Firestore; perubahan tersinkron otomatis.</p>
        </section>
      </div>}
      {activeMenu === "customers" && <div className="admin-content"><div className="admin-page-heading"><div><span className="section-kicker">CUSTOMER MANAGEMENT</span><h1>Data Pelanggan</h1><p>Daftar pelanggan dengan format tabel menyerupai lembar Excel.</p></div><div className="action-group"><button className="excel-button" onClick={() => {const legacyNo=Math.max(0,...customers.map(c=>c.legacyNo))+1;setEditor({kind:"customer",data:{legacyNo:String(legacyNo),name:"",customerId:"",phone:"",address:"",packageName:"Reguler",basePrice:"100000",dueDay:"10",status:"active"}});}}><Users size={15}/> Tambah Pelanggan</button><button className="excel-button" onClick={() => downloadExcel({sourceSheet:"Data Pelanggan",year:2026,customers,bills})}><ArrowDownToLine size={15}/> Backup Excel</button></div></div><div className="excel-panel"><div className="table-toolbar"><div className="admin-search"><Search size={16}/><input value={search} onChange={e=>setSearch(e.target.value)} placeholder="Cari nomor atau nama pelanggan"/></div><span>{filteredCustomers.length} pelanggan</span></div><div className="excel-table-wrap"><table className="excel-table"><thead><tr><th>No</th><th>Nama Pelanggan</th><th>Status</th><th>ID Login</th><th>ID Legacy</th><th>Aksi</th></tr></thead><tbody>{filteredCustomers.map(c=><tr key={c.legacyNo}><td>{c.legacyNo}</td><td>{c.name}</td><td><span className={c.status==="off"?"sheet-status off":"sheet-status"}>{c.status==="off"?"OFF":"Aktif"}</span></td><td>{c.customerId || "—"}</td><td>legacy-{String(c.legacyNo).padStart(3,"0")}</td><td><button className="row-action" onClick={()=>setEditor({kind:"customer",index:c.legacyNo,data:{legacyNo:String(c.legacyNo),name:c.name,customerId:c.customerId||"",phone:c.phone||"",address:c.address||"",packageName:c.packageName||"Reguler",basePrice:String(c.basePrice??100000),dueDay:String(c.dueDay??10),status:c.status||"active"}})}>Edit</button><button className="row-action danger" onClick={()=>{if(window.confirm("Hapus pelanggan "+c.name+"?")){void (async()=>{if(!db)return;try{await deleteDoc(doc(db,"customers",String(c.legacyNo)));const batch=writeBatch(db);bills.filter(b=>b.legacyNo===c.legacyNo).forEach(b=>{if(b.id)batch.delete(doc(db!,"bills",b.id));});await batch.commit();}catch(error){setDbError(error instanceof Error?error.message:"Gagal menghapus pelanggan.");}})();}}}>Hapus</button></td></tr>)}</tbody></table></div></div></div>}
      {activeMenu === "billing" && <div className="admin-content"><div className="admin-page-heading"><div><span className="section-kicker">BILLING MANAGEMENT</span><h1>Tagihan Bulanan</h1><p>Catatan tagihan per pelanggan dan per bulan.</p></div><div className="action-group"><button className="excel-button" onClick={() => {const customer=customers[0];if(!customer){setNotice("Tambahkan pelanggan terlebih dahulu.");return;}setEditor({kind:"bill",data:{legacyNo:String(customer.legacyNo),customerName:customer.name,billingMonth:"Oktober",year:"2026",amount:String(customer.basePrice??100000),discount:"0",paidAt:"",paymentNote:"",status:"unpaid"}});}}><ReceiptText size={15}/> Tambah Tagihan</button><button className="excel-button" onClick={() => downloadExcel({sourceSheet:"Tagihan Bulanan",year:2026,customers,bills})}><ArrowDownToLine size={15}/> Backup Excel</button></div></div><div className="excel-panel"><div className="table-toolbar"><label>Bulan <select value={billMonth} onChange={e=>setBillMonth(e.target.value)}><option>Semua</option>{["Januari","Februari","Maret","April","Mei","Juni","Juli","Agustus","September","Oktober","November","Desember"].map(m=><option key={m}>{m}</option>)}</select></label><label>Tahun <input className="year-filter" type="number" value={billYear} onChange={e=>setBillYear(e.target.value)}/></label></div><div className="excel-table-wrap"><table className="excel-table"><thead><tr><th>No Pelanggan</th><th>Nama Pelanggan</th><th>Bulan</th><th>Tahun</th><th>Nominal</th><th>Keterangan</th><th>Status</th><th>Aksi</th></tr></thead><tbody>{bills.filter(b=>(billMonth==="Semua"||b.billingMonth===billMonth)&&String(b.year||2026)===billYear).slice(0,100).map((b,i)=><tr key={i}><td>{b.legacyNo}</td><td>{b.customerName}</td><td>{b.billingMonth}</td><td>{b.year||2026}</td><td>{b.amount == null ? (b.amountRaw || "—") : new Intl.NumberFormat("id-ID",{style:"currency",currency:"IDR",maximumFractionDigits:0}).format(b.amount)}</td><td>{b.paymentNote || "—"}</td><td><span className={b.status==="paid"?"sheet-status":"sheet-status off"}>{b.status || "Belum ada status"}</span></td><td><button className="row-action" onClick={()=>setEditor({kind:"bill",index:bills.indexOf(b),data:{legacyNo:String(b.legacyNo),customerName:b.customerName,billingMonth:b.billingMonth,year:String(b.year||2026),amount:String(b.amount??b.amountRaw??""),discount:String(b.discount??0),paidAt:b.paidAt||"",paymentNote:b.paymentNote||"",status:b.status||"unpaid"}})}>Edit</button><button className="row-action danger" onClick={()=>{if(window.confirm("Hapus catatan tagihan ini?"))void (async()=>{if(!db||!b.id)return;try{await deleteDoc(doc(db,"bills",b.id));}catch(error){setDbError(error instanceof Error?error.message:"Gagal menghapus tagihan.");}})();}}>Hapus</button></td></tr>)}</tbody></table>{bills.length===0 && <div className="empty-sheet">Belum ada data tagihan di sesi ini. Muat backup JSON untuk melihat catatan.</div>}</div></div></div>}
      {activeMenu === "backup" && <div className="admin-content"><div className="admin-page-heading"><div><span className="section-kicker">DATA SAFETY</span><h1>Backup & Restore</h1><p>Simpan salinan data atau muat kembali file backup Excel.</p></div></div><div className="backup-grid"><section className="backup-card"><div className="backup-icon"><ArrowDownToLine size={22}/></div><h3>Download Backup</h3><p>Unduh salinan data pelanggan dan tagihan yang saat ini tersedia di panel.</p><div className="backup-counts"><span>{customers.length} pelanggan</span><span>{bills.length} tagihan</span></div><button className="button-primary" onClick={()=>downloadExcel({sourceFile:"IndimediaWifi",sourceSheet:"Backup Admin",year:2026,customers,bills})}><ArrowDownToLine size={16}/> Download file backup</button></section><section className="backup-card"><div className="backup-icon upload"><HardDriveUpload size={22}/></div><h3>Upload / Restore Backup</h3><p>Pilih backup IndimediaWifi atau file sumber Catatan Indi Media WiFi.xlsx (sheet CAT WIFI 2026). Data akan divalidasi sebelum disimpan ke Firestore.</p><div className="backup-file-note">Format: <b>.xlsx / .xls</b> · CAT WIFI 2026 atau sheet Data Pelanggan + Tagihan Bulanan</div><input ref={fileRef} className="hidden-file-input" type="file" accept=".xlsx,.xls" onChange={handleBackupUpload}/><button className="button-primary" onClick={()=>fileRef.current?.click()}><HardDriveUpload size={16}/> Pilih file backup</button></section></div>{backupNotice && <div className="backup-result" role="status">{backupNotice}</div>}<div className="backup-warning"><LockKeyhole size={18}/><div><b>Catatan penting</b><p>Restore akan menulis data yang lolos validasi ke Firestore. File CAT WIFI 2026 dipetakan menjadi pelanggan dan tagihan per bulan. Unduh backup terlebih dahulu sebelum restore agar ada salinan data saat ini.</p></div></div></div>}
      {activeMenu === "settings" && <div className="admin-content"><div className="admin-page-heading"><div><span className="section-kicker">SYSTEM CONFIGURATION</span><h1>Pengaturan</h1><p>Informasi sistem dan status koneksi database.</p></div></div><div className="excel-panel settings-panel"><div><span>Nama usaha</span><b>Indimedia Wifi</b></div><div><span>Slogan</span><b>Internet murah untuk semua</b></div><div><span>Sumber data</span><b>Cloud Firestore</b></div><div><span>Firestore</span><b className={dbStatus==="Terhubung"?"":"pending-text"}>{dbStatus}</b></div><div><span>Login admin</span><b>Firebase Authentication</b></div></div></div>}
    </section>

      {editor && <div className="modal-backdrop" onMouseDown={e=>{if(e.target===e.currentTarget)setEditor(null)}}><section className="login-modal admin-edit-modal" role="dialog" aria-modal="true" aria-labelledby="edit-title"><button className="modal-close" onClick={()=>setEditor(null)} aria-label="Tutup"><X size={19}/></button><div className="modal-kicker">ADMIN DATA MANAGEMENT</div><h2 id="edit-title">{editor.kind==="customer"?(editor.index===undefined?"Tambah Pelanggan":"Edit Pelanggan"):(editor.index===undefined?"Tambah Tagihan":"Edit Tagihan Bulanan")}</h2><form onSubmit={async e=>{e.preventDefault();const d=editor.data;if (!db) { setDbError("Firestore belum tersedia."); return; }
      setSaving(true);
      try {
        if (editor.kind==="customer") {
          const row:Customer={legacyNo:Number(d.legacyNo),name:d.name.trim(),customerId:String(d.customerId||"").trim(),phone:d.phone,address:d.address,packageName:d.packageName,basePrice:Math.max(0,Number(d.basePrice)||0),dueDay:Math.min(28,Math.max(1,Number(d.dueDay)||10)),status:d.status};
          if(!row.name)return;
          await setDoc(doc(db,"customers",String(row.legacyNo)),{...row,updatedAt:serverTimestamp()},{merge:true});
        } else {
          const customer=customers.find(x=>x.legacyNo===Number(d.legacyNo));if(!customer)return;
          const id=editor.index===undefined?("bill-"+Date.now()):(bills[editor.index]?.id||("bill-"+Date.now()));
          const row:Bill={id,legacyNo:customer.legacyNo,customerName:customer.name,billingMonth:d.billingMonth,year:Number(d.year),amount:Math.max(0,Number(d.amount)||0),discount:Math.max(0,Number(d.discount)||0),paidAt:d.paidAt,paymentNote:d.paymentNote,status:d.status};
          await setDoc(doc(db,"bills",id),{...row,updatedAt:serverTimestamp()},{merge:true});
        }
        setEditor(null);setDbError("");
      } catch(error) { setDbError(error instanceof Error?error.message:"Gagal menyimpan ke Firestore."); }
      finally { setSaving(false); }}}><div className="admin-edit-grid">{(editor.kind==="customer"?[["name","Nama pelanggan","text"],["customerId","ID Login Pelanggan","text"],["phone","Nomor HP","tel"],["address","Alamat","text"],["packageName","Paket internet","text"],["basePrice","Harga dasar bulanan (Rp)","number"],["dueDay","Tanggal jatuh tempo (1–28)","number"]]:[["legacyNo","Nomor pelanggan","customer"],["billingMonth","Bulan tagihan","month"],["year","Tahun","number"],["amount","Nominal tagihan (Rp)","number"],["discount","Diskon (Rp)","number"],["paidAt","Tanggal pembayaran","date"],["paymentNote","Keterangan","text"]]).map(([key,label,type])=><label key={key}>{label}{type==="customer"?<select value={editor.data[key]||""} onChange={e=>setEditor({...editor,data:{...editor.data,[key]:e.target.value}})}>{customers.map(c=><option key={c.legacyNo} value={c.legacyNo}>{c.legacyNo} — {c.name}</option>)}</select>:type==="month"?<select value={editor.data[key]||"Januari"} onChange={e=>setEditor({...editor,data:{...editor.data,[key]:e.target.value}})}>{["Januari","Februari","Maret","April","Mei","Juni","Juli","Agustus","September","Oktober","November","Desember"].map(m=><option key={m}>{m}</option>)}</select>:<input type={type==="month"?"text":type} value={editor.data[key]||""} required={["name","billingMonth","year","amount"].includes(key)} onChange={e=>setEditor({...editor,data:{...editor.data,[key]:e.target.value}})}/>}</label>)}<label>Status {editor.kind==="customer"?"layanan":"pembayaran"}<select value={editor.data.status||"active"} onChange={e=>setEditor({...editor,data:{...editor.data,status:e.target.value}})}>{(editor.kind==="customer"?[["active","Aktif"],["off","OFF"]]:[["unpaid","Belum Bayar"],["paid","Lunas"]]).map(([v,l])=><option key={v} value={v}>{l}</option>)}</select></label></div><div className="admin-edit-actions"><button type="button" className="excel-button" onClick={()=>setEditor(null)}>Batal</button><button type="submit" className="button-primary" disabled={saving}>{saving?"Menyimpan…":"Simpan ke Firestore"}</button></div></form></section></div>}
  </main>;

  return <main className="site-shell" id="home">
    <header className="topbar"><Brand compact/><nav className="topnav" aria-label="Navigasi utama"><a href="#status">Status jaringan</a><a href="#bantuan">Pusat bantuan</a></nav><button className="nav-login" onClick={()=>openLogin("customer")}>Login <ArrowRight size={16}/></button></header>
    <section className="hero"><div className="hero-copy"><div className="eyebrow"><i/> INTERNET UNTUK SEMUA</div><h1>Koneksi yang<br/>menghubungkan <span>semua.</span></h1><p className="hero-description">Informasi jaringan, layanan pelanggan, dan bantuan koneksi IndimediaWiFi dalam satu portal yang mudah diakses.</p><div className="hero-actions"><button className="button-primary" onClick={()=>openLogin("customer")}>Login Pelanggan <ArrowRight size={18}/></button><button className="button-secondary" onClick={()=>openLogin("admin")}><ShieldCheck size={17}/> Login Admin</button></div><div className="trust-row"><span><CheckCircle2 size={15}/> Portal layanan terpadu</span><span><LockKeyhole size={15}/> Akses terkontrol</span></div></div><div className="hero-visual" aria-label="Ilustrasi jaringan WiFi"><div className="visual-grid"/><div className="orbit orbit-a"/><div className="orbit orbit-b"/><div className="network-core"><div className="signal-rings"><Wifi size={100} strokeWidth={2.4}/></div><div className="router-device"><span/><span/><span/><span/></div><div className="router-base"/></div><div className="floating-card floating-top"><RadioTower size={19}/><span><b>Network center</b><small>Monitoring jaringan</small></span></div><div className="floating-card floating-bottom"><i/><span><b>Terhubung lebih dekat</b><small>IndimediaWiFi service portal</small></span></div></div></section>
    <section className="status-section" id="status"><div className="section-heading"><div><span className="section-kicker">NETWORK OVERVIEW</span><h2>Status layanan pusat</h2><p>Informasi status akan mengikuti hasil monitoring server yang terhubung.</p></div><span className="demo-label"><Activity size={14}/> {serverStatus.label}</span></div><div className="status-card"><div className="status-symbol"><Signal size={25}/></div><div className="status-main"><b>Server pusat IndimediaWiFi</b><span>{serverStatus.target || "Belum terhubung ke sistem monitoring"}</span></div><div className="status-state"><i/> {serverStatus.label}</div><div className="status-divider"/><div className="status-meta"><span>STATUS</span><b>{serverStatus.checkedAt ? new Date(serverStatus.checkedAt).toLocaleString("id-ID") : "Menunggu konfigurasi"}</b></div></div><p className="status-note"><CircleHelp size={15}/> Status target diperiksa berkala. Indikator ini tidak menjamin koneksi Wi-Fi pelanggan secara individual.</p></section>
    <section className="help-section" id="bantuan"><div className="help-icon"><Wifi size={24}/></div><div className="help-copy"><h3>Ada kendala koneksi?</h3><p>Periksa status server pusat terlebih dahulu. Jika status pusat normal tetapi koneksi rumah terganggu, tim teknis dapat melakukan pengecekan lanjutan di area pelanggan.</p></div><button className="help-link" onClick={()=>openLogin("customer")}>Buka portal pelanggan <ArrowRight size={16}/></button></section>
    <footer className="footer"><Brand compact/><span>© {new Date().getFullYear()} IndimediaWiFi · Internet murah untuk semua</span><span><i/> Monitoring belum aktif</span></footer>
    {loginType && <div className="modal-backdrop" onMouseDown={event=>{if(event.target===event.currentTarget)openLogin(null);}}><section className="login-modal" role="dialog" aria-modal="true" aria-labelledby="login-title"><button className="modal-close" onClick={()=>openLogin(null)} aria-label="Tutup dialog"><X size={19}/></button><div className="modal-brand"><Brand/></div><div className="modal-kicker">{loginType==="customer"?"CUSTOMER PORTAL":"ADMIN CONTROL CENTER"}</div><h2 id="login-title">{loginType==="customer"?"Login Pelanggan":"Login Admin"}</h2><p className="modal-description">{loginType==="customer"?"Masukkan ID pelanggan untuk melanjutkan ke portal layanan Anda.":"Masuk ke pusat kontrol pengelolaan IndimediaWiFi."}</p><form onSubmit={submitLogin}>{loginType==="customer"?<><label htmlFor="customer-id">ID Pelanggan</label><input id="customer-id" autoFocus value={customerId} onChange={event=>setCustomerId(event.target.value)} placeholder="Contoh: WIFI-00125" autoComplete="username" required/></>:<><label htmlFor="admin-username">Email admin Firebase</label><input id="admin-username" type="email" autoFocus value={username} onChange={event=>setUsername(event.target.value)} placeholder="admin@email.com" autoComplete="username" required/><label htmlFor="admin-password">Password</label><input id="admin-password" type="password" value={password} onChange={event=>setPassword(event.target.value)} placeholder="Password" autoComplete="current-password" required/></>}<button className="button-primary form-submit" type="submit">Login <ArrowRight size={17}/></button></form>{notice && <div className="notice-box" role="status">{notice}</div>}{loginType==="admin" && <div className="demo-credential">Gunakan akun email/password yang dibuat di Firebase Authentication.</div>}<div className="modal-security"><LockKeyhole size={14}/> Autentikasi dan penyimpanan data menggunakan Firebase project IndimediaWifi.</div></section></div>}
  </main>;
}
