import { useState, useRef } from "react";
import * as XLSX from "xlsx";

/* ══════════════════════════════════════════
   TRANSLATIONS
══════════════════════════════════════════ */
const TRANSLATIONS = {
  en: {
    dir: "ltr",
    appSub: "IT Asset Management System",
    dashboard: "Dashboard",
    inventory: "Inventory",
    settings: "Settings",
    exportBtn: "Export",
    importBtn: "Import",
    printBtn: "Print",
    totalAssets: "Total Assets",
    available: "Available",
    inUse: "In Use",
    damaged: "Damaged",
    maintenance: "Maintenance",
    departments: "Departments",
    acrossLabel: (n) => `across ${n} departments`,
    clickManage: "click card to manage",
    assetStatus: "Asset Status",
    deptOverview: "Department Overview",
    categoryOverview: "Category Overview",
    clickRow: "click row → manage",
    allDepts: "All Departments",
    newDept: "+ New Department",
    manageAssets: "Manage assets →",
    addDevice: "+ Add Device",
    allStatus: "All Status",
    allCats: "All Categories",
    searchPlaceholder: "Search name, serial, user, model...",
    noDevices: "No devices found",
    noDevicesHint: (a) => <><strong style={{color:a}}>+ Add Device</strong> to register your first asset</>,
    recordsShown: (n) => `${n} record${n!==1?"s":""} displayed`,
    backDash: "← Dashboard",
    editDevice: "Edit Device",
    registerDevice: "Register New Device",
    department: "Department",
    category: "Category",
    deviceName: "Device Name *",
    brand: "Brand",
    model: "Model",
    serial: "Serial Number",
    status: "Status",
    assignedUser: "Assigned User",
    location: "Location / Room",
    entryDate: "Entry Date",
    notes: "Notes",
    cancel: "Cancel",
    saveChanges: "Save Changes",
    addDeviceBtn: "Add Device",
    addNewDept: "Add New Department",
    editDept: "Edit Department",
    deptNameLabel: "Department Name *",
    iconLabel: "Icon (emoji)",
    colorLabel: "Accent Color",
    catsLabel: "Asset Categories (comma-separated)",
    catsHint: "Leave blank → defaults: Laptop, Monitor, Mouse, Keyboard",
    preview: "Preview",
    createDept: "Create Department",
    saveDeptBtn: "Save Changes",
    nameRequired: "Device name is required",
    deptRequired: "Department name required",
    deviceUpdated: "Device updated ✓",
    deviceAdded: "Device added ✓",
    deviceDeleted: "Device deleted",
    exportedMsg: "Exported to Excel ✓",
    importedMsg: (n) => `Imported ${n} records ✓`,
    deleteConfirm: "Delete this device?",
    deleteDeptConfirm: "Delete this department and all its devices?",
    deptAddedMsg: (n) => `"${n}" added ✓`,
    deptUpdatedMsg: (n) => `"${n}" updated ✓`,
    deptDeletedMsg: "Department deleted",
    total: "Total",
    available2: "Available",
    inUse2: "In Use",
    damaged2: "Damaged",
    avail: "available",
    other: "other",
    dmg: "dmg",
    tableHeaders: ["#","Category","Device Name","Brand","Model","Serial","Status","User","Location","Notes","Date",""],
    deptNamePlaceholder: "e.g. Legal Department",
    catPlaceholder: "Laptop, Monitor, Mouse, Keyboard",
    brandPlaceholder: "Dell / HP / Cisco...",
    modelPlaceholder: "Model number",
    serialPlaceholder: "S/N",
    userPlaceholder: "Employee name",
    locationPlaceholder: "e.g. Server Room B2",
    notesPlaceholder: "Any additional notes...",
    deviceNamePlaceholder: "e.g. Dell PowerEdge R740",
    editBtn: "Edit",
    delBtn: "✕",
    printTitle: (dept) => `IT Asset Inventory${dept?" — "+dept:""} | ${new Date().toLocaleDateString()}`,
    // Settings page
    settingsTitle: "Settings & Administration",
    deptManagement: "Department Management",
    categoryManagement: "Category Management",
    addCategory: "+ Add Category",
    catName: "Category Name",
    catDept: "Assigned Departments",
    catEnabled: "Enabled",
    shared: "Shared (all depts)",
    moveCat: "Move to department",
    cameraType: "Camera Type",
    cameraWireless: "Wireless",
    wirelessNote: "Ezviz cameras are automatically wireless",
    // Camera
    cameraBrand: "Camera Brand",
    // Mobile
    simDetails: "SIM Details",
    simNumber: "SIM Number",
    simCarrier: "Carrier",
    simPlan: "Plan",
  },
  ar: {
    dir: "rtl",
    appSub: "نظام إدارة أصول تقنية المعلومات",
    dashboard: "لوحة التحكم",
    inventory: "الإنفنتوري",
    settings: "الإعدادات",
    exportBtn: "تصدير",
    importBtn: "استيراد",
    printBtn: "طباعة",
    totalAssets: "إجمالي الأجهزة",
    available: "متاح",
    inUse: "مستخدم",
    damaged: "هالك",
    maintenance: "صيانة",
    departments: "الأقسام",
    acrossLabel: (n) => `موزعة على ${n} أقسام`,
    clickManage: "اضغط على البطاقة للإدارة",
    assetStatus: "حالة الأجهزة",
    deptOverview: "نظرة عامة على الأقسام",
    categoryOverview: "نظرة عامة على الفئات",
    clickRow: "اضغط على الصف ← للإدارة",
    allDepts: "جميع الأقسام",
    newDept: "+ قسم جديد",
    manageAssets: "← إدارة الأجهزة",
    addDevice: "+ إضافة جهاز",
    allStatus: "كل الحالات",
    allCats: "كل الفئات",
    searchPlaceholder: "بحث بالاسم، السيريال، المستخدم...",
    noDevices: "لا توجد أجهزة",
    noDevicesHint: (a) => <>اضغط <strong style={{color:a}}>+ إضافة جهاز</strong> لتسجيل أول جهاز</>,
    recordsShown: (n) => `${n} سجل معروض`,
    backDash: "→ لوحة التحكم",
    editDevice: "تعديل الجهاز",
    registerDevice: "تسجيل جهاز جديد",
    department: "القسم",
    category: "الفئة",
    deviceName: "اسم الجهاز *",
    brand: "الماركة",
    model: "الموديل",
    serial: "الرقم التسلسلي",
    status: "الحالة",
    assignedUser: "المستخدم",
    location: "الموقع / الغرفة",
    entryDate: "تاريخ الإدخال",
    notes: "ملاحظات",
    cancel: "إلغاء",
    saveChanges: "حفظ التعديلات",
    addDeviceBtn: "إضافة الجهاز",
    addNewDept: "إضافة قسم جديد",
    editDept: "تعديل القسم",
    deptNameLabel: "اسم القسم *",
    iconLabel: "أيقونة (إيموجي)",
    colorLabel: "اللون",
    catsLabel: "فئات الأجهزة (مفصولة بفواصل)",
    catsHint: "اتركه فارغاً ← افتراضي: لابتوب، شاشة، ماوس، كيبورد",
    preview: "معاينة",
    createDept: "إنشاء القسم",
    saveDeptBtn: "حفظ التعديلات",
    nameRequired: "اسم الجهاز مطلوب",
    deptRequired: "اسم القسم مطلوب",
    deviceUpdated: "تم تعديل الجهاز ✓",
    deviceAdded: "تم إضافة الجهاز ✓",
    deviceDeleted: "تم الحذف",
    exportedMsg: "تم التصدير إلى Excel ✓",
    importedMsg: (n) => `تم استيراد ${n} سجل ✓`,
    deleteConfirm: "تأكيد حذف هذا الجهاز؟",
    deleteDeptConfirm: "حذف هذا القسم وجميع أجهزته؟",
    deptAddedMsg: (n) => `تم إضافة "${n}" ✓`,
    deptUpdatedMsg: (n) => `تم تحديث "${n}" ✓`,
    deptDeletedMsg: "تم حذف القسم",
    total: "الإجمالي",
    available2: "متاح",
    inUse2: "مستخدم",
    damaged2: "هالك",
    avail: "متاح",
    other: "غير متاح",
    dmg: "هالك",
    tableHeaders: ["#","الفئة","اسم الجهاز","الماركة","الموديل","السيريال","الحالة","المستخدم","الموقع","ملاحظات","التاريخ",""],
    deptNamePlaceholder: "مثال: القسم القانوني",
    catPlaceholder: "لابتوب، شاشة، ماوس، كيبورد",
    brandPlaceholder: "Dell / HP / Cisco...",
    modelPlaceholder: "رقم الموديل",
    serialPlaceholder: "S/N",
    userPlaceholder: "اسم الموظف",
    locationPlaceholder: "مثال: غرفة السيرفر B2",
    notesPlaceholder: "أي ملاحظات إضافية...",
    deviceNamePlaceholder: "مثال: Dell PowerEdge R740",
    editBtn: "تعديل",
    delBtn: "✕",
    printTitle: (dept) => `إنفنتوري${dept?" — "+dept:""} | ${new Date().toLocaleDateString("ar-EG")}`,
    settingsTitle: "الإعدادات والإدارة",
    deptManagement: "إدارة الأقسام",
    categoryManagement: "إدارة الفئات",
    addCategory: "+ إضافة فئة",
    catName: "اسم الفئة",
    catDept: "الأقسام المرتبطة",
    catEnabled: "مفعّلة",
    shared: "مشتركة (كل الأقسام)",
    moveCat: "نقل إلى قسم",
    cameraType: "نوع الكاميرا",
    cameraWireless: "لاسلكي",
    wirelessNote: "كاميرات Ezviz تكون لاسلكية تلقائياً",
    cameraBrand: "ماركة الكاميرا",
    simDetails: "تفاصيل SIM",
    simNumber: "رقم SIM",
    simCarrier: "المشغل",
    simPlan: "الباقة",
  },
};

const STATUS_KEYS = {
  Available:   { en:"Available",   ar:"متاح",    color:"#059669", bg:"#d1fae5", dot:"#059669", border:"#6ee7b7" },
  In_Use:      { en:"In Use",      ar:"مستخدم",  color:"#0284c7", bg:"#e0f2fe", dot:"#0284c7", border:"#7dd3fc" },
  Damaged:     { en:"Damaged",     ar:"هالك",    color:"#dc2626", bg:"#fee2e2", dot:"#dc2626", border:"#fca5a5" },
  Maintenance: { en:"Maintenance", ar:"صيانة",   color:"#d97706", bg:"#fef3c7", dot:"#d97706", border:"#fcd34d" },
};

// Category icons map
const CAT_ICONS = {
  "Laptop": "💻", "Monitor": "🖥️", "Mouse": "🖱️", "Keyboard": "⌨️",
  "Server": "🗄️", "Switch": "🔀", "Router": "📡", "Rack": "🏗️",
  "Cables": "🔌", "Camera": "📹", "NVR / DVR": "📼",
  "Laptop Spare Parts": "🔩", "Network Equipment": "🌐",
  "IP Phone": "☎️", "iPad": "📱", "Cash Drawer": "💰",
  "Receipt Printer": "🧾", "PoE Switch": "⚡", "Network Switch": "🔄",
  "Access Point": "📶", "A4 Printer": "🖨️", "HP Charger": "🔋",
  "Dell Charger": "🔋", "HP Laptop Battery": "🔌", "Dell Laptop Battery": "🔌",
  "Mobile Phone": "📱", "Mobile SIM Line": "📲",
};

// SHARED category that appears in all departments
const SHARED_CATS = ["A4 Printer"];

const INIT_DEPTS = [
  { id:"server-room",    nameEn:"Server Room",         nameAr:"غرفة السيرفر",         icon:"🗄️",  color:"#0284c7", cats:["Server","Switch","Router","Rack","Cables","Camera","NVR / DVR","Laptop Spare Parts","iPad","Cash Drawer","Receipt Printer","PoE Switch","Network Switch","Access Point","HP Charger","Dell Charger","HP Laptop Battery","Dell Laptop Battery"] },
  { id:"general-mgmt",  nameEn:"General Management",  nameAr:"الإدارة العامة",        icon:"🏢",  color:"#7c3aed", cats:["Laptop","Monitor","Switch","Mouse","Keyboard","Mobile Phone","Mobile SIM Line"] },
  { id:"it-dept",       nameEn:"IT",                  nameAr:"IT",                    icon:"💻",  color:"#0369a1", cats:["Laptop","Monitor","Switch","Mouse","Keyboard","Server","Network Equipment","Mobile Phone","Mobile SIM Line"] },
  { id:"finance-karak", nameEn:"Finance – Karak Boy", nameAr:"المالية – Karak Boy",   icon:"💳",  color:"#059669", cats:["Laptop","Monitor","Mouse","Keyboard","Mobile Phone","Mobile SIM Line"] },
  { id:"finance-kimbo", nameEn:"Finance – Kimbo",     nameAr:"المالية – Kimbo",       icon:"📊",  color:"#0d9488", cats:["Laptop","Monitor","Mouse","Keyboard","Mobile Phone","Mobile SIM Line"] },
  { id:"cost-control",  nameEn:"Cost Control",        nameAr:"Cost Control",          icon:"📉",  color:"#b45309", cats:["Laptop","Monitor","Mouse","Keyboard","Mobile Phone","Mobile SIM Line"] },
  { id:"hr",            nameEn:"HR",                  nameAr:"HR",                    icon:"👥",  color:"#be185d", cats:["Laptop","Monitor","Mouse","Keyboard","Mobile Phone","Mobile SIM Line"] },
  { id:"procurement",   nameEn:"Procurement",         nameAr:"المشتريات",             icon:"🛒",  color:"#c2410c", cats:["Laptop","Monitor","Mouse","Keyboard","Mobile Phone","Mobile SIM Line"] },
  { id:"reception",     nameEn:"Reception",           nameAr:"الاستقبال",             icon:"🛎️", color:"#6d28d9", cats:["Laptop","Monitor","Mouse","Keyboard","IP Phone","Mobile Phone","Mobile SIM Line"] },
  { id:"sales",         nameEn:"Sales",               nameAr:"المبيعات",              icon:"📈",  color:"#15803d", cats:["Laptop","Monitor","Mouse","Keyboard","Mobile Phone","Mobile SIM Line"] },
  { id:"builders",      nameEn:"Builders",            nameAr:"Builders",              icon:"🏗️",  color:"#92400e", cats:["Laptop","Monitor","Mouse","Keyboard","Mobile Phone","Mobile SIM Line"] },
  { id:"supply-chain",  nameEn:"Supply Chain",        nameAr:"سلسلة التوريد",         icon:"🔗",  color:"#065f46", cats:["Laptop","Monitor","Mouse","Keyboard","Mobile Phone","Mobile SIM Line"] },
  { id:"wechai",        nameEn:"Wechai",              nameAr:"Wechai",                icon:"⚙️",  color:"#1e3a5f", cats:["Laptop","Monitor","Mouse","Keyboard","Mobile Phone","Mobile SIM Line"] },
];

const newItem = (deptId, cats) => ({
  id: crypto.randomUUID(), deptId, category: cats[0]||"",
  name:"", brand:"", model:"", serial:"",
  status:"Available", user:"", location:"", notes:"",
  date: new Date().toISOString().slice(0,10),
  // camera extras
  cameraType:"", cameraBrand:"", cameraWireless:false,
  // sim extras
  simNumber:"", simCarrier:"", simPlan:"",
});

/* ══ LIGHT THEME TOKENS ══ */
const L = {
  bg:       "#f0f4f8",
  surface:  "#ffffff",
  card:     "#ffffff",
  cardHov:  "#f8faff",
  border:   "#e2e8f0",
  borderHi: "#cbd5e1",
  accent:   "#1d4ed8",
  accentBg: "#eff6ff",
  text:     "#0f172a",
  sub:      "#475569",
  muted:    "#94a3b8",
  font:     "'IBM Plex Sans','Segoe UI',sans-serif",
  mono:     "'IBM Plex Mono','Courier New',monospace",
  shadow:   "0 1px 4px #0000000e, 0 4px 16px #0000000a",
  shadowMd: "0 4px 24px #00000014",
};

/* ══ MINI COMPONENTS ══ */
function StatusBadge({ status, lang }) {
  const s = STATUS_KEYS[status] || STATUS_KEYS.Available;
  const label = lang === "ar" ? s.ar : s.en;
  return (
    <span style={{ display:"inline-flex", alignItems:"center", gap:5,
      background:s.bg, color:s.color,
      border:`1px solid ${s.border}`,
      borderRadius:20, padding:"3px 10px", fontSize:11, fontWeight:700,
      letterSpacing:.3, whiteSpace:"nowrap" }}>
      <span style={{ width:6, height:6, borderRadius:"50%", background:s.dot, flexShrink:0 }} />
      {label}
    </span>
  );
}

function KpiCard({ label, value, color, icon, sub }) {
  return (
    <div style={{ background:L.card, border:`1px solid ${L.border}`, borderRadius:14,
      padding:"18px 22px", flex:1, minWidth:120, position:"relative", overflow:"hidden",
      display:"flex", flexDirection:"column", gap:3, boxShadow:L.shadow }}>
      <div style={{ position:"absolute", top:-6, right:10, fontSize:44, opacity:.06, pointerEvents:"none" }}>{icon}</div>
      <span style={{ fontSize:10, fontWeight:700, color:L.muted, textTransform:"uppercase", letterSpacing:1.2 }}>{label}</span>
      <span style={{ fontSize:30, fontWeight:800, color, fontFamily:L.mono, lineHeight:1.1 }}>{value}</span>
      {sub && <span style={{ fontSize:11, color:L.muted }}>{sub}</span>}
      <div style={{ position:"absolute", bottom:0, left:0, right:0, height:3,
        background:`linear-gradient(90deg,${color},${color}33)`, borderRadius:"0 0 14px 14px" }} />
    </div>
  );
}

function SectionLabel({ children }) {
  return (
    <div style={{ fontSize:10, fontWeight:700, letterSpacing:2, textTransform:"uppercase",
      color:L.muted, marginBottom:12, display:"flex", alignItems:"center", gap:8 }}>
      <span style={{ display:"inline-block", width:16, height:2, background:L.accent, borderRadius:2 }} />
      {children}
    </div>
  );
}

function Modal({ children, onClose, dir, wide }) {
  return (
    <div style={{ position:"fixed", inset:0, background:"#00000055", zIndex:1000,
      display:"flex", alignItems:"center", justifyContent:"center", backdropFilter:"blur(4px)" }}
      onClick={e => { if (e.target === e.currentTarget) onClose(); }}>
      <div style={{ background:L.surface, border:`1px solid ${L.borderHi}`, borderRadius:18,
        padding:"28px 32px", width:"94%", maxWidth:wide?860:640, maxHeight:"92vh",
        overflowY:"auto", boxShadow:L.shadowMd, position:"relative", direction:dir }}>
        <button onClick={onClose} style={{ position:"absolute", top:14, right:14,
          background:L.bg, border:`1px solid ${L.border}`, color:L.sub, width:28, height:28,
          borderRadius:"50%", cursor:"pointer", fontSize:14, display:"flex",
          alignItems:"center", justifyContent:"center" }}>✕</button>
        {children}
      </div>
    </div>
  );
}

function Field({ label, children }) {
  return (
    <div>
      <label style={{ display:"block", fontSize:10, fontWeight:700, color:L.sub,
        letterSpacing:1.1, textTransform:"uppercase", marginBottom:5 }}>{label}</label>
      {children}
    </div>
  );
}

function Donut({ segments, total, size=130 }) {
  const r = 38, cx = size/2, cy = size/2, circ = 2*Math.PI*r;
  let offset = 0;
  return (
    <svg width={size} height={size} viewBox={`0 0 ${size} ${size}`}>
      <circle cx={cx} cy={cy} r={r} fill="none" stroke={L.border} strokeWidth={10} />
      {segments.map((s,i) => {
        if (!s.value) return null;
        const dash = (s.value/(total||1)) * circ;
        const rotate = (offset/(total||1)) * 360 - 90;
        offset += s.value;
        return <circle key={i} cx={cx} cy={cy} r={r} fill="none" stroke={s.color} strokeWidth={10}
          strokeDasharray={`${dash} ${circ-dash}`} strokeDashoffset={circ*0.25}
          style={{ transformOrigin:`${cx}px ${cy}px`, transform:`rotate(${rotate}deg)`, transition:"all .5s" }}
          strokeLinecap="butt" />;
      })}
      <text x={cx} y={cy-4} textAnchor="middle" fill={L.text} fontSize={20} fontWeight={800} fontFamily={L.mono}>{total}</text>
      <text x={cx} y={cy+13} textAnchor="middle" fill={L.muted} fontSize={8} fontWeight={600} letterSpacing={1} fontFamily={L.font}>TOTAL</text>
    </svg>
  );
}

/* ══════════════════════════════════════════
   MAIN APP
══════════════════════════════════════════ */
export default function App() {
  const [lang, setLang]         = useState("en");
  const [depts, setDepts]       = useState(INIT_DEPTS);
  const [items, setItems]       = useState([]);
  const [view, setView]         = useState("dashboard");
  const [activeDeptId, setActiveDeptId] = useState(null);
  const [form, setForm]         = useState(null);
  const [editId, setEditId]     = useState(null);
  const [showForm, setShowForm] = useState(false);
  const [showDeptModal, setShowDeptModal] = useState(false);
  const [editingDept, setEditingDept] = useState(null); // dept being edited
  const [deptForm, setDeptForm] = useState({ nameEn:"", nameAr:"", icon:"🖥️", color:"#1d4ed8", cats:"" });
  const [search, setSearch]     = useState("");
  const [fStatus, setFStatus]   = useState("All");
  const [fCat, setFCat]         = useState("All");
  const [toast, setToast]       = useState(null);
  const [settingsTab, setSettingsTab] = useState("departments"); // departments | categories
  const [newCatName, setNewCatName] = useState("");
  const fileRef = useRef();

  const t = TRANSLATIONS[lang];
  const dir = t.dir;
  const deptName = (d) => lang === "ar" ? (d.nameAr||d.nameEn) : d.nameEn;

  const showMsg = (msg, type="ok") => { setToast({msg,type}); setTimeout(()=>setToast(null),2800); };

  const activeDept = depts.find(d => d.id === activeDeptId);

  // All unique cats across all departments + shared
  const allCats = [...new Set([...depts.flatMap(d=>d.cats), ...SHARED_CATS])].sort();

  // Dept's effective cats = own cats + shared
  const deptCats = (d) => d ? [...new Set([...d.cats, ...SHARED_CATS])] : [];
  const currentCats = deptCats(activeDept);

  const deptCount     = (id) => items.filter(i=>i.deptId===id).length;
  const deptSt        = (id,st) => items.filter(i=>i.deptId===id&&i.status===st).length;
  const totalBySt     = Object.fromEntries(Object.keys(STATUS_KEYS).map(k=>[k,items.filter(i=>i.status===k).length]));
  const donutSegs     = Object.entries(STATUS_KEYS).map(([k,s])=>({color:s.color,value:totalBySt[k]}));
  const catCount      = (cat) => items.filter(i=>i.category===cat).length;

  const rows = items.filter(it => {
    if (activeDeptId && it.deptId !== activeDeptId) return false;
    const q = search.toLowerCase();
    if (q && ![it.name,it.brand,it.model,it.serial,it.user,it.notes].join(" ").toLowerCase().includes(q)) return false;
    if (fStatus!=="All" && it.status!==fStatus) return false;
    if (fCat!=="All" && it.category!==fCat) return false;
    return true;
  });

  const openAddItem = () => {
    if (!activeDept) return;
    setForm(newItem(activeDeptId, deptCats(activeDept)));
    setEditId(null); setShowForm(true);
  };
  const openEdit = (item) => { setForm({...item}); setEditId(item.id); setShowForm(true); };

  // Camera logic
  const isCameraCategory = (cat) => cat === "Camera";
  const isMobilePhone = (cat) => cat === "Mobile Phone";
  const isMobileSIM = (cat) => cat === "Mobile SIM Line";

  const saveItem = () => {
    if (!form.name.trim()) { showMsg(t.nameRequired,"err"); return; }
    // Auto wireless for Ezviz
    const finalForm = {...form};
    if (isCameraCategory(form.category) && form.cameraBrand === "Ezviz") {
      finalForm.cameraWireless = true;
    }
    if (editId) setItems(p=>p.map(i=>i.id===editId?{...finalForm,id:editId}:i));
    else        setItems(p=>[...p,{...finalForm,id:crypto.randomUUID()}]);
    showMsg(editId ? t.deviceUpdated : t.deviceAdded);
    setShowForm(false);
  };

  const delItem = (id) => {
    if (!confirm(t.deleteConfirm)) return;
    setItems(p=>p.filter(i=>i.id!==id));
    showMsg(t.deviceDeleted,"err");
  };

  const exportXLSX = () => {
    const src = activeDeptId ? rows : items;
    const ws = XLSX.utils.json_to_sheet(src.map((it,i)=>({
      "#":i+1, Department:depts.find(d=>d.id===it.deptId)?.nameEn||it.deptId,
      Category:it.category, Name:it.name, Brand:it.brand, Model:it.model,
      Serial:it.serial, Status:it.status, User:it.user,
      Location:it.location, Notes:it.notes, Date:it.date,
      CameraType:it.cameraType||"", CameraWireless:it.cameraWireless?"Yes":"",
      SIMNumber:it.simNumber||"", SIMCarrier:it.simCarrier||"", SIMPlan:it.simPlan||"",
    })));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Inventory");
    XLSX.writeFile(wb, `UEG_Inventory_${new Date().toISOString().slice(0,10)}.xlsx`);
    showMsg(t.exportedMsg);
  };

  const importXLSX = (e) => {
    const file = e.target.files[0]; if (!file) return;
    const rd = new FileReader();
    rd.onload = ev => {
      const wb = XLSX.read(ev.target.result,{type:"binary"});
      const rws = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]]);
      const imp = rws.map(r => {
        const dept = depts.find(d=>d.nameEn===r["Department"])||depts[0];
        return { id:crypto.randomUUID(), deptId:dept.id, category:r["Category"]||"",
          name:r["Name"]||"", brand:r["Brand"]||"", model:r["Model"]||"",
          serial:r["Serial"]||"", status:r["Status"]||"Available",
          user:r["User"]||"", location:r["Location"]||"",
          notes:r["Notes"]||"", date:r["Date"]||new Date().toISOString().slice(0,10),
          cameraType:r["CameraType"]||"", cameraWireless:r["CameraWireless"]==="Yes",
          simNumber:r["SIMNumber"]||"", simCarrier:r["SIMCarrier"]||"", simPlan:r["SIMPlan"]||"",
        };
      });
      setItems(p=>[...p,...imp]);
      showMsg(t.importedMsg(imp.length));
    };
    rd.readAsBinaryString(file);
    e.target.value="";
  };

  const printView = () => {
    const src = activeDeptId ? rows : items;
    const dname = activeDept ? deptName(activeDept) : "";
    const trs = src.map((it,i)=>`<tr>
      <td>${i+1}</td><td>${depts.find(d=>d.id===it.deptId)?.nameEn||it.deptId}</td>
      <td>${it.category}</td><td>${it.name}</td><td>${it.brand}</td>
      <td>${it.model}</td><td>${it.serial}</td><td>${it.status}</td>
      <td>${it.user}</td><td>${it.location}</td><td>${it.notes}</td><td>${it.date}</td></tr>`).join("");
    const w = window.open("","_blank");
    w.document.write(`<html><head><title>UEG – IT Inventory</title>
    <style>body{font-family:Arial;font-size:11px}h2,h3{text-align:center;margin:3px}
    table{width:100%;border-collapse:collapse;margin-top:10px}
    th{background:#1e3a5f;color:#fff;padding:6px;border:1px solid #ccc;font-size:10px}
    td{padding:5px;border:1px solid #e5e5e5}tr:nth-child(even){background:#f0f6ff}</style></head>
    <body><h2>Unlimited Egypt Group</h2><h3>${t.printTitle(dname)}</h3>
    <table><thead><tr><th>#</th><th>Dept</th><th>Category</th><th>Name</th><th>Brand</th>
    <th>Model</th><th>Serial</th><th>Status</th><th>User</th><th>Location</th><th>Notes</th><th>Date</th>
    </tr></thead><tbody>${trs}</tbody></table>
    <script>window.print()</script></body></html>`);
    w.document.close();
  };

  // ── DEPT CRUD ──
  const openAddDept = () => {
    setEditingDept(null);
    setDeptForm({nameEn:"",nameAr:"",icon:"🖥️",color:"#1d4ed8",cats:""});
    setShowDeptModal(true);
  };

  const openEditDept = (d) => {
    setEditingDept(d);
    setDeptForm({nameEn:d.nameEn,nameAr:d.nameAr||"",icon:d.icon||"🖥️",color:d.color||"#1d4ed8",cats:d.cats.join(", ")});
    setShowDeptModal(true);
  };

  const deleteDept = (id) => {
    if (!confirm(t.deleteDeptConfirm)) return;
    setDepts(p=>p.filter(d=>d.id!==id));
    setItems(p=>p.filter(i=>i.deptId!==id));
    if (activeDeptId===id) setActiveDeptId(null);
    showMsg(t.deptDeletedMsg,"err");
  };

  const saveDept = () => {
    if (!deptForm.nameEn.trim()) { showMsg(t.deptRequired,"err"); return; }
    const cats = deptForm.cats.split(",").map(s=>s.trim()).filter(Boolean);
    if (!cats.length) cats.push("Laptop","Monitor","Mouse","Keyboard");
    if (editingDept) {
      setDepts(p=>p.map(d=>d.id===editingDept.id ? {
        ...d, nameEn:deptForm.nameEn, nameAr:deptForm.nameAr||deptForm.nameEn,
        icon:deptForm.icon||"🖥️", color:deptForm.color, cats,
      } : d));
      showMsg(t.deptUpdatedMsg(deptForm.nameEn));
    } else {
      const nd = {
        id: deptForm.nameEn.toLowerCase().replace(/\s+/g,"-")+"-"+Date.now(),
        nameEn: deptForm.nameEn, nameAr: deptForm.nameAr || deptForm.nameEn,
        icon: deptForm.icon||"🖥️", color: deptForm.color, cats,
      };
      setDepts(p=>[...p,nd]);
      showMsg(t.deptAddedMsg(nd.nameEn));
    }
    setDeptForm({nameEn:"",nameAr:"",icon:"🖥️",color:"#1d4ed8",cats:""});
    setShowDeptModal(false);
    setEditingDept(null);
  };

  const goInventory = (deptId, catFilter) => {
    setActiveDeptId(deptId||null);
    setSearch("");
    setFStatus("All");
    setFCat(catFilter||"All");
    setView("inventory");
    if (!deptId && depts.length) setActiveDeptId(depts[0]?.id);
  };

  const inp = {
    width:"100%", padding:"9px 13px", borderRadius:8,
    border:`1px solid ${L.borderHi}`, background:L.bg,
    color:L.text, fontSize:13, outline:"none", boxSizing:"border-box",
    fontFamily:L.font,
  };
  const selStyle = {
    padding:"9px 13px", borderRadius:9, border:`1px solid ${L.borderHi}`,
    background:L.bg, color:L.text, fontSize:12, cursor:"pointer", fontFamily:L.font,
  };

  // Category summary for dashboard
  const catSummary = allCats
    .map(cat=>({ cat, icon:CAT_ICONS[cat]||"📦", count:catCount(cat) }))
    .filter(x=>x.count>0 || true) // show all, even 0
    .sort((a,b)=>b.count-a.count);

  /* ════════════════════ RENDER ════════════════════ */
  return (
    <div style={{ minHeight:"100vh", background:L.bg, color:L.text, fontFamily:L.font, direction:dir }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Sans:wght@400;500;600;700;800&family=IBM+Plex+Mono:wght@400;500;600;700&display=swap');
        *{box-sizing:border-box;margin:0;padding:0}
        ::-webkit-scrollbar{width:5px;height:5px}
        ::-webkit-scrollbar-track{background:#f0f4f8}
        ::-webkit-scrollbar-thumb{background:#cbd5e1;border-radius:4px}
        .dept-card:hover{background:#f8faff !important;border-color:#93c5fd !important;transform:translateY(-2px);box-shadow:0 8px 28px #1d4ed81a !important}
        .tr-row:hover td{background:#f8faff !important}
        .tab-btn:hover{background:#eff6ff !important;color:#1d4ed8 !important}
        .action-btn:hover{opacity:.85;transform:translateY(-1px)}
        .cat-card:hover{border-color:#93c5fd !important;transform:translateY(-2px);box-shadow:0 6px 20px #1d4ed81a !important}
        select option{background:#fff;color:#0f172a}
        input[type=date]::-webkit-calendar-picker-indicator{opacity:.5}
        .settings-row:hover{background:#f8faff !important}
      `}</style>

      {/* ── TOPBAR ── */}
      <header style={{ background:L.surface, borderBottom:`1px solid ${L.border}`,
        padding:"0 24px", height:60, display:"flex", alignItems:"center",
        justifyContent:"space-between", position:"sticky", top:0, zIndex:200,
        boxShadow:"0 1px 8px #0000000d" }}>

        {/* Logo */}
        <div style={{ display:"flex", alignItems:"center", gap:12 }}>
          <div style={{ width:38, height:38, borderRadius:10,
            background:"linear-gradient(135deg,#1d4ed8,#7c3aed)",
            display:"flex", alignItems:"center", justifyContent:"center",
            boxShadow:"0 2px 12px #1d4ed844" }}>
            <span style={{ color:"#fff", fontWeight:900, fontSize:14, fontFamily:L.mono, letterSpacing:-1 }}>UEG</span>
          </div>
          <div>
            <div style={{ fontWeight:800, fontSize:14, color:L.text, letterSpacing:.1 }}>Unlimited Egypt Group</div>
            <div style={{ fontSize:9, color:L.muted, letterSpacing:1.4, textTransform:"uppercase", fontFamily:L.mono }}>{t.appSub}</div>
          </div>
        </div>

        {/* Nav */}
        <div style={{ display:"flex", gap:4 }}>
          {[["dashboard",t.dashboard,"📊"],["inventory",t.inventory,"📦"],["settings",t.settings,"⚙️"]].map(([v,lb,ic])=>(
            <button key={v} className="tab-btn" onClick={()=>{ setView(v); if(v==="inventory"&&!activeDeptId) setActiveDeptId(depts[0]?.id); }}
              style={{ padding:"7px 18px", border:`1.5px solid ${view===v?"#1d4ed888":L.border}`,
                borderRadius:9, cursor:"pointer", background:view===v?"#eff6ff":L.surface,
                color:view===v?"#1d4ed8":L.sub, fontWeight:view===v?700:500,
                fontSize:12, fontFamily:L.font, transition:"all .15s", display:"flex", alignItems:"center", gap:6 }}>
              {ic} {lb}
            </button>
          ))}
        </div>

        {/* Actions + lang toggle */}
        <div style={{ display:"flex", gap:8, alignItems:"center" }}>
          {[
            [t.exportBtn,"#059669","↑",exportXLSX],
            [t.importBtn,"#0284c7","↓",()=>fileRef.current.click()],
            [t.printBtn, "#7c3aed","⎙",printView],
          ].map(([lb,c,ic,fn])=>(
            <button key={lb} className="action-btn" onClick={fn}
              style={{ padding:"6px 14px", border:`1.5px solid ${c}44`,
                background:c+"12", color:c, borderRadius:8, cursor:"pointer",
                fontWeight:700, fontSize:11, fontFamily:L.mono,
                transition:"all .15s", display:"flex", alignItems:"center", gap:4 }}>
              {ic} {lb}
            </button>
          ))}
          <div style={{ display:"flex", border:`1.5px solid ${L.borderHi}`, borderRadius:8, overflow:"hidden" }}>
            {["en","ar"].map(l=>(
              <button key={l} onClick={()=>setLang(l)}
                style={{ padding:"6px 12px", border:"none", cursor:"pointer",
                  background:lang===l?"#1d4ed8":L.surface,
                  color:lang===l?"#fff":L.sub, fontWeight:700, fontSize:11,
                  fontFamily:L.mono, transition:"all .15s" }}>
                {l.toUpperCase()}
              </button>
            ))}
          </div>
          <input ref={fileRef} type="file" accept=".xlsx,.xls" style={{display:"none"}} onChange={importXLSX} />
        </div>
      </header>

      {/* ══════════════ DASHBOARD ══════════════ */}
      {view === "dashboard" && (
        <main style={{ padding:"28px 24px", maxWidth:1440, margin:"0 auto" }}>

          {/* KPI */}
          <div style={{ display:"flex", gap:12, marginBottom:24, flexWrap:"wrap" }}>
            <KpiCard label={t.totalAssets}  value={items.length}              color="#1d4ed8" icon="💻" sub={t.acrossLabel(depts.length)} />
            <KpiCard label={t.available}    value={totalBySt.Available}       color="#059669" icon="✅" />
            <KpiCard label={t.inUse}        value={totalBySt.In_Use}          color="#0284c7" icon="👤" />
            <KpiCard label={t.damaged}      value={totalBySt.Damaged}         color="#dc2626" icon="⚠️" />
            <KpiCard label={t.maintenance}  value={totalBySt.Maintenance}     color="#d97706" icon="🔧" />
            <KpiCard label={t.departments}  value={depts.length}              color="#7c3aed" icon="🏢" sub={t.clickManage} />
          </div>

          {/* Donut + Dept bars */}
          <div style={{ display:"grid", gridTemplateColumns:"260px 1fr", gap:16, marginBottom:24 }}>
            <div style={{ background:L.card, border:`1px solid ${L.border}`, borderRadius:16,
              padding:22, display:"flex", flexDirection:"column", gap:18,
              alignItems:"center", boxShadow:L.shadow }}>
              <SectionLabel>{t.assetStatus}</SectionLabel>
              <Donut segments={donutSegs} total={items.length} size={130} />
              <div style={{ width:"100%", display:"flex", flexDirection:"column", gap:9 }}>
                {Object.entries(STATUS_KEYS).map(([k,s])=>(
                  <div key={k} style={{ display:"flex", justifyContent:"space-between", alignItems:"center" }}>
                    <div style={{ display:"flex", alignItems:"center", gap:8 }}>
                      <div style={{ width:8,height:8,borderRadius:"50%",background:s.color }} />
                      <span style={{ fontSize:12, color:L.sub }}>{lang==="ar"?s.ar:s.en}</span>
                    </div>
                    <div style={{ display:"flex", gap:6, alignItems:"center" }}>
                      <span style={{ fontSize:13, fontWeight:700, color:s.color, fontFamily:L.mono }}>{totalBySt[k]}</span>
                      <span style={{ fontSize:10, color:L.muted, fontFamily:L.mono }}>
                        {items.length ? Math.round(totalBySt[k]/items.length*100) : 0}%
                      </span>
                    </div>
                  </div>
                ))}
              </div>
            </div>

            <div style={{ background:L.card, border:`1px solid ${L.border}`, borderRadius:16,
              padding:22, boxShadow:L.shadow }}>
              <div style={{ display:"flex", justifyContent:"space-between", alignItems:"flex-start", marginBottom:16 }}>
                <SectionLabel>{t.deptOverview}</SectionLabel>
                <span style={{ fontSize:10, color:L.muted }}>{t.clickRow}</span>
              </div>
              <div style={{ display:"flex", flexDirection:"column", gap:9, maxHeight:290, overflowY:"auto" }}>
                {depts.map(d => {
                  const tot = deptCount(d.id);
                  const pct = items.length ? tot/items.length : 0;
                  const avail = deptSt(d.id,"Available");
                  const dmg   = deptSt(d.id,"Damaged");
                  return (
                    <div key={d.id} onClick={()=>goInventory(d.id)} style={{ cursor:"pointer",
                      padding:"9px 13px", borderRadius:10, border:`1px solid ${L.border}`,
                      background:L.bg, transition:"all .15s" }}
                      onMouseEnter={e=>{e.currentTarget.style.background="#eff6ff";e.currentTarget.style.borderColor="#93c5fd"}}
                      onMouseLeave={e=>{e.currentTarget.style.background=L.bg;e.currentTarget.style.borderColor=L.border}}>
                      <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:5 }}>
                        <div style={{ display:"flex", alignItems:"center", gap:7 }}>
                          <span style={{ fontSize:14 }}>{d.icon}</span>
                          <span style={{ fontSize:12, fontWeight:600, color:L.text }}>{deptName(d)}</span>
                        </div>
                        <div style={{ display:"flex", gap:8, alignItems:"center" }}>
                          {dmg>0 && <span style={{ fontSize:10, color:"#dc2626", fontFamily:L.mono }}>{dmg} {t.dmg}</span>}
                          <span style={{ fontSize:14, fontWeight:700, color:d.color, fontFamily:L.mono }}>{tot}</span>
                        </div>
                      </div>
                      <div style={{ height:4, background:L.border, borderRadius:99, overflow:"hidden" }}>
                        <div style={{ height:"100%", width:`${pct*100}%`, borderRadius:99, transition:"width .6s",
                          background:`linear-gradient(90deg,${d.color},${d.color}66)` }} />
                      </div>
                      <div style={{ display:"flex", gap:10, marginTop:5 }}>
                        <span style={{ fontSize:9, color:"#059669", fontFamily:L.mono }}>{avail} {t.avail}</span>
                        <span style={{ fontSize:9, color:L.muted, fontFamily:L.mono }}>{tot-avail} {t.other}</span>
                      </div>
                    </div>
                  );
                })}
              </div>
            </div>
          </div>

          {/* Category Overview Cards — quick-access */}
          <div style={{ marginBottom:24 }}>
            <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:14 }}>
              <SectionLabel>{t.categoryOverview}</SectionLabel>
            </div>
            <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fill,minmax(150px,1fr))", gap:10 }}>
              {catSummary.map(({cat,icon,count}) => (
                <div key={cat} className="cat-card"
                  onClick={()=>{ setActiveDeptId(null); setFCat(cat); setFStatus("All"); setSearch(""); setView("inventory"); if(depts.length) setActiveDeptId(depts[0]?.id); }}
                  style={{ background:L.card, border:`1px solid ${L.border}`, borderRadius:12,
                    padding:"14px 16px", cursor:"pointer", transition:"all .2s", boxShadow:L.shadow,
                    display:"flex", flexDirection:"column", gap:6 }}>
                  <div style={{ fontSize:26 }}>{icon}</div>
                  <div style={{ fontSize:11, fontWeight:700, color:L.text, lineHeight:1.3 }}>{cat}</div>
                  <div style={{ fontSize:20, fontWeight:800, color:L.accent, fontFamily:L.mono }}>{count}</div>
                  <div style={{ fontSize:9, color:L.muted, textTransform:"uppercase", letterSpacing:.8 }}>
                    {count === 1 ? "item" : "items"}
                  </div>
                </div>
              ))}
            </div>
          </div>

          {/* Dept Cards */}
          <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:14 }}>
            <SectionLabel>{t.allDepts}</SectionLabel>
            <button onClick={openAddDept}
              style={{ padding:"7px 18px", border:`1.5px solid #1d4ed888`,
                background:"#eff6ff", color:"#1d4ed8", borderRadius:9, cursor:"pointer",
                fontWeight:700, fontSize:12, fontFamily:L.font, letterSpacing:.2 }}>
              {t.newDept}
            </button>
          </div>
          <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fill,minmax(215px,1fr))", gap:12 }}>
            {depts.map(d => {
              const st = { total:deptCount(d.id), Available:deptSt(d.id,"Available"), In_Use:deptSt(d.id,"In_Use"), Damaged:deptSt(d.id,"Damaged") };
              return (
                <div key={d.id} className="dept-card" onClick={()=>goInventory(d.id)}
                  style={{ background:L.card, border:`1px solid ${L.border}`, borderRadius:16,
                    padding:18, cursor:"pointer", transition:"all .2s", position:"relative",
                    overflow:"hidden", boxShadow:L.shadow }}>
                  <div style={{ position:"absolute", top:0, left:0, right:0, height:3,
                    background:`linear-gradient(90deg,${d.color},${d.color}33)` }} />
                  <div style={{ display:"flex", alignItems:"center", gap:10, marginBottom:13 }}>
                    <div style={{ width:38, height:38, borderRadius:10, fontSize:18,
                      background:d.color+"18", border:`1px solid ${d.color}30`,
                      display:"flex", alignItems:"center", justifyContent:"center", flexShrink:0 }}>{d.icon}</div>
                    <div>
                      <div style={{ fontWeight:700, fontSize:12, color:L.text, lineHeight:1.3 }}>{deptName(d)}</div>
                      <div style={{ fontSize:9, color:L.muted, marginTop:2 }}>
                        {d.cats.slice(0,3).join(" · ")}{d.cats.length>3?"…":""}
                      </div>
                    </div>
                  </div>
                  <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:6 }}>
                    {[[t.total,st.total,L.text,"#f0f4f8"],[t.available2,st.Available,"#059669","#f0fdf4"],[t.inUse2,st.In_Use,"#0284c7","#f0f9ff"],[t.damaged2,st.Damaged,"#dc2626","#fff5f5"]].map(([lb,val,cl,bg])=>(
                      <div key={lb} style={{ background:bg, borderRadius:8, padding:"7px 10px", border:`1px solid ${cl}22` }}>
                        <div style={{ fontSize:18, fontWeight:800, color:cl, fontFamily:L.mono, lineHeight:1 }}>{val}</div>
                        <div style={{ fontSize:9, color:L.muted, fontWeight:600, marginTop:2, textTransform:"uppercase", letterSpacing:.8 }}>{lb}</div>
                      </div>
                    ))}
                  </div>
                  <div style={{ marginTop:11, fontSize:10, color:d.color, fontWeight:700 }}>{t.manageAssets}</div>
                </div>
              );
            })}
          </div>
        </main>
      )}

      {/* ══════════════ INVENTORY ══════════════ */}
      {view === "inventory" && (
        <main style={{ padding:"20px 24px", maxWidth:1600, margin:"0 auto" }}>
          {/* Dept scroll tabs */}
          <div style={{ display:"flex", gap:6, overflowX:"auto", paddingBottom:8, marginBottom:18, scrollbarWidth:"none" }}>
            {depts.map(d=>{
              const active = activeDeptId===d.id;
              return (
                <button key={d.id} className="tab-btn" onClick={()=>{setActiveDeptId(d.id);setFCat("All");setFStatus("All");setSearch("");}}
                  style={{ padding:"7px 16px", border:`1.5px solid ${active?d.color+"88":L.border}`,
                    borderRadius:30, cursor:"pointer", whiteSpace:"nowrap",
                    background:active?d.color+"18":L.surface,
                    color:active?d.color:L.sub, fontWeight:active?700:500,
                    fontSize:12, fontFamily:L.font, transition:"all .15s",
                    display:"flex", alignItems:"center", gap:6, boxShadow:active?`0 2px 10px ${d.color}33`:"none" }}>
                  {d.icon} {deptName(d)}
                  <span style={{ background:active?d.color+"33":L.border, borderRadius:10, padding:"1px 7px", fontSize:10, fontFamily:L.mono }}>{deptCount(d.id)}</span>
                </button>
              );
            })}
          </div>

          {/* Status pills */}
          {activeDept && (
            <div style={{ display:"flex", gap:7, marginBottom:14, flexWrap:"wrap" }}>
              <button onClick={()=>setFStatus("All")}
                style={{ padding:"5px 14px", border:`1.5px solid ${fStatus==="All"?"#1d4ed888":L.border}`,
                  borderRadius:20, cursor:"pointer", background:fStatus==="All"?"#eff6ff":L.surface,
                  color:fStatus==="All"?"#1d4ed8":L.sub, fontSize:11, fontWeight:600, fontFamily:L.font }}>
                {t.allStatus} ({deptCount(activeDeptId)})
              </button>
              {Object.entries(STATUS_KEYS).map(([k,s])=>{
                const n = deptSt(activeDeptId,k); const active = fStatus===k;
                return (
                  <button key={k} onClick={()=>setFStatus(active?"All":k)}
                    style={{ padding:"5px 14px", border:`1.5px solid ${active?s.color+"88":L.border}`,
                      borderRadius:20, cursor:"pointer", background:active?s.bg:L.surface,
                      color:active?s.color:L.sub, fontSize:11, fontWeight:600, fontFamily:L.font,
                      display:"flex", alignItems:"center", gap:5 }}>
                    <span style={{ width:6,height:6,borderRadius:"50%",background:s.color,flexShrink:0 }} />
                    {lang==="ar"?s.ar:s.en} ({n})
                  </button>
                );
              })}
            </div>
          )}

          {/* Toolbar */}
          <div style={{ display:"flex", gap:10, marginBottom:14, flexWrap:"wrap", alignItems:"center" }}>
            <div style={{ position:"relative", flex:1, minWidth:220 }}>
              <span style={{ position:"absolute", [dir==="rtl"?"right":"left"]:12, top:"50%",
                transform:"translateY(-50%)", color:L.muted, fontSize:13, pointerEvents:"none" }}>⌕</span>
              <input value={search} onChange={e=>setSearch(e.target.value)}
                placeholder={t.searchPlaceholder}
                style={{ ...inp, [dir==="rtl"?"paddingRight":"paddingLeft"]:34, borderRadius:10 }} />
            </div>
            <select value={fCat} onChange={e=>setFCat(e.target.value)} style={selStyle}>
              <option value="All">{t.allCats}</option>
              {currentCats.map(c=><option key={c}>{c}</option>)}
            </select>
            <button onClick={openAddItem}
              style={{ padding:"9px 22px", border:"none",
                background:`linear-gradient(135deg,${activeDept?.color||"#1d4ed8"},${activeDept?.color||"#1d4ed8"}bb)`,
                color:"#fff", borderRadius:9, cursor:"pointer", fontWeight:700, fontSize:13,
                fontFamily:L.font, whiteSpace:"nowrap",
                boxShadow:`0 3px 14px ${activeDept?.color||"#1d4ed8"}44` }}>
              {t.addDevice}
            </button>
          </div>

          {/* Table */}
          <div style={{ background:L.card, border:`1px solid ${L.border}`, borderRadius:16,
            overflow:"hidden", boxShadow:L.shadow }}>
            <div style={{ overflowX:"auto" }}>
              <table style={{ width:"100%", borderCollapse:"collapse", fontSize:12 }}>
                <thead>
                  <tr style={{ background:"#f8faff" }}>
                    {t.tableHeaders.map((h,i)=>(
                      <th key={i} style={{ padding:"11px 13px", textAlign:dir==="rtl"?"right":"left",
                        fontWeight:700, fontSize:10, textTransform:"uppercase", letterSpacing:.9,
                        color:L.muted, borderBottom:`1.5px solid ${L.border}`, whiteSpace:"nowrap",
                        fontFamily:L.mono }}>
                        {h}
                      </th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {rows.length===0 ? (
                    <tr><td colSpan={12} style={{ textAlign:"center", padding:60, color:L.muted, fontSize:13 }}>
                      <div style={{ fontSize:36, marginBottom:10 }}>📭</div>
                      <div style={{ fontWeight:600, color:L.sub }}>{t.noDevices}</div>
                      <div style={{ fontSize:11, marginTop:5 }}>{t.noDevicesHint(L.accent)}</div>
                    </td></tr>
                  ) : rows.map((it,i)=>(
                    <tr key={it.id} className="tr-row" style={{ borderBottom:`1px solid ${L.border}`, transition:"all .12s" }}>
                      <td style={{ padding:"10px 13px", color:L.muted, fontFamily:L.mono, fontSize:11 }}>{i+1}</td>
                      <td style={{ padding:"10px 13px" }}>
                        <span style={{ background:L.bg, color:L.sub, border:`1px solid ${L.borderHi}`,
                          borderRadius:6, padding:"2px 8px", fontSize:10, fontWeight:600, fontFamily:L.mono }}>
                          {CAT_ICONS[it.category]||""} {it.category}
                        </span>
                      </td>
                      <td style={{ padding:"10px 13px", fontWeight:600, color:L.text }}>
                        {it.name}
                        {isCameraCategory(it.category) && it.cameraBrand && (
                          <div style={{ fontSize:10, color:L.muted, marginTop:2 }}>
                            {it.cameraBrand}{it.cameraType ? ` · ${it.cameraType}` : ""}{it.cameraWireless ? " · 📶 Wireless" : ""}
                          </div>
                        )}
                        {isMobileSIM(it.category) && it.simNumber && (
                          <div style={{ fontSize:10, color:L.muted, marginTop:2 }}>{it.simNumber}{it.simCarrier ? ` · ${it.simCarrier}` : ""}</div>
                        )}
                      </td>
                      <td style={{ padding:"10px 13px", color:L.sub }}>{it.brand||"—"}</td>
                      <td style={{ padding:"10px 13px", color:L.sub, fontFamily:L.mono, fontSize:11 }}>{it.model||"—"}</td>
                      <td style={{ padding:"10px 13px", color:L.sub, fontFamily:L.mono, fontSize:11 }}>{it.serial||"—"}</td>
                      <td style={{ padding:"10px 13px" }}><StatusBadge status={it.status} lang={lang} /></td>
                      <td style={{ padding:"10px 13px", color:L.text }}>{it.user||"—"}</td>
                      <td style={{ padding:"10px 13px", color:L.sub }}>{it.location||"—"}</td>
                      <td style={{ padding:"10px 13px", color:L.sub, maxWidth:120, overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }}>{it.notes||"—"}</td>
                      <td style={{ padding:"10px 13px", color:L.muted, fontFamily:L.mono, fontSize:10 }}>{it.date}</td>
                      <td style={{ padding:"10px 13px" }}>
                        <div style={{ display:"flex", gap:5 }}>
                          <button onClick={()=>openEdit(it)}
                            style={{ background:"#eff6ff", color:"#1d4ed8", border:"1px solid #bfdbfe",
                              borderRadius:6, padding:"4px 11px", cursor:"pointer", fontSize:11, fontWeight:600 }}>{t.editBtn}</button>
                          <button onClick={()=>delItem(it.id)}
                            style={{ background:"#fff5f5", color:"#dc2626", border:"1px solid #fecaca",
                              borderRadius:6, padding:"4px 10px", cursor:"pointer", fontSize:11 }}>{t.delBtn}</button>
                        </div>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
            <div style={{ padding:"10px 14px", borderTop:`1px solid ${L.border}`,
              display:"flex", justifyContent:"space-between", alignItems:"center" }}>
              <span style={{ fontSize:11, color:L.muted, fontFamily:L.mono }}>{t.recordsShown(rows.length)}</span>
              <span onClick={()=>setView("dashboard")}
                style={{ fontSize:11, color:"#1d4ed8", cursor:"pointer", fontWeight:600 }}>{t.backDash}</span>
            </div>
          </div>
        </main>
      )}

      {/* ══════════════ SETTINGS ══════════════ */}
      {view === "settings" && (
        <main style={{ padding:"28px 24px", maxWidth:1100, margin:"0 auto" }}>
          <div style={{ fontWeight:800, fontSize:20, color:L.text, marginBottom:6 }}>{t.settingsTitle}</div>
          <div style={{ fontSize:12, color:L.muted, marginBottom:24 }}>Manage departments and categories without code changes.</div>

          {/* Settings sub-tabs */}
          <div style={{ display:"flex", gap:6, marginBottom:24 }}>
            {[["departments",t.deptManagement,"🏢"],["categories",t.categoryManagement,"📂"]].map(([tab,lb,ic])=>(
              <button key={tab} onClick={()=>setSettingsTab(tab)}
                style={{ padding:"9px 22px", border:`1.5px solid ${settingsTab===tab?"#1d4ed888":L.border}`,
                  borderRadius:10, cursor:"pointer", background:settingsTab===tab?"#eff6ff":L.surface,
                  color:settingsTab===tab?"#1d4ed8":L.sub, fontWeight:settingsTab===tab?700:500,
                  fontSize:13, fontFamily:L.font, display:"flex", alignItems:"center", gap:7 }}>
                {ic} {lb}
              </button>
            ))}
          </div>

          {/* ── DEPT MANAGEMENT ── */}
          {settingsTab === "departments" && (
            <div style={{ background:L.card, border:`1px solid ${L.border}`, borderRadius:16, overflow:"hidden", boxShadow:L.shadow }}>
              <div style={{ padding:"16px 20px", borderBottom:`1px solid ${L.border}`,
                display:"flex", justifyContent:"space-between", alignItems:"center" }}>
                <div style={{ fontWeight:700, fontSize:14, color:L.text }}>{t.deptManagement}</div>
                <button onClick={openAddDept}
                  style={{ padding:"7px 18px", border:"none",
                    background:"linear-gradient(135deg,#1d4ed8,#7c3aed)",
                    color:"#fff", borderRadius:9, cursor:"pointer", fontWeight:700, fontSize:12, fontFamily:L.font }}>
                  {t.newDept}
                </button>
              </div>
              <div>
                {depts.map((d,idx)=>(
                  <div key={d.id} className="settings-row" style={{
                    padding:"14px 20px", borderBottom:idx<depts.length-1?`1px solid ${L.border}`:"none",
                    display:"flex", alignItems:"center", gap:14, transition:"background .12s" }}>
                    <div style={{ width:40, height:40, borderRadius:10, fontSize:20,
                      background:d.color+"18", border:`1px solid ${d.color}30`,
                      display:"flex", alignItems:"center", justifyContent:"center", flexShrink:0 }}>{d.icon}</div>
                    <div style={{ flex:1 }}>
                      <div style={{ fontWeight:700, fontSize:13, color:L.text }}>{d.nameEn}</div>
                      {d.nameAr && d.nameAr !== d.nameEn && (
                        <div style={{ fontSize:11, color:L.muted, direction:"rtl", textAlign:"left" }}>{d.nameAr}</div>
                      )}
                      <div style={{ fontSize:10, color:L.muted, marginTop:3 }}>
                        {deptCats(d).slice(0,6).join(" · ")}{deptCats(d).length>6?` +${deptCats(d).length-6} more`:""}
                      </div>
                    </div>
                    <div style={{ display:"flex", alignItems:"center", gap:6 }}>
                      <span style={{ fontFamily:L.mono, fontSize:13, fontWeight:700, color:d.color }}>{deptCount(d.id)}</span>
                      <span style={{ fontSize:11, color:L.muted }}>devices</span>
                    </div>
                    <div style={{ display:"flex", gap:6 }}>
                      <button onClick={()=>openEditDept(d)}
                        style={{ padding:"6px 14px", background:"#eff6ff", color:"#1d4ed8",
                          border:"1px solid #bfdbfe", borderRadius:7, cursor:"pointer",
                          fontSize:11, fontWeight:600, fontFamily:L.font }}>
                        ✏️ {lang==="ar"?"تعديل":"Edit"}
                      </button>
                      <button onClick={()=>deleteDept(d.id)}
                        style={{ padding:"6px 12px", background:"#fff5f5", color:"#dc2626",
                          border:"1px solid #fecaca", borderRadius:7, cursor:"pointer", fontSize:12 }}>
                        ✕
                      </button>
                    </div>
                  </div>
                ))}
              </div>
            </div>
          )}

          {/* ── CATEGORY MANAGEMENT ── */}
          {settingsTab === "categories" && (
            <div>
              {/* Shared cats info */}
              <div style={{ background:"#eff6ff", border:"1px solid #bfdbfe", borderRadius:12,
                padding:"12px 18px", marginBottom:16, fontSize:12, color:"#1d4ed8" }}>
                <strong>Shared categories</strong> (appear in all departments): {SHARED_CATS.join(", ")}
              </div>

              {/* Add new cat to a dept */}
              <div style={{ background:L.card, border:`1px solid ${L.border}`, borderRadius:16,
                padding:"20px 24px", marginBottom:20, boxShadow:L.shadow }}>
                <div style={{ fontWeight:700, fontSize:14, color:L.text, marginBottom:14 }}>Add Category to Department</div>
                <div style={{ display:"flex", gap:10, flexWrap:"wrap", alignItems:"flex-end" }}>
                  <div style={{ flex:1, minWidth:180 }}>
                    <Field label="Category Name">
                      <input value={newCatName} onChange={e=>setNewCatName(e.target.value)}
                        placeholder="e.g. UPS, Tablet..."
                        style={inp} />
                    </Field>
                  </div>
                  <div style={{ flex:1, minWidth:180 }}>
                    <Field label="Department">
                      <select id="newCatDept" style={selStyle}>
                        <option value="__all__">All Departments</option>
                        {depts.map(d=><option key={d.id} value={d.id}>{d.icon} {d.nameEn}</option>)}
                      </select>
                    </Field>
                  </div>
                  <button onClick={()=>{
                    const catVal = newCatName.trim();
                    if (!catVal) return;
                    const deptSel = document.getElementById("newCatDept").value;
                    if (deptSel === "__all__") {
                      setDepts(p=>p.map(d=>d.cats.includes(catVal)?d:{...d,cats:[...d.cats,catVal]}));
                      showMsg(`"${catVal}" added to all departments ✓`);
                    } else {
                      setDepts(p=>p.map(d=>d.id===deptSel && !d.cats.includes(catVal)?{...d,cats:[...d.cats,catVal]}:d));
                      const dn = depts.find(d=>d.id===deptSel)?.nameEn||"";
                      showMsg(`"${catVal}" added to ${dn} ✓`);
                    }
                    setNewCatName("");
                  }}
                    style={{ padding:"9px 20px", border:"none",
                      background:"linear-gradient(135deg,#1d4ed8,#7c3aed)",
                      color:"#fff", borderRadius:9, cursor:"pointer", fontWeight:700,
                      fontSize:13, fontFamily:L.font, whiteSpace:"nowrap" }}>
                    {t.addCategory}
                  </button>
                </div>
              </div>

              {/* Per-dept category lists */}
              <div style={{ display:"flex", flexDirection:"column", gap:14 }}>
                {depts.map(d=>(
                  <div key={d.id} style={{ background:L.card, border:`1px solid ${L.border}`,
                    borderRadius:14, overflow:"hidden", boxShadow:L.shadow }}>
                    <div style={{ padding:"12px 18px", borderBottom:`1px solid ${L.border}`,
                      background:d.color+"0a", display:"flex", alignItems:"center", gap:10 }}>
                      <span style={{ fontSize:18 }}>{d.icon}</span>
                      <span style={{ fontWeight:700, fontSize:13, color:L.text }}>{d.nameEn}</span>
                      <span style={{ fontSize:11, color:L.muted, fontFamily:L.mono }}>{d.cats.length} categories</span>
                    </div>
                    <div style={{ padding:"12px 18px", display:"flex", flexWrap:"wrap", gap:8 }}>
                      {d.cats.map(cat=>(
                        <div key={cat} style={{ display:"flex", alignItems:"center", gap:5,
                          background:L.bg, border:`1px solid ${L.borderHi}`, borderRadius:20,
                          padding:"4px 12px 4px 10px", fontSize:12 }}>
                          <span>{CAT_ICONS[cat]||"📦"}</span>
                          <span style={{ color:L.text, fontWeight:500 }}>{cat}</span>
                          <button onClick={()=>{
                            setDepts(p=>p.map(x=>x.id===d.id?{...x,cats:x.cats.filter(c=>c!==cat)}:x));
                            showMsg(`"${cat}" removed from ${d.nameEn}`,"err");
                          }} style={{ background:"none", border:"none", color:L.muted, cursor:"pointer",
                            fontSize:13, lineHeight:1, padding:"0 0 0 3px" }}>×</button>
                        </div>
                      ))}
                      {SHARED_CATS.map(cat=>(
                        <div key={cat} style={{ display:"flex", alignItems:"center", gap:5,
                          background:"#eff6ff", border:`1px solid #bfdbfe`, borderRadius:20,
                          padding:"4px 12px 4px 10px", fontSize:12 }}>
                          <span>{CAT_ICONS[cat]||"📦"}</span>
                          <span style={{ color:"#1d4ed8", fontWeight:500 }}>{cat}</span>
                          <span style={{ fontSize:9, color:"#60a5fa", marginLeft:2 }}>shared</span>
                        </div>
                      ))}
                    </div>
                  </div>
                ))}
              </div>
            </div>
          )}
        </main>
      )}

      {/* ══ ITEM MODAL ══ */}
      {showForm && form && (
        <Modal onClose={()=>setShowForm(false)} dir={dir}>
          <div style={{ fontWeight:800, fontSize:17, marginBottom:20, color:L.text }}>
            {editId ? t.editDevice : t.registerDevice}
          </div>
          <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:14 }}>
            <Field label={t.department}>
              <select value={form.deptId} onChange={e=>{ const d=depts.find(x=>x.id===e.target.value); setForm(f=>({...f,deptId:e.target.value,category:deptCats(d)[0]||""})); }} style={inp}>
                {depts.map(d=><option key={d.id} value={d.id}>{d.icon} {deptName(d)}</option>)}
              </select>
            </Field>
            <Field label={t.category}>
              <select value={form.category} onChange={e=>setForm(f=>({...f,category:e.target.value,cameraBrand:"",cameraType:"",cameraWireless:false}))} style={inp}>
                {deptCats(depts.find(d=>d.id===form.deptId)).map(c=><option key={c}>{c}</option>)}
              </select>
            </Field>
            <Field label={t.deviceName}><input value={form.name} onChange={e=>setForm(f=>({...f,name:e.target.value}))} placeholder={t.deviceNamePlaceholder} style={inp} /></Field>

            {/* Camera-specific brand */}
            {isCameraCategory(form.category) ? (
              <Field label={t.cameraBrand}>
                <select value={form.cameraBrand||""} onChange={e=>{
                  const brand=e.target.value;
                  setForm(f=>({...f,cameraBrand:brand,cameraWireless:brand==="Ezviz",cameraType:brand==="Ezviz"?"":f.cameraType}));
                }} style={inp}>
                  <option value="">— Select brand —</option>
                  <option value="Hikvision">Hikvision</option>
                  <option value="Ezviz">Ezviz</option>
                </select>
              </Field>
            ) : (
              <Field label={t.brand}><input value={form.brand} onChange={e=>setForm(f=>({...f,brand:e.target.value}))} placeholder={t.brandPlaceholder} style={inp} /></Field>
            )}

            {/* Camera type — only for Hikvision */}
            {isCameraCategory(form.category) && form.cameraBrand === "Hikvision" && (
              <Field label={t.cameraType}>
                <select value={form.cameraType||""} onChange={e=>setForm(f=>({...f,cameraType:e.target.value}))} style={inp}>
                  <option value="">— Select type —</option>
                  <option value="CCTV">CCTV</option>
                  <option value="IP Camera">IP Camera</option>
                </select>
              </Field>
            )}

            {/* Ezviz wireless badge */}
            {isCameraCategory(form.category) && form.cameraBrand === "Ezviz" && (
              <div style={{ gridColumn:"1/-1", background:"#f0fdf4", border:"1px solid #6ee7b7",
                borderRadius:10, padding:"10px 14px", fontSize:12, color:"#059669", display:"flex", alignItems:"center", gap:8 }}>
                📶 <strong>Wireless</strong> — {t.wirelessNote}
              </div>
            )}

            {/* SIM-specific fields */}
            {isMobileSIM(form.category) && (
              <>
                <Field label={t.simNumber}><input value={form.simNumber||""} onChange={e=>setForm(f=>({...f,simNumber:e.target.value}))} placeholder="+20..." style={inp} /></Field>
                <Field label={t.simCarrier}><input value={form.simCarrier||""} onChange={e=>setForm(f=>({...f,simCarrier:e.target.value}))} placeholder="Vodafone / Orange / WE / Etisalat" style={inp} /></Field>
                <Field label={t.simPlan}><input value={form.simPlan||""} onChange={e=>setForm(f=>({...f,simPlan:e.target.value}))} placeholder="Monthly plan details..." style={inp} /></Field>
              </>
            )}

            {!isCameraCategory(form.category) && (
              <>
                <Field label={t.model}><input value={form.model} onChange={e=>setForm(f=>({...f,model:e.target.value}))} placeholder={t.modelPlaceholder} style={inp} /></Field>
                <Field label={t.serial}><input value={form.serial} onChange={e=>setForm(f=>({...f,serial:e.target.value}))} placeholder={t.serialPlaceholder} style={inp} /></Field>
              </>
            )}
            {isCameraCategory(form.category) && (
              <>
                <Field label={t.model}><input value={form.model} onChange={e=>setForm(f=>({...f,model:e.target.value}))} placeholder={t.modelPlaceholder} style={inp} /></Field>
                <Field label={t.serial}><input value={form.serial} onChange={e=>setForm(f=>({...f,serial:e.target.value}))} placeholder={t.serialPlaceholder} style={inp} /></Field>
              </>
            )}

            <Field label={t.status}>
              <select value={form.status} onChange={e=>setForm(f=>({...f,status:e.target.value}))} style={inp}>
                {Object.entries(STATUS_KEYS).map(([k,s])=><option key={k} value={k}>{lang==="ar"?s.ar:s.en}</option>)}
              </select>
            </Field>
            <Field label={t.assignedUser}><input value={form.user} onChange={e=>setForm(f=>({...f,user:e.target.value}))} placeholder={t.userPlaceholder} style={inp} /></Field>
            <Field label={t.location}><input value={form.location} onChange={e=>setForm(f=>({...f,location:e.target.value}))} placeholder={t.locationPlaceholder} style={inp} /></Field>
            <Field label={t.entryDate}><input type="date" value={form.date} onChange={e=>setForm(f=>({...f,date:e.target.value}))} style={inp} /></Field>
            <div style={{ gridColumn:"1/-1" }}>
              <Field label={t.notes}><textarea value={form.notes} onChange={e=>setForm(f=>({...f,notes:e.target.value}))} rows={2} placeholder={t.notesPlaceholder} style={{...inp,resize:"vertical"}} /></Field>
            </div>
          </div>
          <div style={{ display:"flex", gap:10, justifyContent:"flex-end", marginTop:20 }}>
            <button onClick={()=>setShowForm(false)}
              style={{ padding:"9px 22px", borderRadius:9, border:`1px solid ${L.borderHi}`,
                background:"transparent", color:L.sub, cursor:"pointer", fontWeight:600, fontFamily:L.font }}>{t.cancel}</button>
            <button onClick={saveItem}
              style={{ padding:"9px 28px", borderRadius:9, border:"none",
                background:"linear-gradient(135deg,#1d4ed8,#7c3aed)",
                color:"#fff", cursor:"pointer", fontWeight:700, fontSize:14, fontFamily:L.font,
                boxShadow:"0 4px 16px #1d4ed844" }}>
              {editId ? t.saveChanges : t.addDeviceBtn}
            </button>
          </div>
        </Modal>
      )}

      {/* ══ DEPT MODAL (Add / Edit) ══ */}
      {showDeptModal && (
        <Modal onClose={()=>{setShowDeptModal(false);setEditingDept(null);}} dir={dir}>
          <div style={{ fontWeight:800, fontSize:17, marginBottom:20, color:L.text }}>
            {editingDept ? t.editDept : t.addNewDept}
          </div>
          <div style={{ display:"flex", flexDirection:"column", gap:14 }}>
            <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:14 }}>
              <Field label={t.deptNameLabel + " (EN)"}>
                <input value={deptForm.nameEn} onChange={e=>setDeptForm(f=>({...f,nameEn:e.target.value}))} placeholder="e.g. Legal Department" style={inp} />
              </Field>
              <Field label={t.deptNameLabel + " (AR)"}>
                <input value={deptForm.nameAr} onChange={e=>setDeptForm(f=>({...f,nameAr:e.target.value}))} placeholder="مثال: القسم القانوني" style={{...inp,direction:"rtl"}} />
              </Field>
            </div>
            <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:14 }}>
              <Field label={t.iconLabel}><input value={deptForm.icon} onChange={e=>setDeptForm(f=>({...f,icon:e.target.value}))} placeholder="🖥️" style={inp} /></Field>
              <Field label={t.colorLabel}>
                <div style={{ display:"flex", gap:8 }}>
                  <input type="color" value={deptForm.color} onChange={e=>setDeptForm(f=>({...f,color:e.target.value}))}
                    style={{ width:44,height:38,border:`1px solid ${L.borderHi}`,borderRadius:8,cursor:"pointer",background:L.bg,padding:3,flexShrink:0 }} />
                  <input value={deptForm.color} onChange={e=>setDeptForm(f=>({...f,color:e.target.value}))} style={{...inp,flex:1}} />
                </div>
              </Field>
            </div>
            <Field label={t.catsLabel}>
              <input value={deptForm.cats} onChange={e=>setDeptForm(f=>({...f,cats:e.target.value}))} placeholder={t.catPlaceholder} style={inp} />
              <div style={{ fontSize:10, color:L.muted, marginTop:4 }}>{t.catsHint}</div>
            </Field>
          </div>
          {deptForm.nameEn && (
            <div style={{ marginTop:14, padding:"11px 15px", borderRadius:10,
              border:`1px solid ${deptForm.color}55`, background:deptForm.color+"12",
              display:"flex", alignItems:"center", gap:10 }}>
              <span style={{ fontSize:22 }}>{deptForm.icon||"🖥️"}</span>
              <div>
                <div style={{ fontWeight:700, color:deptForm.color, fontSize:13 }}>{deptForm.nameEn}</div>
                {deptForm.nameAr && <div style={{ fontSize:11, color:L.sub, direction:"rtl" }}>{deptForm.nameAr}</div>}
                <div style={{ fontSize:10, color:L.muted }}>{t.preview}</div>
              </div>
            </div>
          )}
          <div style={{ display:"flex", gap:10, justifyContent:"flex-end", marginTop:20 }}>
            <button onClick={()=>{setShowDeptModal(false);setEditingDept(null);}}
              style={{ padding:"9px 22px", borderRadius:9, border:`1px solid ${L.borderHi}`,
                background:"transparent", color:L.sub, cursor:"pointer", fontWeight:600, fontFamily:L.font }}>{t.cancel}</button>
            <button onClick={saveDept}
              style={{ padding:"9px 28px", borderRadius:9, border:"none",
                background:"linear-gradient(135deg,#1d4ed8,#7c3aed)",
                color:"#fff", cursor:"pointer", fontWeight:700, fontSize:14, fontFamily:L.font }}>
              {editingDept ? t.saveDeptBtn : t.createDept}
            </button>
          </div>
        </Modal>
      )}

      {/* ══ TOAST ══ */}
      {toast && (
        <div style={{ position:"fixed", bottom:24, left:"50%", transform:"translateX(-50%)",
          background:toast.type==="err"?"#fff5f5":"#f0fdf4",
          color:toast.type==="err"?"#dc2626":"#059669",
          border:`1.5px solid ${toast.type==="err"?"#fca5a5":"#6ee7b7"}`,
          borderRadius:12, padding:"12px 24px", fontWeight:700, fontSize:13,
          boxShadow:"0 8px 32px #00000018", zIndex:9999,
          whiteSpace:"nowrap", fontFamily:L.mono }}>
          {toast.msg}
        </div>
      )}
    </div>
  );
}
