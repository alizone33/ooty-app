import { useState, useCallback, useRef } from "react";
import * as XLSX from "xlsx";

// ── Auth Config ────────────────────────────────────────────────────────────────
const USERS = {
  "ootyveg@admin": { password: "ootyveg@3333", role: "admin" },
  "ootyveg@store": { password: "ootyveg@3333", role: "store" },
};

// ── Helpers ────────────────────────────────────────────────────────────────────
const todayISO = () => {
  const d = new Date();
  return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}-${String(d.getDate()).padStart(2,"0")}`;
};
const fmtDate = (iso) => iso ? iso.split("-").reverse().join("-") : "";

// ── Palette ────────────────────────────────────────────────────────────────────
const C = {
  bg:       "#f0f4f2",
  surface:  "#ffffff",
  card:     "#ffffff",
  border:   "#c5d8cd",
  accent:   "#1a6641",
  accentDk: "#1a3a2a",
  accentLt: "#2d6a4f",
  text:     "#1a3a2a",
  textDim:  "#5a7a68",
  textMut:  "#8aaa96",
  white:    "#ffffff",
  danger:   "#c0392b",
  warn:     "#c0610a",
  extra:    "#6b3fa0",
  extraLt:  "#6b3fa0",
  row0:     "#f8fbf9",
  row1:     "#edf2ef",
};

// ── Shared Styles ──────────────────────────────────────────────────────────────
const S = {
  root: {
    minHeight:"100vh", background:C.bg, display:"flex",
    alignItems:"flex-start", justifyContent:"center",
    padding:"28px 12px", fontFamily:"'Courier New',Courier,monospace",
  },
  card: {
    background:C.card, borderRadius:12, border:`1px solid ${C.border}`,
    padding:"28px 24px", width:"100%", maxWidth:560,
    boxShadow:"0 8px 40px rgba(0,0,0,.5)",
  },
  logoRow: { display:"flex", alignItems:"center", gap:10, marginBottom:6 },
  brand: { fontWeight:800, fontSize:15, color:C.accent, letterSpacing:2, textTransform:"uppercase" },
  h1: { fontSize:18, fontWeight:800, color:C.text, margin:"0 0 20px", letterSpacing:1 },
  label: { display:"block", fontSize:11, fontWeight:700, color:C.textDim, letterSpacing:1.2, marginBottom:6, textTransform:"uppercase" },
  input: {
    width:"100%", padding:"10px 14px", borderRadius:8,
    border:`1.5px solid ${C.border}`, background:C.surface,
    fontSize:14, outline:"none", color:C.text, boxSizing:"border-box",
    fontFamily:"'Courier New',Courier,monospace",
  },
  primaryBtn: {
    padding:"11px 22px", borderRadius:8, background:C.accentDk,
    color:"#fff", fontSize:14, fontWeight:800, cursor:"pointer", border:"none",
    letterSpacing:.5, fontFamily:"'Courier New',Courier,monospace",
  },
  ghostBtn: {
    padding:"9px 16px", borderRadius:8, border:`1.5px solid ${C.border}`,
    background:"transparent", color:C.textDim, fontSize:13, fontWeight:600, cursor:"pointer",
    fontFamily:"'Courier New',Courier,monospace",
  },
  tableHeader: {
    display:"grid", background:C.accentDk, padding:"9px 14px", fontSize:11,
    fontWeight:800, color:"#fff", letterSpacing:1, gap:8,
  },
  tableRow: {
    display:"grid", padding:"10px 14px", gap:8, alignItems:"center",
    borderBottom:`1px solid ${C.border}`,
  },
  badge: { borderRadius:5, padding:"2px 8px", fontSize:11, fontWeight:700 },
};

// ── Logo ───────────────────────────────────────────────────────────────────────
function Logo() {
  return (
    <div style={S.logoRow}>
      <span style={{fontSize:22}}>🥬</span>
      <span style={S.brand}>Ooty Veg &amp; Fruits</span>
    </div>
  );
}

// ── Nav Bar ────────────────────────────────────────────────────────────────────
function NavBar({ role, page, onNav, onLogout }) {
  const adminTabs = [
    { id:"upload", label:"📤 Import" },
    { id:"reports", label:"📊 Reports" },
  ];
  const storeTabs = [
    { id:"entry", label:"📝 Entry" },
    { id:"reports", label:"📊 Reports" },
  ];
  const tabs = role === "admin" ? adminTabs : storeTabs;

  return (
    <div style={{
      position:"sticky", top:0, zIndex:100,
      background:C.surface, borderBottom:`1px solid ${C.border}`,
      padding:"10px 20px", display:"flex", alignItems:"center",
      justifyContent:"space-between", fontFamily:"'Courier New',Courier,monospace",
    }}>
      <div style={{display:"flex", alignItems:"center", gap:6}}>
        <span style={{fontSize:18}}>🥬</span>
        <span style={{fontWeight:800, fontSize:13, color:C.accent, letterSpacing:2}}>OVF</span>
      </div>
      <div style={{display:"flex", gap:4}}>
        {tabs.map(t => (
          <button key={t.id} onClick={()=>onNav(t.id)} style={{
            padding:"6px 14px", borderRadius:7, border:"none", cursor:"pointer",
            background: page===t.id ? C.accentDk : "transparent",
            color: page===t.id ? "#fff" : C.textDim,
            fontSize:12, fontWeight:700, letterSpacing:.5,
            fontFamily:"'Courier New',Courier,monospace",
          }}>{t.label}</button>
        ))}
      </div>
      <div style={{display:"flex", alignItems:"center", gap:10}}>
        <span style={{fontSize:11, color:C.textMut, letterSpacing:.5}}>
          {role==="admin"?"👤 Admin":"🏪 Store"}
        </span>
        <button onClick={onLogout} style={{
          padding:"5px 12px", borderRadius:6, border:`1px solid ${C.border}`,
          background:"transparent", color:C.textDim, fontSize:11, cursor:"pointer",
          fontFamily:"'Courier New',Courier,monospace",
        }}>Logout</button>
      </div>
    </div>
  );
}

// ══════════════════════════════════════════════════════════════════════════════
// PAGE 1 — LOGIN
// ══════════════════════════════════════════════════════════════════════════════
function LoginPage({ onLogin }) {
  const [username, setUsername] = useState("");
  const [password, setPassword] = useState("");
  const [error, setError] = useState("");
  const [showPw, setShowPw] = useState(false);

  const handleLogin = () => {
    const user = USERS[username.trim()];
    if (user && user.password === password) {
      onLogin(username.trim(), user.role);
    } else {
      setError("Invalid username or password.");
    }
  };

  return (
    <div style={S.root}>
      <div style={{...S.card, maxWidth:420, marginTop:60}}>
        <div style={{textAlign:"center", marginBottom:28}}>
          <div style={{fontSize:48, marginBottom:8}}>🥬</div>
          <div style={{fontWeight:900, fontSize:20, color:C.accent, letterSpacing:3, textTransform:"uppercase"}}>
            Ooty Veg &amp; Fruits
          </div>
          <div style={{fontSize:12, color:C.textDim, marginTop:4, letterSpacing:1.5}}>
            INDENT &amp; RECEIVING SYSTEM
          </div>
        </div>

        <div style={{marginBottom:14}}>
          <label style={S.label}>Username</label>
          <input
            style={S.input}
            type="text"
            placeholder="ootyveg@admin"
            value={username}
            onChange={e=>{setUsername(e.target.value);setError("");}}
            onKeyDown={e=>e.key==="Enter"&&handleLogin()}
          />
        </div>
        <div style={{marginBottom:20}}>
          <label style={S.label}>Password</label>
          <div style={{position:"relative"}}>
            <input
              style={{...S.input, paddingRight:44}}
              type={showPw?"text":"password"}
              placeholder="••••••••"
              value={password}
              onChange={e=>{setPassword(e.target.value);setError("");}}
              onKeyDown={e=>e.key==="Enter"&&handleLogin()}
            />
            <button
              onClick={()=>setShowPw(p=>!p)}
              style={{
                position:"absolute", right:10, top:"50%", transform:"translateY(-50%)",
                background:"none", border:"none", cursor:"pointer", color:C.textDim, fontSize:16,
              }}
            >{showPw?"🙈":"👁"}</button>
          </div>
        </div>

        {error && (
          <div style={{background:"#fdf0f0", border:`1px solid ${C.danger}`, borderRadius:8, padding:"10px 14px", color:C.danger, fontSize:13, marginBottom:14}}>
            ⚠ {error}
          </div>
        )}

        <button onClick={handleLogin} style={{...S.primaryBtn, width:"100%", padding:"13px"}}>
          Login →
        </button>

        <div style={{marginTop:20, padding:"12px 14px", background:C.surface, borderRadius:8, border:`1px solid ${C.border}`, fontSize:11, color:C.textMut, lineHeight:1.8}}>
          <div>Admin: <span style={{color:C.accentLt}}>ootyveg@admin</span></div>
          <div>Store: <span style={{color:C.accentLt}}>ootyveg@store</span></div>
        </div>
      </div>
    </div>
  );
}

// ══════════════════════════════════════════════════════════════════════════════
// PAGE 2 — ADMIN UPLOAD
// ══════════════════════════════════════════════════════════════════════════════
function AdminUploadPage({ indentData, onUpload }) {
  const [dragging, setDragging] = useState(false);
  const [error, setError] = useState("");
  const [uploading, setUploading] = useState(false);
  const fileRef = useRef();

  const parseFile = useCallback((file) => {
    setError(""); setUploading(true);
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const wb = XLSX.read(e.target.result, { type:"array" });
        // Try INDENT sheet first, then first sheet
        const sheetName = wb.SheetNames.includes("INDENT") ? "INDENT" : wb.SheetNames[0];
        const ws = wb.Sheets[sheetName];
        const rows = XLSX.utils.sheet_to_json(ws, { header:1, defval:0 });
        if (rows.length < 2) throw new Error("File appears empty.");
        const header = rows[0].map(h => String(h||"").trim());
        const catIdx  = header.findIndex(h => /categ/i.test(h));
        const codeIdx = header.findIndex(h => /code/i.test(h));
        const nameIdx = header.findIndex(h => /name/i.test(h));
        if (catIdx===-1||codeIdx===-1||nameIdx===-1)
          throw new Error("Could not find Category / Code / Name columns.");
        const fixedCols = new Set([catIdx,codeIdx,nameIdx]);
        const storeCols = header.map((h,i)=>({h,i}))
          .filter(({h,i})=>!fixedCols.has(i)&&!/total/i.test(h)&&h!=="");
        const storeNames = storeCols.map(s=>s.h);
        const parsed = rows.slice(1)
          .filter(r => r[nameIdx] && String(r[nameIdx]).trim())
          .map(r => {
            const stores = {};
            storeCols.forEach(({h,i}) => { stores[h] = Number(r[i])||0; });
            return {
              category: String(r[catIdx]||"").trim(),
              code: String(r[codeIdx]||"").trim(),
              name: String(r[nameIdx]||"").trim(),
              stores,
            };
          });
        onUpload({ products: parsed, stores: storeNames, uploadedAt: todayISO(), fileName: file.name });
        setUploading(false);
      } catch(err) { setError(err.message||"Failed to parse file."); setUploading(false); }
    };
    reader.readAsArrayBuffer(file);
  }, [onUpload]);

  const handleDrop = useCallback((e) => {
    e.preventDefault(); setDragging(false);
    if (e.dataTransfer.files[0]) parseFile(e.dataTransfer.files[0]);
  }, [parseFile]);

  const cats = indentData ? [...new Set(indentData.products.map(p=>p.category))] : [];

  return (
    <div style={{...S.root, display:"block"}}>
      <div style={{maxWidth:620, margin:"28px auto", padding:"0 12px"}}>
        <div style={S.card}>
          <Logo/>
          <h1 style={S.h1}>Import Indent File</h1>

          <div
            onDragOver={e=>{e.preventDefault();setDragging(true);}}
            onDragLeave={()=>setDragging(false)}
            onDrop={handleDrop}
            onClick={()=>fileRef.current.click()}
            style={{
              border:`2px dashed ${dragging?C.accent:C.border}`,
              borderRadius:12, padding:"40px 20px", textAlign:"center",
              cursor:"pointer", transition:"all .2s", marginBottom:16,
              background: dragging ? "rgba(26,102,65,.07)" : C.surface,
            }}
          >
            <div style={{fontSize:40, marginBottom:10}}>{uploading?"⏳":"📄"}</div>
            <div style={{fontWeight:700, color:C.text, marginBottom:4, letterSpacing:.5}}>
              {uploading?"Parsing file…":"Drop INDENT file here"}
            </div>
            <div style={{fontSize:12, color:C.textDim}}>or click to browse · .xlsx / .xls</div>
            <input ref={fileRef} type="file" accept=".xlsx,.xls" style={{display:"none"}}
              onChange={e=>e.target.files[0]&&parseFile(e.target.files[0])}/>
          </div>

          {error && (
            <div style={{background:"#fdf0f0",border:`1px solid ${C.danger}`,borderRadius:8,padding:"10px 14px",color:C.danger,fontSize:13,marginBottom:14}}>
              ⚠ {error}
            </div>
          )}

          <div style={{padding:"11px 14px",background:C.surface,borderRadius:8,border:`1px solid ${C.border}`,fontSize:12,color:C.textMut,lineHeight:1.7}}>
            <strong style={{color:C.accentLt}}>Expected:</strong> INDENT sheet with columns — Product Category, Product Code, Product Name, then one column per store, optional Total at end.
          </div>
        </div>

        {indentData && (
          <div style={{...S.card, marginTop:16}}>
            <div style={{display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:16}}>
              <div>
                <div style={{fontWeight:800, color:C.accent, fontSize:15, letterSpacing:.5}}>✓ File Loaded</div>
                <div style={{fontSize:12, color:C.textDim, marginTop:2}}>{indentData.fileName} · {fmtDate(indentData.uploadedAt)}</div>
              </div>
              <button onClick={()=>fileRef.current&&fileRef.current.click()} style={{...S.ghostBtn, fontSize:12}}>
                Replace ↻
              </button>
            </div>
            <div style={{display:"grid", gridTemplateColumns:"repeat(3,1fr)", gap:10, marginBottom:16}}>
              {[
                {label:"Products", val: indentData.products.length, icon:"📦"},
                {label:"Stores", val: indentData.stores.length, icon:"🏪"},
                {label:"Categories", val: cats.length, icon:"🗂"},
              ].map(({label,val,icon})=>(
                <div key={label} style={{background:C.surface, borderRadius:8, border:`1px solid ${C.border}`, padding:"14px 10px", textAlign:"center"}}>
                  <div style={{fontSize:22}}>{icon}</div>
                  <div style={{fontSize:22, fontWeight:900, color:C.accent, marginTop:4}}>{val}</div>
                  <div style={{fontSize:11, color:C.textDim, letterSpacing:.5}}>{label}</div>
                </div>
              ))}
            </div>
            <div style={{fontSize:12, color:C.textDim, marginBottom:8, letterSpacing:.5, fontWeight:700}}>CATEGORIES</div>
            <div style={{display:"flex", flexWrap:"wrap", gap:6}}>
              {cats.map(c=>(
                <span key={c} style={{background:C.surface, border:`1px solid ${C.border}`, borderRadius:20, padding:"4px 12px", fontSize:12, color:C.accentLt}}>{c}</span>
              ))}
            </div>
          </div>
        )}
      </div>
    </div>
  );
}

// ══════════════════════════════════════════════════════════════════════════════
// ENTRY FLOW — Store User (Steps: Store Select → Entry → Review → Done)
// ══════════════════════════════════════════════════════════════════════════════

// ── Add Extra Item Modal ───────────────────────────────────────────────────────
function AddExtraModal({ products, onAdd, onClose }) {
  const [code, setCode] = useState("");
  const [found, setFound] = useState(null);
  const [notFound, setNotFound] = useState(false);
  const [receivedQty, setReceivedQty] = useState("");
  const [manualName, setManualName] = useState("");
  const [manualCat, setManualCat] = useState("");

  const searchCode = () => {
    const trimmed = code.trim().toUpperCase();
    const match = products.find(p => p.code.toUpperCase() === trimmed);
    if (match) { setFound(match); setNotFound(false); }
    else { setFound(null); setNotFound(true); }
  };

  const handleAdd = () => {
    if (found) {
      onAdd({ category:found.category, code:found.code, name:found.name, indentQty:0, receivedQty:Number(receivedQty)||0, isExtra:true });
    } else {
      onAdd({ category:manualCat.trim()||"EXTRA", code:code.trim().toUpperCase(), name:manualName.trim()||code.trim(), indentQty:0, receivedQty:Number(receivedQty)||0, isExtra:true });
    }
    onClose();
  };

  const canAdd = (found || (notFound && manualName.trim())) && receivedQty !== "";

  return (
    <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,.65)",display:"flex",alignItems:"center",justifyContent:"center",zIndex:1000,padding:16}}
      onClick={e=>e.target===e.currentTarget&&onClose()}>
      <div style={{background:C.card,borderRadius:14,padding:24,width:"100%",maxWidth:400,border:`1px solid ${C.border}`}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:18}}>
          <h3 style={{margin:0,fontSize:15,fontWeight:800,color:C.text}}>Add Extra Item</h3>
          <button onClick={onClose} style={{background:"none",border:"none",fontSize:22,cursor:"pointer",color:C.textDim}}>×</button>
        </div>

        <label style={S.label}>Product Code</label>
        <div style={{display:"flex",gap:8,marginBottom:14}}>
          <input type="text" placeholder="e.g. VEG001" value={code}
            onChange={e=>{setCode(e.target.value);setFound(null);setNotFound(false);}}
            onKeyDown={e=>e.key==="Enter"&&searchCode()}
            style={{...S.input,margin:0,flex:1}}/>
          <button onClick={searchCode} style={{...S.primaryBtn,padding:"9px 16px",fontSize:13}}>Search</button>
        </div>

        {found && (
          <div style={{background:C.surface,border:`1.5px solid ${C.accentDk}`,borderRadius:8,padding:"10px 12px",marginBottom:14}}>
            <div style={{fontSize:13,color:C.accentLt,fontWeight:700}}>{found.name}</div>
            <div style={{fontSize:11,color:C.textDim,marginTop:2}}>{found.code} · {found.category}</div>
          </div>
        )}

        {notFound && (
          <div style={{background:C.surface,border:`1.5px solid ${C.warn}`,borderRadius:8,padding:"10px 12px",marginBottom:14}}>
            <div style={{fontSize:12,color:C.warn,fontWeight:700,marginBottom:8}}>⚠ Code not found — enter manually</div>
            <label style={{...S.label,marginBottom:4}}>Product Name *</label>
            <input type="text" placeholder="Enter product name" value={manualName}
              onChange={e=>setManualName(e.target.value)} style={{...S.input,marginBottom:10,padding:"8px 10px"}}/>
            <label style={{...S.label,marginBottom:4}}>Category</label>
            <input type="text" placeholder="e.g. VEGETABLES" value={manualCat}
              onChange={e=>setManualCat(e.target.value)} style={{...S.input,padding:"8px 10px"}}/>
          </div>
        )}

        {(found || notFound) && (
          <>
            <label style={S.label}>Received Qty</label>
            <input type="number" min="0" step="0.5" placeholder="0" value={receivedQty}
              onChange={e=>setReceivedQty(e.target.value)} style={{...S.input,marginBottom:18}}/>
          </>
        )}

        <div style={{display:"flex",gap:10}}>
          <button onClick={onClose} style={{...S.ghostBtn,flex:1}}>Cancel</button>
          <button onClick={handleAdd} disabled={!canAdd}
            style={{...S.primaryBtn,flex:1,opacity:canAdd?1:.4,cursor:canAdd?"pointer":"not-allowed"}}>
            Add Item
          </button>
        </div>
      </div>
    </div>
  );
}

// ── Entry Page ─────────────────────────────────────────────────────────────────
function EntryPage({ indentData, savedEntries, onSaveEntry }) {
  const [entryStep, setEntryStep] = useState(1); // 1=select store+date, 2=entry, 3=review
  const [selectedDate, setSelectedDate] = useState(todayISO());
  const [storeSearch, setStoreSearch] = useState("");
  const [selectedStore, setSelectedStore] = useState(null);

  const [received, setReceived] = useState({});
  const [activeTab, setActiveTab] = useState(null);
  const [showExtraModal, setShowExtraModal] = useState(false);
  const [extraItems, setExtraItems] = useState([]);
  const [reviewData, setReviewData] = useState(null);
  const [submitted, setSubmitted] = useState(false);

  const EXTRA_TAB = "__EXTRA__";

  const products = indentData ? indentData.products : [];
  const stores = indentData ? indentData.stores : [];

  // Check if there's existing entry for selected date+store
  const existingEntry = selectedDate && selectedStore
    ? savedEntries.find(e=>e.date===selectedDate && e.store===selectedStore)
    : null;

  const pickStore = (store) => {
    setSelectedStore(store);
    const existing = savedEntries.find(e=>e.date===selectedDate && e.store===store);
    if (existing) {
      // Load existing data
      const init = {};
      existing.items.forEach(item => { if (!item.isExtra) init[item.code] = item.receivedQty; });
      setReceived(init);
      setExtraItems(existing.items.filter(i=>i.isExtra).map(i=>({...i})));
    } else {
      const init = {};
      products.forEach(p => { init[p.code] = ""; });
      setReceived(init);
      setExtraItems([]);
    }
    const cats = [...new Set(products.map(p=>p.category))];
    setActiveTab(cats[0]||null);
    setEntryStep(2);
    setSubmitted(false);
  };

  const categories = [...new Set(products.map(p=>p.category))];

  const storeProducts = (cat) =>
    products.filter(p => p.category===cat && p.stores[selectedStore]>0);

  const totalIndentCat = (cat) => storeProducts(cat).reduce((s,p)=>s+p.stores[selectedStore],0);
  const totalRecvCat = (cat) => storeProducts(cat).reduce((s,p)=>s+(Number(received[p.code])||0),0);
  const enteredCount = (cat) => storeProducts(cat).filter(p=>received[p.code]!=="").length;
  const totalCount = (cat) => storeProducts(cat).length;

  const allWithIndent = products.filter(p=>p.stores[selectedStore]>0);
  const totalEntered = allWithIndent.filter(p=>received[p.code]!=="").length;
  const totalItems = allWithIndent.length;

  const addExtraItem = (item) => {
    setExtraItems(prev => {
      const idx = prev.findIndex(e=>e.code===item.code);
      if (idx>=0) { const u=[...prev]; u[idx]=item; return u; }
      return [...prev, item];
    });
  };
  const removeExtraItem = (code) => setExtraItems(prev=>prev.filter(e=>e.code!==code));
  const updateExtraRecv = (code,val) => setExtraItems(prev=>prev.map(e=>e.code===code?{...e,receivedQty:val}:e));

  const submitAll = () => {
    const cats = categories.map(cat => {
      const prods = storeProducts(cat);
      const indent = prods.reduce((s,p)=>s+p.stores[selectedStore],0);
      const recv = prods.reduce((s,p)=>s+(Number(received[p.code])||0),0);
      const items = prods.map(p => ({
        code:p.code, name:p.name, category:p.category,
        indentQty:p.stores[selectedStore],
        receivedQty:Number(received[p.code])||0,
        diff:(Number(received[p.code])||0)-p.stores[selectedStore],
        isExtra:false,
      }));
      return {cat,indent,recv,items};
    });
    const grandIndent = cats.reduce((s,c)=>s+c.indent,0);
    const grandRecv = cats.reduce((s,c)=>s+c.recv,0);
    const extraRecv = extraItems.reduce((s,e)=>s+(Number(e.receivedQty)||0),0);
    setReviewData({ cats, grandIndent, grandRecv, extraItems, grandRecvWithExtra:grandRecv+extraRecv });
    setEntryStep(3);
  };

  const saveAndFinish = () => {
    // Flatten all items for storage
    const allItems = [];
    reviewData.cats.forEach(({items})=>allItems.push(...items));
    reviewData.extraItems.forEach(e=>allItems.push({...e, isExtra:true}));
    onSaveEntry({ date:selectedDate, store:selectedStore, items:allItems, savedAt:new Date().toISOString() });
    setSubmitted(true);
  };

  const exportExcel = () => {
    const wb = XLSX.utils.book_new();
    const dateLabel = fmtDate(selectedDate);
    const rows = [["DATE","STORE NAME","PRODUCT CATEGORY","PRODUCT CODE","PRODUCT NAME","INDENT QTY","RECEIVED QTY","OVERALL"]];
    reviewData.cats.forEach(({cat,items})=>{
      items.forEach(item=>{
        rows.push([dateLabel, selectedStore, cat, item.code, item.name, item.indentQty, item.receivedQty, item.receivedQty-item.indentQty]);
      });
    });
    reviewData.extraItems.forEach(e=>{
      rows.push([dateLabel, selectedStore, e.category, e.code, e.name, 0, Number(e.receivedQty)||0, Number(e.receivedQty)||0]);
    });
    const ws = XLSX.utils.aoa_to_sheet(rows);
    ws["!cols"] = [12,10,16,12,26,12,14,10].map(w=>({wch:w}));
    XLSX.utils.book_append_sheet(wb, ws, "Receiving Data");
    XLSX.writeFile(wb, `receiving_${selectedStore}_${dateLabel}.xlsx`);
  };

  // ── No indent loaded ───────────────────────────────────────────────────────
  if (!indentData) {
    return (
      <div style={{...S.root, display:"block"}}>
        <div style={{maxWidth:500,margin:"40px auto",padding:"0 12px"}}>
          <div style={{...S.card,textAlign:"center",padding:"40px 24px"}}>
            <div style={{fontSize:40,marginBottom:12}}>⚠️</div>
            <div style={{fontWeight:700,color:C.warn,fontSize:15,marginBottom:8}}>No Indent File Loaded</div>
            <div style={{fontSize:13,color:C.textDim}}>Please ask your admin to upload today's INDENT file.</div>
          </div>
        </div>
      </div>
    );
  }

  // ── Step 1: Select date + store ───────────────────────────────────────────
  if (entryStep===1) {
    const filtered = stores.filter(s=>s.toLowerCase().includes(storeSearch.toLowerCase()));
    return (
      <div style={{...S.root,display:"block"}}>
        <div style={{maxWidth:560,margin:"28px auto",padding:"0 12px"}}>
          <div style={S.card}>
            <Logo/>
            <h1 style={S.h1}>Select Date &amp; Store</h1>
            <div style={{marginBottom:18}}>
              <label style={S.label}>Entry Date</label>
              <input type="date" value={selectedDate} onChange={e=>setSelectedDate(e.target.value)}
                style={{...S.input,width:"auto"}}/>
            </div>
            <label style={S.label}>Search Store</label>
            <input type="text" placeholder="Search…" value={storeSearch}
              onChange={e=>setStoreSearch(e.target.value)} style={{...S.input,marginBottom:14}}/>
            <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:8,marginBottom:8}}>
              {filtered.map(s=>{
                const hasEntry = savedEntries.find(e=>e.date===selectedDate && e.store===s);
                return (
                  <button key={s} onClick={()=>pickStore(s)} style={{
                    display:"flex",flexDirection:"column",alignItems:"center",padding:"12px 8px",
                    borderRadius:10,border:`1.5px solid ${hasEntry?C.accentDk:C.border}`,
                    background:hasEntry?"rgba(26,102,65,.09)":C.surface,cursor:"pointer",gap:3,
                  }}>
                    <span style={{fontSize:18}}>{hasEntry?"✅":"🏪"}</span>
                    <span style={{fontWeight:700,fontSize:11,color:C.text,textAlign:"center"}}>{s}</span>
                    {hasEntry && <span style={{fontSize:10,color:C.accentLt,fontWeight:600}}>Saved</span>}
                  </button>
                );
              })}
            </div>
            {selectedDate !== indentData.uploadedAt && (
              <div style={{padding:"9px 14px",background:C.surface,border:`1px solid ${C.warn}`,borderRadius:8,fontSize:12,color:C.warn,marginTop:8}}>
                ⚠ Indent file is from {fmtDate(indentData.uploadedAt)} · You're entering for {fmtDate(selectedDate)}
              </div>
            )}
          </div>
        </div>
      </div>
    );
  }

  // ── Step 2: Entry ─────────────────────────────────────────────────────────
  if (entryStep===2) {
    const isExtraTab = activeTab===EXTRA_TAB;
    const tabProds = isExtraTab ? [] : storeProducts(activeTab);

    return (
      <div style={{...S.root,display:"block"}}>
        {showExtraModal && <AddExtraModal products={products} onAdd={addExtraItem} onClose={()=>setShowExtraModal(false)}/>}
        <div style={{maxWidth:680,margin:"28px auto",padding:"0 12px"}}>
          <div style={S.card}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:4}}>
              <Logo/>
              <div style={{background:C.accentDk,color:"#fff",borderRadius:20,padding:"4px 12px",fontSize:12,fontWeight:700}}>
                {totalEntered}/{totalItems} items
              </div>
            </div>
            <div style={{fontSize:12,color:C.textDim,marginBottom:16}}>
              <strong style={{color:C.text}}>{selectedStore}</strong> · {fmtDate(selectedDate)}
              {existingEntry && <span style={{marginLeft:8,color:C.accentLt,fontSize:11}}>✓ Editing saved entry</span>}
            </div>

            {/* Category tabs */}
            <div style={{display:"flex",gap:6,marginBottom:16,flexWrap:"wrap",alignItems:"center"}}>
              {categories.map(cat=>{
                const entered=enteredCount(cat),total=totalCount(cat);
                const done=total>0&&entered===total,active=activeTab===cat;
                return (
                  <button key={cat} onClick={()=>setActiveTab(cat)} style={{
                    padding:"6px 14px",borderRadius:20,cursor:"pointer",fontSize:12,fontWeight:700,
                    background:active?C.accentDk:done?"rgba(26,102,65,.12)":C.surface,
                    color:active?"#fff":done?C.accentLt:C.textDim,
                    border:active?"none":done?`1.5px solid ${C.accentDk}`:`1.5px solid ${C.border}`,
                  }}>
                    {cat}
                    <span style={{
                      marginLeft:5,background:active?"rgba(255,255,255,.2)":C.border,
                      color:active?"#fff":C.accentLt,borderRadius:10,padding:"1px 7px",fontSize:10,
                    }}>{entered}/{total}</span>
                  </button>
                );
              })}
              <button onClick={()=>setActiveTab(EXTRA_TAB)} style={{
                padding:"6px 14px",borderRadius:20,cursor:"pointer",fontSize:12,fontWeight:700,
                background:activeTab===EXTRA_TAB?C.extra:extraItems.length>0?"rgba(107,63,160,.12)":C.surface,
                color:activeTab===EXTRA_TAB?"#fff":C.extraLt,
                border:activeTab===EXTRA_TAB?"none":`1.5px solid ${C.extra}44`,
              }}>
                ✦ Extra {extraItems.length>0&&<span style={{marginLeft:5,background:C.extra+"44",color:C.extraLt,borderRadius:10,padding:"1px 7px",fontSize:10}}>{extraItems.length}</span>}
              </button>
            </div>

            {!isExtraTab && (
              <>
                <div style={{borderRadius:10,overflow:"hidden",border:`1px solid ${C.border}`,marginBottom:14}}>
                  <div style={{...S.tableHeader,gridTemplateColumns:"1fr 90px 110px"}}>
                    <span>PRODUCT</span>
                    <span style={{textAlign:"right"}}>INDENT</span>
                    <span style={{textAlign:"right"}}>RECEIVED</span>
                  </div>
                  {tabProds.length===0 ? (
                    <div style={{padding:"20px",textAlign:"center",color:C.textDim,fontSize:13}}>No products indented for this store.</div>
                  ) : tabProds.map((p,i)=>{
                    const indent=p.stores[selectedStore];
                    const recv=Number(received[p.code])||0;
                    const entered=received[p.code]!=="";
                    const short=entered&&recv<indent,full=entered&&recv>=indent;
                    return (
                      <div key={p.code} style={{...S.tableRow,gridTemplateColumns:"1fr 90px 110px",background:i%2===0?C.row0:C.row1}}>
                        <div>
                          <div style={{fontWeight:600,fontSize:13,color:C.text}}>{p.name}</div>
                          <div style={{fontSize:11,color:C.textMut,marginTop:1}}>{p.code}</div>
                        </div>
                        <div style={{textAlign:"right",fontWeight:800,color:C.accent,fontSize:15}}>{indent}</div>
                        <div style={{display:"flex",justifyContent:"flex-end"}}>
                          <input
                            type="number" min="0" step="0.5" placeholder="0"
                            value={received[p.code]??""}
                            onChange={e=>setReceived(prev=>({...prev,[p.code]:e.target.value}))}
                            style={{
                              width:88,padding:"6px 10px",borderRadius:7,textAlign:"right",
                              border:full?`1.5px solid ${C.accentDk}`:short?`1.5px solid ${C.warn}`:`1.5px solid ${C.border}`,
                              background:full?"rgba(26,102,65,.10)":short?"rgba(192,97,10,.10)":C.surface,
                              fontSize:14,fontWeight:800,outline:"none",color:C.text,
                              fontFamily:"'Courier New',Courier,monospace",
                            }}
                          />
                        </div>
                      </div>
                    );
                  })}
                </div>
                <div style={{background:C.surface,borderRadius:8,padding:"9px 14px",display:"flex",justifyContent:"space-between",marginBottom:16,border:`1px solid ${C.border}`}}>
                  <div style={{fontSize:12,color:C.textDim}}>
                    <span style={{color:C.text,fontWeight:700}}>{activeTab}</span>
                    {" · "}Indent: <strong>{totalIndentCat(activeTab)}</strong>
                    {" · "}Received: <strong>{totalRecvCat(activeTab)}</strong>
                  </div>
                  <div style={{fontSize:12,fontWeight:700,color:totalRecvCat(activeTab)<totalIndentCat(activeTab)?C.warn:C.accentLt}}>
                    {totalRecvCat(activeTab)<totalIndentCat(activeTab)
                      ?`−${totalIndentCat(activeTab)-totalRecvCat(activeTab)} short`
                      :totalRecvCat(activeTab)>totalIndentCat(activeTab)
                      ?`+${totalRecvCat(activeTab)-totalIndentCat(activeTab)} extra`:"✓ Matched"}
                  </div>
                </div>
              </>
            )}

            {isExtraTab && (
              <div style={{marginBottom:16}}>
                <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12}}>
                  <div style={{fontSize:13,color:C.textDim}}>Items not in today's indent</div>
                  <button onClick={()=>setShowExtraModal(true)} style={{
                    padding:"7px 14px",borderRadius:8,background:C.extra,color:"#fff",
                    fontSize:12,fontWeight:700,cursor:"pointer",border:"none",
                  }}>+ Add Item</button>
                </div>
                <div style={{borderRadius:10,overflow:"hidden",border:`1px solid ${C.extra}44`}}>
                  <div style={{...S.tableHeader,gridTemplateColumns:"1fr 90px 110px 36px",background:C.extra}}>
                    <span>PRODUCT</span>
                    <span style={{textAlign:"right"}}>CATEGORY</span>
                    <span style={{textAlign:"right"}}>RECEIVED</span>
                    <span/>
                  </div>
                  {extraItems.length===0 ? (
                    <div style={{padding:"24px",textAlign:"center",color:C.extraLt,fontSize:13}}>
                      No extra items yet. Click <strong>+ Add Item</strong>.
                    </div>
                  ) : extraItems.map((e,i)=>(
                    <div key={e.code} style={{...S.tableRow,gridTemplateColumns:"1fr 90px 110px 36px",background:i%2===0?C.row0:C.row1}}>
                      <div>
                        <div style={{fontWeight:600,fontSize:13,color:C.text}}>{e.name}</div>
                        <div style={{fontSize:11,color:C.textMut}}>{e.code}</div>
                      </div>
                      <div style={{textAlign:"right",fontSize:12,color:C.extraLt,fontWeight:600}}>{e.category}</div>
                      <div style={{display:"flex",justifyContent:"flex-end"}}>
                        <input type="number" min="0" step="0.5" value={e.receivedQty}
                          onChange={ev=>updateExtraRecv(e.code,ev.target.value)}
                          style={{width:88,padding:"6px 10px",borderRadius:7,textAlign:"right",border:`1.5px solid ${C.extra}66`,background:"rgba(107,63,160,.08)",fontSize:14,fontWeight:800,outline:"none",color:C.text,fontFamily:"'Courier New',Courier,monospace"}}/>
                      </div>
                      <button onClick={()=>removeExtraItem(e.code)}
                        style={{background:"none",border:"none",cursor:"pointer",color:C.danger,fontSize:18,lineHeight:1,padding:0}}>×</button>
                    </div>
                  ))}
                </div>
              </div>
            )}

            <div style={{display:"flex",gap:10}}>
              <button onClick={()=>setEntryStep(1)} style={S.ghostBtn}>← Back</button>
              <button onClick={submitAll} style={{...S.primaryBtn,flex:1}}>Review &amp; Submit →</button>
            </div>
          </div>
        </div>
      </div>
    );
  }

  // ── Step 3: Review ────────────────────────────────────────────────────────
  if (entryStep===3 && reviewData) {
    const diffLabel = (diff)=>{
      if(diff===0) return {text:"✓ Full",color:C.accentLt,bg:"rgba(26,102,65,.10)"};
      if(diff>0) return {text:`+${diff} Extra`,color:C.accentLt,bg:"rgba(26,102,65,.10)"};
      return {text:`${diff} Short`,color:C.warn,bg:"rgba(192,97,10,.10)"};
    };
    return (
      <div style={{...S.root,display:"block"}}>
        <div style={{maxWidth:720,margin:"28px auto",padding:"0 12px"}}>
          <div style={S.card}>
            <Logo/>
            <h1 style={S.h1}>Review &amp; Submit</h1>
            <div style={{fontSize:12,color:C.textDim,marginBottom:16}}>
              <strong style={{color:C.text}}>{selectedStore}</strong> · {fmtDate(selectedDate)}
            </div>
            <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:10,marginBottom:20}}>
              {[
                {label:"Total Indent",val:reviewData.grandIndent,color:C.accent},
                {label:"Total Received",val:reviewData.grandRecvWithExtra,color:C.accentLt},
                {label:"Extra Items",val:reviewData.extraItems.length,color:C.extraLt},
              ].map(({label,val,color})=>(
                <div key={label} style={{background:C.surface,borderRadius:10,padding:"14px 10px",textAlign:"center",border:`1px solid ${C.border}`}}>
                  <div style={{fontSize:24,fontWeight:900,color}}>{val}</div>
                  <div style={{fontSize:11,color:C.textDim,marginTop:4,letterSpacing:.5}}>{label}</div>
                </div>
              ))}
            </div>

            {reviewData.cats.map(({cat,indent,recv,items})=>(
              <div key={cat} style={{marginBottom:14,borderRadius:10,overflow:"hidden",border:`1px solid ${C.border}`}}>
                <div style={{background:C.accentDk,padding:"8px 14px",display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                  <span style={{color:"#fff",fontWeight:800,fontSize:12,letterSpacing:.8}}>{cat}</span>
                  <span style={{color:"rgba(255,255,255,.75)",fontSize:11}}>Indent {indent} · Recv {recv}</span>
                </div>
                {items.map((item,i)=>{
                  const lbl=diffLabel(item.diff);
                  return (
                    <div key={item.code} style={{...S.tableRow,gridTemplateColumns:"1fr 70px 70px 90px",background:i%2===0?C.row0:C.row1}}>
                      <div>
                        <div style={{fontWeight:600,fontSize:13,color:C.text}}>{item.name}</div>
                        <div style={{fontSize:11,color:C.textMut}}>{item.code}</div>
                      </div>
                      <div style={{textAlign:"right",color:C.accent,fontWeight:700}}>{item.indentQty}</div>
                      <div style={{textAlign:"right",color:C.accentLt,fontWeight:700}}>{item.receivedQty}</div>
                      <div style={{textAlign:"right"}}>
                        <span style={{...S.badge,background:lbl.bg,color:lbl.color}}>{lbl.text}</span>
                      </div>
                    </div>
                  );
                })}
              </div>
            ))}

            {reviewData.extraItems.length>0 && (
              <div style={{marginBottom:14,borderRadius:10,overflow:"hidden",border:`1px solid ${C.extra}44`}}>
                <div style={{background:C.extra,padding:"8px 14px",display:"flex",justifyContent:"space-between"}}>
                  <span style={{color:"#fff",fontWeight:800,fontSize:12}}>✦ Extra Items</span>
                  <span style={{color:"rgba(255,255,255,.75)",fontSize:11}}>{reviewData.extraItems.length} item(s)</span>
                </div>
                {reviewData.extraItems.map((e,i)=>(
                  <div key={e.code} style={{...S.tableRow,gridTemplateColumns:"1fr 80px 80px 90px",background:i%2===0?C.row0:C.row1}}>
                    <div>
                      <div style={{fontWeight:600,fontSize:13,color:C.text}}>{e.name}</div>
                      <div style={{fontSize:11,color:C.textMut}}>{e.code} · {e.category}</div>
                    </div>
                    <div style={{textAlign:"right",color:C.textDim,fontStyle:"italic",fontSize:12}}>—</div>
                    <div style={{textAlign:"right",color:C.extraLt,fontWeight:700}}>{Number(e.receivedQty)||0}</div>
                    <div style={{textAlign:"right"}}>
                      <span style={{...S.badge,background:"rgba(107,63,160,.12)",color:C.extraLt}}>Extra</span>
                    </div>
                  </div>
                ))}
              </div>
            )}

            {submitted ? (
              <div style={{background:"rgba(26,102,65,.08)",border:`1px solid ${C.accentDk}`,borderRadius:10,padding:"18px",textAlign:"center"}}>
                <div style={{fontSize:28,marginBottom:6}}>✅</div>
                <div style={{fontWeight:800,color:C.accent,fontSize:14,marginBottom:4}}>Entry Saved!</div>
                <div style={{fontSize:12,color:C.textDim,marginBottom:14}}>Data saved for {selectedStore} · {fmtDate(selectedDate)}</div>
                <div style={{display:"flex",gap:10,justifyContent:"center",flexWrap:"wrap"}}>
                  <button onClick={exportExcel} style={{...S.primaryBtn,background:C.accentDk,fontSize:13}}>⬇ Export Excel</button>
                  <button onClick={()=>{setEntryStep(1);setSubmitted(false);}} style={S.ghostBtn}>New Entry</button>
                </div>
              </div>
            ) : (
              <div style={{display:"flex",gap:10,marginTop:4}}>
                <button onClick={()=>setEntryStep(2)} style={S.ghostBtn}>← Edit</button>
                <button onClick={saveAndFinish} style={{...S.primaryBtn,flex:1}}>Save Entry ✓</button>
              </div>
            )}
          </div>
        </div>
      </div>
    );
  }

  return null;
}

// ══════════════════════════════════════════════════════════════════════════════
// PAGE 5 — REPORTS
// ══════════════════════════════════════════════════════════════════════════════
function ReportsPage({ savedEntries, indentData }) {
  const [dateFrom, setDateFrom] = useState(todayISO());
  const [dateTo, setDateTo] = useState(todayISO());
  const [storeFilter, setStoreFilter] = useState("ALL");

  const allStores = indentData ? indentData.stores : [];
  const uniqueStoresInEntries = [...new Set(savedEntries.map(e=>e.store))].sort();

  const filtered = savedEntries.filter(e => {
    const inRange = e.date >= dateFrom && e.date <= dateTo;
    const inStore = storeFilter==="ALL" || e.store===storeFilter;
    return inRange && inStore;
  });

  // Flatten to rows
  const rows = [];
  filtered.forEach(entry => {
    entry.items.forEach(item => {
      rows.push({
        date: fmtDate(entry.date),
        store: entry.store,
        category: item.category,
        code: item.code,
        name: item.name,
        indentQty: item.indentQty,
        receivedQty: item.receivedQty,
        overall: item.receivedQty - item.indentQty,
      });
    });
  });

  const totalIndent = rows.reduce((s,r)=>s+r.indentQty,0);
  const totalRecv = rows.reduce((s,r)=>s+r.receivedQty,0);

  const exportReport = () => {
    const wb = XLSX.utils.book_new();
    const header = ["DATE","STORE NAME","PRODUCT CATEGORY","PRODUCT CODE","PRODUCT NAME","INDENT QTY","RECEIVED QTY","OVERALL"];
    const data = [header, ...rows.map(r=>[r.date,r.store,r.category,r.code,r.name,r.indentQty,r.receivedQty,r.overall])];
    const ws = XLSX.utils.aoa_to_sheet(data);
    ws["!cols"] = [12,10,16,12,28,12,14,10].map(w=>({wch:w}));
    XLSX.utils.book_append_sheet(wb, ws, "Report");
    const suffix = storeFilter==="ALL"?"all-stores":storeFilter;
    XLSX.writeFile(wb, `report_${suffix}_${dateFrom}_to_${dateTo}.xlsx`);
  };

  return (
    <div style={{...S.root,display:"block"}}>
      <div style={{maxWidth:900,margin:"28px auto",padding:"0 12px"}}>
        <div style={S.card}>
          <Logo/>
          <h1 style={S.h1}>Reports</h1>

          {/* Filters */}
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr auto",gap:12,marginBottom:20,alignItems:"flex-end"}}>
            <div>
              <label style={S.label}>From Date</label>
              <input type="date" value={dateFrom} onChange={e=>setDateFrom(e.target.value)} style={{...S.input,width:"auto",minWidth:140}}/>
            </div>
            <div>
              <label style={S.label}>To Date</label>
              <input type="date" value={dateTo} onChange={e=>setDateTo(e.target.value)} style={{...S.input,width:"auto",minWidth:140}}/>
            </div>
            <div>
              <label style={S.label}>Store</label>
              <select value={storeFilter} onChange={e=>setStoreFilter(e.target.value)}
                style={{...S.input,width:"auto",minWidth:120,cursor:"pointer"}}>
                <option value="ALL">All Stores</option>
                {(uniqueStoresInEntries.length > 0 ? uniqueStoresInEntries : allStores).map(s=>(
                  <option key={s} value={s}>{s}</option>
                ))}
              </select>
            </div>
            <button onClick={exportReport} disabled={rows.length===0}
              style={{...S.primaryBtn,opacity:rows.length>0?1:.4,cursor:rows.length>0?"pointer":"not-allowed",whiteSpace:"nowrap"}}>
              ⬇ Export
            </button>
          </div>

          {/* Summary cards */}
          <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:10,marginBottom:20}}>
            {[
              {label:"Entries",val:filtered.length,icon:"📋"},
              {label:"Rows",val:rows.length,icon:"📦"},
              {label:"Total Indent",val:totalIndent,icon:"📊"},
              {label:"Total Received",val:totalRecv,icon:"✅"},
            ].map(({label,val,icon})=>(
              <div key={label} style={{background:C.surface,borderRadius:8,border:`1px solid ${C.border}`,padding:"12px 10px",textAlign:"center"}}>
                <div style={{fontSize:18}}>{icon}</div>
                <div style={{fontSize:20,fontWeight:900,color:C.accent,marginTop:4}}>{val}</div>
                <div style={{fontSize:10,color:C.textDim,letterSpacing:.5}}>{label}</div>
              </div>
            ))}
          </div>

          {/* Table */}
          {rows.length===0 ? (
            <div style={{textAlign:"center",padding:"40px 20px",color:C.textDim,border:`1px dashed ${C.border}`,borderRadius:10}}>
              <div style={{fontSize:30,marginBottom:8}}>📭</div>
              <div style={{fontSize:13}}>No entries found for selected filters.</div>
            </div>
          ) : (
            <div style={{borderRadius:10,overflow:"hidden",border:`1px solid ${C.border}`,overflowX:"auto"}}>
              <div style={{...S.tableHeader,gridTemplateColumns:"100px 80px 120px 90px 1fr 80px 90px 80px",minWidth:800}}>
                <span>DATE</span>
                <span>STORE</span>
                <span>CATEGORY</span>
                <span>CODE</span>
                <span>PRODUCT NAME</span>
                <span style={{textAlign:"right"}}>INDENT</span>
                <span style={{textAlign:"right"}}>RECEIVED</span>
                <span style={{textAlign:"right"}}>OVERALL</span>
              </div>
              <div style={{maxHeight:420,overflowY:"auto"}}>
                {rows.map((r,i)=>(
                  <div key={i} style={{...S.tableRow,gridTemplateColumns:"100px 80px 120px 90px 1fr 80px 90px 80px",background:i%2===0?C.row0:C.row1,minWidth:800}}>
                    <span style={{fontSize:12,color:C.textDim}}>{r.date}</span>
                    <span style={{fontSize:12,color:C.accentLt,fontWeight:700}}>{r.store}</span>
                    <span style={{fontSize:11,color:C.textDim}}>{r.category}</span>
                    <span style={{fontSize:11,color:C.textMut}}>{r.code}</span>
                    <span style={{fontSize:13,color:C.text,fontWeight:600}}>{r.name}</span>
                    <span style={{textAlign:"right",color:C.accent,fontWeight:700}}>{r.indentQty}</span>
                    <span style={{textAlign:"right",color:C.accentLt,fontWeight:700}}>{r.receivedQty}</span>
                    <span style={{textAlign:"right",fontWeight:700,color:r.overall<0?C.warn:r.overall>0?C.accent:C.textDim}}>{r.overall>0?`+${r.overall}`:r.overall}</span>
                  </div>
                ))}
              </div>
            </div>
          )}
        </div>
      </div>
    </div>
  );
}

// ══════════════════════════════════════════════════════════════════════════════
// ROOT APP
// ══════════════════════════════════════════════════════════════════════════════
export default function App() {
  const [auth, setAuth] = useState(null); // {username, role}
  const [page, setPage] = useState(null);
  const [indentData, setIndentData] = useState(null); // shared across admin & store
  const [savedEntries, setSavedEntries] = useState([]); // all submitted entries

  const handleLogin = (username, role) => {
    setAuth({username, role});
    setPage(role==="admin" ? "upload" : "entry");
  };

  const handleLogout = () => {
    setAuth(null);
    setPage(null);
  };

  const handleUpload = (data) => {
    setIndentData(data);
  };

  const handleSaveEntry = (entry) => {
    setSavedEntries(prev => {
      // Replace if same date+store
      const filtered = prev.filter(e=>!(e.date===entry.date && e.store===entry.store));
      return [...filtered, entry];
    });
  };

  if (!auth) return <LoginPage onLogin={handleLogin}/>;

  return (
    <div style={{background:C.bg,minHeight:"100vh"}}>
      <NavBar role={auth.role} page={page} onNav={setPage} onLogout={handleLogout}/>
      {auth.role==="admin" && page==="upload" && (
        <AdminUploadPage indentData={indentData} onUpload={handleUpload}/>
      )}
      {auth.role==="store" && page==="entry" && (
        <EntryPage indentData={indentData} savedEntries={savedEntries} onSaveEntry={handleSaveEntry}/>
      )}
      {page==="reports" && (
        <ReportsPage savedEntries={savedEntries} indentData={indentData}/>
      )}
    </div>
  );
}
