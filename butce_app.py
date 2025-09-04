import streamlit as st
import pandas as pd
import numpy as np
import re, io, os, time, datetime as dt, unicodedata
from urllib.parse import unquote

# ================== AYAR ==================
DEFAULT_EXCEL_PATH = "BÜTÇE ÇALIŞMAA.xlsx"

st.set_page_config(page_title="Bütçe Uygulaması", page_icon="💰")
st.title("Bütçe Uygulaması 💰")

# ================== YARDIMCI ==================
def speak(text: str):
    st.components.v1.html(
        "<script>try{const u=new SpeechSynthesisUtterance("
        + repr(str(text)) +
        ");u.lang='tr-TR';speechSynthesis.cancel();speechSynthesis.speak(u);}catch(e){}</script>", height=0
    )

def get_query_param(name: str):
    try:
        qp = st.query_params
        val = qp.get(name)
        if isinstance(val, list):
            return val[0] if val else None
        return val
    except Exception:
        return None

@st.cache_data(show_spinner=False)
def read_excel_path(path: str, mtime: float) -> pd.DataFrame:
    return pd.read_excel(path)

def _strip_accents(s: str) -> str:
    return "".join(c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn")

def _canon(s: str) -> str:
    s = (s or "").strip().lower()
    s = _strip_accents(s)
    s = re.sub(r"[^a-z0-9çğıöşü]+", "", s)
    return s

def tl(x):
    try: return f"{x:,.2f} TL".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception: return f"{x} TL"

def get_numeric(val, default=0.0):
    try:
        if pd.isna(val): return float(default)
        return float(val)
    except Exception: return float(default)

# ==== TR sayı kelimeleri ====
TR1={"sıfır":0,"sifir":0,"bir":1,"iki":2,"üç":3,"uc":3,"dört":4,"dort":4,"beş":5,"bes":5,"altı":6,"alti":6,"yedi":7,"sekiz":8,"dokuz":9}
TR10={"on":10,"yirmi":20,"otuz":30,"kırk":40,"kirk":40,"elli":50,"altmış":60,"altmis":60,"yetmiş":70,"yetmis":70,"seksen":80,"doksan":90}
TRM={"yüz":100,"yuz":100,"bin":1000}
def parse_tr_words(words):
    total=0; cur=0; used=False
    for w in words:
        w=w.lower()
        if w in TR1: cur+=TR1[w]; used=True
        elif w in TR10: cur+=TR10[w]; used=True
        elif w in TRM:
            mul=TRM[w]
            if mul==100: cur=(cur or 1)*100
            else: cur=(cur or 1)*mul; total+=cur; cur=0
            used=True
        else:
            if used: break
    total+=cur
    return total if used and total>0 else None

def splitw(txt):
    return [re.sub(r"[^a-zçğıöşü0-9]", "", w.lower()) for w in txt.split()]

# ==== PersonRef çıkarımı (tutarla karışmaz) ====
def extract_personref(txt):
    txt = txt or ""
    m = re.search(r"(?:person|ref|sicil|kişi|kisi)\D*([0-9][0-9\s]{3,})", txt, re.I)
    if m:
        d = re.sub(r"\D", "", m.group(1))
        if d.isdigit() and len(d) >= 4:
            return int(d), d
    m2 = re.search(r"\b(\d[ \d]{3,})\b", txt)
    if m2:
        d = re.sub(r"\D", "", m2.group(1))
        if d.isdigit() and len(d) >= 4:
            return int(d), d
    return None, None

def extract_amount(txt, pref_digits):
    m=re.search(r"(\d[\d\.\,]*)\s*(tl|lira)?\b", txt, re.I)
    if m:
        raw=m.group(1).replace(".","").replace(",",".")
        try:
            val=float(raw); return val if val>0 else None
        except: pass
    ws=splitw(txt)
    if "tl" in ws or "lira" in ws:
        idxs=[i for i,w in enumerate(ws) if w in ("tl","lira")]
        for idx in reversed(idxs):
            val=parse_tr_words(ws[max(0,idx-6):idx])
            if val: return float(val)
    toks=re.findall(r"\d[\d\.\,]*", txt)
    if pref_digits: toks=[t for t in toks if re.sub(r"\D","",t)!=pref_digits]
    if toks:
        raw=toks[-1].replace(".","").replace(",",".")
        try: val=float(raw); return val if val>0 else None
        except: pass
    val=parse_tr_words(ws[::-1])
    return float(val) if val else None

# ==== İsimden kişi bulma ====
def build_fullname_columns(df: pd.DataFrame) -> pd.DataFrame:
    out=df.copy()
    full = None
    for c in out.columns:
        if _canon(c) in {"adsoyad","adsoyadi","ad soyad","ad soyadi"}:
            full = out[c].astype(str).fillna("").str.strip(); break
    if full is None:
        ad_col=None; soyad_col=None
        for c in out.columns:
            if _canon(c) in {"ad","adi","isim"}: ad_col=c
            if _canon(c) in {"soyad","soyadi"}: soyad_col=c
        if ad_col and soyad_col:
            full = (out[ad_col].astype(str).fillna("") + " " + out[soyad_col].astype(str).fillna("")).str.strip()
    out["FULLNAME"] = full if full is not None else ""
    out["FULLNAME_NORM"]=out["FULLNAME"].astype(str).map(_canon)
    return out

@st.cache_data(show_spinner=False)
def normalize_all(df_in: pd.DataFrame) -> pd.DataFrame:
    df = df_in.copy()
    c2orig = {_canon(c): c for c in df.columns}
    def need(name, alts, default):
        if name in df.columns: return
        src = None
        if _canon(name) in c2orig: src = c2orig[_canon(name)]
        else:
            for a in alts:
                if _canon(a) in c2orig: src = c2orig[_canon(a)]; break
        if src: df.rename(columns={src: name}, inplace=True)
        else: df[name] = default
    need("PersonRef", ["sicil","sicil no","person","employee id","id","ref","personref"], pd.NA)
    need("CurrentSalary", ["mevcut maaş","mevcut ucret","salary","maas"], 0.0)
    need("NewSalary", ["yeni maaş","yeni ucret","new salary"], 0.0)
    need("BÜTÇE DIŞI TALEPLER İLE", ["butce disi","budget extra","ekstra"], 0.0)
    need("DEPARTMAN", ["departman","bölüm","bolum","department","birim"], "")

    for y in ["1.YÖNETİCİSİ","2.YÖNETİCİSİ","3.YÖNETİCİSİ","4.YÖNETİCİSİ"]:
        if y not in df.columns: df[y] = ""
        df[y] = df[y].fillna("").astype(str)
    for c in ["PersonRef","CurrentSalary","NewSalary","BÜTÇE DIŞI TALEPLER İLE"]:
        df[c] = pd.to_numeric(df[c], errors="coerce")
    df = build_fullname_columns(df)

    used = df["CurrentSalary"].fillna(0)*1.4
    df["KULLANILAN BÜTÇE DIŞI DAHİL"] = used
    df["SİSTEM KALAN"] = used - df["NewSalary"].fillna(0)
    df["BÜTÇE DIŞI KALAN"] = used - df["BÜTÇE DIŞI TALEPLER İLE"].fillna(0)
    return df

def find_personref_by_name(df: pd.DataFrame, text: str):
    norm_t=_canon(text)
    best_len=0; best_ref=None; best_name=None
    ser_ref=pd.to_numeric(df["PersonRef"], errors="coerce")
    for i,row in df.iterrows():
        fn=str(row.get("FULLNAME","") or "")
        fnn=str(row.get("FULLNAME_NORM","") or "")
        if not fnn: continue
        if fnn in norm_t and pd.notna(ser_ref.iat[i]):
            L=len(fnn)
            if L>best_len:
                best_len=L; best_ref=int(float(ser_ref.iat[i])); best_name=fn
    return (best_ref, best_name) if best_ref is not None else (None, None)

def parse_op_from_text(text: str, fallback_ui_op: str | None = None) -> str | None:
    t = (text or "").lower()
    act = None
    if re.search(r"\b(düş|dus|düşür|çıkar|cikar|eksilt|azalt)\b", t): act = "düş"
    elif re.search(r"\b(ekle|arttır|artır|yükselt|yukselt)\b", t): act = "ekle"
    is_dis = ("bütçe dış" in t) or ("butce dis" in t) or bool(re.search(r"\bbütçe\b.*\bdış", t)) or bool(re.search(r"\bbutce\b.*\bdis", t))
    pool = "dis" if is_dis else "sistem"
    if act:
        if pool == "sistem":
            return "Bütçeden Düş (Sistem Kalan)" if act == "düş" else "Bütçeye Ekle (Sistem Kalan)"
        else:
            return "Bütçeden Düş (Bütçe Dışı Kalan)" if act == "düş" else "Bütçeye Ekle (Bütçe Dışı Kalan)"
    return fallback_ui_op

def manager_chain(row):
    mans = [str(row.get(k,"")).strip() for k in ["1.YÖNETİCİSİ","2.YÖNETİCİSİ","3.YÖNETİCİSİ","4.YÖNETİCİSİ"]]
    mans = [m for m in mans if m]
    return " > ".join(mans) if mans else ""

def pool_from_op(op: str):
    return "Bütçe Dışı" if ("Bütçe Dışı" in (op or "")) else "Sistem"

# ================== STATE ==================
defaults = {
    "_last_voice": "",
    "history": [],
    "unsaved_ops": [],
    "pending_batch": None,
    "selected_ref": None,
    "force_listen": True,   # Durdur demedikçe açık
    "listening": True,
    "last_final_text": "",
    "sticky_amount": None,
    "sticky_amount_ts": 0.0,
    "auto_apply": True,
}
for k,v in defaults.items():
    if k not in st.session_state: st.session_state[k]=v

def set_sticky_amount(val: float):
    st.session_state.sticky_amount = float(val)
    st.session_state.sticky_amount_ts = time.time()

def get_sticky_amount():
    # 30sn: sesle söyleyip butona geç basarsan kaybolmasın
    if st.session_state.sticky_amount and (time.time()-st.session_state.sticky_amount_ts)<=30.0:
        return float(st.session_state.sticky_amount)
    return None

st.session_state.listening = bool(st.session_state.get("force_listen", True))

# ================== SİDEBAR - AYARLAR ==================
with st.sidebar:
    st.header("⚙️ Ayarlar")
    st.session_state.auto_apply = st.toggle("🎤 Sesle otomatik uygula", value=st.session_state.get("auto_apply", True))

# ================== VERİ YÜKLEME ==================
with st.sidebar:
    st.header("📄 Veri Kaynağı")
    use_default = st.toggle("Varsayılan dosya (BÜTÇE ÇALIŞMAA.xlsx)", value=True)

try:
    file_mtime = os.path.getmtime(DEFAULT_EXCEL_PATH) if use_default and os.path.exists(DEFAULT_EXCEL_PATH) else 0.0
    base_df = read_excel_path(DEFAULT_EXCEL_PATH, file_mtime) if use_default else st.stop()
except FileNotFoundError:
    st.error(f"'{DEFAULT_EXCEL_PATH}' bulunamadı."); st.stop()
except Exception as e:
    st.error(f"Excel okunamadı: {e}"); st.stop()

# --- ÖNEMLİ: Excel'i HER SEFERİNDE ezme! ---
# İlk çalıştırmada Excel'den yükle; sonrasında hep session_state.df'yi koru.
if "df" not in st.session_state or st.session_state.df is None:
    st.session_state.df = normalize_all(base_df)
else:
    # sadece türetilen kolonları tazele
    st.session_state.df = normalize_all(st.session_state.df)

df = st.session_state.df  # bundan sonra hep bunu kullan

# ================== FİLTRE ==================
with st.sidebar:
    st.header("🎛️ Filtreler & İşlemler")
    managers = pd.concat([df["1.YÖNETİCİSİ"],df["2.YÖNETİCİSİ"],df["3.YÖNETİCİSİ"],df["4.YÖNETİCİSİ"]], ignore_index=True)
    opts = sorted([m for m in managers.dropna().unique() if str(m).strip()!=""])
    selected_manager = st.selectbox("Bütçe işlemi yapılacak yönetici", opts if opts else ["(yok)"])

if opts and selected_manager!="(yok)":
    msk = (df["1.YÖNETİCİSİ"]==selected_manager)|(df["2.YÖNETİCİSİ"]==selected_manager)|(df["3.YÖNETİCİSİ"]==selected_manager)|(df["4.YÖNETİCİSİ"]==selected_manager)
    df_filtered = df[msk].copy()
else:
    df_filtered = df.copy()

# ================== KPI ==================
kullanilan=(df_filtered["CurrentSalary"].fillna(0).sum())*1.4
sistem_kalan=(df_filtered["SİSTEM KALAN"].fillna(0).sum())
butce_disi_kalan=(df_filtered["BÜTÇE DIŞI KALAN"].fillna(0).sum())
c1,c2,c3=st.columns(3)
c1.metric("KULLANILAN BÜTÇE DIŞI DAHİL", tl(kullanilan))
c2.metric("SİSTEM KALAN", tl(sistem_kalan))
c3.metric("BÜTÇE DIŞI KALAN", tl(butce_disi_kalan))

# ================== TABLO ==================
cols = ["PersonRef","FULLNAME","DEPARTMAN","1.YÖNETİCİSİ","2.YÖNETİCİSİ","3.YÖNETİCİSİ","4.YÖNETİCİSİ",
        "CurrentSalary","NewSalary","BÜTÇE DIŞI TALEPLER İLE","KULLANILAN BÜTÇE DIŞI DAHİL","SİSTEM KALAN","BÜTÇE DIŞI KALAN"]
for c in cols:
    if c not in df_filtered.columns: df_filtered[c]=np.nan
df_show=df_filtered[cols].copy(); df_show.insert(0,"Seç",False)
st.write("**Bağlı kişiler (satır seç → PersonRef atanır)**")
edited = st.data_editor(df_show, use_container_width=True, hide_index=True, height=420,
                        disabled=[c for c in df_show.columns if c!="Seç"])
sel=None
chosen=edited.index[edited.get("Seç",False)==True].tolist() if "Seç" in edited.columns else []
if chosen:
    try:
        v=edited.loc[chosen[0],"PersonRef"]
        if pd.notna(v): sel=int(float(v))
    except: sel=None
if sel is not None: st.session_state.selected_ref=sel
selected_ref = st.session_state.selected_ref

# ================== SİDEBAR İŞLEM ALANLARI ==================
with st.sidebar:
    st.markdown("---"); st.subheader("🛠️ İşlem")
    if selected_ref is not None: st.success(f"Seçili PersonRef: {selected_ref}")
    manuel_ref = st.text_input("Veya Manuel PersonRef", value="" if selected_ref is None else str(selected_ref))
    tutar = st.number_input("Tutar (TL) — (istersen boş bırak)", step=100.0, min_value=0.0, value=0.0)
    islem = st.radio("İşlem Türü",
        ["Bütçeden Düş (Sistem Kalan)","Bütçeye Ekle (Sistem Kalan)","Bütçeden Düş (Bütçe Dışı Kalan)","Bütçeye Ekle (Bütçe Dışı Kalan)"], index=0)

# ================== İŞLEM FONKSİYONU ==================
def islem_yap(person_ref:int, tutar:float, islem_tipi:str, announce=True, do_rerun=True):
    dff=st.session_state.df.copy()
    ser=pd.to_numeric(dff["PersonRef"], errors="coerce")
    idxs=dff.index[ser==float(person_ref)]
    if len(idxs)==0:
        st.warning("Girilen PersonRef ile eşleşen kişi bulunamadı.")
        if announce: speak("Girilen kişi bulunamadı.")
        return
    i=idxs[0]

    # ---- Önceki değerler ----
    cur_sal = get_numeric(dff.at[i,"CurrentSalary"],0.0)
    new     = get_numeric(dff.at[i,"NewSalary"],0.0)
    bd      = get_numeric(dff.at[i,"BÜTÇE DIŞI TALEPLER İLE"],0.0)
    used    = cur_sal*1.4
    pre_sys = used - new
    pre_dis = used - bd

    # ---- Güncelle ----
    if islem_tipi=="Bütçeden Düş (Sistem Kalan)":
        dff.at[i,"NewSalary"]=new+float(tutar); verb="sistem kalandan düşüldü"; pool="Sistem"
    elif islem_tipi=="Bütçeye Ekle (Sistem Kalan)":
        dff.at[i,"NewSalary"]=new-float(tutar); verb="sistem kalana eklendi"; pool="Sistem"
    elif islem_tipi=="Bütçeden Düş (Bütçe Dışı Kalan)":
        dff.at[i,"BÜTÇE DIŞI TALEPLER İLE"]=bd+float(tutar); verb="bütçe dışı kalandan düşüldü"; pool="Bütçe Dışı"
    elif islem_tipi=="Bütçeye Ekle (Bütçe Dışı Kalan)":
        dff.at[i,"BÜTÇE DIŞI TALEPLER İLE"]=bd-float(tutar); verb="bütçe dışı kalana eklendi"; pool="Bütçe Dışı"
    else:
        st.warning("Bilinmeyen işlem tipi."); return

    # ---- Sonraki (normalize) ----
    dff = normalize_all(dff)
    st.session_state.df = dff

    # Satır tekrar bulunup sonrası metrikleri okunur
    ser2=pd.to_numeric(dff["PersonRef"], errors="coerce")
    j = dff.index[ser2==float(person_ref)][0]
    post_sys = get_numeric(dff.at[j,"SİSTEM KALAN"],0.0)
    post_dis = get_numeric(dff.at[j,"BÜTÇE DIŞI KALAN"],0.0)

    # Kim bilgileri
    row = dff.loc[j]
    fullname = str(row.get("FULLNAME","") or "")
    dep = str(row.get("DEPARTMAN","") or "")
    mans = manager_chain(row)

    # Kaydedilmemiş işlem kaydı (Kaydet'te geçmişe yazılır)
    st.session_state.unsaved_ops.append({
        "Zaman": dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "PersonRef": int(person_ref),
        "AdSoyad": fullname,
        "Departman": dep,
        "Yöneticiler": mans,
        "Tür": islem_tipi,
        "Havuz": pool,
        "Tutar": float(tutar),
        "Önce_SistemKalan": float(pre_sys),
        "Sonra_SistemKalan": float(post_sys),
        "Önce_BütçeDışıKalan": float(pre_dis),
        "Sonra_BütçeDışıKalan": float(post_dis),
    })

    st.success(f"İşlem uygulandı: {tutar:.2f} TL {verb}. (Geçmiş: Kaydet ile)")
    if announce: speak(f"{int(round(float(tutar)))} lira {verb}. Kaydet tuşuyla geçmişe eklenecek.")
    if do_rerun: st.rerun()

# ================== CLICK İÇİN GİRDİ ÇÖZÜMLE ==================
def resolve_click_inputs(manuel_ref, selected_ref, ui_amount, ui_islem, last_text):
    pref = None; pref_digits = None
    # PersonRef: manuel > seçili > Son'dan (ref) > Son'dan (ad-soyad)
    if isinstance(manuel_ref, str):
        trg = manuel_ref.strip()
        if trg:
            try: pref = int(float(trg.replace(",", ".")))
            except: pref = None
    if pref is None and selected_ref is not None:
        try: pref = int(float(selected_ref))
        except: pref = None
    if pref is None and last_text:
        pnum, pdig = extract_personref(last_text)
        if pnum is not None:
            pref = int(pnum); pref_digits = pdig
    if pref is None and last_text:
        pbyname, _nm = find_personref_by_name(st.session_state.df, last_text)
        if pbyname is not None: pref = int(pbyname)

    # Tutar: UI > sticky > Son
    amt = None
    try:
        if ui_amount and float(ui_amount) > 0: amt = float(ui_amount)
    except: pass
    if amt is None:
        stick = get_sticky_amount()
        if stick: amt = float(stick)
    if amt is None and last_text:
        a = extract_amount(last_text, pref_digits)
        if a: amt = float(a)

    # İşlem türü: Son'dan > UI
    op = parse_op_from_text(last_text, fallback_ui_op=ui_islem)
    return pref, amt, op

# ================== BUTONLAR ==================
cA,cB,cC=st.columns([1,1,1])
with cA:
    if st.button("İşlem Yap", use_container_width=True):
        last_text = st.session_state.get("last_final_text", "")
        pref, chosen_amt, op = resolve_click_inputs(manuel_ref, selected_ref, tutar, islem, last_text)
        if pref is None:
            st.warning("Kişi bulunamadı. Tabloda seçin, PersonRef girin veya 'Son' komutta isim/PersonRef geçsin.")
            speak("Kişi bulunamadı. Lütfen kişi seçin ya da PersonRef söyleyin.")
        elif not chosen_amt or float(chosen_amt) <= 0:
            st.warning("Tutar yok. Soldan girin ya da 'Son' cümlede tutarı söyleyin (örn. 80 TL).")
            speak("Tutar algılanmadı. Lütfen tutarı söyleyin veya girin.")
        elif not op:
            st.warning("İşlem türü anlaşılmadı. 'düş' veya 'ekle' deyin; 'bütçe dışı' derseniz oraya uygulanır.")
            speak("İşlem türü anlaşılmadı. Lütfen düş mü ekle mi olduğunu söyleyin.")
        else:
            try:
                islem_yap(int(pref), float(chosen_amt), op, do_rerun=True)
            except:
                st.warning("İşlem uygulanamadı. Girdileri kontrol edin.")

with cB:
    if st.session_state.unsaved_ops: st.info(f"Kaydedilmemiş işlem: {len(st.session_state.unsaved_ops)}")
    if st.button("Kaydet", type="primary", use_container_width=True):
        out=DEFAULT_EXCEL_PATH
        st.session_state.df.to_excel(out, index=False)
        st.cache_data.clear()  # Excel'i güncel oku
        # Geçmişe yaz:
        st.session_state.history.extend(st.session_state.unsaved_ops)
        st.session_state.unsaved_ops=[]
        st.success("Veriler kaydedildi — diğer kullanıcılar da aynı şekilde görecek.")
        speak("Veriler kaydedildi ve geçmişe işlendi.")
        st.rerun()

with cC:
    if st.button("Komut Örnekleri", use_container_width=True):
        st.info("Örnek: 'Bu kişinin sistemden seksen beş düş' | 'Ayşegül Ünal’ın bütçesine 5 TL ekle' | '… işlem yap'")
        speak("Örnek komutlar ekranınızda.")

# ================== CANLI YAZIM (Başlat/Durdur kontrollü) ==================
st.markdown("### 🎧 Canlı Yazım")
st.components.v1.html(f"""
<div style="border:1px dashed #bbb;padding:8px;border-radius:8px;background:#fbfbfb">
  <div><b>Canlı:</b> <span id="stt_live">{'Dinleniyor…' if st.session_state.listening else 'Kapalı'}</span></div>
  <div style="margin-top:6px"><b>Son:</b> <span id="stt_final">{st.session_state.get('last_final_text') or ''}</span></div>
  <div id="stt_dbg" style="margin-top:6px;font-size:12px;color:#888;"></div>
</div>
<script>
(function(){{
  const PY_SHOULD = {str(st.session_state.listening).lower()};
  const SR   = window.SpeechRecognition || window.webkitSpeechRecognition;
  const live = document.getElementById('stt_live');
  const fin  = document.getElementById('stt_final');
  const dbg  = document.getElementById('stt_dbg');
  const setLive = (t)=>{{ if(live) live.textContent = t; localStorage.setItem('stt_live_last', t||''); }};
  const setFinal= (t)=>{{ if(fin)  fin.textContent  = t; localStorage.setItem('stt_final_last', t||''); }};
  const log     = (x)=>{{ if(dbg)  dbg.textContent  = x; }};

  try {{
    const lastL = localStorage.getItem('stt_live_last'); if(lastL && live) live.textContent=lastL;
    const lastF = localStorage.getItem('stt_final_last'); if(lastF && fin) fin.textContent=lastF;
  }} catch(_ ){{}}

  if (!SR) {{ setLive('Tarayıcıda Ses Tanıma yok'); return; }}

  function ensureHandlers(rec){{
    if (rec.__handlersAttached) return;
    rec.__handlersAttached = true;
    rec.onresult = (e) => {{
      let interim = '', finalTxt = '';
      for (let i=e.resultIndex; i<e.results.length; i++) {{
        const t = e.results[i][0].transcript;
        if (e.results[i].isFinal) finalTxt += t; else interim += t;
      }}
      if (interim && interim.trim()) {{
        setLive(interim.trim());
        const it = interim.toLowerCase();
        if (/(i\\s*ş\\s*lem\\s*yap|islem\\s*yap|hemen\\s*uygula|uygula|tamam|onayla)/.test(it)) {{
          finalizeNow((localStorage.getItem('stt_live_last')||'').trim());
          return;
        }}
      }}
      if (finalTxt && finalTxt.trim()) {{
        setFinal(finalTxt.trim());
        finalizeNow(finalTxt.trim());
      }}
    }};
    rec.onerror = (ev) => {{
      if (!window.__stt_should_listen) return;
      const err = (ev && ev.error) || '';
      if (['aborted','no-speech','network','audio-capture'].includes(err)) {{
        setTimeout(()=>{{ try{{ rec.start(); }}catch(_ ){{}} }}, 150);
        return;
      }}
      log('Hata: ' + err);
      setTimeout(()=>{{ try{{ rec.start(); }}catch(_ ){{}} }}, 350);
    }};
    rec.onstart = ()=>{{ window.__stt_running=true; setLive('Dinleniyor…'); }};
    rec.onend   = ()=>{{ window.__stt_running=false; if (window.__stt_should_listen) setTimeout(()=>{{ try{{ rec.start(); }}catch(_ ){{}} }}, 200); else setLive('Kapalı'); }};
  }}

  function finalizeNow(text){{
    const finTxt = (text||'').trim();
    if (!finTxt) return;
    try {{
      if (window.__stt_rec) window.__stt_rec.stop();
    }} catch(_ ){{}}
    const u = new URL(location.href);
    u.searchParams.set('voice', finTxt);
    setTimeout(()=>{{ location.assign(u.toString()); }}, 20);
  }}

  async function askMic(){{
    if (!navigator.mediaDevices || !navigator.mediaDevices.getUserMedia) return true;
    try {{
      const s = await navigator.mediaDevices.getUserMedia({{audio:true}});
      try{{ s.getTracks().forEach(t=>t.stop()); }}catch(_ ){{}}
      return true;
    }} catch(e) {{
      setLive('Mikrofon engellendi — tarayıcı izinlerini kontrol edin');
      log('getUserMedia error: '+e);
      return false;
    }}
  }}

  window.sttStart = async function(){{
    window.__stt_should_listen = true;
    if (!window.__stt_rec) {{
      window.__stt_rec = new SR();
      window.__stt_rec.lang='tr-TR';
      window.__stt_rec.continuous=true;
      window.__stt_rec.interimResults=true;
      window.__stt_rec.maxAlternatives=1;
      ensureHandlers(window.__stt_rec);
    }} else {{
      ensureHandlers(window.__stt_rec);
    }}
    const ok = await askMic();
    if (!ok) return;
    try {{ window.__stt_rec.start(); setLive('Dinleniyor…'); }} catch(_ ){{}}
  }}

  window.sttStop = function(){{
    window.__stt_should_listen = false;
    try {{ window.__stt_rec && window.__stt_rec.stop(); }} catch(_ ){{}}
    setLive('Kapalı');
  }}

  if (PY_SHOULD) window.sttStart(); else window.sttStop();
}})();
</script>
""", height=170)

# ================== BAŞLAT / DURDUR (BUTONLAR JS TETİKLER) ==================
with st.sidebar:
    st.markdown("---"); st.subheader("🎧 Dinleme")
    colS, colT = st.columns(2)
    with colS:
        if st.button("🎤 Başlat", use_container_width=True, disabled=st.session_state.listening):
            st.session_state.force_listen = True
            st.session_state.listening = True
            st.components.v1.html("<script>try{window.sttStart && window.sttStart();}catch(e){}</script>", height=0)
            st.rerun()
    with colT:
        if st.button("⏹️ Durdur", use_container_width=True, disabled=not st.session_state.listening):
            st.session_state.force_listen = False
            st.session_state.listening = False
            st.components.v1.html("<script>try{window.sttStop && window.sttStop();}catch(e){}</script>", height=0)
            st.rerun()

# ================== SES KOMUTU -> PARSE/UYGULA ==================
def handle_command(text: str, ui_amount: float, ui_islem: str, ui_selected_ref: int|None, auto_apply: bool = True):
    df = st.session_state.df
    t = text.lower()

    trigger = any(k in t for k in ["işlem yap","islem yap","hemen uygula","uygula","onayla","tamam"])
    act=None
    if any(k in t for k in ["düş","dus","düşür","çıkar","cikar","eksilt","azalt"]): act="düş"
    elif any(k in t for k in ["ekle","arttır","artır","yükselt","yukselt"]): act="ekle"
    pool = "sistem"
    if any(k in t for k in ["bütçe dış","butce dis","dış","dis"]): pool="dis"
    op = ("Bütçeden Düş (Sistem Kalan)" if act=="düş" else "Bütçeye Ekle (Sistem Kalan)") if act and pool=="sistem" else \
         ("Bütçeden Düş (Bütçe Dışı Kalan)" if act=="düş" else "Bütçeye Ekle (Bütçe Dışı Kalan)") if act and pool=="dis" else None

    pref, pref_digits = extract_personref(t)
    if pref is None and ui_selected_ref is not None:
        pref = ui_selected_ref
    if pref is None:
        pref_by_name, name_found = find_personref_by_name(df, t)
        if pref_by_name is not None:
            pref = pref_by_name
            speak(f"{name_found} bulundu.")

    amt_voice = extract_amount(t, pref_digits)
    if amt_voice: set_sticky_amount(amt_voice)
    amt = float(amt_voice) if (amt_voice and amt_voice>0) else (float(ui_amount) if ui_amount and float(ui_amount)>0 else (get_sticky_amount() or None))

    # Toplu (tüm bağlılar) için kısayol
    managers = pd.concat([df["1.YÖNETİCİSİ"],df["2.YÖNETİCİSİ"],df["3.YÖNETİCİSİ"],df["4.YÖNETİCİSİ"]], ignore_index=True)
    opts = sorted([m for m in managers.dropna().unique() if str(m).strip()!=""])
    lowmap={m.lower():m for m in opts}
    hit=None
    for low,orig in lowmap.items():
        if low and low in t: hit=low; break
    scope=df
    if hit:
        msk=(df["1.YÖNETİCİSİ"].str.lower()==hit)|(df["2.YÖNETİCİSİ"].str.lower()==hit)|(df["3.YÖNETİCİSİ"].str.lower()==hit)|(df["4.YÖNETİCİSİ"].str.lower()==hit)
        scope=df[msk]
    allreq = any(kw in t for kw in ["tüm bağlı","tum bagli","hepsi","tamamı","tüm çalışan","tum calisan"])
    if not allreq and hit and pref is None: allreq=True

    if allreq and hit and op and amt is not None:
        refs = scope["PersonRef"].dropna().astype(float).astype(int).tolist()
        if not refs:
            st.warning("Bu yöneticiye bağlı kimse bulunamadı."); speak("Bu yöneticiye bağlı kimse bulunamadı."); return
        st.session_state.pending_batch={"manager":lowmap[hit],"op":op,"amount":float(amt),"refs":refs}
        speak(f"{lowmap[hit]} yöneticisinin {len(refs)} bağlısına {int(float(amt))} lira işlem için onay gerekiyor. 'Onayla' ya da 'İptal' diyebilirsiniz.")
        st.rerun(); return

    if op and amt is not None and pref is not None:
        if auto_apply or trigger:
            islem_yap(int(pref), float(amt), op)
            return
        else:
            st.info("Komut çözüldü. 'İşlem Yap' butonuyla uygulayabilirsiniz.")
            speak("Komut hazır. İşlem Yap'a basın.")
            return

    if trigger:
        missing=[]
        if pref is None: missing.append("kişi (seçin ya da adını söyleyin)")
        if amt is None:  missing.append("tutar (söyleyin ya da girin)")
        if op is None:   missing.append("işlem türü (düş/ekle)")
        if not missing:
            islem_yap(int(pref), float(amt), op)
            return
        msg=" , ".join(missing) + "."
        st.warning(msg); speak(msg)
    else:
        if amt is None:
            st.warning("Tutar algılanamadı. Cümlede tutarı söyleyin (ör. 'seksen beş', '85 TL') ya da sol taraftan girin.")
            speak("Tutar algılanamadı. Lütfen tutarı söyleyin veya ekrandan girin.")
        else:
            msg="Komut eksik. 'Bu kişinin sistemden 85 TL düş' (tabloda Seç) ya da 'PersonRef 12345 …'."
            st.warning(msg); speak(msg)

voice_param = get_query_param("voice")
if voice_param:
    if st.session_state.force_listen:
        st.session_state.listening = True
    vtxt=unquote(voice_param).strip()
    if vtxt!=st.session_state._last_voice:
        st.session_state._last_voice=vtxt
        st.session_state.last_final_text=vtxt
        st.info(f"🎤 Algılanan komut: **{vtxt}**")
        speak("Komut alındı.")
        st.components.v1.html("<script>const u=new URL(location.href);u.searchParams.delete('voice');history.replaceState({},'',u);</script>", height=0)
        if st.session_state.pending_batch:
            vlow=vtxt.lower()
            if any(w in vlow for w in ["onayla","evet","uygula","tamam"]):
                b=st.session_state.pending_batch
                for ref in b["refs"]:
                    try: islem_yap(int(ref), float(b["amount"]), b["op"], announce=False, do_rerun=False)
                    except: pass
                st.session_state.pending_batch=None
                speak("Toplu işlem uygulandı. Kaydet’e basarak geçmişe işleyin."); st.rerun()
            elif any(w in vlow for w in ["iptal","hayır","hayir","vazgeç","vazgec"]):
                st.session_state.pending_batch=None; speak("Toplu işlem iptal edildi."); st.rerun()
            else:
                handle_command(vtxt, tutar, islem, selected_ref, st.session_state.get("auto_apply", True))
        else:
            handle_command(vtxt, tutar, islem, selected_ref, st.session_state.get("auto_apply", True))

# ================== TOPLU ONAY KARTI ==================
if st.session_state.pending_batch:
    b = st.session_state.pending_batch
    st.warning(f"🧾 Toplu İşlem Bekliyor: **{b['manager']}** yöneticisinin **{len(b['refs'])}** bağlısına **{int(b['amount'])} TL** → **{b['op']}**")
    preview = st.session_state.df[st.session_state.df["PersonRef"].isin(b["refs"])][["PersonRef","FULLNAME","DEPARTMAN","1.YÖNETİCİSİ","2.YÖNETİCİSİ","3.YÖNETİCİSİ","4.YÖNETİCİSİ","CurrentSalary","NewSalary","BÜTÇE DIŞI TALEPLER İLE"]].copy()
    st.dataframe(preview, use_container_width=True, height=260)
    c_ok, c_cancel = st.columns(2)
    with c_ok:
        if st.button("✅ Onayla (Toplu Uygula)", type="primary", use_container_width=True):
            for ref in b["refs"]:
                try: islem_yap(int(ref), float(b["amount"]), b["op"], announce=False, do_rerun=False)
                except: pass
            st.session_state.pending_batch=None
            speak("Toplu işlem uygulandı. Kaydet’e basarak geçmişe işleyin."); st.rerun()
    with c_cancel:
        if st.button("❌ İptal", use_container_width=True):
            st.session_state.pending_batch=None; speak("Toplu işlem iptal edildi."); st.rerun()

# ================== GEÇMİŞ & İNDİR ==================
st.markdown("## 🧾 İşlem Geçmişi")
if not st.session_state.history:
    st.info("Henüz geçmiş kaydı yok. İşlem yap → Kaydet’e bas.")
else:
    hd=pd.DataFrame(st.session_state.history)
    try:
        hd["Zaman_dt"]=pd.to_datetime(hd["Zaman"]); hd=hd.sort_values("Zaman_dt",ascending=False).drop(columns=["Zaman_dt"])
    except: pass
    st.dataframe(hd, use_container_width=True, height=280)
    buf=io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w: hd.to_excel(w,index=False,sheet_name="Islem_Gecmisi")
    st.download_button("⬇️ İşlem Geçmişini İndir (Excel)", data=buf.getvalue(),
        file_name=f"Islem_Gecmisi_{dt.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

st.markdown("## ⬇️ Güncel Veriyi İndir (Excel)")
only = st.checkbox("Sadece seçili yönetici filtresi", value=False)
export = df_filtered.copy() if only else st.session_state.df.copy()
buf2=io.BytesIO()
with pd.ExcelWriter(buf2, engine="openpyxl") as w: export.to_excel(w,index=False,sheet_name="Veri")
st.download_button("⬇️ Veriyi İndir", data=buf2.getvalue(),
    file_name=f"Veri_{'filtreli_' if only else ''}{dt.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
