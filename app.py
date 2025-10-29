# Stok Hero – Görselli ve Eğlenceli Depo-Asistan Prototipi
# --------------------------------------------------------
# Notlar:
# - Ekstra paket gerekmeden çalışır: streamlit, pandas, openpyxl (Excel için)
# - Sol menüde akış adımları: İş Emri, Stok, Teklif, Satın Alma, Durum Takip
# - "Sürükle-bırak" hissi için; Stok listesinden satır seçip "Sepete ekle" butonu ile Teklif sepetine atılır.
# - Teklifte adetler düzenlenir; stok yetersiz olanlar otomatik Satın Alma listesine düşer.
# - Üretim amaçlı örnek şablon Excel indirebilirsiniz.
# --------------------------------------------------------

from __future__ import annotations
import io
from datetime import datetime
from typing import List

import pandas as pd
import streamlit as st

st.set_page_config(page_title="Stok Hero", layout="wide", page_icon="🦸")

# ----------------------
# Yardımcı Fonksiyonlar
# ----------------------

def _normalize_cols(df: pd.DataFrame) -> pd.DataFrame:
    """Kolon adlarını normalize eder ve yaygın karşılıkları yakalamaya çalışır."""
    mapping_sets = {
        "code": {"kod", "ürün kodu", "stok kodu", "malzeme kodu", "code", "sku"},
        "name": {"ad", "adı", "ürün", "ürün adı", "stok adı", "malzeme adı", "name"},
        "stock": {"stok", "mevcut", "miktar", "adet", "stock", "qty"},
        "price": {"fiyat", "birim fiyat", "tutar", "price", "unit price"},
        "unit": {"birim", "unit"},
    }

    new_cols = {}
    for c in df.columns:
        lc = str(c).strip().lower()
        target = None
        for key, names in mapping_sets.items():
            if lc in names:
                target = key
                break
        new_cols[c] = target if target else lc
    df = df.rename(columns=new_cols)

    # Zorunlu alanlar: name, stock. Diğerleri opsiyonel.
    if "name" not in df.columns:
        # en yakın tahmin
        for col in df.columns:
            if any(k in col for k in ["ürün", "malzeme", "stok adı", "name", "açıklama"]):
                df = df.rename(columns={col: "name"})
                break
    if "stock" not in df.columns:
        for col in df.columns:
            if any(k in col for k in ["stok", "miktar", "adet", "qty", "mevcut"]):
                df = df.rename(columns={col: "stock"})
                break

    # Tip düzeltmeleri
    if "stock" in df.columns:
        df["stock"] = pd.to_numeric(df["stock"], errors="coerce").fillna(0).astype(int)
    if "price" in df.columns:
        df["price"] = pd.to_numeric(df["price"], errors="coerce").fillna(0.0)
    if "unit" not in df.columns:
        df["unit"] = "adet"
    if "code" not in df.columns:
        # otomatik kod üret
        df["code"] = [f"MZ-{i:04d}" for i in range(1, len(df) + 1)]

    # Sütun sırası
    ordered = ["code", "name", "unit", "stock", "price"]
    other_cols = [c for c in df.columns if c not in ordered]
    return df[ordered + other_cols]


def _sample_excel() -> bytes:
    """Örnek Excel dosyası üretir (openpyxl varsa)."""
    sample = pd.DataFrame(
        {
            "Malzeme Kodu": ["KBL-1001", "KBL-2002", "KBL-3003", "KBL-4004"],
            "Malzeme Adı": ["Kablo CAT6 10m", "RJ45 Konnektör", "UPS 1kVA", "Rack Vidalı Set"],
            "Birim": ["adet", "adet", "adet", "kutu"],
            "Stok": [25, 300, 2, 15],
            "Birim Fiyat": [350.0, 4.2, 6500.0, 90.0],
        }
    )
    buf = io.BytesIO()
    # openpyxl yüklü değilse ImportError atar → üst seviye fallback kullanıyoruz
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        sample.to_excel(writer, index=False, sheet_name="Stok")
    buf.seek(0)
    return buf.read()


def _sample_csv() -> bytes:
    """Örnek CSV döndürür (openpyxl yoksa fallback)."""
    sample = pd.DataFrame(
        {
            "Malzeme Kodu": ["KBL-1001", "KBL-2002", "KBL-3003", "KBL-4004"],
            "Malzeme Adı": ["Kablo CAT6 10m", "RJ45 Konnektör", "UPS 1kVA", "Rack Vidalı Set"],
            "Birim": ["adet", "adet", "adet", "kutu"],
            "Stok": [25, 300, 2, 15],
            "Birim Fiyat": [350.0, 4.2, 6500.0, 90.0],
        }
    )
    return sample.to_csv(index=False).encode("utf-8")


def get_sample_download():
    """Örnek indirme butonu için veri, dosya adı, mime ve etiket döndürür."""
    try:
        data = _sample_excel()
        return data, "stok_ornek.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Örnek Excel İndir"
    except Exception:
        data = _sample_csv()
        return data, "stok_ornek.csv", "text/csv", "Örnek CSV İndir"


def icon_badge(text: str, color: str = "#2ecc71", emoji: str = "✅") -> str:
    return f"""
    <span style='display:inline-flex;align-items:center;gap:.5rem;background:{color};color:white;padding:.35rem .6rem;border-radius:999px;font-weight:600;'>
        <span style='font-size:1.1rem'>{emoji}</span>
        <span>{text}</span>
    </span>
    """


# ----------------------
# Session State
# ----------------------
if "stock_df" not in st.session_state:
    st.session_state.stock_df = pd.DataFrame()
if "cart_df" not in st.session_state:
    st.session_state.cart_df = pd.DataFrame(columns=["code", "name", "unit", "qty", "price"])  # Teklif
if "purchase_df" not in st.session_state:
    st.session_state.purchase_df = pd.DataFrame(columns=["code", "name", "unit", "needed_qty"])  # Satın alma
if "job_orders" not in st.session_state:
    st.session_state.job_orders = []  # basit liste


# ----------------------
# Üst Başlık ve Aşama Rozetleri
# ----------------------
st.markdown("""
<style>
:root { --accent:#6C5CE7; }
.block-container { padding-top: 1rem; }
h1 { letter-spacing: .5px }
</style>
""", unsafe_allow_html=True)

st.title("🦸 Stok Hero")
st.caption("Güvenlik şirketi iş emirlerini ikonlarla yöneten eğlenceli depo asistanı ✨")

# Durum rozeti
total_items = int(st.session_state.stock_df["stock"].sum()) if not st.session_state.stock_df.empty else 0
cart_count = len(st.session_state.cart_df)
purchase_count = len(st.session_state.purchase_df)

colA, colB, colC, colD = st.columns(4)
with colA:
    st.markdown(icon_badge("Yeni İş Emri", "#0984e3", "📋"), unsafe_allow_html=True)
with colB:
    st.markdown(icon_badge("Malzeme Hazırlanıyor", "#fdcb6e", "🧰"), unsafe_allow_html=True)
with colC:
    st.markdown(icon_badge("Teslim / Teklif", "#00b894", "🚚"), unsafe_allow_html=True)
with colD:
    st.markdown(icon_badge("Satın Alma Kuyruğu", "#d63031", "🛒"), unsafe_allow_html=True)

st.divider()

# ----------------------
# Sol Menü
# ----------------------
menu = st.sidebar.radio(
    "Menü",
    ["📋 İş Emri", "📦 Stok", "💰 Teklif", "🛒 Satın Alma", "🚦 Durum Takip"],
)

# ----------------------
# Veri Yükleme Alanı
# ----------------------
st.sidebar.subheader("📥 Stok Excel Yükle")
x_file = st.sidebar.file_uploader("Stok dosyası (.xlsx veya .csv)", type=["xlsx", "csv"])

if x_file is not None:
    try:
if x_file.name.lower().endswith(".csv"):
    raw = pd.read_csv(x_file)

    else:
        raw = pd.read_excel(x_file, sheet_name=0)
        st.session_state.stock_df = _normalize_cols(raw)
        st.sidebar.success("Stok yüklendi ✅")
    except Exception as e:
        st.sidebar.error(f"Excel okunamadı: {e}")
else:
    data, fname, mime, label = get_sample_download()
st.sidebar.download_button(
        label,
        data=data,
        file_name=fname,
        mime=mime,
        help="Kolonları otomatik tanıyacak şekilde örnek.",
    )

# ----------------------
# 1) İş Emri
# ----------------------
if menu.startswith("📋"):
    st.header("📋 Yeni İş Emri")
    with st.form("job_form", clear_on_submit=True):
        c1, c2 = st.columns([2, 1])
        with c1:
            requester = st.text_input("Talep Eden (Güvenlik Şirketi / Kişi)")
            description = st.text_area("Talep Açıklaması", placeholder="Örn: Kamera kurulum malzemeleri…")
        with c2:
            prio = st.selectbox("Öncelik", ["Normal", "Acil", "Düşük"]) 
            due = st.date_input("Termin Tarihi", value=datetime.today())
        submitted = st.form_submit_button("İş Emrini Kaydet ✍️")
    if submitted:
        st.session_state.job_orders.append(
            {
                "id": f"JOB-{len(st.session_state.job_orders)+1:04d}",
                "requester": requester,
                "desc": description,
                "prio": prio,
                "due": due.strftime("%Y-%m-%d"),
                "created": datetime.now().strftime("%Y-%m-%d %H:%M"),
            }
        )
        st.success("İş emri oluşturuldu. Şimdi Stok'tan malzeme seçebilirsiniz 👉")

    if st.session_state.job_orders:
        st.subheader("Açık İş Emirleri")
        st.dataframe(pd.DataFrame(st.session_state.job_orders))
    else:
        st.info("Henüz iş emri yok. Üstteki formdan ekleyin.")

# ----------------------
# 2) Stok
# ----------------------
elif menu.startswith("📦"):
    st.header("📦 Stok Listesi")
    if st.session_state.stock_df.empty:
        st.warning("Önce Excel stok dosyası yükleyin veya örneği indirin.")
    else:
        df = st.session_state.stock_df.copy()
        c1, c2, c3 = st.columns([3, 2, 2])
        with c1:
            q = st.text_input("Ara (ad/kod)")
        with c2:
            min_stock = st.number_input("Min. Stok Filtresi", min_value=0, value=0, step=1)
        with c3:
            sort_by = st.selectbox("Sırala", ["name", "stock", "price", "code"], index=0)
        if q:
            df = df[df["name"].str.contains(q, case=False, na=False) | df["code"].str.contains(q, case=False, na=False)]
        df = df[df["stock"] >= min_stock].sort_values(sort_by)

        st.caption("Satırları seçip alttan miktar belirleyin, sonra \"Sepete Ekle\".")
        selected = st.data_editor(
            df,
            use_container_width=True,
            hide_index=True,
            column_config={
                "price": st.column_config.NumberColumn("price", format="%.2f"),
                "stock": st.column_config.NumberColumn("stock", step=1),
            },
            disabled=["code", "name", "unit", "stock", "price"],
            selection_mode="multi-row",
            key="stock_editor",
        )

        sel_rows: List[int] = selected.get("selected_rows", []) if isinstance(selected, dict) else []
        qty = st.number_input("Seçilen her ürün için eklenecek adet", min_value=1, value=1, step=1)
        add_btn = st.button("🧲 Sepete Ekle", type="primary", use_container_width=True)
        if add_btn and sel_rows:
            add_items = df.iloc[sel_rows].copy()
            add_items["qty"] = qty
            add_items = add_items[["code", "name", "unit", "qty", "price"]]
            # Mevcut sepetle birleştir (aynı kod varsa qty topla)
            if st.session_state.cart_df.empty:
                st.session_state.cart_df = add_items
            else:
                merged = pd.concat([st.session_state.cart_df, add_items], ignore_index=True)
                st.session_state.cart_df = (
                    merged.groupby(["code", "name", "unit", "price"], as_index=False)["qty"].sum()
                )
            st.success(f"{len(sel_rows)} satır sepete eklendi 🧺")
        elif add_btn and not sel_rows:
            st.info("Önce en az bir satır seçin.")

# ----------------------
# 3) Teklif
# ----------------------
elif menu.startswith("💰"):
    st.header("💰 Teklif Sepeti")
    if st.session_state.cart_df.empty:
        st.info("Sepet boş. Stok'tan ürün ekleyin.")
    else:
        cart = st.session_state.cart_df.copy()
        # Adet düzenlenebilir
        st.caption("Adetleri düzenleyebilirsiniz. Sıfır (0) girerseniz satır silinir.")
        edited = st.data_editor(
            cart,
            hide_index=True,
            use_container_width=True,
            column_config={
                "qty": st.column_config.NumberColumn("qty", min_value=0, step=1),
                "price": st.column_config.NumberColumn("price", format="%.2f"),
            },
            key="cart_editor",
        )
        # 0 adet olanları at
        edited = edited[edited["qty"] > 0]
        st.session_state.cart_df = edited

        # Toplamlar
        edited["line_total"] = edited["qty"] * edited["price"].fillna(0)
        total = float(edited["line_total"].sum())

        # Stok kontrol
        purchase_rows = []
        if not st.session_state.stock_df.empty:
            stock_lookup = st.session_state.stock_df.set_index("code")["stock"].to_dict()
            for _, r in edited.iterrows():
                in_stock = int(stock_lookup.get(r["code"], 0))
                need = int(r["qty"]) - in_stock
                if need > 0:
                    purchase_rows.append({
                        "code": r["code"],
                        "name": r["name"],
                        "unit": r["unit"],
                        "needed_qty": need,
                    })
        st.session_state.purchase_df = pd.DataFrame(purchase_rows)

        c1, c2 = st.columns([2, 1])
        with c1:
            st.subheader("Özet")
            st.metric("Kalem Sayısı", len(edited))
            st.metric("Toplam Tutar", f"{total:,.2f} ₺")
        with c2:
            if not st.session_state.purchase_df.empty:
                st.warning("Stok yetersiz kalemler tespit edildi ve Satın Alma listesine eklendi 🛒")

        st.divider()
        st.subheader("Teklif PDF/Excel Çıktısı")
        # Basit Excel çıktısı
        out_buf = io.BytesIO()
        with pd.ExcelWriter(out_buf, engine="openpyxl") as writer:
            edited.to_excel(writer, index=False, sheet_name="Teklif")
            if not st.session_state.purchase_df.empty:
                st.session_state.purchase_df.to_excel(writer, index=False, sheet_name="SatinAlma")
        out_buf.seek(0)
        st.download_button(
            "📄 Teklif & Satın Alma Excel İndir",
            data=out_buf.read(),
            file_name=f"teklif_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

# ----------------------
# 4) Satın Alma
# ----------------------
elif menu.startswith("🛒"):
    st.header("🛒 Satın Alma Kuyruğu")
    if st.session_state.purchase_df.empty:
        st.info("Satın alma kuyruğu boş.")
    else:
        st.dataframe(st.session_state.purchase_df, use_container_width=True, hide_index=True)
        # Basit CSV çıkışı
        csv = st.session_state.purchase_df.to_csv(index=False).encode("utf-8")
        st.download_button(
            "🧾 CSV İndir",
            data=csv,
            file_name="satin_alma_kuyrugu.csv",
            mime="text/csv",
            use_container_width=True,
        )

# ----------------------
# 5) Durum Takip
# ----------------------
elif menu.startswith("🚦"):
    st.header("🚦 Durum Takip (Akış Haritası)")
    st.markdown(
        """
        İş akışı:
        
        **📋 İş Emri** ➜ **🧰 Malzeme Hazırlanıyor** ➜ **🚚 Teslim/Teklif** ➜ **🛒 Satın Alma** ➜ **✅ Tamamlandı**
        
        Her adımda ilgili ekranı kullanarak ilerleyebilirsiniz. Teklifte stok yetersizse kalemler otomatik Satın Alma kuyruğuna düşer.
        """
    )
    c1, c2, c3 = st.columns(3)
    with c1:
        st.metric("Açık İş Emri", len(st.session_state.job_orders))
    with c2:
        st.metric("Teklifteki Kalem", len(st.session_state.cart_df))
    with c3:
        st.metric("Satın Alma Bekleyen", len(st.session_state.purchase_df))

    if st.session_state.stock_df.empty:
        st.info("Stok yüklenmedi. Sol menüden Excel yükleyin.")
    else:
        total_stock = int(st.session_state.stock_df["stock"].sum())
        st.progress(min(total_stock / 1000, 1.0), text="Depo doluluk göstergesi (temsili)")

st.sidebar.caption("Made with ❤️  •  Stok Hero Prototype")
