# Stok Hero – Görselli ve Eğlenceli Depo-Asistan Prototipi
# --------------------------------------------------------
# Notlar:
# - Ekstra paket gerekmeden çalışır: streamlit, pandas, openpyxl (Excel için)
# - Sol menüde akış adımları: Cari Hesap, Teklif, İş Emri, Stok, Satın Alma, Durum Takip
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
from fpdf import FPDF

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

def generate_quote_pdf(
    company: str,
    contact: str,
    email: str,
    phone: str,
    items: pd.DataFrame,
    subtotal: float,
    labor_cost: float,
    grand_total: float,
    notes: str,
) -> bytes:
    pdf = FPDF()
    pdf.add_page()
    pdf.set_auto_page_break(auto=True, margin=15)

    # Header / Logo area
    pdf.set_fill_color(108, 92, 231)
    pdf.set_text_color(255, 255, 255)
    pdf.set_font("Helvetica", "B", 28)
    pdf.cell(0, 18, "BFM", ln=True, align="C", fill=True)

    pdf.ln(6)
    pdf.set_text_color(0, 0, 0)
    pdf.set_font("Helvetica", "B", 14)
    pdf.cell(0, 10, "Teklif Özeti", ln=True)

    pdf.set_font("Helvetica", "", 11)
    pdf.cell(0, 7, f"Tarih: {datetime.now().strftime('%d.%m.%Y')}", ln=True)
    pdf.cell(0, 7, f"Firma: {company}", ln=True)
    if contact:
        pdf.cell(0, 7, f"Yetkili: {contact}", ln=True)
    if phone:
        pdf.cell(0, 7, f"Telefon: {phone}", ln=True)
    if email:
        pdf.cell(0, 7, f"E-posta: {email}", ln=True)

    pdf.ln(4)
    pdf.set_font("Helvetica", "B", 11)
    pdf.set_fill_color(240, 240, 240)
    headers = ["Kod", "Ürün", "Birim", "Adet", "Birim Fiyat", "Tutar"]
    widths = [25, 75, 20, 20, 28, 28]
    for header, width in zip(headers, widths):
        pdf.cell(width, 9, header, border=1, align="C", fill=True)
    pdf.ln(9)

    pdf.set_font("Helvetica", "", 10)
    for _, row in items.iterrows():
        name = str(row.get("name", ""))
        if len(name) > 42:
            name = name[:39] + "…"
        qty = row.get("qty", 0)
        try:
            qty_val = int(float(qty))
        except (TypeError, ValueError):
            qty_val = 0
        price_val = float(row.get("price", 0) or 0)
        line_total = float(row.get("line_total", price_val * qty_val) or 0)

        cells = [
            str(row.get("code", "")),
            name,
            str(row.get("unit", "")),
            f"{qty_val}",
            f"{price_val:,.2f}",
            f"{line_total:,.2f}",
        ]

        aligns = ["L", "L", "C", "C", "R", "R"]
        for content, width, align in zip(cells, widths, aligns):
            pdf.cell(width, 8, content, border=1, align=align)
        pdf.ln(8)

    pdf.ln(4)
    pdf.set_font("Helvetica", "B", 12)
    pdf.cell(0, 8, "Tutar Özeti", ln=True)
    pdf.set_font("Helvetica", "", 11)
    pdf.cell(0, 7, f"Ürün Toplamı: {subtotal:,.2f} ₺", ln=True)
    pdf.cell(0, 7, f"İşçilik / Hizmet: {labor_cost:,.2f} ₺", ln=True)
    pdf.set_font("Helvetica", "B", 12)
    pdf.cell(0, 8, f"Genel Toplam: {grand_total:,.2f} ₺", ln=True)

    if notes:
        pdf.ln(4)
        pdf.set_font("Helvetica", "B", 12)
        pdf.cell(0, 8, "Notlar", ln=True)
        pdf.set_font("Helvetica", "", 11)
        pdf.multi_cell(0, 6, notes)

    pdf.ln(12)
    pdf.set_font("Helvetica", "", 11)
    pdf.cell(0, 7, "Müşteri Onayı:", ln=True)
    pdf.ln(14)
    pdf.cell(0, 7, "İsim / Ünvan: _______________________________", ln=True)
    pdf.ln(12)
    pdf.cell(0, 7, "İmza: _______________________________", ln=True)

    return pdf.output(dest="S").encode("latin-1")


# ----------------------
# Session State
# ----------------------
if "stock_df" not in st.session_state:
    st.session_state.stock_df = pd.DataFrame()
if "cart_df" not in st.session_state:
    st.session_state.cart_df = pd.DataFrame(columns=["code", "name", "unit", "qty", "price"])  # Teklif
if "purchase_df" not in st.session_state:
    st.session_state.purchase_df = pd.DataFrame(columns=["code", "name", "unit", "needed_qty"])  # Satın alma
if "customer_accounts" not in st.session_state:
    st.session_state.customer_accounts = pd.DataFrame(
        columns=[
            "Firma / Cari",
            "İlgili Kişi",
            "E-posta",
            "Telefon",
            "Notlar",
            "Son Teklif Tutarı",
            "Statü",
            "Kayıt Tarihi",
        ]
    )
if "job_orders" not in st.session_state:
    st.session_state.job_orders = []  # basit liste
if "converted_quotes" not in st.session_state:
    st.session_state.converted_quotes = []
if "quote_counter" not in st.session_state:
    st.session_state.quote_counter = 0
if "selected_customer" not in st.session_state:
    st.session_state.selected_customer = None
if "labor_cost" not in st.session_state:
    st.session_state.labor_cost = 0.0
if "quote_note" not in st.session_state:
    st.session_state.quote_note = ""

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
    st.markdown(icon_badge("Yeni Teklif", "#6C5CE7", "💰"), unsafe_allow_html=True)
with colB:
    st.markdown(icon_badge("İş Emri Hazırlığı", "#0984e3", "📋"), unsafe_allow_html=True)
with colC:
    st.markdown(icon_badge("Malzeme Hazırlanıyor", "#fdcb6e", "🧰"), unsafe_allow_html=True)
with colD:
    st.markdown(icon_badge("Satın Alma Kuyruğu", "#d63031", "🛒"), unsafe_allow_html=True)

st.divider()

# ----------------------
# Sol Menü
# ----------------------
menu = st.sidebar.radio(
    "Menü",
    ["👥 Cari Hesap", "💰 Teklif", "📋 İş Emri", "📦 Stok", "🛒 Satın Alma", "🚦 Durum Takip"],
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
        label=label,
        data=data,
        file_name=fname,
        mime=mime,
        help="Kolonları otomatik tanıyacak şekilde örnek.",
    )

# ----------------------
# 1) Teklif
# ----------------------
if menu.startswith("💰"):
    st.header("💰 Teklif Hazırlığı")

    if st.session_state.customer_accounts.empty:
        st.info("Önce Cari Hesaplar sekmesinden teklif hazırlanacak firmayı ekleyin.")
    else:
        customers = st.session_state.customer_accounts["Firma / Cari"].tolist()
        default_idx = 0
        if st.session_state.selected_customer in customers:
            default_idx = customers.index(st.session_state.selected_customer)
        selected_customer = st.selectbox(
            "Teklif hazırlanacak cari hesap",
            customers,
            index=default_idx,
        )
        st.session_state.selected_customer = selected_customer

        # Cari özet bilgisi
        customer_row = st.session_state.customer_accounts[
            st.session_state.customer_accounts["Firma / Cari"] == selected_customer
        ].iloc[0]

        def _display_customer_value(val: object) -> str:
            if pd.isna(val) or val is None or str(val).strip() == "":
                return "-"
            return str(val)

        c1, c2, c3 = st.columns(3)
        with c1:
            st.metric("İlgili Kişi", _display_customer_value(customer_row.get("İlgili Kişi")))
        with c2:
            st.metric("Telefon", _display_customer_value(customer_row.get("Telefon")))
        with c3:
            st.metric("Statü", _display_customer_value(customer_row.get("Statü")))

        st.divider()

        # Stoktan ürün seçimi
        st.subheader("Ürün Seçimi")
        if st.session_state.stock_df.empty:
            st.warning("Stok verisi yüklenmemiş. Sol menüden Excel yükleyin.")
        else:
            with st.expander("📦 Stok listesinden ürün seç", expanded=st.session_state.cart_df.empty):
                stock_df = st.session_state.stock_df.copy()
                s1, s2 = st.columns([2, 1])
                with s1:
                    stock_query = st.text_input("Ara (ad/kod)", key="quote_stock_search")
                with s2:
                    sort_opt = st.selectbox(
                        "Sırala",
                        ["name", "stock", "price", "code"],
                        key="quote_stock_sort",
                    )
                if stock_query:
                    stock_df = stock_df[
                        stock_df["name"].str.contains(stock_query, case=False, na=False)
                        | stock_df["code"].str.contains(stock_query, case=False, na=False)
                    ]
                stock_df = stock_df.sort_values(sort_opt)

                selection_col = "__select__"
                qty_col = "__qty__"
                selection_view = stock_df.copy()
                selection_view[selection_col] = False
                selection_view[qty_col] = 1

                edited_stock = st.data_editor(
                    selection_view,
                    hide_index=True,
                    use_container_width=True,
                    column_config={
                        "price": st.column_config.NumberColumn("price", format="%.2f"),
                        "stock": st.column_config.NumberColumn("stock", step=1),
                        selection_col: st.column_config.CheckboxColumn("Seç"),
                        qty_col: st.column_config.NumberColumn("Adet", min_value=1, step=1),
                    },
                    disabled=["code", "name", "unit", "stock", "price"],
                    key="quote_stock_editor",
                )

                add_from_stock = st.button("Seçilenleri sepete ekle", type="primary", key="quote_add_btn")
                if add_from_stock:
                    if isinstance(edited_stock, pd.DataFrame):
                        chosen = edited_stock[edited_stock[selection_col]]
                    else:
                        chosen = pd.DataFrame(edited_stock)
                        chosen = chosen[chosen[selection_col]]
                    if not chosen.empty:
                        add_df = chosen[["code", "name", "unit", qty_col, "price"]].copy()
                        add_df.rename(columns={qty_col: "qty"}, inplace=True)
                        add_df["qty"] = pd.to_numeric(add_df["qty"], errors="coerce").fillna(1).astype(int)
                        if st.session_state.cart_df.empty:
                            st.session_state.cart_df = add_df
                        else:
                            merged = pd.concat([st.session_state.cart_df, add_df], ignore_index=True)
                            st.session_state.cart_df = (
                                merged.groupby(["code", "name", "unit", "price"], as_index=False)["qty"].sum()
                            )
                        st.success(f"{len(chosen)} ürün sepete eklendi 🧺")
                    else:
                        st.info("Sepete eklemek için en az bir satır seçin.")

        st.divider()

        st.subheader("Teklif Sepeti")
        if st.session_state.cart_df.empty:
            st.info("Sepet boş. Üstteki stok listesinden ürün seçin.")
        else:
            cart = st.session_state.cart_df.copy()
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
            edited = edited[edited["qty"] > 0]
            st.session_state.cart_df = edited

            edited["line_total"] = edited["qty"] * edited["price"].fillna(0)
            total = float(edited["line_total"].sum())

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

            summary_col, total_col, info_col = st.columns([1.5, 1, 1])
            with summary_col:
                st.subheader("Özet")
                st.metric("Cari", selected_customer)
                st.metric("Kalem Sayısı", len(edited))
                st.metric("Ürün Toplamı", f"{total:,.2f} ₺")
            with total_col:
                st.subheader("İşçilik / Toplam")
                labor_cost = st.number_input(
                    "İşçilik / Hizmet Tutarı (₺)",
                    min_value=0.0,
                    step=50.0,
                    format="%.2f",
                    key="labor_cost",
                )
                grand_total = total + float(labor_cost or 0.0)
                st.metric("Genel Toplam", f"{grand_total:,.2f} ₺")
            with info_col:
                if not st.session_state.purchase_df.empty:
                    st.warning("Stok yetersiz kalemler tespit edildi ve Satın Alma listesine eklendi 🛒")
         
            notes = st.text_area("Teklif Notları / İşçilik Detayı", key="quote_note")

            st.divider()
            st.subheader("Teklif PDF/Excel Çıktısı")
            out_buf = io.BytesIO()
            with pd.ExcelWriter(out_buf, engine="openpyxl") as writer:
                edited.assign(Cari=selected_customer).to_excel(writer, index=False, sheet_name="Teklif")
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

            contact_val = _display_customer_value(customer_row.get("İlgili Kişi"))
            email_val = _display_customer_value(customer_row.get("E-posta"))
            phone_val = _display_customer_value(customer_row.get("Telefon"))

            contact_val = "" if contact_val == "-" else contact_val
            email_val = "" if email_val == "-" else email_val
            phone_val = "" if phone_val == "-" else phone_val

            pdf_bytes = generate_quote_pdf(
                company=selected_customer,
                contact=contact_val,
                email=email_val,
                phone=phone_val,
                items=edited,
                subtotal=total,
                labor_cost=float(labor_cost or 0.0),
                grand_total=grand_total,
                notes=notes,
            )

            st.download_button(
                "🖨️ Teklif PDF İndir",
                data=pdf_bytes,
                file_name=f"teklif_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf",
                mime="application/pdf",
                use_container_width=True,
            )

            st.divider()
            convert_btn = st.button(
                "✅ Müşteri onayı alındı • İş Emri Oluştur",
                type="primary",
                use_container_width=True,
                key="quote_convert_btn",
            )
            if convert_btn:
                if edited.empty:
                    st.warning("Onaylanacak teklif kalemi bulunamadı.")
                else:
                    st.session_state.quote_counter += 1
                    quote_id = f"TEKLIF-{st.session_state.quote_counter:04d}"
                    job_id = f"JOB-{len(st.session_state.job_orders)+1:04d}"
                    short_note = notes.strip() if isinstance(notes, str) else ""
                    if short_note:
                        original_note = short_note
                        short_note = original_note[:140]
                        if len(original_note) > 140:
                            short_note += "…"
                    else:
                        short_note = f"{len(edited)} kalem • {grand_total:,.2f} ₺"

                    st.session_state.job_orders.append(
                        {
                            "id": job_id,
                            "requester": selected_customer,
                            "desc": short_note,
                            "prio": "Normal",
                            "due": datetime.now().strftime("%Y-%m-%d"),
                            "created": datetime.now().strftime("%Y-%m-%d %H:%M"),
                            "source": "Teklif",
                            "quote_id": quote_id,
                        }
                    )

                    st.session_state.converted_quotes.append(
                        {
                            "quote_id": quote_id,
                            "job_id": job_id,
                            "customer": selected_customer,
                            "item_count": len(edited),
                            "subtotal": total,
                            "labor_cost": float(labor_cost or 0.0),
                            "grand_total": grand_total,
                            "approved_at": datetime.now().strftime("%Y-%m-%d %H:%M"),
                        }
                    )

                    st.session_state.cart_df = pd.DataFrame(
                        columns=["code", "name", "unit", "qty", "price"]
                    )
                    st.session_state.purchase_df = pd.DataFrame(
                        columns=["code", "name", "unit", "needed_qty"]
                    )
                    st.session_state.labor_cost = 0.0
                    st.session_state.quote_note = ""

                    st.success(
                        f"{selected_customer} için {quote_id} numaralı teklif iş emrine dönüştürüldü. (İş Emri: {job_id})"
                    )
# ----------------------
# 2) İş Emri
# ----------------------
elif menu.startswith("📋"):
    st.header("📋 Yeni İş Emri")
    if st.session_state.selected_customer:
        st.caption(f"Seçili cari: {st.session_state.selected_customer}")
    if not st.session_state.cart_df.empty:
        st.caption(f"Teklif sepetinde {len(st.session_state.cart_df)} kalem hazır.")
    with st.form("job_form", clear_on_submit=True):
        c1, c2 = st.columns([2, 1])
        with c1:
            requester_default = st.session_state.selected_customer or ""
            requester = st.text_input("Talep Eden (Güvenlik Şirketi / Kişi)", value=requester_default)
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
                "source": "Manuel",
                "quote_id": None,
            }
        )
        st.success("İş emri oluşturuldu. Teklif ve stok adımlarından ilerlemeye devam edebilirsiniz 👉")

    if st.session_state.job_orders:
        st.subheader("Açık İş Emirleri")
        st.dataframe(pd.DataFrame(st.session_state.job_orders))
    else:
        st.info("Henüz iş emri yok. Üstteki formdan ekleyin.")

    if st.session_state.converted_quotes:
        st.subheader("İş Emrine Dönüşmüş Teklifler")
        converted_df = pd.DataFrame(
            [
                {
                    "Teklif No": q["quote_id"],
                    "İş Emri No": q["job_id"],
                    "Cari": q["customer"],
                    "Kalem Sayısı": q["item_count"],
                    "Toplam (₺)": q["grand_total"],
                    "Onay Tarihi": q["approved_at"],
                }
                for q in st.session_state.converted_quotes
            ]
        )
        if not converted_df.empty:
            converted_df = converted_df.sort_values("Onay Tarihi", ascending=False)
            st.dataframe(converted_df, use_container_width=True, hide_index=True)
    else:
        st.info("Onaylanmış teklif bulunmuyor.")
# ----------------------
# 3) Stok
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

        st.caption("Satırları işaretleyip alttan miktar belirleyin, sonra \"Sepete Ekle\".")

        select_col = "__select__"
        display_df = df.copy()
        display_df[select_col] = False

        edited = st.data_editor(
            display_df,
            use_container_width=True,
            hide_index=True,
            column_config={
                "price": st.column_config.NumberColumn("price", format="%.2f"),
                "stock": st.column_config.NumberColumn("stock", step=1),
                    select_col: st.column_config.CheckboxColumn(
                    "Seç",
                    help="Sepete eklenecek satırları işaretleyin.",
                    default=False,
                ),
            },
            disabled=["code", "name", "unit", "stock", "price"],
            key="stock_editor",
        )

        if isinstance(edited, pd.DataFrame):
            sel_rows: List[int] = edited.index[edited[select_col]].tolist()
        else:
            sel_rows = []
        qty = st.number_input("Seçilen her ürün için eklenecek adet", min_value=1, value=1, step=1)
        add_btn = st.button("🧲 Sepete Ekle", type="primary", use_container_width=True)
        if add_btn and sel_rows:
            add_items = df.loc[sel_rows].copy()
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
# 4) Cari Hesaplar
# ----------------------
elif menu.startswith("👥"):
    st.header("👥 Cari Hesaplar")
    st.caption("Teklif vereceğiniz müşteri ve firmaları burada saklayın.")

    with st.form("customer_form", clear_on_submit=True):
        c1, c2 = st.columns(2)
        with c1:
            account_name = st.text_input("Firma / Cari Adı", placeholder="Örn: Parlak Güvenlik A.Ş.")
            contact_name = st.text_input("İlgili Kişi", placeholder="Örn: Ayşe Yılmaz")
            email = st.text_input("E-posta", placeholder="ornek@firma.com")
            phone = st.text_input("Telefon", placeholder="0 (5xx) xxx xx xx")
        with c2:
            status = st.selectbox("Statü", ["Potansiyel", "Teklif Verildi", "Kazandı", "Kaybedildi"])
            last_quote = st.number_input(
                "Son Teklif Tutarı (₺)", min_value=0.0, value=0.0, step=100.0, format="%0.2f"
            )
            notes = st.text_area("Notlar", placeholder="Gereksinimler, hatırlatmalar…")
        submitted = st.form_submit_button("Cari Kaydı Ekle 📇")

    if submitted:
        if account_name:
            new_row = pd.DataFrame(
                [
                    {
                        "Firma / Cari": account_name,
                        "İlgili Kişi": contact_name,
                        "E-posta": email,
                        "Telefon": phone,
                        "Notlar": notes,
                        "Son Teklif Tutarı": last_quote if last_quote else None,
                        "Statü": status,
                        "Kayıt Tarihi": datetime.now().strftime("%Y-%m-%d %H:%M"),
                    }
                ]
            )
            st.session_state.customer_accounts = pd.concat(
                [st.session_state.customer_accounts, new_row], ignore_index=True
            )
            st.success("Cari hesap kaydedildi ✅")
        else:
            st.warning("Firma / Cari adı zorunludur.")

    if st.session_state.customer_accounts.empty:
        st.info("Henüz cari hesap eklenmedi. Yukarıdaki formdan ekleyin.")
    else:
        st.subheader("Kayıtlı Cari Hesaplar")
        st.dataframe(st.session_state.customer_accounts, use_container_width=True, hide_index=True)

        export_buf = io.BytesIO()
        with pd.ExcelWriter(export_buf, engine="openpyxl") as writer:
            st.session_state.customer_accounts.to_excel(writer, index=False, sheet_name="CariHesaplar")
        export_buf.seek(0)
        st.download_button(
            "📥 Cari Hesap Excel İndir",
            data=export_buf.read(),
            file_name="cari_hesaplar.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

# ----------------------
# 5) Satın Alma
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
# 6) Durum Takip
# ----------------------
elif menu.startswith("🚦"):
    st.header("🚦 Durum Takip (Akış Haritası)")
    st.markdown(
        """
        İş akışı:
        
        **💰 Teklif** ➜ **📋 İş Emri** ➜ **🧰 Malzeme Hazırlanıyor** ➜ **🚚 Teslim** ➜ **🛒 Satın Alma** ➜ **✅ Tamamlandı**
        
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
