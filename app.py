#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Sefer Aktarım — Web Sürümü (Streamlit)

PDF sefer bildirimlerini okuyup Excel taşıma kontrol listesine ekler ve
isteğe bağlı olarak her sefer için Boşaltma Kontrol Dökümanı (.docx / .pdf) üretir.
"""

from __future__ import annotations

import io
import tempfile
from datetime import datetime
from pathlib import Path

import streamlit as st

from core_logic import extract_pdf_data, process_pdfs, _DOCX_DESTEGI


st.set_page_config(
    page_title="Sefer Aktarım — Taşımacılık Kontrol Listesi",
    page_icon="🚚",
    layout="wide",
)

st.title("🚚 Sefer Aktarım")
st.caption("PDF sefer bildirimlerini Excel kontrol listesine aktarın ve Boşaltma Kontrol Dökümanı oluşturun.")

# Boşaltma Kontrol Dökümanı kartını ve imza kartını vurgulamak için özel stil
st.markdown(
    """
    <style>
    div.st-key-docx_karti,
    div.st-key-imza_karti {
        background-color: #C8E6C9;
        border: 2px solid #4CAF50 !important;
        border-radius: 10px;
        padding: 0.5rem;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# ---------------------------------------------------------------------------
# Oturum durumu
# ---------------------------------------------------------------------------
if "calisma_klasoru" not in st.session_state:
    st.session_state.calisma_klasoru = tempfile.mkdtemp(prefix="sefer_aktarim_")
if "sonuc" not in st.session_state:
    st.session_state.sonuc = None
if "log_mesajlari" not in st.session_state:
    st.session_state.log_mesajlari = []
if "logo_bytes" not in st.session_state:
    st.session_state.logo_bytes = None

CALISMA_KLASORU = Path(st.session_state.calisma_klasoru)

# ---------------------------------------------------------------------------
# SOL KENAR ÇUBUĞU — Firma Logosu + Boşaltan/Şoför Adı Soyadı
# ---------------------------------------------------------------------------
with st.sidebar:
    st.header("🏢 Firma Bilgileri")

    logo_dosya = st.file_uploader(
        "Firma Logosu",
        type=["png", "jpg", "jpeg"],
        key="logo_uploader",
        help="Excel'in ve Kontrol Dökümanının sol üst köşesine eklenir.",
    )
    if logo_dosya is not None:
        st.session_state.logo_bytes = logo_dosya.getvalue()

    if st.session_state.logo_bytes:
        st.image(st.session_state.logo_bytes, width=160)
        if st.button("🗑️ Logoyu Kaldır", use_container_width=True):
            st.session_state.logo_bytes = None
            st.rerun()
    else:
        st.caption("Logo yüklenmezse çıktılar logosuz oluşturulur.")

    st.divider()

    st.subheader("👤 İmza Bilgileri")
    docx_secili = st.session_state.get("docx_uret_checkbox", False)

    if docx_secili:
        with st.container(border=True, key="imza_karti"):
            st.caption("📋 Boşaltma Kontrol Dökümanı için bu bilgileri doldurun:")
            bosaltan_adi = st.text_input("Boşaltan Adı Soyadı")
            sofor_adi = st.text_input("Taşıyıcı/Şoför Adı Soyadı")
    else:
        bosaltan_adi = ""
        sofor_adi = ""
        st.caption(
            "ℹ️ Bu alanlar, sağdaki **'Boşaltma Kontrol Dökümanı oluştur'** "
            "kutucuğunu işaretlediğinizde burada aktif olacaktır."
        )

# ---------------------------------------------------------------------------
# ANA EKRAN — 1) Excel + 2) PDF yükleme (yan yana)
# ---------------------------------------------------------------------------
col_excel, col_pdf = st.columns(2)

with col_excel:
    st.subheader("1️⃣ Excel Taşıma Kontrol Listesi")
    excel_dosya = st.file_uploader(
        "Mevcut Excel dosyanızı yükleyin (.xlsx)",
        type=["xlsx"],
        help="Sefer verileri bu dosyaya tarih sırasına göre eklenecektir.",
    )

with col_pdf:
    st.subheader("2️⃣ Sefer Bildirimi PDF'leri")
    pdf_dosyalari = st.file_uploader(
        "Bir veya birden fazla sefer bildirimi PDF'i yükleyin",
        type=["pdf"],
        accept_multiple_files=True,
    )

st.divider()

# ---------------------------------------------------------------------------
# 3) Taşıma türü + 4) Kontrol Dökümanı seçeneği (yan yana)
# ---------------------------------------------------------------------------
col_tasima, col_docx = st.columns(2)

with col_tasima:
    st.subheader("3️⃣ Taşıma Türü")
    tasima_turu = st.radio(
        "Taşıma türünü seçin",
        options=["ADR-AMBALAJLI", "ADR-TANK", "ADR-DÖKME"],
        format_func=lambda v: {"ADR-AMBALAJLI": "Ambalajlı", "ADR-TANK": "Tank", "ADR-DÖKME": "Dökme"}[v],
        horizontal=True,
    )
    st.caption("UN 1202/1203 seçimden bağımsız olarak her zaman ADR-TANK olarak yazılır.")

with col_docx:
    st.subheader("4️⃣ Boşaltma Kontrol Dökümanı")
    with st.container(border=True, key="docx_karti"):
        docx_uret = st.checkbox(
            "**📋 Her sefer için Boşaltma Kontrol Dökümanı oluştur**",
            value=False,
            disabled=not _DOCX_DESTEGI,
            key="docx_uret_checkbox",
        )
        if docx_uret and _DOCX_DESTEGI:
            st.success("✅ Aktif — her sefer için ayrıca Kontrol Dökümanı (PDF/Word) üretilecek.")
        elif not _DOCX_DESTEGI:
            st.warning("python-docx kurulu değil, kontrol dökümanı üretimi devre dışı.")
        else:
            st.caption("İşaretlerseniz, aşağıda her plaka için tarih bilgileri istenecektir.")

plaka_ek_tarihler: dict[str, dict] = {}
plaka_muayene_tarihleri: dict[str, str] = {}

if docx_uret and _DOCX_DESTEGI and pdf_dosyalari:
    st.divider()
    st.subheader("🗓️ Araç Muayene ve Geçerlilik Tarihleri")
    st.caption("Her plaka için ayrı ayrı doldurun (GG.AA.YYYY formatında).")

    gecici_plakalar = []
    for pf in pdf_dosyalari:
        try:
            tmp_path = CALISMA_KLASORU / pf.name
            tmp_path.write_bytes(pf.getvalue())
            _, plaka = extract_pdf_data(tmp_path)
            if plaka and plaka not in [p[0] for p in gecici_plakalar]:
                gecici_plakalar.append((plaka, pf.name))
        except Exception:
            pass

    # Plaka kartlarını 2'li grid halinde yan yana diz
    for i in range(0, len(gecici_plakalar), 2):
        cift = gecici_plakalar[i:i + 2]
        grid_cols = st.columns(2)
        for col, (plaka, pdf_adi) in zip(grid_cols, cift):
            with col:
                with st.container(border=True):
                    st.markdown(f"**🚛 Plaka: {plaka}**")
                    st.caption(pdf_adi)
                    ara_muayene = st.text_input("Ara Muayene Tarihi", key=f"ara_{plaka}", placeholder="GG.AA.YYYY")
                    yangin_tup = st.text_input("Yangın Tüpü Geçerlilik Tarihi", key=f"yangin_{plaka}", placeholder="GG.AA.YYYY")
                    tmfb = st.text_input("TMFB Geçerlilik Tarihi", key=f"tmfb_{plaka}", placeholder="GG.AA.YYYY")
                    periyodik = st.text_input("Periyodik Muayene Tarihi", key=f"periyodik_{plaka}", placeholder="GG.AA.YYYY")
                    adr_uygunluk = st.text_input("ADR Uygunluk Belgesi Geçerlilik Tarihi", key=f"adr_{plaka}", placeholder="GG.AA.YYYY")

                if ara_muayene:
                    plaka_muayene_tarihleri[plaka] = ara_muayene
                ek = {}
                if yangin_tup: ek["yangin_tup"] = yangin_tup
                if tmfb: ek["tmfb"] = tmfb
                if adr_uygunluk: ek["adr_uygunluk"] = adr_uygunluk
                if periyodik: ek["periyodik_muayene"] = periyodik
                if ek:
                    plaka_ek_tarihler[plaka] = ek

st.divider()

# ---------------------------------------------------------------------------
# 5) Aktar butonu
# ---------------------------------------------------------------------------
calistir = st.button("▶️ Aktar", type="primary", use_container_width=True)

if calistir:
    if not excel_dosya:
        st.error("Lütfen bir Excel dosyası yükleyin.")
    elif not pdf_dosyalari:
        st.error("Lütfen en az bir PDF dosyası yükleyin.")
    else:
        with st.spinner("İşleniyor…"):
            excel_yolu = CALISMA_KLASORU / "girdi.xlsx"
            excel_yolu.write_bytes(excel_dosya.getvalue())

            pdf_yollari = []
            for pf in pdf_dosyalari:
                p = CALISMA_KLASORU / pf.name
                p.write_bytes(pf.getvalue())
                pdf_yollari.append(p)

            cikti_excel_yolu = CALISMA_KLASORU / f"{Path(excel_dosya.name).stem}_guncel.xlsx"
            docx_cikti_klasoru = CALISMA_KLASORU / "kontrol_dokumanlari"

            log_mesajlari = []

            sonuc = process_pdfs(
                excel_path=excel_yolu,
                pdf_paths=pdf_yollari,
                output_path=cikti_excel_yolu,
                log_cb=log_mesajlari.append,
                muayene_tarihleri=plaka_muayene_tarihleri,
                tasima_turu=tasima_turu,
                ek_tarihler=plaka_ek_tarihler,
                docx_uret=docx_uret,
                docx_cikti_klasor=docx_cikti_klasoru,
                bosaltan_adi=bosaltan_adi,
                sofor_adi=sofor_adi,
                docx_pdf_donustur=True,
                logo_bytes=st.session_state.logo_bytes,
            )

            st.session_state.sonuc = sonuc
            st.session_state.cikti_excel_yolu = cikti_excel_yolu
            st.session_state.log_mesajlari = log_mesajlari

# ---------------------------------------------------------------------------
# Sonuçlar
# ---------------------------------------------------------------------------
if st.session_state.sonuc:
    r = st.session_state.sonuc
    st.divider()
    st.subheader("✅ Sonuç")

    if r["eklenen"] > 0:
        c1, c2, c3 = st.columns(3)
        c1.metric("Yeni Eklenen Sefer", r["eklenen"])
        c2.metric("Atlanan (Zaten Mevcut)", r["atlanan"])
        c3.metric("Toplam Satır", r["toplam"])
    else:
        st.warning("Eklenecek yeni sefer bulunamadı (tüm seferler zaten mevcut olabilir).")

    with st.expander("İşlem Günlüğü", expanded=False):
        st.code("\n".join(st.session_state.log_mesajlari), language=None)

    sonuc_col_excel, sonuc_col_docx = st.columns(2)

    if r["eklenen"] > 0:
        with sonuc_col_excel:
            excel_yolu = st.session_state.cikti_excel_yolu
            st.markdown("#### 📊 Excel Taşıma Kontrol Listesi")
            if r.get("uretilen_dosyalar"):
                st.info(
                    "💡 Excel'deki Sefer No hücrelerine tıklanabilir köprüler eklendi. "
                    "Köprülerin çalışması için indirdiğiniz Kontrol Dökümanı PDF/Word "
                    "dosyalarını, Excel dosyasıyla **aynı klasöre** kaydedin."
                )
            st.download_button(
                "⬇️ Excel'i İndir",
                data=excel_yolu.read_bytes(),
                file_name=excel_yolu.name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

    uretilen = r.get("uretilen_dosyalar") or []
    if uretilen:
        with sonuc_col_docx:
            st.markdown("#### 📋 Boşaltma Kontrol Dökümanları")
            for dosya in uretilen:
                sn = dosya["sefer_no"]
                docx_yolu: Path = dosya["docx"]
                pdf_yolu: Path = dosya["pdf"]

                st.markdown(f"**Sefer No: {sn}**")
                dc1, dc2 = st.columns(2)
                with dc1:
                    if pdf_yolu and pdf_yolu.is_file():
                        st.download_button(
                            "⬇️ PDF",
                            data=pdf_yolu.read_bytes(),
                            file_name=pdf_yolu.name,
                            mime="application/pdf",
                            key=f"pdf_{sn}",
                            use_container_width=True,
                        )
                    else:
                        st.caption("PDF dönüşümü yapılamadı")
                with dc2:
                    if docx_yolu and docx_yolu.is_file():
                        st.download_button(
                            "⬇️ Word",
                            data=docx_yolu.read_bytes(),
                            file_name=docx_yolu.name,
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            key=f"docx_{sn}",
                            use_container_width=True,
                        )
