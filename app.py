#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Sefer Aktarım & Atık Gönderim — Web Sürümü (Streamlit)

İki modül:
  • Boşaltma: PDF sefer bildirimleri → Excel + Boşaltma Kontrol Dökümanı
  • Gönderim: Export XLS (atık gönderimleri) → Excel + Gönderim Kontrol Dökümanı

Yüklenen dosya tipine göre mod otomatik belirlenir:
  • PDF(ler) → Boşaltma modülü
  • XLS/XLSX → Gönderim modülü
"""
from __future__ import annotations

import re
import tempfile
from pathlib import Path

import streamlit as st

from core_logic import extract_pdf_data, process_pdfs, _DOCX_DESTEGI, docx_to_pdf
from export_islem import export_oku
from gonderim_excel import process_export
from gonderim_doldur import gonderim_dokumani_olustur

st.set_page_config(
    page_title="Sefer Aktarım & Atık Gönderim",
    page_icon="🚚",
    layout="wide",
)

# ---------------------------------------------------------------------------
# CSS
# ---------------------------------------------------------------------------
st.markdown("""
<style>
div.st-key-docx_karti, div.st-key-imza_karti,
div.st-key-gonderim_karti, div.st-key-gonderim_ana_karti {
    background-color: #FCE4EC;
    border: 2px solid #F48FB1 !important;
    border-radius: 10px;
    padding: 0.5rem;
}
</style>
""", unsafe_allow_html=True)

# ---------------------------------------------------------------------------
# Oturum durumu
# ---------------------------------------------------------------------------
for key, default in [
    ("calisma_klasoru", None),
    ("sonuc", None),
    ("log_mesajlari", []),
    ("logo_bytes", None),
    ("mod", None),      # "bosaltma" | "gonderim" | None
]:
    if key not in st.session_state:
        st.session_state[key] = default

if st.session_state.calisma_klasoru is None:
    st.session_state.calisma_klasoru = tempfile.mkdtemp(prefix="sefer_aktarim_")

CALISMA_KLASORU = Path(st.session_state.calisma_klasoru)

# ---------------------------------------------------------------------------
# SIDEBAR — Firma Bilgileri (her iki modülde de ortak)
# ---------------------------------------------------------------------------
with st.sidebar:
    st.header("🏢 Firma Bilgileri")

    # 1️⃣ Logo
    st.markdown("**1️⃣ Firma Logosu**")
    logo_dosya = st.file_uploader("Logo", type=["png","jpg","jpeg"],
        key="logo_uploader", label_visibility="collapsed",
        help="Excel ve Kontrol Dökümanlarına eklenir.")
    if logo_dosya:
        st.session_state.logo_bytes = logo_dosya.getvalue()
    if st.session_state.logo_bytes:
        col_l, col_r = st.columns([2,1])
        with col_l:
            st.image(st.session_state.logo_bytes, width=120)
        with col_r:
            if st.button("🗑️", help="Logoyu kaldır"):
                st.session_state.logo_bytes = None
                st.rerun()
    else:
        st.caption("📎 Logo yükleyin (isteğe bağlı)")

    st.divider()

    # 2️⃣ Gönderici Firma
    st.markdown("**2️⃣ Gönderici Firma**")
    gonderici_firma = st.text_input("Firma Unvanı", key="gonderici_firma",
                                     placeholder="Firma adını girin")

    st.divider()

    # 3️⃣ Boşaltma Kontrol Formu
    st.markdown("**3️⃣ Boşaltma Kontrol Formu**")
    docx_secili = st.session_state.get("docx_uret_checkbox", False)
    if docx_secili:
        with st.container(border=True, key="imza_karti"):
            st.caption("📋 Boşaltma Kontrol Dökümanı için doldurun:")
            bosaltan_adi = st.text_input("Boşaltan", key="bosaltan_adi", placeholder="Ad Soyad")
            sofor_adi = st.text_input("Şoför", key="sofor_adi", placeholder="Ad Soyad")
    else:
        bosaltan_adi = sofor_adi = ""
        st.caption("Ana ekranda 'Boşaltma Kontrol Dökümanı oluştur' kutucuğunu işaretleyin.")

    st.divider()

    # 4️⃣ Gönderim Kontrol Formu
    st.markdown("**4️⃣ Gönderim Kontrol Formu**")
    gonderim_docx_secili = st.session_state.get("gonderim_docx_uret", False)
    if gonderim_docx_secili:
        with st.container(border=True, key="gonderim_karti"):
            st.caption("📋 Gönderim Kontrol Dökümanı için doldurun:")
            gonderici_adi = st.text_input("Gönderen", key="gonderici_adi", placeholder="Ad Soyad")
            sofor_adi_g = st.text_input("Şoför", key="sofor_adi_g", placeholder="Ad Soyad")
    else:
        gonderici_adi = sofor_adi_g = ""
        st.caption("Ana ekranda 'Gönderim Kontrol Dökümanı oluştur' kutucuğunu işaretleyin.")

    mod = st.session_state.get("mod")

# ---------------------------------------------------------------------------
# BAŞLIK
# ---------------------------------------------------------------------------
st.title("🚚 Sefer Aktarım & Atık Gönderim")

mod_gosterge = {
    "bosaltma": "🟢 **Boşaltma Modu** — PDF sefer bildirimleri işleniyor",
    "gonderim": "🔵 **Gönderim Modu** — Export XLS (atık gönderimleri) işleniyor",
    None: "⚪ Dosya yükleyin — mod otomatik belirlenecek",
}
st.caption(mod_gosterge.get(mod, ""))

# ---------------------------------------------------------------------------
# 1) ANA DOSYA YÜKLEME ALANI
# ---------------------------------------------------------------------------
col_excel, col_kaynak = st.columns(2)

with col_excel:
    st.subheader("1️⃣ Excel Taşıma Kontrol Listesi")
    excel_modu = st.radio(
        "Mod seçin",
        options=["yeni", "guncelle"],
        format_func=lambda v: "📄 Yeni Liste Oluştur" if v == "yeni" else "📂 Mevcut Listeyi Güncelle",
        horizontal=True,
        key="excel_modu",
    )
    if excel_modu == "guncelle":
        excel_dosya = st.file_uploader(
            "Mevcut Excel dosyanızı yükleyin (.xlsx)",
            type=["xlsx"],
            help="Veriler bu dosyaya tarih sırasına göre eklenir.",
        )
    else:
        excel_dosya = None
        st.info("📋 Program yerleşik boş şablonu kullanacak.")

with col_kaynak:
    st.subheader("2️⃣ Kaynak Dosyalar")
    st.caption("**PDF** → Boşaltma modülü  |  **XLS/XLSX** → Gönderim modülü")
    pdf_dosyalari = st.file_uploader(
        "Sefer bildirimi PDF'leri (boşaltma için)",
        type=["pdf"],
        accept_multiple_files=True,
        key="pdf_uploader",
    )
    export_dosya = st.file_uploader(
        "Export XLS/XLSX dosyası (gönderim için)",
        type=["xls", "xlsx"],
        key="export_uploader",
    )

# Mod algılama
yeni_mod = None
if pdf_dosyalari:
    yeni_mod = "bosaltma"
elif export_dosya:
    yeni_mod = "gonderim"

if yeni_mod and yeni_mod != st.session_state.mod:
    st.session_state.mod = yeni_mod
    st.rerun()

mod = st.session_state.get("mod")

st.divider()

# ---------------------------------------------------------------------------
# BOSALTMA MODÜLÜ — ayarlar
# ---------------------------------------------------------------------------
bosaltma_sonuc = None
gonderim_sonuc = None

if mod == "bosaltma":
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
                st.success("✅ Aktif — her sefer için Kontrol Dökümanı (PDF) üretilecek.")
            elif not _DOCX_DESTEGI:
                st.warning("python-docx kurulu değil.")
            else:
                st.caption("İşaretlerseniz aşağıda plaka başına tarih bilgileri istenecek.")

    plaka_ek_tarihler: dict = {}
    plaka_muayene_tarihleri: dict = {}

    if docx_uret and _DOCX_DESTEGI and pdf_dosyalari:
        st.divider()
        st.subheader("🗓️ Araç Muayene ve Geçerlilik Tarihleri")
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

        for i in range(0, len(gecici_plakalar), 2):
            cift = gecici_plakalar[i:i + 2]
            grid_cols = st.columns(2)
            for col, (plaka, pdf_adi) in zip(grid_cols, cift):
                with col:
                    with st.container(border=True):
                        st.markdown(f"**🚛 Plaka: {plaka}**")
                        st.caption(pdf_adi)
                        ara = st.text_input("Ara Muayene Tarihi", key=f"ara_{plaka}", placeholder="GG.AA.YYYY")
                        yangin = st.text_input("Yangın Tüpü Geçerlilik", key=f"yangin_{plaka}", placeholder="GG.AA.YYYY")
                        tmfb = st.text_input("TMFB Geçerlilik", key=f"tmfb_{plaka}", placeholder="GG.AA.YYYY")
                        periyodik = st.text_input("Periyodik Muayene", key=f"periyodik_{plaka}", placeholder="GG.AA.YYYY")
                        adr = st.text_input("ADR Uygunluk Geçerlilik", key=f"adr_{plaka}", placeholder="GG.AA.YYYY")
                    if ara:
                        plaka_muayene_tarihleri[plaka] = ara
                    ek = {}
                    if yangin: ek["yangin_tup"] = yangin
                    if tmfb: ek["tmfb"] = tmfb
                    if adr: ek["adr_uygunluk"] = adr
                    if periyodik: ek["periyodik_muayene"] = periyodik
                    if ek:
                        plaka_ek_tarihler[plaka] = ek

    st.divider()

# ---------------------------------------------------------------------------
# GÖNDERIM MODÜLÜ — ayarlar
# ---------------------------------------------------------------------------
elif mod == "gonderim":
    col_g1, col_g2 = st.columns([3, 2])
    with col_g1:
        st.subheader("3️⃣ Export Önizleme")
        if export_dosya:
            try:
                tmp_export = CALISMA_KLASORU / export_dosya.name
                tmp_export.write_bytes(export_dosya.getvalue())
                gonderimler_on, uyarilar_on = export_oku(tmp_export)
                st.success(f"✅ {len(gonderimler_on)} grup okundu")
                if uyarilar_on:
                    with st.expander(f"⚠️ {len(uyarilar_on)} uyarı"):
                        for u in uyarilar_on:
                            st.caption(u)
                import pandas as pd
                preview_data = [{
                    "Tarih": g.tarih_str, "Plaka": g.plaka,
                    "Taşıyıcı": g.tasiyici[:25],
                    "Taşıma No": g.tasima_nolari_str[:25],
                    "Miktar (kg)": f"{g.miktar_kg:.0f}",
                    "Muafiyet": g.muafiyet.split('\n')[0],
                } for g in gonderimler_on]
                st.dataframe(pd.DataFrame(preview_data), use_container_width=True, hide_index=True)
            except Exception as e:
                st.error(f"Export okunamadı: {e}")

    with col_g2:
        st.subheader("4️⃣ Taşıma Türü")
        st.info("Atık gönderimleri her zaman **ADR-AMBALAJLI** olarak işlenir.")

        st.subheader("5️⃣ Gönderim Kontrol Dökümanı")
        with st.container(border=True, key="gonderim_ana_karti"):
            gonderim_docx_uret = st.checkbox(
                "**📋 Her grup için Gönderim Kontrol Dökümanı oluştur**",
                value=False,
                key="gonderim_docx_uret",
            )
            if gonderim_docx_uret:
                st.success("✅ Aktif — her Tarih+Plaka+Taşıyıcı grubu için PDF üretilecek.")
            else:
                st.caption("İşaretlerseniz, sol sidebar **4️⃣** bölümünden gönderen ve şoför bilgilerini doldurun.")

    st.divider()

# ---------------------------------------------------------------------------
# AKTAR BUTONU
# ---------------------------------------------------------------------------
calistir = st.button("▶️ Aktar", type="primary", use_container_width=True)

if calistir:
    # ----- BOŞALTMA -----
    if mod == "bosaltma":
        excel_modu_val = st.session_state.get("excel_modu", "guncelle")
        if excel_modu_val == "guncelle" and not excel_dosya:
            st.error("Lütfen bir Excel dosyası yükleyin.")
        elif not pdf_dosyalari:
            st.error("Lütfen en az bir PDF dosyası yükleyin.")
        else:
            with st.spinner("İşleniyor…"):
                excel_yolu = CALISMA_KLASORU / "girdi.xlsx"
                if excel_modu_val == "yeni":
                    sablon = Path(__file__).parent / "bos_sablon.xlsx"
                    excel_yolu.write_bytes(sablon.read_bytes())
                    cikti_excel = CALISMA_KLASORU / "Taşıma_Kontrol_Listesi.xlsx"
                else:
                    excel_yolu.write_bytes(excel_dosya.getvalue())
                    cikti_excel = CALISMA_KLASORU / f"{Path(excel_dosya.name).stem}_guncel.xlsx"

                pdf_yollari = []
                for pf in pdf_dosyalari:
                    p = CALISMA_KLASORU / pf.name
                    p.write_bytes(pf.getvalue())
                    pdf_yollari.append(p)

                log = []
                bosaltma_sonuc = process_pdfs(
                    excel_path=excel_yolu,
                    pdf_paths=pdf_yollari,
                    output_path=cikti_excel,
                    log_cb=log.append,
                    muayene_tarihleri=plaka_muayene_tarihleri,
                    tasima_turu=tasima_turu,
                    ek_tarihler=plaka_ek_tarihler,
                    docx_uret=docx_uret,
                    docx_cikti_klasor=CALISMA_KLASORU / "kontrol_dokumanlari",
                    bosaltan_adi=st.session_state.get("bosaltan_adi", ""),
                    sofor_adi=st.session_state.get("sofor_adi", ""),
                    docx_pdf_donustur=True,
                    logo_bytes=st.session_state.logo_bytes,
                )
                bosaltma_sonuc["_excel_yolu"] = cikti_excel
                bosaltma_sonuc["_log"] = log
                st.session_state.sonuc = {"tip": "bosaltma", "veri": bosaltma_sonuc}

    # ----- GÖNDERİM -----
    elif mod == "gonderim":
        if not export_dosya:
            st.error("Lütfen bir Export XLS dosyası yükleyin.")
        else:
            excel_modu_val = st.session_state.get("excel_modu", "yeni")
            with st.spinner("İşleniyor…"):
                # Export oku
                tmp_export = CALISMA_KLASORU / export_dosya.name
                tmp_export.write_bytes(export_dosya.getvalue())
                try:
                    gonderimler, uyarilar = export_oku(tmp_export)
                except Exception as e:
                    st.error(f"Export okunamadı: {e}")
                    st.stop()

                # Excel şablonu
                excel_yolu = CALISMA_KLASORU / "girdi_gonderim.xlsx"
                if excel_modu_val == "yeni":
                    sablon = Path(__file__).parent / "bos_gonderim_sablon.xlsx"
                    excel_yolu.write_bytes(sablon.read_bytes())
                    cikti_excel = CALISMA_KLASORU / "Atik_Gonderim_Listesi.xlsx"
                else:
                    if not excel_dosya:
                        st.error("'Mevcut Listeyi Güncelle' modunda Excel dosyası yükleyin.")
                        st.stop()
                    excel_yolu.write_bytes(excel_dosya.getvalue())
                    cikti_excel = CALISMA_KLASORU / f"{Path(excel_dosya.name).stem}_guncel.xlsx"

                log = []
                gonderici_firma_val = st.session_state.get("gonderici_firma", "")
                gonderici_adi_val = st.session_state.get("gonderici_adi", "")
                sofor_adi_g_val = st.session_state.get("sofor_adi_g", "")

                excel_sonuc = process_export(
                    sablon_excel_path=excel_yolu,
                    gonderimler=gonderimler,
                    cikti_path=cikti_excel,
                    gonderici_adi=gonderici_adi_val,
                    logo_bytes=st.session_state.logo_bytes,
                    log_cb=log.append,
                )

                # Kontrol dökümanları
                uretilen_pdf = []
                if st.session_state.get("gonderim_docx_uret") and _DOCX_DESTEGI:
                    sablon_docx = Path(__file__).parent / "Gonderim_Kontrol_Sablonu.docx"
                    pdf_klasor = CALISMA_KLASORU / "gonderim_kontrol"
                    pdf_klasor.mkdir(exist_ok=True)
                    for g in gonderimler:
                        if g.tasima_nolari_str not in [u["tasima_no"] for u in uretilen_pdf]:
                            # Sadece yeni eklenenler için üret
                            try:
                                guvenli_firma = re.sub(r'[^\w\s]', '', gonderici_firma_val)[:20].strip().replace(' ', '_')
                                dosya_adi = f"Gonderim_{g.dosya_adi_parcasi}_{guvenli_firma}.docx"
                                docx_yolu = pdf_klasor / dosya_adi
                                gonderim_dokumani_olustur(
                                    sablon_path=sablon_docx,
                                    cikti_path=docx_yolu,
                                    tarih=g.tarih_str,
                                    gonderici_firma=gonderici_firma_val,
                                    tasiyici_firma=g.tasiyici,
                                    plaka=g.plaka,
                                    tasima_evragi_no=g.tasima_nolari_str,
                                    gonderici_adi=gonderici_adi_val,
                                    sofor_adi=sofor_adi_g_val,
                                    logo_bytes=st.session_state.logo_bytes,
                                )
                                pdf_yolu = docx_to_pdf(docx_yolu, pdf_klasor)
                                uretilen_pdf.append({
                                    "tasima_no": g.tasima_nolari_str,
                                    "tarih": g.tarih_str,
                                    "plaka": g.plaka,
                                    "pdf": pdf_yolu,
                                    "dosya_adi": (pdf_yolu.name if pdf_yolu else dosya_adi),
                                })
                                log.append(f"  📋 {dosya_adi} oluşturuldu")
                            except Exception as exc:
                                log.append(f"  ⚠ Kontrol dökümanı hatası ({g.plaka}): {exc}")

                excel_sonuc["_excel_yolu"] = cikti_excel
                excel_sonuc["_log"] = log
                excel_sonuc["_uretilen_pdf"] = uretilen_pdf
                excel_sonuc["_uyarilar"] = uyarilar
                st.session_state.sonuc = {"tip": "gonderim", "veri": excel_sonuc}

    else:
        st.warning("Lütfen önce PDF veya Export XLS dosyası yükleyin.")

# ---------------------------------------------------------------------------
# SONUÇLAR
# ---------------------------------------------------------------------------
if st.session_state.sonuc:
    s = st.session_state.sonuc
    tip = s["tip"]
    r = s["veri"]

    st.divider()
    st.subheader("✅ Sonuç")

    c1, c2, c3 = st.columns(3)
    c1.metric("Yeni Eklenen", r["eklenen"])
    c2.metric("Atlanan", r["atlanan"])
    c3.metric("Toplam", r["toplam"])

    with st.expander("İşlem Günlüğü", expanded=False):
        st.code("\n".join(r.get("_log", [])), language=None)

    # Uyarılar (gönderim modülü)
    if tip == "gonderim" and r.get("_uyarilar"):
        with st.expander(f"⚠️ {len(r['_uyarilar'])} Uyarı"):
            for u in r["_uyarilar"]:
                st.caption(u)

    col_exc, col_dok = st.columns(2)

    # Excel indirme
    if r["eklenen"] > 0:
        with col_exc:
            label = "📊 Excel Taşıma Kontrol Listesi" if tip == "bosaltma" else "📊 Excel Atık Gönderim Listesi"
            st.markdown(f"#### {label}")
            excel_path: Path = r["_excel_yolu"]
            st.download_button(
                "⬇️ Excel'i İndir",
                data=excel_path.read_bytes(),
                file_name=excel_path.name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

    # Kontrol dökümanları indirme
    with col_dok:
        if tip == "bosaltma":
            uretilen = r.get("uretilen_dosyalar") or []
            if uretilen:
                st.markdown("#### 📋 Boşaltma Kontrol Dökümanları")
                for dosya in uretilen:
                    sn = dosya["sefer_no"]
                    pdf_yolu: Path = dosya["pdf"]
                    docx_yolu: Path = dosya["docx"]
                    st.markdown(f"**Sefer: {sn}**")
                    dc1, dc2 = st.columns(2)
                    with dc1:
                        if pdf_yolu and pdf_yolu.is_file():
                            st.download_button(
                                "⬇️ PDF", data=pdf_yolu.read_bytes(),
                                file_name=pdf_yolu.name, mime="application/pdf",
                                key=f"pdf_{sn}", use_container_width=True,
                            )
                    with dc2:
                        if docx_yolu and docx_yolu.is_file():
                            st.download_button(
                                "⬇️ Word", data=docx_yolu.read_bytes(),
                                file_name=docx_yolu.name,
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                key=f"docx_{sn}", use_container_width=True,
                            )

        elif tip == "gonderim":
            uretilen_pdf = r.get("_uretilen_pdf") or []
            if uretilen_pdf:
                st.markdown("#### 📋 Gönderim Kontrol Dökümanları")
                for dok in uretilen_pdf:
                    st.markdown(f"**{dok['tarih']} / {dok['plaka']}**")
                    pdf_p: Path = dok["pdf"]
                    if pdf_p and Path(pdf_p).is_file():
                        st.download_button(
                            f"⬇️ PDF İndir",
                            data=Path(pdf_p).read_bytes(),
                            file_name=dok["dosya_adi"].replace(".docx", ".pdf"),
                            mime="application/pdf",
                            key=f"gpdf_{dok['tasima_no']}",
                            use_container_width=True,
                        )
                    else:
                        st.caption("PDF dönüşümü yapılamadı")
