#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Export XLS (atık gönderimleri) işleme modülü.

Yükleme Noktası bazlı UETS export dosyasını okur, satırları işler ve
Taşımacılık Bilgi Listesi Excel çıktısı + Gönderim Kontrol Formu (PDF) için
hazır veri yapısına dönüştürür.
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Optional

import pandas as pd

# ---------------------------------------------------------------------------
# Atık kodu → ADR bilgileri tablosu
# ---------------------------------------------------------------------------
ATIK_ADR: dict[str, dict] = {
    "04 02 19*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "07 01 03*": {"un_no": "1993", "sinif": "3", "pg": "II", "tasimaKategorisi": 2, "sevkiyat_adi": "ALEVLENİR SIVI"},
    "07 05 08*": {"un_no": "3249", "sinif": "6.2", "pg": "II", "tasimaKategorisi": 2, "sevkiyat_adi": "TIBBİ ATIK"},
    "08 01 11*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "08 01 17*": {"un_no": "3175", "sinif": "4.1", "pg": "II", "tasimaKategorisi": 2, "sevkiyat_adi": "ALEVLENİR SIVI İÇEREN KATI MADDE"},
    "08 03 17*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "08 04 11*": {"un_no": "1133", "sinif": "3", "pg": "II", "tasimaKategorisi": 1, "sevkiyat_adi": "YAPIŞTIRICILAR, B.B.B."},
    "09 01 04*": {"un_no": "3082", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE"},
    "11 01 09*": {"un_no": "3264", "sinif": "8", "pg": "II", "tasimaKategorisi": 2, "sevkiyat_adi": "AŞINDIRICI SIVI, ASİDİK"},
    "12 01 09*": {"un_no": "3082", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, SIVI, B.B.B."},
    "13 01 10*": {"un_no": "3082", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, SIVI, B.B.B."},
    "13 01 13*": {"un_no": "3082", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, SIVI, B.B.B."},
    "13 02 05*": {"un_no": "3082", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, SIVI, B.B.B."},
    "13 05 07*": {"un_no": "3082", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, SIVI, B.B.B."},
    "14 06 02*": {"un_no": "1993", "sinif": "3", "pg": "II", "tasimaKategorisi": 2, "sevkiyat_adi": "ALEVLENİR SIVI, B.B.B."},
    "14 06 03*": {"un_no": "1993", "sinif": "3", "pg": "II", "tasimaKategorisi": 2, "sevkiyat_adi": "ALEVLENİR SIVI, B.B.B."},
    "15 01 10*": {"un_no": "3509", "sinif": "9", "pg": "-", "tasimaKategorisi": 4, "sevkiyat_adi": "BOŞ, TEMİZLENMEMİŞ AMBALAJ"},
    "15 02 02*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI, B.B.B."},
    "16 01 07*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI, B.B.B."},
    "16 01 13*": {"un_no": "3082", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, SIVI"},
    "16 02 13*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI, B.B.B."},
    "16 03 05*": {"un_no": "3082", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, SIVI"},
    "16 05 06*": {"un_no": "2811", "sinif": "6.1", "pg": "II", "tasimaKategorisi": 2, "sevkiyat_adi": "TOKSİK KATI, ORGANİK"},
    "16 06 01*": {"un_no": "2794", "sinif": "8", "pg": "-", "tasimaKategorisi": 3, "sevkiyat_adi": "AKÜLER, ASİT İÇEREN"},
    "16 06 02*": {"un_no": "2795", "sinif": "8", "pg": "-", "tasimaKategorisi": 3, "sevkiyat_adi": "AKÜLER, ALKALİ"},
    "17 02 04*": {"un_no": "3175", "sinif": "4.1", "pg": "II", "tasimaKategorisi": 2, "sevkiyat_adi": "ALEVLENİR KATI"},
    "17 04 09*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "18 01 03*": {"un_no": "3291", "sinif": "6.2", "pg": "II", "tasimaKategorisi": 2, "sevkiyat_adi": "KLİNİK ATIK, B.B.B."},
    "19 08 13*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "19 12 11*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE"},
    "20 01 21*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI, B.B.B."},
    "20 01 26*": {"un_no": "3082", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, SIVI, B.B.B."},
    "20 01 27*": {"un_no": "3082", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, SIVI, B.B.B."},
    "20 01 33*": {"un_no": "3480", "sinif": "9", "pg": "-", "tasimaKategorisi": 2, "sevkiyat_adi": "LİTYUM PİLLER"},
}

# Noktalı gösterim olmadan da eşleştirme yap (ör. "080111" → "08 01 11*")
def _normalize_atik_kodu(kod: str) -> str:
    """'080111' → '08 01 11*' formatına çevirir."""
    kod = str(kod).strip()
    # Zaten standart formattaysa
    if re.match(r'^\d{2}\s\d{2}\s\d{2}\*?$', kod):
        return kod if kod.endswith('*') else kod + '*'
    # 6 haneli sıkışık format
    if re.match(r'^\d{6}$', kod):
        return f"{kod[:2]} {kod[2:4]} {kod[4:6]}*"
    return kod


def _atik_adr_bul(atik_kodu: str) -> Optional[dict]:
    """Atık kodundan ADR bilgilerini döndürür. Bulunamazsa None."""
    norm = _normalize_atik_kodu(atik_kodu)
    return ATIK_ADR.get(norm)


# ---------------------------------------------------------------------------
# Firma adı temizleme
# ---------------------------------------------------------------------------
def _temizle_firma(ham: str) -> str:
    """
    'ATA-34-243 - EKOLOJİK ENERJİ ANONİM ŞİRKETİ' → 'EKOLOJİK ENERJİ ANONİM ŞİRKETİ'
    '111838 - EKOLOJİK ENERJİ A.Ş. ÇORLU ŞUBESİ (ÇKN: 227661262)' → 'EKOLOJİK ENERJİ A.Ş. ÇORLU ŞUBESİ'
    """
    if not ham or str(ham).strip() == 'nan':
        return ''
    ham = str(ham).strip()
    # Baştaki kod kısmını at: "ATA-XX-XXX - " veya "XXXXXX - "
    ham = re.sub(r'^[A-Z0-9\-]+\s*-\s*', '', ham)
    # Sondaki ÇKN parantezi at: "(ÇKN: ...)"
    ham = re.sub(r'\s*\(ÇKN\s*:.*?\)\s*$', '', ham, flags=re.IGNORECASE)
    return ham.strip()


# ---------------------------------------------------------------------------
# Miktar ayrıştırma
# ---------------------------------------------------------------------------
def _parse_miktar(deger) -> float:
    """'6.620' (Türkçe binlik), '320', 1040 → float kg olarak döner."""
    if deger is None or str(deger).strip() in ('', 'nan'):
        return 0.0
    s = str(deger).strip()
    # Türkçe binlik ayracı (nokta) → virgülsüz int/float
    # '6.620' → 6620; '1.040' → 1040; '320' → 320; '1,5' → 1.5
    if re.match(r'^\d{1,3}\.\d{3}$', s):
        return float(s.replace('.', ''))
    return float(s.replace(',', '.'))


# ---------------------------------------------------------------------------
# ADR 1.1.3.6 Muafiyet Hesaplama (Puan Sistemi)
# ---------------------------------------------------------------------------
TC_PUANLARI = {1: 50, 2: 3, 3: 1, 4: 0}
MAX_PUAN = 1000


def _1136_puan_hesapla(satirlar: list[dict]) -> tuple[float, bool]:
    """
    ADR 1.1.3.6 miktar muafiyeti — puan sistemi.
    Her kalem kendi miktarı ve taşıma kategorisiyle ayrı ayrı değerlendirilir.
    TC 4 → 0 puan (sınırsız), TC 1/2/3 → miktar × puan çarpanı.
    Toplam ≤ 1000 → muafiyet EVET, > 1000 → HAYIR.
    """
    toplam = 0.0
    for s in satirlar:
        tc = s.get('tasima_kategorisi', 3)
        puan_carpan = TC_PUANLARI.get(tc, 1)
        toplam += s['miktar'] * puan_carpan
    return toplam, toplam <= MAX_PUAN


def _muafiyet_metni(satirlar: list[dict]) -> str:
    """Her grup için ADR 1.1.3.6 puan tabanlı muafiyet metni üretir."""
    toplam_puan, muaf = _1136_puan_hesapla(satirlar)
    kapsam = 'EVET' if muaf else 'HAYIR'
    return f"{kapsam}\n-ADR 1.1.3.6\nPUAN: {toplam_puan:.0f}/1000"


# ---------------------------------------------------------------------------
# Veri modeli
# ---------------------------------------------------------------------------
@dataclass
class AtikGonderim:
    """Tek bir grubun (Tarih+Plaka+Taşıyıcı) işlenmiş verisi."""
    tarih: datetime
    tasiyici: str                    # Temizlenmiş taşıyıcı firma adı
    plaka: str
    alici: str                       # Temizlenmiş alıcı firma adı
    # Aşağıdakiler birden fazla atık kodu olabilir (virgülle birleştirilir)
    atik_kodlari: list[str] = field(default_factory=list)   # ['08 01 11*', '15 01 10*']
    tasima_nolari: list[str] = field(default_factory=list)  # ['E8433002', 'E8432994']
    un_nolar: list[str] = field(default_factory=list)
    miktar_kg: float = 0.0
    tasima_kategorisi: int = 3       # En kısıtlı kategori

    # Hesaplanan alanlar
    @property
    def tarih_str(self) -> str:
        return self.tarih.strftime('%d.%m.%Y') if self.tarih else ''

    @property
    def atik_kodlari_str(self) -> str:
        return ', '.join(self.atik_kodlari)

    @property
    def tasima_nolari_str(self) -> str:
        return ', '.join(self.tasima_nolari)

    @property
    def un_nolar_str(self) -> str:
        unique = list(dict.fromkeys(self.un_nolar))  # sıra koruyarak deduplicate
        return ', '.join(f'UN {u}' for u in unique)

    # Ham satirlar (puan hesabı için saklanır)
    _satirlar: list = field(default_factory=list)

    @property
    def muafiyet(self) -> str:
        if self._satirlar:
            return _muafiyet_metni(self._satirlar)
        # Fallback: eski mantık
        toplam_puan = self.miktar_kg * TC_PUANLARI.get(self.tasima_kategorisi, 1)
        kapsam = 'EVET' if toplam_puan <= MAX_PUAN else 'HAYIR'
        return f"{kapsam}\n-ADR 1.1.3.6\nPUAN: {toplam_puan:.0f}/1000"

    @property
    def dosya_adi_parcasi(self) -> str:
        """Gönderim Kontrol Formu dosya adı için güvenli parça."""
        tarih = self.tarih.strftime('%Y%m%d') if self.tarih else 'bilinmeyen'
        plaka = re.sub(r'[^\w]', '', self.plaka)
        return f'{tarih}_{plaka}'


# ---------------------------------------------------------------------------
# Ana işleme fonksiyonu
# ---------------------------------------------------------------------------
def export_oku(dosya_path: str | Path) -> tuple[list[AtikGonderim], list[str]]:
    """
    Export XLS/XLSX dosyasını okur, gruplar ve işler.

    Returns:
        (gonderimler, uyarilar)
        gonderimler: her Tarih+Plaka+Taşıyıcı grubu için bir AtikGonderim nesnesi
        uyarilar: bilinmeyen atık kodları veya eksik veri uyarıları
    """
    dosya_path = Path(dosya_path)
    engine = 'xlrd' if dosya_path.suffix.lower() == '.xls' else 'openpyxl'

    try:
        df = pd.read_excel(dosya_path, engine=engine, header=None)
    except Exception as e:
        raise ValueError(f"Dosya okunamadı: {e}")

    # Başlık satırını bul (Taşıma Numarası kolonunu ara)
    header_row = None
    for i, row in df.iterrows():
        vals = [str(v).strip() for v in row if str(v).strip() != 'nan']
        if 'Taşıma Numarası' in vals or 'Taşıma No' in vals:
            header_row = i
            break

    if header_row is None:
        raise ValueError("Export dosyasında başlık satırı bulunamadı.")

    df.columns = df.iloc[header_row]
    df = df.iloc[header_row + 1:].reset_index(drop=True)
    df = df.dropna(how='all')

    # Sütun adlarını normalize et
    col_map = {}
    for col in df.columns:
        col_str = str(col).strip()
        if 'Taşıma Numarası' in col_str or 'Taşıma No' in col_str:
            col_map['tasima_no'] = col
        elif 'Atık' in col_str and 'kodu' not in col_str.lower():
            col_map['atik_kodu'] = col
        elif 'Miktar' in col_str:
            col_map['miktar'] = col
        elif 'Taşıyıcı' in col_str:
            col_map['tasiyici'] = col
        elif 'Plaka' in col_str:
            col_map['plaka'] = col
        elif 'Alıcı' in col_str:
            col_map['alici'] = col
        elif 'Boşaltma Zamanı' in col_str or 'Tarih' in col_str:
            col_map['tarih'] = col

    gerekli = ['tasima_no', 'atik_kodu', 'miktar', 'tasiyici', 'plaka', 'alici', 'tarih']
    eksik = [k for k in gerekli if k not in col_map]
    if eksik:
        raise ValueError(f"Export dosyasında sütunlar bulunamadı: {eksik}")

    uyarilar = []
    # Gruplama: (tarih_gun, plaka, tasiyici) → satırlar
    gruplar: dict[tuple, list[dict]] = {}

    for _, row in df.iterrows():
        tasima_no = str(row[col_map['tasima_no']]).strip()
        if not tasima_no or tasima_no == 'nan':
            continue

        atik_ham = str(row[col_map['atik_kodu']]).strip()
        atik_norm = _normalize_atik_kodu(atik_ham)
        miktar = _parse_miktar(row[col_map['miktar']])
        tasiyici = _temizle_firma(row[col_map['tasiyici']])
        plaka = str(row[col_map['plaka']]).strip()
        alici = _temizle_firma(row[col_map['alici']])

        tarih_ham = row[col_map['tarih']]
        if isinstance(tarih_ham, datetime):
            tarih = tarih_ham
        else:
            try:
                tarih = pd.to_datetime(tarih_ham)
            except Exception:
                tarih = None

        adr = _atik_adr_bul(atik_ham)
        if adr is None:
            uyarilar.append(f"Bilinmeyen atık kodu: {atik_ham} (Taşıma No: {tasima_no})")

        tarih_gun = tarih.date() if tarih else None
        grup_key = (tarih_gun, plaka.upper(), tasiyici)

        if grup_key not in gruplar:
            gruplar[grup_key] = []

        gruplar[grup_key].append({
            'tasima_no': tasima_no,
            'atik_norm': atik_norm,
            'un_no': adr['un_no'] if adr else '',
            'tasima_kategorisi': adr['tasimaKategorisi'] if adr else 3,
            'miktar': miktar,
            'tarih': tarih,
            'tasiyici': tasiyici,
            'plaka': plaka,
            'alici': alici,
        })

    # Her grubu AtikGonderim nesnesine dönüştür
    gonderimler: list[AtikGonderim] = []
    for grup_key, satirlar in gruplar.items():
        ilk = satirlar[0]
        # Toplam miktar, en kısıtlı kategori (en küçük sayı = en kısıtlı)
        toplam_miktar = sum(s['miktar'] for s in satirlar)
        min_kategori = min(s['tasima_kategorisi'] for s in satirlar)
        # Sıra koruyarak tekrarsız listeler
        seen_atik, seen_no, seen_un = set(), set(), []
        atik_listesi, no_listesi, un_listesi = [], [], []
        for s in satirlar:
            if s['atik_norm'] not in seen_atik:
                seen_atik.add(s['atik_norm'])
                atik_listesi.append(s['atik_norm'])
            if s['tasima_no'] not in seen_no:
                seen_no.add(s['tasima_no'])
                no_listesi.append(s['tasima_no'])
            if s['un_no'] and s['un_no'] not in seen_un:
                seen_un.append(s['un_no'])
                un_listesi.append(s['un_no'])

        gonderimler.append(AtikGonderim(
            tarih=ilk['tarih'],
            tasiyici=ilk['tasiyici'],
            plaka=ilk['plaka'],
            alici=ilk['alici'],
            atik_kodlari=atik_listesi,
            tasima_nolari=no_listesi,
            un_nolar=un_listesi,
            miktar_kg=toplam_miktar,
            tasima_kategorisi=min_kategori,
            _satirlar=satirlar,
        ))

    # Tarihe göre sırala
    gonderimler.sort(key=lambda g: g.tarih or datetime.min)
    return gonderimler, uyarilar
