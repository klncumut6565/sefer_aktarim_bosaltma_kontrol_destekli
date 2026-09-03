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
# ================================================================
# Genişletilmiş Atık Kodu → ADR Bilgileri Tablosu (2026-09-02)
# Kaynak: Tüm_Atık_Kodları.xlsx — 408 tehlikeli atık kodu
# ================================================================
# Not: Eski dictionary (40 kod) yerine yeni dictionary (408 kod) konmuştur
# Tüm atık kodları otomatik olarak UN numarası ve ADR bilgileriyle
# eşleştirilmiştir. Bilinmeyen kodlar için _atik_adr_bul() None döndürür.
# ================================================================
# ================================================================
# Genişletilmiş Atık Kodu → ADR Bilgileri Tablosu (2026-09-02)
# Kaynak: Tüm_Atık_Kodları.xlsx — 408 tehlikeli atık kodu
# ================================================================
ATIK_ADR: dict[str, dict] = {
    "10 01 04*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "10 01 09*": {"un_no": "2796", "sinif": "9", "pg": "III", "tasimaKategorisi": 2, "sevkiyat_adi": "ADR SINIFLANDI: 9"},
    "10 01 13*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "10 01 14*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "10 01 16*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "10 01 18*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "10 01 20*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "10 01 22*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "10 02 07*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "10 02 11*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "10 02 13*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "10 03 04*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "10 03 08*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "10 03 09*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "10 03 15*": {"un_no": "3543", "sinif": "9", "pg": "III", "tasimaKategorisi": 4, "sevkiyat_adi": "ADR SINIFLANDI: 9"},
    "10 03 17*": {"un_no": "3082", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, SIVI"},
    "10 03 19*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "10 03 21*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "10 03 23*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "10 03 25*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "10 03 27*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "10 03 29*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "10 04 01*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "10 04 02*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "10 04 03*": {"un_no": "1573", "sinif": "9", "pg": "III", "tasimaKategorisi": 2, "sevkiyat_adi": "ADR SINIFLANDI: 9"},
    "10 04 04*": {"un_no": "1562", "sinif": "3", "pg": "III", "tasimaKategorisi": 2, "sevkiyat_adi": "ADR SINIFLANDI: 3"},
    "10 04 05*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "10 04 06*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "10 04 07*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "10 04 09*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "10 05 03*": {"un_no": "1562", "sinif": "3", "pg": "III", "tasimaKategorisi": 2, "sevkiyat_adi": "ADR SINIFLANDI: 3"},
    "10 05 05*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "10 05 06*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "10 05 08*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "10 05 10*": {"un_no": "2813", "sinif": "9", "pg": "III", "tasimaKategorisi": 0, "sevkiyat_adi": "ADR SINIFLANDI: 9"},
    "10 06 03*": {"un_no": "1562", "sinif": "3", "pg": "III", "tasimaKategorisi": 2, "sevkiyat_adi": "ADR SINIFLANDI: 3"},
    "10 06 06*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "10 06 07*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "10 06 09*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "10 07 07*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "10 08 08*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "10 08 10*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "10 08 12*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "10 08 15*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "10 08 17*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "10 08 19*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "10 09 05*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "10 09 07*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "10 09 09*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "10 09 11*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "10 09 13*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "10 09 15*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "10 10 05*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "10 10 07*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "10 10 09*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "10 10 11*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "10 10 13*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "10 10 15*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "10 11 09*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "10 11 11*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "10 11 13*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "10 11 15*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "10 11 17*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "10 11 19*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "10 12 09*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "10 12 11*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "10 13 09*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "10 13 12*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "10 14 01*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "10 30 4*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "10 30 5*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "10 30 7*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "10 31 0*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "10 40 7*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "10 50 5*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "10 50 6*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "11 01 05*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "11 01 06*": {"un_no": "3082", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, SIVI"},
    "11 01 07*": {"un_no": "3082", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, SIVI"},
    "11 01 08*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "11 01 09*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "11 01 11*": {"un_no": "3082", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, SIVI"},
    "11 01 13*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "11 01 15*": {"un_no": "3082", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, SIVI"},
    "11 01 16*": {"un_no": "1286", "sinif": "3", "pg": "III", "tasimaKategorisi": 2, "sevkiyat_adi": "ADR SINIFLANDI: 3"},
    "11 01 98*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "11 02 02*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "11 02 05*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "11 02 07*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "11 03 01*": {"un_no": "3243", "sinif": "9", "pg": "III", "tasimaKategorisi": 2, "sevkiyat_adi": "ADR SINIFLANDI: 9"},
    "11 03 02*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "11 05 03*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "11 05 04*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "12 01 06*": {"un_no": "3082", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, SIVI"},
    "12 01 07*": {"un_no": "3082", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, SIVI"},
    "12 01 08*": {"un_no": "3082", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, SIVI"},
    "12 01 09*": {"un_no": "3082", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, SIVI"},
    "12 01 10*": {"un_no": "3082", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, SIVI"},
    "12 01 12*": {"un_no": "1223", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ADR SINIFLANDI: 9"},
    "12 01 14*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "12 01 16*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "12 01 18*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "12 01 19*": {"un_no": "3082", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, SIVI"},
    "12 01 20*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "12 03 01*": {"un_no": "3082", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, SIVI"},
    "12 03 02*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "13 01 01*": {"un_no": "2315", "sinif": "9", "pg": "III", "tasimaKategorisi": 0, "sevkiyat_adi": "ADR SINIFLANDI: 9"},
    "13 01 04*": {"un_no": "3082", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, SIVI"},
    "13 01 05*": {"un_no": "3082", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, SIVI"},
    "13 01 09*": {"un_no": "3082", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, SIVI"},
    "13 01 10*": {"un_no": "3082", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, SIVI"},
    "13 01 11*": {"un_no": "3082", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, SIVI"},
    "13 01 12*": {"un_no": "3082", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, SIVI"},
    "13 01 13*": {"un_no": "3082", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, SIVI"},
    "13 02 04*": {"un_no": "3082", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, SIVI"},
    "13 02 05*": {"un_no": "3082", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, SIVI"},
    "13 02 06*": {"un_no": "3082", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, SIVI"},
    "13 02 07*": {"un_no": "3082", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, SIVI"},
    "13 02 08*": {"un_no": "3082", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, SIVI"},
    "13 03 01*": {"un_no": "3082", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, SIVI"},
    "13 03 06*": {"un_no": "3082", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, SIVI"},
    "13 03 07*": {"un_no": "3082", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, SIVI"},
    "13 03 08*": {"un_no": "3082", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, SIVI"},
    "13 03 09*": {"un_no": "3082", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, SIVI"},
    "13 03 10*": {"un_no": "3082", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, SIVI"},
    "13 04 01*": {"un_no": "3082", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, SIVI"},
    "13 04 02*": {"un_no": "3082", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, SIVI"},
    "13 04 03*": {"un_no": "3082", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, SIVI"},
    "13 05 01*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "13 05 02*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "13 05 03*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "13 05 06*": {"un_no": "3082", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, SIVI"},
    "13 05 07*": {"un_no": "3082", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, SIVI"},
    "13 05 08*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "13 07 01*": {"un_no": "1202", "sinif": "3", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ADR SINIFLANDI: 3"},
    "13 07 02*": {"un_no": "1203", "sinif": "9", "pg": "III", "tasimaKategorisi": 2, "sevkiyat_adi": "ADR SINIFLANDI: 9"},
    "13 07 03*": {"un_no": "1993", "sinif": "3", "pg": "II", "tasimaKategorisi": 1, "sevkiyat_adi": "ALEVLENİR SIVI"},
    "13 08 01*": {"un_no": "3082", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, SIVI"},
    "13 08 02*": {"un_no": "1202", "sinif": "3", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ADR SINIFLANDI: 3"},
    "13 08 99*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "14 06 01*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "14 06 02*": {"un_no": "3151", "sinif": "9", "pg": "III", "tasimaKategorisi": 0, "sevkiyat_adi": "ADR SINIFLANDI: 9"},
    "14 06 03*": {"un_no": "1993", "sinif": "3", "pg": "II", "tasimaKategorisi": 1, "sevkiyat_adi": "ALEVLENİR SIVI"},
    "14 06 04*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "14 06 05*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "15 01 10*": {"un_no": "3509", "sinif": "9", "pg": "-", "tasimaKategorisi": 4, "sevkiyat_adi": "BOŞ, TEMİZLENMEMİŞ AMBALAJ"},
    "15 01 11*": {"un_no": "1950", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ADR SINIFLANDI: 9"},
    "15 02 02*": {"un_no": "1373", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ADR SINIFLANDI: 9"},
    "16 01 04*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "16 01 07*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "16 01 08*": {"un_no": "3506", "sinif": "3", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ADR SINIFLANDI: 3"},
    "16 01 09*": {"un_no": "3548", "sinif": "9", "pg": "III", "tasimaKategorisi": 4, "sevkiyat_adi": "ADR SINIFLANDI: 9"},
    "16 01 10*": {"un_no": "3268", "sinif": "9", "pg": "III", "tasimaKategorisi": 4, "sevkiyat_adi": "ADR SINIFLANDI: 9"},
    "16 01 11*": {"un_no": "3363", "sinif": "9", "pg": "III", "tasimaKategorisi": 0, "sevkiyat_adi": "ADR SINIFLANDI: 9"},
    "16 01 13*": {"un_no": "3082", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, SIVI"},
    "16 01 14*": {"un_no": "3082", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, SIVI"},
    "16 01 21*": {"un_no": "3548", "sinif": "9", "pg": "III", "tasimaKategorisi": 4, "sevkiyat_adi": "ADR SINIFLANDI: 9"},
    "16 02 09*": {"un_no": "3432", "sinif": "3", "pg": "III", "tasimaKategorisi": 0, "sevkiyat_adi": "ADR SINIFLANDI: 3"},
    "16 02 10*": {"un_no": "3548", "sinif": "9", "pg": "III", "tasimaKategorisi": 4, "sevkiyat_adi": "ADR SINIFLANDI: 9"},
    "16 02 11*": {"un_no": "3548", "sinif": "9", "pg": "III", "tasimaKategorisi": 4, "sevkiyat_adi": "ADR SINIFLANDI: 9"},
    "16 02 12*": {"un_no": "3548", "sinif": "9", "pg": "III", "tasimaKategorisi": 4, "sevkiyat_adi": "ADR SINIFLANDI: 9"},
    "16 02 13*": {"un_no": "3548", "sinif": "9", "pg": "III", "tasimaKategorisi": 4, "sevkiyat_adi": "ADR SINIFLANDI: 9"},
    "16 02 15*": {"un_no": "3548", "sinif": "9", "pg": "III", "tasimaKategorisi": 4, "sevkiyat_adi": "ADR SINIFLANDI: 9"},
    "16 03 03*": {"un_no": "3178", "sinif": "9", "pg": "III", "tasimaKategorisi": 2, "sevkiyat_adi": "ADR SINIFLANDI: 9"},
    "16 03 05*": {"un_no": "3263", "sinif": "9", "pg": "III", "tasimaKategorisi": 1, "sevkiyat_adi": "ADR SINIFLANDI: 9"},
    "16 03 07*": {"un_no": "2809", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ADR SINIFLANDI: 9"},
    "16 04 01*": {"un_no": "353", "sinif": "5.1", "pg": "III", "tasimaKategorisi": 2, "sevkiyat_adi": "ADR SINIFLANDI: 5.1"},
    "16 04 02*": {"un_no": "353", "sinif": "5.1", "pg": "III", "tasimaKategorisi": 2, "sevkiyat_adi": "ADR SINIFLANDI: 5.1"},
    "16 04 03*": {"un_no": "353", "sinif": "5.1", "pg": "III", "tasimaKategorisi": 2, "sevkiyat_adi": "ADR SINIFLANDI: 5.1"},
    "16 05 04*": {"un_no": "3538", "sinif": "9", "pg": "III", "tasimaKategorisi": 4, "sevkiyat_adi": "ADR SINIFLANDI: 9"},
    "16 05 06*": {"un_no": "3082", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, SIVI"},
    "16 05 07*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "16 05 08*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "16 06 01*": {"un_no": "2794", "sinif": "8", "pg": "-", "tasimaKategorisi": 3, "sevkiyat_adi": "AKÜLER, ASİT İÇEREN"},
    "16 06 02*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "16 06 03*": {"un_no": "3506", "sinif": "3", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ADR SINIFLANDI: 3"},
    "16 06 06*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "16 07 08*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "16 07 09*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "16 08 02*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "16 08 05*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "16 08 06*": {"un_no": "3082", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, SIVI"},
    "16 08 07*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "16 09 01*": {"un_no": "1482", "sinif": "9", "pg": "III", "tasimaKategorisi": 2, "sevkiyat_adi": "ADR SINIFLANDI: 9"},
    "16 09 02*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "16 09 03*": {"un_no": "2984", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ADR SINIFLANDI: 9"},
    "16 09 04*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "16 10 01*": {"un_no": "3082", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, SIVI"},
    "16 10 03*": {"un_no": "3082", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, SIVI"},
    "16 11 01*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "16 11 03*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "16 11 05*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "17 01 06*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "17 02 04*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "17 03 01*": {"un_no": "3082", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, SIVI"},
    "17 03 03*": {"un_no": "3082", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, SIVI"},
    "17 04 09*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "17 04 10*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "17 05 03*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "17 05 05*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "17 05 07*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "17 06 01*": {"un_no": "2590", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ADR SINIFLANDI: 9"},
    "17 06 03*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "17 06 05*": {"un_no": "2590", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ADR SINIFLANDI: 9"},
    "17 08 01*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "17 09 01*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "17 09 02*": {"un_no": "3432", "sinif": "3", "pg": "III", "tasimaKategorisi": 0, "sevkiyat_adi": "ADR SINIFLANDI: 3"},
    "17 09 03*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "18 01 03*": {"un_no": "3291", "sinif": "6.2", "pg": "II", "tasimaKategorisi": 2, "sevkiyat_adi": "KLİNİK ATIK"},
    "18 01 06*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "18 01 08*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "18 01 10*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "18 02 02*": {"un_no": "3291", "sinif": "6.2", "pg": "II", "tasimaKategorisi": 2, "sevkiyat_adi": "KLİNİK ATIK"},
    "18 02 05*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "18 02 07*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "19 01 05*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "19 01 06*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "19 01 07*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "19 01 10*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "19 01 11*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "19 01 13*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "19 01 15*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "19 01 17*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "19 02 04*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "19 02 05*": {"un_no": "3082", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, SIVI"},
    "19 02 07*": {"un_no": "3082", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, SIVI"},
    "19 02 08*": {"un_no": "1993", "sinif": "3", "pg": "II", "tasimaKategorisi": 1, "sevkiyat_adi": "ALEVLENİR SIVI"},
    "19 02 09*": {"un_no": "1325", "sinif": "4.1", "pg": "III", "tasimaKategorisi": 2, "sevkiyat_adi": "ADR SINIFLANDI: 4.1"},
    "19 02 11*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "19 03 04*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "19 03 06*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "19 03 08*": {"un_no": "2025", "sinif": "6.1", "pg": "III", "tasimaKategorisi": 1, "sevkiyat_adi": "ADR SINIFLANDI: 6.1"},
    "19 04 02*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "19 04 03*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "19 07 02*": {"un_no": "3082", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, SIVI"},
    "19 08 06*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "19 08 07*": {"un_no": "3082", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, SIVI"},
    "19 08 08*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "19 08 10*": {"un_no": "3082", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, SIVI"},
    "19 08 11*": {"un_no": "1325", "sinif": "4.1", "pg": "III", "tasimaKategorisi": 2, "sevkiyat_adi": "ADR SINIFLANDI: 4.1"},
    "19 08 13*": {"un_no": "1325", "sinif": "4.1", "pg": "III", "tasimaKategorisi": 2, "sevkiyat_adi": "ADR SINIFLANDI: 4.1"},
    "19 10 03*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "19 10 05*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "19 11 01*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "19 11 02*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "19 11 03*": {"un_no": "3082", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, SIVI"},
    "19 11 04*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "19 11 05*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "19 11 07*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "19 12 06*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "19 12 11*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "19 13 01*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "19 13 03*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "19 13 05*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "19 13 07*": {"un_no": "3082", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, SIVI"},
    "20 01 13*": {"un_no": "3082", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, SIVI"},
    "20 01 14*": {"un_no": "3082", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, SIVI"},
    "20 01 15*": {"un_no": "3082", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, SIVI"},
    "20 01 17*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "20 01 19*": {"un_no": "2588", "sinif": "9", "pg": "III", "tasimaKategorisi": 1, "sevkiyat_adi": "ADR SINIFLANDI: 9"},
    "20 01 21*": {"un_no": "3506", "sinif": "3", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ADR SINIFLANDI: 3"},
    "20 01 23*": {"un_no": "3548", "sinif": "9", "pg": "III", "tasimaKategorisi": 4, "sevkiyat_adi": "ADR SINIFLANDI: 9"},
    "20 01 26*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "20 01 27*": {"un_no": "1263", "sinif": "3", "pg": "II", "tasimaKategorisi": 1, "sevkiyat_adi": "BENZIN"},
    "20 01 29*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "20 01 31*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "20 01 33*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "20 01 35*": {"un_no": "3548", "sinif": "9", "pg": "III", "tasimaKategorisi": 4, "sevkiyat_adi": "ADR SINIFLANDI: 9"},
    "20 01 37*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "20 10 8*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "30 10 4*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "30 20 1*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "30 20 2*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "30 20 3*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "30 20 4*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "30 20 5*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "40 10 3*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "40 21 4*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "40 21 6*": {"un_no": "1263", "sinif": "3", "pg": "II", "tasimaKategorisi": 1, "sevkiyat_adi": "BENZIN"},
    "40 21 9*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "50 10 2*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "50 10 3*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "50 10 4*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "50 10 5*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "50 10 6*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "50 10 7*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "50 10 8*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "50 10 9*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "50 11 1*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "50 11 2*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "50 11 5*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "50 60 1*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "50 60 3*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "50 70 1*": {"un_no": "2025", "sinif": "6.1", "pg": "III", "tasimaKategorisi": 1, "sevkiyat_adi": "ADR SINIFLANDI: 6.1"},
    "60 10 1*": {"un_no": "1830", "sinif": "9", "pg": "III", "tasimaKategorisi": 2, "sevkiyat_adi": "ADR SINIFLANDI: 9"},
    "60 10 2*": {"un_no": "1789", "sinif": "9", "pg": "III", "tasimaKategorisi": 2, "sevkiyat_adi": "ADR SINIFLANDI: 9"},
    "60 10 3*": {"un_no": "1790", "sinif": "9", "pg": "III", "tasimaKategorisi": 1, "sevkiyat_adi": "ADR SINIFLANDI: 9"},
    "60 10 4*": {"un_no": "1805", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ADR SINIFLANDI: 9"},
    "60 10 5*": {"un_no": "2031", "sinif": "9", "pg": "III", "tasimaKategorisi": 1, "sevkiyat_adi": "ADR SINIFLANDI: 9"},
    "60 10 6*": {"un_no": "3264", "sinif": "8", "pg": "II", "tasimaKategorisi": 1, "sevkiyat_adi": "AŞINDIRICI SIVI, ASİDİK"},
    "60 20 1*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "60 20 3*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "60 20 4*": {"un_no": "1813", "sinif": "9", "pg": "III", "tasimaKategorisi": 2, "sevkiyat_adi": "ADR SINIFLANDI: 9"},
    "60 20 5*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "60 31 1*": {"un_no": "1588", "sinif": "9", "pg": "III", "tasimaKategorisi": 1, "sevkiyat_adi": "ADR SINIFLANDI: 9"},
    "60 31 3*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "60 31 5*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "60 40 3*": {"un_no": "1557", "sinif": "9", "pg": "III", "tasimaKategorisi": 1, "sevkiyat_adi": "ADR SINIFLANDI: 9"},
    "60 40 4*": {"un_no": "2025", "sinif": "6.1", "pg": "III", "tasimaKategorisi": 1, "sevkiyat_adi": "ADR SINIFLANDI: 6.1"},
    "60 40 5*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "60 50 2*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "60 60 2*": {"un_no": "1350", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ADR SINIFLANDI: 9"},
    "60 70 1*": {"un_no": "2212", "sinif": "9", "pg": "III", "tasimaKategorisi": 2, "sevkiyat_adi": "ADR SINIFLANDI: 9"},
    "60 70 2*": {"un_no": "1362", "sinif": "9", "pg": "III", "tasimaKategorisi": 4, "sevkiyat_adi": "ADR SINIFLANDI: 9"},
    "60 70 3*": {"un_no": "1564", "sinif": "9", "pg": "III", "tasimaKategorisi": 2, "sevkiyat_adi": "ADR SINIFLANDI: 9"},
    "60 70 4*": {"un_no": "3082", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, SIVI"},
    "60 80 2*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "60 90 3*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "61 00 2*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "61 30 1*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "61 30 2*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "61 30 4*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "61 30 5*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "70 10 1*": {"un_no": "3082", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, SIVI"},
    "70 10 3*": {"un_no": "3082", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, SIVI"},
    "70 10 4*": {"un_no": "3082", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, SIVI"},
    "70 10 7*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "70 10 8*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "70 10 9*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "70 11 0*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "70 11 1*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "70 20 1*": {"un_no": "3082", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, SIVI"},
    "70 20 3*": {"un_no": "3082", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, SIVI"},
    "70 20 4*": {"un_no": "3082", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, SIVI"},
    "70 20 7*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "70 20 8*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "70 20 9*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "70 21 0*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "70 21 1*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "70 21 4*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "70 21 6*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "70 30 1*": {"un_no": "3082", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, SIVI"},
    "70 30 3*": {"un_no": "3082", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, SIVI"},
    "70 30 4*": {"un_no": "3082", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, SIVI"},
    "70 30 7*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "70 30 8*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "70 30 9*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "70 31 0*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "70 31 1*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "70 40 1*": {"un_no": "3082", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, SIVI"},
    "70 40 3*": {"un_no": "3082", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, SIVI"},
    "70 40 4*": {"un_no": "3082", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, SIVI"},
    "70 40 7*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "70 40 8*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "70 40 9*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "70 41 0*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "70 41 1*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "70 41 3*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "70 50 1*": {"un_no": "3082", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, SIVI"},
    "70 50 3*": {"un_no": "3082", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, SIVI"},
    "70 50 4*": {"un_no": "3082", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, SIVI"},
    "70 50 7*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "70 50 8*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "70 50 9*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "70 51 0*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "70 51 1*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "70 51 3*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "70 60 1*": {"un_no": "3082", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, SIVI"},
    "70 60 3*": {"un_no": "3082", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, SIVI"},
    "70 60 4*": {"un_no": "3082", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, SIVI"},
    "70 60 7*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "70 60 8*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "70 60 9*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "70 61 0*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "70 61 1*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "70 70 1*": {"un_no": "3082", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, SIVI"},
    "70 70 3*": {"un_no": "3082", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, SIVI"},
    "70 70 4*": {"un_no": "3082", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, SIVI"},
    "70 70 7*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "70 70 8*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "70 70 9*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "70 71 0*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "70 71 1*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "80 11 1*": {"un_no": "1263", "sinif": "3", "pg": "II", "tasimaKategorisi": 1, "sevkiyat_adi": "BENZIN"},
    "80 11 3*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "80 11 5*": {"un_no": "3175", "sinif": "4.1", "pg": "II", "tasimaKategorisi": 2, "sevkiyat_adi": "ALEVLENİR KATI"},
    "80 11 7*": {"un_no": "3175", "sinif": "4.1", "pg": "II", "tasimaKategorisi": 2, "sevkiyat_adi": "ALEVLENİR KATI"},
    "80 11 9*": {"un_no": "1263", "sinif": "3", "pg": "II", "tasimaKategorisi": 1, "sevkiyat_adi": "BENZIN"},
    "80 12 1*": {"un_no": "1263", "sinif": "3", "pg": "II", "tasimaKategorisi": 1, "sevkiyat_adi": "BENZIN"},
    "80 31 2*": {"un_no": "1210", "sinif": "3", "pg": "III", "tasimaKategorisi": 1, "sevkiyat_adi": "ADR SINIFLANDI: 3"},
    "80 31 4*": {"un_no": "1210", "sinif": "3", "pg": "III", "tasimaKategorisi": 1, "sevkiyat_adi": "ADR SINIFLANDI: 3"},
    "80 31 6*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "80 31 7*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "80 31 9*": {"un_no": "3082", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, SIVI"},
    "80 40 9*": {"un_no": "3175", "sinif": "4.1", "pg": "II", "tasimaKategorisi": 2, "sevkiyat_adi": "ALEVLENİR KATI"},
    "80 41 1*": {"un_no": "3175", "sinif": "4.1", "pg": "II", "tasimaKategorisi": 2, "sevkiyat_adi": "ALEVLENİR KATI"},
    "80 41 3*": {"un_no": "3175", "sinif": "4.1", "pg": "II", "tasimaKategorisi": 2, "sevkiyat_adi": "ALEVLENİR KATI"},
    "80 41 5*": {"un_no": "1993", "sinif": "3", "pg": "II", "tasimaKategorisi": 1, "sevkiyat_adi": "ALEVLENİR SIVI"},
    "80 41 7*": {"un_no": "1286", "sinif": "3", "pg": "III", "tasimaKategorisi": 2, "sevkiyat_adi": "ADR SINIFLANDI: 3"},
    "80 50 1*": {"un_no": "2206", "sinif": "9", "pg": "III", "tasimaKategorisi": 2, "sevkiyat_adi": "ADR SINIFLANDI: 9"},
    "90 10 1*": {"un_no": "3082", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, SIVI"},
    "90 10 2*": {"un_no": "3082", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, SIVI"},
    "90 10 3*": {"un_no": "3082", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, SIVI"},
    "90 10 4*": {"un_no": "3082", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, SIVI"},
    "90 10 5*": {"un_no": "3082", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, SIVI"},
    "90 10 6*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "90 11 1*": {"un_no": "3077", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, KATI"},
    "90 11 3*": {"un_no": "3082", "sinif": "9", "pg": "III", "tasimaKategorisi": 3, "sevkiyat_adi": "ÇEVRE İÇİN TEHLİKELİ MADDE, SIVI"}
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
    _, muaf = _1136_puan_hesapla(satirlar)
    return 'EVET' if muaf else 'HAYIR'


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
        return 'EVET' if toplam_puan <= MAX_PUAN else 'HAYIR'

    @property
    def un_miktar_str(self) -> str:
        """Kontrol formunun Plaka satırına yazılacak 'UN 3509 • 2300 kg' formatı."""
        if not self._satirlar:
            return f"UN {', '.join(self.un_nolar)} • {self.miktar_kg:.0f} kg" if self.un_nolar else ''
        parcalar = []
        for s in self._satirlar:
            un = s.get('un_no', '')
            mik = s.get('miktar', 0)
            if un:
                parcalar.append(f"UN {un} • {mik:.0f} kg")
        return ', '.join(parcalar)

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
        df = pd.read_excel(dosya_path, engine=engine, header=None, dtype=str)
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
