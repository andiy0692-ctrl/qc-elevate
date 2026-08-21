"use client";

import { useState, useEffect, useRef } from 'react';
import * as XLSX from 'xlsx';
import Image from 'next/image';

// ============================================
// DATA INSPEKSI (116 ITEM) - LENGKAP DENGAN KETENTUAN
// ============================================
const inspectionItems = [
  // A. MESIN (1-23)
  { id: 1, category: "MESIN", item: "Dudukan Mesin", ketentuan: "Kuat", weight: 0.0005 },
  { id: 2, category: "MESIN", item: "Rem Mekanik", ketentuan: "Ada, berfungsi baik", weight: 0.04 },
  { id: 3, category: "MESIN", item: "Rem Electric (Brake Switch)", ketentuan: "Ada, berfungsi baik", weight: 0.008 },
  { id: 4, category: "MESIN", item: "Konstruksi Kamar Mesin", ketentuan: "Bebas Air, Kuat, tahan api", weight: 0.0005 },
  { id: 5, category: "MESIN", item: "Ruang Bebas Kamar Mesin", ketentuan: "Didepan alat pengendali ≥ 700 mm", weight: 0.0005 },
  { id: 6, category: "MESIN", item: "Ruang Bebas Kamar Mesin", ketentuan: "Didepan barang bergerak ≥ 500x600 mm", weight: 0.0005 },
  { id: 7, category: "MESIN", item: "Ruang Bebas Kamar Mesin", ketentuan: "Di atas mesin ≥ 500 mm", weight: 0.0005 },
  { id: 8, category: "MESIN", item: "Penerangan Kamar Mesin", ketentuan: "Area kerja ≥ 100 lux", weight: 0.0005 },
  { id: 9, category: "MESIN", item: "Penerangan Kamar Mesin", ketentuan: "Di antara area kerja ≥ 50 lux", weight: 0.0005 },
  { id: 10, category: "MESIN", item: "Ventilasi/Pendingin Ruangan", ketentuan: "Ada, sesuai Spesifikasi", weight: 0.0005 },
  { id: 11, category: "MESIN", item: "Pintu Kamar Mesin", ketentuan: "Membuka keluar, tahan api, lebar ≥ 75 cm, tinggi 2 meter", weight: 0.0005 },
  { id: 12, category: "MESIN", item: "Posisi Panel Hubung Bagi Listrik", ketentuan: "Di kamar mesin", weight: 0.0005 },
  { id: 13, category: "MESIN", item: "Alat Pelindung Benda Berputar", ketentuan: "Ada", weight: 0.0005 },
  { id: 14, category: "MESIN", item: "Pelindung Lubang Tali Baja/sabuk Penggantung", ketentuan: "Tinggi ≥ 50 mm", weight: 0.0005 },
  { id: 15, category: "MESIN", item: "Tangga menuju kamar mesin", ketentuan: "Permanen, pagar pengaman, tahan api", weight: 0.0005 },
  { id: 16, category: "MESIN", item: "Perbedaan ketinggian lantai di kamar mesin", ketentuan: "Tersedia tangga & pagar pengaman", weight: 0.0005 },
  { id: 17, category: "MESIN", item: "Tersedia Alat Pemadam Api Ringan", ketentuan: "isi ≥ 5kg", weight: 0.0005 },
  { id: 18, category: "MESIN", item: "MRL - Penempatan panel kontrol dan PHB listrik", ketentuan: "Berada di lantai yang sama dan berjarak tidak lebih dari 5000 mm", weight: 0.009 },
  { id: 19, category: "MESIN", item: "MRL - Intensitas cahaya area kerja", ketentuan: "≥ 100 lux", weight: 0.0005 },
  { id: 20, category: "MESIN", item: "MRL - Intensitas cahaya diantara area kerja", ketentuan: "≥ 50 lux", weight: 0.0005 },
  { id: 21, category: "MESIN", item: "MRL - Alat pembuka rem mesin", ketentuan: "Ada dan terpasang dengan baik (elektrik/manual)", weight: 0.04 },
  { id: 22, category: "MESIN", item: "MRL - Penempatan APAR", ketentuan: "Dekat pintu elevator paling atas", weight: 0.0005 },
  { id: 23, category: "MESIN", item: "MRL - Emergency stop switch", ketentuan: "Terpasang di dekat panel kontrol", weight: 0.009 },
  // B. TALI/SABUK PENGGANTUNG (24-30)
  { id: 24, category: "TALI/SABUK PENGGANTUNG", item: "Tali / sabuk penggantung", ketentuan: "Tidak memiliki sambungan, kuat, luwes dan memiliki spesifikasi bahan yang seragam", weight: 0.04 },
  { id: 25, category: "TALI/SABUK PENGGANTUNG", item: "Tali/sabuk penggantung", ketentuan: "Tidak menggunakan rantai", weight: 0.009 },
  { id: 26, category: "TALI/SABUK PENGGANTUNG", item: "Nilai faktor keamanan tali / sabuk penggantung", ketentuan: "Kec. 20 - 59 m/menit ≥ 8 kali", weight: 0.009 },
  { id: 27, category: "TALI/SABUK PENGGANTUNG", item: "Tali penggantung Kereta jenis tali dengan bobot imbang", ketentuan: "≥ 6mm, ≥ 3 jalur", weight: 0.009 },
  { id: 28, category: "TALI/SABUK PENGGANTUNG", item: "Tali penggantung Kereta tanpa Bobot imbang", ketentuan: "≥ 6mm, ≥ 2 jalur", weight: 0.009 },
  { id: 29, category: "TALI/SABUK PENGGANTUNG", item: "Sabuk", ketentuan: "≥ 3 x 30 mm, ≥ 2 jalur", weight: 0.009 },
  { id: 30, category: "TALI/SABUK PENGGANTUNG", item: "Alat Pengaman pada elevator tanpa bobot imbang", ketentuan: "Switch otomatis berfungsi dan motor penggerak berhenti", weight: 0.009 },
  // C. TEROMOL (31-33)
  { id: 31, category: "TEROMOL", item: "Alur teromol", ketentuan: "Ada", weight: 0.009 },
  { id: 32, category: "TEROMOL", item: "Diameter teromol Penumpang/barang", ketentuan: "40 : 1", weight: 0.0005 },
  { id: 33, category: "TEROMOL", item: "Diameter teromol Governor", ketentuan: "25 : 1", weight: 0.0005 },
  // D. BANGUNAN RUANG LUNCUR, RUANG ATAS DAN LEKUK DASAR (34-53)
  { id: 34, category: "BANGUNAN", item: "Konstruksi ruang luncur", ketentuan: "Kuat, kokoh, tahan api, dan tertutup rapat", weight: 0.009 },
  { id: 35, category: "BANGUNAN", item: "Dinding ruang luncur (Elevator panorama)", ketentuan: "Dapat dilalui orang dengan tinggi ≥ 2000 mm", weight: 0.0005 },
  { id: 36, category: "BANGUNAN", item: "Ruang luncur", ketentuan: "Bersih, bebas dari instalasi dan peralatan lainnya", weight: 0.0005 },
  { id: 37, category: "BANGUNAN", item: "Penerangan ruang luncur", ketentuan: "≥ 100 lux", weight: 0.0005 },
  { id: 38, category: "BANGUNAN", item: "Pintu darurat (non stop)", ketentuan: "Jarak paling jauh 1100 mm, tinggi ambang pintu paling jauh 300 mm", weight: 0.0005 },
  { id: 39, category: "BANGUNAN", item: "Ukuran pintu darurat", ketentuan: "lebar 700 mm, tinggi 1400 mm, membuka keluar", weight: 0.0005 },
  { id: 40, category: "BANGUNAN", item: "Saklar pengaman pintu darurat", ketentuan: "Tersedia", weight: 0.0005 },
  { id: 41, category: "BANGUNAN", item: "Jembatan bantu dari pintu darurat", ketentuan: "Tersedia, lebar ≥ 500 mm, berpagar", weight: 0.0005 },
  { id: 42, category: "BANGUNAN", item: "Ruang bebas diatas sangkar", ketentuan: "≥ 500 mm", weight: 0.009 },
  { id: 43, category: "BANGUNAN", item: "Ruang bebas lekuk dasar", ketentuan: "≥ 500 mm, kecuali Elevator rumah tinggal ≥ 300 mm", weight: 0.009 },
  { id: 44, category: "BANGUNAN", item: "Tangga lekuk dasar", ketentuan: "Tersedia mulai dari 1000 mm", weight: 0.0005 },
  { id: 45, category: "BANGUNAN", item: "Kekuatan struktur lantai lekuk dasar", ketentuan: "paling sedikit 500 N/meter²", weight: 0.0005 },
  { id: 46, category: "BANGUNAN", item: "Lekuk dasar - Tersedia rem pengaman", ketentuan: "Tersedia", weight: 0.009 },
  { id: 47, category: "BANGUNAN", item: "Lekuk dasar - Tidak sebagai tempat kerja", ketentuan: "Tidak sebagai tempat kerja", weight: 0.0005 },
  { id: 48, category: "BANGUNAN", item: "Akses menuju lekuk dasar", ketentuan: "Tersedia saklar pengaman dengan tinggi 1500 mm, mudah dijangkau, dan 500 mm dari lantai pit", weight: 0.009 },
  { id: 49, category: "BANGUNAN", item: "Lekuk dasar antar 2 Elevator", ketentuan: "Tersedia pit screen dengan tinggi mulai dari 300 mm dari dasar pit sampai 3000 mm keatas", weight: 0.0005 },
  { id: 50, category: "BANGUNAN", item: "Daun pintu ruang luncur", ketentuan: "Tahan api ≥ 1 jam, menutup rapat", weight: 0.0005 },
  { id: 51, category: "BANGUNAN", item: "Interlock / kunci kait pintu ruang luncur", ketentuan: "Tersedia, dapat menutup rapat, pintu hanya terbuka pada zona pemberhentian", weight: 0.009 },
  { id: 52, category: "BANGUNAN", item: "Kerataan lantai", ketentuan: "< 10 mm", weight: 0.009 },
  { id: 53, category: "BANGUNAN", item: "Sekat ruang luncur (2 sangkar)", ketentuan: "≤ 500 mm", weight: 0.0005 },
  // E. KERETA (54-89)
  { id: 54, category: "KERETA", item: "Kerangka", ketentuan: "Dari baja dan kuat", weight: 0.0005 },
  { id: 55, category: "KERETA", item: "Badan kereta", ketentuan: "Tertutup dan ada pintu", weight: 0.0005 },
  { id: 56, category: "KERETA", item: "Tinggi dinding", ketentuan: "≥ 2000 mm", weight: 0.0005 },
  { id: 57, category: "KERETA", item: "Luas lantai", ketentuan: "Sesuai jumlah penumpang", weight: 0.0005 },
  { id: 58, category: "KERETA", item: "Perluasan luas kereta - Pasien", ketentuan: "Max 6%", weight: 0.0005 },
  { id: 59, category: "KERETA", item: "Perluasan luas kereta - Barang", ketentuan: "Max 14%", weight: 0.0005 },
  { id: 60, category: "KERETA", item: "Pintu kereta", ketentuan: "Kokoh, aman, otomatis", weight: 0.0005 },
  { id: 61, category: "KERETA", item: "Ukuran Pintu Kereta", ketentuan: "≥ 700 x 2000 mm", weight: 0.0005 },
  { id: 62, category: "KERETA", item: "Pintu kereta - Kunci kait dan saklar pengaman", ketentuan: "Ada", weight: 0.0005 },
  { id: 63, category: "KERETA", item: "Celah antar ambang pintu kereta", ketentuan: "28 ≤ celah ≤ 32 mm", weight: 0.0005 },
  { id: 64, category: "KERETA", item: "Sisi luar kereta dg balok pemisah ruang luncur", ketentuan: "≥ 250 mm", weight: 0.0005 },
  { id: 65, category: "KERETA", item: "Alarm bell", ketentuan: "Tersedia", weight: 0.04 },
  { id: 66, category: "KERETA", item: "Sumber tenaga cadangan (ARD)", ketentuan: "Tersedia", weight: 0.04 },
  { id: 67, category: "KERETA", item: "Intercom", ketentuan: "Tersedia", weight: 0.04 },
  { id: 68, category: "KERETA", item: "Ventilasi", ketentuan: "Tersedia", weight: 0.009 },
  { id: 69, category: "KERETA", item: "Penerangan darurat", ketentuan: "Tersedia", weight: 0.009 },
  { id: 70, category: "KERETA", item: "Panel operasi", ketentuan: "Tersedia", weight: 0.0005 },
  { id: 71, category: "KERETA", item: "Petunjuk posisi sangkar", ketentuan: "Tersedia", weight: 0.0005 },
  { id: 72, category: "KERETA", item: "Nama pembuat pada panel", ketentuan: "Tersedia", weight: 0.0005 },
  { id: 73, category: "KERETA", item: "Kapasitas beban pada panel", ketentuan: "Tersedia", weight: 0.0005 },
  { id: 74, category: "KERETA", item: "Rambu dilarang merokok", ketentuan: "Tersedia", weight: 0.0005 },
  { id: 75, category: "KERETA", item: "Indikasi beban lebih", ketentuan: "Tersedia", weight: 0.0005 },
  { id: 76, category: "KERETA", item: "Tombol buka dan tutup", ketentuan: "Tersedia", weight: 0.0005 },
  { id: 77, category: "KERETA", item: "Tombol lantai pemberhentian", ketentuan: "Tersedia", weight: 0.0005 },
  { id: 78, category: "KERETA", item: "Tombol bell alarm", ketentuan: "Tersedia", weight: 0.0005 },
  { id: 79, category: "KERETA", item: "Intercom dua arah", ketentuan: "Tersedia", weight: 0.009 },
  { id: 80, category: "KERETA", item: "Kekuatan atap kereta", ketentuan: "≥ 200 Kg", weight: 0.0005 },
  { id: 81, category: "KERETA", item: "Pintu darurat atap kereta", ketentuan: "Berengsel, saklar pengaman, dapat dibuka dari luar, tidak mengganggu instalasi, ukuran ≥ 350 x 450 mm", weight: 0.0005 },
  { id: 82, category: "KERETA", item: "Pintu darurat samping kereta", ketentuan: "Berengsel, dapat dibuka dari luar, dilengkapi Saklar pengaman, ada pegangan tangan, warna kuning, Ukuran ≥ 350 x 1800 mm", weight: 0.0005 },
  { id: 83, category: "KERETA", item: "Pagar pengaman atap kereta", ketentuan: "Warna kuning", weight: 0.0005 },
  { id: 84, category: "KERETA", item: "Pagar pengaman atap kereta", ketentuan: "≥ 90 Kg (kekuatan)", weight: 0.0005 },
  { id: 85, category: "KERETA", item: "Pagar pengaman dengan celah 300 - 850 mm", ketentuan: "Tinggi ≥ 700 mm", weight: 0.0005 },
  { id: 86, category: "KERETA", item: "Pagar pengaman dengan celah lebih dari 850 mm", ketentuan: "Tinggi ≥ 1100 mm", weight: 0.0005 },
  { id: 87, category: "KERETA", item: "Penerangan atap kereta", ketentuan: "≥ 100 Lux dengan kabel lentur 2 m", weight: 0.0005 },
  { id: 88, category: "KERETA", item: "Tombol operasi manual", ketentuan: "Permanen dengan tombol utama", weight: 0.0005 },
  { id: 89, category: "KERETA", item: "Interior kereta", ketentuan: "Bahan tidak mudah pecah dan membahayakan, serta memperhitungkan factor keamanan dan kapasitas motor", weight: 0.0005 },
  // F. GOVERNOR DAN REM PENGAMAN KERETA (90-100)
  { id: 90, category: "GOVERNOR", item: "Penjepit tali / sabuk governor", ketentuan: "Bekerja", weight: 0.09 },
  { id: 91, category: "GOVERNOR", item: "Saklar governor", ketentuan: "Berfungsi", weight: 0.09 },
  { id: 92, category: "GOVERNOR", item: "Fungsi kecepatan rem pengaman", ketentuan: "115% - 140%", weight: 0.009 },
  { id: 93, category: "GOVERNOR", item: "Rem pengaman", ketentuan: "Dipasang pada sangkar, berfungsi secara bertahap, berangsur, dan /mendadak", weight: 0.009 },
  { id: 94, category: "GOVERNOR", item: "Bentuk rem pengaman", ketentuan: "Tidak boleh sistem elektris, hidrolik, atau pneumatis", weight: 0.009 },
  { id: 95, category: "GOVERNOR", item: "Rem pengaman berangsur", ketentuan: "> 60 m/menit", weight: 0.009 },
  { id: 96, category: "GOVERNOR", item: "Rem pengaman mendadak", ketentuan: "< 60 m/menit", weight: 0.009 },
  { id: 97, category: "GOVERNOR", item: "Rem pengaman", ketentuan: "Bekerja kebawah, Bekerja serempak", weight: 0.09 },
  { id: 98, category: "GOVERNOR", item: "Kecepatan ≥ 60 m/menit", ketentuan: "Ada pemutus elektrik, Ps,25/3", weight: 0.009 },
  { id: 99, category: "GOVERNOR", item: "Saklar pengaman lintas batas", ketentuan: "Berfungsi", weight: 0.04 },
  { id: 100, category: "GOVERNOR", item: "Alat pembatas beban lebih", ketentuan: "Berfungsi", weight: 0.04 },
  // G. BOBOT IMBANG, REL PEMANDU DAN PEREDAM (101-106)
  { id: 101, category: "BOBOT IMBANG", item: "Bahan yang dipergunakan", ketentuan: "Beton / Steel Block", weight: 0.0005 },
  { id: 102, category: "BOBOT IMBANG", item: "Pemasangan sekat pengaman bobot imbang", ketentuan: "Dimulai dari 300 mm dari dasar pit, mengelilingi bobot imbang jika terdapat celah > 300 mm", weight: 0.0005 },
  { id: 103, category: "BOBOT IMBANG", item: "Konstruksi rel pemandu", ketentuan: "Kuat memandu jalan", weight: 0.0005 },
  { id: 104, category: "BOBOT IMBANG", item: "Jenis Peredam", ketentuan: "massif kenyal / pegas / hidrolik", weight: 0.0005 },
  { id: 105, category: "BOBOT IMBANG", item: "Fungsi peredaman", ketentuan: "Meredam secara bertahap", weight: 0.04 },
  { id: 106, category: "BOBOT IMBANG", item: "Saklar pengaman untuk kecepatan ≥ 90 m/menit", ketentuan: "Tersedia", weight: 0.009 },
  // H. INSTALASI LISTRIK (107-116)
  { id: 107, category: "LISTRIK", item: "Standar rangkaian instalasi listrik", ketentuan: "SNI dan standar internasional", weight: 0.009 },
  { id: 108, category: "LISTRIK", item: "Panel listrik", ketentuan: "Panel khusus untuk elevator", weight: 0.009 },
  { id: 109, category: "LISTRIK", item: "Catu daya pengganti listrik otomatis (ARD)", ketentuan: "Tersedia", weight: 0.04 },
  { id: 110, category: "LISTRIK", item: "Kabel grounding", ketentuan: "Penampang ≥ 10 mm2", weight: 0.009 },
  { id: 111, category: "LISTRIK", item: "Kabel grounding", ketentuan: "≤ 5 Ω (ohm)", weight: 0.009 },
  { id: 112, category: "LISTRIK", item: "Alarm kebakaran", ketentuan: "Terhubung dan beroperasi otomatis", weight: 0.009 },
  { id: 113, category: "LISTRIK", item: "Lebih dari 10 lantai / 40 meter", ketentuan: "Tersedia sensor gempa", weight: 0.009 },
  { id: 114, category: "LISTRIK", item: "Input signal sensor gempa", ketentuan: "Berhenti lantai terdekat, pintu terbuka, tidak dapat dioperasikan", weight: 0.0005 },
];

// ============================================
// DATA PEMELIHARAAN (20 ITEM)
// ============================================
const maintenanceItems = [
  { id: 1, item: "Pembersihan ruang mesin" },
  { id: 2, item: "Pembersihan ruang luncur" },
  { id: 3, item: "Pembersihan ruang luncur bagian bawah (Pit)" },
  { id: 4, item: "Pembersihan dan pengecekan saringan oli" },
  { id: 5, item: "Pengecekan kebocoran oli" },
  { id: 6, item: "Pelumasan rel pemandu" },
  { id: 7, item: "Pelumasan rantai pengimbang" },
  { id: 8, item: "Pelumasan komponen-komponen mesin" },
  { id: 9, item: "Pengecekan level oli" },
  { id: 10, item: "Pengecekan dan penyetelan rem" },
  { id: 11, item: "Pengecekan dan penyetelan governor" },
  { id: 12, item: "Pengecekan dan penyetelan alat pengaman" },
  { id: 13, item: "Pengecekan dan penyetelan switch batas" },
  { id: 14, item: "Pengecekan dan penyetelan pintu lantai" },
  { id: 15, item: "Pengecekan dan penyetelan pintu kereta" },
  { id: 16, item: "Pengecekan dan penyetelan kabel/sabuk" },
  { id: 17, item: "Pengecekan dan pengencangan baut-baut" },
  { id: 18, item: "Pengecekan dan penggantian saklar pengaman" },
  { id: 19, item: "Pengecekan dan penggantian komponen listrik" },
  { id: 20, item: "Pengecekan dan pengujian fungsi elevator" },
];

// ============================================
// TYPES
// ============================================
type StatusType = 'Good' | 'Not Good' | 'N/A' | '';
type MaintenanceStatusType = 'Good' | 'Not Good' | 'N/A' | '';
type QCStatusType = 'Approved' | 'Revision Required' | '';
type UserRole = 'maintenance' | 'qc' | 'sales' | null;
type MenuType = 'dashboard' | 'newPemeriksaan' | 'newPemeliharaan' | 'viewReport' | 'review';
type SortOrder = 'terbaru' | 'terlama';
type ReportType = 'pemeriksaan' | 'pemeliharaan';

// ============================================
// INTERFACE TEMUAN DENGAN UPLOAD FOTO
// ============================================
interface TemuanKeselamatanType {
  id: number;
  temuan: string;
  dokumentasi: string | null;
  kategori: string;
  solusi: string;
}

interface TemuanPermenakerType {
  id: number;
  temuan: string;
  dokumentasi: string | null;
  kategori: string;
  solusi: string;
}

interface TemuanOperasionalType {
  id: number;
  temuan: string;
  dokumentasi: string | null;
  kategori: string;
  solusi: string;
}

interface TemuanAreaLainType {
  id: number;
  temuan: string;
  dokumentasi: string | null;
  solusi: string;
}

interface InspectionItemType {
  id: number;
  status: StatusType;
  notes: string;
  photoBefore: string | null;
  photoAfter: string | null;
  repairNote: string;
  isApproved: boolean;
}

interface MaintenanceItemType {
  id: number;
  status: MaintenanceStatusType;
  finding: string;
  photoBefore: string | null;
  photoAfter: string | null;
  repairNote: string;
  isApproved: boolean;
}

interface UnitData {
  unitNumber: string;
  projectCode: string;
  customerName: string;
  buildingLocation: string;
  elevatorType: string;
  elevatorBrand: string;
  elevatorModel: string;
  capacity: string;
  speed: string;
  inspectionDate: string;
  qcName: string;
  maintenanceData: MaintenanceItemType[];
}

interface Report {
  id: string;
  reportType: ReportType;
  jumlahUnit: number;
  units: UnitData[];
  teknisiName: string;
  // Data Umum (Header Laporan) - SESUAI WORD
  namaGedung: string;
  alamat: string;
  jenisElevator: string;
  merkTipe: string;
  elevatorNo: string;
  tanggalRiksaUji: string;
  namaPabrikPembuat: string;
  tahunPemasangan: string;
  kapasitasAngkut: string;
  kecepatanAngkut: string;
  melayani: string;
  dataPemeliharaan: string;
  hasilRiksaUjiSebelumnya: string;
  // Data Pemeriksaan
  inspectionData: InspectionItemType[];
  // Temuan Negatif - DINAMIS (bisa tambah/hapus)
  temuanKeselamatan: TemuanKeselamatanType[];
  temuanPermenaker: TemuanPermenakerType[];
  temuanOperasional: TemuanOperasionalType[];
  temuanAreaLain: TemuanAreaLainType[];
  // Lainnya
  qcNote: string;
  attachment: string | null;
  attachmentName: string | null;
  qcVerification: {
    qcName: string;
    qcStatus: QCStatusType;
    qcNote: string;
    verifiedAt: string;
  };
  submittedBy: string;
  submittedAt: string;
  status: 'draft' | 'qc_approved' | 'revision' | 'maintenance_done' | 'approved';
  createdAt: string;
}

// ============================================
// LOCAL STORAGE UTILITIES
// ============================================
const STORAGE_KEY = 'elevateqc_reports';

function loadFromStorage(): Report[] {
  try {
    if (typeof window === 'undefined') return [];
    const raw = localStorage.getItem(STORAGE_KEY);
    return raw ? JSON.parse(raw) : [];
  } catch {
    return [];
  }
}

function saveToStorage(reports: Report[]) {
  try {
    if (typeof window === 'undefined') return;
    localStorage.setItem(STORAGE_KEY, JSON.stringify(reports));
  } catch {}
}

// ============================================
// HELPER FUNCTIONS
// ============================================
const calculateTotalScore = (inspectionData: InspectionItemType[]) => {
  let totalWeight = 0;
  let achievedWeight = 0;
  
  inspectionData.forEach((item) => {
    const originalItem = inspectionItems.find(i => i.id === item.id);
    if (originalItem) {
      totalWeight += originalItem.weight;
      if (item.status === 'Good' || item.status === 'N/A') {
        achievedWeight += originalItem.weight;
      } else if (item.status === 'Not Good' && item.isApproved) {
        achievedWeight += originalItem.weight;
      }
    }
  });
  
  return totalWeight > 0 ? (achievedWeight / totalWeight) * 100 : 0;
};

const calculateHitung = (item: InspectionItemType) => {
  const originalItem = inspectionItems.find(i => i.id === item.id);
  if (!originalItem) return 0;
  
  if (item.status === 'Good' || item.status === 'N/A') {
    return originalItem.weight;
  }
  if (item.status === 'Not Good' && item.isApproved) {
    return originalItem.weight;
  }
  if (item.status === 'Not Good' && !item.isApproved) {
    return -originalItem.weight;
  }
  return 0;
};

// ============================================
// KOMPONEN LOADING SPINNER
// ============================================
function LoadingSpinner({ message = 'Memuat...' }: { message?: string }) {
  return (
    <div className="flex flex-col items-center justify-center min-h-[60vh]">
      <div className="relative">
        <div className="w-16 h-16 border-4 border-gray-200 rounded-full"></div>
        <div className="absolute top-0 left-0 w-16 h-16 border-4 border-blue-600 rounded-full border-t-transparent animate-spin"></div>
      </div>
      <p className="mt-6 text-gray-500 text-sm font-medium tracking-wide">{message}</p>
    </div>
  );
}

// ============================================
// KOMPONEN TOAST NOTIFICATION
// ============================================
function Toast({ message, type, onClose }: { message: string; type: 'success' | 'error' | 'info'; onClose: () => void }) {
  useEffect(() => {
    const timer = setTimeout(onClose, 3000);
    return () => clearTimeout(timer);
  }, [onClose]);

  const colors = {
    success: 'bg-emerald-500',
    error: 'bg-rose-500',
    info: 'bg-blue-500'
  };

  return (
    <div className={`fixed top-6 right-6 z-50 px-6 py-4 rounded-xl shadow-2xl text-white ${colors[type]} transform transition-all duration-500 animate-slide-in max-w-md`}>
      <div className="flex items-center gap-3">
        {type === 'success' && <span>✅</span>}
        {type === 'error' && <span>❌</span>}
        {type === 'info' && <span>ℹ️</span>}
        <span className="font-medium">{message}</span>
      </div>
    </div>
  );
}

// ============================================
// KOMPONEN LOGIN
// ============================================
function LoginPage({ onLogin }: { onLogin: (role: UserRole) => void }) {
  const [username, setUsername] = useState('');
  const [password, setPassword] = useState('');
  const [selectedRole, setSelectedRole] = useState<UserRole>(null);
  const [error, setError] = useState('');
  const [loading, setLoading] = useState(false);

  const handleLogin = async (e: React.FormEvent) => {
    e.preventDefault();
    setError('');
    setLoading(true);

    if (!selectedRole) {
      setError('Silakan pilih peran terlebih dahulu');
      setLoading(false);
      return;
    }

    await new Promise(resolve => setTimeout(resolve, 800));

    if (
      (username === 'admin' && password === '123') ||
      (username === 'teknisi' && password === '123') ||
      (username === 'sales' && password === '123')
    ) {
      const roleMap: Record<string, UserRole> = {
        'admin': 'qc',
        'teknisi': 'maintenance',
        'sales': 'sales'
      };
      const role = roleMap[username];
      onLogin(role);
      setLoading(false);
      return;
    }

    setError('Username atau password salah');
    setLoading(false);
  };

  return (
    <div className="min-h-screen bg-gradient-to-br from-slate-50 via-white to-blue-50 flex items-center justify-center p-4">
      <div className="bg-white/80 backdrop-blur-sm rounded-3xl shadow-2xl max-w-md w-full p-8 transform transition-all duration-700 animate-fade-in border border-gray-100">
        <div className="text-center mb-8">
          <div className="flex justify-center mb-4">
            <Image 
              src="/logo_louser_2022_1696999044 (1).png" 
              alt="Logo Louser" 
              width={200} 
              height={60} 
              className="object-contain"
              priority
            />
          </div>
          <p className="text-sm text-gray-400 mt-1 font-light">Elevator Quality Control System</p>
        </div>

        <form onSubmit={handleLogin} className="space-y-4">
          <div className="grid grid-cols-3 gap-2">
            {(['maintenance', 'qc', 'sales'] as const).map((role) => (
              <button
                key={role}
                type="button"
                onClick={() => setSelectedRole(role)}
                className={`p-3 rounded-xl border-2 text-center transition-all duration-200 text-sm ${
                  selectedRole === role
                    ? role === 'maintenance' ? 'border-blue-500 bg-blue-50 text-blue-700 shadow-md scale-105' :
                      role === 'qc' ? 'border-emerald-500 bg-emerald-50 text-emerald-700 shadow-md scale-105' :
                      'border-amber-500 bg-amber-50 text-amber-700 shadow-md scale-105'
                    : 'border-gray-200 hover:border-gray-300 hover:bg-gray-50'
                }`}
              >
                <div className="text-2xl">
                  {role === 'maintenance' ? '🔧' : role === 'qc' ? '✅' : '💰'}
                </div>
                <div className="text-[10px] font-semibold mt-1 uppercase tracking-wider">
                  {role === 'maintenance' ? 'Maintenance' : role === 'qc' ? 'QC' : 'Sales'}
                </div>
              </button>
            ))}
          </div>

          <input
            type="text"
            value={username}
            onChange={(e) => setUsername(e.target.value)}
            placeholder="Username"
            className="w-full px-4 py-3 bg-gray-50 border border-gray-200 rounded-xl focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-transparent transition"
          />

          <input
            type="password"
            value={password}
            onChange={(e) => setPassword(e.target.value)}
            placeholder="Password"
            className="w-full px-4 py-3 bg-gray-50 border border-gray-200 rounded-xl focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-transparent transition"
          />

          {error && <div className="text-red-500 text-sm text-center bg-red-50 p-3 rounded-xl animate-shake">{error}</div>}

          <button
            type="submit"
            disabled={loading}
            className={`w-full font-semibold py-3.5 rounded-xl transition-all duration-300 ${
              loading 
                ? 'bg-gray-300 text-white cursor-not-allowed' 
                : 'bg-gradient-to-r from-blue-700 to-blue-800 hover:from-blue-800 hover:to-blue-900 text-white shadow-lg hover:shadow-xl transform hover:scale-[1.02]'
            }`}
          >
            {loading ? 'Memproses...' : 'Masuk'}
          </button>

          <p className="text-center text-[10px] text-gray-400 mt-4">© 2026 PT Louserindo Megah Permai</p>
        </form>
      </div>

      <style>{`
        @keyframes fade-in {
          from { opacity: 0; transform: translateY(-20px); }
          to { opacity: 1; transform: translateY(0); }
        }
        @keyframes slide-in {
          from { transform: translateX(100%); opacity: 0; }
          to { transform: translateX(0); opacity: 1; }
        }
        @keyframes shake {
          0%, 100% { transform: translateX(0); }
          25% { transform: translateX(-5px); }
          75% { transform: translateX(5px); }
        }
        .animate-fade-in { animation: fade-in 0.5s ease-out; }
        .animate-slide-in { animation: slide-in 0.3s ease-out; }
        .animate-shake { animation: shake 0.3s ease-in-out; }
      `}</style>
    </div>
  );
}

// ============================================
// KOMPONEN DASHBOARD
// ============================================
function Dashboard({ role, onLogout }: { 
  role: UserRole; 
  onLogout: () => void;
}) {
  const [currentMenu, setCurrentMenu] = useState<MenuType>('dashboard');
  const [selectedReport, setSelectedReport] = useState<Report | null>(null);
  const [reports, setReports] = useState<Report[]>([]);
  const [isReadOnly, setIsReadOnly] = useState(false);
  const [sortOrder, setSortOrder] = useState<SortOrder>('terbaru');
  const [loading, setLoading] = useState(true);
  const [toast, setToast] = useState<{ message: string; type: 'success' | 'error' | 'info' } | null>(null);
  const [activeNote, setActiveNote] = useState<string>('');
  const fileInputRef = useRef<HTMLInputElement>(null);

  const isQC = role === 'qc';
  const isMaintenance = role === 'maintenance';
  const isSales = role === 'sales';

  const showToast = (message: string, type: 'success' | 'error' | 'info' = 'success') => {
    setToast({ message, type });
    setTimeout(() => setToast(null), 3000);
  };

  useEffect(() => {
    setLoading(true);
    setReports(loadFromStorage());
    setLoading(false);
  }, []);

  const handleExportJSON = () => {
    const data = loadFromStorage();
    const today = new Date().toISOString().split('T')[0];
    const blob = new Blob([JSON.stringify(data, null, 2)], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = `elevateQC_export_${today}.json`;
    link.click();
    URL.revokeObjectURL(url);
    showToast('Data berhasil diexport!', 'success');
  };

  const handleImportJSON = (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const importedData = JSON.parse(e.target?.result as string) as Report[];
        const currentData = loadFromStorage();
        const merged = [...currentData];
        importedData.forEach((newItem) => {
          const existingIndex = merged.findIndex(item => item.id === newItem.id);
          if (existingIndex !== -1) {
            merged[existingIndex] = newItem;
          } else {
            merged.push(newItem);
          }
        });
        saveToStorage(merged);
        setReports(merged);
        showToast('Data berhasil diimpor & digabung!', 'success');
      } catch {
        showToast('Gagal mengimpor file! Pastikan file JSON valid.', 'error');
      }
    };
    reader.readAsText(file);
    if (fileInputRef.current) fileInputRef.current.value = '';
  };

  const handleRestoreJSON = (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;
    if (!confirm('⚠️ PERINGATAN! Tindakan ini akan MENGHAPUS TOTAL semua data saat ini dan menggantinya dengan data dari file ini. Lanjutkan?')) {
      if (fileInputRef.current) fileInputRef.current.value = '';
      return;
    }
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const importedData = JSON.parse(e.target?.result as string) as Report[];
        saveToStorage(importedData);
        setReports(importedData);
        showToast('Data berhasil direstore (data lama diganti)!', 'success');
      } catch {
        showToast('Gagal merestore file! Pastikan file JSON valid.', 'error');
      }
    };
    reader.readAsText(file);
    if (fileInputRef.current) fileInputRef.current.value = '';
  };

  const handleNewPemeriksaan = () => {
    setActiveNote('NEW DATA PEMERIKSAAN');
    const newReport: Report = {
      id: Date.now().toString(),
      reportType: 'pemeriksaan',
      jumlahUnit: 1,
      units: [],
      teknisiName: '',
      namaGedung: '',
      alamat: '',
      jenisElevator: '',
      merkTipe: '',
      elevatorNo: '',
      tanggalRiksaUji: '',
      namaPabrikPembuat: '',
      tahunPemasangan: '',
      kapasitasAngkut: '',
      kecepatanAngkut: '',
      melayani: '',
      dataPemeliharaan: '',
      hasilRiksaUjiSebelumnya: '',
      inspectionData: inspectionItems.map(item => ({
        id: item.id,
        status: '' as StatusType,
        notes: '',
        photoBefore: null,
        photoAfter: null,
        repairNote: '',
        isApproved: false,
      })),
      temuanKeselamatan: [{ id: 1, temuan: '', dokumentasi: null, kategori: '', solusi: '' }],
      temuanPermenaker: [{ id: 1, temuan: '', dokumentasi: null, kategori: '', solusi: '' }],
      temuanOperasional: [{ id: 1, temuan: '', dokumentasi: null, kategori: '', solusi: '' }],
      temuanAreaLain: [{ id: 1, temuan: '', dokumentasi: null, solusi: '' }],
      qcNote: '',
      attachment: null,
      attachmentName: null,
      qcVerification: { qcName: '', qcStatus: '' as QCStatusType, qcNote: '', verifiedAt: '' },
      submittedBy: '',
      submittedAt: '',
      status: 'draft',
      createdAt: new Date().toISOString(),
    };
    setSelectedReport(newReport);
    setIsReadOnly(false);
    setCurrentMenu('newPemeriksaan');
  };

  const handleNewPemeliharaan = () => {
    setActiveNote('NEW DATA PEMELIHARAAN');
    const defaultUnit: UnitData = {
      unitNumber: '',
      projectCode: '',
      customerName: '',
      buildingLocation: '',
      elevatorType: '',
      elevatorBrand: '',
      elevatorModel: '',
      capacity: '',
      speed: '',
      inspectionDate: '',
      qcName: '',
      maintenanceData: maintenanceItems.map(item => ({
        id: item.id,
        status: '' as MaintenanceStatusType,
        finding: '',
        photoBefore: null,
        photoAfter: null,
        repairNote: '',
        isApproved: false,
      })),
    };
    const newReport: any = {
      id: Date.now().toString(),
      reportType: 'pemeliharaan',
      jumlahUnit: 1,
      units: [defaultUnit],
      teknisiName: '',
      namaGedung: '',
      alamat: '',
      jenisElevator: '',
      merkTipe: '',
      elevatorNo: '',
      tanggalRiksaUji: '',
      namaPabrikPembuat: '',
      tahunPemasangan: '',
      kapasitasAngkut: '',
      kecepatanAngkut: '',
      melayani: '',
      dataPemeliharaan: '',
      hasilRiksaUjiSebelumnya: '',
      inspectionData: [],
      temuanKeselamatan: [{ id: 1, temuan: '', dokumentasi: null, kategori: '', solusi: '' }],
      temuanPermenaker: [{ id: 1, temuan: '', dokumentasi: null, kategori: '', solusi: '' }],
      temuanOperasional: [{ id: 1, temuan: '', dokumentasi: null, kategori: '', solusi: '' }],
      temuanAreaLain: [{ id: 1, temuan: '', dokumentasi: null, solusi: '' }],
      qcNote: '',
      attachment: null,
      attachmentName: null,
      qcVerification: { qcName: '', qcStatus: '' as QCStatusType, qcNote: '', verifiedAt: '' },
      submittedBy: '',
      submittedAt: '',
      status: 'draft',
      createdAt: new Date().toISOString(),
    };
    setSelectedReport(newReport);
    setIsReadOnly(false);
    setCurrentMenu('newPemeliharaan');
  };

  const handleReviewReport = (report: Report) => {
    setActiveNote('REVIEW - Data dari Maintenance');
    setSelectedReport(report);
    setIsReadOnly(false);
    setCurrentMenu('review');
  };

  const handleViewReport = (report: Report) => {
    setSelectedReport(report);
    if (isSales) setIsReadOnly(true);
    else if (report.status === 'approved') setIsReadOnly(true);
    else if (report.status === 'revision' && isMaintenance) setIsReadOnly(false);
    else if (isMaintenance) setIsReadOnly(report.status !== 'qc_approved' && report.status !== 'revision');
    else if (isQC) setIsReadOnly(report.status !== 'draft' && report.status !== 'maintenance_done' && report.status !== 'revision');
    else setIsReadOnly(true);
    setCurrentMenu('viewReport');
    setActiveNote('');
  };

  const handleSaveReport = (data: any) => {
    const current = loadFromStorage();
    const existingIndex = current.findIndex(r => r.id === data.id);
    if (existingIndex !== -1) {
      current[existingIndex] = data;
    } else {
      current.push(data);
    }
    saveToStorage(current);
    setReports(current);
    showToast('Data berhasil disimpan!', 'success');
    setCurrentMenu('dashboard');
    setActiveNote('');
  };

  const handleDeleteReport = (id: string) => {
    if (!isQC) {
      showToast('Hanya QC yang bisa menghapus laporan!', 'error');
      return;
    }
    if (!confirm('Yakin ingin menghapus laporan ini?')) return;
    const current = loadFromStorage().filter(r => r.id !== id);
    saveToStorage(current);
    setReports(current);
    showToast('Laporan berhasil dihapus!', 'success');
  };

  const handleCancel = () => {
    setSelectedReport(null);
    setIsReadOnly(false);
    setCurrentMenu('dashboard');
    setActiveNote('');
  };

  const handlePrintPDF = () => window.print();

  const getFilteredReports = () => {
    let filtered = reports;
    if (!isQC) {
      filtered = reports.filter(r => 
        r.status === 'qc_approved' || 
        r.status === 'revision' ||
        r.status === 'maintenance_done' || 
        r.status === 'approved'
      );
      if (isSales) {
        filtered = filtered.filter(r => r.reportType === 'pemeriksaan');
      }
    }
    return filtered.sort((a, b) => {
      const dateA = new Date(a.createdAt).getTime();
      const dateB = new Date(b.createdAt).getTime();
      return sortOrder === 'terbaru' ? dateB - dateA : dateA - dateB;
    });
  };

  const filteredReports = getFilteredReports();

  const stats = {
    total: reports.length,
    draft: reports.filter(r => r.status === 'draft').length,
    qcApproved: reports.filter(r => r.status === 'qc_approved').length,
    revision: reports.filter(r => r.status === 'revision').length,
    maintenanceDone: reports.filter(r => r.status === 'maintenance_done').length,
    approved: reports.filter(r => r.status === 'approved').length,
  };

  const formatDate = (dateString: string) => {
    if (!dateString) return '-';
    try {
      const date = new Date(dateString);
      return date.toLocaleDateString('id-ID', {
        day: '2-digit', month: 'short', year: 'numeric', hour: '2-digit', minute: '2-digit'
      });
    } catch { return '-'; }
  };

  const totalTugas = isMaintenance ? reports.filter(r => r.status !== 'approved' && (r.status === 'qc_approved' || r.status === 'revision')).length : 0;
  const needReview = isQC ? reports.filter(r => r.status === 'maintenance_done' || r.status === 'revision').length : 0;

  return (
    <div className="min-h-screen bg-gradient-to-br from-slate-50 via-white to-gray-50">
      {toast && <Toast message={toast.message} type={toast.type} onClose={() => setToast(null)} />}
      <div className="container mx-auto px-4 py-6 max-w-7xl">
        <header className="bg-white/80 backdrop-blur-sm rounded-2xl shadow-sm border border-gray-100/50 p-6 mb-6">
          <div className="flex flex-col md:flex-row md:items-center md:justify-between">
            <div className="flex items-center gap-2">
              <div className="relative w-[150px] h-[40px]">
                <Image 
                  src="/logo_louser_2022_1696999044 (1).png" 
                  alt="Logo Louser" 
                  fill
                  className="object-contain object-left"
                  priority
                />
              </div>
              <div className="ml-2">
                <p className="text-xs text-gray-400 font-light tracking-wide">Elevator Quality Control System</p>
                {isMaintenance && <p className="text-sm text-blue-600 mt-1">{totalTugas} tugas aktif</p>}
                {isSales && <p className="text-sm text-amber-600 mt-1">Mode Lihat</p>}
                {isQC && needReview > 0 && <p className="text-sm text-orange-500 mt-1 animate-pulse">{needReview} data perlu direview dari Maintenance</p>}
              </div>
            </div>
            <div className="mt-4 md:mt-0 flex flex-wrap items-center gap-3">
              <span className={`px-4 py-1.5 rounded-full text-sm font-medium ${isQC ? 'bg-emerald-100 text-emerald-700' : isMaintenance ? 'bg-blue-100 text-blue-700' : 'bg-amber-100 text-amber-700'}`}>
                {isQC ? 'QC' : isMaintenance ? 'Maintenance' : 'Sales'}
              </span>
              <button onClick={onLogout} className="text-sm bg-red-500 hover:bg-red-600 text-white px-4 py-2 rounded-xl transition shadow-sm">Keluar</button>
            </div>
          </div>

          <div className="mt-4 grid grid-cols-3 md:grid-cols-6 gap-2">
            <div className="bg-gray-50/80 rounded-xl p-3 text-center"><div className="text-xl font-bold text-gray-700">{stats.total}</div><div className="text-[10px] text-gray-400 uppercase tracking-wider">Total</div></div>
            <div className="bg-gray-50/80 rounded-xl p-3 text-center"><div className="text-xl font-bold text-gray-400">{stats.draft}</div><div className="text-[10px] text-gray-400 uppercase tracking-wider">Draft</div></div>
            <div className="bg-blue-50/80 rounded-xl p-3 text-center"><div className="text-xl font-bold text-blue-600">{stats.qcApproved}</div><div className="text-[10px] text-blue-400 uppercase tracking-wider">Maintenance</div></div>
            <div className="bg-amber-50/80 rounded-xl p-3 text-center"><div className="text-xl font-bold text-amber-600">{stats.revision}</div><div className="text-[10px] text-amber-400 uppercase tracking-wider">Revisi</div></div>
            <div className="bg-orange-50/80 rounded-xl p-3 text-center"><div className="text-xl font-bold text-orange-600">{stats.maintenanceDone}</div><div className="text-[10px] text-orange-400 uppercase tracking-wider">Verifikasi</div></div>
            <div className="bg-emerald-50/80 rounded-xl p-3 text-center"><div className="text-xl font-bold text-emerald-600">{stats.approved}</div><div className="text-[10px] text-emerald-400 uppercase tracking-wider">Final</div></div>
          </div>
        </header>

        {currentMenu === 'dashboard' ? (
          <div className="space-y-4">
            <div className="flex flex-wrap items-center justify-between gap-3">
              <div className="flex flex-wrap gap-3">
                {isQC && (
                  <>
                    <button onClick={handleNewPemeriksaan} className="bg-blue-600 hover:bg-blue-700 text-white px-5 py-2.5 rounded-xl flex items-center gap-2 shadow-sm hover:shadow-md transition text-sm font-medium">
                      <span>📋</span> New Pemeriksaan
                    </button>
                    <button onClick={handleNewPemeliharaan} className="bg-emerald-600 hover:bg-emerald-700 text-white px-5 py-2.5 rounded-xl flex items-center gap-2 shadow-sm hover:shadow-md transition text-sm font-medium">
                      <span>🔧</span> New Pemeliharaan
                    </button>
                  </>
                )}
                {isQC && needReview > 0 && <span className="text-sm text-orange-500 self-center font-medium animate-pulse">{needReview} data perlu review</span>}
              </div>
              <div className="flex items-center gap-2">
                <span className="text-xs text-gray-400">Urutkan:</span>
                <button onClick={() => setSortOrder('terbaru')} className={`px-3 py-1 rounded-lg text-xs transition ${sortOrder === 'terbaru' ? 'bg-blue-600 text-white' : 'bg-gray-100 text-gray-600 hover:bg-gray-200'}`}>Terbaru</button>
                <button onClick={() => setSortOrder('terlama')} className={`px-3 py-1 rounded-lg text-xs transition ${sortOrder === 'terlama' ? 'bg-blue-600 text-white' : 'bg-gray-100 text-gray-600 hover:bg-gray-200'}`}>Terlama</button>
              </div>
            </div>

            <div className="flex flex-wrap items-center gap-3 p-3 bg-white rounded-xl shadow-sm border border-gray-100">
              <span className="text-sm font-medium text-gray-600">Data Management:</span>
              <button onClick={handleExportJSON} className="bg-blue-500 hover:bg-blue-600 text-white px-4 py-2 rounded-lg text-sm shadow-sm transition">📤 Export</button>
              <div className="relative">
                <button onClick={() => fileInputRef.current?.click()} className="bg-green-500 hover:bg-green-600 text-white px-4 py-2 rounded-lg text-sm shadow-sm transition">📥 Import</button>
                <input type="file" accept=".json" ref={fileInputRef} onChange={handleImportJSON} className="absolute inset-0 opacity-0 w-full h-full cursor-pointer" />
              </div>
              <div className="relative">
                <button onClick={() => {
                  if (fileInputRef.current) {
                    fileInputRef.current.accept = '.json';
                    const restoreHandler = (e: Event) => {
                      const target = e.target as HTMLInputElement;
                      handleRestoreJSON(e as unknown as React.ChangeEvent<HTMLInputElement>);
                      target.removeEventListener('change', restoreHandler as EventListener);
                    };
                    fileInputRef.current.addEventListener('change', restoreHandler as EventListener);
                    fileInputRef.current.click();
                  }
                }} className="bg-red-500 hover:bg-red-600 text-white px-4 py-2 rounded-lg text-sm shadow-sm transition">📂 Restore</button>
              </div>
            </div>

            <div className="bg-white/80 backdrop-blur-sm rounded-2xl border border-gray-100/50 overflow-hidden shadow-sm">
              <div className="px-6 py-4 border-b border-gray-100 flex justify-between items-center">
                <div>
                  <h2 className="font-semibold text-gray-700">{isQC ? 'Semua Laporan' : isMaintenance ? 'Tugas Maintenance' : 'Data Pemeriksaan'}</h2>
                  <span className="text-xs text-gray-400">{filteredReports.length} laporan</span>
                </div>
                <span className="text-[10px] text-emerald-500 font-medium bg-emerald-50 px-3 py-1 rounded-full">Local Storage</span>
              </div>

              {loading ? (
                <LoadingSpinner message="Memuat data..." />
              ) : filteredReports.length === 0 ? (
                <div className="p-12 text-center text-gray-400">
                  {isQC ? 'Belum ada laporan. Klik New Pemeriksaan atau New Pemeliharaan untuk mulai.' : 'Belum ada data.'}
                </div>
              ) : (
                <div className="overflow-x-auto">
                  <table className="w-full text-sm">
                    <thead className="bg-gray-50/80">
                      <tr>
                        <th className="px-4 py-3 text-left text-gray-500 text-xs font-medium uppercase tracking-wider">Tanggal</th>
                        <th className="px-4 py-3 text-left text-gray-500 text-xs font-medium uppercase tracking-wider">Unit</th>
                        <th className="px-4 py-3 text-left text-gray-500 text-xs font-medium uppercase tracking-wider">Proyek</th>
                        <th className="px-4 py-3 text-left text-gray-500 text-xs font-medium uppercase tracking-wider">Pelanggan</th>
                        <th className="px-4 py-3 text-left text-gray-500 text-xs font-medium uppercase tracking-wider">Tipe</th>
                        {isQC && <th className="px-4 py-3 text-left text-gray-500 text-xs font-medium uppercase tracking-wider">Teknisi</th>}
                        <th className="px-4 py-3 text-left text-gray-500 text-xs font-medium uppercase tracking-wider">Status</th>
                        <th className="px-4 py-3 text-left text-gray-500 text-xs font-medium uppercase tracking-wider">Aksi</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-gray-100">
                      {filteredReports.map((report) => {
                        const isReview = report.status === 'maintenance_done' || report.status === 'revision';

                        return (
                          <tr key={report.id} className="hover:bg-gray-50/50 transition">
                            <td className="px-4 py-3 text-gray-500 text-xs whitespace-nowrap">{formatDate(report.createdAt)}</td>
                            <td className="px-4 py-3 font-medium text-gray-700 text-sm">{report.namaGedung || '-'}</td>
                            <td className="px-4 py-3 text-gray-600 text-sm">{report.merkTipe || '-'}</td>
                            <td className="px-4 py-3 text-gray-600 text-sm">{report.elevatorNo || '-'}</td>
                            <td className="px-4 py-3">
                              <span className={`px-2.5 py-1 rounded-full text-[10px] font-medium ${report.reportType === 'pemeriksaan' ? 'bg-blue-100 text-blue-700' : 'bg-emerald-100 text-emerald-700'}`}>
                                {report.reportType === 'pemeriksaan' ? 'Pemeriksaan' : 'Pemeliharaan'}
                              </span>
                            </td>
                            {isQC && <td className="px-4 py-3 text-sm text-gray-500">{report.teknisiName || '-'}</td>}
                            <td className="px-4 py-3">
                              <span className={`inline-flex px-2.5 py-1 rounded-full text-[10px] font-medium ${report.status === 'draft' ? 'bg-gray-200 text-gray-600' : report.status === 'qc_approved' ? 'bg-blue-100 text-blue-700' : report.status === 'revision' ? 'bg-amber-100 text-amber-700' : report.status === 'maintenance_done' ? 'bg-orange-100 text-orange-700' : 'bg-emerald-100 text-emerald-700'}`}>
                                {report.status === 'draft' ? 'Draft' : report.status === 'qc_approved' ? 'Maintenance' : report.status === 'revision' ? 'Revisi' : report.status === 'maintenance_done' ? 'Verifikasi' : 'Final'}
                              </span>
                            </td>
                            <td className="px-4 py-3 flex items-center gap-2">
                              {isQC && isReview ? (
                                <button onClick={() => handleReviewReport(report)} className="bg-amber-50 hover:bg-amber-100 text-amber-600 px-3 py-1 rounded-xl text-sm font-medium transition">Review</button>
                              ) : (
                                <button onClick={() => handleViewReport(report)} className="bg-blue-50 hover:bg-blue-100 text-blue-600 px-3 py-1 rounded-xl text-sm font-medium transition">Lihat</button>
                              )}
                              {isQC && report.status !== 'approved' && (
                                <button onClick={() => handleDeleteReport(report.id)} className="text-red-400 hover:text-red-600 ml-2 text-lg transition" title="Hapus">🗑️</button>
                              )}
                            </td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              )}
            </div>
          </div>
        ) : (
          <div className="bg-white/80 backdrop-blur-sm rounded-2xl border border-gray-100/50 p-6 shadow-sm">
            {activeNote && <div className="mb-4 p-3 bg-blue-50 border border-blue-200 rounded-xl text-blue-700 font-medium text-sm text-center">{activeNote}</div>}
            {selectedReport && (
              <ReportForm
                report={selectedReport}
                onSave={handleSaveReport}
                onCancel={handleCancel}
                onDelete={handleDeleteReport}
                userRole={role}
                isReadOnly={isReadOnly}
                onPrintPDF={handlePrintPDF}
                reportType={currentMenu === 'newPemeliharaan' || selectedReport.reportType === 'pemeliharaan' ? 'pemeliharaan' : 'pemeriksaan'}
                isReview={currentMenu === 'review'}
              />
            )}
          </div>
        )}
      </div>
    </div>
  );
}

// ============================================
// KOMPONEN REPORT FORM - PEMERIKSAAN & PEMELIHARAAN
// ============================================
function ReportForm({
  report,
  onSave,
  onCancel,
  onDelete,
  userRole,
  isReadOnly,
  onPrintPDF,
  reportType,
  isReview = false,
}: {
  report: Report | null;
  onSave: (data: any) => void;
  onCancel: () => void;
  onDelete?: (id: string) => void;
  userRole: UserRole;
  isReadOnly: boolean;
  onPrintPDF?: () => void;
  reportType?: ReportType;
  isReview?: boolean;
}) {
  const isQC = userRole === 'qc';
  const isMaintenance = userRole === 'maintenance';
  const isPemeriksaan = reportType === 'pemeriksaan';
  const isPemeliharaan = reportType === 'pemeliharaan';
  const [isSubmitting, setIsSubmitting] = useState(false);
  const [searchQuery, setSearchQuery] = useState('');
  const [formData, setFormData] = useState<any>(() => {
    if (report) return report;
    return {
      id: Date.now().toString(),
      reportType: reportType || 'pemeriksaan',
      jumlahUnit: 1,
      units: [],
      teknisiName: '',
      namaGedung: '',
      alamat: '',
      jenisElevator: '',
      merkTipe: '',
      elevatorNo: '',
      tanggalRiksaUji: '',
      namaPabrikPembuat: '',
      tahunPemasangan: '',
      kapasitasAngkut: '',
      kecepatanAngkut: '',
      melayani: '',
      dataPemeliharaan: '',
      hasilRiksaUjiSebelumnya: '',
      inspectionData: inspectionItems.map(item => ({
        id: item.id,
        status: '' as StatusType,
        notes: '',
        photoBefore: null,
        photoAfter: null,
        repairNote: '',
        isApproved: false,
      })),
      temuanKeselamatan: [{ id: 1, temuan: '', dokumentasi: null, kategori: '', solusi: '' }],
      temuanPermenaker: [{ id: 1, temuan: '', dokumentasi: null, kategori: '', solusi: '' }],
      temuanOperasional: [{ id: 1, temuan: '', dokumentasi: null, kategori: '', solusi: '' }],
      temuanAreaLain: [{ id: 1, temuan: '', dokumentasi: null, solusi: '' }],
      qcNote: '',
      attachment: null,
      attachmentName: null,
      qcVerification: { qcName: '', qcStatus: '' as QCStatusType, qcNote: '', verifiedAt: '' },
      submittedBy: '',
      submittedAt: '',
      status: 'draft',
      createdAt: new Date().toISOString(),
    };
  });

  // ============================================
  // HANDLER PEMELIHARAAN
  // ============================================
  const handleMaintenanceStatusChange = (unitIndex: number, itemId: number, status: MaintenanceStatusType) => {
    if (isReadOnly || !isQC) return;
    setFormData((prev: any) => {
      const newUnits = [...prev.units];
      const unit = newUnits[unitIndex];
      unit.maintenanceData = unit.maintenanceData.map((item: MaintenanceItemType) =>
        item.id === itemId ? { ...item, status, finding: status === 'Good' || status === 'N/A' ? '' : item.finding, isApproved: false } : item
      );
      return { ...prev, units: newUnits };
    });
  };

  const handleMaintenanceFindingChange = (unitIndex: number, itemId: number, finding: string) => {
    if (isReadOnly || !isQC) return;
    setFormData((prev: any) => {
      const newUnits = [...prev.units];
      const unit = newUnits[unitIndex];
      unit.maintenanceData = unit.maintenanceData.map((item: MaintenanceItemType) =>
        item.id === itemId ? { ...item, finding } : item
      );
      return { ...prev, units: newUnits };
    });
  };

  const handleMaintenancePhotoBefore = (unitIndex: number, itemId: number, event: React.ChangeEvent<HTMLInputElement>) => {
    if (isReadOnly || !isQC) return;
    const file = event.target.files?.[0];
    if (file) {
      const reader = new FileReader();
      reader.onloadend = () => {
        setFormData((prev: any) => {
          const newUnits = [...prev.units];
          const unit = newUnits[unitIndex];
          unit.maintenanceData = unit.maintenanceData.map((item: MaintenanceItemType) =>
            item.id === itemId ? { ...item, photoBefore: reader.result as string } : item
          );
          return { ...prev, units: newUnits };
        });
      };
      reader.readAsDataURL(file);
    }
  };

  const handleMaintenancePhotoAfter = (unitIndex: number, itemId: number, event: React.ChangeEvent<HTMLInputElement>) => {
    if (isReadOnly || !isMaintenance) return;
    const file = event.target.files?.[0];
    if (file) {
      const reader = new FileReader();
      reader.onloadend = () => {
        setFormData((prev: any) => {
          const newUnits = [...prev.units];
          const unit = newUnits[unitIndex];
          unit.maintenanceData = unit.maintenanceData.map((item: MaintenanceItemType) =>
            item.id === itemId ? { ...item, photoAfter: reader.result as string } : item
          );
          return { ...prev, units: newUnits };
        });
      };
      reader.readAsDataURL(file);
    }
  };

  const handleMaintenanceRepairNote = (unitIndex: number, itemId: number, note: string) => {
    if (isReadOnly || !isMaintenance) return;
    setFormData((prev: any) => {
      const newUnits = [...prev.units];
      const unit = newUnits[unitIndex];
      unit.maintenanceData = unit.maintenanceData.map((item: MaintenanceItemType) =>
        item.id === itemId ? { ...item, repairNote: note } : item
      );
      return { ...prev, units: newUnits };
    });
  };

  const handleMaintenanceApprove = (unitIndex: number, itemId: number) => {
    if (isReadOnly || !isQC || (formData.status !== 'maintenance_done' && formData.status !== 'revision')) return;
    setFormData((prev: any) => {
      const newUnits = [...prev.units];
      const unit = newUnits[unitIndex];
      unit.maintenanceData = unit.maintenanceData.map((item: MaintenanceItemType) =>
        item.id === itemId ? { ...item, isApproved: true } : item
      );
      return { ...prev, units: newUnits };
    });
  };

  const handleUnitDataChange = (unitIndex: number, field: string, value: string) => {
    if (isReadOnly || !isQC) return;
    setFormData((prev: any) => {
      const newUnits = [...prev.units];
      newUnits[unitIndex] = { ...newUnits[unitIndex], [field]: value };
      return { ...prev, units: newUnits };
    });
  };

  // Filter maintenance items berdasarkan search
  const getFilteredMaintenanceItems = (unitIndex: number) => {
    const unit = formData.units[unitIndex];
    if (!unit || !unit.maintenanceData) return [];
    if (!searchQuery) return unit.maintenanceData;
    const query = searchQuery.toLowerCase();
    return unit.maintenanceData.filter((item: MaintenanceItemType) => {
      const originalItem = maintenanceItems.find(i => i.id === item.id);
      return originalItem?.item.toLowerCase().includes(query) || 
             item.finding.toLowerCase().includes(query);
    });
  };

  // ============================================
  // HANDLER PEMERIKSAAN (sama seperti sebelumnya)
  // ============================================
  const filteredItems = inspectionItems.filter(item => {
    if (!searchQuery) return true;
    const query = searchQuery.toLowerCase();
    return item.item.toLowerCase().includes(query) || 
           item.ketentuan.toLowerCase().includes(query) ||
           item.category.toLowerCase().includes(query);
  });

  const totalScore = isPemeriksaan ? calculateTotalScore(formData.inspectionData) : 0;
  const totalItems = formData.inspectionData?.length || 0;
  const goodCount = formData.inspectionData?.filter((i: any) => i.status === 'Good' || i.status === 'N/A' || (i.status === 'Not Good' && i.isApproved)).length || 0;
  const notGoodCount = formData.inspectionData?.filter((i: any) => i.status === 'Not Good' && !i.isApproved).length || 0;

  // Handler untuk Temuan Negatif
  const handleTemuanAdd = (table: 'temuanKeselamatan' | 'temuanPermenaker' | 'temuanOperasional' | 'temuanAreaLain') => {
    if (isReadOnly || !isQC) return;
    setFormData((prev: any) => {
      const current = [...prev[table]];
      const maxId = current.reduce((max: number, item: any) => Math.max(max, item.id), 0);
      const newId = maxId + 1;
      let newItem: any;
      if (table === 'temuanAreaLain') {
        newItem = { id: newId, temuan: '', dokumentasi: null, solusi: '' };
      } else {
        newItem = { id: newId, temuan: '', dokumentasi: null, kategori: '', solusi: '' };
      }
      return { ...prev, [table]: [...current, newItem] };
    });
  };

  const handleTemuanRemove = (table: 'temuanKeselamatan' | 'temuanPermenaker' | 'temuanOperasional' | 'temuanAreaLain', id: number) => {
    if (isReadOnly || !isQC) return;
    setFormData((prev: any) => {
      const current = [...prev[table]];
      if (current.length <= 1) {
        alert('Minimal 1 baris temuan harus ada!');
        return prev;
      }
      return { ...prev, [table]: current.filter((item: any) => item.id !== id) };
    });
  };

  const handleTemuanChange = (
    table: 'temuanKeselamatan' | 'temuanPermenaker' | 'temuanOperasional' | 'temuanAreaLain',
    id: number,
    field: string,
    value: string
  ) => {
    if (isReadOnly || !isQC) return;
    setFormData((prev: any) => ({
      ...prev,
      [table]: prev[table].map((item: any) =>
        item.id === id ? { ...item, [field]: value } : item
      ),
    }));
  };

  const handleTemuanDokumentasi = (
    table: 'temuanKeselamatan' | 'temuanPermenaker' | 'temuanOperasional' | 'temuanAreaLain',
    id: number,
    event: React.ChangeEvent<HTMLInputElement>
  ) => {
    if (isReadOnly || !isQC) return;
    const file = event.target.files?.[0];
    if (file) {
      const reader = new FileReader();
      reader.onloadend = () => {
        setFormData((prev: any) => ({
          ...prev,
          [table]: prev[table].map((item: any) =>
            item.id === id ? { ...item, dokumentasi: reader.result as string } : item
          ),
        }));
      };
      reader.readAsDataURL(file);
    }
  };

  const removeTemuanDokumentasi = (
    table: 'temuanKeselamatan' | 'temuanPermenaker' | 'temuanOperasional' | 'temuanAreaLain',
    id: number
  ) => {
    if (isReadOnly || !isQC) return;
    setFormData((prev: any) => ({
      ...prev,
      [table]: prev[table].map((item: any) =>
        item.id === id ? { ...item, dokumentasi: null } : item
      ),
    }));
  };

  const handleDataUmumChange = (field: string, value: string) => {
    if (isReadOnly || !isQC) return;
    setFormData((prev: any) => ({ ...prev, [field]: value }));
  };

  const handleStatusChange = (id: number, status: StatusType) => {
    if (!isQC || isReadOnly || formData.status !== 'draft') return;
    setFormData((prev: any) => ({
      ...prev,
      inspectionData: prev.inspectionData.map((item: any) =>
        item.id === id ? { ...item, status, notes: status === 'Good' || status === 'N/A' ? '' : item.notes, isApproved: false } : item
      ),
    }));
  };

  const handleNotesChange = (id: number, notes: string) => {
    if (!isQC || isReadOnly || formData.status !== 'draft') return;
    setFormData((prev: any) => ({
      ...prev,
      inspectionData: prev.inspectionData.map((item: any) =>
        item.id === id ? { ...item, notes } : item
      ),
    }));
  };

  const handlePhotoBeforeChange = (id: number, event: React.ChangeEvent<HTMLInputElement>) => {
    if (!isQC || isReadOnly || formData.status !== 'draft') return;
    const file = event.target.files?.[0];
    if (file) {
      const reader = new FileReader();
      reader.onloadend = () => {
        setFormData((prev: any) => ({
          ...prev,
          inspectionData: prev.inspectionData.map((item: any) =>
            item.id === id ? { ...item, photoBefore: reader.result as string } : item
          ),
        }));
      };
      reader.readAsDataURL(file);
    }
  };

  const handlePhotoAfterChange = (id: number, event: React.ChangeEvent<HTMLInputElement>) => {
    if (!isMaintenance || isReadOnly) return;
    const file = event.target.files?.[0];
    if (file) {
      const reader = new FileReader();
      reader.onloadend = () => {
        setFormData((prev: any) => ({
          ...prev,
          inspectionData: prev.inspectionData.map((item: any) =>
            item.id === id ? { ...item, photoAfter: reader.result as string } : item
          ),
        }));
      };
      reader.readAsDataURL(file);
    }
  };

  const handleRepairNoteChange = (id: number, note: string) => {
    if (!isMaintenance || isReadOnly) return;
    setFormData((prev: any) => ({
      ...prev,
      inspectionData: prev.inspectionData.map((item: any) =>
        item.id === id ? { ...item, repairNote: note } : item
      ),
    }));
  };

  const handleApproveItem = (id: number) => {
    if (!isQC || isReadOnly || (formData.status !== 'maintenance_done' && formData.status !== 'revision')) return;
    const item = formData.inspectionData.find((i: any) => i.id === id);
    if (!item || item.status !== 'Not Good' || item.isApproved) return;
    setFormData((prev: any) => ({
      ...prev,
      inspectionData: prev.inspectionData.map((item: any) =>
        item.id === id ? { ...item, isApproved: true } : item
      ),
    }));
  };

  const handleTeknisiNameChange = (name: string) => {
    if (!isMaintenance || isReadOnly) return;
    setFormData((prev: any) => ({ ...prev, teknisiName: name }));
  };

  const handleAttachmentChange = (event: React.ChangeEvent<HTMLInputElement>) => {
    if (isReadOnly) return;
    const file = event.target.files?.[0];
    if (file) {
      const reader = new FileReader();
      reader.onloadend = () => {
        setFormData((prev: any) => ({
          ...prev,
          attachment: reader.result as string,
          attachmentName: file.name
        }));
      };
      reader.readAsDataURL(file);
    }
  };

  const handleDownloadAttachment = () => {
    if (formData.attachment) {
      const link = document.createElement('a');
      link.href = formData.attachment;
      link.download = formData.attachmentName || 'attachment';
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
    }
  };

  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    onSave(formData);
  };

  const handleKirim = (action: 'submit' | 'approve' | 'revision') => {
    setIsSubmitting(true);
    let updatedData: any;
    
    if (isQC && formData.status === 'draft') {
      const requiredFields = ['namaGedung', 'alamat', 'jenisElevator', 'merkTipe', 'elevatorNo', 'tanggalRiksaUji'];
      const emptyFields = requiredFields.filter(f => !formData[f] || formData[f].trim() === '');
      if (emptyFields.length > 0) {
        alert(`⚠️ Data Umum belum lengkap: ${emptyFields.join(', ')}`);
        setIsSubmitting(false);
        return;
      }

      if (isPemeriksaan) {
        const emptyItems = formData.inspectionData.filter((item: any) => !item.status);
        if (emptyItems.length > 0) {
          alert(`⚠️ ${emptyItems.length} item belum diisi status.`);
          setIsSubmitting(false);
          return;
        }

        const invalidItems = formData.inspectionData.filter((item: any) => 
          item.status === 'Not Good' && (!item.notes || item.notes.trim() === '')
        );
        if (invalidItems.length > 0) {
          alert(`⚠️ ${invalidItems.length} item Not Good tanpa temuan (Notes).`);
          setIsSubmitting(false);
          return;
        }
      } else {
        // Validasi Pemeliharaan
        const unit = formData.units[0];
        if (!unit.unitNumber || !unit.projectCode || !unit.customerName) {
          alert('⚠️ Data Unit belum lengkap!');
          setIsSubmitting(false);
          return;
        }
        const emptyItems = unit.maintenanceData.filter((item: any) => !item.status);
        if (emptyItems.length > 0) {
          alert(`⚠️ ${emptyItems.length} item pemeliharaan belum diisi status.`);
          setIsSubmitting(false);
          return;
        }
        const invalidItems = unit.maintenanceData.filter((item: any) => 
          item.status === 'Not Good' && (!item.finding || item.finding.trim() === '')
        );
        if (invalidItems.length > 0) {
          alert(`⚠️ ${invalidItems.length} item Not Good tanpa temuan (Finding).`);
          setIsSubmitting(false);
          return;
        }
      }

      updatedData = { ...formData, submittedBy: 'qc', submittedAt: new Date().toLocaleString('id-ID'), status: 'qc_approved' };
      setTimeout(() => { onSave(updatedData); setIsSubmitting(false); alert('Data dikirim ke Maintenance!'); }, 500);
      return;
    } 
    
    if (isQC && formData.status === 'maintenance_done') {
      if (action === 'approve') {
        let unapproved: any[] = [];
        if (isPemeriksaan) {
          unapproved = formData.inspectionData.filter((item: any) => item.status === 'Not Good' && !item.isApproved);
        } else {
          const unit = formData.units[0];
          unapproved = unit.maintenanceData.filter((item: any) => item.status === 'Not Good' && !item.isApproved);
        }
        if (unapproved.length > 0) {
          alert(`⚠️ ${unapproved.length} item Not Good belum di-approve!`);
          setIsSubmitting(false);
          return;
        }
        updatedData = { ...formData, qcVerification: { ...formData.qcVerification, verifiedAt: new Date().toLocaleString('id-ID') }, status: 'approved' };
        setTimeout(() => { onSave(updatedData); setIsSubmitting(false); alert('Laporan FINAL APPROVED!'); }, 500);
      } else if (action === 'revision') {
        updatedData = { ...formData, qcVerification: { ...formData.qcVerification, verifiedAt: new Date().toLocaleString('id-ID'), qcStatus: 'Revision Required' }, status: 'revision' };
        setTimeout(() => { onSave(updatedData); setIsSubmitting(false); alert('REVISION! Dikirim ke Maintenance.'); }, 500);
      }
      return;
    } 
    
    if (isMaintenance && (formData.status === 'qc_approved' || formData.status === 'revision')) {
      if (!formData.teknisiName || formData.teknisiName.trim() === '') {
        alert('Silakan isi nama teknisi!');
        setIsSubmitting(false);
        return;
      }
      let invalid: any[] = [];
      if (isPemeriksaan) {
        invalid = formData.inspectionData.filter((item: any) => 
          item.status === 'Not Good' && (!item.photoAfter || !item.repairNote || item.repairNote.trim() === '')
        );
      } else {
        const unit = formData.units[0];
        invalid = unit.maintenanceData.filter((item: any) => 
          item.status === 'Not Good' && (!item.photoAfter || !item.repairNote || item.repairNote.trim() === '')
        );
      }
      if (invalid.length > 0) {
        alert(`⚠️ ${invalid.length} item Not Good tanpa foto/catatan perbaikan.`);
        setIsSubmitting(false);
        return;
      }
      updatedData = { ...formData, submittedBy: 'maintenance', submittedAt: new Date().toLocaleString('id-ID'), status: 'maintenance_done' };
      setTimeout(() => { onSave(updatedData); setIsSubmitting(false); alert(`Perbaikan selesai oleh ${formData.teknisiName}!`); }, 500);
      return;
    }
    
    alert('Status tidak sesuai.');
    setIsSubmitting(false);
  };

  const canEdit = !isReadOnly;
  const isApproved = formData.status === 'approved' || formData.status === 'qc_approved';
  let allNotGoodApproved = true;
  if (isPemeriksaan) {
    allNotGoodApproved = formData.inspectionData.filter((item: any) => item.status === 'Not Good').every((item: any) => item.isApproved);
  } else {
    const unit = formData.units[0];
    if (unit && unit.maintenanceData) {
      allNotGoodApproved = unit.maintenanceData.filter((item: any) => item.status === 'Not Good').every((item: any) => item.isApproved);
    }
  }
  const canFinalApprove = formData.status === 'maintenance_done' && allNotGoodApproved;

  const renderTeknisiInput = () => {
    if (isMaintenance && !isReadOnly && (formData.status === 'qc_approved' || formData.status === 'revision')) {
      return (
        <div className="bg-amber-50 p-4 rounded-xl border border-amber-200 mb-6">
          <label className="block text-sm font-medium text-gray-700 mb-1">Nama Teknisi Pelaksana <span className="text-red-500">*</span></label>
          <input
            type="text"
            value={formData.teknisiName || ''}
            onChange={(e) => handleTeknisiNameChange(e.target.value)}
            placeholder="Masukkan nama teknisi"
            className="w-full px-4 py-2.5 border border-amber-300 rounded-xl focus:outline-none focus:ring-2 focus:ring-amber-500 bg-white"
          />
          <p className="text-xs text-gray-400 mt-1">Isi nama teknisi pelaksana perbaikan</p>
        </div>
      );
    }
    if (formData.teknisiName) {
      return (
        <div className="bg-blue-50 p-3 rounded-xl border border-blue-200 mb-4">
          <span className="text-sm font-medium text-blue-700">Teknisi: </span>
          <span className="text-blue-600 font-semibold">{formData.teknisiName}</span>
        </div>
      );
    }
    return null;
  };

  // ============================================
  // RENDER PEMELIHARAAN
  // ============================================
  if (isPemeliharaan) {
    const unitIndex = 0;
    const unit = formData.units?.[unitIndex];
    if (!unit) return <div className="text-center py-10 text-gray-400">Data unit tidak ditemukan</div>;

    const filteredMaintenanceItems = getFilteredMaintenanceItems(unitIndex);
    const totalMaintenanceItems = unit.maintenanceData?.length || 0;
    const goodMaintenanceCount = unit.maintenanceData?.filter((i: any) => i.status === 'Good' || i.status === 'N/A' || (i.status === 'Not Good' && i.isApproved)).length || 0;
    const notGoodMaintenanceCount = unit.maintenanceData?.filter((i: any) => i.status === 'Not Good' && !i.isApproved).length || 0;

    return (
      <form onSubmit={handleSubmit} className="space-y-6">
        {/* HEADER */}
        <div className="flex flex-wrap justify-between items-center gap-2">
          <button type="button" onClick={onCancel} className="bg-gray-100 hover:bg-gray-200 text-gray-600 px-4 py-2 rounded-xl text-sm font-medium transition">Kembali</button>
          <div className="flex flex-wrap gap-2">
            {canEdit && <button type="button" onClick={() => onDelete?.(formData.id)} className="bg-red-500 hover:bg-red-600 text-white px-4 py-2 rounded-xl text-sm transition">Hapus</button>}
            <button type="submit" className="bg-blue-600 hover:bg-blue-700 text-white px-5 py-2 rounded-xl text-sm font-medium transition">Simpan</button>
            {isQC && formData.status === 'draft' && (
              <button type="button" onClick={() => handleKirim('submit')} disabled={isSubmitting} className="bg-blue-600 hover:bg-blue-700 text-white px-5 py-2 rounded-xl text-sm font-medium transition">
                {isSubmitting ? 'Memproses...' : 'Submit ke Maintenance'}
              </button>
            )}
            {isQC && formData.status === 'maintenance_done' && (
              <>
                <button type="button" onClick={() => handleKirim('revision')} disabled={isSubmitting} className="bg-amber-500 hover:bg-amber-600 text-white px-5 py-2 rounded-xl text-sm font-medium transition">Revisi</button>
                <button type="button" onClick={() => { if (!canFinalApprove) { alert('Approve semua Not Good dulu!'); return; } handleKirim('approve'); }} disabled={!canFinalApprove || isSubmitting} className={`px-5 py-2 rounded-xl text-sm font-medium transition ${canFinalApprove && !isSubmitting ? 'bg-emerald-600 hover:bg-emerald-700 text-white' : 'bg-gray-300 text-gray-500 cursor-not-allowed'}`}>Final Approve</button>
              </>
            )}
            {(isMaintenance && (formData.status === 'qc_approved' || formData.status === 'revision')) && (
              <button type="button" onClick={() => { if (!formData.teknisiName) { alert('Isi nama teknisi!'); return; } handleKirim('submit'); }} className="bg-orange-500 hover:bg-orange-600 text-white px-5 py-2 rounded-xl text-sm font-medium transition">Selesai Perbaikan</button>
            )}
            {isQC && isApproved && onPrintPDF && (
              <button type="button" onClick={onPrintPDF} className="bg-rose-500 hover:bg-rose-600 text-white px-5 py-2 rounded-xl text-sm font-medium transition">PDF</button>
            )}
          </div>
        </div>

        {/* PRINT HEADER */}
        <div className="text-center border-b pb-4 mb-6 print:block hidden">
          <div className="flex justify-center mb-2">
            <Image src="/logo_louser_2022_1696999044 (1).png" alt="Logo Louser" width={150} height={50} className="object-contain" />
          </div>
          <p className="text-sm text-gray-500">Elevator Quality Control System</p>
          <p className="text-xs text-gray-400">Laporan Pemeliharaan Elevator</p>
        </div>

        {/* HEADER LAPORAN */}
        <div className="text-center mb-6 print:block">
          <h1 className="text-2xl font-bold text-gray-800">LAPORAN PEMELIHARAAN ELEVATOR</h1>
          <p className="text-sm text-gray-500">PT Louserindo Megah Permai</p>
        </div>

        {renderTeknisiInput()}

        {/* DATA UMUM */}
        <section className="bg-gray-50/50 rounded-xl p-6 mb-6 border border-gray-100">
          <h2 className="text-sm font-semibold text-gray-600 mb-4 border-b pb-2">DATA UMUM</h2>
          <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
            <div>
              <label className="block text-xs font-medium text-gray-500 uppercase tracking-wider mb-1">No Unit</label>
              <input
                type="text"
                value={unit.unitNumber || ''}
                onChange={(e) => handleUnitDataChange(unitIndex, 'unitNumber', e.target.value)}
                className="w-full px-3 py-2 bg-white border border-gray-200 rounded-xl focus:outline-none focus:ring-2 focus:ring-blue-500 text-sm"
                disabled={isReadOnly || !isQC}
                placeholder="Contoh: E-001"
              />
            </div>
            <div>
              <label className="block text-xs font-medium text-gray-500 uppercase tracking-wider mb-1">Kode Proyek</label>
              <input
                type="text"
                value={unit.projectCode || ''}
                onChange={(e) => handleUnitDataChange(unitIndex, 'projectCode', e.target.value)}
                className="w-full px-3 py-2 bg-white border border-gray-200 rounded-xl focus:outline-none focus:ring-2 focus:ring-blue-500 text-sm"
                disabled={isReadOnly || !isQC}
                placeholder="Kode proyek"
              />
            </div>
            <div>
              <label className="block text-xs font-medium text-gray-500 uppercase tracking-wider mb-1">Pelanggan</label>
              <input
                type="text"
                value={unit.customerName || ''}
                onChange={(e) => handleUnitDataChange(unitIndex, 'customerName', e.target.value)}
                className="w-full px-3 py-2 bg-white border border-gray-200 rounded-xl focus:outline-none focus:ring-2 focus:ring-blue-500 text-sm"
                disabled={isReadOnly || !isQC}
                placeholder="Nama pelanggan"
              />
            </div>
            <div>
              <label className="block text-xs font-medium text-gray-500 uppercase tracking-wider mb-1">Lokasi Gedung</label>
              <input
                type="text"
                value={unit.buildingLocation || ''}
                onChange={(e) => handleUnitDataChange(unitIndex, 'buildingLocation', e.target.value)}
                className="w-full px-3 py-2 bg-white border border-gray-200 rounded-xl focus:outline-none focus:ring-2 focus:ring-blue-500 text-sm"
                disabled={isReadOnly || !isQC}
                placeholder="Alamat gedung"
              />
            </div>
            <div>
              <label className="block text-xs font-medium text-gray-500 uppercase tracking-wider mb-1">Tipe Elevator</label>
              <input
                type="text"
                value={unit.elevatorType || ''}
                onChange={(e) => handleUnitDataChange(unitIndex, 'elevatorType', e.target.value)}
                className="w-full px-3 py-2 bg-white border border-gray-200 rounded-xl focus:outline-none focus:ring-2 focus:ring-blue-500 text-sm"
                disabled={isReadOnly || !isQC}
                placeholder="Contoh: Penumpang"
              />
            </div>
            <div>
              <label className="block text-xs font-medium text-gray-500 uppercase tracking-wider mb-1">Brand Elevator</label>
              <input
                type="text"
                value={unit.elevatorBrand || ''}
                onChange={(e) => handleUnitDataChange(unitIndex, 'elevatorBrand', e.target.value)}
                className="w-full px-3 py-2 bg-white border border-gray-200 rounded-xl focus:outline-none focus:ring-2 focus:ring-blue-500 text-sm"
                disabled={isReadOnly || !isQC}
                placeholder="Contoh: Mitsubishi"
              />
            </div>
            <div>
              <label className="block text-xs font-medium text-gray-500 uppercase tracking-wider mb-1">Model</label>
              <input
                type="text"
                value={unit.elevatorModel || ''}
                onChange={(e) => handleUnitDataChange(unitIndex, 'elevatorModel', e.target.value)}
                className="w-full px-3 py-2 bg-white border border-gray-200 rounded-xl focus:outline-none focus:ring-2 focus:ring-blue-500 text-sm"
                disabled={isReadOnly || !isQC}
                placeholder="Model elevator"
              />
            </div>
            <div>
              <label className="block text-xs font-medium text-gray-500 uppercase tracking-wider mb-1">Kapasitas</label>
              <input
                type="text"
                value={unit.capacity || ''}
                onChange={(e) => handleUnitDataChange(unitIndex, 'capacity', e.target.value)}
                className="w-full px-3 py-2 bg-white border border-gray-200 rounded-xl focus:outline-none focus:ring-2 focus:ring-blue-500 text-sm"
                disabled={isReadOnly || !isQC}
                placeholder="Contoh: 1000 kg"
              />
            </div>
            <div>
              <label className="block text-xs font-medium text-gray-500 uppercase tracking-wider mb-1">Kecepatan</label>
              <input
                type="text"
                value={unit.speed || ''}
                onChange={(e) => handleUnitDataChange(unitIndex, 'speed', e.target.value)}
                className="w-full px-3 py-2 bg-white border border-gray-200 rounded-xl focus:outline-none focus:ring-2 focus:ring-blue-500 text-sm"
                disabled={isReadOnly || !isQC}
                placeholder="Contoh: 60 m/menit"
              />
            </div>
            <div>
              <label className="block text-xs font-medium text-gray-500 uppercase tracking-wider mb-1">Tanggal Inspeksi</label>
              <input
                type="date"
                value={unit.inspectionDate || ''}
                onChange={(e) => handleUnitDataChange(unitIndex, 'inspectionDate', e.target.value)}
                className="w-full px-3 py-2 bg-white border border-gray-200 rounded-xl focus:outline-none focus:ring-2 focus:ring-blue-500 text-sm"
                disabled={isReadOnly || !isQC}
              />
            </div>
            <div>
              <label className="block text-xs font-medium text-gray-500 uppercase tracking-wider mb-1">Nama QC</label>
              <input
                type="text"
                value={unit.qcName || ''}
                onChange={(e) => handleUnitDataChange(unitIndex, 'qcName', e.target.value)}
                className="w-full px-3 py-2 bg-white border border-gray-200 rounded-xl focus:outline-none focus:ring-2 focus:ring-blue-500 text-sm"
                disabled={isReadOnly || !isQC}
                placeholder="Nama QC"
              />
            </div>
          </div>
        </section>

        {/* STATUS BAR */}
        <div className="flex flex-wrap items-center gap-3 bg-gray-50/50 p-3 rounded-xl mb-6">
          <span className="text-xs font-medium text-gray-500">Status:</span>
          <span className={`px-3 py-1 rounded-full text-xs font-medium ${formData.status === 'draft' ? 'bg-gray-200 text-gray-600' : formData.status === 'qc_approved' ? 'bg-blue-100 text-blue-700' : formData.status === 'revision' ? 'bg-amber-100 text-amber-700' : formData.status === 'maintenance_done' ? 'bg-orange-100 text-orange-700' : 'bg-emerald-100 text-emerald-700'}`}>
            {formData.status === 'draft' ? 'Draft' : formData.status === 'qc_approved' ? 'Maintenance' : formData.status === 'revision' ? 'REVISION' : formData.status === 'maintenance_done' ? 'Selesai Perbaikan' : 'FINAL APPROVED'}
          </span>
          {formData.submittedAt && <span className="text-xs text-gray-400">Dikirim: {formData.submittedAt}</span>}
          {formData.qcVerification?.verifiedAt && <span className="text-xs text-gray-400">Diverifikasi: {formData.qcVerification.verifiedAt}</span>}
        </div>

        {/* TABEL PEMELIHARAAN */}
        <section className="bg-white rounded-xl border border-gray-100 overflow-hidden mb-6">
          <div className="bg-emerald-700 px-6 py-3 flex flex-wrap justify-between items-center">
            <h2 className="text-sm font-semibold text-white">DATA PEMELIHARAAN</h2>
            <div className="flex items-center gap-4">
              <span className="text-xs text-white bg-emerald-600 px-3 py-1 rounded-full">{totalMaintenanceItems} Item</span>
              <div className="relative">
                <input
                  type="text"
                  value={searchQuery}
                  onChange={(e) => setSearchQuery(e.target.value)}
                  placeholder="🔍 Cari item pemeliharaan..."
                  className="px-3 py-1.5 bg-white/20 border border-white/30 rounded-lg text-white placeholder-white/70 text-xs focus:outline-none focus:ring-2 focus:ring-white/50 w-48"
                />
              </div>
            </div>
          </div>

          <div className="overflow-x-auto" style={{ maxHeight: '500px', overflowY: 'auto' }}>
            <table className="w-full text-xs border-collapse" style={{ minWidth: '1100px' }}>
              <thead className="sticky top-0 z-10">
                <tr className="bg-gray-100">
                  <th className="p-2 text-center text-gray-600 font-semibold border-b border-gray-300 w-10">NO</th>
                  <th className="p-2 text-left text-gray-600 font-semibold border-b border-gray-300 min-w-[200px]">Item Pemeliharaan</th>
                  <th className="p-2 text-center text-gray-600 font-semibold border-b border-gray-300 w-28">Status</th>
                  <th className="p-2 text-left text-gray-600 font-semibold border-b border-gray-300 min-w-[200px]">Temuan</th>
                  <th className="p-2 text-center text-gray-600 font-semibold border-b border-gray-300 w-32">Foto Sebelum</th>
                  <th className="p-2 text-center text-gray-600 font-semibold border-b border-gray-300 w-32">Foto Setelah</th>
                  <th className="p-2 text-left text-gray-600 font-semibold border-b border-gray-300 min-w-[150px]">Catatan Perbaikan</th>
                  <th className="p-2 text-center text-gray-600 font-semibold border-b border-gray-300 w-16">Approve</th>
                </tr>
              </thead>
              <tbody>
                {filteredMaintenanceItems.map((item: MaintenanceItemType) => {
                  const originalItem = maintenanceItems.find(i => i.id === item.id);
                  if (!originalItem) return null;

                  const isNotGood = item.status === 'Not Good';
                  const isApproved = item.isApproved;
                  const isEditable = isQC && !isReadOnly && formData.status === 'draft';
                  const isMaintEditable = isMaintenance && !isReadOnly && (formData.status === 'qc_approved' || formData.status === 'revision');
                  const canApprove = isQC && !isReadOnly && (formData.status === 'maintenance_done' || formData.status === 'revision') && isNotGood && !isApproved;

                  return (
                    <tr key={item.id} className="border-b border-gray-100 hover:bg-gray-50/50 transition">
                      <td className="p-2 text-center text-gray-400 font-medium">{item.id}</td>
                      <td className="p-2 text-gray-700 align-top break-words whitespace-normal leading-tight">{originalItem.item}</td>
                      <td className="p-2 align-top text-center">
                        {isEditable ? (
                          <div className="flex flex-col gap-0.5 text-[10px]">
                            <label className="flex items-center gap-1"><input type="radio" name={`m-status-${item.id}`} value="Good" checked={item.status === 'Good'} onChange={() => handleMaintenanceStatusChange(unitIndex, item.id, 'Good')} className="w-3 h-3"/> Good</label>
                            <label className="flex items-center gap-1"><input type="radio" name={`m-status-${item.id}`} value="Not Good" checked={item.status === 'Not Good'} onChange={() => handleMaintenanceStatusChange(unitIndex, item.id, 'Not Good')} className="w-3 h-3"/> Not Good</label>
                            <label className="flex items-center gap-1"><input type="radio" name={`m-status-${item.id}`} value="N/A" checked={item.status === 'N/A'} onChange={() => handleMaintenanceStatusChange(unitIndex, item.id, 'N/A')} className="w-3 h-3"/> N/A</label>
                          </div>
                        ) : (
                          <span className={`px-2 py-0.5 rounded-full text-[10px] font-medium ${
                            item.status === 'Good' ? 'bg-emerald-100 text-emerald-700' : 
                            item.status === 'N/A' ? 'bg-gray-200 text-gray-500' : 
                            item.status === 'Not Good' && isApproved ? 'bg-emerald-200 text-emerald-800' : 
                            item.status === 'Not Good' && !isApproved ? 'bg-rose-100 text-rose-700' : 
                            'bg-gray-100 text-gray-400'
                          }`}>{item.status || '-'}</span>
                        )}
                      </td>
                      <td className="p-2 align-top">
                        {isEditable ? (
                          <textarea 
                            value={item.finding} 
                            onChange={(e) => handleMaintenanceFindingChange(unitIndex, item.id, e.target.value)} 
                            disabled={item.status === 'Good' || item.status === 'N/A'} 
                            placeholder="Temuan..." 
                            className={`w-full px-2 py-1 border rounded-lg text-[10px] ${isNotGood && !item.finding ? 'border-rose-300 bg-rose-50' : 'border-gray-200'} ${item.status === 'Good' || item.status === 'N/A' ? 'bg-gray-50' : 'bg-white'}`} 
                            rows={2}
                          />
                        ) : (
                          <span className="text-gray-600 break-words whitespace-normal text-[10px]">{item.finding || '-'}</span>
                        )}
                      </td>
                      <td className="p-2 align-top text-center">
                        {item.photoBefore ? (
                          <div className="flex flex-col items-center gap-1">
                            <img src={item.photoBefore} alt="Sebelum" className="w-16 h-16 object-cover rounded border" />
                            {isEditable && (
                              <button
                                type="button"
                                onClick={() => {
                                  if (confirm('Hapus foto sebelum?')) {
                                    setFormData((prev: any) => {
                                      const newUnits = [...prev.units];
                                      const unit = newUnits[unitIndex];
                                      unit.maintenanceData = unit.maintenanceData.map((i: any) =>
                                        i.id === item.id ? { ...i, photoBefore: null } : i
                                      );
                                      return { ...prev, units: newUnits };
                                    });
                                  }
                                }}
                                className="text-red-400 hover:text-red-600 text-[10px]"
                              >
                                ✕ Hapus
                              </button>
                            )}
                          </div>
                        ) : (
                          isEditable && (
                            <input
                              type="file"
                              accept="image/*"
                              onChange={(e) => handleMaintenancePhotoBefore(unitIndex, item.id, e)}
                              className="text-[10px] w-full"
                            />
                          )
                        )}
                        {!item.photoBefore && !isEditable && <span className="text-gray-400 text-[10px]">-</span>}
                      </td>
                      <td className="p-2 align-top text-center">
                        {item.photoAfter ? (
                          <div className="flex flex-col items-center gap-1">
                            <img src={item.photoAfter} alt="Setelah" className="w-16 h-16 object-cover rounded border" />
                            {isMaintEditable && (
                              <button
                                type="button"
                                onClick={() => {
                                  if (confirm('Hapus foto setelah?')) {
                                    setFormData((prev: any) => {
                                      const newUnits = [...prev.units];
                                      const unit = newUnits[unitIndex];
                                      unit.maintenanceData = unit.maintenanceData.map((i: any) =>
                                        i.id === item.id ? { ...i, photoAfter: null } : i
                                      );
                                      return { ...prev, units: newUnits };
                                    });
                                  }
                                }}
                                className="text-red-400 hover:text-red-600 text-[10px]"
                              >
                                ✕ Hapus
                              </button>
                            )}
                          </div>
                        ) : (
                          isMaintEditable && isNotGood && (
                            <input
                              type="file"
                              accept="image/*"
                              onChange={(e) => handleMaintenancePhotoAfter(unitIndex, item.id, e)}
                              className="text-[10px] w-full"
                            />
                          )
                        )}
                        {!item.photoAfter && !isMaintEditable && <span className="text-gray-400 text-[10px]">-</span>}
                      </td>
                      <td className="p-2 align-top">
                        {isMaintEditable && isNotGood ? (
                          <textarea 
                            value={item.repairNote} 
                            onChange={(e) => handleMaintenanceRepairNote(unitIndex, item.id, e.target.value)} 
                            placeholder="Catatan perbaikan..." 
                            className="w-full px-2 py-1 border border-gray-200 rounded-lg text-[10px] bg-white" 
                            rows={2}
                          />
                        ) : (
                          <span className="text-gray-600 break-words whitespace-normal text-[10px]">{item.repairNote || '-'}</span>
                        )}
                      </td>
                      <td className="p-2 align-top text-center">
                        {canApprove ? (
                          <button
                            type="button"
                            onClick={() => handleMaintenanceApprove(unitIndex, item.id)}
                            className="bg-emerald-500 hover:bg-emerald-600 text-white px-3 py-1 rounded-lg text-[10px] font-medium transition"
                          >
                            Approve
                          </button>
                        ) : isApproved ? (
                          <span className="text-emerald-600 text-[10px] font-medium">✅ Approved</span>
                        ) : (
                          <span className="text-gray-400 text-[10px]">-</span>
                        )}
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>

          {/* REKAP PEMELIHARAAN */}
          <div className="bg-gray-50 p-4 border-t border-gray-200">
            <div className="grid grid-cols-2 md:grid-cols-3 gap-4">
              <div className="bg-white p-3 rounded-xl text-center shadow-sm">
                <div className="text-xs text-gray-400">Total Item</div>
                <div className="text-xl font-bold text-gray-700">{totalMaintenanceItems}</div>
              </div>
              <div className="bg-emerald-50 p-3 rounded-xl text-center">
                <div className="text-xs text-emerald-500">Good + N/A + Approved</div>
                <div className="text-xl font-bold text-emerald-700">{goodMaintenanceCount}</div>
              </div>
              <div className="bg-rose-50 p-3 rounded-xl text-center">
                <div className="text-xs text-rose-500">Not Good (Belum Approve)</div>
                <div className="text-xl font-bold text-rose-700">{notGoodMaintenanceCount}</div>
              </div>
            </div>
          </div>
        </section>

        {/* CATATAN & ATTACHMENT */}
        <section className="bg-white rounded-xl border border-gray-100 p-6 mb-6">
          <h2 className="text-sm font-semibold text-gray-600 mb-4">Catatan & Attachment</h2>
          
          {isQC && !isReadOnly && (
            <div className="mb-4">
              <label className="block text-xs font-medium text-gray-500 uppercase tracking-wider mb-1">Catatan QC</label>
              <textarea
                value={formData.qcNote || ''}
                onChange={(e) => setFormData((prev: any) => ({ ...prev, qcNote: e.target.value }))}
                placeholder="Tambahkan catatan di sini..."
                className="w-full px-3 py-2 bg-white border border-gray-200 rounded-xl focus:outline-none focus:ring-2 focus:ring-blue-500 text-sm"
                rows={3}
                disabled={isReadOnly}
              />
            </div>
          )}
          {formData.qcNote && isReadOnly && (
            <div className="mb-4 p-3 bg-blue-50 border border-blue-200 rounded-xl text-gray-700 text-sm">
              <span className="font-medium text-blue-700">Catatan QC:</span> {formData.qcNote}
            </div>
          )}

          <div className="border-2 border-dashed border-gray-200 rounded-xl p-6 text-center hover:border-blue-400 transition">
            {isQC && !isReadOnly ? (
              <input
                type="file"
                accept=".pdf,.doc,.docx,.xlsx,.xls,.jpg,.jpeg,.png"
                onChange={handleAttachmentChange}
                className="text-sm text-gray-500 file:mr-3 file:py-2 file:px-4 file:rounded-xl file:border-0 file:text-sm file:font-medium file:bg-blue-50 file:text-blue-700 hover:file:bg-blue-100"
              />
            ) : formData.attachment ? (
              <div className="flex items-center justify-center gap-4 flex-wrap">
                <span className="text-sm text-gray-600">📄 {formData.attachmentName}</span>
                <button type="button" onClick={handleDownloadAttachment} className="bg-blue-600 hover:bg-blue-700 text-white px-4 py-2 rounded-xl text-sm transition">Download</button>
                <a href={formData.attachment} target="_blank" rel="noopener" className="bg-emerald-600 hover:bg-emerald-700 text-white px-4 py-2 rounded-xl text-sm transition">Lihat</a>
              </div>
            ) : (
              <p className="text-sm text-gray-400">Tidak ada file attachment</p>
            )}
          </div>
        </section>

        {/* FOOTER */}
        <div className="border-t pt-4">
          <div className="flex flex-wrap justify-between items-start gap-4">
            <div className="text-left">
              <p className="text-xs text-gray-400">Dicetak: {new Date().toLocaleDateString('id-ID', { day: '2-digit', month: 'long', year: 'numeric' })}</p>
              <p className="text-xs text-gray-400">© PT Louserindo Megah Permai</p>
            </div>
            <div className="text-center">
              <p className="text-xs text-gray-500 mb-1">Tanda Tangan</p>
              <div className="w-40 h-12 border-b-2 border-gray-300"></div>
              <p className="text-[10px] text-gray-400 mt-1">(_____________________)</p>
            </div>
          </div>
          <div className="text-right text-xs text-gray-400 border-t pt-4 mt-4">
            <p>Dibuat: {new Date(formData.createdAt).toLocaleString('id-ID')}</p>
            {formData.submittedAt && <p>Dikirim: {formData.submittedAt}</p>}
            {formData.qcVerification?.verifiedAt && <p>Diverifikasi: {formData.qcVerification.verifiedAt}</p>}
          </div>
        </div>
      </form>
    );
  }

  // ============================================
  // RENDER PEMERIKSAAN (tetap sama seperti sebelumnya)
  // ============================================
  return (
    <form onSubmit={handleSubmit} className="space-y-6">
      {/* HEADER BUTTON */}
      <div className="flex flex-wrap justify-between items-center gap-2">
        <button type="button" onClick={onCancel} className="bg-gray-100 hover:bg-gray-200 text-gray-600 px-4 py-2 rounded-xl text-sm font-medium transition">Kembali</button>
        <div className="flex flex-wrap gap-2">
          {canEdit && <button type="button" onClick={() => onDelete?.(formData.id)} className="bg-red-500 hover:bg-red-600 text-white px-4 py-2 rounded-xl text-sm transition">Hapus</button>}
          {isQC && formData.status === 'draft' && (
            <button type="button" onClick={() => handleKirim('submit')} disabled={isSubmitting} className="bg-blue-600 hover:bg-blue-700 text-white px-5 py-2 rounded-xl text-sm font-medium transition">
              {isSubmitting ? 'Memproses...' : 'Submit ke Maintenance'}
            </button>
          )}
          {isQC && formData.status === 'maintenance_done' && (
            <>
              <button type="button" onClick={() => handleKirim('revision')} disabled={isSubmitting} className="bg-amber-500 hover:bg-amber-600 text-white px-5 py-2 rounded-xl text-sm font-medium transition">Revisi</button>
              <button type="button" onClick={() => { if (!canFinalApprove) { alert('Approve semua Not Good dulu!'); return; } handleKirim('approve'); }} disabled={!canFinalApprove || isSubmitting} className={`px-5 py-2 rounded-xl text-sm font-medium transition ${canFinalApprove && !isSubmitting ? 'bg-emerald-600 hover:bg-emerald-700 text-white' : 'bg-gray-300 text-gray-500 cursor-not-allowed'}`}>Final Approve</button>
            </>
          )}
          {(isMaintenance && (formData.status === 'qc_approved' || formData.status === 'revision')) && (
            <button type="button" onClick={() => { if (!formData.teknisiName) { alert('Isi nama teknisi!'); return; } handleKirim('submit'); }} className="bg-orange-500 hover:bg-orange-600 text-white px-5 py-2 rounded-xl text-sm font-medium transition">Selesai Perbaikan</button>
          )}
          {isQC && isApproved && onPrintPDF && (
            <button type="button" onClick={onPrintPDF} className="bg-rose-500 hover:bg-rose-600 text-white px-5 py-2 rounded-xl text-sm font-medium transition">PDF</button>
          )}
        </div>
      </div>

      {/* PRINT HEADER */}
      <div className="text-center border-b pb-4 mb-6 print:block hidden">
        <div className="flex justify-center mb-2">
          <Image src="/logo_louser_2022_1696999044 (1).png" alt="Logo Louser" width={150} height={50} className="object-contain" />
        </div>
        <p className="text-sm text-gray-500">Elevator Quality Control System</p>
        <p className="text-xs text-gray-400">Laporan Pemeriksaan dan Pengujian Elevator</p>
      </div>

      {renderTeknisiInput()}

      {/* DATA UMUM - PEMERIKSAAN */}
      <section className="bg-gray-50/50 rounded-xl p-6 mb-6 border border-gray-100">
        <h2 className="text-sm font-semibold text-gray-600 mb-4 border-b pb-2">DATA UMUM</h2>
        <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
          <div>
            <label className="block text-xs font-medium text-gray-500 uppercase tracking-wider mb-1">Nama Gedung</label>
            <input
              type="text"
              value={formData.namaGedung || ''}
              onChange={(e) => handleDataUmumChange('namaGedung', e.target.value)}
              className="w-full px-3 py-2 bg-white border border-gray-200 rounded-xl focus:outline-none focus:ring-2 focus:ring-blue-500 text-sm"
              disabled={isReadOnly || !isQC}
              placeholder="Masukkan Nama Gedung"
            />
          </div>
          <div>
            <label className="block text-xs font-medium text-gray-500 uppercase tracking-wider mb-1">Alamat</label>
            <input
              type="text"
              value={formData.alamat || ''}
              onChange={(e) => handleDataUmumChange('alamat', e.target.value)}
              className="w-full px-3 py-2 bg-white border border-gray-200 rounded-xl focus:outline-none focus:ring-2 focus:ring-blue-500 text-sm"
              disabled={isReadOnly || !isQC}
              placeholder="Masukkan Alamat"
            />
          </div>
          <div>
            <label className="block text-xs font-medium text-gray-500 uppercase tracking-wider mb-1">Jenis Elevator</label>
            <input
              type="text"
              value={formData.jenisElevator || ''}
              onChange={(e) => handleDataUmumChange('jenisElevator', e.target.value)}
              className="w-full px-3 py-2 bg-white border border-gray-200 rounded-xl focus:outline-none focus:ring-2 focus:ring-blue-500 text-sm"
              disabled={isReadOnly || !isQC}
              placeholder="Contoh: Penumpang, Barang, Pasien"
            />
          </div>
          <div>
            <label className="block text-xs font-medium text-gray-500 uppercase tracking-wider mb-1">Merk / Tipe</label>
            <input
              type="text"
              value={formData.merkTipe || ''}
              onChange={(e) => handleDataUmumChange('merkTipe', e.target.value)}
              className="w-full px-3 py-2 bg-white border border-gray-200 rounded-xl focus:outline-none focus:ring-2 focus:ring-blue-500 text-sm"
              disabled={isReadOnly || !isQC}
              placeholder="Contoh: Mitsubishi / NEXIEZ"
            />
          </div>
          <div>
            <label className="block text-xs font-medium text-gray-500 uppercase tracking-wider mb-1">Elevator No</label>
            <input
              type="text"
              value={formData.elevatorNo || ''}
              onChange={(e) => handleDataUmumChange('elevatorNo', e.target.value)}
              className="w-full px-3 py-2 bg-white border border-gray-200 rounded-xl focus:outline-none focus:ring-2 focus:ring-blue-500 text-sm"
              disabled={isReadOnly || !isQC}
              placeholder="Nomor unit elevator"
            />
          </div>
          <div>
            <label className="block text-xs font-medium text-gray-500 uppercase tracking-wider mb-1">Tanggal Riksa Uji</label>
            <input
              type="date"
              value={formData.tanggalRiksaUji || ''}
              onChange={(e) => handleDataUmumChange('tanggalRiksaUji', e.target.value)}
              className="w-full px-3 py-2 bg-white border border-gray-200 rounded-xl focus:outline-none focus:ring-2 focus:ring-blue-500 text-sm"
              disabled={isReadOnly || !isQC}
            />
          </div>
          <div>
            <label className="block text-xs font-medium text-gray-500 uppercase tracking-wider mb-1">Nama Pabrik Pembuat</label>
            <input
              type="text"
              value={formData.namaPabrikPembuat || ''}
              onChange={(e) => handleDataUmumChange('namaPabrikPembuat', e.target.value)}
              className="w-full px-3 py-2 bg-white border border-gray-200 rounded-xl focus:outline-none focus:ring-2 focus:ring-blue-500 text-sm"
              disabled={isReadOnly || !isQC}
              placeholder="Nama pabrik pembuat elevator"
            />
          </div>
          <div>
            <label className="block text-xs font-medium text-gray-500 uppercase tracking-wider mb-1">Tahun Pemasangan / Modernisasi</label>
            <input
              type="text"
              value={formData.tahunPemasangan || ''}
              onChange={(e) => handleDataUmumChange('tahunPemasangan', e.target.value)}
              className="w-full px-3 py-2 bg-white border border-gray-200 rounded-xl focus:outline-none focus:ring-2 focus:ring-blue-500 text-sm"
              disabled={isReadOnly || !isQC}
              placeholder="Tahun pemasangan atau modernisasi"
            />
          </div>
          <div>
            <label className="block text-xs font-medium text-gray-500 uppercase tracking-wider mb-1">Kapasitas Angkut</label>
            <input
              type="text"
              value={formData.kapasitasAngkut || ''}
              onChange={(e) => handleDataUmumChange('kapasitasAngkut', e.target.value)}
              className="w-full px-3 py-2 bg-white border border-gray-200 rounded-xl focus:outline-none focus:ring-2 focus:ring-blue-500 text-sm"
              disabled={isReadOnly || !isQC}
              placeholder="Contoh: 1000 kg"
            />
          </div>
          <div>
            <label className="block text-xs font-medium text-gray-500 uppercase tracking-wider mb-1">Kecepatan Angkut</label>
            <input
              type="text"
              value={formData.kecepatanAngkut || ''}
              onChange={(e) => handleDataUmumChange('kecepatanAngkut', e.target.value)}
              className="w-full px-3 py-2 bg-white border border-gray-200 rounded-xl focus:outline-none focus:ring-2 focus:ring-blue-500 text-sm"
              disabled={isReadOnly || !isQC}
              placeholder="Contoh: 60 m/menit"
            />
          </div>
          <div>
            <label className="block text-xs font-medium text-gray-500 uppercase tracking-wider mb-1">Melayani (Lantai)</label>
            <input
              type="text"
              value={formData.melayani || ''}
              onChange={(e) => handleDataUmumChange('melayani', e.target.value)}
              className="w-full px-3 py-2 bg-white border border-gray-200 rounded-xl focus:outline-none focus:ring-2 focus:ring-blue-500 text-sm"
              disabled={isReadOnly || !isQC}
              placeholder="Contoh: Lantai 1 - 10"
            />
          </div>
          <div className="md:col-span-2">
            <label className="block text-xs font-medium text-gray-500 uppercase tracking-wider mb-1">Data Pemeliharaan</label>
            <textarea
              value={formData.dataPemeliharaan || ''}
              onChange={(e) => handleDataUmumChange('dataPemeliharaan', e.target.value)}
              className="w-full px-3 py-2 bg-white border border-gray-200 rounded-xl focus:outline-none focus:ring-2 focus:ring-blue-500 text-sm"
              disabled={isReadOnly || !isQC}
              placeholder="Riwayat pemeliharaan elevator"
              rows={2}
            />
          </div>
          <div className="md:col-span-2">
            <label className="block text-xs font-medium text-gray-500 uppercase tracking-wider mb-1">Hasil Riksa Uji Sebelumnya</label>
            <textarea
              value={formData.hasilRiksaUjiSebelumnya || ''}
              onChange={(e) => handleDataUmumChange('hasilRiksaUjiSebelumnya', e.target.value)}
              className="w-full px-3 py-2 bg-white border border-gray-200 rounded-xl focus:outline-none focus:ring-2 focus:ring-blue-500 text-sm"
              disabled={isReadOnly || !isQC}
              placeholder="Hasil riksa uji sebelumnya"
              rows={2}
            />
          </div>
        </div>
      </section>

      {/* STATUS BAR */}
      <div className="flex flex-wrap items-center gap-3 bg-gray-50/50 p-3 rounded-xl mb-6">
        <span className="text-xs font-medium text-gray-500">Status:</span>
        <span className={`px-3 py-1 rounded-full text-xs font-medium ${formData.status === 'draft' ? 'bg-gray-200 text-gray-600' : formData.status === 'qc_approved' ? 'bg-blue-100 text-blue-700' : formData.status === 'revision' ? 'bg-amber-100 text-amber-700' : formData.status === 'maintenance_done' ? 'bg-orange-100 text-orange-700' : 'bg-emerald-100 text-emerald-700'}`}>
          {formData.status === 'draft' ? 'Draft' : formData.status === 'qc_approved' ? 'Maintenance' : formData.status === 'revision' ? 'REVISION' : formData.status === 'maintenance_done' ? 'Selesai Perbaikan' : 'FINAL APPROVED'}
        </span>
        {formData.submittedAt && <span className="text-xs text-gray-400">Dikirim: {formData.submittedAt}</span>}
        {formData.qcVerification?.verifiedAt && <span className="text-xs text-gray-400">Diverifikasi: {formData.qcVerification.verifiedAt}</span>}
      </div>

      {/* TABEL PEMERIKSAAN */}
      <section className="bg-white rounded-xl border border-gray-100 overflow-hidden mb-6">
        <div className="bg-blue-700 px-6 py-3 flex flex-wrap justify-between items-center">
          <h2 className="text-sm font-semibold text-white">PEMERIKSAAN DAN PENGUJIAN</h2>
          <div className="flex items-center gap-4">
            <span className="text-xs text-white bg-blue-600 px-3 py-1 rounded-full">{formData.inspectionData.length} Item</span>
            <div className="relative">
              <input
                type="text"
                value={searchQuery}
                onChange={(e) => setSearchQuery(e.target.value)}
                placeholder="🔍 Cari komponen / ketentuan..."
                className="px-3 py-1.5 bg-white/20 border border-white/30 rounded-lg text-white placeholder-white/70 text-xs focus:outline-none focus:ring-2 focus:ring-white/50 w-48"
              />
            </div>
          </div>
        </div>

        <div className="overflow-x-auto" style={{ maxHeight: '600px', overflowY: 'auto' }}>
          <table className="w-full text-xs border-collapse" style={{ minWidth: '1100px' }}>
            <thead className="sticky top-0 z-10">
              <tr className="bg-gray-100">
                <th className="p-2 text-center text-gray-600 font-semibold border-b border-gray-300 w-10">NO</th>
                <th className="p-2 text-left text-gray-600 font-semibold border-b border-gray-300 min-w-[150px] max-w-[200px]">Komponen</th>
                <th className="p-2 text-left text-gray-600 font-semibold border-b border-gray-300 min-w-[180px] max-w-[250px]">Ketentuan</th>
                <th className="p-2 text-center text-gray-600 font-semibold border-b border-gray-300 w-24">Hasil</th>
                <th className="p-2 text-left text-gray-600 font-semibold border-b border-gray-300 min-w-[120px] max-w-[180px]">Notes</th>
                <th className="p-2 text-center text-gray-600 font-semibold border-b border-gray-300 w-16">Bobot</th>
                <th className="p-2 text-center text-gray-600 font-semibold border-b border-gray-300 w-16">Hitung</th>
              </tr>
            </thead>
            <tbody>
              {filteredItems.map((originalItem, index) => {
                const item = formData.inspectionData.find((i: any) => i.id === originalItem.id);
                if (!item) return null;

                const prevItem = index > 0 ? filteredItems[index - 1] : null;
                const showCategoryHeader = !prevItem || prevItem.category !== originalItem.category;

                const isNotGood = item.status === 'Not Good';
                const isApproved = item.isApproved;
                const isEditable = isQC && !isReadOnly && formData.status === 'draft';
                const isMaintEditable = isMaintenance && !isReadOnly && (formData.status === 'qc_approved' || formData.status === 'revision');
                const canApprove = isQC && !isReadOnly && (formData.status === 'maintenance_done' || formData.status === 'revision') && isNotGood && !isApproved;
                const hitung = calculateHitung(item);
                const hitungPersen = hitung * 100;

                return (
                  <>
                    {showCategoryHeader && (
                      <tr className="bg-blue-50">
                        <td colSpan={7} className="p-2 font-bold text-blue-800 text-xs uppercase tracking-wider border-t-2 border-blue-300">
                          {originalItem.category}
                        </td>
                      </tr>
                    )}
                    <tr className="border-b border-gray-100 hover:bg-gray-50/50 transition">
                      <td className="p-2 text-center text-gray-400 font-medium">{originalItem.id}</td>
                      <td className="p-2 text-gray-700 align-top break-words whitespace-normal leading-tight">{originalItem.item}</td>
                      <td className="p-2 text-gray-500 align-top break-words whitespace-normal leading-tight text-[10px]">{originalItem.ketentuan}</td>
                      <td className="p-2 align-top text-center">
                        {isEditable ? (
                          <div className="flex flex-col gap-0.5 text-[10px]">
                            <label className="flex items-center gap-1"><input type="radio" name={`status-${item.id}`} value="Good" checked={item.status === 'Good'} onChange={() => handleStatusChange(item.id, 'Good')} className="w-3 h-3"/> Good</label>
                            <label className="flex items-center gap-1"><input type="radio" name={`status-${item.id}`} value="Not Good" checked={item.status === 'Not Good'} onChange={() => handleStatusChange(item.id, 'Not Good')} className="w-3 h-3"/> Not Good</label>
                            <label className="flex items-center gap-1"><input type="radio" name={`status-${item.id}`} value="N/A" checked={item.status === 'N/A'} onChange={() => handleStatusChange(item.id, 'N/A')} className="w-3 h-3"/> N/A</label>
                          </div>
                        ) : (
                          <span className={`px-2 py-0.5 rounded-full text-[10px] font-medium ${
                            item.status === 'Good' ? 'bg-emerald-100 text-emerald-700' : 
                            item.status === 'N/A' ? 'bg-gray-200 text-gray-500' : 
                            item.status === 'Not Good' && isApproved ? 'bg-emerald-200 text-emerald-800' : 
                            item.status === 'Not Good' && !isApproved ? 'bg-rose-100 text-rose-700' : 
                            'bg-gray-100 text-gray-400'
                          }`}>{item.status || '-'}</span>
                        )}
                      </td>
                      <td className="p-2 align-top">
                        {isEditable ? (
                          <textarea 
                            value={item.notes} 
                            onChange={(e) => handleNotesChange(item.id, e.target.value)} 
                            disabled={item.status === 'Good' || item.status === 'N/A'} 
                            placeholder="Temuan..." 
                            className={`w-full px-2 py-1 border rounded-lg text-[10px] ${isNotGood && !item.notes ? 'border-rose-300 bg-rose-50' : 'border-gray-200'} ${item.status === 'Good' || item.status === 'N/A' ? 'bg-gray-50' : 'bg-white'}`} 
                            rows={2}
                          />
                        ) : (
                          <span className="text-gray-600 break-words whitespace-normal text-[10px]">{item.notes || '-'}</span>
                        )}
                      </td>
                      <td className="p-2 text-center font-medium text-gray-600">
                        {(originalItem.weight * 100).toFixed(2).replace('.', ',')}%
                      </td>
                      <td className="p-2 text-center">
                        <span className={`font-bold ${hitungPersen >= 0 ? 'text-emerald-600' : 'text-rose-600'}`}>
                          {hitungPersen.toFixed(2).replace('.', ',')}%
                        </span>
                      </td>
                    </tr>
                  </>
                );
              })}
            </tbody>
          </table>
        </div>

        {/* REKAP SKOR */}
        <div className="bg-gray-50 p-4 border-t border-gray-200">
          <div className="grid grid-cols-2 md:grid-cols-4 gap-4">
            <div className="bg-white p-3 rounded-xl text-center shadow-sm">
              <div className="text-xs text-gray-400">Total Item</div>
              <div className="text-xl font-bold text-gray-700">{totalItems}</div>
            </div>
            <div className="bg-emerald-50 p-3 rounded-xl text-center">
              <div className="text-xs text-emerald-500">Good + N/A + Approved</div>
              <div className="text-xl font-bold text-emerald-700">{goodCount}</div>
            </div>
            <div className="bg-rose-50 p-3 rounded-xl text-center">
              <div className="text-xs text-rose-500">Not Good (Belum Approve)</div>
              <div className="text-xl font-bold text-rose-700">{notGoodCount}</div>
            </div>
            <div className="bg-blue-50 p-3 rounded-xl text-center">
              <div className="text-xs text-blue-500">Total Nilai</div>
              <div className="text-xl font-bold text-blue-700">{totalScore.toFixed(2).replace('.', ',')}%</div>
            </div>
          </div>
        </div>
      </section>

      {/* TEMUAN NEGATIF */}
      <section className="bg-white rounded-xl border border-gray-100 p-6 mb-6">
        <h2 className="text-sm font-semibold text-gray-600 mb-4 border-b pb-2">TEMUAN NEGATIF</h2>

        {/* A. KINERJA KESELAMATAN */}
        <div className="mb-6">
          <div className="flex items-center justify-between mb-2">
            <h3 className="text-xs font-bold text-gray-700 bg-gray-100 px-3 py-1 rounded-lg">A. KINERJA KESELAMATAN</h3>
            {isQC && !isReadOnly && (
              <button
                type="button"
                onClick={() => handleTemuanAdd('temuanKeselamatan')}
                className="bg-blue-500 hover:bg-blue-600 text-white px-3 py-1 rounded-lg text-xs transition"
              >
                + Tambah Baris
              </button>
            )}
          </div>
          <div className="overflow-x-auto">
            <table className="w-full text-xs border-collapse">
              <thead>
                <tr className="bg-gray-50">
                  <th className="p-2 border border-gray-200 text-center w-10">NO</th>
                  <th className="p-2 border border-gray-200 text-left min-w-[150px]">Temuan</th>
                  <th className="p-2 border border-gray-200 text-left min-w-[150px]">Dokumentasi (Foto)</th>
                  <th className="p-2 border border-gray-200 text-left min-w-[120px]">Kategori</th>
                  <th className="p-2 border border-gray-200 text-left min-w-[150px]">Solusi</th>
                  {isQC && !isReadOnly && <th className="p-2 border border-gray-200 text-center w-10">Aksi</th>}
                </tr>
              </thead>
              <tbody>
                {formData.temuanKeselamatan?.map((t: any) => (
                  <tr key={t.id}>
                    <td className="p-2 border border-gray-200 text-center text-gray-400">{t.id}</td>
                    <td className="p-1 border border-gray-200">
                      <input
                        type="text"
                        value={t.temuan}
                        onChange={(e) => handleTemuanChange('temuanKeselamatan', t.id, 'temuan', e.target.value)}
                        className="w-full px-2 py-1 border-0 focus:ring-1 focus:ring-blue-500 rounded text-xs bg-transparent"
                        disabled={isReadOnly || !isQC}
                        placeholder="Temuan..."
                      />
                    </td>
                    <td className="p-1 border border-gray-200">
                      {t.dokumentasi ? (
                        <div className="flex items-center gap-2">
                          <img src={t.dokumentasi} alt="dokumentasi" className="w-12 h-12 object-cover rounded border" />
                          {isQC && !isReadOnly && (
                            <button
                              type="button"
                              onClick={() => removeTemuanDokumentasi('temuanKeselamatan', t.id)}
                              className="text-red-500 hover:text-red-700 text-sm"
                              title="Hapus foto"
                            >
                              ✕
                            </button>
                          )}
                        </div>
                      ) : (
                        isQC && !isReadOnly && (
                          <input
                            type="file"
                            accept="image/*"
                            onChange={(e) => handleTemuanDokumentasi('temuanKeselamatan', t.id, e)}
                            className="text-[10px] w-full"
                          />
                        )
                      )}
                      {!t.dokumentasi && isReadOnly && <span className="text-gray-400 text-[10px]">-</span>}
                    </td>
                    <td className="p-1 border border-gray-200">
                      <input
                        type="text"
                        value={t.kategori}
                        onChange={(e) => handleTemuanChange('temuanKeselamatan', t.id, 'kategori', e.target.value)}
                        className="w-full px-2 py-1 border-0 focus:ring-1 focus:ring-blue-500 rounded text-xs bg-transparent"
                        disabled={isReadOnly || !isQC}
                        placeholder="Kategori..."
                      />
                    </td>
                    <td className="p-1 border border-gray-200">
                      <input
                        type="text"
                        value={t.solusi}
                        onChange={(e) => handleTemuanChange('temuanKeselamatan', t.id, 'solusi', e.target.value)}
                        className="w-full px-2 py-1 border-0 focus:ring-1 focus:ring-blue-500 rounded text-xs bg-transparent"
                        disabled={isReadOnly || !isQC}
                        placeholder="Solusi..."
                      />
                    </td>
                    {isQC && !isReadOnly && (
                      <td className="p-1 border border-gray-200 text-center">
                        <button
                          type="button"
                          onClick={() => handleTemuanRemove('temuanKeselamatan', t.id)}
                          className="text-red-400 hover:text-red-600 text-sm"
                          title="Hapus baris"
                        >
                          ✕
                        </button>
                      </td>
                    )}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>

        {/* B. KESESUAIAN PERMENAKER */}
        <div className="mb-6">
          <div className="flex items-center justify-between mb-2">
            <h3 className="text-xs font-bold text-gray-700 bg-gray-100 px-3 py-1 rounded-lg">B. KESESUAIAN PERMENAKER</h3>
            {isQC && !isReadOnly && (
              <button
                type="button"
                onClick={() => handleTemuanAdd('temuanPermenaker')}
                className="bg-blue-500 hover:bg-blue-600 text-white px-3 py-1 rounded-lg text-xs transition"
              >
                + Tambah Baris
              </button>
            )}
          </div>
          <div className="overflow-x-auto">
            <table className="w-full text-xs border-collapse">
              <thead>
                <tr className="bg-gray-50">
                  <th className="p-2 border border-gray-200 text-center w-10">NO</th>
                  <th className="p-2 border border-gray-200 text-left min-w-[150px]">Temuan</th>
                  <th className="p-2 border border-gray-200 text-left min-w-[150px]">Dokumentasi (Foto)</th>
                  <th className="p-2 border border-gray-200 text-left min-w-[120px]">Kategori</th>
                  <th className="p-2 border border-gray-200 text-left min-w-[150px]">Solusi</th>
                  {isQC && !isReadOnly && <th className="p-2 border border-gray-200 text-center w-10">Aksi</th>}
                </tr>
              </thead>
              <tbody>
                {formData.temuanPermenaker?.map((t: any) => (
                  <tr key={t.id}>
                    <td className="p-2 border border-gray-200 text-center text-gray-400">{t.id}</td>
                    <td className="p-1 border border-gray-200">
                      <input
                        type="text"
                        value={t.temuan}
                        onChange={(e) => handleTemuanChange('temuanPermenaker', t.id, 'temuan', e.target.value)}
                        className="w-full px-2 py-1 border-0 focus:ring-1 focus:ring-blue-500 rounded text-xs bg-transparent"
                        disabled={isReadOnly || !isQC}
                        placeholder="Temuan..."
                      />
                    </td>
                    <td className="p-1 border border-gray-200">
                      {t.dokumentasi ? (
                        <div className="flex items-center gap-2">
                          <img src={t.dokumentasi} alt="dokumentasi" className="w-12 h-12 object-cover rounded border" />
                          {isQC && !isReadOnly && (
                            <button
                              type="button"
                              onClick={() => removeTemuanDokumentasi('temuanPermenaker', t.id)}
                              className="text-red-500 hover:text-red-700 text-sm"
                              title="Hapus foto"
                            >
                              ✕
                            </button>
                          )}
                        </div>
                      ) : (
                        isQC && !isReadOnly && (
                          <input
                            type="file"
                            accept="image/*"
                            onChange={(e) => handleTemuanDokumentasi('temuanPermenaker', t.id, e)}
                            className="text-[10px] w-full"
                          />
                        )
                      )}
                      {!t.dokumentasi && isReadOnly && <span className="text-gray-400 text-[10px]">-</span>}
                    </td>
                    <td className="p-1 border border-gray-200">
                      <input
                        type="text"
                        value={t.kategori}
                        onChange={(e) => handleTemuanChange('temuanPermenaker', t.id, 'kategori', e.target.value)}
                        className="w-full px-2 py-1 border-0 focus:ring-1 focus:ring-blue-500 rounded text-xs bg-transparent"
                        disabled={isReadOnly || !isQC}
                        placeholder="Kategori..."
                      />
                    </td>
                    <td className="p-1 border border-gray-200">
                      <input
                        type="text"
                        value={t.solusi}
                        onChange={(e) => handleTemuanChange('temuanPermenaker', t.id, 'solusi', e.target.value)}
                        className="w-full px-2 py-1 border-0 focus:ring-1 focus:ring-blue-500 rounded text-xs bg-transparent"
                        disabled={isReadOnly || !isQC}
                        placeholder="Solusi..."
                      />
                    </td>
                    {isQC && !isReadOnly && (
                      <td className="p-1 border border-gray-200 text-center">
                        <button
                          type="button"
                          onClick={() => handleTemuanRemove('temuanPermenaker', t.id)}
                          className="text-red-400 hover:text-red-600 text-sm"
                          title="Hapus baris"
                        >
                          ✕
                        </button>
                      </td>
                    )}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>

        {/* C. KINERJA OPERASIONAL */}
        <div className="mb-6">
          <div className="flex items-center justify-between mb-2">
            <h3 className="text-xs font-bold text-gray-700 bg-gray-100 px-3 py-1 rounded-lg">C. KINERJA OPERASIONAL</h3>
            {isQC && !isReadOnly && (
              <button
                type="button"
                onClick={() => handleTemuanAdd('temuanOperasional')}
                className="bg-blue-500 hover:bg-blue-600 text-white px-3 py-1 rounded-lg text-xs transition"
              >
                + Tambah Baris
              </button>
            )}
          </div>
          <div className="overflow-x-auto">
            <table className="w-full text-xs border-collapse">
              <thead>
                <tr className="bg-gray-50">
                  <th className="p-2 border border-gray-200 text-center w-10">NO</th>
                  <th className="p-2 border border-gray-200 text-left min-w-[150px]">Temuan</th>
                  <th className="p-2 border border-gray-200 text-left min-w-[150px]">Dokumentasi (Foto)</th>
                  <th className="p-2 border border-gray-200 text-left min-w-[120px]">Kategori</th>
                  <th className="p-2 border border-gray-200 text-left min-w-[150px]">Solusi</th>
                  {isQC && !isReadOnly && <th className="p-2 border border-gray-200 text-center w-10">Aksi</th>}
                </tr>
              </thead>
              <tbody>
                {formData.temuanOperasional?.map((t: any) => (
                  <tr key={t.id}>
                    <td className="p-2 border border-gray-200 text-center text-gray-400">{t.id}</td>
                    <td className="p-1 border border-gray-200">
                      <input
                        type="text"
                        value={t.temuan}
                        onChange={(e) => handleTemuanChange('temuanOperasional', t.id, 'temuan', e.target.value)}
                        className="w-full px-2 py-1 border-0 focus:ring-1 focus:ring-blue-500 rounded text-xs bg-transparent"
                        disabled={isReadOnly || !isQC}
                        placeholder="Temuan..."
                      />
                    </td>
                    <td className="p-1 border border-gray-200">
                      {t.dokumentasi ? (
                        <div className="flex items-center gap-2">
                          <img src={t.dokumentasi} alt="dokumentasi" className="w-12 h-12 object-cover rounded border" />
                          {isQC && !isReadOnly && (
                            <button
                              type="button"
                              onClick={() => removeTemuanDokumentasi('temuanOperasional', t.id)}
                              className="text-red-500 hover:text-red-700 text-sm"
                              title="Hapus foto"
                            >
                              ✕
                            </button>
                          )}
                        </div>
                      ) : (
                        isQC && !isReadOnly && (
                          <input
                            type="file"
                            accept="image/*"
                            onChange={(e) => handleTemuanDokumentasi('temuanOperasional', t.id, e)}
                            className="text-[10px] w-full"
                          />
                        )
                      )}
                      {!t.dokumentasi && isReadOnly && <span className="text-gray-400 text-[10px]">-</span>}
                    </td>
                    <td className="p-1 border border-gray-200">
                      <input
                        type="text"
                        value={t.kategori}
                        onChange={(e) => handleTemuanChange('temuanOperasional', t.id, 'kategori', e.target.value)}
                        className="w-full px-2 py-1 border-0 focus:ring-1 focus:ring-blue-500 rounded text-xs bg-transparent"
                        disabled={isReadOnly || !isQC}
                        placeholder="Kategori..."
                      />
                    </td>
                    <td className="p-1 border border-gray-200">
                      <input
                        type="text"
                        value={t.solusi}
                        onChange={(e) => handleTemuanChange('temuanOperasional', t.id, 'solusi', e.target.value)}
                        className="w-full px-2 py-1 border-0 focus:ring-1 focus:ring-blue-500 rounded text-xs bg-transparent"
                        disabled={isReadOnly || !isQC}
                        placeholder="Solusi..."
                      />
                    </td>
                    {isQC && !isReadOnly && (
                      <td className="p-1 border border-gray-200 text-center">
                        <button
                          type="button"
                          onClick={() => handleTemuanRemove('temuanOperasional', t.id)}
                          className="text-red-400 hover:text-red-600 text-sm"
                          title="Hapus baris"
                        >
                          ✕
                        </button>
                      </td>
                    )}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>

        {/* D. TEMUAN PADA AREA LAIN */}
        <div>
          <div className="flex items-center justify-between mb-2">
            <h3 className="text-xs font-bold text-gray-700 bg-gray-100 px-3 py-1 rounded-lg">D. TEMUAN PADA AREA LAIN</h3>
            {isQC && !isReadOnly && (
              <button
                type="button"
                onClick={() => handleTemuanAdd('temuanAreaLain')}
                className="bg-blue-500 hover:bg-blue-600 text-white px-3 py-1 rounded-lg text-xs transition"
              >
                + Tambah Baris
              </button>
            )}
          </div>
          <div className="overflow-x-auto">
            <table className="w-full text-xs border-collapse">
              <thead>
                <tr className="bg-gray-50">
                  <th className="p-2 border border-gray-200 text-center w-10">NO</th>
                  <th className="p-2 border border-gray-200 text-left min-w-[180px]">Temuan</th>
                  <th className="p-2 border border-gray-200 text-left min-w-[150px]">Dokumentasi (Foto)</th>
                  <th className="p-2 border border-gray-200 text-left min-w-[180px]">Solusi</th>
                  {isQC && !isReadOnly && <th className="p-2 border border-gray-200 text-center w-10">Aksi</th>}
                </tr>
              </thead>
              <tbody>
                {formData.temuanAreaLain?.map((t: any) => (
                  <tr key={t.id}>
                    <td className="p-2 border border-gray-200 text-center text-gray-400">{t.id}</td>
                    <td className="p-1 border border-gray-200">
                      <input
                        type="text"
                        value={t.temuan}
                        onChange={(e) => handleTemuanChange('temuanAreaLain', t.id, 'temuan', e.target.value)}
                        className="w-full px-2 py-1 border-0 focus:ring-1 focus:ring-blue-500 rounded text-xs bg-transparent"
                        disabled={isReadOnly || !isQC}
                        placeholder="Temuan..."
                      />
                    </td>
                    <td className="p-1 border border-gray-200">
                      {t.dokumentasi ? (
                        <div className="flex items-center gap-2">
                          <img src={t.dokumentasi} alt="dokumentasi" className="w-12 h-12 object-cover rounded border" />
                          {isQC && !isReadOnly && (
                            <button
                              type="button"
                              onClick={() => removeTemuanDokumentasi('temuanAreaLain', t.id)}
                              className="text-red-500 hover:text-red-700 text-sm"
                              title="Hapus foto"
                            >
                              ✕
                            </button>
                          )}
                        </div>
                      ) : (
                        isQC && !isReadOnly && (
                          <input
                            type="file"
                            accept="image/*"
                            onChange={(e) => handleTemuanDokumentasi('temuanAreaLain', t.id, e)}
                            className="text-[10px] w-full"
                          />
                        )
                      )}
                      {!t.dokumentasi && isReadOnly && <span className="text-gray-400 text-[10px]">-</span>}
                    </td>
                    <td className="p-1 border border-gray-200">
                      <input
                        type="text"
                        value={t.solusi}
                        onChange={(e) => handleTemuanChange('temuanAreaLain', t.id, 'solusi', e.target.value)}
                        className="w-full px-2 py-1 border-0 focus:ring-1 focus:ring-blue-500 rounded text-xs bg-transparent"
                        disabled={isReadOnly || !isQC}
                        placeholder="Solusi..."
                      />
                    </td>
                    {isQC && !isReadOnly && (
                      <td className="p-1 border border-gray-200 text-center">
                        <button
                          type="button"
                          onClick={() => handleTemuanRemove('temuanAreaLain', t.id)}
                          className="text-red-400 hover:text-red-600 text-sm"
                          title="Hapus baris"
                        >
                          ✕
                        </button>
                      </td>
                    )}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      </section>

      {/* CATATAN QC & ATTACHMENT */}
      <section className="bg-white rounded-xl border border-gray-100 p-6 mb-6">
        <h2 className="text-sm font-semibold text-gray-600 mb-4">Catatan & Attachment</h2>
        
        {isQC && !isReadOnly && (
          <div className="mb-4">
            <label className="block text-xs font-medium text-gray-500 uppercase tracking-wider mb-1">Catatan QC</label>
            <textarea
              value={formData.qcNote || ''}
              onChange={(e) => setFormData((prev: any) => ({ ...prev, qcNote: e.target.value }))}
              placeholder="Tambahkan catatan di sini..."
              className="w-full px-3 py-2 bg-white border border-gray-200 rounded-xl focus:outline-none focus:ring-2 focus:ring-blue-500 text-sm"
              rows={3}
              disabled={isReadOnly}
            />
          </div>
        )}
        {formData.qcNote && isReadOnly && (
          <div className="mb-4 p-3 bg-blue-50 border border-blue-200 rounded-xl text-gray-700 text-sm">
            <span className="font-medium text-blue-700">Catatan QC:</span> {formData.qcNote}
          </div>
        )}

        <div className="border-2 border-dashed border-gray-200 rounded-xl p-6 text-center hover:border-blue-400 transition">
          {isQC && !isReadOnly ? (
            <input
              type="file"
              accept=".pdf,.doc,.docx,.xlsx,.xls,.jpg,.jpeg,.png"
              onChange={handleAttachmentChange}
              className="text-sm text-gray-500 file:mr-3 file:py-2 file:px-4 file:rounded-xl file:border-0 file:text-sm file:font-medium file:bg-blue-50 file:text-blue-700 hover:file:bg-blue-100"
            />
          ) : formData.attachment ? (
            <div className="flex items-center justify-center gap-4 flex-wrap">
              <span className="text-sm text-gray-600">📄 {formData.attachmentName}</span>
              <button type="button" onClick={handleDownloadAttachment} className="bg-blue-600 hover:bg-blue-700 text-white px-4 py-2 rounded-xl text-sm transition">Download</button>
              <a href={formData.attachment} target="_blank" rel="noopener" className="bg-emerald-600 hover:bg-emerald-700 text-white px-4 py-2 rounded-xl text-sm transition">Lihat</a>
            </div>
          ) : (
            <p className="text-sm text-gray-400">Tidak ada file attachment</p>
          )}
        </div>
      </section>

      {/* FOOTER */}
      <div className="border-t pt-4">
        <div className="flex flex-wrap justify-between items-start gap-4">
          <div className="text-left">
            <p className="text-xs text-gray-400">Dicetak: {new Date().toLocaleDateString('id-ID', { day: '2-digit', month: 'long', year: 'numeric' })}</p>
            <p className="text-xs text-gray-400">© PT Louserindo Megah Permai</p>
          </div>
          <div className="text-center">
            <p className="text-xs text-gray-500 mb-1">Tanda Tangan</p>
            <div className="w-40 h-12 border-b-2 border-gray-300"></div>
            <p className="text-[10px] text-gray-400 mt-1">(_____________________)</p>
          </div>
        </div>
        <div className="text-right text-xs text-gray-400 border-t pt-4 mt-4">
          <p>Dibuat: {new Date(formData.createdAt).toLocaleString('id-ID')}</p>
          {formData.submittedAt && <p>Dikirim: {formData.submittedAt}</p>}
          {formData.qcVerification?.verifiedAt && <p>Diverifikasi: {formData.qcVerification.verifiedAt}</p>}
        </div>
      </div>
    </form>
  );
}

// ============================================
// KOMPONEN HOME
// ============================================
export default function Home() {
  const [isLoggedIn, setIsLoggedIn] = useState(false);
  const [userRole, setUserRole] = useState<UserRole>(null);

  const handleLogin = (role: UserRole) => {
    setUserRole(role);
    setIsLoggedIn(true);
  };

  const handleLogout = () => {
    setIsLoggedIn(false);
    setUserRole(null);
  };

  if (!isLoggedIn) {
    return <LoginPage onLogin={handleLogin} />;
  }

  return <Dashboard role={userRole} onLogout={handleLogout} />;
}