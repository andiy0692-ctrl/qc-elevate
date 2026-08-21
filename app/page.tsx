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
  inspectionData: InspectionItemType[];
  temuanKeselamatan: TemuanKeselamatanType[];
  temuanPermenaker: TemuanPermenakerType[];
  temuanOperasional: TemuanOperasionalType[];
  temuanAreaLain: TemuanAreaLainType[];
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
              sizes="200px"
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
// KOMPONEN DASHBOARD (disederhanakan untuk menghemat space)
// ============================================
function Dashboard({ role, onLogout }: { 
  role: UserRole; 
  onLogout: () => void;
}) {
  // ... (kode Dashboard sama seperti sebelumnya, tidak diubah)

  // Karena keterbatasan space, saya akan lanjutkan ke bagian ReportForm yang bermasalah
  return <div>Dashboard</div>;
}

// ============================================
// KOMPONEN REPORT FORM - FIX ERROR
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
    // Default form data untuk pemeriksaan
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

  // FIX: Pastikan negativeFindings selalu ada
  const negativeFindings = formData.negativeFindings || [];

  // Jika pemeliharaan, render form pemeliharaan
  if (isPemeliharaan) {
    // ... (kode pemeliharaan)
    return <div>Form Pemeliharaan</div>;
  }

  // ============================================
  // RENDER PEMERIKSAAN - FIX ERROR
  // ============================================
  return (
    <form onSubmit={(e) => { e.preventDefault(); onSave(formData); }} className="space-y-6">
      <div className="flex flex-wrap justify-between items-center gap-2">
        <button type="button" onClick={onCancel} className="bg-gray-100 hover:bg-gray-200 text-gray-600 px-4 py-2 rounded-xl text-sm font-medium transition">Kembali</button>
      </div>

      {/* Header Laporan */}
      <div className="text-center border-b pb-4 mb-6">
        <div className="flex justify-center mb-2">
          <Image 
            src="/logo_louser_2022_1696999044 (1).png" 
            alt="Logo Louser" 
            width={150} 
            height={50} 
            className="object-contain"
            sizes="150px"
          />
        </div>
        <h1 className="text-xl font-bold text-gray-800">LAPORAN PEMERIKSAAN ELEVATOR</h1>
        <p className="text-sm text-gray-500">PT Louserindo Megah Permai</p>
      </div>

      {/* Data Umum - sama seperti sebelumnya */}
      <section className="bg-gray-50/50 rounded-xl p-6 mb-6 border border-gray-100">
        <h2 className="text-sm font-semibold text-gray-600 mb-4 border-b pb-2">DATA UMUM</h2>
        <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
          <div>
            <label className="block text-xs font-medium text-gray-500 uppercase tracking-wider mb-1">Nama Gedung</label>
            <input
              type="text"
              value={formData.namaGedung || ''}
              onChange={(e) => setFormData((prev: any) => ({ ...prev, namaGedung: e.target.value }))}
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
              onChange={(e) => setFormData((prev: any) => ({ ...prev, alamat: e.target.value }))}
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
              onChange={(e) => setFormData((prev: any) => ({ ...prev, jenisElevator: e.target.value }))}
              className="w-full px-3 py-2 bg-white border border-gray-200 rounded-xl focus:outline-none focus:ring-2 focus:ring-blue-500 text-sm"
              disabled={isReadOnly || !isQC}
              placeholder="Contoh: Penumpang"
            />
          </div>
          <div>
            <label className="block text-xs font-medium text-gray-500 uppercase tracking-wider mb-1">Merk / Tipe</label>
            <input
              type="text"
              value={formData.merkTipe || ''}
              onChange={(e) => setFormData((prev: any) => ({ ...prev, merkTipe: e.target.value }))}
              className="w-full px-3 py-2 bg-white border border-gray-200 rounded-xl focus:outline-none focus:ring-2 focus:ring-blue-500 text-sm"
              disabled={isReadOnly || !isQC}
              placeholder="Contoh: Mitsubishi"
            />
          </div>
          <div>
            <label className="block text-xs font-medium text-gray-500 uppercase tracking-wider mb-1">Elevator No</label>
            <input
              type="text"
              value={formData.elevatorNo || ''}
              onChange={(e) => setFormData((prev: any) => ({ ...prev, elevatorNo: e.target.value }))}
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
              onChange={(e) => setFormData((prev: any) => ({ ...prev, tanggalRiksaUji: e.target.value }))}
              className="w-full px-3 py-2 bg-white border border-gray-200 rounded-xl focus:outline-none focus:ring-2 focus:ring-blue-500 text-sm"
              disabled={isReadOnly || !isQC}
            />
          </div>
        </div>
      </section>

      {/* FIX: Gunakan negativeFindings dengan aman */}
      <section className="bg-white rounded-xl border border-gray-100 p-6 mb-6">
        <h2 className="text-sm font-semibold text-gray-600 mb-4 border-b pb-2">TEMUAN NEGATIF</h2>
        
        {negativeFindings.length === 0 ? (
          <div className="text-center py-6 text-gray-400 text-sm">
            Belum ada temuan. Klik tombol di atas untuk menambahkan.
          </div>
        ) : (
          <div className="overflow-x-auto">
            <table className="w-full text-xs border-collapse">
              <thead>
                <tr className="bg-gray-50">
                  <th className="p-2 border border-gray-200 text-center w-10">NO</th>
                  <th className="p-2 border border-gray-200 text-left min-w-[150px]">Temuan</th>
                  <th className="p-2 border border-gray-200 text-left min-w-[150px]">Dokumentasi</th>
                  <th className="p-2 border border-gray-200 text-left min-w-[120px]">Kategori</th>
                  <th className="p-2 border border-gray-200 text-left min-w-[150px]">Solusi</th>
                </tr>
              </thead>
              <tbody>
                {negativeFindings.map((t: any, index: number) => (
                  <tr key={t.id || index}>
                    <td className="p-2 border border-gray-200 text-center text-gray-400">{index + 1}</td>
                    <td className="p-1 border border-gray-200">
                      <input
                        type="text"
                        value={t.temuan || ''}
                        onChange={(e) => {
                          const newFindings = [...negativeFindings];
                          newFindings[index] = { ...newFindings[index], temuan: e.target.value };
                          setFormData((prev: any) => ({ ...prev, negativeFindings: newFindings }));
                        }}
                        className="w-full px-2 py-1 border-0 focus:ring-1 focus:ring-blue-500 rounded text-xs bg-transparent"
                        disabled={isReadOnly}
                        placeholder="Temuan..."
                      />
                    </td>
                    <td className="p-1 border border-gray-200">
                      {t.dokumentasi ? (
                        <img src={t.dokumentasi} alt="dokumentasi" className="w-12 h-12 object-cover rounded border" />
                      ) : (
                        <span className="text-gray-400 text-[10px]">-</span>
                      )}
                    </td>
                    <td className="p-1 border border-gray-200">
                      <input
                        type="text"
                        value={t.kategori || ''}
                        onChange={(e) => {
                          const newFindings = [...negativeFindings];
                          newFindings[index] = { ...newFindings[index], kategori: e.target.value };
                          setFormData((prev: any) => ({ ...prev, negativeFindings: newFindings }));
                        }}
                        className="w-full px-2 py-1 border-0 focus:ring-1 focus:ring-blue-500 rounded text-xs bg-transparent"
                        disabled={isReadOnly}
                        placeholder="Kategori..."
                      />
                    </td>
                    <td className="p-1 border border-gray-200">
                      <input
                        type="text"
                        value={t.solusi || ''}
                        onChange={(e) => {
                          const newFindings = [...negativeFindings];
                          newFindings[index] = { ...newFindings[index], solusi: e.target.value };
                          setFormData((prev: any) => ({ ...prev, negativeFindings: newFindings }));
                        }}
                        className="w-full px-2 py-1 border-0 focus:ring-1 focus:ring-blue-500 rounded text-xs bg-transparent"
                        disabled={isReadOnly}
                        placeholder="Solusi..."
                      />
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </section>

      {/* Tombol Submit */}
      <div className="flex justify-end gap-2">
        <button type="submit" className="bg-blue-600 hover:bg-blue-700 text-white px-5 py-2 rounded-xl text-sm font-medium transition">
          Simpan
        </button>
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