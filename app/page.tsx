"use client";

import { useState, useEffect } from 'react';
import * as XLSX from 'xlsx';
import { db } from './lib/firebase';
import { 
  collection, 
  getDocs, 
  addDoc, 
  updateDoc, 
  deleteDoc, 
  doc, 
  onSnapshot,
  query,
  orderBy
} from 'firebase/firestore';
import Image from 'next/image';

// ============================================
// DATA INSPEKSI (116 ITEM) - BOBOT SESUAI EXCEL
// ============================================
const inspectionItems = [
  { id: 1, category: "Mesin", item: "Dudukan Mesin", priority: "Prioritas_3", weight: 0.0005 },
  { id: 2, category: "Mesin", item: "Rem Mekanik", priority: "Prioritas_1", weight: 0.04 },
  { id: 3, category: "Mesin", item: "Rem Electric (Brake Switch)", priority: "Prioritas_2", weight: 0.008 },
  { id: 4, category: "Mesin", item: "Konstruksi Kamar Mesin", priority: "Prioritas_3", weight: 0.0005 },
  { id: 5, category: "Mesin", item: "Ruang Bebas Kamar Mesin - Depan alat pengendali ≥ 700 mm", priority: "Prioritas_3", weight: 0.0005 },
  { id: 6, category: "Mesin", item: "Ruang Bebas Kamar Mesin - Depan barang bergerak ≥ 500x600 mm", priority: "Prioritas_3", weight: 0.0005 },
  { id: 7, category: "Mesin", item: "Ruang Bebas Kamar Mesin - Di atas mesin ≥ 500 mm", priority: "Prioritas_3", weight: 0.0005 },
  { id: 8, category: "Mesin", item: "Penerangan Kamar Mesin - Area kerja ≥ 100 lux", priority: "Prioritas_3", weight: 0.0005 },
  { id: 9, category: "Mesin", item: "Penerangan Kamar Mesin - Di antara area kerja ≥ 50 lux", priority: "Prioritas_3", weight: 0.0005 },
  { id: 10, category: "Mesin", item: "Ventilasi/Pendingin Ruangan", priority: "Prioritas_3", weight: 0.0005 },
  { id: 11, category: "Mesin", item: "Pintu Kamar Mesin", priority: "Prioritas_3", weight: 0.0005 },
  { id: 12, category: "Mesin", item: "Posisi Panel Hubung Bagi Listrik", priority: "Prioritas_3", weight: 0.0005 },
  { id: 13, category: "Mesin", item: "Alat Pelindung Benda Berputar", priority: "Prioritas_3", weight: 0.0005 },
  { id: 14, category: "Mesin", item: "Pelindung Lubang Tali Baja/sabuk Penggantung", priority: "Prioritas_3", weight: 0.0005 },
  { id: 15, category: "Mesin", item: "Tangga menuju kamar mesin", priority: "Prioritas_3", weight: 0.0005 },
  { id: 16, category: "Mesin", item: "Perbedaan ketinggian lantai di kamar mesin > 500mm", priority: "Prioritas_3", weight: 0.0005 },
  { id: 17, category: "Mesin", item: "Tersedia Alat Pemadam Api Ringan", priority: "Prioritas_3", weight: 0.0005 },
  { id: 18, category: "Mesin", item: "MRL - Penempatan panel kontrol dan PHB listrik", priority: "Prioritas_2", weight: 0.009 },
  { id: 19, category: "Mesin", item: "MRL - Intensitas cahaya area kerja di kamar mesin", priority: "Prioritas_3", weight: 0.0005 },
  { id: 20, category: "Mesin", item: "MRL - Intensitas cahaya diantara area kerja", priority: "Prioritas_3", weight: 0.0005 },
  { id: 21, category: "Mesin", item: "MRL - Alat pembuka rem mesin secara elektrik/ manual", priority: "Prioritas_1", weight: 0.04 },
  { id: 22, category: "Mesin", item: "MRL - Penempatan APAR", priority: "Prioritas_3", weight: 0.0005 },
  { id: 23, category: "Mesin", item: "MRL - Emergency stop switch", priority: "Prioritas_2", weight: 0.009 },
  { id: 24, category: "TALI/SABUK", item: "Tali / sabuk penggantung - Tidak memiliki sambungan", priority: "Prioritas_1", weight: 0.04 },
  { id: 25, category: "TALI/SABUK", item: "Tali/sabuk penggantung - Tidak menggunakan rantai", priority: "Prioritas_2", weight: 0.009 },
  { id: 26, category: "TALI/SABUK", item: "Nilai faktor keamanan tali / sabuk penggantung", priority: "Prioritas_2", weight: 0.009 },
  { id: 27, category: "TALI/SABUK", item: "Tali penggantung Kereta jenis tali dengan bobot imbang", priority: "Prioritas_2", weight: 0.009 },
  { id: 28, category: "TALI/SABUK", item: "Tali penggantung Kereta tanpa Bobot imbang", priority: "Prioritas_2", weight: 0.009 },
  { id: 29, category: "TALI/SABUK", item: "Sabuk - ≥ 3 x 30 mm, ≥ 2 jalur", priority: "Prioritas_2", weight: 0.009 },
  { id: 30, category: "TALI/SABUK", item: "Alat Pengaman pada elevator tanpa bobot imbang", priority: "Prioritas_2", weight: 0.009 },
  { id: 31, category: "TEROMOL", item: "Alur teromol", priority: "Prioritas_2", weight: 0.009 },
  { id: 32, category: "TEROMOL", item: "Diameter teromol Penumpang/barang", priority: "Prioritas_3", weight: 0.0005 },
  { id: 33, category: "TEROMOL", item: "Diameter teromol Governor", priority: "Prioritas_3", weight: 0.0005 },
  { id: 34, category: "BANGUNAN", item: "Konstruksi ruang luncur", priority: "Prioritas_2", weight: 0.009 },
  { id: 35, category: "BANGUNAN", item: "Dinding ruang luncur - Elevator panorama", priority: "Prioritas_3", weight: 0.0005 },
  { id: 36, category: "BANGUNAN", item: "Ruang luncur - Bersih, bebas instalasi", priority: "Prioritas_3", weight: 0.0005 },
  { id: 37, category: "BANGUNAN", item: "Penerangan ruang luncur ≥ 100 lux", priority: "Prioritas_3", weight: 0.0005 },
  { id: 38, category: "BANGUNAN", item: "Pintu darurat (non stop)", priority: "Prioritas_3", weight: 0.0005 },
  { id: 39, category: "BANGUNAN", item: "Ukuran pintu darurat", priority: "Prioritas_3", weight: 0.0005 },
  { id: 40, category: "BANGUNAN", item: "Saklar pengaman pintu darurat", priority: "Prioritas_3", weight: 0.0005 },
  { id: 41, category: "BANGUNAN", item: "Jembatan bantu dari pintu darurat", priority: "Prioritas_3", weight: 0.0005 },
  { id: 42, category: "BANGUNAN", item: "Ruang bebas diatas sangkar ≥ 500 mm", priority: "Prioritas_2", weight: 0.009 },
  { id: 43, category: "BANGUNAN", item: "Ruang bebas lekuk dasar ≥ 500 mm", priority: "Prioritas_2", weight: 0.009 },
  { id: 44, category: "BANGUNAN", item: "Tangga lekuk dasar", priority: "Prioritas_3", weight: 0.0005 },
  { id: 45, category: "BANGUNAN", item: "Kekuatan struktur lantai lekuk dasar", priority: "Prioritas_3", weight: 0.0005 },
  { id: 46, category: "BANGUNAN", item: "Lekuk dasar - Tersedia rem pengaman", priority: "Prioritas_2", weight: 0.009 },
  { id: 47, category: "BANGUNAN", item: "Lekuk dasar - Tidak sebagai tempat kerja", priority: "Prioritas_3", weight: 0.0005 },
  { id: 48, category: "BANGUNAN", item: "Akses menuju lekuk dasar", priority: "Prioritas_2", weight: 0.009 },
  { id: 49, category: "BANGUNAN", item: "Lekuk dasar antar 2 Elevator", priority: "Prioritas_3", weight: 0.0005 },
  { id: 50, category: "BANGUNAN", item: "Daun pintu ruang luncur", priority: "Prioritas_3", weight: 0.0005 },
  { id: 51, category: "BANGUNAN", item: "Interlock / kunci kait pintu ruang luncur", priority: "Prioritas_2", weight: 0.009 },
  { id: 52, category: "BANGUNAN", item: "Kerataan lantai < 10 mm", priority: "Prioritas_2", weight: 0.009 },
  { id: 53, category: "BANGUNAN", item: "Sekat ruang luncur (2 sangkar)", priority: "Prioritas_3", weight: 0.0005 },
  { id: 54, category: "KERETA", item: "Kerangka - Dari baja dan kuat", priority: "Prioritas_3", weight: 0.0005 },
  { id: 55, category: "KERETA", item: "Badan kereta - Tertutup dan ada pintu", priority: "Prioritas_3", weight: 0.0005 },
  { id: 56, category: "KERETA", item: "Tinggi dinding ≥ 2000 mm", priority: "Prioritas_3", weight: 0.0005 },
  { id: 57, category: "KERETA", item: "Luas lantai - Sesuai jumlah penumpang", priority: "Prioritas_3", weight: 0.0005 },
  { id: 58, category: "KERETA", item: "Perluasan luas kereta - Pasien Max 6%", priority: "Prioritas_3", weight: 0.0005 },
  { id: 59, category: "KERETA", item: "Perluasan luas kereta - Barang Max 14%", priority: "Prioritas_3", weight: 0.0005 },
  { id: 60, category: "KERETA", item: "Pintu kereta - Kokoh, aman, otomatis", priority: "Prioritas_3", weight: 0.0005 },
  { id: 61, category: "KERETA", item: "Ukuran Pintu Kereta ≥ 700 x 2000 mm", priority: "Prioritas_3", weight: 0.0005 },
  { id: 62, category: "KERETA", item: "Pintu kereta - Kunci kait dan saklar pengaman", priority: "Prioritas_3", weight: 0.0005 },
  { id: 63, category: "KERETA", item: "Celah antar ambang pintu kereta", priority: "Prioritas_3", weight: 0.0005 },
  { id: 64, category: "KERETA", item: "Sisi luar kereta dg balok pemisah ruang luncur", priority: "Prioritas_3", weight: 0.0005 },
  { id: 65, category: "KERETA", item: "Alarm bell", priority: "Prioritas_1", weight: 0.04 },
  { id: 66, category: "KERETA", item: "Sumber tenaga cadangan (ARD)", priority: "Prioritas_1", weight: 0.04 },
  { id: 67, category: "KERETA", item: "Intercom", priority: "Prioritas_1", weight: 0.04 },
  { id: 68, category: "KERETA", item: "Ventilasi", priority: "Prioritas_2", weight: 0.009 },
  { id: 69, category: "KERETA", item: "Penerangan darurat", priority: "Prioritas_2", weight: 0.009 },
  { id: 70, category: "KERETA", item: "Panel operasi", priority: "Prioritas_3", weight: 0.0005 },
  { id: 71, category: "KERETA", item: "Petunjuk posisi sangkar", priority: "Prioritas_3", weight: 0.0005 },
  { id: 72, category: "KERETA", item: "Nama pembuat pada panel", priority: "Prioritas_3", weight: 0.0005 },
  { id: 73, category: "KERETA", item: "Kapasitas beban pada panel", priority: "Prioritas_3", weight: 0.0005 },
  { id: 74, category: "KERETA", item: "Rambu dilarang merokok", priority: "Prioritas_3", weight: 0.0005 },
  { id: 75, category: "KERETA", item: "Indikasi beban lebih", priority: "Prioritas_3", weight: 0.0005 },
  { id: 76, category: "KERETA", item: "Tombol buka dan tutup", priority: "Prioritas_3", weight: 0.0005 },
  { id: 77, category: "KERETA", item: "Tombol lantai pemberhentian", priority: "Prioritas_3", weight: 0.0005 },
  { id: 78, category: "KERETA", item: "Tombol bell alarm", priority: "Prioritas_3", weight: 0.0005 },
  { id: 79, category: "KERETA", item: "Intercom dua arah", priority: "Prioritas_2", weight: 0.009 },
  { id: 80, category: "KERETA", item: "Kekuatan atap kereta ≥ 200 Kg", priority: "Prioritas_3", weight: 0.0005 },
  { id: 81, category: "KERETA", item: "Pintu darurat atap kereta", priority: "Prioritas_3", weight: 0.0005 },
  { id: 82, category: "KERETA", item: "Pintu darurat samping kereta", priority: "Prioritas_3", weight: 0.0005 },
  { id: 83, category: "KERETA", item: "Pagar pengaman atap kereta - Warna kuning", priority: "Prioritas_3", weight: 0.0005 },
  { id: 84, category: "KERETA", item: "Pagar pengaman atap kereta - Kekuatan ≥ 90 Kg", priority: "Prioritas_3", weight: 0.0005 },
  { id: 85, category: "KERETA", item: "Pagar pengaman - Celah 300-850 mm, tinggi ≥ 700 mm", priority: "Prioritas_3", weight: 0.0005 },
  { id: 86, category: "KERETA", item: "Pagar pengaman - Celah > 850 mm, tinggi ≥ 1100 mm", priority: "Prioritas_3", weight: 0.0005 },
  { id: 87, category: "KERETA", item: "Penerangan atap kereta ≥ 100 Lux", priority: "Prioritas_3", weight: 0.0005 },
  { id: 88, category: "KERETA", item: "Tombol operasi manual", priority: "Prioritas_3", weight: 0.0005 },
  { id: 89, category: "KERETA", item: "Interior kereta - Bahan tidak mudah pecah", priority: "Prioritas_3", weight: 0.0005 },
  { id: 90, category: "GOVERNOR", item: "Penjepit tali / sabuk governor - Bekerja", priority: "Prioritas_1", weight: 0.09 },
  { id: 91, category: "GOVERNOR", item: "Saklar governor - Berfungsi", priority: "Prioritas_1", weight: 0.09 },
  { id: 92, category: "GOVERNOR", item: "Fungsi kecepatan rem pengaman 115% - 140%", priority: "Prioritas_2", weight: 0.009 },
  { id: 93, category: "GOVERNOR", item: "Rem pengaman - Dipasang pada sangkar", priority: "Prioritas_2", weight: 0.009 },
  { id: 94, category: "GOVERNOR", item: "Bentuk rem pengaman - Bukan sistem elektris", priority: "Prioritas_2", weight: 0.009 },
  { id: 95, category: "GOVERNOR", item: "Rem pengaman berangsur > 60 m/menit", priority: "Prioritas_2", weight: 0.009 },
  { id: 96, category: "GOVERNOR", item: "Rem pengaman mendadak < 60 m/menit", priority: "Prioritas_2", weight: 0.009 },
  { id: 97, category: "GOVERNOR", item: "Rem pengaman - Bekerja kebawah dan serempak", priority: "Prioritas_1", weight: 0.09 },
  { id: 98, category: "GOVERNOR", item: "Kecepatan ≥ 60 m/menit - Ada pemutus elektrik", priority: "Prioritas_2", weight: 0.009 },
  { id: 99, category: "GOVERNOR", item: "Saklar pengaman lintas batas - Berfungsi", priority: "Prioritas_1", weight: 0.04 },
  { id: 100, category: "GOVERNOR", item: "Alat pembatas beban lebih - Berfungsi", priority: "Prioritas_1", weight: 0.04 },
  { id: 101, category: "BOBOT IMBANG", item: "Bahan bobot imbang - Beton/Steel Block", priority: "Prioritas_3", weight: 0.0005 },
  { id: 102, category: "BOBOT IMBANG", item: "Pemasangan sekat pengaman bobot imbang", priority: "Prioritas_3", weight: 0.0005 },
  { id: 103, category: "BOBOT IMBANG", item: "Konstruksi rel pemandu - Kuat memandu jalan", priority: "Prioritas_3", weight: 0.0005 },
  { id: 104, category: "BOBOT IMBANG", item: "Jenis Peredam", priority: "Prioritas_3", weight: 0.0005 },
  { id: 105, category: "BOBOT IMBANG", item: "Fungsi peredaman - Meredam secara bertahap", priority: "Prioritas_1", weight: 0.04 },
  { id: 106, category: "BOBOT IMBANG", item: "Saklar pengaman untuk kecepatan ≥ 90 m/menit", priority: "Prioritas_2", weight: 0.009 },
  { id: 107, category: "LISTRIK", item: "Standar rangkaian instalasi listrik", priority: "Prioritas_2", weight: 0.009 },
  { id: 108, category: "LISTRIK", item: "Panel listrik - Khusus untuk elevator", priority: "Prioritas_2", weight: 0.009 },
  { id: 109, category: "LISTRIK", item: "Catu daya pengganti listrik otomatis (ARD)", priority: "Prioritas_1", weight: 0.04 },
  { id: 110, category: "LISTRIK", item: "Kabel grounding - Penampang ≥ 10 mm2", priority: "Prioritas_2", weight: 0.009 },
  { id: 111, category: "LISTRIK", item: "Kabel grounding - ≤ 5 Ω (ohm)", priority: "Prioritas_2", weight: 0.009 },
  { id: 112, category: "LISTRIK", item: "Alarm kebakaran - Terhubung dan otomatis", priority: "Prioritas_2", weight: 0.009 },
  { id: 115, category: "SENSOR", item: "Sensor gempa - > 10 lantai / 40 meter", priority: "Prioritas_2", weight: 0.009 },
  { id: 116, category: "SENSOR", item: "Input signal sensor gempa - Berhenti lantai terdekat", priority: "Prioritas_3", weight: 0.0005 },
];

const maintenanceItems = [
  { id: 1, item: "Kondisi ruang mesin" },
  { id: 2, item: "Kondisi controller" },
  { id: 3, item: "Kondisi mesin" },
  { id: 4, item: "Kondisi Governor" },
  { id: 5, item: "Kondisi alat evakuasi manual" },
  { id: 6, item: "Kondisi wire rope" },
  { id: 7, item: "Kondisi car top" },
  { id: 8, item: "Kondisi safety line" },
  { id: 9, item: "Kondisi level rata lantai" },
  { id: 10, item: "Kondisi interior (Fan & Light)" },
  { id: 11, item: "Kondisi tombol dan indicator sangkar" },
  { id: 12, item: "Kondisi buka tutup car door" },
  { id: 13, item: "Kondisi sensor pintu" },
  { id: 14, item: "Kondisi hall door" },
  { id: 15, item: "Kondisi tombol dan indicator hall call" },
  { id: 16, item: "Kondisi ruang lekuk dasar (PIT)" },
  { id: 17, item: "Kondisi dan fungsi interphone" },
  { id: 18, item: "Kondisi dan fungsi bell/alarm" },
  { id: 19, item: "Kondisi dan fungsi lampu emergency" },
  { id: 20, item: "Data log book" },
];

type StatusType = 'Good' | 'Not Good' | 'N/A' | '';
type QCStatusType = 'Approved' | 'Revision Required' | '';
type UserRole = 'maintenance' | 'qc' | 'sales' | null;
type MenuType = 'dashboard' | 'newPemeriksaan' | 'newPemeliharaan' | 'viewReport' | 'review';
type SortOrder = 'terbaru' | 'terlama';
type ReportType = 'pemeriksaan' | 'pemeliharaan';

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
  maintenanceData: {
    id: number;
    status: StatusType;
    description: string;
    finding: string;
    photoBefore: string | null;
    photoAfter: string | null;
    repairNote: string;
    isApproved: boolean;
  }[];
}

interface Report {
  id: string;
  reportType: ReportType;
  jumlahUnit: number;
  units: UnitData[];
  teknisiName: string;
  inspectionData: {
    id: number;
    status: StatusType;
    finding: string;
    photoBefore: string | null;
    photoAfter: string | null;
    repairNote: string;
    isApproved: boolean;
  }[];
  qcNote: string; // Tambahan catatan QC
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
  unitNumber?: string;
  projectCode?: string;
  customerName?: string;
  buildingLocation?: string;
  elevatorType?: string;
  elevatorBrand?: string;
  elevatorModel?: string;
  capacity?: string;
  speed?: string;
  inspectionDate?: string;
  qcName?: string;
}

const calculateItemScore = (status: StatusType, weight: number, isApproved: boolean = false): number => {
  if (status === 'Good' || status === 'N/A') {
    return weight;
  } else if (status === 'Not Good' && isApproved) {
    return weight;
  } else if (status === 'Not Good' && !isApproved) {
    return -weight;
  }
  return 0;
};

// ============================================
// LOAD DATA DARI FIRESTORE
// ============================================
async function loadDataFromFirestore() {
  try {
    const q = query(collection(db, 'reports'), orderBy('createdAt', 'desc'));
    const snapshot = await getDocs(q);
    const data: Report[] = [];
    snapshot.forEach((doc) => {
      data.push({ ...doc.data(), id: doc.id } as Report);
    });
    return data;
  } catch (error) {
    console.error('Error loading from Firestore:', error);
    return [];
  }
}

// ============================================
// SAVE DATA KE FIRESTORE
// ============================================
async function saveDataToFirestore(report: Report) {
  try {
    const { id, ...data } = report;
    if (id && id.length > 20) {
      await updateDoc(doc(db, 'reports', id), data);
    } else {
      const docRef = await addDoc(collection(db, 'reports'), data);
      return { ...report, id: docRef.id };
    }
    return report;
  } catch (error) {
    console.error('Error saving to Firestore:', error);
    throw error;
  }
}

// ============================================
// DELETE DATA DARI FIRESTORE
// ============================================
async function deleteDataFromFirestore(id: string) {
  try {
    await deleteDoc(doc(db, 'reports', id));
  } catch (error) {
    console.error('Error deleting from Firestore:', error);
    throw error;
  }
}

// ============================================
// REALTIME LISTENER
// ============================================
function subscribeToReports(callback: (data: Report[]) => void) {
  const q = query(collection(db, 'reports'), orderBy('createdAt', 'desc'));
  return onSnapshot(q, (snapshot) => {
    const data: Report[] = [];
    snapshot.forEach((doc) => {
      data.push({ ...doc.data(), id: doc.id } as Report);
    });
    callback(data);
  });
}

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

    const users = {
      'qc': { password: 'qc8#Kx92!mN4', role: 'qc' },
      'mainten': { password: 'mai7&Hp3@wQ9', role: 'maintenance' },
      'sales': { password: 'sAl2$Rt6#vB1', role: 'sales' }
    };

    if (users[username as keyof typeof users]?.password === password) {
      const role = users[username as keyof typeof users].role as UserRole;
      if (role === selectedRole) {
        onLogin(role);
        setLoading(false);
        return;
      } else {
        setError(`Role tidak sesuai! Anda login sebagai ${role}, tetapi memilih ${selectedRole}`);
        setLoading(false);
        return;
      }
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

  const isQC = role === 'qc';
  const isMaintenance = role === 'maintenance';
  const isSales = role === 'sales';

  const showToast = (message: string, type: 'success' | 'error' | 'info' = 'success') => {
    setToast({ message, type });
    setTimeout(() => setToast(null), 3000);
  };

  useEffect(() => {
    setLoading(true);
    const unsubscribe = subscribeToReports((data) => {
      setReports(data);
      setLoading(false);
    });
    return () => unsubscribe();
  }, []);

  const handleNewPemeriksaan = () => {
    setActiveNote('NEW DATA PEMERIKSAAN');
    const newReport: any = {
      reportType: 'pemeriksaan',
      jumlahUnit: 1,
      units: [],
      teknisiName: '',
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
      qcNote: '',
      inspectionData: inspectionItems.map(item => ({
        id: item.id,
        status: '' as StatusType,
        finding: '',
        photoBefore: null,
        photoAfter: null,
        repairNote: '',
        isApproved: false,
      })),
      attachment: null,
      attachmentName: null,
      qcVerification: {
        qcName: '',
        qcStatus: '' as QCStatusType,
        qcNote: '',
        verifiedAt: '',
      },
      submittedBy: '',
      submittedAt: '',
      status: 'draft' as const,
      createdAt: new Date().toISOString(),
    };
    setSelectedReport(newReport as Report);
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
        status: '' as StatusType,
        description: '',
        finding: '',
        photoBefore: null,
        photoAfter: null,
        repairNote: '',
        isApproved: false,
      })),
    };

    const newReport: any = {
      reportType: 'pemeliharaan',
      jumlahUnit: 1,
      units: [defaultUnit],
      teknisiName: '',
      inspectionData: [],
      attachment: null,
      attachmentName: null,
      qcVerification: {
        qcName: '',
        qcStatus: '' as QCStatusType,
        qcNote: '',
        verifiedAt: '',
      },
      submittedBy: '',
      submittedAt: '',
      status: 'draft' as const,
      createdAt: new Date().toISOString(),
    };
    setSelectedReport(newReport as Report);
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
    if (isSales) {
      setIsReadOnly(true);
    } else if (report.status === 'approved') {
      setIsReadOnly(true);
    } else if (report.status === 'revision' && isMaintenance) {
      setIsReadOnly(false);
    } else if (report.status === 'revision' && !isMaintenance) {
      setIsReadOnly(true);
    } else if (isMaintenance) {
      setIsReadOnly(report.status !== 'qc_approved' && report.status !== 'revision');
    } else if (isQC) {
      setIsReadOnly(report.status !== 'draft' && report.status !== 'maintenance_done' && report.status !== 'revision');
    } else {
      setIsReadOnly(true);
    }
    setCurrentMenu('viewReport');
    setActiveNote('');
  };

  const handleSaveReport = async (data: any) => {
    try {
      await saveDataToFirestore(data);
      showToast('Data berhasil disimpan!', 'success');
      setCurrentMenu('dashboard');
      setActiveNote('');
    } catch (error) {
      showToast('Gagal menyimpan data!', 'error');
    }
  };

  const handleDeleteReport = async (id: string) => {
    if (!isQC) {
      showToast('Hanya QC yang bisa menghapus laporan!', 'error');
      return;
    }
    if (confirm('Yakin ingin menghapus laporan ini?')) {
      try {
        await deleteDataFromFirestore(id);
        showToast('Laporan berhasil dihapus!', 'success');
      } catch (error) {
        showToast('Gagal menghapus laporan!', 'error');
      }
    }
  };

  const handleCancel = () => {
    setSelectedReport(null);
    setIsReadOnly(false);
    setCurrentMenu('dashboard');
    setActiveNote('');
  };

  const handlePrintPDF = () => {
    window.print();
  };

  const getFilteredReports = () => {
    let filtered: Report[] = [];
    if (isQC) {
      filtered = reports;
    } else if (isMaintenance) {
      filtered = reports.filter(r => 
        r.status === 'qc_approved' || 
        r.status === 'revision' ||
        r.status === 'maintenance_done' || 
        r.status === 'approved'
      );
    } else if (isSales) {
      filtered = reports.filter(r => 
        r.reportType === 'pemeriksaan' &&
        (r.status === 'qc_approved' || 
         r.status === 'maintenance_done' || 
         r.status === 'approved')
      );
    }
    return filtered.sort((a, b) => {
      const dateA = new Date(a.createdAt).getTime();
      const dateB = new Date(b.createdAt).getTime();
      return sortOrder === 'terbaru' ? dateB - dateA : dateA - dateB;
    });
  };

  const filteredReports = getFilteredReports();

  const getStats = () => {
    const all = reports;
    return {
      total: all.length,
      draft: all.filter(r => r.status === 'draft').length,
      qcApproved: all.filter(r => r.status === 'qc_approved').length,
      revision: all.filter(r => r.status === 'revision').length,
      maintenanceDone: all.filter(r => r.status === 'maintenance_done').length,
      approved: all.filter(r => r.status === 'approved').length,
    };
  };
  const stats = getStats();

  const formatDate = (dateString: string) => {
    if (!dateString) return '-';
    try {
      const date = new Date(dateString);
      return date.toLocaleDateString('id-ID', {
        day: '2-digit',
        month: 'short',
        year: 'numeric',
        hour: '2-digit',
        minute: '2-digit'
      });
    } catch {
      return '-';
    }
  };

  const totalTugas = isMaintenance ? reports.filter(r => r.status !== 'approved' && (r.status === 'qc_approved' || r.status === 'revision')).length : 0;
  const needReview = isQC ? reports.filter(r => r.status === 'maintenance_done' || r.status === 'revision').length : 0;

  return (
    <div className="min-h-screen bg-gradient-to-br from-slate-50 via-white to-gray-50">
      {toast && <Toast message={toast.message} type={toast.type} onClose={() => setToast(null)} />}
      
      <div className="container mx-auto px-4 py-6 max-w-7xl">
        {/* HEADER */}
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
                {isMaintenance && (
                  <p className="text-sm text-blue-600 mt-1">{totalTugas} tugas aktif</p>
                )}
                {isSales && (
                  <p className="text-sm text-amber-600 mt-1">Mode Lihat</p>
                )}
                {isQC && needReview > 0 && (
                  <p className="text-sm text-orange-500 mt-1 animate-pulse">{needReview} data perlu direview dari Maintenance</p>
                )}
              </div>
            </div>
            <div className="mt-4 md:mt-0 flex flex-wrap items-center gap-3">
              <span className={`px-4 py-1.5 rounded-full text-sm font-medium ${
                isQC ? 'bg-emerald-100 text-emerald-700' : 
                isMaintenance ? 'bg-blue-100 text-blue-700' : 
                'bg-amber-100 text-amber-700'
              }`}>
                {isQC ? 'QC' : isMaintenance ? 'Maintenance' : 'Sales'}
              </span>
              {isQC && stats.maintenanceDone > 0 && (
                <span className="px-3 py-1 rounded-full text-xs font-medium bg-amber-100 text-amber-700 animate-pulse">
                  {stats.maintenanceDone} verifikasi
                </span>
              )}
              <button 
                onClick={onLogout} 
                className="text-sm bg-red-500 hover:bg-red-600 text-white px-4 py-2 rounded-xl transition shadow-sm"
              >
                Keluar
              </button>
            </div>
          </div>

          {/* STATISTIK */}
          <div className="mt-4 grid grid-cols-3 md:grid-cols-6 gap-2">
            <div className="bg-gray-50/80 rounded-xl p-3 text-center">
              <div className="text-xl font-bold text-gray-700">{stats.total}</div>
              <div className="text-[10px] text-gray-400 uppercase tracking-wider">Total</div>
            </div>
            <div className="bg-gray-50/80 rounded-xl p-3 text-center">
              <div className="text-xl font-bold text-gray-400">{stats.draft}</div>
              <div className="text-[10px] text-gray-400 uppercase tracking-wider">Draft</div>
            </div>
            <div className="bg-blue-50/80 rounded-xl p-3 text-center">
              <div className="text-xl font-bold text-blue-600">{stats.qcApproved}</div>
              <div className="text-[10px] text-blue-400 uppercase tracking-wider">Maintenance</div>
            </div>
            <div className="bg-amber-50/80 rounded-xl p-3 text-center">
              <div className="text-xl font-bold text-amber-600">{stats.revision}</div>
              <div className="text-[10px] text-amber-400 uppercase tracking-wider">Revisi</div>
            </div>
            <div className="bg-orange-50/80 rounded-xl p-3 text-center">
              <div className="text-xl font-bold text-orange-600">{stats.maintenanceDone}</div>
              <div className="text-[10px] text-orange-400 uppercase tracking-wider">Verifikasi</div>
            </div>
            <div className="bg-emerald-50/80 rounded-xl p-3 text-center">
              <div className="text-xl font-bold text-emerald-600">{stats.approved}</div>
              <div className="text-[10px] text-emerald-400 uppercase tracking-wider">Final</div>
            </div>
          </div>
        </header>

        {currentMenu === 'dashboard' ? (
          <div className="space-y-4">
            <div className="flex flex-wrap items-center justify-between gap-3">
              <div className="flex flex-wrap gap-3">
                {isQC && (
                  <>
                    <button 
                      onClick={handleNewPemeriksaan} 
                      className="bg-blue-600 hover:bg-blue-700 text-white px-5 py-2.5 rounded-xl flex items-center gap-2 shadow-sm hover:shadow-md transition hover:scale-[1.02] text-sm font-medium"
                    >
                      New Pemeriksaan
                    </button>
                    <button 
                      onClick={handleNewPemeliharaan} 
                      className="bg-emerald-600 hover:bg-emerald-700 text-white px-5 py-2.5 rounded-xl flex items-center gap-2 shadow-sm hover:shadow-md transition hover:scale-[1.02] text-sm font-medium"
                    >
                      New Pemeliharaan
                    </button>
                  </>
                )}
                {isMaintenance && (
                  <p className="text-sm text-gray-500 self-center">Tugas dari QC</p>
                )}
                {isSales && (
                  <p className="text-sm text-amber-500 self-center">Mode Lihat Saja</p>
                )}
                {isQC && needReview > 0 && (
                  <span className="text-sm text-orange-500 self-center font-medium animate-pulse">{needReview} data perlu review</span>
                )}
              </div>
              <div className="flex items-center gap-2">
                <span className="text-xs text-gray-400">Urutkan:</span>
                <button 
                  onClick={() => setSortOrder('terbaru')} 
                  className={`px-3 py-1 rounded-lg text-xs transition ${
                    sortOrder === 'terbaru' 
                      ? 'bg-blue-600 text-white' 
                      : 'bg-gray-100 text-gray-600 hover:bg-gray-200'
                  }`}
                >
                  Terbaru
                </button>
                <button 
                  onClick={() => setSortOrder('terlama')} 
                  className={`px-3 py-1 rounded-lg text-xs transition ${
                    sortOrder === 'terlama' 
                      ? 'bg-blue-600 text-white' 
                      : 'bg-gray-100 text-gray-600 hover:bg-gray-200'
                  }`}
                >
                  Terlama
                </button>
              </div>
            </div>

            <div className="bg-white/80 backdrop-blur-sm rounded-2xl border border-gray-100/50 overflow-hidden shadow-sm">
              <div className="px-6 py-4 border-b border-gray-100 flex justify-between items-center">
                <div>
                  <h2 className="font-semibold text-gray-700">
                    {isQC ? 'Semua Laporan' : isMaintenance ? 'Tugas Maintenance' : 'Data Pemeriksaan'}
                  </h2>
                  <span className="text-xs text-gray-400">{filteredReports.length} laporan</span>
                </div>
                <span className="text-[10px] text-emerald-500 font-medium bg-emerald-50 px-3 py-1 rounded-full">Realtime</span>
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
                        const firstUnit = report.units?.[0];
                        const unitDisplay = report.reportType === 'pemeliharaan' && report.jumlahUnit > 1 ? `${report.jumlahUnit} Unit` : firstUnit?.unitNumber || report.unitNumber || '-';
                        const projectDisplay = report.reportType === 'pemeliharaan' && report.jumlahUnit > 1 ? `${report.jumlahUnit} Proyek` : firstUnit?.projectCode || report.projectCode || '-';
                        const customerDisplay = report.reportType === 'pemeliharaan' && report.jumlahUnit > 1 ? `${report.jumlahUnit} Pelanggan` : firstUnit?.customerName || report.customerName || '-';

                        const isReview = report.status === 'maintenance_done' || report.status === 'revision';

                        return (
                          <tr key={report.id} className="hover:bg-gray-50/50 transition">
                            <td className="px-4 py-3 text-gray-500 text-xs whitespace-nowrap">{formatDate(report.createdAt)}</td>
                            <td className="px-4 py-3 font-medium text-gray-700 text-sm">{unitDisplay}</td>
                            <td className="px-4 py-3 text-gray-600 text-sm">{projectDisplay}</td>
                            <td className="px-4 py-3 text-gray-600 text-sm">{customerDisplay}</td>
                            <td className="px-4 py-3">
                              <span className={`px-2.5 py-1 rounded-full text-[10px] font-medium ${
                                report.reportType === 'pemeriksaan' 
                                  ? 'bg-blue-100 text-blue-700' 
                                  : 'bg-emerald-100 text-emerald-700'
                              }`}>
                                {report.reportType === 'pemeriksaan' ? 'Pemeriksaan' : 'Pemeliharaan'}
                              </span>
                            </td>
                            {isQC && <td className="px-4 py-3 text-sm text-gray-500">{report.teknisiName || '-'}</td>}
                            <td className="px-4 py-3">
                              <span className={`inline-flex px-2.5 py-1 rounded-full text-[10px] font-medium ${
                                report.status === 'draft' ? 'bg-gray-200 text-gray-600' : 
                                report.status === 'qc_approved' ? 'bg-blue-100 text-blue-700' : 
                                report.status === 'revision' ? 'bg-amber-100 text-amber-700' : 
                                report.status === 'maintenance_done' ? 'bg-orange-100 text-orange-700' : 
                                'bg-emerald-100 text-emerald-700'
                              }`}>
                                {report.status === 'draft' ? 'Draft' : 
                                 report.status === 'qc_approved' ? 'Maintenance' : 
                                 report.status === 'revision' ? 'Revisi' : 
                                 report.status === 'maintenance_done' ? 'Verifikasi' : 
                                 'Final'}
                              </span>
                            </td>
                            <td className="px-4 py-3 flex items-center gap-2">
                              {isQC && isReview ? (
                                <button 
                                  onClick={() => handleReviewReport(report)} 
                                  className="bg-amber-50 hover:bg-amber-100 text-amber-600 px-3 py-1 rounded-xl text-sm font-medium transition"
                                >
                                  Review
                                </button>
                              ) : (
                                <button 
                                  onClick={() => handleViewReport(report)} 
                                  className="bg-blue-50 hover:bg-blue-100 text-blue-600 px-3 py-1 rounded-xl text-sm font-medium transition"
                                >
                                  Lihat
                                </button>
                              )}
                              {isQC && report.status !== 'approved' && (
                                <button 
                                  onClick={() => handleDeleteReport(report.id)} 
                                  className="text-red-400 hover:text-red-600 ml-2 text-lg transition"
                                  title="Hapus"
                                >
                                  🗑️
                                </button>
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
            {activeNote && (
              <div className="mb-4 p-3 bg-blue-50 border border-blue-200 rounded-xl text-blue-700 font-medium text-sm text-center">
                {activeNote}
              </div>
            )}
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
// KOMPONEN REPORT FORM
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
  const [selectedJumlahUnit, setSelectedJumlahUnit] = useState<number>(report?.jumlahUnit || 1);
  const [formData, setFormData] = useState<any>(() => {
    if (report) return report;
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
        status: '' as StatusType,
        description: '',
        finding: '',
        photoBefore: null,
        photoAfter: null,
        repairNote: '',
        isApproved: false,
      })),
    };
    return {
      id: Date.now().toString(),
      reportType: reportType || 'pemeriksaan',
      jumlahUnit: 1,
      units: [defaultUnit],
      teknisiName: '',
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
      qcNote: '',
      inspectionData: inspectionItems.map(item => ({
        id: item.id,
        status: '' as StatusType,
        finding: '',
        photoBefore: null,
        photoAfter: null,
        repairNote: '',
        isApproved: false,
      })),
      attachment: null,
      attachmentName: null,
      qcVerification: { qcName: '', qcStatus: '' as QCStatusType, qcNote: '', verifiedAt: '' },
      submittedBy: '',
      submittedAt: '',
      status: 'draft' as const,
      createdAt: new Date().toISOString(),
    };
  });

  const calculateTotalScore = (inspectionData: any[]) => {
    let totalWeight = 0, achievedWeight = 0;
    inspectionData.forEach((item: any) => {
      const originalItem = inspectionItems.find(i => i.id === item.id);
      if (originalItem) {
        totalWeight += originalItem.weight;
        if (item.status === 'Good' || item.status === 'N/A') achievedWeight += originalItem.weight;
        else if (item.status === 'Not Good' && item.isApproved) achievedWeight += originalItem.weight;
      }
    });
    const score = totalWeight > 0 ? (achievedWeight / totalWeight) * 100 : 0;
    return Math.round(score * 100) / 100;
  };

  const calculateHitung = (item: any, originalItem: any) => {
    if (!originalItem) return 0;
    if (item.status === 'Good' || item.status === 'N/A') return originalItem.weight;
    if (item.status === 'Not Good' && item.isApproved) return originalItem.weight;
    if (item.status === 'Not Good' && !item.isApproved) return -originalItem.weight;
    return 0;
  };

  const handleStatusChange = (id: number, status: StatusType) => {
    if (!isQC || isReadOnly || formData.status !== 'draft') return;
    setFormData((prev: any) => ({
      ...prev,
      inspectionData: prev.inspectionData.map((item: any) =>
        item.id === id ? { ...item, status, finding: status === 'Good' || status === 'N/A' ? '' : item.finding, isApproved: false } : item
      ),
    }));
  };

  const handleFindingChange = (id: number, finding: string) => {
    if (!isQC || isReadOnly || formData.status !== 'draft') return;
    setFormData((prev: any) => ({
      ...prev,
      inspectionData: prev.inspectionData.map((item: any) =>
        item.id === id ? { ...item, finding } : item
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

  const handleApproveMaintenanceItem = (unitIndex: number, id: number) => {
    if (!isQC || isReadOnly || (formData.status !== 'maintenance_done' && formData.status !== 'revision')) return;
    const item = formData.units[unitIndex].maintenanceData.find((i: any) => i.id === id);
    if (!item || item.status !== 'Not Good' || item.isApproved) return;
    setFormData((prev: any) => {
      const newUnits = [...prev.units];
      const newData = newUnits[unitIndex].maintenanceData.map((item: any) =>
        item.id === id ? { ...item, isApproved: true } : item
      );
      newUnits[unitIndex] = { ...newUnits[unitIndex], maintenanceData: newData };
      return { ...prev, units: newUnits };
    });
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

  const handleJumlahUnitChange = (value: number) => {
    if (!isQC || isReadOnly || formData.status !== 'draft') return;
    if (isPemeriksaan) return;
    const newCount = Math.min(Math.max(value, 1), 100);
    setSelectedJumlahUnit(newCount);
    const currentUnits = formData.units || [];
    let newUnits = [...currentUnits];
    if (newCount > currentUnits.length) {
      for (let i = currentUnits.length; i < newCount; i++) {
        newUnits.push({
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
            status: '' as StatusType,
            description: '',
            finding: '',
            photoBefore: null,
            photoAfter: null,
            repairNote: '',
            isApproved: false,
          })),
        });
      }
    } else if (newCount < currentUnits.length) {
      newUnits = newUnits.slice(0, newCount);
    }
    setFormData((prev: any) => ({ ...prev, jumlahUnit: newCount, units: newUnits }));
  };

  const updateUnitData = (index: number, field: string, value: string) => {
    if (isReadOnly || !isQC || formData.status !== 'draft') return;
    setFormData((prev: any) => {
      const newUnits = [...prev.units];
      newUnits[index] = { ...newUnits[index], [field]: value };
      return { ...prev, units: newUnits };
    });
  };

  const handleMaintenanceStatusChange = (unitIndex: number, id: number, status: StatusType) => {
    if (!isQC || isReadOnly || formData.status !== 'draft') return;
    setFormData((prev: any) => {
      const newUnits = [...prev.units];
      const newData = newUnits[unitIndex].maintenanceData.map((item: any) =>
        item.id === id ? { ...item, status, isApproved: false } : item
      );
      newUnits[unitIndex] = { ...newUnits[unitIndex], maintenanceData: newData };
      return { ...prev, units: newUnits };
    });
  };

  const handleMaintenanceFindingChange = (unitIndex: number, id: number, finding: string) => {
    if (!isQC || isReadOnly || formData.status !== 'draft') return;
    setFormData((prev: any) => {
      const newUnits = [...prev.units];
      const newData = newUnits[unitIndex].maintenanceData.map((item: any) =>
        item.id === id ? { ...item, finding } : item
      );
      newUnits[unitIndex] = { ...newUnits[unitIndex], maintenanceData: newData };
      return { ...prev, units: newUnits };
    });
  };

  const handleMaintenancePhotoBeforeChange = (unitIndex: number, id: number, event: React.ChangeEvent<HTMLInputElement>) => {
    if (!isQC || isReadOnly || formData.status !== 'draft') return;
    const file = event.target.files?.[0];
    if (file) {
      const reader = new FileReader();
      reader.onloadend = () => {
        setFormData((prev: any) => {
          const newUnits = [...prev.units];
          const newData = newUnits[unitIndex].maintenanceData.map((item: any) =>
            item.id === id ? { ...item, photoBefore: reader.result as string } : item
          );
          newUnits[unitIndex] = { ...newUnits[unitIndex], maintenanceData: newData };
          return { ...prev, units: newUnits };
        });
      };
      reader.readAsDataURL(file);
    }
  };

  const handleMaintenancePhotoAfterChange = (unitIndex: number, id: number, event: React.ChangeEvent<HTMLInputElement>) => {
    if (!isMaintenance || isReadOnly) return;
    const file = event.target.files?.[0];
    if (file) {
      const reader = new FileReader();
      reader.onloadend = () => {
        setFormData((prev: any) => {
          const newUnits = [...prev.units];
          const newData = newUnits[unitIndex].maintenanceData.map((item: any) =>
            item.id === id ? { ...item, photoAfter: reader.result as string } : item
          );
          newUnits[unitIndex] = { ...newUnits[unitIndex], maintenanceData: newData };
          return { ...prev, units: newUnits };
        });
      };
      reader.readAsDataURL(file);
    }
  };

  const handleMaintenanceRepairNoteChange = (unitIndex: number, id: number, note: string) => {
    if (!isMaintenance || isReadOnly) return;
    setFormData((prev: any) => {
      const newUnits = [...prev.units];
      const newData = newUnits[unitIndex].maintenanceData.map((item: any) =>
        item.id === id ? { ...item, repairNote: note } : item
      );
      newUnits[unitIndex] = { ...newUnits[unitIndex], maintenanceData: newData };
      return { ...prev, units: newUnits };
    });
  };

  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    onSave(formData);
  };

  const handleKirim = (action: 'submit' | 'approve' | 'revision') => {
    setIsSubmitting(true);
    let updatedData: any;
    if (isQC && formData.status === 'draft') {
      if (isPemeriksaan) {
        const emptyItems = formData.inspectionData.filter((item: any) => !item.status);
        if (emptyItems.length > 0) { alert(`⚠️ ${emptyItems.length} item belum diisi status.`); setIsSubmitting(false); return; }
        const invalidItems = formData.inspectionData.filter((item: any) => (item.status === 'Not Good') && !item.finding);
        if (invalidItems.length > 0) { alert(`⚠️ ${invalidItems.length} item Not Good tanpa temuan.`); setIsSubmitting(false); return; }
      } else if (isPemeliharaan) {
        let hasError = false;
        formData.units.forEach((unit: UnitData, idx: number) => {
          const emptyItems = unit.maintenanceData.filter((item: any) => !item.status);
          if (emptyItems.length > 0) { alert(`⚠️ Unit ${idx + 1}: ${emptyItems.length} item belum diisi.`); hasError = true; }
        });
        if (hasError) { setIsSubmitting(false); return; }
      }
      updatedData = { ...formData, submittedBy: 'qc', submittedAt: new Date().toLocaleString('id-ID'), status: 'qc_approved' };
      setTimeout(() => { onSave(updatedData); setIsSubmitting(false); alert('Data dikirim ke Maintenance!'); }, 500);
      return;
    } else if (isQC && formData.status === 'maintenance_done') {
      if (action === 'approve') {
        if (isPemeriksaan) {
          const unapproved = formData.inspectionData.filter((item: any) => item.status === 'Not Good' && !item.isApproved);
          if (unapproved.length > 0) { alert(`⚠️ ${unapproved.length} item Not Good belum di-approve!`); setIsSubmitting(false); return; }
        } else if (isPemeliharaan) {
          let unapprovedCount = 0;
          formData.units.forEach((unit: UnitData) => {
            unapprovedCount += unit.maintenanceData.filter((item: any) => item.status === 'Not Good' && !item.isApproved).length;
          });
          if (unapprovedCount > 0) { alert(`⚠️ ${unapprovedCount} item Not Good belum di-approve!`); setIsSubmitting(false); return; }
        }
        updatedData = { ...formData, qcVerification: { ...formData.qcVerification, verifiedAt: new Date().toLocaleString('id-ID') }, status: 'approved' };
        setTimeout(() => { onSave(updatedData); setIsSubmitting(false); alert('Laporan FINAL APPROVED!'); }, 500);
      } else if (action === 'revision') {
        updatedData = { ...formData, qcVerification: { ...formData.qcVerification, verifiedAt: new Date().toLocaleString('id-ID'), qcStatus: 'Revision Required' }, status: 'revision' };
        setTimeout(() => { onSave(updatedData); setIsSubmitting(false); alert('REVISION! Dikirim ke Maintenance.'); }, 500);
      }
    } else if (isMaintenance && (formData.status === 'qc_approved' || formData.status === 'revision')) {
      if (!formData.teknisiName || formData.teknisiName.trim() === '') { alert('Silakan isi nama teknisi!'); setIsSubmitting(false); return; }
      if (isPemeriksaan) {
        const invalid = formData.inspectionData.filter((item: any) => item.status === 'Not Good' && (!item.photoAfter || !item.repairNote));
        if (invalid.length > 0) { alert(`⚠️ ${invalid.length} item Not Good tanpa foto/catatan.`); setIsSubmitting(false); return; }
      } else if (isPemeliharaan) {
        let hasError = false;
        formData.units.forEach((unit: UnitData, idx: number) => {
          const invalid = unit.maintenanceData.filter((item: any) => item.status === 'Not Good' && (!item.photoAfter || !item.repairNote));
          if (invalid.length > 0) { alert(`Unit ${idx + 1}: ${invalid.length} item Not Good tanpa foto/catatan.`); hasError = true; }
        });
        if (hasError) { setIsSubmitting(false); return; }
      }
      updatedData = { ...formData, submittedBy: 'maintenance', submittedAt: new Date().toLocaleString('id-ID'), status: 'maintenance_done' };
      setTimeout(() => { onSave(updatedData); setIsSubmitting(false); alert(`Perbaikan selesai oleh ${formData.teknisiName}!`); }, 500);
    } else {
      alert('Status tidak sesuai.');
      setIsSubmitting(false);
      return;
    }
  };

  const canEdit = !isReadOnly;
  const isApproved = formData.status === 'approved' || formData.status === 'qc_approved';
  const typeLabel = isPemeriksaan ? 'Pemeriksaan' : 'Pemeliharaan';
  const allNotGoodApproved = isPemeriksaan ? formData.inspectionData.filter((item: any) => item.status === 'Not Good').every((item: any) => item.isApproved) : true;
  const canFinalApprove = formData.status === 'maintenance_done' && allNotGoodApproved;
  const totalScore = isPemeriksaan ? calculateTotalScore(formData.inspectionData) : 0;

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
          <p className="text-xs text-gray-400 mt-1">Isi nama teknisi pelaksana</p>
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

  return (
    <form onSubmit={handleSubmit} className="space-y-6">
      <div className="flex flex-wrap justify-between items-center gap-2">
        <button type="button" onClick={onCancel} className="bg-gray-100 hover:bg-gray-200 text-gray-600 px-4 py-2 rounded-xl text-sm font-medium transition">
          Kembali
        </button>
        <div className="flex flex-wrap gap-2">
          {canEdit && (
            <button type="button" onClick={() => onDelete?.(formData.id)} className="bg-red-500 hover:bg-red-600 text-white px-4 py-2 rounded-xl text-sm transition">Hapus</button>
          )}
          {isQC && formData.status === 'draft' && (
            <button type="button" onClick={() => handleKirim('submit')} disabled={isSubmitting} className="bg-blue-600 hover:bg-blue-700 text-white px-5 py-2 rounded-xl text-sm font-medium transition">Submit</button>
          )}
          {isQC && formData.status === 'maintenance_done' && (
            <>
              <button type="button" onClick={() => handleKirim('revision')} disabled={isSubmitting} className="bg-amber-500 hover:bg-amber-600 text-white px-5 py-2 rounded-xl text-sm font-medium transition">Revisi</button>
              <button type="button" onClick={() => { if (!canFinalApprove) { alert('Approve semua Not Good dulu!'); return; } handleKirim('approve'); }} disabled={!canFinalApprove || isSubmitting} className={`px-5 py-2 rounded-xl text-sm font-medium transition ${canFinalApprove && !isSubmitting ? 'bg-emerald-600 hover:bg-emerald-700 text-white' : 'bg-gray-300 text-gray-500 cursor-not-allowed'}`}>Approve</button>
            </>
          )}
          {(isMaintenance && (formData.status === 'qc_approved' || formData.status === 'revision')) && (
            <button type="button" onClick={() => { if (!formData.teknisiName) { alert('Isi nama teknisi!'); return; } handleKirim('submit'); }} className="bg-orange-500 hover:bg-orange-600 text-white px-5 py-2 rounded-xl text-sm font-medium transition">Selesai</button>
          )}
          {isQC && isApproved && onPrintPDF && (
            <button type="button" onClick={onPrintPDF} className="bg-rose-500 hover:bg-rose-600 text-white px-5 py-2 rounded-xl text-sm font-medium transition">PDF</button>
          )}
          {isReadOnly && <span className="text-xs text-gray-400 self-center">Read-Only</span>}
          {isQC && formData.status === 'maintenance_done' && (
            <span className={`text-xs self-center ${allNotGoodApproved ? 'text-emerald-500' : 'text-amber-500'}`}>
              {allNotGoodApproved ? 'Siap di-approve' : 'Approve Not Good'}
            </span>
          )}
          {isReview && (
            <span className="text-xs text-orange-500 self-center font-medium">Mode Review - Data dari Maintenance</span>
          )}
        </div>
      </div>

      <div id="print-content">
        <div className="text-center border-b pb-4 mb-6 print:block hidden">
          <div className="flex justify-center mb-2">
             <Image src="/logo_louser_2022_1696999044 (1).png" alt="Logo Louser" width={150} height={50} className="object-contain" />
          </div>
          <p className="text-sm text-gray-500">Elevator Quality Control System</p>
          <p className="text-xs text-gray-400">Laporan {typeLabel}</p>
        </div>

        {renderTeknisiInput()}

        {isPemeriksaan && (
          <>
            <section className="bg-gray-50/50 rounded-xl p-6 mb-6 border border-gray-100">
              <h2 className="text-sm font-semibold text-gray-600 mb-4">Data Umum</h2>
              <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                <div>
                  <label className="block text-xs font-medium text-gray-500 uppercase tracking-wider mb-1">No Unit</label>
                  <input
                    type="text"
                    value={formData.unitNumber || ''}
                    onChange={(e) => setFormData((prev: any) => ({ ...prev, unitNumber: e.target.value }))}
                    className="w-full px-3 py-2 bg-white border border-gray-200 rounded-xl focus:outline-none focus:ring-2 focus:ring-blue-500 text-sm"
                    disabled={isReadOnly}
                    placeholder="Masukkan No Unit"
                  />
                </div>
                <div>
                  <label className="block text-xs font-medium text-gray-500 uppercase tracking-wider mb-1">Kode Proyek</label>
                  <input
                    type="text"
                    value={formData.projectCode || ''}
                    onChange={(e) => setFormData((prev: any) => ({ ...prev, projectCode: e.target.value }))}
                    className="w-full px-3 py-2 bg-white border border-gray-200 rounded-xl focus:outline-none focus:ring-2 focus:ring-blue-500 text-sm"
                    disabled={isReadOnly}
                    placeholder="Masukkan Kode Proyek"
                  />
                </div>
                <div>
                  <label className="block text-xs font-medium text-gray-500 uppercase tracking-wider mb-1">Nama Pelanggan</label>
                  <input
                    type="text"
                    value={formData.customerName || ''}
                    onChange={(e) => setFormData((prev: any) => ({ ...prev, customerName: e.target.value }))}
                    className="w-full px-3 py-2 bg-white border border-gray-200 rounded-xl focus:outline-none focus:ring-2 focus:ring-blue-500 text-sm"
                    disabled={isReadOnly}
                    placeholder="Masukkan Nama Pelanggan"
                  />
                </div>
                <div>
                  <label className="block text-xs font-medium text-gray-500 uppercase tracking-wider mb-1">Lokasi Gedung</label>
                  <input
                    type="text"
                    value={formData.buildingLocation || ''}
                    onChange={(e) => setFormData((prev: any) => ({ ...prev, buildingLocation: e.target.value }))}
                    className="w-full px-3 py-2 bg-white border border-gray-200 rounded-xl focus:outline-none focus:ring-2 focus:ring-blue-500 text-sm"
                    disabled={isReadOnly}
                    placeholder="Masukkan Lokasi Gedung"
                  />
                </div>
                <div>
                  <label className="block text-xs font-medium text-gray-500 uppercase tracking-wider mb-1">Tipe / Merk Elevator</label>
                  <input
                    type="text"
                    value={formData.elevatorType || ''}
                    onChange={(e) => setFormData((prev: any) => ({ ...prev, elevatorType: e.target.value }))}
                    className="w-full px-3 py-2 bg-white border border-gray-200 rounded-xl focus:outline-none focus:ring-2 focus:ring-blue-500 text-sm"
                    disabled={isReadOnly}
                    placeholder="Masukkan Tipe / Merk"
                  />
                </div>
                <div>
                  <label className="block text-xs font-medium text-gray-500 uppercase tracking-wider mb-1">Kapasitas Angkut (kg)</label>
                  <input
                    type="text"
                    value={formData.capacity || ''}
                    onChange={(e) => setFormData((prev: any) => ({ ...prev, capacity: e.target.value }))}
                    className="w-full px-3 py-2 bg-white border border-gray-200 rounded-xl focus:outline-none focus:ring-2 focus:ring-blue-500 text-sm"
                    disabled={isReadOnly}
                    placeholder="Masukkan Kapasitas"
                  />
                </div>
                <div>
                  <label className="block text-xs font-medium text-gray-500 uppercase tracking-wider mb-1">Kecepatan (m/menit)</label>
                  <input
                    type="text"
                    value={formData.speed || ''}
                    onChange={(e) => setFormData((prev: any) => ({ ...prev, speed: e.target.value }))}
                    className="w-full px-3 py-2 bg-white border border-gray-200 rounded-xl focus:outline-none focus:ring-2 focus:ring-blue-500 text-sm"
                    disabled={isReadOnly}
                    placeholder="Masukkan Kecepatan"
                  />
                </div>
                <div>
                  <label className="block text-xs font-medium text-gray-500 uppercase tracking-wider mb-1">Tanggal Pemeriksaan</label>
                  <input
                    type="date"
                    value={formData.inspectionDate || ''}
                    onChange={(e) => setFormData((prev: any) => ({ ...prev, inspectionDate: e.target.value }))}
                    className="w-full px-3 py-2 bg-white border border-gray-200 rounded-xl focus:outline-none focus:ring-2 focus:ring-blue-500 text-sm"
                    disabled={isReadOnly}
                  />
                </div>
                <div>
                  <label className="block text-xs font-medium text-gray-500 uppercase tracking-wider mb-1">Nama QC</label>
                  <input
                    type="text"
                    value={formData.qcName || ''}
                    onChange={(e) => setFormData((prev: any) => ({ ...prev, qcName: e.target.value }))}
                    className="w-full px-3 py-2 bg-white border border-gray-200 rounded-xl focus:outline-none focus:ring-2 focus:ring-blue-500 text-sm"
                    disabled={isReadOnly}
                    placeholder="Masukkan Nama QC"
                  />
                </div>
              </div>
            </section>

            <div className="flex flex-wrap items-center gap-3 bg-gray-50/50 p-3 rounded-xl mb-6">
              <span className="text-xs font-medium text-gray-500">Status:</span>
              <span className={`px-3 py-1 rounded-full text-xs font-medium ${
                formData.status === 'draft' ? 'bg-gray-200 text-gray-600' :
                formData.status === 'qc_approved' ? 'bg-blue-100 text-blue-700' :
                formData.status === 'revision' ? 'bg-amber-100 text-amber-700' :
                formData.status === 'maintenance_done' ? 'bg-orange-100 text-orange-700' :
                'bg-emerald-100 text-emerald-700'
              }`}>
                {formData.status === 'draft' ? 'Draft' :
                 formData.status === 'qc_approved' ? 'Maintenance' :
                 formData.status === 'revision' ? 'REVISION' :
                 formData.status === 'maintenance_done' ? 'Selesai' :
                 'FINAL'}
              </span>
              {formData.submittedAt && <span className="text-xs text-gray-400">Dikirim: {formData.submittedAt}</span>}
              {formData.qcVerification?.verifiedAt && <span className="text-xs text-gray-400">Diverifikasi: {formData.qcVerification.verifiedAt}</span>}
            </div>

            <section className="bg-white rounded-xl border border-gray-100 overflow-hidden mb-6">
              <div className="bg-blue-700 px-6 py-3 flex justify-between items-center">
                <h2 className="text-sm font-semibold text-white">Daftar Simak Pemeriksaan (116 Item)</h2>
                <span className="text-xs text-white bg-blue-600 px-3 py-1 rounded-full">{formData.inspectionData.length} Item</span>
              </div>
              {/* ✅ PERBAIKAN: Tabel dipadatkan, border dalam dihapus, muat 1 halaman */}
              <div className="overflow-x-auto p-0">
                <table className="w-full text-xs border-collapse">
                  <thead className="bg-gray-50">
                    <tr>
                      <th className="p-1.5 text-center text-gray-500 font-medium border-b border-gray-200 w-6">No</th>
                      <th className="p-1.5 text-left text-gray-500 font-medium border-b border-gray-200 w-20">Kategori</th>
                      <th className="p-1.5 text-left text-gray-500 font-medium border-b border-gray-200 min-w-[120px] max-w-[200px]">Komponen</th>
                      <th className="p-1.5 text-left text-gray-500 font-medium border-b border-gray-200 w-24">Prioritas</th>
                      <th className="p-1.5 text-center text-gray-500 font-medium border-b border-gray-200 w-12">Bobot</th>
                      <th className="p-1.5 text-left text-gray-500 font-medium border-b border-gray-200 w-20">Status</th>
                      <th className="p-1.5 text-left text-gray-500 font-medium border-b border-gray-200 min-w-[100px] max-w-[150px]">Temuan</th>
                      <th className="p-1.5 text-center text-gray-500 font-medium border-b border-gray-200 w-24">Foto Sblm</th>
                      <th className="p-1.5 text-center text-gray-500 font-medium border-b border-gray-200 w-24">Foto Sdh</th>
                      <th className="p-1.5 text-left text-gray-500 font-medium border-b border-gray-200 min-w-[100px] max-w-[150px]">Catatan</th>
                      <th className="p-1.5 text-center text-gray-500 font-medium border-b border-gray-200 w-14">Approve</th>
                      <th className="p-1.5 text-center text-gray-500 font-medium border-b border-gray-200 w-14">Hitung</th>
                    </tr>
                  </thead>
                  <tbody>
                    {formData.inspectionData.map((item: any) => {
                      const originalItem = inspectionItems.find(i => i.id === item.id);
                      const isNotGood = item.status === 'Not Good';
                      const isApproved = item.isApproved;
                      const isEditable = isQC && !isReadOnly && formData.status === 'draft';
                      const isMaintEditable = isMaintenance && !isReadOnly && (formData.status === 'qc_approved' || formData.status === 'revision');
                      const canApprove = isQC && !isReadOnly && (formData.status === 'maintenance_done' || formData.status === 'revision') && isNotGood && !isApproved;
                      const hitung = calculateHitung(item, originalItem);
                      const hitungPersen = hitung * 100;

                      return (
                        <tr key={item.id} className="border-b border-gray-100 hover:bg-gray-50/50 transition">
                          <td className="p-1.5 text-center text-gray-400 align-top">{item.id}</td>
                          <td className="p-1.5 text-gray-500 uppercase align-top">{originalItem?.category || '-'}</td>
                          <td className="p-1.5 text-gray-700 align-top break-words whitespace-normal leading-tight">{originalItem?.item || '-'}</td>
                          <td className="p-1.5 text-gray-600 align-top">{originalItem?.priority || '-'}</td>
                          <td className="p-1.5 text-center text-gray-400 align-top">{(originalItem?.weight || 0) * 100}%</td>
                          <td className="p-1.5 align-top">
                            {isEditable ? (
                              <div className="flex flex-col gap-0.5">
                                <label className="flex items-center gap-1 text-[10px]"><input type="radio" name={`status-${item.id}`} value="Good" checked={item.status === 'Good'} onChange={() => handleStatusChange(item.id, 'Good')} className="w-3 h-3"/> Good</label>
                                <label className="flex items-center gap-1 text-[10px]"><input type="radio" name={`status-${item.id}`} value="Not Good" checked={item.status === 'Not Good'} onChange={() => handleStatusChange(item.id, 'Not Good')} className="w-3 h-3"/> Not Good</label>
                                <label className="flex items-center gap-1 text-[10px]"><input type="radio" name={`status-${item.id}`} value="N/A" checked={item.status === 'N/A'} onChange={() => handleStatusChange(item.id, 'N/A')} className="w-3 h-3"/> N/A</label>
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
                          <td className="p-1.5 align-top">
                            {isEditable ? (
                              <textarea value={item.finding} onChange={(e) => handleFindingChange(item.id, e.target.value)} disabled={item.status === 'Good' || item.status === 'N/A'} placeholder="Temuan" className={`w-full px-1 py-0.5 border rounded text-[10px] ${isNotGood && !item.finding ? 'border-rose-300 bg-rose-50' : 'border-gray-200'} ${item.status === 'Good' ? 'bg-gray-50' : 'bg-white'}`} rows={2}/>
                            ) : <span className="text-gray-600 break-words whitespace-normal">{item.finding || '-'}</span>}
                          </td>
                          <td className="p-1.5 text-center align-top">
                            {isEditable && isNotGood ? (
                              <input type="file" accept="image/*" onChange={(e) => handlePhotoBeforeChange(item.id, e)} className="text-[10px] w-full"/>
                            ) : item.photoBefore ? <img src={item.photoBefore} alt="before" className="w-16 h-16 object-cover rounded border mx-auto"/> : <span className="text-gray-300">-</span>}
                          </td>
                          <td className="p-1.5 text-center align-top">
                            {isMaintEditable ? (
                              <div>
                                <input type="file" accept="image/*" onChange={(e) => handlePhotoAfterChange(item.id, e)} className="text-[10px] w-full"/>
                                {isNotGood && !item.photoAfter && <p className="text-rose-500 text-[8px]">Wajib</p>}
                              </div>
                            ) : item.photoAfter ? <img src={item.photoAfter} alt="after" className="w-16 h-16 object-cover rounded border mx-auto"/> : <span className="text-gray-300">-</span>}
                          </td>
                          <td className="p-1.5 align-top">
                            {isMaintEditable ? (
                              <textarea value={item.repairNote || ''} onChange={(e) => handleRepairNoteChange(item.id, e.target.value)} placeholder="Catatan" className={`w-full px-1 py-0.5 border rounded text-[10px] ${isNotGood && !item.repairNote ? 'border-rose-300 bg-rose-50' : 'border-gray-200'}`} rows={2}/>
                            ) : <span className="text-gray-600 break-words whitespace-normal">{item.repairNote || '-'}</span>}
                          </td>
                          <td className="p-1.5 text-center align-top">
                            {canApprove ? (
                              <button type="button" onClick={() => handleApproveItem(item.id)} className="bg-emerald-500 hover:bg-emerald-600 text-white px-2 py-0.5 rounded text-[10px] transition">Approve</button>
                            ) : isApproved ? <span className="text-emerald-500 text-xs">Approved</span> : item.status === 'Not Good' ? <span className="text-gray-300 text-xs">Menunggu</span> : <span className="text-gray-200 text-xs">-</span>}
                          </td>
                          <td className="p-1.5 text-center align-top">
                            {/* ✅ PERBAIKAN: .toFixed(2) memaksa 0.90% */}
                            <span className={`font-bold ${hitungPersen >= 0 ? 'text-emerald-600' : 'text-rose-600'}`}>{hitungPersen.toFixed(2)}%</span>
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            </section>

            <section className="bg-white rounded-xl border border-gray-100 p-6 mb-6">
              <h2 className="text-sm font-semibold text-gray-600 mb-4">Rekap</h2>
              
              {isQC && !isReadOnly && (
                <div className="mb-4">
                  <label className="block text-xs font-medium text-gray-500 uppercase tracking-wider mb-1">Catatan QC</label>
                  <textarea
                    value={formData.qcNote || ''}
                    onChange={(e) => setFormData((prev: any) => ({ ...prev, qcNote: e.target.value }))}
                    placeholder="Tambahkan catatan di sini..."
                    className="w-full px-3 py-2 bg-white border border-gray-200 rounded-xl focus:outline-none focus:ring-2 focus:ring-blue-500 text-sm"
                    rows={3}
                  />
                </div>
              )}
              {formData.qcNote && isReadOnly && (
                <div className="mb-4 p-3 bg-blue-50 border border-blue-200 rounded-xl text-gray-700 text-sm">
                  <span className="font-medium text-blue-700">Catatan QC:</span> {formData.qcNote}
                </div>
              )}

              <div className="grid grid-cols-2 md:grid-cols-4 gap-4">
                <div className="bg-gray-50 p-4 rounded-xl text-center"><div className="text-xs text-gray-400">Total</div><div className="text-xl font-bold text-gray-700">{formData.inspectionData.length}</div></div>
                <div className="bg-emerald-50 p-4 rounded-xl text-center"><div className="text-xs text-emerald-500">Good + N/A</div><div className="text-xl font-bold text-emerald-700">{formData.inspectionData.filter((i: any) => i.status === 'Good' || i.status === 'N/A' || (i.status === 'Not Good' && i.isApproved)).length}</div></div>
                <div className="bg-rose-50 p-4 rounded-xl text-center"><div className="text-xs text-rose-500">Not Good</div><div className="text-xl font-bold text-rose-700">{formData.inspectionData.filter((i: any) => i.status === 'Not Good' && !i.isApproved).length}</div></div>
                <div className="bg-blue-50 p-4 rounded-xl text-center"><div className="text-xs text-blue-500">Nilai</div><div className="text-xl font-bold text-blue-700">{totalScore.toFixed(2)}%</div></div>
              </div>
            </section>

            <section className="bg-white rounded-xl border border-gray-100 p-6 mb-6">
              <h2 className="text-sm font-semibold text-gray-600 mb-4">Attachment</h2>
              <div className="border-2 border-dashed border-gray-200 rounded-xl p-6 text-center hover:border-blue-400 transition">
                {isQC && !isReadOnly ? (
                  <input type="file" accept=".pdf,.doc,.docx,.xlsx,.xls,.jpg,.jpeg,.png" onChange={handleAttachmentChange} disabled={isReadOnly} className="text-sm text-gray-500 file:mr-3 file:py-2 file:px-4 file:rounded-xl file:border-0 file:text-sm file:font-medium file:bg-blue-50 file:text-blue-700 hover:file:bg-blue-100"/>
                ) : formData.attachment ? (
                  <div className="flex items-center justify-center gap-4 flex-wrap">
                    <span className="text-sm text-gray-600">📄 {formData.attachmentName}</span>
                    <button type="button" onClick={handleDownloadAttachment} className="bg-blue-600 hover:bg-blue-700 text-white px-4 py-2 rounded-xl text-sm transition">Download</button>
                    <a href={formData.attachment} target="_blank" rel="noopener" className="bg-emerald-600 hover:bg-emerald-700 text-white px-4 py-2 rounded-xl text-sm transition">Lihat</a>
                  </div>
                ) : <p className="text-sm text-gray-400">Tidak ada file</p>}
                <p className="text-[10px] text-gray-400 mt-2">PDF, DOC, XLSX, Gambar</p>
              </div>
            </section>
          </>
        )}

        {isPemeliharaan && formData.units?.map((unit: UnitData, unitIndex: number) => {
          const totalItems = unit.maintenanceData?.length || 0;
          const goodCount = unit.maintenanceData?.filter((item: any) => item.status === 'Good' || item.status === 'N/A' || (item.status === 'Not Good' && item.isApproved)).length || 0;
          const score = totalItems > 0 ? Math.round((goodCount / totalItems) * 100) : 0;

          return (
            <div key={unitIndex} className="bg-white rounded-xl border border-gray-100 overflow-hidden mb-6">
              <div className="bg-emerald-700 px-6 py-3 flex justify-between items-center">
                <h2 className="text-sm font-semibold text-white">Unit #{unitIndex + 1}</h2>
                <span className="text-xs text-white bg-emerald-600 px-3 py-1 rounded-full">{totalItems} Item</span>
              </div>
              <div className="p-4 border-b">
                <h3 className="text-xs font-semibold text-gray-500 uppercase tracking-wider mb-3">Data Umum</h3>
                <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
                  {['unitNumber','projectCode','customerName','buildingLocation','elevatorType','capacity','speed','inspectionDate','qcName'].map((field) => (
                    <div key={field}>
                      <label className="block text-[10px] font-medium text-gray-400 uppercase tracking-wider">
                        {field === 'unitNumber' ? 'No Unit' : field === 'projectCode' ? 'Kode Proyek' : field === 'customerName' ? 'Pelanggan' : field === 'buildingLocation' ? 'Lokasi' : field === 'elevatorType' ? 'Tipe/Merk' : field === 'capacity' ? 'Kapasitas' : field === 'speed' ? 'Kecepatan' : field === 'inspectionDate' ? 'Tanggal' : 'Nama QC'}
                      </label>
                      <input 
                        type={field === 'inspectionDate' ? 'date' : 'text'} 
                        value={String(unit[field as keyof UnitData] || '')} 
                        onChange={(e) => updateUnitData(unitIndex, field, e.target.value)} 
                        className="w-full px-3 py-2 bg-white border border-gray-200 rounded-xl focus:outline-none focus:ring-2 focus:ring-emerald-500 text-sm" 
                        disabled={isReadOnly || !isQC || formData.status !== 'draft'} 
                        placeholder="-"
                      />
                    </div>
                  ))}
                </div>
              </div>
              <div className="p-4 overflow-x-auto">
                <h3 className="text-xs font-semibold text-gray-500 uppercase tracking-wider mb-3">Data Pemeliharaan</h3>
                <table className="w-full text-xs" style={{ minWidth: '800px' }}>
                  <thead className="bg-gray-50">
                    <tr>
                      <th className="p-2 text-left text-gray-500 font-medium w-8">No</th>
                      <th className="p-2 text-left text-gray-500 font-medium w-32">Item</th>
                      <th className="p-2 text-left text-gray-500 font-medium w-20">Status</th>
                      <th className="p-2 text-left text-gray-500 font-medium w-32">Temuan</th>
                      <th className="p-2 text-left text-gray-500 font-medium w-16">Foto Sebelum</th>
                      <th className="p-2 text-left text-gray-500 font-medium w-16">Foto Setelah</th>
                      <th className="p-2 text-left text-gray-500 font-medium w-32">Catatan</th>
                      <th className="p-2 text-left text-gray-500 font-medium w-12">Approve</th>
                    </tr>
                  </thead>
                  <tbody>
                    {unit.maintenanceData.map((item: any) => {
                      const originalItem = maintenanceItems.find(i => i.id === item.id);
                      const isNotGood = item.status === 'Not Good';
                      const isApproved = item.isApproved;
                      const isEditable = isQC && !isReadOnly && formData.status === 'draft';
                      const isMaintEditable = isMaintenance && !isReadOnly && (formData.status === 'qc_approved' || formData.status === 'revision');
                      const canApprove = isQC && !isReadOnly && (formData.status === 'maintenance_done' || formData.status === 'revision') && isNotGood && !isApproved;

                      return (
                        <tr key={item.id} className="border-t border-gray-50 hover:bg-gray-50/50 transition">
                          <td className="p-2 text-gray-400 align-top text-center">{item.id}</td>
                          <td className="p-2 text-gray-700 align-top break-words whitespace-normal leading-tight">{originalItem?.item || '-'}</td>
                          <td className="p-2 align-top">
                            {isEditable ? (
                              <div className="flex flex-col gap-0.5">
                                <label className="flex items-center gap-1 text-[10px]"><input type="radio" name={`unit-${unitIndex}-status-${item.id}`} value="Good" checked={item.status === 'Good'} onChange={() => handleMaintenanceStatusChange(unitIndex, item.id, 'Good')} className="w-3 h-3"/> Good</label>
                                <label className="flex items-center gap-1 text-[10px]"><input type="radio" name={`unit-${unitIndex}-status-${item.id}`} value="Not Good" checked={item.status === 'Not Good'} onChange={() => handleMaintenanceStatusChange(unitIndex, item.id, 'Not Good')} className="w-3 h-3"/> Not Good</label>
                                <label className="flex items-center gap-1 text-[10px]"><input type="radio" name={`unit-${unitIndex}-status-${item.id}`} value="N/A" checked={item.status === 'N/A'} onChange={() => handleMaintenanceStatusChange(unitIndex, item.id, 'N/A')} className="w-3 h-3"/> N/A</label>
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
                              <textarea value={item.finding} onChange={(e) => handleMaintenanceFindingChange(unitIndex, item.id, e.target.value)} disabled={item.status === 'Good' || item.status === 'N/A'} placeholder="Temuan" className={`w-full min-w-[100px] px-2 py-1 border rounded-lg text-xs ${isNotGood && !item.finding ? 'border-rose-300 bg-rose-50' : 'border-gray-200'}`} rows={2}/>
                            ) : <span className="text-gray-600 text-xs break-words whitespace-normal">{item.finding || '-'}</span>}
                          </td>
                          <td className="p-2 align-top">
                            {isEditable && isNotGood ? (
                              <input type="file" accept="image/*" onChange={(e) => handleMaintenancePhotoBeforeChange(unitIndex, item.id, e)} className="text-[10px]"/>
                            ) : item.photoBefore ? <img src={item.photoBefore} alt="before" className="w-12 h-12 object-cover rounded-lg border"/> : <span className="text-gray-300 text-[10px]">-</span>}
                          </td>
                          <td className="p-2 align-top">
                            {isMaintEditable ? (
                              <div>
                                <input type="file" accept="image/*" onChange={(e) => handleMaintenancePhotoAfterChange(unitIndex, item.id, e)} className="text-[10px]"/>
                                {isNotGood && !item.photoAfter && <p className="text-rose-500 text-[8px]">Wajib</p>}
                              </div>
                            ) : item.photoAfter ? <img src={item.photoAfter} alt="after" className="w-12 h-12 object-cover rounded-lg border"/> : <span className="text-gray-300 text-[10px]">-</span>}
                          </td>
                          <td className="p-2 align-top">
                            {isMaintEditable ? (
                              <textarea value={item.repairNote || ''} onChange={(e) => handleMaintenanceRepairNoteChange(unitIndex, item.id, e.target.value)} placeholder="Catatan" className={`w-full min-w-[100px] px-2 py-1 border rounded-lg text-xs ${isNotGood && !item.repairNote ? 'border-rose-300 bg-rose-50' : 'border-gray-200'}`} rows={2}/>
                            ) : <span className="text-gray-600 text-xs break-words whitespace-normal">{item.repairNote || '-'}</span>}
                          </td>
                          <td className="p-2 align-top text-center">
                            {canApprove ? (
                              <button type="button" onClick={() => handleApproveMaintenanceItem(unitIndex, item.id)} className="bg-emerald-500 hover:bg-emerald-600 text-white px-2 py-0.5 rounded-lg text-[10px] transition">Approve</button>
                            ) : isApproved ? <span className="text-emerald-500 text-xs">Approved</span> : item.status === 'Not Good' ? <span className="text-gray-300 text-xs">Menunggu</span> : <span className="text-gray-200 text-xs">-</span>}
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
              <div className="p-4 bg-gray-50 border-t">
                <div className="grid grid-cols-2 md:grid-cols-4 gap-3">
                  <div className="bg-white p-3 rounded-xl text-center"><div className="text-[10px] text-gray-400">Total</div><div className="font-bold text-gray-700">{totalItems}</div></div>
                  <div className="bg-emerald-50 p-3 rounded-xl text-center"><div className="text-[10px] text-emerald-500">Good + N/A</div><div className="font-bold text-emerald-700">{goodCount}</div></div>
                  <div className="bg-rose-50 p-3 rounded-xl text-center"><div className="text-[10px] text-rose-500">Not Good</div><div className="font-bold text-rose-700">{totalItems - goodCount}</div></div>
                  <div className="bg-blue-50 p-3 rounded-xl text-center"><div className="text-[10px] text-blue-500">Score</div><div className="font-bold text-blue-700">{score}%</div></div>
                </div>
              </div>
            </div>
          );
        })}

        <div className="text-center text-[10px] text-gray-300 border-t pt-4 mt-4 print:block hidden">
          <p>Dicetak: {new Date().toLocaleString('id-ID')}</p>
          <p>© 2026 PT Louserindo Megah Permai</p>
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