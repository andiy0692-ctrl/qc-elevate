"use client";

import { useState, useEffect } from 'react';
import * as XLSX from 'xlsx';

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
type MenuType = 'dashboard' | 'newPemeriksaan' | 'newPemeliharaan' | 'viewReport';
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
// 🔥 FUNGSI BACKUP - DOWNLOAD JSON (PASTI BERHASIL)
// ============================================
async function backupToGoogleDrive(data: any[]) {
  try {
    localStorage.setItem('elevateQC_reports', JSON.stringify(data));
    
    const jsonString = JSON.stringify(data, null, 2);
    const blob = new Blob([jsonString], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    const fileName = `elevateQC_backup_${new Date().toISOString().replace(/[:.]/g, '-')}.json`;
    link.download = fileName;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
    
    alert(`✅ Backup berhasil! File: ${fileName}`);
    return { success: true, fileName: fileName };
  } catch (error) {
    alert('❌ Backup gagal!');
    return { success: false, error: error instanceof Error ? error.message : 'Unknown error' };
  }
}

// ============================================
// 🔥 FUNGSI RESTORE - UPLOAD JSON (PASTI BERHASIL)
// ============================================
async function restoreFromGoogleDrive() {
  try {
    return new Promise((resolve) => {
      const input = document.createElement('input');
      input.type = 'file';
      input.accept = '.json';
      input.onchange = (e: any) => {
        const file = e.target.files[0];
        if (!file) {
          resolve(null);
          return;
        }
        const reader = new FileReader();
        reader.onload = (event: any) => {
          try {
            const data = JSON.parse(event.target.result);
            if (!Array.isArray(data)) {
              alert('❌ File tidak valid!');
              resolve(null);
              return;
            }
            localStorage.setItem('elevateQC_reports', JSON.stringify(data));
            alert(`✅ Restore berhasil! ${data.length} laporan dimuat.`);
            resolve(data);
          } catch (err) {
            alert('❌ File tidak valid!');
            resolve(null);
          }
        };
        reader.readAsText(file);
      };
      input.click();
    });
  } catch (error) {
    alert('❌ Restore gagal!');
    return null;
  }
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

    const users = {
      'qc': { password: 'qc123', role: 'qc' },
      'mainten': { password: 'm123', role: 'maintenance' },
      'sales': { password: 's123', role: 'sales' }
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
    <div className="min-h-screen bg-gradient-to-br from-blue-50 to-gray-100 flex items-center justify-center p-4">
      <div className="bg-white rounded-2xl shadow-xl max-w-md w-full p-8">
        <div className="text-center mb-8">
          <div className="inline-block bg-blue-700 rounded-full p-4 mb-4">
            <svg className="w-10 h-10 text-white" fill="none" stroke="currentColor" viewBox="0 0 24 24">
              <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M19 21V5a2 2 0 00-2-2H7a2 2 0 00-2 2v16m14 0h2m-2 0h-5m-9 0H3m2 0h5M9 7h1m-1 4h1m4-4h1m-1 4h1m-5 10v-5a1 1 0 011-1h2a1 1 0 011 1v5m-4 0h4" />
            </svg>
          </div>
          <h1 className="text-2xl font-bold text-gray-800">LOUSERINDO MEGAH PERMAI</h1>
          <p className="text-sm text-gray-500">Elevator Quality Control System</p>
        </div>

        <form onSubmit={handleLogin} className="space-y-4">
          <div className="grid grid-cols-3 gap-2">
            {(['maintenance', 'qc', 'sales'] as const).map((role) => (
              <button
                key={role}
                type="button"
                onClick={() => setSelectedRole(role)}
                className={`p-2 rounded-lg border-2 text-center transition text-sm ${
                  selectedRole === role
                    ? role === 'maintenance' ? 'border-blue-600 bg-blue-50 text-blue-700' :
                      role === 'qc' ? 'border-green-600 bg-green-50 text-green-700' :
                      'border-orange-600 bg-orange-50 text-orange-700'
                    : 'border-gray-200 hover:border-blue-300'
                }`}
              >
                <div className="text-xl">
                  {role === 'maintenance' ? '🔧' : role === 'qc' ? '✅' : '💰'}
                </div>
                <div className="text-[10px] font-semibold">
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
            className="w-full px-4 py-2 border rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500"
          />

          <input
            type="password"
            value={password}
            onChange={(e) => setPassword(e.target.value)}
            placeholder="Password"
            className="w-full px-4 py-2 border rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500"
          />

          {error && <div className="text-red-500 text-sm">{error}</div>}

          <div className="bg-gray-50 rounded-lg p-3 text-xs text-gray-500 text-center border border-gray-200">
            <p className="font-semibold text-blue-700">PT Louserindo Megah Permai</p>
            <p className="text-gray-400">Elevator Quality Control System v1.0</p>
            <p className="text-[10px] text-gray-400 mt-1">© 2026 - All Rights Reserved</p>
          </div>

          <button
            type="submit"
            disabled={loading}
            className={`w-full font-bold py-3 rounded-lg transition ${
              loading 
                ? 'bg-gray-400 text-white cursor-not-allowed' 
                : 'bg-blue-700 hover:bg-blue-800 text-white'
            }`}
          >
            {loading ? '⏳ Memuat...' : 'Masuk'}
          </button>
        </form>
      </div>
    </div>
  );
}

// ============================================
// 🔥 KOMPONEN DASHBOARD - REVISI FINAL
// ============================================
function Dashboard({ role, onLogout }: { 
  role: UserRole; 
  onLogout: () => void;
}) {
  const [currentMenu, setCurrentMenu] = useState<MenuType>('dashboard');
  const [selectedReport, setSelectedReport] = useState<Report | null>(null);
  const [reports, setReports] = useState<Report[]>([]);
  const [isReadOnly, setIsReadOnly] = useState(false);
  const [refreshKey, setRefreshKey] = useState(0);
  const [sortOrder, setSortOrder] = useState<SortOrder>('terbaru');
  const [loading, setLoading] = useState(false);

  const isQC = role === 'qc';
  const isMaintenance = role === 'maintenance';
  const isSales = role === 'sales';

  const loadData = async () => {
    setLoading(true);
    try {
      // Cek localStorage dulu
      const savedData = localStorage.getItem('elevateQC_reports');
      if (savedData) {
        try {
          const parsed = JSON.parse(savedData);
          if (parsed && parsed.length > 0) {
            setReports(parsed);
            console.log('📦 Loaded from localStorage:', parsed.length);
            setLoading(false);
            return;
          }
        } catch (err) {}
      }
      
      // Kalau kosong, coba restore dari file
      const driveData = await restoreFromGoogleDrive();
      if (driveData && driveData.length > 0) {
        setReports(driveData);
        localStorage.setItem('elevateQC_reports', JSON.stringify(driveData));
        console.log('📦 Restore dari file:', driveData.length);
        setLoading(false);
        return;
      }
      
      setReports([]);
    } catch (e) {
      console.error('Error loading reports:', e);
      setReports([]);
    }
    setLoading(false);
  };

  const saveReports = async (newReports: Report[]) => {
    localStorage.setItem('elevateQC_reports', JSON.stringify(newReports));
    setReports(newReports);
    
    try {
      await backupToGoogleDrive(newReports);
      console.log('💾 Backup success');
    } catch (e) {
      console.warn('⚠️ Backup error:', e);
    }
  };

  const handleRefresh = async () => {
    const result = await restoreFromGoogleDrive();
    if (result && result.length > 0) {
      setReports(result);
      localStorage.setItem('elevateQC_reports', JSON.stringify(result));
      alert(`✅ Data berhasil di-restore! ${result.length} laporan dimuat.`);
    } else {
      alert('📂 Silakan pilih file JSON untuk di-restore');
    }
    setRefreshKey(prev => prev + 1);
  };

  const handleManualBackup = async () => {
    if (reports.length === 0) {
      alert('⚠️ Tidak ada data untuk di-backup!');
      return;
    }
    const result = await backupToGoogleDrive(reports);
    if (result?.success) {
      alert(`✅ Backup berhasil!\nFile: ${result.fileName}`);
    } else {
      alert(`❌ Backup gagal: ${result?.error || 'Unknown error'}`);
    }
  };

  const handleNewPemeriksaan = () => {
    const newReport: Report = {
      id: Date.now().toString(),
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
    setSelectedReport(newReport);
    setIsReadOnly(false);
    setCurrentMenu('newPemeriksaan');
  };

  const handleNewPemeliharaan = () => {
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

    const newReport: Report = {
      id: Date.now().toString(),
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
    setSelectedReport(newReport);
    setIsReadOnly(false);
    setCurrentMenu('newPemeliharaan');
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
  };

  const handleSaveReport = async (data: any) => {
    const existingIndex = reports.findIndex(r => r.id === data.id);
    let newReports: Report[];
    if (existingIndex >= 0) {
      newReports = [...reports];
      newReports[existingIndex] = data;
    } else {
      newReports = [...reports, data];
    }
    
    await saveReports(newReports);
    setSelectedReport(data);
    setCurrentMenu('dashboard');
  };

  const handleDeleteReport = async (id: string) => {
    if (!isQC) {
      alert('⚠️ Hanya QC yang bisa menghapus laporan!');
      return;
    }
    if (confirm('Yakin ingin menghapus laporan ini?')) {
      const newReports = reports.filter(r => r.id !== id);
      localStorage.setItem('elevateQC_reports', JSON.stringify(newReports));
      setReports(newReports);
      await backupToGoogleDrive(newReports);
      setCurrentMenu('dashboard');
      alert('🗑️ Laporan berhasil dihapus!');
    }
  };

  const handleCancel = () => {
    setSelectedReport(null);
    setIsReadOnly(false);
    setCurrentMenu('dashboard');
  };

  const handlePrintPDF = () => {
    window.print();
  };

  const handleExportExcel = () => {
    if (!selectedReport) return;
    alert('📊 Export Excel (fitur dalam pengembangan)');
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
  const totalPemeriksaan = isQC ? reports.filter(r => r.reportType === 'pemeriksaan').length : 0;
  const totalPemeliharaan = isQC ? reports.filter(r => r.reportType === 'pemeliharaan').length : 0;

  useEffect(() => {
    loadData();
  }, []);

  return (
    <div className="min-h-screen bg-gray-50">
      <div className="container mx-auto px-4 py-6 max-w-7xl">
        <header className="no-print bg-white shadow-sm rounded-lg p-6 mb-8 border border-gray-200">
          <div className="flex flex-col md:flex-row md:items-center md:justify-between">
            <div>
              <h1 className="text-3xl font-bold text-blue-700">LOUSERINDO MEGAH PERMAI</h1>
              <p className="text-sm text-gray-600 mt-1">Elevator Quality Control System</p>
              {isMaintenance && (
                <p className="text-sm text-blue-600 mt-1">🔧 Login sebagai: Kepala Maintenance ({totalTugas} tugas aktif)</p>
              )}
              {isSales && (
                <p className="text-sm text-orange-600 mt-1">👀 Sales - Hanya melihat Data Pemeriksaan</p>
              )}
            </div>
            <div className="mt-4 md:mt-0 flex flex-wrap items-center gap-4">
              <span className={`inline-flex items-center px-3 py-1 rounded-full text-sm font-medium ${
                isQC ? 'bg-green-100 text-green-800' : isMaintenance ? 'bg-blue-100 text-blue-800' : 'bg-orange-100 text-orange-800'
              }`}>
                {isQC ? '✅ QC' : isMaintenance ? '🔧 Maintenance' : '💰 Sales'}
              </span>
              {isQC && stats.maintenanceDone > 0 && <span className="inline-flex items-center px-3 py-1 rounded-full text-sm font-medium bg-yellow-100 text-yellow-800 animate-pulse">📬 {stats.maintenanceDone} menunggu verifikasi</span>}
              {isQC && stats.revision > 0 && <span className="inline-flex items-center px-3 py-1 rounded-full text-sm font-medium bg-red-100 text-red-800">🔄 {stats.revision} revisi</span>}
              {isMaintenance && (stats.qcApproved + stats.revision > 0) && <span className="inline-flex items-center px-3 py-1 rounded-full text-sm font-medium bg-blue-100 text-blue-800 animate-pulse">🔧 {stats.qcApproved + stats.revision} perlu dikerjakan</span>}
              <button onClick={handleRefresh} className="text-sm bg-blue-100 hover:bg-blue-200 text-blue-700 px-3 py-1 rounded-md transition flex items-center gap-1">🔄 Upload & Restore</button>
              <button onClick={handleManualBackup} className="text-sm bg-green-100 hover:bg-green-200 text-green-700 px-3 py-1 rounded-md transition flex items-center gap-1">💾 Backup Drive</button>
              <button onClick={onLogout} className="text-sm text-gray-500 hover:text-gray-700 hover:bg-gray-100 px-3 py-1 rounded-md transition">Logout</button>
            </div>
          </div>

          <div className="mt-3 grid grid-cols-4 md:grid-cols-6 gap-2 text-xs">
            <div className="bg-gray-50 p-2 rounded text-center"><div className="font-bold text-gray-700">{stats.total}</div><div className="text-gray-400">Total</div></div>
            <div className="bg-gray-50 p-2 rounded text-center"><div className="font-bold text-gray-500">{stats.draft}</div><div className="text-gray-400">Draft QC</div></div>
            <div className="bg-blue-50 p-2 rounded text-center"><div className="font-bold text-blue-600">{stats.qcApproved}</div><div className="text-blue-400">Maintenance</div></div>
            <div className="bg-yellow-50 p-2 rounded text-center"><div className="font-bold text-yellow-600">{stats.revision}</div><div className="text-yellow-400">Revisi</div></div>
            <div className="bg-orange-50 p-2 rounded text-center"><div className="font-bold text-orange-600">{stats.maintenanceDone}</div><div className="text-orange-400">Verifikasi</div></div>
            <div className="bg-green-50 p-2 rounded text-center"><div className="font-bold text-green-600">{stats.approved}</div><div className="text-green-400">Final</div></div>
          </div>

          {isQC && (
            <div className="mt-2 flex flex-wrap gap-3 text-xs">
              <span className="bg-blue-50 px-3 py-1 rounded-full text-blue-600">📋 Pemeriksaan: {totalPemeriksaan}</span>
              <span className="bg-green-50 px-3 py-1 rounded-full text-green-600">🔧 Pemeliharaan: {totalPemeliharaan}</span>
            </div>
          )}
        </header>

        {currentMenu === 'dashboard' ? (
          <div className="space-y-6">
            <div className="no-print flex flex-wrap items-center justify-between gap-4">
              <div className="flex flex-wrap gap-3">
                {isQC && (
                  <>
                    <button onClick={handleNewPemeriksaan} className="bg-blue-600 hover:bg-blue-700 text-white px-6 py-3 rounded-lg flex items-center gap-2 shadow-md">
                      <span className="text-xl">📋</span> New Pemeriksaan
                    </button>
                    <button onClick={handleNewPemeliharaan} className="bg-green-600 hover:bg-green-700 text-white px-6 py-3 rounded-lg flex items-center gap-2 shadow-md">
                      <span className="text-xl">🔧</span> New Pemeliharaan
                    </button>
                  </>
                )}
                {isMaintenance && <p className="text-sm text-gray-500 self-center">🔧 Tugas dari QC - Isi nama teknisi pelaksana</p>}
                {isSales && <p className="text-sm text-orange-500 self-center">👀 Hanya melihat Data Pemeriksaan</p>}
              </div>

              <div className="flex items-center gap-2">
                <span className="text-sm text-gray-500">Urutkan:</span>
                <button onClick={() => setSortOrder('terbaru')} className={`px-3 py-1 rounded-md text-sm transition ${sortOrder === 'terbaru' ? 'bg-blue-600 text-white' : 'bg-gray-200 text-gray-700 hover:bg-gray-300'}`}>📅 Terbaru</button>
                <button onClick={() => setSortOrder('terlama')} className={`px-3 py-1 rounded-md text-sm transition ${sortOrder === 'terlama' ? 'bg-blue-600 text-white' : 'bg-gray-200 text-gray-700 hover:bg-gray-300'}`}>📅 Terlama</button>
              </div>
            </div>

            <div className="bg-white shadow-sm rounded-lg border border-gray-200 overflow-hidden">
              <div className="bg-gray-50 px-6 py-3 border-b border-gray-200 flex justify-between items-center">
                <div>
                  <h2 className="font-semibold text-gray-700">
                    {isQC ? '📋 Semua Data' : isMaintenance ? '🔧 Tugas Maintenance' : '📋 Data Pemeriksaan (Sales)'}
                  </h2>
                  <span className="text-sm text-gray-500">{filteredReports.length} laporan</span>
                </div>
              </div>

              {loading ? (
                <div className="p-8 text-center text-gray-500">⏳ Memuat data...</div>
              ) : (() => {
                if (filteredReports.length === 0) {
                  return (
                    <div className="p-8 text-center text-gray-500">
                      {isQC ? <p>Belum ada laporan. Klik tombol <strong>"New Pemeriksaan"</strong> atau <strong>"New Pemeliharaan"</strong> untuk mulai.</p> : isMaintenance ? <p>Belum ada tugas dari QC.</p> : <p>Belum ada laporan Pemeriksaan dari QC.</p>}
                    </div>
                  );
                }

                return (
                  <div className="overflow-x-auto">
                    <table className="w-full" style={{ fontSize: '14px' }}>
                      <thead className="bg-gray-50">
                        <tr>
                          <th className="px-4 py-2 text-left text-gray-600" style={{ fontSize: '13px' }}>Tanggal</th>
                          <th className="px-4 py-2 text-left text-gray-600" style={{ fontSize: '13px' }}>Unit</th>
                          <th className="px-4 py-2 text-left text-gray-600" style={{ fontSize: '13px' }}>Kode Proyek</th>
                          <th className="px-4 py-2 text-left text-gray-600" style={{ fontSize: '13px' }}>Pelanggan</th>
                          <th className="px-4 py-2 text-left text-gray-600" style={{ fontSize: '13px' }}>Tipe</th>
                          {isQC && <th className="px-4 py-2 text-left text-gray-600" style={{ fontSize: '13px' }}>Teknisi</th>}
                          <th className="px-4 py-2 text-left text-gray-600" style={{ fontSize: '13px' }}>Status</th>
                          <th className="px-4 py-2 text-left text-gray-600" style={{ fontSize: '13px' }}>Aksi</th>
                        </tr>
                      </thead>
                      <tbody>
                        {filteredReports.map((report) => {
                          const firstUnit = report.units && report.units.length > 0 ? report.units[0] : null;
                          const unitDisplay = report.reportType === 'pemeliharaan' && report.jumlahUnit > 1 ? `${report.jumlahUnit} Unit` : firstUnit?.unitNumber || report.unitNumber || '-';
                          const projectDisplay = report.reportType === 'pemeliharaan' && report.jumlahUnit > 1 ? `${report.jumlahUnit} Proyek` : firstUnit?.projectCode || report.projectCode || '-';
                          const customerDisplay = report.reportType === 'pemeliharaan' && report.jumlahUnit > 1 ? `${report.jumlahUnit} Pelanggan` : firstUnit?.customerName || report.customerName || '-';

                          return (
                            <tr key={report.id} className="border-t border-gray-100 hover:bg-gray-50">
                              <td className="px-4 py-2 text-gray-500 whitespace-nowrap" style={{ fontSize: '13px' }}>{formatDate(report.createdAt)}</td>
                              <td className="px-4 py-2 font-medium" style={{ fontSize: '13px' }}>{unitDisplay}</td>
                              <td className="px-4 py-2" style={{ fontSize: '13px' }}>{projectDisplay}</td>
                              <td className="px-4 py-2" style={{ fontSize: '13px' }}>{customerDisplay}</td>
                              <td className="px-4 py-2">
                                <span className={`px-2 py-1 rounded-full ${report.reportType === 'pemeriksaan' ? 'bg-blue-100 text-blue-700' : 'bg-green-100 text-green-700'}`} style={{ fontSize: '12px' }}>
                                  {report.reportType === 'pemeriksaan' ? '📋 Pemeriksaan' : '🔧 Pemeliharaan'}
                                </span>
                              </td>
                              {isQC && <td className="px-4 py-2" style={{ fontSize: '13px' }}>{report.teknisiName || 'Belum ditunjuk'}</td>}
                              <td className="px-4 py-2">
                                <span className={`inline-flex px-2 py-1 rounded-full font-semibold ${report.status === 'draft' ? 'bg-gray-200 text-gray-700' : report.status === 'qc_approved' ? 'bg-blue-100 text-blue-700' : report.status === 'revision' ? 'bg-yellow-100 text-yellow-700' : report.status === 'maintenance_done' ? 'bg-orange-100 text-orange-700' : report.status === 'approved' ? 'bg-green-100 text-green-700' : 'bg-gray-200 text-gray-700'}`} style={{ fontSize: '12px' }}>
                                  {report.status === 'draft' ? '📝 Draft QC' : report.status === 'qc_approved' ? '📤 Menunggu Maintenance' : report.status === 'revision' ? '🔄 Revisi' : report.status === 'maintenance_done' ? '🔧 Verifikasi QC' : report.status === 'approved' ? '✅ Final Approved' : '📝 Draft'}
                                </span>
                              </td>
                              <td className="px-4 py-2">
                                <button onClick={() => handleViewReport(report)} className="text-blue-600 hover:text-blue-800 font-medium" style={{ fontSize: '13px' }}>👁️ Lihat</button>
                                {isQC && report.status !== 'approved' && (
                                  <button onClick={() => handleDeleteReport(report.id)} className="text-red-500 hover:text-red-700 font-medium ml-2" style={{ fontSize: '13px' }}>🗑️</button>
                                )}
                              </td>
                            </tr>
                          );
                        })}
                      </tbody>
                    </table>
                  </div>
                );
              })()}
            </div>
          </div>
        ) : (
          <div className="bg-white shadow-sm rounded-lg border border-gray-200 p-6">
            {selectedReport && (
              <ReportForm
                report={selectedReport}
                onSave={handleSaveReport}
                onCancel={handleCancel}
                onDelete={handleDeleteReport}
                userRole={role}
                isReadOnly={isReadOnly}
                onPrintPDF={handlePrintPDF}
                onExportExcel={handleExportExcel}
                reportType={currentMenu === 'newPemeliharaan' || selectedReport.reportType === 'pemeliharaan' ? 'pemeliharaan' : 'pemeriksaan'}
              />
            )}
          </div>
        )}
      </div>

      <style jsx global>{`
        @media print {
          .no-print { display: none !important; }
          body { background: white !important; padding: 0 !important; margin: 0 !important; font-size: 10px !important; }
          .container { max-width: 100% !important; padding: 10px !important; margin: 0 !important; }
          #print-content { display: block !important; }
          table { page-break-inside: auto; width: 100% !important; font-size: 9px !important; }
          tr { page-break-inside: avoid; page-break-after: auto; }
          thead { display: table-header-group; }
          .bg-blue-700, .bg-green-700, .bg-orange-600 { background: #1a365d !important; -webkit-print-color-adjust: exact !important; print-color-adjust: exact !important; }
          .bg-blue-700 h2, .bg-green-700 h2, .bg-orange-600 h2 { color: white !important; }
          .bg-gray-100 { background: #f3f4f6 !important; -webkit-print-color-adjust: exact !important; print-color-adjust: exact !important; }
          .bg-green-100 { background: #d1fae5 !important; -webkit-print-color-adjust: exact !important; print-color-adjust: exact !important; }
          .bg-red-100 { background: #fee2e2 !important; -webkit-print-color-adjust: exact !important; print-color-adjust: exact !important; }
          .bg-blue-100 { background: #dbeafe !important; -webkit-print-color-adjust: exact !important; print-color-adjust: exact !important; }
          .bg-orange-100 { background: #ffedd5 !important; -webkit-print-color-adjust: exact !important; print-color-adjust: exact !important; }
          .bg-yellow-100 { background: #fef9c3 !important; -webkit-print-color-adjust: exact !important; print-color-adjust: exact !important; }
          .border { border: 1px solid #e5e7eb !important; }
          .shadow-sm { box-shadow: none !important; }
          .text-sm { font-size: 9px !important; }
          .text-xs { font-size: 8px !important; }
          .text-lg { font-size: 14px !important; }
          .text-2xl { font-size: 18px !important; }
          .p-6 { padding: 8px !important; }
          .p-4 { padding: 6px !important; }
          .px-6 { padding-left: 8px !important; padding-right: 8px !important; }
          .py-3 { padding-top: 4px !important; padding-bottom: 4px !important; }
          .py-2 { padding-top: 3px !important; padding-bottom: 3px !important; }
          .px-2 { padding-left: 4px !important; padding-right: 4px !important; }
          .mb-6 { margin-bottom: 8px !important; }
          .gap-4 { gap: 4px !important; }
          .print\\:break-inside-avoid { break-inside: avoid !important; page-break-inside: avoid !important; }
          .print\\:table-fixed { table-layout: fixed !important; }
          .print\\:block { display: block !important; }
          .print\\:hidden { display: none !important; }
          img { max-width: 80px !important; max-height: 80px !important; }
        }
      `}</style>
    </div>
  );
}

// ============================================
// KOMPONEN REPORT FORM - FULL LENGKAP
// ============================================
function ReportForm({
  report,
  onSave,
  onCancel,
  onDelete,
  userRole,
  isReadOnly,
  onPrintPDF,
  onExportExcel,
  reportType,
}: {
  report: Report | null;
  onSave: (data: any) => void;
  onCancel: () => void;
  onDelete?: (id: string) => void;
  userRole: UserRole;
  isReadOnly: boolean;
  onPrintPDF?: () => void;
  onExportExcel?: () => void;
  reportType?: ReportType;
}) {
  const isQC = userRole === 'qc';
  const isMaintenance = userRole === 'maintenance';
  const isPemeriksaan = reportType === 'pemeriksaan';
  const isPemeliharaan = reportType === 'pemeliharaan';

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
  });

  const calculateTotalScore = (inspectionData: any[]) => {
    let totalWeight = 0;
    let achievedWeight = 0;
    
    inspectionData.forEach((item: any) => {
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

  const calculateHitung = (item: any, originalItem: any) => {
    if (!originalItem) return 0;
    if (item.status === 'Good' || item.status === 'N/A') return originalItem.weight;
    if (item.status === 'Not Good' && item.isApproved) return originalItem.weight;
    if (item.status === 'Not Good' && !item.isApproved) return -originalItem.weight;
    return 0;
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
    
    setFormData((prev: any) => ({
      ...prev,
      jumlahUnit: newCount,
      units: newUnits,
    }));
  };

  const updateUnitData = (index: number, field: string, value: string) => {
    if (isReadOnly || !isQC || formData.status !== 'draft') return;
    setFormData((prev: any) => {
      const newUnits = [...prev.units];
      newUnits[index] = { ...newUnits[index], [field]: value };
      return { ...prev, units: newUnits };
    });
  };

  const handleTeknisiNameChange = (name: string) => {
    if (!isMaintenance || isReadOnly) return;
    setFormData((prev: any) => ({ ...prev, teknisiName: name }));
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
    let updatedData: any;

    if (isQC && formData.status === 'draft') {
      if (isPemeriksaan) {
        const emptyItems = formData.inspectionData.filter((item: any) => !item.status);
        if (emptyItems.length > 0) {
          alert(`⚠️ Terdapat ${emptyItems.length} item yang belum diisi status.`);
          return;
        }
        const invalidItems = formData.inspectionData.filter(
          (item: any) => (item.status === 'Not Good') && !item.finding
        );
        if (invalidItems.length > 0) {
          alert(`⚠️ Terdapat ${invalidItems.length} item Not Good tanpa temuan.`);
          return;
        }
      } else if (isPemeliharaan) {
        let hasError = false;
        formData.units.forEach((unit: UnitData, idx: number) => {
          const emptyItems = unit.maintenanceData.filter((item: any) => !item.status);
          if (emptyItems.length > 0) {
            alert(`⚠️ Unit ${idx + 1}: Terdapat ${emptyItems.length} item yang belum diisi status.`);
            hasError = true;
          }
        });
        if (hasError) return;
      }

      updatedData = {
        ...formData,
        submittedBy: 'qc',
        submittedAt: new Date().toLocaleString('id-ID'),
        status: 'qc_approved',
      };
      
      onSave(updatedData);
      alert(`✅ Data ${isPemeriksaan ? 'Pemeriksaan' : 'Pemeliharaan'} dikirim ke Maintenance!`);
      return;
    } else if (isQC && formData.status === 'maintenance_done') {
      if (action === 'approve') {
        if (isPemeriksaan) {
          const notGoodItems = formData.inspectionData.filter((item: any) => item.status === 'Not Good');
          const unapprovedItems = notGoodItems.filter((item: any) => !item.isApproved);
          if (unapprovedItems.length > 0) {
            alert(`⚠️ Masih ada ${unapprovedItems.length} item Not Good yang belum di-approve!`);
            return;
          }
        } else if (isPemeliharaan) {
          let unapprovedCount = 0;
          formData.units.forEach((unit: UnitData) => {
            const notGoodItems = unit.maintenanceData.filter((item: any) => item.status === 'Not Good' && !item.isApproved);
            unapprovedCount += notGoodItems.length;
          });
          if (unapprovedCount > 0) {
            alert(`⚠️ Masih ada ${unapprovedCount} item Not Good yang belum di-approve!`);
            return;
          }
        }
        
        updatedData = {
          ...formData,
          qcVerification: {
            ...formData.qcVerification,
            verifiedAt: new Date().toLocaleString('id-ID'),
          },
          status: 'approved',
        };
        alert('✅ Laporan FINAL APPROVED!');
      } else if (action === 'revision') {
        updatedData = {
          ...formData,
          qcVerification: {
            ...formData.qcVerification,
            verifiedAt: new Date().toLocaleString('id-ID'),
            qcStatus: 'Revision Required',
          },
          status: 'revision',
        };
        alert(`📝 REVISION! Dikirim ke Maintenance untuk perbaikan.`);
      }
    } else if (isMaintenance && (formData.status === 'qc_approved' || formData.status === 'revision')) {
      if (!formData.teknisiName || formData.teknisiName.trim() === '') {
        alert('Silakan isi nama teknisi yang bertugas terlebih dahulu!');
        return;
      }

      if (isPemeriksaan) {
        const notGoodItems = formData.inspectionData.filter((item: any) => item.status === 'Not Good');
        const invalidItems = notGoodItems.filter((item: any) => !item.photoAfter || !item.repairNote);
        if (invalidItems.length > 0) {
          alert(`Terdapat ${invalidItems.length} item Not Good tanpa foto setelah perbaikan atau catatan.`);
          return;
        }
      } else if (isPemeliharaan) {
        let hasError = false;
        formData.units.forEach((unit: UnitData, idx: number) => {
          const notGoodItems = unit.maintenanceData.filter((item: any) => item.status === 'Not Good');
          const invalidItems = notGoodItems.filter((item: any) => !item.photoAfter || !item.repairNote);
          if (invalidItems.length > 0) {
            alert(`Unit ${idx + 1}: Terdapat ${invalidItems.length} item Not Good tanpa foto setelah perbaikan atau catatan.`);
            hasError = true;
          }
        });
        if (hasError) return;
      }
      
      updatedData = {
        ...formData,
        submittedBy: 'maintenance',
        submittedAt: new Date().toLocaleString('id-ID'),
        status: 'maintenance_done',
      };
      alert(`✅ Perbaikan selesai oleh ${formData.teknisiName}! Dikirim ke QC untuk verifikasi.`);
    } else {
      alert('⚠️ Status tidak sesuai untuk tindakan ini.');
      return;
    }
    onSave(updatedData);
  };

  const canEdit = !isReadOnly;
  const isApproved = formData.status === 'approved' || formData.status === 'qc_approved';
  const typeLabel = isPemeriksaan ? 'Pemeriksaan' : 'Pemeliharaan';
  
  const allNotGoodApproved = isPemeriksaan ? 
    formData.inspectionData.filter((item: any) => item.status === 'Not Good').every((item: any) => item.isApproved === true) : 
    true;

  const canFinalApprove = formData.status === 'maintenance_done' && allNotGoodApproved;
  const totalScore = isPemeriksaan ? calculateTotalScore(formData.inspectionData) : 0;

  const renderTeknisiInput = () => {
    if (isMaintenance && !isReadOnly && (formData.status === 'qc_approved' || formData.status === 'revision')) {
      return (
        <div className="bg-yellow-50 p-4 rounded-lg border border-yellow-200 mb-6">
          <label className="block text-sm font-medium text-gray-700 mb-1">
            Nama Teknisi Pelaksana <span className="text-red-500">*</span>
          </label>
          <div className="flex flex-col sm:flex-row sm:items-center gap-3">
            <input
              type="text"
              value={formData.teknisiName || ''}
              onChange={(e) => handleTeknisiNameChange(e.target.value)}
              placeholder="Masukkan nama teknisi yang bertugas"
              className="flex-1 px-3 py-2 border border-blue-400 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500 bg-white"
            />
            {formData.teknisiName && (
              <span className="text-xs text-green-600 font-medium">
                ✅ Nama tersimpan: {formData.teknisiName}
              </span>
            )}
          </div>
          <p className="text-xs text-gray-400 mt-1">
            ✏️ Isi nama teknisi yang akan melaksanakan perbaikan (berlaku untuk semua unit)
          </p>
        </div>
      );
    }

    if (formData.teknisiName) {
      return (
        <div className="bg-blue-50 p-3 rounded-lg border border-blue-200 mb-6">
          <span className="text-sm font-medium text-blue-700">👨‍🔧 Teknisi Pelaksana:</span>
          <span className="ml-2 text-blue-600 font-semibold">{formData.teknisiName}</span>
          {isQC && formData.status === 'maintenance_done' && (
            <span className="ml-2 text-xs text-orange-500">(Menunggu verifikasi)</span>
          )}
        </div>
      );
    }

    if (isMaintenance && (formData.status === 'qc_approved' || formData.status === 'revision') && !formData.teknisiName) {
      return (
        <div className="bg-red-50 p-4 rounded-lg border border-red-200 mb-6">
          <p className="text-sm text-red-600 font-medium">
            ⚠️ Nama teknisi belum diisi! Silakan isi nama teknisi pelaksana di atas.
          </p>
        </div>
      );
    }

    return null;
  };

  return (
    <form onSubmit={handleSubmit} className="space-y-6">
      <div className="no-print flex flex-wrap justify-between items-center gap-2">
        <button type="button" onClick={onCancel} className="text-gray-500 hover:text-gray-700 flex items-center gap-2">← Kembali</button>
        <div className="flex flex-wrap gap-2">
          {canEdit && <button type="button" onClick={() => onDelete?.(formData.id)} className="bg-red-600 hover:bg-red-700 text-white px-4 py-2 rounded-lg">🗑️ Hapus</button>}
          
          {isQC && formData.status === 'draft' && (
            <button type="button" onClick={() => handleKirim('submit')} className="bg-blue-600 hover:bg-blue-700 text-white px-6 py-2 rounded-lg">📤 Submit ke Maintenance</button>
          )}

          {isQC && formData.status === 'maintenance_done' && (
            <>
              <button type="button" onClick={() => handleKirim('revision')} className="bg-yellow-600 hover:bg-yellow-700 text-white px-6 py-2 rounded-lg">🔄 Revisi</button>
              <button 
                type="button" 
                onClick={() => {
                  if (!canFinalApprove) {
                    alert('⚠️ Masih ada item Not Good yang belum di-approve!');
                    return;
                  }
                  handleKirim('approve');
                }} 
                className={`px-6 py-2 rounded-lg ${canFinalApprove ? 'bg-green-600 hover:bg-green-700 text-white' : 'bg-gray-400 text-gray-600 cursor-not-allowed'}`}
                disabled={!canFinalApprove}
              >
                ✅ Approve (Final)
              </button>
            </>
          )}

          {(isMaintenance && (formData.status === 'qc_approved' || formData.status === 'revision')) && (
            <button
              type="button"
              onClick={() => {
                if (!formData.teknisiName || formData.teknisiName.trim() === '') {
                  alert('Silakan isi nama teknisi yang bertugas terlebih dahulu!');
                  return;
                }
                if (isPemeriksaan) {
                  const notGoodItems2 = formData.inspectionData.filter((item: any) => item.status === 'Not Good');
                  const invalidItems = notGoodItems2.filter((item: any) => !item.photoAfter || !item.repairNote);
                  if (invalidItems.length > 0) {
                    alert(`Terdapat ${invalidItems.length} item Not Good tanpa foto setelah perbaikan atau catatan.`);
                    return;
                  }
                } else if (isPemeliharaan) {
                  let hasError = false;
                  formData.units.forEach((unit: UnitData, idx: number) => {
                    const notGoodItems = unit.maintenanceData.filter((item: any) => item.status === 'Not Good');
                    const invalidItems = notGoodItems.filter((item: any) => !item.photoAfter || !item.repairNote);
                    if (invalidItems.length > 0) {
                      alert(`Unit ${idx + 1}: Terdapat ${invalidItems.length} item Not Good tanpa foto setelah perbaikan atau catatan.`);
                      hasError = true;
                    }
                  });
                  if (hasError) return;
                }
                const updatedData = {
                  ...formData,
                  submittedBy: 'maintenance',
                  submittedAt: new Date().toLocaleString('id-ID'),
                  status: 'maintenance_done',
                };
                onSave(updatedData);
                alert(`✅ Perbaikan selesai oleh ${formData.teknisiName}! Dikirim ke QC untuk verifikasi.`);
              }}
              className="bg-orange-600 hover:bg-orange-700 text-white px-6 py-2 rounded-lg"
            >
              🔧 Kirim ke QC
            </button>
          )}

          {isQC && isApproved && onPrintPDF && (
            <button type="button" onClick={onPrintPDF} className="bg-red-600 hover:bg-red-700 text-white px-6 py-2 rounded-lg flex items-center gap-2">📄 Download PDF</button>
          )}

          {isPemeliharaan && onExportExcel && (
            <button type="button" onClick={onExportExcel} className="bg-green-600 hover:bg-green-700 text-white px-6 py-2 rounded-lg flex items-center gap-2">📊 Export Excel</button>
          )}

          {isReadOnly && <span className="text-sm text-gray-400 self-center">🔒 Mode Read-Only</span>}
          
          {isQC && formData.status === 'maintenance_done' && (
            <span className="text-sm text-orange-500 self-center">
              {allNotGoodApproved ? '✅ Semua Not Good sudah di-approve!' : '⚠️ Approve setiap item Not Good yang sudah diperbaiki'}
            </span>
          )}
        </div>
      </div>

      <div id="print-content" className="print-content">
        <div className="text-center border-b-2 border-gray-300 pb-4 mb-6 print:block hidden">
          <h1 className="text-2xl font-bold text-blue-700">LOUSERINDO MEGAH PERMAI</h1>
          <p className="text-sm text-gray-500">Elevator Quality Control System</p>
          <p className="text-xs text-gray-400 mt-1">Laporan {typeLabel}</p>
        </div>

        {renderTeknisiInput()}

        {isPemeliharaan && isQC && !isReadOnly && formData.status === 'draft' && (
          <section className="bg-white shadow-sm rounded-lg border border-gray-200 overflow-hidden mb-6">
            <div className="bg-blue-700 px-6 py-3">
              <h2 className="text-lg font-semibold text-white">📊 Jumlah Unit Proyek</h2>
            </div>
            <div className="p-6">
              <div className="flex flex-col items-center">
                <div className="w-full max-w-md">
                  <label className="block text-sm font-medium text-gray-700 mb-2">
                    Pilih Jumlah Unit <span className="text-red-500">*</span>
                  </label>
                  <input
                    type="number"
                    min="1"
                    max="100"
                    value={selectedJumlahUnit}
                    onChange={(e) => handleJumlahUnitChange(Number(e.target.value))}
                    className="w-full px-4 py-3 text-lg border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500 text-center"
                  />
                  <p className="text-xs text-gray-400 mt-2 text-center">Maksimal 100 unit.</p>
                  <p className="text-sm text-blue-600 mt-2 text-center font-semibold">
                    {selectedJumlahUnit} Unit {selectedJumlahUnit > 1 ? 'terpilih' : 'terpilih'}
                  </p>
                </div>
              </div>
            </div>
          </section>
        )}

        {isPemeliharaan && (isReadOnly || !isQC || formData.status !== 'draft') && formData.jumlahUnit && (
          <div className="bg-blue-50 p-4 rounded-lg border border-blue-200 mb-6">
            <div className="flex items-center justify-between">
              <span className="text-sm font-medium text-blue-700">📊 Jumlah Unit Proyek:</span>
              <span className="text-2xl font-bold text-blue-800">{formData.jumlahUnit} Unit</span>
            </div>
          </div>
        )}

        {isPemeriksaan && (
          <>
            <section className="bg-white shadow-sm rounded-lg border border-gray-200 overflow-hidden mb-6">
              <div className="bg-blue-700 px-6 py-3"><h2 className="text-lg font-semibold text-white">Data Umum</h2></div>
              <div className="p-6">
                <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                  <div>
                    <label className="block text-sm font-medium text-gray-700 mb-1">No Unit</label>
                    <input
                      type="text"
                      value={formData.unitNumber || ''}
                      onChange={(e) => setFormData((prev: any) => ({ ...prev, unitNumber: e.target.value }))}
                      className="w-full px-3 py-2 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500"
                      disabled={isReadOnly}
                      placeholder="Masukkan No Unit"
                    />
                  </div>
                  <div>
                    <label className="block text-sm font-medium text-gray-700 mb-1">Kode Proyek</label>
                    <input
                      type="text"
                      value={formData.projectCode || ''}
                      onChange={(e) => setFormData((prev: any) => ({ ...prev, projectCode: e.target.value }))}
                      className="w-full px-3 py-2 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500"
                      disabled={isReadOnly}
                      placeholder="Masukkan Kode Proyek"
                    />
                  </div>
                  <div>
                    <label className="block text-sm font-medium text-gray-700 mb-1">Nama Pelanggan</label>
                    <input
                      type="text"
                      value={formData.customerName || ''}
                      onChange={(e) => setFormData((prev: any) => ({ ...prev, customerName: e.target.value }))}
                      className="w-full px-3 py-2 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500"
                      disabled={isReadOnly}
                      placeholder="Masukkan Nama Pelanggan"
                    />
                  </div>
                  <div>
                    <label className="block text-sm font-medium text-gray-700 mb-1">Lokasi Gedung</label>
                    <input
                      type="text"
                      value={formData.buildingLocation || ''}
                      onChange={(e) => setFormData((prev: any) => ({ ...prev, buildingLocation: e.target.value }))}
                      className="w-full px-3 py-2 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500"
                      disabled={isReadOnly}
                      placeholder="Masukkan Lokasi Gedung"
                    />
                  </div>
                  <div>
                    <label className="block text-sm font-medium text-gray-700 mb-1">Tipe / Merk Elevator</label>
                    <input
                      type="text"
                      value={formData.elevatorType || ''}
                      onChange={(e) => setFormData((prev: any) => ({ ...prev, elevatorType: e.target.value }))}
                      className="w-full px-3 py-2 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500"
                      disabled={isReadOnly}
                      placeholder="Masukkan Tipe / Merk"
                    />
                  </div>
                  <div>
                    <label className="block text-sm font-medium text-gray-700 mb-1">Kapasitas Angkut (kg)</label>
                    <input
                      type="text"
                      value={formData.capacity || ''}
                      onChange={(e) => setFormData((prev: any) => ({ ...prev, capacity: e.target.value }))}
                      className="w-full px-3 py-2 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500"
                      disabled={isReadOnly}
                      placeholder="Masukkan Kapasitas"
                    />
                  </div>
                  <div>
                    <label className="block text-sm font-medium text-gray-700 mb-1">Kecepatan (m/menit)</label>
                    <input
                      type="text"
                      value={formData.speed || ''}
                      onChange={(e) => setFormData((prev: any) => ({ ...prev, speed: e.target.value }))}
                      className="w-full px-3 py-2 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500"
                      disabled={isReadOnly}
                      placeholder="Masukkan Kecepatan"
                    />
                  </div>
                  <div>
                    <label className="block text-sm font-medium text-gray-700 mb-1">Tanggal Pemeriksaan</label>
                    <input
                      type="date"
                      value={formData.inspectionDate || ''}
                      onChange={(e) => setFormData((prev: any) => ({ ...prev, inspectionDate: e.target.value }))}
                      className="w-full px-3 py-2 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500"
                      disabled={isReadOnly}
                    />
                  </div>
                  <div>
                    <label className="block text-sm font-medium text-gray-700 mb-1">Nama QC</label>
                    <input
                      type="text"
                      value={formData.qcName || ''}
                      onChange={(e) => setFormData((prev: any) => ({ ...prev, qcName: e.target.value }))}
                      className="w-full px-3 py-2 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500"
                      disabled={isReadOnly}
                      placeholder="Masukkan Nama QC"
                    />
                  </div>
                </div>
              </div>
            </section>

            <div className="bg-gray-50 p-3 rounded-lg border border-gray-200 mb-6">
              <div className="flex flex-wrap items-center gap-4">
                <span className="text-sm font-medium text-gray-600">Status:</span>
                <span className={`inline-flex px-3 py-1 rounded-full text-sm font-medium ${
                  formData.status === 'draft' ? 'bg-gray-200 text-gray-700' :
                  formData.status === 'qc_approved' ? 'bg-blue-100 text-blue-700' :
                  formData.status === 'revision' ? 'bg-yellow-100 text-yellow-700' :
                  formData.status === 'maintenance_done' ? 'bg-orange-100 text-orange-700' :
                  formData.status === 'approved' ? 'bg-green-100 text-green-700' :
                  'bg-gray-200 text-gray-700'
                }`}>
                  {formData.status === 'draft' ? '📝 Draft QC' :
                   formData.status === 'qc_approved' ? '📤 Menunggu Maintenance' :
                   formData.status === 'revision' ? '🔄 REVISION' :
                   formData.status === 'maintenance_done' ? '🔧 Maintenance Selesai' :
                   formData.status === 'approved' ? '✅ FINAL APPROVED' :
                   '📝 Draft'}
                </span>
                {formData.submittedAt && <span className="text-xs text-gray-400">Dikirim: {formData.submittedAt}</span>}
                {formData.qcVerification?.verifiedAt && <span className="text-xs text-gray-400">Diverifikasi: {formData.qcVerification.verifiedAt}</span>}
              </div>
            </div>

            <section className="bg-white shadow-sm rounded-lg border border-gray-200 overflow-hidden mb-6 print:break-inside-avoid">
              <div className="bg-blue-700 px-6 py-3 flex items-center justify-between">
                <h2 className="text-lg font-semibold text-white">Daftar Simak Pemeriksaan (116 Item)</h2>
                <span className="text-sm text-white bg-blue-600 px-3 py-1 rounded-full">{formData.inspectionData.length} Item</span>
              </div>
              <div className="p-4 overflow-x-auto">
                <table className="w-full border-collapse print:table-fixed" style={{ fontSize: '14px' }}>
                  <thead>
                    <tr className="bg-gray-100 border-b border-gray-200">
                      <th className="text-left py-2 px-2 font-semibold text-gray-700 w-10" style={{ fontSize: '13px' }}>No</th>
                      <th className="text-left py-2 px-2 font-semibold text-gray-700 w-20" style={{ fontSize: '13px' }}>Kategori</th>
                      <th className="text-left py-2 px-2 font-semibold text-gray-700" style={{ fontSize: '13px' }}>Item</th>
                      <th className="text-left py-2 px-2 font-semibold text-gray-700 w-24" style={{ fontSize: '13px' }}>Prioritas</th>
                      <th className="text-left py-2 px-2 font-semibold text-gray-700 w-16" style={{ fontSize: '13px' }}>Bobot</th>
                      <th className="text-left py-2 px-2 font-semibold text-gray-700 w-24" style={{ fontSize: '13px' }}>Status</th>
                      <th className="text-left py-2 px-2 font-semibold text-gray-700 w-32" style={{ fontSize: '13px' }}>Temuan QC</th>
                      <th className="text-left py-2 px-2 font-semibold text-gray-700 w-24" style={{ fontSize: '13px' }}>Foto Sebelum</th>
                      <th className="text-left py-2 px-2 font-semibold text-gray-700 w-24" style={{ fontSize: '13px' }}>Foto Setelah</th>
                      <th className="text-left py-2 px-2 font-semibold text-gray-700 w-32" style={{ fontSize: '13px' }}>Catatan Perbaikan</th>
                      <th className="text-left py-2 px-2 font-semibold text-gray-700 w-20" style={{ fontSize: '13px' }}>Approve QC</th>
                      <th className="text-left py-2 px-2 font-semibold text-gray-700 w-16" style={{ fontSize: '13px' }}>Hitung</th>
                    </tr>
                  </thead>
                  <tbody>
                    {formData.inspectionData.map((item: any) => {
                      const originalItem = inspectionItems.find(i => i.id === item.id);
                      const isNotGood = item.status === 'Not Good';
                      const isApproved = item.isApproved;
                      const isEditable = isQC && !isReadOnly && (formData.status === 'draft');
                      const isMaintEditable = isMaintenance && !isReadOnly && (formData.status === 'qc_approved' || formData.status === 'revision');
                      const canApprove = isQC && !isReadOnly && (formData.status === 'maintenance_done' || formData.status === 'revision') && isNotGood && !isApproved;
                      
                      const hitung = calculateHitung(item, originalItem);
                      const bobotPersen = (originalItem?.weight || 0) * 100;
                      const hitungPersen = hitung * 100;

                      return (
                        <tr key={item.id} className="border-b border-gray-100 print:break-inside-avoid">
                          <td className="py-2 px-2 text-gray-600 align-top" style={{ fontSize: '13px' }}>{item.id}</td>
                          <td className="py-2 px-2 text-gray-600 align-top uppercase font-medium" style={{ fontSize: '12px' }}>{originalItem?.category || '-'}</td>
                          <td className="py-2 px-2 text-gray-800 align-top" style={{ fontSize: '13px' }}>{originalItem?.item || '-'}</td>
                          <td className="py-2 px-2 text-gray-600 align-top" style={{ fontSize: '12px' }}>{originalItem?.priority || '-'}</td>
                          <td className="py-2 px-2 text-gray-600 align-top" style={{ fontSize: '12px' }}>{bobotPersen.toFixed(1)}%</td>
                          <td className="py-2 px-2 align-top">
                            {isEditable ? (
                              <div className="flex flex-col gap-1">
                                <label className="inline-flex items-center" style={{ fontSize: '12px' }}>
                                  <input type="radio" name={`status-${item.id}`} value="Good" checked={item.status === 'Good'} onChange={() => handleStatusChange(item.id, 'Good')} className="w-4 h-4 text-green-600"/>
                                  <span className="ml-1">Good</span>
                                </label>
                                <label className="inline-flex items-center" style={{ fontSize: '12px' }}>
                                  <input type="radio" name={`status-${item.id}`} value="Not Good" checked={item.status === 'Not Good'} onChange={() => handleStatusChange(item.id, 'Not Good')} className="w-4 h-4 text-red-600"/>
                                  <span className="ml-1">Not Good</span>
                                </label>
                                <label className="inline-flex items-center" style={{ fontSize: '12px' }}>
                                  <input type="radio" name={`status-${item.id}`} value="N/A" checked={item.status === 'N/A'} onChange={() => handleStatusChange(item.id, 'N/A')} className="w-4 h-4 text-gray-400"/>
                                  <span className="ml-1">N/A</span>
                                </label>
                              </div>
                            ) : (
                              <span className={`inline-flex px-2 py-1 rounded-full font-semibold ${
                                item.status === 'Good' ? 'bg-green-100 text-green-800' :
                                item.status === 'N/A' ? 'bg-gray-300 text-gray-700' :
                                item.status === 'Not Good' && isApproved ? 'bg-green-300 text-green-900' :
                                item.status === 'Not Good' && !isApproved ? 'bg-red-100 text-red-800' :
                                'bg-gray-100 text-gray-500'
                              }`} style={{ fontSize: '12px' }}>
                                {item.status === 'Not Good' && isApproved ? '✅ Approved' : item.status || '-'}
                              </span>
                            )}
                          </td>
                          <td className="py-2 px-2 align-top">
                            {isEditable ? (
                              <textarea value={item.finding} onChange={(e) => handleFindingChange(item.id, e.target.value)} disabled={item.status === 'Good' || item.status === 'N/A'} placeholder={isNotGood ? 'Wajib diisi' : 'Temuan'} className={`w-full px-2 py-1 border rounded-md ${isNotGood && !item.finding ? 'border-red-300 bg-red-50' : 'border-gray-300'} ${item.status === 'Good' ? 'bg-gray-100' : 'bg-white'}`} rows={2} style={{ fontSize: '13px' }}/>
                            ) : (
                              <div style={{ fontSize: '13px' }} className="text-gray-600">{item.finding || '-'}</div>
                            )}
                          </td>
                          <td className="py-2 px-2 align-top">
                            {isEditable && isNotGood ? (
                              <div>
                                <input type="file" accept="image/*" onChange={(e) => handlePhotoBeforeChange(item.id, e)} style={{ fontSize: '12px' }}/>
                                {item.photoBefore && <div className="relative inline-block mt-1"><img src={item.photoBefore} alt="Before" className="w-20 h-20 object-cover rounded border"/></div>}
                              </div>
                            ) : (
                              item.photoBefore ? <img src={item.photoBefore} alt="Before" className="w-20 h-20 object-cover rounded border"/> : <span style={{ fontSize: '12px' }} className="text-gray-400">-</span>
                            )}
                          </td>
                          <td className="py-2 px-2 align-top">
                            {isMaintEditable ? (
                              <div>
                                <input type="file" accept="image/*" onChange={(e) => handlePhotoAfterChange(item.id, e)} style={{ fontSize: '12px' }}/>
                                {item.photoAfter && <div className="relative inline-block mt-1"><img src={item.photoAfter} alt="After" className="w-20 h-20 object-cover rounded border"/></div>}
                                {isNotGood && !item.photoAfter && <p style={{ fontSize: '12px' }} className="text-red-500">Wajib upload</p>}
                              </div>
                            ) : (
                              item.photoAfter ? <img src={item.photoAfter} alt="After" className="w-20 h-20 object-cover rounded border"/> : <span style={{ fontSize: '12px' }} className="text-gray-400">-</span>
                            )}
                          </td>
                          <td className="py-2 px-2 align-top">
                            {isMaintEditable ? (
                              <textarea value={item.repairNote || ''} onChange={(e) => handleRepairNoteChange(item.id, e.target.value)} placeholder="Catatan perbaikan..." className={`w-full px-2 py-1 border rounded-md ${isNotGood && !item.repairNote ? 'border-red-300 bg-red-50' : 'border-gray-300'}`} rows={2} style={{ fontSize: '13px' }}/>
                            ) : (
                              <div style={{ fontSize: '13px' }} className="text-gray-600">{item.repairNote || '-'}</div>
                            )}
                          </td>
                          <td className="py-2 px-2 align-top text-center">
                            {canApprove ? (
                              <button
                                type="button"
                                onClick={() => handleApproveItem(item.id)}
                                className="bg-green-500 hover:bg-green-600 text-white px-3 py-1 rounded-lg text-xs font-medium transition"
                              >
                                ✅ Approve
                              </button>
                            ) : isApproved ? (
                              <span className="text-xs text-green-600 font-semibold">✅ Approved</span>
                            ) : item.status === 'Not Good' && !isApproved ? (
                              <span className="text-xs text-gray-400">Menunggu Approve</span>
                            ) : (
                              <span className="text-xs text-gray-300">-</span>
                            )}
                          </td>
                          <td className="py-2 px-2 text-center align-top">
                            <span className={`font-semibold ${hitungPersen >= 0 ? 'text-green-600' : 'text-red-600'}`} style={{ fontSize: '13px' }}>
                              {hitungPersen.toFixed(1)}%
                            </span>
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            </section>

            <section className="bg-white shadow-sm rounded-lg border border-gray-200 overflow-hidden mb-6">
              <div className="bg-blue-700 px-6 py-3">
                <h2 className="text-lg font-semibold text-white">📊 Rekap Hasil Pemeriksaan</h2>
              </div>
              <div className="p-6">
                <div className="grid grid-cols-1 md:grid-cols-4 gap-4">
                  <div className="bg-gray-50 p-4 rounded-lg text-center border border-gray-200">
                    <div className="text-sm text-gray-500">Total Item</div>
                    <div className="text-2xl font-bold text-gray-700">{formData.inspectionData.length}</div>
                  </div>
                  <div className="bg-green-50 p-4 rounded-lg text-center border border-green-200">
                    <div className="text-sm text-green-600">Good + N/A</div>
                    <div className="text-2xl font-bold text-green-700">
                      {formData.inspectionData.filter((i: any) => i.status === 'Good' || i.status === 'N/A' || (i.status === 'Not Good' && i.isApproved)).length}
                    </div>
                  </div>
                  <div className="bg-red-50 p-4 rounded-lg text-center border border-red-200">
                    <div className="text-sm text-red-600">Not Good</div>
                    <div className="text-2xl font-bold text-red-700">
                      {formData.inspectionData.filter((i: any) => i.status === 'Not Good' && !i.isApproved).length}
                    </div>
                  </div>
                  <div className="bg-blue-50 p-4 rounded-lg text-center border border-blue-200">
                    <div className="text-sm text-blue-600">Nilai Akhir</div>
                    <div className="text-2xl font-bold text-blue-700">
                      {totalScore.toFixed(1)}%
                    </div>
                  </div>
                </div>
              </div>
            </section>

            <section className="bg-white shadow-sm rounded-lg border border-gray-200 overflow-hidden mb-6">
              <div className="bg-blue-700 px-6 py-3">
                <h2 className="text-lg font-semibold text-white">📎 Upload Attachment</h2>
              </div>
              <div className="p-6">
                <div className="flex flex-col items-center justify-center p-4 border-2 border-dashed border-gray-300 rounded-lg hover:border-blue-400 transition">
                  {isQC && !isReadOnly ? (
                    <input
                      type="file"
                      accept=".pdf,.doc,.docx,.xlsx,.xls,.jpg,.jpeg,.png"
                      onChange={handleAttachmentChange}
                      disabled={isReadOnly}
                      className="w-full text-sm text-gray-500 file:mr-4 file:py-2 file:px-4 file:rounded-md file:border-0 file:text-sm file:font-medium file:bg-blue-50 file:text-blue-700 hover:file:bg-blue-100 disabled:opacity-50"
                    />
                  ) : (
                    <div className="w-full text-center">
                      <p className="text-sm text-gray-600 mb-2">📎 File Attachment</p>
                      {formData.attachment ? (
                        <div className="flex items-center justify-center gap-4">
                          <span className="text-sm text-gray-700">📄 {formData.attachmentName || 'File'}</span>
                          <button
                            type="button"
                            onClick={handleDownloadAttachment}
                            className="bg-blue-600 hover:bg-blue-700 text-white px-4 py-2 rounded-lg text-sm"
                          >
                            ⬇️ Download
                          </button>
                          {formData.attachment && (
                            <a
                              href={formData.attachment}
                              target="_blank"
                              rel="noopener noreferrer"
                              className="bg-green-600 hover:bg-green-700 text-white px-4 py-2 rounded-lg text-sm"
                            >
                              👁️ Lihat
                            </a>
                          )}
                        </div>
                      ) : (
                        <p className="text-sm text-gray-400">Tidak ada file attachment</p>
                      )}
                    </div>
                  )}
                  {formData.attachment && isQC && !isReadOnly && (
                    <div className="mt-2 text-sm text-green-600">✅ File terupload: {formData.attachmentName}</div>
                  )}
                  <p className="text-xs text-gray-400 mt-2">Upload file pendukung (PDF, DOC, XLSX, Gambar, dll)</p>
                </div>
              </div>
            </section>
          </>
        )}

        {isPemeliharaan && (
          <>
            {formData.units && formData.units.map((unit: UnitData, unitIndex: number) => {
              const totalItems = unit.maintenanceData?.length || 0;
              const goodCount = unit.maintenanceData?.filter((item: any) => item.status === 'Good' || item.status === 'N/A' || (item.status === 'Not Good' && item.isApproved)).length || 0;
              const notGoodCount = unit.maintenanceData?.filter((item: any) => item.status === 'Not Good' && !item.isApproved).length || 0;
              const score = totalItems > 0 ? Math.round((goodCount / totalItems) * 100) : 0;
              
              return (
                <div key={unitIndex} className="bg-white shadow-sm rounded-lg border border-gray-200 overflow-hidden mb-8">
                  <div className="bg-green-700 px-6 py-3 flex items-center justify-between">
                    <h2 className="text-lg font-semibold text-white">🔧 Unit #{unitIndex + 1}</h2>
                    <span className="text-sm text-white bg-green-600 px-3 py-1 rounded-full">{totalItems} Item</span>
                  </div>
                  
                  <div className="p-6 border-b border-gray-200">
                    <h3 className="text-sm font-semibold text-gray-600 mb-3">Data Umum</h3>
                    <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                      <div><label className="block text-sm font-medium text-gray-700 mb-1">No Unit</label><input type="text" value={unit.unitNumber || ''} onChange={(e) => updateUnitData(unitIndex, 'unitNumber', e.target.value)} className="w-full px-3 py-2 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500" disabled={isReadOnly || !isQC || formData.status !== 'draft'} placeholder="Masukkan No Unit" /></div>
                      <div><label className="block text-sm font-medium text-gray-700 mb-1">Kode Proyek</label><input type="text" value={unit.projectCode || ''} onChange={(e) => updateUnitData(unitIndex, 'projectCode', e.target.value)} className="w-full px-3 py-2 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500" disabled={isReadOnly || !isQC || formData.status !== 'draft'} placeholder="Masukkan Kode Proyek" /></div>
                      <div><label className="block text-sm font-medium text-gray-700 mb-1">Nama Pelanggan</label><input type="text" value={unit.customerName || ''} onChange={(e) => updateUnitData(unitIndex, 'customerName', e.target.value)} className="w-full px-3 py-2 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500" disabled={isReadOnly || !isQC || formData.status !== 'draft'} placeholder="Masukkan Nama Pelanggan" /></div>
                      <div><label className="block text-sm font-medium text-gray-700 mb-1">Lokasi Gedung</label><input type="text" value={unit.buildingLocation || ''} onChange={(e) => updateUnitData(unitIndex, 'buildingLocation', e.target.value)} className="w-full px-3 py-2 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500" disabled={isReadOnly || !isQC || formData.status !== 'draft'} placeholder="Masukkan Lokasi Gedung" /></div>
                      <div><label className="block text-sm font-medium text-gray-700 mb-1">Tipe / Merk</label><input type="text" value={unit.elevatorType || ''} onChange={(e) => updateUnitData(unitIndex, 'elevatorType', e.target.value)} className="w-full px-3 py-2 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500" disabled={isReadOnly || !isQC || formData.status !== 'draft'} placeholder="Masukkan Tipe / Merk" /></div>
                      <div><label className="block text-sm font-medium text-gray-700 mb-1">Kapasitas (kg)</label><input type="text" value={unit.capacity || ''} onChange={(e) => updateUnitData(unitIndex, 'capacity', e.target.value)} className="w-full px-3 py-2 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500" disabled={isReadOnly || !isQC || formData.status !== 'draft'} placeholder="Masukkan Kapasitas" /></div>
                      <div><label className="block text-sm font-medium text-gray-700 mb-1">Kecepatan (m/menit)</label><input type="text" value={unit.speed || ''} onChange={(e) => updateUnitData(unitIndex, 'speed', e.target.value)} className="w-full px-3 py-2 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500" disabled={isReadOnly || !isQC || formData.status !== 'draft'} placeholder="Masukkan Kecepatan" /></div>
                      <div><label className="block text-sm font-medium text-gray-700 mb-1">Tanggal Pemeriksaan</label><input type="date" value={unit.inspectionDate || ''} onChange={(e) => updateUnitData(unitIndex, 'inspectionDate', e.target.value)} className="w-full px-3 py-2 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500" disabled={isReadOnly || !isQC || formData.status !== 'draft'} /></div>
                      <div><label className="block text-sm font-medium text-gray-700 mb-1">Nama QC</label><input type="text" value={unit.qcName || ''} onChange={(e) => updateUnitData(unitIndex, 'qcName', e.target.value)} className="w-full px-3 py-2 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500" disabled={isReadOnly || !isQC || formData.status !== 'draft'} placeholder="Masukkan Nama QC" /></div>
                    </div>
                  </div>

                  <div className="p-4 overflow-x-auto">
                    <h3 className="text-sm font-semibold text-gray-600 mb-3">Data Pemeliharaan</h3>
                    <table className="w-full border-collapse print:table-fixed" style={{ fontSize: '14px' }}>
                      <thead><tr className="bg-gray-100 border-b border-gray-200">
                        <th className="text-left py-2 px-2 font-semibold text-gray-700 w-10" style={{ fontSize: '13px' }}>No</th>
                        <th className="text-left py-2 px-2 font-semibold text-gray-700" style={{ fontSize: '13px' }}>Item</th>
                        <th className="text-left py-2 px-2 font-semibold text-gray-700 w-24" style={{ fontSize: '13px' }}>Status</th>
                        <th className="text-left py-2 px-2 font-semibold text-gray-700 w-32" style={{ fontSize: '13px' }}>Temuan QC</th>
                        <th className="text-left py-2 px-2 font-semibold text-gray-700 w-24" style={{ fontSize: '13px' }}>Foto Sebelum</th>
                        <th className="text-left py-2 px-2 font-semibold text-gray-700 w-24" style={{ fontSize: '13px' }}>Foto Setelah</th>
                        <th className="text-left py-2 px-2 font-semibold text-gray-700 w-32" style={{ fontSize: '13px' }}>Catatan Perbaikan</th>
                        <th className="text-left py-2 px-2 font-semibold text-gray-700 w-20" style={{ fontSize: '13px' }}>Approve QC</th>
                      </tr></thead>
                      <tbody>
                        {unit.maintenanceData.map((item: any) => {
                          const originalItem = maintenanceItems.find(i => i.id === item.id);
                          const isNotGood = item.status === 'Not Good';
                          const isApproved = item.isApproved;
                          const isEditable = isQC && !isReadOnly && (formData.status === 'draft');
                          const isMaintEditable = isMaintenance && !isReadOnly && (formData.status === 'qc_approved' || formData.status === 'revision');
                          const canApprove = isQC && !isReadOnly && (formData.status === 'maintenance_done' || formData.status === 'revision') && isNotGood && !isApproved;

                          return (
                            <tr key={item.id} className="border-b border-gray-100 print:break-inside-avoid">
                              <td className="py-2 px-2 text-gray-600 align-top" style={{ fontSize: '13px' }}>{item.id}</td>
                              <td className="py-2 px-2 text-gray-800 align-top" style={{ fontSize: '13px' }}>{originalItem?.item || '-'}</td>
                              <td className="py-2 px-2 align-top">
                                {isEditable ? (
                                  <div className="flex flex-col gap-1">
                                    <label className="inline-flex items-center" style={{ fontSize: '12px' }}><input type="radio" name={`unit-${unitIndex}-status-${item.id}`} value="Good" checked={item.status === 'Good'} onChange={() => handleMaintenanceStatusChange(unitIndex, item.id, 'Good')} className="w-4 h-4 text-green-600"/><span className="ml-1">Good</span></label>
                                    <label className="inline-flex items-center" style={{ fontSize: '12px' }}><input type="radio" name={`unit-${unitIndex}-status-${item.id}`} value="Not Good" checked={item.status === 'Not Good'} onChange={() => handleMaintenanceStatusChange(unitIndex, item.id, 'Not Good')} className="w-4 h-4 text-red-600"/><span className="ml-1">Not Good</span></label>
                                    <label className="inline-flex items-center" style={{ fontSize: '12px' }}><input type="radio" name={`unit-${unitIndex}-status-${item.id}`} value="N/A" checked={item.status === 'N/A'} onChange={() => handleMaintenanceStatusChange(unitIndex, item.id, 'N/A')} className="w-4 h-4 text-gray-400"/><span className="ml-1">N/A</span></label>
                                  </div>
                                ) : (
                                  <span className={`inline-flex px-2 py-1 rounded-full font-semibold ${item.status === 'Good' ? 'bg-green-100 text-green-800' : item.status === 'N/A' ? 'bg-gray-300 text-gray-700' : item.status === 'Not Good' && isApproved ? 'bg-green-300 text-green-900' : item.status === 'Not Good' && !isApproved ? 'bg-red-100 text-red-800' : 'bg-gray-100 text-gray-500'}`} style={{ fontSize: '12px' }}>
                                    {item.status === 'Not Good' && isApproved ? '✅ Approved' : item.status || '-'}
                                  </span>
                                )}
                              </td>
                              <td className="py-2 px-2 align-top">
                                {isEditable ? <textarea value={item.finding} onChange={(e) => handleMaintenanceFindingChange(unitIndex, item.id, e.target.value)} disabled={item.status === 'Good' || item.status === 'N/A'} placeholder={isNotGood ? 'Wajib diisi' : 'Temuan'} className={`w-full px-2 py-1 border rounded-md ${isNotGood && !item.finding ? 'border-red-300 bg-red-50' : 'border-gray-300'} ${item.status === 'Good' ? 'bg-gray-100' : 'bg-white'}`} rows={2} style={{ fontSize: '13px' }}/>
                                : <div style={{ fontSize: '13px' }} className="text-gray-600">{item.finding || '-'}</div>}
                              </td>
                              <td className="py-2 px-2 align-top">
                                {isEditable && isNotGood ? <div><input type="file" accept="image/*" onChange={(e) => handleMaintenancePhotoBeforeChange(unitIndex, item.id, e)} style={{ fontSize: '12px' }}/>{item.photoBefore && <div className="relative inline-block mt-1"><img src={item.photoBefore} alt="Before" className="w-20 h-20 object-cover rounded border"/></div>}</div>
                                : item.photoBefore ? <img src={item.photoBefore} alt="Before" className="w-20 h-20 object-cover rounded border"/> : <span style={{ fontSize: '12px' }} className="text-gray-400">-</span>}
                              </td>
                              <td className="py-2 px-2 align-top">
                                {isMaintEditable ? <div><input type="file" accept="image/*" onChange={(e) => handleMaintenancePhotoAfterChange(unitIndex, item.id, e)} style={{ fontSize: '12px' }}/>{item.photoAfter && <div className="relative inline-block mt-1"><img src={item.photoAfter} alt="After" className="w-20 h-20 object-cover rounded border"/></div>}{isNotGood && !item.photoAfter && <p style={{ fontSize: '12px' }} className="text-red-500">Wajib upload</p>}</div>
                                : item.photoAfter ? <img src={item.photoAfter} alt="After" className="w-20 h-20 object-cover rounded border"/> : <span style={{ fontSize: '12px' }} className="text-gray-400">-</span>}
                              </td>
                              <td className="py-2 px-2 align-top">
                                {isMaintEditable ? <textarea value={item.repairNote || ''} onChange={(e) => handleMaintenanceRepairNoteChange(unitIndex, item.id, e.target.value)} placeholder="Catatan perbaikan..." className={`w-full px-2 py-1 border rounded-md ${isNotGood && !item.repairNote ? 'border-red-300 bg-red-50' : 'border-gray-300'}`} rows={2} style={{ fontSize: '13px' }}/>
                                : <div style={{ fontSize: '13px' }} className="text-gray-600">{item.repairNote || '-'}</div>}
                              </td>
                              <td className="py-2 px-2 align-top text-center">
                                {canApprove ? <button type="button" onClick={() => handleApproveMaintenanceItem(unitIndex, item.id)} className="bg-green-500 hover:bg-green-600 text-white px-3 py-1 rounded-lg text-xs font-medium transition">✅ Approve</button>
                                : isApproved ? <span className="text-xs text-green-600 font-semibold">✅ Approved</span>
                                : item.status === 'Not Good' && !isApproved ? <span className="text-xs text-gray-400">Menunggu Approve</span>
                                : <span className="text-xs text-gray-300">-</span>}
                              </td>
                            </tr>
                          );
                        })}
                      </tbody>
                    </table>
                  </div>

                  <div className="p-4 bg-gray-50 border-t border-gray-200">
                    <div className="grid grid-cols-2 md:grid-cols-4 gap-3">
                      <div className="bg-white p-3 rounded-lg text-center border border-gray-200"><div className="text-xs text-gray-500">Total Item</div><div className="text-lg font-bold text-gray-700">{totalItems}</div></div>
                      <div className="bg-green-50 p-3 rounded-lg text-center border border-green-200"><div className="text-xs text-green-600">Good + N/A</div><div className="text-lg font-bold text-green-700">{goodCount}</div></div>
                      <div className="bg-red-50 p-3 rounded-lg text-center border border-red-200"><div className="text-xs text-red-600">Not Good</div><div className="text-lg font-bold text-red-700">{notGoodCount}</div></div>
                      <div className="bg-blue-50 p-3 rounded-lg text-center border border-blue-200 col-span-2 md:col-span-1"><div className="text-xs text-blue-600">Score</div><div className="text-lg font-bold text-blue-700">{score}%</div></div>
                    </div>
                  </div>
                </div>
              );
            })}
          </>
        )}

        <div className="text-center text-[10px] text-gray-400 border-t border-gray-200 pt-3 mt-4 print:block hidden">
          <p>Dicetak: {new Date().toLocaleString('id-ID')}</p>
          <p>© 2026 PT Louserindo Megah Permai - ELEVATE QC</p>
        </div>
      </div>
    </form>
  );
}

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