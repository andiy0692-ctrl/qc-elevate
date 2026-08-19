import { NextResponse } from 'next/server';

// 🔥 GANTI DENGAN URL YANG BENAR (PAKAI YANG DI DEVICE 1 BERHASIL)
const APPS_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbxvs83cuXXZ5le-SvVRd36hg2v37QTz1UGPWplSL0yvmR1s9eey-AVvYj_K-sH3jiQEPQ/exec';

export async function POST(request: Request) {
  try {
    const { data } = await request.json();
    const response = await fetch(APPS_SCRIPT_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(data),
    });
    const result = await response.json();
    return NextResponse.json({
      success: true,
      fileId: result.fileId,
      fileName: result.fileName,
    });
  } catch (error: any) {
    return NextResponse.json({
      success: false,
      error: error.message || 'Unknown error'
    }, { status: 500 });
  }
}

export async function GET() {
  try {
    const response = await fetch(APPS_SCRIPT_URL);
    const result = await response.json();
    return NextResponse.json({
      success: true,
      data: result.data || null
    });
  } catch (error: any) {
    return NextResponse.json({
      success: false,
      error: error.message || 'Unknown error'
    }, { status: 500 });
  }
}