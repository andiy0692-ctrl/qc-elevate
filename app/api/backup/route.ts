import { NextResponse } from 'next/server';

// 🔥 GANTI DENGAN URL APPS SCRIPT BARU ANDA
const APPS_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbzyDFxCHIBJE2_F5AsbDoqNn7XrzN9QceZCEfO6t3vhJkaLo-qlT7NtvX_vzie_MJPT0g/exec';

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
      message: 'Backup ke Google Drive berhasil!'
    });
  } catch (error: any) {
    console.error('❌ Backup error:', error);
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
    console.error('❌ Restore error:', error);
    return NextResponse.json({
      success: false,
      error: error.message || 'Unknown error'
    }, { status: 500 });
  }
}