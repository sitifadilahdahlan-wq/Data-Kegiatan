// =========================================================================
// KONFIGURASI UTAMA PORTAL PTW - LRT JAKARTA
// =========================================================================
const PARENT_FOLDER_ID = "1OtXREPuBAGpG00mI_Si8PYqxsjaQvR9w";

/**
 * 1. MEMBUAT MENU TOMBOL DI ATAS GOOGLE SHEETS
 */
function onOpen() {
  try {
    SpreadsheetApp.getUi()
      .createMenu('📋 Input Laporan WA')
      .addItem('➕ Buka Portal Input & Pencarian PTW', 'tampilkanModalForm')
      .addToUi();
  } catch (e) {
    Logger.log("Info: Menu UI hanya dimuat saat membuka Google Sheets langsung.");
  }
}

/**
 * 2. MENAMPILKAN POP-UP MAKSIMAL DI DALAM GOOGLE SHEETS
 */
function tampilkanModalForm() {
  var html = HtmlService.createHtmlOutput(getHtmlForm())
    .setWidth(1200)
    .setHeight(850)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  SpreadsheetApp.getUi().showModalDialog(html, 'Portal Input & Pencarian Laporan PTW');
}

/**
 * 3. MENDUKUNG AKSES SEBAGAI HALAMAN WEB APP MANDIRI (HP / Browser)
 */
function doGet() {
  return HtmlService.createHtmlOutput(getHtmlForm())
    .setTitle('Portal Input & Pencarian PTW - LRT Jakarta')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * 4. PEMROSESAN TEKS, FOLDER DRIVE, SIMPAN SHEET & FORMAT PESAN WA
 */
function prosesChatDanFoto(payload) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    var teksChat = (payload && payload.teksChat) ? payload.teksChat : "";
    var filePtw = payload ? payload.filePtw : null;
    var filesGiat = (payload && payload.filesGiat) ? payload.filesGiat : [];

    if (!teksChat || teksChat.trim() === "") {
      throw new Error("Teks chat masih kosong. Silakan paste laporan WhatsApp terlebih dahulu.");
    }
    if (!filePtw || !filePtw.data) {
      throw new Error("Foto PTW wajib diunggah (1 foto).");
    }
    if (filesGiat.length < 3) {
      throw new Error("Foto Giat wajib minimal 3 foto! (Saat ini baru dipilih " + filesGiat.length + " foto).");
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Sheet1") || ss.getSheets()[0];

    // Buat Header otomatis jika Sheet masih kosong
    if (sheet.getLastRow() === 0) {
      sheet.appendRow([
        "Tanggal Input", "Pekerja", "Area", "PT / Vendor", "No PTW",
        "Tgl Berlaku PTW", "PIC", "Jam Mulai", "Jam Selesai", "Jenis Pekerjaan",
        "Titik Area Pengerjaan", "Plan Output / Hasil", "Link Foto Giat", "Link Foto PTW", "APD"
      ]);
      sheet.getRange(1, 1, 1, 15).setFontWeight("bold").setBackground("#164e78").setFontColor("#ffffff");
    }

    var tglKirim = Utilities.formatDate(new Date(), "Asia/Jakarta", "dd/MM/yyyy HH:mm:ss");

    var hasilPekerja   = ekstrakPoin(teksChat, "1\\.\\s*Pekerja(?:an)?");
    var pekerja        = (hasilPekerja === "-" || hasilPekerja === "") ? "Vendor" : hasilPekerja;
    var area           = ekstrakPoin(teksChat, "2\\.\\s*Area");
    var pt             = ekstrakPoin(teksChat, "3\\.\\s*PT");
    var noPtw          = ekstrakPoin(teksChat, "4\\.\\s*No\\.?\\s*PTW");
    var tglBerlakuPtw  = ekstrakPoin(teksChat, "5\\.\\s*Tgl\\s*berlaku\\s*PTW");
    var pic            = ekstrakPoin(teksChat, "6\\.\\s*PIC[^:\\n]*");
    var jamMulai       = ekstrakPoin(teksChat, "7\\.\\s*Jam\\s*mulai");
    var jamSelesai     = ekstrakPoin(teksChat, "8\\.\\s*Jam\\s*selesai[^:\\n]*");
    var jenisPekerjaan = ekstrakPoin(teksChat, "9\\.\\s*Jenis\\s*pe(?:k|ng)erjaan");
    var titikArea      = ekstrakPoin(teksChat, "10\\.\\s*(?:List\\s*)?(?:titik\\s*)?(?:area\\s*)?pe(?:k|ng)erjaan");
    var hasilPlan      = ekstrakPoin(teksChat, "(?:11|12)\\.\\s*Plan\\s*output\\s*\\/\\s*Hasil\\s*pe(?:k|ng)erjaan");
    var planOutput     = (hasilPlan === "-" || hasilPlan === "") ? "Giat Sedang Berlangsung" : hasilPlan;
    var apd            = ekstrakPoin(teksChat, "(?:11|12)\\.\\s*APD");

    // --- PROSES FOLDER GOOGLE DRIVE ---
    var parentFolder = DriveApp.getFolderById(PARENT_FOLDER_ID);
    var namaFolderPtw = (noPtw !== "-" && noPtw !== "")
      ? noPtw.replace(/[\/\\:*?"<>|]/g, "-").trim()
      : "PTW_Tanpa_Nomor_" + Utilities.formatDate(new Date(), "Asia/Jakarta", "ddMMyyyy_HHmm");

    var ptwFolder;
    var folderIterator = parentFolder.getFoldersByName(namaFolderPtw);
    if (folderIterator.hasNext()) {
      ptwFolder = folderIterator.next();
    } else {
      ptwFolder = parentFolder.createFolder(namaFolderPtw);
      try {
        ptwFolder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      } catch (e) {}
    }

    // 1. Simpan Foto PTW
    var extPtw = dapatkanEkstensi(filePtw.name);
    var blobPtw = Utilities.newBlob(Utilities.base64Decode(filePtw.data), filePtw.mimeType, "FOTO_PTW_" + namaFolderPtw + extPtw);
    var savedFilePtw = ptwFolder.createFile(blobPtw);
    var linkFotoPtw = savedFilePtw.getUrl();

    // 2. Simpan Foto Giat
    var listLinkGiat = [];
    for (var i = 0; i < filesGiat.length; i++) {
      var fg = filesGiat[i];
      var extGiat = dapatkanEkstensi(fg.name);
      var namaFileGiat = "FOTO_GIAT_" + (i + 1) + "_" + namaFolderPtw + extGiat;
      var blobGiat = Utilities.newBlob(Utilities.base64Decode(fg.data), fg.mimeType, namaFileGiat);
      var savedFileGiat = ptwFolder.createFile(blobGiat);
      listLinkGiat.push(savedFileGiat.getUrl());
    }

    var linkFotoGiatGabung = listLinkGiat.join("\n");

    // Simpan ke baris Sheet1
    sheet.appendRow([
      tglKirim, pekerja, area, pt, noPtw, tglBerlakuPtw, pic,
      jamMulai, jamSelesai, jenisPekerjaan, titikArea, planOutput,
      linkFotoGiatGabung, linkFotoPtw, apd
    ]);

    // Format Link Gambar QR Code PTW
    var qrCodeUrl = (noPtw !== "-" && noPtw !== "")
      ? "https://api.qrserver.com/v1/create-qr-code/?size=300x300&data=" + encodeURIComponent(noPtw)
      : "";

    // Susun Pesan WhatsApp Lengkap termasuk Link Folder & QR Code PTW
    var pesanWaLengkap = teksChat.trim() + "\n\n📁 *Folder Dokumen GDrive*:\n" + ptwFolder.getUrl();
    if (qrCodeUrl !== "") {
      pesanWaLengkap += "\n\n🏷️ *Link QR Code PTW Kontraktor*:\n" + qrCodeUrl;
    }

    return {
      status: "success",
      noPtw: noPtw,
      namaFolder: namaFolderPtw,
      folderUrl: ptwFolder.getUrl(),
      jmlFotoGiat: filesGiat.length,
      pesanWaLengkap: pesanWaLengkap
    };
  } finally {
    lock.releaseLock();
  }
}

/**
 * 5. FUNGSI PENCARIAN DATA
 */
function cariDataPtw(keyword) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Sheet1") || ss.getSheets()[0];
  var data = sheet.getDataRange().getDisplayValues();

  if (data.length <= 1) return [];

  var q = (keyword || "").toString().toLowerCase().trim();
  var hasil = [];

  for (var i = data.length - 1; i >= 1; i--) {
    var row = data[i];
    if (row.join("").trim() === "") continue;

    var gabungTeksBaris = row.join(" ").toLowerCase();
    if (q === "" || gabungTeksBaris.indexOf(q) !== -1) {
      var rawGiat = row[12] || "";
      var rawPtw  = row[13] || "";

      hasil.push({
        barisKe: i + 1,
        tglKirim: row[0] || "-",
        pekerja: row[1] || "-",
        area: row[2] || "-",
        pt: row[3] || "-",
        noPtw: row[4] || "-",
        tglBerlakuPtw: row[5] || "-",
        pic: row[6] || "-",
        jamMulai: row[7] || "-",
        jamSelesai: row[8] || "-",
        jenisPekerjaan: row[9] || "-",
        titikArea: row[10] || "-",
        planOutput: row[11] || "-",
        fotoGiat: parseDriveLinks(rawGiat),
        fotoPtw: parseDriveLinks(rawPtw),
        apd: row[14] || "-"
      });

      if (hasil.length >= 35) break;
    }
  }

  return hasil;
}

function parseDriveLinks(teksUrl) {
  if (!teksUrl) return [];
  var baris = teksUrl.split(/\r?\n/);
  var list = [];
  for (var i = 0; i < baris.length; i++) {
    var url = baris[i].trim();
    if (!url) continue;
    var fileId = ekstrakDriveFileId(url);
    if (fileId) {
      list.push({
        fileId: fileId,
        viewUrl: "https://drive.google.com/file/d/" + fileId + "/view",
        thumbUrl: "https://drive.google.com/thumbnail?id=" + fileId + "&sz=w600"
      });
    } else if (url.indexOf("http") === 0) {
      list.push({ fileId: "", viewUrl: url, thumbUrl: url });
    }
  }
  return list;
}

function ekstrakDriveFileId(url) {
  var match1 = url.match(/\/d\/([a-zA-Z0-9_-]{15,})/);
  if (match1 && match1[1]) return match1[1];
  var match2 = url.match(/[?&]id=([a-zA-Z0-9_-]{15,})/);
  if (match2 && match2[1]) return match2[1];
  return null;
}

function ekstrakPoin(teks, polaLabel) {
  var teksBersih = teks.replace(/[\u200B-\u200D\uFEFF]/g, "").replace(/\u00A0/g, " ");
  var regex = new RegExp(polaLabel + "\\s*:\\s*([\\s\\S]*?)(?=\\n\\s*\\d+\\.\\s*|\\n\\s*Demikian|$)", "i");
  var match = teksBersih.match(regex);
  if (match && match[1]) {
    var hasil = match[1].replace(/\r?\n/g, " ").replace(/\s+/g, " ").trim();
    return hasil !== "" ? hasil : "-";
  }
  return "-";
}

function dapatkanEkstensi(namaFile) {
  if (!namaFile) return ".jpg";
  var idx = namaFile.lastIndexOf(".");
  return idx !== -1 ? namaFile.substring(idx) : ".jpg";
}

/**
 * 7. TAMPILAN ANTARMUKA WEB APP DENGAN FITUR FULLSCREEN
 */
function getHtmlForm() {
  return `
  <!DOCTYPE html>
  <html>
  <head>
    <meta charset="utf-8">
    <script src="https://unpkg.com/html5-qrcode@2.3.8/html5-qrcode.min.js" type="text/javascript"></script>
    <style>
      * { box-sizing: border-box; }
      html, body {
        height: 100%;
        margin: 0;
        padding: 0;
      }
      body {
        font-family: "Book Antiqua", Palatino, "Palatino Linotype", serif;
        background-color: #f4f7f9;
        color: #1e293b;
        display: flex;
        flex-direction: column;
      }

      /* CONTAINER FULLSCREEN FLUID */
      .app-container {
        width: 100%;
        min-height: 100vh;
        background: #ffffff;
        display: flex;
        flex-direction: column;
      }
      .app-header {
        background: #164e78;
        color: #ffffff;
        padding: 10px 16px;
        display: flex;
        justify-content: space-between;
        align-items: center;
        box-shadow: 0 2px 6px rgba(0,0,0,0.15);
      }
      .app-title {
        font-size: 16px;
        font-weight: bold;
        letter-spacing: 0.5px;
        display: flex;
        align-items: center;
        gap: 8px;
      }
      .header-actions {
        display: flex;
        gap: 8px;
        align-items: center;
      }
      .btn-fullscreen {
        background: rgba(255, 255, 255, 0.18);
        color: #ffffff;
        border: 1px solid rgba(255, 255, 255, 0.4);
        padding: 6px 12px;
        font-size: 12px;
        border-radius: 6px;
        cursor: pointer;
        display: inline-flex;
        align-items: center;
        gap: 5px;
        transition: 0.2s;
        font-family: "Book Antiqua", Palatino, serif;
      }
      .btn-fullscreen:hover {
        background: rgba(255, 255, 255, 0.3);
      }

      .app-body {
        padding: 16px;
        flex: 1;
        max-width: 1200px;
        width: 100%;
        margin: 0 auto;
      }

      /* TABS */
      .tabs {
        display: flex;
        gap: 8px;
        margin-bottom: 16px;
        border-bottom: 2px solid #e2e8f0;
        padding-bottom: 10px;
      }
      .tab-btn {
        font-family: "Book Antiqua", Palatino, serif;
        padding: 9px 16px;
        border-radius: 6px;
        border: 1px solid #cbd5e1;
        background: #f8fafc;
        color: #334155;
        font-size: 14px;
        font-weight: bold;
        cursor: pointer;
        transition: 0.2s;
        flex: 1;
        text-align: center;
      }
      .tab-btn.active {
        background: #164e78;
        color: #ffffff;
        border-color: #164e78;
      }
      .tab-content { display: none; }
      .tab-content.active { display: block; }

      h3 { margin-top: 0; margin-bottom: 4px; color: #164e78; font-size: 18px; }
      p.desc { margin-top: 0; margin-bottom: 12px; font-size: 13px; color: #475569; }
      
      textarea {
        width: 100%;
        height: 200px;
        padding: 12px;
        border: 1px solid #cbd5e1;
        border-radius: 6px;
        font-family: "Book Antiqua", Palatino, serif;
        font-size: 13px;
        resize: vertical;
        line-height: 1.45;
      }
      textarea:focus, input[type="text"]:focus {
        outline: none;
        border-color: #164e78;
        box-shadow: 0 0 0 2px rgba(22,78,120,0.15);
      }

      .upload-grid {
        display: grid;
        grid-template-columns: 1fr 1fr;
        gap: 12px;
        margin-top: 12px;
      }
      @media (max-width: 600px) {
        .upload-grid { grid-template-columns: 1fr; }
        .app-body { padding: 10px; }
      }
      .upload-box {
        background: #f8fafc;
        border: 1px dashed #94a3b8;
        border-radius: 8px;
        padding: 10px 12px;
      }
      .upload-box label { display: block; font-weight: bold; font-size: 13px; color: #164e78; margin-bottom: 4px; }
      .upload-box small { display: block; font-size: 11px; color: #64748b; margin-bottom: 6px; }
      input[type="file"] { font-family: "Book Antiqua", Palatino, serif; font-size: 12px; width: 100%; }
      .file-badge { margin-top: 6px; font-size: 12px; font-weight: bold; }
      .ok { color: #166534; }
      .warn { color: #b91c1c; }

      .btn-row { display: flex; gap: 10px; margin-top: 14px; flex-wrap: wrap; }
      button {
        font-family: "Book Antiqua", Palatino, serif;
        padding: 10px 14px;
        border: none;
        border-radius: 6px;
        font-size: 14px;
        font-weight: bold;
        cursor: pointer;
        transition: 0.2s;
      }
      .btn-primary { background-color: #164e78; color: white; flex: 2; }
      .btn-primary:hover { background-color: #0f3857; }
      .btn-primary:disabled { background-color: #94a3b8; cursor: not-allowed; }
      .btn-secondary { background-color: #e2e8f0; color: #334155; flex: 1; }
      .btn-secondary:hover { background-color: #cbd5e1; }
      .btn-success { background-color: #15803d; color: white; }
      .btn-success:hover { background-color: #166534; }
      .btn-danger { background-color: #fee2e2; color: #991b1b; border: 1px solid #fca5a5; }

      .btn-icon {
        flex: 0 0 42px;
        height: 40px;
        padding: 0;
        display: inline-flex;
        align-items: center;
        justify-content: center;
        font-size: 17px;
        border-radius: 6px;
      }
      .btn-toggle-active { background-color: #164e78 !important; color: #ffffff !important; }

      #notif {
        margin-top: 12px;
        padding: 12px;
        border-radius: 6px;
        font-size: 13px;
        display: none;
        line-height: 1.45;
        white-space: pre-line;
      }
      .sukses { background-color: #dcfce7; color: #166534; border: 1px solid #86efac; }
      .gagal { background-color: #fee2e2; color: #991b1b; border: 1px solid #fca5a5; }

      .search-bar { display: flex; gap: 6px; margin-bottom: 12px; align-items: center; }
      .search-bar input[type="text"] {
        flex: 1;
        height: 40px;
        padding: 8px 12px;
        border: 1px solid #cbd5e1;
        border-radius: 6px;
        font-family: "Book Antiqua", Palatino, serif;
        font-size: 14px;
      }

      /* QR PANEL */
      .qr-panel {
        display: none;
        background: #f8fafc;
        border: 2px dashed #164e78;
        border-radius: 8px;
        padding: 14px;
        margin-bottom: 14px;
        text-align: center;
      }
      #reader { width: 100%; max-width: 320px; margin: 0 auto 10px auto; border-radius: 8px; overflow: hidden; }

      /* RESULT CARD */
      .result-card {
        background: #f8fafc;
        border: 1px solid #cbd5e1;
        border-left: 5px solid #164e78;
        border-radius: 8px;
        padding: 14px;
        margin-bottom: 14px;
      }
      .result-header {
        display: flex;
        justify-content: space-between;
        align-items: center;
        border-bottom: 1px solid #e2e8f0;
        padding-bottom: 8px;
        margin-bottom: 10px;
        flex-wrap: wrap;
        gap: 6px;
      }
      .timestamp { font-size: 12px; color: #64748b; }
      
      .data-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 6px 16px; font-size: 13px; margin-bottom: 12px; }
      @media (max-width: 600px) {
        .data-grid { grid-template-columns: 1fr; }
        .span-2 { grid-column: span 1 !important; }
      }
      .span-2 { grid-column: span 2; }
      .data-item b { color: #0f3857; }

      .btn-header-wa {
        font-size: 12px;
        padding: 5px 12px;
        border-radius: 5px;
        display: inline-flex;
        align-items: center;
        gap: 4px;
      }

      .photo-section-title { font-size: 13px; font-weight: bold; color: #164e78; margin: 10px 0 6px 0; }
      .photo-gallery { display: grid; grid-template-columns: repeat(auto-fill, minmax(130px, 1fr)); gap: 10px; }
      .photo-card {
        background: #ffffff;
        border: 1px solid #cbd5e1;
        border-radius: 6px;
        overflow: hidden;
        text-align: center;
      }
      .photo-card img { width: 100%; height: 110px; object-fit: cover; display: block; background: #e2e8f0; }
      .photo-card a { display: block; padding: 5px; font-size: 11px; color: #164e78; text-decoration: none; font-weight: bold; }
      .empty-state { text-align: center; padding: 30px; color: #64748b; font-size: 14px; }
    </style>
  </head>
  <body>
    <div class="app-container" id="appContainer">
      <!-- HEADER APLIKASI DENGAN TOMBOL FULLSCREEN -->
      <div class="app-header">
        <div class="app-title">🚇 PORTAL PTW & KEGIATAN KONTRAKTOR</div>
        <div class="header-actions">
          <button type="button" class="btn-fullscreen" id="btnToggleFs" onclick="toggleFullScreen()">
            ⛶ Layar Penuh
          </button>
        </div>
      </div>

      <div class="app-body">
        <div class="tabs">
          <button class="tab-btn active" id="tabBtnInput" onclick="bukaTab('input')">➕ Laporan Kegiatan</button>
          <button class="tab-btn" id="tabBtnCari" onclick="bukaTab('cari')">🔍 Pencarian & Kirim WA</button>
        </div>

        <!-- ==================== TAB 1: INPUT ==================== -->
        <div id="tabInput" class="tab-content active">
          <h3>📋 FORMULIR KEGIATAN & KIRIM KE WHATSAPP</h3>
          <p class="desc">Paste teks laporan WA dan unggah foto. Data otomatis tersimpan ke Sheets, GDrive, dan <b>langsung membuka WhatsApp</b>.</p>

          <textarea id="inputChat" placeholder="Izin Info Pekerjaan Vendor/Kontraktor LRT Jakarta&#10;&#10;1. Pekerjaan : Vendor&#10;2. Area : &#10;3. PT : ..."></textarea>

          <div class="upload-grid">
            <div class="upload-box">
              <label>📄 Upload Foto PTW</label>
              <small>Wajib 1 foto lembar dokumen PTW</small>
              <input type="file" id="fotoPtw" accept="image/*" onchange="cekJumlahFile()">
              <div id="statusPtw" class="file-badge warn">Belum ada foto dipilih</div>
            </div>
            <div class="upload-box">
              <label>👷 Upload Foto Giat</label>
              <small>Wajib minimal 3 foto kegiatan</small>
              <input type="file" id="fotoGiat" accept="image/*" multiple onchange="cekJumlahFile()">
              <div id="statusGiat" class="file-badge warn">0 foto dipilih (Minimal 3)</div>
            </div>
          </div>

          <div class="btn-row">
            <button class="btn-secondary" onclick="isiContoh()">Isi Contoh</button>
            <button class="btn-secondary" onclick="bersihkan()">Bersihkan</button>
            <button class="btn-primary" id="btnSimpan" onclick="kirimData()">🚀 Simpan & Buka WhatsApp</button>
          </div>

          <div id="notif"></div>
        </div>

        <!-- ==================== TAB 2: PENCARIAN & KIRIM ULANG ==================== -->
        <div id="tabCari" class="tab-content">
          <h3>DATA KONTRAKTOR & PENGIRIMAN WHATSAPP</h3>
          <p class="desc">Cari data dengan No. PTW / Area / Scan QR. Klik tombol <b>📲 Kirim WA</b> di setiap kartu data untuk mengirimkan laporan.</p>

          <div class="search-bar">
            <input type="text" id="inputKeyword" placeholder="Cari No. PTW / Vendor / Area..." onkeydown="if(event.key==='Enter') jalankanPencarian()">
            <button class="btn-primary btn-icon" id="btnScanQR" onclick="toggleScannerPanel()" title="Scan QR Code">📷</button>
            <button class="btn-primary btn-icon" id="btnCari" onclick="jalankanPencarian()" title="Cari Data">🔍</button>
            <button class="btn-secondary btn-icon" id="btnToggleSemua" onclick="toggleTampilkanSemua()" title="Tampilkan Semua">👁️</button>
          </div>

          <div id="qrPanel" class="qr-panel">
            <div style="font-weight:bold; color:#164e78; margin-bottom:8px;">Arahkan Kamera ke QR Code PTW</div>
            <div id="reader"></div>
            <div id="qrStatusMsg" style="font-size:12px; margin-bottom:8px;">Menyiapkan kamera...</div>
            <button type="button" class="btn-danger" style="padding:6px 12px; font-size:12px;" onclick="tutupScanner()">Tutup Scanner</button>
          </div>

          <div id="wadahHasil">
            <div class="empty-state">Silakan cari data atau klik <b>👁️ Tampilkan Semua Data</b> untuk melihat laporan kontraktor.</div>
          </div>
        </div>
      </div>
    </div>

    <script>
      var sedangTampilSemua = false;
      var html5QrCode = null;
      var listHasilPencarianGlobal = [];

      // =========================================================================
      // FITUR FULLSCREEN (LAYAR PENUH)
      // =========================================================================
      function toggleFullScreen() {
        var btn = document.getElementById("btnToggleFs");
        if (!document.fullscreenElement && !document.webkitFullscreenElement) {
          var el = document.documentElement;
          if (el.requestFullscreen) {
            el.requestFullscreen();
          } else if (el.webkitRequestFullscreen) {
            el.webkitRequestFullscreen();
          }
        } else {
          if (document.exitFullscreen) {
            document.exitFullscreen();
          } else if (document.webkitExitFullscreen) {
            document.webkitExitFullscreen();
          }
        }
      }

      function updateFsBtnText() {
        var btn = document.getElementById("btnToggleFs");
        if (document.fullscreenElement || document.webkitFullscreenElement) {
          btn.innerHTML = "🗗 Keluar Layar Penuh";
        } else {
          btn.innerHTML = "⛶ Layar Penuh";
        }
      }

      document.addEventListener("fullscreenchange", updateFsBtnText);
      document.addEventListener("webkitfullscreenchange", updateFsBtnText);

      // =========================================================================
      // TABS & FORM
      // =========================================================================
      function bukaTab(namaTab) {
        document.getElementById("tabBtnInput").classList.remove("active");
        document.getElementById("tabBtnCari").classList.remove("active");
        document.getElementById("tabInput").classList.remove("active");
        document.getElementById("tabCari").classList.remove("active");

        if (namaTab === "input") {
          document.getElementById("tabBtnInput").classList.add("active");
          document.getElementById("tabInput").classList.add("active");
          tutupScanner();
        } else {
          document.getElementById("tabBtnCari").classList.add("active");
          document.getElementById("tabCari").classList.add("active");
        }
      }

      function cekJumlahFile() {
        var fPtw = document.getElementById("fotoPtw").files;
        var fGiat = document.getElementById("fotoGiat").files;
        var stPtw = document.getElementById("statusPtw");
        stPtw.className = fPtw.length === 1 ? "file-badge ok" : "file-badge warn";
        stPtw.innerText = fPtw.length === 1 ? "✔ 1 Foto PTW siap" : "Belum ada foto dipilih";

        var stGiat = document.getElementById("statusGiat");
        stGiat.className = fGiat.length >= 3 ? "file-badge ok" : "file-badge warn";
        stGiat.innerText = fGiat.length >= 3 ? "✔ " + fGiat.length + " Foto Giat siap" : "⚠ Baru " + fGiat.length + " foto (Wajib min. 3)";
      }

      function bacaDanKompresGambar(file) {
        return new Promise(function(resolve, reject) {
          var reader = new FileReader();
          reader.onload = function(e) {
            var img = new Image();
            img.onload = function() {
              var canvas = document.createElement("canvas");
              var maxDim = 1600;
              var w = img.width, h = img.height;
              if (w > maxDim || h > maxDim) {
                if (w > h) { h = Math.round(h * (maxDim / w)); w = maxDim; }
                else { w = Math.round(w * (maxDim / h)); h = maxDim; }
              }
              canvas.width = w; canvas.height = h;
              var ctx = canvas.getContext("2d");
              ctx.drawImage(img, 0, 0, w, h);
              resolve({
                name: file.name.replace(/\\.[^/.]+$/, "") + ".jpg",
                mimeType: "image/jpeg",
                data: canvas.toDataURL("image/jpeg", 0.82).split(",")[1]
              });
            };
            img.onerror = reject;
            img.src = e.target.result;
          };
          reader.onerror = reject;
          reader.readAsDataURL(file);
        });
      }

      // =========================================================================
      // SIMPAN DATA FORM INPUT & BUKA WHATSAPP
      // =========================================================================
      async function kirimData() {
        var teks = document.getElementById("inputChat").value;
        var fPtw = document.getElementById("fotoPtw").files;
        var fGiat = document.getElementById("fotoGiat").files;
        var btn = document.getElementById("btnSimpan");
        var notif = document.getElementById("notif");

        if (!teks.trim()) { tampilNotif("Silakan paste teks laporan WhatsApp!", false); return; }
        if (fPtw.length === 0) { tampilNotif("Wajib memilih 1 Foto PTW!", false); return; }
        if (fGiat.length < 3) { tampilNotif("Foto Giat kurang! Wajib minimal 3 foto kegiatan.", false); return; }

        try {
          btn.disabled = true;
          btn.innerText = "⏳ Menyimpan ke Google Drive & Sheets...";
          notif.style.display = "none";

          var ptwObj = await bacaDanKompresGambar(fPtw[0]);
          var giatArr = [];
          for (var i = 0; i < fGiat.length; i++) giatArr.push(await bacaDanKompresGambar(fGiat[i]));

          var payload = {
            teksChat: teks,
            filePtw: ptwObj,
            filesGiat: giatArr
          };

          google.script.run
            .withSuccessHandler(function(res) {
              btn.disabled = false;
              btn.innerText = "🚀 Simpan & Buka WhatsApp";

              var pesanSukses = "✅ Berhasil disimpan ke Sheets & Google Drive!\\n📁 Folder: " + res.namaFolder + "\\n📲 Mengalihkan ke WhatsApp...";
              tampilNotif(pesanSukses, true);

              bukaWhatsAppUrl(res.pesanWaLengkap);
              bersihkanFormSetelahSimpan();
            })
            .withFailureHandler(function(err) {
              btn.disabled = false;
              btn.innerText = "🚀 Simpan & Buka WhatsApp";
              tampilNotif("❌ Gagal menyimpan: " + err.message, false);
            })
            .prosesChatDanFoto(payload);

        } catch (err) {
          btn.disabled = false;
          btn.innerText = "🚀 Simpan & Buka WhatsApp";
          tampilNotif("❌ Gagal memproses gambar: " + err.message, false);
        }
      }

      function bukaWhatsAppUrl(teksPesan) {
        var urlWa = "https://api.whatsapp.com/send?text=" + encodeURIComponent(teksPesan);
        window.open(urlWa, "_blank");
      }

      // =========================================================================
      // PENCARIAN & KIRIM WA PER-BARIS DATA
      // =========================================================================
      function toggleTampilkanSemua() {
        tutupScanner();
        var wadah = document.getElementById("wadahHasil");
        var btnToggle = document.getElementById("btnToggleSemua");
        if (sedangTampilSemua) {
          sedangTampilSemua = false;
          btnToggle.innerText = "👁️";
          btnToggle.classList.remove("btn-toggle-active");
          wadah.innerHTML = '<div class="empty-state">Data disembunyikan. Silakan ketik kata kunci atau klik <b>👁️</b>.</div>';
        } else {
          document.getElementById("inputKeyword").value = "";
          sedangTampilSemua = true;
          btnToggle.innerText = "🙈";
          btnToggle.classList.add("btn-toggle-active");
          muatDataPencarian("");
        }
      }

      function jalankanPencarian() {
        tutupScanner();
        var keyword = document.getElementById("inputKeyword").value;
        sedangTampilSemua = (keyword.trim() === "");
        muatDataPencarian(keyword);
      }

      function muatDataPencarian(keyword) {
        var wadah = document.getElementById("wadahHasil");
        var btnCari = document.getElementById("btnCari");
        btnCari.disabled = true; btnCari.innerText = "⏳";
        wadah.innerHTML = '<div class="empty-state">⏳ Mengambil data dari Google Sheets...</div>';

        google.script.run
          .withSuccessHandler(function(listData) {
            btnCari.disabled = false; btnCari.innerText = "🔍";
            listHasilPencarianGlobal = listData || [];
            renderHasilPencarian(listHasilPencarianGlobal);
          })
          .withFailureHandler(function(err) {
            btnCari.disabled = false; btnCari.innerText = "🔍";
            wadah.innerHTML = '<div class="empty-state" style="color:#b91c1c;">❌ Gagal memuat data: ' + escapeHtml(err.message) + '</div>';
          })
          .cariDataPtw(keyword);
      }

      function siapkanKirimItem(idx) {
        var item = listHasilPencarianGlobal[idx];
        if (!item) return;

        var teksFormat = 
          "Izin Info Pekerjaan Vendor/Kontraktor LRT Jakarta\\n\\n" +
          "1. Pekerjaan : " + item.pekerja + "\\n" +
          "2. Area : " + item.area + "\\n" +
          "3. PT : " + item.pt + "\\n" +
          "4. No PTW : " + item.noPtw + "\\n" +
          "5. Tgl berlaku PTW : " + item.tglBerlakuPtw + "\\n" +
          "6. PIC : " + item.pic + "\\n" +
          "7. Jam mulai : " + item.jamMulai + "\\n" +
          "8. Jam selesai : " + item.jamSelesai + "\\n" +
          "9. Jenis pekerjaan : " + item.jenisPekerjaan + "\\n" +
          "10. List titik area pekerjaan : " + item.titikArea + "\\n" +
          "11. APD : " + item.apd + "\\n" +
          "12. Plan output/Hasil pekerjaan : " + item.planOutput;

        if (item.noPtw && item.noPtw !== "-") {
          var qrUrl = "https://api.qrserver.com/v1/create-qr-code/?size=300x300&data=" + encodeURIComponent(item.noPtw);
          teksFormat += "\\n\\n🏷️ *Link QR Code PTW Kontraktor*:\\n" + qrUrl;
        }

        if (item.fotoPtw && item.fotoPtw.length > 0) {
          teksFormat += "\\n\\n📄 *Foto Dokumen PTW*:\\n" + item.fotoPtw[0].viewUrl;
        }

        bukaWhatsAppUrl(teksFormat);
      }

      function renderHasilPencarian(list) {
        var wadah = document.getElementById("wadahHasil");
        if (!list || list.length === 0) {
          wadah.innerHTML = '<div class="empty-state">Tidak ditemukan data kegiatan.</div>';
          return;
        }

        var html = "";
        for (var i = 0; i < list.length; i++) {
          var item = list[i];
          html += '<div class="result-card">';
          html += '  <div class="result-header">';
          html += '    <span class="timestamp">🕒 ' + escapeHtml(item.tglKirim) + ' (Baris #' + item.barisKe + ')</span>';
          html += '    <button type="button" class="btn-success btn-header-wa" onclick="siapkanKirimItem(' + i + ')">📲 Kirim WA</button>';
          html += '  </div>';

          html += '  <div class="data-grid">';
          html += '    <div class="data-item"><b>No. PTW:</b> <span style="color:#164e78; font-weight:bold;">' + escapeHtml(item.noPtw) + '</span></div>';
          html += '    <div class="data-item"><b>Pekerja:</b> ' + escapeHtml(item.pekerja) + '</div>';
          html += '    <div class="data-item"><b>PT / Vendor:</b> ' + escapeHtml(item.pt) + '</div>';
          html += '    <div class="data-item"><b>Area:</b> ' + escapeHtml(item.area) + '</div>';
          html += '    <div class="data-item"><b>Tgl Berlaku PTW:</b> ' + escapeHtml(item.tglBerlakuPtw) + '</div>';
          html += '    <div class="data-item"><b>PIC:</b> ' + escapeHtml(item.pic) + '</div>';
          html += '    <div class="data-item span-2"><b>Waktu:</b> ' + escapeHtml(item.jamMulai) + ' s/d ' + escapeHtml(item.jamSelesai) + '</div>';
          html += '    <div class="data-item span-2"><b>Jenis Pekerjaan:</b> ' + escapeHtml(item.jenisPekerjaan) + '</div>';
          html += '    <div class="data-item span-2"><b>Titik Area Pengerjaan:</b> ' + escapeHtml(item.titikArea) + '</div>';
          html += '    <div class="data-item"><b>Hasil / Plan Output:</b> ' + escapeHtml(item.planOutput) + '</div>';
          html += '    <div class="data-item"><b>APD:</b> ' + escapeHtml(item.apd) + '</div>';
          html += '  </div>';

          html += '  <div class="photo-section-title">📸 Dokumentasi Foto:</div>';
          html += '  <div class="photo-gallery">';
          if (item.fotoPtw && item.fotoPtw.length > 0) {
            html += '<div class="photo-card"><a href="' + item.fotoPtw[0].viewUrl + '" target="_blank"><img src="' + item.fotoPtw[0].thumbUrl + '">📄 Foto PTW</a></div>';
          }
          if (item.fotoGiat && item.fotoGiat.length > 0) {
            for (var g = 0; g < item.fotoGiat.length; g++) {
              html += '<div class="photo-card"><a href="' + item.fotoGiat[g].viewUrl + '" target="_blank"><img src="' + item.fotoGiat[g].thumbUrl + '">👷 Giat ' + (g+1) + '</a></div>';
            }
          }
          html += '  </div>';

          html += '</div>';
        }
        wadah.innerHTML = html;
      }

      // =========================================================================
      // QR SCANNER
      // =========================================================================
      function toggleScannerPanel() {
        var panel = document.getElementById("qrPanel");
        if (panel.style.display === "block") tutupScanner();
        else mulaiScanner();
      }

      function mulaiScanner() {
        var panel = document.getElementById("qrPanel");
        panel.style.display = "block";
        if (!html5QrCode) html5QrCode = new Html5Qrcode("reader");

        html5QrCode.start(
          { facingMode: "environment" },
          { fps: 10, qrbox: { width: 230, height: 230 } },
          function(decodedText) {
            document.getElementById("inputKeyword").value = decodedText.trim();
            tutupScanner();
            jalankanPencarian();
          },
          function(err) {}
        ).catch(function(err) {
          document.getElementById("qrStatusMsg").innerText = "Gagal membuka kamera: " + err;
        });
      }

      function tutupScanner() {
        var panel = document.getElementById("qrPanel");
        if (html5QrCode) {
          html5QrCode.stop().then(function() { panel.style.display = "none"; }).catch(function() { panel.style.display = "none"; });
        } else {
          panel.style.display = "none";
        }
      }

      function escapeHtml(str) {
        if (!str) return "";
        return String(str).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
      }

      function tampilNotif(pesan, isSukses) {
        var notif = document.getElementById("notif");
        notif.className = isSukses ? "sukses" : "gagal";
        notif.innerText = pesan; notif.style.display = "block";
      }

      function bersihkanFormSetelahSimpan() {
        document.getElementById("inputChat").value = "";
        document.getElementById("fotoPtw").value = "";
        document.getElementById("fotoGiat").value = "";
        cekJumlahFile();
      }

      function bersihkan() {
        bersihkanFormSetelahSimpan();
        document.getElementById("notif").style.display = "none";
      }

      function isiContoh() {
        document.getElementById("inputChat").value =
"Izin Info Pekerjaan Vendor/Kontraktor LRT Jakarta\\n\\n" +
"1. Pekerjaan : Vendor\\n" +
"2. Area : Stasiun Pegangsaan Dua\\n" +
"3. PT : PT Solusi Fasilitas Mandiri\\n" +
"4. No PTW : PTW/LRTJ/X/2026/088\\n" +
"5. Tgl berlaku PTW : 20/10/2026\\n" +
"6. PIC : Dimas (0812999888)\\n" +
"7. Jam mulai : 23:00 WIB\\n" +
"8. Jam selesai : 04:00 WIB\\n" +
"9. Jenis pekerjaan : Maintenance Eskalator\\n" +
"10. List titik area pekerjaan : Lantai Concourse Stasiun\\n" +
"11. APD : Lengkap sesuai standar K3\\n" +
"12. Plan output/Hasil pekerjaan : Sedang Berlangsung\\n\\n" +
"Demikian disampaikan, terima kasih.";
      }
    </script>
  </body>
  </html>
  `;
}
