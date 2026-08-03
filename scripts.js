import { initializeApp } from "https://www.gstatic.com/firebasejs/10.12.0/firebase-app.js";
import { getFirestore, collection, addDoc, doc, getDoc } from "https://www.gstatic.com/firebasejs/10.12.0/firebase-firestore.js";

const firebaseConfig = {
    apiKey: "AIzaSyA47cTihvHBnqYcn4HDlOKJf88O7MmzINo",
    authDomain: "absenit-92dd9.firebaseapp.com",
    projectId: "absenit-92dd9",
    storageBucket: "absenit-92dd9.firebasestorage.app",
    messagingSenderId: "219381976986",
    appId: "1:219381976986:web:7f1ac4bd70470a52e63423"
};

const app = initializeApp(firebaseConfig);
const db = getFirestore(app);

const canvas = document.getElementById('signature-pad');
const signaturePad = new SignaturePad(canvas);

function resizeCanvas() {
    const ratio = Math.max(window.devicePixelRatio || 1, 1);
    canvas.width = canvas.offsetWidth * ratio;
    canvas.height = canvas.offsetHeight * ratio;
    canvas.getContext("2d").scale(ratio, ratio);
    signaturePad.clear();
}
window.addEventListener("resize", resizeCanvas);
resizeCanvas();

let dataSiswa = {};

async function loadStudentData() {
    try {
        const response = await fetch('assets/Absensi.xlsx');
        const arrayBuffer = await response.arrayBuffer();
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(arrayBuffer);

        // Mengambil worksheet 'JULI' atau sheet terakhir (sheet aktif)
        let worksheet = workbook.getWorksheet('JULI');
        if (!worksheet) {
            worksheet = workbook.worksheets[workbook.worksheets.length - 1];
        }
        const datalist = document.getElementById('studentList');

        // Kosongkan dulu datalist untuk jaga-jaga
        datalist.innerHTML = "";
        dataSiswa = {};

        worksheet.eachRow((row, rowNumber) => {
            // Mulai membaca dari baris 8 (karena baris 7 adalah HEADER)
            if (rowNumber >= 8) {
                const getVal = (colIndex) => {
                    const cell = row.getCell(colIndex);
                    if (!cell || cell.value === null || cell.value === undefined) return '';

                    if (typeof cell.value === 'object') {
                        if (cell.value.result !== undefined) return String(cell.value.result).trim();
                        if (cell.value.richText) return cell.value.richText.map(t => t.text).join('').trim();
                        if (cell.value.text !== undefined) return String(cell.value.text).trim();
                    }

                    return String(cell.text || cell.value || '').trim();
                };

                const no = getVal(1); // Kolom A (NO)
                const nama = getVal(2); // Kolom B (NAMA)
                const kelas = getVal(3); // Kolom C (KELAS)

                // Hanya ambil baris jika Kolom A berisi angka NO (nomor urut siswa) & Nama tidak kosong
                if (/^\d+$/.test(no) && nama !== "") {
                    dataSiswa[nama] = kelas;

                    const option = document.createElement('option');
                    option.value = nama;
                    datalist.appendChild(option);
                }
            }
        });

        console.log("Data siswa berhasil dimuat:", Object.keys(dataSiswa).length, "siswa.");
    } catch (err) {
        console.error("Gagal load file Excel:", err);
    }
}

// Auto-fill Kelas
document.getElementById('nameSearch').addEventListener('input', (e) => {
    const inputVal = e.target.value;
    const kelasInput = document.getElementById('kelas');
    if (dataSiswa[inputVal]) {
        kelasInput.value = dataSiswa[inputVal];
    } else {
        kelasInput.value = "";
    }
});

document.getElementById('clear').addEventListener('click', () => signaturePad.clear());

document.getElementById('submitBtn').addEventListener('click', async () => {
    const nama = document.getElementById('nameSearch').value;
    const kelas = document.getElementById('kelas').value;
    const tokenUser = document.getElementById('tokenUser').value.trim();

    // 1. Validasi Input Dasar
    if (!dataSiswa[nama]) {
        alert("Silakan pilih nama yang tersedia di daftar!");
        return;
    }

    if (tokenUser === "") {
        alert("Masukkan token terlebih dahulu!");
        return;
    }

    if (signaturePad.isEmpty()) {
        alert("Tanda tangan tidak boleh kosong!");
        return;
    }

    try {
        // 2. Validasi Token ke Firebase
        const tokenRef = doc(db, "system", "token_aktif");
        const tokenSnap = await getDoc(tokenRef);

        if (!tokenSnap.exists()) {
            alert("Sistem belum siap. Hubungi admin untuk aktivasi token.");
            return;
        }

        const { token, createdAt, duration } = tokenSnap.data();

        if (tokenUser.toUpperCase() !== token.toUpperCase()) {
            alert("Token salah!");
            return;
        }

        const timeCreated = createdAt.toDate().getTime();
        const timeNow = new Date().getTime();
        const diffInMinutes = (timeNow - timeCreated) / (1000 * 60);

        if (diffInMinutes > duration) {
            alert("Token sudah expired! Silakan minta token baru ke admin.");
            return;
        }

        // 3. Simpan Data
        const ttdData = signaturePad.toDataURL();
        await addDoc(collection(db, "presensi"), {
            nama: nama,
            kelas: kelas,
            ttd: ttdData,
            timestamp: new Date()
        });

        alert("Absensi Berhasil Terkirim!");
        resetForm();

    } catch (e) {
        console.error(e);
        alert("Terjadi kesalahan: " + e.message);
    }
});

function resetForm() {
    document.getElementById('nameSearch').value = "";
    document.getElementById('kelas').value = "";
    document.getElementById('tokenUser').value = "";
    document.getElementById('kelas').placeholder = "Otomatis terisi";
    signaturePad.clear();
}

document.getElementById('cancelBtn').addEventListener('click', resetForm);
loadStudentData();
