<!DOCTYPE html>
<html lang="id">
<head>
    <meta charset="UTF-8">
    <title>Uji Coba Absensi & Rekap</title>
</head>
<body>
    <h2>Form Absensi Sekolah Rakyat</h2>
    
    <!-- Pilihan Kegiatan -->
    <label>Pilih Kegiatan:</label>
    <select id="pilihanKegiatan" onchange="muatRekap()">
        <option value="Sholat Subuh">Sholat Subuh</option>
        <option value="Apel Pagi">Apel Pagi</option>
        <option value="Belajar Mandiri">Belajar Mandiri</option>
    </select>
    <br><br>

    <!-- Form Input ID Siswa untuk Tes -->
    <label>ID Siswa / Scan Kartu:</label>
    <input type="text" id="idSiswaInput" placeholder="Masukkan ID Siswa..." autofocus>
    <button onclick="kirimAbsen()">Kirim Absen</button>
    
    <p id="infoStatus" style="font-weight: bold;"></p>

    <h3>Tabel Rekap Kehadiran</h3>
    <table border="1" cellpadding="5" cellspacing="0">
        <thead>
            <tr>
                <th>No</th>
                <th>ID</th>
                <th>Nama</th>
                <th>Kelas</th>
                <th>Status</th>
                <th>Waktu</th>
                <th>Petugas</th>
            </tr>
        </thead>
        <tbody id="tabelRekapBody">
            <tr><td colspan="7">Memuat data rekap...</td></tr>
        </tbody>
    </table>

    <script>
        // Ganti dengan URL Web App Apps Script Anda setelah di-deploy
        const scriptURL = "https://script.google.com/macros/s/AKfycb.../exec"; 
        const namaPetugas = "Wali Asuh Jaga"; // Contoh nama petugas

        // Muat rekap saat pertama kali halaman dibuka
        window.onload = function() {
            muatRekap();
        };

        // Fungsi Kirim Absen / Scan
        function kirimAbsen() {
            const idSiswa = document.getElementById("idSiswaInput").value.trim();
            const kegiatan = document.getElementById("pilihanKegiatan").value;

            if (!idSiswa) {
                alert("ID Siswa tidak boleh kosong!");
                return;
            }

            document.getElementById("infoStatus").innerText = "Memproses...";

            fetch(scriptURL, {
                method: "POST",
                body: JSON.stringify({
                    idSiswa: idSiswa,
                    kegiatan: kegiatan,
                    petugas: namaPetugas
                })
            })
            .then(res => res.json())
            .then(data => {
                const info = document.getElementById("infoStatus");
                if (data.status === "success") {
                    info.style.color = "green";
                    info.innerText = "Berhasil Absen: " + data.siswa.nama;
                    document.getElementById("idSiswaInput").value = ""; // Bersihkan input
                    
                    // 🟢 KUNCI UTAMA: Langsung muat ulang tabel rekap agar berubah jadi "Hadir"
                    muatRekap(); 
                } else if (data.status === "warning") {
                    info.style.color = "orange";
                    info.innerText = data.pesan;
                } else {
                    info.style.color = "red";
                    info.innerText = data.pesan;
                }
            })
            .catch(err => {
                console.error(err);
                document.getElementById("infoStatus").innerText = "Terjadi kesalahan koneksi.";
            });
        }

        // Fungsi Ambil & Render Rekap Keseluruhan
        function muatRekap() {
            const kegiatan = document.getElementById("pilihanKegiatan").value;

            fetch(scriptURL, {
                method: "POST",
                body: JSON.stringify({
                    action: "getRekapKeseluruhan",
                    kegiatan: kegiatan
                })
            })
            .then(res => res.json())
            .then(data => {
                if (data.status === "success") {
                    const tbody = document.getElementById("tabelRekapBody");
                    tbody.innerHTML = "";

                    data.rekap.forEach(row => {
                        let warnaStatus = row.status === "Hadir" ? "green" : "red";
                        let tr = `<tr>
                            <td>${row.no}</td>
                            <td>${row.id}</td>
                            <td>${row.nama}</td>
                            <td>${row.kelas}</td>
                            <td style="color: ${warnaStatus}; font-weight: bold;">${row.status}</td>
                            <td>${row.waktu}</td>
                            <td>${row.petugas}</td>
                        </tr>`;
                        tbody.innerHTML += tr;
                    });
                }
            })
            .catch(err => console.error("Gagal memuat rekap:", err));
        }
    </script>
</body>
</html>
