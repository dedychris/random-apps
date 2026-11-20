# Cara bikin Wedding Quest jadi web biasa (gratis, tanpa akun Claude)

Hasil akhirnya satu link `https://script.google.com/macros/s/.../exec` yang bisa dibuka siapa aja di HP atau laptop, tanpa login.
Semua data (status quest, budget, tamu, vendor) tersimpan di Google Sheet milik kamu, dan dua HP yang buka link yang sama akan saling sinkron (update tiap ±12 detik).

Waktu yang dibutuhkan sekitar 10 menit. Yang dibutuhkan cuma akun Google kamu.

## 1. Bikin Google Sheet untuk penyimpanan

1. Buka https://sheets.new (Google Sheet baru).
2. Ganti namanya, misalnya `Wedding Quest Data`.

## 2. Pasang script-nya

1. Di Google Sheet itu, klik menu **Extensions → Apps Script** (Ekstensi → Apps Script).
2. Di editor yang terbuka ada file `Code.gs`. Hapus semua isinya, lalu tempel seluruh isi file **`Code.gs`** dari folder ini.
3. Di baris paling atas, ganti PIN:
   ```js
   const PIN = '1234';
   ```
   jadi kode rahasia kalian sendiri, misalnya `'270915'`. PIN ini nanti diketik sekali di tiap HP.
4. Klik tombol **+** di sebelah "Files" → **HTML**. Beri nama persis `Index` (tanpa `.html`, huruf I besar).
5. Hapus isi bawaannya, lalu tempel seluruh isi file **`Index.html`** dari folder ini.
6. Klik ikon **Save** (disket).

## 3. Beri izin (sekali saja)

1. Di toolbar atas editor, pilih fungsi **`setup`** di dropdown, lalu klik **Run**.
2. Google akan minta izin. Klik **Review permissions**, pilih akun kamu.
3. Kalau muncul "Google hasn't verified this app", klik **Advanced → Go to ... (unsafe)**. Ini aman karena script-nya milik kamu sendiri.
4. Klik **Allow**. Di Google Sheet akan muncul tab baru `wq_data`.

## 4. Publikasikan sebagai web

1. Klik **Deploy → New deployment**.
2. Klik ikon gerigi di sebelah "Select type" → pilih **Web app**.
3. Isi:
   - **Execute as:** `Me`
   - **Who has access:** `Anyone` (bukan "Anyone with Google account", supaya tidak perlu login)
4. Klik **Deploy**, lalu salin **Web app URL** (yang berakhiran `/exec`).

## 5. Bagikan

Kirim link `/exec` itu dan PIN-nya ke pasangan. Di HP:
- Buka link-nya, masukkan PIN, selesai.
- Biar terasa seperti aplikasi: di Chrome pilih **⋮ → Add to Home screen**, di Safari pilih **Share → Add to Home Screen**.

Catatan:
- Google selalu menampilkan baris kecil di atas halaman ("This application was created by a Google Apps Script user"). Itu normal.
- Siapa pun yang punya link **dan** PIN bisa mengubah data, jadi jangan sebar PIN-nya.
- Jangan edit tab `wq_data` secara manual. Tab itu database aplikasinya.

## Kalau nanti ada versi baru

1. Ganti isi file `Index.html` (dan `Code.gs` kalau berubah) di editor Apps Script, lalu Save.
2. **Deploy → Manage deployments** → klik ikon pensil → **Version: New version** → **Deploy**.

Link-nya tetap sama dan datanya tidak hilang.
