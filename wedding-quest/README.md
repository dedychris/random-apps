# Wedding Quest

Wedding planner bergaya game untuk pasangan Indonesia–Korea (dan kombinasi lain).

- `quests.json`: sumber data quest (dua bahasa: ID/KO), lengkap dengan kondisi kapan quest berlaku
  (kewarganegaraan, lokasi nikah, agama, rencana tinggal) dan jarak deadline ideal dari Hari H.
- `src/template.html` + `build.py` → versi game:
  - `web/Index.html` + `web/Code.gs`: web app Google Apps Script (bisa dibuka tanpa akun, data di Google Sheet,
    dikunci PIN). Cara pasang: `web/PANDUAN-DEPLOY.md`.
  - `wedding-quest.html`: versi Claude Artifact.
- `formal.json` + `build_xlsx.py` → `Rencana-Pernikahan.xlsx` (versi spreadsheet berbahasa formal, aturan sama
  dalam bentuk formula).

```bash
python3 wedding-quest/build.py
```

```bash
python3 wedding-quest/build_xlsx.py
```

## Aturan deadline

Setiap quest punya `b` (hari sebelum Hari H untuk rencana 12 bulan) dan `m` (batas minimal).
Deadline = Hari H − max(m, round(b × min(1, durasi/365))). Kalau hasilnya jatuh sebelum tanggal mulai,
deadline dipindah ke tanggal mulai dan quest ditandai "Mepet". Nilai `b` negatif berarti hari setelah Hari H.

Syarat dokumen dirangkum per November 2026 dan wajib dikonfirmasi ulang ke KUA/Dukcapil, 구청, dan kedutaan.
