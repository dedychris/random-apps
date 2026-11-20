"""Build Rencana-Pernikahan.xlsx: the formal, simple spreadsheet version.

Same tasks and rules as the game (quests.json), with formal wording from
formal.json. The Pengaturan sheet holds the couple's choices; a block of flag
formulas turns them into TRUE/FALSE, and each task's "Berlaku" column
evaluates that task's condition from those flags.
"""
import json, pathlib, re
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule, CellIsRule, DataBarRule
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.utils import get_column_letter as COL

here = pathlib.Path(__file__).parent
DATA = json.loads((here / "quests.json").read_text(encoding="utf-8"))
FORMAL = json.loads((here / "formal.json").read_text(encoding="utf-8"))
missing = [q["id"] for q in DATA["quests"] if q["id"] not in FORMAL["quests"]]
assert not missing, missing
OUT = here / "Rencana-Pernikahan.xlsx"

FONT = "Arial"
INK = "333333"
F = lambda **k: Font(name=FONT, color=k.pop("color", INK), **k)
fill = lambda c: PatternFill("solid", start_color=c, end_color=c)
INPUT = fill("FFF2CC")     # cells to fill in
HEAD = fill("D9E2F3")
LIGHT = fill("F2F2F2")
thin = Side(style="thin", color="BFBFBF")
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
WRAP = Alignment(wrap_text=True, vertical="top")
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
DATE_FMT = "dd mmm yyyy"
DAYS_FMT = '0" hari";"lewat "0" hari";"hari ini"'
IDR_FMT = '"Rp" #,##0;-"Rp" #,##0;"-"'
KRW_FMT = '"₩"#,##0;-"₩"#,##0;"-"'

STATUSES = ["Belum Dimulai (시작 전)", "Sedang Dikerjakan (진행 중)", "Tertunda (보류)", "Ditinjau (검토 중)", "Selesai (완료)"]
ST0, DONE = STATUSES[0], STATUSES[-1]
LATE, TIGHT = "Terlambat (지연)", "Waktu Sempit (촉박)"
YN = ["Ya", "Tidak"]
NATS = ["Indonesia", "Korea", "Lainnya"]
RELS = ["Islam", "Kristen Protestan", "Katolik", "Hindu", "Buddha", "Khonghucu", "Tidak beragama"]
LOCS = ["Indonesia", "Korea", "Indonesia dan Korea"]
LIVES = ["Indonesia", "Korea", "Belum ditentukan"]
WHO = {
    "both": "Kedua mempelai / 신랑·신부", "bride": "Mempelai wanita / 신부", "groom": "Mempelai pria / 신랑",
    "KR": "Mempelai WN Korea / 한국 국적 배우자", "ID": "Mempelai WNI / 인도네시아 국적 배우자",
    "notID": "Mempelai WNA / 외국 국적 배우자", "notKR": "Mempelai bukan WN Korea / 비한국 국적 배우자",
    "OT": "Mempelai WN lain / 제3국 국적 배우자",
}
BCATS = ["Tempat acara", "Katering", "Busana dan rias", "Foto dan video", "Dekorasi", "Cincin", "Undangan dan suvenir",
         "Dokumen dan legal", "Perjalanan dan akomodasi", "Seserahan / sangjit / 예단", "Adat dan upacara agama", "Bulan madu", "Lain-lain"]
VCATS = ["Tempat acara / 웨딩홀", "Wedding organizer", "Katering", "Foto dan video", "Busana dan rias / 스드메", "Dekorasi",
         "Pembawa acara dan musik", "Cincin", "Undangan dan suvenir", "Penerjemah", "Notaris / legal", "Vihara / pandita", "Perjalanan dan hotel", "Lain-lain"]
VSTAT = ["Survei", "Negosiasi", "Dipesan", "Uang muka dibayar", "Lunas"]
AREA = {a["id"]: FORMAL["areas"][a["id"]][0] + " / " + FORMAL["areas"][a["id"]][1] for a in DATA["areas"]}

wb = Workbook()


def dv_list(ws, options, rng):
    dv = DataValidation(type="list", formula1='"' + ",".join(options) + '"', allow_blank=True)
    dv.error, dv.errorTitle = "Silakan pilih dari daftar. / 목록에서 선택하십시오.", "Pilihan tidak valid"
    ws.add_data_validation(dv)
    dv.add(rng)


def name(n, ref):
    wb.defined_names[n] = DefinedName(n, attr_text=ref)


def header(ws, row, labels, widths):
    for i, lab in enumerate(labels, 1):
        c = ws.cell(row=row, column=i, value=lab)
        c.font, c.fill, c.alignment, c.border = F(bold=True), HEAD, CENTER, BOX
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[COL(i)].width = w
    ws.row_dimensions[row].height = 30


def title(ws, text, sub):
    ws["A1"] = text
    ws["A1"].font = F(bold=True, size=14)
    ws["A2"] = sub
    ws["A2"].font = F(italic=True, size=9, color="595959")


# ------------------------------------------------------------------ Pengaturan
ws = wb.active
ws.title = "Pengaturan"
ws.sheet_view.showGridLines = False
for c, w in zip("ABCDEFG", [3, 46, 24, 4, 4, 18, 10]):
    ws.column_dimensions[c].width = w
ws["B1"] = "Rencana Pernikahan — Pengaturan / 결혼 준비 — 설정"
ws["B1"].font = F(bold=True, size=14)
ws["B2"] = "Isi sel berwarna kuning. Daftar tugas, tenggat, dan ringkasan akan menyesuaikan secara otomatis. / 노란색 칸을 입력하면 업무 목록과 마감일이 자동으로 반영됩니다."
ws["B2"].font = F(italic=True, size=9, color="595959")
ws["B2"].alignment = WRAP
ws.merge_cells("B2:C2")
ws.row_dimensions[2].height = 30

sections = [(3, "Mempelai wanita / 신부"), (7, "Mempelai pria / 신랑"), (11, "Pernikahan / 결혼"), (18, "Tanggal / 날짜"), (22, "Anggaran / 예산")]
inputs = [
    (4, "Nama / 이름", None, None),
    (5, "Kewarganegaraan / 국적", "Korea", NATS),
    (6, "Agama / 종교", "Buddha", RELS),
    (8, "Nama / 이름", None, None),
    (9, "Kewarganegaraan / 국적", "Indonesia", NATS),
    (10, "Agama / 종교", "Buddha", RELS),
    (12, "Lokasi pernikahan / 결혼식 장소", "Indonesia", LOCS),
    (13, "Pencatatan resmi pertama (jika di dua negara) / 우선 혼인 등록 국가", "Indonesia", ["Indonesia", "Korea"]),
    (14, "Negara tempat tinggal setelah menikah / 결혼 후 거주 국가", "Belum ditentukan", LIVES),
    (15, "Menggunakan adat Tionghoa / 중국계 전통 예식", "Ya", YN),
    (16, "Menggunakan adat Korea / 한국 전통 예식", "Ya", YN),
    (19, "Tanggal mulai persiapan / 준비 시작일", "=TODAY()", None),
    (20, "Tanggal pernikahan / 결혼식 날짜", "=EDATE(C19,12)", None),
    (21, "Tanggal acara di negara kedua (jika ada) / 두 번째 예식 날짜", None, None),
    (23, "Anggaran total / 총예산", 0, None),
    (24, "Mata uang anggaran / 예산 통화", "IDR", ["IDR", "KRW"]),
    (25, "Kurs: 1 KRW = ... IDR / 환율", 11.5, None),
]
for r, txt in sections:
    ws.cell(row=r, column=2, value=txt).font = F(bold=True, size=11)
for r, label, val, opts in inputs:
    ws.cell(row=r, column=2, value=label).font = F()
    c = ws.cell(row=r, column=3, value=val)
    c.fill, c.border, c.font = INPUT, BOX, F(color="0000FF")
    if opts:
        dv_list(ws, opts, f"C{r}")
for r in (19, 20, 21):
    ws[f"C{r}"].number_format = DATE_FMT
ws["C23"].number_format = "#,##0"
ws["C25"].number_format = "0.00"
ws["B27"] = ("Catatan: Agama Buddha dan adat Tionghoa serta Korea telah diisi sesuai rencana awal; ubah bila berbeda. "
             "Tanggal mulai dan tanggal pernikahan berisi rumus contoh (hari ini dan 12 bulan kemudian); ganti dengan tanggal sebenarnya. "
             "Kurs 11,5 merupakan perkiraan; perbarui sesuai kurs terkini.")
ws["B27"].font = F(italic=True, size=9, color="595959")
ws["B27"].alignment = WRAP
ws.merge_cells("B27:C29")

ws["F3"] = "Aturan otomatis (jangan diubah)"
ws["F3"].font = F(bold=True, size=8, color="808080")
FL = [
    ("hasID", 'OR($C$5="Indonesia",$C$9="Indonesia")'),
    ("hasKR", 'OR($C$5="Korea",$C$9="Korea")'),
    ("hasOT", 'OR($C$5="Lainnya",$C$9="Lainnya")'),
    ("nonID", 'OR($C$5<>"Indonesia",$C$9<>"Indonesia")'),
    ("nonKR", 'OR($C$5<>"Korea",$C$9<>"Korea")'),
    ("mixedNat", '$C$5<>$C$9'),
    ("regID", 'OR($C$12="Indonesia",AND($C$12="Indonesia dan Korea",$C$13<>"Korea"))'),
    ("regKR", 'OR($C$12="Korea",AND($C$12="Indonesia dan Korea",$C$13="Korea"))'),
    ("cerID", '$C$12<>"Korea"'),
    ("cerKR", '$C$12<>"Indonesia"'),
    ("relSet", 'AND($C$6<>"",$C$10<>"")'),
    ("mixedRel", 'AND(f_relSet,$C$6<>$C$10)'),
    ("noRel", 'OR($C$6="Tidak beragama",$C$10="Tidak beragama")'),
    ("bothIslam", 'AND($C$6="Islam",$C$10="Islam")'),
    ("anyIslam", 'OR($C$6="Islam",$C$10="Islam")'),
    ("anyBuddha", 'OR($C$6="Buddha",$C$10="Buddha")'),
    ("kua", 'AND(f_regID,f_bothIslam)'),
    ("dukcapil", 'AND(f_regID,f_relSet,NOT(f_bothIslam),NOT(f_mixedRel),NOT(f_noRel))'),
    ("relP", 'AND(f_relSet,NOT(f_mixedRel),$C$6="Kristen Protestan")'),
    ("relK", 'AND(f_relSet,NOT(f_mixedRel),$C$6="Katolik")'),
    ("relH", 'AND(f_relSet,NOT(f_mixedRel),$C$6="Hindu")'),
    ("relB", 'AND(f_relSet,NOT(f_mixedRel),$C$6="Buddha")'),
    ("relC", 'AND(f_relSet,NOT(f_mixedRel),$C$6="Khonghucu")'),
    ("adatCN", '$C$15="Ya"'),
    ("adatKR", '$C$16="Ya"'),
    ("needF6", 'AND(f_hasKR,f_nonKR,OR($C$14="Korea",$C$14="Belum ditentukan"))'),
    ("needKITAS", 'AND(f_hasID,f_nonID,OR($C$14="Indonesia",$C$14="Belum ditentukan"))'),
    ("reportKR", 'AND(f_regID,f_hasKR)'),
    ("travelFam", 'OR(AND(f_hasKR,f_cerID),AND(f_hasID,f_cerKR))'),
    ("idFamToKR", 'AND(f_hasID,f_cerKR)'),
    ("krFamToID", 'AND(f_hasKR,f_cerID)'),
    ("bilingual", 'AND(f_hasKR,f_hasID)'),
]
helpers = [
    ("start", "=$C$19", DATE_FMT), ("wedding", "=$C$20", DATE_FMT), ("date2", "=$C$21", DATE_FMT),
    ("ratio", "=MIN(1,MAX(1,$C$20-$C$19)/365)", "0.00"),
    ("second", '=IF($C$12="Indonesia dan Korea",IF($C$13="Korea","ID","KR"),"")', "@"),
    ("rate", "=$C$25", "0.00"),
]
r = 4
for key, formula in FL:
    ws.cell(row=r, column=6, value=key).font = F(size=8, color="808080")
    ws.cell(row=r, column=7, value="=" + formula).font = F(size=8, color="808080")
    name("f_" + key, f"Pengaturan!$G${r}")
    r += 1
for key, formula, fmt in helpers:
    ws.cell(row=r, column=6, value=key).font = F(size=8, color="808080")
    c = ws.cell(row=r, column=7, value=formula)
    c.font, c.number_format = F(size=8, color="808080"), fmt
    name("f_" + key, f"Pengaturan!$G${r}")
    r += 1


def to_excel(src):
    """Translate a quest condition (idents, !, &&, ||, parentheses) into an Excel expression."""
    if not src:
        return "TRUE"
    toks = re.findall(r"[A-Za-z_]\w*|&&|\|\||!|\(|\)", src)
    pos = [0]

    def peek():
        return toks[pos[0]] if pos[0] < len(toks) else None

    def nxt():
        pos[0] += 1
        return toks[pos[0] - 1]

    def prim():
        tk = nxt()
        if tk == "!":
            return f"NOT({prim()})"
        if tk == "(":
            a = orx()
            nxt()
            return a
        return "f_" + tk

    def join(fn, op, sub):
        parts = [sub()]
        while peek() == op:
            nxt()
            parts.append(sub())
        return parts[0] if len(parts) == 1 else fn + "(" + ",".join(parts) + ")"

    def andx():
        return join("AND", "&&", prim)

    def orx():
        return join("OR", "||", andx)

    return orx()


# ------------------------------------------------------------------ Daftar Tugas
# Columns: A No | B Tahap | C Tema | D Tugas | E 업무 | F Penanggung jawab | G Berlaku | H Tenggat | I Sisa waktu
#          J Status | K Keterangan | L Catatan | M Tenggat manual | N Uraian | O 설명 | P Tautan | hidden Q b, R m, S ev, T auto, U tight
qs = wb.create_sheet("Daftar Tugas")
cols = ["No.", "Tahap / 단계", "Tema / 테마", "Tugas", "업무", "Penanggung jawab / 담당", "Berlaku / 해당", "Tenggat / 마감일", "Sisa waktu / 남은 기간",
        "Status / 상태", "Keterangan / 비고", "Catatan / 메모", "Tenggat manual (opsional) / 수동 마감일", "Uraian", "설명", "Tautan / 링크",
        "b", "m", "ev", "auto", "tight"]
header(qs, 1, cols, [5, 20, 24, 40, 32, 22, 9, 13, 13, 22, 16, 30, 16, 60, 50, 28, 5, 5, 5, 12, 6])
qs.freeze_panes = "E2"
GROUPS = {g["id"]: g for g in DATA["groups"]}
THEME = {gid: FORMAL["groups"][gid][0] + " / " + FORMAL["groups"][gid][1] for gid in GROUPS}
gorder = {g["id"]: i for i, g in enumerate(DATA["groups"])}
quests = sorted(DATA["quests"], key=lambda q: (gorder[q["g"]], -q["b"] if q["b"] > 0 else 10000 - q["b"]))
for n, q in enumerate(quests, 1):
    r = n + 1
    fm = FORMAL["quests"][q["id"]]
    for ci, v in enumerate([n, AREA[q["a"]], THEME[q["g"]], fm["t"][0], fm["t"][1], WHO[q["w"]]], 1):
        qs.cell(row=r, column=ci, value=v)
    qs.cell(row=r, column=7, value=f'=IF({to_excel(q.get("c"))},"Ya","Tidak")')
    qs.cell(row=r, column=17, value=q["b"])
    qs.cell(row=r, column=18, value=q["m"])
    qs.cell(row=r, column=19, value=q.get("ev", ""))
    base = f'IF(AND(S{r}<>"",S{r}=f_second,f_date2<>""),f_date2,f_wedding)'
    qs.cell(row=r, column=20, value=f"=IF(Q{r}<=0,{base}-Q{r},MAX(f_start,{base}-MAX(R{r},ROUND(Q{r}*f_ratio,0))))")
    qs.cell(row=r, column=21, value=f"=AND(Q{r}>0,{base}-MAX(R{r},ROUND(Q{r}*f_ratio,0))<f_start)")
    qs.cell(row=r, column=8, value=f'=IF(M{r}<>"",M{r},T{r})')
    qs.cell(row=r, column=9, value=f'=IF(G{r}="Ya",H{r}-TODAY(),"")')
    qs.cell(row=r, column=10, value=ST0)
    qs.cell(row=r, column=11, value=f'=IF(G{r}<>"Ya","",IF(AND(J{r}<>"{DONE}",H{r}<TODAY()),"{LATE}",IF(AND(U{r},M{r}="",J{r}<>"{DONE}"),"{TIGHT}","")))')
    qs.cell(row=r, column=14, value=fm["d"][0])
    qs.cell(row=r, column=15, value=fm["d"][1])
    if q.get("l"):
        c = qs.cell(row=r, column=16, value=q["l"])
        c.hyperlink = q["l"]
last_q = len(quests) + 1
row = last_q + 1
qs.cell(row=row, column=4, value="Tugas tambahan (isi tahap, tema, tugas, dan tenggat manual) / 추가 업무").font = F(bold=True, italic=True)
custom_start = row + 1
last_row = custom_start + 19
for r in range(custom_start, last_row + 1):
    qs.cell(row=r, column=1, value=f'=IF(D{r}="","",ROW()-1)')
    qs.cell(row=r, column=7, value=f'=IF(D{r}="","","Ya")')
    qs.cell(row=r, column=8, value=f'=IF(M{r}="","",M{r})')
    qs.cell(row=r, column=9, value=f'=IF(AND(D{r}<>"",M{r}<>""),M{r}-TODAY(),"")')
    qs.cell(row=r, column=11, value=f'=IF(AND(D{r}<>"",M{r}<>"",J{r}<>"{DONE}",M{r}<TODAY()),"{LATE}","")')
    for ci in (2, 3, 4, 5, 6):
        qs.cell(row=r, column=ci).fill = INPUT
for r in list(range(2, last_q + 1)) + list(range(custom_start, last_row + 1)):
    for ci in range(1, 22):
        c = qs.cell(row=r, column=ci)
        c.font, c.border, c.alignment = F(size=9), BOX, WRAP
    for ci in (8, 13, 20):
        qs.cell(row=r, column=ci).number_format = DATE_FMT
    qs.cell(row=r, column=9).number_format = DAYS_FMT
    for ci in (10, 12, 13):
        qs.cell(row=r, column=ci).fill = INPUT
dv_list(qs, STATUSES, f"J2:J{last_row}")
dv_list(qs, list(AREA.values()), f"B{custom_start}:B{last_row}")
for c in "QRSTU":
    qs.column_dimensions[c].hidden = True
qs.auto_filter.ref = f"A1:P{last_row}"
rng = f"A2:P{last_row}"
qs.conditional_formatting.add(rng, FormulaRule(formula=['$G2="Tidak"'], font=Font(name=FONT, color="BFBFBF"), stopIfTrue=True))
qs.conditional_formatting.add(rng, FormulaRule(formula=[f'$J2="{DONE}"'], fill=fill("E2EFDA")))
qs.conditional_formatting.add(f"K2:K{last_row}", CellIsRule(operator="equal", formula=[f'"{LATE}"'], fill=fill("F8CBAD"), font=Font(name=FONT, bold=True, color="9C0006")))
qs.conditional_formatting.add(f"K2:K{last_row}", CellIsRule(operator="equal", formula=[f'"{TIGHT}"'], fill=fill("FFE699")))

# ------------------------------------------------------------------ Anggaran
bs = wb.create_sheet("Anggaran")
title(bs, "Anggaran / 예산", "Isi kolom A–F dan M. Kolom IDR dan KRW dihitung dari kurs pada lembar Pengaturan. Baris 5 adalah contoh dan dapat dihapus.")
header(bs, 4, ["Uraian biaya / 항목", "Kategori / 분류", "Mata uang / 통화", "Rencana / 예상", "Realisasi / 실제", "Telah dibayar / 지급액",
               "Rencana (IDR)", "Realisasi (IDR)", "Dibayar (IDR)", "Rencana (KRW)", "Realisasi (KRW)", "Sisa pembayaran (IDR)", "Catatan / 메모"],
       [28, 24, 11, 15, 15, 15, 16, 16, 16, 15, 15, 17, 30])
B0, B1 = 5, 154
for r in range(B0, B1 + 1):
    for ci in list(range(1, 7)) + [13]:
        bs.cell(row=r, column=ci).fill = INPUT
    for ci, src in ((7, "D"), (8, "E"), (9, "F")):
        bs.cell(row=r, column=ci, value=f'=IF($A{r}="","",IF($C{r}="KRW",{src}{r}*f_rate,N({src}{r})))').number_format = IDR_FMT
    bs.cell(row=r, column=10, value=f'=IF($A{r}="","",IF(f_rate=0,0,G{r}/f_rate))').number_format = KRW_FMT
    bs.cell(row=r, column=11, value=f'=IF($A{r}="","",IF(f_rate=0,0,H{r}/f_rate))').number_format = KRW_FMT
    bs.cell(row=r, column=12, value=f'=IF($A{r}="","",MAX(0,IF(H{r}>0,H{r},G{r})-I{r}))').number_format = IDR_FMT
    for ci in (4, 5, 6):
        bs.cell(row=r, column=ci).number_format = "#,##0"
    for ci in range(1, 14):
        bs.cell(row=r, column=ci).border = BOX
        bs.cell(row=r, column=ci).font = F(size=9)
for ci, v in enumerate(["Uang muka gedung (contoh)", "Tempat acara", "IDR", 60000000, 60000000, 20000000], 1):
    bs.cell(row=B0, column=ci, value=v)
dv_list(bs, BCATS, f"B{B0}:B{B1}")
dv_list(bs, ["IDR", "KRW"], f"C{B0}:C{B1}")
bs.freeze_panes = "B5"

# ------------------------------------------------------------------ Tamu
gs = wb.create_sheet("Tamu")
title(gs, "Daftar Tamu / 하객 명단", "Satu baris untuk setiap keluarga atau rombongan. Kolom acara diisi Ya atau Tidak. Baris 5 adalah contoh.")
header(gs, 4, ["Nama / 이름", "Pihak / 소속", "Kelompok / 관계", "Negara / 국가", "Jumlah orang / 인원", "Upacara agama / 종교 예식",
               "Resepsi Indonesia / 인도네시아 피로연", "Upacara Korea / 한국 예식", "Undangan terkirim / 청첩장 발송", "Kehadiran / 참석 여부",
               "Kontak / 연락처", "Catatan / 메모"],
       [26, 16, 13, 12, 12, 14, 16, 14, 14, 15, 18, 28])
G0, G1 = 5, 404
for r in range(G0, G1 + 1):
    for ci in range(1, 13):
        c = gs.cell(row=r, column=ci)
        c.fill, c.border, c.font = INPUT, BOX, F(size=9)
for ci, v in enumerate(["Keluarga Kim (contoh)", "Mempelai wanita", "Keluarga", "Korea", 4, "Ya", "Ya", "Tidak", "Ya", "Hadir", "", "Datang dari Seoul"], 1):
    gs.cell(row=G0, column=ci, value=v)
dv_list(gs, ["Mempelai wanita", "Mempelai pria", "Bersama"], f"B{G0}:B{G1}")
dv_list(gs, ["Keluarga", "Teman", "Rekan kerja", "Lainnya"], f"C{G0}:C{G1}")
dv_list(gs, NATS, f"D{G0}:D{G1}")
dv_list(gs, YN, f"F{G0}:I{G1}")
dv_list(gs, ["Belum konfirmasi", "Hadir", "Tidak hadir"], f"J{G0}:J{G1}")
gs.freeze_panes = "B5"
GR = lambda c: f"Tamu!${c}${G0}:${c}${G1}"

# ------------------------------------------------------------------ Vendor
vs = wb.create_sheet("Vendor")
title(vs, "Daftar Vendor / 업체 목록", "Catat seluruh vendor yang dipertimbangkan. Baris 5 adalah contoh.")
header(vs, 4, ["Nama vendor / 업체명", "Kategori / 분류", "Status / 상태", "Narahubung / 담당자", "Telepon / 전화", "Email / Instagram",
               "Harga / 가격", "Mata uang / 통화", "Harga (IDR)", "Jatuh tempo / 잔금 기한", "Penilaian (1–5) / 평가", "Catatan / 메모"],
       [26, 24, 18, 18, 18, 22, 15, 11, 16, 14, 12, 32])
V0, V1 = 5, 104
for r in range(V0, V1 + 1):
    for ci in range(1, 13):
        c = vs.cell(row=r, column=ci)
        c.border, c.font = BOX, F(size=9)
        if ci != 9:
            c.fill = INPUT
    vs.cell(row=r, column=5).number_format = "@"
    vs.cell(row=r, column=7).number_format = "#,##0"
    vs.cell(row=r, column=9, value=f'=IF(A{r}="","",IF(H{r}="KRW",G{r}*f_rate,N(G{r})))').number_format = IDR_FMT
    vs.cell(row=r, column=10).number_format = DATE_FMT
for ci, v in enumerate(["Studio Contoh", "Foto dan video", "Negosiasi", "Ibu Rina", "+62 812 0000 0000", "@studiocontoh", 25000000, "IDR", None, None, 4, "Paket 8 jam dan album"], 1):
    if v is not None:
        vs.cell(row=V0, column=ci, value=v)
dv_list(vs, VCATS, f"B{V0}:B{V1}")
dv_list(vs, VSTAT, f"C{V0}:C{V1}")
dv_list(vs, ["IDR", "KRW"], f"H{V0}:H{V1}")
dv_list(vs, ["1", "2", "3", "4", "5"], f"K{V0}:K{V1}")
vs.freeze_panes = "B5"

# ------------------------------------------------------------------ Ringkasan
ds = wb.create_sheet("Ringkasan", 1)
ds.sheet_view.showGridLines = False
for c, w in zip("ABCDEF", [3, 46, 18, 18, 14, 12]):
    ds.column_dimensions[c].width = w
ds["B1"] = '=IF(OR(Pengaturan!C4="",Pengaturan!C8=""),"Ringkasan Persiapan Pernikahan","Ringkasan Persiapan Pernikahan "&Pengaturan!C4&" dan "&Pengaturan!C8)'
ds["B1"].font = F(bold=True, size=14)
ds["B2"] = "Seluruh angka dihitung otomatis. / 모든 수치는 자동으로 계산됩니다."
ds["B2"].font = F(italic=True, size=9, color="595959")
TS = "'Daftar Tugas'"
QF, QI, QB, QJ, QH, QC = (f"{TS}!${c}$2:${c}${last_row}" for c in "GJBKIC")
ds["B4"] = "Progres / 진행 현황"
ds["B4"].font = F(bold=True, size=11)
kpis = [
    ("Tanggal pernikahan / 결혼식 날짜", "=f_wedding", DATE_FMT),
    ("Sisa waktu / 남은 기간", '=IF(f_wedding="","",f_wedding-TODAY())', DAYS_FMT),
    ("Jumlah tugas yang berlaku / 해당 업무 수", f'=COUNTIF({QF},"Ya")', "0"),
    ("Tugas selesai / 완료된 업무", f'=COUNTIFS({QF},"Ya",{QI},"{DONE}")', "0"),
    ("Persentase selesai / 완료율", "=IF(C7=0,0,C8/C7)", "0%"),
    ("Tugas terlambat / 지연 업무", f'=COUNTIF({QJ},"{LATE}")', "0"),
    ("Jatuh tempo dalam 30 hari / 30일 이내 마감", f'=COUNTIFS({QF},"Ya",{QI},"<>{DONE}",{QH},">=0",{QH},"<=30")', "0"),
]
for i, (lab, f, fmt) in enumerate(kpis):
    r = 5 + i
    ds.cell(row=r, column=2, value=lab).font = F()
    c = ds.cell(row=r, column=3, value=f)
    c.font, c.number_format, c.fill, c.border = F(bold=True), fmt, LIGHT, BOX
ds.conditional_formatting.add("C10", CellIsRule(operator="greaterThan", formula=["0"], fill=fill("F8CBAD"), font=Font(name=FONT, bold=True, color="9C0006")))

r = 13
ds.cell(row=r, column=2, value="Progres per tahap / 단계별 진행").font = F(bold=True, size=11)
r += 1
for ci, lab in enumerate(["Tahap / 단계", "Berlaku / 해당", "Selesai / 완료", "Persentase / 비율", "Terlambat / 지연"], 2):
    c = ds.cell(row=r, column=ci, value=lab)
    c.font, c.fill, c.border, c.alignment = F(bold=True), HEAD, BOX, CENTER
first = r + 1
for a in DATA["areas"]:
    r += 1
    ds.cell(row=r, column=2, value=AREA[a["id"]])
    ds.cell(row=r, column=3, value=f'=COUNTIFS({QB},B{r},{QF},"Ya")')
    ds.cell(row=r, column=4, value=f'=COUNTIFS({QB},B{r},{QF},"Ya",{QI},"{DONE}")')
    ds.cell(row=r, column=5, value=f"=IF(C{r}=0,0,D{r}/C{r})").number_format = "0%"
    ds.cell(row=r, column=6, value=f'=COUNTIFS({QB},B{r},{QJ},"{LATE}")')
    for ci in range(2, 7):
        ds.cell(row=r, column=ci).border = BOX
        ds.cell(row=r, column=ci).font = F()
ds.conditional_formatting.add(f"E{first}:E{r}", DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1, color="5B9BD5"))

r += 2
ds.cell(row=r, column=2, value="Progres per tema / 테마별 진행").font = F(bold=True, size=11)
ds.cell(row=r + 1, column=2, value="Tema dengan 0 tugas berlaku tidak relevan untuk pilihan pada lembar Pengaturan.").font = F(italic=True, size=9, color="595959")
r += 2
for ci, lab in enumerate(["Tema / 테마", "Berlaku / 해당", "Selesai / 완료", "Persentase / 비율", "Terlambat / 지연"], 2):
    c = ds.cell(row=r, column=ci, value=lab)
    c.font, c.fill, c.border, c.alignment = F(bold=True), HEAD, BOX, CENTER
tfirst = r + 1
for g in DATA["groups"]:
    r += 1
    ds.cell(row=r, column=2, value=THEME[g["id"]])
    ds.cell(row=r, column=3, value=f'=COUNTIFS({QC},B{r},{QF},"Ya")')
    ds.cell(row=r, column=4, value=f'=COUNTIFS({QC},B{r},{QF},"Ya",{QI},"{DONE}")')
    ds.cell(row=r, column=5, value=f"=IF(C{r}=0,0,D{r}/C{r})").number_format = "0%"
    ds.cell(row=r, column=6, value=f'=COUNTIFS({QC},B{r},{QJ},"{LATE}")')
    for ci in range(2, 7):
        ds.cell(row=r, column=ci).border = BOX
        ds.cell(row=r, column=ci).font = F()
ds.conditional_formatting.add(f"E{tfirst}:E{r}", DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1, color="70AD47"))
ds.conditional_formatting.add(f"B{tfirst}:F{r}", FormulaRule(formula=[f"$C{tfirst}=0"], font=Font(name=FONT, color="BFBFBF")))

r += 2
ds.cell(row=r, column=2, value="Anggaran / 예산").font = F(bold=True, size=11)
r += 1
for ci, lab in enumerate(["", "IDR", "KRW"], 2):
    c = ds.cell(row=r, column=ci, value=lab)
    c.font, c.fill, c.border, c.alignment = F(bold=True), HEAD, BOX, CENTER
b_first = r + 1
for lab, f in [("Anggaran total / 총예산", '=IF(Pengaturan!$C$24="KRW",Pengaturan!$C$23*f_rate,Pengaturan!$C$23)'),
               ("Rencana biaya / 예상 비용", f"=SUM(Anggaran!$G${B0}:$G${B1})"),
               ("Realisasi biaya / 실제 비용", f"=SUM(Anggaran!$H${B0}:$H${B1})"),
               ("Telah dibayar / 지급액", f"=SUM(Anggaran!$I${B0}:$I${B1})")]:
    r += 1
    ds.cell(row=r, column=2, value=lab).font = F()
    ds.cell(row=r, column=3, value=f).number_format = IDR_FMT
    ds.cell(row=r, column=4, value=f"=IF(f_rate=0,0,C{r}/f_rate)").number_format = KRW_FMT
r += 1
ds.cell(row=r, column=2, value="Sisa anggaran / 잔여 예산").font = F(bold=True)
ds.cell(row=r, column=3, value=f"=C{b_first}-MAX(C{b_first + 2},0)").number_format = IDR_FMT
ds.cell(row=r, column=4, value=f"=IF(f_rate=0,0,C{r}/f_rate)").number_format = KRW_FMT
ds.conditional_formatting.add(f"C{r}:D{r}", CellIsRule(operator="lessThan", formula=["0"], fill=fill("F8CBAD")))
for rr in range(b_first, r + 1):
    for ci in (2, 3, 4):
        ds.cell(row=rr, column=ci).border = BOX
        if ci > 2:
            ds.cell(row=rr, column=ci).font = F(bold=True)

r += 2
ds.cell(row=r, column=2, value="Tamu (jumlah orang) / 하객 (인원)").font = F(bold=True, size=11)
for lab, f in [
    ("Total diundang / 초대 인원", f"=SUM({GR('E')})"),
    ("Hadir / 참석", f'=SUMIF({GR("J")},"Hadir",{GR("E")})'),
    ("Belum konfirmasi / 미응답", f'=SUM({GR("E")})-SUMIF({GR("J")},"Hadir",{GR("E")})-SUMIF({GR("J")},"Tidak hadir",{GR("E")})'),
    ("Tidak hadir / 불참", f'=SUMIF({GR("J")},"Tidak hadir",{GR("E")})'),
    ("Pihak mempelai wanita / 신부 측", f'=SUMIF({GR("B")},"Mempelai wanita",{GR("E")})'),
    ("Pihak mempelai pria / 신랑 측", f'=SUMIF({GR("B")},"Mempelai pria",{GR("E")})'),
    ("Dari Korea / 한국", f'=SUMIF({GR("D")},"Korea",{GR("E")})'),
    ("Dari Indonesia / 인도네시아", f'=SUMIF({GR("D")},"Indonesia",{GR("E")})'),
    ("Undangan terkirim / 청첩장 발송", f'=SUMIF({GR("I")},"Ya",{GR("E")})'),
]:
    r += 1
    ds.cell(row=r, column=2, value=lab).font = F()
    c = ds.cell(row=r, column=3, value=f)
    c.number_format, c.border, c.font = "#,##0", BOX, F(bold=True)

# ------------------------------------------------------------------ Panduan
ps = wb.create_sheet("Panduan")
ps.sheet_view.showGridLines = False
ps.column_dimensions["A"].width = 3
ps.column_dimensions["B"].width = 110
lines = [
    ("Petunjuk Penggunaan", True),
    ("1. Isi sel berwarna kuning pada lembar Pengaturan: nama, kewarganegaraan, agama, lokasi, adat, tanggal, anggaran, dan kurs.", False),
    ("2. Pada lembar Daftar Tugas, kolom Berlaku terisi Ya atau Tidak secara otomatis. Gunakan filter pada judul kolom untuk menampilkan tugas yang berlaku saja.", False),
    ("3. Perbarui kolom Status melalui daftar pilihan: Belum Dimulai, Sedang Dikerjakan, Tertunda, Ditinjau, atau Selesai.", False),
    ("4. Tenggat dihitung mundur dari tanggal pernikahan (acuan 12 bulan) dan dipadatkan apabila waktu persiapan lebih singkat. Isi kolom Tenggat manual untuk mengubahnya.", False),
    ("5. Keterangan \"Waktu Sempit\" berarti tenggat ideal sudah terlewati sebelum persiapan dimulai. \"Terlambat\" berarti tenggat telah lewat dan tugas belum selesai.", False),
    ("6. Lembar Anggaran, Tamu, dan Vendor diisi pada baris berwarna kuning. Konversi IDR dan KRW menggunakan kurs pada lembar Pengaturan.", False),
    ("7. Untuk digunakan bersama, unggah berkas ini ke Google Drive, buka dengan Google Sheets, lalu bagikan kepada pasangan sebagai Editor.", False),
    ("", False),
    ("사용 방법", True),
    ("1. 설정 시트의 노란색 칸(이름, 국적, 종교, 장소, 전통 예식, 날짜, 예산, 환율)을 입력합니다.", False),
    ("2. 업무 목록(Daftar Tugas) 시트의 '해당' 열은 자동으로 Ya(해당) 또는 Tidak(해당 없음)으로 표시됩니다. 필터를 사용하여 해당 업무만 볼 수 있습니다.", False),
    ("3. 상태 열은 목록에서 선택합니다: 시작 전, 진행 중, 보류, 검토 중, 완료.", False),
    ("4. 마감일은 결혼식 날짜를 기준으로 자동 계산됩니다. 변경하려면 수동 마감일 칸에 날짜를 입력합니다.", False),
    ("5. 함께 사용하려면 Google 드라이브에 업로드하여 Google 스프레드시트로 열고 상대방을 편집자로 공유합니다.", False),
    ("", False),
    ("Catatan mengenai persyaratan dokumen / 서류 요건 안내", True),
    ("Persyaratan dirangkum dari peraturan dan sumber publik per November 2026 (PMA 30/2024, SEMA 2/2023, ketentuan apostille Indonesia sejak 4 Juni 2022, "
     "dan ketentuan visa F-6). Persyaratan dapat berbeda di setiap daerah dan dapat berubah. Konfirmasikan kepada KUA atau Dukcapil, vihara atau majelis, 구청, "
     "Kedutaan Besar Korea di Jakarta, dan KBRI Seoul sebelum mengurus dokumen.", False),
    ("본 자료는 2026년 11월 기준 공개 자료를 정리한 것으로, 지역에 따라 다르거나 변경될 수 있습니다. 서류 준비 전 관할 기관과 대사관에 반드시 확인하시기 바랍니다.", False),
    ("", False),
    ("Sumber / 출처", True),
]
sources = [
    "https://news.detik.com/berita/d-7725372/syarat-pernikahan-campuran-di-indonesia-bagi-wna-cek-di-sini",
    "https://news.detik.com/berita/d-6831689/tertutup-pintu-pencatatan-nikah-beda-agama-usai-terbit-surat-edaran-ma",
    "https://www.hukumonline.com/klinik/a/menikah-dengan-warga-negara-korea-selatan-lt4fea846707c98/",
    "https://www.hukumonline.com/klinik/detail/cl4623/hak-milik-untuk-wni-yg-kawin-dengan-wna",
    "https://www.hcch.net/es/news-archive/details/?varevent=861",
    "https://apostille.ahu.go.id",
    "https://overseas.mofa.go.kr/us-losangeles-en/wpge/m_20421/contents.do",
    "https://www.moj.go.kr/bbs/moj/156/540963/artclView.do",
]
r = 1
for txt, bold in lines:
    c = ps.cell(row=r, column=2, value=txt)
    c.font, c.alignment = F(bold=bold, size=12 if bold else 10), WRAP
    r += 1
for u in sources:
    c = ps.cell(row=r, column=2, value=u)
    c.hyperlink = u
    c.font = F(size=9, color="0563C1", underline="single")
    r += 1

for sheet, colr in [("Pengaturan", "FFC000"), ("Ringkasan", "5B9BD5"), ("Daftar Tugas", "70AD47"), ("Anggaran", "FFC000"), ("Tamu", "A5A5A5"), ("Vendor", "A5A5A5"), ("Panduan", "A5A5A5")]:
    wb[sheet].sheet_properties.tabColor = colr
wb._sheets = [wb[n] for n in ["Pengaturan", "Ringkasan", "Daftar Tugas", "Anggaran", "Tamu", "Vendor", "Panduan"]]
wb.active = 0
wb.save(OUT)
print("saved", OUT.name, "tasks:", len(quests))
