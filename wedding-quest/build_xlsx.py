"""Build Wedding-Quest.xlsx (spreadsheet version) from quests.json.

Same quests and rules as the web version: the Mulai sheet holds the couple's
choices, a block of flag formulas turns them into TRUE/FALSE, and each quest's
"Berlaku?" column evaluates that quest's condition from those flags.
"""
import json, pathlib, re
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule, CellIsRule
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.utils import get_column_letter as COL

here = pathlib.Path(__file__).parent
DATA = json.loads((here / "quests.json").read_text(encoding="utf-8"))
OUT = here / "Wedding-Quest.xlsx"

FONT = "Arial"
INK = "5A3A2C"
F = lambda **k: Font(name=FONT, color=k.pop("color", INK), **k)
fill = lambda c: PatternFill("solid", start_color=c, end_color=c)
INPUT = fill("FFF2B3")      # cells to fill in
HEAD = fill("F8BFCB")
CREAM = fill("FFF8E6")
MINT = fill("D6F0EA")
thin = Side(style="thin", color="C9B29A")
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
WRAP = Alignment(wrap_text=True, vertical="top")
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
DATE_FMT = "dd mmm yyyy"
IDR_FMT = '"Rp" #,##0;-"Rp" #,##0;"-"'
KRW_FMT = '"₩"#,##0;-"₩"#,##0;"-"'

STATUSES = ["Belum mulai", "On progress", "Pending", "Review", "Done"]
NATS = ["Indonesia", "Korea", "Lainnya"]
RELS = ["Islam", "Kristen Protestan", "Katolik", "Hindu", "Buddha", "Khonghucu", "Tidak beragama"]
LOCS = ["Indonesia", "Korea", "Dua-duanya"]
LIVES = ["Indonesia", "Korea", "Belum tahu"]
WHO = {
    "both": "Berdua / 함께", "bride": "Pengantin wanita / 신부", "groom": "Pengantin pria / 신랑",
    "KR": "Pasangan WN Korea / 한국 국적자", "ID": "Pasangan WNI / 인도네시아 국적자",
    "notID": "Pasangan WNA / 외국인 배우자", "notKR": "Pasangan non-Korea / 비한국 국적 배우자",
    "OT": "Pasangan WN lain / 제3국 국적자",
}
LEVELS = ["Tunangan Pemula / 새내기 예비부부", "Penjelajah Rencana / 계획 탐험가", "Pemburu Dokumen / 서류 사냥꾼",
          "Penakluk Vendor / 업체 정복자", "Ahli Budget / 예산의 달인", "Pasangan Tangguh / 든든한 커플",
          "Siap Pelaminan / 결혼 준비 완료", "Pengantin Legendaris / 전설의 신랑신부"]
BCATS = ["Venue", "Catering", "Busana & MUA", "Foto & video", "Dekor & bunga", "Cincin", "Undangan & souvenir",
         "Dokumen & legal", "Perjalanan & akomodasi", "Seserahan / 예단", "Bulan madu", "Lainnya"]
VCATS = ["Venue / 웨딩홀", "Wedding organizer", "Catering", "Foto & video", "스드메 / MUA & busana", "Dekor & bunga",
         "MC & musik", "Cincin", "Undangan & souvenir", "Penerjemah", "Notaris / legal", "Travel & hotel", "Lainnya"]
VSTAT = ["Riset", "Negosiasi", "Booked", "DP dibayar", "Lunas"]

wb = Workbook()


def dv_list(ws, options, rng, allow_blank=True):
    dv = DataValidation(type="list", formula1='"' + ",".join(options) + '"', allow_blank=allow_blank)
    dv.error, dv.errorTitle = "Pilih dari daftar / 목록에서 선택하세요", "Pilihan tidak valid"
    ws.add_data_validation(dv)
    dv.add(rng)


def name(n, ref):
    wb.defined_names[n] = DefinedName(n, attr_text=ref)


def header(ws, row, labels, widths=None):
    for i, lab in enumerate(labels, 1):
        c = ws.cell(row=row, column=i, value=lab)
        c.font, c.fill, c.alignment, c.border = F(bold=True), HEAD, CENTER, BOX
    if widths:
        for i, w in enumerate(widths, 1):
            ws.column_dimensions[COL(i)].width = w


# ------------------------------------------------------------------ Mulai
ws = wb.active
ws.title = "Mulai"
ws.sheet_view.showGridLines = False
ws.column_dimensions["A"].width = 3
ws.column_dimensions["B"].width = 44
ws.column_dimensions["C"].width = 26
ws.column_dimensions["D"].width = 4
ws.column_dimensions["E"].width = 4
ws.column_dimensions["F"].width = 22
ws.column_dimensions["G"].width = 10
ws["B1"] = "💍 Wedding Quest — Mulai di sini / 여기서 시작"
ws["B1"].font = F(bold=True, size=16)
ws["B2"] = "Isi sel kuning. Quest, deadline, dan dashboard menyesuaikan otomatis. / 노란 칸만 채우면 퀘스트와 마감일이 자동으로 바뀌어요."
ws["B2"].font = F(italic=True, size=9)

inputs = [
    (4, "Nama pengantin wanita / 신부 이름", "Ji-yeon (contoh)", None),
    (5, "Kewarganegaraan pengantin wanita / 신부 국적", "Korea", NATS),
    (6, "Agama pengantin wanita / 신부 종교", "", RELS),
    (8, "Nama pengantin pria / 신랑 이름", "Budi (contoh)", None),
    (9, "Kewarganegaraan pengantin pria / 신랑 국적", "Indonesia", NATS),
    (10, "Agama pengantin pria / 신랑 종교", "", RELS),
    (12, "Menikah di / 결혼식 장소", "Indonesia", LOCS),
    (13, "Pencatatan resmi pertama di (kalau dua-duanya) / 먼저 혼인 등록할 나라", "Indonesia", ["Indonesia", "Korea"]),
    (14, "Tinggal setelah menikah / 결혼 후 거주지", "Belum tahu", LIVES),
    (16, "Mulai persiapan / 준비 시작일", "=TODAY()", None),
    (17, "Tanggal pernikahan (Hari H) / 결혼식 날짜", "=EDATE(C16,12)", None),
    (18, "Tanggal acara di negara kedua (opsional) / 두 번째 예식 날짜", None, None),
    (20, "Budget total / 총예산", 0, None),
    (21, "Mata uang budget / 예산 통화", "IDR", ["IDR", "KRW"]),
    (22, "Kurs: 1 KRW = ... IDR / 환율", 11.5, None),
]
for r, label, val, opts in inputs:
    ws.cell(row=r, column=2, value=label).font = F()
    c = ws.cell(row=r, column=3, value=val)
    c.fill, c.border, c.font = INPUT, BOX, F(color="0000FF")
    if opts:
        dv_list(ws, opts, f"C{r}")
for r in (16, 17, 18):
    ws[f"C{r}"].number_format = DATE_FMT
ws["C20"].number_format = "#,##0"
ws["C22"].number_format = "0.00"
for r, txt in [(3, "Pengantin wanita / 신부"), (7, "Pengantin pria / 신랑"), (11, "Pernikahan / 결혼"), (15, "Tanggal / 날짜"), (19, "Budget / 예산")]:
    ws.cell(row=r, column=2, value=txt).font = F(bold=True, size=12)
ws["B24"] = "Catatan: tanggal mulai & Hari H berisi formula contoh (hari ini, +12 bulan). Ganti dengan tanggal asli kalian. Kurs 11,5 hanya perkiraan: perbarui sesuai kurs saat ini."
ws["B24"].font = F(italic=True, size=9)
ws["B24"].alignment = WRAP
ws.merge_cells("B24:C25")
ws["C16"].comment = None

# Flags (rule engine). Column F name, G formula.
ws["F3"] = "Aturan otomatis (jangan diubah)"
ws["F3"].font = F(bold=True, size=9)
FL = [
    ("hasID", 'OR($C$5="Indonesia",$C$9="Indonesia")'),
    ("hasKR", 'OR($C$5="Korea",$C$9="Korea")'),
    ("hasOT", 'OR($C$5="Lainnya",$C$9="Lainnya")'),
    ("nonID", 'OR($C$5<>"Indonesia",$C$9<>"Indonesia")'),
    ("nonKR", 'OR($C$5<>"Korea",$C$9<>"Korea")'),
    ("mixedNat", '$C$5<>$C$9'),
    ("regID", 'OR($C$12="Indonesia",AND($C$12="Dua-duanya",$C$13<>"Korea"))'),
    ("regKR", 'OR($C$12="Korea",AND($C$12="Dua-duanya",$C$13="Korea"))'),
    ("cerID", '$C$12<>"Korea"'),
    ("cerKR", '$C$12<>"Indonesia"'),
    ("relSet", 'AND($C$6<>"",$C$10<>"")'),
    ("mixedRel", 'AND(f_relSet,$C$6<>$C$10)'),
    ("noRel", 'OR($C$6="Tidak beragama",$C$10="Tidak beragama")'),
    ("bothIslam", 'AND($C$6="Islam",$C$10="Islam")'),
    ("anyIslam", 'OR($C$6="Islam",$C$10="Islam")'),
    ("kua", 'AND(f_regID,f_bothIslam)'),
    ("dukcapil", 'AND(f_regID,f_relSet,NOT(f_bothIslam),NOT(f_mixedRel),NOT(f_noRel))'),
    ("relP", 'AND(f_relSet,NOT(f_mixedRel),$C$6="Kristen Protestan")'),
    ("relK", 'AND(f_relSet,NOT(f_mixedRel),$C$6="Katolik")'),
    ("relH", 'AND(f_relSet,NOT(f_mixedRel),$C$6="Hindu")'),
    ("relB", 'AND(f_relSet,NOT(f_mixedRel),$C$6="Buddha")'),
    ("relC", 'AND(f_relSet,NOT(f_mixedRel),$C$6="Khonghucu")'),
    ("needF6", 'AND(f_hasKR,f_nonKR,OR($C$14="Korea",$C$14="Belum tahu"))'),
    ("needKITAS", 'AND(f_hasID,f_nonID,OR($C$14="Indonesia",$C$14="Belum tahu"))'),
    ("reportKR", 'AND(f_regID,f_hasKR)'),
    ("travelFam", 'OR(AND(f_hasKR,f_cerID),AND(f_hasID,f_cerKR))'),
    ("idFamToKR", 'AND(f_hasID,f_cerKR)'),
    ("krFamToID", 'AND(f_hasKR,f_cerID)'),
    ("bilingual", 'AND(f_hasKR,f_hasID)'),
]
r = 4
for key, formula in FL:
    ws.cell(row=r, column=6, value=key).font = F(size=8, color="8C6B58")
    c = ws.cell(row=r, column=7, value="=" + formula)
    c.font = F(size=8, color="8C6B58")
    name("f_" + key, f"Mulai!$G${r}")
    r += 1
# date helpers
helpers = [
    ("start", "=$C$16", DATE_FMT), ("wedding", "=$C$17", DATE_FMT), ("date2", "=$C$18", DATE_FMT),
    ("ratio", "=MIN(1,MAX(1,$C$17-$C$16)/365)", "0.00"),
    ("second", '=IF($C$12="Dua-duanya",IF($C$13="Korea","ID","KR"),"")', "@"),
    ("rate", "=$C$22", "0.00"),
]
for key, formula, fmt in helpers:
    ws.cell(row=r, column=6, value=key).font = F(size=8, color="8C6B58")
    c = ws.cell(row=r, column=7, value=formula)
    c.font, c.number_format = F(size=8, color="8C6B58"), fmt
    name("f_" + key, f"Mulai!$G${r}")
    r += 1


# ------------------------------------------------------------------ condition -> Excel
def to_excel(src):
    if not src:
        return "TRUE"
    toks = re.findall(r"[A-Za-z_]\w*|&&|\|\||!|\(|\)", src)
    i = 0

    def peek():
        return toks[i] if i < len(toks) else None

    def nxt():
        nonlocal i
        i += 1
        return toks[i - 1]

    def prim():
        tk = nxt()
        if tk == "!":
            return f"NOT({prim()})"
        if tk == "(":
            a = orx()
            nxt()
            return a
        return "f_" + tk

    def andx():
        parts = [prim()]
        while peek() == "&&":
            nxt()
            parts.append(prim())
        return parts[0] if len(parts) == 1 else "AND(" + ",".join(parts) + ")"

    def orx():
        parts = [andx()]
        while peek() == "||":
            nxt()
            parts.append(andx())
        return parts[0] if len(parts) == 1 else "OR(" + ",".join(parts) + ")"

    return orx()


# ------------------------------------------------------------------ Quest
qs = wb.create_sheet("Quest")
cols = ["ID", "Area / 구역", "Quest", "퀘스트", "Untuk / 담당", "Berlaku? / 해당", "Deadline / 마감", "Sisa hari / D-day",
        "Status / 상태", "Peringatan / 알림", "XP", "Catatan / 메모", "Deadline manual (opsional)", "Penjelasan", "설명", "Link",
        "b", "m", "ev", "auto", "tight", "XP didapat"]
widths = [6, 18, 38, 30, 20, 10, 13, 10, 14, 12, 6, 30, 14, 60, 50, 28, 5, 5, 5, 12, 6, 8]
header(qs, 1, cols, widths)
qs.row_dimensions[1].height = 32
qs.freeze_panes = "D2"
area_name = {a["id"]: a["icon"] + " " + a["t"][0] + " / " + a["t"][1] for a in DATA["areas"]}
order = {a["id"]: i for i, a in enumerate(DATA["areas"])}
quests = sorted(DATA["quests"], key=lambda q: (order[q["a"]], -q["b"] if q["b"] > 0 else 10000 - q["b"]))
row = 2
for q in quests:
    r = row
    vals = [q["id"], area_name[q["a"]], q["t"][0], q["t"][1], WHO[q["w"]]]
    for ci, v in enumerate(vals, 1):
        qs.cell(row=r, column=ci, value=v)
    qs.cell(row=r, column=6, value=f'=IF({to_excel(q.get("c"))},"Ya","Tidak")')
    qs.cell(row=r, column=17, value=q["b"])
    qs.cell(row=r, column=18, value=q["m"])
    qs.cell(row=r, column=19, value=q.get("ev", ""))
    base = f'IF(AND(S{r}<>"",S{r}=f_second,f_date2<>""),f_date2,f_wedding)'
    qs.cell(row=r, column=20, value=f"=IF(Q{r}<=0,{base}-Q{r},MAX(f_start,{base}-MAX(R{r},ROUND(Q{r}*f_ratio,0))))")
    qs.cell(row=r, column=21, value=f"=AND(Q{r}>0,{base}-MAX(R{r},ROUND(Q{r}*f_ratio,0))<f_start)")
    qs.cell(row=r, column=7, value=f'=IF(M{r}<>"",M{r},T{r})')
    qs.cell(row=r, column=8, value=f'=IF(F{r}="Ya",G{r}-TODAY(),"")')
    qs.cell(row=r, column=9, value="Belum mulai")
    qs.cell(row=r, column=10, value=f'=IF(F{r}<>"Ya","",IF(AND(I{r}<>"Done",G{r}<TODAY()),"Telat",IF(AND(U{r},M{r}="",I{r}<>"Done"),"Mepet","")))')
    qs.cell(row=r, column=11, value=q["xp"])
    qs.cell(row=r, column=14, value=q["d"][0])
    qs.cell(row=r, column=15, value=q["d"][1])
    if q.get("l"):
        c = qs.cell(row=r, column=16, value=q["l"])
        c.hyperlink = q["l"]
    qs.cell(row=r, column=22, value=f'=IF(AND(F{r}="Ya",I{r}="Done"),K{r},0)')
    row += 1
last_q = row - 1
# blank rows for custom quests
qs.cell(row=row, column=3, value="⬇ Quest tambahan kalian (isi judul, area, dan deadline manual) / 직접 추가하는 퀘스트").font = F(bold=True, italic=True)
row += 1
custom_start = row
for r in range(custom_start, custom_start + 20):
    qs.cell(row=r, column=6, value=f'=IF(C{r}="","","Ya")')
    qs.cell(row=r, column=7, value=f'=IF(M{r}="","",M{r})')
    qs.cell(row=r, column=8, value=f'=IF(AND(C{r}<>"",M{r}<>""),M{r}-TODAY(),"")')
    qs.cell(row=r, column=10, value=f'=IF(AND(C{r}<>"",M{r}<>"",I{r}<>"Done",M{r}<TODAY()),"Telat","")')
    qs.cell(row=r, column=11, value=f'=IF(C{r}="","",20)')
    qs.cell(row=r, column=22, value=f'=IF(AND(C{r}<>"",I{r}="Done"),20,0)')
    for ci in (2, 3, 4, 5, 9, 12, 13):
        qs.cell(row=r, column=ci).fill = INPUT
last_row = custom_start + 19
for r in range(2, last_row + 1):
    for ci in range(1, 23):
        c = qs.cell(row=r, column=ci)
        c.font = F(size=9)
        c.border = BOX
        c.alignment = WRAP
    for ci in (7, 13, 20):
        qs.cell(row=r, column=ci).number_format = DATE_FMT
    qs.cell(row=r, column=8).number_format = '"D-"0;"D+"0;"D-DAY"'
    for ci in (9, 12, 13):
        if r <= last_q:
            qs.cell(row=r, column=ci).fill = INPUT
dv_list(qs, STATUSES, f"I2:I{last_row}")
dv_list(qs, [a["icon"] + " " + a["t"][0] + " / " + a["t"][1] for a in DATA["areas"]], f"B{custom_start}:B{last_row}")
for c in ("Q", "R", "S", "T", "U", "V"):
    qs.column_dimensions[c].hidden = True
qs.auto_filter.ref = f"A1:P{last_row}"
rng = f"A2:P{last_row}"
qs.conditional_formatting.add(rng, FormulaRule(formula=['$F2="Tidak"'], font=Font(name=FONT, color="BDB2A6"), fill=fill("F4F1EC"), stopIfTrue=True))
qs.conditional_formatting.add(rng, FormulaRule(formula=['$I2="Done"'], fill=fill("DDF3E0")))
qs.conditional_formatting.add(f"J2:J{last_row}", CellIsRule(operator="equal", formula=['"Telat"'], fill=fill("F5B5AC"), font=Font(name=FONT, bold=True, color="9C1C10")))
qs.conditional_formatting.add(f"J2:J{last_row}", CellIsRule(operator="equal", formula=['"Mepet"'], fill=fill("FCE3A0")))
for st, colr in [("On progress", "CDE8F8"), ("Pending", "FCE9B8"), ("Review", "E6DAF8"), ("Done", "9FDCA8")]:
    qs.conditional_formatting.add(f"I2:I{last_row}", CellIsRule(operator="equal", formula=[f'"{st}"'], fill=fill(colr)))

# ------------------------------------------------------------------ Budget
bs = wb.create_sheet("Budget")
bcols = ["Pengeluaran / 항목", "Kategori / 카테고리", "Mata uang / 통화", "Rencana / 예상", "Aktual / 실제", "Sudah dibayar / 지급",
         "Rencana (IDR)", "Aktual (IDR)", "Dibayar (IDR)", "Rencana (KRW)", "Aktual (KRW)", "Sisa bayar (IDR)", "Catatan / 메모"]
header(bs, 4, bcols, [28, 20, 10, 15, 15, 15, 16, 16, 16, 15, 15, 16, 30])
bs["A1"] = "💰 Kantong Gold / 골드 주머니"
bs["A1"].font = F(bold=True, size=14)
bs["A2"] = "Isi kolom A–F (kuning). Kolom IDR/KRW dihitung dari kurs di sheet Mulai. Baris 5 adalah contoh: hapus atau timpa."
bs["A2"].font = F(italic=True, size=9)
B0, B1 = 5, 154
for r in range(B0, B1 + 1):
    for ci in range(1, 7):
        bs.cell(row=r, column=ci).fill = INPUT
    bs.cell(row=r, column=13).fill = INPUT
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
ex = ["DP gedung (contoh)", "Venue", "IDR", 60000000, 60000000, 20000000]
for ci, v in enumerate(ex, 1):
    bs.cell(row=B0, column=ci, value=v)
dv_list(bs, BCATS, f"B{B0}:B{B1}")
dv_list(bs, ["IDR", "KRW"], f"C{B0}:C{B1}")
bs.freeze_panes = "B5"
name("b_planIDR", f"Budget!$G${B0}:$G${B1}")
name("b_actIDR", f"Budget!$H${B0}:$H${B1}")
name("b_paidIDR", f"Budget!$I${B0}:$I${B1}")
name("b_cat", f"Budget!$B${B0}:$B${B1}")

# ------------------------------------------------------------------ Tamu
gs = wb.create_sheet("Tamu")
gcols = ["Nama / 이름", "Pihak / 소속", "Kelompok / 그룹", "Negara / 국가", "Jumlah orang / 인원", "Akad/Pemberkatan / 혼례",
         "Resepsi Indonesia / 인니 피로연", "예식 Korea / 한국 예식", "Undangan terkirim / 청첩장", "Konfirmasi / 참석", "Kontak / 연락처", "Catatan / 메모"]
header(gs, 4, gcols, [26, 15, 13, 12, 12, 14, 14, 14, 13, 13, 18, 28])
gs["A1"] = "👥 Daftar Tamu / 하객 명단"
gs["A1"].font = F(bold=True, size=14)
gs["A2"] = "Satu baris per keluarga/rombongan. Kolom acara: isi Ya. Baris 5 contoh."
gs["A2"].font = F(italic=True, size=9)
G0, G1 = 5, 404
for r in range(G0, G1 + 1):
    for ci in range(1, 13):
        c = gs.cell(row=r, column=ci)
        c.fill, c.border, c.font = INPUT, BOX, F(size=9)
for ci, v in enumerate(["Keluarga Kim (contoh)", "Pihak wanita", "Keluarga", "Korea", 4, "Ya", "Ya", "", "Ya", "Hadir", "", "Datang dari Seoul"], 1):
    gs.cell(row=G0, column=ci, value=v)
dv_list(gs, ["Pihak wanita", "Pihak pria", "Bersama"], f"B{G0}:B{G1}")
dv_list(gs, ["Keluarga", "Teman", "Kantor", "Lainnya"], f"C{G0}:C{G1}")
dv_list(gs, NATS, f"D{G0}:D{G1}")
dv_list(gs, ["Ya", "Tidak"], f"F{G0}:I{G1}")
dv_list(gs, ["Belum jawab", "Hadir", "Tidak hadir"], f"J{G0}:J{G1}")
gs.freeze_panes = "B5"
gs.conditional_formatting.add(f"J{G0}:J{G1}", CellIsRule(operator="equal", formula=['"Hadir"'], fill=fill("9FDCA8")))
gs.conditional_formatting.add(f"J{G0}:J{G1}", CellIsRule(operator="equal", formula=['"Tidak hadir"'], fill=fill("F8BFCB")))
GR = lambda c: f"Tamu!${c}${G0}:${c}${G1}"

# ------------------------------------------------------------------ Vendor
vs = wb.create_sheet("Vendor")
vcols = ["Nama vendor / 업체명", "Kategori / 카테고리", "Status / 상태", "Narahubung / 담당자", "Telepon / 전화", "Email / Instagram",
         "Harga / 가격", "Mata uang / 통화", "Harga (IDR)", "Jatuh tempo / 잔금 기한", "Rating (1-5)", "Catatan / 메모"]
header(vs, 4, vcols, [26, 22, 13, 18, 18, 22, 15, 10, 16, 14, 10, 32])
vs["A1"] = "🛍️ Daftar Vendor / 업체 목록"
vs["A1"].font = F(bold=True, size=14)
vs["A2"] = "Simpan semua vendor yang dilirik, termasuk yang belum jadi. Baris 5 contoh."
vs["A2"].font = F(italic=True, size=9)
V0, V1 = 5, 104
for r in range(V0, V1 + 1):
    for ci in range(1, 13):
        c = vs.cell(row=r, column=ci)
        c.border, c.font = BOX, F(size=9)
        if ci != 9:
            c.fill = INPUT
    vs.cell(row=r, column=7).number_format = "#,##0"
    vs.cell(row=r, column=9, value=f'=IF(A{r}="","",IF(H{r}="KRW",G{r}*f_rate,N(G{r})))').number_format = IDR_FMT
    vs.cell(row=r, column=10).number_format = DATE_FMT
    vs.cell(row=r, column=5).number_format = "@"
for ci, v in enumerate(["Studio Contoh", "Foto & video", "Negosiasi", "Mbak Rina", "+62 812 0000 0000", "@studiocontoh", 25000000, "IDR", None, None, 4, "Paket 8 jam + album"], 1):
    if v is not None:
        vs.cell(row=V0, column=ci, value=v)
dv_list(vs, VCATS, f"B{V0}:B{V1}")
dv_list(vs, VSTAT, f"C{V0}:C{V1}")
dv_list(vs, ["IDR", "KRW"], f"H{V0}:H{V1}")
dv_list(vs, ["1", "2", "3", "4", "5"], f"K{V0}:K{V1}")
vs.freeze_panes = "B5"
vs.conditional_formatting.add(f"C{V0}:C{V1}", CellIsRule(operator="equal", formula=['"Lunas"'], fill=fill("9FDCA8")))
vs.conditional_formatting.add(f"C{V0}:C{V1}", FormulaRule(formula=[f'OR(C{V0}="Booked",C{V0}="DP dibayar")'], fill=fill("D6F0EA")))

# ------------------------------------------------------------------ Dashboard
ds = wb.create_sheet("Dashboard", 1)
ds.sheet_view.showGridLines = False
for c, w in zip("ABCDEFG", [3, 34, 16, 16, 16, 14, 4]):
    ds.column_dimensions[c].width = w
ds["B1"] = '=IF(OR(Mulai!C4="",Mulai!C8=""),"💍 Wedding Quest","💍 "&Mulai!C4&" & "&Mulai!C8)'
ds["B1"].font = F(bold=True, size=18)
ds["B2"] = "Semua angka di sini dihitung otomatis dari sheet lain. / 모든 숫자는 자동 계산돼요."
ds["B2"].font = F(italic=True, size=9)
QF = f"Quest!$F$2:$F${last_row}"
QI = f"Quest!$I$2:$I${last_row}"
QB = f"Quest!$B$2:$B${last_row}"
QJ = f"Quest!$J$2:$J${last_row}"
QH = f"Quest!$H$2:$H${last_row}"
QK = f"Quest!$K$2:$K${last_row}"
QV = f"Quest!$V$2:$V${last_row}"
kpis = [
    ("Hari H / D-Day", "=f_wedding", DATE_FMT),
    ("Hitung mundur / 남은 날", '=IF(f_wedding="","",f_wedding-TODAY())', '"D-"0;"D+"0;"D-DAY"'),
    ("Quest aktif / 활성 퀘스트", f'=COUNTIF({QF},"Ya")', "0"),
    ("Quest selesai / 완료", f'=COUNTIFS({QF},"Ya",{QI},"Done")', "0"),
    ("Progress / 진행률", "=IF(C6=0,0,C7/C6)", "0%"),
    ("Quest telat / 지연", f'=COUNTIF({QJ},"Telat")', "0"),
    ("Jatuh tempo 30 hari ke depan / 30일 이내", f'=COUNTIFS({QF},"Ya",{QI},"<>Done",{QH},">=0",{QH},"<=30")', "0"),
    ("XP didapat / 획득 XP", f"=SUM({QV})", "#,##0"),
    ("XP total / 전체 XP", f'=SUMIF({QF},"Ya",{QK})', "#,##0"),
    ("Level", '=MIN(8,INT(IF(C12=0,0,C11/C12)*8)+1)', "0"),
    ("Gelar / 칭호", "=INDEX($I$4:$I$11,C13)", "@"),
]
ds["B3"] = "Progress"
ds["B3"].font = F(bold=True, size=12)
for i, (lab, f, fmt) in enumerate(kpis):
    r = 4 + i
    ds.cell(row=r, column=2, value=lab).font = F()
    c = ds.cell(row=r, column=3, value=f)
    c.font, c.number_format, c.fill, c.border = F(bold=True, size=11), fmt, CREAM, BOX
ds.conditional_formatting.add("C9", CellIsRule(operator="greaterThan", formula=["0"], fill=fill("F5B5AC"), font=Font(name=FONT, bold=True, color="9C1C10")))
# level names (reference list)
ds["I3"] = "Daftar gelar"
ds["I3"].font = F(size=8, color="8C6B58")
for i, lv in enumerate(LEVELS):
    ds.cell(row=4 + i, column=9, value=lv).font = F(size=8, color="8C6B58")
ds.column_dimensions["I"].width = 34

r = 17
ds.cell(row=r, column=2, value="Peta area / 구역별 진행").font = F(bold=True, size=12)
r += 1
for ci, lab in enumerate(["Area / 구역", "Aktif", "Selesai", "Progress", "Telat"], 2):
    c = ds.cell(row=r, column=ci, value=lab)
    c.font, c.fill, c.border, c.alignment = F(bold=True), HEAD, BOX, CENTER
for a in DATA["areas"]:
    r += 1
    an = area_name[a["id"]]
    ds.cell(row=r, column=2, value=an)
    ds.cell(row=r, column=3, value=f'=COUNTIFS({QB},B{r},{QF},"Ya")')
    ds.cell(row=r, column=4, value=f'=COUNTIFS({QB},B{r},{QF},"Ya",{QI},"Done")')
    ds.cell(row=r, column=5, value=f"=IF(C{r}=0,0,D{r}/C{r})").number_format = "0%"
    ds.cell(row=r, column=6, value=f'=COUNTIFS({QB},B{r},{QJ},"Telat")')
    for ci in range(2, 7):
        ds.cell(row=r, column=ci).border = BOX
        ds.cell(row=r, column=ci).font = F()
area_first, area_last = r - len(DATA["areas"]) + 1, r
from openpyxl.formatting.rule import DataBarRule
ds.conditional_formatting.add(f"E{area_first}:E{area_last}", DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1, color="6CC04F"))

r += 2
ds.cell(row=r, column=2, value="Budget / 예산").font = F(bold=True, size=12)
r += 1
for ci, lab in enumerate(["", "IDR", "KRW"], 2):
    c = ds.cell(row=r, column=ci, value=lab)
    c.font, c.fill, c.border, c.alignment = F(bold=True), HEAD, BOX, CENTER
budget_rows = [
    ("Budget total / 총예산", '=IF(Mulai!$C$21="KRW",Mulai!$C$20*f_rate,Mulai!$C$20)'),
    ("Rencana / 예상", "=SUM(b_planIDR)"),
    ("Terpakai / 실제", "=SUM(b_actIDR)"),
    ("Sudah dibayar / 지급", "=SUM(b_paidIDR)"),
]
b_first = r + 1
for lab, f in budget_rows:
    r += 1
    ds.cell(row=r, column=2, value=lab).font = F()
    ds.cell(row=r, column=3, value=f).number_format = IDR_FMT
    ds.cell(row=r, column=4, value=f"=IF(f_rate=0,0,C{r}/f_rate)").number_format = KRW_FMT
r += 1
ds.cell(row=r, column=2, value="Sisa budget / 남은 예산").font = F(bold=True)
ds.cell(row=r, column=3, value=f"=C{b_first}-MAX(C{b_first + 2},0)").number_format = IDR_FMT
ds.cell(row=r, column=4, value=f"=IF(f_rate=0,0,C{r}/f_rate)").number_format = KRW_FMT
ds.conditional_formatting.add(f"C{r}:D{r}", CellIsRule(operator="lessThan", formula=["0"], fill=fill("F5B5AC")))
for rr in range(b_first, r + 1):
    for ci in (2, 3, 4):
        ds.cell(row=rr, column=ci).border = BOX
        if ci > 2:
            ds.cell(row=rr, column=ci).font = F(bold=True)

r += 2
ds.cell(row=r, column=2, value="Tamu / 하객 (orang / 명)").font = F(bold=True, size=12)
guest_rows = [
    ("Total diundang / 초대 인원", f"=SUM({GR('E')})"),
    ("Hadir / 참석", f'=SUMIF({GR("J")},"Hadir",{GR("E")})'),
    ("Belum jawab / 미응답", f'=SUM({GR("E")})-SUMIF({GR("J")},"Hadir",{GR("E")})-SUMIF({GR("J")},"Tidak hadir",{GR("E")})'),
    ("Tidak hadir / 불참", f'=SUMIF({GR("J")},"Tidak hadir",{GR("E")})'),
    ("Pihak wanita / 신부 측", f'=SUMIF({GR("B")},"Pihak wanita",{GR("E")})'),
    ("Pihak pria / 신랑 측", f'=SUMIF({GR("B")},"Pihak pria",{GR("E")})'),
    ("Dari Korea / 한국에서", f'=SUMIF({GR("D")},"Korea",{GR("E")})'),
    ("Dari Indonesia / 인도네시아에서", f'=SUMIF({GR("D")},"Indonesia",{GR("E")})'),
    ("Undangan sudah dikirim / 청첩장 발송", f'=SUMIF({GR("I")},"Ya",{GR("E")})'),
]
for lab, f in guest_rows:
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
    ("📖 Cara pakai / 사용법", True),
    ("1. Isi sel kuning di sheet Mulai: nama, kewarganegaraan, agama, lokasi, tanggal, budget, kurs.", False),
    ("2. Buka sheet Quest. Kolom 'Berlaku?' otomatis Ya/Tidak sesuai pilihan kalian. Pakai filter di header untuk menampilkan yang Ya saja.", False),
    ("3. Ubah Status lewat dropdown: Belum mulai → On progress → Pending → Review → Done. Quest Done memberi XP.", False),
    ("4. Deadline dihitung mundur dari Hari H (rencana 12 bulan, dipadatkan otomatis kalau waktunya lebih pendek). Isi 'Deadline manual' kalau mau mengganti.", False),
    ("5. 'Mepet' artinya deadline ideal sudah lewat sebelum persiapan dimulai: kerjakan secepatnya. 'Telat' artinya lewat deadline dan belum Done.", False),
    ("6. Budget, Tamu, Vendor: isi baris kuning. Konversi IDR ↔ KRW memakai kurs di sheet Mulai.", False),
    ("7. Mau dipakai berdua? Upload file ini ke Google Drive, buka dengan Google Sheets, lalu Share ke pasangan sebagai Editor.", False),
    ("", False),
    ("1. Mulai 시트의 노란 칸(이름, 국적, 종교, 장소, 날짜, 예산, 환율)을 채우세요.", False),
    ("2. Quest 시트의 'Berlaku? / 해당' 열이 자동으로 Ya(해당)/Tidak(해당 없음)으로 바뀌어요. 필터로 Ya만 보세요.", False),
    ("3. 상태는 드롭다운으로: Belum mulai(시작 전) → On progress(진행 중) → Pending(보류) → Review(검토) → Done(완료).", False),
    ("4. 마감일은 D-Day 기준으로 자동 계산돼요. 직접 정하려면 'Deadline manual' 칸에 날짜를 넣으세요.", False),
    ("5. 둘이 함께 쓰려면 Google 드라이브에 올려 Google 스프레드시트로 열고 상대방을 편집자로 공유하세요.", False),
    ("", False),
    ("⚠️ Tentang syarat dokumen / 서류 요건 안내", True),
    ("Syarat dirangkum dari peraturan dan sumber publik per November 2026 (PMA 30/2024, SEMA 2/2023, aturan apostille Indonesia sejak 4 Juni 2022, info visa F-6). "
     "Syarat bisa berbeda per kota dan sering berubah. Selalu konfirmasi ke KUA / Dukcapil, 구청, Kedutaan Besar Korea di Jakarta, dan KBRI Seoul sebelum mengurus.", False),
    ("2026년 11월 기준 공개 자료를 정리한 것으로, 지역마다 다르고 자주 바뀌어요. 반드시 관할 기관과 대사관에 먼저 확인하세요.", False),
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
    c.font = F(bold=bold, size=13 if bold else 10)
    c.alignment = WRAP
    r += 1
for u in sources:
    c = ps.cell(row=r, column=2, value=u)
    c.hyperlink = u
    c.font = F(size=9, color="8559BF", underline="single")
    r += 1

# tab colors & order
for sheet, colr in [("Mulai", "F5C34E"), ("Dashboard", "F39AAE"), ("Quest", "A47AD8"), ("Budget", "F5C34E"), ("Tamu", "A3DCCF"), ("Vendor", "8FC7EA"), ("Panduan", "C9B29A")]:
    wb[sheet].sheet_properties.tabColor = colr
wb._sheets = [wb[n] for n in ["Mulai", "Dashboard", "Quest", "Budget", "Tamu", "Vendor", "Panduan"]]
wb.active = 0
wb.save(OUT)
print("saved", OUT, "quests:", len(quests))
