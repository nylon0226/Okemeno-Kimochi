# ========================
# 必要なライブラリのインポート
# ========================

# Excelを制御するCOM初期化用（安全な運用のためにに使う）
import pythoncom

# 実行時にExcelファイルパスを受け取るため
import sys

# 体重などの数値を文字列から抽出するための正規表現処理
import re

# Excelアプリケーションを操作するためのCOMライブラリ（win32経由でExcelを開く・書き込む）
import win32com.client as win32 #COMは現在起動中のExcelを直接操作可能（open pyxl では不可）

# 日付関連処理（日付の加算、文字列→日付変換など）
from datetime import date, timedelta, datetime

# ==============================
# 定数設定（全体仕様の根幹）
# ==============================

# 対象のExcelシート名、データ開始行、生後追跡日数（91日分 = 0〜90日）
SHEET, START, SPAN = "オカメの気持ち", 6, 90

# 使用する列番号を定義（1始まり、Excel列と一致）
# A:日付 / B:体重 / C:平均 / D:3日平均 / E:5日平均 / F:判定 / G:備考
COL_DATE, COL_WT, COL_AVG, COL_A3, COL_A5, COL_JUDGE, COL_NOTE = range(1, 8)

# さし餌回数変更の判定条件設定
MIN_DAY_3TO2 = 41       # 3→2回への判定を始められる生後日数の最小値
TOLERANCE = 2.0         # 平均体重との差の許容誤差（−2g以内）
REQ_3TO2 = 3            # 3→2判定に必要な連続日数
REQ_2TO1 = 5            # 2→1判定に必要な連続日数
REQ_1TO0 = 7            # 1→0（完全自立）判定に必要な連続日数

# 前回処理時の生年月日を保存するセル（リセット防止）
CELL_PREV_BD = "Z1"

# ==============================
# 日付・体重・列出力ユーティリティ関数
# ==============================

# Python日付をExcelのシリアル日付に変換する（1900年起点）
def as_serial(d): return (d - date(1899, 12, 30)).days

# セルの値を日付型に変換（datetime, int, float, string どれでもOK）
def to_date(v):
    if isinstance(v, (date, datetime)): return v.date() if isinstance(v, datetime) else v
    if isinstance(v, (int, float)): return date(1899, 12, 30) + timedelta(days=int(v))
    try: return datetime.fromisoformat(str(v).strip().replace("/", "-")).date()
    except: return None

# セルの値から体重（数値）を抽出する（"80g" や "８１．０ｇ" でもOK）
def parse_wt(v):
    if isinstance(v, (int, float)): return float(v)
    s = str(v).translate(str.maketrans("０１２３４５６７８９．，", "0123456789.."))
    m = re.search(r'[-+]?\d+(?:\.\d+)?', s)
    return float(m.group()) if m else None

# 任意の列にリストの値を書き込む（文字列として書き込み、書式は文字列）
def write_col(ws, r0, col, vals):
    rng = ws.Range(ws.Cells(r0, col), ws.Cells(r0 + len(vals) - 1, col))
    rng.Value = tuple(("" if v is None else str(v),) for v in vals)
    rng.NumberFormatLocal = "@"

# 指定行の備考欄の背景色を赤にする or 解除する
def set_red(ws, row, on): ws.Cells(row, COL_NOTE).Interior.ColorIndex = 3 if on else 0

# ==============================
# メイン処理本体
# ==============================

def main():
    if len(sys.argv) < 2: raise RuntimeError("パス未指定")  # Excelファイルのパスが指定されていない場合は終了
    path = sys.argv[1]
    pythoncom.CoInitialize()  # COM初期化（Excel制御に必要）
    xl = win32.Dispatch("Excel.Application")
    wb = xl.Workbooks.Open(path)
    ws = wb.Worksheets(SHEET)

    try:
        # B1セル（生年月日）とZ1セル（前回記録）を日付に変換
        bd, prev_bd = to_date(ws.Range("B1").Value), to_date(ws.Range(CELL_PREV_BD).Value)
        if not bd: raise RuntimeError("B1（生年月日）不正")

        # B3セルに入力された下限体重を取得
        low = parse_wt(ws.Range("B3").Value)

        # 初回 or 生年月日が更新された場合、A列に日付を再生成
        if bd != prev_bd:
            dates = [as_serial(bd + timedelta(days=i)) for i in range(SPAN + 1)]  # 生後0日〜90日までの日付
            rng = ws.Range(ws.Cells(START, COL_DATE), ws.Cells(START + SPAN, COL_DATE))
            rng.Value, rng.NumberFormatLocal = tuple((d,) for d in dates), "yyyy/mm/dd"
            ws.Range("B2").Value = dates[-1]; ws.Range("B2").NumberFormatLocal = "yyyy/mm/dd"
            ws.Range(CELL_PREV_BD).Value = bd.strftime("%Y-%m-%d")

        # B列（体重列）からデータを取得・数値化
        vals = ws.Range(ws.Cells(START, COL_WT), ws.Cells(START + SPAN, COL_WT)).Value
        wts = [parse_wt(v[0]) for v in vals]

        # 各種平均の準備
        cum, a3s, a5s, s, c = [], [], [], 0.0, 0

# ==============================
# 繰り返し処理：平均計算とさし餌判定
# ==============================

        # 各日付に対し：体重が存在する場合は累積平均、3日平均、5日平均を算出
        for i, w in enumerate(wts):
            if w is not None:
                s += w
                c += 1
            # 累積平均（最初の行からその行まで）
            cum.append(round(s/c,1) if c else None)
            # 直近3日間の体重が全て存在するなら3日平均
            a3s.append(round(sum(wts[i-j] for j in range(3))/3,1) if i>=2 and all(wts[i-j] is not None for j in range(3)) else None)
            # 直近5日間の体重が全て存在するなら5日平均
            a5s.append(round(sum(wts[i-j] for j in range(5))/5,1) if i>=4 and all(wts[i-j] is not None for j in range(5)) else None)

        # 各平均を対応する列に出力（C〜E列）
        write_col(ws, START, COL_AVG, [f"{v:.1f}" if v else None for v in cum])
        write_col(ws, START, COL_A3,  [f"{v:.1f}" if v else None for v in a3s])
        write_col(ws, START, COL_A5,  [f"{v:.1f}" if v else None for v in a5s])

        # f: 現在のさし餌回数、override: 一時的な上書き指示
        # flags3: 3→2の条件を満たした日 / flags2: 2→1の条件を満たした日 / flags1: 1→0の条件を満たした日 / reds: 赤色警告を出した日
        f, override, flags3, flags2, flags1, reds = 3, None, [], [], [], []

        # 全行に対して、体重・平均値・過去判定状態をもとにさし餌回数の候補を判定する
        for i, (w, a3, a5, avg) in enumerate(zip(wts, a3s, a5s, cum)):
            r, msg, note = START + i, "", ""

            # 強制的な戻し指示が前回で出ていた場合、ここで適用
            if override is not None:
                f, override = override, None

            # 生後41日目の行に「さし餌減少判断開始日」を記録する
            if i == 40:
                cell = ws.Cells(r, COL_NOTE)
                if not cell.Value or "さし餌減少判断開始日" not in str(cell.Value):
                    cell.Value = "さし餌減少判断開始日" if not cell.Value else f"さし餌減少判断開始日 / {cell.Value}"

            # データが存在しない日はスキップし、Falseフラグで埋めておく
            if w is None:
                flags3.append(False); flags2.append(False); flags1.append(False); reds.append(False)
                continue

            # 2回以下のさし餌で下限体重を下回ったら、翌日から1回増加＋赤アラート
            if low is not None and f <= 2 and w < low:
                msg, note, override = "戻す", "体重低下 → 明日からさし餌１回増加", 3
                set_red(ws, r, True)
                flags3.append(False); flags2.append(False); flags1.append(False); reds.append(True)
                ws.Cells(r, COL_JUDGE).Value = msg
                ws.Cells(r, COL_NOTE).Value = note
                continue

            # 通常処理：赤判定ではない行
            reds.append(False)

            # 3→2判定条件（従来通り：3日平均と累積平均を比較）
            ok3 = (
                i >= MIN_DAY_3TO2 and
                all(x is not None for x in [a3, avg, w]) and
                a3 >= avg - TOLERANCE and
                w  >= low
            )

            # 2→1判定条件（差し餌2回モードで体重が「累積平均−2」以上）
            ok2 = (
                f == 2 and
                w is not None and avg is not None and
                w >= avg - TOLERANCE
            )

            # 1→0判定条件（差し餌1回モードで体重が「累積平均−2」以上）
            ok1 = (
                f == 1 and
                w is not None and avg is not None and
                w >= avg - TOLERANCE
            )

            flags3.append(ok3)
            flags2.append(ok2)
            flags1.append(ok1)

            # 直近k日間の条件が連続してTrueであるかを確認する関数
            def last_k_true(lst, k): return len(lst) >= k and all(lst[-k:])

            # 置き餌用ペレット量（10%換算）を算出
            pellet = f"{round(w * 0.1,1)}g"

            # 2→1の切り替え条件を連続5日満たしたら
            if f == 2 and last_k_true(flags2, REQ_2TO1):
                msg, note, override = "2→1候補", f"明日からさし餌1回 / 置き餌 {pellet}", 1
                set_red(ws, r, False)

            # 1→0の切り替え条件を連続7日満たしたら（完全自立）
            elif f == 1 and last_k_true(flags1, REQ_1TO0):
                msg, note, override = "1→0候補", "明日から差し餌なし（完全自立）", 0
                set_red(ws, r, False)

            # 3→2の切り替え条件を連続3日満たしたら
            elif f >= 3 and last_k_true(flags3, REQ_3TO2):
                msg, note, override = "3→2候補", f"明日からさし餌2回 / 置き餌 {pellet}", 2
                set_red(ws, r, False)

            # 条件未満の場合は赤色解除のみ
            else:
                set_red(ws, r, False)

            # 判定結果を書き込み
            if msg: ws.Cells(r, COL_JUDGE).Value = msg
            if note: ws.Cells(r, COL_NOTE).Value = note

        # ファイル保存
        wb.Save()
    finally:
        try: xl.StatusBar = False
        except: pass
        pythoncom.CoUninitialize()

if __name__ == "__main__":
    main()





























