
import win32com.client
import time
from pathlib import Path

# ===== 設定 =====
########################################################
# ここを自分のフォルダ階層に設定
BASE_DIR  = Path(r"c:\@VScode\260325_mergeDWG") 
########################################################
A_DIR     = BASE_DIR / "merge_a"
B_DIR     = BASE_DIR / "merge_b"
OUT_DIR   = BASE_DIR / "output" # ＜--- 自動でフォルダ生成される
COMPARE_WAIT = 15           # COMPARE の完了待ち秒数（図面サイズに応じて調整）
# ================


def connect_autocad():
    """起動中の AutoCAD COM オブジェクトを取得する"""
    try:
        acad = win32com.client.GetActiveObject("AutoCAD.Application")
        acad.Visible = True
        return acad
    except Exception:
        raise RuntimeError(
            "AutoCAD が見つかりません。AutoCAD を起動してから再実行してください。"
        )


def send(doc, cmd: str, wait: float = 1.0):
    """AutoCAD にコマンドを送信して wait 秒待機する"""
    doc.SendCommand(cmd)
    time.sleep(wait)


def wait_idle(acad, timeout: float = 60.0):
    """AutoCAD がコマンド待機状態になるまで最大 timeout 秒待つ"""
    deadline = time.time() + timeout
    while time.time() < deadline:
        try:
            if acad.GetAcadState().IsQuiescent:
                return True
        except Exception:
            pass
        time.sleep(3)
    return False


def process_pair(acad, file_a: Path, file_b: Path, out_dir: Path):
    """1 ペアを処理: 開く → 比較 → DWG 保存 → 閉じる"""

    print(f"  [開く]    {file_a.name}")
    doc = acad.Documents.Open(str(file_a))
    time.sleep(10)
    wait_idle(acad, 30)

    send(doc, "FILEDIA\n0\n", 3.5)

    # 全体表示
    send(doc, "ZOOM\nE\n", 2.0)

    # ---- DWG 比較 ----
    print(f"  [比較]    {file_b.name}")
    # FILEDIA=0 のとき COMPARE はダイアログを出さずコマンドラインでパスを受け取る
    send(doc, f'COMPARE\n"{file_b}"\n', COMPARE_WAIT)
    wait_idle(acad, COMPARE_WAIT + 5)

    # 比較後に全体表示
    send(doc, "ZOOM\nE\n", 2.0)

    # ---- 出力ファイル名 ----
    base = f"{file_a.stem}_vs_{file_b.stem}"

    # ---- 比較結果を COMPAREEXPORT で DWG として書き出し ----
    dwg_path = out_dir / f"{base}.dwg"
    print(f"  [DWG]     {dwg_path.name}")
    # COMPAREEXPORT のプロンプト順序:
    send(doc, f'COMPAREEXPORT\n\n"{dwg_path}"\ny\n1\n', 8.0)
    wait_idle(acad, 15)

    # ファイルダイアログを元に戻す
    send(doc, "FILEDIA\n1\n", 3.5)

    # ドキュメントを閉じる（元ファイルは保存しない）
    doc.Close(False)
    time.sleep(10)

    print(f"  [完了]    -> {dwg_path.parent.name}/{dwg_path.name}")


def pair_files(a_files: list, b_files: list) -> list:
    """ファイルをソート順にペアリングする"""
    a_sorted = sorted(a_files)
    b_sorted = sorted(b_files)
    pairs = list(zip(a_sorted, b_sorted))
    if len(a_files) != len(b_files):
        print(f"  警告: ファイル数が異なります "
              f"(merge_a: {len(a_files)}, merge_b: {len(b_files)})")
        print(f"  先頭 {len(pairs)} ペアのみ処理します")
    return pairs


def main():
    OUT_DIR.mkdir(exist_ok=True)

    a_files = list(A_DIR.glob("*.dwg"))
    b_files = list(B_DIR.glob("*.dwg"))

    if not a_files:
        print(f"エラー: {A_DIR} に DWG ファイルがありません")
        return
    if not b_files:
        print(f"エラー: {B_DIR} に DWG ファイルがありません")
        return

    pairs = pair_files(a_files, b_files)
    total = len(pairs)

    print(f"\n{total} ペアを処理します")
    print(f"出力先: {OUT_DIR}\n")
    print("-" * 50)

    acad = connect_autocad()
    print(f"AutoCAD 接続完了: {acad.Name} {acad.Version}\n")

    errors = []
    for i, (fa, fb) in enumerate(pairs, 1):
        print(f"\n[{i}/{total}] {fa.name}  <->  {fb.name}")
        try:
            process_pair(acad, fa, fb, OUT_DIR)
        except Exception as e:
            msg = f"[{i}] {fa.name}: {e}"
            print(f"  エラー: {e}")
            errors.append(msg)

    print("\n" + "=" * 50)
    print(f"処理完了: {total - len(errors)}/{total} 成功")
    if errors:
        print("エラー一覧:")
        for err in errors:
            print(f"  {err}")
    print(f"出力フォルダ: {OUT_DIR}")


if __name__ == "__main__":
    main()