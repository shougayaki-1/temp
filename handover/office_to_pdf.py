import os
import win32com.client
from tqdm import tqdm

# ==========================================
# 設定エリア
# ==========================================
# ダウンロードしたフォルダの「フルパス」を指定してください
# 例: r"C:\Users\Name\Downloads\DriveDownload"
TARGET_FOLDER = r"C:\Users\あなたのユーザー名\Desktop\ドライブダウンロードフォルダ"

# ==========================================
# 変換ロジック
# ==========================================

def convert_office_to_pdf(base_folder):
    # アプリケーションのインスタンス初期化
    word = None
    excel = None
    ppt = None

    # ファイルリストの作成
    files_to_convert = []
    for root, dirs, files in os.walk(base_folder):
        for file in files:
            ext = file.lower().split('.')[-1]
            if ext in ['docx', 'doc', 'xlsx', 'xls', 'pptx', 'ppt']:
                files_to_convert.append(os.path.join(root, file))

    print(f"変換対象ファイル数: {len(files_to_convert)}")
    
    try:
        # アプリケーションを起動（高速化のためループの外で起動）
        try:
            word = win32com.client.Dispatch("Word.Application")
            word.Visible = False
        except: pass

        try:
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False # 保存時の警告などを無視
        except: pass

        try:
            ppt = win32com.client.Dispatch("PowerPoint.Application")
            # PPTは仕様上、ウィンドウを表示しないと動かない場合があるが、最小化などで対応
        except: pass

        # ループ処理
        for file_path in tqdm(files_to_convert, desc="PDF変換中"):
            pdf_path = os.path.splitext(file_path)[0] + ".pdf"
            
            # すでにPDFがある場合はスキップ
            if os.path.exists(pdf_path):
                continue

            ext = file_path.lower().split('.')[-1]
            abs_path = os.path.abspath(file_path) # COMは絶対パス必須
            abs_pdf_path = os.path.abspath(pdf_path)

            try:
                # --- Word ---
                if ext in ['docx', 'doc'] and word:
                    doc = word.Documents.Open(abs_path)
                    # 17 = wdFormatPDF
                    doc.SaveAs(abs_pdf_path, FileFormat=17)
                    doc.Close()

                # --- Excel ---
                elif ext in ['xlsx', 'xls'] and excel:
                    wb = excel.Workbooks.Open(abs_path)
                    # シートを全選択してPDF化（これがないとアクティブシートしか変換されない）
                    wb.WorkSheets.Select() 
                    # 0 = xlTypePDF
                    wb.ActiveSheet.ExportAsFixedFormat(0, abs_pdf_path)
                    wb.Close(False)

                # --- PowerPoint ---
                elif ext in ['pptx', 'ppt'] and ppt:
                    presentation = ppt.Presentations.Open(abs_path, WithWindow=False)
                    # 32 = ppSaveAsPDF
                    presentation.SaveAs(abs_pdf_path, 32)
                    presentation.Close()

            except Exception as e:
                print(f"\n[エラー] {os.path.basename(file_path)}: {e}")
                continue

    finally:
        # 後始末：必ずアプリを終了させる
        print("アプリケーションを終了しています...")
        if word: word.Quit()
        if excel: excel.Quit()
        if ppt: ppt.Quit()

if __name__ == "__main__":
    convert_office_to_pdf(TARGET_FOLDER)