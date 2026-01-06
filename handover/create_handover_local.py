import os
import time
import google.generativeai as genai

# ==========================================
# 設定エリア
# ==========================================
API_KEY = "AIzaSyDmFvm0vje2LJuxCPP79AwGJxRlRKMNspY"  # ここにキーを入れる
PDF_FOLDER_PATH = "./handover_docs" # PDFが入っているローカルフォルダのパス
MODEL_NAME = "gemini-2.5-flash-lite"

# APIキーの設定
genai.configure(api_key=API_KEY)

def upload_files(folder_path):
    """フォルダ内のPDFをGeminiのFile APIにアップロードする"""
    uploaded_files = []
    
    print(f"フォルダ '{folder_path}' 内のファイルをスキャン中...")
    
    for filename in os.listdir(folder_path):
        if filename.lower().endswith(".pdf"):
            file_path = os.path.join(folder_path, filename)
            print(f"アップロード中: {filename} ...")
            
            try:
                # Geminiへファイルをアップロード
                # mime_typeを指定すると処理がスムーズです
                uploaded_file = genai.upload_file(path=file_path, mime_type="application/pdf")
                uploaded_files.append(uploaded_file)
            except Exception as e:
                print(f"エラーが発生しました ({filename}): {e}")

    print(f"\n合計 {len(uploaded_files)} 個のファイルをアップロードしました。")
    return uploaded_files

def wait_for_files_active(files):
    """ファイルが処理完了(ACTIVE)になるまで待機する"""
    print("ファイルの処理完了を待機しています...")
    for file in files:
        current_file = genai.get_file(file.name)
        while current_file.state.name == "PROCESSING":
            print(".", end="", flush=True)
            time.sleep(2)
            current_file = genai.get_file(file.name)
        
        if current_file.state.name != "ACTIVE":
            print(f"\n警告: ファイル {file.display_name} の処理に失敗しました。State: {current_file.state.name}")
            
    print("\n全ファイルの準備が完了しました。")

def create_handover_material():
    # 1. ファイルをアップロード
    files_to_analyze = upload_files(PDF_FOLDER_PATH)
    
    if not files_to_analyze:
        print("PDFファイルが見つかりませんでした。パスを確認してください。")
        return

    # 2. 処理完了待ち
    wait_for_files_active(files_to_analyze)

    # 3. モデルの準備
    # system_instructionで役割を与えます
    model = genai.GenerativeModel(
        model_name=MODEL_NAME,
        system_instruction="あなたは優秀なプロジェクトマネージャー兼ドキュメント作成者です。提供された資料を元に引き継ぎ資料を作成します。"
    )

    # 4. プロンプト作成
    prompt = """
    アップロードされた全てのPDFファイルは、Googleドライブからダウンロードした業務資料です。
    これらを分析し、次年度担当者への「業務引き継ぎ資料」を作成してください。

    以下のフォーマットで出力してください。

    # 1. 業務概要とプロジェクト状況
    資料から読み取れる全プロジェクトのステータス、目的、主要な成果物をまとめてください。

    # 2. キーマンと関係者
    登場する人物や組織、およびその関係性を整理してください。

    # 3. 作成者（私）へのインタビューリスト
    ここが最も重要です。資料だけでは読み取れない「背景」「経緯」「苦労した点」「未解決のトラブル」を補完するため、
    私に対して行うべき具体的な質問を15個〜20個リストアップしてください。
    
    回答は詳細な日本語でお願いします。
    """

    print("\nAIが思考中...（資料量により数分かかる場合があります）")

    # 5. 生成実行（ファイルリストとプロンプトを渡す）
    # requestのリストにプロンプト(文字列)とファイルオブジェクトを混ぜて渡せます
    request_content = [prompt] + files_to_analyze
    
    try:
        response = model.generate_content(
            request_content,
            generation_config=genai.types.GenerationConfig(
                temperature=0.7,
                max_output_tokens=8192 # 長文出力用
            )
        )
        
        # 6. 結果の保存
        output_filename = "引き継ぎ資料ドラフト.txt"
        with open(output_filename, "w", encoding="utf-8") as f:
            f.write(response.text)
            
        print(f"\n完了しました！ '{output_filename}' を確認してください。")
        print(f"消費トークン数（目安）: {response.usage_metadata.prompt_token_count}")

    except Exception as e:
        print(f"生成中にエラーが発生しました: {e}")

if __name__ == "__main__":
    create_handover_material()