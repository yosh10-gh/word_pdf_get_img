import os
import csv
import pandas as pd
import zipfile
import shutil
from PIL import Image
import tempfile
from datetime import datetime

def detect_encoding(file_path):
    """
    ファイルの文字エンコーディングを自動判定する
    
    Args:
        file_path (str): CSVファイルのパス
        
    Returns:
        str: 検出されたエンコーディング
    """
    encodings = ['cp932', 'shift_jis', 'utf-8', 'utf-8-sig', 'euc-jp', 'iso-2022-jp']
    
    for encoding in encodings:
        try:
            with open(file_path, 'r', encoding=encoding) as f:
                f.read()
            print(f"  → エンコーディング検出: {encoding}")
            return encoding
        except UnicodeDecodeError:
            continue
    
    print("  → 警告: エンコーディングを自動検出できませんでした。utf-8を使用します。")
    return 'utf-8'

def read_replacement_csv(csv_path):
    """
    差し替え指示CSVファイルを読み込む
    
    Args:
        csv_path (str): CSVファイルのパス
        
    Returns:
        list: 差し替え指示のリスト
    """
    if not os.path.exists(csv_path):
        print(f"エラー: CSVファイルが見つかりません: {csv_path}")
        return []
    
    print(f"CSVファイルを読み込み中: {csv_path}")
    
    try:
        # エンコーディングを自動検出
        encoding = detect_encoding(csv_path)
        
        # CSVファイルを読み込み
        df = pd.read_csv(csv_path, encoding=encoding)
        
        replacements = []
        for index, row in df.iterrows():
            file_path = str(row.iloc[0])  # 最初の列がファイルパス
            
            # 残りの列から画像名と差し替えパスのペアを抽出
            replacement_pairs = []
            for i in range(1, len(row), 2):  # 1列目から2列ずつ
                if i + 1 < len(row) and pd.notna(row.iloc[i]) and pd.notna(row.iloc[i + 1]):
                    image_name = str(row.iloc[i]).strip()
                    replacement_path = str(row.iloc[i + 1]).strip()
                    if image_name and replacement_path:
                        replacement_pairs.append((image_name, replacement_path))
            
            if replacement_pairs:
                replacements.append({
                    'file_path': file_path,
                    'replacements': replacement_pairs
                })
        
        print(f"  → {len(replacements)} ファイルの差し替え指示を読み込みました")
        return replacements
        
    except Exception as e:
        print(f"CSVファイル読み込みエラー: {str(e)}")
        return []

def find_extracted_structure(file_path, extracted_base_dir="extracted_structures"):
    """
    ファイルパスに対応する展開済み構造ディレクトリを見つける
    
    Args:
        file_path (str): 元のWordファイルパス
        extracted_base_dir (str): 展開済み構造のベースディレクトリ
        
    Returns:
        str: 展開済み構造ディレクトリのパス（見つからない場合はNone）
    """
    # ファイルパスを安全な文字に変換（file_extractor.pyと同じロジック）
    relative_path = os.path.relpath(file_path, ".")
    safe_path = relative_path.replace("\\", "_").replace("/", "_").replace(":", "_")
    safe_name = safe_path.replace(".docx", "")
    
    extracted_dir = os.path.join(extracted_base_dir, safe_name)
    
    if os.path.exists(extracted_dir):
        print(f"  → 展開済み構造を発見: {extracted_dir}")
        return extracted_dir
    else:
        print(f"  → 警告: 展開済み構造が見つかりません: {extracted_dir}")
        return None

def replace_image_in_structure(extracted_dir, image_name, replacement_path):
    """
    展開済み構造内の画像ファイルを差し替える
    
    Args:
        extracted_dir (str): 展開済み構造ディレクトリ
        image_name (str): 差し替え対象の画像ファイル名
        replacement_path (str): 新しい画像ファイルのパス
        
    Returns:
        bool: 成功した場合True
    """
    try:
        # mediaディレクトリのパス
        media_dir = os.path.join(extracted_dir, "word", "media")
        target_image_path = os.path.join(media_dir, image_name)
        
        if not os.path.exists(target_image_path):
            print(f"    → エラー: 対象画像が見つかりません: {image_name}")
            return False
        
        if not os.path.exists(replacement_path):
            print(f"    → エラー: 差し替え画像が見つかりません: {replacement_path}")
            return False
        
        # 差し替え画像の形式を確認・変換
        print(f"    → 画像を差し替え中: {image_name} ← {replacement_path}")
        
        # 新しい画像を読み込み
        with Image.open(replacement_path) as img:
            # RGBモードに変換（必要に応じて）
            if img.mode in ('RGBA', 'LA', 'P'):
                img = img.convert('RGB')
            
            # 元の画像の拡張子を取得
            _, original_ext = os.path.splitext(image_name)
            
            # 適切な形式で保存
            if original_ext.lower() in ['.jpg', '.jpeg']:
                img.save(target_image_path, 'JPEG', quality=95, optimize=True)
            elif original_ext.lower() == '.png':
                img.save(target_image_path, 'PNG', optimize=True)
            else:
                # デフォルトはJPEG
                img.save(target_image_path, 'JPEG', quality=95, optimize=True)
        
        print(f"    → 差し替え完了: {image_name}")
        return True
        
    except Exception as e:
        print(f"    → 画像差し替えエラー: {str(e)}")
        return False

def rebuild_docx_from_structure(extracted_dir, output_path):
    """
    展開済み構造からWordファイルを再構築する
    
    Args:
        extracted_dir (str): 展開済み構造ディレクトリ
        output_path (str): 出力Wordファイルパス
        
    Returns:
        bool: 成功した場合True
    """
    try:
        print(f"    → Wordファイルを再構築中: {output_path}")
        
        # 出力ディレクトリを作成
        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        
        # zipファイルとして圧縮
        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            # extracted_dir内のすべてのファイルを追加
            for root, dirs, files in os.walk(extracted_dir):
                for file in files:
                    file_path = os.path.join(root, file)
                    # zip内のパスを計算（extracted_dirを除く）
                    arc_path = os.path.relpath(file_path, extracted_dir)
                    # Windowsのパス区切りをスラッシュに変換
                    arc_path = arc_path.replace('\\', '/')
                    zipf.write(file_path, arc_path)
        
        print(f"    → 再構築完了: {output_path}")
        return True
        
    except Exception as e:
        print(f"    → 再構築エラー: {str(e)}")
        return False

def process_image_replacements(csv_path, output_dir="replace_data/replace_result"):
    """
    画像差し替え処理のメイン関数
    
    Args:
        csv_path (str): 差し替え指示CSVファイルのパス
        output_dir (str): 出力ディレクトリ
    """
    print("=" * 60)
    print("extracted_structures活用型 画像差し替え処理開始")
    print("=" * 60)
    
    # 差し替え指示を読み込み
    replacements = read_replacement_csv(csv_path)
    if not replacements:
        print("差し替え指示がありません。処理を終了します。")
        return
    
    # 出力ディレクトリを作成（タイムスタンプなし、直接指定されたディレクトリに出力）
    os.makedirs(output_dir, exist_ok=True)
    
    success_count = 0
    total_count = len(replacements)
    
    for replacement in replacements:
        file_path = replacement['file_path']
        replacement_pairs = replacement['replacements']
        
        print(f"\n処理中: {file_path}")
        print(f"  差し替え対象: {len(replacement_pairs)} 個の画像")
        
        # 展開済み構造を見つける
        extracted_dir = find_extracted_structure(file_path)
        if not extracted_dir:
            print(f"  → スキップ: 展開済み構造が見つかりません")
            continue
        
        # 展開済み構造をコピー（元を保護）
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        work_dir = os.path.join(tempfile.gettempdir(), f"docx_work_{timestamp}_{success_count}")
        if os.path.exists(work_dir):
            shutil.rmtree(work_dir)
        shutil.copytree(extracted_dir, work_dir)
        print(f"  → 作業用ディレクトリにコピー: {work_dir}")
        
        # 各画像を差し替え
        replacement_success = True
        for image_name, replacement_path in replacement_pairs:
            print(f"  差し替え: {image_name} ← {replacement_path}")
            if not replace_image_in_structure(work_dir, image_name, replacement_path):
                replacement_success = False
        
        # Wordファイルを再構築
        if replacement_success:
            output_filename = os.path.basename(file_path)
            output_path = os.path.join(output_dir, output_filename)
            
            if rebuild_docx_from_structure(work_dir, output_path):
                print(f"  → 成功: {output_path}")
                success_count += 1
            else:
                print(f"  → 失敗: ファイル再構築エラー")
        else:
            print(f"  → 失敗: 画像差し替えエラー")
        
        # 作業用ディレクトリを削除
        try:
            shutil.rmtree(work_dir)
        except:
            pass
    
    print("\n" + "=" * 60)
    print(f"画像差し替え処理完了: {success_count}/{total_count} ファイル成功")
    print(f"出力ディレクトリ: {output_dir}")
    print("=" * 60)

def main():
    """
    メイン実行関数
    """
    # 指定された依頼書を使用
    csv_path = "replace_data/画像差し替え依頼.csv"
    
    # 依頼書の存在確認
    if not os.path.exists(csv_path):
        print(f"エラー: 依頼書が見つかりません: {csv_path}")
        exit(1)
    
    # extracted_structuresディレクトリの存在確認
    if not os.path.exists("extracted_structures"):
        print("エラー: extracted_structuresディレクトリが見つかりません。")
        print("最初にfile_extractor.pyを実行してWord文書を展開してください。")
        exit(1)
    
    print(f"依頼書: {csv_path}")
    print(f"展開済み構造: extracted_structures")
    print(f"出力先: replace_data/replace_result")
    print("")
    
    # 画像差し替え処理を実行
    process_image_replacements(csv_path, "replace_data/replace_result")

if __name__ == "__main__":
    main() 