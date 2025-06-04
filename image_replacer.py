import os
import pandas as pd
import zipfile
import shutil
from PIL import Image
import io
import tempfile
from datetime import datetime
import re

def load_replacement_orders(csv_file_path):
    """
    画像差し替え指示CSVファイルを読み込む（新形式対応）
    
    形式: ファイルパス,画像名1,差し替えパス1,画像名2,差し替えパス2,...
    
    Args:
        csv_file_path (str): CSVファイルのパス
        
    Returns:
        list: 差し替え指示の辞書のリスト
    """
    replacement_orders = []
    
    try:
        # 一般的なエンコーディングで順次試行
        encodings_to_try = ['cp932', 'shift-jis', 'utf-8', 'utf-8-sig', 'latin1']
        
        df = None
        for enc in encodings_to_try:
            try:
                print(f"エンコーディング '{enc}' で読み込み試行中...")
                df = pd.read_csv(csv_file_path, encoding=enc, header=None)  # ヘッダーなしで読み込み
                print(f"エンコーディング '{enc}' で読み込み成功")
                break
            except Exception as e:
                print(f"エンコーディング '{enc}' で読み込み失敗: {str(e)}")
                continue
        
        if df is None:
            print("CSVファイルを読み込めませんでした")
            return replacement_orders
        
        # 1行目がヘッダーの場合はスキップ
        start_row = 1 if df.iloc[0, 0] == 'ファイルパス' else 0
        
        for index in range(start_row, len(df)):
            row = df.iloc[index]
            file_path = row.iloc[0]  # ファイルパス
            
            # 空の行をスキップ
            if pd.isna(file_path) or file_path == '':
                continue
            
            order = {
                'file_path': file_path,
                'replacements': []
            }
            
            # 1列目以降のペア（画像名、差し替えパス）を処理
            for col_idx in range(1, len(row), 2):
                if col_idx + 1 < len(row):
                    target_image = row.iloc[col_idx]
                    replacement_path = row.iloc[col_idx + 1]
                    
                    # 両方の値が存在し、空でない場合に追加
                    if (not pd.isna(target_image) and not pd.isna(replacement_path) and 
                        target_image != '' and replacement_path != ''):
                        order['replacements'].append({
                            'target': target_image,
                            'replacement_path': replacement_path
                        })
            
            # 差し替え指示がある場合のみ追加
            if order['replacements']:
                replacement_orders.append(order)
                print(f"読み込み: {file_path} ({len(order['replacements'])}個の差し替え指示)")
    
    except Exception as e:
        print(f"CSVファイル読み込みエラー: {str(e)}")
    
    return replacement_orders

def get_image_index(target_image):
    """
    image1, image2, image3, ...の文字列を配列インデックス（0,1,2,...）に変換
    
    Args:
        target_image (str): "image1", "image2", "image3", ...のいずれか
        
    Returns:
        int: 配列インデックス（0,1,2,...）または-1（無効な場合）
    """
    try:
        # 正規表現でimage後の数字を抽出
        match = re.match(r'image(\d+)', target_image.lower())
        if match:
            image_num = int(match.group(1))
            return image_num - 1  # 1-based index を 0-based index に変換
        return -1
    except Exception:
        return -1

def prepare_replacement_image(image_path):
    """
    差し替え用画像を準備する（適切な形式に変換）
    
    Args:
        image_path (str): 差し替え用画像のパス
        
    Returns:
        bytes: 変換された画像データ
    """
    try:
        if not os.path.exists(image_path):
            print(f"  → 差し替え画像が見つかりません: {image_path}")
            return None
        
        # 画像を読み込み
        with Image.open(image_path) as img:
            # RGBに変換
            if img.mode in ('RGBA', 'LA', 'P'):
                img = img.convert('RGB')
            
            # JPEGバイトデータに変換
            output = io.BytesIO()
            img.save(output, format='JPEG', quality=95)
            return output.getvalue()
    
    except Exception as e:
        print(f"  → 差し替え画像の準備エラー: {str(e)}")
        return None

def replace_images_in_docx(file_path, replacements, output_dir):
    """
    .docxファイル内の画像を差し替える（全ての画像に対応）
    
    Args:
        file_path (str): 元の.docxファイルパス
        replacements (list): 差し替え指示のリスト
        output_dir (str): 出力ディレクトリ
        
    Returns:
        str: 出力されたファイルのパス（成功時）、None（失敗時）
    """
    try:
        print(f"  → Wordファイルの画像差し替え開始: {file_path}")
        
        # ファイル名のみを取得（ディレクトリ構造は無視）
        file_name = os.path.basename(file_path)
        
        # 出力ファイルパスを準備（直接出力ディレクトリに配置）
        output_path = os.path.join(output_dir, file_name)
        
        # 出力ディレクトリを作成
        os.makedirs(output_dir, exist_ok=True)
        
        # 一時ファイルを作成して新しいzipファイルを構築
        temp_fd, temp_path = tempfile.mkstemp(suffix='.docx')
        os.close(temp_fd)
        
        try:
            with zipfile.ZipFile(file_path, 'r') as original_zip:
                # word/media/内の画像ファイルをリストアップ
                media_files = []
                for file_info in original_zip.filelist:
                    if file_info.filename.startswith('word/media/') and not file_info.is_dir():
                        media_files.append(file_info.filename)
                
                media_files.sort()  # ファイル名順にソート
                print(f"    → 検出された画像ファイル数: {len(media_files)}")
                
                # 差し替え対象の画像データを準備
                replacement_data = {}
                for replacement in replacements:
                    target_image = replacement['target']
                    replacement_path = replacement['replacement_path']
                    
                    # 画像ファイル名で直接検索
                    image_index = get_image_index_by_filename(media_files, target_image)
                    if image_index == -1:
                        print(f"    → 対象画像が見つかりません: {target_image}")
                        continue
                    
                    # 差し替え画像を準備
                    new_image_data = prepare_replacement_image(replacement_path)
                    if new_image_data is None:
                        continue
                    
                    # 対象画像のファイル名
                    target_filename = media_files[image_index]
                    replacement_data[target_filename] = {
                        'data': new_image_data,
                        'target': target_image,
                        'source': os.path.basename(replacement_path)
                    }
                
                # 新しいzipファイルを作成
                with zipfile.ZipFile(temp_path, 'w', zipfile.ZIP_DEFLATED) as new_zip:
                    for item in original_zip.filelist:
                        if item.filename in replacement_data:
                            # 差し替え画像を書き込み
                            new_zip.writestr(item.filename, replacement_data[item.filename]['data'])
                            print(f"    → {replacement_data[item.filename]['target']} を差し替えました: {replacement_data[item.filename]['source']}")
                        else:
                            # 既存のファイルをコピー
                            data = original_zip.read(item.filename)
                            new_zip.writestr(item.filename, data)
            
            # 一時ファイルを最終的な出力先に移動
            shutil.move(temp_path, output_path)
            
            print(f"  → 差し替え完了: {output_path}")
            return output_path
        
        except Exception as e:
            # エラー時は一時ファイルを削除
            if os.path.exists(temp_path):
                os.unlink(temp_path)
            raise e
    
    except Exception as e:
        print(f"  → Wordファイル差し替えエラー: {str(e)}")
        return None

def replace_images_in_pdf(file_path, replacements, output_dir):
    """
    .pdfファイル内の画像を差し替える（現在は対応不可）
    
    Args:
        file_path (str): 元の.pdfファイルパス
        replacements (list): 差し替え指示のリスト
        output_dir (str): 出力ディレクトリ
        
    Returns:
        str: None（PDFは現在対応不可）
    """
    print(f"  → PDFファイルの画像差し替えは現在対応していません: {file_path}")
    print(f"    → PyMuPDF等の追加ライブラリが必要です")
    return None

def process_image_replacement(csv_file_path, output_base_dir="replace_data/replace_result"):
    """
    CSVファイルの指示に基づいて画像差し替えを実行する
    
    Args:
        csv_file_path (str): 差し替え指示CSVファイルのパス
        output_base_dir (str): 出力ディレクトリのパス
    """
    print("=" * 60)
    print("画像差し替え処理開始")
    print("=" * 60)
    
    # CSVから差し替え指示を読み込み
    replacement_orders = load_replacement_orders(csv_file_path)
    
    if not replacement_orders:
        print("差し替え指示が見つかりませんでした。")
        return
    
    print(f"差し替え対象ファイル数: {len(replacement_orders)}")
    
    # 出力ディレクトリを設定（タイムスタンプなし）
    output_dir = output_base_dir
    os.makedirs(output_dir, exist_ok=True)
    
    print(f"出力ディレクトリ: {output_dir}")
    print("-" * 60)
    
    success_count = 0
    fail_count = 0
    
    # 各ファイルの差し替え処理
    for order in replacement_orders:
        file_path = order['file_path']
        replacements = order['replacements']
        
        print(f"処理中: {file_path}")
        print(f"  → 差し替え指示数: {len(replacements)}")
        
        # ファイルの存在確認
        if not os.path.exists(file_path):
            print(f"  → ファイルが見つかりません")
            fail_count += 1
            continue
        
        # ファイルタイプに応じて処理
        if file_path.endswith('.docx'):
            result = replace_images_in_docx(file_path, replacements, output_dir)
            if result:
                success_count += 1
            else:
                fail_count += 1
        elif file_path.endswith('.pdf'):
            result = replace_images_in_pdf(file_path, replacements, output_dir)
            if result:
                success_count += 1
            else:
                fail_count += 1
        else:
            print(f"  → 対応していないファイル形式です")
            fail_count += 1
        
        print()
    
    print("=" * 60)
    print(f"画像差し替え処理完了")
    print(f"成功: {success_count} ファイル")
    print(f"失敗: {fail_count} ファイル")
    print(f"出力ディレクトリ: {output_dir}")
    print("=" * 60)

def get_image_index_by_filename(media_files, target_filename):
    """
    画像ファイル名から配列インデックスを取得する
    
    Args:
        media_files (list): メディアファイル名のリスト（ソート済み）
        target_filename (str): 対象の画像ファイル名
        
    Returns:
        int: 配列インデックス（見つからない場合は-1）
    """
    try:
        for i, media_file in enumerate(media_files):
            # ファイル名のみを比較（パスを除去）
            if os.path.basename(media_file) == target_filename:
                return i
        return -1
    except Exception:
        return -1

def main():
    import sys
    
    # コマンドライン引数からCSVファイルパスを取得
    if len(sys.argv) > 1:
        csv_file = sys.argv[1]
    else:
        # デフォルトのファイル名
        csv_file = "order_edit_image.csv"
    
    if not os.path.exists(csv_file):
        print(f"差し替え指示ファイルが見つかりません: {csv_file}")
        return
    
    # 画像差し替え処理を実行
    process_image_replacement(csv_file)

if __name__ == "__main__":
    main() 