import os
import glob
import zipfile
import shutil
from datetime import datetime

def extract_docx_structure(docx_file_path, output_dir):
    """
    .docxファイルの構成要素を指定ディレクトリに展開する
    
    Args:
        docx_file_path (str): .docxファイルのパス
        output_dir (str): 出力ディレクトリのパス
        
    Returns:
        tuple: (成功フラグ, 展開ファイル数, 総サイズ)
    """
    try:
        if not os.path.exists(docx_file_path):
            return False, 0, 0
        
        # 出力ディレクトリを作成（既存の場合は削除して再作成）
        if os.path.exists(output_dir):
            shutil.rmtree(output_dir)
        
        os.makedirs(output_dir, exist_ok=True)
        
        extracted_files = 0
        total_extracted_size = 0
        
        with zipfile.ZipFile(docx_file_path, 'r') as zip_file:
            file_list = zip_file.filelist
            
            for file_info in file_list:
                if file_info.is_dir():
                    # ディレクトリの場合は作成のみ
                    dir_path = os.path.join(output_dir, file_info.filename)
                    os.makedirs(dir_path, exist_ok=True)
                    continue
                
                # ファイルを展開
                file_data = zip_file.read(file_info.filename)
                output_file_path = os.path.join(output_dir, file_info.filename)
                
                # 必要に応じてディレクトリを作成
                os.makedirs(os.path.dirname(output_file_path), exist_ok=True)
                
                # ファイルを書き込み
                with open(output_file_path, 'wb') as output_file:
                    output_file.write(file_data)
                
                extracted_files += 1
                total_extracted_size += file_info.file_size
        
        return True, extracted_files, total_extracted_size
        
    except (zipfile.BadZipFile, PermissionError, OSError) as e:
        print(f"    → エラー: {str(e)}")
        return False, 0, 0
    except Exception as e:
        print(f"    → 予期しないエラー: {str(e)}")
        return False, 0, 0

def extract_pdf_structure(pdf_file_path, output_dir):
    """
    .pdfファイルを指定ディレクトリにコピー（PDFは単一ファイルのため）
    
    Args:
        pdf_file_path (str): .pdfファイルのパス
        output_dir (str): 出力ディレクトリのパス
        
    Returns:
        tuple: (成功フラグ, ファイル数, ファイルサイズ)
    """
    try:
        if not os.path.exists(pdf_file_path):
            return False, 0, 0
        
        # 出力ディレクトリを作成
        os.makedirs(output_dir, exist_ok=True)
        
        # PDFファイルをコピー
        filename = os.path.basename(pdf_file_path)
        output_path = os.path.join(output_dir, filename)
        shutil.copy2(pdf_file_path, output_path)
        
        file_size = os.path.getsize(pdf_file_path)
        
        # PDF構造情報をテキストファイルで出力
        info_file = os.path.join(output_dir, "pdf_info.txt")
        with open(info_file, 'w', encoding='utf-8') as f:
            f.write(f"PDFファイル情報\n")
            f.write(f"=" * 50 + "\n")
            f.write(f"ファイル名: {filename}\n")
            f.write(f"ファイルサイズ: {file_size:,} bytes ({file_size / 1024 / 1024:.2f} MB)\n")
            f.write(f"更新日時: {datetime.fromtimestamp(os.path.getmtime(pdf_file_path))}\n")
            f.write(f"\n注意: PDFファイルは単一構造のため、元ファイルをコピーしました。\n")
            f.write(f"詳細な構造解析にはPyMuPDF等の専用ライブラリが必要です。\n")
        
        return True, 2, file_size  # PDFファイル + 情報ファイル
        
    except Exception as e:
        print(f"    → PDFコピーエラー: {str(e)}")
        return False, 0, 0

def get_safe_dirname(file_path):
    """
    ファイルパスから安全なディレクトリ名を生成する
    
    Args:
        file_path (str): ファイルパス
        
    Returns:
        str: 安全なディレクトリ名
    """
    # ファイル名（拡張子なし）を取得
    basename = os.path.splitext(os.path.basename(file_path))[0]
    
    # 安全でない文字を置換
    safe_name = basename.replace('\\', '_').replace('/', '_').replace(':', '_')
    safe_name = safe_name.replace('*', '_').replace('?', '_').replace('"', '_')
    safe_name = safe_name.replace('<', '_').replace('>', '_').replace('|', '_')
    
    return safe_name

def count_media_files(structure_dir):
    """
    展開された構造内の画像ファイル数をカウントする
    
    Args:
        structure_dir (str): 構造ディレクトリのパス
        
    Returns:
        int: 画像ファイル数
    """
    media_dir = os.path.join(structure_dir, "word", "media")
    if not os.path.exists(media_dir):
        return 0
    
    media_files = [f for f in os.listdir(media_dir) if os.path.isfile(os.path.join(media_dir, f))]
    return len(media_files)

def batch_extract_all_files(search_directory="test_directry", output_base_dir="extracted_structures"):
    """
    指定ディレクトリ内の全ての.docxと.pdfファイルを分解・展開する
    
    Args:
        search_directory (str): 検索対象のディレクトリ
        output_base_dir (str): 出力ベースディレクトリ
    """
    print("=" * 80)
    print("一括ファイル構造展開処理開始")
    print("=" * 80)
    
    if not os.path.exists(search_directory):
        print(f"検索ディレクトリが見つかりません: {search_directory}")
        return
    
    # .docx と .pdf ファイルを検索
    docx_pattern = os.path.join(search_directory, "**", "*.docx")
    pdf_pattern = os.path.join(search_directory, "**", "*.pdf")
    
    docx_files = glob.glob(docx_pattern, recursive=True)
    pdf_files = glob.glob(pdf_pattern, recursive=True)
    
    all_files = [(f, 'docx') for f in docx_files] + [(f, 'pdf') for f in pdf_files]
    
    if not all_files:
        print("対象ファイルが見つかりませんでした。")
        return
    
    print(f"検索対象ディレクトリ: {os.path.abspath(search_directory)}")
    print(f"見つかったファイル数: {len(all_files)} (DOCX: {len(docx_files)}, PDF: {len(pdf_files)})")
    print(f"出力ベースディレクトリ: {os.path.abspath(output_base_dir)}")
    print()
    
    # 出力ベースディレクトリを作成
    os.makedirs(output_base_dir, exist_ok=True)
    
    success_count = 0
    fail_count = 0
    total_extracted_files = 0
    total_extracted_size = 0
    total_images = 0
    
    print("処理中...")
    print(f"{'No':<4} {'ファイル名':<40} {'種別':<6} {'結果':<8} {'ファイル数':<8} {'サイズ':<12} {'画像数'}")
    print("-" * 95)
    
    for idx, (file_path, file_type) in enumerate(all_files, 1):
        # 相対パスを取得
        rel_path = os.path.relpath(file_path, search_directory)
        filename = os.path.basename(file_path)
        
        # 安全な出力ディレクトリ名を生成
        safe_dirname = get_safe_dirname(file_path)
        output_dir = os.path.join(output_base_dir, safe_dirname)
        
        try:
            if file_type == 'docx':
                success, file_count, file_size = extract_docx_structure(file_path, output_dir)
                image_count = count_media_files(output_dir) if success else 0
            else:  # pdf
                success, file_count, file_size = extract_pdf_structure(file_path, output_dir)
                image_count = 0  # PDFの画像カウントは現在未対応
            
            if success:
                success_count += 1
                total_extracted_files += file_count
                total_extracted_size += file_size
                total_images += image_count
                result = "成功"
            else:
                fail_count += 1
                result = "失敗"
                file_count = 0
                file_size = 0
                image_count = 0
            
            # サイズの表示形式
            size_str = f"{file_size / 1024 / 1024:.1f}MB" if file_size > 0 else "0MB"
            
            print(f"{idx:<4} {filename:<40} {file_type.upper():<6} {result:<8} {file_count:<8} {size_str:<12} {image_count}")
            
        except Exception as e:
            fail_count += 1
            print(f"{idx:<4} {filename:<40} {file_type.upper():<6} {'エラー':<8} {'0':<8} {'0MB':<12} 0")
            print(f"    → エラー詳細: {str(e)}")
    
    print("-" * 95)
    print()
    print("=" * 80)
    print("一括展開処理完了")
    print("=" * 80)
    print(f"処理済みファイル数: {len(all_files)}")
    print(f"成功: {success_count} ファイル")
    print(f"失敗: {fail_count} ファイル")
    print(f"総展開ファイル数: {total_extracted_files:,}")
    print(f"総展開サイズ: {total_extracted_size:,} bytes ({total_extracted_size / 1024 / 1024:.2f} MB)")
    print(f"総画像ファイル数: {total_images}")
    print(f"出力ディレクトリ: {os.path.abspath(output_base_dir)}")
    print("=" * 80)
    
    # 出力された構造を確認
    if success_count > 0:
        print()
        print("展開されたディレクトリ一覧:")
        print("-" * 50)
        try:
            for item in sorted(os.listdir(output_base_dir)):
                item_path = os.path.join(output_base_dir, item)
                if os.path.isdir(item_path):
                    # ディレクトリ内のファイル数を数える
                    file_count = sum(len(files) for _, _, files in os.walk(item_path))
                    dir_size = sum(os.path.getsize(os.path.join(dirpath, filename))
                                 for dirpath, _, filenames in os.walk(item_path)
                                 for filename in filenames)
                    print(f"  {item:<30} ({file_count} ファイル, {dir_size / 1024 / 1024:.1f} MB)")
        except Exception:
            print("  ディレクトリ一覧の取得に失敗しました")

def main():
    # test_directryの全ファイルを分解
    search_dir = "test_directry"
    output_dir = "extracted_structures"
    
    batch_extract_all_files(search_dir, output_dir)

if __name__ == "__main__":
    main() 