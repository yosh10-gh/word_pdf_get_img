import os
import glob
import pandas as pd
import zipfile
from PIL import Image
import io
import tempfile
from openpyxl import Workbook
from openpyxl.drawing.image import Image as OpenpyxlImage
from openpyxl.utils.dataframe import dataframe_to_rows

def extract_docx_pdf_files(directory_path):
    """
    指定されたディレクトリを再帰的にクロールして、
    .docx と .pdf ファイルのパスを抽出する
    
    Args:
        directory_path (str): 検索対象のディレクトリパス
        
    Returns:
        list: 見つかった .docx と .pdf ファイルのパスのリスト
    """
    found_files = []
    
    try:
        # ディレクトリの存在確認
        if not os.path.exists(directory_path):
            print(f"警告: ディレクトリが存在しません: {directory_path}")
            return found_files
        
        # .docx ファイルを検索
        docx_pattern = os.path.join(directory_path, "**", "*.docx")
        docx_files = glob.glob(docx_pattern, recursive=True)
        # 有効なファイルのみを追加
        for file_path in docx_files:
            if os.path.isfile(file_path) and os.access(file_path, os.R_OK):
                found_files.append(file_path)
        
        # .pdf ファイルを検索
        pdf_pattern = os.path.join(directory_path, "**", "*.pdf")
        pdf_files = glob.glob(pdf_pattern, recursive=True)
        # 有効なファイルのみを追加
        for file_path in pdf_files:
            if os.path.isfile(file_path) and os.access(file_path, os.R_OK):
                found_files.append(file_path)
    
    except Exception as e:
        print(f"ファイル検索エラー: {str(e)}")
    
    return found_files

def has_images_in_docx(file_path):
    """
    .docxファイルに画像が含まれているかチェックする
    
    Args:
        file_path (str): .docxファイルのパス
        
    Returns:
        bool: 画像が含まれている場合True
    """
    try:
        # ファイルの存在とアクセス権限をチェック
        if not os.path.exists(file_path) or not os.access(file_path, os.R_OK):
            return False
            
        with zipfile.ZipFile(file_path, 'r') as zip_file:
            # word/media/ フォルダ内のファイルをチェック
            for file_info in zip_file.filelist:
                if file_info.filename.startswith('word/media/'):
                    return True
        return False
    except (zipfile.BadZipFile, PermissionError, OSError) as e:
        print(f"  → docxファイル読み込みエラー: {str(e)}")
        return False
    except Exception:
        return False

def get_media_filenames(file_path):
    """
    .docxファイル内の画像ファイル名のリストを取得する
    
    Args:
        file_path (str): .docxファイルのパス
        
    Returns:
        list: 画像ファイル名のリスト（ソート済み）
    """
    media_filenames = []
    try:
        # ファイルの存在とアクセス権限をチェック
        if not os.path.exists(file_path) or not os.access(file_path, os.R_OK):
            return media_filenames
            
        with zipfile.ZipFile(file_path, 'r') as zip_file:
            # word/media/ フォルダ内のファイルを検索
            for file_info in zip_file.filelist:
                if file_info.filename.startswith('word/media/') and not file_info.is_dir():
                    # ファイル名のみを取得
                    filename = os.path.basename(file_info.filename)
                    media_filenames.append(filename)
        
        # ファイル名順にソート
        media_filenames.sort()
        
    except (zipfile.BadZipFile, PermissionError, OSError) as e:
        print(f"  → docxファイル処理エラー: {str(e)}")
    except Exception:
        pass
    return media_filenames

def extract_images_from_docx(file_path):
    """
    .docxファイルから画像を抽出する
    
    Args:
        file_path (str): .docxファイルのパス
        
    Returns:
        list: 抽出された画像のPILImageオブジェクトのリスト
    """
    images = []
    try:
        # ファイルの存在とアクセス権限をチェック
        if not os.path.exists(file_path) or not os.access(file_path, os.R_OK):
            return images
            
        with zipfile.ZipFile(file_path, 'r') as zip_file:
            # word/media/ フォルダ内のファイルを検索
            for file_info in zip_file.filelist:
                if file_info.filename.startswith('word/media/') and not file_info.is_dir():
                    try:
                        # 画像データを読み込み
                        image_data = zip_file.read(file_info.filename)
                        if len(image_data) == 0:  # 空のファイルをスキップ
                            continue
                            
                        # PILImageオブジェクトに変換
                        image = Image.open(io.BytesIO(image_data))
                        # 画像形式を確認してRGBに変換
                        if image.mode in ('RGBA', 'LA', 'P'):
                            image = image.convert('RGB')
                        images.append(image)
                    except (Image.UnidentifiedImageError, OSError) as e:
                        print(f"    → 画像読み込みエラー ({file_info.filename}): {str(e)}")
                        continue
                    except Exception:
                        continue
    except (zipfile.BadZipFile, PermissionError, OSError) as e:
        print(f"  → docxファイル処理エラー: {str(e)}")
    except Exception:
        pass
    return images

def has_images_in_pdf(file_path):
    """
    .pdfファイルに画像が含まれているかチェックする
    
    Args:
        file_path (str): .pdfファイルのパス
        
    Returns:
        bool: 画像が含まれている場合True
    """
    try:
        # ファイルの存在とアクセス権限をチェック
        if not os.path.exists(file_path) or not os.access(file_path, os.R_OK):
            return False
            
        import PyPDF2
        with open(file_path, 'rb') as pdf_file:
            pdf_reader = PyPDF2.PdfReader(pdf_file)
            
            # 暗号化されたPDFの場合
            if pdf_reader.is_encrypted:
                try:
                    pdf_reader.decrypt('')  # 空のパスワードで試行
                except:
                    print(f"  → 暗号化されたPDFです: {file_path}")
                    return False
            
            for page_num, page in enumerate(pdf_reader.pages):
                try:
                    if '/Resources' in page and '/XObject' in page['/Resources']:
                        xObject = page['/Resources']['/XObject'].get_object()
                        for obj in xObject:
                            if xObject[obj]['/Subtype'] == '/Image':
                                return True
                except Exception:
                    continue
        return False
    except (PyPDF2.errors.PdfReadError, PermissionError, OSError) as e:
        print(f"  → PDFファイル読み込みエラー: {str(e)}")
        return False
    except Exception:
        return False

def extract_images_from_pdf(file_path):
    """
    .pdfファイルから画像を抽出する
    
    Args:
        file_path (str): .pdfファイルのパス
        
    Returns:
        list: 抽出された画像のPILImageオブジェクトのリスト
    """
    images = []
    try:
        # ファイルの存在とアクセス権限をチェック
        if not os.path.exists(file_path) or not os.access(file_path, os.R_OK):
            return images
            
        import PyPDF2
        with open(file_path, 'rb') as pdf_file:
            pdf_reader = PyPDF2.PdfReader(pdf_file)
            
            # 暗号化されたPDFの場合
            if pdf_reader.is_encrypted:
                try:
                    pdf_reader.decrypt('')  # 空のパスワードで試行
                except:
                    print(f"  → 暗号化されたPDFをスキップ: {file_path}")
                    return images
            
            for page_num, page in enumerate(pdf_reader.pages):
                try:
                    if '/Resources' in page and '/XObject' in page['/Resources']:
                        xObject = page['/Resources']['/XObject'].get_object()
                        for obj in xObject:
                            if xObject[obj]['/Subtype'] == '/Image':
                                try:
                                    # 画像データを抽出
                                    img_obj = xObject[obj]
                                    if '/Filter' in img_obj:
                                        if img_obj['/Filter'] == '/DCTDecode':
                                            # JPEG画像
                                            img_data = img_obj._data
                                            if len(img_data) > 0:  # 空のデータをスキップ
                                                image = Image.open(io.BytesIO(img_data))
                                                if image.mode in ('RGBA', 'LA', 'P'):
                                                    image = image.convert('RGB')
                                                images.append(image)
                                        elif img_obj['/Filter'] == '/FlateDecode':
                                            # PNG/その他の圧縮画像
                                            try:
                                                width = int(img_obj['/Width'])
                                                height = int(img_obj['/Height'])
                                                img_data = img_obj._data
                                                if len(img_data) > 0 and '/ColorSpace' in img_obj:
                                                    if img_obj['/ColorSpace'] == '/DeviceRGB':
                                                        expected_size = width * height * 3
                                                        if len(img_data) >= expected_size:
                                                            image = Image.frombytes('RGB', (width, height), img_data[:expected_size])
                                                            images.append(image)
                                            except (ValueError, TypeError):
                                                continue
                                except (Image.UnidentifiedImageError, OSError, ValueError) as e:
                                    print(f"    → PDF画像抽出エラー (page {page_num + 1}): {str(e)}")
                                    continue
                                except Exception:
                                    continue
                except Exception:
                    continue
    except (PyPDF2.errors.PdfReadError, PermissionError, OSError) as e:
        print(f"  → PDFファイル処理エラー: {str(e)}")
    except Exception:
        pass
    return images

def filter_files_with_images(file_list):
    """
    ファイルリストから画像が含まれているファイルのみを抽出する
    
    Args:
        file_list (list): ファイルパスのリスト
        
    Returns:
        list: 画像が含まれているファイルのリスト
    """
    files_with_images = []
    
    for file_path in file_list:
        try:
            print(f"チェック中: {file_path}")
            
            if file_path.endswith('.docx'):
                if has_images_in_docx(file_path):
                    files_with_images.append(file_path)
                    print(f"  → 画像あり")
                else:
                    print(f"  → 画像なし")
                    
            elif file_path.endswith('.pdf'):
                if has_images_in_pdf(file_path):
                    files_with_images.append(file_path)
                    print(f"  → 画像あり")
                else:
                    print(f"  → 画像なし")
        except Exception as e:
            print(f"  → ファイルチェックエラー: {str(e)}")
            continue
    
    return files_with_images

def resize_image_to_100px(image):
    """
    画像を100px×100pxにリサイズする（アスペクト比を保持）
    
    Args:
        image: PILImageオブジェクト
        
    Returns:
        PILImageオブジェクト: リサイズされた画像
    """
    try:
        # 画像のサイズをチェック
        if image.size[0] == 0 or image.size[1] == 0:
            raise ValueError("無効な画像サイズ")
        
        # アスペクト比を保持して100px以内にリサイズ
        image.thumbnail((100, 100), Image.Resampling.LANCZOS)
        return image
    except Exception as e:
        print(f"    → 画像リサイズエラー: {str(e)}")
        # フォールバック: 最小サイズの白い画像を作成
        fallback_image = Image.new('RGB', (50, 50), 'white')
        return fallback_image

def save_to_excel_with_images(file_list, output_dir="result", output_filename="検索結果.xlsx"):
    """
    ファイルリストと画像をExcelファイルに保存する（実際の画像ファイル名を使用）
    
    Args:
        file_list (list): ファイルパスのリスト
        output_dir (str): 出力ディレクトリ
        output_filename (str): 出力するExcelファイル名
    """
    if not file_list:
        print("保存するファイルがありません。")
        return
    
    # 出力ディレクトリを作成
    os.makedirs(output_dir, exist_ok=True)
    output_path = os.path.join(output_dir, output_filename)
    
    # 新しいワークブックを作成
    wb = Workbook()
    ws = wb.active
    
    # 全ファイルの画像ファイル名を調査
    all_image_filenames = set()
    file_image_map = {}
    
    print("画像ファイル名を調査中...")
    for file_path in file_list:
        try:
            if file_path.endswith('.docx'):
                media_filenames = get_media_filenames(file_path)
                file_image_map[file_path] = media_filenames
                all_image_filenames.update(media_filenames)
            elif file_path.endswith('.pdf'):
                images = extract_images_from_pdf(file_path)
                pdf_filenames = [f"pdf_image{i+1}" for i in range(len(images))]
                file_image_map[file_path] = pdf_filenames
                all_image_filenames.update(pdf_filenames)
        except Exception:
            file_image_map[file_path] = []
    
    # 画像ファイル名をソート
    sorted_image_filenames = sorted(all_image_filenames)
    max_images = len(sorted_image_filenames)
    
    print(f"ユニークな画像ファイル名数: {max_images}")
    
    # ヘッダーを動的に設定
    ws['A1'] = 'ファイルパス'
    for i, filename in enumerate(sorted_image_filenames):
        col_letter = chr(ord('B') + i) if i < 25 else f"A{chr(ord('A') + i - 25)}"
        ws[f'{col_letter}1'] = filename
    
    # 行の高さを設定（100px用）
    for row in range(2, len(file_list) + 2):
        ws.row_dimensions[row].height = 75  # Excelのポイント単位
    
    # 列の幅を設定
    ws.column_dimensions['A'].width = 50
    for i in range(max_images):
        col_letter = chr(ord('B') + i) if i < 25 else f"A{chr(ord('A') + i - 25)}"
        ws.column_dimensions[col_letter].width = 15
    
    temp_files = []  # 一時ファイルのリストを保持
    
    try:
        for idx, file_path in enumerate(file_list, start=2):
            try:
                print(f"画像抽出中: {file_path}")
                
                # A列にファイルパスを設定
                ws[f'A{idx}'] = file_path
                
                # 画像を抽出
                images = []
                current_filenames = file_image_map.get(file_path, [])
                
                if file_path.endswith('.docx'):
                    images = extract_images_from_docx(file_path)
                elif file_path.endswith('.pdf'):
                    images = extract_images_from_pdf(file_path)
                
                # 実際のファイル名と画像を対応付けて配置
                for img_idx, (filename, image) in enumerate(zip(current_filenames, images)):
                    try:
                        # 該当する列名を計算
                        if filename in sorted_image_filenames:
                            col_idx = sorted_image_filenames.index(filename)
                            col_letter = chr(ord('B') + col_idx) if col_idx < 25 else f"A{chr(ord('A') + col_idx - 25)}"
                            
                            # 画像を100px×100pxにリサイズ
                            resized_image = resize_image_to_100px(image.copy())
                            
                            # 一時ファイルを作成（手動で削除）
                            temp_fd, temp_path = tempfile.mkstemp(suffix='.png')
                            temp_files.append(temp_path)  # 削除用にリストに追加
                            
                            try:
                                # ファイルディスクリプタを閉じる
                                os.close(temp_fd)
                                
                                # 画像を保存
                                resized_image.save(temp_path, 'PNG', optimize=True)
                                
                                # ファイルサイズをチェック（異常に大きい場合はスキップ）
                                if os.path.getsize(temp_path) > 10 * 1024 * 1024:  # 10MB制限
                                    print(f"    → {filename}: ファイルサイズが大きすぎます")
                                    continue
                                
                                # Excelに画像を挿入
                                img = OpenpyxlImage(temp_path)
                                img.width = 100
                                img.height = 100
                                
                                cell_location = f'{col_letter}{idx}'
                                ws.add_image(img, cell_location)
                                
                                print(f"  → {filename}: 配置完了")
                            except (OSError, PermissionError) as e:
                                print(f"  → {filename}: ファイル操作エラー ({str(e)})")
                            except Exception as e:
                                print(f"  → {filename}: エラー ({str(e)})")
                    except Exception as e:
                        print(f"  → {filename}: 画像処理エラー ({str(e)})")
                
                print(f"  → 合計 {len(images)} 個の画像を処理")
            except Exception as e:
                print(f"  → ファイル処理エラー: {str(e)}")
                continue
        
        # Excelファイルを保存
        try:
            wb.save(output_path)
            print(f"Excelファイルを保存しました: {output_path}")
        except (PermissionError, OSError) as e:
            print(f"Excelファイル保存エラー: {str(e)}")
            # 代替ファイル名で保存を試行
            import time
            alt_filename = f"検索結果_{int(time.time())}.xlsx"
            alt_path = os.path.join(output_dir, alt_filename)
            wb.save(alt_path)
            print(f"代替ファイル名で保存しました: {alt_path}")
    
    finally:
        # 一時ファイルをクリーンアップ
        for temp_path in temp_files:
            try:
                if os.path.exists(temp_path):
                    os.unlink(temp_path)
            except Exception:
                pass  # 削除できなくても続行

def save_to_csv_with_image_info(file_list, output_dir="result", output_filename="検索結果.csv"):
    """
    ファイルリストと画像情報をCSVファイルに保存する（実際のファイル名を使用）
    
    Args:
        file_list (list): ファイルパスのリスト
        output_dir (str): 出力ディレクトリ
        output_filename (str): 出力するCSVファイル名
    """
    if not file_list:
        print("保存するファイルがありません。")
        return
    
    # 出力ディレクトリを作成
    os.makedirs(output_dir, exist_ok=True)
    output_path = os.path.join(output_dir, output_filename)
    
    # CSVデータを準備
    csv_data = []
    
    print("CSV用データを作成中...")
    for file_path in file_list:
        try:
            print(f"画像情報抽出中: {file_path}")
            
            row = {'ファイルパス': file_path}
            
            if file_path.endswith('.docx'):
                media_filenames = get_media_filenames(file_path)
                # 最初の3つの画像ファイル名を記録（拡張可能）
                for i, filename in enumerate(media_filenames[:6]):  # 最大6個まで
                    row[f'画像{i+1}_ファイル名'] = filename
                    
            elif file_path.endswith('.pdf'):
                images = extract_images_from_pdf(file_path)
                # PDFの場合は便宜的な名前を使用
                for i in range(min(len(images), 6)):  # 最大6個まで
                    row[f'画像{i+1}_ファイル名'] = f"pdf_image{i+1}"
            
            csv_data.append(row)
            
        except Exception as e:
            print(f"  → ファイル処理エラー: {str(e)}")
            # エラーの場合もファイルパスは記録
            csv_data.append({'ファイルパス': file_path})
            continue
    
    # DataFrameを作成してCSV保存
    try:
        df = pd.DataFrame(csv_data)
        df.to_csv(output_path, index=False, encoding='utf-8-sig')
        print(f"CSVファイルを保存しました: {output_path}")
    except Exception as e:
        print(f"CSVファイル保存エラー: {str(e)}")

def main():
    # 検索対象のディレクトリを指定（現在のディレクトリから開始）
    search_directory = "."
    
    print(f"ディレクトリをクロール中: {os.path.abspath(search_directory)}")
    print("-" * 50)
    
    # ファイルを抽出
    files = extract_docx_pdf_files(search_directory)
    
    # 結果を表示
    if files:
        print(f"見つかったファイル数: {len(files)}")
        
        # 画像が含まれているファイルのみを抽出
        print("\n" + "-" * 50)
        print("画像が含まれているファイルをチェック中...")
        print("-" * 50)
        
        files_with_images = filter_files_with_images(files)
        
        print("\n" + "-" * 50)
        print(f"画像が含まれているファイル数: {len(files_with_images)}")
        print("\n画像が含まれているファイル:")
        for i, file_path in enumerate(files_with_images, 1):
            print(f"{i:3d}. {file_path}")
        
        # Excelファイルに保存（実際の画像ファイル名でヘッダー作成）
        if files_with_images:
            print("\n" + "-" * 50)
            print("実際の画像ファイル名を使用したExcelファイルを作成中...")
            print("-" * 50)
            save_to_excel_with_images(files_with_images, "result", "検索結果.xlsx")
            
            print("\n" + "-" * 50)
            print("実際の画像ファイル名を使用したCSVファイルを作成中...")
            print("-" * 50)
            save_to_csv_with_image_info(files_with_images, "result", "検索結果.csv")
        else:
            print("\n画像が含まれているファイルがありませんでした。")
    else:
        print("該当するファイルが見つかりませんでした。")
    
    print("-" * 50)
    print("抽出完了")

if __name__ == "__main__":
    main() 