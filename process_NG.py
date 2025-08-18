from openpyxl import load_workbook, Workbook
from openpyxl.drawing.image import Image
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
import os
from datetime import datetime
import shutil
import re
import glob
import sys
import subprocess
from PIL import Image as PILImage
import pytesseract
from pytesseract import Output
import traceback

from extract_zip_files import start_extract_zip


# 检查并安装必要的依赖
def install_dependencies():
    dependencies = ['pytesseract', 'pillow']
    installed = False

    # 检查并安装pytesseract
    try:
        import pytesseract
    except ImportError:
        print("正在安装必要的依赖库 pytesseract...")
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", "pytesseract"])
            print("pytesseract 安装成功！")
            import pytesseract
            installed = True
        except Exception as e:
            print(f"安装 pytesseract 失败: {str(e)}")
            print("请手动安装: pip install pytesseract")
            sys.exit(1)

    # 检查并安装Pillow
    try:
        import PIL
    except ImportError:
        print("正在安装必要的依赖库 Pillow...")
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", "pillow"])
            print("Pillow 安装成功！")
            installed = True
        except Exception as e:
            print(f"安装 Pillow 失败: {str(e)}")
            print("请手动安装: pip install pillow")
            sys.exit(1)

    # 检查Tesseract OCR引擎
    try:
        # 尝试自动查找Tesseract路径
        possible_paths = [
            r'C:\Program Files\Tesseract-OCR\tesseract.exe',
            r'D:\Tesseract-OCR\tesseract.exe',
            r'/usr/bin/tesseract',
            r'/usr/local/bin/tesseract'
        ]

        for path in possible_paths:
            if os.path.exists(path):
                pytesseract.pytesseract.tesseract_cmd = path
                print(f"找到Tesseract: {path}")
                break
        else:
            # 如果自动查找失败，尝试环境变量路径
            try:
                pytesseract.get_tesseract_version()
            except EnvironmentError:
                print("未找到Tesseract OCR引擎，请按以下步骤安装:")
                print("1. Windows: 下载安装包 https://github.com/UB-Mannheim/tesseract/wiki")
                print("2. macOS: brew install tesseract")
                print("3. Linux: sudo apt install tesseract-ocr")
                print("安装后请确保tesseract命令在系统路径中")
                return False
    except Exception as e:
        print(f"Tesseract检查失败: {str(e)}")
        return False

    if installed:
        print("所有依赖已成功安装，请重新运行脚本")
        return True
    return True


# 安装依赖
if not install_dependencies():
    print("依赖安装失败，请手动安装必要组件")
    sys.exit(1)

# 忽略openpyxl的样式警告
import warnings

warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl.styles.stylesheet')

# 图片处理设置
IMAGE_HEIGHT = 120  # 图片高度（像素）
IMAGE_MARGIN = 15  # 图片间距（像素）


def contains_ng_text(img_path):
    """
    严格验证图片中是否包含'NG'文字，排除无文字图片
    """
    try:
        # 打开图片
        img = PILImage.open(img_path)

        # 方法0: 检查图片是否完全是空白或噪点（无实际内容）
        if is_blank_image(img):
            print(f"跳过空白图片: {os.path.basename(img_path)}")
            return False

        # 方法1: 直接OCR识别
        ocr_result = pytesseract.image_to_string(img, lang='eng', config='--psm 6')
        if re.search(r'\bNG\b', ocr_result, re.IGNORECASE):
            return True

        # 方法2: 使用OCR获取文本位置信息，专门检测"NG"区域
        ocr_data = pytesseract.image_to_data(img, output_type=Output.DICT, lang='eng', config='--psm 6')
        total_text = ""
        ng_detected = False

        for i, text in enumerate(ocr_data['text']):
            text = text.strip()
            total_text += text + " "

            # 检查置信度
            conf = float(ocr_data['conf'][i]) if ocr_data['conf'][i] != '-1' else 0

            # 检测NG文本
            if re.search(r'\bNG\b', text, re.IGNORECASE) and conf > 60:
                ng_detected = True
                break

        if ng_detected:
            return True

        # 检查总文本长度 - 如果几乎没有文字，则不是NG图片
        clean_text = re.sub(r'\s+', '', total_text)  # 移除所有空白字符
        if len(clean_text) < 3:  # 少于3个字符视为无文字
            print(f"跳过无文字图片: {os.path.basename(img_path)}")
            return False

        # 方法3: 图像预处理后再次识别
        # 转换为灰度图
        gray = img.convert('L')
        # 二值化处理
        threshold = 150
        binary = gray.point(lambda p: p > threshold and 255)
        # 再次OCR识别
        ocr_result = pytesseract.image_to_string(binary, lang='eng', config='--psm 6')
        if re.search(r'\bNG\b', ocr_result, re.IGNORECASE):
            return True

        # 如果所有OCR方法都失败，检查文件名是否包含'NG'
        filename = os.path.basename(img_path).upper()
        if 'NG' in filename and not any(word in filename for word in ['ANG', 'ING', 'ONG', 'UNG']):
            print(f"警告: 基于文件名包含NG但未识别内容: {filename}")
            return True

        # 最终检查：如果没有任何文字，排除图片
        if len(clean_text) == 0:
            print(f"跳过无内容图片: {os.path.basename(img_path)}")
            return False

        return False
    except Exception as e:
        print(f"OCR处理失败: {str(e)}")
        # 如果OCR失败，严格检查文件名
        filename = os.path.basename(img_path).upper()
        return 'NG' in filename and not any(word in filename for word in ['ANG', 'ING', 'ONG', 'UNG'])


def find_ng_images(sn, data_dir='data'):
    """
    在data目录中查找包含指定SN和"NG"的图片，增加多重过滤
    """
    matched_images = []
    sn_str = str(sn).strip() if sn is not None else ""
    min_file_size = 10 * 1024  # 10KB最小文件大小

    if not sn_str:
        return matched_images

    # 检查data_dir是否存在
    if not os.path.exists(data_dir):
        print(f"目录 {data_dir} 不存在，跳过图片搜索")
        return matched_images

    # 构建搜索模式
    patterns = [
        f"*{re.escape(sn_str)}*NG*.jpg",
        f"*{re.escape(sn_str)}*NG*.jpeg",
        f"*{re.escape(sn_str)}*NG*.png"
    ]

    # 递归搜索所有匹配的图片
    for pattern in patterns:
        for img_path in glob.glob(os.path.join(data_dir, "**", pattern), recursive=True):
            # 检查文件大小
            if os.path.getsize(img_path) < min_file_size:
                print(f"跳过小文件: {os.path.basename(img_path)} (大小: {os.path.getsize(img_path) / 1024:.1f}KB)")
                continue

            # 检查是否为有效图片
            try:
                with PILImage.open(img_path) as img:
                    img.verify()  # 验证图片完整性
            except Exception as e:
                print(f"跳过损坏图片: {os.path.basename(img_path)} - {str(e)}")
                continue

            matched_images.append(img_path)

    # 添加不区分大小写的匹配
    if not matched_images:
        for root, dirs, files in os.walk(data_dir):
            for file in files:
                if file.lower().endswith(('.jpg', '.jpeg', '.png')):
                    file_lower = file.lower()
                    if "ng" in file_lower and sn_str.lower() in file_lower:
                        img_path = os.path.join(root, file)
                        # 检查文件大小
                        if os.path.getsize(img_path) < min_file_size:
                            print(f"跳过小文件: {file} (大小: {os.path.getsize(img_path) / 1024:.1f}KB)")
                            continue
                        # 检查是否为有效图片
                        try:
                            with PILImage.open(img_path) as img:
                                img.verify()  # 验证图片完整性
                        except Exception as e:
                            print(f"跳过损坏图片: {file} - {str(e)}")
                            continue
                        matched_images.append(img_path)

    # 使用严格验证检查图片是否确实包含NG
    verified_images = []
    for img_path in matched_images:
        if contains_ng_text(img_path):
            verified_images.append(img_path)
        else:
            print(f"跳过图片（未检测到NG）: {os.path.basename(img_path)}")

    return verified_images


def is_blank_image(img):
    """
    检测图片是否完全是空白或噪点（无实际内容）
    """
    try:
        # 转换为灰度图
        gray = img.convert('L')

        # 计算像素值方差
        pixels = list(gray.getdata())
        mean = sum(pixels) / len(pixels)
        variance = sum((p - mean) ** 2 for p in pixels) / len(pixels)

        # 如果方差很小，说明图片很均匀（可能是空白）
        if variance < 100:  # 经验值，可根据需要调整
            return True

        # 检查是否有大量相同颜色的像素
        color_counts = {}
        for pixel in pixels:
            color_counts[pixel] = color_counts.get(pixel, 0) + 1

        # 如果某个颜色占比超过95%，可能是空白图片
        max_count = max(color_counts.values())
        if max_count / len(pixels) > 0.95:
            return True

        return False
    except Exception as e:
        print(f"空白检测失败: {str(e)}")
        return False


def calculate_image_size(img_path, target_height):
    """
    计算调整后的图片大小（保持宽高比）
    """
    try:
        with PILImage.open(img_path) as img:
            # 保持宽高比调整大小
            width_ratio = target_height / img.height
            new_width = int(img.width * width_ratio)
            return new_width, target_height
    except Exception as e:
        print(f"计算图片大小失败: {str(e)}")
        return 100, 100  # 默认大小


def insert_images_horizontally(worksheet, row_idx, start_col, image_paths):
    """
    终极优化：精确控制图片行高度，消除多余空白
    """
    if not image_paths:
        return {}, []

    col_widths = {}
    x_offset = 0
    max_img_height = 0
    images_added = []

    # 先计算所有图片的尺寸
    img_sizes = []
    for img_path in image_paths:
        try:
            width, height = calculate_image_size(img_path, IMAGE_HEIGHT)
            img_sizes.append((width, height))
            if height > max_img_height:
                max_img_height = height
        except Exception as e:
            print(f"计算图片尺寸失败: {str(e)}")
            continue

    # 设置行高（关键修改：使用精确计算方式）
    if max_img_height > 0:
        # Excel行高单位换算：1单位 ≈ 1/7.2毫米 ≈ 1.33像素
        # 精确计算公式：行高 = (图片高度 + 上边距) / 1.33
        exact_row_height = (max_img_height + IMAGE_MARGIN) / 1.33
        worksheet.row_dimensions[row_idx].height = exact_row_height
        print(f"精确设置行 {row_idx} 高度: {exact_row_height:.2f} (像素高度: {max_img_height}+{IMAGE_MARGIN})")

    # 插入图片
    for idx, (img_path, (width, height)) in enumerate(zip(image_paths, img_sizes)):
        try:
            current_col = start_col + idx
            col_widths[current_col] = width

            img = Image(img_path)
            img.width = width
            img.height = height

            # 关键修改：使用单元格锚定方式（消除额外空白）
            cell_anchor = f"{get_column_letter(current_col)}{row_idx}"
            img.anchor = cell_anchor
            worksheet.add_image(img)
            images_added.append(img)

            x_offset += width + IMAGE_MARGIN
        except Exception as e:
            print(f"插入图片失败: {str(e)}")

    return col_widths, images_added


def apply_image_dimensions(worksheet, col_widths):
    """
    应用图片尺寸到工作表列宽
    """
    # 像素到Excel列宽单位的转换因子
    PIXELS_TO_EXCEL_UNITS = 0.14

    for col_idx, width_px in col_widths.items():
        # 计算列宽（Excel单位）
        col_width = max(15, width_px * PIXELS_TO_EXCEL_UNITS)
        col_letter = get_column_letter(col_idx)
        worksheet.column_dimensions[col_letter].width = col_width
        print(f"设置列 {col_letter} 宽度为: {col_width} (基于图片宽度: {width_px} 像素)")


def format_excel(worksheet):
    """
    格式化Excel表格，使其更易读，并优化行高
    """
    # 移除原有的行高设置逻辑，因为已经在插入时精确设置
    for row in range(2, worksheet.max_row + 1):
        # 只设置没有图片的行的默认高度
        if not any(img.anchor.endswith(str(row)) for img in worksheet._images):
            worksheet.row_dimensions[row].height = 15  # 更小的默认行高

        # 单元格样式设置保持不变...
        # 设置基础列宽
        column_widths = {
            'A': 25,  # SN
            'B': 25,  # QPL-Station Name
            'C': 15,  # Gantry
            'D': 20,  # Time(end)
            'E': 25,  # locate picture
        }

        # 应用基础列宽
        for col_letter, width in column_widths.items():
            worksheet.column_dimensions[col_letter].width = width

        # 设置标题样式
        header_font = Font(bold=True, color="FFFFFF")
        header_fill = PatternFill(start_color="0070C0", end_color="0070C0", fill_type="solid")
        thin_border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )

        # 应用标题样式
        max_col = worksheet.max_column
        for col in range(1, max_col + 1):
            cell = worksheet.cell(row=1, column=col)
            cell.font = header_font
            cell.fill = header_fill
            cell.border = thin_border
            cell.alignment = Alignment(horizontal='center', vertical='center')

        # 设置数据行样式
        min_row_height = 20  # 默认最小行高（像素）

        # 设置行高 - 只对没有图片的行设置默认行高
        # 有图片的行已经在插入时设置了精确高度
        for row in range(2, worksheet.max_row + 1):
            # 检查该行是否有图片
            has_images = False
            for image in worksheet._images:
                # 从图片的anchor字符串中提取行号
                anchor_str = image.anchor
                if isinstance(anchor_str, str):
                    match = re.search(r'\d+$', anchor_str)
                    if match:
                        image_row = int(match.group())
                        if image_row == row:
                            has_images = True
                            break

            # 如果没有图片，设置默认行高
            if not has_images:
                worksheet.row_dimensions[row].height = min_row_height

            # 设置单元格样式
            for col in range(1, max_col + 1):
                cell = worksheet.cell(row=row, column=col)
                cell.border = thin_border
                cell.alignment = Alignment(vertical='center', horizontal='center', wrap_text=True)

                # 设置日期格式
                if col == 4 and isinstance(cell.value, datetime):
                    cell.number_format = 'yyyy-mm-dd hh:mm:ss'

        # 冻结首行
        worksheet.freeze_panes = 'A2'


def extract_columns(argv):
    # 创建result目录（如果不存在）
    result_dir = "result"
    os.makedirs(result_dir, exist_ok=True)

    # 创建图片目录
    images_dir = os.path.join(result_dir, "images")
    os.makedirs(images_dir, exist_ok=True)

    # # 查找最新的Detail文件
    # detail_files = glob.glob("*_Detail*.xlsx")
    # if not detail_files:
    #     raise FileNotFoundError("未找到任何Detail文件（模式：*_Detail*.xlsx）")

    # 按修改时间排序，获取最新的文件
    input_file = argv[1]
    print(f"使用文件: {input_file}")

    # 加载工作簿
    wb = load_workbook(input_file, data_only=True)

    # 查找包含"不良明细"的工作表
    target_sheet = None
    for sheet_name in wb.sheetnames:
        if "不良明细" in sheet_name:
            target_sheet = sheet_name
            break

    if not target_sheet:
        available_sheets = "\n".join(wb.sheetnames)
        raise ValueError(f"未找到包含'不良明细'的工作表。可用工作表有：\n{available_sheets}")

    print(f"找到目标工作表: {target_sheet}")

    # 获取目标工作表
    sheet = wb[target_sheet]

    # 查找列索引
    header_row = 1  # 假设标题在第一行
    col_index = {}
    required_columns = ['SN', 'Station Name', 'Time End']

    for cell in sheet[header_row]:
        if cell.value in required_columns:
            col_index[cell.value] = cell.column_letter

    # 检查是否找到所有列
    missing_cols = [col for col in required_columns if col not in col_index]

    if missing_cols:
        # 尝试大小写不敏感匹配
        col_index = {}
        for cell in sheet[header_row]:
            cell_value = str(cell.value).lower() if cell.value else ""
            for req_col in required_columns:
                if req_col.lower() == cell_value:
                    col_index[req_col] = cell.column_letter

        # 再次检查
        missing_cols = [col for col in required_columns if col not in col_index]

        if missing_cols:
            available_cols = [cell.value for cell in sheet[header_row] if cell.value]
            raise ValueError(f"以下列在目标工作表中不存在: {', '.join(missing_cols)}\n"
                             f"可用列: {', '.join(filter(None, available_cols))}")

    # 创建新工作簿
    new_wb = Workbook()
    new_sheet = new_wb.active
    new_sheet.title = "不良明细汇总"

    # 添加新标题（按指定顺序）
    new_headers = [
        'SN',
        'QPL-Station Name',
        'Gantry',
        'Time(end)',
        'locate picture',
        'NG picture'
    ]

    # 写入新标题
    for col_idx, header in enumerate(new_headers, start=1):
        new_sheet.cell(row=1, column=col_idx, value=header)

    # 存储所有列的最大宽度
    all_col_widths = {}

    # 复制数据
    row_idx = 2  # 数据从第2行开始
    for src_row in range(header_row + 1, sheet.max_row + 1):
        # 获取所需列的值
        sn_value = sheet[f"{col_index['SN']}{src_row}"].value
        station_value = sheet[f"{col_index['Station Name']}{src_row}"].value
        time_value = sheet[f"{col_index['Time End']}{src_row}"].value

        # 跳过空行
        if not any([sn_value, station_value, time_value]):
            continue

        try:
            # 写入基础数据
            new_sheet.cell(row=row_idx, column=1, value=sn_value)
            new_sheet.cell(row=row_idx, column=2, value=station_value)
            new_sheet.cell(row=row_idx, column=3, value=None)  # Gantry
            new_sheet.cell(row=row_idx, column=4, value=time_value)
            new_sheet.cell(row=row_idx, column=5, value=None)  # locate picture

            # 查找并验证NG图片
            ng_images = []
            if sn_value:
                ng_images = find_ng_images(sn_value)

                if ng_images:
                    print(f"为SN {sn_value} 找到 {len(ng_images)} 张可能有的NG图片")

                    # 二次验证：确保图片确实包含NG
                    verified_images = []
                    for img_path in ng_images:
                        if contains_ng_text(img_path):
                            verified_images.append(img_path)

                    print(f"经过二次验证，{len(verified_images)} 张图片确认包含NG")

                    # 复制有效的NG图片到结果目录
                    copied_images = []
                    for img_path in verified_images:
                        try:
                            img_name = os.path.basename(img_path)
                            dest_path = os.path.join(images_dir, img_name)
                            shutil.copy2(img_path, dest_path)
                            copied_images.append(dest_path)
                        except Exception as e:
                            print(f"复制图片失败: {str(e)}")

                    # 插入图片到工作表（水平排列）
                    if copied_images:
                        # 在NG picture列开始插入图片
                        col_widths, inserted_images = insert_images_horizontally(new_sheet, row_idx, 6, copied_images)

                        # 更新全局列宽记录
                        for col_idx, width in col_widths.items():
                            if col_idx not in all_col_widths or width > all_col_widths[col_idx]:
                                all_col_widths[col_idx] = width

            # 移动到下一行
            row_idx += 1
        except Exception as e:
            print(f"警告: 处理行 {src_row} 时出错 - {str(e)}")
            traceback.print_exc()
            row_idx += 1

    # 应用图片尺寸到列宽
    if all_col_widths:
        apply_image_dimensions(new_sheet, all_col_widths)

    # 格式化Excel表格（包含行高优化）
    if new_sheet.max_row > 1:  # 确保有数据行
        format_excel(new_sheet)

    # 保存新工作簿到result目录
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_file = os.path.join(result_dir, f"不良明细汇总_{timestamp}.xlsx")
    new_wb.save(output_file)

    print(f"成功创建新文件: {output_file}")
    print(f"处理了 {row_idx - 2} 条记录")
    print(f"源工作表: {target_sheet}")


if __name__ == "__main__":
    try:
        start_extract_zip()
        # 确保data目录存在
        if not os.path.exists("data"):
            print("警告: data目录不存在，将跳过图片搜索")
            os.makedirs("data", exist_ok=True)

        extract_columns(sys.argv)
    except Exception as e:
        print(f"发生错误: {str(e)}")
        traceback.print_exc()
        print("请确保:")
        print("1. 原始Excel文件存在且路径正确")
        print("2. 文件没有被其他程序占用")
        print("3. 包含'不良明细'的工作表存在")
        print("4. 工作表中包含SN, Station Name, Time End三列")
        print("5. Tesseract OCR已正确安装")
