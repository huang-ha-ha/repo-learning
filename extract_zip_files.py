import os
import zipfile
import shutil
from pathlib import Path
import time


def move_current_dir_zips_to_zip_dir():
    """
    将当前目录下的所有ZIP文件移动到zip目录中
    """
    moved_count = 0
    # 获取当前目录下的所有文件
    for file in os.listdir('.'):
        if file.lower().endswith('.zip'):
            src_path = os.path.join('.', file)
            dest_path = os.path.join('zip', file)

            # 创建zip目录（如果不存在）
            os.makedirs('zip', exist_ok=True)

            # 移动文件
            shutil.move(src_path, dest_path)
            print(f"已将 {file} 移动到 zip 目录")
            moved_count += 1

    return moved_count


def remove__MACOSX_directories(directory):
    """
    递归删除目录中的所有__MACOSX目录
    """
    removed_count = 0
    for root, dirs, _ in os.walk(directory):
        # 注意：我们需要在遍历过程中修改dirs列表
        if '__MACOSX' in dirs:
            macosx_dir = os.path.join(root, '__MACOSX')
            try:
                shutil.rmtree(macosx_dir)
                print(f"已删除: {os.path.relpath(macosx_dir, directory)}")
                removed_count += 1
                # 从当前遍历中移除这个目录，避免后续遍历
                dirs.remove('__MACOSX')
            except Exception as e:
                print(f"删除失败 {macosx_dir}: {str(e)}")

    return removed_count


def recursive_unzip(source_dir, target_dir):
    """
    递归解压source_dir中的所有ZIP文件到target_dir对应位置
    """
    total_extracted = 0

    # 遍历源目录中的所有项目
    for item in os.listdir(source_dir):
        source_path = os.path.join(source_dir, item)
        target_path = os.path.join(target_dir, item)
        relative_path = os.path.relpath(source_path, "zip")

        # 如果是ZIP文件
        if os.path.isfile(source_path) and item.lower().endswith('.zip'):
            # 创建目标目录（使用ZIP文件名作为目录名）
            folder_name = Path(item).stem
            extract_dir = os.path.join(target_dir, folder_name)
            os.makedirs(extract_dir, exist_ok=True)

            print(f"解压: {relative_path} -> {os.path.relpath(extract_dir, 'data')}")

            try:
                # 解压ZIP文件
                with zipfile.ZipFile(source_path, 'r') as zip_ref:
                    zip_ref.extractall(extract_dir)

                total_extracted += 1

                removed = remove__MACOSX_directories(extract_dir)
                if removed > 0:
                    print(f"已删除 {removed} 个__MACOSX目录")

                # 递归处理新解压的目录
                total_extracted += recursive_unzip(extract_dir, extract_dir)

            except Exception as e:
                print(f"处理 {relative_path} 时出错: {str(e)}")

        # 如果是目录，递归处理
        elif os.path.isdir(source_path):
            # 在目标目录中创建对应的子目录
            os.makedirs(target_path, exist_ok=True)
            total_extracted += recursive_unzip(source_path, target_path)

    return total_extracted


def remove_zip_files(directory):
    """
    递归删除目录中的所有ZIP文件
    """
    removed_count = 0
    for root, _, files in os.walk(directory):
        for file in files:
            if file.lower().endswith('.zip'):
                zip_path = os.path.join(root, file)
                try:
                    os.remove(zip_path)
                    print(f"已删除: {os.path.relpath(zip_path, directory)}")
                    removed_count += 1
                except Exception as e:
                    print(f"删除失败 {zip_path}: {str(e)}")

    return removed_count


def start_extract_zip():
    start_time = time.time()

    print("=" * 50)
    print("ZIP文件递归解压工具")
    print("=" * 50)
    print("将解压zip目录中的所有ZIP文件到data目录")
    print("保持相同的目录结构，并删除data目录中的ZIP文件")

    # 确保目录存在
    os.makedirs("zip", exist_ok=True)
    os.makedirs("data", exist_ok=True)

    print("\n移动当前目录的ZIP文件到zip目录...")
    moved_count = move_current_dir_zips_to_zip_dir()
    print(f"已移动 {moved_count} 个ZIP文件到zip目录")

    # 检查是否有ZIP文件
    zip_files = []
    for root, _, files in os.walk("zip"):
        for file in files:
            if file.lower().endswith('.zip'):
                zip_files.append(os.path.join(root, file))

    if not zip_files:
        print("\n错误: zip目录中没有找到任何ZIP文件")
        print("请将ZIP文件放入zip目录中")
        return

    print(f"\n在zip目录中找到 {len(zip_files)} 个ZIP文件")

    # 递归解压所有文件
    total_extracted = recursive_unzip("zip", "data")

    print("\n最终清理所有__MACOSX目录...")
    total_removed = remove__MACOSX_directories("data")
    print(f"已删除 {total_removed} 个__MACOSX目录")

    # 删除data目录中的所有ZIP文件
    print("\n删除data目录中的ZIP文件...")
    removed_count = remove_zip_files("data")

    # 计算耗时
    end_time = time.time()
    elapsed = end_time - start_time

    print("\n" + "=" * 50)
    print(f"处理完成! 共解压 {total_extracted} 个ZIP文件")
    print(f"删除 {removed_count} 个ZIP文件")
    print(f"总耗时: {elapsed:.2f} 秒")
    print("=" * 50)

    # 显示目录结构
    print("\n最终目录结构:")
    print("zip/")
    for root, dirs, files in os.walk("zip"):
        level = root.replace("zip", '').count(os.sep)
        indent = '│   ' * level
        print(f"{indent}├── {os.path.basename(root)}/")
        for f in files:
            print(f"{indent}│   ├── {f}")

    print("\ndata/")
    for root, dirs, files in os.walk("data"):
        # 跳过根目录
        if root == "data":
            continue

        level = root.replace("data", '').count(os.sep)
        indent = '│   ' * level
        print(f"{indent}├── {os.path.basename(root)}/")
        for f in files:
            print(f"{indent}│   ├── {f}")


if __name__ == "__main__":
    start_extract_zip()
