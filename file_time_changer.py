#!/usr/bin/env python3
"""
一键修改文件时间属性与照片 EXIF 拍摄时间。

支持的功能：
- 修改文件的访问时间（atime）、修改时间（mtime）、创建时间（ctime，Windows）
- 修改照片 EXIF 中的拍摄时间（DateTimeOriginal / DateTimeDigitized / DateTime）
- 支持单个文件、多个文件、目录及递归处理
- 预览模式（dry-run）：先看效果再执行
- 多种日期时间格式自动识别
- 从文件名推断日期（如 IMG_20240101_120000.jpg）
- 进度显示与操作统计
- 详细的错误提示和日志

依赖安装（可选）：
    pip install piexif      # 修改照片 EXIF
    pip install pywin32     # Windows 下修改创建时间（推荐）
"""

import os
import sys
import re
import argparse
import logging
from datetime import datetime
from pathlib import Path
from typing import Optional, List, Tuple, Dict

# ---------- 可选依赖检测 ----------
try:
    import win32file
    import pywintypes
    HAS_WIN32 = True
except ImportError:
    HAS_WIN32 = False

try:
    import piexif
    HAS_PIEXIF = True
except ImportError:
    HAS_PIEXIF = False

# ---------- 常量配置 ----------
IMAGE_EXTENSIONS = {".jpg", ".jpeg", ".tiff", ".tif", ".png", ".webp", ".heic", ".heif"}

# 支持的日期时间格式（按优先级排序，靠前的优先尝试）
DATETIME_FORMATS = [
    "%Y-%m-%d %H:%M:%S",
    "%Y-%m-%d %H:%M",
    "%Y/%m/%d %H:%M:%S",
    "%Y/%m/%d %H:%M",
    "%Y%m%d %H%M%S",
    "%Y%m%d_%H%M%S",
    "%Y-%m-%d",
    "%Y/%m/%d",
    "%Y%m%d",
]

# 从文件名提取日期的正则模式（按优先级排序）
FILENAME_DATE_PATTERNS = [
    # IMG_20240101_120000.jpg / VID_20240101_120000.mp4
    re.compile(r"(?P<y>20\d{2})(?P<m>\d{2})(?P<d>\d{2})[_\-](?P<H>\d{2})(?P<M>\d{2})(?P<S>\d{2})"),
    # 2024-01-01 12.00.00.jpg
    re.compile(r"(?P<y>20\d{2})[-_](?P<m>\d{2})[-_](?P<d>\d{2})[ _](?P<H>\d{2})[._](?P<M>\d{2})[._](?P<S>\d{2})"),
    # 20240101.jpg
    re.compile(r"(?P<y>20\d{2})(?P<m>\d{2})(?P<d>\d{2})"),
    # 2024-01-01.jpg / 2024_01_01.jpg
    re.compile(r"(?P<y>20\d{2})[-_](?P<m>\d{2})[-_](?P<d>\d{2})"),
]

# ---------- 日志配置 ----------
logger = logging.getLogger("file_time_changer")


def setup_logging(verbose: bool = False) -> None:
    """配置日志输出。"""
    level = logging.DEBUG if verbose else logging.INFO
    handler = logging.StreamHandler(sys.stdout)
    handler.setFormatter(logging.Formatter("%(levelname)s: %(message)s"))
    logger.setLevel(level)
    logger.addHandler(handler)
    logger.propagate = False


# ---------- 日期解析 ----------
def parse_datetime(date_str: str) -> Optional[datetime]:
    """尝试用多种格式解析日期时间字符串。"""
    date_str = date_str.strip()
    for fmt in DATETIME_FORMATS:
        try:
            return datetime.strptime(date_str, fmt)
        except ValueError:
            continue
    return None


def parse_date_from_filename(filename: str) -> Optional[datetime]:
    """从文件名中尝试提取日期时间。"""
    name = Path(filename).stem
    for pattern in FILENAME_DATE_PATTERNS:
        match = pattern.search(name)
        if match:
            groups = match.groupdict()
            try:
                year = int(groups.get("y", 0))
                month = int(groups.get("m", 0))
                day = int(groups.get("d", 0))
                hour = int(groups.get("H", 0))
                minute = int(groups.get("M", 0))
                second = int(groups.get("S", 0))
                if 1970 <= year <= 2100 and 1 <= month <= 12 and 1 <= day <= 31:
                    return datetime(year, month, day, hour, minute, second)
            except (ValueError, TypeError):
                continue
    return None


# ---------- 文件系统时间操作 ----------
def set_file_times(path: str, dt: datetime, set_ctime: bool = True, dry_run: bool = False) -> Tuple[bool, str]:
    """
    修改文件的时间属性。

    Returns:
        (是否成功, 描述信息)
    """
    timestamp = dt.timestamp()
    results = []

    try:
        if not dry_run:
            os.utime(path, (timestamp, timestamp))
        results.append("mtime/atime 已更新")
    except Exception as e:
        return False, f"修改 mtime/atime 失败: {e}"

    if set_ctime:
        if sys.platform == "win32":
            success, msg = _set_ctime_windows(path, dt, dry_run)
            results.append(msg)
            return success, "; ".join(results)
        elif sys.platform == "darwin":
            success, msg = _set_ctime_macos(path, dt, dry_run)
            results.append(msg)
            return success, "; ".join(results)
        else:
            results.append("ctime: Linux 下通常无法直接修改创建时间")
            return True, "; ".join(results)

    return True, "; ".join(results)


def _set_ctime_windows(path: str, dt: datetime, dry_run: bool) -> Tuple[bool, str]:
    """Windows 下修改文件创建时间。"""
    if dry_run:
        return True, "ctime 将更新（预览模式）"

    if HAS_WIN32:
        try:
            handle = win32file.CreateFile(
                path,
                win32file.GENERIC_WRITE,
                win32file.FILE_SHARE_READ | win32file.FILE_SHARE_WRITE | win32file.FILE_SHARE_DELETE,
                None,
                win32file.OPEN_EXISTING,
                win32file.FILE_FLAG_BACKUP_SEMANTICS,
                None,
            )
            ft = pywintypes.Time(dt.timestamp())
            win32file.SetFileTime(handle, ft, ft, ft)
            handle.Close()
            return True, "ctime 已更新（pywin32）"
        except Exception as e:
            return False, f"修改 ctime 失败（pywin32）: {e}"
    else:
        # ctypes 回退方案
        try:
            import ctypes
            from ctypes import wintypes

            kernel32 = ctypes.windll.kernel32

            def dt_to_filetime(dt_val: datetime) -> int:
                delta = dt_val - datetime(1601, 1, 1)
                return int(delta.total_seconds() * 10_000_000)

            ft = dt_to_filetime(dt)
            handle = kernel32.CreateFileW(
                path, 0x40000000, 0x7, None, 3, 0x80, None
            )
            if handle != -1:
                kernel32.SetFileTime(
                    handle,
                    ctypes.byref(ctypes.c_ulonglong(ft)),
                    ctypes.byref(ctypes.c_ulonglong(ft)),
                    ctypes.byref(ctypes.c_ulonglong(ft)),
                )
                kernel32.CloseHandle(handle)
                return True, "ctime 已更新（ctypes，可能需要管理员权限）"
            else:
                return False, "无法打开文件（可能需要管理员权限）"
        except Exception as e:
            return False, f"修改 ctime 失败（ctypes）: {e}"


def _set_ctime_macos(path: str, dt: datetime, dry_run: bool) -> Tuple[bool, str]:
    """macOS 下修改文件创建时间（使用 SetFile 命令）。"""
    import subprocess

    # SetFile 格式：MM/DD/YYYY HH:MM:SS
    date_str = dt.strftime("%m/%d/%Y %H:%M:%S")
    cmd = ["SetFile", "-d", date_str, path]

    if dry_run:
        return True, "ctime 将更新（SetFile，预览模式）"

    try:
        result = subprocess.run(cmd, capture_output=True, text=True)
        if result.returncode == 0:
            return True, "ctime 已更新（SetFile）"
        else:
            return False, f"修改 ctime 失败（SetFile）: {result.stderr.strip()}"
    except FileNotFoundError:
        return False, "未找到 SetFile 命令（请安装 Xcode Command Line Tools）"
    except Exception as e:
        return False, f"修改 ctime 失败: {e}"


# ---------- EXIF 时间操作 ----------
def set_exif_times(path: str, dt: datetime, dry_run: bool = False) -> Tuple[bool, str]:
    """修改照片的 EXIF 拍摄时间。"""
    ext = Path(path).suffix.lower()
    if ext not in IMAGE_EXTENSIONS:
        return False, f"不支持的图片格式: {ext}"

    if not HAS_PIEXIF:
        return False, "未安装 piexif，请运行: pip install piexif"

    if dry_run:
        return True, "EXIF 时间将更新（预览模式）"

    try:
        try:
            exif_dict = piexif.load(path)
        except Exception:
            exif_dict = {"0th": {}, "Exif": {}, "GPS": {}, "1st": {}, "thumbnail": None}

        date_str = dt.strftime("%Y:%m:%d %H:%M:%S")

        if "Exif" not in exif_dict:
            exif_dict["Exif"] = {}
        if "0th" not in exif_dict:
            exif_dict["0th"] = {}

        exif_dict["Exif"][piexif.ExifIFD.DateTimeOriginal] = date_str
        exif_dict["Exif"][piexif.ExifIFD.DateTimeDigitized] = date_str
        exif_dict["0th"][piexif.ImageIFD.DateTime] = date_str

        exif_bytes = piexif.dump(exif_dict)
        piexif.insert(exif_bytes, path)
        return True, "EXIF 时间已更新"
    except Exception as e:
        return False, f"写入 EXIF 失败: {e}"


# ---------- 文件收集 ----------
def collect_files(paths: List[str], recursive: bool = False) -> List[str]:
    """收集所有需要处理的文件路径。"""
    files = []
    for p in paths:
        p = os.path.abspath(p)
        if os.path.isfile(p):
            files.append(p)
        elif os.path.isdir(p):
            if recursive:
                for root, _, filenames in os.walk(p):
                    for f in filenames:
                        files.append(os.path.join(root, f))
            # 目录本身不加入文件列表，单独处理
        else:
            logger.warning(f"路径不存在: {p}")
    return files


def collect_dirs(paths: List[str]) -> List[str]:
    """收集所有需要处理的目录路径。"""
    dirs = []
    for p in paths:
        p = os.path.abspath(p)
        if os.path.isdir(p):
            dirs.append(p)
    return dirs


# ---------- 统计信息 ----------
class Stats:
    """操作统计。"""

    def __init__(self) -> None:
        self.total_files = 0
        self.total_dirs = 0
        self.fs_success = 0
        self.fs_failed = 0
        self.exif_success = 0
        self.exif_failed = 0
        self.exif_skipped = 0
        self.errors: List[str] = []

    def add_error(self, msg: str) -> None:
        self.errors.append(msg)

    def summary(self) -> str:
        lines = [
            "",
            "=" * 50,
            "处理完成统计",
            "=" * 50,
            f"处理文件数: {self.total_files}",
            f"处理目录数: {self.total_dirs}",
            f"文件系统时间 - 成功: {self.fs_success}, 失败: {self.fs_failed}",
            f"EXIF 时间 - 成功: {self.exif_success}, 失败: {self.exif_failed}, 跳过: {self.exif_skipped}",
        ]
        if self.errors:
            lines.append("")
            lines.append(f"错误列表（共 {len(self.errors)} 条）:")
            for i, err in enumerate(self.errors[:10], 1):
                lines.append(f"  {i}. {err}")
            if len(self.errors) > 10:
                lines.append(f"  ... 还有 {len(self.errors) - 10} 条错误")
        lines.append("=" * 50)
        return "\n".join(lines)


# ---------- 核心处理 ----------
def process_file(
    path: str,
    dt: datetime,
    set_ctime: bool = True,
    set_exif: bool = True,
    dry_run: bool = False,
) -> Tuple[bool, bool]:
    """
    处理单个文件。

    Returns:
        (fs_ok, exif_ok) — 文件系统时间是否成功，EXIF 是否成功
    """
    fs_ok = False
    exif_ok = False
    exif_skipped = False

    # 修改文件系统时间
    fs_ok, fs_msg = set_file_times(path, dt, set_ctime, dry_run)
    if fs_ok:
        logger.info(f"  [FS] {os.path.basename(path)} — {fs_msg}")
    else:
        logger.error(f"  [FS] {os.path.basename(path)} — {fs_msg}")

    # 修改 EXIF 时间
    if set_exif and Path(path).suffix.lower() in IMAGE_EXTENSIONS:
        exif_ok, exif_msg = set_exif_times(path, dt, dry_run)
        if exif_ok:
            logger.info(f"  [EXIF] {os.path.basename(path)} — {exif_msg}")
        else:
            logger.warning(f"  [EXIF] {os.path.basename(path)} — {exif_msg}")
    else:
        exif_skipped = True

    return fs_ok, exif_ok or exif_skipped


def process_all(
    paths: List[str],
    dt: datetime,
    recursive: bool = False,
    set_ctime: bool = True,
    set_exif: bool = True,
    dry_run: bool = False,
) -> Stats:
    """处理所有路径。"""
    stats = Stats()

    # 收集文件
    files = collect_files(paths, recursive)
    dirs = collect_dirs(paths)
    stats.total_files = len(files)
    stats.total_dirs = len(dirs)

    if dry_run:
        logger.info("=== 预览模式（不会实际修改） ===")
    else:
        logger.info(f"=== 开始处理（共 {len(files)} 个文件，{len(dirs)} 个目录） ===")

    # 处理文件
    for i, f in enumerate(files, 1):
        logger.info(f"[{i}/{len(files)}] {f}")
        fs_ok, exif_ok = process_file(f, dt, set_ctime, set_exif, dry_run)
        if fs_ok:
            stats.fs_success += 1
        else:
            stats.fs_failed += 1
            stats.add_error(f"FS 失败: {f}")
        if set_exif and Path(f).suffix.lower() in IMAGE_EXTENSIONS:
            if exif_ok:
                stats.exif_success += 1
            else:
                stats.exif_failed += 1
                stats.add_error(f"EXIF 失败: {f}")
        else:
            stats.exif_skipped += 1

    # 处理目录自身的时间
    for d in dirs:
        logger.info(f"[目录] {d}")
        fs_ok, _ = process_file(d, dt, set_ctime, False, dry_run)
        if fs_ok:
            stats.fs_success += 1
        else:
            stats.fs_failed += 1
            stats.add_error(f"FS 失败（目录）: {d}")

    return stats


# ---------- 命令行入口 ----------
def build_parser() -> argparse.ArgumentParser:
    """构建命令行参数解析器。"""
    parser = argparse.ArgumentParser(
        description="一键修改文件时间属性与照片 EXIF 拍摄时间",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例:
  # 修改单个文件的时间
  %(prog)s photo.jpg -d "2024-01-01 12:00:00"

  # 预览模式：先看效果再决定
  %(prog)s ./photos -d "2024-01-01" -r --dry-run

  # 递归处理目录，同时修改文件系统时间和 EXIF
  %(prog)s ./photos -d "2024-06-15 14:30:00" -r

  # 只修改文件系统时间，不修改 EXIF
  %(prog)s ./docs -d "2024-01-01" --no-exif

  # 从文件名推断日期（如 IMG_20240101_120000.jpg）
  %(prog)s ./photos --from-filename -r

  # 从文件的修改时间作为目标时间
  %(prog)s ./photos --from-mtime -r --set-exif
        """,
    )

    # 路径参数
    parser.add_argument(
        "path",
        nargs="+",
        help="文件或目录路径，可多个",
    )

    # 日期时间来源（互斥组）
    date_group = parser.add_mutually_exclusive_group()
    date_group.add_argument(
        "-d",
        "--datetime",
        help='目标日期时间，如 "2024-01-01 12:00:00"，支持多种格式自动识别',
    )
    date_group.add_argument(
        "--from-filename",
        action="store_true",
        help="从文件名推断日期时间（如 IMG_20240101_120000.jpg）",
    )
    date_group.add_argument(
        "--from-mtime",
        action="store_true",
        help="使用文件当前的修改时间作为目标时间（常用于同步 EXIF）",
    )

    # 处理选项
    parser.add_argument(
        "-r",
        "--recursive",
        action="store_true",
        help="递归处理目录下的所有文件",
    )
    parser.add_argument(
        "--no-ctime",
        action="store_true",
        help="不修改创建时间（Windows/macOS 有效）",
    )
    parser.add_argument(
        "--no-exif",
        action="store_true",
        help="不修改照片 EXIF 时间",
    )
    parser.add_argument(
        "--set-exif-only",
        action="store_true",
        help="只修改 EXIF 时间，不修改文件系统时间",
    )

    # 其他选项
    parser.add_argument(
        "-n",
        "--dry-run",
        action="store_true",
        help="预览模式，不实际修改任何内容",
    )
    parser.add_argument(
        "-v",
        "--verbose",
        action="store_true",
        help="显示详细调试信息",
    )

    return parser


def resolve_target_datetime(
    args: argparse.Namespace,
    file_path: Optional[str] = None,
) -> Optional[datetime]:
    """根据参数解析目标日期时间。"""
    if args.datetime:
        dt = parse_datetime(args.datetime)
        if dt is None:
            logger.error(f"无法解析日期时间: {args.datetime}")
            logger.error("支持的格式示例:")
            for fmt in DATETIME_FORMATS:
                example = datetime(2024, 1, 15, 14, 30, 0).strftime(fmt)
                logger.error(f"  {fmt}  →  {example}")
        return dt

    elif args.from_filename and file_path:
        dt = parse_date_from_filename(file_path)
        if dt is None:
            logger.debug(f"无法从文件名提取日期: {os.path.basename(file_path)}")
        return dt

    elif args.from_mtime and file_path:
        try:
            mtime = os.path.getmtime(file_path)
            return datetime.fromtimestamp(mtime)
        except Exception as e:
            logger.error(f"读取文件修改时间失败: {e}")
            return None

    return None


def main() -> None:
    parser = build_parser()
    args = parser.parse_args()

    # 配置日志
    setup_logging(args.verbose)

    # 参数校验
    if not any([args.datetime, args.from_filename, args.from_mtime]):
        parser.error("请指定日期来源：-d/--datetime, --from-filename, 或 --from-mtime")

    # 依赖提示
    set_ctime = not args.no_ctime
    set_exif = not args.no_exif and not args.set_exif_only
    only_exif = args.set_exif_only

    if only_exif:
        set_ctime = False
        set_exif_flag = True
    else:
        set_exif_flag = set_exif

    if set_ctime and sys.platform == "win32" and not HAS_WIN32:
        logger.warning("未安装 pywin32，修改创建时间将使用 ctypes（可能需要管理员权限）")
        logger.warning("安装命令: pip install pywin32")

    if set_exif_flag and not HAS_PIEXIF:
        logger.warning("未安装 piexif，无法修改照片 EXIF。请运行: pip install piexif")

    # 处理逻辑
    if args.from_filename or args.from_mtime:
        # 每个文件有自己的时间，需要逐个处理
        stats = Stats()
        files = collect_files(args.path, args.recursive)
        dirs = collect_dirs(args.path)
        stats.total_files = len(files)
        stats.total_dirs = len(dirs)

        if args.dry_run:
            logger.info("=== 预览模式（不会实际修改） ===")
        else:
            logger.info(f"=== 开始处理（共 {len(files)} 个文件） ===")

        for i, f in enumerate(files, 1):
            logger.info(f"[{i}/{len(files)}] {f}")
            dt = resolve_target_datetime(args, f)
            if dt is None:
                logger.warning("  跳过：无法确定目标时间")
                stats.fs_failed += 1
                stats.add_error(f"无法确定时间: {f}")
                continue

            fs_ok, exif_ok = process_file(f, dt, set_ctime if not only_exif else False, set_exif_flag, args.dry_run)
            if fs_ok or only_exif:
                if not only_exif:
                    stats.fs_success += 1
            else:
                stats.fs_failed += 1
                stats.add_error(f"FS 失败: {f}")
            if set_exif_flag and Path(f).suffix.lower() in IMAGE_EXTENSIONS:
                if exif_ok:
                    stats.exif_success += 1
                else:
                    stats.exif_failed += 1
                    stats.add_error(f"EXIF 失败: {f}")
            else:
                stats.exif_skipped += 1

        print(stats.summary())

    else:
        # 所有文件使用统一时间
        dt = resolve_target_datetime(args)
        if dt is None:
            sys.exit(1)

        logger.info(f"目标时间: {dt.strftime('%Y-%m-%d %H:%M:%S')}")

        stats = process_all(
            args.path,
            dt,
            recursive=args.recursive,
            set_ctime=set_ctime if not only_exif else False,
            set_exif=set_exif_flag,
            dry_run=args.dry_run,
        )
        print(stats.summary())

    if args.dry_run:
        logger.info("以上为预览结果，未实际修改任何文件。去掉 --dry-run 即可执行。")


if __name__ == "__main__":
    main()
