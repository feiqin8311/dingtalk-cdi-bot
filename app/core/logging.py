import datetime
import os
import sys


def get_log_dir(module_filename: str) -> str:
    base_dir = os.path.dirname(sys.executable) if getattr(sys, "frozen", False) else os.path.dirname(os.path.abspath(module_filename))
    return os.path.join(base_dir, "Logs")


def log_message(log_file_path: str, message: str) -> None:
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    log_line = f"[{timestamp}] {message}"
    print(log_line)
    try:
        with open(log_file_path, "a", encoding="utf-8") as f:
            f.write(log_line + "\n")
    except Exception as exc:
        print(f"写入日志文件失败: {exc}, 路径: {log_file_path}")
