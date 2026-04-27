import os
import shutil


def copy_template_file(template_path: str, output_dir: str, output_name: str) -> str:
    output_path = os.path.join(output_dir, output_name)
    shutil.copy(template_path, output_path)
    return output_path
