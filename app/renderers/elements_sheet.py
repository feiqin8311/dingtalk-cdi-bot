def render_elements_sheet(generator, elements_file_path: str, excel_file_path: str, image_path: str, wb=None, **kwargs):
    return generator.add_elements_sheet(elements_file_path, excel_file_path, image_path, wb=wb, **kwargs)
