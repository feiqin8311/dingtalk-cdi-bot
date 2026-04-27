def render_detailed_list(generator, excel_file_path: str, image_path: str, wb=None, **kwargs):
    return generator.add_detailed_list_sheet(excel_file_path, image_path, wb=wb, **kwargs)
