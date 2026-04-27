def render_customs_declaration(generator, excel_file_path: str, image_path: str, wb=None, **kwargs):
    return generator.add_customs_declaration_sheet(excel_file_path, image_path, wb=wb, **kwargs)
