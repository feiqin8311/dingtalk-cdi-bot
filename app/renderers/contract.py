def render_contract(generator, excel_file_path: str, image_path: str, wb=None, **kwargs):
    return generator.add_contract_sheet(excel_file_path, image_path, wb=wb, **kwargs)
