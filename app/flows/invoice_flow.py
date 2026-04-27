import datetime
import os

import openpyxl

from app.renderers.bill import render_bill
from app.renderers.common import ensure_default_sheet_removed, finalize_workbook_page_setup
from app.renderers.contract import render_contract
from app.renderers.customs_declaration import render_customs_declaration
from app.renderers.detailed_list import render_detailed_list
from app.renderers.elements_sheet import render_elements_sheet
from app.renderers.packing_list import render_packing_list


class InvoiceFlow:
    def __init__(self, generator):
        self.generator = generator

    @staticmethod
    def _log_timed_step(generator, label: str, start_time: datetime.datetime) -> None:
        logger = getattr(generator, "_log_timed_step", None)
        if callable(logger):
            logger(label, start_time)

    def generate_all_documents(self, invoice_filepath: str, product_filepath: str,
                               product_name_filepath: str, site_address_filepath: str,
                               elements_filepath: str, output_filepath: str, pictures_path: str,
                               flag: bool, generate_declaration_file: bool = True) -> None:
        generator = self.generator
        generator.lcl_flag = flag
        start_time = datetime.datetime.now()
        source_file = os.path.basename(invoice_filepath)

        try:
            generator.log_message(
                f"开始生成报关资料(发货单) | 发货单: {os.path.basename(invoice_filepath)}, "
                f"输出目录: {output_filepath}, 拼箱: {generator.lcl_flag}"
            )

            step_started_at = datetime.datetime.now()
            if flag:
                print("拼箱")
                generator.lcl_situation(
                    invoice_filepath, product_filepath, product_name_filepath, site_address_filepath
                )
            else:
                print("不拼箱")
                generator.process_invoice(
                    invoice_filepath, product_filepath, product_name_filepath, site_address_filepath
                )
            self._log_timed_step(generator, "准备报关数据", step_started_at)

            output_dir = output_filepath
            output_filepath = os.path.join(output_dir, f"{generator.shipping_details['invoice_number']} 报关资料.xlsx")
            generator.log_message(f"报关资料输出文件: {output_filepath}")

            step_started_at = datetime.datetime.now()
            wb = openpyxl.Workbook()
            render_detailed_list(generator, output_filepath, pictures_path, wb=wb)
            ensure_default_sheet_removed(wb)
            render_packing_list(generator, output_filepath, pictures_path, wb=wb)
            render_bill(generator, output_filepath, pictures_path, wb=wb)
            render_contract(generator, output_filepath, pictures_path, wb=wb)
            render_customs_declaration(generator, output_filepath, pictures_path, wb=wb)
            render_elements_sheet(generator, elements_filepath, output_filepath, pictures_path, wb=wb)
            finalize_workbook_page_setup(wb)
            wb.save(output_filepath)
            self._log_timed_step(generator, "渲染报关资料工作簿", step_started_at)

            if generate_declaration_file:
                step_started_at = datetime.datetime.now()
                declaration_template_path = generator.copy_declaration_template_file(output_dir)
                generator.fill_declaration_template_file(declaration_template_path)
                self._log_timed_step(generator, "填充申报模板", step_started_at)
            else:
                generator.log_message("申报文件开关为否，跳过申报模板复制与填充")

            step_started_at = datetime.datetime.now()
            generator.set_excel_page_view(output_filepath)
            self._log_timed_step(generator, "设置分页视图", step_started_at)
            generator.log_message("报关资料生成完成(发货单)")
            generator.append_call_record(
                module_name="发货单",
                action="生成报关资料",
                success=True,
                duration_seconds=(datetime.datetime.now() - start_time).total_seconds(),
                invoice_number=generator.shipping_details.get('invoice_number', ''),
                source_file=source_file,
            )
        except Exception as e:
            print(f"生成文档时出错: {e}")
            generator.log_message(f"报关资料生成失败(发货单): {e}")
            generator.append_call_record(
                module_name="发货单",
                action="生成报关资料",
                success=False,
                duration_seconds=(datetime.datetime.now() - start_time).total_seconds(),
                invoice_number=generator.shipping_details.get('invoice_number', ''),
                source_file=source_file,
                error_message=str(e),
            )
            raise

    def register_all_documents(self, *args, **kwargs):
        return self.generator.register_all_documents(*args, **kwargs)
