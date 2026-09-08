# 报关资料生成工具业务与代码逻辑

本文档记录本项目的业务逻辑和代码逻辑，便于后续维护、排查问题和新增规则。

## 1. 项目定位

本项目是一个基于 PyQt6 的桌面工具，用于根据发货单或备货单生成报关资料 Excel、PDF，并按需生成申报文件、登记出运明细和报关开票明细。

程序不是常驻服务。用户通过界面选择输入 Excel、输出目录和选项，然后点击“生成报关资料”。实际处理在后台线程中执行，避免界面卡死。

核心输入文件包括：

- 发货单或备货单，二选一。
- 产品信息单证专用.xlsx。
- 品名.xlsx。
- 全站点地址.xlsx。
- 要素.xlsx。
- 登记相关台账，如出运明细、各公司报关明细。

选择输出目录后，程序按两层目录写入文件：

- 每票创建 `{发票号} 报关资料` 文件夹，存放该票的报关资料。
- 申报模板直接放在用户选择的输出目录，不放入每票文件夹；同店铺后续订单复用同一模板并追加。

每票文件夹的核心输出包括：

- `{发票号} 报关资料.xlsx` 和 `{发票号} 报关资料.pdf`。
- `{发票号} 销售合同.pdf`，仅包含报关资料工作簿的“合同”Sheet。
- 可选写入登记台账。

## 2. 程序入口

界面入口是 `ui_pyqt.py`。

主要调用链：

1. `QTInterface.process_data()`
2. 收集界面文件路径和开关状态。
3. 判断使用发货单还是备货单。
4. 校验必填路径。
5. 创建 `ProcessingThread`。
6. 在线程中调用对应业务类。

发货单和备货单只能二选一：

- 填了发货单路径，不填备货单路径：走发货单模式。
- 填了备货单路径，不填发货单路径：走备货单模式。
- 两个都填：界面报“文件冲突”。
- 两个都不填：界面报“文件缺失”。

`ProcessingThread.run()` 中的分流：

- 发货单：实例化 `invoice_data.CustomsDeclaration()`。
- 备货单：实例化 `stock_preparation.CustomsDeclaration()`。

两条线都会先调用 `flow.generate_all_documents()` 生成报关资料。如果界面勾选“登记”，再调用 `flow.register_all_documents()` 写入登记台账。

## 3. 代码分层

当前项目同时保留了顶层大文件和拆分后的 `app/` 分层。

顶层文件：

- `ui_pyqt.py`：桌面界面、路径收集、线程调度。
- `invoice_data.py`：发货单模式的核心业务类。
- `stock_preparation.py`：备货单模式的核心业务类。

`app/flows`：

- `invoice_flow.py`：发货单生成报关资料的总流程。
- `stock_flow.py`：备货单生成报关资料的总流程。

`app/renderers`：

- `detailed_list.py`：明细单。
- `packing_list.py`：装箱单。
- `bill.py`：发票。
- `contract.py`：合同。
- `customs_declaration.py`：报关单。
- `elements_sheet.py`：要素。
- `common.py`：删除默认 Sheet、统一分页设置等公共渲染逻辑。

`app/services`：

- `freight_service.py`：运费计算和分摊口径。
- `declared_price_service.py`：申报单价列选择。
- `elements_service.py`：要素表分组匹配。
- `register_service.py`：登记表查重、查最后数据行等公共逻辑。

`app/core`：

- `excel_io.py`：Excel 读取缓存，避免同一文件重复读取。
- `page_setup.py`：Excel 分页预览和缩放设置。
- `logging.py`：日志写入。
- `metrics.py`：调用统计写入。
- `models.py`：登记目标选择等简单模型函数。
- `pdf_export.py`：通过 WPS/Excel 导出整本工作簿或指定 Sheet 的 PDF。
- `text_layout.py`：按文本换行估算行高，避免长地址在 PDF 中被截断。

## 4. 发货单模式业务逻辑

发货单模式入口类是 `invoice_data.CustomsDeclaration`。

主流程在 `app/flows/invoice_flow.py` 的 `InvoiceFlow.generate_all_documents()`：

1. 设置 `generator.lcl_flag`。
2. 根据拼箱开关调用：
   - 拼箱：`generator.lcl_situation()`。
   - 不拼箱：`generator.process_invoice()`。
3. 上一步会生成 `generator.result_df`。
4. 在所选输出目录下创建 `{发票号} 报关资料` 文件夹，并基于 `result_df` 渲染多个 Sheet。
5. 保存 `{发票号} 报关资料.xlsx`。
6. 如果开启申报文件，在所选输出目录复用或复制申报模板，并向模板追加本票数据。
7. 设置 Excel 分页视图。
8. 生成 `{发票号} 报关资料.pdf` 和 `{发票号} 销售合同.pdf`。
9. 写入调用统计。

### 4.1 发货单基础信息

发货单公共前置逻辑在 `invoice_data.py` 的 `get_basic_information()`。

它读取三类数据：

- 发货单 `装箱信息`。
- 发货单 `发货单详情`。
- `产品信息单证专用.xlsx`。

它写入 `self.shipping_details` 的关键字段：

- `today`：英文日期。
- `mark`：货件单号拼接，并追加 `MADE IN CHINA`。
- `country_zh`：中文国家。
- `country`：国家英文/编码，如美国映射为 `USA`。
- `invoice_number`：来自 `发货单详情` 的 `物流商单号`。
- `logistics_provider`：物流商。
- `Logistics_center`：物流中心编码。
- `recipient`：从 `全站点地址.xlsx` 按国家和目的港匹配。
- `brand`：从店铺字段截取并大写。
- 公司主体信息，如发票抬头、发货人、银行信息、报关单价 Sheet、申报模板等。

公司主体映射：

- `EZARC` -> 易达，`company_info_index = 0`。
- `LIBRATON` / `PACMANN` -> 鲁米亚，`company_info_index = 1`。
- 未匹配的默认走素桐，`company_info_index = 2`。
- `SIXPRAY` -> 六浪，`company_info_index = 3`。
- `TOLESA` -> 匠晟，`company_info_index = 4`。

体积兜底：

- 如果 `装箱信息` 里的 `总体积（m³）-箱子` 为 0，会改成 `0.01`。
- 这一步用于避免后续按体积分摊时出现 0。

### 4.2 发货单不拼箱

不拼箱走 `invoice_data.py` 的 `process_invoice()`。

主要步骤：

1. 调用 `get_basic_information()`。
2. 如果同一 SKU 出现多行，先按 SKU 汇总：
   - 汇总 `发货数量`、`箱数`、`总体积（m³）-箱子`、`总重量（kg）-箱子`。
   - 保留第一条 `单箱数量`、`箱子毛重（kg）`、`箱子长度（cm）`、`箱子宽度（cm）`、`品名`。
3. 计算整票运费。
4. 根据运费口径分摊到每个品名：
   - `volume`：按体积分摊。
   - `weight`：按重量分摊。
5. 合并产品资料和品名资料：
   - 报关中文品名。
   - 英文品名。
   - HS CODE。
   - 单位。
   - 材质。
   - 货源地。
   - 供应商。
   - 申报单价。
6. 计算报关字段：
   - 申报总价。
   - 含运费申报单价。
   - 含运费申报总价。
   - 总 PCS。
   - 是否含金刚石。
7. 按 `报关中文品名 + 是否含金刚石` 汇总生成 `self.result_df`。

不拼箱的特点：

- 箱数基本沿用原始装箱信息累计结果。
- 不按重量重新估箱数。
- 最终登记和报关资料以汇总后的 `self.result_df` 为准。

### 4.3 发货单拼箱

拼箱走 `invoice_data.py` 的 `lcl_situation()`。

拼箱的核心目标是：同一箱里有多个 SKU 或品名时，要把箱重、体积、箱数合理分摊到报关品名层级。

主要步骤：

1. 调用 `get_basic_information()`。
2. 调用 `update_shipment_quantity()`：
   - 如果同一个发货单号里同一个 SKU 重复出现，重算 `发货数量 = 箱数 * 单箱数量`。
3. 合并产品资料，得到单品毛重、材质、报关中文品名等字段。
4. 合并申报单价。
5. 计算 `总PCS`、`申报总价`、`是否含金刚石`。
6. 先按 `报关中文品名 + 是否含金刚石` 做一次汇总。
7. 遍历装箱组合：
   - 如果某组只有一行，直接计入总体积、总重量、总箱数。
   - 如果某组多行，视为共箱，按单品毛重占比分摊箱重和 CBM。
8. 根据分摊后的重量和体积，再按报关品名汇总。
9. 对拼箱品名按重量估算箱数：
   - 基本规则是 `floor(总重量 / 15)`。
   - 小于 1 时按 1 箱处理。
10. 对箱数差额做回补。
11. 计算高度、含运费申报单价、净重、总体积等字段。
12. 合并品名表，生成 `self.result_df`。

拼箱的特点：

- 不是简单按 SKU 汇总。
- 共箱情况下会重新分摊箱重和体积。
- 箱数可能被重新估算和回补。
- `result_df` 是后续所有输出 Sheet 和报关开票登记的基础。

## 5. 备货单模式业务逻辑

备货单模式入口类是 `stock_preparation.CustomsDeclaration`。

主流程在 `app/flows/stock_flow.py` 的 `StockFlow.generate_all_documents()`：

1. 调用 `generator._prepare_declaration_data()`。
2. 根据拼箱开关进入备货单拼箱或不拼箱逻辑。
3. 生成 `generator.result_df`。
4. 在所选输出目录下创建 `{发票号} 报关资料` 文件夹，渲染并保存 `{发票号} 报关资料.xlsx`。
5. 如果开启申报文件，在所选输出目录复用或复制申报模板，并向模板追加本票数据。
6. 设置 Excel 分页视图。
7. 生成 `{发票号} 报关资料.pdf` 和 `{发票号} 销售合同.pdf`。
8. 写入调用统计。

### 5.1 备货单基础信息

备货单公共前置逻辑在 `stock_preparation.py` 的 `get_basic_information()`。

关键数据来源：

- 备货单 `备货单详情`。
- 备货单 `轨迹信息（新版）`。
- 备货单 `装箱信息`。
- `产品信息单证专用.xlsx`。

备货单 SKU 会做标准化：

- `normalize_stock_sku()` 会去掉 SKU 前缀 `AE`。

备货单的店铺/品牌主要从产品信息表读取，并映射公司主体：

- 包含 `EZARC` 或 `TOLESA`：当前归到易达主体。
- 包含 `LIBRATON`：归到鲁米亚主体。
- 其他未匹配会报错，需要维护映射规则。

备货单发票号通常来自 `轨迹信息（新版）` 的 `查询单号`。

### 5.2 备货单生成结果

备货单也有拼箱和不拼箱两条线：

- 不拼箱：`stock_preparation.py` 的 `process_invoice()`。
- 拼箱：`stock_preparation.py` 的 `lcl_situation()`。

两条线最终都生成 `self.result_df`，后续渲染逻辑复用 `app/renderers` 里的 Sheet 渲染函数。

备货单与发货单一样，发票和合同 Sheet 仍保留在报关资料工作簿中；不再额外生成独立的“发票合同”工作簿及 PDF。

## 6. 运费逻辑

运费计算在 `app/services/freight_service.py`。

入口函数：

- `calculate_shipping_costs(transportation_method, logistics_provider, country_zh, currency, total_volume, total_weight)`

返回值：

- 运费金额。
- 分摊口径：`volume` 或 `weight`。

主要规则：

- `海卡`：按体积，部分国家和物流商有特殊单价。
- `海派`：按重量，默认最低计费重量 12kg。
- `空派`：按重量，默认最低计费重量 12kg。
- `卡航`：按重量。
- `美国海运`：按体积。

如果币别是 USD，运费用 `7.2` 汇率折算；其他币别按 `1.0`。

生成报关结果时，会根据返回的 `freight_basis` 决定运费分摊：

- `weight`：按 `总重量（kg）-箱子` 分摊。
- `volume`：按 `总体积（m³）-箱子` 分摊。

## 7. 申报单价逻辑

申报单价选择在 `app/services/declared_price_service.py`。

入口函数：

- `select_declared_price_columns()`

核心规则：

- 易达报关单价：优先找列名包含 `易达-亚马逊` 和 `申报单价` 的列。
- 匠晟报关单价：使用 `匠晟-亚马逊\n申报单价(USD)`。
- 备货单中易达且非 AWD：使用 `易达-欣耀\n申报单价（CNY）`。
- 其他情况默认使用 `申报单价`。

如果合并后某个 SKU 找不到申报单价，代码会抛错，提示去 `产品信息单证专用.xlsx` 维护。

## 8. 要素表逻辑

要素匹配在 `app/services/elements_service.py`。

入口函数：

- `find_data_group(elements_file_path, result_df, cache, sheet_name_transform=None)`

处理方式：

1. 根据 `result_df` 最后一列判断要读取的要素 Sheet，通常是店铺名。
2. 读取对应 Sheet 的第一列。
3. 用空行把要素内容拆成多个组。
4. 每组内匹配：
   - HS CODE。
   - 报关中文品名。
   - 是否含金刚石。
5. 匹配成功后，把 SKU 追加到要素组末尾。
6. 如果单位是 `套`，还会追加形如 `10套=20个` 的换算说明。

如果要素表缺少对应品名或 HS CODE，会抛出明确错误，提示去 `要素.xlsx` 对应 Sheet 维护。

## 9. 输出文件逻辑

生成报关资料时，流程层创建一个 `openpyxl.Workbook()`，再依次调用 `app/renderers`：

1. `render_detailed_list()`：明细单。
2. `ensure_default_sheet_removed()`：删除默认 Sheet。
3. `render_packing_list()`：装箱单。
4. `render_bill()`：发票。
5. `render_contract()`：合同。
6. `render_customs_declaration()`：报关单。
7. `render_elements_sheet()`：要素。
8. `finalize_workbook_page_setup()`：统一分页设置。
9. `wb.save()`：保存工作簿。

所有 Sheet 主要依赖两个对象：

- `generator.shipping_details`：票据级信息，如发票号、国家、收件人、公司主体、总箱数、总重量、总金额。
- `generator.result_df`：品名/SKU 汇总后的报关明细。

图片路径来自界面的 `pictures_path`，默认是网络共享目录下的 `pictures` 文件夹。

工作簿保存并设置分页视图后，通过 `app/core/pdf_export.py` 调用 WPS 或 Excel 的 `ExportAsFixedFormat`：

- 整本报关资料工作簿导出为 `{发票号} 报关资料.pdf`。
- “合同”Sheet 单独导出为 `{发票号} 销售合同.pdf`。

分页设置和 PDF 导出通过 `DispatchEx` 启动独立的 WPS/Excel 进程。处理完成后只关闭该独立进程，不会连接、关闭或触发保存用户已打开的其他工作簿。

明细单、装箱单、发票和合同里的发件人/收件人区域会按显式换行和预估自动换行计算行高。此规则用于避免 WPS/Excel 导出 PDF 时，长地址被下方表格覆盖或裁切。

## 10. 申报文件逻辑

如果界面勾选“是否生成申报文件”，生成报关资料后会继续执行：

1. `copy_declaration_template_file(所选输出目录)`。
2. `fill_declaration_template_file(declaration_template_path)`。

申报模板不放在 `{发票号} 报关资料` 文件夹内，而是保留在用户选择的输出目录，例如桌面。这样同一店铺的后续订单会复用同一个模板文件，在已有明细后继续追加，不会每票生成一份孤立申报文件。

模板按店铺/公司主体选择。

发货单模板映射大致为：

- `EZARC`：EA B2B 订单导入模板。
- `LIBRATON` / `PACMANN`：LT B2B 订单导入模板。
- `SIXPRAY`：六浪模板。
- `TOLESA`：匠晟模板。
- 未匹配：默认 YP B2B 订单导入模板。

申报文件填充时会使用 `self.result_df` 和 `self.shipping_details` 中的发票号、平台编码、平台名称、数量、价格等字段。

填充规则：

- 模板的第一个 Sheet 从第 4 行开始写入；已有数据时追加到最后一条有效明细之后。
- 发票号写入“订单编号”；同一个模板已存在该发票号时会报错，防止重复追加。
- 单位代码从模板的 `计量单位代码` Sheet 匹配，写入时保留三位字符串格式。
- 发货单使用 `发货数量`；备货单优先使用 `备货数量`，没有该列时回退到 `发货数量`。

## 11. 登记逻辑

登记只在界面勾选“是否登记”为“是”时执行。

发货单和备货单都会登记两类文件：

1. 出运明细。
2. 报关开票登记表。

这两类登记不是重复登记，而是写入两套不同台账。

### 11.1 出运明细登记

发货单入口：

- `invoice_data.py` 的 `register_shipping_invoice_number()`。

备货单入口：

- `stock_preparation.py` 的 `register_shipping_invoice_number()`。

写入目标：

- 界面“登记设置”里的 `出运明细` 文件。
- Sheet 固定为 `出运明细`。

发货单写入字段包括：

- 发货日期。
- 物流商单号。
- 站点。
- 物流中心编码。
- 国家。
- 运输方式。
- 货件单号。
- 店铺。
- SKU。
- 发货数量。
- 箱数。
- 总重量。
- 总体积。
- 关税。
- VAT/消费税。
- 货运费。
- 发货仓库。
- 物流商。

备货单字段类似，但发票号、站点、物流中心编码、出运方式等会结合仓库和轨迹信息重新映射。

写入前会检查发票号是否已经登记：

- 发货单使用 `check_invoice_number()`。
- 备货单使用 `check_invoice_number_in_sheet()`。

如果已登记，会抛出“发票号已登记”。

### 11.2 报关开票登记

发货单入口：

- `invoice_data.py` 的 `build_invoicing_register_df()` 构建数据。
- `write_invoicing_register()` 写入文件。

备货单入口：

- `stock_preparation.py` 的 `build_invoicing_register_df()` 构建数据。
- `write_invoicing_register()` 写入文件。

目标文件由 `app/core/models.py` 的 `select_register_filepath()` 根据 `company_info_index` 选择：

- `0`：易达报关明细。
- `1`：鲁米亚报关明细。
- `2`：素桐报关开票明细。
- `3`：六浪报关明细。
- `4`：匠晟报关明细。
- 未知值默认走素桐。

发货单中：

- 素桐写入 Sheet `2026`。
- 其他主体写入 Sheet `2026（内部开票）`。
- 报关渠道固定 `亚马逊`。
- 报关方式固定 `9710`。
- 成交方式固定 `CFR`。
- 供应商名称默认 `上海友谊卷笔刀有限公司`。

备货单中：

- AWD：报关渠道 `亚马逊`，报关方式 `9710`。
- 非 AWD：报关渠道 `欣耀报关`，报关方式 `一般贸易`。
- 成交方式 `CFR`。
- 供应商名称默认 `上海友谊卷笔刀有限公司`。

### 11.3 发货单报关开票登记的数据来源

发货单登记构建时有两个数据源：

- `register_combined_invoice()`：从发货单装箱信息和发货单详情合并出来的明细级数据。
- `self.result_df`：报关资料生成阶段得到的报关汇总数据。

素桐主体的报关开票登记以 `self.result_df` 为准。

原因：

- `self.result_df` 已经按报关品名、SKU、是否含金刚石等规则完成汇总。
- 如果直接按 SKU 合并明细级发货单数据，同一 SKU 在多个发货单中出现时会被重复放大。
- 当前逻辑会先把明细级字段按 SKU 折叠成一行，再和 `self.result_df` 合并。

典型例子：

- 同一票 `26YP065` 中，SKU `425202` 出现 5 次，SKU `425205` 出现 5 次。
- 出运明细登记需要保留 10 行，因为它记录实际发货明细。
- 报关开票登记只应保留 2 行，因为它记录报关汇总品名：
  - `425202` 剪刀。
  - `425205` 橡皮。

### 11.4 登记写入校验

登记写入后会做校验：

- 出运明细校验实际总行数是否等于预计总行数。
- 报关开票登记校验合同号/发票号是否能在目标列中找到。
- 如果目标文件被打开占用，写入前的 `ensure_file_writable()` 会报错。

## 12. Excel 读取缓存

`app/core/excel_io.py` 的 `ExcelCache` 用来缓存 `pd.read_excel()` 结果。

特点：

- 缓存 key 包含文件路径、Sheet 名和读取参数。
- 返回时使用 `copy()`，避免调用方修改缓存中的原始 DataFrame。
- 对 `dtype` 和 `usecols` 做了兼容处理，防止 pandas 因 dtype 指定了未读取列而报错。

发货单和备货单类中都通过 `_read_excel_cached()` 使用这个缓存。

## 13. 日志与调用统计

日志：

- 发货单日志文件：`invoice_data_log.txt`。
- 备货单日志文件：`stock_preparation_log.txt`。
- 路径通过 `app/core/logging.py` 计算。

调用统计：

- 通过 `app/core/metrics.py` 追加到共享盘的 `报关资料-调用统计.xlsx`。
- 记录字段包括记录时间、模块、动作、状态、耗时、发票号、源文件、错误信息。

主要统计动作：

- 生成报关资料。
- 登记数据。

## 14. 常见问题定位

### 14.1 填了两个单据路径

现象：

- 界面报“请只填写发货单或备货单路径，不要同时填写。”

原因：

- `process_data()` 要求发货单和备货单二选一。

处理：

- 清空其中一个路径。

### 14.2 发票号为空

发货单：

- 检查 `发货单详情` 的 `物流商单号`。

备货单：

- 检查 `轨迹信息（新版）` 的 `查询单号`。

### 14.3 SKU 找不到产品资料

现象：

- 报错提示某些 SKU 在产品信息单证专用表里找不到。

原因：

- 发货单/备货单里的 SKU 没有在 `产品信息单证专用.xlsx` 中维护。

处理：

- 在产品信息表对应 Sheet 添加 SKU。
- 备货单注意 `AE` 前缀会被去掉后再匹配。

### 14.4 申报单价为空

现象：

- 报错提示某些 SKU 申报单价为空。

原因：

- 根据公司主体选择到的报关单价 Sheet 或列中没有维护该 SKU 的申报单价。

处理：

- 到 `产品信息单证专用.xlsx` 对应报关单价 Sheet 维护。

### 14.5 要素表找不到品名

现象：

- 报错提示某 HS CODE、报关中文品名需要去要素表添加。

原因：

- `要素.xlsx` 对应店铺 Sheet 中没有匹配的要素组。

处理：

- 在对应 Sheet 添加包含品名和 HS CODE 的要素组。

### 14.6 登记提示发票号已登记

原因：

- 出运明细或报关开票登记表中已经存在当前发票号/合同号。

处理：

- 核对是否重复生成。
- 如果确实需要重登，需要先按业务规则处理原登记记录。

### 14.7 报关开票登记出现重复 SKU 行

判断方式：

- 出运明细保留明细行是正常的。
- 报关开票登记应以报关汇总行为准，通常同一个 SKU/报关品名只出现一行。

相关代码：

- 发货单：`invoice_data.py` 的 `build_invoicing_register_df()`。
- 关键数据源：`self.result_df`。

处理思路：

- 优先检查 `self.result_df` 是否已经重复。
- 如果 `self.result_df` 正确，但登记结果重复，检查明细级 `merged_df` 是否在按 SKU 合并时放大。
- 素桐主体应先把明细级字段按 SKU 折叠后再合并。

## 15. 测试

测试目录是 `tests/`。

当前测试覆盖重点：

- Excel 缓存返回副本。
- 公司登记文件选择。
- 运费计算。
- 要素匹配。
- 分页设置持久化。
- flow 层不会递归调用旧入口。
- 发货单和备货单部分登记规则。
- 发货单素桐报关开票登记同 SKU 多明细时只登记汇总行。

常用测试命令：

```bash
python -m unittest discover -s tests -p 'test_*.py'
```

发货单专项测试：

```bash
python -m unittest discover -s tests -p 'test_invoice_data.py'
```

## 16. 维护建议

新增业务规则时优先确认规则属于哪一层：

- 界面路径、开关、必填校验：改 `ui_pyqt.py`。
- 发货单业务计算：改 `invoice_data.py`。
- 备货单业务计算：改 `stock_preparation.py`。
- 输出 Sheet 样式或字段：改 `app/renderers`。
- 运费、申报单价、要素、登记公共规则：优先改 `app/services` 或 `app/core`。
- 整体生成顺序：改 `app/flows`。

修改登记逻辑时要特别区分：

- 出运明细：记录发货/备货明细，允许同 SKU 多行。
- 报关开票登记：通常记录报关汇总结果，不应被发货明细行数放大。

修改拼箱逻辑时要重点验证：

- 总箱数。
- 总重量。
- 总体积。
- 申报总价。
- 含运费申报单价。
- `self.result_df` 是否符合登记和输出需求。
