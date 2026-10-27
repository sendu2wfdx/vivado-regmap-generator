# Vivado Register Map Generator

一个以 JSON 为单一数据源的寄存器表生成工具，可同时输出：

- 便于评审和交付的 Excel 寄存器表；
- 适合放入仓库的 Markdown 文档；
- 驱动和应用可直接引用的 C 头文件。

当前仓库附带 ADS4229/Zynq 采集项目的示例定义。生成器本身不依赖 Vivado。

## 安装与生成

```bash
python -m pip install -r requirements.txt
python gen_regmap.py ads4229_register_map.json --out-dir .
```

生成结果为 `ads4229_register_map.xlsx`、`.md` 和 `.h`。

## RTL 一致性检查

`check_regmap_against_rtl.py` 是 ADS4229 示例工程专用检查器，可将 JSON 中的地址
与 RTL 寄存器定义进行比对：

```bash
python check_regmap_against_rtl.py \
  ads4229_register_map.json \
  ../zynq-ads4229-daq/rtl/ads4229_interface/ads4229_regs.v \
  ../zynq-ads4229-daq/rtl/ads4229_accumulator/sample_regs.v
```

`import_legacy_regmap.py` 用于导入旧版 Excel 数据，建议导入后人工复核字段语义。

## 许可

本仓库目前未附带开源许可证。公开可见不等于获得复制、修改或再分发授权；若要
供第三方复用，发布前应补充明确许可证。
