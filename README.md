# china-bond-announcement-scraper
A data scraping tool for Chinese Government Bond (CGB) business announcements and issuance data.

### Chinese Government Bond (CGB) Business Announcements Scraper
Quickly scrape all government bond business announcements from the official website of the Ministry of Finance (MOF) of the PRC and organize them into structured Excel sheets.

### Quick Start
1. Prerequisites
Ensure you have Python 3 and the required dependencies installed:

```bash
pip install requests beautifulsoup4 openpyxl
```

2. Run the Scraper

```bash
cd scripts
python3 crawl_treasury_bonds.py
```

3. Check the Output
Once the script completes, a file named 国债业务公告数据.xlsx (CGB_Business_Announcements.xlsx) will be generated in the current directory, containing 3 sheets:

| Sheet | Description | Data Volume |
|-------|------|--------|
| Government bond issuance data | Standard Issuance + Re-issuance announcements | ~130 records |
| Savings Government Bond Data | Savings bonds data (split by tranche/phase) | ~4 records |
| Market-making operational data | Market-making support operation announcements | ~16 records |

(2025/01/08 - 2026/06/24)
___

### Features
✅ Auto-Pagination: Automatically traverses all list pages without manual page turning.

✅ Multi-Type Support: Supports coupon-bearing bonds (附息国债), discount bonds (贴现国债), special government bonds (特别国债), and more.

✅ Smart Tenor Inference: Automatically deduces the tenor/maturity of re-issued bonds by calculating the time until the repayment date.

✅ Savings Bond Splitting: Automatically splits combined announcements of multi-tranche savings bonds into individual data rows.

✅ Market-Making Data Extraction: Intelligently extracts tables from market-making operations (with merged cell support).

✅ Formatted Excel Output: Generates stylized Excel files featuring custom headers, frozen top rows, and auto-fitted column widths.

✅ Polite Scraping: Implements a request delay of 0.3 to 0.5 seconds to respect the server.

___

### Data Sample
* CGB Issuance Data(国债发行数据)

| Announcement Title| Type | Planned Issuance| Tenor | Coupon/Yield |
|---------|------|---------|------|------|
| No. 105 of 2026 (2026年第105号) | Coupon-bearing | 85.0B RMB | 30Y | 2.23% |
| No. 105 of 2026 (2026年第104号) | Discount | 20.0B RMB | 91D | 0.83% |
| No. 105 of 2026 (2026年第102号) | Coupon-bearing (Re-issue) | 90.0B RMB | 10Y | 1.72% |

* Market-Making Operation Data (做市操作数据)

| Target Bond | Tenor | Operation Amount | Yield |
|---------|------|--------|--------|
| 2026 Book-entry Coupon-bearing Treasury Bond (Issue 4) (2026年记账式附息（四期）国债) | 3Y | 270M RMB | 1.33% |
| 2026 Book-entry Coupon-bearing Treasury Bond (Issue 4) (2026年记账式附息（六期）国债) | 2Y | 270M RMB | 1.31% |

### Data Source
Ministry of Finance of the PRC - Government Bond Business Announcements:
🔗 https://zwgls.mof.gov.cn/ywgg/

### Detailed Documentation
For more comprehensive instructions, please refer to SKILL.md, which covers:

* Full scraping workflow

* Troubleshooting & FAQ

* Data validation checklist

* Field definition guide

___

### Disclaimer
⚠️ Research Only: This data is for research and educational purposes only and does not constitute investment advice.

⚠️ Scraping Ethics: Please respect the website's scraping rules and avoid sending high-frequency requests.

⚠️ Data Accuracy: For any critical or institutional data, please verify with the original text on the official MOF website.

___

# 国债业务公告数据抓取工具

快速从财政部官网抓取所有国债业务公告数据，整理成结构化 Excel 表格。

## 快速开始

### 1. 环境准备

确保已安装 Python 3 和必要的依赖：

```bash
pip install requests beautifulsoup4 openpyxl
```

### 2. 运行爬虫

```bash
cd scripts
python3 crawl_treasury_bonds.py
```

### 3. 查看结果

脚本运行完成后，会在当前目录生成 `国债业务公告数据.xlsx`，包含 3 个 sheet：

| Sheet | 说明 | 数据量 |
|-------|------|--------|
| 国债发行数据 | 普通发行 + 续发行公告 | ~130 条 |
| 储蓄国债数据 | 储蓄国债多期拆分 | ~4 条 |
| 做市操作数据 | 做市支持操作公告 | ~16 条 |

## 功能特性

- ✅ 自动遍历所有列表页，无需手动翻页
- ✅ 支持附息国债、贴现国债、特别国债等多种类型
- ✅ 自动推断续发行公告的期限（从偿还日期计算）
- ✅ 自动拆分储蓄国债多期合并公告
- ✅ 智能提取做市操作表格数据（支持合并单元格）
- ✅ 生成格式化的 Excel（表头样式、冻结首行、自适应列宽）
- ✅ 礼貌爬取（请求间隔 0.3-0.5 秒）

## 数据示例

### 国债发行数据

| 公告标题 | 类型 | 计划发行 | 期限 | 利率 |
|---------|------|---------|------|------|
| 2026年第105号 | 附息国债 | 850亿 | 30年 | 2.23% |
| 2026年第104号 | 贴现国债 | 200亿 | 91天 | 0.83% |
| 2026年第102号 | 附息国债（续发行） | 900亿 | 10年 | 1.72% |

### 做市操作数据

| 操作券种 | 期限 | 操作额 | 收益率 |
|---------|------|--------|--------|
| 2026年记账式附息（四期）国债 | 3年 | 2.7亿 | 1.33% |
| 2026年记账式附息（六期）国债 | 2年 | 2.7亿 | 1.31% |

## 数据来源

财政部官网 - 国债业务公告：https://zwgls.mof.gov.cn/ywgg/

## 详细文档

更多详细说明请查看 [SKILL.md](./SKILL.md)，包含：
- 完整工作流程
- 常见问题与解决方案
- 数据验证检查清单
- 字段定义说明

## 注意事项

⚠️ 数据仅供研究参考，不构成投资建议
⚠️ 请遵守网站爬取规则，不要请求过于频繁
⚠️ 重要数据请以财政部官网原文为准
