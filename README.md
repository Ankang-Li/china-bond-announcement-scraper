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
