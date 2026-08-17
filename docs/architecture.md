# Architecture

This document describes the internal design of `crawl_treasury_bonds.py` for contributors and maintainers.

## Layers

| Layer | Responsibility | Key functions |
|-------|----------------|---------------|
| **Transport** | HTTP fetch with polite delay and encoding fix | `get_list_page`, `get_detail_page` |
| **Link extraction** | Parse list HTML, match `YYYY年第N号 国债业务公告` titles | `extract_announcements_from_list` |
| **Classification** | Decide announcement category from page text | `is_market_operation_announcement`, `is_savings_bond_announcement`, `is_reopening_announcement` |
| **Structured extraction** | Regex/numeric parsing into typed records | `extract_normal_bond_data`, `extract_savings_bond_data`, `extract_market_operation_data` |
| **Persistence** | Write styled multi-sheet Excel | `create_excel` |
| **Orchestration** | Pipeline control and console reporting | `extract_all_data`, `main` |

## Request lifecycle

```
main()
  └─ extract_all_data()
       ├─ loop pages: get_list_page → extract_announcements_from_list
       │    └─ dedup by (year, num), sort newest-first
       └─ for each announcement:
            get_detail_page → classify → dispatch extractor
            → append to normal_bonds / savings_bonds / market_operations
  └─ create_excel(normal, savings, market_ops, OUTPUT_FILE)
  └─ print statistics and previews
```

## Classification rules

- **Market-making** — page text contains `国债做市支持操作` or `做市支持操作`.
- **Savings** — contains `储蓄国债` and a tranche marker (`第一期`…`第四期`).
- **Normal / re-issue** — otherwise; re-issue flagged when text contains `续发行`.

## Extraction highlights

- **Tenor inference for re-issues** — when no explicit `期限X年` is found, the maturity is derived as `repayment_year − issue_year` extracted from `YYYY年…日偿还本金` and `YYYY年记账式`.
- **Savings splitting** — a single combined announcement is expanded with the pattern `第(一二三…)期期限(\d+)年，票面年利率为([\d.]+)%，最大发行额([\d.]+)亿元`; Chinese numerals are mapped to integers.
- **Merged-cell resilience** — market-making table rows inherit the previous row's operation direction when empty, and fall back to positional matching when the header match produces an invalid bond name.
- **Multi-form rate parsing** — tries `票面利率为X%`, `票面利率X%`, `折合年收益率X%`, `年收益率X%`, and `票面年利率X%` in priority order.

## Configuration contract

All tunables read from environment variables with constants as fallback:

| Variable | Fallback |
|----------|----------|
| `CGB_BASE_URL` | `https://zwgls.mof.gov.cn/ywgg/` |
| `CGB_USER_AGENT` | Chrome UA string |
| `CGB_DELAY_MIN` / `CGB_DELAY_MAX` | `0.3` / `0.5` |
| `CGB_OUTPUT_FILE` | `国债业务公告数据.xlsx` |

## Extension points

- **Incremental mode** — persist the max `(year, num)` seen and only fetch newer announcements.
- **Alternate sinks** — after `extract_all_data`, write to CSV/JSON or a database instead of (or in addition to) Excel.
- **Scheduling** — wrap `main()` in a cron / scheduler job for periodic refresh.
- **Validation hook** — reuse the checklist in `SKILL.md` §5 as an automated post-run assertion.

## Limitations

- The script depends on the current MOF HTML structure; selectors are text-pattern based and may need adjustment if the site changes.
- Output is best-effort parsing of semi-structured Chinese text; spot-check critical rows against the source.
