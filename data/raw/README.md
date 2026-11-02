# Raw inputs

These files are preserved unchanged from the development snapshot. No cells, dates, column names, missing-value markers or numeric units were rewritten for publication.

## Treasury CSV

- File: `treasury_yields_2026.csv`.
- Reported source: [U.S. Treasury Daily Par Yield Curve Rates, 2026](https://home.treasury.gov/resource-center/data-chart-center/interest-rates/TextView?field_tdr_date_value=2026&type=daily_treasury_yield_curve).
- Direct download URL recorded in the supplied file's metadata: https://home.treasury.gov/resource-center/data-chart-center/interest-rates/daily-treasury-rates.csv/2026/all?field_tdr_date_value=2026&type=daily_treasury_yield_curve&page&_format=csv .
- The original manual download timestamp/timezone is unknown. Preservation/inspection was recorded at 2026-10-29T21:57:58-04:00 (America/New_York), not as proof of download time. The metadata URL was not an independently observed HTTP retrieval.
- 208 records, 15 columns; observed range January 2-October 29, 2026, descending by date. Source dates are month/day/four-digit year.
- Units: annual par yields in percent; for example, 4.70 means 4.70%, not 0.0470%. These are benchmark-curve yields, not individual bond prices. See [Treasury methodology](https://home.treasury.gov/policy-issues/financing-the-government/interest-rate-statistics/treasury-yield-curve-methodology).
- Selected analysis columns: Date, 2 Yr, 5 Yr, 10 Yr and 30 Yr.
- SHA-256: `04ff6da859cf8d7a98c239b67d462ab213f47e7fca657ded34319c383c5b41c3`.

## Synthetic portfolio XLSX

- File: `synthetic_portfolio.xlsx`, worksheet `Holdings`, three fictional holdings. It contains no real account or market-price data.
- Every position has USD 1,000,000 principal. Annual coupons are 4%, 4.5% and 5%; maturities are October 27, 2028/2031/2036. Issue/as-of dates are October 27, 2026.
- Original Excel dates and cell storage are preserved; the reader exposes date/datetime objects.
- SHA-256: `11b481dfe9c1b8acc1ca1be6c9cf0c433a678093600cde80b5bf997ae553040a`.

| Column | Meaning |
|---|---|
| bond_id | Identifier for the fictional holding |
| face_value_usd | Full-position principal in USD, not market value |
| coupon_rate_pct | Annual coupon in percent; 4.5 means 4.5% |
| issue_date | Stated issue date |
| maturity_date | Date of final coupon and principal repayment |
| as_of_date | Holdings reference date |
| currency | USD currency code |
| coupon_payments_per_year | Coupon frequency; 2 means semiannual |

This public provenance note is adapted from the development records to remove machine-specific paths. It is not an unedited transcript of those records.
