# Evaluating Investment Returns & Index Performance — Excel mini-project (margin return, IOC fills, equal/cap-weighted indices)


**Scenarios covered**

| Module | What it shows | Key output |
|---|---|---|
| Investment Rate of Return (with margin, dividends, commissions, interest) | Equity return on a margined trade | **21.31%** leveraged RoR |
| Average Trade Price (IOC, partial fills) | VWAP across price levels given order book sizes | Buy: **$140.8424** (139k shrs); Sell: **$140.4405** (37k shrs) |
| Equal-Weighted Index (price & total return) | Index level, divisor, component contributions | Price: **13.906%**, Total: **21.997%** |
| Capitalization-Weighted Index (price & total return) | Market-cap weights, divisor, index return | Price: **4.679%**, Total: **7.523%** |

---

## Case Description
You’ll see three independent but related analyst tasks often performed on trading/portfolio desks:
1) calculate equity investors’ realized return when leverage, fees, and cash flows matter;  
2) compute average execution price for IOC orders with partial fills from an order book;  
3) build both equal-weighted and cap-weighted index return calculations (price & total).

---

## Tasks
- Compute a **leveraged rate of return** including dividends, buy/sell commissions, and interest on margin.
- Derive **VWAP** for an IOC buy order and an IOC sell order given price levels and available sizes.
- Build **equal-weighted** and **cap-weighted** index worksheets that:
  - track **price** and **total** return,
  - use a **divisor** to anchor the index level,
  - surface which constituents drive return.

---

## Accounting/Analytics Steps
### 1) Investment Rate of Return (margin, dividends, costs)
Inputs (from `Investment Rate of Return.xlsx`, “Case 1”):
- Shares: 200,000  
- Buy price: $180.00  
- Initial margin: 50%  
- Sell price (1 year): $200.00  
- Interest on borrowed funds: 4%  
- Dividends per share: $3.00  
- Commission per share: $0.10 (both sides)

Step-by-step:
1. Total purchase cost = 200,000 × 180 = **$36,000,000**  
2. Initial equity (margin) = 36,000,000 × 50% = **$18,000,000**  
3. Borrowed amount = 36,000,000 − 18,000,000 = **$18,000,000**  
4. Buy commission = 200,000 × 0.10 = **$20,000**  
5. Sell proceeds = 200,000 × 200 = **$40,000,000**  
6. Sell commission = **$20,000**  
7. Dividends = 200,000 × 3 = **$600,000**  
8. Interest = 18,000,000 × 4% = **$720,000**  
9. **Net cash to investor at exit** = 40,000,000 − 20,000 − 18,000,000 (loan repay) + 600,000 − 720,000 = **$21,860,000**  
10. **Initial equity outlay** = 18,000,000 + 20,000 = **$18,020,000**  
11. **Profit** = 21,860,000 − 18,020,000 = **$3,840,000**  
12. **Leveraged RoR** = Profit / Initial equity = **21.3097%** (≈ **21.31%**)

### 2) Average Trade Price (IOC partial fills)
**IOC Buy (order 140,000 shrs)**  
Fills across offers:  
- 40,000 @ 140.70 ; 12,000 @ 140.80 ; 75,000 @ 140.90 ; 12,000 @ 141.00 → **139,000 executed**, **1,000 canceled**  
VWAP = (∑ price × qty) / (∑ qty) = **$140.8424**

**IOC Sell (order 40,000 shrs, limit $140.35)**  
Fills across bids:  
- 15,000 @ 140.50 ; 22,000 @ 140.40 → **37,000 executed**, **3,000 canceled**  
VWAP = **$140.4405**

### 3) Equal-Weighted Index (price & total return)
From `Equal-weighted index.xlsx`, “Case 3”:
- **Divisor**: **5,292,000**  
- **Price Return**: **13.9063%**  
- **Total Return**: **21.9965%**  
- Notes: Equal weighting amplifies small-cap influence; frequent rebalancing needed to maintain weights.

### 4) Capitalization-Weighted Index (price & total return)
From `Capitalization-weighted index.xlsx`, “Case 4”:
- **Divisor**: **2,477,450.08**  
- **Price Return**: **4.6788%**  
- **Total Return**: **7.5233%**  
- Notes: Larger caps dominate index behavior; price swings in top names materially move the index.

---

## Trial Balance / Data Summary (table + totals/checks)
**A) Margin Trade — cash flow summary**

| Item | Amount |
|---|---:|
| Purchase cost | $36,000,000 |
| Initial equity (incl. buy commission) | $18,020,000 |
| Borrowed amount | $18,000,000 |
| Sell proceeds (net of sell commission) | $39,980,000 |
| Loan repayment | $(18,000,000)$ |
| Dividends received | $600,000 |
| Interest paid | $(720,000)$ |
| **Ending cash to investor** | **$21,860,000** |
| **Profit** | **$3,840,000** |
| **RoR (Profit / Initial equity)** | **21.31%** |

_Checks:_ Ending cash − Initial equity = Profit ✔︎

**B) IOC Orders — execution summary**

| Order | Executed Qty | Canceled Qty | VWAP |
|---|---:|---:|---:|
| Buy 140,000 shrs (IOC) | 139,000 | 1,000 | $140.8424 |
| Sell 40,000 shrs (IOC, Lmt 140.35) | 37,000 | 3,000 | $140.4405 |

_Checks:_ VWAP within min/max executed prices ✔︎; Executed ≤ Order qty ✔︎

**C) Indices — summary**

| Index | Divisor | Price Return | Total Return |
|---|---:|---:|---:|
| Equal-Weighted | 5,292,000 | 13.906% | 21.997% |
| Cap-Weighted | 2,477,450.08 | 4.679% | 7.523% |

_Checks:_ Weights sum to 100% by construction; Divisor anchors index level ✔︎

---

## Financial Statements / Results
- **Investor outcome (leveraged trade):** Profit **$3.84m**, **21.31%** equity return over holding period.  
- **Execution quality:** IOC Buy VWAP **$140.8424** on 139k shrs; IOC Sell VWAP **$140.4405** on 37k shrs.  
- **Index performance:** Equal-weighted outperforms cap-weighted in this sample (higher total return), consistent with greater small-cap influence.

---

## Mapping / Logic (line flow & formulas)
- **Leveraged RoR pipeline**
  - Equity_outlay = Initial_margin × (Shares × BuyPrice) + BuyCommission  
  - Exit_cash = (Shares × SellPrice − SellCommission) − Borrowed + Dividends − Interest  
  - Profit = Exit_cash − Equity_outlay  
  - RoR = Profit / Equity_outlay
- **IOC VWAP**
  - For each eligible price level within limit, take `fill_qty = min(remaining, available_size)`  
  - VWAP = `SUMPRODUCT(price_level, fill_qty) / SUM(fill_qty)`  
  - Remaining cancels on IOC.
- **Equal-Weighted Index**
  - Each name weight = `1 / N` at (re)balance.  
  - Index Price Return = `Σ (equal_weight × stock_price_return)`  
  - Total Return adds dividends (reinvested).  
  - Divisor chosen so base level = 100.
- **Cap-Weighted Index**
  - MarketCapᵢ = Priceᵢ × SharesOutstandingᵢ ; Weightᵢ = MarketCapᵢ / Σ MarketCap  
  - Index return = `Σ (weightᵢ × constituent returnᵢ)`; Total Return adds dividends.  
  - Divisor ensures continuity through corporate actions.

---

## How I Built It (tools & example formulas)
- **Tools:** Microsoft Excel (formulas, structured references), basic data validation.
- **Key Excel techniques**
  - Cash flows & RoR:  
    - `=Shares*BuyPrice`  
    - `=Borrowed*InterestRate`  
    - `=DividendsPerShare*Shares`  
    - `=Profit/InitialEquity`
  - VWAP from order book:  
    - `=SUMPRODUCT(PriceRange, QtyRange) / SUM(QtyRange)`
  - Equal-Weighted Index:  
    - `=AVERAGE(ReturnsRange)` for price return; add dividends for total return  
    - Store **Divisor** and compute **Index Level** with: `=BaseLevel*(1+PriceReturn)`
  - Cap-Weighted Index:  
    - `=MarketCap_i / SUM(MarketCaps)` for weights  
    - `=SUMPRODUCT(Weights, Returns)`  
    - Maintain a **Divisor** to keep index level continuity.

---

## What I Learned
- How leverage, carry costs, and fees transform a seemingly simple trade into a materially different **equity return**.
- How IOC order constraints change realized **VWAP** and leave measurable unexecuted remainders.
- Why **weighting schemes** (equal vs. cap-weighted) lead to different risk/return profiles and require different maintenance (rebalancing vs. drift).

---
