# Fundamental Booster – Excel Practice Workbook

A hands-on Excel workbook for practising core functions: logical tests, text, rounding, date, lookup, and conditional aggregation. It has three sheets, each with 20 records.

**File:** `Fundamental_booster.xlsx`

---

## Sheets at a Glance

| Sheet | Records | Topic | Main functions |
|---|---|---|---|
| Students Grade | 20 | Marks, percentage, grading | `IF`, `AND`, `ROUND`, `UPPER`, `LOWER`, `COUNTIF`, `AVERAGEIFS`, `VLOOKUP` |
| Sales Data | 20 | Regional discounts and eligibility | `IF`, `OR`, `SUMIFS`, `TODAY`, `ROUND`, `CEILING`, `FLOOR`, `VLOOKUP` |
| Employee Data | 20 | Salary rounding | `UPPER`, `LOWER`, `ROUND`, `CEILING`, `FLOOR` |

---

## 1. Students Grade

**Input columns (A–F):** Student ID, Name, Math, Science, English, Enrollment Date

**Calculated columns:**

| Column | Formula (row 2) | Purpose |
|---|---|---|
| Total (G) | `=C2+D2+E2` | Sum of three subjects |
| Percentage (H) | `=ROUND((G2/300)*100,2)` | Out of 300 marks |
| Grade (I) | `=IF(H2>80,"A",IF(H2>60,"B",IF(H2>40,"C")))` | A > 80, B > 60, C > 40 |
| Scored >80 in Math and Science (J) | `=IF(AND(C2>80,D2>80),"Yes","No")` | Both subjects above 80 |
| Average (K) | `=ROUND((C2+D2+E2)/3,2)` | Mean of the three subjects |
| UPPERCASE / LOWERCASE (L, M) | `=UPPER(B2)` / `=LOWER(B2)` | Name case conversion |

**Exercises below the table:**

- **COUNTIFS:** Count students with Math > 50 → `=COUNTIF(C2:C21,">50")` (result: 19)
- **AVERAGEIFS:** Average percentage of students scoring above 60 in all three subjects → `=ROUND(AVERAGEIFS(H2:H21,C2:C21,">60",D2:D21,">60",E2:E21,">60"),2)` (result: 79.33)
- **VLOOKUP:** Fetch a student name from Student ID using the separate list at the bottom.

---

## 2. Sales Data

**Input columns (A–F):** Sales ID, Product, Region, Salesperson, Amount, Date

**Calculated columns:**

| Column | Formula (row 2) | Purpose |
|---|---|---|
| Discount (G) | Nested `IF` on Region | South 10%, West 15%, East 2.5%, North 5% |
| Amount After Discount (H) | `=E2*(1-G2)` | Net amount |
| Eligible / Not eligible (I) | `=IF(OR(E2>10000,B2="Monitor"),"Eligible","Not Eligible")` | Amount > 10,000 or product is a Monitor |
| UPPERCASE / LOWERCASE (J, K) | `=UPPER(D2)` / `=LOWER(D2)` | Salesperson name case |
| Days from date (L) | `=TODAY()-F2` | Age of the sale in days |
| Round / Ceiling / Floor (M–O) | `=ROUND(H2,-3)`, `=CEILING(H2,1000)`, `=FLOOR(H2,1000)` | Net amount rounded to the nearest thousand |

**Exercises below the table:**

- **SUMIFS:** Total sales for a region → `=SUMIFS(E2:E21,C2:C21,"South")` (result: 73,188)
- **VLOOKUP:** Retrieve the amount for a product from the lookup list.

---

## 3. Employee Data

**Input columns (A–E):** Employee ID, Name, Department, Salary, Joining Date

**Calculated columns:**

| Column | Formula (row 2) | Purpose |
|---|---|---|
| UPPERCASE / LOWERCASE (F, G) | `=UPPER(B2)` / `=LOWER(B2)` | Name case |
| ROUND (H) | `=ROUND(D2,-3)` | Nearest 1,000 |
| Ceiling (I) | `=CEILING(D2,1000)` | Round up to next 1,000 |
| Floor (J) | `=FLOOR(D2,1000)` | Round down to previous 1,000 |

Departments: Finance, HR, IT, Marketing.

---

## Function Reference

| Function | What it does |
|---|---|
| `IF` / nested `IF` | Returns a value based on a condition |
| `AND` / `OR` | Combines multiple conditions |
| `COUNTIF` | Counts cells meeting one condition |
| `SUMIFS` | Sums values meeting conditions |
| `AVERAGEIFS` | Averages values meeting multiple conditions |
| `VLOOKUP` | Looks up a value in the first column and returns a column from the same row |
| `UPPER` / `LOWER` | Converts text case |
| `ROUND` | Rounds to nearest (`-3` = nearest thousand) |
| `CEILING` / `FLOOR` | Rounds up / down to a multiple |
| `TODAY` | Current date |

---

## Things to Know

- **Dates:** Date columns are Excel serial numbers. Format the cells as Date if they show as plain numbers.
- **Days from date:** `TODAY()` recalculates every time the file opens, so these values change daily.
- **Grade formula:** A percentage of 40 or below returns `FALSE`, because the last `IF` has no else value. Exactly 80 and 60 fall into the lower grade (B and C) since the tests use `>`.
- **VLOOKUP ranges:** The lookup ranges are relative (for example `A1:K21`, `A2:K22`, ...), so they shift down each row. This works for the current data but is fragile. Use a fixed range such as `$A$2:$B$21` for reliability.
- **Discount column:** Discounts are stored as text (`"10%"`). Excel converts them automatically in the multiplication, but a numeric value (`0.1`) is safer.

---

## How to Use

1. Open `Fundamental_booster.xlsx` in Excel, LibreOffice, or Google Sheets.
2. Click any formula cell to see how it is built.
3. Change the input data (marks, amounts, salaries) and watch the calculated columns update.
4. Try the exercises under each table with your own conditions.
