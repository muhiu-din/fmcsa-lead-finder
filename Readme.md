# FMCSA Active Interstate Carrier Lead Finder 🚚

A Python tool that scans a range of **USDOT numbers** against the official free **FMCSA QCMobile API** and exports qualified trucking carriers into an Excel spreadsheet for cold-calling and lead generation.

The script finds carriers that meet these conditions:

- ✅ `allowedToOperate == "Y"` (Active operating authority)
- ✅ Carrier operation classification contains **Interstate**
- ✅ Exported into an Excel `.xlsx` file ready for sales outreach

---

## Why Use This Approach?

Instead of scraping the FMCSA SAFER website, this project uses the official **FMCSA QCMobile API**.

### Advantages:

✅ No HTML scraping  
✅ Less likely to be blocked or rate-limited  
✅ Uses structured JSON responses  
✅ Directly verifies active operating status  
✅ Targets USDOT numbers (the current authority identifier)

---

## Important: USDOT vs MC Numbers

FMCSA retired **MC/MX numbers on October 1, 2025**.

New carrier authorities are now issued under the **USDOT number**, meaning:

- Old MC/MX scanning misses newer carriers
- USDOT scanning finds the newest and most valuable leads
- `--mode dot` is recommended for current lead generation

Legacy MC/MX scanning is still available for older carriers.

---

# Features

- 🔎 Scan thousands of USDOT numbers automatically
- 🚛 Filter only active interstate carriers
- 📊 Export results to Excel
- 💾 Save progress automatically
- 🔐 Uses official FMCSA API authentication
- ⏱ Configurable API request delay
- 🛡 Handles API errors and rate limits

---

# Requirements

- Python 3.9+
- FMCSA QCMobile API WebKey

Install dependencies:

```bash
pip install requests openpyxl

# Zip File
- It is extension fopr safer scrapping add to chrome and run.