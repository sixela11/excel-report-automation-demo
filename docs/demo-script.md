# 60-Second Demo Script

Setup: open the `sample_data/input` folder, a terminal in the project folder, and Excel closed.

| Time | Show | Say |
|---|---|---|
| 0:00–0:10 | Input folder with 7 files | "Here are seven months of sales exports. Merging them by hand each month means hours of copy-paste, and mistakes slip through." |
| 0:10–0:15 | Open `sales_2026-03.xlsx` and point at the customer ID column | "Typical problem: Excel has stripped the leading zeros, so 000104 became 104." |
| 0:15–0:25 | Run `python consolidate.py --input sample_data/input --output sample_output` | "One command processes every file." |
| 0:25–0:35 | Terminal statistics | "On this sample data: 243 rows read, 5 invalid rows excluded, 3 duplicates removed and one incomplete file rejected." |
| 0:35–0:45 | Report: **Summary** then **Consolidated** sheets | "The result is a formatted report. Revenue by month and product, and clean data with every ID intact. Each row shows which file and row it came from." |
| 0:45–0:55 | **Duplicates** and **Exceptions** sheets | "Nothing is silently dropped. Every duplicate and every error is listed with the exact file, row and reason, so your team reviews exceptions instead of hunting for them." |
| 0:55–1:00 | Back to the input folder | "Original files are never changed. If your team rebuilds a report like this every month, I can automate it for a fixed price." |

Note: always say "sample data" or "demo". Never present these figures as client results.

Recording tips: 1080p screen capture (free: Windows Snipping Tool screen recording or OBS), terminal font 16pt+, hide personal folders and notifications.
