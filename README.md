#From Clarity to Excel output file:
Select scroll and Mobile
Download all CSV files
Then add to Sharepoint
Change all files in data folder
Do this for every month in the beginning of the month(first 5 business days)

--> please be aware that quality check on missing months is now for ALL months starting from January.
I will plan to update a UQI that shows the  first month of every URL. Then it would be less Warnings on quality check and we are not going to try to download all missing months.

## How to run

Open Git Bash in the project folder and run:

```bash
source .venv/Scripts/activate
python scripts/clarity_csv_to_excel.py