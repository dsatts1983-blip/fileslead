# WebsitesDoIt Lead Finder — v0.5.0

This Windows test release is designed for the full Companies House basic-company CSV, including the supplied 2.81 GB expanded dataset.

## Find companies

1. Import the Companies House ZIP or CSV from the blue button. ZIP files are read directly and copied into the app's local data folder. Replacing a prior import keeps a timestamped backup.
2. Enter a town, postcode, company-name fragment, or a category. The first search scans the imported file, counts every match, and learns the SIC categories from the data. It shows progress and can be stopped safely.
3. After the first scan, the category field suggests real categories as you type their description or SIC code. Search results are shown 100 at a time, with totals and working Previous / Next buttons across every match.
4. Saved searches remain separate. The app stores compact row references rather than copying all matching company records into each search. If you replace the imported file, the old file is retained as a backup so earlier searches still point to their original data.
5. Check rows on the current page, add candidate websites, and run verification. The search saves checked rows and website entries as you work.

Each search reads the imported file so the total covers every matching company; on a multi-gigabyte CSV this can take time. Progress appears in the status bar, and **Stop search** cancels without saving partial results. The first completed search builds the SIC suggestion list; later searches reuse it while still scanning the imported file for an accurate total.

## Verified leads and CSV export

Verified leads are saved separately for each search. Choose a saved search on the **Verified leads** tab and export the four requested columns: `Company`, `Email`, `Website`, and `Telephone`.

Website checks read public pages from the website address you provide and only save a lead when the registered company name is found. Google searches open in your browser for manual review; the app does not harvest Google results automatically.

## Local data and disk space

Imported company data and saved searches are kept on the PC under `%LOCALAPPDATA%\WebsitesDoIt\LeadFinder`. Keep enough free space for the expanded company CSV (about 2.81 GB for the supplied ZIP), the original ZIP if retained in Downloads, and any timestamped backups made when replacing data.

## Rebuild

Run `Build-Release.ps1` on Windows to bundle `LeadFinder.ps1` into `WebsitesDoIt-Lead-Finder.exe`.
