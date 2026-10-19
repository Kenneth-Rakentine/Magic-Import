# Opera XML Grabber (Gmail → zip)

The companion to Magic Import. If OPERA's Reports Scheduler emails the Manager Flash and Journal reports to a
Gmail address, this Google Apps Script collects one month of them into a single zip, ready for the bulk import.

- Reads every email on one Gmail label (default `Automated XML Reports - Night Audit`).
- Uses the date in the **file name**, not the email date, because the morning email carries the previous
  day's business date.
- When a day was re-run, keeps the file with the higher job number, so a manual re-run beats the scheduled one.
- Lists every day that is missing a manager or journal file, so you know what to re-run in OPERA before
  importing. For the current month it checks through yesterday.

## Privacy

The reports are financial records, so v4 never puts them anywhere someone else could reach:

- **The web page downloads the zip straight to your browser.** Nothing is saved to, or shared from,
  Google Drive.
- **Only you can use the web app**, plus anyone you list in `ALLOWED_USERS`. Visitors Google can't identify are
  refused, so a forwarded or leaked link shows "You do not have access to this tool."
- **Editor runs** (`downloadThisMonth` and the others) save the zip to your Drive with private sharing; the
  logged link only works for you.
- **Upgrading from an older version?** Older versions shared the zips in the `Opera XML Exports` Drive folder
  by link. Run `lockDownOldZips` once to make them all private, or delete that folder.

## Setup

1. Go to **script.google.com** → **New project**, and name it `Opera XML Grabber`.
2. Replace the contents of `Code.gs` with this folder's `Code.gs`.
3. Click **+** next to Files → **HTML**, name it exactly `Index`, and paste in `Index.html`. Save both.
4. In `Code.gs`, check `LABEL` matches your Gmail label exactly. Run `listMyLabels` to print your labels.
   The first run asks for permission: **Advanced → Go to Opera XML Grabber (unsafe) → Allow**. That warning
   appears for any script you write yourself.

## Use it from the editor

Pick `downloadThisMonth`, `downloadLastMonth` or `downloadPickedMonth` (set `PICK_YEAR` and `PICK_MONTH`)
and click **Run**. The log shows the gap report and a private download link.

## Use it as a web page

1. **Deploy → New deployment** → gear → **Web app**.
2. **Execute as: Me.**
3. **Who has access:** *Only myself*, or *Anyone within [your organization]* if coworkers on the same
   Google Workspace domain will use it. In that case add their addresses to `ALLOWED_USERS`.
   Don't choose *Anyone*.
4. **Deploy** and bookmark the Web app URL.

After editing the code, use **Deploy → Manage deployments → pencil → Version: New version → Deploy**, so the
URL stays the same.

Pick a month, wait up to a minute, and click the green **Download** button. Unzip into the folder you'll
point the Magic Import bulk import at.

## Troubleshooting

- `Label not found`: run `listMyLabels` and copy the label exactly into `LABEL`.
- `No reports found`: run `debugSample` to see the newest attachment names and whether they match
  `manager20260811_12345.xml` / `journal20260811_12345.xml`.
