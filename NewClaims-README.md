# New Claims List VBA Guide

Extracts Claim Number, Date Received, Insured, and Adjuster from "New Task To Assign" emails into a local Excel tracker. Appends new rows on every run and skips duplicates by Claim Number.

## 1. Move relevant emails into a dedicated folder

Do this **manually** rather than with an Outlook rule — rules match the phrase anywhere in the subject, not just at the start, so they'll also catch unrelated emails that happen to contain the same words.

- Create a subfolder under your Inbox (e.g. `New Claims`).

![Outlook folder pane showing the New Claims subfolder under Inbox](screenshots/01-folder-pane.png)

- As "New Task To Assign" emails come in, drag them into that folder.

![Dragging an email into the New Claims folder](screenshots/02-drag-email.png)

## 2. Load the VBA script

- In Outlook, press `Alt` + `F11` to open the VBA editor.
- In the Project Explorer (left pane), right-click your project → **Import File...** → select [`NewClaimFiles.bas`](NewClaimFiles.bas).

![VBA editor Project Explorer with the Import File menu open](screenshots/03-import-file.png)

## 3. Match the script to your setup

Open the `NewClaimFiles` module and update two lines near the top of `ExtractEmailTableDataToDesktop`:

```vb
targetFolderName = "New Claims"        ' must exactly match the folder name from Step 1
```

```vb
filePath = ... & "\New_Claims_List.xlsx"   ' change the filename here if you want something different
```

![The targetFolderName and filePath lines in the code, with what to edit circled](screenshots/04-edit-config-lines.png)

## 4. Add buttons to the ribbon

- **File → Options → Quick Access Toolbar**
- Set "Choose commands from" to **Macros**

![Quick Access Toolbar setup screen with the Macros dropdown selected](screenshots/05-qat-macros-dropdown.png)

- Add `NewClaimFiles.ExtractEmailTableDataToDesktop` → click **Modify** → rename to **Refresh New Claims List**
- Add `NewClaimFiles.ResetClaimsSheet` → **Modify** → rename to **Reset New Claims List** (pick a distinct icon — this one is destructive)

![Final Quick Access Toolbar showing both buttons](screenshots/06-final-toolbar.png)

## 5. Weekly cleanup

1. Once claims are resolved, delete or archive their emails out of the `New Claims` folder.
2. Click **Reset New Claims List** — wipes the spreadsheet back to just the header row.
3. Click **Refresh New Claims List** — repopulates from whatever's currently in the folder.

## Known errors

**Runtime error 1004 / "file is open in another application"**
A previous run didn't close Excel properly, leaving it running invisibly in the background with the file locked.
→ Task Manager (`Ctrl+Shift+Esc`) → **Details** tab → end any `EXCEL.EXE` processes → re-run the macro.

**Macros stop running after reopening Outlook**
Outlook's macro security resets each session by default.
→ **File → Options → Trust Center → Trust Center Settings → Macro Settings** → select "Notifications for all macros," then click **Enable Content** when prompted each time Outlook restarts.

**A claim doesn't show up after running the macro**
Usually means it's already in the sheet (duplicate-check by Claim Number is working as intended) — check existing rows before assuming it failed.
