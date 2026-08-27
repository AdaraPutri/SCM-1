# New Claims List VBA Guide

Extracts Claim Number, Date Received, Insured, and Adjuster from "New Task To Assign" emails into a local Excel tracker. Appends new rows on every run and skips duplicates by Claim Number.

## 1. Move relevant emails into a dedicated folder

Do this **manually** rather than with an Outlook rule — rules match the phrase anywhere in the subject, not just at the start, so they'll also catch unrelated emails that happen to contain the same words.

- Create a subfolder under your Inbox (e.g. `New Claims`).

![Outlook folder pane showing the New Claims subfolder under Inbox](Screenshots/1.png)

- As "New Task To Assign" emails come in, drag them into that folder.

![Dragging an email into the New Claims folder](Screenshots/2.png)

## 2. Load the VBA script

- In Outlook, press `Alt` + `F11` to open the VBA editor.
- Press `Enable Macros` if a pop up message appears.
- In the Project Explorer (left pane), right-click your project → **Import File...** → select [`NewClaimFiles.bas`](NewClaimFiles.bas) that is downloaded to your device.

![VBA editor Project Explorer with the Import File menu open](Screenshots/4.png)

## 3. Match the script to your setup

Open the `NewClaimFiles` module and update two lines near the top of `ExtractEmailTableDataToDesktop` and near the bottom `ResetClaimsSheet` as you'd like or leave it be:

```vb
targetFolderName = "New Claims"        ' must exactly match the folder name from Step 1
```

```vb
filePath = ... & "\New_Claims_List.xlsx"   ' change the filename here if you want something different
```

![The targetFolderName and filePath lines in the code, with what to edit circled](Screenshots/5.png)
![The targetFolderName and filePath lines in the code, with what to edit circled](Screenshots/6.png)
![The targetFolderName and filePath lines in the code, with what to edit circled](Screenshots/7.png)

- Type `CTRL` + `s` to save any changes made

## 4. Add buttons to the ribbon

- Go back to Outlook
- **File → Options → Quick Access Toolbar**

![Navigating to Quick Access Toolbar](Screenshots/8.png)
![Navigating to Quick Access Toolbar](Screenshots/9.png)
![Navigating to Quick Access Toolbar](Screenshots/10.png)

- Set "Choose commands from" to **Macros**

![Quick Access Toolbar setup screen with the Macros dropdown selected](Screenshots/11.png)

- Add `Project1.ExtractEmailTableDataToDesktop` → click **Modify** → rename to **Update New Claims List** (pick a distinct icon)
- Add `Project1.ResetClaimsSheet` → **Modify** → rename to **Reset Claims Sheet** (pick a distinct icon)

![Quick Access Toolbar showing both buttons](Screenshots/12.png)

- Press OK
- The buttons should show up on the top of your screen:

![Ribbon with new buttons](Screenshots/13.png)

## 5. Weekly cleanup

1. Once claims are resolved, delete or archive their emails out of the `New Claims` folder.
2. Click **Update New Claims List** — wipes the spreadsheet back to just the header row.
3. Click **Reset Claims Sheet** — repopulates from whatever's currently in the folder.

## Known errors

**Runtime error 1004 / "file is open in another application"**
A previous run didn't close Excel properly, leaving it running invisibly in the background with the file locked.
→ Task Manager (`Ctrl+Shift+Esc`) → **Details** tab → end any `EXCEL.EXE` processes → re-run the macro.

**Macros stop running after reopening Outlook**
Outlook's macro security resets each session by default.
→ **File → Options → Trust Center → Trust Center Settings → Macro Settings** → select "Notifications for all macros," then click **Enable Content** when prompted each time Outlook restarts.

**A claim doesn't show up after running the macro**
Usually means it's already in the sheet (duplicate-check by Claim Number is working as intended) — check existing rows before assuming it failed.