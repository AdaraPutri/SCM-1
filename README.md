# New Claims List & Claims To Set Up VBA Guide

Two Outlook VBA trackers that turn claim-related emails into living Excel to-do lists.

- **New Claims**: extracts Claim Number, Date Received, Insured, and Adjuster from "New Task To Assign" emails. Appends new rows and skips duplicates by Claim Number.
- **Claims to Setup**: extracts unique email subjects (RE:/FW:/FWD: stripped) and Date Received. Appends new rows and skips duplicates by subject.

## 1. Move relevant emails into a dedicated folder

Do this **manually** rather than with an Outlook rule as it matches the phrase anywhere in the subject, not just at the start, so they'll also catch unrelated emails that happen to contain the same words.

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

## 5. Repeat the procedure with `ClaimsToSetup.bas`
Repeat steps 1–4 with these differences:

| Item | New Claims | Claims to Setup |
|------|------------|-----------------|
| Folder name | `New Claims` | `Claims to Set Up` |
| `.bas` file | `NewClaimFiles.bas` | `ClaimsToSetup.bas` |
| Excel file | `New_Claims_List.xlsx` | `Claims_To_Setup_Data.xlsx` |
| Update macro | `ExtractEmailTableDataToDesktop` | `ExtractUniqueSubjectsToDesktop` |
| Reset macro | `ResetClaimsSheet` | `ResetClaimsToSetupSheet` |
| Button names | Update New Claims List / Reset Claims Sheet | Update Claims to Set Up List / Reset Claims to Set Up List |

Everything else (manual folder move, VBA import, Quick Access Toolbar, weekly cleanup order, known errors) is identical.

## 6. Suggested Use/ Workflow

1. At the end of each week, archive/ move the emails out of the `New Claims` and `Claims to Set Up` folder and click **Reset Claims Sheet** and **Reset Claims To Set Up Sheet**.
2. Click **Update New Claims List** and/ or **Update Claims To Set Up List** whenever a new email has been added to respective folders.

**NOTE**: Always exit out of the spreadsheet before updating or resetting the list.

Below is how the generated spreadsheet looks like, note that the highlight colors and file number are to be added manually:
- New Claims List
![Sample Spreadsheet 1](Screenshots/SpreadsheetSample_NewClaims.png)
- Claims to Set Up List
![Sample Spreadsheet 2](Screenshots/SpreadsheetSample_ClaimsToSetup.png)



## Known errors

**Runtime error 1004 / "file is open in another application"**
A previous run didn't close Excel properly, leaving it running invisibly in the background with the file locked.
→ Task Manager (`Ctrl+Shift+Esc`) → **Details** tab → end any `EXCEL.EXE` processes → re-run the macro.

**Macros stop running after reopening Outlook**
Outlook's macro security resets each session by default.
→ **File → Options → Trust Center → Trust Center Settings → Macro Settings** → select "Notifications for all macros," then click **Enable Content** when prompted each time Outlook restarts.

**A claim doesn't show up after running the macro**
Usually means it's already in the sheet (duplicate-check by Claim Number is working as intended) — check existing rows before assuming it failed.