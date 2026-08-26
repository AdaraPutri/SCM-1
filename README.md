# SCM-1

Keeping track of programmed automations I use at work. 

## CAT Claims Workflow
During CAT cases, handle these 4 types of claim submissions accordingly:

**1. New Claim File Assignments**
- See [New Claim File Assignments — Setup Guide](NewClaims-README.md) for full instructions
- Module generates a local Excel spreadsheet, appending new claims and skipping duplicates on each run

**2. ECS Calls**
- Set Outlook rule to migrate relevant emails to correct folder
- Download ECS list from Claimspace
- Run Data-Driven Email Automation (TBD)

**3. Emails for Claim Setup**
- Set Outlook rule to migrate relevant emails to correct folder
- Open Outlook VBA editor: `ALT` + `F11`
- Run [ClaimsToSetup VBA module](ClaimsToSetup.bas) on Outlook VBA editor
  - Module currently generates Excel spreadsheet. Refactoring is in progress to link to Sharepoint and remove duplicates.

**4. Updates from Prev (non-CAT tasks)**
- TBD
