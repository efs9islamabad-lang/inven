Google Apps Script Deployment Guide

Follow these steps to set up, configure, and deploy your Google Apps Script web application.



Step 1: Create a New Google Sheet

1. Go to [Google Sheets](https://sheets.google.com) and create a new blank spreadsheet.
2. Title your spreadsheet (e.g., Inventory Management System).



Step 2: Add the Backend Code (Code.gs)

1. In the Google Sheet top menu, click Extensions > Apps Script.
2. By default, a file named Code.gs will open in the editor.
3. Clear any existing placeholder code in Code.gs.
4. Copy and paste your JavaScript backend code into the editor.
5. Click the Save icon (or press Ctrl + S / Cmd + S).


Step 3: Add the Frontend HTML (Index.html)

1. In the left panel of the Apps Script editor, click the + (Add a file) icon.
2. Select HTML from the dropdown menu.
3. Name the file Index (the editor will automatically append .html) and press Enter.
4. Clear any default boilerplate code in Index.html.
5. Copy and paste your complete HTML/CSS/JS frontend code into the file.
6. Click the Save icon (or press Ctrl + S / Cmd + S).

Step 4: Initial Run (Sheet Initialization)

1. At the top of the Apps Script editor, locate the Select function dropdown.
2. Select initializeSheets (or doGet).
3. Click Run.
4. If prompted with an Authorization Required dialog:
 Click Review Permissions.
 Choose your Google Account.
 Click Advanced > Go to (Untitled project) (unsafe).
 Click Allow to grant necessary permissions.


Step 5: Deploy as a Web Application

1. In the upperright corner of the editor, click Deploy > New deployment.
2. Click the Gear icon (Select type) next to "Select type" and select Web app.
3. Fill in the deployment details:
 Description: Enter a title (e.g., Inventory System Web App v1.0).
 Execute as: Set to Me (youremail@gmail.com).
 Who has access: Select Anyone (or Anyone within your organization if using a Workspace account).

4. Click Deploy.
5. Once deployment is complete, copy the provided Web app URL.
6. Open the link in a new browser tab to access your live application.

Modifying Your Code After Deployment

Whenever you make changes to Code.gs or Index.html after deploying:

1. Click Deploy > Manage deployments.
2. Click the Edit (pencil icon) on your active deployment.
3. Change Version to New version.
4. Click Deploy to publish your updates.

## Important product and safety notes
- This project is a Google Apps Script web app backed by Google Sheets; it is not an Excel VBA workbook. Buyers need a Google account and must authorize the script.
- Before using live business data, make a copy of the spreadsheet and test add, adjust, and remove flows with sample records.
- Deploy only to the intended audience. “Anyone” access may expose the app to people who obtain its URL; review your data-sharing and authorization requirements before selecting it.
- After editing code, deploy a new version and test the deployed URL—not only the Apps Script preview.
- The app currently supports stock transactions and reports; it is not a full accounting, purchase-order, barcode, or multi-user permission system. Do not advertise features that are not implemented.
