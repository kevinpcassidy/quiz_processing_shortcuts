# quiz_processing_shortcuts
This code can be added as a google sheets script to create two menu items: 1) Fill Average Formulas, which will fill a column with the average of the two highest of the preceding 4 columns, and 2) Check all Scores for Updates, which will loop through another tab and update scores by column and student name. 

Directions for installation:
🧩 Step 1 — Open the Script Editor
In your Google Sheet, click Extensions → Apps Script.
A new tab will open showing the Google Apps Script editor.

📝 Step 2 — Add the Code
Delete anything in the editor.
Paste in the full script (see https://github.com/kevinpcassidy/quiz_processing_shortcuts/blob/main/quiz_processing_shortcuts.gs )
Click the 💾 Save icon and name the project (e.g., “Quiz Processing Shortcuts”).

▶️ Step 3 — Run It
Click Run ▶️ once.

The first time, you’ll need to authorize permissions:
Click “Review permissions” → choose your account → “Allow.”
It may have the pop up saying that the script can view/edit/delete google files. This script has no code that will delete any google files. Also, it can only view/edit the google sheet you are running it in, and only runs when you choose the new menu bar shortcuts.

After it runs once, you can close the script editor.

Step 4 - Refresh
Refresh your Google Sheet. Once it fully loads, you should have a new menu bar with the two functions. 

