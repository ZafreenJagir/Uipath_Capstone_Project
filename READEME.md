# CAPSTONE PROJECT

## AIM
To automate a Speech-to-Text workflow using UiPath, where the user speaks a sentence, the text is captured from Notepad, and the data is stored in an Excel sheet automatically.

## PROCEDURE:
1. Create a Trigger to Start Recording
Use Keyboard Trigger or Monitor Events to start the process.
When you press the hotkey → workflow begins.
2. Open Notepad for Capturing Speech Output
Use Use Application/Browser
App: Notepad
Selector appears like:
"<wnd app='notepad.exe' ... />"
This Notepad will display the converted speech text.
3. Perform Speech-to-Text
Use Say / Speech Recognition or the package you installed.
Assign output to variable:
spokenText (String)
Earlier, the issue happened because you did not assign the recognized text.
This step fixes that.
4. Write the Captured Text into Notepad

Use Type Into

Target: Notepad

Value: spokenText

If spokenText is empty → Notepad stays empty → you get “No Text Captured” message.

5. Open Excel Automatically

Use Use Excel File

File path: "YourFile.xlsx"

Create file if not exists → ON

6. Read Existing Data (Optional)

Your Read Range has only:

File Path

Range

It is okay to leave Range blank → it reads the entire sheet.

Assign the output to:

dt (DataTable)

7. Append the New Speech Text into Excel

Use Write Cell OR Write Range:

✔ If you use Write Cell:

Cell: "A" + (dt.Rows.Count + 1).ToString

Value: spokenText

This ensures:

If row count = 1 → writes at A2

No overwritten data

8. Save and Close Excel

The Excel Scope automatically saves.

9. Show Success Message

Example:

MessageBox: "Text saved successfully!"
