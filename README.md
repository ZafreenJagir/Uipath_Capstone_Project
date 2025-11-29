# CAPSTONE PROJECT
##### NAME : ZAFREEN J 
##### REGISTER NUMBER : 212223040252



## AIM:

To automate a Speech-to-Text workflow using UiPath, where the user speaks a sentence, the text is captured from Notepad, and the data is stored in an Excel sheet automatically.

## PROCEDURE:

#### 1. Create a Trigger to Start Recording

Use Keyboard Trigger or Monitor Events to start the process.

When you press the hotkey → workflow begins.

#### 2. Open Notepad for Capturing Speech Output

Use Use Application/Browser

App: Notepad

Selector appears like:
"<wnd app='notepad.exe' ... />"

This Notepad will display the converted speech text.

#### 3. Perform Speech-to-Text

Use Say / Speech Recognition or the package you installed.

Assign output to variable:

spokenText (String)


⚠️ Earlier, the issue happened because you did not assign the recognized text.
This step fixes that.

#### 4. Write the Captured Text into Notepad

Use Type Into

Target: Notepad

Value: spokenText

If spokenText is empty → Notepad stays empty → you get “No Text Captured” message.

#### 5. Open Excel Automatically

Use Use Excel File

File path: "YourFile.xlsx"

Create file if not exists → ON

#### 6. Read Existing Data (Optional)

Your Read Range has only:

File Path

Range

It is okay to leave Range blank → it reads the entire sheet.

Assign the output to:

dt (DataTable)

#### 7. Append the New Speech Text into Excel

Use Write Cell OR Write Range:

✔ If you use Write Cell:

Cell: "A" + (dt.Rows.Count + 1).ToString

Value: spokenText

This ensures:

If row count = 1 → writes at A2

No overwritten data

#### 8. Save and Close Excel

The Excel Scope automatically saves.

#### 9. Show Success Message

MessageBox: "Text saved successfully!"

## DESIGN WORKFLOW:

<img width="1003" height="950" alt="image" src="https://github.com/user-attachments/assets/87fcf11f-446f-4c7d-93fa-327eb4624f77" />

<img width="991" height="927" alt="image" src="https://github.com/user-attachments/assets/591ec7a6-e107-4e54-8189-ed76b25e300c" />


<img width="1004" height="934" alt="image" src="https://github.com/user-attachments/assets/3e0cc100-0d8b-4170-920d-70526ddef1ce" />


<img width="1004" height="937" alt="image" src="https://github.com/user-attachments/assets/272ae358-d0da-4706-8e6d-6eac6aa58be9" />


<img width="994" height="922" alt="image" src="https://github.com/user-attachments/assets/8a144477-1848-4abd-8ae2-6618b71adf1c" />


<img width="1009" height="936" alt="image" src="https://github.com/user-attachments/assets/050e067f-2bcf-45a4-8b36-80b7241526c6" />

<img width="990" height="938" alt="image" src="https://github.com/user-attachments/assets/ba9481a2-202f-486a-84ec-0b81fff28fb8" />

<img width="991" height="687" alt="image" src="https://github.com/user-attachments/assets/94cda9ba-64f6-4882-b8ec-9255a1361170" />

<img width="992" height="684" alt="image" src="https://github.com/user-attachments/assets/b3801333-8cc8-4080-ac23-c749168e802d" />

## OUTPUT:

<img width="1914" height="956" alt="image" src="https://github.com/user-attachments/assets/2ea1994e-36ce-40d8-81b3-812a8c6af92c" />

<img width="1897" height="703" alt="image" src="https://github.com/user-attachments/assets/9f2f283d-45b7-4b26-b9d5-3856461cba50" />

<img width="863" height="440" alt="image" src="https://github.com/user-attachments/assets/98685cbf-8043-4ad2-9120-78326a3f9928" />

<img width="879" height="425" alt="image" src="https://github.com/user-attachments/assets/80517c7b-8269-42e5-ac3b-6609d6d20df6" />

<img width="920" height="403" alt="image" src="https://github.com/user-attachments/assets/1fdfce95-72ca-414e-b441-1f6f886b5952" />

## RESULT:

Automating a Speech-to-Text workflow using UiPath, where the user speaks a sentence, the text is captured from Notepad, and the data is stored in an Excel sheet automatically is implemented Successfully.

