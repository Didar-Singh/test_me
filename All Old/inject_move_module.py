"""
Injects a new VBA module "MoveFilesToFolder" into Split Into Folders.xlsm.
Run this script once; afterwards open the .xlsm and run the macro from the VBA editor
or assign it to a button on any sheet.

Requirements: pywin32  (pip install pywin32)
"""

import os
import sys
import tempfile
import win32com.client as win32

# ---------------------------------------------------------------------------
# VBA source for the new module
# ---------------------------------------------------------------------------
VBA_CODE = r'''
'===========================================================================
' Module : MoveFilesToFolder
' Purpose: Reads a sheet where
'            Column A = filename  (e.g. ABC.xlsx)
'            Column B = subfolder (e.g. folder1)
'          Moves every listed file from <BaseFolder>\<subfolder>\<filename>
'          into <DestFolder>, then removes now-empty source sub-folders.
'
' Usage  : Open the workbook, go to the sheet that holds the list,
'          run  MoveFilesToFolder  from the Macros dialog or a button.
'===========================================================================
Option Explicit

Sub MoveFilesToFolder()

    Dim ws          As Worksheet
    Dim lastRow     As Long
    Dim i           As Long
    Dim fileName    As String
    Dim subFolder   As String
    Dim srcFile     As String
    Dim dstFile     As String
    Dim baseFolder  As String
    Dim destFolder  As String
    Dim fso         As Object
    Dim movedCount  As Long
    Dim skippedLog  As String

    ' -----------------------------------------------------------------------
    ' 1.  Ask the user for the base folder that CONTAINS the sub-folders
    ' -----------------------------------------------------------------------
    baseFolder = BrowseForFolder("Select the BASE folder that contains the sub-folders")
    If baseFolder = "" Then
        MsgBox "Cancelled – no base folder selected.", vbExclamation, "Move Files"
        Exit Sub
    End If

    ' -----------------------------------------------------------------------
    ' 2.  Ask the user for the DESTINATION (single) folder
    ' -----------------------------------------------------------------------
    destFolder = BrowseForFolder("Select the DESTINATION folder to move files into")
    If destFolder = "" Then
        MsgBox "Cancelled – no destination folder selected.", vbExclamation, "Move Files"
        Exit Sub
    End If

    ' Ensure trailing backslash
    If Right(baseFolder, 1) <> "\" Then baseFolder = baseFolder & "\"
    If Right(destFolder, 1) <> "\" Then destFolder = destFolder & "\"

    ' -----------------------------------------------------------------------
    ' 3.  Locate the active sheet (or the sheet named "FileList" if it exists)
    ' -----------------------------------------------------------------------
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("FileList")
    On Error GoTo 0
    If ws Is Nothing Then Set ws = ActiveSheet

    ' Find last used row in Column A
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "No data found in the list (expected headers in row 1).", vbExclamation, "Move Files"
        Exit Sub
    End If

    ' -----------------------------------------------------------------------
    ' 4.  Create FSO and destination folder if it doesn't exist
    ' -----------------------------------------------------------------------
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(destFolder) Then fso.CreateFolder destFolder

    movedCount = 0
    skippedLog = ""

    ' -----------------------------------------------------------------------
    ' 5.  Loop through the list
    ' -----------------------------------------------------------------------
    For i = 2 To lastRow                          ' row 1 is the header

        fileName  = Trim(CStr(ws.Cells(i, 1).Value))
        subFolder = Trim(CStr(ws.Cells(i, 2).Value))

        If fileName = "" Then GoTo NextRow        ' skip blank rows

        srcFile = baseFolder & subFolder & "\" & fileName
        dstFile = destFolder & fileName

        If Not fso.FileExists(srcFile) Then
            skippedLog = skippedLog & "Row " & i & ": NOT FOUND – " & srcFile & vbCrLf
            GoTo NextRow
        End If

        ' If a file with the same name already exists in dest, rename with counter
        Dim counter As Long
        counter = 1
        Dim baseName As String, ext As String, candidate As String
        baseName = fso.GetBaseName(fileName)
        ext      = fso.GetExtensionName(fileName)
        candidate = dstFile
        Do While fso.FileExists(candidate)
            candidate = destFolder & baseName & "_" & counter & "." & ext
            counter = counter + 1
        Loop
        dstFile = candidate

        fso.MoveFile srcFile, dstFile
        movedCount = movedCount + 1

NextRow:
    Next i

    ' -----------------------------------------------------------------------
    ' 6.  Remove now-empty source sub-folders
    ' -----------------------------------------------------------------------
    Dim removedFolders As Long
    removedFolders = 0
    Dim folderPath As String
    Dim subFoldersSeen As New Collection

    ' Collect unique sub-folder names from the list
    On Error Resume Next
    For i = 2 To lastRow
        subFolder = Trim(CStr(ws.Cells(i, 2).Value))
        If subFolder <> "" Then subFoldersSeen.Add subFolder, subFolder
    Next i
    On Error GoTo 0

    Dim sf As Variant
    For Each sf In subFoldersSeen
        folderPath = baseFolder & CStr(sf)
        If fso.FolderExists(folderPath) Then
            ' Only delete if the folder is empty
            If fso.GetFolder(folderPath).Files.Count = 0 And _
               fso.GetFolder(folderPath).SubFolders.Count = 0 Then
                fso.DeleteFolder folderPath, True
                removedFolders = removedFolders + 1
            End If
        End If
    Next sf

    ' -----------------------------------------------------------------------
    ' 7.  Summary
    ' -----------------------------------------------------------------------
    Dim msg As String
    msg = "Done!" & vbCrLf & vbCrLf & _
          "Files moved    : " & movedCount & vbCrLf & _
          "Folders removed: " & removedFolders

    If skippedLog <> "" Then
        msg = msg & vbCrLf & vbCrLf & _
              "Skipped (file not found):" & vbCrLf & skippedLog
    End If

    MsgBox msg, vbInformation, "Move Files – Complete"

End Sub

' ---------------------------------------------------------------------------
' Helper: Browse-for-folder dialog
' ---------------------------------------------------------------------------
Private Function BrowseForFolder(prompt As String) As String
    Dim shell  As Object
    Dim folder As Object
    Set shell = CreateObject("Shell.Application")
    Set folder = shell.BrowseForFolder(0, prompt, 0, "")
    If Not folder Is Nothing Then
        BrowseForFolder = folder.Self.Path
    Else
        BrowseForFolder = ""
    End If
End Function
'''

# ---------------------------------------------------------------------------
# Inject the module via COM
# ---------------------------------------------------------------------------
XLSM_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                         "Split Into Folders.xlsm")
MODULE_NAME = "MoveFilesToFolder"


def inject():
    if not os.path.exists(XLSM_PATH):
        print(f"ERROR: File not found: {XLSM_PATH}")
        sys.exit(1)

    # Write VBA to a temp .bas file
    tmp = tempfile.NamedTemporaryFile(suffix=".bas", delete=False,
                                      mode="w", encoding="utf-8")
    tmp.write(f"Attribute VB_Name = \"{MODULE_NAME}\"\r\n")
    tmp.write(VBA_CODE)
    tmp.close()

    xl = None
    wb = None
    try:
        xl = win32.Dispatch("Excel.Application")
        xl.Visible = False
        xl.DisplayAlerts = False

        wb = xl.Workbooks.Open(XLSM_PATH)
        vbp = wb.VBProject

        # Remove existing module with the same name (idempotent)
        for comp in list(vbp.VBComponents):
            if comp.Name == MODULE_NAME:
                vbp.VBComponents.Remove(comp)
                print(f"  Removed existing module '{MODULE_NAME}'")
                break

        # Import fresh module
        vbp.VBComponents.Import(tmp.name)
        print(f"  Imported module '{MODULE_NAME}'")

        wb.Save()
        print(f"  Saved: {XLSM_PATH}")

    except Exception as exc:
        print(f"ERROR: {exc}")
        print("\nTip: Make sure 'Trust access to the VBA project object model' is enabled.")
        print("     File > Options > Trust Center > Trust Center Settings")
        print("     > Macro Settings > check 'Trust access to the VBA project object model'")
        sys.exit(1)

    finally:
        if wb:
            try:
                wb.Close(SaveChanges=False)
            except Exception:
                pass
        if xl:
            try:
                xl.Quit()
            except Exception:
                pass
        os.unlink(tmp.name)

    print("\nAll done. Open Split Into Folders.xlsm and run 'MoveFilesToFolder'.")


if __name__ == "__main__":
    inject()
