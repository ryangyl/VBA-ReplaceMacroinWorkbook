Attribute VB_Name = "Module1"
Option Explicit
Sub replacemacro()
Dim folders As Variant
Dim folder As Variant
Dim folderpath As String
Dim filename As Variant
Dim file As Variant
Dim wb As Workbook
Dim macrocode As String
Dim i As Integer
i = 0
macrocode = _
    "Private Sub Workbook_BeforeClose(Cancel As Boolean)" & vbCrLf & _
    "    Dim newWorkbook As Workbook" & vbCrLf & _
    "    On Error Resume Next" & vbCrLf & _
    "    Set newWorkbook = Workbooks.Open(""\\siwdsntv002\SG_PSC_SG1_PL_08_Control_WHse\Daily Tank Reading\Solvent Tracking Macro.xlsm"")" & vbCrLf & _
    "    If Not newWorkbook Is Nothing Then" & vbCrLf & _
    "        MsgBox ""Solvent Tracking Macro opened.""" & vbCrLf & _
    "Thisworkbook.Close Savechanges := True" & vbCrLf & _
    "    Else" & vbCrLf & _
    "        MsgBox ""Error: Could not open Solvent Tracking Macro.""" & vbCrLf & _
    "        Cancel = True" & vbCrLf & _
    "        Exit Sub" & vbCrLf & _
    "    End If" & vbCrLf & _
    "End Sub"

folders = Array("\Oct 24\", "\Nov 24\", "\Dec 24\")

For Each folder In folders
folderpath = "\\siwdsntv002\SG_PSC_SG1_PL_08_Control_WHse\Daily Tank Reading\Tanker reading year 2024" & folder
filename = Dir(folderpath)

    Do While filename <> ""
    i = i + 1
    file = folderpath & filename
    Set wb = Workbooks.Open(file)
    If wb Is Nothing Then
        MsgBox "Invalid File"
        Exit Sub
        End If
'Set ws = Sheets("Sheet1")
    Debug.Print "Processing file: " & file
'Debug no need click yes no
    With wb.VBProject.VBComponents("ThisWorkbook").CodeModule
            .DeleteLines 1, .CountOfLines ' Remove any existing code in ThisWorkbook
            .AddFromString macrocode      ' Add the Workbook_BeforeClose macro
        End With
MsgBox filename & " Done"
wb.Close Savechanges:=True
filename = Dir
Loop
Next folder
MsgBox "Done!"
End Sub






