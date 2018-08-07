Sub Insert_SeqNo()

i = InputBox("Major number", , " ")

If IsNumeric(i) Then
    For x = 1 To Selection.Cells.Count
        Selection.Cells(x) = Str(i) + "." + Str(x) + ": " + Selection.Cells(x)
    Next
End If
End Sub

Sub Paste_values_transpose()
'
' Paste_transpose Macro
'

'
    ActiveCell.Offset.Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=True
End Sub


Sub InsertMissing()
'
' InsertMissing Macro
'
' Keyboard Shortcut: Ctrl+m
'
    Dim iRow As Integer
    iRow = activecell.Row
        
    Cells(iRow, 5).Value = Cells(iRow, 2).Value
        
    Range(Cells(iRow, 5), Cells(iRow, 6)).Font.Color = vbRed
    
    Cells(iRow, 5).Interior.ThemeColor = xlThemeColorAccent6
    Cells(iRow, 5).Interior.TintAndShade = 0.8
    Cells(iRow + 1, 5).Select
    
End Sub

Sub auto_open()
Application.OnKey "{F1}", "InsertMissing"
Application.OnKey "^q", "clearStyles" ' Ctrl + q
Application.OnKey "^+q", "GetFileOwner" ' Ctrl + Shift + q
End Sub

Sub clearStyles()
    Dim sy As Style
    On Error Resume Next
    For Each sy In ActiveWorkbook.Styles
    sy.Delete
    Next
	MsgBox "Styles cleared."
End Sub

Sub GetFileOwner()
    Dim secUtil As Object
    Dim secDesc As Object
    Dim sFileName, File_Shortname As String
    Dim fileDir As String
     
     sFileName = Application.GetOpenFilename(, , "Choose a file:")
     File_Shortname = Dir(sFileName)
    fileDir = Left(sFileName, InStr(1, sFileName, File_Shortname) - 1)
    Set secUtil = CreateObject("ADsSecurityUtility")
    Set secDesc = secUtil.GetSecurityDescriptor(fileDir & File_Shortname, 1, 1)
         MsgBox secDesc.owner
End Sub

Sub Filter_By_Tab()
	Dim sCriteria As String
	sCriteria = Selection.Value
	Sheets("Ownership").Select
	Range("E1").Select
	If ActiveSheet.FilterMode Then ActiveSheet.ShowAllData
	Sheets("Ownership").ListObjects("Ownership").Range.AutoFilter Field:=3, Criteria1:=sCriteria
End Sub

Sub MDD_Filter()
'
' Macro1 Macro
' filter for CAN ID

' Keyboard Shortcut: Ctrl+Shift+X
'
On Error Resume Next
    Application.ScreenUpdating = False
    
    Dim sCriteria As String
        
'    sCriteria = Trim(ActiveSheet.Cells(ActiveCell.Row, 9).Value)
    sCriteria = Trim(ActiveCell.Value)
    ActiveSheet.Previous.Select
    If ActiveSheet.FilterMode Then ActiveSheet.ShowAllData
    ActiveSheet.Range("A1").Select
     ActiveSheet.ListObjects("MDD").Range.AutoFilter Field:=ActiveSheet.ListObjects("MDD").ListColumns(ActiveCell.Text).Index, _
     Criteria1:=Trim(sCriteria)
    
    Application.ScreenUpdating = True
End Sub

Sub Highlight_Differences()
Dim sLeft, sRight As Variant
Dim lTextLength, i As Long

Selection.Font.ColorIndex = xlAutomatic
sRight = Selection.Value
sLeft = Selection.Offset(0, -1).Value
lTextLength = Len(sRight)

For i = 1 To lTextLength
    If Mid(sRight, i, 1) <> Mid(sLeft, i, 1) Then _
        ActiveCell.Characters(Start:=i, Length:=1).Font.Color = -16776961
Next i
End Sub

Sub CreateIndex()
' Links cell A1 in all sheets to Index page
Dim i As Long

With ActiveWorkbook
If .Sheets(1).Name <> "Index" Then
    .Sheets.Add Before:=Sheets(1)
    .Sheets(1).Name = "Index"
End If

For i = 2 To .Sheets.Count
    .Sheets("Index").Cells(i, 1).Hyperlinks.Add Anchor:=.Sheets("Index").Cells(i, 1), Address:="", SubAddress:= _
        .Sheets(i).Name & "!A1", TextToDisplay:=.Sheets(i).Name

Next i
If MsgBox("Link to Index page added to all sheets." + Chr(10) + Chr(10) + "Do you want to create link to Index in all sheets?", vbYesNo + vbQuestion) = vbYes Then
    For i = 2 To .Sheets.Count
            .Sheets(i).Cells(1, 1).Hyperlinks.Add Anchor:=.Sheets(i).Cells(1, 1), Address:="", SubAddress:= _
        "Index!A" & CStr(i) ', TextToDisplay:=Selection.Value
    Next i
    MsgBox "Link to Index sheet created in all sheets", vbInformation
End If

End With
End Sub

Sub Get_Sheets_n_Headers()
Dim i, lRow, lActiveCol As Long
Application.ScreenUpdating = False

With ActiveWorkbook
If .Sheets(1).Name <> "Index" Then
    .Sheets.Add Before:=Sheets(1)
    .Sheets(1).Name = "Index"
Else
    .Sheets(1).Range("A:B").Clear
End If
.Sheets(1).Range("A1") = "Tab"
.Sheets(1).Range("B1") = "Attribute"
lRow = 2
For i = IIf(.Sheets(2).Name = "Submission_header", 3, 2) To .Sheets.Count
    If .Sheets(i).Visible = True Then
        .Sheets("Index").Cells(lRow, 1).Hyperlinks.Add Anchor:=.Sheets("Index").Cells(lRow, 1), Address:="", SubAddress:= _
            .Sheets(i).Name & "!A1", TextToDisplay:=.Sheets(i).Name
    
        lActiveCol = .Sheets(i).Range("A3").End(xlToRight).Column
        .Sheets(i).Range(.Sheets(i).Cells(3, 1), .Sheets(i).Cells(3, lActiveCol)).Copy
        .Sheets("Index").Cells(lRow, 2).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:= _
            False, Transpose:=True
        lRow = .Sheets(1).UsedRange.Rows.Count + 1
    End If

Next i
.Sheets(1).Columns("A:B").AutoFit

' Links cell A1 in all sheets to Index page
If MsgBox("Link to Index page added to all sheets." + Chr(10) + Chr(10) + "Do you want to create link to Index in all sheets?", vbYesNo + vbQuestion) = vbYes Then
    For i = 2 To .Sheets.Count
            .Sheets(i).Cells(1, 1).Hyperlinks.Add Anchor:=.Sheets(i).Cells(1, 1), Address:="", SubAddress:= _
        "Index!A" & CStr(i) ', TextToDisplay:=Selection.Value
    Next i
    MsgBox "Link to Index sheet created in all sheets", vbInformation
End If

End With

Application.ScreenUpdating = True
End Sub

Sub SplitFillAndLookup()  ' split_fill_lookup
Dim sValues() As String
Dim sVal As Variant
Dim lActiveRows, lCol, i, j, lRow, k As Long

lActiveRows = ActiveWorkbook.Sheets(2).UsedRange.Rows.Count
lCol = ActiveCell.Column
For i = 2 To lActiveRows
    sValues = Split(ActiveWorkbook.Sheets(2).Cells(i, lCol).Value, Chr(10))
    On Error Resume Next
    For Each sVal In sValues
        j = Left(sVal, 1)
        ActiveWorkbook.Sheets(2).Cells(i, lCol + j).Value = ActiveWorkbook.Sheets(2).Cells(i, lCol + j).Value & Chr(10) & sVal
        lRow = ActiveWorkbook.Sheets(3).Range("A1:A250").Find(sVal).Row
         If Err.Number = 0 Then
            For k = 3 To 8
                If IsEmpty(ActiveWorkbook.Sheets(2).Cells(i, lCol + k + 3).Value) Then _
                    ActiveWorkbook.Sheets(2).Cells(i, lCol + k + 3).Value = ActiveWorkbook.Sheets(3).Cells(lRow, k).Value
            Next k
        Else
            lRow = 0
            Err.Clear
        End If
    Next sVal
Next i
MsgBox "Split & Fill Done", vbExclamation, "Split & Fill"

End Sub

Sub FormatDefn()
Dim lRow, lTargetRow As Long

For lRow = 6 To 2284
    If Not IsEmpty(Cells(lRow, 1)) Then
        lTargetRow = lRow
    Else
        If Not IsEmpty(Cells(lRow, 2)) Then
            Cells(lTargetRow, 2).Value = Cells(lTargetRow, 2).Value & Chr(10) & Cells(lRow, 2).Value
            Cells(lRow, 2).Value = ""
        End If

        If Not IsEmpty(Cells(lRow, 3)) Then
            Cells(lTargetRow, 3).Value = Cells(lTargetRow, 3).Value & Chr(10) & Cells(lRow, 3).Value
            Cells(lRow, 3).Value = ""
        End If
'        ActiveSheet.Cells(lRow, 1).EntireRow.Delete
    End If
Next lRow

lTargetRow = 0
lTargetRow = Range("A5").SpecialCells(xlLastCell).Row

With ActiveSheet.Sort
    .SortFields.Clear
    .SortFields.Add Key:=Range( _
    "A5:A" & CStr(lTargetRow)), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
    xlSortNormal
    .SetRange Range("A5:C" & CStr(lTargetRow))
    .Header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With

MsgBox "Definitions formated and sorted", vbCritical

End Sub

' +++
' https://www.mrexcel.com/forum/excel-questions/86056-any-way-visual-basic-applications-change-file-modified-property.html
Option Explicit

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Private Const GENERIC_WRITE = &H40000000
Private Const OPEN_EXISTING = 3
Private Const FILE_SHARE_READ = &H1
Private Const FILE_SHARE_WRITE = &H2
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function SetFileTime Lib "kernel32" (ByVal hFile As Long, lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long
Private Declare Function SystemTimeToFileTime Lib "kernel32" (lpSystemTime As SYSTEMTIME, lpFileTime As FILETIME) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function LocalFileTimeToFileTime Lib "kernel32" (lpLocalFileTime As FILETIME, lpFileTime As FILETIME) As Long
Private Declare Function GetLastError Lib "kernel32" () As Long

Private Function AdjustFileTime(strFilePath As String, WriteFileDate As Date, CreateFileDate As Date, AccessFileDate As Date) As Long
Dim NewWriteDate As Date, NewCreateDate As Date, NewAccessDate As Date, lngHandle As Long

Dim udtWriteTime As FILETIME
Dim udtCreateTime As FILETIME
Dim udtAccessTime As FILETIME

Dim udtSysCreateTime As SYSTEMTIME
Dim udtSysAccessTime As SYSTEMTIME
Dim udtSysWriteTime As SYSTEMTIME

Dim udtLocalCreateTime As FILETIME
Dim udtLocalAccessTime As FILETIME
Dim udtLocalWriteTime As FILETIME

NewCreateDate = Format(CreateFileDate, "DD-MM-YY HH:mm:SS")
NewAccessDate = Format(AccessFileDate, "DD-MM-YY HH:mm:SS")
NewWriteDate = Format(WriteFileDate, "DD-MM-YY HH:mm:SS")

With udtSysCreateTime
    .wYear = Year(NewCreateDate)
    .wMonth = Month(NewCreateDate)
    .wDay = Day(NewCreateDate)
    .wDayOfWeek = Weekday(NewCreateDate) - 1
    .wHour = Hour(NewCreateDate)
    .wMinute = Minute(NewCreateDate)
    .wSecond = Second(NewCreateDate)
    .wMilliseconds = 0
End With

With udtSysAccessTime
    .wYear = Year(NewAccessDate)
    .wMonth = Month(NewAccessDate)
    .wDay = Day(NewAccessDate)
    .wDayOfWeek = Weekday(NewAccessDate) - 1
    .wHour = Hour(NewAccessDate)
    .wMinute = Minute(NewAccessDate)
    .wSecond = Second(NewAccessDate)
    .wMilliseconds = 0
End With

With udtSysWriteTime
    .wYear = Year(NewWriteDate)
    .wMonth = Month(NewWriteDate)
    .wDay = Day(NewWriteDate)
    .wDayOfWeek = Weekday(NewWriteDate) - 1
    .wHour = Hour(NewWriteDate)
    .wMinute = Minute(NewWriteDate)
    .wSecond = Second(NewWriteDate)
    .wMilliseconds = 0
End With
Dim ret As Long
ret = SystemTimeToFileTime(udtSysCreateTime, udtLocalCreateTime)
If ret <> 1 Then Err.Raise GetLastError
ret = LocalFileTimeToFileTime(udtLocalCreateTime, udtCreateTime)
If ret <> 1 Then Err.Raise GetLastError

ret = SystemTimeToFileTime(udtSysAccessTime, udtLocalAccessTime)
If ret <> 1 Then Err.Raise GetLastError
ret = LocalFileTimeToFileTime(udtLocalAccessTime, udtAccessTime)
If ret <> 1 Then Err.Raise GetLastError

ret = SystemTimeToFileTime(udtSysWriteTime, udtLocalWriteTime)
If ret <> 1 Then Err.Raise GetLastError
ret = LocalFileTimeToFileTime(udtLocalWriteTime, udtWriteTime)
If ret <> 1 Then Err.Raise GetLastError

lngHandle = CreateFile(strFilePath, GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, OPEN_EXISTING, 0, 0)
If lngHandle = -1 Then Err.Raise 53

'                                 create,      access,      write
ret = SetFileTime(lngHandle, udtCreateTime, udtAccessTime, udtWriteTime)
CloseHandle lngHandle
AdjustFileTime = 1
If ret <> 1 Then Err.Raise GetLastError

End Function
' ++++


Sub foo()
Dim i As Long
On Error GoTo ERR_HANDLER
i = AdjustFileTime("c:\text.txt", CDate("23/03/2017 11:13:17"), CDate("23/03/2017 16:49:43"), CDate("23/03/2017 18:09:26"))
Exit Sub
ERR_HANDLER:
MsgBox Err.Description
On Error GoTo 0
End Sub

 Sub Repeat_Row_Values()
Dim lMainRow, lSplitRow, lCol As Long

lMainRow = Selection.Row
lCol = Selection.Column

With ActiveWorkbook.ActiveSheet
While lMainRow < 12185
    lSplitRow = .Cells(lMainRow, lCol).End(xlDown).Row - 1
    .Range(.Cells(lMainRow + 1, lCol), .Cells(lSplitRow, lCol)) = .Cells(lMainRow, lCol)
    lMainRow = .Cells(lMainRow, lCol).End(xlDown).Row
Wend
End With
MsgBox "Repeat row values done", vbExclamation, "Repeast row values"
End Sub

Sub fill_ref()
Dim from_num, to_num, i  As Integer
Dim lRow, lCol As Long

lRow = Selection.Row
lCol = Selection.Column
from_num = Int(InputBox("From number: "))
to_num = Int(InputBox("To number: "))

With ActiveWorkbook.ActiveSheet
For i = 1 To (to_num + 1 - from_num)
    .Cells(lRow, lCol) = .Cells(lRow, lCol) & Trim(Str(.Cells(lRow, 1))) & "." & Trim(Str(from_num)) & Chr(10)
    from_num = from_num + 1
Next i
Selection.Offset(0, 1).Select
End With

'oldStatusBar = Application.DisplayStatusBar 
'Application.DisplayStatusBar = True 
'Application.StatusBar = "Please be patient..." 
'Application.StatusBar = False 
'Application.DisplayStatusBar = oldStatusBar

End Sub

Sub get_tabs_and_attributes()
Dim i, lRow, lUptoRow As Long
Application.ScreenUpdating = False

With ActiveWorkbook
If .Sheets(1).Name <> "Index" Then
    .Sheets.Add Before:=Sheets(1)
    .Sheets(1).Name = "Index"
Else
    .Sheets(1).Range("A:C").Clear
End If
.Sheets(1).Range("A1") = "Template"
.Sheets(1).Range("B1") = "Tab"
.Sheets(1).Range("C1") = "Attribute"
lRow = 2
For i = 2 To .Sheets.Count
    If .Sheets(i).Visible = True Then
        If .Sheets(i).Cells(9, 2) = "Field Name" Then
            .Sheets(1).Cells(lRow, 2) = .Sheets(i).Name
            lUptoRow = .Sheets(i).Range("B10").End(xlDown).Row
            .Sheets(i).Range("B10:B" & lUptoRow).Copy
            .Sheets(1).Cells(lRow, 3).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:= _
            False
            lRow = .Sheets(1).UsedRange.Rows.Count + 1
        Else
            .Sheets(1).Cells(lRow, 1) = .Sheets(i).Name
        End If
    End If

Next i
.Sheets(1).Columns("A:C").AutoFit
End With
Application.ScreenUpdating = True
MsgBox "Get Tabs & Fields", vbExclamation, "Get Tabs & Fields done!"
End Sub

Sub Merge_Split_Cols()
Dim lMainRow, lSplitRow As Long
Dim iColToMerge As Integer


lMainRow = Selection.Row

With ActiveWorkbook.ActiveSheet
While .Cells(lMainRow, 1) <> "Until this row"
    If .Cells(lMainRow + 1, 1) = "" Then
            lSplitRow = .Cells(lMainRow, 1).End(xlDown).Row - 1
        While lMainRow < lSplitRow
            iColToMerge = 10
            While iColToMerge > 0
            Range(.Cells(lMainRow, iColToMerge), .Cells(lSplitRow, iColToMerge)).Merge
            iColToMerge = iColToMerge - 1
            Wend
            lMainRow = lSplitRow
        Wend
    End If
    lMainRow = lMainRow + 1
Wend
End With
MsgBox "Merge of Split Cols done", vbExclamation, "Split Merge Cols"
End Sub

Sub Move2VisibleCell()
    Do
        ActiveCell.Offset(1, 0).Select
    
    Loop While ActiveCell.EntireRow.Hidden
    
End Sub