Sub HighlightChangesInExcel(ByVal sQry As String)
   
    Dim xl As Object
    Dim fd As FileDialog
    Dim lTotalRows, lRow, lCol As Long
    Dim rsExport As DAO.QueryDef
        
    ' Creating an external file dialog box object to save/create an excel file
    Set fd = Application.FileDialog(msoFileDialogSaveAs)

    'Assigning name to the excel file with the date
    fd.InitialFileName = "Highlight_DD_Changes_" & Format(Now(), "ddMMMyyyy_hhmm") & ".xlsx"

    ' Show the dialog box. If the .Show method returns True, the user picked at least one file
    If fd.Show = True Then
    'If atleast one item is selected
        If Format(fd.SelectedItems(1)) <> vbNullString Then
            ' if the query already exists then delete it
          If QueryExists("HighlightChanges_" & sUserLogin) Then CurrentDb.QueryDefs.Delete "HighlightChanges_" & sUserLogin
            'creating query definition with the SQL parameter
            ' HighlightChanges Query is like a stored procedure which executes on the fly to fetch the result.
            ' The query displays the previous approved version of submitted record (Approved) OR -  if approved doesn't exists -  the previous month existing records (Existing), the corresponding modified records(Modified) in the next line and also the newly added records for the current month (New)

            Set rsExport = CurrentDb.CreateQueryDef("HighlightChanges_" & sUserLogin, sQry)
            
            'Exporting the result and transferring it to the excel
            DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12Xml, rsExport.Name, fd.SelectedItems(1), True
            CurrentDb.QueryDefs.Delete rsExport.Name

            ' Compare and highlight changes
            
            'Creating an Excel file object
            Set xl = CreateObject("Excel.Application")
            'Opening the exported file
            xl.workbooks.Open (fd.SelectedItems(1))
                        
            ' Hiding the visibility of Excel and screen updates
            xl.Visible = False
            xl.screenupdating = False
            With xl.activeworkbook.ActiveSheet
            'Selecting the sheet with the name Highlight DD Changes
            .Name = "Highlight DD Changes"
                'Counting Total used rows in the excel
                lTotalRows = .usedrange.rows.Count
                'Entire used range of cells are made to no fill color
                .usedrange.Interior.ColorIndex = 0
                'making the First Row headings bold
                .Range("A1:Az1").Font.Bold = True
                
                'The query result in the excel file is such that the existing and the modified records for the same MDD_ID are in adjacent rows and hence we use this loop to compare one cell (existing value in DD)and the cell below the selected cell
                '(The modified value in current month's DD)
                For lRow = 2 To lTotalRows
                'Changing the column color for all the DD Cells to one particular color if they are of deprectaed status
                    If .Cells(lRow, 28) = "Deprecated" Then .Range(.Cells(lRow, 1), .Cells(lRow, 36)).Interior.ColorIndex = 15
                'Comparing the Modified DD cell values to the existing DD cell values
                    If (.Cells(lRow, 1) = "Modified" And .Cells(lRow, 2) = .Cells(lRow - 1, 2) And (.Cells(lRow, 3) = .Cells(lRow - 1, 3))) Then
                        For lCol = 3 To 36
                        'Changing the column color for all the Modified DD Cells to one particular color
                            If .Cells(lRow, lCol) <> .Cells(lRow - 1, lCol) Then .Cells(lRow, lCol).Interior.ColorIndex = 6
                            
                        Next lCol
                    End If
                Next lRow
                .Cells(1, 1).Select
            End With
            
            '  save the workbook

                xl.activeworkbook.Save
                MsgBox "Spreadsheet ready for reivew.", , "Highlight DD Changes"
                xl.screenupdating = True
                xl.Visible = True

                Set xl = Nothing
        End If
    End If
End Sub

Sub RemoveUsageDetail()
' function to remove usage detail - one at a time, for all MDD entries
Dim sUsgToRemove As String
Dim iEntriesUpdated As Long
Static db As DAO.Database

If db Is Nothing Then Set db = CurrentDb()
' get Usage to remove from user
sUsgToRemove = UCase(InputBox("Removes the given usage from ALL entries in MDD." & Chr(13) & Chr(13) & "Enter Usage to be removed (one at a time)", "What to remove?"))

' validate the Usage given
If Nz(sUsgToRemove, "") = "" Then
    MsgBox "Blank usage provided, nothing to do.", vbOKOnly
    Exit Sub
End If
If DCount("USAGE_DETAILS", "USAGE_DETAILS", "USAGE_DETAILS = '" & sUsgToRemove & "'") < 1 Then
    MsgBox "Usage detail invalid.  Refer to Usage Details master list for valid usage detail."
    Exit Sub
End If
If InStr(sUsgToRemove, "*") Then
    MsgBox "Usage detail input  cannot contain '*'."
    Exit Sub
End If

' diplay a warning about removal from all entries in MDD
' Quit if the user chooses not to proceed
If MsgBox("Removes '" & sUsgToRemove & "' from ALL entries in MDD." & Chr(13) & Chr(13) & "Do you want to proceed?", vbYesNo + vbQuestion, "Usage Details removal") = vbNo Then Exit Sub

' remove from MDD entries which has sigle usage
db.Execute "UPDATE MDD_Temp SET Usage_Details = null where Status <> 'Deprecated' and Usage_Details = '" & sUsgToRemove & "'", dbFailOnError
iEntriesUpdated = db.RecordsAffected

' remove usage present at the begining
db.Execute "UPDATE MDD_Temp SET Usage_Details = MID(Usage_Details," & Len(sUsgToRemove) + 3 & ",500 ) where Status <> 'Deprecated' and  Usage_Details like '" & sUsgToRemove & "; *'", dbFailOnError
iEntriesUpdated = iEntriesUpdated + db.RecordsAffected

' remove usage from the middle
db.Execute "UPDATE MDD_Temp SET Usage_Details = LEFT(Usage_Details,LEN(Usage_Details) -" & Len(sUsgToRemove) + 2 & ") where Status <> 'Deprecated' and  Usage_Details like '*; " & sUsgToRemove & "'", dbFailOnError
iEntriesUpdated = iEntriesUpdated + db.RecordsAffected

'remove usage present at the end
db.Execute "UPDATE MDD_Temp SET Usage_Details = Replace(Usage_Details, '; " & sUsgToRemove & "; ' , '; ' ) where Status <> 'Deprecated' and  Usage_Details like '*; " & sUsgToRemove & "; *' ", dbFailOnError
iEntriesUpdated = iEntriesUpdated + db.RecordsAffected

' Log activity
db.Execute "insert into Audit_Log_TBL (UserID, ENTRIES_UPLOADED, ENTRIES_ERROR, ERROR_CATEGORY, UPLOAD_STATUS)  VALUES " & _
    " ('" & sUserLogin & "',' " & CStr(iEntriesUpdated) & "', Null, 'Removed Usage detail: " & sUsgToRemove & "' ,'Successful' ) "

'Display completion message
MsgBox "Usage detail '" & sUsgToRemove & "' removed from " & CStr(iEntriesUpdated) & " MDD entries.", vbExclamation + vbOKOnly, "Removed usage detail"

End Sub

Public Sub LoadDataGroups(ByRef oTreeView As Object, Optional ByVal bIncludeAll As Boolean = False)
    Dim tv As MSComctlLib.TreeView
    Dim nodX, nodX1 As MSComctlLib.Node
    Dim rsReqs As DAO.Recordset
    Dim rsreq As DAO.Recordset
    Dim sDataGroup As String
        
    Set tv = oTreeView.Object
    tv.Nodes.Clear
    
    If bIncludeAll Then Set nodX = tv.Nodes.Add(, , , "All")
    
    Set rsReqs = CurrentDb.OpenRecordset("SELECT DISTINCT DATA_GROUP,DATA_SUB_GROUP FROM DATA_GROUP_TBL", dbOpenDynaset) 'uncomment
         
    rsReqs.MoveFirst
    sDataGroup = rsReqs!DATA_GROUP
    Set nodX = tv.Nodes.Add(, , , rsReqs!DATA_GROUP)
    
    Set nodX1 = tv.Nodes.Add(nodX, tvwChild, , rsReqs!Data_Sub_Group)
    rsReqs.MoveNext
    
    Do Until rsReqs.EOF
       If rsReqs!DATA_GROUP = sDataGroup Then
            Set nodX1 = tv.Nodes.Add(nodX, tvwChild, , rsReqs!Data_Sub_Group)
            
       Else
            Set nodX = tv.Nodes.Add(, , , rsReqs!DATA_GROUP)
            Set nodX1 = tv.Nodes.Add(nodX, tvwChild, , rsReqs!Data_Sub_Group)
            sDataGroup = rsReqs!DATA_GROUP
       End If
        rsReqs.MoveNext
    Loop

End Sub
' function to export Error Log
Sub ExportErrLog()
    Dim db As DAO.Database, rsExport As DAO.QueryDef
    Dim exportSQL As String
    
    Set db = CurrentDb
    'select all the errors from the error log table
    exportSQL = " SELECT MDD_ID, CAN_ID, REF_ID, NAME, ERR_CODE, ERR_DESC, DD_VALUE, EXPECTED_VALUE  FROM ERROR_LOG_TBL WHERE USERID = '" & sUserLogin & "'"
      ' if the query already exists then delete it
      If QueryExists("ERROR_LOG_TBL_Export") Then CurrentDb.QueryDefs.Delete "ERROR_LOG_TBL_Export"
      
    Set rsExport = db.CreateQueryDef("ERROR_LOG_TBL_Export", exportSQL)
    
    'extracting all the error log into a file at save it on the shared drive
    DoCmd.TransferText acExportDelim, , "ERROR_LOG_TBL_Export", CurrentProject.Path & "\ERROR_LOG_FILES\ERROR_LOG_" & sUserLogin & "_" & Format(CStr(Now), "yyyy-mm-dd_hhmm") & ".csv", True
    
    'cleanup
    CurrentDb.QueryDefs.Delete rsExport.Name
    db.Close
    Set db = Nothing
End Sub
' Function to check if the file is already open
Function IsFileOpen(FileName As String)
    Dim iFilenum As Long
    Dim iErr As Long
     
    On Error Resume Next
    iFilenum = FreeFile()
    Open FileName For Input Lock Read Write As #iFilenum
    iErr = err
    Close iFilenum
    
    On Error GoTo 0
     
    Select Case iErr
    Case 0:    IsFileOpen = False
    Case 70:   IsFileOpen = True
    Case Else: Error iErr
    End Select
     
End Function

' funtion to check if a query exists
Function QueryExists(queryName As String)
    Dim qryDefn As QueryDef
    ' default value if false
    QueryExists = False
    
    ' check all queries in current DB
    For Each qryDefn In CurrentDb.QueryDefs
        If qryDefn.Name = queryName Then
         QueryExists = True ' set flag to true
        Exit For
        End If
    Next qryDefn
    
End Function

'Fucntion to return the filename for the selected file
Private Function selectFile()

Dim fd As FileDialog

On Error GoTo ErrorHandler
 
Set fd = Application.FileDialog(msoFileDialogFilePicker)

fd.AllowMultiSelect = False

With fd
.Title = "Select the Excel file to import"
.AllowMultiSelect = False
.Filters.Clear

If fd.Show = True Then
    If fd.SelectedItems(1) <> vbNullString Then
        FileName = fd.SelectedItems(1)
    End If
Else
    'Exit code if no file is selected
    End
End If
End With
    'Return Selected FileName
    selectFile = FileName
 
    Set fd = Nothing
 
Exit Function
 
ErrorHandler:
    Set fd = Nothing
    MsgBox "Error " & err & ": " & Error(err)
 
End Function

'function to Export MDD entries from MDD tool and save it to another external file based on the SQL query received as parameters
Private Sub exportMDD(ByVal exportSQL As String, ByVal sfileName As String)
Dim db As DAO.Database, qd As DAO.QueryDef
Dim fd As FileDialog

Set fd = Application.FileDialog(msoFileDialogSaveAs)
fd.InitialFileName = sfileName

Set db = CurrentDb

If fd.Show = True Then
    If Format(fd.SelectedItems(1)) <> vbNullString Then
         ' if the query already exists then delete it
          If QueryExists("ExportedEntries") Then CurrentDb.QueryDefs.Delete "ExportedEntries"
          
        ' executing the SQL query received as parameter to the function
        Set qd = db.CreateQueryDef("ExportedEntries", exportSQL)
        'Exporting the result and transferring it to the excel
                DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12Xml, qd.Name, fd.SelectedItems(1), True
        
        CurrentDb.QueryDefs.Delete qd.Name
        
    Else
        MsgBox "File name cannot be blank."
        
    End If
Else
    Exit Sub
End If

'cleanup

db.Close
Set db = Nothing
Set qd = Nothing
Set fd = Nothing

End Sub

'Function to add / remove  usage details
Sub ChangeUsageDetails()
Dim rsBulkUpdates As DAO.Recordset
Dim sSplitUsage() As String
Dim sRevisedUsage, sAddedUsages, sRemovedUsages, sUsageNotFound As String
Dim colUniqueUsages As New Collection, aUsage
Dim i As Integer
sUserLogin = Environ("Username")
     
'select the submitted usage_details  mdd_id and status
Set rsBulkUpdates = CurrentDb.OpenRecordset("SELECT mdd_id, usage_details, remove_usage_details, status FROM REF_ID_TRACEABILITY_CHANGES WHERE STATUS IS NULL AND UserID = '" & sUserLogin & "' AND ( nz(REMOVE_USAGE_DETAILS,'') <> '' OR nz(USAGE_DETAILS,'') <> '') ")

If Not (rsBulkUpdates.BOF And rsBulkUpdates.EOF) Then ' atleast one record to work on
    ' Loop till end of file
    Do Until rsBulkUpdates.EOF
        sRevisedUsage = ""
        ' get existing usage detail from MDD Temp
        sRevisedUsage = DLookup("USAGE_DETAILS", "MDD_TEMP", " MDD_ID = " & Nz(rsBulkUpdates(0), -1))
        
        'Splitting all the Usage details based on ';'
        sSplitUsage = Split(Nz(sRevisedUsage, ""), "; ", , vbBinaryCompare)
        
        'Add splitted usage details into a collection.  Note that collection by design rejects duplicates
        On Error Resume Next
        Set colUniqueUsages = Nothing
        For Each aUsage In sSplitUsage
           colUniqueUsages.Add Trim(aUsage), Trim(aUsage)
        Next
        sAddedUsages = ""
        sRemovedUsages = ""
        sUsageNotFound = ""
        
        For i = 1 To 2 ' 1 = add usage details, 2 = remove
             sSplitUsage = Empty
            'Splitting all Usage details to be added based on ';'
             sSplitUsage = Split(Nz(rsBulkUpdates(i), ""), "; ", , vbBinaryCompare)
             
            'If splitted usage details are not null
             If UBound(sSplitUsage) > -1 Then
                ' Validate each usage detail against master.
                ' if its invalid, then log issue
                ' if its valid, then add to collection is round 1 (i=1)  and remove it from collection in round 2
                For Each aUsage In sSplitUsage
                    aUsage = Trim(aUsage) ' trim spaces
                    If Len(aUsage) > 0 Then
                        If Right(aUsage, 1) = ";" Then aUsage = Left(aUsage, Len(aUsage) - 1)
                            
                       'Counting the usage detail in USAGE_DETAILS table
                        If (DCount("USAGE_DETAILS", "USAGE_DETAILS", "USAGE_DETAILS ='" & aUsage & "'") < 1) Then
                            ' Usage details NOT present in master table
                            'Updating the corresponding records in bulk update table with invalid usage detail provided
                            DoCmd.RunSQL "UPDATE REF_ID_TRACEABILITY_CHANGES SET Status = nz(Status,'') + '" & aUsage & "; ' where MDD_ID = " & rsBulkUpdates(0)
        
                        Else ' usage detail is valid
                            If i = 1 Then ' Usage to be added
                                colUniqueUsages.Add aUsage, aUsage
                                sAddedUsages = sAddedUsages & aUsage & "; "
                            Else ' Usage to be removed
                                ' check if the usage to remove is present in current mdd
                                If Contains(colUniqueUsages, aUsage) = False Then
                                    sUsageNotFound = sUsageNotFound & aUsage & "; "
                                Else
                                    colUniqueUsages.Remove aUsage
                                    sRemovedUsages = sRemovedUsages & aUsage & "; "
                                End If
                            End If
                        End If
                    End If
                Next aUsage
            End If
        Next i
        
        'colUniqueUsages contains final list of usage details
        'convert it back to ; separated list
        sRevisedUsage = ""
        For Each aUsage In colUniqueUsages
            sRevisedUsage = sRevisedUsage & aUsage & "; "
        Next aUsage
        
        'Remove the last ';' and space
        sRevisedUsage = Mid(sRevisedUsage, 1, Len(sRevisedUsage) - 2)
        
        'Update current month MDD with the new usage details
        DoCmd.RunSQL "UPDATE MDD_Temp SET USAGE_DETAILS = '" & sRevisedUsage & "',  ERR_CODE = '', ERR_DESC = '' where MDD_ID = " & Nz(rsBulkUpdates(0), -1)
        
        ' Log additions to  Usages
        If Len(sAddedUsages) > 0 Then
            DoCmd.RunSQL "insert into Audit_Log_TBL (UserID, ENTRIES_UPLOADED, ENTRIES_ERROR, ERROR_CATEGORY, UPLOAD_STATUS)  VALUES " & _
                " ('" & sUserLogin & "',' MDD_ID: " & CStr(Nz(rsBulkUpdates(0), -1)) & "', '', ' " & sAddedUsages & "' ,'Usage added' ) "
        End If
        ' Log removal of  Usages
        If Len(sRemovedUsages) > 0 Then
            DoCmd.RunSQL "insert into Audit_Log_TBL (UserID, ENTRIES_UPLOADED, ENTRIES_ERROR, ERROR_CATEGORY, UPLOAD_STATUS)  VALUES " & _
                " ('" & sUserLogin & "',' MDD_ID: " & CStr(Nz(rsBulkUpdates(0), -1)) & "', '', ' " & sRemovedUsages & "' ,'Usage removed' ) "
        End If

        ' Log unfound Usages
        If Len(sUsageNotFound) > 0 Then
            DoCmd.RunSQL "insert into Audit_Log_TBL (UserID, ENTRIES_UPLOADED, ENTRIES_ERROR, ERROR_CATEGORY, UPLOAD_STATUS)  VALUES " & _
                " ('" & sUserLogin & "',' MDD_ID: " & CStr(Nz(rsBulkUpdates(0), -1)) & "', '', ' " & sUsageNotFound & "' ,'Usage not found' ) "
        End If
        
        rsBulkUpdates.MoveNext
    Loop
End If

' Refine the Status message in bulk update table
' In all submissions with invalid usage, Status has list of invalid usage details; add 'Invalid usage details:' in front of it
DoCmd.RunSQL "UPDATE REF_ID_TRACEABILITY_CHANGES SET Status = 'Invalid Usage Details: ' + Status  WHERE UserID = '" & sUserLogin & "' AND Status like '*; *' AND Status NOT LIKE 'Invalid Usage *' "

 End Sub
 
 
' function to check if an item exists in a collection

Public Function Contains(col As Collection, key As Variant) As Boolean
Dim obj As Variant
On Error GoTo err
    Contains = True
    obj = col(key)
    Exit Function
err:

    Contains = False
End Function