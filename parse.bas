''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' Flexible(-ish) parser of semi-regular data
''' Alexander Ivashkin, April 2017
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'     .--.              .--.
'    : (\ ". _......_ ." /) :
'     '.    `        `    .'
'      /'   _        _   `\
'     /     0}      {0     \
'    |       /      \       |
'    |     /'        `\     |
'     \   | .  .==.  . |   /
'      '._ \.' \__/ './ _.'
' jgs  /  ``'._-''-_.'``  \
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''' CUSTOMISABLE CONSTANTS
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' What to look for to find the beginning and the end of the Index Sheet
Const cStartOfIndexTable = "Index Details"
Const cEndOfIndexTable = " "
' Sanity check: if the Index Table is larger than this constant, then something is fishy
Const cMaxSizeOfIndexTable = 2000

' Column offset in the Index Sheet to get to the hyperlink
Const cColOffset_Hyperlink = 14

' Max. size of a project.
' Used for: when we don't have a next project to define the size of the table with the project details
'   and to find the beginning of the next project
Const cMaxHeightOfProject = 45
' Max width of a project (used to select a Range that would include every possible column in a project)
Const cMaxWidthOfProject = 30
' From where in Current Project to start looking for the Next Project
Const cRowNextProjectSearchOffset = 3
' How to find the next project
Const cNextProjectSignature = "Project Manager: "
' How much rows/columns to subtract from that signature to get to the next project's beginning
Const cRowNextProjectOffset = -2
Const cColNextProjectOffset = 0

' How many rows of bullsh... erm, fillers, to print.
' Dirty dirty hack to force Access choose the right data type on import from Excel
Const cCountOfDummyFillers = 25

' Where to save the exported file by default
Const cExportFileDir As String = "\\server\share"

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''' END OF CUSTOMISABLE CONSTANTS.
'''''' Venture forth at your own risk.
'''''' Here be dragons.
'''''' Or Russian bears.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



''' TODO replace hard-coded subscripts in Data Structure arrays with variables (populated dynamically to reflect any changes)
    

Option Explicit
Option Base 1
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''' Caveat:
'''''' All the arrays are being populated starting from subscript 1,
'''''' for consistency with arrays auto-created from Excel Tables
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' Global arrays with data structures
Public aDataStructure_ProjectDetails As Variant
Public aDataStructure_IndexSheet As Variant
' Count of special cases in data structures - upd: No more special cases as of 14 April
'Public iDataStructure_ProjectDetailsSpecialCases As Integer
'Public iDataStructure_IndexSheetSpecialCases As Integer

' Beginning and end of the Index Shit
Public iRow_IndexBeginning As Integer, iRow_IndexEnd As Integer


Private Sub ConvertPSR()
    
    Dim wbPSR As Workbook
    Dim vPSRFileName As Variant
    
    On Error GoTo Hell
    
    vPSRFileName = Application.GetOpenFilename("Excel Workbooks (*.xls;*.xlsx),*.xls;*.xlsx", , "Show me the file!", , False)

    If vPSRFileName <> False Then
        Set wbPSR = Application.Workbooks.Open(vPSRFileName)
        If Not wbPSR Is Nothing Then
            Application.ScreenUpdating = False
            Application.DisplayAlerts = False
            Call TraverseProjects(wbPSR.Sheets(1))
            wbPSR.Close (False)
        Else
            MsgBox "Something went horribly wrong during opening of the " & vPSRFileName & " file", , "A terrible tradegy!"
            Call PSR_Cleanup
            End
        End If
    Else
        MsgBox "Hey! I wanted to munch on a new PSR file!", , "Cancelled. Boring!"
        Call PSR_Cleanup
        End
    End If

    Call PSR_Cleanup
Exit Sub
Hell:
    Call PSR_Cleanup
    Call CrashDump("ConvertPSR")
End Sub


Public Sub ConvertThisShit()
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Call TraverseProjects(ActiveSheet)
    Call PSR_Cleanup
End Sub

Private Sub TraverseProjects(shSheet As Worksheet)

    On Error GoTo Hell ' Go To statement considered harmful
    
    Dim dTimer As Double
    dTimer = Timer
    
    
    'Application.ScreenUpdating = False
    Application.Cursor = xlWait

    'iDataStructure_ProjectDetailsSpecialCases = 0
    'iDataStructure_IndexSheetSpecialCases = 0
    
    shSheet.Activate
    
    Call PopulateDataStructureArrays
    Call FindIndexTable
    
    If Not IsIndexSheetDataStructureKosher Then
        'Application.ScreenUpdating = True
        Application.Cursor = xlDefault
        Exit Sub
    End If

    Dim sBack As String
    
    Dim aProjectDetails() As String
    Dim aProjectDetailsTemp() As String
    Dim sTempProjectDetails As String
    
    Dim i As Integer
    Dim iRow_Current As Integer
    Dim iCurrentProjectIdx As Integer
    
    Dim iProjectDetails_Dimension1 As Integer
    iProjectDetails_Dimension1 = UBound(aDataStructure_ProjectDetails) + UBound(aDataStructure_IndexSheet)
    
    ' 1000 rows times DataStructure columns. Redimmed once every 1000 rows. Waste of memory but saving of CPU cycles.
    ' Redimmed at the very end to trim the array (and calculate the number of exported projects).
    ReDim aProjectDetails(iProjectDetails_Dimension1, 1000)
    
    ReDim aProjectDetailsTemp(iProjectDetails_Dimension1)
    
    
    iCurrentProjectIdx = 1
    Cells(iRow_IndexBeginning + 1, 1).Activate
    iRow_Current = ActiveCell.Row
    
    Dim rProjectStart As Range
    Dim rProjectEnd As Range
    Dim iMaxProjectDetailsSize As Integer
    iMaxProjectDetailsSize = 0
    
    While iRow_Current <= iRow_IndexEnd
    
        ' Fill array with data from the Index Sheet
        For i = 1 To UBound(aDataStructure_IndexSheet)
            sTempProjectDetails = Cells(iRow_Current, aDataStructure_IndexSheet(i, 2)).Value
            
            ' Any regex parsing?
            If aDataStructure_IndexSheet(i, 3) <> "" Then
                sTempProjectDetails = PSR_RegExp_Match(sTempProjectDetails, aDataStructure_IndexSheet(i, 3), False, 0, 0)
            End If
        
            aProjectDetails(i, iCurrentProjectIdx) = Bowdlerise(sTempProjectDetails)
        
        Next i
        ' Finished grabbing data from the Index Sheet

        
        ' Down the rabbit hole!
        sBack = ActiveCell.Address
        'ActiveCell.Offset(0, cColOffset_Hyperlink).Range("A1").Select
        'Selection.Hyperlinks(1).Follow NewWindow:=False, AddHistory:=True
        Cells(ActiveCell.Row, ActiveCell.Column + cColOffset_Hyperlink).Hyperlinks(1).Follow NewWindow:=False, AddHistory:=True
        
        ' Let's save the beginning of the current project and find the beginning of the next one (or lack thereof)
        
        Set rProjectStart = ActiveCell
        'Range(sBack).Activate
        
        Set rProjectEnd = Nothing
        Set rProjectEnd = Range(rProjectStart.Offset(cRowNextProjectSearchOffset, 0), rProjectStart.Offset(cMaxHeightOfProject, cMaxWidthOfProject)).Find(What:=cNextProjectSignature, After:=rProjectStart.Offset(cRowNextProjectSearchOffset, 0), LookIn:= _
                    xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:= _
                    xlNext, MatchCase:=True, SearchFormat:=False)
        
        If rProjectEnd Is Nothing Then
            ' It must be the last project!
            'MsgBox "Could not find the next project's signature: """ & cNextProjectSignature & """" & vbCrLf & "Project start address: " & cProjectStart.Address, vbCritical, "What is going on here?!"
            'Call PSR_Cleanup
            'End
            Set rProjectEnd = rProjectStart.Offset(cMaxHeightOfProject - 1, cMaxWidthOfProject)
        Else
            Set rProjectEnd = rProjectEnd.Offset(cRowNextProjectOffset, cColNextProjectOffset + cMaxWidthOfProject)
        End If
                
        If iMaxProjectDetailsSize < rProjectEnd.Row - rProjectStart.Row Then
            iMaxProjectDetailsSize = rProjectEnd.Row - rProjectStart.Row
            'Debug.Assert iMaxProjectDetailsSize < 50
        End If

        Range(sBack).Activate
        
        
        '
        ''' OLD version: hopping hyperlinks one-by-one to find the next project. Won't work due to some projects being out of order.
        '
        'If Cells(ActiveCell.Row + 1, ActiveCell.Column + cColOffset_Hyperlink).Hyperlinks.Count > 0 Then
        '    Cells(ActiveCell.Row + 1, ActiveCell.Column + cColOffset_Hyperlink).Hyperlinks(1).Follow NewWindow:=False, AddHistory:=True
        '    Set rProjectEnd = ActiveCell.Offset(-1, cMaxWidthOfProject)
        '    If iMaxProjectDetailsSize < rProjectEnd.Row - rProjectStart.Row Then
        '        iMaxProjectDetailsSize = rProjectEnd.Row - rProjectStart.Row
        '        Debug.Assert iMaxProjectDetailsSize < 50
        '    End If
        '    'rProjectStart.Activate
        '    Range(sBack).Activate
        'Else
        '    Set rProjectEnd = rProjectStart.Offset(cMaxHeightOfProject - 1, cMaxWidthOfProject)
        'End If


        ' Offset: where to start puting the second part of the data into the array (i.e. Project Details after the Index Sheet)
        Dim iTempProjectDetailsOffset As Integer
        iTempProjectDetailsOffset = UBound(aDataStructure_IndexSheet)
        
        'Debug.Assert iCurrentProjectIdx <> 16
        aProjectDetailsTemp() = GetProjectDetails(rProjectStart, rProjectEnd)
        
        For i = 1 To UBound(aProjectDetailsTemp)
            aProjectDetails(i + iTempProjectDetailsOffset, iCurrentProjectIdx) = aProjectDetailsTemp(i)
        Next i
        
        If iCurrentProjectIdx = UBound(aProjectDetails, 2) Then ReDim Preserve aProjectDetails(iProjectDetails_Dimension1, UBound(aProjectDetails, 2) + 1000)
        
        Range(sBack).Activate
        
        ''' Next row
        'Cells(ActiveCell.Row + 1, 1).Select
        ActiveCell.Offset(1, 0).Select
        iRow_Current = ActiveCell.Row
        
        iCurrentProjectIdx = iCurrentProjectIdx + 1
        
    Wend
    
    iCurrentProjectIdx = iCurrentProjectIdx - 1
    
    ReDim Preserve aProjectDetails(iProjectDetails_Dimension1, iCurrentProjectIdx)
    
    dTimer = Timer - dTimer
    
    Dim vExportResult As Variant
    vExportResult = ExportProjectDetailsToNewBook(aProjectDetails())
    
    Dim sSuccessMessage As String
    sSuccessMessage = "Successfully parsed " & iCurrentProjectIdx & " projects." & vbCrLf
    If vExportResult = False Then
        sSuccessMessage = sSuccessMessage & "Export was cancelled."
    Else
        sSuccessMessage = sSuccessMessage & "Exported to: " & vExportResult
    End If
    sSuccessMessage = sSuccessMessage & vbCrLf & "Execution took " & dTimer & " seconds."
    sSuccessMessage = sSuccessMessage & vbCrLf & "Max project height (rows): " & iMaxProjectDetailsSize
    
    'Application.ScreenUpdating = True
    Application.Cursor = xlDefault
    MsgBox sSuccessMessage, vbInformation, "Great success! Nice!"
    'Debug.Print iMaxProjectDetailsSize
    
    Exit Sub

Hell:
    Call PSR_Cleanup
    Call CrashDump("TraverseProjects")
End Sub




Private Function GetProjectDetails(rProjectStart As Range, rProjectEnd As Range) As String()

    'Application.ScreenUpdating = False
    
    On Error GoTo Hell

    'If IsEmpty(aDataStructure_ProjectDetails) Then Call PopulateDataStructureArrays

    Dim i As Integer

    ' Unneeded
    'Dim iStartRow As Integer
    'Dim iStartColumn As Integer
    'iStartRow = ActiveCell.Row
    'iStartColumn = ActiveCell.Column
    
    Dim iCountProjectFields
    iCountProjectFields = UBound(aDataStructure_ProjectDetails)
    
    Dim aProjectDetails() As String
    

    ReDim aProjectDetails(iCountProjectFields)
    
    Dim rFoundStartingRange As Range
    Dim rFoundEndingRange As Range
    Dim rTempCell As Range
    Dim vLookAt As Variant
    Dim bMoreThanOneLine As Boolean
    Dim sTempString As String
    
    
    For i = 1 To UBound(aProjectDetails)
        
        ' What action shall we take to get to the data
        Select Case aDataStructure_ProjectDetails(i, 2)
            
            ' The simple case
            Case "Get one cell"
                ''''' Not using .Offset for bloody merged cells
                'aProjectDetails(i) = Trim(rProjectStart.Offset(CInt(aDataStructure_ProjectDetails(i, 3)), CInt(aDataStructure_ProjectDetails(i, 4))).Value)
                aProjectDetails(i) = Trim(Cells(rProjectStart.Row + CInt(aDataStructure_ProjectDetails(i, 3)), rProjectStart.Column + CInt(aDataStructure_ProjectDetails(i, 4))).Value)
       
            ' The funnier cases
            Case "Find and get one cell"
                
                'For Each rTempCell In Range(rProjectStart, rProjectEnd)
                '    If rTempCell.Value = aDataStructure_ProjectDetails(i, 5) Then
                '        Set rFoundStartingRange = rTempCell
                '        Exit For
                '    End If
                'Next rTempCell
                
                If aDataStructure_ProjectDetails(i, 5) = "" Then
                    MsgBox "Empty 'Find: starting text' in the Data Structure table (" & aDataStructure_ProjectDetails(i, 1) & ")", vbCritical, "Data structure is wicked, deprave and corrupt"
                    Call PSR_Cleanup
                    End
                End If
                
                Select Case aDataStructure_ProjectDetails(i, 7)
                    Case "xlWhole"
                        vLookAt = xlWhole
                    Case "xlPart"
                        vLookAt = xlPart
                    Case Else
                        MsgBox "Unsupported 'Match mode' in the Data Structure table: " & aDataStructure_ProjectDetails(i, 7) & vbCrLf & "(" & aDataStructure_ProjectDetails(i, 1) & ")", vbCritical, "Data structure is wicked, deprave and corrupt"
                        Call PSR_Cleanup
                        End
                End Select
                
                Set rFoundStartingRange = Nothing
                Set rFoundStartingRange = Range(rProjectStart, rProjectEnd).Find(What:=aDataStructure_ProjectDetails(i, 5), After:=rProjectStart, LookIn:= _
                    xlFormulas, LookAt:=vLookAt, SearchOrder:=xlByRows, SearchDirection:= _
                    xlNext, MatchCase:=True, SearchFormat:=False)
                
                ' If we haven't found the field's caption, something is utterly wrong here.
                If rFoundStartingRange Is Nothing Then
                    MsgBox "Could not find the following text: """ & aDataStructure_ProjectDetails(i, 5) & """" & vbCrLf & "Project start address: " & rProjectStart.Address & vbCrLf & "Hint: check Const cMaxHeightOfProject, it could be the culprit", vbCritical, "What is going on here?!"
                    Call PSR_Cleanup
                    End
                End If
                
                ' YES, we're invoking .Offset this time to accomodate for those pesky merged cells
                aProjectDetails(i) = rFoundStartingRange.Offset(CInt(aDataStructure_ProjectDetails(i, 3)), CInt(aDataStructure_ProjectDetails(i, 4))).Value

            Case "Find and get all cells between texts"

                

                If aDataStructure_ProjectDetails(i, 5) = "" Then
                    MsgBox "Empty 'Find: starting text' in the Data Structure table (" & aDataStructure_ProjectDetails(i, 1) & ")", vbCritical, "Data structure is wicked, deprave and corrupt"
                    Call PSR_Cleanup
                    End
                End If
                If aDataStructure_ProjectDetails(i, 6) = "" Then
                    MsgBox "Empty 'Find: ending text' in the Data Structure table (" & aDataStructure_ProjectDetails(i, 1) & ")", vbCritical, "Data structure is wicked, deprave and corrupt"
                    Call PSR_Cleanup
                    End
                End If
                
                Select Case aDataStructure_ProjectDetails(i, 7)
                    Case "xlWhole"
                        vLookAt = xlWhole
                    Case "xlPart"
                        vLookAt = xlPart
                    Case Else
                        MsgBox "Unsupported 'Match mode' in the Data Structure table: " & aDataStructure_ProjectDetails(i, 7) & vbCrLf & " (" & aDataStructure_ProjectDetails(i, 1) & ")", vbCritical, "Data structure is wicked, deprave and corrupt"
                        Call PSR_Cleanup
                        End
                End Select
                
                ' Let's find the beginning of the range
                Set rFoundStartingRange = Nothing
                Set rFoundStartingRange = Range(rProjectStart, rProjectEnd).Find(What:=aDataStructure_ProjectDetails(i, 5), After:=rProjectStart, LookIn:= _
                    xlFormulas, LookAt:=vLookAt, SearchOrder:=xlByRows, SearchDirection:= _
                    xlNext, MatchCase:=True, SearchFormat:=False)
                
                ' If we haven't found the field's caption, something is utterly wrong here.
                If rFoundStartingRange Is Nothing Then
                    MsgBox "Could not find the following text: """ & aDataStructure_ProjectDetails(i, 5) & """" & vbCrLf & "Project start address: " & rProjectStart.Address, vbCritical, "What is going on here?!"
                    Call PSR_Cleanup
                    End
                End If
                
                ' Let's find the end of the range
                Set rFoundEndingRange = Nothing
                Set rFoundEndingRange = Range(rProjectStart, rProjectEnd).Find(What:=aDataStructure_ProjectDetails(i, 6), After:=rProjectStart, LookIn:= _
                    xlFormulas, LookAt:=vLookAt, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True, SearchFormat:=False)
                
                ' If we haven't found the field's caption, something is utterly wrong here.
                If rFoundEndingRange Is Nothing Then
                    MsgBox "Could not find the following text: """ & aDataStructure_ProjectDetails(i, 6) & """" & vbCrLf & "Project start address: " & rProjectStart.Address, vbCritical, "What is going on here?!"
                    Call PSR_Cleanup
                    End
                End If
                
                ' Get all the cells between:
                ' - the Starting Range with Offset from the DataStructure
                ' - the Ending Range minus one row (i.e. not taking the Ending Text).
                ' Should have been a separate Row/ColumnOffset in the DataStructure for the Ending, but that'd make everything bigger and clumsier.
                sTempString = ""
                bMoreThanOneLine = False
                Set rFoundEndingRange = Cells(rFoundEndingRange.Row, rFoundEndingRange.Column + cMaxWidthOfProject)
                For Each rTempCell In Range(rFoundStartingRange.Offset(aDataStructure_ProjectDetails(i, 3), aDataStructure_ProjectDetails(i, 4)), rFoundEndingRange.Offset(-1, 0))
                    If rTempCell.Value <> "" Then
                        If bMoreThanOneLine = True Then sTempString = sTempString + vbCrLf
                        sTempString = sTempString + CStr(rTempCell.Value)
                        bMoreThanOneLine = True
                    End If
                Next rTempCell
                
                aProjectDetails(i) = Trim(sTempString)
    
        Case Else
            MsgBox "Unsupported 'Type of action' in the Data Structure table: " & aDataStructure_ProjectDetails(i, 2) & vbCrLf & " (" & aDataStructure_ProjectDetails(i, 1) & ")", vbCritical, "Data structure is wicked, deprave and corrupt"
            Call PSR_Cleanup
            End
        End Select
                        
        ' Any regex parsing?
        If aDataStructure_ProjectDetails(i, 8) <> "" Then
            aProjectDetails(i) = PSR_RegExp_Match(aProjectDetails(i), aDataStructure_ProjectDetails(i, 8), False, 0, 0)
        End If
        aProjectDetails(i) = Bowdlerise(aProjectDetails(i))
        
    Next i

    
    GetProjectDetails = aProjectDetails

    Exit Function

Hell:
    Call PSR_Cleanup
    Call CrashDump("GetProjectDetails")

End Function

Private Function ExportProjectDetailsToNewBook(ByRef aData() As String) As Variant

    Dim wbTempWb As Workbook
    Dim vFileName As Variant
    Dim iTemp As Integer
        
    On Error GoTo Hell
    
    iTemp = Application.SheetsInNewWorkbook
    Application.SheetsInNewWorkbook = 1
    Set wbTempWb = Application.Workbooks.Add
    Application.SheetsInNewWorkbook = iTemp
    
    Dim iRow_Current As Integer, iCol_Current As Integer, i As Integer, j As Integer
    
    ''''''''''''''''''''''''''''''''''''''''
    ' Print the headers
    
    iRow_Current = 1
    iCol_Current = 1
    
    For i = 1 To UBound(aDataStructure_IndexSheet)
        Cells(iRow_Current, iCol_Current).Value = aDataStructure_IndexSheet(i, 1)
        iCol_Current = iCol_Current + 1
    Next i
    
    For i = 1 To UBound(aDataStructure_ProjectDetails)
        Cells(iRow_Current, iCol_Current).Value = aDataStructure_ProjectDetails(i, 1)
        iCol_Current = iCol_Current + 1
    Next i
    
    '
    ''''''''''''''''''''''''''''''''''''''''
    
    
    
    ''''''''''''''''''''''''''''''''''''''''
    ' Paste the dummy fillers
    
    iRow_Current = 2
    iCol_Current = 1
    
    For j = 1 To cCountOfDummyFillers
    
        For i = 1 To UBound(aDataStructure_IndexSheet)
            Cells(iRow_Current, iCol_Current).Value = aDataStructure_IndexSheet(i, 4)
            iCol_Current = iCol_Current + 1
        Next i
        
        For i = 1 To UBound(aDataStructure_ProjectDetails)
            Cells(iRow_Current, iCol_Current).Value = aDataStructure_ProjectDetails(i, 9)
            iCol_Current = iCol_Current + 1
        Next i
        
        iRow_Current = iRow_Current + 1
        iCol_Current = 1
        
    Next j
    
    '
    ''''''''''''''''''''''''''''''''''''''''
    
    
    
    ''''''''''''''''''''''''''''''''''''''''
    ' Paste the data
    
    'iRow_Current = iRow_Current + 1
    iCol_Current = 1
    Dim bIsRowEmpty
        
    For i = 1 To UBound(aData, 2)
    
        bIsRowEmpty = True
        
        For j = 1 To UBound(aData, 1)
        
            If aData(j, i) <> "" Then
                Cells(iRow_Current, iCol_Current).Value = aData(j, i)
                bIsRowEmpty = False
            End If
            iCol_Current = iCol_Current + 1
        
        Next j
        
        If Not bIsRowEmpty Then iRow_Current = iRow_Current + 1
        iCol_Current = 1
    
    Next i
    
    '
    ''''''''''''''''''''''''''''''''''''''''
    
    
    vFileName = Application.GetSaveAsFilename(cExportFileDir & "PSR for import " & Date, "Excel Workbooks (*.xlsx),*.xlsx)", , "Place to save the processed PSR")
    If vFileName <> False Then wbTempWb.SaveAs (vFileName)
    wbTempWb.Close
    
    ExportProjectDetailsToNewBook = vFileName
    'Application.ScreenUpdating = True
    'Application.DisplayAlerts = True
    Exit Function

Hell:
    Call PSR_Cleanup
    Call CrashDump("ExportProjectDetailsToNewBook")
End Function


Private Sub PopulateDataStructureArrays()

    On Error GoTo Hell
    
    Dim tblDataStructure As ListObject
    Dim i As Integer

    ' "Project details"
    Set tblDataStructure = ThisWorkbook.Sheets(1).ListObjects("tblProjectDetailsDataStructure")
    aDataStructure_ProjectDetails = tblDataStructure.DataBodyRange
    
    ' "Index sheet"
    Set tblDataStructure = ThisWorkbook.Sheets(1).ListObjects("tblIndexSheetDataStructure")
    aDataStructure_IndexSheet = tblDataStructure.DataBodyRange
    
    
    Exit Sub

Hell:
    Call PSR_Cleanup
    Call CrashDump("PopulateDataStructureArrays")
End Sub


Private Sub FindIndexTable()
    On Error GoTo Hell
    
    Range("A1").Select
    Cells.Find(What:=cStartOfIndexTable, After:=ActiveCell, LookIn:= _
        xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:= _
        xlNext, MatchCase:=False, SearchFormat:=False).Activate
            
    iRow_IndexBeginning = ActiveCell.Row + 1
    'ActiveCell.Offset(2, 0).Range("A1").Select
    
    Cells.Find(What:=cEndOfIndexTable, After:=ActiveCell, LookIn:= _
        xlFormulas, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:= _
        xlNext, MatchCase:=False, SearchFormat:=False).Activate

    iRow_IndexEnd = ActiveCell.Row - 1
    
    ' Sanity check: size of the Index Table (too large a table might indicate failure to find the end of it)
    If iRow_IndexEnd - iRow_IndexBeginning > cMaxSizeOfIndexTable Then
        Select Case MsgBox("Size of the Index Table is " & iRow_IndexEnd - iRow_IndexBeginning & ". It is larger than the arbitrary threshold of " & cMaxSizeOfIndexTable & ".", vbAbortRetryIgnore, "Something fishy is going on here...")
            Case vbRetry
                MsgBox "Well I could certainly calculate the size again for you, but I assure you it would turn out to be the same.", , "Please do not waste my CPU cycles"
                Call PSR_Cleanup
                End
            Case vbIgnore
                MsgBox "As you wish. May the wrath be on your head!", , "Loving the risk aren't we?.."
            Case vbAbort
                Call PSR_Cleanup
                End
        End Select
    End If
    
    Exit Sub
Hell:
    Call PSR_Cleanup
    Call CrashDump("FindIndexTable")
End Sub


Private Function IsIndexSheetDataStructureKosher() As Boolean

    Dim i As Integer
    Dim sExpected As String
    Dim sGot As String

    For i = 1 To UBound(aDataStructure_IndexSheet)
    
        sExpected = aDataStructure_IndexSheet(i, 1)
        sGot = Cells(iRow_IndexBeginning, aDataStructure_IndexSheet(i, 2)).Value
    
        If sExpected <> sGot Then
            MsgBox "Blimey! THEY have changed the PSR format again! " & vbCrLf & vbCrLf & "Expected field: " & sExpected & vbCrLf & "Got: " & sGot, vbCritical, "Conspiracies everywhere"
            
            IsIndexSheetDataStructureKosher = False
            Exit Function
        End If
        
    Next i

    IsIndexSheetDataStructureKosher = True

End Function
