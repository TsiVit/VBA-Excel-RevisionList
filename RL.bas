Option Explicit
Dim SPTDL As Range          'Start Point To Do List
Dim SPRL As Range           'Start Point Revision List
Dim ClmnA As Byte           'Columns List Amount
Dim COff As UDT_ColumnOffset

Type UDT_ColumnOffset
    Version As Byte
    Changes As Byte
    Priority As Byte
    Date As Byte
    Deadline As Byte
    Due As Byte
End Type


Sub ini_Rev_List()
    Set SPTDL = Range("B6")     '<---------------------------
    
    ClmnA = 6                   '<---------------------------
    COff.Version = 0            '<---------------------------
    COff.Changes = 1            '<---------------------------
    COff.Priority = 2           '<---------------------------
    COff.Date = 3               '<---------------------------
    COff.Deadline = 4           '<---------------------------
    COff.Due = 5                '<---------------------------
    
    Set SPRL = SPTDL.EntireColumn.Find( _
                What:="Version", _
                After:=SPTDL).Offset(1, 0)
End Sub

Sub Rev_List()
    Dim StopMark As Boolean
    Dim DoneTsk As Range
    Dim NewRecRL As Range
    Dim VolMark As Byte         ' volume of transfered tasks
        
    If SheetNameCheck() = False Then Exit Sub
    Call ToggleEvents(False)
    VolMark = 0
   
    Do
        Call ini_Rev_List
        
        ' Search any Mark in TDL column (version ID) that means the task was implemented
        On Error GoTo ErrHandle
            Set DoneTsk = Range(SPTDL, SPRL.Offset(-4, 0)).Find("*").Resize(1, ClmnA)
        On Error GoTo 0
        If StopMark Then Exit Do
        
        ' Create new blank record row in Revision List
        Set NewRecRL = SPRL.Resize(1, ClmnA)
        
        ' Transfer task from TDL to Revision List
        Call New_Req(DoneTsk, NewRecRL)
        SPRL.Offset(-1, COff.Due).Value = Format(Date, "yyyy/mm/dd")
        
        ' Remove done task from TDL
        DoneTsk.EntireRow.Delete (xlUp)
        VolMark = VolMark + 1
    Loop
    
    If VolMark = 0 Then
        MsgBox "Done tasks MUST be marked by version ID to perform " & _
                "transition into Revision List"
    End If
    
    Set SPTDL = Nothing
    Set SPRL = Nothing
    Set DoneTsk = Nothing
    Set NewRecRL = Nothing

    Call ToggleEvents(True)
    Range("A1").Select
    
Exit Sub
ErrHandle:
    StopMark = True
    Resume Next
End Sub

Sub New_Req(ByVal RngFrom As Range, ByVal RngTo As Range)
' Insert a new top line with autospread the format condition rules
    
    ' Insert new line in work range
    RngTo.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow
    
    ' Copy Input row data
    RngFrom.Copy
    ' Paste Input data as formulas into top line
    RngTo.Offset(-1, 0).PasteSpecial _
            Paste:=xlPasteFormulas, _
            Operation:=xlNone, _
            SkipBlanks:=False, _
            Transpose:=False
    
    Application.CutCopyMode = False
    
    ' Clear values only
    RngFrom.SpecialCells(xlCellTypeConstants).ClearContents

    Application.CutCopyMode = False
End Sub

Sub TDL()
    Dim NewReqTDL As Range
    
    If SheetNameCheck() = False Then Exit Sub
    
    Call ini_Rev_List
    Call ToggleEvents(False)
    
    With SPTDL
        Set NewReqTDL = .Offset(-1, 0).Resize(1, ClmnA)
        
        If .Offset(-1, COff.Changes) = vbNullString Then
            MsgBox "Changes field in Request Row must be non-blank"
        Else
            ' Date
            If .Offset(-1, COff.Date) = vbNullString Then
                .Offset(-1, COff.Date).Value = Format(Date, "yyyy/mm/dd")
            End If
            ' Write New Request
            Call New_Req(NewReqTDL, NewReqTDL.Offset(1, 0))
        End If
        
        ' Clear Request Line
        .Offset(-2, 0).Resize(1, ClmnA).ClearContents
    End With
    
    Set NewReqTDL = Nothing
    Call ToggleEvents(True)

    Range("A1").Select
End Sub

Function SheetNameCheck() As Boolean
    ' If macro referenced throug hot keys then check sheet to avoid misprinting
    SheetNameCheck = True
    
    If ActiveSheet.Name <> "RL" Then
        MsgBox "Selected sheet is not appropriate for requested action" & Chr(10) & _
                "(The Macros check the sheet name to avoid missprinting hot keys) ", _
                vbCritical + vbOKOnly
        SheetNameCheck = False
    End If
    
End Function
