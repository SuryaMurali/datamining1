

Function Last(choice As Long, rng As Range)
' 1 = last row
' 2 = last column
' 3 = last cell
    Dim lrw As Long
    Dim lcol As Long

    Select Case choice

    Case 1:
        On Error Resume Next
        Last = rng.Find(What:="*", _
                        After:=rng.Cells(1), _
                        Lookat:=xlPart, _
                        LookIn:=xlFormulas, _
                        SearchOrder:=xlByRows, _
                        SearchDirection:=xlPrevious, _
                        MatchCase:=False).Row
        On Error GoTo 0

    Case 2:
        On Error Resume Next
        Last = rng.Find(What:="*", _
                        After:=rng.Cells(1), _
                        Lookat:=xlPart, _
                        LookIn:=xlFormulas, _
                        SearchOrder:=xlByColumns, _
                        SearchDirection:=xlPrevious, _
                        MatchCase:=False).Column
        On Error GoTo 0

    Case 3:
        On Error Resume Next
        lrw = rng.Find(What:="*", _
                       After:=rng.Cells(1), _
                       Lookat:=xlPart, _
                       LookIn:=xlFormulas, _
                       SearchOrder:=xlByRows, _
                       SearchDirection:=xlPrevious, _
                       MatchCase:=False).Row
        On Error GoTo 0

        On Error Resume Next
        lcol = rng.Find(What:="*", _
                        After:=rng.Cells(1), _
                        Lookat:=xlPart, _
                        LookIn:=xlFormulas, _
                        SearchOrder:=xlByColumns, _
                        SearchDirection:=xlPrevious, _
                        MatchCase:=False).Column
        On Error GoTo 0

        On Error Resume Next
        Last = rng.Parent.Cells(lrw, lcol).Address(False, False)
        If Err.Number > 0 Then
            Last = rng.Cells(1).Address(False, False)
            Err.Clear
        End If
        On Error GoTo 0

    End Select
End Function

Sub Macro1()

Dim ALT As Workbook
Application.AskToUpdateLinks = False
Set ALT = Workbooks.Open("C:\Users\surya.murali\Desktop\Macros\ALT.xlsx")
Dim ALTs As Worksheet
Set ALTs = ALT.Sheets("Annual Layouts")
    Dim lastALT As Long
    lastALT = ALTs.UsedRange.SpecialCells(xlCellTypeLastCell).Row
    'Debug.Print lastALT
Dim SHAFT As Workbook
Set SHAFT = Workbooks.Open("C:\Users\surya.murali\Desktop\Macros\SHAFT.xls")
Dim SHAFTs As Worksheet
Set SHAFTs = SHAFT.Sheets("Plan")
    Dim lastSHAFT As Long
    lastSHAFT = SHAFTs.UsedRange.SpecialCells(xlCellTypeLastCell).Row
    SHAFTs.Activate
    SHAFTs.Range("C5:C" & lastSHAFT).Copy
    ALTs.Activate
    ALTs.Range("N2:N" & lastSHAFT).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone
    'Debug.Print lastSHAFT
    Application.CutCopyMode = False
    SHAFT.Close savechanges:=False
Dim TSA As Workbook
Set TSA = Workbooks.Open("C:\Users\surya.murali\Desktop\Macros\TSA.xls")
Dim TSAs As Worksheet
Set TSAs = TSA.Sheets("Tube Plan")
    Dim lastTSA As Long
    lastTSA = TSAs.UsedRange.SpecialCells(xlCellTypeLastCell).Row
    TSAs.Activate
    TSAs.Range("D6:D" & lastTSA).Copy
    ALTs.Activate
    Dim a As Long
    a = Last(1, ALTs.Range("N1:N100000"))
    b = a + lastTSA + 4
    ALTs.Range(ALTs.Cells(a + 2, 14), ALTs.Cells(b, 14)).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone
    'Debug.Print lastTSA
    Application.CutCopyMode = False
    TSA.Close savechanges:=False
Dim TULIP As Workbook
Set TULIP = Workbooks.Open("C:\Users\surya.murali\Desktop\Macros\TULIP.xls")
Dim TULIPs As Worksheet
Set TULIPs = TULIP.Sheets("Plan")
    Dim lastTULIP As Long
    lastTULIP = TULIPs.UsedRange.SpecialCells(xlCellTypeLastCell).Row
    'Debug.Print lastTULIP
    TULIPs.Activate
    TULIPs.Range("D6:D" & lastTULIP).Copy
    ALTs.Activate
    a = Last(1, ALTs.Range("N1:N100000"))
    b = a + lastTSA + 4
    ALTs.Range(ALTs.Cells(a + 2, 14), ALTs.Cells(b, 14)).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone
    Application.CutCopyMode = False
    TULIP.Close savechanges:=False
Dim FiOR As Workbook
Set FiOR = Workbooks.Open("C:\Users\surya.murali\Desktop\Macros\FiOR.xls")
Dim FiORs As Worksheet
Set FiORs = FiOR.Sheets("Plan")
    Dim lastFiOR As Long
    lastFiOR = FiORs.UsedRange.SpecialCells(xlCellTypeLastCell).Row
    'Debug.Print lastFiOR
    FiORs.Activate
    FiORs.Range("B6:B" & lastFiOR).Copy
    ALTs.Activate
    a = Last(1, ALTs.Range("N1:N100000"))
    b = a + lastTSA + 4
    ALTs.Range(ALTs.Cells(a + 2, 14), ALTs.Cells(b, 14)).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone
    Application.CutCopyMode = False
    FiOR.Close savechanges:=False
On Error GoTo Errorcatch

Dim counter As Integer, i As Integer
counter = 0
For i = 1 To Last(1, ALTs.Range("N1:N100000"))
    If Cells(i, 14).Value <> "" Then
        Cells(counter + 2, 13).Value = Cells(i, 14).Value
        counter = counter + 1
    End If
Next i
 Columns("N").EntireColumn.Delete
 
 For i = 2 To Last(1, ALTs.Range("A1:A100000"))
    For Z = 2 To Last(1, ALTs.Range("M1:M100000"))
        If InStr(1, Cells(i, 1).Value, Cells(Z, 13).Value, vbTextCompare) > 0 Then
            ALTs.Cells(i, 14).Value = "Yes"
        Else
        End If
    Next Z
Next i
ALTs.Cells(1, 13).Value = "Currently Running Parts"
ALTs.Cells(1, 14).Value = "Matching Parts"
 
'Assigning variables for the search algorithm
     
    Dim srchterm As String
    Dim srchdvalue As String
    
    
    Dim cellrow As Integer
    
    Dim layoutdate As Date
    Dim newlayoutdate As Date
    Dim todaydate As Date
    
    Dim intYear As Integer
    Dim intMonth As Integer
    Dim intDay As Integer
    Dim j As Integer
    
    
    
For j = 2 To lastALT
    If ALTs.Cells(j, 14).Value = "Yes" Then
        If IsEmpty(Range("D" & j).Value) = False Then
                layoutdate = ALTs.Range("D" & j).Value   'Calculate a new layout date that is actual layout date plus 350 days
                todaydate = Date
                intYear = Year(layoutdate)
                intMonth = Month(layoutdate)
                intDay = Day(layoutdate)
                newlayoutdate = DateSerial(intYear, intMonth, intDay + 350)
                        If todaydate - layoutdate >= 350 And todaydate - layoutdate < 365 And (ALTs.Range("F" & j).Value = "A" Or ALTs.Range("F" & j).Value = "a") Then '5th IF - If due date is coming up in two weeks or lesser, highlight whole row in yellow.
                        ALTs.Range("A" & j, "J" & j).Select
                        With Selection.Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .Color = 65535
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                        End With
                        Else
                        
                        If todaydate - layoutdate < 350 And (ALTs.Range("F" & j).Value = "A" Or ALTs.Range("F" & j).Value = "a") Then
                        ALTs.Range("A" & j, "J" & j).Select   'Else, highlight in green to state its alright
                        With Selection.Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .Color = 5296274
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                        End With
                        
                        Else
                        ALTs.Range("A" & j, "J" & j).Select 'highlight the whole line in red
                        With Selection.Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .Color = 255
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                        End With
                        End If
                End If '5th ENDIF
                Else   'If the layout date field is blank
                ALTs.Range("A" & j, "J" & j).Select 'highlight the whole line in red
                With Selection.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .Color = 255
                .TintAndShade = 0
                .PatternTintAndShade = 0
                End With
                End If
                End If
                
    Next j
    
         ALTs.Activate
         With ALTs

            .AutoFilterMode = False

            .Range("A1:N" & a).AutoFilter

            .Range("A1:N" & a).AutoFilter Field:=14, Criteria1:="<>"

    End With
         
                  
         
 Exit Sub
Errorcatch:
MsgBox "The following error occured:" & vbLf & "Error #: " & Err.Number & vbLf & "Description: " & Err.Description, _
vbCritical, "An Error Has Occured", Err.HelpFile, Err.HelpContext
End Sub

