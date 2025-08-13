Option Strict Off
Option Explicit On
Option Compare Text

Imports System
Imports System.Reflection
Imports System.Reflection.Emit
Imports System.Runtime.InteropServices
Imports System.Text.RegularExpressions
Imports Microsoft.Office.Interop.Excel

Module Comparison
    Public Structure InputType3
        Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, ByRef lParam As VariantType) As Integer

        Public BOMSheet As String
        Public NewStdDwg As String
        '-------------------------------------------------------------Start Shipping list items.
        Public StrLineNo, StrLineNo2 As Integer
        Public OldShipListSht, NewShipListSht As Worksheet
        Public Shared HeaderType As String

        Public Shared Function ReadBOM(ByRef BomArray As Object, ByRef SheetToUse As Object) As Object
            Dim iA As Object                                                'Used to read contents of Ship List
            Dim FoundLast As Boolean                                        'Create Array with all information on the BOM
            Dim LineNo As Short, jA As Short, LineDel As Short
            Dim Test As String, Test2 As String
            '                                   Below is correct.
            Test = SheetToUse.Name              'Possible problem Found Sheet name equal to STD Items on first pass 
            FoundLast = False
            LineNo = 4
            LineDel = 0

            Do Until FoundLast = True           'This part of the program is looking for the last Item
                LineNo = LineNo + 1             'on the spreadsheet.
                Test = SheetToUse.Range("A" & LineNo).Value
                Test2 = SheetToUse.Range("A" & LineNo).Interior.ColorIndex
                '---------------Program was looking for color RED and then would quit looking for true total .
                'If SheetToUse.Range("A" & LineNo).Value = "" Or SheetToUse.Range("A" & LineNo).Interior.ColorIndex = 3 Then
                '    LineNo = LineNo - 1
                '    FoundLast = True
                'End If

                If SheetToUse.Range("A" & LineNo).Value = "" Then
                    LineNo = LineNo - 1
                    FoundLast = True
                End If
                If SheetToUse.Range("A" & LineNo).Interior.ColorIndex = 3 Then
                    LineDel = LineDel + 1
                End If
            Loop

            LineNo = LineNo - LineDel

            ReDim BomArray(13, LineNo - 4)

            For iA = 5 To LineNo
                For jA = 1 To 13
                    If SheetToUse.Range("A" & iA).Interior.ColorIndex <> 3 Then
                        Test = SheetToUse.Range(Chr(jA + 64) & iA).Value
                        BomArray(jA, iA - 4) = SheetToUse.Range(Chr(jA + 64) & iA).Value
                    Else
                        GoTo NextiA
                    End If
                Next jA
NextiA:
            Next iA

        End Function

        Public Shared Function ReadBulkBOM(ByRef BomArray As Object, ByRef SheetToUse As Object) As Object
            Dim iA As Object                                                'Used to read contents of ShippingList
            Dim FoundLast As Boolean                                        'Create Array with all information on the BOM
            Dim LineNo As Short, jA As Short
            Dim Test As String, Test2 As String

            FoundLast = False
            LineNo = 4

            Do Until FoundLast = True
                LineNo = LineNo + 1
                Test = SheetToUse.Range("A" & LineNo).Value                 'At this point SheetToUse = OldBlukBOM, but is set to each BOM Type.
                Test2 = SheetToUse.Range("A" & LineNo).Interior.ColorIndex
                If SheetToUse.Range("A" & LineNo).Value = "" Or SheetToUse.Range("A" & LineNo).Interior.ColorIndex = 3 Then
                    LineNo = LineNo - 1
                    FoundLast = True
                End If
            Loop

            ReDim BomArray(11, LineNo - 4)

            For iA = 5 To LineNo
                For jA = 1 To 11
                    Test = SheetToUse.Range(Chr(jA + 64) & iA).Value
                    BomArray(jA, iA - 4) = SheetToUse.Range(Chr(jA + 64) & iA).Value
                Next jA
            Next iA

            'MsgBox("Count is " & iA)

        End Function

        Public Shared Function ReadFindSTDs(ByRef BomArray As Object, ByRef SheetToUse As Object) As Object
            Dim iA As Object                        'Used to read contents of STDs BOM.
            Dim FoundLast As Boolean                'Create Array with all information on the STDs to Find.
            Dim LineNo As Short, jA As Short
            Dim Test As String

            FoundLast = False
            LineNo = 4

            Do Until FoundLast = True
                LineNo = LineNo + 1
                If SheetToUse.Range("A" & LineNo).Value = "" Then
                    LineNo = LineNo - 1
                    FoundLast = True
                End If
            Loop

            ReDim BomArray(14, LineNo - 4)              'ReDim BomArray(11, LineNo - 4)

            For iA = 5 To LineNo
                For jA = 1 To 13
                    Test = SheetToUse.Range(Chr(jA + 64) & iA).Value
                    BomArray(jA, iA - 4) = SheetToUse.Range(Chr(jA + 64) & iA).Value
                Next jA
            Next iA

        End Function

        ' Search predicate returns true if a string ends in "MX0503F".
        'Private Function EndsWithSaurus(ByVal s As String) As Boolean
        '    ' AndAlso prevents evaluation of the second Boolean expression if the string is so
        '    ' short that an error would occur.
        '    If (s.Length > 5) AndAlso
        '        (s.Substring(s.Length - 6).ToLower() = NewStdDwg) Then                   '(s.Substring(s.Length - 6).ToLower() = "saurus") Then
        '        Return True
        '    Else
        '        Return False
        '    End If

        'End Function

        Public Shared Function CompareArraysTank(ByRef NewArray As Object, ByRef OldArray As Object) As Object
            '--------DJL-------10-31-2023--------Yes this one is ok.
            '-------Compares two different Array's.
            Dim Excel, iB, jB As Object
            Dim kB As Short
            Dim ArrayCnt As Integer
            Dim FoundBrk, FoundTab, FoundNewLFd, FoundDash, FoundSpace, JBFound As Integer
            Dim TestN, TestO, JobNoCustNew, JobNoCustOld, DwgNoNew, DwgNoOld, RevNoNew, RevNoOld, ShpMkNew, ShpMkOld, DescNew, DescOld As String
            Dim QtyNew, QtyOld, FoundItemOnOldBOM, FoundNew, FoundOld, Test11, Test12, Test13, Test14, NewCust, FixCust As String
            Dim Test11First, Test11Second, Test12First, Test12Second, DashFound, NewInvNo, OldInvNo, NewQNo, OldQNo As String
            Dim NewStdNo, OldStdNo, NewMat, OldMat, NewWeight, OldWeight, DescNewFirst, DescNewSecond, DescOldFirst, DescOldSecond As String
            Dim Errno, PrgName, ErrMsg, ErrSource, FormString, ErrDll As String
            Dim PriPrg, ErrLastLineX, SearchException, JobNoCustOldReplace As String
            Dim ExceptPos, CntExcept, DWPos As Integer
            Dim Title, Msg, Style, Response As Object           '-------DJL-07-30-2025      'Added

            On Error GoTo Err_CompareArraysTank

            TestN = UBound(NewArray, 2)
            TestO = UBound(OldArray, 2)
            Test12 = ""
            FixCust = Nothing
            JBFound = 1                         'Yes
            ShippingList_Menu.Label2.Text = "Comparing New Shipping List to Old Shipping List"              '-------DJL-06-30-2025
            ShippingList_Menu.ProgressBar1.Maximum = UBound(NewArray, 2)

            For iB = 1 To UBound(NewArray, 2)                           'get dwgnumber of first item in newarray
                For jB = JBFound To UBound(OldArray, 2)   'or jB = 1 To UBound(OldArray, 2)       'compare to all dwg numbers in oldarray.
                    '--------Change All Test1 thru Test18 to True Variable.-------DJL-------10-31-2023

                    JobNoCustNew = NewArray(1, iB)         'get dwgnumber of first item in newarray.
                    JobNoCustOld = OldArray(1, jB)         'get dwgnumber of first item in oldarray.
                    DwgNoNew = NewArray(2, iB)         'get dwgnumber of first item in newarray.
                    DwgNoOld = OldArray(2, jB)
                    FoundItemOnOldBOM = OldArray(0, jB)       'FoundItemOnOldBOM = OldArray(12, jB)    'Test11 = OldArray(13, jB)      Make sure part has not already been found.

                    If InStr(DwgNoNew, "-DW-") Then     '-------DJL-06-30-2025      'New problem on Job Example:    5606-1107-210A-DW-08D01-01.dwg
                        DWPos = InStr(DwgNoNew, "-DW-")

                        DwgNoNew = Mid(DwgNoNew, DWPos + 4, Len(DwgNoNew))
                    End If

                    If InStr(DwgNoOld, "-DW-") Then     '-------DJL-07-30-2025      'New problem on Job Example:    5606-1107-210A-DW-08D01-01.dwg
                        DWPos = InStr(DwgNoOld, "-DW-")

                        DwgNoOld = Mid(DwgNoOld, DWPos + 4, Len(DwgNoOld))
                    End If

                    If InStr(JobNoCustNew, GenInfo3135.FullJobNo) = 0 Then '-------DJL-06-30-2025      'New problem on Job Example:    212-25-006_05A.DWG
                        JobNoCustNew = (GenInfo3135.FullJobNo & "/" & JobNoCustNew)
                    End If

                    '-------DJL-07-30-2025      'Old Spreadsheets may not have Job No\Customer info.
                    If JobNoCustNew <> JobNoCustOld And InStr(JobNoCustNew, JobNoCustOld) = 1 Then
                        If JobNoCustOldReplace = "" Then           '-------DJL-07-30-2025      'Added
                            '-------DJL-07-30-2025      'Added
                            Msg = ("Job Number\Customer Info on old spreadsheet does not match new spreadsheet is this correct? " & JobNoCustNew)
                            Style = MsgBoxStyle.YesNo
                            Title = "Found Job Number and Customer information does not match."
                            Response = MsgBox(Msg, Style, Title)

                            If Response = 6 Then
                                JobNoCustOldReplace = JobNoCustNew
                                JobNoCustOld = JobNoCustOldReplace
                            Else
                                JobNoCustOldReplace = InputBox("Please type in the correct information for your Job Number: Format = JobNo/Customer Info Example: 5608-1179/DOMINION ENERGY")
                            End If
                        Else
                            JobNoCustOld = JobNoCustOldReplace           '-------DJL-07-30-2025      'Added
                        End If
                    End If

                    If JobNoCustNew = "" And DwgNoNew = "" Then                         '-------DJL-08-12-2025      'Added
                        GoTo NextiB
                    End If

                    Select Case True
                        Case FoundItemOnOldBOM = "FOUND"                       'This Column Number has changed.
                            JBFound = jB
                            GoTo NextjB2
                        Case DwgNoNew = DwgNoOld   'Case NewArray(2, iB) = OldArray(2, jB)--------DJL--------10-31-2023'--'If JB Has been found Why return to begining.
                            '----------------------------------------if Reference dwg numbers match
                            RevNoNew = NewArray(3, iB)         'If revision number is new mark as revised.
                            '------------------------Found Problem "Nothing" = RevNoOld = OldArray(3, jB).
                            '------Problem must be in previous program "ReadShipList(OldBOM, OldShipListSht)".
                            RevNoOld = OldArray(3, jB)         'If NewArray(3, iB) = OldArray(3, jB) Then 

                            '---------------------Found new problem Dwg 12C had rev A on prev shiplist Rev 0 Job 2211-3909.
                            If RevNoNew <> RevNoOld Then      'Revision Number Changed.   'If NewArray(3, iB) <> OldArray(3, jB) Then
                                NewArray(0, iB) = "REVNO"      'NewArray(11, iB) = "REVNO"         'Added later in prg --  'HighlightLine(iC + StrLineNo, "Y", NewArray)
                            End If

                            '-------Found new ship mark do not match due to space add before ship Mark on new BOM-Shipping List.
                            ShpMkNew = NewArray(4, iB)                     'if ship marks match
                            ShpMkOld = OldArray(4, jB)

                            If ShpMkNew <> ShpMkOld Then
                                ShpMkNew = LTrim(ShpMkNew)
                                ShpMkNew = RTrim(ShpMkNew)
                            End If

                            If ShpMkOld <> ShpMkOld Then
                                ShpMkOld = LTrim(ShpMkOld)
                                ShpMkOld = RTrim(ShpMkOld)
                            End If

                            If ShpMkNew = ShpMkOld Then       'If NewArray(4, iB) = OldArray(4, jB) Then
                                If JobNoCustNew = JobNoCustOld Then
FixNames:
ContFind:                           DescNew = NewArray(8, iB)           '-------DJL-07-24-2025      'DescNew = NewArray(6, iB)
                                    DescOld = OldArray(6, jB)

                                    If DescNew <> DescOld Then             'if Description does not match remove extra spaces Etc...
                                        DescNew = LTrim(DescNew)
                                        DescNew = RTrim(DescNew)

                                        FoundTab = InStr(DescNew, Chr(9))      'Tab
                                        Select Case FoundTab
                                            Case 1
                                                DescNew = Mid(DescNew, 2, Len(DescNew))
                                            Case Is > 1
                                                MsgBox("This is a new problem, Please Create Ticket for Tab found on " & FoundItemOnOldBOM)            'MsgBox("This is a new problem, Please Create Ticket for Tab found on " & Test11)
                                        End Select

                                        FoundNewLFd = InStr(DescNew, Chr(10))     'New Line Feed
                                        Select Case FoundNewLFd
                                            Case 1
                                                FoundItemOnOldBOM = Mid(DescNew, 2, Len(DescNew))          'Test11 = Mid(DescNew, 2, Len(DescNew))
                                            Case Is > 1
                                                DescNewFirst = Mid(DescNew, 1, (FoundNewLFd - 1))
                                                DescNewSecond = Mid(DescNew, (FoundNewLFd + 1), Len(DescNew))
                                                DescNew = (DescNewFirst + DescNewSecond)
                                        End Select

                                        FoundDash = InStr(DescNew, Chr(45))     'Dash
                                        While FoundDash > 0
                                            DescNewFirst = Mid(DescNew, 1, (FoundDash - 1))
                                            DescNewSecond = Mid(DescNew, (FoundDash + 1), Len(DescNew))
                                            DescNew = DescNewFirst & DescNewSecond
                                            FoundDash = InStr(DescNew, Chr(45))
                                        End While

                                        DescNew = LTrim(DescNew)
                                        DescNew = RTrim(DescNew)
                                        FoundSpace = InStr(DescNew, Chr(32))         'Space
                                        While FoundSpace > 0
                                            DescNewFirst = Mid(DescNew, 1, (FoundSpace - 1))
                                            DescNewSecond = Mid(DescNew, (FoundSpace + 1), Len(DescNew))
                                            DescNew = DescNewFirst & DescNewSecond
                                            FoundSpace = InStr(DescNew, Chr(32))         'Space
                                        End While

                                        '---------------------------Test12 ----Fix
                                        DescOld = LTrim(DescOld)
                                        DescOld = RTrim(DescOld)

                                        FoundTab = InStr(DescOld, Chr(9))      'Tab
                                        Select Case FoundTab
                                            Case 1
                                                DescOld = Mid(DescOld, 2, Len(DescOld))
                                            Case Is > 1
                                                MsgBox("This is a new problem, Please Create Ticket for Tab found on " & Test12)
                                        End Select

                                        FoundNewLFd = InStr(DescOld, Chr(10))     'New Line Feed
                                        Select Case FoundNewLFd
                                            Case 1
                                                DescOld = Mid(DescOld, 2, Len(DescOld))
                                            Case Is > 1
                                                DescOldFirst = Mid(DescOld, 1, (FoundNewLFd - 1))
                                                DescOldSecond = Mid(DescOld, (FoundNewLFd + 1), Len(DescOld))
                                                DescOld = (DescOldFirst + DescOldSecond)
                                        End Select

                                        FoundDash = InStr(DescOld, Chr(45))     'Dash
                                        While FoundDash > 0
                                            DescOldFirst = Mid(DescOld, 1, (FoundDash - 1))
                                            DescOldSecond = Mid(DescOld, (FoundDash + 1), Len(DescOld))
                                            DescOld = DescOldFirst & DescOldSecond
                                            FoundDash = InStr(DescOld, Chr(45))
                                        End While

                                        DescOld = LTrim(DescOld)
                                        DescOld = RTrim(DescOld)
                                        FoundSpace = InStr(DescOld, Chr(32))         'Space
                                        While FoundSpace > 0
                                            DescOldFirst = Mid(DescOld, 1, (FoundSpace - 1))
                                            DescOldSecond = Mid(DescOld, (FoundSpace + 1), Len(DescOld))
                                            DescOld = DescOldFirst & DescOldSecond
                                            FoundSpace = InStr(DescOld, Chr(32))         'Space
                                        End While
                                    End If

                                    If DescNew = DescOld Then
                                        QtyNew = NewArray(6, iB)                            '-------DJL-07-24-2025      'QtyNew = NewArray(5, iB)
                                        QtyOld = OldArray(5, jB)
                                        Test13 = OldArray(11, jB)

                                        If QtyNew <> QtyOld Then
                                            QtyNew = LTrim(QtyNew)
                                            QtyNew = RTrim(QtyNew)
                                            If QtyNew <> QtyOld Then
                                                QtyOld = LTrim(QtyOld)
                                                QtyOld = RTrim(QtyOld)
                                            End If
                                        End If

                                        If QtyNew = QtyOld Then

                                            '----------------------------------------------------------------------------
                                            'Fix below for each like above--------DJL-------11-1-2023
                                            '7 = Std Part No -WB18
                                            '8 = Std No MX1701A
                                            '9 = Material
                                            '10 = Weight
                                            '11 = RevNo, or Found on Old        '-------Below does not need to look at Items 11 & 12 unless it was just a RevNo Only.
                                            '12 = Found, No Chg
                                            '-----------------------------------------------------------------------------------
                                            NewQNo = NewArray(7, iB)          'Maybe always blank
                                            OldQNo = OldArray(7, jB)

                                            If NewQNo <> OldQNo Then
                                                FixFormat(NewQNo, OldQNo)
                                                NewQNo = GenInfo3135.NewArrayItem
                                                OldQNo = GenInfo3135.OldArrayItem
                                            End If

                                            NewInvNo = NewArray(9, iB)          '-------DJL-07-24-2025      'NewInvNo = NewArray(8, iB)
                                            OldInvNo = OldArray(9, jB)          '-------DJL-07-24-2025      'OldInvNo = OldArray(8, jB)

                                            If NewInvNo <> OldInvNo Then
                                                FixFormat(NewInvNo, OldInvNo)
                                                NewInvNo = GenInfo3135.NewArrayItem
                                                OldInvNo = GenInfo3135.OldArrayItem
                                            End If

                                            NewStdNo = NewArray(10, iB)          '-------DJL-07-24-2025      'NewStdNo = NewArray(9, iB)
                                            OldStdNo = OldArray(10, jB)          '-------DJL-07-24-2025      'OldStdNo = OldArray(9, jB)

                                            If NewStdNo <> OldStdNo Then
                                                FixFormat(NewStdNo, OldStdNo)
                                                NewStdNo = GenInfo3135.NewArrayItem
                                                OldStdNo = GenInfo3135.OldArrayItem
                                            End If

                                            NewMat = NewArray(11, iB)           '-------DJL-07-24-2025      'NewMat = NewArray(10, iB)
                                            OldMat = OldArray(11, jB)           '-------DJL-07-24-2025      'OldMat = OldArray(10, jB)

                                            If NewMat <> OldMat Then
                                                FixFormat(NewMat, OldMat)
                                                NewMat = GenInfo3135.NewArrayItem
                                                OldMat = GenInfo3135.OldArrayItem
                                            End If

                                            If NewMat <> OldMat Then                '-------DJL-08-13-2025      'Added
                                                NewArray(0, iB) = "REVISED"
                                                OldArray(0, jB) = "FOUND"
                                                GoTo NextiB
                                            End If

                                            NewWeight = NewArray(14, iB)           '-------DJL-07-24-2025      'NewWeight = NewArray(11, iB)
                                            OldWeight = OldArray(12, jB)            '-------DJL-07-24-2025      'OldWeight = OldArray(11, jB)

                                            If NewWeight <> OldWeight Then          'And NewWeight <> vbLf Then
                                                FixFormat(NewWeight, OldWeight)
                                                NewWeight = GenInfo3135.NewArrayItem
                                                OldWeight = GenInfo3135.OldArrayItem
                                            End If

                                            If NewQNo <> OldQNo Or NewInvNo <> OldInvNo Then
                                                NewArray(0, iB) = "REVISED"
                                                OldArray(0, jB) = "FOUND"
                                                GoTo NextiB
                                            Else
                                                If NewStdNo <> OldStdNo Or NewMat <> OldMat Then
                                                    NewArray(0, iB) = "REVISED"
                                                    OldArray(0, jB) = "FOUND"
                                                    GoTo NextiB
                                                Else
                                                    If NewWeight <> OldWeight Then
                                                        NewArray(0, iB) = "Weight"                           'NewArray(0, iB) = "REVISED"
                                                        OldArray(0, jB) = "FOUND"
                                                        GoTo NextiB
                                                    Else
                                                        NewArray(0, iB) = "NO CHANGE"
                                                        OldArray(0, jB) = "FOUND"
                                                        GoTo NextiB
                                                    End If
                                                End If
                                            End If
                                        Else
                                            NewArray(0, iB) = "REVISED"
                                            OldArray(0, jB) = "FOUND"            '-------DJL-6-19-2024 'OldArray(UBound(OldArray, 0), jB) = "FOUND"
                                            GoTo NextiB
                                        End If
                                    Else
                                        NewArray(0, iB) = "REVISED"
                                        OldArray(0, jB) = "FOUND"                         '-------DJL-6-19-2024   'OldArray(UBound(OldArray, 0), jB) = "FOUND"
                                        GoTo NextiB
                                    End If

                                Else    '-------Look at issue for number 72-----Look for issue when prev ship list has dwg customer name changes.
                                    If JobNoCustOld <> JobNoCustNew Then

                                        If IsNothing(FixCust) = True Then
                                            NewCust = InputBox("This job has two customer names please type in the correct Customer Name ? " & JobNoCustNew & " or " & JobNoCustOld)
                                            NewCust = UCase(NewCust)
                                            FixCust = NewCust
                                        Else
                                            NewCust = FixCust             'If names do not match set to what user said was correct.
                                        End If

                                        GoTo FixNames
                                    End If

                                    If NewArray(4, iB) = vbNullString Then
                                        GoTo ContFind
                                    End If
                                    GoTo NextjB
                                End If

                            End If
                    End Select

NextjB:
                    If jB = UBound(OldArray, 2) And NewArray(0, UBound(NewArray, 2)) = vbNullString Then
                        NewArray(0, iB) = "NEW"
                    End If
NextjB2:
                Next jB

                If jB > UBound(OldArray, 2) And NewArray(0, iB) = vbNullString Then
                    NewArray(0, iB) = "NEW"
                End If
NextiB:
                ShippingList_Menu.ProgressBar1.Value = iB
            Next iB

Err_CompareArraysTank:
            Errno = Err.Number

            If Errno <> 0 Then
                PrgName = "CompareArraysTank"
                PriPrg = "ShipListReadBOMAutoCAD"
                ErrMsg = Err.Description
                ErrSource = Err.Source
                ErrDll = Err.LastDllError
                ErrLastLineX = Err.Erl
                'ErrException = Err.GetException

                Dim st As New StackTrace(Err.GetException, True)
                CntFrames = st.FrameCount
                GetFramesSrt = st.GetFrames
                PrgLineNo = GetFramesSrt(CntFrames - 1).ToString
                PrgLineNo = PrgLineNo.Replace("@", "at")
                LenPrgLineNo = (Len(PrgLineNo))
                PrgLineNo = Mid(PrgLineNo, 1, (LenPrgLineNo - 2))

                ShippingList_Menu.HandleErrSQL(PrgName, Errno, ErrMsg, ErrSource, PriPrg, ErrDll, DwgItem, PrgLineNo)
                If System.Environment.UserName = "dlong" Then       '-------DJL-06-30-2025      'If (ShippingList_Menu.UserNamex()) = "dlong" Then
                    MsgBox(ErrMsg)
                    Stop
                    Resume
                Else
                    ExceptPos = 0
                    SearchException = "Exception"
                    ExceptPos = InStr(ErrMsg, 1)
                    If ExceptPos > 0 Then
                        CntExcept = (CntExcept + 1)
                        If CntExcept < 6 Then
                            Resume
                        End If
                    End If
                End If
            End If

        End Function

        Public Shared Function CheckOldBOM(ByRef SheetToUse As Object) As Boolean
            Dim HeaderArray As Object
            Dim iD As Object

            HeaderArray = New Object() {"Ref", "DWG", "REV", "SHIP MARK", "PIECE MARK", "QTY", "DESCRIPTION", "INV-1", "INV-2", "MAT'L", "WEIGHT", "REQ TYPE", "PROD CODE"}

            For iD = 1 To 11                    'For iD = 1 To 12
                With SheetToUse
                    'Test = .Range(Chr(iD + 64) & "4").Value
                    If .Range(Chr(iD + 64) & "4").Value <> HeaderArray(iD - 1) Then
                        CheckOldBOM = False
                        Exit Function
                    End If
                End With
            Next iD

            CheckOldBOM = True

        End Function

        Public Shared Function FixFormat(ByRef NewArrayItem As String, ByRef OldArrayItem As String) As Object
            '--------DJL-------11-29-2023--------
            '-------Compares two different Array items to see what is different.
            Dim FoundBrk, FoundTab, FoundNewLFd, FoundDash, FoundSpace, JBFound As Integer
            Dim NewArrayItemFirst, NewArrayItemSecond, OldArrayItemFirst, OldArrayItemSecond, DashFound, NewInvNo, OldInvNo, NewQNo, OldQNo As String
            Dim NewStdNo, OldStdNo, NewMat, OldMat, NewWeight, OldWeight, DescNewFirst, DescNewSecond, DescOldFirst, DescOldSecond As String
            Dim Errno, PrgName, ErrMsg, ErrSource, FormString, ErrDll As String

            On Error GoTo Err_FixFormat

            'Fix below for each like above--------DJL-------11-1-2023
            '7 = Std Part No -WB18
            '8 = Std No MX1701A
            '9 = Material
            '10 = Weight
            '11 = RevNo, or Found on Old        '-------Below does not need to look at Items 11 & 12 unless it was just a RevNo Only.
            '12 = Found, No Chg

            If NewArrayItem <> OldArrayItem Then
                NewArrayItem = LTrim(NewArrayItem)
                NewArrayItem = RTrim(NewArrayItem)

                FoundTab = InStr(NewArrayItem, Chr(9))      'Tab

                Select Case FoundTab
                    Case 1
                        NewArrayItem = Mid(NewArrayItem, 2, Len(NewArrayItem))
                    Case Is > 1
                        MsgBox("This is a new problem, Please Create Ticket for Tab found on " & NewArrayItem)
                End Select

                FoundNewLFd = InStr(NewArrayItem, Chr(10))     'New Line Feed

                Select Case FoundNewLFd
                    Case 1
                        NewArrayItem = Mid(NewArrayItem, 2, Len(NewArrayItem))
                    Case Is > 1
                        NewArrayItemFirst = Mid(NewArrayItem, 1, (FoundNewLFd - 1))
                        NewArrayItemSecond = Mid(NewArrayItem, (FoundNewLFd + 1), Len(NewArrayItem))
                        NewArrayItem = (NewArrayItemFirst + NewArrayItemSecond)
                End Select

                FoundDash = InStr(NewArrayItem, Chr(45))     'Dash

                While FoundDash > 0
                    NewArrayItemFirst = Mid(NewArrayItem, 1, (FoundDash - 1))
                    NewArrayItemSecond = Mid(NewArrayItem, (FoundDash + 1), Len(NewArrayItem))
                    NewArrayItem = NewArrayItemFirst & NewArrayItemSecond
                    FoundDash = InStr(NewArrayItem, Chr(45))
                End While

                NewArrayItem = LTrim(NewArrayItem)
                NewArrayItem = RTrim(NewArrayItem)
                FoundSpace = InStr(NewArrayItem, Chr(32))         'Space

                While FoundSpace > 0
                    NewArrayItemFirst = Mid(NewArrayItem, 1, (FoundSpace - 1))
                    NewArrayItemSecond = Mid(NewArrayItem, (FoundSpace + 1), Len(NewArrayItem))
                    NewArrayItem = NewArrayItemFirst & NewArrayItemSecond
                    FoundSpace = InStr(NewArrayItem, Chr(32))         'Space
                End While

                '---------------------------OldArrayItem ----Fix
                OldArrayItem = LTrim(OldArrayItem)
                OldArrayItem = RTrim(OldArrayItem)

                FoundTab = InStr(OldArrayItem, Chr(9))      'Tab

                Select Case FoundTab
                    Case 1
                        OldArrayItem = Mid(OldArrayItem, 2, Len(OldArrayItem))
                    Case Is > 1
                        MsgBox("This is a new problem, Please Create Ticket for Tab found on " & OldArrayItem)
                End Select

                FoundNewLFd = InStr(OldArrayItem, Chr(10))     'New Line Feed
                Select Case FoundNewLFd
                    Case 1
                        OldArrayItem = Mid(OldArrayItem, 2, Len(OldArrayItem))
                    Case Is > 1
                        'MsgBox("This is a new problem, Please Create Ticket for Tab found on " & OldArrayItem)
                        OldArrayItemFirst = Mid(OldArrayItem, 1, (FoundNewLFd - 1))
                        OldArrayItemSecond = Mid(OldArrayItem, (FoundNewLFd + 1), Len(OldArrayItem))
                        OldArrayItem = (OldArrayItemFirst + OldArrayItemSecond)
                End Select

                FoundDash = InStr(OldArrayItem, Chr(45))     'Dash

                While FoundDash > 0
                    OldArrayItemFirst = Mid(OldArrayItem, 1, (FoundDash - 1))
                    OldArrayItemSecond = Mid(OldArrayItem, (FoundDash + 1), Len(OldArrayItem))
                    OldArrayItem = OldArrayItemFirst & OldArrayItemSecond
                    FoundDash = InStr(OldArrayItem, Chr(45))
                End While

                OldArrayItem = LTrim(OldArrayItem)
                OldArrayItem = RTrim(OldArrayItem)
                FoundSpace = InStr(OldArrayItem, Chr(32))         'Space

                While FoundSpace > 0
                    OldArrayItemFirst = Mid(OldArrayItem, 1, (FoundSpace - 1))
                    OldArrayItemSecond = Mid(OldArrayItem, (FoundSpace + 1), Len(OldArrayItem))
                    OldArrayItem = OldArrayItemFirst & OldArrayItemSecond
                    FoundSpace = InStr(OldArrayItem, Chr(32))         'Space
                End While
            End If

            GenInfo3135.NewArrayItem = NewArrayItem
            GenInfo3135.OldArrayItem = OldArrayItem

Err_FixFormat:
            Errno = Err.Number

            If Errno <> 0 Then
                ErrMsg = Err.Description
                MsgBox(ErrMsg)
                Stop
                Resume

            End If
        End Function

        'Function ReadShipList(ByRef ShipListArray As Object, ByVal SheetToUse As Object) As Object
        '    '--------------------------------------Used to read contents of Shipping List
        '    Dim iA, jA, LineNo As Integer
        '    Dim FoundLast As Boolean

        '    FoundLast = False
        '    If SheetToUse.Range("A" & "29").Value = "Job No: " Then
        '        LineNo = 30
        '        StrLineNo = 30
        '    ElseIf SheetToUse.Range("A" & "31").Value = "Job No: " Then
        '        LineNo = 32
        '        StrLineNo = 32
        '    ElseIf SheetToUse.Range("A" & "43").Value = "Job No: " Then
        '        LineNo = 44
        '        StrLineNo = 44
        '    End If

        '    Do Until FoundLast = True
        '        LineNo = LineNo + 1
        '        If SheetToUse.Range("B" & LineNo).Value = "" Or SheetToUse.Range("B" & LineNo).Interior.ColorIndex = 3 Then
        '            LineNo = LineNo - 1
        '            FoundLast = True
        '        End If
        '    Loop

        '    ReDim ShipListArray(11, LineNo - StrLineNo)

        '    For iA = StrLineNo + 1 To LineNo
        '        For jA = 1 To 10
        '            Select Case jA
        '                Case 1 To 6
        '                    If SheetToUse.Range(Chr(jA + 65) & iA).Interior.ColorIndex = 3 Then
        '                    Else
        '                        ShipListArray(jA, iA - StrLineNo) = SheetToUse.Range(Chr(jA + 65) & iA).Value
        '                    End If
        '                Case 7 To 10
        '                    If SheetToUse.Range(Chr(jA + 67) & iA).Interior.ColorIndex = 3 Then
        '                    Else
        '                        ShipListArray(jA, iA - StrLineNo) = SheetToUse.Range(Chr(jA + 67) & iA).Value
        '                    End If
        '            End Select
        '        Next jA
        '    Next iA
        'End Function

        '        Function CompareArraysTank2(ByVal NewArray As Object, ByVal OldArray As Object) As Object
        '            Dim iB, jB, kB As Integer

        '            For iB = 1 To UBound(NewArray, 2) 'get dwgnumber of first item in newarray
        '                For jB = 1 To UBound(OldArray, 2) 'compare to all dwg numbers in oldarray
        '                    If NewArray(2, iB) = OldArray(2, jB) Then 'if dwg numbers match
        '                        If NewArray(4, iB) = OldArray(4, jB) Then 'if ship marks match
        '                            If NewArray(6, iB) = OldArray(6, jB) Then 'if Description match
        '                                If NewArray(5, iB) = OldArray(5, jB) Then
        '                                    For kB = 5 To UBound(NewArray, 1) - 1 'check each remaining value for non matches
        '                                        Select Case kB
        '                                            Case 5 To UBound(NewArray, 1) - 2
        '                                                If NewArray(kB, iB) <> OldArray(kB, jB) Then
        '                                                    NewArray(UBound(NewArray, 1), iB) = "REVISED" 'revised from previous bom
        '                                                    OldArray(UBound(OldArray, 1), jB) = "FOUND" 'mark as found in oldarray
        '                                                    GoTo NextiB 'next item in newarray
        '                                                End If
        '                                            Case UBound(NewArray, 1) - 1
        '                                                If NewArray(kB, iB) <> OldArray(kB, jB) Then
        '                                                    NewArray(UBound(NewArray, 1), iB) = "REVISED" 'revised from previous bom
        '                                                    OldArray(UBound(OldArray, 1), jB) = "FOUND" 'mark as found in oldarray
        '                                                    GoTo NextiB 'next item in newarray
        '                                                Else
        '                                                    NewArray(UBound(NewArray, 1), iB) = "NO CHANGE" 'no change from previous bom
        '                                                    OldArray(UBound(OldArray, 1), jB) = "FOUND" 'mark as found in oldarray
        '                                                    GoTo NextiB 'next item in newarray
        '                                                End If
        '                                        End Select
        '                                    Next kB
        '                                Else
        '                                    NewArray(UBound(NewArray, 1), iB) = "REVISED"
        '                                    OldArray(UBound(OldArray, 1), jB) = "FOUND"
        '                                    GoTo NextiB
        '                                End If
        '                            End If
        '                        End If
        '                    End If
        '                    'if entire oldarray searched and no matches found, label "NEW"
        '                    If jB = UBound(OldArray, 2) And NewArray(UBound(NewArray, 1), iB) = vbNullString Then
        '                        NewArray(UBound(NewArray, 1), iB) = "NEW"
        '                    End If
        '                Next jB
        'NextiB:
        '            Next iB
        '            OldShipListSht.Activate()
        '            ExcelApp.Application.ActiveWorkbook.Close(False)
        '        End Function

        '        Function CompareArraysSeal2(ByVal NewArray As Object, ByVal OldArray As Object)
        '            Dim iB, jB, kB As Integer

        '            For iB = 1 To UBound(NewArray, 2) 'get dwgnumber of first item in newarray
        '                For jB = 1 To UBound(OldArray, 2) 'compare to all dwg numbers in oldarray
        '                    If NewArray(2, iB) = OldArray(2, jB) Then 'if dwg numbers match
        '                        If NewArray(4, iB) = OldArray(4, jB) Then 'if ship marks match
        '                            If NewArray(6, iB) = OldArray(6, jB) Then 'if descriptions match
        '                                For kB = 5 To UBound(NewArray, 1) - 1 'check each remaining value for non matches
        '                                    Select Case kB
        '                                        Case 5, 7 To UBound(NewArray, 1) - 2
        '                                            If NewArray(kB, iB) <> OldArray(kB, jB) Then
        '                                                NewArray(UBound(NewArray, 1), iB) = "REVISED" 'revised from previous bom
        '                                                OldArray(UBound(OldArray, 1), jB) = "FOUND" 'mark as found in oldarray
        '                                                GoTo NextiB 'next item in newarray
        '                                            End If
        '                                        Case 6
        '                                            'do nothing, description already checked
        '                                        Case UBound(NewArray, 1) - 1
        '                                            If NewArray(kB, iB) <> OldArray(kB, jB) Then
        '                                                NewArray(UBound(NewArray, 1), iB) = "REVISED" 'revised from previous bom
        '                                                OldArray(UBound(OldArray, 1), jB) = "FOUND" 'mark as found in oldarray
        '                                                GoTo NextiB 'next item in newarray
        '                                            Else
        '                                                NewArray(UBound(NewArray, 1), iB) = "NO CHANGE" 'no change from previous bom
        '                                                OldArray(UBound(OldArray, 1), jB) = "FOUND" 'mark as found in oldarray
        '                                                GoTo NextiB 'next item in newarray
        '                                            End If
        '                                    End Select
        '                                Next kB
        '                            End If
        '                        End If
        '                    End If
        '                    'if entire oldarray searched and no matches found, label "NEW"
        '                    If jB = UBound(OldArray, 2) And NewArray(UBound(NewArray, 1), iB) = vbNullString Then
        '                        NewArray(UBound(NewArray, 1), iB) = "NEW"
        '                    End If
        '                Next jB
        'NextiB:
        '            Next iB
        '            OldShipListSht.Activate()
        '            ExcelApp.Application.ActiveWorkbook.Close(False)

        '        End Function

        'Public Function FormatNewShipList(ByVal NewArray As Object, ByVal OldArray As Object, ByVal FileToFormat As Object)
        '    Dim iC, jC, LineNo As Integer
        '    Dim MultiLineMatl As Boolean

        '    FileToFormat.Activate()
        '    If StrLineNo = 42 Then
        '        StrLineNo = 44
        '    End If

        '    For iC = 1 To UBound(NewArray, 2)
        '        Select Case NewArray(UBound(NewArray, 1), iC)
        '            Case "REVISED"
        '                HighlightLine(iC + StrLineNo, "Y", NewArray)        'HighlightLine(iC + StrLineNo, "Y")
        '            Case "NEW"
        '                HighlightLine(iC + StrLineNo, "G", NewArray)        'HighlightLine(iC + StrLineNo, "G")
        '            Case "NO CHANGE"
        '                HighlightLine(iC + StrLineNo, "N", NewArray)        'HighlightLine(iC + StrLineNo, "N")
        '        End Select
        '    Next iC

        '    LineNo = UBound(NewArray, 2) + StrLineNo

        '    For iC = 1 To UBound(OldArray, 2)
        '        If OldArray(UBound(OldArray, 1), iC) = vbNullString Then
        '            MultiLineMatl = False
        '            LineNo = LineNo + 1
        '            For jC = 1 To UBound(OldArray, 1) - 1
        '                If jC > 6 Then
        '                    'Range(Chr(67 + jC) & LineNo).Value = OldArray(jC, iC) 'Changed for alignment
        '                Else
        '                    'Range(Chr(65 + jC) & LineNo).Value = OldArray(jC, iC) 'Changed for alignment
        '                End If
        '                Select Case jC
        '                    Case 9
        '                        If InStr(1, OldArray(jC, iC), Chr(10)) <> 0 Then
        '                            MultiLineMatl = True
        '                        End If
        '                End Select
        '            Next jC
        '            FormatLine(LineNo, MultiLineMatl)
        '            HighlightLine(LineNo, "R", BOMSheet)
        '        End If
        '    Next iC
        'End Function

        Public Shared Function CheckOldShipList(ByVal SheetToUse As Object) As Boolean
            Dim HeaderArray As Object
            Dim iD, StrLineNo2 As Integer               ', jD
            'Dim Test As String

            '-------------------------------Old version of Shipping list that does not have Column "LINE"
            HeaderArray = New Object() {"FAB ORDER NUMBER", "CUSTOMER PO NUMBER", "DWG", "REV", "SHIP", "QTY", "DESCRIPTION"}
            HeaderType = "OldHeader"

            If SheetToUse.Range("A" & "29").Value = "Job No: " Then     'Moved by Dennis J. Long StrLineNo2 only needs to be set once.
                StrLineNo2 = 30
            ElseIf SheetToUse.Range("A" & "31").Value = "Job No: " Then
                StrLineNo2 = 32
            ElseIf SheetToUse.Range("A" & "43").Value = "Job No: " Then
                StrLineNo2 = 44
            End If

            For iD = 1 To 7
                With SheetToUse

                    If .Range(Chr(iD + 64) & CStr(StrLineNo2)).Value <> HeaderArray(iD - 1) Then
                        'Test = .Range(Chr(iD + 64) & CStr(StrLineNo2)).Value
                        'Test = HeaderArray(iD - 1)
                        'CheckOldShipList = False
                        'Exit Function
                        GoTo ChkNewShipList
                    End If
                End With
            Next iD
            GoTo OldVersion

            '---------------------------------------------------------------------------------------------------
ChkNewShipList:
            HeaderArray = New Object() {"FAB ORDER NUMBER", "LINE", "CUSTOMER PO NUMBER", "DWG", "REV", "SHIP", "QTY", "DESCRIPTION"}
            HeaderType = "True"

            If SheetToUse.Range("A" & "29").Value = "Job No: " Then     'Moved by Dennis J. Long StrLineNo2 only needs to be set once.
                StrLineNo2 = 30
            ElseIf SheetToUse.Range("A" & "31").Value = "Job No: " Then
                StrLineNo2 = 32
            ElseIf SheetToUse.Range("A" & "43").Value = "Job No: " Then
                StrLineNo2 = 44
            End If

            For iD = 1 To 7
                With SheetToUse

                    If .Range(Chr(iD + 64) & CStr(StrLineNo2)).Value <> HeaderArray(iD - 1) Then
                        'Test = .Range(Chr(iD + 64) & CStr(StrLineNo2)).Value
                        'Test = HeaderArray(iD - 1)
                        CheckOldShipList = False
                        HeaderType = "False"
                        Exit Function
                    End If
                End With
            Next iD
            '---------------------------------------------------------------------------------------------------
            CheckOldShipList = True

OldVersion:

        End Function

    End Structure
End Module