Option Strict Off
Option Explicit On
Option Compare Text

Imports System
Imports System.Reflection
Imports System.Runtime.InteropServices
Imports Microsoft.Office.Interop.Excel

Module Formatting
    Public WorkShtName, PriPrg, ErrNo, ErrMsg, ErrSource, ErrDll, ErrLastLineX, PrgName As String
    Public ErrException As System.Exception
    Public ExcelApp As Object
    'Public BOMWrkSht As Worksheet
    Public Workbooks As Excel.Workbooks
    Public FileToOpen As String
    Public CntFrames As Integer
    Public GetFramesSrt
    Public PrgLineNo As String
    Public FindEndStr, LenPrgLineNo As Integer

    Function HighlightLine(ByRef LineNo As Short, ByRef Highlight As String, ByRef BOMSheet As String) As Object
        Dim xlColorIndexNone As Object
        Dim Range As Object
        Dim WorkBooks As Workbooks
        PrgName = "HighlightLine"

        On Error Resume Next

        ExcelApp = GetObject(, "Excel.Application")

        If Err.Number Then
            Information.Err.Clear()
            ExcelApp = CreateObject("Excel.Application")
            If Err.Number Then
                MsgBox(Err.Description)
                Exit Function
            End If
        End If

        On Error GoTo Err_HighlightLine

        WorkBooks = ExcelApp.WorkBooks

        NewBulkBOM = WorkBooks.Application.Worksheets(BOMSheet)
        NewBulkBOM.Activate()
        With NewBulkBOM
            With .Range("A" & LineNo & ":U" & LineNo).Interior
                Select Case Highlight
                    Case "Y"                                'color yellow fo revised
                        .ColorIndex = 6
                    Case "G"                                'color green for new
                        .ColorIndex = 4
                    Case "R"                                'color red for deleted
                        .ColorIndex = 3
                    Case "N"                                'no color for unchenged cells
                        'xlColorIndexNone = XlColorIndex.xlColorIndexNone
                        .ColorIndex = -4142                 'xlColorIndexNone 'xlNone
                End Select
            End With

            Select Case Highlight
                Case "G"                                            'color green for new
                    .Range("U" & LineNo).Value = "1"
                Case "Y"                                            'color yellow for revised
                    .Range("U" & LineNo).Value = "2"
                Case "R"                                            'color red for deleted
                    .Range("U" & LineNo).Value = "3"
                Case "N"                                            'no color for unchenged cells
                    .Range("U" & LineNo).Value = "8"
                Case "X"
                    .Range("U" & LineNo).Value = "7"
            End Select

        End With

FinishPrg:

Err_HighlightLine:
        ErrNo = Err.Number

        If ErrNo <> 0 Then
            PrgName = "HighlightLine"
            PriPrg = "ShipListReadBOMAutoCAD"
            ErrMsg = Err.Description
            ErrSource = Err.Source
            ErrDll = Err.LastDllError
            ErrLastLineX = Err.Erl
            ErrException = Err.GetException

            Dim st As New StackTrace(Err.GetException, True)
            CntFrames = st.FrameCount
            GetFramesSrt = st.GetFrames
            PrgLineNo = GetFramesSrt(CntFrames - 1).ToString
            PrgLineNo = PrgLineNo.Replace("@", "at")
            LenPrgLineNo = (Len(PrgLineNo))
            PrgLineNo = Mid(PrgLineNo, 1, (LenPrgLineNo - 2))

            ShippingList_Menu.HandleErrSQL(PrgName, ErrNo, ErrMsg, ErrSource, PriPrg, ErrDll, DwgItem, PrgLineNo)
            If ErrNo = "20" And ErrMsg = "Resume without error." Then
                Err.Clear()
                GoTo FinishPrg
            End If
            If GenInfo3135.UserName = "dlong" Then
                MsgBox(ErrMsg)
                Err.Clear()
                Stop
                Resume
            Else
                ExceptionPos = InStr(1, ErrMsg, "Exception")
                CallPos = InStr(1, ErrMsg, "Call was rejected by callee")
                CntExcept = (CntExcept + 1)

                If CntExcept < 20 Then
                    If ExceptionPos > 0 Then         'If program errors out due to use input, then repeat process.
                        Resume
                    End If
                    If CallPos > 0 Then         'If program errors out due to use input, then repeat process.
                        Resume
                    End If
                End If
            End If
        End If

    End Function

    Function HighlightRevNO(ByRef LineNo As Short, ByRef Highlight As String, ByRef BOMSheet As String) As Object
        Dim xlColorIndexNone As Object
        Dim Range As Object
        Dim TestColor, LookAtHeaderD, LookAtHeaderE As String
        Dim WorkBooks As Workbooks
        PrgName = "HighlightRevNO"

        On Error Resume Next

        ExcelApp = GetObject(, "Excel.Application")

        If Err.Number Then
            Information.Err.Clear()
            ExcelApp = CreateObject("Excel.Application")
            If Err.Number Then
                MsgBox(Err.Description)
                Exit Function
            End If
        End If

        On Error GoTo Err_HighlightRevNO

        WorkBooks = ExcelApp.WorkBooks

        NewBulkBOM = WorkBooks.Application.Worksheets(BOMSheet)
        NewBulkBOM.Activate()
        With NewBulkBOM
            LookAtHeaderD = .Range("D44").value
            LookAtHeaderE = .Range("E44").value

            If LookAtHeaderD = "REV" Then
                With .Range("D" & LineNo & ":D" & LineNo).Interior
                    Select Case Highlight
                        Case "Y"                                'color yellow fo revised
                            .ColorIndex = 6
                        Case "G"                                'color green for new
                            .ColorIndex = 4
                        Case "R"                                'color red for deleted
                            .ColorIndex = 3
                        Case "N"                                'no color for unchenged cells
                            .ColorIndex = -4142
                    End Select
                End With
            Else
                If LookAtHeaderE = "REV" Then
                    With .Range("E" & LineNo & ":E" & LineNo).Interior
                        Select Case Highlight
                            Case "Y"                                'color yellow fo revised
                                .ColorIndex = 6
                            Case "G"                                'color green for new
                                .ColorIndex = 4
                            Case "R"                                'color red for deleted
                                .ColorIndex = 3
                            Case "N"                                'no color for unchenged cells
                                .ColorIndex = -4142
                        End Select
                    End With
                Else
                    MsgBox("Spreadsheet Header Columns are in wrong order, Please contact IT with a Ticket.")
                End If
            End If

            TestColor = .Range("H" & LineNo & ":H" & LineNo).Interior.ColorIndex

            Select Case Highlight
                Case "G"                                            'color green for new
                    .Range("U" & LineNo).Value = "1"
                Case "Y"                                            'color yellow for revised
                    If TestColor = "-4142" And TestColor <> "6" Then
                        .Range("U" & LineNo).Value = "7"
                    End If
                Case "R"                                            'color red for deleted
                    .Range("U" & LineNo).Value = "3"
                Case "N"                                            'no color for unchenged cells
                    .Range("U" & LineNo).Value = "8"
                Case "X"
                    .Range("U" & LineNo).Value = "7"
            End Select
        End With

Err_HighlightRevNO:
        ErrNo = Err.Number

        If ErrNo <> 0 Then
            PrgName = "HighlightRevNo"
            PriPrg = "ShipListReadBOMAutoCAD"
            ErrMsg = Err.Description
            ErrSource = Err.Source
            ErrDll = Err.LastDllError
            ErrLastLineX = Err.Erl
            ErrException = Err.GetException

            Dim st As New StackTrace(Err.GetException, True)
            CntFrames = st.FrameCount
            GetFramesSrt = st.GetFrames
            PrgLineNo = GetFramesSrt(CntFrames - 1).ToString
            PrgLineNo = PrgLineNo.Replace("@", "at")
            LenPrgLineNo = (Len(PrgLineNo))
            PrgLineNo = Mid(PrgLineNo, 1, (LenPrgLineNo - 2))

            ShippingList_Menu.HandleErrSQL(PrgName, ErrNo, ErrMsg, ErrSource, PriPrg, ErrDll, DwgItem, PrgLineNo)
            If GenInfo3135.UserName = "dlong" Then
                MsgBox(ErrMsg)
                Err.Clear()
                Stop
                Resume
            Else
                ExceptionPos = InStr(1, ErrMsg, "Exception")
                CallPos = InStr(1, ErrMsg, "Call was rejected by callee")
                CntExcept = (CntExcept + 1)

                If CntExcept < 20 Then
                    If ExceptionPos > 0 Then         'If program errors out due to use input, then repeat process.
                        Resume
                    End If
                    If CallPos > 0 Then         'If program errors out due to use input, then repeat process.
                        Resume
                    End If
                End If
            End If
        End If

    End Function

    Function HighlightWeight(ByRef LineNo As Short, ByRef Highlight As String, ByRef BOMSheet As String) As Object
        Dim xlColorIndexNone As Object
        Dim Range As Object
        Dim TestDesc, TestColor, LookAtHeaderD As String
        Dim WorkBooks As Workbooks
        PrgName = "HighlightWeight"

        On Error Resume Next

        ExcelApp = GetObject(, "Excel.Application")

        If Err.Number Then
            Information.Err.Clear()
            ExcelApp = CreateObject("Excel.Application")
            If Err.Number Then
                MsgBox(Err.Description)
                Exit Function
            End If
        End If

        On Error GoTo Err_HighlightWeight

        WorkBooks = ExcelApp.WorkBooks
        NewBulkBOM = WorkBooks.Application.Worksheets(BOMSheet)
        NewBulkBOM.Activate()

        With NewBulkBOM
            LookAtHeaderD = .Range("N44").value

            If LookAtHeaderD = "WEIGHT" Then
                With .Range("N" & LineNo & ":N" & LineNo).Interior
                    Select Case Highlight
                        Case "Y"                                'color yellow fo revised
                            .ColorIndex = 6
                        Case "G"                                'color green for new
                            .ColorIndex = 4
                        Case "R"                                'color red for deleted
                            .ColorIndex = 3
                        Case "N"                                'no color for unchenged cells
                            .ColorIndex = -4142
                    End Select
                End With
            Else
                MsgBox("Spreadsheet Header Columns are in wrong order, Please contact IT with a Ticket.")
            End If

            TestColor = .Range("N" & LineNo & ":N" & LineNo).Interior.ColorIndex
            TestDesc = .Range("H" & LineNo & ":H" & LineNo).Interior.ColorIndex

            Select Case Highlight
                Case "G"                                            'color green for new
                    .Range("U" & LineNo).Value = "1"
                Case "Y"                                            'color yellow for revised
                    If TestDesc = "-4142" And TestColor = "6" Then
                        .Range("U" & LineNo).Value = "10"
                    End If
                Case "R"                                            'color red for deleted
                    .Range("U" & LineNo).Value = "3"
                Case "N"                                            'no color for unchenged cells
                    .Range("U" & LineNo).Value = "8"
                Case "X"
                    .Range("U" & LineNo).Value = "7"
            End Select
        End With

Err_HighlightWeight:
        ErrNo = Err.Number

        If ErrNo <> 0 Then
            PrgName = "HighLightWeight"
            PriPrg = "ShipListReadBOMAutoCAD"
            ErrMsg = Err.Description
            ErrSource = Err.Source
            ErrDll = Err.LastDllError
            ErrLastLineX = Err.Erl
            ErrException = Err.GetException

            Dim st As New StackTrace(Err.GetException, True)
            CntFrames = st.FrameCount
            GetFramesSrt = st.GetFrames
            PrgLineNo = GetFramesSrt(CntFrames - 1).ToString
            PrgLineNo = PrgLineNo.Replace("@", "at")
            LenPrgLineNo = (Len(PrgLineNo))
            PrgLineNo = Mid(PrgLineNo, 1, (LenPrgLineNo - 2))

            ShippingList_Menu.HandleErrSQL(PrgName, ErrNo, ErrMsg, ErrSource, PriPrg, ErrDll, DwgItem, PrgLineNo)
            If GenInfo3135.UserName = "dlong" Then
                MsgBox(ErrMsg)
                Err.Clear()
                Stop
                Resume
            Else
                ExceptionPos = InStr(1, ErrMsg, "Exception")
                CallPos = InStr(1, ErrMsg, "Call was rejected by callee")
                CntExcept = (CntExcept + 1)

                If CntExcept < 20 Then
                    If ExceptionPos > 0 Then
                        Resume
                    End If
                    If CallPos > 0 Then
                        Resume
                    End If
                End If
            End If
        End If

    End Function

    Function FormatLine(ByRef LineNo As Object, ByRef FileToOpen As String, Optional ByRef MultiLineMatl As Boolean = False) As Object
        Dim xlInsideHorizontal As Object, xlLeft As Object, xlInsideVertical As Object, xlEdgeRight As Object
        Dim xlEdgeBottom As Object, xlEdgeTop As Object, xlAutomatic As Object, xlThin As Object, xlEdgeLeft As Object
        Dim xlContinuous As Object, xlDiagonalUp As Object, xlDiagonalDown As Object, xlNone As Object, xlCenter As Object
        Dim Range As Object, xlDown As Object, Rows As Object, ExcelApp As Object
        Dim CallPos, ExceptionPos As Integer
        Dim WorkBooks As Workbooks
        Dim BOMWrkSht As Worksheet
        Dim WorkSht As Worksheet
        PrgName = "FormatLine"

        On Error Resume Next

        ExcelApp = GetObject(, "Excel.Application")

        If Err.Number Then
            Information.Err.Clear()
            ExcelApp = CreateObject("Excel.Application")
            If Err.Number Then
                MsgBox(Err.Description)
                Exit Function
            End If
        End If

        On Error GoTo Err_FormatLine

        WorkBooks = ExcelApp.Workbooks
        WorkShtName = "Shipping List"
        BOMWrkSht = WorkBooks.Application.Worksheets(WorkShtName)

        With BOMWrkSht
            .Rows(LineNo + 2 & ":" & LineNo + 2).Insert()
            .Rows(LineNo & ":" & LineNo).RowHeight = 18

            With .Range("A" & LineNo & ":S" & LineNo)
                .HorizontalAlignment = XlHAlign.xlHAlignCenter
                .VerticalAlignment = XlVAlign.xlVAlignCenter
                .Font.Name = "Arial"
                .Font.FontStyle = "Regular"
                .Font.Size = 9
                With .Borders.Item(XlBordersIndex.xlDiagonalDown)
                    .LineStyle = XlLineStyle.xlLineStyleNone
                End With
                With .Borders.Item(XlBordersIndex.xlDiagonalUp)
                    .LineStyle = XlLineStyle.xlLineStyleNone
                End With
                With .Borders.Item(XlBordersIndex.xlEdgeLeft)
                    .LineStyle = XlLineStyle.xlContinuous
                    .Weight = XlBorderWeight.xlThin
                    .ColorIndex = XlColorIndex.xlColorIndexAutomatic
                End With
                With .Borders.Item(XlBordersIndex.xlEdgeTop)
                    .LineStyle = XlLineStyle.xlContinuous
                    .Weight = XlBorderWeight.xlThin
                    .ColorIndex = XlColorIndex.xlColorIndexAutomatic
                End With
                With .Borders.Item(XlBordersIndex.xlEdgeBottom)
                    .LineStyle = XlLineStyle.xlContinuous
                    .Weight = XlBorderWeight.xlThin
                    .ColorIndex = XlColorIndex.xlColorIndexAutomatic
                End With
                With .Borders.Item(XlBordersIndex.xlEdgeRight)
                    .LineStyle = XlLineStyle.xlContinuous
                    .Weight = XlBorderWeight.xlThin
                    .ColorIndex = XlColorIndex.xlColorIndexAutomatic
                End With
                With .Borders.Item(XlBordersIndex.xlInsideVertical)
                    .LineStyle = XlLineStyle.xlContinuous
                    .Weight = XlBorderWeight.xlThin
                    .ColorIndex = XlColorIndex.xlColorIndexAutomatic
                End With
            End With

            With .Range("G" & LineNo & ":I" & LineNo)
                .HorizontalAlignment = XlHAlign.xlHAlignLeft
                .VerticalAlignment = XlVAlign.xlVAlignCenter

                With .Borders.Item(XlBordersIndex.xlDiagonalDown)
                    .LineStyle = XlLineStyle.xlLineStyleNone
                End With
                With .Borders.Item(XlBordersIndex.xlDiagonalUp)
                    .LineStyle = XlLineStyle.xlLineStyleNone
                End With
                With .Borders.Item(XlBordersIndex.xlEdgeLeft)
                    .LineStyle = XlLineStyle.xlContinuous
                    .Weight = XlBorderWeight.xlThin
                    .ColorIndex = XlColorIndex.xlColorIndexAutomatic
                End With
                With .Borders.Item(XlBordersIndex.xlEdgeTop)
                    .LineStyle = XlLineStyle.xlContinuous
                    .Weight = XlBorderWeight.xlThin
                    .ColorIndex = XlColorIndex.xlColorIndexAutomatic
                End With
                With .Borders.Item(XlBordersIndex.xlEdgeBottom)
                    .LineStyle = XlLineStyle.xlContinuous
                    .Weight = XlBorderWeight.xlThin
                    .ColorIndex = XlColorIndex.xlColorIndexAutomatic
                End With
                With .Borders.Item(XlBordersIndex.xlEdgeRight)
                    .LineStyle = XlLineStyle.xlContinuous
                    .Weight = XlBorderWeight.xlThin
                    .ColorIndex = XlColorIndex.xlColorIndexAutomatic
                End With
                With .Borders.Item(XlBordersIndex.xlInsideVertical)
                    .LineStyle = XlLineStyle.xlLineStyleNone
                End With
                With .Borders.Item(XlBordersIndex.xlInsideHorizontal)
                    .LineStyle = XlLineStyle.xlLineStyleNone
                End With
            End With

            With .Range("B" & LineNo)
                .HorizontalAlignment = XlHAlign.xlHAlignLeft
            End With

            If MultiLineMatl = True Then
                With .Range("L" & LineNo)
                    .HorizontalAlignment = XlHAlign.xlHAlignCenter
                    .VerticalAlignment = XlVAlign.xlVAlignCenter
                    .Font.Size = 7
                End With
            End If
        End With

Err_FormatLine:
        ErrNo = Err.Number

        If ErrNo <> 0 Then
            PrgName = "FormatLine"
            PriPrg = "ShipListReadBOMAutoCAD"
            ErrMsg = Err.Description
            ErrSource = Err.Source
            ErrDll = Err.LastDllError
            ErrLastLineX = Err.Erl
            ErrException = Err.GetException

            Dim st As New StackTrace(Err.GetException, True)
            CntFrames = st.FrameCount
            GetFramesSrt = st.GetFrames
            PrgLineNo = GetFramesSrt(CntFrames - 1).ToString
            PrgLineNo = PrgLineNo.Replace("@", "at")
            LenPrgLineNo = (Len(PrgLineNo))
            PrgLineNo = Mid(PrgLineNo, 1, (LenPrgLineNo - 2))

            ShippingList_Menu.HandleErrSQL(PrgName, ErrNo, ErrMsg, ErrSource, PriPrg, ErrDll, DwgItem, PrgLineNo)
            If GenInfo3135.UserName = "dlong" Then
                MsgBox(ErrMsg)
                Stop
                Resume
            Else
                ExceptionPos = InStr(1, ErrMsg, "Exception")
                CallPos = InStr(1, ErrMsg, "Call was rejected by callee")
                CntExcept = (CntExcept + 1)

                If CntExcept < 20 Then
                    If ExceptionPos > 0 Then
                        Resume
                    End If
                    If CallPos > 0 Then
                        Resume
                    End If
                End If
            End If
        End If

    End Function

    Function FormatLine2(ByRef LineNo As Object, ByRef FileToOpen As String, Optional ByRef MultiLineMatl As Boolean = False) As Object
        Dim xlInsideHorizontal As Object, xlLeft As Object, xlInsideVertical As Object, xlEdgeRight As Object
        Dim xlEdgeBottom As Object, xlEdgeTop As Object, xlAutomatic As Object, xlThin As Object, xlEdgeLeft As Object
        Dim xlContinuous As Object, xlDiagonalUp As Object, xlDiagonalDown As Object, xlNone As Object, xlCenter As Object
        Dim Range As Object, xlDown As Object, Rows As Object, ExcelApp As Object
        Dim CallPos, ExceptionPos As Integer
        Dim WorkBooks As Workbooks
        Dim BOMWrkSht As Worksheet
        Dim WorkSht As Worksheet
        PrgName = "FormatLine2"

        On Error Resume Next

        ExcelApp = GetObject(, "Excel.Application")

        If Err.Number Then
            Information.Err.Clear()
            ExcelApp = CreateObject("Excel.Application")
            If Err.Number Then
                MsgBox(Err.Description)
                Exit Function
            End If
        End If

        On Error GoTo Err_FormatLine2

        WorkBooks = ExcelApp.Workbooks
        WorkShtName = FileToOpen
        BOMWrkSht = WorkBooks.Application.Worksheets(WorkShtName)

        With BOMWrkSht
            .Rows(LineNo + 2 & ":" & LineNo + 2).Insert()
            .Rows(LineNo & ":" & LineNo).RowHeight = 18
            .Range("A" & LineNo).NumberFormat = "@"

            With .Range("A" & LineNo & ":M" & LineNo)
                .HorizontalAlignment = XlHAlign.xlHAlignCenter
                .VerticalAlignment = XlVAlign.xlVAlignCenter
                .Font.Name = "Arial"
                .Font.FontStyle = "Regular"
                .Font.Size = 9
                With .Borders.Item(XlBordersIndex.xlDiagonalDown)
                    .LineStyle = XlLineStyle.xlLineStyleNone
                End With
                With .Borders.Item(XlBordersIndex.xlDiagonalUp)
                    .LineStyle = XlLineStyle.xlLineStyleNone
                End With
                With .Borders.Item(XlBordersIndex.xlEdgeLeft)
                    .LineStyle = XlLineStyle.xlContinuous
                    .Weight = XlBorderWeight.xlThin
                    .ColorIndex = XlColorIndex.xlColorIndexAutomatic
                End With
                With .Borders.Item(XlBordersIndex.xlEdgeTop)
                    .LineStyle = XlLineStyle.xlContinuous
                    .Weight = XlBorderWeight.xlThin
                    .ColorIndex = XlColorIndex.xlColorIndexAutomatic
                End With
                With .Borders.Item(XlBordersIndex.xlEdgeBottom)
                    .LineStyle = XlLineStyle.xlContinuous
                    .Weight = XlBorderWeight.xlThin
                    .ColorIndex = XlColorIndex.xlColorIndexAutomatic
                End With
                With .Borders.Item(XlBordersIndex.xlEdgeRight)
                    .LineStyle = XlLineStyle.xlContinuous
                    .Weight = XlBorderWeight.xlThin
                    .ColorIndex = XlColorIndex.xlColorIndexAutomatic
                End With
                With .Borders.Item(XlBordersIndex.xlInsideVertical)
                    .LineStyle = XlLineStyle.xlContinuous
                    .Weight = XlBorderWeight.xlThin
                    .ColorIndex = XlColorIndex.xlColorIndexAutomatic
                End With
            End With

            With .Range("G" & LineNo)
                .HorizontalAlignment = XlHAlign.xlHAlignLeft
                .VerticalAlignment = XlVAlign.xlVAlignCenter

                With .Borders.Item(XlBordersIndex.xlDiagonalDown)
                    .LineStyle = XlLineStyle.xlLineStyleNone
                End With
                With .Borders.Item(XlBordersIndex.xlDiagonalUp)
                    .LineStyle = XlLineStyle.xlLineStyleNone
                End With
                With .Borders.Item(XlBordersIndex.xlEdgeLeft)
                    .LineStyle = XlLineStyle.xlContinuous
                    .Weight = XlBorderWeight.xlThin
                    .ColorIndex = XlColorIndex.xlColorIndexAutomatic
                End With
                With .Borders.Item(XlBordersIndex.xlEdgeTop)
                    .LineStyle = XlLineStyle.xlContinuous
                    .Weight = XlBorderWeight.xlThin
                    .ColorIndex = XlColorIndex.xlColorIndexAutomatic
                End With
                With .Borders.Item(XlBordersIndex.xlEdgeBottom)
                    .LineStyle = XlLineStyle.xlContinuous
                    .Weight = XlBorderWeight.xlThin
                    .ColorIndex = XlColorIndex.xlColorIndexAutomatic
                End With
                With .Borders.Item(XlBordersIndex.xlEdgeRight)
                    .LineStyle = XlLineStyle.xlContinuous
                    .Weight = XlBorderWeight.xlThin
                    .ColorIndex = XlColorIndex.xlColorIndexAutomatic
                End With
                With .Borders.Item(XlBordersIndex.xlInsideVertical)
                    .LineStyle = XlLineStyle.xlLineStyleNone
                End With
                With .Borders.Item(XlBordersIndex.xlInsideHorizontal)
                    .LineStyle = XlLineStyle.xlLineStyleNone
                End With
            End With

            If MultiLineMatl = True Then
                With .Range("I" & LineNo)
                    .HorizontalAlignment = XlHAlign.xlHAlignCenter
                    .VerticalAlignment = XlVAlign.xlVAlignCenter
                    .Font.Size = 7
                End With
            End If
        End With

Err_FormatLine2:
        ErrNo = Err.Number

        If ErrNo <> 0 Then
            PrgName = "FormatLine2"
            PriPrg = "ShipListReadBOMAutoCAD"
            ErrMsg = Err.Description
            ErrSource = Err.Source
            ErrDll = Err.LastDllError
            ErrLastLineX = Err.Erl
            ErrException = Err.GetException

            Dim st As New StackTrace(Err.GetException, True)
            CntFrames = st.FrameCount
            GetFramesSrt = st.GetFrames
            PrgLineNo = GetFramesSrt(CntFrames - 1).ToString
            PrgLineNo = PrgLineNo.Replace("@", "at")
            LenPrgLineNo = (Len(PrgLineNo))
            PrgLineNo = Mid(PrgLineNo, 1, (LenPrgLineNo - 2))

            ShippingList_Menu.HandleErrSQL(PrgName, ErrNo, ErrMsg, ErrSource, PriPrg, ErrDll, DwgItem, PrgLineNo)
            If GenInfo3135.UserName = "dlong" Then
                MsgBox(ErrMsg)
                Stop
                Resume
            Else
                ExceptionPos = InStr(1, ErrMsg, "Exception")
                CallPos = InStr(1, ErrMsg, "Call was rejected by callee")
                CntExcept = (CntExcept + 1)

                If CntExcept < 20 Then
                    If ExceptionPos > 0 Then
                        Resume
                    End If
                    If CallPos > 0 Then
                        Resume
                    End If
                End If
            End If
        End If

    End Function

    Function FormatLine3(ByRef LineNo As Object, ByRef FileToOpen As String, Optional ByRef MultiLineMatl As Boolean = False) As Object
        Dim xlInsideHorizontal As Object, xlLeft As Object, xlInsideVertical As Object, xlEdgeRight As Object
        Dim xlEdgeBottom As Object, xlEdgeTop As Object, xlAutomatic As Object, xlThin As Object, xlEdgeLeft As Object
        Dim xlContinuous As Object, xlDiagonalUp As Object, xlDiagonalDown As Object, xlNone As Object, xlCenter As Object
        Dim Range As Object, xlDown As Object, Rows As Object, ExcelApp As Object
        Dim CallPos, ExceptionPos As Integer
        Dim WorkBooks As Workbooks
        Dim BOMWrkSht As Worksheet
        Dim WorkSht As Worksheet
        PrgName = "FormatLine3"

        On Error Resume Next

        ExcelApp = GetObject(, "Excel.Application")

        If Err.Number Then
            Information.Err.Clear()
            ExcelApp = CreateObject("Excel.Application")
            If Err.Number Then
                MsgBox(Err.Description)
                Exit Function
            End If
        End If

        On Error GoTo Err_FormatLine3

        WorkBooks = ExcelApp.Workbooks

        If FileToOpen = "Purchase BOM" Then
            WorkShtName = "Other BOM"
        Else
            WorkShtName = FileToOpen
        End If

        BOMWrkSht = WorkBooks.Application.Worksheets(WorkShtName)

        With BOMWrkSht
            .Rows(LineNo & ":" & LineNo).Insert()
            .Rows(LineNo & ":" & LineNo).RowHeight = 18

            With .Range("A" & LineNo & ":L" & LineNo)
                .HorizontalAlignment = XlHAlign.xlHAlignCenter
                .VerticalAlignment = XlVAlign.xlVAlignCenter
                .Font.Name = "Arial"
                .Font.FontStyle = "Regular"
                .Font.Size = 9
                With .Borders.Item(XlBordersIndex.xlDiagonalDown)
                    .LineStyle = XlLineStyle.xlLineStyleNone
                End With
                With .Borders.Item(XlBordersIndex.xlDiagonalUp)
                    .LineStyle = XlLineStyle.xlLineStyleNone
                End With
                With .Borders.Item(XlBordersIndex.xlEdgeLeft)
                    .LineStyle = XlLineStyle.xlContinuous
                    .Weight = XlBorderWeight.xlThin
                    .ColorIndex = XlColorIndex.xlColorIndexAutomatic
                End With
                With .Borders.Item(XlBordersIndex.xlEdgeTop)
                    .LineStyle = XlLineStyle.xlContinuous
                    .Weight = XlBorderWeight.xlThin
                    .ColorIndex = XlColorIndex.xlColorIndexAutomatic
                End With
                With .Borders.Item(XlBordersIndex.xlEdgeBottom)
                    .LineStyle = XlLineStyle.xlContinuous
                    .Weight = XlBorderWeight.xlThin
                    .ColorIndex = XlColorIndex.xlColorIndexAutomatic
                End With
                With .Borders.Item(XlBordersIndex.xlEdgeRight)
                    .LineStyle = XlLineStyle.xlContinuous
                    .Weight = XlBorderWeight.xlThin
                    .ColorIndex = XlColorIndex.xlColorIndexAutomatic
                End With
                With .Borders.Item(XlBordersIndex.xlInsideVertical)
                    .LineStyle = XlLineStyle.xlContinuous
                    .Weight = XlBorderWeight.xlThin
                    .ColorIndex = XlColorIndex.xlColorIndexAutomatic
                End With
            End With

            With .Range("F" & LineNo)
                .HorizontalAlignment = XlHAlign.xlHAlignLeft
                .VerticalAlignment = XlVAlign.xlVAlignCenter

                With .Borders.Item(XlBordersIndex.xlDiagonalDown)
                    .LineStyle = XlLineStyle.xlLineStyleNone
                End With
                With .Borders.Item(XlBordersIndex.xlDiagonalUp)
                    .LineStyle = XlLineStyle.xlLineStyleNone
                End With
                With .Borders.Item(XlBordersIndex.xlEdgeLeft)
                    .LineStyle = XlLineStyle.xlContinuous
                    .Weight = XlBorderWeight.xlThin
                    .ColorIndex = XlColorIndex.xlColorIndexAutomatic
                End With
                With .Borders.Item(XlBordersIndex.xlEdgeTop)
                    .LineStyle = XlLineStyle.xlContinuous
                    .Weight = XlBorderWeight.xlThin
                    .ColorIndex = XlColorIndex.xlColorIndexAutomatic
                End With
                With .Borders.Item(XlBordersIndex.xlEdgeBottom)
                    .LineStyle = XlLineStyle.xlContinuous
                    .Weight = XlBorderWeight.xlThin
                    .ColorIndex = XlColorIndex.xlColorIndexAutomatic
                End With
                With .Borders.Item(XlBordersIndex.xlEdgeRight)
                    .LineStyle = XlLineStyle.xlContinuous
                    .Weight = XlBorderWeight.xlThin
                    .ColorIndex = XlColorIndex.xlColorIndexAutomatic
                End With
                With .Borders.Item(XlBordersIndex.xlInsideVertical)
                    .LineStyle = XlLineStyle.xlLineStyleNone
                End With
                With .Borders.Item(XlBordersIndex.xlInsideHorizontal)
                    .LineStyle = XlLineStyle.xlLineStyleNone
                End With
            End With

            If MultiLineMatl = True Then
                With .Range("I" & LineNo)
                    .HorizontalAlignment = XlHAlign.xlHAlignCenter
                    .VerticalAlignment = XlVAlign.xlVAlignCenter
                    .Font.Size = 7
                End With
            End If
        End With

Err_FormatLine3:
        ErrNo = Err.Number

        If ErrNo <> 0 Then
            PrgName = "FormatLine3"
            PriPrg = "ShipListReadBOMAutoCAD"
            ErrMsg = Err.Description
            ErrSource = Err.Source
            ErrDll = Err.LastDllError
            ErrLastLineX = Err.Erl
            ErrException = Err.GetException

            Dim st As New StackTrace(Err.GetException, True)
            CntFrames = st.FrameCount
            GetFramesSrt = st.GetFrames
            PrgLineNo = GetFramesSrt(CntFrames - 1).ToString
            PrgLineNo = PrgLineNo.Replace("@", "at")
            LenPrgLineNo = (Len(PrgLineNo))
            PrgLineNo = Mid(PrgLineNo, 1, (LenPrgLineNo - 2))

            ShippingList_Menu.HandleErrSQL(PrgName, ErrNo, ErrMsg, ErrSource, PriPrg, ErrDll, DwgItem, PrgLineNo)
            If GenInfo3135.UserName = "dlong" Then
                MsgBox(ErrMsg)
                Stop
                Resume
            Else
                ExceptionPos = InStr(1, ErrMsg, "Exception")
                CallPos = InStr(1, ErrMsg, "Call was rejected by callee")
                CntExcept = (CntExcept + 1)

                If CntExcept < 20 Then
                    If ExceptionPos > 0 Then
                        Resume
                    End If
                    If CallPos > 0 Then
                        Resume
                    End If
                End If
            End If
        End If

    End Function

    '    Function FormatLineShipList(ByVal LineNo As VariantType, Optional ByVal MultiLineMatl As Boolean = False) As Object
    '        PrgName = "FormatLine3"

    '        On Error Resume Next

    '        ExcelApp = GetObject(, "Excel.Application")

    '        If Err.Number Then
    '            Information.Err.Clear()
    '            ExcelApp = CreateObject("Excel.Application")
    '            If Err.Number Then
    '                MsgBox(Err.Description)
    '                Exit Function
    '            End If
    '        End If

    '        On Error GoTo Err_FormatLineShipList

    '        Workbooks = ExcelApp.Workbooks

    '        If FileToOpen = "Purchase BOM" Then
    '            WorkShtName = "Other BOM"
    '        Else
    '            WorkShtName = FileToOpen
    '        End If

    '        BOMWrkSht = Workbooks.Application.Worksheets(WorkShtName)

    '        With BOMWrkSht
    '            .Rows(LineNo + 2 & ":" & LineNo + 2).Insert()
    '            .Rows(LineNo & ":" & LineNo).RowHeight = 18
    '            With .Range("A" & LineNo & ":U" & LineNo)
    '                .HorizontalAlignment = XlHAlign.xlHAlignCenter
    '                .VerticalAlignment = XlVAlign.xlVAlignCenter
    '                .Font.Name = "Arial"
    '                .Font.FontStyle = "Regular"
    '                .Font.Size = 9
    '                With .Borders.Item(XlBordersIndex.xlDiagonalDown)
    '                    .LineStyle = XlLineStyle.xlLineStyleNone
    '                End With
    '                With .Borders.Item(XlBordersIndex.xlDiagonalUp)
    '                    .LineStyle = XlLineStyle.xlLineStyleNone
    '                End With
    '                With .Borders.Item(XlBordersIndex.xlEdgeLeft)
    '                    .LineStyle = XlLineStyle.xlContinuous
    '                    .Weight = XlBorderWeight.xlThin
    '                    .ColorIndex = XlColorIndex.xlColorIndexAutomatic
    '                End With
    '                With .Borders.Item(XlBordersIndex.xlEdgeTop)
    '                    .LineStyle = XlLineStyle.xlContinuous
    '                    .Weight = XlBorderWeight.xlThin
    '                    .ColorIndex = XlColorIndex.xlColorIndexAutomatic
    '                End With
    '                With .Borders.Item(XlBordersIndex.xlEdgeBottom)
    '                    .LineStyle = XlLineStyle.xlContinuous
    '                    .Weight = XlBorderWeight.xlThin
    '                    .ColorIndex = XlColorIndex.xlColorIndexAutomatic
    '                End With
    '                With .Borders.Item(XlBordersIndex.xlEdgeRight)
    '                    .LineStyle = XlLineStyle.xlContinuous
    '                    .Weight = XlBorderWeight.xlThin
    '                    .ColorIndex = XlColorIndex.xlColorIndexAutomatic
    '                End With
    '                With .Borders.Item(XlBordersIndex.xlInsideVertical)
    '                    .LineStyle = XlLineStyle.xlContinuous
    '                    .Weight = XlBorderWeight.xlThin
    '                    .ColorIndex = XlColorIndex.xlColorIndexAutomatic
    '                End With
    '            End With

    '            With .Range("G" & LineNo & ":I" & LineNo)
    '                .HorizontalAlignment = XlHAlign.xlHAlignLeft
    '                .VerticalAlignment = XlVAlign.xlVAlignCenter

    '                With .Borders.Item(XlBordersIndex.xlDiagonalDown)
    '                    .LineStyle = XlLineStyle.xlLineStyleNone
    '                End With
    '                With .Borders.Item(XlBordersIndex.xlDiagonalUp)
    '                    .LineStyle = XlLineStyle.xlLineStyleNone
    '                End With
    '                With .Borders.Item(XlBordersIndex.xlEdgeLeft)
    '                    .LineStyle = XlLineStyle.xlContinuous
    '                    .Weight = XlBorderWeight.xlThin
    '                    .ColorIndex = XlColorIndex.xlColorIndexAutomatic
    '                End With
    '                With .Borders.Item(XlBordersIndex.xlEdgeTop)
    '                    .LineStyle = XlLineStyle.xlContinuous
    '                    .Weight = XlBorderWeight.xlThin
    '                    .ColorIndex = XlColorIndex.xlColorIndexAutomatic
    '                End With
    '                With .Borders.Item(XlBordersIndex.xlEdgeBottom)
    '                    .LineStyle = XlLineStyle.xlContinuous
    '                    .Weight = XlBorderWeight.xlThin
    '                    .ColorIndex = XlColorIndex.xlColorIndexAutomatic
    '                End With
    '                With .Borders.Item(XlBordersIndex.xlEdgeRight)
    '                    .LineStyle = XlLineStyle.xlContinuous
    '                    .Weight = XlBorderWeight.xlThin
    '                    .ColorIndex = XlColorIndex.xlColorIndexAutomatic
    '                End With
    '                With .Borders.Item(XlBordersIndex.xlInsideVertical)
    '                    .LineStyle = XlLineStyle.xlLineStyleNone
    '                End With
    '                With .Borders.Item(XlBordersIndex.xlInsideHorizontal)
    '                    .LineStyle = XlLineStyle.xlLineStyleNone
    '                End With
    '            End With
    '            With .Range("B" & LineNo)
    '                .HorizontalAlignment = XlHAlign.xlHAlignLeft
    '            End With
    '            If MultiLineMatl = True Then
    '                With .Range("L" & LineNo)
    '                    .HorizontalAlignment = XlHAlign.xlHAlignCenter
    '                    .VerticalAlignment = XlVAlign.xlVAlignCenter
    '                    .Font.Size = 7
    '                End With
    '            End If
    '        End With

    'Err_FormatLineShipList:
    '        ErrNo = Err.Number

    '        If ErrNo <> 0 Then
    '            PrgName = "FormatLineShipList"
    '            PriPrg = "ShipListReadBOMAutoCAD"
    '            ErrMsg = Err.Description
    '            ErrSource = Err.Source
    '            ErrDll = Err.LastDllError
    '            ErrLastLineX = Err.Erl
    '            ErrException = Err.GetException

    '            Dim st As New StackTrace(Err.GetException, True)
    '            CntFrames = st.FrameCount
    '            GetFramesSrt = st.GetFrames
    '            PrgLineNo = GetFramesSrt(CntFrames - 1).ToString
    '            PrgLineNo = PrgLineNo.Replace("@", "at")
    '            LenPrgLineNo = (Len(PrgLineNo))
    '            PrgLineNo = Mid(PrgLineNo, 1, (LenPrgLineNo - 2))

    '            ShippingList_Menu.HandleErrSQL(PrgName, ErrNo, ErrMsg, ErrSource, PriPrg, ErrDll, DwgItem, PrgLineNo)
    '            If GenInfo3135.UserName = "dlong" Then
    '                MsgBox(ErrMsg)
    '                Stop
    '                Resume
    '            Else
    '                ExceptionPos = InStr(1, ErrMsg, "Exception")
    '                CallPos = InStr(1, ErrMsg, "Call was rejected by callee")
    '                CntExcept = (CntExcept + 1)

    '                If CntExcept < 20 Then
    '                    If ExceptionPos > 0 Then
    '                        Resume
    '                    End If
    '                    If CallPos > 0 Then
    '                        Resume
    '                    End If
    '                End If
    '            End If
    '        End If

    '    End Function

    '    Function FormatBulkBOM(ByRef BOMList As Object) As Object
    '        Dim Range As Object, Worksheets As Object, i As Object
    '        Dim LinesToDelete As Object, DupShipMk As Object, AssyMK As Object, AssyMKtmp As Object
    '        Dim RefDwgLen As Short, j As Short
    '        Dim RefDwgLetter, RefMk, RefDwg, ShpMark As String
    '        Dim BOMWrkSht As Worksheet
    '        Dim WorkSht As Worksheet
    '        'Dim ExcelApp As Excel.Workbooks
    '        Dim Workbooks As Excel.Workbooks
    '        Dim ExceptionPos As Integer
    '        Dim CallPos As Integer

    '        PrgName = "FormatBulkBOM"
    '        AssyMK = Nothing

    '        On Error Resume Next

    '        ExcelApp = GetObject(, "Excel.Application")

    '        If Err.Number Then
    '            Information.Err.Clear()
    '            ExcelApp = CreateObject("Excel.Application")
    '            If Err.Number Then
    '                MsgBox(Err.Description)
    '                Exit Function
    '            End If
    '        End If

    '        On Error GoTo Err_FormatBulkBOM

    '        Workbooks = ExcelApp.Workbooks
    '        WorkShtName = "Bulk BOM"
    '        BOMWrkSht = Workbooks.Application.Worksheets(WorkShtName)
    '        WorkSht = Workbooks.Application.ActiveSheet
    '        WorkShtName = WorkSht.Name

    '        ReDim LinesToDelete(1)
    '        ReDim DupShipMk(1)

    '        For i = 4 To (UBound(BOMList, 2) + 4)       'Remove any cells that contain only a space
    '            For j = 1 To 10
    '                With BOMWrkSht
    '                    Dim Test
    '                    Test = (Chr(64 + j) & i)
    '                    With .Range(Test)
    '                        If .Value <> vbNullString Then
    '                            If .Value.ToString = " " Then
    '                                .Value = vbNullString
    '                            End If
    '                        End If
    '                    End With
    '                End With
    '            Next j
    '        Next i

    '        '----------------------------------Remove headers for parts listed on sheet not requiring STD Lookup.
    '        Dim FCODwgNo As String
    '        For i = 5 To (UBound(BOMList, 2) + 4)
    '            With BOMWrkSht                                              'With Worksheets("Bulk BOM")
    '                If .Range("C" & i).Value <> "" Then                     'if ship mark is not empty
    '                    If .Range("G" & i).Value = "" Then                  'if inventory number is empty
    '                        '---------------------------------------special code for use with flush cleanouts
    '                        If InStr(1, .Range("F" & i).Value, "FLUSH CLEAN") <> 0 Then

    '                            MarkForDelete(LinesToDelete, i)
    '                            AssyMK = .Range("C" & i).Value
    '                            FCODwgNo = Left(.Range("A" & i).Value, 2)

    '                            Do Until Left(.Range("A" & i).Value, 2) <> FCODwgNo
    '                                If .Range("C" & i).Value = "" And .Range("D" & i).Value = "" And .Range("E" & i).Value = "" Then
    '                                    MarkForDelete(LinesToDelete, i)
    '                                ElseIf .Range("C" & i).Value <> "" And .Range("D" & i).Value = "" And .Range("G" & i).Value = "" Then

    '                                    MarkForDelete(LinesToDelete, i)
    '                                    AssyMKtmp = .Range("C" & i).Value
    '                                    i = i + 1
    '                                    If .Range("C" & i).Value = "" And Left(.Range("A" & i).Value, 2) = FCODwgNo Then
    '                                        .Range("C" & i).Value = AssyMKtmp
    '                                    End If
    '                                Else
    '                                    If .Range("C" & i).Value = "" Then
    '                                        .Range("C" & i).Value = AssyMK
    '                                    End If
    '                                End If
    '                                i = i + 1
    '                            Loop
    '                            i = i - 1

    '                        ElseIf .Range("C" & i).Value <> Nothing And .Range("D" & i).Value <> Nothing And .Range("E" & i).Value.ToString <> Nothing Then
    '                        Else                                            'section for typical material callouts.
    '                            If .Range("C" & i).Value <> "" And .Range("C" & i + 1).Value <> "" Then
    '                                If .Range("G" & i).Value = "" And .Range("H" & i).Value = "" Then
    '                                    MarkForDelete(LinesToDelete, i)
    '                                End If
    '                            Else
    '                                If i < (UBound(BOMList, 2)) Then
    '                                    AssyMK = .Range("C" & i).Value
    '                                    MarkForDelete(LinesToDelete, i)             'Deletes Assembly Marks....
    '                                End If
    '                            End If
    '                        End If
    '                    End If

    '                    '------------------special code for use with references to sub-assemblies on other drawings
    '                ElseIf InStr(1, .Range("F" & i).Value, "SEE DWG") <> 0 Or InStr(1, .Range("F" & i).Value, "SEE DRAWING") <> 0 Then  'if ship mark is empty AND reference to see drawing
    '                    .Range("C" & i).Value = AssyMK
    '                Else                        '---------------------------------if ship mark is empty
    '                    If .Range("C" & i).Value <> Nothing And .Range("D" & i).Value = Nothing And .Range("E" & i).Value <> Nothing Then
    '                        If .Range("C" & i).Value = "" And .Range("D" & i).Value.ToString = "" And .Range("E" & i).Value.ToString = "" Then
    '                            MarkForDelete(LinesToDelete, i)
    '                        Else
    '                            .Range("C" & i).Value = AssyMK
    '                        End If
    '                    End If
    '                End If
    '            End With
    '        Next i

    '        If UBound(LinesToDelete) <> 1 Then      'Now delete header lines not required.
    '            For i = UBound(LinesToDelete) To LBound(LinesToDelete) Step -1
    '                With BOMWrkSht
    '                    If i <> 0 Then
    '                        .Rows(LinesToDelete(i)).Delete()
    '                    Else
    '                        GoTo NextI
    '                    End If
    '                End With
    'NextI:      Next i
    '        End If

    'Err_FormatBulkBOM:
    '        ErrNo = Err.Number

    '        If ErrNo <> 0 Then
    '            PrgName = "FormatBulkBOM"
    '            PriPrg = "ShipListReadBOMAutoCAD"
    '            ErrMsg = Err.Description
    '            ErrSource = Err.Source
    '            ErrDll = Err.LastDllError
    '            ErrLastLineX = Err.Erl
    '            ErrException = Err.GetException

    '            Dim st As New StackTrace(Err.GetException, True)
    '            CntFrames = st.FrameCount
    '            GetFramesSrt = st.GetFrames
    '            PrgLineNo = GetFramesSrt(CntFrames - 1).ToString
    '            PrgLineNo = PrgLineNo.Replace("@", "at")
    '            LenPrgLineNo = (Len(PrgLineNo))
    '            PrgLineNo = Mid(PrgLineNo, 1, (LenPrgLineNo - 2))

    '            ShippingList_Menu.HandleErrSQL(PrgName, ErrNo, ErrMsg, ErrSource, PriPrg, ErrDll, DwgItem, PrgLineNo)
    '            If GenInfo3135.UserName = "dlong" Then
    '                MsgBox(ErrMsg)
    '                Stop
    '                Resume
    '            Else
    '                ExceptionPos = InStr(1, ErrMsg, "Exception")
    '                CallPos = InStr(1, ErrMsg, "Call was rejected by callee")
    '                CntExcept = (CntExcept + 1)

    '                If CntExcept < 20 Then
    '                    If ExceptionPos > 0 Then
    '                        Resume
    '                    End If
    '                    If CallPos > 0 Then
    '                        Resume
    '                    End If
    '                End If
    '            End If

    '        End If

    '    End Function

    '    Function UpdateShpMarks(ByRef BOMList As Object) As Object
    '        Dim Range As Object, Worksheets As Object, i As Object
    '        Dim LinesToDelete As Object, DupShipMk As Object, AssyMK As Object, AssyMKtmp As Object
    '        Dim RefDwgLen As Short, j As Short
    '        Dim RefDwgLetter, RefMk, RefDwg, ShpMark As String
    '        Dim BOMWrkSht As Worksheet, WorkSht As Worksheet
    '        Dim Workbooks As Excel.Workbooks
    '        Dim ExceptionPos As Integer, CallPos As Integer

    '        PrgName = "UpdateShpMarks"
    '        ShpMark = Nothing

    '        On Error Resume Next

    '        ExcelApp = GetObject(, "Excel.Application")

    '        If Err.Number Then
    '            Information.Err.Clear()
    '            ExcelApp = CreateObject("Excel.Application")
    '            If Err.Number Then
    '                MsgBox(Err.Description)
    '                Exit Function
    '            End If
    '        End If

    '        On Error GoTo Err_UpdateShpMarks

    '        Workbooks = ExcelApp.Workbooks
    '        WorkShtName = "Bulk BOM"
    '        BOMWrkSht = Workbooks.Application.Worksheets(WorkShtName)
    '        WorkSht = Workbooks.Application.ActiveSheet
    '        WorkShtName = WorkSht.Name

    '        For i = 5 To (UBound(BOMList, 2) + 3)
    '            With BOMWrkSht                                              'With Worksheets("Bulk BOM")
    '                If .Range("C" & i).Value <> "" Then                     'if ship mark is not empty
    '                    ShpMark = .Range("C" & i).Value
    '                Else                                    '---------------------------------if ship mark is empty.
    '                    .Range("C" & i).Value = ShpMark
    '                End If
    '            End With
    '        Next i

    'Err_UpdateShpMarks:
    '        ErrNo = Err.Number

    '        If ErrNo <> 0 Then
    '            PrgName = "UpdateShpMarks"
    '            PriPrg = "ShipListReadBOMAutoCAD"
    '            ErrMsg = Err.Description
    '            ErrSource = Err.Source
    '            ErrDll = Err.LastDllError
    '            ErrLastLineX = Err.Erl
    '            ErrException = Err.GetException

    '            Dim st As New StackTrace(Err.GetException, True)
    '            CntFrames = st.FrameCount
    '            GetFramesSrt = st.GetFrames
    '            PrgLineNo = GetFramesSrt(CntFrames - 1).ToString
    '            PrgLineNo = PrgLineNo.Replace("@", "at")
    '            LenPrgLineNo = (Len(PrgLineNo))
    '            PrgLineNo = Mid(PrgLineNo, 1, (LenPrgLineNo - 2))

    '            ShippingList_Menu.HandleErrSQL(PrgName, ErrNo, ErrMsg, ErrSource, PriPrg, ErrDll, DwgItem, PrgLineNo)
    '            If GenInfo3135.UserName = "dlong" Then
    '                MsgBox(ErrMsg)
    '                Stop
    '                Resume
    '            Else
    '                ExceptionPos = InStr(1, ErrMsg, "Exception")
    '                CallPos = InStr(1, ErrMsg, "Call was rejected by callee")
    '                CntExcept = (CntExcept + 1)

    '                If CntExcept < 20 Then
    '                    If ExceptionPos > 0 Then
    '                        Resume
    '                    End If
    '                    If CallPos > 0 Then
    '                        Resume
    '                    End If
    '                End If
    '            End If

    '        End If

    '    End Function

    '    Function FormatSTDItems(ByRef StdsBOMList As Object) As Object
    '        Dim Range As Object
    '        Dim Worksheets As Object
    '        Dim i As Object
    '        Dim j As Short
    '        Dim LinesToDelete As Object
    '        Dim DupShipMk As Object
    '        Dim AssyMK As Object
    '        Dim AssyMKtmp As Object
    '        Dim RefDwgLen As Short
    '        Dim RefDwgLetter, RefMk, RefDwg As String
    '        Dim BOMWrkSht As Worksheet
    '        Dim StdItemsWrkSht As Worksheet
    '        Dim WorkSht As Worksheet
    '        Dim Workbooks As Excel.Workbooks

    '        PrgName = "FormatSTDItems"
    '        AssyMK = Nothing

    '        On Error Resume Next

    '        ExcelApp = GetObject(, "Excel.Application")

    '        If Err.Number Then
    '            Information.Err.Clear()
    '            ExcelApp = CreateObject("Excel.Application")
    '            If Err.Number Then
    '                MsgBox(Err.Description)
    '                Exit Function
    '            End If
    '        End If

    '        On Error GoTo Err_FormatSTDItems

    '        Workbooks = ExcelApp.Workbooks
    '        WorkShtName = "STD Items"
    '        StdItemsWrkSht = Workbooks.Application.Worksheets(WorkShtName)
    '        WorkSht = Workbooks.Application.ActiveSheet
    '        WorkShtName = WorkSht.Name
    '        ReDim LinesToDelete(1)
    '        ReDim DupShipMk(1)

    '        'remove any cells that contain only a space
    '        For i = 5 To (UBound(StdsBOMList, 2) + 4)
    '            For j = 1 To 10
    '                With StdItemsWrkSht
    '                    Dim Test                        'Do not make a string type or next lines will not work.
    '                    Test = (Chr(64 + j) & i)
    '                    With .Range(Test)
    '                        If .Value <> vbNullString Then
    '                            If .Value.ToString = " " Then
    '                                .Value = vbNullString
    '                            End If
    '                        End If
    '                    End With
    '                End With
    '            Next j
    '        Next i

    '        Dim FCODwgNo As String
    '        For i = 5 To (UBound(StdsBOMList, 2) + 4)
    '            With StdItemsWrkSht                  'With Worksheets("Bulk BOM")
    '                If .Range("C" & i).Value <> "" Then 'if ship mark is not empty
    '                    If .Range("G" & i).Value = "" Then 'if inventory number is empty
    '                        'special code for use with flush cleanouts
    '                        If InStr(1, .Range("F" & i).Value, "FLUSH CLEAN") <> 0 Then

    '                            MarkForDelete(LinesToDelete, i)
    '                            AssyMK = .Range("C" & i).Value
    '                            FCODwgNo = Left(.Range("A" & i).Value, 2)

    '                            Do Until Left(.Range("A" & i).Value, 2) <> FCODwgNo
    '                                If .Range("C" & i).Value = "" And .Range("D" & i).Value = "" And .Range("E" & i).Value = "" Then
    '                                    MarkForDelete(LinesToDelete, i)
    '                                ElseIf .Range("C" & i).Value <> "" And .Range("D" & i).Value = "" And .Range("G" & i).Value = "" Then

    '                                    MarkForDelete(LinesToDelete, i)
    '                                    AssyMKtmp = .Range("C" & i).Value
    '                                    i = i + 1
    '                                    If .Range("C" & i).Value = "" And Left(.Range("A" & i).Value, 2) = FCODwgNo Then
    '                                        .Range("C" & i).Value = AssyMKtmp
    '                                    End If
    '                                Else
    '                                    If .Range("C" & i).Value = "" Then
    '                                        .Range("C" & i).Value = AssyMK
    '                                    End If
    '                                End If
    '                                i = i + 1
    '                            Loop
    '                            i = i - 1
    '                            'special code for completed lines
    '                        ElseIf .Range("C" & i).Value <> "" And .Range("D" & i).Value <> "" And .Range("E" & i).Value.ToString <> "" Then
    '                        Else 'section for typical material callouts
    '                            If .Range("C" & i).Value <> "" And .Range("C" & i + 1).Value <> "" Then
    '                                'Stop    'Need to fix the next line.
    '                                If InStr(1, .Range("F" & i).Value, "SEE D") <> 0 Then
    '                                    MarkForDelete(LinesToDelete, i)
    '                                End If
    '                            Else
    '                                AssyMK = .Range("C" & i).Value
    '                                MarkForDelete(LinesToDelete, i)             'Deletes Assembly Marks....
    '                            End If
    '                        End If
    '                    End If
    '                    'special code for use with references to
    '                    'sub-assemblies on other drawings
    '                ElseIf InStr(1, .Range("F" & i).Value, "SEE DWG") <> 0 Or InStr(1, .Range("F" & i).Value, "SEE DRAWING") <> 0 Then  'if ship mark is empty AND reference to see drawing
    '                    .Range("C" & i).Value = AssyMK
    '                Else 'if ship mark is empty
    '                    If .Range("C" & i).Value <> Nothing And .Range("D" & i).Value = Nothing And .Range("E" & i).Value <> Nothing Then
    '                        If .Range("C" & i).Value = "" And .Range("D" & i).Value.ToString = "" And .Range("E" & i).Value.ToString = "" Then
    '                            MarkForDelete(LinesToDelete, i)
    '                        Else
    '                            .Range("C" & i).Value = AssyMK
    '                        End If
    '                    End If
    '                End If
    '            End With
    '        Next i

    '        If UBound(LinesToDelete) <> 1 Then
    '            For i = UBound(LinesToDelete) To LBound(LinesToDelete) Step -1
    '                With StdItemsWrkSht
    '                    If i <> 0 Then
    '                        .Rows(LinesToDelete(i)).Delete()
    '                    Else
    '                        GoTo NextI
    '                    End If
    '                End With
    'NextI:      Next i
    '        End If

    'Err_FormatSTDItems:
    '        ErrNo = Err.Number

    '        If ErrNo <> 0 Then
    '            PrgName = "FormatSTDItems"
    '            PriPrg = "ShipListReadBOMAutoCAD"
    '            ErrMsg = Err.Description
    '            ErrSource = Err.Source
    '            ErrDll = Err.LastDllError
    '            ErrLastLineX = Err.Erl
    '            ErrException = Err.GetException

    '            Dim st As New StackTrace(Err.GetException, True)
    '            CntFrames = st.FrameCount
    '            GetFramesSrt = st.GetFrames
    '            PrgLineNo = GetFramesSrt(CntFrames - 1).ToString
    '            PrgLineNo = PrgLineNo.Replace("@", "at")
    '            LenPrgLineNo = (Len(PrgLineNo))
    '            PrgLineNo = Mid(PrgLineNo, 1, (LenPrgLineNo - 2))

    '            ShippingList_Menu.HandleErrSQL(PrgName, ErrNo, ErrMsg, ErrSource, PriPrg, ErrDll, DwgItem, PrgLineNo)
    '            If GenInfo3135.UserName = "dlong" Then
    '                MsgBox(ErrMsg)
    '                Stop
    '                Resume
    '            Else
    '                ExceptionPos = InStr(1, ErrMsg, "Exception")
    '                CallPos = InStr(1, ErrMsg, "Call was rejected by callee")
    '                CntExcept = (CntExcept + 1)

    '                If CntExcept < 20 Then
    '                    If ExceptionPos > 0 Then
    '                        Resume
    '                    End If
    '                    If CallPos > 0 Then
    '                        Resume
    '                    End If
    '                End If
    '            End If
    '        End If

    '    End Function

    '    Function MarkForDelete(ByRef LinesToDelete As Object, ByVal LineNo As Short) As Object
    '        PrgName = "MarkForDelete"

    '        On Error GoTo Err_MarkForDelete

    '        Dim Test
    '        Test = LinesToDelete(UBound(LinesToDelete))

    '        If LinesToDelete(UBound(LinesToDelete)) = Nothing Then
    '            If LinesToDelete(UBound(LinesToDelete)) <> LineNo Then
    '                LinesToDelete(UBound(LinesToDelete)) = LineNo
    '            End If
    '        Else
    '            If LinesToDelete(UBound(LinesToDelete)) <> LineNo Then
    '                ReDim Preserve LinesToDelete(UBound(LinesToDelete) + 1)
    '                LinesToDelete(UBound(LinesToDelete)) = LineNo
    '            End If
    '        End If

    'Err_MarkForDelete:
    '        ErrNo = Err.Number

    '        If ErrNo <> 0 Then
    '            PrgName = "MarkForDelete"
    '            PriPrg = "ShipListReadBOMAutoCAD"
    '            ErrMsg = Err.Description
    '            ErrSource = Err.Source
    '            ErrDll = Err.LastDllError
    '            ErrLastLineX = Err.Erl
    '            ErrException = Err.GetException

    '            Dim st As New StackTrace(Err.GetException, True)
    '            CntFrames = st.FrameCount
    '            GetFramesSrt = st.GetFrames
    '            PrgLineNo = GetFramesSrt(CntFrames - 1).ToString
    '            PrgLineNo = PrgLineNo.Replace("@", "at")
    '            LenPrgLineNo = (Len(PrgLineNo))
    '            PrgLineNo = Mid(PrgLineNo, 1, (LenPrgLineNo - 2))

    '            ShippingList_Menu.HandleErrSQL(PrgName, ErrNo, ErrMsg, ErrSource, PriPrg, ErrDll, DwgItem, PrgLineNo)
    '            If GenInfo3135.UserName = "dlong" Then
    '                MsgBox(ErrMsg)
    '                Stop
    '                Resume
    '            Else
    '                ExceptionPos = InStr(1, ErrMsg, "Exception")
    '                CallPos = InStr(1, ErrMsg, "Call was rejected by callee")
    '                CntExcept = (CntExcept + 1)

    '                If CntExcept < 20 Then
    '                    If ExceptionPos > 0 Then
    '                        Resume
    '                    End If
    '                    If CallPos > 0 Then
    '                        Resume
    '                    End If
    '                End If
    '            End If
    '        End If

    '    End Function
End Module