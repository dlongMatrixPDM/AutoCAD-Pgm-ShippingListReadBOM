Option Strict Off
Option Explicit On
Option Compare Text

'Imports VB = Microsoft.VisualBasic
Imports System
'Imports System.Drawing
Imports System.Reflection
Imports System.Runtime.InteropServices
Imports System.Text.RegularExpressions
'Imports Autodesk.AutoCAD.Interop
'Imports Autodesk.AutoCAD.Interop.Common
'Imports Autodesk.AutoCAD
'Imports Autodesk.AutoCAD.Runtime
'Imports Autodesk.AutoCAD.ApplicationServices
'Imports Autodesk.AutoCAD.DatabaseServices
'Imports AutoCAD = Autodesk.AutoCAD.Interop
Imports Microsoft.Office.Interop.Excel

Public Module Misc
    Public Structure InputType2
        'Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, ByRef lParam As Any) As Integer
        Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, ByRef lParam As VariantType) As Integer

        Public Structure FILETIME
            Dim dwLowDateTime As Integer
            Dim dwHighDateTime As Integer
        End Structure

        Public Structure SECURITY_ATTRIBUTES
            Dim nLength As Integer
            Dim lpSecurityDescriptor As Integer
            Dim bInheritHandle As Boolean
        End Structure

        Public PassFilename As String
        Public ReadyToContinue As Boolean
        Public SBclicked As Boolean
        Public CBclicked As Boolean
        Public errorExist As Boolean
        Public BOMList() As String
        Public ShippingList() As String
        Public RowNo As String
        Public RowNo2 As String
        'Public AcadApp As Object
        Public AcadDoc As Object
        Public Shared NewBulkBOM As Object
        Public NewPlateBOM As Object
        Public NewStickBOM As Object
        Public NewPurchaseBOM As Object
        Public OldBOMFile As Excel.Workbook
        Public OldBulkBOM As Object
        Public OldBulkBOMFile As String
        Public OldPlateBOM As Object
        Public OldStickBOM As Object
        Public OldPurchaseBOM As Object
        Public NewBOM As Object
        Public OldBOM As Object
        Public OldStdItems As Object
        'Public GetStdFilter As Object
        '	Public MainBOMFile As Object
        'Public BOMType As String                           '-------DJL-------10-31-2023---Not needed anymore.
        Public BOMSheet As String
        Public RevNo As String
        Public RevNo2 As String
        Public Shared MainBOMFile As Excel.Workbook

        Public Continue_Renamed As Boolean
        Public SortListing As Boolean
        Public ExceptionPos As Integer
        Public CallPos As Integer
        Public CntExcept As Integer
        Public Shared ErrNo As String
        Public ErrMsg As String
        Public ErrSource As String
        Public ErrDll As String
        Public PriPrg As String
        Public PrgName As String
        Public ErrException As System.Exception
        Public ErrLastLineX As Integer
        Public OldStdDwg As String
        Public NewStdDwg As String
        Public FuncGetDataNew As String
        Public Count As Integer
        Public MatInch As Double
        Public FoundDir As String
        Public SearchException As String
        Public ExceptPos As Integer
        'Public ThisDrawing As AcadApp.ActiveDocument     'AcadApp.ActiveDocument
        'Public Mospace As AcadModelSpace = ThisDrawing.ModelSpace
        'Public Paspace As AcadPaperSpace = ThisDrawing.PaperSpace
        'Public UtilObj As AcadUtility = ThisDrawing.Utility
        Public LytHid As Boolean
        Public PurchaseProb As Boolean
        Public ProbPart, BOMorShip As String
        Public CountBOM As Integer
        Public STDsList
        Public NewShipListSht As Object
        Public OldShipListSht As Object
        Public MainShipListFile As Object
        'Public OldShipListFile As String
        Public ShipListType As String
        Public CustomerPO, Cust1, Cust2 As String
        Shared AcadOpen As Boolean
        Shared ThisDrawing As AutoCAD.AcadDocument
        Shared Mospace, Paspace As Object
        Shared UtilObj As Object
        Public Shared AcadApp As AutoCAD.AcadApplication       'Dim AcadApp As Object
        Public Shared AcadApp2 As AutoCAD.AcadApplication       'Dim AcadApp As Object

        Function GetCadVersion() As Object
            Dim CadVersion As String

            PrgName = "GetCadVersion"

            On Error GoTo Err_GetCadVersion
            'CadVersion = AutoCAD.ApplicationServices.Application.Version.ToString

            'Select Case CadVersion
            '    Case "17.0s (LMS Tech)"
            '        FoundDir = "k:/CAD/Lisp/"       '"k:/cad/lisp/bolts/"   ----Used Below
            '    Case "18.0s (LMS Tech)"
            '        FoundDir = "k:/CAD/Lisp/"                   '"k:/CAD/Lisp2010/"   'No Longer required Network is not mult versions.
            '    Case "18.1s (LMS Tech)"
            '        FoundDir = "k:/CAD/Lisp/"                   '"k:/CAD/Lisp2011/"
            '    Case Else
            '        MsgBox("AutoCAD Version not found, please create an IT Ticket. ")
            'End Select

Err_GetCadVersion:
            ErrNo = Err.Number

            If ErrNo <> 0 Then
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
                If System.Environment.UserName = "dlong" Then           '-------DJL-06-30-2025     'If (ShippingList_Menu.UserNamex()) = "dlong" Then
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

        Public Shared Function IsArrayEmpty(ByRef Temparray As Object) As Boolean
            On Error Resume Next                'added back by monty lowe to over come error on remove 11/15/06
            IsArrayEmpty = UBound(Temparray)
            IsArrayEmpty = Err.Number
        End Function

        Public Function StripNulls(ByRef startstr As String) As String
            Dim pos As Integer

            pos = InStr(startstr, Chr(0))
            If pos Then
                StripNulls = Mid(startstr, 1, pos - 1)
                startstr = Mid(startstr, pos + 1, Len(startstr))
            End If
        End Function

        Public Shared Function vbdApiTrim(ByVal strName As String) As String
            Dim intLoc As Short

            intLoc = InStr(strName, Chr(0))
            If intLoc > 0 Then
                vbdApiTrim = Left(strName, intLoc - 1)
            Else
                vbdApiTrim = strName
            End If
        End Function

        Public Shared Sub vbdQsort(ByRef varArray As Object, Optional ByRef intLbound As Object = Nothing, Optional ByRef intUbound As Object = Nothing)
            Dim intLow, intHigh, intLeft, intRight, intMid As Short
            Dim varTestVal, varHold As Object

            If IsNothing(intLbound) Then
                intLeft = LBound(varArray)
            Else
                intLeft = intLbound
            End If

            If IsNothing(intUbound) Then
                intRight = UBound(varArray)
            Else
                intRight = intUbound
            End If

            If intLeft < intRight Then
                intLow = intLeft
                intHigh = intRight
                intMid = (intLow + intHigh) \ 2
                varTestVal = varArray(intMid)
                Do
                    Do While varArray(intLow) < varTestVal
                        intLow = intLow + 1
                    Loop

                    Do While varArray(intHigh) > varTestVal
                        intHigh = intHigh - 1
                    Loop
                    If intLow <= intHigh Then
                        varHold = varArray(intHigh)
                        varArray(intHigh) = varArray(intLow)
                        varArray(intLow) = varHold
                        intLow = intLow + 1
                        intHigh = intHigh - 1
                    End If
                Loop Until intLow > intHigh
                If intHigh <= intMid Then
                    Call vbdQsort(varArray, intLeft, intHigh)
                    Call vbdQsort(varArray, intLow, intRight)
                Else
                    Call vbdQsort(varArray, intLow, intRight)
                    Call vbdQsort(varArray, intLeft, intHigh)
                End If
            End If
        End Sub

        ' SearchLine is input, SearchFor is what to search for, ReplaceWith is the replacement
        Public Shared Function sReplace(ByRef SearchLine As String, ByRef SearchFor As String, ByRef ReplaceWith As String) As Object
            Dim vSearchLine As String
            Dim found As Short

            found = InStr(SearchLine, SearchFor) : vSearchLine = SearchLine
            If found <> 0 Then
                vSearchLine = ""
                If found > 1 Then vSearchLine = Left(SearchLine, found - 1)
                vSearchLine = vSearchLine & ReplaceWith
                If found + Len(SearchFor) - 1 < Len(SearchLine) Then vSearchLine = vSearchLine & Right(SearchLine, Len(SearchLine) - found - Len(SearchFor) + 1)
            End If
            sReplace = vSearchLine

        End Function

        Public Sub Open_AutoCAD()
            Dim Mospace As AutoCAD.AcadModelSpace = ThisDrawing.ModelSpace
            Dim Paspace As AutoCAD.AcadPaperSpace = ThisDrawing.PaperSpace
            Dim UtilObj As AutoCAD.AcadUtility = ThisDrawing.Utility

            On Error Resume Next
            If Err.Number Then
                Err.Clear()
            End If

            AcadApp = GetObject(, "AutoCAD.Application.")

            If Err.Number Then
                AcadApp = CreateObject("AutoCAD.Application.")
            End If

            AcadApp.Visible = False          '        AcadApp.Visible = False
            ThisDrawing = AcadApp.ActiveDocument
            Mospace = ThisDrawing.ModelSpace
            Paspace = ThisDrawing.PaperSpace
            UtilObj = ThisDrawing.Utility

        End Sub

        'Sub LayerChange(ByVal TmpLayer As String)
        '    Dim TmpStr As String
        '    LytHid = False
        '    If TmpLayer = "HIDDEN" Then
        '        For i As Integer = 0 To AcadApp.ActiveDocument.Layers.Count() - 1
        '            If AcadApp.ActiveDocument.Layers.Item(i).Name = "LAYOUT_HIDDEN" Then
        '                LytHid = True
        '            End If
        '        Next
        '        If LytHid = False Then
        '            AcadApp.ActiveDocument.Layers.Add("LAYOUT_HIDDEN")
        '            TmpLayer = "LAYOUT_HIDDEN"
        '            AcadApp.ActiveDocument.Layers.Item("LAYOUT_HIDDEN").Linetype = "Hidden"
        '            AcadApp.ActiveDocument.Layers.Item("LAYOUT_HIDDEN").TrueColor = AcadApp.ActiveDocument.Layers.Item("SMALL").TrueColor
        '        Else
        '            TmpLayer = "LAYOUT_HIDDEN"
        '        End If
        '    End If
        '    TmpStr = "clayer" & Chr(13)
        '    TmpStr = TmpStr & Trim$(TmpLayer) & Chr(13)

        '    AcadDoc.SendCommand(TmpStr)
        'End Sub

        'Sub LayerMake(ByVal TmpLn)
        '    Dim TmpStr As String

        '    TmpStr = "(command " '& Chr(13)
        '    TmpStr = TmpStr & Chr(34) & "-layer" & Chr(34)
        '    TmpStr = TmpStr & Chr(34) & "m" & Chr(34)
        '    TmpStr = TmpStr & Chr(34) & Trim$(TmpLn) & Chr(34) ' & Chr(13)
        '    TmpStr = TmpStr & Chr(34) & "c" & Chr(34)
        '    TmpStr = TmpStr & Chr(34) & "2" & Chr(34) ' & Chr(13)
        '    TmpStr = TmpStr & Chr(34) & Chr(34) & " " & Chr(34) & Chr(34) & ")" & Chr(13)

        '    acadDoc.SendCommand(TmpStr)
        'End Sub

        'Sub InsertDwg(ByVal fname As String, ByVal ps1 As Pxy, ByVal scfx As Double, ByVal inang As Integer)
        '    Dim TmpStr As String

        '    TmpStr = "-insert" & Chr(13) '01.09.02 MTeague changed .insert to -insert
        '    TmpStr = TmpStr & "*" & Trim$(fname) & Chr(13)
        '    TmpStr = TmpStr & ps1.X & "," & ps1.Y & Chr(13)
        '    TmpStr = TmpStr & scfx & Chr(13)
        '    TmpStr = TmpStr & inang & Chr(13)
        '    'TmpStr = TmpStr & "]"

        '    'acadDoc.ModelSpace.InsertBlock("", "*" & fname.Trim, ps1.X, ps1.Y, scfx, inang)


        '    acadDoc.SendCommand(TmpStr)
        'End Sub

        '<CommandMethod("ConnectToAcad")> _
        Public Sub ConnectToAcad()
            Dim acAppComObj As AutoCAD.AcadApplication
            Dim strProgId As String = "AutoCAD.Application."

            On Error Resume Next

            '' Get a running instance of AutoCAD
            acAppComObj = GetObject(, strProgId)

            '' An error occurs if no instance is running
            If Err.Number > 0 Then
                Err.Clear()

                '' Create a new instance of AutoCAD
                acAppComObj = CreateObject("AutoCAD.Application.")

                '' Check to see if an instance of AutoCAD was created
                If Err.Number > 0 Then
                    Err.Clear()

                    '' If an instance of AutoCAD is not created then message and exit
                    MsgBox("Instance of 'AutoCAD.Application' could not be created.")

                    Exit Sub
                End If
            End If

            '' Display the application and return the name and version
            acAppComObj.Visible = True
            MsgBox("Now running " & acAppComObj.Name & " version " & acAppComObj.Version)

            '' Get the active document
            Dim acDocComObj As AutoCAD.AcadDocument
            acDocComObj = acAppComObj.ActiveDocument

            '' Optionally, load your assembly and start your command or if your assembly
            '' is demandloaded, simply start the command of your in-process assembly.
            acDocComObj.SendCommand("(command " & Chr(34) & "NETLOAD" & Chr(34) & " " & _
                                    Chr(34) & "c:/myapps/mycommands.dll" & Chr(34) & ") ")

            acDocComObj.SendCommand("MyCommand ")
        End Sub

        '<CommandMethod("SetLayerCurrent")> _
        'Public Sub SetLayerCurrent()                    ' Get the current document and database
        '    Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        '    Dim acCurDb As Database = acDoc.Database

        '    Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()    ' Start a transaction
        '        Dim acLyrTbl As LayerTable                              ' Open the Layer table for read
        '        acLyrTbl = acTrans.GetObject(acCurDb.LayerTableId, _
        '                                     OpenMode.ForRead)

        '        Dim sLayerName As String = "Center"

        '        If acLyrTbl.Has(sLayerName) = True Then
        '            acCurDb.Clayer = acLyrTbl(sLayerName)               ' Set the layer Center current
        '            acTrans.Commit()                                    ' Save the changes
        '        End If
        '    End Using                                                   ' Dispose of the transaction

        'End Sub

        '        Public Shared Function SetLayerCurrent(ByVal FindLayer As String, ByVal FoundStatus As String) As String
        '            Dim colLayers As AutoCAD.AcadLayers
        '            Dim objNewLayer As AutoCAD.AcadLayer
        '            'Dim FoundStatus As Boolean
        '            Dim objLayer As AutoCAD.AcadLayer
        '            FoundStatus = False

        '            On Error Resume Next

        '            If Err.Number Then
        '                Err.Clear()
        '            End If

        '            AcadApp = GetObject(, "AutoCAD.Application.")     'AcadApp = GetObject(, "Autocad.Application")
        '            AcadOpen = True
        '            ErrNo = Err.Number

        '            If ErrNo > 0 Then
        '                'MsgBox(Err.Description)
        '                Information.Err.Clear()
        '                AcadApp = CreateObject("AutoCAD.Application.")    'AcadApp = CreateObject("Autocad.Application")
        '                AcadOpen = True             'AcadOpen = False
        '                If Err.Number Then
        '                    Information.Err.Clear()
        '                    AcadApp = CreateObject("AutoCAD.Application.")
        '                    If Err.Number Then
        '                        'MsgBox(Err.Description)
        '                        'MsgBox("Instance of 'AutoCAD.Application' could not be created.")
        '                        'AcadApp.Visible = True
        '                        'MsgBox("Now running " & AcadApp.Name & " version " & AcadApp.Version)

        '                        If (ShippingList_Menu.UserNamex()) = "dlong" Then
        '                            MsgBox(Err.Description)
        '                            MsgBox("Instance of 'AutoCAD.Application' could not be created.")

        '                            AcadApp.Visible = True
        '                            MsgBox("Now running " & AcadApp.Name & " version " & AcadApp.Version)
        '                            Stop
        '                            Resume
        '                        Else
        '                            Exit Function
        '                        End If
        '                    End If
        '                End If
        '            End If

        '            AcadApp.Visible = True
        '            ThisDrawing = AcadApp.ActiveDocument
        '            Mospace = ThisDrawing.ModelSpace
        '            Paspace = ThisDrawing.PaperSpace
        '            UtilObj = ThisDrawing.Utility

        '            colLayers = ThisDrawing.Layers

        '            For Each objLayer In colLayers
        '                Select Case (objLayer.Name)
        '                    Case FindLayer
        '                        FoundStatus = True
        '                        GoTo LayerFound
        '                End Select
        '            Next objLayer

        'LayerFound:
        '            If FoundStatus = False Then
        '                objNewLayer = ThisDrawing.Layers.Add(FindLayer)
        '                objNewLayer.color = CShort("40")
        '                ThisDrawing.SetVariable("CLayer", FindLayer)
        '            Else
        '                ThisDrawing.SetVariable("CLayer", FindLayer)
        '            End If

        '        End Function

        '<CommandMethod("SetLayerCurrent")> _
        'Public Sub SearchLayer(ByVal FindLayer As String, ByVal FoundStatus As String)          'SetLayerCurrent()  'Get the current document and database
        '    Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        '    Dim acCurDb As Database = acDoc.Database
        '    Dim acLyrTbl As LayerTable

        '    Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()    ' Start a transaction
        '        acLyrTbl = acTrans.GetObject(acCurDb.LayerTableId, OpenMode.ForRead)    'Open the Layer table for read

        '        If acLyrTbl.Has(FindLayer) = True Then                 'Set the layer Center current
        '            FoundStatus = True
        '            'acCurDb.Clayer = acLyrTbl(FindLayer)
        '            'acTrans.Commit()                                    'Save the changes
        '        End If
        '    End Using                                                   'Dispose of the transaction
        'End Sub

        'Public Shared Function SearchLayer(ByVal FindLayer As String, ByVal FoundStatus As String) As String
        '    Dim colLayers As AutoCAD.AcadLayers
        '    Dim objNewLayer As AutoCAD.AcadLayer
        '    Dim objLayer

        '    colLayers = ThisDrawing.Layers

        '    For Each objLayer In colLayers
        '        Select Case (objLayer.Name)
        '            Case FindLayer
        '                FoundStatus = True
        '        End Select
        '    Next objLayer

        '    If FoundStatus = "FALSE" Then
        '        objNewLayer = ThisDrawing.Layers.Add(FindLayer)
        '        objNewLayer.color = CShort("40")
        '        ThisDrawing.SetVariable("CLayer", FindLayer)
        '    Else
        '        ThisDrawing.SetVariable("CLayer", FindLayer)
        '    End If

        'End Function

    End Structure
End Module