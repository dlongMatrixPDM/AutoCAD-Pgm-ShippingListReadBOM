Option Strict Off
Option Explicit On
Option Compare Text

Imports System
Imports System.Reflection
Imports System.Runtime.InteropServices.ComTypes
'Imports Autodesk.AutoCAD.Interop
'Imports Autodesk.AutoCAD.Interop.Common
'Imports Autodesk.AutoCAD
'Imports Autodesk.AutoCAD.Runtime
'Imports Autodesk.AutoCAD.ApplicationServices
'Imports AutoCAD = Autodesk.AutoCAD.Interop
Imports Microsoft.Office.Interop.Excel

Module File

    Private Const BFFM_INITIALIZED As Short = 1
    Private Const BFFM_SETSELECTION As Integer = &H466
    Private Const BIF_DONTGOBELOWDOMAIN As Short = 2
    Private Const BIF_RETURNONLYFSDIRS As Short = 1
    Private Const MAX_PATH As Short = 260

    Private Declare Function CreateDirectory Lib "kernel32.dll" Alias "CreateDirectoryA" (ByVal lpPathName As String, ByRef lpSecurityAttributes As Misc.InputType2.SECURITY_ATTRIBUTES) As Integer
    Private Declare Function FindClose Lib "kernel32.dll" (ByVal hFindFile As Integer) As Integer
    Private Declare Function FindFirstFile Lib "kernel32.dll" Alias "FindFirstFileA" (ByVal lpFileName As String, ByRef lpFindFileData As WIN32_FIND_DATA) As Integer
    Private Declare Function FindNextFile Lib "kernel32.dll" Alias "FindNextFileA" (ByVal hFindFile As Integer, ByRef lpFindFileData As WIN32_FIND_DATA) As Integer
    Private Declare Function GetPrivateProfileInt Lib "kernel32.dll" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Integer, ByVal lpFileName As String) As Integer
    Private Declare Function SHBrowseForFolder Lib "shell32.dll" (ByRef lpbi As BrowseInfo) As Integer
    Private Declare Function SHGetPathFromIDList Lib "shell32.dll" (ByVal pidList As Integer, ByVal lpBuffer As String) As Integer
    Private Declare Function GetCurrentVBAProject Lib "vba332.dll" Alias "EbGetExecutingProj" (ByRef hProject As Integer) As Integer
    Private Declare Function GetAddr Lib "vba332.dll" Alias "TipGetLpfnOfFunctionId" (ByVal hProject As Integer, ByVal strFunctionId As String, ByRef lpfn As Integer) As Integer
    Private Declare Function GetFuncID Lib "vba332.dll" Alias "TipGetFunctionId" (ByVal hProject As Integer, ByVal strFunctionName As String, ByRef strFunctionId As String) As Integer
    Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (ByRef pOpenfilename As OPENFILENAME) As Integer

    Dim WorkShtName, PriPrg, ErrNo, ErrMsg, ErrSource, ErrDll, ErrLastLineX, PrgName As String
    Dim ErrException As System.Exception
    Public ExcelApp As Object
    Public FirstTimeThru As String
    Public FuncGetDataNew As String
    Public NewBulkBOM As Object
    Public MainBOMFile As Excel.Workbook
    Public NewPlateBOM As Object
    Public NewStickBOM As Object
    Public NewPurchaseBOM As Object
    Public OldBOMFile As Excel.Workbook
    Public OldBulkBOM As Object
    Public OldBulkBOMFile As String
    Public OldShipListFile As Excel.Workbook                'Public OldShipListFile As String
    Public OldPlateBOM As Object
    Public OldStickBOM As Object
    Public OldPurchaseBOM As Object
    Public NewBOM As Object
    Public OldBOM As Object
    Public OldStdItems As Object
    'Public BOMType As String                           '-------DJL-------10-31-2023---Not needed anymore.
    Public BOMSheet As String
    Public RowNo As String
    Public RowNo2 As String
    Public OldStdDwg As String
    Public NewStdDwg As String
    Public ExceptionPos As Integer
    Public CallPos As Integer
    Public CntExcept As Integer
    Public Count As Integer
    Public PassFilename As String
    Public ReadyToContinue As Boolean
    Public SBclicked As Boolean
    Public CBclicked As Boolean
    Public errorExist As Boolean
    'Public BOMList() As String
    Public AcadApp As Object
    Public AcadDoc As Object
    'Public GetStdFilter As Object
    '	Public MainBOMFile As Object
    Public RevNo As String
    Public RevNo2 As String
    Public Continue_Renamed As Boolean
    Public SortListing As Boolean
    Public MatInch As Double
    Public FoundDir As String
    Public SearchException As String
    Public ExceptPos As Integer
    Public ThisDrawing As AutoCAD.AcadDocument     'AcadApp.ActiveDocument
    'Public Mospace As AcadModelSpace = ThisDrawing.ModelSpace
    'Public Paspace As AcadPaperSpace = ThisDrawing.PaperSpace
    'Public UtilObj As AcadUtility = ThisDrawing.Utility
    Public LytHid As Boolean
    Public DwgItem As String
    Public xlNormal As Object
    Public ShipListType As String
    Public NewShipListSht As Worksheet

    Private Structure BrowseInfo
        Dim hwndOwner As Integer
        Dim pIDLRoot As Integer
        Dim pszDisplayName As String
        Dim lpszTitle As String
        Dim ulFlags As Integer
        Dim lpfnCallback As Integer
        Dim lParam As Integer
        Dim iImage As Integer
    End Structure

    Private Structure WIN32_FIND_DATA
        Dim dwFileAttributes As Integer
        Dim ftCreationTime As FILETIME
        Dim ftLastAccessTime As FILETIME
        Dim ftLastWriteTime As FILETIME
        Dim nFileSizeHigh As Integer
        Dim nFileSizeLow As Integer
        Dim dwReserved0 As Integer
        Dim dwReserved1 As Integer
        <VBFixedString(MAX_PATH), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst:=MAX_PATH)> Public cFileName() As Char
        <VBFixedString(14), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst:=14)> Public cAlternate() As Char
    End Structure

    Private Structure OPENFILENAME
        Dim lStructSize As Integer
        Dim hwndOwner As Integer
        Dim hInstance As Integer
        Dim lpstrFilter As String
        Dim lpstrCustomFilter As String
        Dim nMaxCustFilter As Integer
        Dim nFilterIndex As Integer
        Dim lpstrFile As String
        Dim nMaxFile As Integer
        Dim lpstrFileTitle As String
        Dim nMaxFileTitle As Integer
        Dim lpstrInitialDir As String
        Dim lpstrTitle As String
        Dim flags As Integer
        Dim nFileOffset As Short
        Dim nFileExtension As Short
        Dim lpstrDefExt As String
        Dim lCustData As Integer
        Dim lpfnHook As Integer
        Dim lpTemplateName As String
    End Structure

    Public currentDir As String

    Public Function AddrOf(ByRef strFuncName As String) As Integer
        Dim hProject As Integer
        Dim lngResult As Integer
        Dim strID As String
        Dim lpfn As Integer
        Dim strFuncNameUnicode As String

        Const NO_ERROR As Short = 0
        strID = Nothing

        strFuncNameUnicode = StrConv(strFuncName, VbStrConv.None)        'strFuncNameUnicode = StrConv(strFuncName, vbUnicode)
        Call GetCurrentVBAProject(hProject)
        If hProject <> 0 Then
            lngResult = GetFuncID(hProject, strFuncNameUnicode, strID)
            If lngResult = NO_ERROR Then
                lngResult = GetAddr(hProject, strID, lpfn)
                If lngResult = NO_ERROR Then
                    AddrOf = lpfn
                End If
            End If
        End If
    End Function

    Public Function BrowseCallbackProc(ByVal hwnd As Integer, ByVal uMsg As Integer, ByVal lParam As Integer, ByVal lpData As Integer) As Integer
        Dim retval As Integer

        Select Case uMsg
            Case BFFM_INITIALIZED
                retval = Misc.InputType2.SendMessage(hwnd, BFFM_SETSELECTION, CInt(1), currentDir)
        End Select
        BrowseCallbackProc = 0
    End Function

    Public Function CreatePath(ByRef DirPath As String) As Object
        Dim tempDir As String
        Dim tempChar As String
        Dim Count As Short
        Dim secattr As Misc.InputType2.SECURITY_ATTRIBUTES

        Count = Len(DirPath)

        If Dir(DirPath) = "" Then
            Do
                tempChar = Mid(DirPath, Count, 1)
                Count = Count - 1
            Loop Until tempChar = "\"
            tempDir = Left(DirPath, Count)
            CreatePath(tempDir)
            CreateDirectory(DirPath, secattr)
        End If
    End Function

    Function GetDir(ByRef szTitle As String) As String
        Dim lpIDList As Integer
        Dim sBuffer As String
        Dim tBrowseInfo As BrowseInfo

        With tBrowseInfo
            .lpszTitle = szTitle
            .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
            .lpfnCallback = AddrOf("BrowseCallbackProc")
        End With

        lpIDList = SHBrowseForFolder(tBrowseInfo)

        If (lpIDList) Then
            sBuffer = Space(MAX_PATH)
            SHGetPathFromIDList(lpIDList, sBuffer)
            sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
            If Right(sBuffer, 1) <> "\" Then sBuffer = sBuffer & "\"
            GetDir = sBuffer
        Else
            GetDir = ""
        End If

    End Function

    Function GetFile(ByRef startdir As String) As String

        Dim OpenFile As OPENFILENAME
        Dim lReturn As Integer
        Dim sFilter As String
        Dim s As String

        OpenFile.lStructSize = Len(OpenFile)
        sFilter = "Excel Worksheet(*.xls)" & Chr(0) & "*.xls" & Chr(0)
        OpenFile.lpstrFilter = sFilter
        OpenFile.nFilterIndex = 1
        OpenFile.lpstrFile = New String(Chr(0), 257)
        OpenFile.nMaxFile = Len(OpenFile.lpstrFile) - 1
        OpenFile.lpstrFileTitle = OpenFile.lpstrFile
        OpenFile.nMaxFileTitle = OpenFile.nMaxFile
        OpenFile.lpstrInitialDir = startdir
        OpenFile.lpstrTitle = "Select file to Open"
        'OpenFile.flags = cdlOFNFileMustExist + cdlOFNHideReadOnly
        lReturn = GetOpenFileName(OpenFile)
        If lReturn = 0 Then
            GetFile = ""
        Else
            s = Trim(OpenFile.lpstrFile)
            If InStr(s, Chr(0)) Then s = Left(s, InStr(s, Chr(0)) - 1)
            GetFile = s
            OldBulkBOMFile = GetFile
            OldShipListFile = GetFile
            'OpenExcelFile GetFile
        End If
    End Function

    Function ViewFiles(ByVal BOMorShip As String) As Object
        Dim colFiles As New Collection
        Dim intCnt As Short
        Dim lngFile As Integer
        Dim lngReturn As Integer
        Dim varFileArray As Object
        Dim WinDat As WIN32_FIND_DATA
        Dim ViewFilesDialog As ViewFilesDialog
        ViewFilesDialog = ViewFilesDialog

        ViewFilesDialog.Filelist.Items.Clear()
        WinDat.cFileName = ""
        lngFile = FindFirstFile("*.*" & Chr(0), WinDat)

        If lngFile < 0 Then
            Exit Function
        End If

        Do                  'BOMorShip = "*BOM*.XLS*" or "*SH*.XLS*"
            If UCase(WinDat.cFileName) Like BOMorShip Then
                colFiles.Add(Misc.InputType2.vbdApiTrim(UCase(WinDat.cFileName)))
            End If
            WinDat.cFileName = ""
            lngReturn = FindNextFile(lngFile, WinDat)
        Loop Until lngReturn = False

        lngReturn = FindClose(lngFile)

        ReDim varFileArray(colFiles.Count())

        For intCnt = 1 To UBound(varFileArray) 'LBound(varFileArray)
            varFileArray(intCnt) = colFiles.Item(intCnt)
        Next

        Misc.InputType2.vbdQsort(varFileArray)

        For intCnt = 1 To UBound(varFileArray) 'LBound(varFileArray)
            ViewFilesDialog.Filelist.Items.Add(varFileArray(intCnt))
        Next

    End Function

    Function CopyBOMFile(ByVal OldFileNam As String, ByVal RevNo As String) As Object
        Dim xlNormal As Object
        Dim Excel As Object
        Dim Sheets As Object
        Dim Range As Object
        Dim Worksheets As Object
        Dim FileDir, JobNo, BomListRev, BomListFileName As String
        Dim BOMMnu As ShippingList_Menu
        Dim SaveAsFilename As SaveAsFilename
        Dim BOMWrkSht As Worksheet
        Dim ShopCutWrkSht As Worksheet
        Dim WorkSht As Worksheet
        Dim Workbooks As Excel.Workbooks
        BOMMnu = ShippingList_Menu

        SaveAsFilename = SaveAsFilename

        PrgName = "CopyBOMFile"
        BomListFileName = Nothing

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

        On Error GoTo Err_CopyBOMFile

        FileDir = OldFileNam              'FileDir = Main_Menu.PathBox.Text & "\"

        Workbooks = ExcelApp.Workbooks
        WorkShtName = "Shipping List"
        BOMWrkSht = Workbooks.Application.Worksheets(WorkShtName)
        WorkSht = Workbooks.Application.ActiveSheet
        WorkShtName = WorkSht.Name

        With BOMWrkSht                          'Worksheets("Bulk BOM")
            JobNo = .Range("E3").Value
        End With

        BomListRev = RevNo                      'Main_Menu.ComboBox1.Text

        'Select Case BOMMnu.BOMType
        '    Case "Tank"
        BomListFileName = FileDir & JobNo & "-ShippingList-R" & BomListRev & ".xls"

        '    Case "Seal"
        '        BomListFileName = FileDir & JobNo & "-SEALShippingList-R" & BomListRev & ".xls"
        '    Case "Intent"
        '        BomListFileName = FileDir & JobNo & "-INTENTShippingList-R" & BomListRev & ".xls"
        'End Select

        BOMMnu.MainBOMFile.Worksheets.Copy()        'MainBOMFile.Worksheets.Copy()  '----Found Problem that needs to be fixed.
        NewBulkBOM = BOMWrkSht          'Excel.Application.ActiveWorkbook.Sheets("Bulk BOM")

CheckFileName:
        Dim Style, Msg, Title As String
        Dim Response As Object

        MsgBox("After about 30 seconds check your spreadsheet make sure it is not waiting you to Save/Continue?")

        If Dir(BomListFileName) <> vbNullString Then
            Msg = "File " & BomListFileName & " already exists. Do you want to overwrite it?"
            Style = CStr(MsgBoxStyle.YesNo)
            Title = "Save Bulk BOM"

            Response = MsgBox(Msg, CDbl(Style), Title)

            If Response = MsgBoxResult.Yes Then
                Kill((BomListFileName))
                'Excel.Application.ActiveWorkbook.SaveAs(FileName:=BomListFileName, FileFormat:=xlNormal, Password:="", WriteResPassword:="", ReadOnlyRecommended:=False, CreateBackup:=False, AddToMru:=True)
                Workbooks.Application.ActiveWorkbook.SaveAs(Filename:=BomListFileName, FileFormat:=XlFileFormat.xlWorkbookNormal, Password:="", WriteResPassword:="", ReadOnlyRecommended:=False, CreateBackup:=False, AddToMru:=True)
            ElseIf Response = MsgBoxResult.No Then
                SaveAsFilename.Show()
                If PassFilename <> vbNullString And ReadyToContinue = True Then
                    If Right(PassFilename, 4) = ".xls" Then
                        BomListFileName = FileDir & PassFilename
                        Excel.Application.ActiveWorkbook.SaveAs(FileName:=BomListFileName, FileFormat:=xlNormal, Password:="", WriteResPassword:="", ReadOnlyRecommended:=False, CreateBackup:=False, AddToMru:=True)
                    Else
                        BomListFileName = FileDir & PassFilename & ".xls"
                        Excel.Application.ActiveWorkbook.SaveAs(FileName:=BomListFileName, FileFormat:=xlNormal, Password:="", WriteResPassword:="", ReadOnlyRecommended:=False, CreateBackup:=False, AddToMru:=True)
                    End If
                ElseIf PassFilename = "CancelProgram" And ReadyToContinue = False Then
                    Exit Function
                Else
                    GoTo CheckFileName
                End If
            End If
        Else
            Workbooks.Application.ActiveWorkbook.SaveAs(Filename:=BomListFileName, FileFormat:=XlFileFormat.xlWorkbookNormal, Password:="", WriteResPassword:="", ReadOnlyRecommended:=False, CreateBackup:=False, AddToMru:=True)
        End If

Err_CopyBOMFile:
        ErrNo = Err.Number

        If ErrNo <> 0 Then
            PrgName = "CopyBOMFile"
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


            If GenInfo.UserName = "dlong" Then
                MsgBox(ErrMsg)
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

    Function CopyShipListFile()
        Dim FileDir, JobNo, ShipListRev, ShipListFileName As String
        Dim ShippingMnu As ShippingList_Menu
        ShippingMnu = ShippingMnu

        FileDir = ShippingMnu.ComboBox1.Text & "\"

        'With Worksheets("Shipping List")
        '    JobNo = .Range("E3").Value
        'End With

        ShipListRev = ShippingMnu.ComboBox1.Text

        If ShipListType = "TANK" Then
            ShipListFileName = FileDir & JobNo & "-SHIPLIST-R" & ShipListRev & ".xls"
        Else
            ShipListFileName = FileDir & JobNo & "-SEALSHIPLIST-R" & ShipListRev & ".xls"
        End If

        'Sheets("Shipping List").Select()
        'Sheets("Shipping List").Copy()

        NewShipListSht = ExcelApp.Application.ActiveWorkbook.Sheets("Shipping List")

CheckFileName:
        If Dir(ShipListFileName) <> vbNullString Then
            Dim Msg, Style, Title As String
            Dim Response As Object

            ShippingMnu.Hide()                 'MainMenu.Hide()
            Msg = "File " & ShipListFileName & " already exists. Do you want to overwrite it?"
            Style = vbYesNo
            Title = "Save Shipping List"

            Response = MsgBox(Msg, Style, Title)

            If Response = vbYes Then
                Kill(ShipListFileName)
                ExcelApp.Application.ActiveWorkbook.SaveAs(FileName:=ShipListFileName, FileFormat:=xlNormal, Password:="", WriteResPassword:="", ReadOnlyRecommended:= _
                    False, CreateBackup:=False, AddToMru:=True)
            ElseIf Response = vbNo Then
                SaveAsFilename.Show()
                If PassFilename <> vbNullString And ReadyToContinue = True Then
                    If Right(PassFilename, 4) = ".xls" Then
                        ShipListFileName = FileDir & PassFilename
                        ExcelApp.Application.ActiveWorkbook.SaveAs(FileName:=ShipListFileName, FileFormat _
                            :=xlNormal, Password:="", WriteResPassword:="", ReadOnlyRecommended:= _
                            False, CreateBackup:=False, AddToMru:=True)
                    Else
                        ShipListFileName = FileDir & PassFilename & ".xls"
                        ExcelApp.Application.ActiveWorkbook.SaveAs(FileName:=ShipListFileName, FileFormat _
                            :=xlNormal, Password:="", WriteResPassword:="", ReadOnlyRecommended:= _
                            False, CreateBackup:=False, AddToMru:=True)
                    End If
                ElseIf PassFilename = "CancelProgram" And ReadyToContinue = False Then
                    Exit Function
                Else
                    GoTo CheckFileName
                End If
            End If
        Else
            ExcelApp.Application.ActiveWorkbook.SaveAs(FileName:=ShipListFileName, FileFormat _
                :=xlNormal, Password:="", WriteResPassword:="", ReadOnlyRecommended:= _
                False, CreateBackup:=False, AddToMru:=True)
        End If

        'Excel.Application.ActiveWorkbook.Close True
    End Function

    Sub OpenExcelFile(ByRef FileToOpen As String)
        Dim xlMinimized, ExcelApp, ExcelWorkbook As Object

        On Error Resume Next

        ExcelApp = GetObject(, "Excel.Application")

        If Err.Number Then
            Information.Err.Clear()
            ExcelApp = CreateObject("Excel.Application")
            If Err.Number Then
                MsgBox(Err.Description)
                Exit Sub
            End If
        End If

        ExcelWorkbook = ExcelApp.Application.Workbooks.Open(FileToOpen)

        ExcelApp.Application.WindowState = xlMinimized

    End Sub

    Function ShopCUT(ByVal RevNo As String) As Object
        Dim xlGuess As Object, ActiveSheet As Object, Range As Object, Worksheets As Object, myarray As Object
        Dim myarray2 As Object
        'Dim PurchaseArray As Object, PurchaseArray2 As Object
        Dim PIPE As Object, L As Object, W As Object, C As Object, TUBE As Object, PL As Object, SHT As Object
        Dim RB As Object, FB As Object, Purchase As Object
        Dim Fnd, Desc, Ref, Leng, ROD, PipePos, ShellPos, SeeDwgPos, Plug, LenPart, LenRbSize As Integer
        Dim LenPipeSize, LenMaterial As Integer
        Dim FoundSch, FoundStd, FoundX, FoundInch, FoundFoot, FoundFraction, LenMatPart, PartFoundX As Integer
        Dim FoundDash, FoundSpace, LenPipe, FootPos, FootPos2, FootPos3, LenFoot As Integer
        Dim LenPlateNotes, FoundNote As Integer
        Dim iA As Object
        Dim jA As Short
        Dim FoundLast As Boolean
        Dim LineNo As Short
        Dim SheetToUse As Object
        Dim CutCount As Short, Count As Short
        Dim FullJobNo, FileToOpen, PrgName, PriPrg, ErrNo, ErrMsg, ErrSource, ErrDll, ErrLastLineX, WorkShtName As String
        Dim MatType, PlateRowNo, StickRowNo, GratingRowNo, PurchaseRowNo, ItemFound, AddItem, Test As String
        Dim SearchPipe, SearchX, SearchInch, SearchSch, SearchStd, SearchShell, SearchDwg, SearchFoot As String
        Dim SearchFraction, SearchDash, SearchNote, PipeTest, Test1, FindFoot, FindFoot1, FindFoot2 As String
        Dim PipeSize, PipeLen, PipeSch, Material, MaterialPart, MaterialPart2, MatLength, RbSize As String
        Dim DecInches, SearchSpace, PurchaseItem, GetMatType, GetAngle, PipeFirstPart, PipeSecondPart As String
        Dim GetGrating, PlateNotes As String
        Dim FeetLength, FeetTotal, InchFraction, FstFraction, SndFraction, DecFraction As Double
        'Dim MatInch As Double                          'Now on Misc sheet as a public share....
        Dim i As Short, i2 As Short
        Dim ErrException As System.Exception            'System.EventHandler 
        Dim BOMMnu As ShippingList_Menu
        BOMMnu = ShippingList_Menu                           'Main_Menu = Main_Menu
        Dim BOMWrkSht As Worksheet
        Dim ShopCutWrkSht As Worksheet
        Dim PlateWrkSht As Worksheet
        Dim StickWrkSht As Worksheet
        Dim GratingWrkSht As Worksheet
        Dim PurchaseWrkSht As Worksheet
        Dim WorkSht As Worksheet
        Dim Workbooks As Excel.Workbooks
        Dim PurchaseProb As Boolean
        Dim ProbPart As String

        PrgName = "ShopCUT"

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

        On Error GoTo Err_ShopCUT

        Workbooks = ExcelApp.Workbooks
        WorkShtName = "Bulk BOM"
        BOMWrkSht = Workbooks.Application.Worksheets(WorkShtName)
        WorkSht = Workbooks.Application.ActiveSheet
        WorkShtName = WorkSht.Name

        'WorkShtName = "Shop Cut BOM"
        'ShopCutWrkSht = Workbooks.Application.Worksheets(WorkShtName)

        'NewSheet
        'Worksheets("Bulk BOM").Activate()
        With BOMWrkSht
            FullJobNo = .Range("C3").Value                          'FullJobNo = .Range("B3").Value
            FoundLast = False
            LineNo = 4
            CutCount = 1
            Count = 1
            Do Until FoundLast = True
                LineNo = LineNo + 1
                If .Range("G" & LineNo).Value = "" Then              'If .Range("B" & LineNo).Value = "" Then
                    LineNo = LineNo - 1
                    FoundLast = True
                Else                                                'Remove Double foot Marks '' use inch instead
CheckAgainFoot:
                    FindFoot = Nothing
                    FindFoot1 = Nothing
                    FindFoot2 = Nothing
                    FootPos = 0
                    FootPos2 = 0
                    FootPos3 = 0
                    SearchFoot = Chr(39)
                    FindFoot = .Range("G" & LineNo).Value
                    LenFoot = Len(FindFoot)
                    FootPos = InStr(1, FindFoot, SearchFoot)

                    If FootPos > 0 Then
                        FootPos2 = InStr((FootPos + 1), FindFoot, SearchFoot)
                    End If

                    If FootPos2 > 0 Then
                        FootPos3 = InStr((FootPos2 + 1), FindFoot, SearchFoot)
                    End If

                    If FootPos2 = (FootPos + 1) Then
                        FindFoot1 = Mid(FindFoot, 1, (FootPos - 1))
                        FindFoot2 = Mid(FindFoot, (FootPos + 2), (LenFoot - (FootPos + 1)))
                        Test = FindFoot1 & Chr(34) & FindFoot2
                        .Range("G" & LineNo).Value = FindFoot1 & Chr(34) & FindFoot2
                        GoTo CheckAgainFoot
                    Else
                        If FootPos3 = (FootPos2 + 1) Then
                            FindFoot1 = Mid(FindFoot, 1, (FootPos2 - 1))
                            If FootPos3 = LenFoot Then
                                'Do Nothing
                            Else
                                FindFoot2 = Mid(FindFoot, (FootPos2 + 2), (LenFoot - (FootPos2 + 1)))
                            End If
                            'Test = FindFoot1 & Chr(34) & FindFoot2
                            .Range("G" & LineNo).Value = FindFoot1 & Chr(34) & FindFoot2
                            GoTo CheckAgainFoot
                        End If
                    End If
                End If
            Loop

            ReDim myarray(14, LineNo - 4)
            ReDim myarray2(14, CutCount)
            For iA = 5 To LineNo
                'If iA = 10 Then
                '    Stop
                'End If
                For jA = 1 To 14
                    myarray(jA, iA - 4) = .Range(Chr(jA + 64) & iA).Value        'myarray(jA, iA - 4) = ActiveSheet.Range(Chr(jA + 64) & iA).Value
                    Test = .Range(Chr(jA + 64) & iA).Value
                Next jA
            Next iA

            '            Dim i As Short
            '            Dim i2 As Short
            ItemFound = "No"
            SearchX = " x "
            SearchInch = Chr(34)                'Find " Inch
            SearchFoot = Chr(39)                'Find ' Feet
            SearchSch = "SCH"
            SearchStd = "STD"
            SearchFraction = "/"
            SearchDash = "-"
            SearchSpace = " "

            BOMMnu.ProgressBar1.Maximum = iA                'Add Progressbar info
            For i = 1 To (iA - 5)                           'Looping structure to look at array.
                'Test = (myarray(6, i))                     'Just looking for issues
                'PIPE = InStr(1, Test, "SHELL")
                'If PIPE > 0 Then
                '    Stop
                'End If

                BOMMnu.ProgressBar1.Value = i

                InchFraction = 0
                DecFraction = 0
                FeetTotal = 0
                FoundFraction = 0
                FstFraction = 0
                SndFraction = 0
                FoundInch = 0
                FeetLength = 0
                Test1 = (myarray(8, i))                     'Found problem were " " was found
                If Test1 = " " Then
                    GoTo FindPart
                End If
                Test = (myarray(7, i))

                'Dim FindProb
                'FindProb = "NIPPLE"                         'Program Errored out on ShopCut
                'If InStr(Test, FindProb) > 0 Then
                '    'Stop
                '    GoTo PurchaseItem
                'End If

                If IsNothing(myarray(8, i)) Then            'If IsNothing(myarray(7, i)) Then
FindPart:           PIPE = 0
                    W = 0
                    C = 0
                    TUBE = 0
                    FB = 0
                    ROD = 0
                    Plug = 0
                    SHT = 0
                    L = 0
                    RB = 0
                    PurchaseItem = 0
                    PL = 0
                    PipeTest = Nothing
                    GetMatType = Left(myarray(7, i), 1)

                    Select Case UCase(GetMatType)
                        Case "U"                '------------------------------U Bolts
                            If InStr(Left(myarray(7, i), 6), "U-BOLT") = 1 Then
                                GoTo PurchaseItem
                            Else
                                GoTo PurchaseItem
                            End If
                        Case "P"                '------------------------------Pipe
                            Select Case Left(myarray(7, i), 6)
                                Case "PIPE 1"
                                    GoTo PipeItem
                                Case "PIPE 2"
                                    GoTo PipeItem
                                Case "PIPE 3"
                                    GoTo PipeItem
                                Case "PIPE 4"
                                    GoTo PipeItem
                                Case "PIPE 5"
                                    GoTo PipeItem
                                Case "PIPE 6"
                                    GoTo PipeItem
                                Case "PIPE 7"
                                    GoTo PipeItem
                                Case "PIPE 8"
                                    GoTo PipeItem
                                Case "PIPE 9"
                                    GoTo PipeItem
                                Case Else
                                    Test = (myarray(7, i))      'Found issue were user listed pipe as 20" SCH 10 PIPE X 8"
                                    PIPE = InStr(1, Test, "PIPE")
                                    If PIPE > 0 Then
                                        'Stop     'Removed stop           'May need to reprogram some of the Pipe from Sch 20 Pipe to STD.
                                        GoTo PipeItem
                                    End If

                                    Plug = InStr(1, myarray(7, i), "PLUG")      'Found issue were Plug was added to plate steel
                                    If Plug > 0 Then
                                        GoTo PurchaseItem
                                    Else
                                        PL = InStr(Left(myarray(7, i), 2), "PL")
                                        If PL > 0 Then
                                            GoTo FlatPlateItems
                                        Else
                                            GoTo PurchaseItem
                                        End If
                                    End If
                            End Select

                        Case "W"                '------------------------------Beams
                            Select Case Left(myarray(7, i), 2)
                                Case "W "
                                    GoTo BeamItem
                                Case "W1"
                                    GoTo BeamItem
                                Case "W2"
                                    GoTo BeamItem
                                Case "W3"
                                    GoTo BeamItem
                                Case "W4"
                                    GoTo BeamItem
                                Case "W5"
                                    GoTo BeamItem
                                Case "W6"
                                    GoTo BeamItem
                                Case "W7"
                                    GoTo BeamItem
                                Case "W8"
                                    GoTo BeamItem
                                Case "W9"
                                    GoTo BeamItem
                                Case Else
                                    GoTo PurchaseItem
                            End Select
                        Case "S"                '------------------------------S Beams
                            Select Case Left(myarray(7, i), 2)
                                Case "S "                                   'S BEAMS 
                                    GoTo BeamItem
                                Case "S1"
                                    GoTo BeamItem
                                Case "S2"
                                    GoTo BeamItem
                                Case "S3"
                                    GoTo BeamItem
                                Case "S4"
                                    GoTo BeamItem
                                Case "S5"
                                    GoTo BeamItem
                                Case "S6"
                                    GoTo BeamItem
                                Case "S7"
                                    GoTo BeamItem
                                Case "S8"
                                    GoTo BeamItem
                                Case "S9"
                                    GoTo BeamItem
                                Case "SH"
                                    SHT = InStr(Left(myarray(7, i), 3), "SHT")
                                    If SHT > 0 Then
                                        GoTo SheetItems
                                    Else
                                        GoTo PurchaseItem
                                    End If
                                Case Else
                                    GoTo PurchaseItem
                            End Select
                        Case "M"                '----------------------------------M Beams
                            Select Case Left(myarray(7, i), 2)
                                Case "M "                               'M BEAMs
                                    GoTo BeamItem
                                Case "M1"
                                    GoTo BeamItem
                                Case "M2"
                                    GoTo BeamItem
                                Case "M3"
                                    GoTo BeamItem
                                Case "M4"
                                    GoTo BeamItem
                                Case "M5"
                                    GoTo BeamItem
                                Case "M6"
                                    GoTo BeamItem
                                Case "M7"
                                    GoTo BeamItem
                                Case "M8"
                                    GoTo BeamItem
                                Case "M9"
                                    GoTo BeamItem
                                Case "MC"           '---------------------------MC CHANNELS
                                    GoTo ChannelItems
                                Case Else
                                    GoTo PurchaseItem
                            End Select
                        Case "C"                '-------------------------------------Channels
                            Select Case Left(myarray(7, i), 2)
                                Case "C "                           'Channels
                                    GoTo ChannelItems
                                Case "C1"
                                    GoTo ChannelItems
                                Case "C2"
                                    GoTo ChannelItems
                                Case "C3"
                                    GoTo ChannelItems
                                Case "C4"
                                    GoTo ChannelItems
                                Case "C5"
                                    GoTo ChannelItems
                                Case "C6"
                                    GoTo ChannelItems
                                Case "C7"
                                    GoTo ChannelItems
                                Case "C8"
                                    GoTo ChannelItems
                                Case "C9"
                                    GoTo ChannelItems
                                Case Else
                                    GoTo PurchaseItem
                            End Select

                        Case "L"                    '---------------------------Angle Material
                            GetAngle = Left(myarray(7, i), 2)   'Many Issues Must Hunt for L2 and L 2 And add Purchase Items

                            Select Case GetAngle
                                Case "L "
                                    GoTo AngleItems
                                Case "L1"
                                    GoTo AngleItems
                                Case "L2"
                                    GoTo AngleItems
                                Case "L3"
                                    GoTo AngleItems
                                Case "L4"
                                    GoTo AngleItems
                                Case "L5"
                                    GoTo AngleItems
                                Case "L6"
                                    GoTo AngleItems
                                Case "L7"
                                    GoTo AngleItems
                                Case "L8"
                                    GoTo AngleItems
                                Case "L9"
                                    GoTo AngleItems
                                Case Else
                                    GoTo PurchaseItem
                            End Select

                        Case "G"                    '---------------------------Grating Material
                            GetGrating = Left(myarray(7, i), 12)   'Many Issues Must Hunt for Grating and Grating Straps, and add to Purchase Items

                            Select Case GetGrating                ' , and add to Purchase Items.
                                Case "GRATING CLIP"
                                    GoTo PurchaseItem
                                Case "GRATING STRU"
                                    GoTo PurchaseItem
                                Case Else
                                    GetGrating = Left(myarray(7, i), 7)

                                    If GetGrating = "GRATING" Then
                                        GoTo GratingItem
                                    Else
                                        GoTo PurchaseItem
                                    End If
                            End Select

                        Case "T"                '------------------------------Tubbing
                            TUBE = InStr(Left(myarray(7, i), 4), "TUBE")
                            If TUBE > 0 Then
                                LenMaterial = Len(myarray(7, i))
                                LenPipeSize = InStr(1, myarray(7, i), " ")
                                PipeSize = Mid(myarray(7, i), 1, 7)
                                myarray(14, i).Value = PipeSize
                                LenPart = Len(myarray(7, i))
                                PipeLen = Mid(myarray(7, i), 7, LenPart)
                                FoundX = InStr(1, PipeLen, SearchX)
                                While SearchX > 0
                                    PipeLen = Mid(myarray(7, i), FoundX, LenPart)
                                    FoundX = InStr(1, PipeLen, SearchX)
                                    FoundInch = InStr(1, PipeLen, SearchInch)
                                    FoundSch = InStr(1, PipeLen, SearchSch)
                                    FoundStd = InStr(1, PipeLen, SearchStd)
                                End While
                                If FoundStd > 0 Then
                                    PipeSch = Mid(myarray(7, i), FoundStd, 6)
                                    myarray(12, i).Value = PipeSch
                                End If
                                If FoundSch > 0 Then
                                    PipeSch = Mid(myarray(7, i), FoundSch, 6)
                                    myarray(12, i).Value = PipeSch
                                End If
                                If FoundInch > 0 Then
                                    PipeLen = Mid(myarray(7, i), 1, FoundInch)
                                    myarray(13, i).Value = PipeLen
                                End If
                                '----------------------------------
                                myarray(14, i) = "TUBE"
                                For i2 = 1 To 14
                                    ReDim Preserve myarray2(14, CutCount)
                                    myarray2(i2, Count) = myarray(i2, i)
                                Next i2
                                CutCount = CutCount + 1
                                Count = Count + 1
                                ItemFound = "Yes"
                                GoTo NextItem
                            Else
                                GoTo PurchaseItem
                            End If

                        Case "R"
                            RB = InStr(Left(myarray(7, i), 2), "RB")
                            If RB > 0 Then
                                LenMaterial = Len(myarray(7, i))
                                Material = myarray(7, i)
                                Material = LTrim(Material)
                                Material = RTrim(Material)

                                FoundX = InStr(1, Material, SearchX)
                                MaterialPart = Material

                                While FoundX > 0
                                    LenMatPart = Len(MaterialPart)
                                    MaterialPart = Mid(MaterialPart, (FoundX + 3), LenMatPart)
                                    PartFoundX = FoundX
                                    FoundX = InStr(1, MaterialPart, SearchX)
                                End While

                                LenMatPart = Len(MaterialPart)
                                MaterialPart = LTrim(MaterialPart)
                                MaterialPart = RTrim(MaterialPart)
                                MaterialPart2 = Mid(Material, 1, (LenMaterial - (LenMatPart + 3)))
                                MaterialPart2 = LTrim(MaterialPart2)
                                MaterialPart2 = RTrim(MaterialPart2)

                                If Len(MaterialPart2) > 0 Then
                                    myarray(12, i) = MaterialPart2
                                End If

                                FeetInchesToDecInches(MaterialPart)
                                myarray(13, i) = MatInch
                                myarray(14, i) = "ROUND BAR"
                                For i2 = 1 To 14
                                    ReDim Preserve myarray2(14, CutCount)
                                    myarray2(i2, Count) = myarray(i2, i)
                                Next i2
                                CutCount = CutCount + 1
                                Count = Count + 1
                                ItemFound = "Yes"
                                GoTo NextItem
                            End If

                            Test = (myarray(7, i))      'Found issue were user listed ROUND BAR as ROD 3/8" X 2'-10 1/4" (COLD ROLLED)
                            ROD = InStr(1, myarray(7, i), "ROD")
                            If ROD > 0 Then
                                LenMaterial = Len(myarray(7, i))
                                Material = myarray(7, i)
                                Material = LTrim(Material)
                                Material = RTrim(Material)

                                FoundX = InStr(1, Material, SearchX)
                                MaterialPart = Material

                                While FoundX > 0
                                    LenMatPart = Len(MaterialPart)
                                    MaterialPart = Mid(MaterialPart, (FoundX + 3), LenMatPart)
                                    PartFoundX = FoundX
                                    FoundX = InStr(1, MaterialPart, SearchX)
                                End While

                                LenMatPart = Len(MaterialPart)
                                MaterialPart = LTrim(MaterialPart)
                                MaterialPart = RTrim(MaterialPart)
                                MaterialPart2 = Mid(Material, 1, (LenMaterial - (LenMatPart + 3)))
                                MaterialPart2 = LTrim(MaterialPart2)
                                MaterialPart2 = RTrim(MaterialPart2)

                                If Len(MaterialPart2) > 0 Then
                                    myarray(12, i) = MaterialPart2
                                End If

                                FeetInchesToDecInches(MaterialPart)
                                myarray(13, i) = MatInch
                                myarray(14, i) = "ROUND BAR"
                                For i2 = 1 To 14
                                    ReDim Preserve myarray2(14, CutCount)
                                    myarray2(i2, Count) = myarray(i2, i)
                                Next i2
                                CutCount = CutCount + 1
                                Count = Count + 1
                                ItemFound = "Yes"
                                GoTo NextItem
                            Else
                                GoTo PurchaseItem
                            End If

                        Case "F"

                            FB = InStr(Left(myarray(7, i), 2), "FB")
                            If FB > 0 Then
                                '----------------------------------------------New Part
                                LenMaterial = Len(myarray(7, i))
                                Material = myarray(7, i)
                                Material = LTrim(Material)
                                Material = RTrim(Material)

                                FoundX = InStr(1, Material, SearchX)
                                MaterialPart = Material

                                While FoundX > 0
                                    LenMatPart = Len(MaterialPart)
                                    MaterialPart = Mid(MaterialPart, (FoundX + 3), LenMatPart)
                                    PartFoundX = FoundX
                                    FoundX = InStr(1, MaterialPart, SearchX)
                                End While

                                LenMatPart = Len(MaterialPart)
                                MaterialPart = LTrim(MaterialPart)
                                MaterialPart = RTrim(MaterialPart)
                                MaterialPart2 = Mid(Material, 1, (LenMaterial - (LenMatPart + 3)))         'MaterialPart2 = Mid(Material, 1, (LenMaterial - (PartFoundX + LenMatPart - 1)))
                                MaterialPart2 = LTrim(MaterialPart2)
                                MaterialPart2 = RTrim(MaterialPart2)

                                If Len(MaterialPart2) > 0 Then
                                    myarray(12, i) = MaterialPart2
                                End If

                                FeetInchesToDecInches(MaterialPart)
                                myarray(13, i) = MatInch
                                myarray(14, i) = "FLAT BAR"
                                For i2 = 1 To 14
                                    ReDim Preserve myarray2(14, CutCount)
                                    myarray2(i2, Count) = myarray(i2, i)
                                Next i2
                                CutCount = CutCount + 1
                                Count = Count + 1
                                ItemFound = "Yes"
                                GoTo NextItem
                            Else
                                GoTo PurchaseItem
                            End If
                        Case Else
                            Test = (myarray(7, i))      'Found issue were user listed pipe as 20" SCH 10 PIPE X 8"
                            PIPE = InStr(1, Test, "PIPE")
                            If PIPE > 0 Then
                                'Stop                'May need to reprogram some of the Pipe from Sch 20 Pipe to STD.
                                LenPipe = Len(Test)
                                PipeFirstPart = Mid(Test, 1, (PIPE - 1))
                                PipeSecondPart = Mid(Test, (PIPE + 5), (LenPipe - (PIPE + 5)))
                                'Stop
                                PipeTest = "PIPE " & PipeFirstPart & PipeSecondPart
                                GoTo PipeItem
                            Else
                                GoTo PurchaseItem
                            End If
                    End Select

                    '----------------------------------
SheetItems:
                    LenMaterial = Len(myarray(7, i))
                    Material = myarray(7, i)
                    Material = LTrim(Material)
                    Material = RTrim(Material)

                    FoundX = InStr(1, Material, SearchX)
                    MaterialPart = Material

                    While FoundX > 0
                        LenMatPart = Len(MaterialPart)
                        MaterialPart = Mid(MaterialPart, (FoundX + 3), LenMatPart)
                        PartFoundX = FoundX
                        FoundX = InStr(1, MaterialPart, SearchX)
                    End While

                    LenMatPart = Len(MaterialPart)
                    MaterialPart = LTrim(MaterialPart)
                    MaterialPart = RTrim(MaterialPart)
                    If LenMaterial = LenMatPart Then
                        MaterialPart2 = Material
                    Else
                        MaterialPart2 = Mid(Material, 1, (LenMaterial - (LenMatPart + 3)))
                    End If

                    '----------------------------------------Per Ken Erdmann remove width add Material type.

                    FoundX = InStr(1, MaterialPart2, SearchX)

                    While FoundX > 0
                        MaterialPart2 = Mid(MaterialPart2, 1, (FoundX - 1))
                        FoundX = InStr(1, MaterialPart2, SearchX)
                    End While

                    MaterialPart2 = LTrim(MaterialPart2)
                    MaterialPart2 = RTrim(MaterialPart2)
                    MatType = myarray(10, i)
                    MaterialPart2 = MaterialPart2 & " - " & MatType

                    If Len(MaterialPart2) > 0 Then
                        myarray(12, i) = MaterialPart2
                    End If
                    '-----------------------------------------Material Found Completed

                    LenPlateNotes = Len(myarray(7, i))        '--------------------------Find PlateNotes
                    PlateNotes = myarray(7, i)
                    SearchNote = "NOTE"

                    FoundNote = InStr(1, PlateNotes, SearchNote)

                    If FoundNote > 0 Then
                        LenPlateNotes = Len(PlateNotes)
                        PlateNotes = Mid(PlateNotes, FoundNote, (LenPlateNotes - (FoundNote - 1)))
                        PlateNotes = LTrim(PlateNotes)
                        PlateNotes = RTrim(PlateNotes)
                    Else
                        PlateNotes = ""
                    End If                          '------------------------------Notes Found Completed.

                    'FeetInchesToDecInches(MaterialPart)
                    myarray(13, i) = PlateNotes                        'myarray(13, i) = MatInch
                    myarray(14, i) = "SHEET STEEL-GA"
                    For i2 = 1 To 14
                        ReDim Preserve myarray2(14, CutCount)
                        myarray2(i2, Count) = myarray(i2, i)
                    Next i2
                    CutCount = CutCount + 1
                    Count = Count + 1
                    ItemFound = "Yes"
                    GoTo NextItem

AngleItems:
                    LenMaterial = Len(myarray(7, i))
                    Material = myarray(7, i)
                    Material = LTrim(Material)
                    Material = RTrim(Material)

                    FoundX = InStr(1, Material, SearchX)
                    MaterialPart = Material

                    While FoundX > 0
                        LenMatPart = Len(MaterialPart)
                        MaterialPart = Mid(MaterialPart, (FoundX + 3), LenMatPart)
                        PartFoundX = FoundX
                        FoundX = InStr(1, MaterialPart, SearchX)
                    End While

                    LenMatPart = Len(MaterialPart)
                    MaterialPart = LTrim(MaterialPart)
                    MaterialPart = RTrim(MaterialPart)
                    If LenMaterial = LenMatPart Then
                        MaterialPart2 = Material
                    Else
                        MaterialPart2 = Mid(Material, 1, (LenMaterial - (LenMatPart + 3)))     'MaterialPart2 = Mid(Material, 1, (LenMaterial - (PartFoundX + LenMatPart - 1)))
                    End If
                    MaterialPart2 = LTrim(MaterialPart2)
                    MaterialPart2 = RTrim(MaterialPart2)

                    If Len(MaterialPart2) > 0 Then
                        myarray(12, i) = MaterialPart2
                    End If

                    FeetInchesToDecInches(MaterialPart)           'MatInch = FeetInchesToDecInches(MaterialPart)
                    myarray(13, i) = MatInch
                    '----------------------------------
                    myarray(14, i) = "ANGLE"
                    For i2 = 1 To 14
                        ReDim Preserve myarray2(14, CutCount)
                        myarray2(i2, Count) = myarray(i2, i)
                    Next i2
                    CutCount = CutCount + 1
                    Count = Count + 1
                    ItemFound = "Yes"
                    GoTo NextItem

FlatPlateItems:
                    LenMaterial = Len(myarray(7, i))        '--------Find Material Information
                    Material = myarray(7, i)
                    Material = LTrim(Material)
                    Material = RTrim(Material)

                    FoundX = InStr(1, Material, SearchX)
                    MaterialPart = Material

                    While FoundX > 0
                        LenMatPart = Len(MaterialPart)
                        MaterialPart = Mid(MaterialPart, (FoundX + 3), LenMatPart)
                        PartFoundX = FoundX
                        FoundX = InStr(1, MaterialPart, SearchX)
                    End While

                    LenMatPart = Len(MaterialPart)
                    MaterialPart = LTrim(MaterialPart)
                    MaterialPart = RTrim(MaterialPart)
                    If LenMaterial = LenMatPart Then
                        MaterialPart2 = Material
                    Else
                        MaterialPart2 = Mid(Material, 1, (LenMaterial - (LenMatPart + 3)))
                    End If

                    '----------------------------------------Per Ken Erdmann remove width add Material type.

                    FoundX = InStr(1, MaterialPart2, SearchX)

                    While FoundX > 0
                        MaterialPart2 = Mid(MaterialPart2, 1, (FoundX - 1))
                        FoundX = InStr(1, MaterialPart2, SearchX)
                    End While

                    MaterialPart2 = LTrim(MaterialPart2)
                    MaterialPart2 = RTrim(MaterialPart2)
                    MatType = myarray(10, i)
                    MaterialPart2 = MaterialPart2 & " - " & MatType

                    If Len(MaterialPart2) > 0 Then
                        myarray(12, i) = MaterialPart2
                    End If
                    '-----------------------------------------Material Found Completed

                    LenPlateNotes = Len(myarray(7, i))        '--------------------------Find PlateNotes
                    PlateNotes = myarray(7, i)
                    SearchNote = "NOTE"

                    FoundNote = InStr(1, PlateNotes, SearchNote)

                    If FoundNote > 0 Then
                        LenPlateNotes = Len(PlateNotes)
                        PlateNotes = Mid(PlateNotes, FoundNote, (LenPlateNotes - (FoundNote - 1)))
                        PlateNotes = LTrim(PlateNotes)
                        PlateNotes = RTrim(PlateNotes)
                    Else
                        PlateNotes = ""
                    End If                          '------------------------------Notes Found Completed.

                    'FeetInchesToDecInches(MaterialPart)
                    myarray(13, i) = PlateNotes                        'myarray(13, i) = MatInch
                    myarray(14, i) = "PLATE STEEL"
                    For i2 = 1 To 14
                        ReDim Preserve myarray2(14, CutCount)
                        myarray2(i2, Count) = myarray(i2, i)
                    Next i2
                    CutCount = CutCount + 1
                    Count = Count + 1
                    ItemFound = "Yes"
                    GoTo NextItem

ChannelItems:       LenMaterial = Len(myarray(7, i))
                    Material = myarray(7, i)
                    Material = LTrim(Material)
                    Material = RTrim(Material)

                    FoundX = InStr(1, Material, SearchX)
                    MaterialPart = Material

                    While FoundX > 0
                        LenMatPart = Len(MaterialPart)
                        MaterialPart = Mid(MaterialPart, (FoundX + 3), LenMatPart)
                        PartFoundX = FoundX
                        FoundX = InStr(1, MaterialPart, SearchX)
                    End While

                    LenMatPart = Len(MaterialPart)
                    MaterialPart = LTrim(MaterialPart)
                    MaterialPart = RTrim(MaterialPart)
                    MaterialPart2 = Mid(Material, 1, (LenMaterial - (LenMatPart + 3)))
                    MaterialPart2 = LTrim(MaterialPart2)
                    MaterialPart2 = RTrim(MaterialPart2)

                    If Len(MaterialPart2) > 0 Then
                        myarray(12, i) = MaterialPart2
                    End If

                    FeetInchesToDecInches(MaterialPart)
                    myarray(13, i) = MatInch
                    myarray(14, i) = "CHANNEL"
                    For i2 = 1 To 14
                        ReDim Preserve myarray2(14, CutCount)
                        myarray2(i2, Count) = myarray(i2, i)
                    Next i2
                    CutCount = CutCount + 1
                    Count = Count + 1
                    ItemFound = "Yes"
                    GoTo NextItem

BeamItem:           LenMaterial = Len(myarray(7, i))
                    LenPipeSize = InStr(1, myarray(7, i), " ")
                    PipeSize = Mid(myarray(7, i), 1, 7)
                    myarray(14, i).Value = PipeSize
                    LenPart = Len(myarray(7, i))
                    PipeLen = Mid(myarray(7, i), 7, LenPart)
                    FoundX = InStr(1, PipeLen, SearchX)
                    While SearchX > 0
                        PipeLen = Mid(myarray(7, i), FoundX, LenPart)
                        FoundX = InStr(1, PipeLen, SearchX)
                        FoundInch = InStr(1, PipeLen, SearchInch)
                        FoundSch = InStr(1, PipeLen, SearchSch)
                        FoundStd = InStr(1, PipeLen, SearchStd)
                    End While
                    If FoundStd > 0 Then
                        PipeSch = Mid(myarray(7, i), FoundStd, 6)
                        myarray(12, i).Value = PipeSch
                    End If
                    If FoundSch > 0 Then
                        PipeSch = Mid(myarray(7, i), FoundSch, 6)
                        myarray(12, i).Value = PipeSch
                    End If
                    If FoundInch > 0 Then
                        PipeLen = Mid(myarray(7, i), 1, FoundInch)
                        myarray(13, i).Value = PipeLen
                    End If
                    '----------------------------------
                    myarray(14, i) = "IBEAM"
                    For i2 = 1 To 14
                        ReDim Preserve myarray2(14, CutCount)
                        myarray2(i2, Count) = myarray(i2, i)
                    Next i2
                    CutCount = CutCount + 1
                    Count = Count + 1
                    ItemFound = "Yes"
                    GoTo NextItem

PipeItem:
                    If PipeTest <> Nothing Then
                        LenMaterial = Len(PipeTest)
                        Material = PipeTest
                        myarray(7, i) = PipeTest
                    Else
                        LenMaterial = Len(myarray(7, i))
                        Material = myarray(7, i)
                    End If

                    Material = LTrim(Material)
                    Material = RTrim(Material)

                    FoundX = InStr(1, Material, SearchX)
                    MaterialPart = Material

                    While FoundX > 0
                        LenMatPart = Len(MaterialPart)
                        MaterialPart = Mid(MaterialPart, (FoundX + 3), LenMatPart)
                        PartFoundX = FoundX
                        FoundX = InStr(1, MaterialPart, SearchX)
                    End While

                    LenMatPart = Len(MaterialPart)
                    MaterialPart = LTrim(MaterialPart)
                    MaterialPart = RTrim(MaterialPart)
                    MaterialPart2 = Mid(Material, 1, (LenMaterial - (LenMatPart + 3)))
                    If PurchaseProb = True Then
                        GoTo PurchaseItem
                    End If

                    MaterialPart2 = LTrim(MaterialPart2)
                    MaterialPart2 = RTrim(MaterialPart2)

                    If Len(MaterialPart2) > 0 Then
                        myarray(12, i) = MaterialPart2
                    End If

                    FeetInchesToDecInches(MaterialPart)
                    myarray(13, i) = MatInch
                    myarray(14, i) = "PIPE"
                    For i2 = 1 To 14
                        ReDim Preserve myarray2(14, CutCount)
                        myarray2(i2, Count) = myarray(i2, i)
                    Next i2
                    CutCount = CutCount + 1
                    Count = Count + 1
                    ItemFound = "Yes"
                    GoTo NextItem

GratingItem:
                    LenMaterial = Len(myarray(7, i))
                    Material = myarray(7, i)

                    Material = LTrim(Material)
                    Material = RTrim(Material)

                    FoundX = InStr(1, Material, SearchX)
                    MaterialPart = Material

                    While FoundX > 0
                        LenMatPart = Len(MaterialPart)
                        MaterialPart = Mid(MaterialPart, (FoundX + 3), LenMatPart)
                        PartFoundX = FoundX
                        FoundX = InStr(1, MaterialPart, SearchX)
                    End While

                    LenMatPart = Len(MaterialPart)
                    MaterialPart = LTrim(MaterialPart)
                    MaterialPart = RTrim(MaterialPart)
                    MaterialPart2 = Mid(Material, 1, (LenMaterial - (LenMatPart + 3)))
                    MaterialPart2 = LTrim(MaterialPart2)
                    MaterialPart2 = RTrim(MaterialPart2)

                    If Len(MaterialPart2) > 0 Then
                        myarray(12, i) = MaterialPart2
                    End If

                    FeetInchesToDecInches(MaterialPart)
                    myarray(13, i) = MatInch
                    myarray(14, i) = "GRATING"
                    For i2 = 1 To 14
                        ReDim Preserve myarray2(14, CutCount)
                        myarray2(i2, Count) = myarray(i2, i)
                    Next i2
                    CutCount = CutCount + 1
                    Count = Count + 1
                    ItemFound = "Yes"
                    GoTo NextItem

PurchaseItem:
                    myarray(14, i) = "PURCHASE ITEM"
                    For i2 = 1 To 14
                        ReDim Preserve myarray2(14, CutCount)
                        myarray2(i2, Count) = myarray(i2, i)
                    Next i2
                    CutCount = CutCount + 1
                    Count = Count + 1
                    ItemFound = "Yes"
                    PurchaseProb = False
                    GoTo NextItem

NextItem:

                    If ItemFound = "No" Then
                        Fnd = 0
                        Desc = 0
                        Ref = 0
                        SearchDwg = "(SEE DWG"
                        SearchShell = "SHELL"
                        Test = (myarray(14, i))
                        Fnd = InStr(1, myarray(14, i), "Found")
                        Desc = InStr(1, myarray(14, i), "Description")
                        Ref = InStr(1, myarray(14, i), "Reference")

                        'Test = (myarray(12, i))                         'Look for Std Reference if so do not add to Purchase Items

                        SearchShell = "Shell"
                        SearchDwg = "(See Dwg"
                        SeeDwgPos = InStr(1, myarray(7, i), SearchDwg)
                        ShellPos = InStr(1, myarray(7, i), SearchShell)
                        Test = (myarray(7, i))

                        'If SeeDwgPos > 0 Then
                        '    AddItem = "No"
                        '    'Else
                        '    '    If ShellPos > 0 Then
                        '    '        Stop
                        '    '    End If
                        'End If

                        If Fnd > 0 Then
                            AddItem = "No"
                        Else
                            If Ref > 0 Then
                                AddItem = "Yes"
                            Else
                                If Desc > 0 Then
                                    AddItem = "No"         'Standard Reference only, No additional parts.
                                Else
                                    If SeeDwgPos > 0 Then
                                        AddItem = "No"
                                    Else
                                        AddItem = "Yes"         'Add to Purchase Items if not found above
                                    End If
                                End If
                            End If
                        End If

                        If AddItem = "Yes" Then
                            myarray(14, i) = "Purchase Item"
                            For i2 = 1 To 14
                                ReDim Preserve myarray2(14, CutCount)
                                myarray2(i2, Count) = myarray(i2, i)
                            Next i2
                            CutCount = CutCount + 1
                            Count = Count + 1
                        End If
                    End If
                Else
                    If ItemFound = "No" Then
                        Fnd = 0
                        Desc = 0
                        Ref = 0
                        SearchDwg = "(SEE DWG"
                        SearchShell = "SHELL"
                        Test = (myarray(14, i))
                        Fnd = InStr(1, myarray(14, i), "Found")               'Fnd = InStr(Left(myarray(13, i), 2), "Found")
                        Desc = InStr(1, myarray(14, i), "Description")
                        Ref = InStr(1, myarray(14, i), "Reference")

                        SearchShell = "Shell"
                        SearchDwg = "(See Dwg"
                        SeeDwgPos = InStr(1, myarray(7, i), SearchDwg)
                        ShellPos = InStr(1, myarray(7, i), SearchShell)
                        Test = (myarray(7, i))

                        'If SeeDwgPos > 0 Then
                        '    AddItem = "No"
                        '    'Else
                        '    '    If ShellPos > 0 Then                    'To many Shell combinations do not look for.
                        '    '        Stop
                        '    '    End If
                        'End If

                        If Fnd > 0 Then
                            AddItem = "No"
                        Else
                            If Ref > 0 Then
                                AddItem = "Yes"
                            Else
                                If Desc > 0 Then
                                    AddItem = "No"
                                Else
                                    If SeeDwgPos > 0 Then
                                        AddItem = "No"
                                    Else
                                        AddItem = "Yes"
                                    End If
                                End If
                            End If
                        End If

                        If AddItem = "Yes" Then
                            myarray(14, i) = "Purchase Item"
                            For i2 = 1 To 14
                                ReDim Preserve myarray2(14, CutCount)
                                myarray2(i2, Count) = myarray(i2, i)
                            Next i2
                            CutCount = CutCount + 1
                            Count = Count + 1
                        End If
                    End If
                End If
                Fnd = 0
                Desc = 0
                Ref = 0
                ItemFound = "No"

            Next
        End With

        BOMMnu.ProgressBar1.Value = 0

        'WorkShtName = "Shop Cut BOM"
        'ShopCutWrkSht = Workbooks.Application.Worksheets(WorkShtName)

        WorkShtName = "Plate BOM"
        PlateWrkSht = Workbooks.Application.Worksheets(WorkShtName)

        WorkShtName = "Stick BOM"
        StickWrkSht = Workbooks.Application.Worksheets(WorkShtName)

        WorkShtName = "Grating BOM"
        GratingWrkSht = Workbooks.Application.Worksheets(WorkShtName)

        WorkShtName = "Other BOM"
        PurchaseWrkSht = Workbooks.Application.Worksheets(WorkShtName)

        i = 1
        'ShopCutWrkSht.Activate()            'Worksheets("Shop Cut BOM").Activate()

        'With ShopCutWrkSht
        '    .Range("B3").Value = FullJobNo
        '    .Range("I3").Value = Today
        '    .Range("F3").Value = RevNo                 'Main_Menu.ComboBox1.Text
        'End With

        PlateWrkSht.Activate()            'Worksheets("Shop Cut BOM").Activate()

        With PlateWrkSht
            .Range("C3").Value = FullJobNo
            .Range("J3").Value = Today
            .Range("G3").Value = RevNo                 'Main_Menu.ComboBox1.Text
        End With

        StickWrkSht.Activate()            'Worksheets("Shop Cut BOM").Activate()

        With StickWrkSht
            .Range("C3").Value = FullJobNo
            .Range("J3").Value = Today
            .Range("G3").Value = RevNo                 'Main_Menu.ComboBox1.Text
        End With

        GratingWrkSht.Activate()

        With GratingWrkSht
            .Range("C3").Value = FullJobNo
            .Range("J3").Value = Today
            .Range("G3").Value = RevNo
        End With

        PurchaseWrkSht.Activate()

        With PurchaseWrkSht
            .Range("C3").Value = FullJobNo
            .Range("J3").Value = Today
            .Range("G3").Value = RevNo
        End With

        RowNo = CStr(i + 4)
        PlateRowNo = CStr(i + 4)
        StickRowNo = CStr(i + 4)
        GratingRowNo = CStr(i + 4)
        PurchaseRowNo = CStr(i + 4)

        'Test = UBound(myarray2, 2)

        For i = 1 To Count - 1
            MatType = myarray2(14, i)
            'If i = 260 Then
            '    Stop
            '    Test = myarray2(7, i)
            'End If

            Select Case MatType
                Case "PIPE"
                    FileToOpen = "Stick BOM"
                    FormatLine2(StickRowNo, FileToOpen)
                    With StickWrkSht
                        .Range("A" & StickRowNo).Value = myarray2(1, i)             'Dwg & Ref No
                        .Range("B" & StickRowNo).Value = myarray2(2, i)             'DWG
                        .Range("C" & StickRowNo).Value = myarray2(3, i)             'Rev
                        .Range("D" & StickRowNo).Value = myarray2(4, i)             'Ship Mark
                        .Range("E" & StickRowNo).Value = myarray2(5, i)             'Piece Mark
                        .Range("F" & StickRowNo).Value = myarray2(6, i)             'QTY
                        .Range("G" & StickRowNo).Value = myarray2(7, i)             'Description
                        .Range("H" & StickRowNo).Value = myarray2(8, i)             'Inv-1
                        .Range("I" & StickRowNo).Value = myarray2(9, i)             'Inv-2 Standard Number MX1001A
                        .Range("J" & StickRowNo).Value = myarray2(10, i)             'Material
                        .Range("K" & StickRowNo).Value = myarray2(11, i)            'Weight
                        .Range("L" & StickRowNo).Value = myarray2(12, i)
                        .Range("M" & StickRowNo).Value = myarray2(13, i)            'Material Type - Pipe - Angle - ETC...
                        .Range("N" & StickRowNo).Value = myarray2(14, i)
                    End With
                    StickRowNo = CStr(CDbl(StickRowNo) + 1)
                Case "ANGLE"
                    FileToOpen = "Stick BOM"
                    FormatLine2(StickRowNo, FileToOpen)
                    With StickWrkSht
                        .Range("A" & StickRowNo).Value = myarray2(1, i)             'Dwg & Ref No
                        .Range("B" & StickRowNo).Value = myarray2(2, i)             'DWG
                        .Range("C" & StickRowNo).Value = myarray2(3, i)             'Rev
                        .Range("D" & StickRowNo).Value = myarray2(4, i)             'Ship Mark
                        .Range("E" & StickRowNo).Value = myarray2(5, i)             'Piece Mark
                        .Range("F" & StickRowNo).Value = myarray2(6, i)             'QTY
                        .Range("G" & StickRowNo).Value = myarray2(7, i)             'Description
                        .Range("H" & StickRowNo).Value = myarray2(8, i)             'Inv-1
                        .Range("I" & StickRowNo).Value = myarray2(9, i)             'Inv-2 Standard Number MX1001A
                        .Range("J" & StickRowNo).Value = myarray2(10, i)             'Material
                        .Range("K" & StickRowNo).Value = myarray2(11, i)            'Weight
                        .Range("L" & StickRowNo).Value = myarray2(12, i)
                        .Range("M" & StickRowNo).Value = myarray2(13, i)            'Material Type - Pipe - Angle - ETC...
                        .Range("N" & StickRowNo).Value = myarray2(14, i)
                    End With
                    StickRowNo = CStr(CDbl(StickRowNo) + 1)
                Case "FLAT BAR"
                    FileToOpen = "Stick BOM"
                    FormatLine2(StickRowNo, FileToOpen)
                    With StickWrkSht
                        .Range("A" & StickRowNo).Value = myarray2(1, i)             'Dwg & Ref No
                        .Range("B" & StickRowNo).Value = myarray2(2, i)             'DWG
                        .Range("C" & StickRowNo).Value = myarray2(3, i)             'Rev
                        .Range("D" & StickRowNo).Value = myarray2(4, i)             'Ship Mark
                        .Range("E" & StickRowNo).Value = myarray2(5, i)             'Piece Mark
                        .Range("F" & StickRowNo).Value = myarray2(6, i)             'QTY
                        .Range("G" & StickRowNo).Value = myarray2(7, i)             'Description
                        .Range("H" & StickRowNo).Value = myarray2(8, i)             'Inv-1
                        .Range("I" & StickRowNo).Value = myarray2(9, i)             'Inv-2 Standard Number MX1001A
                        .Range("J" & StickRowNo).Value = myarray2(10, i)             'Material
                        .Range("K" & StickRowNo).Value = myarray2(11, i)            'Weight
                        .Range("L" & StickRowNo).Value = myarray2(12, i)
                        .Range("M" & StickRowNo).Value = myarray2(13, i)            'Material Type - Pipe - Angle - ETC...
                        .Range("N" & StickRowNo).Value = myarray2(14, i)
                    End With
                    StickRowNo = CStr(CDbl(StickRowNo) + 1)
                Case "ROUND BAR"
                    FileToOpen = "Stick BOM"
                    FormatLine2(StickRowNo, FileToOpen)
                    With StickWrkSht
                        .Range("A" & StickRowNo).Value = myarray2(1, i)             'Dwg & Ref No
                        .Range("B" & StickRowNo).Value = myarray2(2, i)             'DWG
                        .Range("C" & StickRowNo).Value = myarray2(3, i)             'Rev
                        .Range("D" & StickRowNo).Value = myarray2(4, i)             'Ship Mark
                        .Range("E" & StickRowNo).Value = myarray2(5, i)             'Piece Mark
                        .Range("F" & StickRowNo).Value = myarray2(6, i)             'QTY
                        .Range("G" & StickRowNo).Value = myarray2(7, i)             'Description
                        .Range("H" & StickRowNo).Value = myarray2(8, i)             'Inv-1
                        .Range("I" & StickRowNo).Value = myarray2(9, i)             'Inv-2 Standard Number MX1001A
                        .Range("J" & StickRowNo).Value = myarray2(10, i)             'Material
                        .Range("K" & StickRowNo).Value = myarray2(11, i)            'Weight
                        .Range("L" & StickRowNo).Value = myarray2(12, i)
                        .Range("M" & StickRowNo).Value = myarray2(13, i)            'Material Type - Pipe - Angle - ETC...
                        .Range("N" & StickRowNo).Value = myarray2(14, i)
                    End With
                    StickRowNo = CStr(CDbl(StickRowNo) + 1)
                Case "TUBE"
                    FileToOpen = "Stick BOM"
                    FormatLine2(StickRowNo, FileToOpen)
                    With StickWrkSht
                        .Range("A" & StickRowNo).Value = myarray2(1, i)             'Dwg & Ref No
                        .Range("B" & StickRowNo).Value = myarray2(2, i)             'DWG
                        .Range("C" & StickRowNo).Value = myarray2(3, i)             'Rev
                        .Range("D" & StickRowNo).Value = myarray2(4, i)             'Ship Mark
                        .Range("E" & StickRowNo).Value = myarray2(5, i)             'Piece Mark
                        .Range("F" & StickRowNo).Value = myarray2(6, i)             'QTY
                        .Range("G" & StickRowNo).Value = myarray2(7, i)             'Description
                        .Range("H" & StickRowNo).Value = myarray2(8, i)             'Inv-1
                        .Range("I" & StickRowNo).Value = myarray2(9, i)             'Inv-2 Standard Number MX1001A
                        .Range("J" & StickRowNo).Value = myarray2(10, i)             'Material
                        .Range("K" & StickRowNo).Value = myarray2(11, i)            'Weight
                        .Range("L" & StickRowNo).Value = myarray2(12, i)
                        .Range("M" & StickRowNo).Value = myarray2(13, i)            'Material Type - Pipe - Angle - ETC...
                        .Range("N" & StickRowNo).Value = myarray2(14, i)
                    End With
                    StickRowNo = CStr(CDbl(StickRowNo) + 1)
                Case "IBEAM"
                    FileToOpen = "Stick BOM"
                    FormatLine2(StickRowNo, FileToOpen)
                    With StickWrkSht
                        .Range("A" & StickRowNo).Value = myarray2(1, i)             'Dwg & Ref No
                        .Range("B" & StickRowNo).Value = myarray2(2, i)             'DWG
                        .Range("C" & StickRowNo).Value = myarray2(3, i)             'Rev
                        .Range("D" & StickRowNo).Value = myarray2(4, i)             'Ship Mark
                        .Range("E" & StickRowNo).Value = myarray2(5, i)             'Piece Mark
                        .Range("F" & StickRowNo).Value = myarray2(6, i)             'QTY
                        .Range("G" & StickRowNo).Value = myarray2(7, i)             'Description
                        .Range("H" & StickRowNo).Value = myarray2(8, i)             'Inv-1
                        .Range("I" & StickRowNo).Value = myarray2(9, i)             'Inv-2 Standard Number MX1001A
                        .Range("J" & StickRowNo).Value = myarray2(10, i)             'Material
                        .Range("K" & StickRowNo).Value = myarray2(11, i)            'Weight
                        .Range("L" & StickRowNo).Value = myarray2(12, i)
                        .Range("M" & StickRowNo).Value = myarray2(13, i)            'Material Type - Pipe - Angle - ETC...
                        .Range("N" & StickRowNo).Value = myarray2(14, i)
                    End With
                    StickRowNo = CStr(CDbl(StickRowNo) + 1)
                Case "CHANNEL"
                    FileToOpen = "Stick BOM"
                    FormatLine2(StickRowNo, FileToOpen)
                    With StickWrkSht
                        .Range("A" & StickRowNo).Value = myarray2(1, i)             'Dwg & Ref No
                        .Range("B" & StickRowNo).Value = myarray2(2, i)             'DWG
                        .Range("C" & StickRowNo).Value = myarray2(3, i)             'Rev
                        .Range("D" & StickRowNo).Value = myarray2(4, i)             'Ship Mark
                        .Range("E" & StickRowNo).Value = myarray2(5, i)             'Piece Mark
                        .Range("F" & StickRowNo).Value = myarray2(6, i)             'QTY
                        .Range("G" & StickRowNo).Value = myarray2(7, i)             'Description
                        .Range("H" & StickRowNo).Value = myarray2(8, i)             'Inv-1
                        .Range("I" & StickRowNo).Value = myarray2(9, i)             'Inv-2 Standard Number MX1001A
                        .Range("J" & StickRowNo).Value = myarray2(10, i)             'Material
                        .Range("K" & StickRowNo).Value = myarray2(11, i)            'Weight
                        .Range("L" & StickRowNo).Value = myarray2(12, i)
                        .Range("M" & StickRowNo).Value = myarray2(13, i)            'Material Type - Pipe - Angle - ETC...
                        .Range("N" & StickRowNo).Value = myarray2(14, i)
                    End With
                    StickRowNo = CStr(CDbl(StickRowNo) + 1)
                Case "PLATE STEEL"
                    FileToOpen = "Plate BOM"
                    FormatLine2(PlateRowNo, FileToOpen)
                    With PlateWrkSht
                        .Range("A" & PlateRowNo).Value = myarray2(1, i)             'Dwg & Ref No
                        .Range("B" & PlateRowNo).Value = myarray2(2, i)             'DWG
                        .Range("C" & PlateRowNo).Value = myarray2(3, i)             'Rev
                        .Range("D" & PlateRowNo).Value = myarray2(4, i)             'Ship Mark
                        .Range("E" & PlateRowNo).Value = myarray2(5, i)             'Piece Mark
                        .Range("F" & PlateRowNo).Value = myarray2(6, i)             'QTY
                        .Range("G" & PlateRowNo).Value = myarray2(7, i)             'Description
                        .Range("H" & PlateRowNo).Value = myarray2(8, i)             'Inv-1
                        .Range("I" & PlateRowNo).Value = myarray2(9, i)             'Inv-2 Standard Number MX1001A
                        .Range("J" & PlateRowNo).Value = myarray2(10, i)             'Material
                        .Range("K" & PlateRowNo).Value = myarray2(11, i)            'Weight
                        .Range("L" & PlateRowNo).Value = myarray2(12, i)
                        .Range("M" & PlateRowNo).Value = myarray2(13, i)            'Material Type - Pipe - Angle - ETC...
                        .Range("N" & PlateRowNo).Value = myarray2(14, i)
                    End With
                    PlateRowNo = CStr(CDbl(PlateRowNo) + 1)
                Case "SHEET STEEL-GA"
                    FileToOpen = "Plate BOM"
                    FormatLine2(PlateRowNo, FileToOpen)
                    With PlateWrkSht
                        .Range("A" & PlateRowNo).Value = myarray2(1, i)             'Dwg & Ref No
                        .Range("B" & PlateRowNo).Value = myarray2(2, i)             'DWG
                        .Range("C" & PlateRowNo).Value = myarray2(3, i)             'Rev
                        .Range("D" & PlateRowNo).Value = myarray2(4, i)             'Ship Mark
                        .Range("E" & PlateRowNo).Value = myarray2(5, i)             'Piece Mark
                        .Range("F" & PlateRowNo).Value = myarray2(6, i)             'QTY
                        .Range("G" & PlateRowNo).Value = myarray2(7, i)             'Description
                        .Range("H" & PlateRowNo).Value = myarray2(8, i)             'Inv-1
                        .Range("I" & PlateRowNo).Value = myarray2(9, i)             'Inv-2 Standard Number MX1001A
                        .Range("J" & PlateRowNo).Value = myarray2(10, i)             'Material
                        .Range("K" & PlateRowNo).Value = myarray2(11, i)            'Weight
                        .Range("L" & PlateRowNo).Value = myarray2(12, i)
                        .Range("M" & PlateRowNo).Value = myarray2(13, i)            'Material Type - Pipe - Angle - ETC...
                        .Range("N" & PlateRowNo).Value = myarray2(14, i)
                    End With
                    PlateRowNo = CStr(CDbl(PlateRowNo) + 1)
                Case "GRATING"
                    FileToOpen = "Grating BOM"
                    FormatLine2(GratingRowNo, FileToOpen)
                    With GratingWrkSht
                        .Range("A" & GratingRowNo).Value = myarray2(1, i)             'Dwg & Ref No
                        .Range("B" & GratingRowNo).Value = myarray2(2, i)             'DWG
                        .Range("C" & GratingRowNo).Value = myarray2(3, i)             'Rev
                        .Range("D" & GratingRowNo).Value = myarray2(4, i)             'Ship Mark
                        .Range("E" & GratingRowNo).Value = myarray2(5, i)             'Piece Mark
                        .Range("F" & GratingRowNo).Value = myarray2(6, i)             'QTY
                        .Range("G" & GratingRowNo).Value = myarray2(7, i)             'Description
                        .Range("H" & GratingRowNo).Value = myarray2(8, i)             'Inv-1
                        .Range("I" & GratingRowNo).Value = myarray2(9, i)             'Inv-2 Standard Number MX1001A
                        .Range("J" & GratingRowNo).Value = myarray2(10, i)             'Material
                        .Range("K" & GratingRowNo).Value = myarray2(11, i)            'Weight
                        .Range("L" & GratingRowNo).Value = myarray2(12, i)
                        .Range("M" & GratingRowNo).Value = myarray2(13, i)            'Material Type - Pipe - Angle - ETC...
                        .Range("N" & GratingRowNo).Value = myarray2(14, i)
                    End With
                    GratingRowNo = CStr(CDbl(GratingRowNo) + 1)
                Case "PURCHASE ITEM"
                    FileToOpen = "Other BOM"
                    FormatLine2(PurchaseRowNo, FileToOpen)
                    With PurchaseWrkSht
                        .Range("A" & PurchaseRowNo).Value = myarray2(1, i)             'Dwg & Ref No
                        .Range("B" & PurchaseRowNo).Value = myarray2(2, i)             'DWG
                        .Range("C" & PurchaseRowNo).Value = myarray2(3, i)             'Rev
                        .Range("D" & PurchaseRowNo).Value = myarray2(4, i)             'Ship Mark
                        .Range("E" & PurchaseRowNo).Value = myarray2(5, i)             'Piece Mark
                        .Range("F" & PurchaseRowNo).Value = myarray2(6, i)             'QTY
                        .Range("G" & PurchaseRowNo).Value = myarray2(7, i)             'Description
                        .Range("H" & PurchaseRowNo).Value = myarray2(8, i)             'Inv-1
                        .Range("I" & PurchaseRowNo).Value = myarray2(9, i)             'Inv-2 Standard Number MX1001A
                        .Range("J" & PurchaseRowNo).Value = myarray2(10, i)             'Material
                        .Range("K" & PurchaseRowNo).Value = myarray2(11, i)            'Weight
                        .Range("L" & PurchaseRowNo).Value = myarray2(12, i)
                        .Range("M" & PurchaseRowNo).Value = myarray2(13, i)            'Material Type - Pipe - Angle - ETC...
                        .Range("N" & PurchaseRowNo).Value = myarray2(14, i)
                    End With
                    PurchaseRowNo = CStr(CDbl(PurchaseRowNo) + 1)
                Case Else
                    ' Stop          'Removed stop
                    'FileToOpen = "Shop Cut BOM"
                    'FormatLine(RowNo, FileToOpen)
                    'With ShopCutWrkSht          'With Worksheets("Shop Cut BOM")
                    '    .Range("A" & RowNo).Value = myarray2(1, i)                      'DWG
                    '    .Range("B" & RowNo).Value = myarray2(2, i)                      'Rev
                    '    .Range("C" & RowNo).Value = myarray2(3, i)                      'Ship Mark
                    '    .Range("D" & RowNo).Value = myarray2(4, i)                      'Piece Mark
                    '    .Range("E" & RowNo).Value = myarray2(5, i)                      'QTY
                    '    .Range("F" & RowNo).Value = myarray2(6, i)                      'Description
                    '    .Range("G" & RowNo).Value = myarray2(7, i)                      'Inv-1
                    '    .Range("H" & RowNo).Value = myarray2(8, i)                      'Inv-2 Standard Number MX1001A
                    '    .Range("I" & RowNo).Value = myarray2(9, i)                      'Material
                    '    .Range("J" & RowNo).Value = myarray2(10, i)                     'Weight
                    '    .Range("K" & RowNo).Value = ""                                  'Blank
                    '    .Range("L" & RowNo).Value = myarray2(11, i)                     'Material Type - Pipe - Angle - ETC...
                    'End With
                    'RowNo = CStr(CDbl(RowNo) + 1)
            End Select
        Next i

        RowNo = CStr(CDbl(RowNo) - 1)
        PlateRowNo = CStr(CDbl(PlateRowNo) - 1)
        StickRowNo = CStr(CDbl(StickRowNo) - 1)
        PurchaseRowNo = CStr(CDbl(PurchaseRowNo) - 1)
        GratingRowNo = CStr(CDbl(GratingRowNo) - 1)

        With PlateWrkSht
            With .Range("A4:N" & PlateRowNo)
                .Sort(Key1:=.Columns("L"), Order1:=XlSortOrder.xlAscending, Header:=XlYesNoGuess.xlYes, Orientation:=XlSortOrientation.xlSortColumns)
            End With
            .Range("J5:J" & PlateRowNo).Font.Size = 7
        End With

        With StickWrkSht
            With .Range("A4:N" & StickRowNo)
                .Sort(Key1:=.Columns("G"), Order1:=XlSortOrder.xlAscending, Header:=XlYesNoGuess.xlYes, Orientation:=XlSortOrientation.xlSortColumns)
            End With
            .Range("J5:J" & StickRowNo).Font.Size = 7
        End With

        With GratingWrkSht
            With .Range("A4:N" & GratingRowNo)
                .Sort(Key1:=.Columns("A"), Order1:=XlSortOrder.xlAscending, Header:=XlYesNoGuess.xlYes, Orientation:=XlSortOrientation.xlSortColumns)
            End With
            .Range("J5:J" & StickRowNo).Font.Size = 7
        End With

        With PurchaseWrkSht
            With .Range("A4:N" & PurchaseRowNo)
                .Sort(Key1:=.Columns("G"), Order1:=XlSortOrder.xlAscending, Header:=XlYesNoGuess.xlYes, Orientation:=XlSortOrientation.xlSortColumns)
            End With
            .Range("J5:J" & PurchaseRowNo).Font.Size = 7
            BOMWrkSht.Activate()
        End With

Err_ShopCUT:
        ErrNo = Err.Number

        If ErrNo <> 0 Then
            PrgName = "ShopCut"
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
            If GenInfo.UserName = "dlong" Then
                If ErrNo = "5" And InStr(ErrMsg, "must be greater or equal to zero.") > 0 Then
                    ProbPart = (myarray(7, i))      'Program Errored out on ShopCut put part in Purchase Item List.
                    'GoTo PurchaseItem           'Program will not allow this to happen if For LOOP
                    PurchaseProb = True
                    'Stop
                    Resume Next
                End If
                MsgBox(ErrMsg)
                Stop
                Resume
            Else
                If ErrNo = "5" And InStr(ErrMsg, "must be greater or equal to zero.") > 0 Then
                    ProbPart = (myarray(7, i))      'Program Errored out on ShopCut put part in Purchase Item List.
                    'GoTo PurchaseItem           'Program will not allow this to happen if For LOOP
                    PurchaseProb = True
                    Resume Next
                End If
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

    Function FeetInchesToDecInches(ByVal MaterialPart As String) As Object
        Dim myarray As Object
        Dim myarray2 As Object
        Dim FoundSch, FoundStd, FoundX, FoundInch, FoundFoot, FoundFraction, LenMatPart, PartFoundX As Integer
        Dim FoundDash, FoundSpace, FoundFoot2, Test As Integer
        Dim SearchPipe, SearchX, SearchInch, SearchSch, SearchStd, SearchShell, SearchDwg, SearchFoot As String
        Dim SearchFraction, SearchDash As String
        Dim MatLength As String
        Dim DecInches, SearchSpace As String
        Dim FeetLength, FeetTotal, InchFraction, FstFraction, SndFraction, DecFraction As Double
        'Dim MatInch As Double
        Dim i As Short, i2 As Short
        'Dim ErrException As System.Exception            'System.EventHandler 

        PrgName = "FeetInchesToDecInches"

        On Error GoTo Err_FeetInchesToDecInches

        SearchX = " x "
        SearchInch = Chr(34)                'Find " Inch
        SearchFoot = Chr(39)                'Find ' Feet
        SearchSch = "SCH"
        SearchStd = "STD"
        SearchFraction = "/"
        SearchDash = "-"
        SearchSpace = " "

        InchFraction = 0
        DecFraction = 0
        FeetTotal = 0
        FoundFraction = 0
        FstFraction = 0
        SndFraction = 0
        FoundInch = 0
        FeetLength = 0
        FoundFoot = 0
        FoundFoot2 = 0

        '--------------------------------------------------------Convert to Decimal Inches
        DecInches = MaterialPart
        FoundInch = InStr(1, DecInches, SearchInch)

        If FoundInch > 0 Then                   'First remove extra characters at end Example: 3'-5 1/4" MOE
            DecInches = Mid(DecInches, 1, FoundInch)
        Else
            DecInches = MaterialPart
        End If

        FoundFoot = InStr(1, DecInches, SearchFoot)
        FoundDash = InStr(1, DecInches, SearchDash)

        If FoundFoot > 0 Then
            If FoundDash > 0 Then
                FeetLength = Mid(DecInches, 1, (FoundFoot - 1))
                FeetTotal = (FeetLength * 12)
                DecInches = Mid(DecInches, (FoundDash + 1), (Len(DecInches) - FoundDash + 1))
            Else
                FoundFoot2 = InStr((FoundFoot + 1), DecInches, SearchFoot)                 'Make sure user did not enter two Foot Marks together ''  Instead of an Inch mark "

                If FoundFoot2 > 0 Then      'Make sure user did not enter two Foot Marks together ''  Instead of an Inch mark "
                    If FoundFoot2 = (FoundFoot + 1) Then
                        FeetTotal = 0
                        FoundSpace = InStr((FoundFoot - 1), DecInches, SearchSpace)
                        If FoundSpace > 0 Then
                            DecInches = Mid(DecInches, 1, (FoundSpace - 1))
                        Else
                            DecInches = Mid(DecInches, 1, (FoundSpace - 3))
                        End If
                    End If
                Else
                    FoundSpace = InStr(1, DecInches, SearchSpace)
                    DecInches = Mid(DecInches, (FoundFoot + 1), (Len(DecInches) - FoundFoot + 1))
                End If

                End If
        End If

        FoundFraction = InStr(1, DecInches, SearchFraction)
        FoundSpace = InStr(1, DecInches, SearchSpace)
        FoundInch = InStr(1, DecInches, SearchInch)
        FoundFoot = InStr(1, DecInches, SearchFoot)
        FoundFoot2 = InStr((FoundFoot + 1), DecInches, SearchFoot)

        If FoundFraction > 0 Then
            If FoundSpace > 0 Then
                If FoundInch > 0 Then
                    InchFraction = Mid(DecInches, 1, (FoundSpace - 1))          'Fraction Found
                    DecInches = Mid(DecInches, (FoundSpace + 1), (Len(DecInches) - (FoundSpace + 1)))   'Remove Inch mark as well "
                    FoundFraction = InStr(1, DecInches, SearchFraction)
                    FstFraction = Mid(DecInches, 1, (FoundFraction - 1))
                    SndFraction = Mid(DecInches, (FoundFraction + 1), (Len(DecInches)))
                    DecFraction = (FstFraction / SndFraction)
                Else
                    If FoundFoot2 = (FoundFoot + 1) Then
                        InchFraction = Mid(DecInches, 1, (FoundSpace - 1))          'Fraction Found
                        Test = (Len(DecInches))
                        Test = (Len(DecInches) - FoundFoot)
                        DecInches = Mid(DecInches, (FoundSpace + 1), (Len(DecInches) - (FoundSpace + 1 + Test)))   'Remove Double Foot marks as well ''
                        FoundFraction = InStr(1, DecInches, SearchFraction)
                        FstFraction = Mid(DecInches, 1, (FoundFraction - 1))
                        SndFraction = Mid(DecInches, (FoundFraction + 1), (Len(DecInches)))
                        DecFraction = (FstFraction / SndFraction)
                    Else
                        'MsgBox("A programming problem has been found that need to be looked at, contact Software Adminstration")
                        'Stop
                        InchFraction = Mid(DecInches, 1, (FoundSpace - 1))          'Fraction Found
                        Test = (Len(DecInches))
                        Test = (Len(DecInches) - FoundFoot)
                        'DecInches = "8 5/16"
                        DecInches = Mid(DecInches, (FoundSpace + 1), (Len(DecInches) - (FoundSpace)))   'Remove Double Foot marks as well ''
                        FoundFraction = InStr(1, DecInches, SearchFraction)
                        FstFraction = Mid(DecInches, 1, (FoundFraction - 1))
                        Dim LastPart As String
                        Dim SearchBBE, SearchBBL, SearchBOE, SearchBW, SearchCBE, SearchCOE As String
                        Dim SearchG, SearchGAL, SearchMBE, SearchMOE, SearchNPT, SearchPBE As String
                        Dim SearchSOL, SearchTBE, SearchTOE, SearchTOL, SearchTS, SearchWOL As String
                        Dim LenBBE As Integer
                        Dim FoundBBE, FoundBBL, FoundBOE, FoundBW, FoundCBE, FoundCOE, FoundG As Integer
                        Dim FoundGAL, FoundMBE, FoundMOE, FoundNPT, FoundPBE, FoundSOL, FoundTBE As Integer
                        Dim FoundTOE, FoundTOL, FoundTS, FoundWOL As Integer

                        SearchBBE = " BBE"
                        SearchBBL = " BBL"
                        SearchBOE = " BOE"
                        SearchBW = " BW"
                        SearchCBE = " CBE"
                        SearchCOE = " COE"
                        SearchG = " G"
                        SearchGAL = " GAL"
                        SearchMBE = " MBE"
                        SearchMOE = " MOE"
                        SearchNPT = " NPT"
                        SearchPBE = " PBE"
                        SearchSOL = " SOL"
                        SearchTBE = " TBE"
                        SearchTOE = " TOE"
                        SearchTOL = " TOL"
                        SearchTS = " TS"
                        SearchWOL = " WOL"
                        FoundBBE = InStr(1, DecInches, SearchBBE)
                        FoundBBL = InStr(1, DecInches, SearchBBL)
                        FoundBOE = InStr(1, DecInches, SearchBOE)
                        FoundBW = InStr(1, DecInches, SearchBW)
                        FoundCBE = InStr(1, DecInches, SearchCBE)
                        FoundCOE = InStr(1, DecInches, SearchCOE)
                        FoundG = InStr(1, DecInches, SearchG)
                        FoundGAL = InStr(1, DecInches, SearchGAL)
                        FoundMBE = InStr(1, DecInches, SearchMBE)
                        FoundMOE = InStr(1, DecInches, SearchMOE)
                        FoundNPT = InStr(1, DecInches, SearchNPT)
                        FoundPBE = InStr(1, DecInches, SearchPBE)
                        FoundSOL = InStr(1, DecInches, SearchSOL)
                        FoundTBE = InStr(1, DecInches, SearchTBE)
                        FoundTOE = InStr(1, DecInches, SearchTOE)
                        FoundTOL = InStr(1, DecInches, SearchTOL)
                        FoundTS = InStr(1, DecInches, SearchTS)
                        FoundWOL = InStr(1, DecInches, SearchWOL)

                        Select Case 0
                            Case Is < FoundBBE
                                LenBBE = Len(FoundBBE)
                                LastPart = Mid(DecInches, FoundBBE, (Len(DecInches) - (LenBBE - 1)))
                                SndFraction = Mid(DecInches, (FoundFraction + 1), (Len(DecInches) - (LenBBE + FoundFraction)))
                                DecFraction = (FstFraction / SndFraction)
                            Case Is < FoundBBL
                                LenBBE = Len(FoundBBL)
                                LastPart = Mid(DecInches, FoundBBL, (Len(DecInches) - (LenBBE - 1)))
                                SndFraction = Mid(DecInches, (FoundFraction + 1), (Len(DecInches) - (LenBBE + FoundFraction)))
                                DecFraction = (FstFraction / SndFraction)
                            Case Is < FoundBOE
                                LenBBE = Len(FoundBOE)
                                LastPart = Mid(DecInches, FoundBOE, (Len(DecInches) - (LenBBE - 1)))
                                SndFraction = Mid(DecInches, (FoundFraction + 1), (Len(DecInches) - (LenBBE + FoundFraction)))
                                DecFraction = (FstFraction / SndFraction)
                            Case Is < FoundBW
                                LenBBE = Len(FoundBW)
                                LastPart = Mid(DecInches, FoundBW, (Len(DecInches) - (LenBBE - 1)))
                                SndFraction = Mid(DecInches, (FoundFraction + 1), (Len(DecInches) - (LenBBE + FoundFraction)))
                                DecFraction = (FstFraction / SndFraction)
                            Case Is < FoundCBE
                                LenBBE = Len(FoundCBE)
                                LastPart = Mid(DecInches, FoundCBE, (Len(DecInches) - (LenBBE - 1)))
                                SndFraction = Mid(DecInches, (FoundFraction + 1), (Len(DecInches) - (LenBBE + FoundFraction)))
                                DecFraction = (FstFraction / SndFraction)
                            Case Is < FoundCOE
                                LenBBE = Len(FoundCOE)
                                LastPart = Mid(DecInches, FoundCOE, (Len(DecInches) - (LenBBE - 1)))
                                SndFraction = Mid(DecInches, (FoundFraction + 1), (Len(DecInches) - (LenBBE + FoundFraction)))
                                DecFraction = (FstFraction / SndFraction)
                            Case Is < FoundG
                                LenBBE = Len(FoundG)
                                LastPart = Mid(DecInches, FoundG, (Len(DecInches) - (LenBBE - 1)))
                                SndFraction = Mid(DecInches, (FoundFraction + 1), (Len(DecInches) - (LenBBE + FoundFraction)))
                                DecFraction = (FstFraction / SndFraction)
                            Case Is < FoundGAL
                                LenBBE = Len(FoundGAL)
                                LastPart = Mid(DecInches, FoundGAL, (Len(DecInches) - (LenBBE - 1)))
                                SndFraction = Mid(DecInches, (FoundFraction + 1), (Len(DecInches) - (LenBBE + FoundFraction)))
                                DecFraction = (FstFraction / SndFraction)
                            Case Is < FoundMBE
                                LenBBE = Len(FoundMBE)
                                LastPart = Mid(DecInches, FoundMBE, (Len(DecInches) - (LenBBE - 1)))
                                SndFraction = Mid(DecInches, (FoundFraction + 1), (Len(DecInches) - (LenBBE + FoundFraction)))
                                DecFraction = (FstFraction / SndFraction)
                            Case Is < FoundMOE
                                LenBBE = Len(FoundMOE)
                                LastPart = Mid(DecInches, FoundMOE, (Len(DecInches) - (LenBBE - 1)))
                                SndFraction = Mid(DecInches, (FoundFraction + 1), (Len(DecInches) - (LenBBE + FoundFraction)))
                                DecFraction = (FstFraction / SndFraction)
                            Case Is < FoundNPT
                                LenBBE = Len(FoundNPT)
                                LastPart = Mid(DecInches, FoundNPT, (Len(DecInches) - (LenBBE - 1)))
                                SndFraction = Mid(DecInches, (FoundFraction + 1), (Len(DecInches) - (LenBBE + FoundFraction)))
                                DecFraction = (FstFraction / SndFraction)
                            Case Is < FoundPBE
                                LenBBE = Len(FoundPBE)
                                LastPart = Mid(DecInches, FoundPBE, (Len(DecInches) - (LenBBE - 1)))
                                SndFraction = Mid(DecInches, (FoundFraction + 1), (Len(DecInches) - (LenBBE + FoundFraction)))
                                DecFraction = (FstFraction / SndFraction)
                            Case Is < FoundSOL
                                LenBBE = Len(FoundSOL)
                                LastPart = Mid(DecInches, FoundSOL, (Len(DecInches) - (LenBBE - 1)))
                                SndFraction = Mid(DecInches, (FoundFraction + 1), (Len(DecInches) - (LenBBE + FoundFraction)))
                                DecFraction = (FstFraction / SndFraction)
                            Case Is < FoundTBE
                                LenBBE = Len(FoundTBE)
                                LastPart = Mid(DecInches, FoundTBE, (Len(DecInches) - (LenBBE - 1)))
                                SndFraction = Mid(DecInches, (FoundFraction + 1), (Len(DecInches) - (LenBBE + FoundFraction)))
                                DecFraction = (FstFraction / SndFraction)
                            Case Is < FoundTOE
                                LenBBE = Len(FoundTOE)
                                LastPart = Mid(DecInches, FoundTOE, (Len(DecInches) - (LenBBE - 1)))
                                SndFraction = Mid(DecInches, (FoundFraction + 1), (Len(DecInches) - (LenBBE + FoundFraction)))
                                DecFraction = (FstFraction / SndFraction)
                            Case Is < FoundTOL
                                LenBBE = Len(FoundTOL)
                                LastPart = Mid(DecInches, FoundTOL, (Len(DecInches) - (LenBBE - 1)))
                                SndFraction = Mid(DecInches, (FoundFraction + 1), (Len(DecInches) - (LenBBE + FoundFraction)))
                                DecFraction = (FstFraction / SndFraction)
                            Case Is < FoundTS
                                LenBBE = Len(FoundTS)
                                LastPart = Mid(DecInches, FoundTS, (Len(DecInches) - (LenBBE - 1)))
                                SndFraction = Mid(DecInches, (FoundFraction + 1), (Len(DecInches) - (LenBBE + FoundFraction)))
                                DecFraction = (FstFraction / SndFraction)
                            Case Is < FoundWOL
                                LenBBE = Len(FoundWOL)
                                LastPart = Mid(DecInches, FoundWOL, (Len(DecInches) - (LenBBE - 1)))
                                SndFraction = Mid(DecInches, (FoundFraction + 1), (Len(DecInches) - (LenBBE + FoundFraction)))
                                DecFraction = (FstFraction / SndFraction)
                            Case Else
                                SndFraction = Mid(DecInches, (FoundFraction + 1), (Len(DecInches)))
                                DecFraction = (FstFraction / SndFraction)
                        End Select

                        'Moved to above
                        'SndFraction = Mid(DecInches, (FoundFraction + 1), (Len(DecInches)))
                        'DecFraction = (FstFraction / SndFraction)
                    End If
                    End If
            Else
                    FstFraction = Mid(DecInches, 1, (FoundFraction - 1))        'Fraction Not Found
                    SndFraction = Mid(DecInches, 1, (FoundFraction + 1))
                    DecFraction = (FstFraction / SndFraction)
            End If
            Else
                FoundInch = InStr(1, DecInches, SearchInch)
                If FoundInch > 0 Then
                    InchFraction = Mid(DecInches, 1, (FoundInch - 1))           'Remove Inch mark as well "
                Else
                    FoundFoot = InStr(1, DecInches, SearchFoot)
                    FoundFoot2 = InStr((FoundFoot + 1), DecInches, SearchFoot)

                    If FoundFoot2 = (FoundFoot + 1) Then    'Make sure user did not enter two Foot Marks together ''  Instead of an Inch mark "
                        InchFraction = Mid(DecInches, 1, (FoundFoot - 1))
                    Else
                    '

                    Select Case Left(MaterialPart, 1)
                        Case 1
                            InchFraction = MaterialPart
                        Case 2
                            InchFraction = MaterialPart
                        Case 3
                            InchFraction = MaterialPart
                        Case 4
                            InchFraction = MaterialPart
                        Case 5
                            InchFraction = MaterialPart
                        Case 6
                            InchFraction = MaterialPart
                        Case 7
                            InchFraction = MaterialPart
                        Case 8
                            InchFraction = MaterialPart
                        Case 9
                            InchFraction = MaterialPart
                        Case 0
                            InchFraction = MaterialPart
                        Case Else
                            'MsgBox("A programming problem has been found that need to be looked at, contact Software Adminstration")
                            InchFraction = 0
                    End Select

                    End If

                End If
                FstFraction = 0        'Fraction Not Found
                SndFraction = 0
                DecFraction = 0
            End If

                FoundInch = InStr(1, MaterialPart, SearchInch)

                If FoundInch > 0 Then
                    MatInch = (InchFraction + DecFraction + FeetTotal)
                    MatLength = Mid(MaterialPart, 1, FoundInch)
                    'myarray(13, i) = MatInch                      'myarray(13, i) = MatLength
                Else
                    MatInch = (InchFraction + DecFraction + FeetTotal)
                    MatLength = MaterialPart
                    'myarray(13, i) = MatInch                      'myarray(13, i) = MatLength
                End If

Err_FeetInchesToDecInches:
                ErrNo = Err.Number

        If ErrNo <> 0 Then
            PrgName = "FeetInchestoDecInches"
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
            If GenInfo.UserName = "dlong" Then
                MsgBox(ErrMsg)
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

End Module