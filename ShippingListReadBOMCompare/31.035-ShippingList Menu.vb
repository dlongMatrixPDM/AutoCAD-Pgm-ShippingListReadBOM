Option Strict Off
Option Explicit On
Option Compare Text

Imports System
Imports System.Reflection
Imports System.Windows.Forms
Imports System.Data
Imports System.Data.OleDb
Imports System.IO
Imports System.IO.Directory
Imports System.IO.DirectoryInfo
'Imports System.IO.FileInfo
Imports System.IO.Stream
Imports ADODB
Imports System.String
Imports System.Runtime.InteropServices
Imports System.Runtime.InteropServices.ComTypes
Imports System.Reflection.Module
Imports System.Text.RegularExpressions
Imports Microsoft.SqlServer
Imports Microsoft.SqlServer.Server
Imports Microsoft.SqlServer.Server.SqlDataRecord
Imports System.Data.SqlClient
Imports Microsoft.Office.Interop.Excel

Imports MSforms

Public Structure GenInfo
    Dim FlangeDesc As String
    Public Shared UserName As String
    Public Shared SRList(17, 1)
End Structure

Public Structure ManwayInfo3135
    Public Shared MWElev As String
    Public Shared MWSize As String
    Public Shared FlangeOD As String
    Public Shared FlangeID As String
    Public Shared FlangeThk As String
    Public Shared FlangeHoles As String
    Public Shared FlangeChord As String
    Public Shared FlangePartNo As String
    Public Shared FlangeStockNo As String
    Public Shared FlangeDesc As String
    Public Shared UserName As String
    Public Shared AssyMat(0) As Object
    Dim Test As String
    Public Shared OldDesc As String
    Public Shared GetDesc As String
    Public Shared GetMatLen As String
    Public Shared GetMatQty As String
    Public Shared GetShipMk As String
    Public Shared GetPieceMk As String
    Public Shared MatType As String
    Public Shared RevNo As String

    Public Shared AddMat As String
    Public Shared AddMatSize As String
    Public Shared AddMatQty As String

    Public Shared FabInvListData(3, 1) As Object
    Public Shared StructureData(3, 1) As Object
    Public Shared FittingData(3, 1) As Object
    Public Shared GasketData(3, 1) As Object
    Public Shared FlangeData(3, 1) As Object
    Public Shared PipeData(3, 1) As Object
    Public Shared HardwareData(3, 1) As Object
    Public Shared MiscData(3, 1) As Object
    Public Shared SubMFGData(3, 1) As Object
    Public Shared SubPURData(3, 1) As Object
    Public Shared NamePlateData(3, 1) As Object
    Public Shared PlateData(3, 1) As Object
    Public Shared ShipSuplData(3, 1) As Object
    Public Shared TkSealData(3, 1) As Object
    Public Shared AllMatSortDesc(3, 1) As Object
End Structure

Public Structure GenInfo3135
    Dim TestNum As String
    Public Shared MaxDLLevel As String
    Public Shared DLLevel As String

    Public Shared UserName As String
    Public Shared AssyMat(0) As Object
    Public Shared PartType As String
    Public Shared CCFolder As String
    Public Shared CCFamily As String

    Public Shared Number2 As String
    Public Shared FileName As String
    Public Shared FullJobNo As String
    Public Shared CustomerPO As String
    Public Shared ShippingList(18, 1)

    Public Shared FileDir As String
    Public Shared JobDir As String
    Public Shared CopyFromFile As String
    Public Shared StrLineNo As Integer

    Public Shared NewArrayItem As String            'Used in FixFormat Function
    Public Shared OldArrayItem As String

    Public Shared DefaultPrt As String
    Public Shared DefaultPrtSet As Boolean
    Public Shared BlockRef 'As AcadBlockReference

    Public Shared NewLayout 'As AcadLayout
    Public Shared NewPlotConfig 'As AcadPlotConfiguration

    Public Shared StartAdept As Boolean

    Public Shared LLCornerVPort
    Public Shared URCornerVPort
    Public Shared CenterVPort
    Public Shared HeightVPort
    Public Shared WidthVPort

    Public Shared ExDwgs1 As Object
    Public Shared PrgReqSpreadSht As String            'Program Requesting SpreadSheet Info.
    Public Shared SpreadSht As String                   'Spreadsheet user Selected.
    Public Shared SpreadshtLoc As String                'Spreadsheet location.
    Public Shared HVECIssue As String
End Structure

Public Class ShippingList_Menu

    Inherits System.Windows.Forms.Form
    Dim Blockarry() As String
    Dim fs As Scripting.FileSystemObject
    Dim Fle As Scripting.File
    Dim db As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim ThisDrawing As AutoCAD.AcadDocument
    Public Shared AcadApp As AutoCAD.AcadApplication
    Public Shared ExcelApp As Object
    Dim XdataType, XdataValue As Object
    Dim Sset As AutoCAD.AcadSelectionSet
    Dim Collction As New Collection
    Dim SaveAction As Boolean
    Dim sysVarLogin, UserNam, DwgStatus As Object
    Dim NewLenDwgNo2, LenDwgNo2, DrawingNo2 As String
    Dim NewLenDwgNo3, DwgNo2Eng, DwgNo2, DwgNo2Job, DwgNo2Eng2, DwgNo2End, DrawingNo3 As String
    Dim ErrMsg, ErrNo, ErrSource, ErrDll, PriPrg, PrgName As String
    Dim ErrException As System.Exception
    Dim ErrLastLineX As Integer
    Dim ListItm3 As System.Windows.Forms.ListViewItem
    Dim SSget_Tblock, ErrFnd As Object
    Dim FileToOpen As String
    Public CBclicked As Boolean
    Public SBclicked As Boolean
    Public SortListing As Boolean
    Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, ByRef lParam As VariantType) As Integer

    Public FirstTimeThru As String
    Public FuncGetDataNew As String
    Public NewBulkBOM As Worksheet
    Public ReadBulkBOM As Worksheet         '-------DJL-07-15-2025
    Public NewShippingList As Worksheet
    Public OldShippingList As Worksheet
    'Public OldShipListSht As Worksheet
    Public WorkSht As Worksheet, BOMWrkSht As Worksheet, ShpCutWrkSht As Worksheet, StdsWrkSht As Worksheet
    Public StdItemsWrkSht, ShipListSht, OldShipListSht As Worksheet

    Public MainBOMFile As Excel.Workbook
    Public OldShipListFile As Excel.Workbook
    Public NewGratingBOM As Object
    Public NewPlateBOM As Object
    Public NewStickBOM As Object
    Public NewStickImport As Object
    Public NewStickImportBOM As Object
    Public NewPurchaseBOM As Object
    Public OldBOMFile As Excel.Workbook
    Public ReadBOMFile As Excel.Workbook
    Public OldBulkBOM As Object
    Public OldBulkBOMFile As String
    Public OldGratingBOM As Object
    Public OldPlateBOM As Object
    Public OldStickBOM As Object
    Public OldStickImportBOM As Object
    Public OldPurchaseBOM As Object
    Public NewBOM As Object
    Public OldBOM As Object
    Public FindSTD As Object
    Public OldStdItems As Object
    'Public BOMType As String                           '-------DJL-------10-31-2023---Not needed anymore.
    Public BOMSheet As String
    Public RowNo As Integer
    Public RowNo2 As String
    Public OldStdDwg As String
    Public NewStdDwg As String
    Public ExceptionPos As Integer
    Public CallPos As Integer
    Public CntExcept As Integer
    Public Count As Integer
    Public PassFilename As String
    Public ReadyToContinue As Boolean
    Public errorExist As Boolean
    Public BOMList() As String
    Public StdBOMList() As String
    Public AcadDoc As Object
    Public RevNo, RevNo2 As String
    Public Continue_Renamed As Boolean
    Public MatInch As Double
    Public FoundDir, SearchException As String
    Public ExceptPos As Integer
    Public LytHid As Boolean
    Public CountBOM, CountNewItems, StrLineNo As Integer
    Public CustomerPO, Cust1, Cust2, OldShipListFileStr As String
    Public OldDwgItem As String
    Public DwgName, SearchSlash As String
    Public SearchPos As Integer, DwgSize As Integer
    Public SecondPart
    Public DwgDetails
    Public Dwg2
    Public TitleBlkName As String
    Public db_String As String
    Public CurrentTask, LastNam, FirstNam, FChrTest, CodeName, DeptName As String
    Public JobNum, SubNum, TaskName, Emp, TableName, SourceName, PrevLastNam, MInit, Title, Funct As String
    Public DeptNam, HireDate, PTOPrevYr, Birth, LastNamChg, DwgItem1, ErrFound As String
    Public dbSQL_String, WorkStationID1, OldWorkStation, FuncName, FindUser, StartDir, FileNam, NewDir, SecondChk, StartAdept As String
    Public strSQL2, ByName, ByNam, TableNam, FindDwg, FindRev, FindRev1, FindRev2, FileSaveAS, FullJobNo, JobNoDash As String
    Public ExtPos, ExtPos2, MxPos1, MxPos2, MxPos3, MxPos4, ChPos1, ChPos2, ChPos3, ChPos4, CaPos1, CaPos2, CaPos3, CaPos4, CntItems As Integer
    Public Shared CntFrames As Integer
    Public GetFramesSrt
    Public PrgLineNo As String
    Public FindEndStr, LenPrgLineNo, CntSpreadSht, BOMPos, FoundRev, RevPos As Integer
    Public ShippingList(18, 1)
    Public ShippingListColl(18, 1)
    Public SRList(18, 1)

    Public DsTLengStaff As New DataSet
    Declare Function apiGetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, ByRef nSize As Integer) As Integer
    Public Sapi As Object = CreateObject("SAPI.spvoice")

    Private Sub ComboBox1_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles ComboBox1.TextChanged
        If ComboBox1.Items.Count > 0 Then
            CheckBox1.Checked = True
            File2.Enabled = True
        End If

        If Me.SelectList.Items.Count > 0 Then
            Me.ComboBox1.BackColor = System.Drawing.Color.White
            Me.BtnStart2.BackColor = System.Drawing.Color.GreenYellow
            'Me.BtnSpeedTest.BackColor = System.Drawing.Color.GreenYellow                            '-------DJL-06-02-2025
            '06/05/2025 Adam had a reference using DBX per the Document it was faster found it is not 44:32 mins versa current system 19 mins from Tulsa '-------DJL-06-05-2025
        End If
    End Sub

    Private Sub CheckBox1_CheckChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CheckBox1.CheckedChanged
        Select Case CheckBox1.CheckState
            Case CheckState.Unchecked
                File2.Enabled = False
            Case CheckState.Checked
                File2.Enabled = True
        End Select
    End Sub

    Private Sub AddButton_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles BtnAdd.Click
        Dim Count, i, CountDwgs, CntDwgs As Integer
        Dim DwgListArray As Object

        On Error GoTo Err_AddButton_Click

        PrgName = "AddButton_Click"
        CntDwgs = 0
        CountDwgs = CollectBOMList.SelectedItems.Count
        ReDim DwgListArray(CountDwgs)

        If CountDwgs > 1 Then
            MsgBox("Only one BOM Spreadsheet is allowed.")
            GoTo MoreThanOneBOMSelected
        End If

        For i = 0 To (CountDwgs - 1)
            SelectList.Items.Add(CollectBOMList.SelectedItems.Item(i))
            DwgListArray(CntDwgs) = CollectBOMList.SelectedItems.Item(i)
            CntDwgs = (CntDwgs + 1)
        Next i

        For i = 0 To (CountDwgs - 1)
            CollectBOMList.Items.Remove(DwgListArray(i))
        Next

        CollectBOMList.Sorted = True
        SelectList.Sorted = True

MoreThanOneBOMSelected:
        If SelectList.Items.Count = 0 Then
            BtnRemove.Enabled = False
            BtnClear.Enabled = False
        Else
            BtnRemove.Enabled = True
            BtnClear.Enabled = True
        End If

        If Me.SelectList.Items.Count <> 0 Then
            Me.BtnStart2.Enabled = True
            SelectList.BackColor = System.Drawing.Color.LawnGreen       '-------DJL-07-21-2025
            CollectBOMList.BackColor = System.Drawing.Color.White
            'Me.BtnSpeedTest.Enabled = True                            '-------DJL-06-02-2025
            '06/05/2025 Adam had a reference using DBX per the Document it was faster found it is not 44:32 mins versa current system 19 mins from Tulsa '-------DJL-06-05-2025
        End If

Err_AddButton_Click:
        ErrNo = Err.Number

        If ErrNo <> 0 Then
            PrgName = "BtnAddClick"
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

            Me.HandleErrSQL(PrgName, ErrNo, ErrMsg, ErrSource, PriPrg, ErrDll, DwgItem, PrgLineNo)

            If System.Environment.UserName = "dlong" Then      '-------DJL-06-30-2025      'If (Me.UserNamex()) = "dlong" Then
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

    End Sub

    Private Sub CancelButton_Renamed_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CancelButton_Renamed.Click
        Me.Close()
    End Sub

    Private Sub ClearButton_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles BtnClear.Click
        Dim Count, i, CountDwgs, CntDwgs As Integer
        Dim DwgListArray As Object

        'On Error GoTo Err_ClearButton_Click

        PrgName = "ClearButton_Click"
        CntDwgs = 0
        CountDwgs = SelectList.Items.Count
        ReDim DwgListArray(CountDwgs)

        For i = 0 To (CountDwgs - 1)
            CollectBOMList.Items.Add(SelectList.Items.Item(i))          '-------DJL-07-15-2025
            DwgListArray(CntDwgs) = SelectList.Items.Item(i)          '-------DJL-07-15-2025
            CntDwgs = (CntDwgs + 1)
        Next i

        For i = 0 To (CountDwgs - 1)
            SelectList.Items.Remove(DwgListArray(i))
        Next

        CollectBOMList.Sorted = True
        SelectList.Sorted = True

MoreThanOneBOMSelected:
        If SelectList.Items.Count = 0 Then
            BtnRemove.Enabled = False
            BtnClear.Enabled = False
        Else
            BtnRemove.Enabled = True
            BtnClear.Enabled = True
        End If

        If Me.SelectList.Items.Count = 1 Then
            Me.BtnStart2.Enabled = True
            'Me.BtnSpeedTest.Enabled = True                            '-------DJL-06-02-2025
            '06/05/2025 Adam had a reference using DBX per the Document it was faster found it is not 44:32 mins versa current system 19 mins from Tulsa '-------DJL-06-05-2025
        End If

        'SelectList.Items.Clear()
        'BtnRemove.Enabled = False
        'BtnClear.Enabled = False
        'BtnStart2.Enabled = False
        'BtnSpeedTest.Enabled = False                            '-------DJL-06-02-2025
        '06/05/2025 Adam had a reference using DBX per the Document it was faster found it is not 44:32 mins versa current system 19 mins from Tulsa '-------DJL-06-05-2025
    End Sub

    Private Sub RemoveButton_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles BtnRemove.Click
        Dim i, DwgPos, CountDwgs, CntDwgs As Integer
        Dim GetDwg As String
        Dim DwgListArray As Object

        On Error GoTo Err_RemoveButton_Click

        PrgName = "RemoveButton_Click"
        CountDwgs = SelectList.SelectedItems.Count
        CntDwgs = 0
        ReDim DwgListArray(CountDwgs)

        For i = 0 To (CountDwgs - 1)
            DwgPos = 0
            GetDwg = SelectList.SelectedItems.Item(i).ToString

            Me.CollectBOMList.Items.Add(GetDwg)                         '-------DJL-07-15-2025
            DwgListArray(CntDwgs) = SelectList.SelectedItems.Item(i)
            CntDwgs = (CntDwgs + 1)
        Next i

        For i = 0 To (CountDwgs - 1)
            SelectList.Items.Remove(DwgListArray(i))
        Next

        If SelectList.Items.Count = 0 Then
            BtnRemove.Enabled = False
            BtnClear.Enabled = False
            BtnStart2.Enabled = False
            'BtnSpeedTest.Enabled = False                            '-------DJL-06-02-2025
            '06/05/2025 Adam had a reference using DBX per the Document it was faster found it is not 44:32 mins versa current system 19 mins from Tulsa '-------DJL-06-05-2025
        End If

        Refresh()
        CollectBOMList.Sorted = True                         '-------DJL-07-15-2025
        SelectList.Sorted = True

Err_RemoveButton_Click:
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

            HandleErrSQL(PrgName, ErrNo, ErrMsg, ErrSource, PriPrg, ErrDll, DwgItem, PrgLineNo)

            If GenInfo.UserName = "dlong" Then
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

    End Sub


    Private Sub SelectList_DblClick(ByVal Cancel As Boolean)
        PrgName = "SelectList_Start"                         '-------DJL-12-19-2024

        If Me.SelectList.SelectedIndex <> -1 Then
            If Me.SelectList.GetSelected(Me.SelectList.SelectedIndex) = True Then
                SelectList.Items.RemoveAt((Me.SelectList.SelectedIndex))
            End If
        End If

        If SelectList.Items.Count = 0 Then
            BtnRemove.Enabled = False
            BtnClear.Enabled = False
            BtnStart2.Enabled = False
            'BtnSpeedTest.Enabled = False                            '-------DJL-06-02-2025
            '06/05/2025 Adam had a reference using DBX per the Document it was faster found it is not 44:32 mins versa current system 19 mins from Tulsa '-------DJL-06-05-2025
        End If

        PrgName = "SelectList_DblClick"                         '-------DJL-12-19-2024

        On Error GoTo Err_StartButton_Click
Err_StartButton_Click:
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

            HandleErrSQL(PrgName, ErrNo, ErrMsg, ErrSource, PriPrg, ErrDll, DwgItem, PrgLineNo)

            If GenInfo.UserName = "dlong" Then
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

    End Sub

    Private Sub UserForm_Initialize()
        Dim RevList(20) As Short
        Dim i As Short

        PrgName = "UserForm_Initialize"

        On Error GoTo Err_UserForm_Initialize

        SBclicked = False
        File1.Text = "Click to Browse -->"
        File2.Text = "Click to Browse -->"
        File1.Focus()
        Me.Label2.Text = "Progress........"
        File2.Enabled = False

        For i = 0 To 20
            Me.ComboBox1.Items.Add(i)
        Next i

Err_UserForm_Initialize:
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

            HandleErrSQL(PrgName, ErrNo, ErrMsg, ErrSource, PriPrg, ErrDll, DwgItem, PrgLineNo)

            If GenInfo.UserName = "dlong" Then
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

    End Sub

    Private Sub PathBox2_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
Begin:
        On Error GoTo DriveError
        ChDrive(File2.Text)
        ChDir(File2.Text)
        'Call Dir1_Change(eventSender, eventArgs)
        File1.DisplayMember = "Existing Bulk BOM Located at " & File2.Text
        Exit Sub
DriveError:
        If Err.Description = "Device unavailable" Then
            If MsgBox("Drive " & UCase(Mid(File2.Text, 1, 1)) & ": is not ready or does not exist.", 21, "Bulk BOM Generator - Drive Error") = MsgBoxResult.Cancel Then
                Exit Sub
            Else
                Resume Begin
            End If
        End If
    End Sub

    Private Sub PathBox2_DropButtonClick()
        Dim NewDir As String

        currentDir = File1.Text
        NewDir = File.GetFile("Read Existing Bulk BOM From:")

        If NewDir <> "" Then
            File1.Text = Mid(NewDir, 1, Len(NewDir))
        End If
    End Sub

    'Private Sub Drive1_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Drive1.SelectedIndexChanged
    '    Dir1.Path = Drive1.Drive
    'End Sub

    '    Private Sub Dir1_Change(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Dir1.Change
    '        Dim i As Short
    '        Dim FindRev, FindRev1, FindDwg As String
    '        Dim FindRev2, FoundRev As Integer
    '        Dim RevPos, ExtPos, BOMPos, CntSpreadSht As Integer         'DJL-------06-02-2025

    '        PrgName = "UserForm_Initialize"
    '        FindRev = Nothing

    '        On Error GoTo Err_Dir1_Change

    '        Me.File1.Pattern = "*.dwg"
    '        Me.File2.Pattern = "*.xls"
    '        'btnNewDir.Enabled = True

    '        File1.Path = Dir1.Path
    '        Drive2.Drive = Drive1.Drive
    '        File2.Path = Dir1.Path
    '        GroupBox5.Text = "Current Directory = " & Dir1.Path

    '        If File1.Items.Count > 0 Then
    '            Me.BtnAdd.Enabled = True
    '            'btnNewDir.Enabled = True
    '        Else
    '            Me.File1.Items.Add("")
    '            Me.BtnAdd.Enabled = False
    '            'btnNewDir.Enabled = False
    '        End If

    '        If Me.File2.Items.Count > 0 Then
    'RestartFile2:
    '            FoundRev = 0
    '            CntSpreadSht = (Me.File2.Items.Count - 1)

    '            For i = 0 To CntSpreadSht
    '                BOMPos = 0

    '                If i > CntSpreadSht Then
    '                    GoTo Nexti
    '                End If

    '                FindRev = File2.Items.Item(i)
    '                RevPos = InStr(1, FindRev, "-R")
    '                BOMPos = InStr(1, FindRev, "BULKBOM")

    '                If BOMPos > 0 Then
    '                    GoTo UpdateBOMRev
    '                End If

    '                FindRev = Mid(FindRev, (RevPos + 2), Len(FindRev))
    '                FindRev1 = FindRev.Replace(".xls", "")

    '                Dim pattern As String = "^[0-9]"                    'Also Look at \Nd  Number with dec.
    '                Dim matches As MatchCollection = Regex.Matches(FindRev1, pattern)
    '                If Regex.IsMatch(FindRev1, pattern) Then
    '                    For Each match As Match In matches
    '                        FindRev2 = match.Value
    '                    Next
    '                End If

    '                If FoundRev < FindRev2 Then
    '                    FoundRev = FindRev2
    '                End If

    '                If BOMPos > 0 Then
    'UpdateBOMRev:
    '                    File2.Items.RemoveAt(i)
    '                    CntSpreadSht = (CntSpreadSht - 1)
    '                    GoTo RestartFile2
    '                End If
    'Nexti:
    '            Next i

    '            Dim index As Integer
    '            index = ComboBox1.FindString(FoundRev + 1)
    '            ComboBox1.SelectedIndex = index
    '        End If

    '        If Me.File1.Items.Count > 0 Then
    '            If DwgList.Items.Count > 0 Then
    '                DwgList.Items.Clear()
    '            End If

    '            For i = 0 To (Me.File1.Items.Count - 1)
    '                ExtPos = 0
    '                MxPos1 = 0
    '                MxPos2 = 0
    '                MxPos3 = 0
    '                MxPos4 = 0
    '                ChPos1 = 0         'DJL-------06-02-2025
    '                ChPos2 = 0
    '                ChPos3 = 0
    '                ChPos4 = 0
    '                CaPos1 = 0         'DJL-------06-02-2025
    '                CaPos2 = 0
    '                CaPos3 = 0
    '                CaPos4 = 0

    '                FindDwg = File1.Items.Item(i)
    '                ExtPos = InStr(1, FindDwg, ".dwg")
    '                MxPos1 = InStr(1, FindDwg, "_MX")
    '                MxPos2 = InStr(1, FindDwg, "_Mx")
    '                MxPos3 = InStr(1, FindDwg, "-MX")
    '                MxPos4 = InStr(1, FindDwg, "-Mx")

    '                ChPos1 = InStr(1, FindDwg, "_CH")         'DJL-------06-02-2025
    '                ChPos2 = InStr(1, FindDwg, "_Ch")
    '                ChPos3 = InStr(1, FindDwg, "-CH")
    '                ChPos4 = InStr(1, FindDwg, "-Ch")

    '                CaPos1 = InStr(1, FindDwg, "_CA")         'DJL-------06-02-2025
    '                CaPos2 = InStr(1, FindDwg, "_Ca")
    '                CaPos3 = InStr(1, FindDwg, "-CA")
    '                CaPos4 = InStr(1, FindDwg, "-Ca")

    '                If ExtPos > 0 Then
    '                    If MxPos1 = 0 And MxPos2 = 0 And MxPos3 = 0 And MxPos4 = 0 Then
    '                        If ChPos1 = 0 And ChPos2 = 0 And ChPos3 = 0 And ChPos4 = 0 Then         'DJL-------06-02-2025
    '                            If CaPos1 = 0 And CaPos2 = 0 And CaPos3 = 0 And CaPos4 = 0 Then         'DJL-------06-02-2025
    '                                DwgList.Items.Add(File1.Items.Item(i))
    '                            End If
    '                        End If
    '                    End If
    '                End If
    '            Next i
    '        End If

    'Err_Dir1_Change:
    '        ErrNo = Err.Number

    '        If ErrNo <> 0 Then
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

    '            HandleErrSQL(PrgName, ErrNo, ErrMsg, ErrSource, PriPrg, ErrDll, DwgItem, PrgLineNo)

    '            If GenInfo.UserName = "dlong" Then
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

    '    End Sub

    Private Sub BOM_Menu_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim i As Short
        PrgName = "BOM_Menu_Load"

        On Error GoTo Err_BOM_Menu_Load

        SBclicked = False

        File2.Text = "Click to Browse -->"
        File1.Focus()
        Me.Label2.Text = "Progress........"
        File2.Enabled = False
        'btnNewDir.Enabled = False
        Me.BtnGetMWInfo.BackColor = System.Drawing.Color.LawnGreen
        Me.PathBox.BackColor = System.Drawing.Color.LawnGreen

        For i = 0 To 20
            Me.ComboBox1.Items.Add(i)
        Next i

        fs = CreateObject("Scripting.FileSystemObject")

        If Directory.Exists("K:\AWA\" & System.Environment.UserName & "\AdeptWork\") = True Then
            PathBox.Text = "K:\AWA\" & System.Environment.UserName & "\AdeptWork"
            'GroupBox5.Text = "Current Directory = K:\AWA\" & System.Environment.UserName & "\AdeptWork"
        Else
            PathBox.Text = "C:\Adeptwork"
            'GroupBox5.Text = "Current Directory = C:\Adeptwork"
        End If

Err_BOM_Menu_Load:
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

            HandleErrSQL(PrgName, ErrNo, ErrMsg, ErrSource, PriPrg, ErrDll, DwgItem, PrgLineNo)

            If GenInfo.UserName = "dlong" Then
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

    End Sub

    Private Sub File1_DoubleClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
        If fs.FileExists(fs.GetFile(File1.Path & "\" & File1.FileName).Path) Then
            SelectList.Items.Add(File1.Path & "\" & File1.FileName)
        End If
    End Sub

    Private Sub BtnStart2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnStart2.Click
        Dim TempAttributes As Object
        Dim Title, Msg, Style, Response, ShipListItem As Object
        Dim JAt, z, i, j, x, ChkCnt, CountVal, CntWorkbook, WrkBooksCnt, FoundMM1, FoundMM2, BOMCnt, CollectedItemsCnt As Integer
        Dim WrkBookName, WorkShtName, NextSht, BomWrkShtNam, FileSaveAS, OldFileNam, UpdDesc, DwgNoNew2, DwgNoOld2, WorkBookName, NextChk, CurrentDwgRev, CurrentDwgNo, Chk10D, Chk10E As String
        Dim GetDesc, GetShpMk, GetCity, GetState, Get2DShipMk, BOMFile, PrevCustPO, ErrAt, GetAllParts, fixedCustomerPO, fixedPrevCustPO, PrevDwgItem, GetPieceMk As String                           'NPath, TPath, OldAttName, NewAttNameGetShpMkAtt,
        Dim GetPrev2DShipMk, GetPrevShpMkDesc, FinishCollectInfo As String
        Dim DwgNoOld, RevNoOld, PcMkOld, QtyOld, DescOld, TestOldDesc, GetReqNo, GetNext2DShipMk As String
        Dim DwgNoNew, RevNoNew, PcMkNew, QtyNew, DescNew, TestNewDesc, JobNoDash, GetNextDesc, ProblemAt As String
        Dim Workbooks As Excel.Workbooks
        Dim DwgItem3

        PrgName = "StartButton2_Click"                          '-------DJL-12-19-2024

        On Error GoTo Err_StartButton_Click

        TestOldDesc = Nothing
        FullJobNo = Nothing
        PrevCustPO = Nothing
        DescOld = Nothing
        ExceptPos = 0
        SortListing = True
        GetAllParts = "No"
        Me.BtnAdd.BackColor = System.Drawing.Color.White
        Me.BtnClear.BackColor = System.Drawing.Color.White
        Me.BtnRemove.BackColor = System.Drawing.Color.White
        Me.TxtBoxCntDown.BackColor = System.Drawing.Color.LawnGreen
        PrgName = "StartButton2-CollectingDATA"                          '-------DJL-12-19-2024

        'Turned Compare process back on per OU 212 PM request                          'DJL-12-25-2023
        If Me.ComboBox1.Items.Count > 0 Then
            If InStr(1, Me.ShipListBox.Text, ".xls") = 0 Then
                Msg = "No Shipping List was selected to compare to, do you want to continue?"
                Style = MsgBoxStyle.YesNo
                Title = "Shipping List"
                Response = MsgBox(Msg, Style, Title)
                If Response = 6 Then                    'if user clicks yes
                    Me.File2.Enabled = False
                    Me.CheckBox1.CheckState = False
                Else 'if user clicks no
                    Me.File2.Enabled = True
                    Me.CheckBox1.CheckState = CheckState.Checked
                    Exit Sub
                End If
            End If
        End If

        PrgName = "StartButton2-OpenExcel"                          '-------DJL-12-19-2024

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

        On Error GoTo Err_StartButton_Click

        If Me.ComboBox1.Text = vbNullString Then
            MsgBox("Please Select a revision number For Shipping List")
            Exit Sub
        End If

        PrgName = "StartButton2-OpenAutoCAD"                          '-------DJL-12-19-2024
        Me.Label2.Text = "Gathering Information from BOM Spreadsheet........Please Wait"           '-------DJL-07-15-2025      'Me.Label2.Text = "Gathering Information from AutoCAD Drawings........Please Wait"
        Me.Refresh()

        On Error Resume Next
        If Err.Number Then
            Err.Clear()
        End If

        On Error GoTo Err_StartButton_Click

        '-----------------------------------------------------------------------------------------------------------------
        '-------DJL-07-15-2025      'Open BOM Spreadsheet instead of AutoCAD Drawings.
        '-----------------------------------------------------------------------------------------------------------------
        BOMFile = SelectList.Items.Item(0).ToString

        ReadBOMFile = ExcelApp.Application.Workbooks.Open(PathBox.Text & BOMFile)
        Workbooks = ExcelApp.Workbooks
        ExcelApp.Application.Visible = True
        WrkBooksCnt = Workbooks.Count
        ReadBulkBOM = ReadBOMFile.Application.ActiveWorkbook.Sheets("Bulk BOM")
        BOMCnt = ReadBulkBOM.Range("A4000").End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Row

        Me.Refresh()
        ProgressBar1.Value = 0
        ProgressBar1.Maximum = BOMCnt                   '-------DJL-07-15-2025'ProgressBar1.Maximum = Count
        ProgressBar1.Visible = True                        'True                           'DJL-------06/02/2025
        CountVal = 0

        If BOMCnt <> 0 Then                          '-------DJL-07-15-2025      'If Count <> 0 Then
            PrgName = "StartButton2-CollectTags"                          '-------DJL-12-19-2024

            For z = 0 To BOMCnt                            '-------DJL-07-15-2025      'For Each DwgItem3 In VarSelArray
                TxtBoxCntDown.Text = BOMCnt - z               '-------DJL-07-15-2025      'TxtBoxCntDown.Text = Count - CountVal
                ProblemAt = "ReadingSpreadsheet"                         '-------DJL-07-15-2025      'ProblemAt = "ReadingTitleBlock"

                GetAllParts = "No"
                Me.Refresh()

                With ReadBulkBOM
                    If GenInfo3135.FullJobNo = Nothing Then
                        GenInfo3135.FullJobNo = .Range("B3").Value
                    End If

                    If GenInfo3135.CustomerPO = Nothing Then                            '-------Need to add customer information to BOM so shipping list has everything required.
                        GenInfo3135.CustomerPO = .Range("Q3").Value                '-------DJL-07-25-2025      'GenInfo3135.CustomerPO = .Range("AA3").Value
                        PrevCustPO = GenInfo3135.CustomerPO
                    End If

                    Get2DShipMk = .Range("C" & (z + 5)).Value      'Get2DShipMk
                    GetPieceMk = .Range("D" & (z + 5)).Value      'GetPieceMk

                    If Get2DShipMk <> "" And GetPieceMk = " " Then              '-------DJL-07-28-2025  'New problem GetPieceMk " "
                        If Get2DShipMk <> " " Then
                            GoTo CollectSRInfo
                        End If
                    End If

                    If Get2DShipMk <> "" And GetPieceMk = "" Then     'CurrentDwgNo <> Nothing and GetPieceMk = "-"
                        If Get2DShipMk = " " Then                           '-------DJL-07-28-2025  'New problem users have " " for Ship Mark
                            GoTo FoundItemsNeeded
                        End If

                        If Get2DShipMk = GetPrev2DShipMk And GetPieceMk = "" Then       '-------DJL-07-28-2025      'New problem Ship mark is correct but piece mark has been left blank.
                            GoTo FoundItemsNeeded
                        End If
CollectSRInfo:
                        If InStr(GenInfo3135.CustomerPO, GenInfo3135.FullJobNo) = 0 Then
                            ShippingList(1, UBound(ShippingList, 2)) = GenInfo3135.FullJobNo & "/" & GenInfo3135.CustomerPO '-------DJL-07-24-2025      'Found problem where part 1 on Old BOM is = Customer PO Number = 5606-1107/DOMINION ENERGY
                        Else
                            ShippingList(1, UBound(ShippingList, 2)) = GenInfo3135.CustomerPO
                        End If

                        ShippingList(2, UBound(ShippingList, 2)) = .Range("A" & (z + 5)).Value      'CurrentDwgNo
                        ShippingList(3, UBound(ShippingList, 2)) = .Range("B" & (z + 5)).Value      'CurrentDwgRev
                        ShippingList(4, UBound(ShippingList, 2)) = .Range("C" & (z + 5)).Value      'Get2DShipMk        '-------DJL-07-24-2025      ''ShippingList(4, UBound(ShippingList, 2)) = Get2DShipQty
                        ShippingList(5, UBound(ShippingList, 2)) = .Range("D" & (z + 5)).Value      'GetPieceMk
                        ShippingList(6, UBound(ShippingList, 2)) = .Range("E" & (z + 5)).Value      'GetQty
                        'ShippingList(7, UBound(ShippingList, 2)) = GetShipDesc

                        'If GetAllParts = "Yes" Then
                        '    ShippingList(8, UBound(ShippingList, 2)) = GetDesc
                        'Else

                        GetDesc = .Range("F" & (z + 5)).Value

                        If GetPrevShpMkDesc = "SHELL PLATE" And InStr(Get2DShipMk, "SR") > 0 Then                                   '-------DJL-07-24-2025      '   Added       'If GetPrevShpMkDesc = "SHELL PLATE" And InStr(Get2DShipMk, "SR") > 0 Then
                            ShippingList(8, UBound(ShippingList, 2)) = .Range("F" & (z + 5)).Value & " " & Get2DShipMk     'GetDesc       '-------DJL-07-24-2025      '   Added
                        Else
                            'GetPrev2DShipMk = Get2DShipMk       '-------DJL-07-24-2025
                            'z = (z + 1)                         '-------DJL-07-24-2025
                            'GoTo CollectSRInfo                  '-------DJL-07-24-2025
                            ShippingList(8, UBound(ShippingList, 2)) = .Range("F" & (z + 5)).Value      'GetDesc
                        End If

                        If InStr(ShippingList(4, (CollectedItemsCnt + 1)), "SR") = 1 Then     '-------07-24-2025      'If InStr(ShippingList(3, CollectedItemsCnt), "SR") = 1 Then        '-------7-16-2025       'If InStr(GetShpMk, "SR") = 1 Then
                            GetNext2DShipMk = .Range("C" & (z + 5 + 1)).Value
                            GetNextDesc = .Range("F" & (z + 5 + 1)).Value
                            GetDesc = .Range("F" & (z + 5)).Value

                            If GetPrev2DShipMk <> GetNext2DShipMk And GetDesc <> "SHELL PLATE" Then         '-------7-18-2025       'This will advance to next SR2 shell plate.
                                GetPrevShpMkDesc = Nothing

                                If FinishCollectInfo = "Yes" Then         '-------7-18-2025
                                    GoTo FinishCollectingInfo
                                End If

                                GoTo NextSRFound                'If Next SHELL PLATE is found.
                            Else
                                GetPrevShpMkDesc = GetDesc
                            End If

                            Select Case 0
                                Case Is < InStr(Get2DShipMk, "SR1")
FoundSR:
                                    FinishCollectInfo = "Yes"
                                    GetPrev2DShipMk = Get2DShipMk
                                    z = (z + 1)
                                    GoTo CollectSRInfo                          '-------DJL-07-17-2025      'GoTo FoundSRPlates
                                Case Is < InStr(Get2DShipMk, "SR2")
                                    GoTo FoundSR
                                Case Is < InStr(Get2DShipMk, "SR3")
                                    GoTo FoundSR
                                Case Is < InStr(Get2DShipMk, "SR4")
                                    GoTo FoundSR
                                Case Is < InStr(Get2DShipMk, "SR5")
                                    GoTo FoundSR
                                Case Is < InStr(Get2DShipMk, "SR6")
                                    GoTo FoundSR
                                Case Is < InStr(Get2DShipMk, "SR7")
                                    GoTo FoundSR
                                Case Is < InStr(Get2DShipMk, "SR8")
                                    GoTo FoundSR
                                Case Is < InStr(Get2DShipMk, "SR9")
                                    GoTo FoundSR
                                Case Is < InStr(Get2DShipMk, "SR10")
                                    GoTo FoundSR
                                Case Is < InStr(Get2DShipMk, "SR11")
                                    GoTo FoundSR
                                Case Is < InStr(Get2DShipMk, "SR12")
                                    GoTo FoundSR
                                Case Is < InStr(Get2DShipMk, "SR13")
                                    GoTo FoundSR
                                Case Is < InStr(Get2DShipMk, "SR14")
                                    GoTo FoundSR
                                Case Is < InStr(Get2DShipMk, "SR15")
                                    GoTo FoundSR
                            End Select
                        End If
FinishCollectingInfo:
                        ShippingList(9, UBound(ShippingList, 2)) = .Range("G" & (z + 5)).Value      'GetInv1
                        ShippingList(10, UBound(ShippingList, 2)) = .Range("H" & (z + 5)).Value      'GetInv2
                        ShippingList(11, UBound(ShippingList, 2)) = .Range("I" & (z + 5)).Value      'GetMat
                        'ShippingList(12, UBound(ShippingList, 2)) = GetMat2
                        'ShippingList(13, UBound(ShippingList, 2)) = Getmat3
                        ShippingList(14, UBound(ShippingList, 2)) = .Range("J" & (z + 5)).Value      'GetWt
                        'ShippingList(15, UBound(ShippingList, 2)) = CurrentDwgNo

                        'If GetPrevShpMkDesc = Nothing Then
                        '    ShippingList(16, UBound(ShippingList, 2)) = .Range("F" & (z + 5)).Value        'No only collect when is not Nothing
                        'Else

                        ShippingList(16, UBound(ShippingList, 2)) = GetPrevShpMkDesc
                        'End If

                        ShippingList(17, UBound(ShippingList, 2)) = CurrentDwgNo

                        ShippingList(18, UBound(ShippingList, 2)) = .Range("M" & (z + 5)).Value      'GetProcurement
                        ReDim Preserve ShippingList(18, UBound(ShippingList, 2) + 1)
                        CollectedItemsCnt = (CollectedItemsCnt + 1)                     '-------7-16-2025
                        GetAllParts = "No"
                        FinishCollectInfo = "No"

                        '-------DJL-07-24-2025      'Moved Above.
                        'If InStr(ShippingList(4, CollectedItemsCnt), "SR") = 1 Then     '-------07-24-2025      'If InStr(ShippingList(3, CollectedItemsCnt), "SR") = 1 Then        '-------7-16-2025       'If InStr(GetShpMk, "SR") = 1 Then
                        '    GetNext2DShipMk = .Range("C" & (z + 5 + 1)).Value
                        '    GetNextDesc = .Range("F" & (z + 5 + 1)).Value
                        '    GetDesc = .Range("F" & (z + 5)).Value

                        '    If GetPrev2DShipMk <> GetNext2DShipMk And GetDesc <> "SHELL PLATE" Then         '-------7-18-2025       'This will advance to next SR2 shell plate.
                        '        GetPrevShpMkDesc = Nothing
                        '        GoTo NextSRFound                'If Next SHELL PLATE is found.
                        '    Else
                        '        GetPrevShpMkDesc = GetDesc
                        '    End If

                        '    Select Case 0
                        '        Case Is < InStr(Get2DShipMk, "SR1")
                        '            GetPrev2DShipMk = Get2DShipMk
                        '            z = (z + 1)
                        '            GoTo CollectSRInfo                          '-------DJL-07-17-2025      'GoTo FoundSRPlates
                        '        Case Is < InStr(Get2DShipMk, "SR2")
                        '            GetPrev2DShipMk = Get2DShipMk
                        '            z = (z + 1)
                        '            GoTo CollectSRInfo                          '-------DJL-07-17-2025      'GoTo FoundSRPlates
                        '        Case Is < InStr(Get2DShipMk, "SR3")
                        '            GetPrev2DShipMk = Get2DShipMk
                        '            z = (z + 1)
                        '            GoTo CollectSRInfo                          '-------DJL-07-17-2025      'GoTo FoundSRPlates
                        '        Case Is < InStr(Get2DShipMk, "SR4")
                        '            GetPrev2DShipMk = Get2DShipMk
                        '            z = (z + 1)
                        '            GoTo CollectSRInfo                          '-------DJL-07-17-2025      'GoTo FoundSRPlates
                        '        Case Is < InStr(Get2DShipMk, "SR5")
                        '            GetPrev2DShipMk = Get2DShipMk
                        '            z = (z + 1)
                        '            GoTo CollectSRInfo                          '-------DJL-07-17-2025      'GoTo FoundSRPlates
                        '        Case Is < InStr(Get2DShipMk, "SR6")
                        '            z = (z + 1)
                        '            GoTo CollectSRInfo                          '-------DJL-07-17-2025      'GoTo FoundSRPlates
                        '        Case Is < InStr(Get2DShipMk, "SR7")
                        '            GetPrev2DShipMk = Get2DShipMk
                        '            z = (z + 1)
                        '            GoTo CollectSRInfo                          '-------DJL-07-17-2025      'GoTo FoundSRPlates
                        '        Case Is < InStr(Get2DShipMk, "SR8")
                        '            GetPrev2DShipMk = Get2DShipMk
                        '            z = (z + 1)
                        '            GoTo CollectSRInfo                          '-------DJL-07-17-2025      'GoTo FoundSRPlates
                        '        Case Is < InStr(Get2DShipMk, "SR9")
                        '            GetPrev2DShipMk = Get2DShipMk
                        '            z = (z + 1)
                        '            GoTo CollectSRInfo                          '-------DJL-07-17-2025      'GoTo FoundSRPlates
                        '        Case Is < InStr(Get2DShipMk, "SR10")
                        '            GetPrev2DShipMk = Get2DShipMk
                        '            z = (z + 1)
                        '            GoTo CollectSRInfo                          '-------DJL-07-17-2025      'GoTo FoundSRPlates
                        '        Case Is < InStr(Get2DShipMk, "SR11")
                        '            GetPrev2DShipMk = Get2DShipMk
                        '            z = (z + 1)
                        '            GoTo CollectSRInfo
                        '        Case Is < InStr(Get2DShipMk, "SR12")
                        '            GetPrev2DShipMk = Get2DShipMk
                        '            z = (z + 1)
                        '            GoTo CollectSRInfo
                        '        Case Is < InStr(Get2DShipMk, "SR13")
                        '            GetPrev2DShipMk = Get2DShipMk
                        '            z = (z + 1)
                        '            GoTo CollectSRInfo
                        '        Case Is < InStr(Get2DShipMk, "SR14")
                        '            GetPrev2DShipMk = Get2DShipMk
                        '            z = (z + 1)
                        '            GoTo CollectSRInfo
                        '        Case Is < InStr(Get2DShipMk, "SR15")
                        '            GetPrev2DShipMk = Get2DShipMk
                        '            z = (z + 1)
                        '            GoTo CollectSRInfo
                        '    End Select
                        'End If
                    End If
                End With

FoundItemsNeeded:
                If CustomerPO = "" Then
                    CustomerPO = UCase(GenInfo3135.FullJobNo)              '-------CustomerPO = UCase(CustomerPO)
                End If

                If CustomerPO <> PrevCustPO And CustomerPO <> "" Then               '-------Tulsa Old Titleblocks.
                    If IsNothing(PrevCustPO) = True Then
                        PrevCustPO = CustomerPO
                    Else
                        If IsNothing(PrevCustPO) = True Then
                            CustomerPO = InputBox("This job has two customer names please type In the correct Customer Name ? " & CustomerPO & " Or " & PrevCustPO)
                            CustomerPO = UCase(CustomerPO)
                            PrevCustPO = CustomerPO
                        Else
                            If CustomerPO <> PrevCustPO And PrevCustPO <> Nothing Then
                                fixedCustomerPO = CustomerPO.Replace(" ", "")           'USERS MUST HAVE EXTRA SPACEZ ON DESCRIPTION.
                                fixedPrevCustPO = PrevCustPO.Replace(" ", "")
                            End If

                            If fixedCustomerPO <> fixedPrevCustPO Then
                                If InStr(PrevCustPO, CustomerPO) = 0 Then
                                    PrevCustPO = CustomerPO & "/" & PrevCustPO                          '-------DJL-07-16-2025
                                    GenInfo3135.CustomerPO = PrevCustPO
                                End If
                            End If
                        End If
                    End If
                Else                            '-------Pittsburgh Changes in Title Block -------DJL-------12-18-2023
FoundMPDMSTDTITLE:
                    If PrevCustPO = "" Then
                        PrevCustPO = Cust1                         '-------PrevCustPO = CustomerPO
                    Else
                        If PrevCustPO <> Cust1 And Cust1 <> "" Then
                            PrevCustPO = Cust1
                        End If
                    End If

                    If CustomerPO = "" Then
                        CustomerPO = UCase(GenInfo3135.FullJobNo)              '-------CustomerPO = UCase(CustomerPO)
                    End If

                    If IsNothing(PrevCustPO) = True Then
                        If PrevCustPO = "" Then
                            PrevCustPO = InputBox("This job has no customer name On the spreadsheet, please type In the correct Customer Name For Job No ? " & CustomerPO)
                        Else
                            PrevCustPO = InputBox("This job has two customer names please type In the correct Customer Name ? " & CustomerPO & " Or " & PrevCustPO)
                        End If

                        PrevCustPO = UCase(PrevCustPO)
                        PrevCustPO = CustomerPO & "/" & PrevCustPO                          '-------DJL-07-16-2025
                        GenInfo3135.CustomerPO = PrevCustPO                                 '-------DJL-07-16-2025
                    Else
                        If CustomerPO <> PrevCustPO And PrevCustPO <> Nothing Then
                            fixedCustomerPO = CustomerPO.Replace(" ", "")           'USERS MUST HAVE EXTRA SPACEZ ON DESCRIPTION.
                            fixedPrevCustPO = PrevCustPO.Replace(" ", "")
                        End If

                        If fixedCustomerPO <> fixedPrevCustPO And fixedCustomerPO <> GenInfo3135.FullJobNo Then
                            CustomerPO = (CustomerPO & "/" & PrevCustPO & " - " & GetCity & GetState)       '-------DJL-06-30-2025
                            GenInfo3135.CustomerPO = CustomerPO                         '-------DJL-06-30-2025
                        End If
                    End If
                End If
FoundSRPlates:

NextSRFound:                            '-------DJL-07-18-2025      'When new Shell Ring is found go to normal collection.

                If z > 0 Then
                    ProgressBar1.Value = z - 1                                 '-------DJL-07-16-2025      'ProgressBar1.Value = CountVal
                End If

                If GetDesc <> "SHELL PLATE" Then
                    GetDesc = Nothing
                End If

                If Get2DShipMk <> "" Then
                    GetPrev2DShipMk = Get2DShipMk
                End If
NextDwg:
            Next z                          '-------DJL-07-15-2025      'DwgItem3

            ProgressBar1.Value = 0              '------------Export Shipping List info to Ship List now.
            Me.Label2.Text = "Outputting Information To Shipping List........Please Wait."
            Me.Refresh()
        End If                          '-------Moved here from after WriteToExcel

        '-----------------------------------------------------------------------------------------------------------------------------
        PrgName = "StartButton2-WriteToExcel"                          '-------DJL-12-19-2024
        WriteToExcel(ShippingList)
        ShippingList = GenInfo3135.ShippingList
        ProgressBar1.Value = 0

        RevNo = Me.ComboBox1.Text
        RevNo2 = RevNo
        WorkBookName = MainBOMFile.Application.ActiveWorkbook.Name
        OldFileNam = PathBox.Text

        '-----------------------------------------------------------------------------------------------------------------
        '-------User copied STD BOM line item and added metric information when he should have been using BOM for Metric.
        '-------DJL-07-21-2025      'Program is looking for dupliate parts were one is a metric version of Part.
        '-----------------------------------------------------------------------------------------------------------------
        CntItems = ShipListSht.Range("H4000").End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Row      'CntItems = (UBound(ShippingList, 2) - 1)
        StrLineNo = GenInfo3135.StrLineNo                       '-------DJL-07-21-2025      'StrLineNo is equal to zero.
        PrgName = "StartButton2-ReadSpreadSht"
        ProgressBar1.Maximum = CntItems                          '-------DJL-07-21-2025

        With ShipListSht
            For j = 1 To CntItems
                If StrLineNo = 42 Then
                    RowNo = (42 + j)
                Else
                    RowNo = (44 + j)
                End If

                If j = 1 Then
                    DwgNoOld = .Range("D" & RowNo).Value              'Dwg     
                    RevNoOld = .Range("E" & RowNo).Value              'Rev      
                    PcMkOld = .Range("F" & RowNo).Value              'Pc Mark  
                    QtyOld = .Range("G" & RowNo).Value              'Qty      
                    DescOld = .Range("H" & RowNo).Value             'Desc      
                    FoundMM1 = InStr(DescOld, "MM")                 'MM

                    TestOldDesc = DwgNoOld & RevNoOld & PcMkOld & QtyOld
                Else
                    DwgNoNew = .Range("D" & RowNo).Value               'Dwg 
                    RevNoNew = .Range("E" & RowNo).Value              'Rev     
                    PcMkNew = .Range("F" & RowNo).Value              'Pc Mark   
                    QtyNew = .Range("G" & RowNo).Value          'Qty           
                    DescNew = .Range("H" & RowNo).Value          'Desc    
                    FoundMM2 = InStr(DescNew, "MM")

                    TestNewDesc = DwgNoNew & RevNoNew & PcMkNew & QtyNew

                    TestOldDesc = TestOldDesc

                    If TestOldDesc = TestNewDesc And TestOldDesc <> "" Then
                        If DescOld = DescNew Then           'Do not delete duplicates they are allowed on same drawing
                            DwgNoOld = DwgNoNew              'Dwg
                            RevNoOld = RevNoNew              'Rev
                            PcMkOld = PcMkNew              'Pc Mark
                            QtyOld = QtyNew              'Qty
                            DescOld = DescNew
                            TestOldDesc = DwgNoOld & RevNoOld & PcMkOld & QtyOld
                            j = (j - 1)
                            CntItems = (CntItems - 1)
                        Else
                            If FoundMM1 > 0 Then
                                .Range("A" & (RowNo - 1) & ":" & "T" & (RowNo - 1)).Delete(Shift:=XlDeleteShiftDirection.xlShiftUp)
                                '-------DJL will need to create a delete for Array.
                                DwgNoOld = DwgNoNew              'Dwg
                                RevNoOld = RevNoNew              'Rev
                                PcMkOld = PcMkNew              'Pc Mark
                                QtyOld = QtyNew              'Qty
                                DescOld = DescNew
                                TestOldDesc = DwgNoOld & RevNoOld & PcMkOld & QtyOld
                                j = (j - 1)
                                CntItems = (CntItems - 1)
                                If FoundMM2 > 0 Then
                                    .Range("A" & (RowNo - 1) & ":" & "T" & (RowNo - 1)).Delete(Shift:=XlDeleteShiftDirection.xlShiftUp)
                                    '-------DJL will need to create a delete for Array.
                                    DwgNoOld = DwgNoNew              'Dwg
                                    RevNoOld = RevNoNew              'Rev
                                    PcMkOld = PcMkNew              'Pc Mark
                                    QtyOld = QtyNew              'Qty
                                    DescOld = DescNew
                                    TestOldDesc = DwgNoOld & RevNoOld & PcMkOld & QtyOld
                                    j = (j - 1)
                                    CntItems = (CntItems - 1)
                                End If
                            End If
                        End If
                    Else
                        DwgNoOld = DwgNoNew              'Dwg
                        RevNoOld = RevNoNew              'Rev
                        PcMkOld = PcMkNew              'Pc Mark
                        QtyOld = QtyNew              'Qty
                        DescOld = DescNew
                        TestOldDesc = DwgNoOld & RevNoOld & PcMkOld & QtyOld
                    End If
                End If
                If j >= CntItems Then
                    GoTo EndProcess
                End If

                Me.ProgressBar1.Value = j                           '-------DJL-07-21-2025      'Me.ProgressBar1.Value = RowNo
                Me.TxtBoxCntDown.Text = (CntItems - j) '-------DJL-07-28-2025      'Me.TxtBoxCntDown.Text = ((CntItems - StrLineNo) - j)
                Me.Label2.Text = "Looking for Metric BOM Items that are duplicates on Shipping List........Please Wait."
                Me.Refresh()
            Next j
        End With
EndProcess:


        '-------DJL-------10-31-2023----------------------------------------------------------------Move to function CreateSpSht
        '------------------------------------------------Need to remove extra Line items that are referenced two times.
        '-------221-22-00103 and 221-22-00102 and 221-22-00107 has duplicated items on drawings 33A and 33B.
        '-------30" SHELL MIXER MANWAY
        '-------30" SHELL MIXER MANWAY REPAD
        '-------N14 and N14R

        'CntItems = CountBOM                            '-------Done above.
        PrgName = "StartButton2-RemoveDup"                          '-------DJL-12-19-2024
        CntItems = ShipListSht.Range("H4000").End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Row      '-------DJL-07-25-2025      'Added
        Me.ProgressBar1.Maximum = CntItems              'CountBOM
        Label2.Text = "Remove Duplicated Items"         'Example Job 221-22-00107   Dwgs 30A, 30B ---> Items, M11, M11R

        '-------DJL-------10-31-2023----------------------------------------------------------------Move to function CreateSpSht
        With ShipListSht
            JAt = 1
StartNextCheck:
            If NextChk = "Yes" Then
                GoTo GetNextItem
            End If

GetNextItem: For j = JAt To CntItems

                If StrLineNo = 42 Then
                    RowNo = (42 + j)
                Else
                    RowNo = (44 + j)
                End If

                If NextChk = "Yes" Then
                    NextChk = "No"
                    GoTo NextItemToFind
                End If

                If j = 1 Then
NextItemToFind:
                    DwgNoOld = .Range("D" & RowNo).Value
                    RevNoOld = .Range("E" & RowNo).Value
                    PcMkOld = .Range("F" & RowNo).Value
                    QtyOld = .Range("G" & RowNo).Value
                    DescOld = .Range("H" & RowNo).Value
                    Me.ProgressBar1.Value = j
                    JAt = j

                    If DwgNoOld = "1 = NEW ENTRY" Then
                        GoTo EndProcess2
                    End If

                    If JAt > CntItems Then
                        GoTo EndProcess2
                    End If

                    If RowNo > (CntItems + StrLineNo) Then
                        GoTo EndProcess2
                    End If

                    If IsNothing(DescOld) = True And IsNothing(PcMkOld) = True Then
                        If RowNo = CntItems Then
                            CntItems = (CntItems - 1)
                        End If

                        If RowNo > CntItems Then
                            If JAt = 1 Then
                                j = 1
                            Else
                                j = JAt
                            End If

                            If RowNo > CntItems And JAt + 44 >= CntItems Then
                                GoTo FoundAllItems
                            End If

                            GoTo NextJAt
                        End If

                        GoTo NextJ2
                    End If
                Else
                    If RowNo > (CntItems + StrLineNo) And JAt > 0 Then
                        JAt = (JAt + 1)
                        GoTo NextJ2
                    End If

                    'DwgNoNew = .Range("D" & RowNo).Value             ''Moved below
                    'RevNoNew = .Range("E" & RowNo).Value           
                    PcMkNew = .Range("F" & RowNo).Value
                    'QtyNew = .Range("G" & RowNo).Value         
                    DescNew = .Range("H" & RowNo).Value

                    If IsNothing(DescOld) = True And IsNothing(PcMkOld) = True Then
                        If RowNo = CntItems Then
                            CntItems = (CntItems - 1)
                        End If

                        If RowNo > CntItems Then
                            If JAt = 1 Then
                                j = 1
                            Else
                                j = JAt
                            End If

                            GoTo NextJAt
                        End If

                        GoTo NextJ2
                    End If

                    If IsNothing(DescNew) = True And IsNothing(PcMkNew) = True Then
                        If RowNo = CntItems Then
                            CntItems = (CntItems - 1)
                        End If

                        If RowNo > CntItems Then
                            If JAt = 1 Then
                                j = 1
                            Else
                                j = JAt
                            End If

                            GoTo NextJAt
                        End If

                        GoTo NextJ2
                    End If

                    '-------DJL-07-21-2025      'Drawing numbers can be a total of 8 Characters Now for Tulsa and Pittsburgh.
                    DwgNoNew = .Range("D" & RowNo).Value            'This works for Tulsa, but sucks for Pittsburg.
                    'Chk10D = InStr(DwgNoNew, "10D")
                    'Chk10E = InStr(DwgNoNew, "10E")
                    'LenDwgNo2 = Len(DwgNoNew)

                    'If Chk10D > 0 And LenDwgNo2 > 4 Then
                    '    DwgNoNew2 = Mid(DwgNoNew, 4, Len(DwgNoNew))                         '-------Pittsburgh
                    '    DwgNoOld2 = Mid(DwgNoOld, 4, Len(DwgNoOld))
                    '    ChkCnt = 0
                    'Else
                    'If Chk10E > 0 And LenDwgNo2 > 4 Then
                    '    If Chk10E > 0 And InStr(DwgNoOld, "10D") = 1 Then
                    '        GoTo NextJ2
                    '    End If

                    '    DwgNoNew2 = Mid(DwgNoNew, 4, Len(DwgNoNew))                         '-------Pittsburgh
                    '    DwgNoOld2 = Mid(DwgNoOld, 4, Len(DwgNoOld))
                    '    ChkCnt = 0
                    'Else

                    'DwgNoNew2 = Mid(DwgNoNew, 1, 5)        '-------Tulsa   '-------DJL-06-30-2025  10D and 10E are not on drawings and Tulsa has started using Pitts Numbers.      'DwgNoNew2 = Mid(DwgNoNew, 1, 2) 
                    'DwgNoOld2 = Mid(DwgNoOld, 1, 5)        '-------DJL-06-30-2025  10D and 10E are not on drawings and Tulsa has started using Pitts Numbers.     'DwgNoOld2 = Mid(DwgNoOld, 1, 2)
                    ChkCnt = 0
                    'End If
                    'End If

                    If DwgNoOld = DwgNoNew Then       '-------DJL-07-21-2025      'If DwgNoOld2 = DwgNoNew2 Then
                        If PcMkOld = PcMkNew And IsNothing(PcMkOld) = False Then     'PartNo, And Description are equal then delete
                            'DwgNoNew = .Range("D" & RowNo).Value           '-------DJL-07-21-2025      'Not needed.
                            RevNoNew = .Range("E" & RowNo).Value
                            QtyNew = .Range("G" & RowNo).Value
ChkDesc:
                            If DescOld = DescNew And IsNothing(DescOld) = False Then
                                If RevNoOld = RevNoNew And IsNothing(RevNoOld) = False Then
                                    If QtyOld = QtyNew And IsNothing(QtyOld) = False Then
                                        .Range("A" & RowNo & ":X" & RowNo).Delete(Shift:=XlDeleteShiftDirection.xlShiftUp)
                                        '-------DJL-------11-29-2023-------Need to delete item from the array.
                                        CntItems = (CntItems - 1)

                                        'On 221-22-00101 found that drawing reference on 33B should have been N15 instead of N14
                                        'Only remove one duplicate the rest could be referencing the wrong part number.
                                        If JAt = 1 Then
                                            j = 1
                                        Else
                                            j = JAt
                                        End If

                                        JAt = (JAt + 1)
                                        NextChk = "Yes"
                                        GoTo StartNextCheck
                                    End If
                                End If
                            Else
                                If DescOld <> DescNew And ChkCnt < 2 Then
                                    DescNew = LTrim(DescNew)             '-------221-22-00107 dwgs 39A thru 40B Item M4 The Description had an extra space at the end of the decription.
                                    DescNew = RTrim(DescNew)

                                    DescOld = LTrim(DescOld)
                                    DescOld = RTrim(DescOld)
                                    ChkCnt = (ChkCnt + 1)
                                    GoTo ChkDesc
                                End If
                            End If

                            ChkCnt = 0
                        Else
                            GoTo NextJ2
                        End If
                    Else
                        If DwgNoOld < DwgNoNew Then       '-------DJL-07-21-2025      'If DwgNoOld2 < DwgNoNew2 Then
                            If JAt = 1 Then
                                j = 1
                            Else
                                j = JAt
                            End If

                            JAt = (JAt + 1)
                            NextChk = "Yes"
                            GoTo StartNextCheck             '-------221-22-00101 N14
                        Else                                '-------If Pittsburgh is going from 10D to 10E then go to next drawing.
                            If Mid(DwgNoOld, 1, 3) = "10D" And Mid(DwgNoNew, 1, 3) = "10E" Then
                                If JAt = 1 Then
                                    j = 1
                                Else
                                    j = JAt
                                End If

                                JAt = (JAt + 1)
                                NextChk = "Yes"
                                GoTo StartNextCheck             '-------221-22-00101 N14
                            End If
                        End If
                    End If
                End If
                If j >= CntItems And JAt = CntItems Then
                    GoTo EndProcess2
                End If
NextJ2:
                ChkCnt = 0
                Me.TxtBoxCntDown.Text = ((CntItems - StrLineNo) - j)
                Me.Label2.Text = "Looking for duplicate ship marks on Shipping List........Please Wait."
                Me.Refresh()
            Next j
NextJAt:
            If JAt < RowNo And JAt > 0 Then
                JAt = (JAt + 1)
                NextChk = "Yes"

                If RowNo < (CntItems + StrLineNo) Then
                    Me.ProgressBar1.Value = JAt
                End If

                GoTo StartNextCheck
            End If
FoundAllItems:
        End With

        Me.ProgressBar1.Value = 0
EndProcess2:

        '----------------------------------------------------------------------------------------------------
        'Turn Compare process back on per OU 212 PM manager                          'DJL-210-25-2023
        '----------------------------------------------------------------------------------------------------
        PrgName = "StartButton2-CompareSpreadSht"                          '-------DJL-12-19-2024
        If Me.CheckBox1.CheckState = 1 Then
            NewShipListSht = ExcelApp.Application.ActiveWorkbook.Sheets("Shipping List")

            If PathBox.Text <> PathBox2.Text Then
                OldShipListFileStr = PathBox2.Text & ShipListBox.SelectedItem.ToString
            Else
                OldShipListFileStr = PathBox.Text & ShipListBox.SelectedItem.ToString            'OldShipListFileStr = PathBox.Text & "\" & ShipListBox.SelectedItem
            End If

            Me.Refresh()
            '-------DJL-------10-31-2023----------------------------------------------------------------Move to function CreateSpSht
            CompareShipList()       '-------------Compare program look for Deleted, New Items ETC.
        Else
            If CurrentDwgRev = "0" Then
                '----------------------------------------Fix All Lines to Green when Revision Is Zero....
                NewShipListSht = ExcelApp.Application.ActiveWorkbook.Sheets("Shipping List")
                FormatShipListRev0(NewBOM, OldBOM, ShipListSht, BOMSheet, CntItems, StrLineNo)
            End If
        End If

        PrgName = "StartButton2-CompareDone"                          '-------DJL-12-19-2024
        ProgressBar1.Value = 0
        ProgressBar1.Maximum = 5
        ProgressBar1.Visible = True                        'True                           'DJL-------06/02/2025
        Me.Label2.Text = "Processing Information........Please Wait."
        ProgressBar1.Value = 1
        RevNo = RevNo2
        CopyBOMFile(OldFileNam, RevNo)
        'AcadApp.ActiveDocument.Close()
Cancel:

        '-------DJL-------10-31-2023----------------------------------------------------------------Now write out to Excel.
        PrgName = "StartButton2-CloseAutoCAD"                          '-------DJL-12-19-2024
        ProgressBar1.Value = 2
        ExcelApp.Application.Visible = True

        PrgName = "StartButton2-CloseSpreadSht"                          '-------DJL-12-19-2024
        ProgressBar1.Value = 3
        Workbooks = ExcelApp.Workbooks
        CntWorkbook = Workbooks.Count
        WorkBookName = MainBOMFile.Application.ActiveWorkbook.Name
        Dim WBook As Excel.Workbook
        Dim TempFile, TempFile2 As Workbook
        Dim TempName As String
        TempFile2 = Nothing

        For Each WBook In ExcelApp.Workbooks
            If InStr(WBook.Name, "ShipListVBNet") > 0 Then
                WBook.Close(False)
            End If
        Next

        ProgressBar1.Value = 4

        If IsNothing(TempFile2) <> True Then
            TempFile.Close(False)
            TempFile = Nothing
        End If


        ProgressBar1.Value = 5
        Me.Label2.Text = "Your Shipping List has been Created."
        MsgBox("Your Shipping List has been Created.")
        ExcelApp.Application.Visible = True
        Me.Close()

Err_StartButton_Click:
        ErrNo = Err.Number

        If ErrNo <> 0 Then
            PriPrg = "ShipListReadBOMAutoCAD"
            ErrMsg = Err.Description
            ErrSource = Err.Source
            ErrDll = Err.LastDllError
            ErrLastLineX = Err.Erl
            ErrException = Err.GetException

            If ErrNo = -2145320885 And ErrMsg = "Problem in unloading DVB file" Then
                Resume Next
            End If

            If ErrNo = -2145320885 And InStr(ErrMsg, "Exception from HRESULT") > 0 Then
                Threading.Thread.Sleep(35)
                Resume Next
            End If

            If ErrNo = 91 And ErrMsg = "Problem in unloading DVB file" Then
                Resume Next
            End If

            If ErrMsg = "Cannot create ActiveX component." Then
                MsgBox("This program is having a problem opening AutoCAD, Please open AutoCAD then pick this OK button.")
                AcadApp = GetObject(, "AutoCAD.Application")
                Threading.Thread.Sleep(25)
                Resume Next
            End If

            If ErrNo = -2147418111 And Mid(ErrMsg, 1, 28) = "Call was rejected by callee." Then
                AcadApp = GetObject(, "AutoCAD.Application")
                Err.Clear()
                Threading.Thread.Sleep(25)
                Resume
            End If

            If ErrNo = 91 And Mid(ErrMsg, 1, 17) = "The RPC server is" Then
                AcadApp = CreateObject("AutoCAD.Application")
                Threading.Thread.Sleep(25)
                Resume
            End If

            If ErrNo = 91 And ErrMsg = "Problem in unloading DVB file" Then
                Resume Next
            End If

            If ErrNo = 91 And ErrMsg = "" Then
                GoTo EndPrg
            End If

            If ErrNo = 9 And InStr(ErrMsg, "Index was outside the bounds of the array.") > 0 Then
                Resume Next
            End If

            Dim st As New StackTrace(Err.GetException, True)
            CntFrames = st.FrameCount
            GetFramesSrt = st.GetFrames
            PrgLineNo = GetFramesSrt(CntFrames - 1).ToString
            PrgLineNo = PrgLineNo.Replace("@", "at")
            LenPrgLineNo = (Len(PrgLineNo))
            PrgLineNo = Mid(PrgLineNo, 1, (LenPrgLineNo - 2))

            HandleErrSQL(PrgName, ErrNo, ErrMsg, ErrSource, PriPrg, ErrDll, DwgItem3, PrgLineNo)

            'If ErrNo = "462" And Mid(ErrMsg, 1, 30) = "The RPC server is unavailable." Then
            '    If Err.Number Then
            '        Err.Clear()
            '    End If

            '    If IsNothing(AcadApp) = True Then
            '        AcadApp = CreateObject("AutoCAD.Application")
            '    Else
            '        Err.Clear()
            '        MsgBox("Please Open AutoCAD then pick this OK button.")
            '        AcadApp = GetObject(, "AutoCAD.Application")
            '        AcadOpen = True
            '    End If

            '    If Err.Number Then
            '        MsgBox(Err.Description)
            '        Information.Err.Clear()
            '        AcadApp = CreateObject("AutoCAD.Application")
            '    End If

            '    If ProblemAt = "ReadingTitleBlock" Then
            '        ProblemAt = "ReadingTitleBlock2"
            '        Resume Next
            '    End If

            '    Resume
            'End If

            'If ErrNo = "-2147418113" And Mid(ErrMsg, 1, 30) = "Internal application error." Then
            '    If Err.Number Then
            '        Err.Clear()
            '    End If

            '    If IsNothing(AcadApp) = True Then
            '        AcadApp = CreateObject("AutoCAD.Application")
            '        Threading.Thread.Sleep(25)
            '        AcadOpen = True
            '        Resume
            '    Else
            '        MsgBox("Please Open AutoCAD then pick this OK button.")
            '        AcadApp = GetObject(, "AutoCAD.Application")
            '        Threading.Thread.Sleep(25)
            '        AcadOpen = True
            '        Resume
            '    End If
            'End If

            If ErrNo = "20" And ErrMsg = "Resume without error." Then
                Exit Sub
            End If

            If ErrNo = "-2145320900" And ErrMsg = "Failed to get the Document object" Then
                If ErrAt = "Unload Adept" Then
                    AcadApp.Documents.Add()
                    Resume
                End If
            End If

            If IsNothing(GenInfo.UserName) = True Then
                GenInfo.UserName = Environment.UserName
            End If

            If GenInfo.UserName = "dlong" Then
                ExceptPos = 0
                SearchException = "Exception"
                ExceptPos = InStr(ErrMsg, 1)
                If ExceptPos > 0 Then
                    CntExcept = (CntExcept + 1)
                    If CntExcept < 20 Then
                        If IsNothing(AcadApp) = True Then
                            AcadApp = GetObject(, "AutoCAD.Application")
                            Threading.Thread.Sleep(25)
                        End If

                        Resume
                    End If
                End If

                MsgBox(ErrMsg)
                Stop
                Resume
            Else
                ExceptPos = 0
                SearchException = "Exception"
                ExceptPos = InStr(ErrMsg, 1)
                If ExceptPos > 0 Then
                    CntExcept = (CntExcept + 1)
                    If CntExcept < 20 Then
                        If IsNothing(AcadApp) = True Then
                            AcadApp = GetObject(, "AutoCAD.Application")
                            Threading.Thread.Sleep(25)
                        End If
                        Resume
                    End If
                End If
            End If
        End If
EndPrg:

    End Sub

    Public Function CompareBOM() As Object
        Dim Excel As Object
        Dim Test As String
        Dim OldBOMTest As Boolean, OldPlateTest As Boolean, OldStickTest As Boolean, OldPurchaseTest As Boolean
        Dim OldGratingTest As Boolean, OldStickImportTest As Boolean
        Dim Workbooks As Excel.Workbooks
        Dim WrkBooksCnt As Integer
        Dim BOMMnu As ShippingList_Menu
        BOMMnu = Me

        On Error Resume Next

        Excel = GetObject(, "Excel.Application")

        If Err.Number Then
            Information.Err.Clear()
            Excel = CreateObject("Excel.Application")
            If Err.Number Then
                MsgBox(Err.Description)
                Exit Function
            End If
        End If

        Workbooks = Excel.Workbooks
        WrkBooksCnt = Workbooks.Count
        NewBulkBOM = MainBOMFile.Application.ActiveWorkbook.Sheets("Bulk BOM")
        'NewGratingBOM = MainBOMFile.Application.ActiveWorkbook.Sheets("Grating BOM")
        'NewPlateBOM = MainBOMFile.Application.ActiveWorkbook.Sheets("Plate BOM")
        'NewStickBOM = MainBOMFile.Application.ActiveWorkbook.Sheets("Stick BOM")
        'NewStickImportBOM = MainBOMFile.Application.ActiveWorkbook.Sheets("Stick Import")
        'NewPurchaseBOM = MainBOMFile.Application.ActiveWorkbook.Sheets("Other BOM")

        '-------------------------------------------------------------Open previous bulk bom to compare.
        OldBOMFile = Excel.Application.Workbooks.Open(OldBulkBOMFile)
        OldBulkBOM = OldBOMFile.Application.ActiveWorkbook.Sheets("Bulk BOM")
        'OldGratingBOM = OldBOMFile.Application.ActiveWorkbook.Sheets("Grating BOM")
        'OldPlateBOM = OldBOMFile.Application.ActiveWorkbook.Sheets("Plate BOM")
        'OldStickBOM = OldBOMFile.Application.ActiveWorkbook.Sheets("Stick BOM")
        'OldStickImportBOM = OldBOMFile.Application.ActiveWorkbook.Sheets("Stick Import")
        'OldPurchaseBOM = OldBOMFile.Application.ActiveWorkbook.Sheets("Other BOM")

        '-----------------------Check to see if selected file is a bulk bom-----"Verify Headers are the same.
        OldBOMTest = Comparison.InputType3.CheckOldBOM(OldBulkBOM)
        OldGratingTest = Comparison.InputType3.CheckOldBOM(OldGratingBOM)
        OldPlateTest = Comparison.InputType3.CheckOldBOM(OldPlateBOM)
        OldStickTest = Comparison.InputType3.CheckOldBOM(OldStickBOM)
        OldPurchaseTest = Comparison.InputType3.CheckOldBOM(OldPurchaseBOM)

        If OldBOMTest = True And OldGratingTest = True And OldPlateTest = True And OldStickTest = True And OldPurchaseTest = True Then
            NewBulkBOM.Activate()
            NewBOM = Nothing
            OldBOM = Nothing
            Comparison.InputType3.ReadBOM(NewBOM, NewBulkBOM)
            NewBOM = NewBOM
            Comparison.InputType3.ReadBOM(OldBOM, OldBulkBOM)
            OldBOM = OldBOM
            BOMSheet = "Bulk BOM"

            'If BOMType = "TANK" Then                           '-------DJL-------10-31-2023---Not needed anymore.
            Comparison.InputType3.CompareArraysTank(NewBOM, OldBOM)
            'ElseIf BOMType = "SEAL" Then
            '    Comparison.InputType3.CompareArraysSeal(NewBOM, OldBOM)
            'End If

            FormatNewBulkBOM(NewBOM, OldBOM, NewBulkBOM, BOMSheet)
            NewPlateBOM.activate()
            NewBOM = Nothing
            OldBOM = Nothing
            Comparison.InputType3.ReadBOM(NewBOM, NewPlateBOM)
            NewBOM = NewBOM
            Comparison.InputType3.ReadBOM(OldBOM, OldPlateBOM)
            OldBOM = OldBOM
            BOMSheet = "Plate BOM"

            'If BOMType = "TANK" Then                           '-------DJL-------10-31-2023---Not needed anymore.
            Comparison.InputType3.CompareArraysTank(NewBOM, OldBOM)
            'ElseIf BOMType = "SEAL" Then
            '    Comparison.InputType3.CompareArraysSeal(NewBOM, OldBOM)
            'End If

            FormatNewBulkBOM(NewBOM, OldBOM, NewPlateBOM, BOMSheet)
            NewStickBOM.activate()
            NewBOM = Nothing
            OldBOM = Nothing
            Comparison.InputType3.ReadBOM(NewBOM, NewStickBOM)      '-----------------------Look at Stick BOM
            NewBOM = NewBOM
            Comparison.InputType3.ReadBOM(OldBOM, OldStickBOM)
            OldBOM = OldBOM
            BOMSheet = "Stick BOM"

            'If BOMType = "TANK" Then                           '-------DJL-------10-31-2023---Not needed anymore.
            Comparison.InputType3.CompareArraysTank(NewBOM, OldBOM)
            'ElseIf BOMType = "SEAL" Then
            '    Comparison.InputType3.CompareArraysSeal(NewBOM, OldBOM)
            'End If

            FormatNewBulkBOM(NewBOM, OldBOM, NewStickBOM, BOMSheet)
            NewPurchaseBOM.activate()
            NewBOM = Nothing
            OldBOM = Nothing
            Comparison.InputType3.ReadBOM(NewBOM, NewPurchaseBOM)   '-----------------------Look at Purchase BOM
            NewBOM = NewBOM
            Comparison.InputType3.ReadBOM(OldBOM, OldPurchaseBOM)
            OldBOM = OldBOM
            BOMSheet = "Other BOM"

            'If BOMType = "TANK" Then                           '-------DJL-------10-31-2023---Not needed anymore.
            Comparison.InputType3.CompareArraysTank(NewBOM, OldBOM)
            'ElseIf BOMType = "SEAL" Then
            '    Comparison.InputType3.CompareArraysSeal(NewBOM, OldBOM)
            'End If

            FormatNewBulkBOM(NewBOM, OldBOM, NewPurchaseBOM, BOMSheet)
        Else
            BOMMnu.Hide()
            MsgBox("No comparison done. One of the selected BOM file does not appear to be a Bulk BOM")
            OldBulkBOM.Activate()
            Excel.Application.ActiveWorkbook.Close(False)
        End If
    End Function

    Public Function CompareShipList() As Object
        '--------DJL-------10-31-2023--------Look at doing this in an array
        Dim Excel As Object
        Dim Test As String
        Dim OldBOMTest, OldPlateTest, OldStickTest, OldPurchaseTest, OldGratingTest, OldStickImportTest As Boolean
        Dim Workbooks As Excel.Workbooks
        Dim WrkBooksCnt As Integer
        Dim BOMMnu As ShippingList_Menu
        Dim OldShipListTest As String
        BOMMnu = Me
        PrgName = "CompareShipList"

        On Error Resume Next

        Excel = GetObject(, "Excel.Application")

        If Err.Number Then
            Information.Err.Clear()
            Excel = CreateObject("Excel.Application")
            If Err.Number Then
                MsgBox(Err.Description)
                Exit Function
            End If
        End If

        On Error GoTo Err_CompareShipList

        Workbooks = Excel.Workbooks
        WrkBooksCnt = Workbooks.Count

        If IsNothing(NewShipListSht) = True Then
            NewShipListSht = MainBOMFile.Application.ActiveWorkbook.Sheets("Shipping List")
        End If

        Test = NewShipListSht.Name
        OldShipListFile = Excel.Application.Workbooks.Open(OldShipListFileStr)
        OldShipListSht = OldShipListFile.Application.ActiveWorkbook.Sheets("Shipping List")
        Test = OldShipListSht.Name
        Comparison.InputType3.CheckOldShipList(OldShipListSht)              'Just looking at headers.
        OldShipListTest = Comparison.InputType3.HeaderType

        Select Case OldShipListTest
            Case "True"
                'NewShipListSht.Activate()
                'NewBOM = Nothing
                'OldBOM = Nothing
                'Label2.Text = "Reading New BOM Spreadsheet"
                ''-------Look at using existing Array ShippingList(18, ?)---------Instead of reading Spreadsheet again--------DJL-10-23-2023
                ''Need to have program write out to new Array when it is reading for write to spreadsheet and craeate new array.-------DJL-11-21-2023
                'ReadShipList(NewBOM, NewShipListSht)           Already done in the function WriteToExcel
                'NewBOM = NewBOM

                OldBOM = Nothing
                Label2.Text = "Reading Old Shipping List"
                ReadShipList(OldBOM, OldShipListSht)                            '--------Yes this one is ok--------DJL-------10-31-2023
                OldBOM = OldBOM
                BOMSheet = "Shipping List"

                If IsNothing(OldShipListFile) = False Then
                    OldShipListFile.Close(False)
                End If

                If IsNothing(ReadBOMFile) = False Then              '-------DJL-07-30-2025      'Close BOM information has been collected.
                    ReadBOMFile.Close(False)
                End If

                'If BOMType = "TANK" Then                           '-------DJL-------10-31-2023---Not needed anymore.
                Comparison.InputType3.CompareArraysTank(ShippingList, OldBOM)                         'Comparison.InputType3.CompareArraysTank(NewBOM, OldBOM)
                'ElseIf BOMType = "SEAL" Then
                '    MsgBox("Seals part of this program still needs to be tested.")
                '    Comparison.InputType3.CompareArraysSeal(NewBOM, OldBOM)
                'End If

                FormatNewShipList(ShippingList, OldBOM, NewShipListSht, BOMSheet)                         'FormatNewShipList(NewBOM, OldBOM, NewShipListSht, BOMSheet)
            Case "OldHeader"
                'NewShipListSht.Activate()
                'NewBOM = Nothing
                'ReadShipList(NewBOM, NewShipListSht)
                'NewBOM = NewBOM

                OldBOM = Nothing
                ReadShipListOld(OldBOM, OldShipListSht)
                OldBOM = OldBOM
                BOMSheet = "Shipping List"

                'If BOMType = "TANK" Then                           '-------DJL-------10-31-2023---Not needed anymore.
                Comparison.InputType3.CompareArraysTank(ShippingList, OldBOM)         'Comparison.InputType3.CompareArraysTank(NewBOM, OldBOM)
                'ElseIf BOMType = "SEAL" Then
                '    MsgBox("Seals part of this program still needs to be tested.")
                '    Comparison.InputType3.CompareArraysSeal(NewBOM, OldBOM)
                'End If

                FormatNewShipList(NewBOM, OldBOM, NewShipListSht, BOMSheet)
            Case Else
                BOMMnu.Hide()
                MsgBox("No comparison done. The selected Shipping List file does not appear to be a Shipping List.")
                OldShipListFile.Activate()
        End Select                              'End If

Err_CompareShipList:
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

            HandleErrSQL(PrgName, ErrNo, ErrMsg, ErrSource, PriPrg, ErrDll, DwgItem, PrgLineNo)

            If IsNothing(GenInfo.UserName) = True Then
                GenInfo.UserName = Environment.UserName
            End If

            If GenInfo.UserName = "dlong" Then
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

    Function FormatShipListRev0(ByVal NewArray As Object, ByVal OldArray As Object, ByVal ShipListSht As Object, ByRef BOMSheet As String, ByRef CntItems As Integer, ByRef StrLineNo As Integer) As Object
        Dim iC, jC, LineNo, j As Integer
        Dim FoundLast As Boolean
        PrgName = "FormatShipListRev0"

        On Error GoTo Err_FormatShipListRev0

        ShipListSht.Activate()
        If StrLineNo = 42 Then
            StrLineNo = 44
        End If

        Count = ShipListSht.Range("E4000").End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Row
        BOMSheet = "Shipping List"

        With ShipListSht
            For j = 1 To (Count - 44)
                HighlightLine(j + StrLineNo, "G", BOMSheet)
            Next j
        End With

Err_FormatShipListRev0:
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

            HandleErrSQL(PrgName, ErrNo, ErrMsg, ErrSource, PriPrg, ErrDll, DwgItem, PrgLineNo)

            If IsNothing(GenInfo.UserName) = True Then
                GenInfo.UserName = Environment.UserName
            End If

            If GenInfo.UserName = "dlong" Then
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

    Function FormatNewShipList(ByVal NewArray As Object, ByVal OldArray As Object, ByVal FileToFormat As Object, ByRef BOMSheet As String) As Object
        Dim iC, jC, LineNo As Integer
        Dim Test, foundDeleted, GetNewDesc As String
        Dim MultiLineMatl As Boolean
        PrgName = "FormatNewShipList"

        On Error GoTo Err_FormatNewShipList

        FileToFormat.Activate()
        If StrLineNo = 42 Then
            StrLineNo = 44
        End If

        Test = UBound(NewArray, 2)
        Label2.Text = "Comparing New Shipping List to Old Shipping List"
        Me.ProgressBar1.Maximum = UBound(NewArray, 2)

        For iC = 1 To UBound(NewArray, 2)
            Select Case NewArray(0, iC)
                Case "REVISED"
                    HighlightLine(iC + StrLineNo, "Y", BOMSheet)                            '2 Color number
                Case "NO CHANGE"
                    HighlightLine(iC + StrLineNo, "N", BOMSheet)                            '0 Color number
                Case "NEW"
                    If NewArray(8, iC) <> vbNullString And NewArray(1, iC) <> vbNullString Then     '-------DJL-07-25-2025       'If NewArray(1, iC) <> vbNullString Then
                        HighlightLine(iC + StrLineNo, "G", BOMSheet)                            '1 Color number
                    End If
            End Select

            Me.ProgressBar1.Value = iC
        Next iC

        Select Case Comparison.InputType3.HeaderType        '-------DJL-------07-25-2025    'Used to highlight if Rev or Weight Only changed.
            Case "True"
                For iC = 1 To UBound(NewArray, 2)
                    Test = NewArray(0, iC)

                    If NewArray(0, iC) = "REVNO" Then
                        HighlightRevNO(iC + StrLineNo, "Y", BOMSheet)
                    End If

                    If NewArray(0, iC) = "Weight" Then
                        HighlightWeight(iC + StrLineNo, "Y", BOMSheet)
                    End If
                Next iC
            Case "False"
                For iC = 1 To UBound(NewArray, 2)                       '-------Look at Revision numbers.
                    Test = NewArray(12, iC)

                    If NewArray(12, iC) = "REVNO" Then
                        HighlightRevNO(iC + StrLineNo, "Y", BOMSheet)   'Now is Yellow was Red   
                    End If

                    If NewArray(0, iC) = "Weight" Then
                        HighlightWeight(iC + StrLineNo, "Y", BOMSheet)
                    End If
                Next iC
            Case "OldHeader"
                For iC = 1 To UBound(NewArray, 2)
                    Test = NewArray(0, iC)

                    If NewArray(0, iC) = "REVNO" Then
                        HighlightRevNO(iC + StrLineNo, "Y", BOMSheet)
                    End If

                    If NewArray(0, iC) = "Weight" Then
                        HighlightWeight(iC + StrLineNo, "Y", BOMSheet)
                    End If
                Next iC
        End Select

        RowNo = ShipListSht.Range("E4000").End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Row
        LineNo = UBound(NewArray, 2) + StrLineNo

        If LineNo < RowNo Then
            LineNo = (RowNo + 2)
        End If

        Label2.Text = "Looking at Old Shipping List to see what got deleted"
        Me.ProgressBar1.Maximum = (UBound(OldArray, 2) - 1)                     'LineNo
        foundDeleted = "No"

        For iC = 1 To (UBound(OldArray, 2) - 1)                   '-----------------------------------Look for information that was deleted.
            If OldArray(0, iC) = vbNullString And OldArray(6, iC) <> vbNullString Then
                MultiLineMatl = False
                LineNo = LineNo + 1

                If foundDeleted = "Yes" Then
                    With NewShipListSht
                        .Rows(LineNo & ":" & LineNo).Select()
                        .Rows(LineNo & ":" & LineNo).Insert()
                    End With
                Else
                    With NewShipListSht
                        .Rows(LineNo & ":" & LineNo).Select()           'Insert two blank lines first.
                        .Rows(LineNo & ":" & LineNo).Insert()

                        .Rows(LineNo & ":" & LineNo).Select()
                        .Rows(LineNo & ":" & LineNo).Insert()
                    End With
                End If

                NewShipListSht.Range("C" & LineNo).Value = OldArray(1, iC)
                NewShipListSht.Range("D" & LineNo).Value = OldArray(2, iC)
                NewShipListSht.Range("E" & LineNo).Value = OldArray(3, iC)
                NewShipListSht.Range("F" & LineNo).Value = OldArray(4, iC)
                NewShipListSht.Range("G" & LineNo).Value = OldArray(5, iC)
                NewShipListSht.Range("H" & LineNo).Value = OldArray(6, iC)
                NewShipListSht.Range("I" & LineNo).Value = OldArray(7, iC)
                NewShipListSht.Range("J" & LineNo).Value = OldArray(8, iC)
                NewShipListSht.Range("K" & LineNo).Value = OldArray(9, iC)
                NewShipListSht.Range("L" & LineNo).Value = OldArray(10, iC)
                If OldArray(11, iC) <> vbLf Then
                    NewShipListSht.Range("M" & LineNo).Value = OldArray(11, iC)
                End If
                NewShipListSht.Range("N" & LineNo).Value = OldArray(12, iC)

                If OldArray(1, iC) <> Nothing Then
                    If foundDeleted = "No" Then
                        FormatLine(LineNo, MultiLineMatl)
                        FormatLine((LineNo + 1), MultiLineMatl)
                        foundDeleted = "Yes"
                    End If

                    HighlightLine(LineNo, "R", BOMSheet)
                    Me.ProgressBar1.Value = iC
                End If
            End If
        Next iC

Err_FormatNewShipList:
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

            HandleErrSQL(PrgName, ErrNo, ErrMsg, ErrSource, PriPrg, ErrDll, DwgItem, PrgLineNo)

            If IsNothing(GenInfo.UserName) = True Then
                GenInfo.UserName = Environment.UserName
            End If

            If GenInfo.UserName = "dlong" Then
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

    Function WriteToExcel(ShippingList)
        '-------Move to new function-------WritetoExcel-------DJL-11-22-2023            
        '------------------------------------------------------------------------------------------------
        '-------Creator:        Dennis J. Long
        '-------Date:           11/22/2023
        '-------Description:    Write data to Excel Spreadsheet
        '-------
        '-------Updates:        Description:
        '-------11-22-2023       Read Array and write to Excel what was collected from AutoCAD.     
        '-------                
        '-------                
        '------------------------------------------------------------------------------------------------
        '-------Create new function to write to excel for all of the code below-------DJL-------10-31-2023
        Dim i, j, k, l, PrevCnt, LineNo, DwgPos As Integer
        'Dim UpdDesc, GetX, GetY, GetRowNo, FoundDwgNo, FoundX, FoundY, FoundItem, GetNextDesc,
        Dim GetDwgNo, GetShipMk, GetDesc As String
        Dim GetPrevDesc, GetPrevShipMk As String

        On Error GoTo Err_WriteToExcel

        FileToOpen = "K:\CAD\VBA\XLTSheets\ShipListVBNet.xlt"
        MainBOMFile = ExcelApp.Application.Workbooks.Open(FileToOpen)
        FullJobNo = GenInfo3135.FullJobNo
        CustomerPO = GenInfo3135.CustomerPO
        FileSaveAS = PathBox.Text & FullJobNo & "-ShipList-R" & Me.ComboBox1.Text & ".xls"
        Workbooks = ExcelApp.Workbooks

        WorkShtName = "Shipping List"
        ShipListSht = Workbooks.Application.Worksheets(WorkShtName)
        WorkSht = Workbooks.Application.ActiveSheet
        WorkShtName = WorkSht.Name
        ExcelApp.Visible = True

        With ShipListSht
            .Range("E3").Value = FullJobNo
            .Range("L3").Value = Today
            .Range("I3").Value = Me.ComboBox1.Text
        End With

        FileToOpen = "Shipping List"
        ShipListSht.Activate()
        CntItems = (UBound(ShippingList, 2) - 1)
        ProgressBar1.Maximum = ((UBound(ShippingList, 2) - 1))                '-------DJL-07-16-2025      'ProgressBar1.Maximum = ((UBound(ShippingList, 2) - 1) + (UBound(GenInfo.SRList, 2) - 1))

        If ShipListSht.Range("A" & "43").Value = "Job No: " Then
            StrLineNo = 44

            RowNo = ShipListSht.Range("A4000").End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Row
            If RowNo = 44 Then
                RowNo = RowNo + 1
            End If
        Else
            StrLineNo = 42
            RowNo = ShipListSht.Range("A4000").End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Row
            If RowNo = 42 Then
                RowNo = RowNo + 1
            End If
        End If

        JobNoDash = FullJobNo

        For i = 1 To (UBound(ShippingList, 2) - 1)
            If ShippingList(4, i) <> "" Then                            'Get2DShipMk        '-------DJL-07-24-2025      'If ShippingList(3, i) <> "" Then

                If RowNo >= 45 Then
                    If RowNo = "45" Or RowNo = "43" Then
                        FormatLine(RowNo, FileToOpen)
                        FormatLine((RowNo + 1), FileToOpen)            '-------DJL-07-16-2025      'Only need to format oine line.
                        GoTo AddRow
                    Else
AddRow:
                        With ShipListSht
                            .Rows(RowNo & ":" & RowNo).Select()
                            .Rows((RowNo + 1) & ":" & (RowNo + 1)).Insert()         '.Rows((RowNo + 1) & ":" & (RowNo + 1)).Insert()
                        End With
                    End If

                    LineNo = RowNo
                End If
            Else
                If i = 1 Then          'If i = 1 Or i = 2 Then
                    RowNo = i + StrLineNo

                    If RowNo = "45" Or RowNo = "43" Then
                        FormatLine(RowNo, FileToOpen)
                        FormatLine((RowNo + 1), FileToOpen)
                    Else
                        With ShipListSht
                            .Rows(RowNo & ":" & RowNo).Select()
                            .Rows(RowNo & ":" & RowNo).Insert()
                        End With
                    End If

                    LineNo = RowNo
                Else
                    GoTo NextI
                End If
            End If

            With ShipListSht
                If ShippingList(4, i) <> "" Then                                        'Get2DShipMk        '-------DJL-07-24-2025      'If ShippingList(3, i) <> "" Then
                    j = (j + 1)

                    'If InStr(CustomerPO, JobNoDash) = 0 Then                            '-------DJL-06-30-2025
                    '    .Range("C" & RowNo).Value = JobNoDash & "/" & CustomerPO
                    'Else
                    '    .Range("C" & RowNo).Value = CustomerPO
                    'End If

                    .Range("C" & RowNo).Value = ShippingList(1, i)          '-------DJL-07-24-2025

                    GetDwgNo = ShippingList(2, i)                          '-------DJL-07-25-2025       'GetDwgNo = ShippingList(1, i)

                    If InStr(GetDwgNo, "-DW-") > 0 Then                          '-------DJL-06-12-2025
                        DwgPos = InStr(GetDwgNo, "-DW-")
                        GetDwgNo = Mid(GetDwgNo, (DwgPos + 4), Len(GetDwgNo))
                    End If

                    .Range("D" & RowNo).Value = GetDwgNo                          '-------DJL-06-12-2025      '.Range("D" & RowNo).Value = ShippingList(1, i)
                    .Range("E" & RowNo).Value = ShippingList(3, i)              'CurrentDwgRev      '-------07-24-2025      '.Range("E" & RowNo).Value = ShippingList(2, i)

                    '-------DJL-07-21-2025      'No need to collect for sort again found solution at Sort Array above.--------------------
                    '-------------------------------------------------DJL-07-16-2025    'Move to higher number  'CustomerPO     'So that values are the same for Shipping List as BOM
                    'ShippingListColl(1, j) = ShippingList(1, i)     'ColumnC    'CurrentDwgNo        '-------DJL-06-30-2025      'ShippingListColl(1, j) = JobNoDash & "/" & CustomerPO
                    'ShippingListColl(2, j) = ShippingList(2, i)     'ColumnD    'CurrentDwgRev
                    'ShippingListColl(3, j) = ShippingList(3, i)     'ColumnE    'Get2DShipMk
                    'ShippingListColl(4, j) = CustomerPO             'Job No / Customer info         '-------DJL-07-16-2025      'Added.

                    GetShipMk = ShippingList(4, i)                  '-------DJL-07-24-2025  'GetShipMk = ShippingList(3, i)
                    GetPrevShipMk = ShippingList(4, (i - 1))        '-------DJL-07-24-2025  'GetPrevShipMk = ShippingList(3, (i - 1))

                    .Range("F" & RowNo).Value = ShippingList(4, i)     'ColumnF    'Get2DShipMk       '-------DJL-07-24-2025        '.Range("F" & RowNo).Value = ShippingList(3, i)
                    .Range("Y" & RowNo).Value = ShippingList(4, i)                                  '-------DJL-07-24-2025          '.Range("Y" & RowNo).Value = ShippingList(3, i) 

                    'If InStr(GetShipMk, "SR") = 1 Then                                 '-------DJL-07-18-2025
                    '    UpdDesc = ShippingList(8, i)                   'GetDesc       '-------DJL-07-16-2025      'UpdDesc = ShippingList(3, i)
                    'Else
                    '    UpdDesc = Nothing
                    'End If

                    'ShippingListColl(5, j) = ShippingList(5, i)     'Piece Mark     'GetQty     '-------DJL-07-16-2025      'Added

                    .Range("G" & RowNo).Value = ShippingList(6, i)     'ColumnG        'GetQty
                    'ShippingListColl(6, j) = ShippingList(6, i)     'ColumnG        'GetQty     '-------DJL-07-16-2025      'ShippingListColl(5, j) = ShippingList(6, i)


                    '-------DJL-07-16-2025  'Will always be column 8
                    If InStr(1, ShippingList(8, i), "%%D") <> 0 Then        'GetDesc        '-------DJL-07-16-2025      'Was labeled GetInv but was wrong
                        '-------DJL-07-18-2025      'Replace with below........
                        GetDesc = (ShippingList(8, i))
                        GetDesc = GetDesc.Replace("%%D", " DEG.")              'ShippingList(8, i) = Misc.InputType2.sReplace(ShippingList(8, i), "%%D", " DEG.")

                        ShippingList(8, i) = ShippingList(8, i).Replace("%%D", " DEG.")         '-------07-18-2025
                    Else
                        GetDesc = ShippingList(8, i)
                    End If

                    'If UpdDesc <> Nothing And GetDesc <> UpdDesc Then           '-------DJL-07-18-2025      'If UpdDesc <> Nothing Then
                    '    GetDesc = GetDesc & " " & UpdDesc
                    'End If

                    '-------DJL-07-21-2025          'Not required anymore.
                    '-------DJL-07-16-2025  'Will always be column 8
                    'If InStr(1, ShippingList(8, (i + 1)), "%%D") <> 0 Then
                    '        GetNextDesc = ShippingList(8, (i + 1))
                    '        GetNextDesc = GetNextDesc.Replace("%%D", " DEG.")        'GetNextDesc = Misc.InputType2.sReplace(ShippingList(8, (i + 1)), "%%D", " DEG.")
                    '    Else
                    '        GetNextDesc = ShippingList(8, (i + 1))
                    '    End If


                    .Range("H" & RowNo).Value = GetDesc     '-------DJL-07-21-2025      '.Range("H" & RowNo).Value = ShippingList(8, i)

                    If ShippingList(16, i) <> "" Then
                        .Range("Z" & RowNo).Value = ShippingList(16, i)     '-------DJL-07-21-2025   'Already exist --> GetPrevDesc   '.Range("Z" & RowNo).Value = GetDesc 
                    Else
                        .Range("Z" & RowNo).Value = GetDesc     '-------DJL-07-21-2025      '.Range("Z" & RowNo).Value = ShippingList(8, i)
                    End If

                    'ShippingListColl(8, j) = ShippingList(8, i)       '-------DJL-07-16-2025      'ShippingListColl(6, j) = ShippingList(8, i)      'ShippingListColl(6, j) = GetDesc

                    '--------------------------------------------------------------------------------
                    '--------------------------------------------------------------------------------
                    If InStr(GetShipMk, "SR") = 1 Then                            'If InStr(.Range("F" & RowNo).Value, "SR") = 1 Then       '-------DJL-------11-28-2023
                        If InStr(GetDesc, "PLATE") = 0 Then                         'If InStr(.Range("H" & RowNo).Value, "PLATE") = 0 Then
                            GoTo SRPlateNotFound
                        End If

                        '-------DJL-07-18-2025      'Going to do this at time of Data Collection.
                        'If GetDesc = "SHELL PLATE" Then                                           'GetDesc = .Range("H" & RowNo).Value
                        '    GetPrevShpMkDesc = GetDesc
                        'End If

                        '-------DJL-07-18-2025      'Not required anymore.
                        'GetNextDesc = ShippingList(8, (i + 1))              '-------DJL-07-18-2025       'GetNextDesc = .Range("H" & (RowNo + 1)).Value

                        If InStr(ShippingList(8, (i + 1)), "PL ") > 0 Then               '-------DJL-07-21-2025      'If InStr(GetNextDesc, "PL ") > 0 Then
                            'GetDesc = GetDesc & " " & GetNextDesc          '-------DJL-07-18-2025      'Do not do this anymore.
                            '.Range("H" & RowNo).Value = GetDesc                     'GetDesc & " " & GetNextDesc
                            'ShippingListColl(6, j) = GetDesc                        'ShippingListColl(6, j) = GetDesc       'GetDesc & " " & GetNextDesc    
                            .Range("A" & RowNo & ":" & "X" & RowNo).Delete()        '-------DJL-07-18-2024      '.Range("A" & (RowNo + 1) & ":" & "X" & (RowNo + 1)).Delete()
                            j = (j - 1)                                     '-------DJL-07-21-2025
                            GoTo NextI                                      '-------DJL-07-18-2025
                            'LineNo = (LineNo - 1)                          '-------DJL-07-18-2024
                        Else
                            .Range("H" & RowNo).Value = GetDesc
                            'ShippingListColl(6, j) = GetDesc                        'ShippingListColl(6, j) = GetDesc
                        End If

                        'If RowNo > LineNo Then                         '-------DJL-07-18-2025
                        '    GoTo LastRowFound
                        'End If
                    Else
SRPlateNotFound:
                        If InStr(GetShipMk, "SR") = 1 And InStr(GetDesc, "PLATE") > 0 Then      '-------Looking for way to delete Insert Plate.

                            Select Case 0
                                Case Is < InStr(GetShipMk, "SR1")
                                    GetPrevDesc = GetDesc     '-------DJL-07-21-2025
                                Case Is < InStr(GetShipMk, "SR2")
                                    GetPrevDesc = GetDesc
                                Case Is < InStr(GetShipMk, "SR3")
                                    GetPrevDesc = GetDesc
                                Case Is < InStr(GetShipMk, "SR4")
                                    GetPrevDesc = GetDesc
                                Case Is < InStr(GetShipMk, "SR5")
                                    GetPrevDesc = GetDesc
                                Case Is < InStr(GetShipMk, "SR6")
                                    GetPrevDesc = GetDesc
                                Case Is < InStr(GetShipMk, "SR7")
                                    GetPrevDesc = GetDesc
                                Case Is < InStr(GetShipMk, "SR8")
                                    GetPrevDesc = GetDesc
                                Case Is < InStr(GetShipMk, "SR9")
                                    GetPrevDesc = GetDesc
                                Case Is < InStr(GetShipMk, "SR10")
                                    GetPrevDesc = GetDesc
                                Case Else
                                    If RowNo > LineNo Then                          'even thou LineNo is now at 200 the orginal number was 224 so program must be stoped.
                                        GoTo LastRowFound
                                    End If
                                    .Range("A" & (RowNo - 1) & ":X" & (RowNo - 1)).Delete()     '-------DJL-07-21-2025     '.Range("A" & RowNo & ":" & "X" & RowNo).Delete()

                                    RowNo = (RowNo - 1)
                                    LineNo = (LineNo - 1)
                            End Select
                        Else                            'Also need to look at plate that is not part of SR1 and has column F as Blank.
                            If IsNothing(GetShipMk) = True And InStr(GetPrevShipMk, "SR") = 0 Then       '-------Make sure that previous Ship Mark is not "SR"        'If IsNothing(.Range("F" & RowNo).Value) = True And InStr(.Range("F" & (RowNo - 1)).Value, "SR") = 0 Then
                                If RowNo > LineNo Then
                                    GoTo LastRowFound
                                End If
                                .Range("A" & RowNo & ":" & "X" & RowNo).Delete()
                                RowNo = (RowNo - 1)
                                LineNo = (LineNo - 1)
                            End If
                        End If
                    End If

LastRowFound:
                    .Range("K" & RowNo).Value = ShippingList(9, i)                 'ColumnK        'GetInv1     '.Range("V" & RowNo).Value = ShippingList(9, i)
                    .Range("L" & RowNo).Value = ShippingList(10, i)               'ColumnL        'GetInv2      '.Range("W" & RowNo).Value = ShippingList(10, i)

                    'ShippingListColl(9, j) = ShippingList(9, i)                 'ColumnK        'GetInv1
                    'ShippingListColl(10, j) = ShippingList(10, i)               'ColumnL        'GetInv2

                    '-------DJL-07-21-2025      'Not required anymore.-------------------------------------------------------------------
                    'If ShippingList(11, i) = vbNullString Or Mid(ShippingList(11, i), 1, 1) = " " Then      'Column M       'GertMat
                    '    With ShipListSht
                    '        With .Range("M" & RowNo)
                    '            .HorizontalAlignment = XlHAlign.xlHAlignCenter
                    '            .VerticalAlignment = XlVAlign.xlVAlignCenter
                    '            .Font.Name = "Arial"
                    '            .Font.Size = 7
                    '            .Value = ShippingList(12, i) & Chr(10) & ShippingList(13, i)
                    '            'ShippingListColl(11, j) = ShippingList(12, i) & Chr(10) & ShippingList(13, i)               'ColumnM
                    '        End With
                    '    End With
                    'Else
                    .Range("M" & RowNo).Value = ShippingList(11, i)         'Column M       'GertMat
                    'ShippingListColl(11, j) = ShippingList(11, i)           'Column M       'GertMat
                    'End If

                    'ShippingListColl(14, j) = ShippingList(14, i)               'ColumnN        'GetPds     '-------DJL-07-17-2025      'ShippingListColl(12, j) = ShippingList(14, i)
                    .Range("N" & RowNo).Value = ShippingList(14, i)               'ColumnN        'GetPds
                    'ShippingListColl(16, j) = ShippingList(16, i)
                    'ReDim Preserve ShippingListColl(18, UBound(ShippingListColl, 2) + 1)
                    .Range("AA" & RowNo).Value = RowNo
                    RowNo = (RowNo + 1)
                End If
            End With
            ProgressBar1.Value = i
            Me.TxtBoxCntDown.Text = (CntItems - i)           '-------DJL-07-25-2025
NextI:
        Next i

        With ShipListSht
            If RowNo = 0 Then
                MsgBox("No Shipping List Items found on drawings.")
                GoTo EndPrg
            End If

            'With .Range("A" & StrLineNo & ":Z" & (RowNo + 1))          'Not Required anymore.
            '    .Sort(Key1:= .Range("V5"), Order1:=XlSortOrder.xlAscending, Key2:= .Range("W5"), Order2:=XlSortOrder.xlAscending, Key3:= .Range("X5"), Order3:=XlSortOrder.xlDescending, Header:=XlYesNoGuess.xlYes, OrderCustom:=1, MatchCase:=False, Orientation:=XlSortOrientation.xlSortColumns)            'This is correct Top down for column Y Decending        '-------DJL-------11-27-2023
            'End With
        End With

        '-------DJL-07-21-2025-------------Below is not required anymore.------------------------------------------------------------------
        '-------DJL-------11-27-2023-------New process create sort.
        'k = 1
        'ProgressBar1.Maximum = RowNo + 1
        'RowNo = ShipListSht.Range("E4000").End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Row
        'ReDim ShippingList(18, 1)
        '        With ShipListSht

        '            For i = StrLineNo To RowNo
        '                GetDwgNo = .Range("D" & i).Value                            '-------DJL-07-18-2025      'GetDwgNo = .Range("V" & i).Value
        '                GetX = .Range("W" & i).Value
        '                GetY = .Range("X" & i).Value
        '                GetRowNo = .Range("Z" & i).Value

        '                If i = 44 And GetDwgNo = Nothing Then
        '                    GoTo Nexti2
        '                Else
        '                    If i = 45 And GetDwgNo = Nothing Then
        '                        GoTo Nexti2
        '                    End If
        '                End If

        '                If PrevCnt = 0 Then
        '                    PrevCnt = 1
        '                End If

        '                For j = PrevCnt To (UBound(ShippingListColl, 2) - 1)
        '                    FoundDwgNo = ShippingListColl(1, j)             'ColumnV       'FoundDwgNo = ShippingListColl(15, j)  
        '                    FoundX = ShippingListColl(16, j)                 'ColumnW
        '                    FoundY = ShippingListColl(17, j)                 'ColumnX
        '                    FoundItem = ShippingListColl(18, j)


        '                    If FoundItem = "Found" Then
        '                        If ShippingListColl(18, PrevCnt) = "Found" And PrevCnt = (j - 1) Then
        '                            If ShippingListColl(18, PrevCnt) = "Found" Then
        '                                PrevCnt = j
        '                            End If
        '                        Else
        '                            For l = PrevCnt To j
        '                                If ShippingListColl(18, l) = "Found" And ShippingListColl(18, PrevCnt) = "Found" Then
        '                                    PrevCnt = l + 1
        '                                End If
        '                            Next l
        '                        End If

        '                        GoTo Nextj
        '                    Else
        '                        If PrevCnt > j And ShippingListColl(18, PrevCnt) = "Found" Then
        '                            PrevCnt = j
        '                        End If
        '                    End If

        '                    If GetDwgNo = FoundDwgNo And GetX = FoundX Then
        '                        If GetY = FoundY Then
        '                            ShippingList(1, k) = ShippingListColl(1, j)                 'ColumnC    'Job Number     'ShippingList(1, j) = ShippingListColl(1, (GetRowNo - StrLineNo))      'ShippingList(1, j) = ShippingListColl(1, i)
        '                            ShippingList(2, k) = ShippingListColl(2, j)                 'ColumnD    'Dwg No.
        '                            ShippingList(3, k) = ShippingListColl(3, j)                 'ColumnE    'Rev No.
        '                            ShippingList(4, k) = ShippingListColl(4, j)                 'ColumnF    'Ship No.
        '                            ShippingList(5, k) = ShippingListColl(5, j)                 'ColumnG    'Qty
        '                            ShippingList(6, k) = ShippingListColl(6, j)                  'ColumnH   'Desc           'ShippingList(6, k) = ShippingListColl(6, j) 

        '                            ShippingList(9, k) = ShippingListColl(9, j)                 'ColumnK    'Inv No.
        '                            ShippingList(10, k) = ShippingListColl(10, j)               'ColumnL    'Std No.
        '                            ShippingList(11, k) = ShippingListColl(11, j)               'ColumnM    'Material
        '                            ShippingList(12, k) = ShippingListColl(12, j)               'ColumnN    'Weight

        '                            ShippingList(15, k) = ShippingListColl(15, j)               'ColumnV    'Dwg No
        '                            ShippingList(16, k) = ShippingListColl(16, j)               'ColumnW    'X
        '                            ShippingList(17, k) = ShippingListColl(17, j)               'ColumnX    'Y
        '                            ShippingList(18, k) = k                     '"Found"    'We know it was found just need row number.
        '                            ShippingListColl(18, j) = "Found"
        '                            ReDim Preserve ShippingListColl(18, UBound(ShippingListColl, 2))
        '                            ReDim Preserve ShippingList(18, UBound(ShippingList, 2) + 1)
        '                            k = (k + 1)

        '                            If PrevCnt + 1 = j And FoundItem <> "Found" Then
        '                                If PrevCnt < 3 And ShippingListColl(18, (j - 1)) <> "Found" Then
        '                                    'Do nothing             'PrevCnt = j
        '                                Else
        '                                    If ShippingListColl(18, (j - 1)) = "Found" And ShippingListColl(18, j) = "Found" Then
        '                                        If ShippingListColl(18, PrevCnt) = "Found" Then
        '                                            PrevCnt = j
        '                                        End If
        '                                    End If
        '                                End If
        '                            Else
        '                                If PrevCnt = j And ShippingListColl(18, j) = "Found" Then
        '                                    If ShippingListColl(18, PrevCnt) = "Found" Then
        '                                        PrevCnt = j + 1
        '                                    End If
        '                                End If
        '                            End If

        '                            GoTo Nexti2
        '                        Else
        '                            GoTo Nextj
        '                        End If
        '                    Else
        '                        GoTo Nextj
        '                    End If

        'Nextj:
        '                Next j

        'Nexti2:
        '                ProgressBar1.Value = i
        '            Next i

        '            With .Range("V:Z")                          'With .Range("V:X")
        '                .Delete(Shift:=XlDeleteShiftDirection.xlShiftUp)
        '            End With
        '        End With

        GenInfo3135.StrLineNo = StrLineNo                           '-------DJL-07-21-2025      'Used in the next part of the program.
EndPrg:

        GenInfo3135.ShippingList = ShippingList
Err_WriteToExcel:
        ErrNo = Err.Number

        If ErrNo <> 0 Then
            PriPrg = "ShipListReadBOMAutoCAD"
            ErrMsg = Err.Description
            ErrSource = Err.Source
            ErrDll = Err.LastDllError
            ErrLastLineX = Err.Erl
            ErrException = Err.GetException

            If ErrNo = "1004" And InStr(ErrMsg, "Select method of Range class") > 0 Then    '-------DJL-07-18-2025      'Basically the program timed out and on resume it is working again.
                Resume
            End If

            If System.Environment.UserName = "dlong" Or System.Environment.UserName = "aclemmer" Then
                ErrMsg = Err.Description
                Stop
                Resume
            End If
        End If

    End Function

    Function ReadShipList(ByRef ShipListArray As Object, ByVal SheetToUse As Object) As Object
        '--------------------------------------Used to read contents of Shipping List
        Dim iA, jA, LineNo As Integer
        Dim FoundLast As Boolean
        'Dim Test1, Test2 As String

        On Error GoTo Err_ReadShipList

        FoundLast = False
        If SheetToUse.Range("A" & "29").Value = "Job No: " Then
            LineNo = 30
            StrLineNo = 30
        ElseIf SheetToUse.Range("A" & "31").Value = "Job No: " Then
            LineNo = 32
            StrLineNo = 32
        ElseIf SheetToUse.Range("A" & "43").Value = "Job No: " Then
            LineNo = 44
            StrLineNo = 44
        End If

        LineNo = SheetToUse.Range("E4000").End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Row

        ReDim ShipListArray(12, LineNo - StrLineNo)
        Me.ProgressBar1.Maximum = LineNo

        For iA = StrLineNo + 1 To LineNo
            'For jA = 1 To 11
            '-------DJL-------Redoing Old Shipping List-------------------------------------Start at 66---------65 below is equal to "A"
            'Test1 = SheetToUse.Range(Chr(jA + 66) & iA).Value
            'Test2 = SheetToUse.Range(Chr(jA + 68) & iA).Value

            'Select Case jA
            '    Case 1 To 6                                                 'Case 1 To 6
            '        If SheetToUse.Range(Chr(jA + 66) & iA).Interior.ColorIndex = 3 Then
            '        Else
            '            ShipListArray(jA, iA - StrLineNo) = SheetToUse.Range(Chr(jA + 66) & iA).Value
            '        End If
            '    Case 7 To 11                                                'Case 7 To 10
            '        If SheetToUse.Range(Chr(jA + 68) & iA).Interior.ColorIndex = 3 Then
            '        Else
            '            ShipListArray(jA, iA - StrLineNo) = SheetToUse.Range(Chr(jA + 68) & iA).Value
            '        End If
            'End Select

            If SheetToUse.Range("H" & iA).Interior.ColorIndex <> 3 And SheetToUse.Range("H" & iA).Value <> "DESCRIPTION" Then
                'ShipListArray(1, iA - StrLineNo) = SheetToUse.Range("C" & iA).Value

                ShipListArray(1, iA - StrLineNo) = SheetToUse.Range("C" & iA).Value
                ShipListArray(2, iA - StrLineNo) = SheetToUse.Range("D" & iA).Value
                ShipListArray(3, iA - StrLineNo) = SheetToUse.Range("E" & iA).Value
                ShipListArray(4, iA - StrLineNo) = SheetToUse.Range("F" & iA).Value
                ShipListArray(5, iA - StrLineNo) = SheetToUse.Range("G" & iA).Value
                ShipListArray(6, iA - StrLineNo) = SheetToUse.Range("H" & iA).Value
                ShipListArray(7, iA - StrLineNo) = SheetToUse.Range("I" & iA).Value
                ShipListArray(8, iA - StrLineNo) = SheetToUse.Range("J" & iA).Value
                ShipListArray(9, iA - StrLineNo) = SheetToUse.Range("K" & iA).Value
                ShipListArray(10, iA - StrLineNo) = SheetToUse.Range("L" & iA).Value
                ShipListArray(11, iA - StrLineNo) = SheetToUse.Range("M" & iA).Value
                ShipListArray(12, iA - StrLineNo) = SheetToUse.Range("N" & iA).Value           'Not required -------DJL-------11-28-2023.
                'ShipListArray(14, iA - StrLineNo) = SheetToUse.Range("O" & iA).Value
                'ShipListArray(15, iA - StrLineNo) = SheetToUse.Range("P" & iA).Value
                'ShipListArray(16, iA - StrLineNo) = SheetToUse.Range("Q" & iA).Value
                'ShipListArray(17, iA - StrLineNo) = SheetToUse.Range("R" & iA).Value
                'ShipListArray(18, iA - StrLineNo) = SheetToUse.Range("S" & iA).Value
            End If

            'Next jA

            Me.TxtBoxCntDown.Text = (LineNo - iA)           '-------DJL-07-28-2025       'Me.TxtBoxCntDown.Text = ((LineNo - StrLineNo) - iA) 
            Me.Label2.Text = "Reading Old Shipping List for Compare process........Please Wait."
            Me.Refresh()

            Me.ProgressBar1.Value = iA
        Next iA

Err_ReadShipList:
        ErrNo = Err.Number

        If ErrNo <> 0 Then
            ErrMsg = Err.Description

            MsgBox(ErrMsg)
            Stop
            Resume
        End If

    End Function

    Function ReadShipListOld(ByRef ShipListArray As Object, ByVal SheetToUse As Object) As Object
        '--------------------------------------Used to read contents of Shipping List
        Dim iA, jA, LineNo As Integer
        Dim FoundLast As Boolean
        Dim Test1, Test2 As String

        FoundLast = False
        If SheetToUse.Range("A" & "29").Value = "Job No: " Then
            LineNo = 30
            StrLineNo = 30
        ElseIf SheetToUse.Range("A" & "31").Value = "Job No: " Then
            LineNo = 32
            StrLineNo = 32
        ElseIf SheetToUse.Range("A" & "43").Value = "Job No: " Then
            LineNo = 44
            StrLineNo = 44
        End If

        'Do Until FoundLast = True
        '    LineNo = LineNo + 1
        '    If SheetToUse.Range("B" & LineNo).Value = "" Or SheetToUse.Range("B" & LineNo).Interior.ColorIndex = 3 Then
        '        LineNo = LineNo - 1
        '        FoundLast = True
        '    End If
        'Loop

        LineNo = SheetToUse.Range("E4000").End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Row

        ReDim ShipListArray(12, LineNo - StrLineNo)
        Me.ProgressBar1.Maximum = LineNo

        For iA = StrLineNo + 1 To LineNo
            For jA = 1 To 11
                Test1 = SheetToUse.Range(Chr(jA + 65) & iA).Value
                Test2 = SheetToUse.Range(Chr(jA + 67) & iA).Value
                Select Case jA
                    Case 1
                        If SheetToUse.Range(Chr(1 + 65) & iA).Interior.ColorIndex = 3 Then
                        Else
                            ShipListArray(jA, iA - StrLineNo) = SheetToUse.Range(Chr(jA + 65) & iA).Value
                        End If
                    Case 2 To 6
                        If SheetToUse.Range(Chr(jA + 65) & iA).Interior.ColorIndex = 3 Then
                        Else
                            ShipListArray(jA, iA - StrLineNo) = SheetToUse.Range(Chr(jA + 65) & iA).Value
                        End If
                    Case 7 To 11
                        If SheetToUse.Range(Chr(jA + 67) & iA).Interior.ColorIndex = 3 Then
                        Else
                            ShipListArray(jA, iA - StrLineNo) = SheetToUse.Range(Chr(jA + 67) & iA).Value
                        End If
                End Select
            Next jA
            Me.ProgressBar1.Value = iA
        Next iA

    End Function

    Public Function FormatNewBulkBOM(ByRef NewArray As Object, ByRef OldArray As Object, ByRef FileToFormat As Object, ByRef BOMSheet As String) As Object
        Dim iC As Object
        Dim jC As Short, LineNo As Short
        Dim MultiLineMatl As Boolean
        Dim Test, FileToOpen As String

        FileToFormat.Activate()

        Test = UBound(NewArray, 2)

        For iC = 1 To UBound(NewArray, 2)
            Select Case NewArray(13, iC)
                Case "REVISED"
                    HighlightLine(iC + 4, "Y", BOMSheet)
                Case "NEW"
                    HighlightLine(iC + 4, "G", BOMSheet)
                Case "NO CHANGE"
                    HighlightLine(iC + 4, "N", BOMSheet)
            End Select
        Next iC

        LineNo = UBound(NewArray, 2) + 4

        For iC = 1 To UBound(OldArray, 2)
            If OldArray(13, iC) = vbNullString Then
                MultiLineMatl = False
                LineNo = LineNo + 1
                For jC = 1 To UBound(OldArray, 1) - 1
                    Select Case BOMSheet
                        Case "Bulk BOM"
                            With NewBulkBOM
                                .Range(Chr(64 + jC) & LineNo).Value = OldArray(jC, iC)
                                Select Case jC
                                    Case 9
                                        If InStr(1, OldArray(jC, iC), Chr(10)) <> 0 Then
                                            MultiLineMatl = True
                                        End If
                                End Select
                            End With
                    End Select
                Next jC

                FileToOpen = BOMSheet
                FormatLine(LineNo, FileToOpen, MultiLineMatl)
                HighlightLine(LineNo, "R", BOMSheet)
            End If
        Next iC

    End Function

    'Private Sub btnNewDir_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnNewDir.Click
    '    Dir2.Visible = True
    '    Drive2.Visible = True
    'End Sub

    'Private Sub Drive2_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Drive2.SelectedIndexChanged
    '    Dir2.Path = Drive2.Drive
    '    GroupBox5.Text = "Current Directory = " & Dir2.Path
    'End Sub

    '    Private Sub Dir2_Change(ByVal sender As Object, ByVal e As System.EventArgs) Handles Dir2.Change
    '        Dim i As Short
    '        Dim k As Integer
    '        Dim FindRev, FindRev1, FindDwg, PrevRev As String
    '        Dim FindRev2, FoundRev As Integer
    '        Dim RevPos, ExtPos, MxPos As Integer

    '        PrgName = "UserForm_Initialize"
    '        FindRev = Nothing

    '        On Error GoTo Err_Dir1_Change

    '        Me.File3.Pattern = "*.xls"
    '        File3.Path = Dir2.Path
    '        File2.Path = Dir2.Path

    '        If Me.File2.Items.Count > 0 Then
    '            FoundRev = 0

    '            For i = 0 To (Me.File2.Items.Count - 1)
    '                FindRev = File2.Items.Item(i)
    '                RevPos = InStr(1, FindRev, "-R")
    '                ExtPos = InStr(1, FindRev, ".xls")
    '                ExtPos = (ExtPos - 4)
    '                FindRev1 = Mid(FindRev, (RevPos + 2), ExtPos)
    '                FindRev1 = FindRev1.Replace(".xls", "")
    '                FindRev1 = FindRev1.Replace(".XLS", "")

    '                '    For Each match As Match In matches
    '                '        FindRev2 = match.Value
    '                '    Next


    '                'If FoundRev < FindRev2 Then
    '                '    FoundRev = FoundRev & FindRev2                          'FoundRev = FindRev2
    '                'End If

    '                Dim pattern As String = "^[0-99]"                    'Also Look at \Nd  Number with dec.
    '                Dim matches As MatchCollection = Regex.Matches(FindRev1, pattern)
    '                If Regex.IsMatch(FindRev1, pattern) Then
    '                    'For k = 1 To Len(FindRev1)          'For Each match As Match In matches     'Is not stepping through each number example: 13
    '                    If FindRev1 > FoundRev Then
    '                        FoundRev = FindRev1
    '                    End If
    '                    'Next k
    '                End If
    '            Next i

    '            Dim index As Integer
    '            index = ComboBox1.FindString(FoundRev + 1)
    '            ComboBox1.SelectedIndex = index
    '            File2.SelectedItem = FindRev
    '        End If

    '        If Me.File2.Items.Count > 0 Then
    '            If ShipListBox.Items.Count > 0 Then         'If DwgList.Items.Count > 0 Then
    '                ShipListBox.Items.Clear()               'DwgList.Items.Clear()
    '            End If

    '            For i = 0 To (Me.File2.Items.Count - 1)
    '                ExtPos = 0
    '                MxPos1 = 0         'DJL-------06-02-2025
    '                MxPos2 = 0
    '                MxPos3 = 0
    '                MxPos4 = 0
    '                ChPos1 = 0         'DJL-------06-02-2025
    '                ChPos2 = 0
    '                ChPos3 = 0
    '                ChPos4 = 0
    '                CaPos1 = 0         'DJL-------06-02-2025
    '                CaPos2 = 0
    '                CaPos3 = 0
    '                CaPos4 = 0

    '                FindDwg = File2.Items.Item(i)
    '                ExtPos = InStr(1, FindDwg, ".xls")                          'ExtPos = InStr(1, FindDwg, ".dwg")
    '                MxPos1 = InStr(1, FindDwg, "_MX")         'DJL-------06-02-2025
    '                MxPos2 = InStr(1, FindDwg, "_Mx")
    '                MxPos3 = InStr(1, FindDwg, "-MX")
    '                MxPos4 = InStr(1, FindDwg, "-Mx")
    '                ChPos1 = InStr(1, FindDwg, "_CH")         'DJL-------06-02-2025
    '                ChPos2 = InStr(1, FindDwg, "_Ch")
    '                ChPos3 = InStr(1, FindDwg, "-CH")
    '                ChPos4 = InStr(1, FindDwg, "-Ch")
    '                CaPos1 = InStr(1, FindDwg, "_CA")         'DJL-------06-02-2025
    '                CaPos2 = InStr(1, FindDwg, "_Ca")
    '                CaPos3 = InStr(1, FindDwg, "-CA")
    '                CaPos4 = InStr(1, FindDwg, "-Ca")

    '                If ExtPos > 0 Then
    '                    If MxPos1 = 0 And MxPos2 = 0 And MxPos3 = 0 And MxPos4 = 0 Then
    '                        If ChPos1 = 0 And ChPos2 = 0 And ChPos3 = 0 And ChPos4 = 0 Then         'DJL-------06-02-2025
    '                            If CaPos1 = 0 And CaPos2 = 0 And CaPos3 = 0 And CaPos4 = 0 Then         'DJL-------06-02-2025
    '                                ShipListBox.Items.Add(File2.Items.Item(i))          'DwgList.Items.Add(File2.Items.Item(i))
    '                            End If
    '                        End If

    '                    End If
    '                End If
    '            Next i
    '        End If

    '        GroupBox5.Text = "Current Directory = " & Dir2.Path

    'Err_Dir1_Change:
    '        ErrNo = Err.Number

    '        If ErrNo <> 0 Then
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

    '            HandleErrSQL(PrgName, ErrNo, ErrMsg, ErrSource, PriPrg, ErrDll, DwgItem, PrgLineNo)

    '            If GenInfo.UserName = "dlong" Then
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
    '    End Sub

    Public Function FindStdsBOM() As Object
        Dim NewBulkBOM As Worksheet
        Dim xlShiftUp As Object, Columns As Object, xlTopToBottom As Object, xlGuess As Object, xlDescending As Object
        Dim xlToRight As Object
        Dim xlAscending As Object, xlCenter As Object, Selection As Object, Range As Object, Worksheets As Object
        Dim VarSelArray As Object, Excel As Object, BlockData(1) As Object, DwgItem As Object, BomItem As Object
        Dim TempAttributes As Object, Temparray As Object, gvntSDIvar As Object, InsertionPT As Object
        Dim Title, Msg, Style, Response As Object, ShipListItem As Object
        Dim TagInsPnt1 As Object, TagInsPnt2 As Object, xlMinimized As Object, ExcelWorkbook As Object
        Dim BlockSel As AutoCAD.AcadSelectionSet
        Dim AMWdisabled As Boolean, AcadOpen As Boolean, MultiLineMatl As Boolean
        Dim BatchYN As Short, GroupCode(1) As Short
        Dim AcadPref As AutoCAD.AcadPreferencesSystem
        Dim CompareX1 As Double, CompareX2 As Double, Dimscale As Double
        Dim i, j, k, i1, x, Count, Shopwatch, CntWorkbook, CntWrkShts, CountVal, CutCount, Testi As Integer
        Dim SeeNotePos, ChkPos, SeeDwgPos, ExceptPos, CntExcept, RevPos, ExtPos, jA, FoundPart As Integer
        Dim CntOldStdItems, CntNewBulkBOM, Startj, l, GetNewjA, StartjA, Totalj, NewCount, TotaljA As Integer
        Dim Note2Pos, NewSeeNotePos, NewChkPos, NewSeeDwgPos, NewNote2Pos, NewBOMPos, StartCnt As Integer
        Dim FullJobNo, CurrentDwgRev, CurrentDwgNo, FilePath, CurrentDWG, PrgName, WorkBookName As String
        Dim WrkBookName, WorkShtName, NextSht, BomWrkShtNam, FileSaveAS, OldFileNam As String
        Dim LookForStd, LookForShipMk, FoundLast, LineNo, LineNo2, LineNo3, LineNo4, FoundRev As String
        Dim AdeptPrg, AdeptStd, Dwg, ExistSTD, Test, SearchException, FindRev, FindRev1, FindRev2 As String
        Dim ShpM, ShpM2, ShpQ, ShpQ2, Desc, Desc2, Ref1, Ref2, Matl, Matl2, Matl3, Wght, NewItem As String
        Dim OldDwg, OldRev, OldShpMk, OldPcMk, OldQty, OldDesc, OldDesc2, OldInv, OldStd, OldMatl As String
        Dim OldWht, OldReq, OldProd As String
        Dim NewDwg, NewRev, NewShpMk, NewPcMk, NewQty, NewDesc, NewDesc2, NewInv, NewStd, NewMatl, NewWht As String
        Dim NewReq, NewProd As String
        Dim FoundItem, SearchSeeNote, SearchNote, SearchNote2, SearchDwg, CompDesc, FirstTimeThru As String
        Dim OldRef, MxPcMk As String
        Dim NTest, NTest1, NTest2, NTest3, Ntest4, NTest5, Ntest6, NTest7, NTest8, NTest9, NTest10 As String
        Dim NTest11, NTest12, DescFixed As String
        Dim OTest, OTest1, OTest2, OTest3, Otest4, OTest5, Otest6, OTest7, OTest8, OTest9, OTest10 As String
        Dim OTest11, OTest12, pattern As String
        Dim Workbooks As Excel.Workbooks
        Dim WorkSht As Worksheet, BOMWrkSht As Worksheet, ShpCutWrkSht As Worksheet, StdsWrkSht As Worksheet
        Dim BOMSTDsSht As Worksheet
        Dim StdItemsWrkSht As Worksheet
        Dim FileToOpen As String
        Dim ExcelApp As Object
        Dim BOMMnu As ShippingList_Menu
        BOMMnu = Me

        PrgName = "FindStdsBOM"
        NewDesc = Nothing
        NewInv = Nothing
        NewDesc2 = Nothing
        NewQty = Nothing
        NewPcMk = Nothing
        CompDesc = ""
        FoundItem = Nothing
        OldQty = Nothing

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

        On Error GoTo Err_FindStdsBOM

        BOMMnu.Label2.Text = "Copy Standards to Shipping List........Please Wait"
        BOMMnu.Refresh()
        Workbooks = ExcelApp.Workbooks
        WorkShtName = "STDs BOM"
        BOMSTDsSht = Workbooks.Application.Worksheets(WorkShtName)
        BOMSTDsSht.Activate()

        '---------------------------------------------Create new sheet for Standards found.
        WorkShtName = "STD Items"
        StdItemsWrkSht = Workbooks.Application.Worksheets(WorkShtName)

        LineNo2 = BOMSTDsSht.Range("E4000").End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Row

        LineNo4 = StdItemsWrkSht.Range("E4000").End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Row

        Count = 0
        CountVal = 0
        i = 1
        SearchSeeNote = "(SEE NOTE"
        SearchNote = " ("
        SearchDwg = "(SEE DWG"
        SearchNote2 = "NOTE"

RptGetData: Testi = (i)
        If Testi > (LineNo2 + CountVal) Then
            BOMMnu.ProgressBar1.Maximum = Testi
        Else
            BOMMnu.ProgressBar1.Maximum = (LineNo2 + CountVal)                 'At Start of loop LineNo2 + CountVal is greater
        End If

        For i = Testi To (LineNo2 + CountVal)
            RowNo = i + 4

            NewBulkBOM = MainBOMFile.Application.ActiveWorkbook.Sheets("Bulk BOM")
            FindStdsBOM = MainBOMFile.Application.ActiveWorkbook.Sheets("STDs BOM")         'New Line look for Stds only on sheet STDs BOM
            OldStdItems = MainBOMFile.Application.ActiveWorkbook.Sheets("STD Items")

            NewBOM = Nothing
            OldBOM = Nothing
            FindSTD = Nothing
            '----Below is not required Array's already exist for this, but could not get Array's to share with this Module....
            Comparison.InputType3.ReadBulkBOM(NewBOM, NewBulkBOM)   'Reads spreadsheet "Bulk BOM" and creates Array--Look at Bulk BOM
            Comparison.InputType3.ReadFindSTDs(FindSTD, FindStdsBOM)  'Reads Spreadshhet "STDs BOM" and creates Array--Look at STDs BOM
            Comparison.InputType3.ReadBOM(OldBOM, OldStdItems)      'Reads spreadsheet "STD Items" and creates Array--Look at Std Items
            '---------------------------------------Create Complete list for MX BOM Std's
            OldDesc = ""
            OldInv = ""
            OldStdDwg = ""
            OTest = ""
            OTest2 = ""
            NTest = ""
            NTest2 = ""
            Dim CntFindStds As String
            '-------------------------------Replace below
            CntNewBulkBOM = UBound(NewBOM, 2)
            CntOldStdItems = UBound(OldBOM, 2)
            CntFindStds = UBound(FindSTD, 2)

GetDataNew:
            If GetNewjA < 1 Then
                GetNewjA = 1
            Else
                GetNewjA = (GetNewjA + 1)
            End If

            If GetNewjA > BOMMnu.ProgressBar1.Maximum Then
                BOMMnu.ProgressBar1.Maximum = (GetNewjA + CountVal + 4)
                BOMMnu.ProgressBar1.Value = GetNewjA
            Else
                BOMMnu.ProgressBar1.Value = GetNewjA
            End If

            DescFixed = "No"                                'Looking for new part set DescFixed to No
            TotaljA = UBound(FindSTD, 2)

            If GetNewjA > TotaljA Then
                GoTo FoundAllParts                          'All Parts have been found now go to Compare part of program.
            End If

            For jA = GetNewjA To UBound(FindSTD, 2)
                NewStdDwg = FindSTD(8, jA)
                If Mid(NewStdDwg, 1, 2) = "MX" Then
                    'NTest1 = FindSTD(1, jA)               'A    
                    'NTest2 = FindSTD(2, jA)               'B  
                    'NTest3 = FindSTD(3, jA)               'C    
                    'Ntest4 = FindSTD(4, jA)               'D    
                    NewQty = FindSTD(5, jA)                'E   
                    'Ntest6 = FindSTD(6, jA)               'F   

                    NewDesc = FindSTD(6, jA)               'G  
                    NewDesc = RTrim(NewDesc)                'Remove extra blank spaces on end.
                    NewDesc2 = NewDesc
                    NewInv = FindSTD(7, jA)                'H  'INV-1        
                    NewStdDwg = FindSTD(8, jA)             'I  'Std Dwg No.
                    'NTest7 = FindSTD(7, jA)               'G
                    'NTest8 = FindSTD(8, jA)               'H
                    'NTest9 = FindSTD(9, jA)               'I
                    'NTest10 = FindSTD(10, jA)             'J
                    'NTest11 = FindSTD(11, jA)             'K
                    'NTest12 = FindSTD(12, jA)             'L
                    NewBOMPos = FindSTD(13, jA)             'M

                    GoTo GetStdInfo
                End If

            Next jA

GetStdInfo:
            Totalj = UBound(OldBOM, 2)                      'OldBOM = Matrix Std's for BOM's

            '--------------------------------------------------------------------
            For j = 0 To UBound(OldBOM, 2)                      'First Find MX Standard
                LookForStd = OldBOM(1, j)                       'LookForStd = .Range("H" & RowNo).Value
                'May want to look at modifing program to look for Standard match before checking every line
                'in Spreadsheet see example below:   (Hold for now.)<---Has Been Done
                '(Program needs to find second line before getting parts on Standard.
                '24"x12" FLAT BOTTOM SUMP (TYPE-A)
                '24"x10" DISHED BOTTOM SUMP (TYPE-B)

                If Mid(LookForStd, 1) = NewStdDwg Then
                    'OTest1 = OldBOM(1, j)               'A
                    'OTest2 = OldBOM(2, j)               'B
                    'OTest3 = OldBOM(3, j)               'C
                    OldQty = OldBOM(4, j)               'D
                    'OTest5 = OldBOM(5, j)               'E
                    'Otest6 = OldBOM(6, j)              'F
                    OldDesc = OldBOM(6, j)              'G  'Description    
                    OldDesc2 = OldDesc                      'Temporary set to be the same below it is changed if condition is meet.
                    'OTest7 = OldBOM(7, j)              'G
                    OldInv = OldBOM(7, j)               'H  'INV-1    
                    'OTest8 = OldBOM(8, j)              'H
                    OldStdDwg = OldBOM(8, j)            'I  'Std Dwg No.
                    'OTest9 = OldBOM(9, j)               'I
                    'OTest10 = OldBOM(10, j)             'J
                    'OTest11 = OldBOM(11, j)             'K
                    'OTest12 = OldBOM(12, j)             'L

                    SeeNotePos = 0
                    ChkPos = 0
                    SeeDwgPos = 0
                    Note2Pos = 0

                    SeeNotePos = InStr(1, OldDesc, SearchSeeNote)
                    ChkPos = InStr(1, OldDesc, SearchNote)
                    SeeDwgPos = InStr(1, OldDesc, SearchDwg)
                    Note2Pos = InStr(1, OldDesc, SearchNote2)

                    Select Case 0
                        Case Is < SeeNotePos
                            OldDesc = Mid(OldDesc, 1, (SeeNotePos - 2))
                        Case Is < ChkPos
                            OldDesc2 = Mid(OldDesc, 1, (ChkPos - 1))
                        Case Is < SeeDwgPos
                            GoTo GetData4
                        Case Is < Note2Pos
                            OldDesc2 = Mid(OldDesc, 1, (Note2Pos - 1))
                    End Select

                    '-----------------------Could not see values
                    NewDesc = NewDesc
                    OldDesc2 = OldDesc2
                    NewInv = NewInv

                    '-------------------Same Problem exist in NewDesc as OldDesc
                    NewSeeNotePos = 0
                    NewChkPos = 0
                    NewSeeDwgPos = 0
                    NewNote2Pos = 0

                    NewSeeNotePos = InStr(1, NewDesc, SearchSeeNote)
                    NewChkPos = InStr(1, NewDesc, SearchNote)
                    NewSeeDwgPos = InStr(1, NewDesc, SearchDwg)
                    NewNote2Pos = InStr(1, NewDesc, SearchNote2)

                    Select Case 0
                        Case Is < NewSeeNotePos
                            NewDesc = Mid(NewDesc, 1, (NewSeeNotePos - 2))         'Question should this be minus 1
                            NewDesc = RTrim(NewDesc)
                        Case Is < NewChkPos
                            NewDesc2 = Mid(NewDesc, 1, (NewChkPos - 1))
                            NewDesc2 = RTrim(NewDesc2)
                        Case Is < NewSeeDwgPos
                            GoTo GetData4
                        Case Is < NewNote2Pos
                            NewDesc2 = Mid(NewDesc, 1, (NewNote2Pos - 2))
                            NewDesc2 = RTrim(NewDesc2)
                    End Select

                    NewDesc = NewDesc
                    NewDesc2 = NewDesc2
                    OldDesc2 = OldDesc2
                    NewInv = NewInv

                    Select Case NewDesc
                        Case OldDesc                    '----------------Check Out Old description First.
                            If OldInv = NewInv Then
                                Startj = j
                                StartCnt = 0
                                GoTo GetData5
                            Else
                                GoTo GetData4
                            End If
                        Case OldDesc2
                            If OldInv = NewInv Then
                                Startj = j
                                GoTo GetData5
                            Else
                                GoTo GetData4
                            End If
                        Case Else
                            Select Case NewDesc2        '------------Setup search for NewDesc without Notes.
                                Case OldDesc
                                    If OldInv = NewInv Then
                                        Startj = j
                                        GoTo GetData5
                                    Else
                                        GoTo GetData4   'If Inventory numbers do not match go to next item in list.
                                    End If
                                Case OldDesc2
                                    If OldInv = NewInv Then
                                        Startj = j
                                        GoTo GetData5
                                    Else
                                        GoTo GetData4   'If Inventory numbers do not match go to next item in list.
                                    End If
                                Case Else
                                    GoTo GetData4
                            End Select
                    End Select

                Else
                    GoTo GetData4
                End If

                pattern = NTest                         '----------------------BOM Std that needs parts.
                If pattern <> "" Then
                    Dim matches As MatchCollection = Regex.Matches(OTest, pattern)       'Look at Std's List find BOM std and get parts.
                    If Regex.IsMatch(OTest, pattern) Then
                        For Each match As Match In matches
                            FoundPart = match.Value
                        Next
                    End If
                End If
GetData5:
                FirstTimeThru = "Yes"
                OldQty = NewQty                             'Transfer Qty's
                CountNewItems = (CountNewItems + 1)         'Count equal to items inserted from Matrix STD's

                For l = (Startj + 1) To UBound(OldBOM, 2)       '-------Found Part Now get Items for Standard.
                    NewDwg = OldBOM(1, l)               'A
                    NewRev = OldBOM(2, l)               'B
                    NewShpMk = OldBOM(3, l)             'C
                    NewPcMk = OldBOM(4, l)              'D
                    NewQty = OldBOM(5, l)               'E
                    NewDesc = OldBOM(6, l)              'F
                    NewInv = OldBOM(7, l)               'G
                    NewStd = OldBOM(8, l)               'H
                    NewMatl = OldBOM(9, l)              'I
                    NewWht = OldBOM(10, l)              'J
                    NewReq = OldBOM(11, l)              'K
                    'NewProd = OldBOM(12, l)             'L         'Index Out of Bounds OldBOM only looks at first 11 columns.

                    If NewPcMk = Nothing Then
                        If CompDesc <> "" Then
                            With BOMWrkSht
                                If FirstTimeThru = "Yes" Then
                                    If CompDesc = "" Then
                                        .Range("M" & (jA + 4)).Value = "Standard was not found."
                                        With .Range("A" & (jA + 4) & ":P" & (jA + 4))
                                            With .Interior
                                                .ColorIndex = 7
                                                .Pattern = Constants.xlSolid
                                            End With
                                        End With
                                    Else
                                        .Range("M" & (jA + 4)).Value = "Standard Reference only, No additional parts. " & CompDesc
                                        With .Range("A" & (jA + 4) & ":P" & (jA + 4))
                                            With .Interior
                                                .ColorIndex = 45
                                                .Pattern = Constants.xlSolid
                                            End With
                                        End With
                                    End If
                                Else
                                    .Range("M" & (jA + 4)).Value = CompDesc
                                    With .Range("A" & (jA + 4) & ":P" & (jA + 4))
                                        With .Interior
                                            .ColorIndex = 45
                                            .Pattern = Constants.xlSolid
                                        End With
                                    End With
                                End If
                            End With
                        Else
                            If FirstTimeThru = "Yes" Then
                                Count = BOMWrkSht.Range("E4000").End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Row

                                Totalj = UBound(NewBOM, 2)
                                Count = (Count - UBound(NewBOM, 2))
                                Count = (Count - 4)

                                With BOMWrkSht
                                    .Range("M" & (NewBOMPos + Count)).Value = "Standard Reference only, No additional parts."
                                    With .Range("A" & (NewBOMPos + Count) & ":P" & (NewBOMPos + Count))
                                        With .Interior
                                            .ColorIndex = 8
                                            .Pattern = Constants.xlSolid
                                        End With
                                    End With
                                End With
                            End If
                        End If
                        CompDesc = ""
                        FirstTimeThru = "No"
                        GetNewjA = jA
                        GoTo GetDataNew
                    End If

                    WorkShtName = "Bulk BOM"
                    BOMWrkSht = Workbooks.Application.Worksheets(WorkShtName)
                    Count = BOMWrkSht.Range("E4000").End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Row

                    Totalj = UBound(NewBOM, 2)
                    Count = (Count - UBound(NewBOM, 2))
                    Count = (Count - 4)
                    'Count = the amount of Item that needs to be added to the position
                    '----------------------------------------------------------

                    With BOMWrkSht                  '----------Insert Items into Bulk BOM.
                        FirstTimeThru = "No"

                        If FoundItem = "Yes" Then
                            LineNo = jA
                        Else
                            LineNo = jA
                            FoundItem = "Yes"
                        End If

                        If Count <> NewCount Then       'LineNo did not advance for dup part on different drawing.
                            If DescFixed = "No" Then
                                LineNo = (LineNo + Count + 3)
                            Else
                                LineNo = (LineNo + Count + 2)
                            End If
                        Else
                            If Startj > LineNo Then
                                LineNo = (Startj + Count + 4)   'Look at replacing LineNo with Startj
                            Else
                                LineNo = (LineNo + Count + 3)   'First time thru is correct, and next parts for Standard.
                            End If
                        End If                              'Second Std with parts is correct.

                        OldDesc = OldDesc

                        'Found problem here were program is inserting duplicates due to LineNo is wrong
                        'above for new part which is the same part on different drawing.
                        With BOMWrkSht      'Need to fix condition were we use NewBomPos or NewBomPos Plus Count.
                            '.Rows(NewBOMPos + 1 & ":" & NewBOMPos + 1).Insert()     '.Rows(LineNo + 2 & ":" & LineNo + 2).Insert()
                            'CountNewItems = Running count of Items inserted after material was found for
                            ' Standard Found.
                            If Count > 0 Then   'Make sure to format all lines inserted if not you will have no lines.
                                LineNo = (NewBOMPos + Count + 1)

                                FileToOpen = "Bulk BOM"
                                FormatLine3(LineNo, FileToOpen)  'This inserts a line and formats the line replaced above insert
                            Else
                                LineNo = (NewBOMPos + 1)

                                FileToOpen = "Bulk BOM"
                                FormatLine3(LineNo, FileToOpen)
                            End If
                        End With

                        .Range("A" & LineNo).Value = NewDwg
                        .Range("B" & LineNo).Value = NewRev
                        .Range("C" & LineNo).Value = NewShpMk
                        .Range("D" & LineNo).Value = NewPcMk
                        .Range("E" & LineNo).Value = (NewQty * OldQty)
                        .Range("F" & LineNo).Value = NewDesc
                        .Range("G" & LineNo).Value = NewInv
                        .Range("H" & LineNo).Value = NewStd
                        .Range("I" & LineNo).Value = NewMatl
                        If NewWht = "-" Then
                            .Range("J" & LineNo).Value = NewWht
                        Else
                            .Range("J" & LineNo).Value = (NewWht * (NewQty * OldQty))
                        End If
                        .Range("K" & LineNo).Value = NewReq
                        '.Range("L" & LineNo).Value = NewProd

                        If CompDesc <> "" Then
                            .Range("M" & RowNo).Value = CompDesc
                            With .Range("A" & RowNo & ":P" & RowNo)
                                With .Interior
                                    .ColorIndex = 45
                                    .Pattern = Constants.xlSolid
                                End With
                            End With
                            With .Range("A" & (RowNo + 1) & ":P" & RowNo)
                                With .Interior
                                    .ColorIndex = 45
                                    .Pattern = Constants.xlSolid
                                End With
                            End With
                        Else
                            If DescFixed = "Yes" Then
                                With .Range("A" & LineNo & ":P" & LineNo)
                                    With .Interior
                                        .ColorIndex = 8
                                        .Pattern = Constants.xlSolid
                                    End With
                                End With
                            Else
                                .Range("M" & (LineNo - 1)).Value = "Found Standard Information"
                                With .Range("A" & (LineNo - 1) & ":P" & (LineNo - 1))
                                    With .Interior
                                        .ColorIndex = 8
                                        .Pattern = Constants.xlSolid
                                    End With
                                End With
                                DescFixed = "Yes"
                                With .Range("A" & LineNo & ":P" & LineNo)
                                    With .Interior
                                        .ColorIndex = 8
                                        .Pattern = Constants.xlSolid
                                    End With
                                End With
                            End If
                        End If
                    End With
                    NewCount = (Count + 1)
                    Count = (Count + 1)

                Next l

GetData4:
            Next j

            '----------------------------------------------------------------------------------
            '-----------Create new loop to recheck for Part No When Description does not match.
            '----------------------------------------------------------------------------------

            With BOMWrkSht          '------------------------Replace Below with above.  No No No
                If j > UBound(OldBOM, 2) Then
                    Totalj = UBound(OldBOM, 2)          'Description does not match
                    CompDesc = ""

                    For j = 0 To UBound(OldBOM, 2)
                        LookForStd = OldBOM(1, j)

                        If Mid(LookForStd, 1) = NewStdDwg Then
                            'OTest1 = OldBOM(1, j)               'A
                            'OTest2 = OldBOM(2, j)               'B
                            'OTest3 = OldBOM(3, j)               'C
                            'Otest4 = OldBOM(4, j)               'D
                            'OTest5 = OldBOM(5, j)               'E
                            'Otest6 = OldBOM(6, j)              'F
                            OldDesc = OldBOM(6, j)              'G  'Description   
                            'OTest7 = OldBOM(7, j)              'G
                            OldInv = OldBOM(7, j)               'H  'INV-1        
                            'OTest8 = OldBOM(8, j)              'H
                            OldStdDwg = OldBOM(8, j)            'I  'Std Dwg No.
                            'OTest9 = OldBOM(9, j)               'I
                            'OTest10 = OldBOM(10, j)             'J
                            'OTest11 = OldBOM(11, j)             'K
                            'OTest12 = OldBOM(12, j)             'L

                            If OldInv = Nothing Then
                                'Do nothing except go to next line
                            Else
                                If OldInv = NewInv Then
                                    CompDesc = "Description Did Not Match double check Item"
                                    Startj = j
                                    StartjA = (jA + 4)
                                    OldQty = NewQty
                                    GetStdInfo(CompDesc, Startj, StartjA, OldQty, FuncGetDataNew, NewBOMPos)   'GoTo GetData5  'GoTo GetData
                                    DescFixed = "Yes"
                                    CompDesc = ""
                                    FirstTimeThru = "No"
                                    GetNewjA = jA
                                    GoTo GetDataNew
                                End If
                            End If

                        End If

                    Next j

                    If IsNothing(BOMWrkSht) = True Then
                        WorkShtName = "Bulk BOM"
                        BOMWrkSht = Workbooks.Application.Worksheets(WorkShtName)
                    End If

                    With BOMWrkSht                  '----------Insert Items into Bulk BOM.
                        DescFixed = DescFixed
                        If DescFixed = "No" Then        'Found new problem were Item was found, and while looking thru list
                            OldDesc = OldDesc           ' end was found but part was found.
                            NewDesc = NewDesc

                            '-----------------------For some reason the count is off by 4 times.
                            'Get new total count for parts inserted from Matrix standards.
                            Count = BOMWrkSht.Range("E4000").End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Row

                            Totalj = UBound(NewBOM, 2)
                            Count = (Count - UBound(NewBOM, 2))
                            Count = (Count - 4)

                            .Range("M" & (NewBOMPos + Count)).Value = "Standard was not found."
                            With .Range("A" & (NewBOMPos + Count) & ":P" & (NewBOMPos + Count))
                                With .Interior
                                    .ColorIndex = 7
                                    .Pattern = Constants.xlSolid
                                End With
                            End With
                        End If
                    End With
                    If jA > 1 Then
                        GoTo GetDataNew
                    Else
                        GoTo GetData2
                    End If
                End If

                If NewPcMk = Nothing Then
                    NewPcMk = "GetData"
                End If
GetData:
                FirstTimeThru = "Yes"

                For j = (RowNo2 + 1) To LineNo4
                    With StdItemsWrkSht
                        NewDwg = .Range("A" & j).Value         'Dwg
                        NewRev = .Range("B" & j).Value         'Rev
                        NewShpMk = .Range("C" & j).Value       'Ship Mark
                        NewPcMk = .Range("D" & j).Value        'New Piece Mark
                        NewQty = .Range("E" & j).Value         'Qty
                        NewDesc = .Range("F" & j).Value        'Description
                        NewInv = .Range("G" & j).Value         'Inv-1
                        NewStd = .Range("H" & j).Value         'Standard Number Example:MX1001A
                        NewMatl = .Range("I" & j).Value        'Material
                        NewWht = .Range("J" & j).Value         'Weight
                        NewReq = .Range("K" & j).Value         'Required Type
                        NewProd = .Range("L" & j).Value        'Production Code
                        If NewPcMk = Nothing Then
                            If CompDesc <> "" Then
                                With BOMWrkSht
                                    If FirstTimeThru = "Yes" Then
                                        If CompDesc = "" Then
                                            .Range("M" & RowNo).Value = "Standard was not found."
                                            With .Range("A" & RowNo & ":P" & RowNo)
                                                With .Interior
                                                    .ColorIndex = 7
                                                    .Pattern = Constants.xlSolid
                                                End With
                                            End With
                                        Else
                                            .Range("M" & RowNo).Value = "Standard Reference only, No additional parts. " & CompDesc
                                            With .Range("A" & RowNo & ":P" & RowNo)
                                                With .Interior
                                                    .ColorIndex = 45
                                                    .Pattern = Constants.xlSolid
                                                End With
                                            End With
                                        End If
                                    Else
                                        .Range("M" & RowNo).Value = CompDesc
                                        With .Range("A" & RowNo & ":P" & RowNo)
                                            With .Interior
                                                .ColorIndex = 45
                                                .Pattern = Constants.xlSolid
                                            End With
                                        End With
                                    End If
                                End With
                            Else
                                If FirstTimeThru = "Yes" Then
                                    With BOMWrkSht
                                        .Range("M" & RowNo).Value = "Standard Reference only, No additional parts."
                                        With .Range("A" & RowNo & ":P" & RowNo)
                                            With .Interior
                                                .ColorIndex = 8
                                                .Pattern = Constants.xlSolid
                                            End With
                                        End With
                                    End With
                                End If
                            End If
                            CompDesc = ""
                            FirstTimeThru = "No"
                            GoTo GetData2
                        End If
                    End With

                    With BOMWrkSht                  '----------Insert Items into Bulk BOM.
                        FirstTimeThru = "No"
                        If FoundItem = "Yes" Then
                            LineNo = ((RowNo - 1) + Count)
                        Else
                            LineNo = (RowNo - 1)
                            FoundItem = "Yes"
                        End If
                        FileToOpen = "Bulk BOM"
                        FormatLine(LineNo, FileToOpen)
                        .Range("A" & ((RowNo + 1) + Count)).Value = NewDwg
                        .Range("B" & ((RowNo + 1) + Count)).Value = NewRev
                        .Range("C" & ((RowNo + 1) + Count)).Value = NewShpMk
                        .Range("D" & ((RowNo + 1) + Count)).Value = NewPcMk
                        .Range("E" & ((RowNo + 1) + Count)).Value = (NewQty * OldQty)
                        .Range("F" & ((RowNo + 1) + Count)).Value = NewDesc
                        .Range("G" & ((RowNo + 1) + Count)).Value = NewInv
                        .Range("H" & ((RowNo + 1) + Count)).Value = NewStd
                        .Range("I" & ((RowNo + 1) + Count)).Value = NewMatl
                        If NewWht = "-" Then
                            .Range("J" & ((RowNo + 1) + Count)).Value = NewWht
                        Else
                            .Range("J" & ((RowNo + 1) + Count)).Value = (NewWht * OldQty)
                        End If
                        .Range("K" & ((RowNo + 1) + Count)).Value = NewReq
                        .Range("L" & ((RowNo + 1) + Count)).Value = NewProd

                        If CompDesc <> "" Then
                            .Range("M" & RowNo).Value = CompDesc
                            With .Range("A" & RowNo & ":P" & RowNo)
                                With .Interior
                                    .ColorIndex = 45
                                    .Pattern = Constants.xlSolid
                                End With
                            End With
                            With .Range("A" & (RowNo + 1) & ":P" & RowNo)
                                With .Interior
                                    .ColorIndex = 45
                                    .Pattern = Constants.xlSolid
                                End With
                            End With
                        Else
                            .Range("M" & RowNo).Value = "Found Standard Information"
                            With .Range("A" & RowNo & ":P" & RowNo)
                                With .Interior
                                    .ColorIndex = 8
                                    .Pattern = Constants.xlSolid
                                End With
                            End With
                            With .Range("A" & (RowNo + 1) & ":P" & RowNo)
                                With .Interior
                                    .ColorIndex = 8
                                    .Pattern = Constants.xlSolid
                                End With
                            End With
                        End If
                    End With
                    Count = (Count + 1)

                Next j
GetData2:
                CountVal = (CountVal + Count)
                Count = 0
                FoundItem = "No"
                CompDesc = ""
                BOMMnu.ProgressBar1.Value = i
            End With
        Next i

        If (LineNo2 + CountVal) > i Then
            GoTo RptGetData
        End If

FoundAllParts:

Err_FindStdsBOM:
        ErrNo = Err.Number

        If ErrNo <> 0 Then
            PrgName = "FindStdsBOM"
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

            HandleErrSQL(PrgName, ErrNo, ErrMsg, ErrSource, PriPrg, ErrDll, DwgItem, PrgLineNo)

            If GenInfo.UserName = "dlong" Then
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

    Public Function GetStdInfo(ByVal CompDesc As String, ByVal Startj As Integer, ByVal StartjA As Integer, ByVal OldQty As Integer, ByVal FuncGetDataNew As String, ByVal NewBOMPos As Integer) As Object
        Dim ExceptPos, jA, l, GetNewjA, TotalOnNewBOM As Integer
        Dim WorkShtName, FoundLast, LineNo, SearchException, FoundItem, SearchSeeNote As String
        Dim SearchNote2, SearchNote, SearchDwg, pattern As String
        Dim FileToOpen, NewDwg, NewRev, NewShpMk, NewPcMk, NewQty, NewDesc, NewInv, NewStd As String
        Dim NewProd, DescFixed, NewMatl, NewWht, NewReq As String
        Dim Workbooks As Excel.Workbooks
        Dim BOMWrkSht As Worksheet
        Dim ExcelApp As Object
        Dim BOMMnu As ShippingList_Menu
        BOMMnu = Me

        jA = StartjA
        PrgName = "GetStdInfo"
        DescFixed = "No"
        FoundItem = Nothing

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

        On Error GoTo Err_GetStdInfo

        Workbooks = ExcelApp.Workbooks
        WorkShtName = "Bulk BOM"
        BOMWrkSht = Workbooks.Application.Worksheets(WorkShtName)
        BOMWrkSht.Activate()
        Count = BOMWrkSht.Range("E4000").End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Row

        TotalOnNewBOM = UBound(NewBOM, 2)
        Count = (Count - UBound(NewBOM, 2))
        Count = (Count - 4)
        Dim OldCount As Integer
        OldCount = UBound(OldBOM, 2)
        FirstTimeThru = "Yes"

        For l = (Startj + 1) To UBound(OldBOM, 2)                  '-------Found Part Now get Items for Standard.
            '                                   With StdItemsWrkSht
            NewDwg = OldBOM(1, l)               'A  'NewDwg = .Range("A" & j).Value         'Dwg
            NewRev = OldBOM(2, l)               'B  'NewRev = .Range("B" & j).Value         'Rev
            NewShpMk = OldBOM(3, l)             'C  'NewShpMk = .Range("C" & j).Value       'Ship Mark
            NewPcMk = OldBOM(4, l)              'D  'NewPcMk = .Range("D" & j).Value        'New Piece Mark
            NewQty = OldBOM(5, l)               'E  'NewQty = .Range("E" & j).Value         'Qty
            NewDesc = OldBOM(6, l)              'F  'NewDesc = .Range("F" & j).Value        'Description
            NewInv = OldBOM(7, l)               'G  'NewInv = .Range("G" & j).Value         'Inv-1
            NewStd = OldBOM(8, l)               'H  'NewStd = .Range("H" & j).Value         'Standard Number Example:MX1001A
            NewMatl = OldBOM(9, l)              'I  'NewMatl = .Range("I" & j).Value        'Material
            NewWht = OldBOM(10, l)              'J  'NewWht = .Range("J" & j).Value         'Weight
            NewReq = OldBOM(11, l)              'K  'NewReq = .Range("K" & j).Value         'Required Type
            'NewProd = OldBOM(12, l)            'L  'Index Out of Bounds OldBOM only looks at first 11 columns.     'Production Code

            If NewPcMk = Nothing Then
                If CompDesc <> "" Then
                    With BOMWrkSht
                        If FirstTimeThru = "Yes" Then
                            If CompDesc = "" Then
                                .Range("M" & (NewBOMPos + Count)).Value = "Standard was not found."
                                With .Range("A" & (NewBOMPos + Count) & ":P" & (NewBOMPos + Count))
                                    With .Interior
                                        .ColorIndex = 7
                                        .Pattern = Constants.xlSolid
                                    End With
                                End With
                            Else
                                .Range("M" & (NewBOMPos + Count)).Value = "Standard Reference only, No additional parts. " & CompDesc
                                With .Range("A" & (NewBOMPos + Count) & ":P" & (NewBOMPos + Count))
                                    With .Interior
                                        .ColorIndex = 45
                                        .Pattern = Constants.xlSolid
                                    End With
                                End With
                            End If
                        Else
                            If FoundItem <> "Yes" Then
                                .Range("M" & NewBOMPos).Value = CompDesc
                                With .Range("A" & NewBOMPos & ":P" & NewBOMPos)
                                    With .Interior
                                        .ColorIndex = 45
                                        .Pattern = Constants.xlSolid
                                    End With
                                End With
                            End If
                        End If
                    End With
                Else
                    If FirstTimeThru = "Yes" Then
                        With BOMWrkSht
                            .Range("M" & NewBOMPos).Value = "Standard Reference only, No additional parts."
                            With .Range("A" & NewBOMPos & ":P" & NewBOMPos)
                                With .Interior
                                    .ColorIndex = 8
                                    .Pattern = Constants.xlSolid
                                End With
                            End With
                        End With
                    End If
                End If
                CompDesc = ""
                FirstTimeThru = "No"
                GetNewjA = jA
                FuncGetDataNew = "GetDataNew"
                GoTo Err_GetStdInfo
            End If

            RowNo = StartjA

            With BOMWrkSht                  '----------Insert Items into Bulk BOM.
                FirstTimeThru = "No"           'Move to bottom
                If FoundItem = "Yes" Then
                    LineNo = (NewBOMPos - 1)
                Else
                    LineNo = (NewBOMPos - 1)
                    FoundItem = "Yes"
                End If
                FileToOpen = "Bulk BOM"

                LineNo = (LineNo + Count)

                If LineNo > 600 Then
                    Stop
                End If

                FormatLine(LineNo, FileToOpen)
                .Range("A" & ((NewBOMPos + 1) + Count)).Value = NewDwg      '
                .Range("B" & ((NewBOMPos + 1) + Count)).Value = NewRev      '
                .Range("C" & ((NewBOMPos + 1) + Count)).Value = NewShpMk    '
                .Range("D" & ((NewBOMPos + 1) + Count)).Value = NewPcMk     '
                .Range("E" & ((NewBOMPos + 1) + Count)).Value = (NewQty * OldQty)
                .Range("F" & ((NewBOMPos + 1) + Count)).Value = NewDesc     '
                .Range("G" & ((NewBOMPos + 1) + Count)).Value = NewInv      '
                .Range("H" & ((NewBOMPos + 1) + Count)).Value = NewStd      '
                .Range("I" & ((NewBOMPos + 1) + Count)).Value = NewMatl     '
                If NewWht = "-" Then
                    .Range("J" & ((NewBOMPos + 1) + Count)).Value = NewWht
                Else
                    .Range("J" & ((NewBOMPos + 1) + Count)).Value = (NewWht * OldQty)
                End If
                .Range("K" & ((NewBOMPos + 1) + Count)).Value = NewReq
                '.Range("L" & ((RowNo + 1) + Count)).Value = NewProd             'Is never set Remove this line

                If CompDesc <> "" Then
                    If DescFixed = "No" Then
                        .Range("M" & (NewBOMPos + Count)).Value = CompDesc
                        DescFixed = "Yes"
                    End If

                    With .Range("A" & (NewBOMPos + Count) & ":P" & (NewBOMPos + Count))
                        With .Interior
                            .ColorIndex = 45
                            .Pattern = Constants.xlSolid
                        End With
                    End With

                    With .Range("A" & (NewBOMPos + 1 + Count) & ":P" & (NewBOMPos + 1 + Count))
                        With .Interior
                            .ColorIndex = 45
                            .Pattern = Constants.xlSolid
                        End With
                    End With
                Else
                    .Range("M" & NewBOMPos).Value = "Found Standard Information"

                    With .Range("A" & NewBOMPos & ":P" & NewBOMPos)
                        With .Interior
                            .ColorIndex = 8
                            .Pattern = Constants.xlSolid
                        End With
                    End With

                    With .Range("A" & (NewBOMPos + 1) & ":P" & NewBOMPos)
                        With .Interior
                            .ColorIndex = 8
                            .Pattern = Constants.xlSolid
                        End With
                    End With
                End If
            End With
            FirstTimeThru = "No"
            Count = (Count + 1)

        Next l

Err_GetStdInfo:
        ErrNo = Err.Number

        If ErrNo <> 0 Then
            PrgName = "GetStdInfo"
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

            HandleErrSQL(PrgName, ErrNo, ErrMsg, ErrSource, PriPrg, ErrDll, DwgItem, PrgLineNo)

            If GenInfo.UserName = "dlong" Then
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

    Sub BulkBom()
        Dim BOMMnu As ShippingList_Menu
        BOMMnu = Me
        BOMMnu.Show()
    End Sub

    Function HandleErrSQL(ByVal PrgName As String, ByVal ErrNo As String, ByVal ErrMsg As String, ByVal ErrSource As String, ByVal PriPrg As String, ByVal ErrDll As String, ByVal DwgItem As String, ByVal PrgLineNo As String)
        Dim sqlConn As ADODB.Connection
        Dim rs As ADODB.Recordset
        Dim sqlStr As String
        Dim ErrDate As Date
        Dim QuoteMkPos As Integer
        Dim UserName, ProgramNotes, ErrMsgPart1, ErrMsgPart2 As String

        sqlConn = New ADODB.Connection
        ErrDate = Now

        If IsNothing(GenInfo.UserName) = True Then
            GenInfo.UserName = System.Environment.UserName      '-------DJL-06-30-2025      'GenInfo.UserName = UserNamex()
        End If

        UserName = GenInfo.UserName
        ProgramNotes = "VB Net 64bit for Inventor Programming Testing"

        QuoteMkPos = InStr(ErrMsg, Chr(39))

        While QuoteMkPos > 0                                        'Remove single quotes that cause database errors.
            ErrMsgPart1 = Mid(ErrMsg, 1, (QuoteMkPos - 1))
            ErrMsgPart2 = Mid(ErrMsg, (QuoteMkPos + 1), Len(ErrMsg))
            ErrMsg = (ErrMsgPart1 & ErrMsgPart2)

            QuoteMkPos = InStr(ErrMsg, Chr(39))
        End While

        '-----------------------------------------------------Save Errors to sql database
        sqlStr = "INSERT  INTO ErrCollection (PrimaryPrg, PrgName, ErrNo, ErrMsg, ErrDate, UserName, ProgramNotes, ErrSource, ErrDll, DwgName) " &
        "VALUES ('" & PriPrg & "', '" & PrgName & "', '" & ErrNo & "', '" & ErrMsg & "', '" & ErrDate & "', '" & UserName & "', '" & ProgramNotes & "', '" & ErrSource & "', '" & ErrDll & "', '" & DwgItem & "')"
        db_String = "Server=MTX16SQL09\Engineering;Database=HandleErrors;User=devDennis;Password=d3v3lop3r;Trusted_Connection=False"

        Dim ConJobLog As New SqlClient.SqlConnection
        ConJobLog = New SqlClient.SqlConnection
        ConJobLog.ConnectionString = db_String
        ConJobLog.Open()
        Dim command2 As New SqlCommand(sqlStr, ConJobLog)
        Dim Writer2
        Writer2 = command2.ExecuteReader
        ConJobLog.Close()
        ConJobLog = Nothing
        sqlConn = Nothing
        rs = Nothing
    End Function

    Private Sub BtnGetMWInfo_Click(sender As Object, e As EventArgs) Handles BtnGetMWInfo.Click
        Dim ErrHandler2 As New ErrHandler3135
        Dim i, j, k As Integer
        Dim BadPart As String       ', TryYes As String
        'Dim TestNew
        Dim FileNam2, FindRev3, PrevRev As String
        Dim lReturn As Int64

        On Error GoTo Err_BtnGetMWInfo

        '-----------------------------------------------------------------------------------------
        SelectList.Items.Clear()
        StartDir = System.Environment.SpecialFolder.Recent

        If StartDir = "0" Then
            StartDir = Me.PathBox.Text                          'StartDir = "C:\AdeptWork\"
        End If

        OpenFileDialog1.InitialDirectory = StartDir
        OpenFileDialog1.Filter = "BOM Spreadsheet(*.xls*)|*.xls*;"              '-------DJL-7-15-2025
        OpenFileDialog1.Title = "Select BOM file to Open"                       '-------DJL-7-15-2025
        lReturn = OpenFileDialog1.ShowDialog()

        FileNam = OpenFileDialog1.FileName                          'Has to be 64 bit Inorder for this to work.
        FileNam2 = OpenFileDialog1.SafeFileName
        NewDir = FileNam.Replace(FileNam2, "")

        GenInfo3135.JobDir = NewDir
        PathBox.Text = NewDir
        PathBox2.Text = NewDir

        Me.PathBox.BackColor = System.Drawing.Color.White
        Me.BtnGetMWInfo.BackColor = System.Drawing.Color.White
        Me.CollectBOMList.BackColor = System.Drawing.Color.GreenYellow              '-------DJL-7-15-2025
        Me.BtnAdd.BackColor = System.Drawing.Color.GreenYellow
        Me.BtnClear.BackColor = System.Drawing.Color.GreenYellow
        Me.BtnRemove.BackColor = System.Drawing.Color.GreenYellow
        '-----------------------------------------------------------------------------------------
        sender = "Excel"
        SecondChk = "First"

        If Directory.Exists("K:\AWA\" & System.Environment.UserName & "\AdeptWork\") = False Then
            StartAdept = 0
            ClosePrg("Excel", "First", StartAdept)
            'ClosePrg("Acad", "First", StartAdept)          '-------DJL-07-15-2025
        End If

        ErrFound = ""

        Dim Dir1 As DirectoryInfo = New DirectoryInfo(GenInfo3135.JobDir)
        Dim BOMItem

        If CollectBOMList.Items.Count = 0 Then              '-------DJL-7-15-2025
            For Each BOMItem In Dir1.GetFiles("*.xls*")              '-------DJL-7-15-2025
                ExtPos = 0
                ExtPos2 = 0              '-------DJL-7-15-2025
                BOMPos = 0              '-------DJL-7-15-2025

                FindDwg = BOMItem.ToString
                ExtPos = InStr(1, FindDwg, ".xls")              '-------DJL-7-15-2025
                ExtPos2 = InStr(1, FindDwg, ".xlsx")              '-------DJL-7-15-2025
                BOMPos = InStr(1, FindDwg, "BulkBOM")              '-------DJL-7-15-2025

                If ExtPos > 0 Or ExtPos2 > 0 Then              '-------DJL-7-15-2025
                    If BOMPos > 0 Then              '-------DJL-7-15-2025
                        CollectBOMList.Items.Add(BOMItem)              '-------DJL-7-15-2025
                    End If
                End If
            Next BOMItem
        End If

        BtnAdd.Enabled = True
        BtnRemove.Enabled = True
        BtnClear.Enabled = True

        'btnNewDir.Enabled = True
        File2.Path = PathBox.Text
        'GroupBox5.Text = "Current Directory = " & PathBox.Text

        If Me.File2.Items.Count > 0 Then
RestartFile2:
            FoundRev = 0
            CntSpreadSht = (Me.File2.Items.Count - 1)

            For i = 0 To CntSpreadSht
                BOMPos = 0

                If i > CntSpreadSht Then
                    GoTo Nexti
                End If

                FindRev = File2.Items.Item(i)
                RevPos = InStr(1, FindRev, "-R")
                BOMPos = InStr(1, FindRev, "ShippingList")

                If BOMPos > 0 Then
                    '    GoTo UpdateBOMRev
                    'End If
                    ShipListBox.Items.Add(FindRev)

                    FindRev = Mid(FindRev, (RevPos + 2), Len(FindRev))
                    FindRev1 = FindRev.Replace(".xls", "")

                    Dim pattern As String = "^[0-9]"                    'Also Look at \Nd  Number with dec.
                    Dim matches As MatchCollection = Regex.Matches(FindRev1, pattern)
                    If Regex.IsMatch(FindRev1, pattern) Then
                        For k = 1 To Len(FindRev1)          'For Each match As Match In matches     'Is not stepping through each number example: 13
                            FindRev2 = Mid(FindRev1, k, 1)

                            If Regex.IsMatch(FindRev2, pattern) Then
                                FindRev3 = FindRev3 & FindRev2
                            End If
                        Next k
                    End If

                    If FoundRev < FindRev3 Then
                        FoundRev = FindRev3
                    End If
UpdateBOMRev:
                End If
Nexti:
                If PrevRev < FoundRev Then
                    PrevRev = FoundRev
                    FindRev3 = Nothing
                End If
            Next i

            FoundRev = PrevRev
            Dim index As Integer
            index = ComboBox1.FindString(FoundRev + 1)
            ComboBox1.SelectedIndex = index
        End If

        If Me.ShipListBox.Items.Count = 1 Then
            ShipListBox.SelectedItem = Me.ShipListBox.Items.Item(0)
        End If

        ShipListBox.Sorted = True
        CollectBOMList.Sorted = True              '-------DJL-7-15-2025
        Me.BringToFront()
        Me.Show()

Err_BtnGetMWInfo:
        ErrNo = Err.Number

        If ErrNo <> 0 Then
            PrgName = "BtnGetMWInfo"
            PriPrg = "ShipListReadBOMAutoCAD"
            PrgName = "BtnGetMWInfo_Click"
            ErrMsg = Err.Description
            ErrSource = Err.Source
            ErrDll = Err.LastDllError
            ErrLastLineX = Err.Erl
            ErrException = Err.GetException

            If IsNothing(ManwayInfo3135.UserName) = True Then
                ManwayInfo3135.UserName = System.Environment.UserName
            End If

            Select Case ErrNo
                Case Is = 13 And InStr(ErrMsg, "Conversion from string") > 0
                    MsgBox("Program is having a time converting this item to String Error Msg = " & ErrMsg)
                    Stop
                    Resume
                Case Is = "91" And Mid(ErrMsg, 1, 24) = "Object reference not set"
                    Sapi.speak("Please open Primary Assembly in Inventor, and then pick OK on this box.")
                    MsgBox("Please open Primary Assembly in Inventor, and then pick OK on this box.")
                    ErrFound = "No Manway at Startup"
                    Resume Next
                    'Case Is = 429 And InStr(ErrMsg, "Operation unavailable") <> 0
                    '    'InvApp = GetObject(, "Inventor.Application")
                    '    InvApp = CreateInstance(InvAppType)

                    '    If IsNothing(InvApp) = True Then
                    '        Sapi.speak("Please open Primary Assembly in Inventor, and then pick OK on this box.")
                    '        MsgBox("Please open Inventor for some reason program cannot restart Inventor.")
                    '    Else
                    '        InvApp.Visible = True
                    '    End If

                    '    Resume Next
                    'Case Is = 462 And InStr(ErrMsg, "The RPC server is unavailable") <> 0
                    '    InvApp = CreateInstance(InvAppType)

                    '    If IsNothing(InvApp) = True Then
                    '        Sapi.speak("Please open Primary Assembly in Inventor, and then pick OK on this box.")
                    '        MsgBox("Please open Inventor for some reason program cannot restart Inventor.")
                    '    Else
                    '        InvApp.Visible = True
                    '    End If
                    '    Resume
                Case Is = -2147023170 And InStr(ErrMsg, "The remote procedure call failed") <> 0
                    MsgBox("Part " & BadPart & " Needs to be replaced using content center, some how this file is corrupt.")
                    Stop
                    Resume
                    'Case Is = "5" And Mid(ErrMsg, 1, 26) = "The parameter is incorrect"
                    '    Select Case FindParameter & PartFound
                    '        Case "TankDia" & PartFound
                    '            Err.Clear()
                    '        Case "d64" & PartFound                                  '-----Does not exist in Davit Manway
                    '            FindParameter = "DoesNotExist"                       'FindParameter = "d64" Does not exist on Davit Manway
                    '            Err.Clear()
                    '            Resume Next
                    '        Case "FlangeThk" & PartFound
                    '            FindParameter = PartFound & "Thickness"
                    '            Err.Clear()
                    '            Resume
                    '        Case "DimensionB" & PartFound                                           '-----Does not exist in Round Repad     'RepadOD
                    '            FindParameter = PartFound & "OD"
                    '            Err.Clear()
                    '            Resume
                    '        Case "DoesNotExist" & PartFound
                    '            FindParameter = "DoesNotExist"
                    '            Err.Clear()
                    '            Resume
                    '        Case Else
                    '            If ManwayInfo3134.UserName = "dlong" Then
                    '                MsgBox("Error Number " & ErrNo & ErrMsg & ErrSource)
                    '                Stop
                    '                Resume
                    '            Else
                    '                MsgBox("Program has new error, Error Msg = " & ErrMsg)
                    '                HandleErrSQL(PrgName, ErrNo, ErrMsg, ErrSource, PriPrg, ErrDll, DwgItem, PrgLineNo)
                    '            End If
                    '    End Select

                    '    Resume Next
                Case Else
                    If ManwayInfo3135.UserName = "dlong" Then
                        MsgBox("Error Number " & ErrNo & ErrMsg & ErrSource)
                        Stop
                        Resume
                    Else
                        MsgBox("Program has new error, Error Msg = " & ErrMsg)

                        Dim st As New StackTrace(Err.GetException, True)
                        CntFrames = st.FrameCount
                        GetFramesSrt = st.GetFrames
                        PrgLineNo = GetFramesSrt(CntFrames - 1).ToString
                        PrgLineNo = PrgLineNo.Replace("@", "at")
                        LenPrgLineNo = (Len(PrgLineNo))
                        PrgLineNo = Mid(PrgLineNo, 1, (LenPrgLineNo - 2))

                        HandleErrSQL(PrgName, ErrNo, ErrMsg, ErrSource, PriPrg, ErrDll, DwgItem, PrgLineNo)
                        'Stop
                    End If
            End Select
        End If

ExitCal:
    End Sub

    Public Class ErrHandler3135
        Inherits Exception

        Public ErrException As System.Exception
        Public ErrLastLineX As Integer
    End Class

    Public Function ClosePrg(ByVal sender As System.Object, ByVal SecondChk As String, ByVal StartAdept As Boolean)
        Dim myProcesses() As Process
        Dim instance As Process
        Dim Title, Msg, Style, Response As Object

        'ChkForAdditional:  'For some reason if Application does not close, this thoughs prg into endless loop.

        myProcesses = Process.GetProcessesByName(sender)                '-------Get Process if open example Excel.
        StartAdept = "False"

        If IsNothing(myProcesses) <> True Then
            For Each instance In myProcesses
                Select Case sender & SecondChk
                    Case "Adept" & "First"                            '-----------------------First Check
                        'GenInfo3233.StartAdept = True
                        'GenInfo3233.StartAdept = True
                        GoTo NextInstance
                    Case "acad" & "First"
                        'MsgBox("AutoCAD was closed due to Issues.")

                        '-------Looking for solution to remove message boxes in the back ground. & skip when users want to work on files in excel at the same time.
                        SecondChk = "Second"
                        Msg = "Do you have any AutoCAD Drawings open? if so please save and close them, Have you saved your work?"
                        Style = MsgBoxStyle.YesNo
                        Title = "Found AutoCAD Drawings open and need to make sure user saves files."
                        Response = MsgBox(Msg, Style, Title)

                        If Response = 6 Then
                            GoTo KillOpenFiles
                        Else
                            GoTo NextInstance
                        End If
                    Case "Excel" & "First"
                        '-------Looking for solution to remove message boxes in the back ground. & skip when users want to work on files in excel at the same time.
                        SecondChk = "Second"
                        'MsgBox("Excel was closed due to Issues.")

                        Msg = "Do you have any Excel spreadsheets open? if so please save and close them, Have you saved your work?"
                        Style = MsgBoxStyle.YesNo
                        Title = "Found Excel is open and need to make sure user saves files."
                        Response = MsgBox(Msg, Style, Title)

                        If Response = 6 Then
                            GoTo KillOpenFiles
                        Else
                            GoTo NextInstance
                        End If
                    Case "Word" & "First"
                        MsgBox("Word was closed due to Issues.")
                    Case "WinWord" & "First"
                        MsgBox("Word was closed due to Issues.")
                    Case "BulkBOM" & "First"
                        MsgBox("BulkBOM was closed due to Issues.")
                    Case "Adept" & "Second"                            '-----------------------Second Check
                        GenInfo3135.StartAdept = True
                    Case "acad" & "Second"
                    Case "Excel" & "Second"
                    Case "Word" & "First"
                        MsgBox("Word was closed due to Issues.")
                    Case "MatrixPrograms" & "First"
                        MsgBox("Matrix Programs was closed due to Issues.")
                    Case "Inventor" & "First"
                        SecondChk = "Second"
                        MsgBox("Inventor was closed due to Issues.")
                    Case "senddmp" & "First"
                        MsgBox("Inventor has had a Hard Crash, program will now try to recover.")
                        SecondChk = "Second"
                    Case "senddmp" & "Second"
                        MsgBox("Inventor has had a Hard Crash, program will now try to recover.")
                End Select
KillOpenFiles:
                instance.Kill()
                instance.CloseMainWindow()
                instance.Close()
NextInstance:
            Next instance
        End If

    End Function

    Private Sub PathBox_Click(sender As Object, e As EventArgs) Handles PathBox.Click
        'Test = Environment.CurrentDirectory
        'Test = Environment.SpecialFolder.Recent.GetName(8).ToString
        'Test = Environment.GetFolderPath(Environment.SpecialFolder.Recent)
        'Test = Envirnoment.UserDataPaths.Recent
        'Test = Environment.SpecialFolder.Recent
        'Test = System.Envirnoment.Path.GetTempPath

        'Test = System.IO.Path.GetDirectoryName

        'Get-PublicFolderClientPermission "\My Public Folder"
        'Get-PublicFolderClientPermission -Identity("\My Public Folder" -User Chris , Format-List)

        Me.PathBox.BackColor = System.Drawing.Color.White
        Me.BtnGetMWInfo.BackColor = System.Drawing.Color.White
        Me.CollectBOMList.BackColor = System.Drawing.Color.GreenYellow     '-------DJL-07-21-2025       'Me.DwgList.BackColor = System.Drawing.Color.GreenYellow

        BtnGetMWInfo.PerformClick()

        Me.Refresh()
    End Sub

    '    Private Sub PathBox2_Click(sender As Object, e As EventArgs) Handles PathBox2.Click
    '        Dim ErrHandler2 As New ErrHandler3135
    '        Dim i, k As Integer    ', j, k As Integer
    '        'Dim  TryYes As String
    '        Dim TestNew
    '        Dim BadPart, FileNam2, FindRev3, PrevRev As String
    '        Dim lReturn As Int64

    '        On Error GoTo Err_PathBox2_Click

    '        '-----------------------------------------------------------------------------------------
    '        ShipListBox.Items.Clear()
    '        Me.Refresh()
    '        StartDir = System.Environment.SpecialFolder.Recent

    '        If StartDir = "0" Then
    '            StartDir = Me.PathBox.Text
    '        End If

    '        OpenFileDialog1.InitialDirectory = StartDir
    '        OpenFileDialog1.Filter = "Excel Spreadsheets(*.xls)|*.xls;"
    '        OpenFileDialog1.Title = "Select file to Open"
    '        lReturn = OpenFileDialog1.ShowDialog()

    '        FileNam = OpenFileDialog1.FileName                          'Has to be 64 bit Inorder for this to work.
    '        FileNam2 = OpenFileDialog1.SafeFileName
    '        NewDir = FileNam.Replace(FileNam2, "")
    '        PathBox2.Text = NewDir

    '        Me.PathBox2.BackColor = System.Drawing.Color.White
    '        Me.ShipListBox.BackColor = System.Drawing.Color.GreenYellow

    '        ErrFound = ""
    '        Dim Dir1 As DirectoryInfo = New DirectoryInfo(NewDir)
    '        Dim DwgItem2

    '        If ShipListBox.Items.Count > 0 Then
    '            ShipListBox.Items.Clear()
    '        End If

    '        If ShipListBox.Items.Count = 0 Then
    '            For Each DwgItem2 In Dir1.GetFiles("*.xls")
    '                FindDwg = DwgItem2.ToString
    '                ExtPos = InStr(1, FindDwg, ".xls")

    '                If ExtPos > 0 Then
    '                    ShipListBox.Items.Add(DwgItem2)
    '                End If
    '            Next DwgItem2
    '        End If

    '        'btnNewDir.Enabled = True
    '        File2.Path = PathBox2.Text
    '        GroupBox5.Text = "Current Directory = " & PathBox2.Text

    '        If Me.File2.Items.Count > 0 Then
    'RestartFile2:
    '            FoundRev = 0
    '            CntSpreadSht = (Me.File2.Items.Count - 1)

    '            For i = 0 To CntSpreadSht
    '                BOMPos = 0

    '                If i > CntSpreadSht Then
    '                    GoTo Nexti
    '                End If

    '                FindRev = File2.Items.Item(i)
    '                RevPos = InStr(1, FindRev, "-R")
    '                BOMPos = InStr(1, FindRev, "ShippingList")

    '                If BOMPos > 0 Then
    '                    'ShipListBox.Items.Add(FindRev)                         '-------Added above.
    '                    FindRev = Mid(FindRev, (RevPos + 2), Len(FindRev))
    '                    FindRev1 = FindRev.Replace(".xls", "")

    '                    Dim pattern As String = "^[0-99]"
    '                    Dim matches As MatchCollection = Regex.Matches(FindRev1, pattern)
    '                    If Regex.IsMatch(FindRev1, pattern) Then
    '                        FindRev3 = FindRev1
    '                        'For k = 1 To Len(FindRev1)
    '                        'FindRev2 = Mid(FindRev1, k, 1)

    '                        'If Regex.IsMatch(FindRev2, pattern) Then
    '                        '    FindRev3 = FindRev3 & FindRev2
    '                        'End If

    '                        'Next k
    '                    End If

    '                    If FoundRev < FindRev3 Then
    '                        FoundRev = FindRev3
    '                    End If
    'UpdateBOMRev:
    '                End If
    'Nexti:
    '                If PrevRev < FoundRev Then
    '                    PrevRev = FoundRev
    '                    FindRev3 = Nothing
    '                End If
    '            Next i

    '            FoundRev = PrevRev
    '            Dim index As Integer
    '            index = ComboBox1.FindString(FoundRev + 1)
    '            ComboBox1.SelectedIndex = index
    '        End If

    '        If Me.ShipListBox.Items.Count = 1 Then
    '            ShipListBox.SelectedItem = Me.ShipListBox.Items.Item(0)
    '        End If

    '        ShipListBox.Sorted = True
    '        Me.Refresh()

    'Err_PathBox2_Click:
    '        ErrNo = Err.Number

    '        If ErrNo <> 0 Then
    '            PriPrg = "ShipListReadBOMAutoCAD"
    '            PrgName = "PathBox2_Click_Click"
    '            ErrMsg = Err.Description
    '            ErrSource = Err.Source
    '            ErrDll = Err.LastDllError
    '            ErrLastLineX = Err.Erl
    '            ErrException = Err.GetException

    '            If IsNothing(ManwayInfo3135.UserName) = True Then
    '                ManwayInfo3135.UserName = System.Environment.UserName
    '            End If

    '            Select Case ErrNo
    '                Case Is = 13 And InStr(ErrMsg, "Conversion from string") > 0
    '                    MsgBox("Program is having a time converting this item to String Error Msg = " & ErrMsg)
    '                    Stop
    '                    Resume
    '                Case Is = "91" And Mid(ErrMsg, 1, 24) = "Object reference not set"
    '                    Sapi.speak("Please open Primary Assembly in Inventor, and then pick OK on this box.")
    '                    MsgBox("Please open Primary Assembly in Inventor, and then pick OK on this box.")
    '                    ErrFound = "No Manway at Startup"
    '                    Resume Next
    '                Case Is = -2147023170 And InStr(ErrMsg, "The remote procedure call failed") <> 0
    '                    MsgBox("Part " & BadPart & " Needs to be replaced using content center, some how this file is corrupt.")
    '                    Stop
    '                    Resume
    '                Case Else
    '                    If ManwayInfo3135.UserName = "dlong" Then
    '                        MsgBox("Error Number " & ErrNo & ErrMsg & ErrSource)
    '                        Stop
    '                        Resume
    '                    Else
    '                        MsgBox("Program has new error, Error Msg = " & ErrMsg)

    '                        Dim st As New StackTrace(Err.GetException, True)
    '                        CntFrames = st.FrameCount
    '                        GetFramesSrt = st.GetFrames
    '                        PrgLineNo = GetFramesSrt(CntFrames - 1).ToString
    '                        PrgLineNo = PrgLineNo.Replace("@", "at")
    '                        LenPrgLineNo = (Len(PrgLineNo))
    '                        PrgLineNo = Mid(PrgLineNo, 1, (LenPrgLineNo - 2))

    '                        HandleErrSQL(PrgName, ErrNo, ErrMsg, ErrSource, PriPrg, ErrDll, DwgItem, PrgLineNo)
    '                    End If
    '            End Select
    '        End If

    'ExitCal:
    '    End Sub

    Private Sub BtnNewDirPathbox_Click(sender As Object, e As EventArgs) Handles BtnNewDirPathbox.Click  'Handles PathBox2.Click
        Dim ErrHandler2 As New ErrHandler3135
        Dim BadPart, FileNam2, FindRev3, PrevRev As String
        Dim lReturn As Int64

        On Error GoTo Err_PathBox2_Click

        ShipListBox.Items.Clear()
        Me.Refresh()
        StartDir = System.Environment.SpecialFolder.Recent

        If StartDir = "0" Then
            StartDir = Me.PathBox.Text
        End If

        OpenFileDialog1.InitialDirectory = StartDir
        OpenFileDialog1.Filter = "Excel Spreadsheets(*.xls)|*.xls;"
        OpenFileDialog1.Title = "Select file to Open"
        lReturn = OpenFileDialog1.ShowDialog()

        FileNam = OpenFileDialog1.FileName                          'Has to be 64 bit Inorder for this to work.
        FileNam2 = OpenFileDialog1.SafeFileName
        NewDir = FileNam.Replace(FileNam2, "")
        PathBox2.Text = NewDir

        Me.PathBox2.BackColor = System.Drawing.Color.White
        Me.ShipListBox.BackColor = System.Drawing.Color.GreenYellow

        ErrFound = ""
        Dim Dir1 As DirectoryInfo = New DirectoryInfo(NewDir)
        Dim DwgItem2

        If ShipListBox.Items.Count > 0 Then
            ShipListBox.Items.Clear()
        End If

        If ShipListBox.Items.Count = 0 Then
            For Each DwgItem2 In Dir1.GetFiles("*.xls")
                ExtPos = 0                          '-------DJL-07-15-2025
                ExtPos2 = 0                          '-------DJL-07-15-2025

                FindDwg = DwgItem2.ToString
                ExtPos = InStr(1, FindDwg, ".xls")
                ExtPos2 = InStr(1, FindDwg, "Shipping")                          '-------DJL-07-15-2025

                If ExtPos > 0 And ExtPos2 > 0 Then                          '-------DJL-07-15-2025
                    ShipListBox.Items.Add(DwgItem2)

                    FindRev = DwgItem2.ToString
                    RevPos = InStr(1, FindRev, "-R")
                    BOMPos = InStr(1, FindRev, "ShippingList")
                    FindRev = Mid(FindRev, (RevPos + 2), Len(FindRev))                          '-------DJL-07-15-2025
                    FindRev1 = FindRev.Replace(".xls", "")                          '-------DJL-07-15-2025

                    Dim pattern As String = "^[0-99]"                          '-------DJL-07-15-2025
                    Dim matches As MatchCollection = Regex.Matches(FindRev1, pattern)                          '-------DJL-07-15-2025
                    If Regex.IsMatch(FindRev1, pattern) Then                          '-------DJL-07-15-2025
                        FindRev3 = FindRev1
                    End If

                    If FoundRev < FindRev3 Then                          '-------DJL-07-15-2025
                        FoundRev = FindRev3
                    Else
                        If FoundRev = FindRev3 Then                          '-------DJL-07-15-2025
                            FoundRev = FindRev3
                        End If
                    End If
                End If
            Next DwgItem2
        End If

        'btnNewDir.Enabled = True
        File2.Path = PathBox2.Text
        'GroupBox5.Text = "Current Directory = " & PathBox2.Text

        Dim index As Integer
        index = ComboBox1.FindString(FoundRev + 1)
        ComboBox1.SelectedIndex = index

        If Me.ShipListBox.Items.Count = 1 Then
            ShipListBox.SelectedItem = Me.ShipListBox.Items.Item(0)
        End If

        ShipListBox.Sorted = True
        Me.Refresh()

Err_PathBox2_Click:
        ErrNo = Err.Number

        If ErrNo <> 0 Then
            PrgName = "BtnNewDirPathBox"
            PriPrg = "ShipListReadBOMAutoCAD"
            PrgName = "PathBox2_Click_Click"
            ErrMsg = Err.Description
            ErrSource = Err.Source
            ErrDll = Err.LastDllError
            ErrLastLineX = Err.Erl
            ErrException = Err.GetException

            If IsNothing(ManwayInfo3135.UserName) = True Then
                ManwayInfo3135.UserName = System.Environment.UserName
            End If

            Select Case ErrNo
                Case Is = 13 And InStr(ErrMsg, "Conversion from string") > 0
                    MsgBox("Program is having a time converting this item to String Error Msg = " & ErrMsg)
                    Stop
                    Resume
                Case Is = "91" And Mid(ErrMsg, 1, 24) = "Object reference not set"
                    Sapi.speak("Please open Primary Assembly in Inventor, and then pick OK on this box.")
                    MsgBox("Please open Primary Assembly in Inventor, and then pick OK on this box.")
                    ErrFound = "No Manway at Startup"
                    Resume Next
                Case Is = -2147023170 And InStr(ErrMsg, "The remote procedure call failed") <> 0
                    MsgBox("Part " & BadPart & " Needs to be replaced using content center, some how this file is corrupt.")
                    Stop
                    Resume
                Case Else
                    If ManwayInfo3135.UserName = "dlong" Then
                        MsgBox("Error Number " & ErrNo & ErrMsg & ErrSource)
                        Stop
                        Resume
                    Else
                        MsgBox("Program has new error, Error Msg = " & ErrMsg)

                        Dim st As New StackTrace(Err.GetException, True)
                        CntFrames = st.FrameCount
                        GetFramesSrt = st.GetFrames
                        PrgLineNo = GetFramesSrt(CntFrames - 1).ToString
                        PrgLineNo = PrgLineNo.Replace("@", "at")
                        LenPrgLineNo = (Len(PrgLineNo))
                        PrgLineNo = Mid(PrgLineNo, 1, (LenPrgLineNo - 2))

                        HandleErrSQL(PrgName, ErrNo, ErrMsg, ErrSource, PriPrg, ErrDll, DwgItem, PrgLineNo)
                    End If
            End Select
        End If

ExitCal:
    End Sub

    Private Sub BtnSpeedTest_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnSpeedTest.Click
        '----------------------------------------------------------------
        '06/05/2025 Adam had a reference using DBX per the Document it was faster found it is not 44:32 mins versa current system 19 mins from Tulsa '-------DJL-06-05-2025
        '---------------------------------------------------------------
        Dim BlockData(1) As Object
        Dim InsertionPT, VarSelArray, gvntSDIvar, Temparray, TempAttributes As Object
        Dim Title, Msg, Style, Response, ShipListItem, BOMBlkInsPoint As Object
        Dim TagInsPnt1 As Object, TagInsPnt2 As Object, xlMinimized As Object, ExcelWorkbook As Object
        Dim BlockSel As AutoCAD.AcadSelectionSet
        Dim BlockSelBOM As AutoCAD.AcadSelectionSet
        Dim DelBlockSel As AutoCAD.AcadEntity
        Dim AMWdisabled As Boolean, AcadOpen As Boolean, MultiLineMatl As Boolean
        Dim BatchYN As Short, GroupCode(1) As Short
        Dim AcadPref As AutoCAD.AcadPreferencesSystemClass
        Dim CompareX1 As Double, CompareX2 As Double, Dimscale As Double
        Dim JAt, i, j, k, l, i1, x, Count, Shopwatch, CntWorkbook, CntWrkShts, CountVal, CutCount, Testi, ChkCnt, CntBlks As Integer
        Dim SeeNotePos, ChkPos, SeeDwgPos, ExceptPos, CntExcept, RevPos, ExtPos, jA, FoundPart As Integer
        Dim FootPos, FootPos2, LenFoot, StrLineNo, CntBOM, CntDwgsNotFound, CntItems, LenDwg2 As Integer
        Dim FoundMM1, FoundMM2, SRListCnt, LineNo, CntItemsFound, NotePos As Integer
        Dim FullJobNo, CurrentDwgRev, CurrentDwgNo, FilePath, CurrentDWG, PrgName, WorkBookName, NextChk, Chk10D, Chk10E As String
        Dim WrkBookName, WorkShtName, NextSht, BomWrkShtNam, FileSaveAS, OldFileNam, UpdDesc, DwgNoNew2, DwgNoOld2 As String
        Dim GetDesc, GetMat, GetNotes, GetPartNo, GetQty, GetWeight, GetDesc3D, GetLen As String
        Dim NewDwg, NewRev, NewShpMk, NewPcMk, NewQty, NewDesc, NewInv, NewStd, NewMatl, NewWht As String
        Dim GetShipMk, GetPartNoAtt, NPath, TPath, BOMBlock, OldAttName, NewAttName, GetCity, GetState As String
        Dim OldRef, MxPcMk, Test10, Test11, GetAllParts, fixedCustomerPO, fixedPrevCustPO As String
        Dim FoundItem, SearchSeeNote, SearchNote, SearchNote2, SearchDwg, CompDesc, FirstTimeThru As String
        Dim test, Test1, pattern, ErrAt, PrevCustPO, BadBlksFound, BadBlksFound2, BadDwgFound As String
        Dim DwgNoOld, RevNoOld, PcMkOld, QtyOld, DescOld, TestOldDesc, GetReqNo, GetPrevPartNo, GetPrev2DShipMk As String
        Dim DwgNoNew, RevNoNew, PcMkNew, QtyNew, DescNew, TestNewDesc, JobNoDash, GetNextDesc, ProblemAt As String
        Dim Get2DShipMk, GetShipDesc, GetInv1, GetInv2, Get2DShipQty, GetMat2, Getmat3, GetWt, BOMBlkName As String
        Dim ScaleFactor, BOMBlkScale As Long
        Dim Workbooks As Excel.Workbooks
        Dim StickWrkSht As Excel.Workbook
        Dim ThisDrawing As AutoCAD.AcadDocument
        Dim OldShipListTest As String
        Dim BlockRefObjs As AutoCAD.AcadBlockReference
        Dim NewAttributes As Object
        Dim Pt0(2) As Double
        Dim Pt1(2) As Double
        Dim PTX, PTY, PTZ As Double
        Dim DwgItem3

        PrgName = "StartButton2_Click"                          '-------DJL-12-19-2024

        On Error GoTo Err_StartButton_Click

        TestOldDesc = Nothing
        FullJobNo = Nothing
        CurrentDwgNo = Nothing
        CurrentDwgRev = Nothing
        PrevCustPO = Nothing
        BOMBlock = Nothing
        DescOld = Nothing
        ExceptPos = 0
        Shopwatch = 0
        SortListing = True
        TPath = "K:\CAD\Blocks\Engineering Tulsa\"
        OldDwgItem = Nothing
        BadDwgFound = "No"
        GetAllParts = "No"
        Me.BtnAdd.BackColor = System.Drawing.Color.White
        Me.BtnClear.BackColor = System.Drawing.Color.White
        Me.BtnRemove.BackColor = System.Drawing.Color.White
        Me.TxtBoxCntDown.BackColor = System.Drawing.Color.LawnGreen
        PrgName = "StartButton2-CollectingDATA"                          '-------DJL-12-19-2024

        'Turned Compare process back on per OU 212 PM request                          'DJL-12-25-2023
        If Me.ComboBox1.Items.Count > 0 Then
            If InStr(1, Me.ShipListBox.Text, ".xls") = 0 Then
                Msg = "No Shipping List was selected to compare to, do you want to continue?"
                Style = MsgBoxStyle.YesNo
                Title = "Bulk BOM"
                Response = MsgBox(Msg, Style, Title)
                If Response = 6 Then                    'if user clicks yes
                    Me.File2.Enabled = False
                    Me.CheckBox1.CheckState = False
                Else 'if user clicks no
                    Me.File2.Enabled = True
                    Me.CheckBox1.CheckState = CheckState.Checked
                    Exit Sub
                End If
            End If
        End If

        PrgName = "StartButton2-OpenExcel"                          '-------DJL-12-19-2024

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

        On Error GoTo Err_StartButton_Click

        If Me.ComboBox1.Text = vbNullString Then
            MsgBox("Please select a revision number for Shipping List")
            Exit Sub
        End If

        PrgName = "StartButton2-OpenAutoCAD"                          '-------DJL-12-19-2024
        Me.Label2.Text = "Gathering Information from AutoCAD Drawings........Please Wait"
        Me.Refresh()

        On Error Resume Next
        If Err.Number Then
            Err.Clear()
        End If

        AcadApp = GetObject(, "AutoCAD.Application")
        Threading.Thread.Sleep(25)
        AcadOpen = True

        If Err.Number Then
            Information.Err.Clear()
            AcadApp = CreateObject("AutoCAD.Application")
            Threading.Thread.Sleep(25)
            AcadOpen = False            'True                           'DJL-------06/02/2025
            If Err.Number Then
                Information.Err.Clear()
                AcadApp = GetObject(, "AutoCAD.Application")
                Threading.Thread.Sleep(25)
                If Err.Number Then
                    Information.Err.Clear()
                    AcadApp = CreateObject("AutoCAD.Application")
                    Threading.Thread.Sleep(25)
                    If Err.Number Then
                        MsgBox("This Program is having problems opening AutoCAD, Please open AutoCAD then pick this Ok Button")
                        AcadApp = GetObject(, "AutoCAD.Application")
                        AcadApp.Visible = False
                        MsgBox("Now running " & AcadApp.Name & " version " & AcadApp.Version)
                    End If
                End If
            End If
        End If

        If IsNothing(AcadApp) = True Then
            AcadApp = GetObject(, "AutoCAD.Application")
            Threading.Thread.Sleep(25)
        End If

        AcadApp.Visible = False                         'True                           'DJL-------06/02/2025
        ThisDrawing = AcadApp.ActiveDocument

        If IsNothing(ThisDrawing) = True Then
            Threading.Thread.Sleep(250)
            ThisDrawing = AcadApp.ActiveDocument            '-------For some reason AutoCAD is not setting ThisDrawing the first time ----DJL-------11-16-2023
        End If

        PrgName = "StartButton2-AutoCADSettings"                          '-------DJL-12-19-2024

        On Error GoTo Err_StartButton_Click

        AcadPref = AcadApp.Preferences.System

        If AcadPref.SingleDocumentMode = True Then
            gvntSDIvar = AcadPref.SingleDocumentMode
        End If

        If gvntSDIvar = True Then
            AcadPref.SingleDocumentMode = False
        End If

        ErrAt = "Unload Adept"
        AcadApp.ActiveDocument.SendCommand(("unloadadept "))
        ErrAt = ""

        AcadApp.WindowState = AutoCAD.AcWindowState.acMin                           'DJL-------06/02/2025

        '-------Only needed in one place below
        ''-------------Looks like layout.dvb is creating a Bad Layout when autocad acts up.
        PrgName = "StartButton2-Layout-dvb"                          '-------DJL-12-19-2024
        'AcadApp.UnloadDVB("K:\cad\vba\autocad\Layout.dvb")
        AcadApp.Visible = False                        'True                           'DJL-------06/02/2025
        Me.Refresh()
        SBclicked = True
        CBclicked = False
        Count = Me.SelectList.Items.Count
        ProgressBar1.Value = 0
        ProgressBar1.Maximum = Count
        ProgressBar1.Visible = False                        'True                           'DJL-------06/02/2025
        CountVal = 0

        If Count <> 0 Then
            ReDim VarSelArray(Count)
            For i = 1 To Count
                VarSelArray(i) = Me.SelectList.Items.Item(i - 1)
            Next i

            PrgName = "StartButton2-CollectTags"                          '-------DJL-12-19-2024

            For Each DwgItem3 In VarSelArray
                CntItemsFound = 0
                TitleBlkName = Nothing
                CurrentDwgNo = Nothing
                FullJobNo = Nothing
                CurrentDwgRev = Nothing
                CustomerPO = Nothing
                Cust1 = Nothing
                Cust2 = Nothing
                GetCity = Nothing
                GetState = Nothing
                Get2DShipMk = Nothing
                GetPrev2DShipMk = Nothing
                Get2DShipQty = Nothing
                GetPartNo = Nothing
                GetPrevPartNo = Nothing
                GetQty = Nothing
                GetShipDesc = Nothing
                GetDesc = Nothing
                GetShipDesc = Nothing
                GetInv1 = Nothing
                GetInv2 = Nothing
                GetLen = Nothing
                GetMat = Nothing
                GetNotes = Nothing
                GetMat2 = Nothing
                Getmat3 = Nothing
                GetNotes = Nothing
                GetWt = Nothing
                GetReqNo = Nothing
                PrevCustPO = Nothing

                If CBclicked Then
                    If MsgBox("Are you sure?", 36, "Bulk BOM Generator - Cancel") = 6 Then
                        GoTo Cancel
                    End If
                Else
                    If IsNothing(DwgItem3) = True Then                          'If DwgItem3 = Nothing Then
                        GoTo NextDwg
                    End If

ReOpenAcad:
                    TxtBoxCntDown.Text = Count - CountVal

                    Dim dbxDoc As Object = AcadApp.GetInterfaceObject("ObjectDBX.AxDbDocument.24")
                    dbxDoc.Open(PathBox.Text & DwgItem3.ToString, "")                    'dbxDoc.Open(dwgPath, "")
                    Dim ms As Object = dbxDoc.paperspace                                    'Dim ms As Object = dbxDoc.ModelSpace

                    'AcadApp.Documents.Open(PathBox.Text & DwgItem3.ToString)
                    Threading.Thread.Sleep(25)

                    If BadDwgFound = "Yes" Then
                        BadDwgFound = "No"
                        GoTo NextDwg
                    End If

                    If IsNothing(OldDwgItem) = False Then
                        For Each Dwg2 In AcadApp.Documents

                            If InStr(OldDwgItem, DwgName) = 0 Or DwgName = Nothing Then
                                DwgName = OldDwgItem
                                SearchSlash = "\"
                                SearchPos = InStr(DwgName, SearchSlash)

                                While SearchPos > 0
                                    SecondPart = Mid(DwgName, (SearchPos + 1), (Len(DwgName) - SearchPos))
                                    DwgName = SecondPart
                                    SearchPos = InStr(DwgName, SearchSlash)
                                End While
                            End If

                            If Dwg2.Name = DwgName Then
                                Dwg2.Close(False)
                                GoTo Start
                            End If
                        Next
                    End If
Start:
                    For Each ent As Object In ms
                        If ent.ObjectName = "AcDbBlockReference" Then
                            Dim blkRef As Object = ent
                            If blkRef.Name = "MX_TITLE*" Or blkRef.Name = "MX_TITLE-11x17" Then                        'DJL-------06-02-2025
                                TitleBlkName = blkRef.Name                        'DJL-------06-03-2025
                                Dim atts As Object = blkRef.GetAttributes()
                                For Each att As Object In atts
Start2:
                                    ProblemAt = "ReadingTitleBlock"

                                    GetAllParts = "No"
                                    AcadApp.Visible = False
                                    Me.Refresh()

                                    If InStr(TitleBlkName, "MPDM_STD_TITLE_") = 0 Then
FindAttTagString:
                                        Select Case att.TagString                          '-------DJL-06-02-2025      'Select Case Temparray(i).TagString
                                            Case "DN"
                                                CurrentDwgNo = att.TextString               'Temparray(i).TextString
                                                CntItemsFound = (CntItemsFound + 1)
                                            Case "JN"
                                                FullJobNo = att.TextString               'Temparray(i).TextString
                                                If GenInfo3135.FullJobNo = Nothing Then
                                                    GenInfo3135.FullJobNo = FullJobNo
                                                End If
                                                CntItemsFound = (CntItemsFound + 1)
                                            Case "RN"
                                                CurrentDwgRev = att.TextString               'Temparray(i).TextString
                                                CntItemsFound = (CntItemsFound + 1)
                                            Case "C"
                                                CustomerPO = att.TextString               'Temparray(i).TextString

                                                If GenInfo3135.CustomerPO = Nothing Then
                                                    GenInfo3135.CustomerPO = CustomerPO
                                                End If

                                                CntItemsFound = (CntItemsFound + 1)
                                            Case "C1"
                                                Cust1 = att.TextString               'Temparray(i).TextString
                                                CntItemsFound = (CntItemsFound + 1)
                                            Case "C2"
                                                Cust2 = att.TextString               'Temparray(i).TextString
                                                CntItemsFound = (CntItemsFound + 1)
                                        End Select

                                        If CntItemsFound = 6 Then       'There is no need to look at all 31 items if what is needed has been found.-------DJL-------10-31-2023
                                            GoTo FoundItemsNeeded
                                        Else
                                            If CntItemsFound = 4 And Cust2 = Nothing Then                         '-------DJL-------10-31-2023
                                                GoTo FoundItemsNeeded
                                            End If
                                        End If
                                    Else
                                        For i = 0 To UBound(Temparray)           '-------Designed for the Newer Title Blocks-------DJL-12-18-2023
                                            test = Temparray(i).TagString        '-------Items went from 31 to 81 ask Pittsburgh to change the order items 
                                            Test1 = Temparray(i).TextString      '-------are in the list so that the programs will run faster-DJL-12-18-2023

                                            Select Case Temparray(i).TagString
                                                Case "DN"                           '-------Drawing No.
                                                    CurrentDwgNo = att.TextString               'Temparray(i).TextString
                                                    CntItemsFound = (CntItemsFound + 1)
                                                Case "PN"                           '-------"JN"-------Job Number.
                                                    FullJobNo = att.TextString               'Temparray(i).TextString
                                                    If GenInfo3135.FullJobNo = Nothing Then
                                                        GenInfo3135.FullJobNo = FullJobNo
                                                    End If
                                                    CntItemsFound = (CntItemsFound + 1)
                                                Case "REV"                          '-------"RN"-------Revision Number
                                                    CurrentDwgRev = att.TextString               'Temparray(i).TextString
                                                    CntItemsFound = (CntItemsFound + 1)
                                                Case "C1"                           '-------
                                                    Cust1 = att.TextString               'Temparray(i).TextString
                                                    CntItemsFound = (CntItemsFound + 1)
                                                Case "C2"
                                                    Cust2 = att.TextString               'Temparray(i).TextString
                                                    CntItemsFound = (CntItemsFound + 1)
                                                Case "CY"
                                                    GetCity = att.TextString               'Temparray(i).TextString
                                                    CntItemsFound = (CntItemsFound + 1)
                                                Case "ST"
                                                    GetState = att.TextString               'Temparray(i).TextString
                                                    CntItemsFound = (CntItemsFound + 1)
                                            End Select

                                            If CntItemsFound = 7 Then       'There is no need to look at all 31 items if what is needed has been found.-------DJL-------10-31-2023
                                                GoTo FoundItemsNeeded       'Pittsburgh forces us to look at all of them because the last 4 are found at the end.
                                                'Else
                                                '    If CntItemsFound = 4 And Cust2 = Nothing Then                         '-------DJL-------10-31-2023
                                                '        GoTo FoundItemsNeeded
                                                '    End If
                                            End If
                                            '-------Why look at every item?                             '-------DJL-10-9-2023
                                        Next i

                                    End If
                                Next att                            '-------DJL-06-03-2025
FoundItemsNeeded:
                                If CustomerPO <> PrevCustPO And CustomerPO <> "" Then               '-------Tulsa Old Titleblocks.
                                    If IsNothing(PrevCustPO) = True Then
                                        PrevCustPO = CustomerPO
                                    Else
                                        If IsNothing(PrevCustPO) = True Then
                                            CustomerPO = InputBox("This job has two customer names please type In the correct Customer Name ? " & CustomerPO & " Or " & PrevCustPO)
                                            CustomerPO = UCase(CustomerPO)
                                            PrevCustPO = CustomerPO
                                        Else
                                            If CustomerPO <> PrevCustPO And PrevCustPO <> Nothing Then
                                                fixedCustomerPO = CustomerPO.Replace(" ", "")           'USERS MUST HAVE EXTRA SPACEZ ON DESCRIPTION.
                                                fixedPrevCustPO = PrevCustPO.Replace(" ", "")
                                            End If

                                            If fixedCustomerPO <> fixedPrevCustPO Then
                                                If InStr(TitleBlkName, "MPDM_STD_TITLE_") > 0 Then
                                                    GoTo FoundMPDMSTDTITLE
                                                Else
                                                    CustomerPO = PrevCustPO
                                                End If
                                            End If
                                        End If
                                    End If

FoundMPDMSTDTITLE:
                                End If

                                If InStr(TitleBlkName, "MPDM_STD_TITLE_") = 0 Then
                                    If CustomerPO = "" Then
                                        CustomerPO = Cust1 & " " & Cust2
                                    End If
                                Else
                                    If CustomerPO = "" Then
                                        CustomerPO = Cust1 & " " & FullJobNo
                                    End If
                                End If

                                CurrentDwgNo = CurrentDwgNo
                                FullJobNo = FullJobNo
                                CurrentDwgRev = CurrentDwgRev

                            End If                          '-------DJL-06-03-2025
                        End If                              '-------DJL-06-03-2025
                    Next ent                                   '-------DJL-06-03-2025

                    For Each ent As Object In ms
                        If ent.ObjectName = "AcDbBlockReference" Then
                            Dim blkRef As Object = ent
                            If blkRef.Name = "STANDARD_BILL_OF_MATERIAL" Then                        'DJL-------06-02-2025
                                BOMBlkName = blkRef.Name                        'DJL-------06-03-2025
                                BOMBlkInsPoint = blkRef.InsertionPoint
                                BOMBlkScale = blkRef.XScaleFactor
                                Dim atts As Object = blkRef.GetAttributes()
                                GetPartNo = Nothing
                                Get2DShipMk = Nothing
                                Get2DShipQty = Nothing
                                GetPartNo = Nothing
                                GetQty = Nothing
                                GetShipDesc = Nothing
                                GetDesc = Nothing
                                GetShipDesc = Nothing
                                GetInv1 = Nothing
                                GetInv2 = Nothing
                                GetLen = Nothing
                                GetMat2 = Nothing
                                Getmat3 = Nothing
                                GetNotes = Nothing
                                GetWt = Nothing
                                GetReqNo = Nothing

                                For Each att As Object In atts
                                    '-------------------------------------------------------------------------------------------------
                                    '-------Problem in code below items are collected by the way they were edited, so this will not work when Plate material
                                    '       for SR1 thru SR10 must be renamed. So we need to collect everything then after spreadsheet is filled out sort
                                    '       out what must be deleted.
                                    '-----------------------------------------------------------------------------------------------

                                    Select Case att.Tagstring                           '-------DJL-06-03-2025      'TempAttributes(x).TagString
                                        Case "SLM"                                          '-------Ship Mark"
                                            Get2DShipMk = att.TextString                    '-------DJL-06-03-2025      'TempAttributes(x).TextString
                                        Case "SLQ"                                          '-------Ship Mark Qty"
                                            Get2DShipQty = att.TextString                    '-------DJL-06-03-2025      'TempAttributes(x).TextString
                                        Case "SM"                                           '-------Ship Part"
                                            GetPartNo = att.TextString                    '-------DJL-06-03-2025      'TempAttributes(x).TextString
                                        Case "Q"                                            '-------Ship Part Qty"
                                            GetQty = att.TextString                    '-------DJL-06-03-2025      'TempAttributes(x).TextString
                                        Case "SD"
                                            GetShipDesc = att.TextString                    '-------DJL-06-03-2025      'TempAttributes(x).TextString
                                        Case "D"
                                            GetDesc = att.TextString                    '-------DJL-06-03-2025      'TempAttributes(x).TextString
                                        Case "D2"
                                            GetShipDesc = att.TextString                    '-------DJL-06-03-2025      'TempAttributes(x).TextString
                                        Case "IU"
                                            GetInv1 = att.TextString                    '-------DJL-06-03-2025      'TempAttributes(x).TextString
                                        Case "IL"
                                            GetInv2 = att.TextString                    '-------DJL-06-03-2025      'TempAttributes(x).TextString
                                        Case "L"
                                            GetLen = att.TextString                    '-------DJL-06-03-2025      'TempAttributes(x).TextString
                                        Case "M"
                                            GetMat = att.TextString                    '-------DJL-06-03-2025      'TempAttributes(x).TextString

                                            If InStr(GetMat, "NOTE") > 0 And NotePos > 1 Then
                                                NotePos = InStr(GetMat, "NOTE")
                                                GetNotes = Mid(GetMat, NotePos, Len(GetMat))
                                                GetMat = Mid(GetMat, 1, (NotePos - 1))
                                            Else
                                                NotePos = InStr(GetMat, "NOTE")

                                                If NotePos > 0 Then
                                                    GetNotes = Mid(GetMat, NotePos, Len(GetMat))
                                                    'GetMat = GetMat
                                                End If
                                            End If

                                        Case "M2"
                                            GetMat2 = att.TextString                    '-------DJL-06-03-2025      'TempAttributes(x).TextString
                                        Case "M3"
                                            Getmat3 = att.TextString                    '-------DJL-06-03-2025      'TempAttributes(x).TextString
                                        Case "N"
                                            GetNotes = att.TextString                    '-------DJL-06-03-2025      'TempAttributes(x).TextString
                                        Case "W"
                                            GetWt = att.TextString                    '-------DJL-06-03-2025      'TempAttributes(x).TextString
                                        Case "P"
                                            GetReqNo = att.TextString                    '-------DJL-06-03-2025      'TempAttributes(x).TextString            'Requestion number
                                    End Select

                                    If GetPrevPartNo <> GetPartNo Then                    '-------DJL-06-04-2025
                                        If IsNothing(GetPartNo) = False And GetPartNo <> "" Then                    '-------DJL-06-04-2025
                                            GetPrevPartNo = GetPartNo                    '-------DJL-06-04-2025
                                        End If
                                    End If

                                    If GetPrev2DShipMk <> Get2DShipMk Then                    '-------DJL-06-04-2025
                                        If IsNothing(Get2DShipMk) = False And Get2DShipMk <> "" Then                    '-------DJL-06-04-2025
                                            GetPrev2DShipMk = Get2DShipMk                    '-------DJL-06-04-2025
                                        End If
                                    End If

                                Next                 '-------DJL-06-03-2025

                                '----------------------------------------------------------------------------------------------------------------
                                '-------Program is not collecting each shipping list item.
                                '----------------------------------------------------------------------------------------------------------------

                                If GetPrev2DShipMk <> "" And IsNothing(GetPrev2DShipMk) = False Then
                                    If GetPartNo = "" And Get2DShipMk = "" Then      '-------DJL-06-03-2025 
                                        GetPartNo = GetPrev2DShipMk
                                    End If
                                End If


                                If GetPartNo <> "" And Get2DShipMk = "" Then      '-------DJL-06-03-2025      'If TempAttributes(x).TagString = "SM" And GetPartNo <> "" Then
                                    'GetShipMk = GetPartNo                    '-------DJL-06-04-2025
                                    GoTo FoundSR
                                End If


                                If Get2DShipMk <> "" Then      '-------DJL-06-03-2025      'If TempAttributes(x).TagString = "SLM" And Get2DShipMk <> "" Then
                                    GetPartNo = Get2DShipMk
FoundSR:
                                    GetPartNo = LTrim(GetPartNo)
                                    GetPartNo = RTrim(GetPartNo)

                                    If InStr(GetPartNo, "SR") = 1 Then
                                        Select Case 0
                                            Case Is < InStr(GetPartNo, "SR1")
                                                GetAllParts = "Yes"
                                                GoTo FoundSRPlates
                                            Case Is < InStr(GetPartNo, "SR2")
                                                GetAllParts = "Yes"
                                                GoTo FoundSRPlates
                                            Case Is < InStr(GetPartNo, "SR3")
                                                GetAllParts = "Yes"
                                                GoTo FoundSRPlates
                                            Case Is < InStr(GetPartNo, "SR4")
                                                GetAllParts = "Yes"
                                                GoTo FoundSRPlates
                                            Case Is < InStr(GetPartNo, "SR5")
                                                GetAllParts = "Yes"
                                                GoTo FoundSRPlates
                                            Case Is < InStr(GetPartNo, "SR6")
                                                GetAllParts = "Yes"
                                                GoTo FoundSRPlates
                                            Case Is < InStr(GetPartNo, "SR7")
                                                GetAllParts = "Yes"
                                                GoTo FoundSRPlates
                                            Case Is < InStr(GetPartNo, "SR8")
                                                GetAllParts = "Yes"
                                                GoTo FoundSRPlates
                                            Case Is < InStr(GetPartNo, "SR9")
                                                GetAllParts = "Yes"
                                                GoTo FoundSRPlates
                                            Case Is < InStr(GetPartNo, "SR10")
                                                GetAllParts = "Yes"
                                                GoTo FoundSRPlates
                                        End Select
                                    End If
                                End If

                                'Next x

FoundSRPlates:
                                PrgName = "StartButton2-UpdateArray"                          '-------DJL-12-19-2024
                                ShippingList(1, UBound(ShippingList, 2)) = CurrentDwgNo     '-------Moved up from below
                                ShippingList(2, UBound(ShippingList, 2)) = CurrentDwgRev

                                Get2DShipMk = LTrim(Get2DShipMk)
                                Get2DShipMk = RTrim(Get2DShipMk)
                                GetPrev2DShipMk = LTrim(GetPrev2DShipMk)      '-------DJL-06-04-2025
                                GetPrev2DShipMk = RTrim(GetPrev2DShipMk)      '-------DJL-06-04-2025

                                If Get2DShipMk = Nothing Or IsNothing(Get2DShipMk) = True Then      '-------DJL-06-04-2025
                                    ShippingList(3, UBound(ShippingList, 2)) = GetPrev2DShipMk          '-------DJL-06-04-2025
                                Else
                                    ShippingList(3, UBound(ShippingList, 2)) = Get2DShipMk          'GetPartNo
                                End If

                                ShippingList(4, UBound(ShippingList, 2)) = Get2DShipQty         'ship GetQty       'Qty

                                GetPartNo = LTrim(GetPartNo)
                                GetPartNo = RTrim(GetPartNo)
                                ShippingList(5, UBound(ShippingList, 2)) = GetPartNo          'GetPartNo
                                ShippingList(6, UBound(ShippingList, 2)) = GetQty       'Qty
                                ShippingList(7, UBound(ShippingList, 2)) = GetShipDesc
                                ShippingList(8, UBound(ShippingList, 2)) = GetDesc
                                ShippingList(9, UBound(ShippingList, 2)) = GetInv1
                                ShippingList(10, UBound(ShippingList, 2)) = GetInv2
                                ShippingList(11, UBound(ShippingList, 2)) = GetMat
                                ShippingList(12, UBound(ShippingList, 2)) = GetMat2
                                ShippingList(13, UBound(ShippingList, 2)) = Getmat3
                                ShippingList(14, UBound(ShippingList, 2)) = GetWt
                                ShippingList(15, UBound(ShippingList, 2)) = CurrentDwgNo

                                ShippingList(18, UBound(ShippingList, 2)) = GetReqNo

                                ScaleFactor = BOMBlkScale     'ScaleFactor = ShipListItem.XScaleFactor        '-------DJL-06-03-0225

                                Pt1 = BOMBlkInsPoint          'Pt1 = ShipListItem.InsertionPoint       '-------DJL-06-03-0225

                                '-------New requirements demand that the shipping list be reformatted so that "SHELL PLATE" reads.
                                '-------Example: "SHELL PLATE SR1 PL 0.859 x 95 x 33'-8" NOTE-4.5,4.8, - A36 G2"
                                '-------Create new Module when SR is found collect all parts in different Array.

                                If GetAllParts = "Yes" Then
                                    If GetPartNo = "" And InStr(GetDesc, "PL ") > 0 Then
                                        GoTo GetAllPlates
                                    End If
                                End If

                                '-------------------------------------------------------------------------------------------------------------------------------
                                '-------DJL-------10-31-2023
                                '-------Need to fix this mess below so that it is looking for the Attributes, SLM, SLQ, SM, Q, D, D2, IU, IL, M, M2, M3, W, etc.
                                '-------This way if anyone messes with the Attrubute structure or the order EVERYTHING will still be correct.

                                If Mid(GetPartNo, 1, 1) Like "[A-Za-z0-9#]" Then
GetAllPlates:
                                    Select Case TitleBlkName           'Select Case ShipListItem.Name
                                        Case "SUBTITLE"
                                            SortListing = False
                                            ShippingList(3, UBound(ShippingList, 2)) = GetPartNo
                                            ShippingList(4, UBound(ShippingList, 2)) = TempAttributes(1).TextString
                                            ShippingList(11, UBound(ShippingList, 2)) = TempAttributes(4).TextString
                                            ShippingList(7, UBound(ShippingList, 2)) = TempAttributes(5).TextString
                                            ShippingList(10, UBound(ShippingList, 2)) = TempAttributes(7).TextString
                                            ShippingList(14, UBound(ShippingList, 2)) = TempAttributes(8).TextString
                                        Case "A_BILL_OF_MATERIAL" 'works on A_BILL_OF_MATERIAL
                                            For i = 3 To (UBound(TempAttributes) + 4)
                                                Select Case i
                                                    Case 3 To 8
                                                        If i = 3 Then
                                                            ShippingList(i, UBound(ShippingList, 2)) = GetPartNo
                                                        Else
                                                            ShippingList(i, UBound(ShippingList, 2)) = TempAttributes(i - 3).TextString
                                                        End If
                                                    Case 9
                                                        If UBound(TempAttributes) = 10 Then
                                                            ShippingList(9, UBound(ShippingList, 2)) = TempAttributes(6).TextString
                                                            i = 10
                                                        End If
                                                    Case 10 To 14
                                                        If UBound(TempAttributes) = 10 Then
                                                            ShippingList(i, UBound(ShippingList, 2)) = TempAttributes(i - 4).TextString
                                                        Else
                                                            ShippingList(i, UBound(ShippingList, 2)) = TempAttributes(i - 5).TextString
                                                        End If
                                                End Select
                                            Next i
                                    End Select
                                End If

AutoCAD3D:
                                PrgName = "StartButton2-XY-Cor"                          '-------DJL-12-19-2024
                                InsertionPT = BOMBlkInsPoint                          '-------DJL-12-19-2024        'ShipListItem.InsertionPoint
                                Dimscale = BOMBlkScale                           '-------DJL-12-19-2024        'ShipListItem.XScaleFactor
                                CompareX1 = 10.5 * Dimscale
                                CompareX1 = InsertionPT(0) - CompareX1
                                CompareX1 = CompareX1 / Dimscale
                                CompareX2 = 6 * Dimscale
                                CompareX2 = InsertionPT(0) - CompareX2
                                CompareX2 = CompareX2 / Dimscale

                                If CompareX1 < 1 Or CompareX2 < 1 Then
                                    If CompareX1 > 0 Or CompareX2 > 0 Then
                                        ShippingList(16, UBound(ShippingList, 2)) = CStr(1)
                                    Else
                                        ShippingList(16, UBound(ShippingList, 2)) = CStr(2)
                                    End If
                                Else
                                    ShippingList(16, UBound(ShippingList, 2)) = CStr(2)
                                End If

                                ShippingList(17, UBound(ShippingList, 2)) = InsertionPT(1)
                                ReDim Preserve ShippingList(18, UBound(ShippingList, 2) + 1)

                                CntBOM = (CntBOM + 1)
                            End If         '-------DJL-06-03-2025
                        End If             '-------DJL-06-03-2025



                    Next    '-------DJL-06-03-2025        'ent       'For Each ent As Object In ms       'DJL-------06-02-2025
                End If

                CountVal = (CountVal + 1)

                ProgressBar1.Value = CountVal
                OldDwgItem = DwgItem3.ToString
NextDwg:
            Next DwgItem3

            If Not IsNothing(gvntSDIvar) Then
                AcadPref.SingleDocumentMode = gvntSDIvar
            End If

            ProgressBar1.Value = 0              '------------Export Shipping List info to Ship List now.
            Me.Label2.Text = "Outputting Information To Shipping List........Please Wait."
            Me.Refresh()
        End If                          '-------Moved here from after WriteToExcel

        '-----------------------------------------------------------------------------------------------------------------------------
        PrgName = "StartButton2-WriteToExcel"                          '-------DJL-12-19-2024
        WriteToExcel(ShippingList)
        ShippingList = GenInfo3135.ShippingList
        ProgressBar1.Value = 0

        RevNo = Me.ComboBox1.Text
        RevNo2 = RevNo
        WorkBookName = MainBOMFile.Application.ActiveWorkbook.Name
        OldFileNam = PathBox.Text

        '-------DJL--------11-28-2023---------------------------------------------------------------Do not see where this "BELOW" is doing anything?
        '''''''-------Look at removinf this part.
        CntItems = (UBound(ShippingList, 2) - 1)
        StrLineNo = GenInfo3135.StrLineNo
        PrgName = "StartButton2-ReadSpreadSht"                          '-------DJL-12-19-2024

        With ShipListSht
            For j = 1 To CntItems
                If StrLineNo = 42 Then
                    RowNo = (42 + j)
                Else
                    RowNo = (44 + j)
                End If

                If j = 1 Then
                    DwgNoOld = .Range("D" & RowNo).Value              'Dwg     
                    RevNoOld = .Range("E" & RowNo).Value              'Rev      
                    PcMkOld = .Range("F" & RowNo).Value              'Pc Mark  
                    QtyOld = .Range("G" & RowNo).Value              'Qty      
                    DescOld = .Range("H" & RowNo).Value             'Desc      
                    FoundMM1 = InStr(DescOld, "MM")

                    TestOldDesc = DwgNoOld & RevNoOld & PcMkOld & QtyOld
                Else
                    DwgNoNew = .Range("D" & RowNo).Value               'Dwg 
                    RevNoNew = .Range("E" & RowNo).Value              'Rev     
                    PcMkNew = .Range("F" & RowNo).Value              'Pc Mark   
                    QtyNew = .Range("G" & RowNo).Value          'Qty           
                    DescNew = .Range("H" & RowNo).Value          'Desc    
                    FoundMM2 = InStr(DescNew, "MM")

                    TestNewDesc = DwgNoNew & RevNoNew & PcMkNew & QtyNew

                    TestOldDesc = TestOldDesc

                    If TestOldDesc = TestNewDesc And TestOldDesc <> "" Then
                        If DescOld = DescNew Then           'Do not delete duplicates they are allowed on same drawing
                            DwgNoOld = DwgNoNew              'Dwg
                            RevNoOld = RevNoNew              'Rev
                            PcMkOld = PcMkNew              'Pc Mark
                            QtyOld = QtyNew              'Qty
                            DescOld = DescNew
                            TestOldDesc = DwgNoOld & RevNoOld & PcMkOld & QtyOld
                            j = (j - 1)
                            CntItems = (CntItems - 1)
                        Else
                            If FoundMM1 > 0 Then
                                .Range("A" & (RowNo - 1) & ":" & "T" & (RowNo - 1)).Delete(Shift:=XlDeleteShiftDirection.xlShiftUp)
                                '-------DJL will need to create a delete for Array.
                                DwgNoOld = DwgNoNew              'Dwg
                                RevNoOld = RevNoNew              'Rev
                                PcMkOld = PcMkNew              'Pc Mark
                                QtyOld = QtyNew              'Qty
                                DescOld = DescNew
                                TestOldDesc = DwgNoOld & RevNoOld & PcMkOld & QtyOld
                                j = (j - 1)
                                CntItems = (CntItems - 1)
                                If FoundMM2 > 0 Then
                                    .Range("A" & (RowNo - 1) & ":" & "T" & (RowNo - 1)).Delete(Shift:=XlDeleteShiftDirection.xlShiftUp)
                                    '-------DJL will need to create a delete for Array.
                                    DwgNoOld = DwgNoNew              'Dwg
                                    RevNoOld = RevNoNew              'Rev
                                    PcMkOld = PcMkNew              'Pc Mark
                                    QtyOld = QtyNew              'Qty
                                    DescOld = DescNew
                                    TestOldDesc = DwgNoOld & RevNoOld & PcMkOld & QtyOld
                                    j = (j - 1)
                                    CntItems = (CntItems - 1)
                                End If
                            End If
                        End If
                    Else
                        DwgNoOld = DwgNoNew              'Dwg
                        RevNoOld = RevNoNew              'Rev
                        PcMkOld = PcMkNew              'Pc Mark
                        QtyOld = QtyNew              'Qty
                        DescOld = DescNew
                        TestOldDesc = DwgNoOld & RevNoOld & PcMkOld & QtyOld
                    End If
                End If
                If j >= CntItems Then
                    GoTo EndProcess
                End If

                Me.ProgressBar1.Value = RowNo
            Next j
        End With
EndProcess:


        '-------DJL-------10-31-2023----------------------------------------------------------------Move to function CreateSpSht
        '------------------------------------------------Need to remove extra Line items that are referenced two times.
        '-------221-22-00103 and 221-22-00102 and 221-22-00107 has duplicated items on drawings 33A and 33B.
        '-------30" SHELL MIXER MANWAY
        '-------30" SHELL MIXER MANWAY REPAD
        '-------N14 and N14R

        'CntItems = CountBOM                            '-------Done above.
        PrgName = "StartButton2-RemoveDup"                          '-------DJL-12-19-2024
        Me.ProgressBar1.Maximum = CntItems              'CountBOM
        Label2.Text = "Remove Duplicated Items"         'Example Job 221-22-00107   Dwgs 30A, 30B ---> Items, M11, M11R

        '-------DJL-------10-31-2023----------------------------------------------------------------Move to function CreateSpSht
        With ShipListSht
            JAt = 1
StartNextCheck:
            If NextChk = "Yes" Then
                GoTo GetNextItem
            End If

GetNextItem: For j = JAt To CntItems

                If StrLineNo = 42 Then
                    RowNo = (42 + j)
                Else
                    RowNo = (44 + j)
                End If

                If NextChk = "Yes" Then
                    NextChk = "No"
                    GoTo NextItemToFind
                End If

                If j = 1 Then
NextItemToFind:
                    DwgNoOld = .Range("D" & RowNo).Value
                    RevNoOld = .Range("E" & RowNo).Value
                    PcMkOld = .Range("F" & RowNo).Value
                    QtyOld = .Range("G" & RowNo).Value
                    DescOld = .Range("H" & RowNo).Value
                    Me.ProgressBar1.Value = j
                    JAt = j

                    If DwgNoOld = "1 = NEW ENTRY" Then
                        GoTo EndProcess2
                    End If

                    If JAt > CntItems Then
                        GoTo EndProcess2
                    End If

                    If RowNo > (CntItems + StrLineNo) Then
                        GoTo EndProcess2
                    End If

                    If IsNothing(DescOld) = True And IsNothing(PcMkOld) = True Then
                        If RowNo = CntItems Then
                            CntItems = (CntItems - 1)
                        End If

                        If RowNo > CntItems Then
                            If JAt = 1 Then
                                j = 1
                            Else
                                j = JAt
                            End If

                            If RowNo > CntItems And JAt + 44 >= CntItems Then
                                GoTo FoundAllItems
                            End If

                            GoTo NextJAt
                        End If

                        GoTo NextJ2
                    End If
                Else
                    If RowNo > (CntItems + StrLineNo) And JAt > 0 Then
                        JAt = (JAt + 1)
                        GoTo NextJ2
                    End If

                    'DwgNoNew = .Range("D" & RowNo).Value             ''Moved below
                    'RevNoNew = .Range("E" & RowNo).Value           
                    PcMkNew = .Range("F" & RowNo).Value
                    'QtyNew = .Range("G" & RowNo).Value         
                    DescNew = .Range("H" & RowNo).Value

                    If IsNothing(DescOld) = True And IsNothing(PcMkOld) = True Then
                        If RowNo = CntItems Then
                            CntItems = (CntItems - 1)
                        End If

                        If RowNo > CntItems Then
                            If JAt = 1 Then
                                j = 1
                            Else
                                j = JAt
                            End If

                            GoTo NextJAt
                        End If

                        GoTo NextJ2
                    End If

                    If IsNothing(DescNew) = True And IsNothing(PcMkNew) = True Then
                        If RowNo = CntItems Then
                            CntItems = (CntItems - 1)
                        End If

                        If RowNo > CntItems Then
                            If JAt = 1 Then
                                j = 1
                            Else
                                j = JAt
                            End If

                            GoTo NextJAt
                        End If

                        GoTo NextJ2
                    End If

                    DwgNoNew = .Range("D" & RowNo).Value            'This works for Tulsa, but sucks for Pittsburg.
                    Chk10D = InStr(DwgNoNew, "10D")
                    Chk10E = InStr(DwgNoNew, "10E")
                    LenDwgNo2 = Len(DwgNoNew)

                    If Chk10D > 0 And LenDwgNo2 > 4 Then
                        DwgNoNew2 = Mid(DwgNoNew, 4, Len(DwgNoNew))                         '-------Pittsburgh
                        DwgNoOld2 = Mid(DwgNoOld, 4, Len(DwgNoOld))
                        ChkCnt = 0
                    Else
                        If Chk10E > 0 And LenDwgNo2 > 4 Then
                            If Chk10E > 0 And InStr(DwgNoOld, "10D") = 1 Then
                                GoTo NextJ2
                            End If

                            DwgNoNew2 = Mid(DwgNoNew, 4, Len(DwgNoNew))                         '-------Pittsburgh
                            DwgNoOld2 = Mid(DwgNoOld, 4, Len(DwgNoOld))
                            ChkCnt = 0
                        Else
                            DwgNoNew2 = Mid(DwgNoNew, 1, 2)                         '-------Tulsa
                            DwgNoOld2 = Mid(DwgNoOld, 1, 2)
                            ChkCnt = 0
                        End If
                    End If

                    If DwgNoOld2 = DwgNoNew2 Then
                        If PcMkOld = PcMkNew And IsNothing(PcMkOld) = False Then     'PartNo, And Description are equal then delete
                            DwgNoNew = .Range("D" & RowNo).Value
                            RevNoNew = .Range("E" & RowNo).Value
                            QtyNew = .Range("G" & RowNo).Value
ChkDesc:
                            If DescOld = DescNew And IsNothing(DescOld) = False Then
                                If RevNoOld = RevNoNew And IsNothing(RevNoOld) = False Then
                                    If QtyOld = QtyNew And IsNothing(QtyOld) = False Then
                                        .Range("A" & RowNo & ":X" & RowNo).Delete(Shift:=XlDeleteShiftDirection.xlShiftUp)
                                        '-------DJL-------11-29-2023-------Need to delete item from the array.
                                        CntItems = (CntItems - 1)

                                        'On 221-22-00101 found that drawing reference on 33B should have been N15 instead of N14
                                        'Only remove one duplicate the rest could be referencing the wrong part number.
                                        If JAt = 1 Then
                                            j = 1
                                        Else
                                            j = JAt
                                        End If

                                        JAt = (JAt + 1)
                                        NextChk = "Yes"
                                        GoTo StartNextCheck
                                    End If
                                End If
                            Else
                                If DescOld <> DescNew And ChkCnt < 2 Then
                                    DescNew = LTrim(DescNew)             '-------221-22-00107 dwgs 39A thru 40B Item M4 The Description had an extra space at the end of the decription.
                                    DescNew = RTrim(DescNew)

                                    DescOld = LTrim(DescOld)
                                    DescOld = RTrim(DescOld)
                                    ChkCnt = (ChkCnt + 1)
                                    GoTo ChkDesc
                                End If
                            End If

                            ChkCnt = 0
                        Else
                            GoTo NextJ2
                        End If
                    Else
                        If DwgNoOld2 < DwgNoNew2 Then
                            If JAt = 1 Then
                                j = 1
                            Else
                                j = JAt
                            End If

                            JAt = (JAt + 1)
                            NextChk = "Yes"
                            GoTo StartNextCheck             '-------221-22-00101 N14
                        Else                                '-------If Pittsburgh is going from 10D to 10E then go to next drawing.
                            If Mid(DwgNoOld, 1, 3) = "10D" And Mid(DwgNoNew, 1, 3) = "10E" Then
                                If JAt = 1 Then
                                    j = 1
                                Else
                                    j = JAt
                                End If

                                JAt = (JAt + 1)
                                NextChk = "Yes"
                                GoTo StartNextCheck             '-------221-22-00101 N14
                            End If
                        End If
                    End If
                End If
                If j >= CntItems And JAt = CntItems Then
                    GoTo EndProcess2
                End If
NextJ2:
                ChkCnt = 0
            Next j
NextJAt:
            If JAt < RowNo And JAt > 0 Then
                JAt = (JAt + 1)
                NextChk = "Yes"

                If RowNo < (CntItems + StrLineNo) Then
                    Me.ProgressBar1.Value = JAt
                End If

                GoTo StartNextCheck
            End If
FoundAllItems:
        End With

        Me.ProgressBar1.Value = 0
EndProcess2:

        'Turning Compare process back on per OU 212 PM manager                          'DJL-210-25-2023
        PrgName = "StartButton2-CompareSpreadSht"                          '-------DJL-12-19-2024
        If Me.CheckBox1.CheckState = 1 Then
            NewShipListSht = ExcelApp.Application.ActiveWorkbook.Sheets("Shipping List")

            If PathBox.Text <> PathBox2.Text Then
                OldShipListFileStr = PathBox2.Text & ShipListBox.SelectedItem.ToString
            Else
                OldShipListFileStr = PathBox.Text & ShipListBox.SelectedItem.ToString
            End If

            Me.Refresh()

            CompareShipList()       '-------------Compare program look for Deleted, New Items ETC.
        Else
            If CurrentDwgRev = "0" Then
                '----------------------------------------Fix All Lines to Green when Revision Is Zero....
                NewShipListSht = ExcelApp.Application.ActiveWorkbook.Sheets("Shipping List")
                FormatShipListRev0(NewBOM, OldBOM, ShipListSht, BOMSheet, CntItems, StrLineNo)
            End If
        End If

        PrgName = "StartButton2-CompareDone"                          '-------DJL-12-19-2024
        ProgressBar1.Value = 0
        ProgressBar1.Maximum = 5
        ProgressBar1.Visible = False                        'True                           'DJL-------06/02/2025
        Me.Label2.Text = "Processing Information........Please Wait."
        ProgressBar1.Value = 1
        RevNo = RevNo2
        CopyBOMFile(OldFileNam, RevNo)
        AcadApp.ActiveDocument.Close()
Cancel:

        '-------DJL-------10-31-2023----------------------------------------------------------------Now write out to Excel.
        PrgName = "StartButton2-CloseAutoCAD"                          '-------DJL-12-19-2024
        ProgressBar1.Value = 2

        If AcadOpen = False Then
            If gvntSDIvar = True Then
                AcadPref.SingleDocumentMode = False
                AcadApp.Quit()
            End If
        Else
            AcadApp.Quit()
        End If

        AcadDoc = Nothing
        AcadApp = Nothing
        ExcelApp.Application.Visible = True

        PrgName = "StartButton2-CloseSpreadSht"                          '-------DJL-12-19-2024
        ProgressBar1.Value = 3
        Workbooks = ExcelApp.Workbooks
        CntWorkbook = Workbooks.Count
        WorkBookName = MainBOMFile.Application.ActiveWorkbook.Name
        Dim WBook As Excel.Workbook
        Dim TempFile, TempFile2 As Workbook
        Dim TempName As String
        TempFile2 = Nothing

        For Each WBook In ExcelApp.Workbooks
            If InStr(WBook.Name, "ShipListVBNet") > 0 Then
                WBook.Close(False)
            End If
        Next

        ProgressBar1.Value = 4

        If IsNothing(TempFile2) <> True Then
            TempFile.Close(False)
            TempFile = Nothing
        End If


        ProgressBar1.Value = 5
        Me.Label2.Text = "Your Shipping List has been Created."
        MsgBox("Your Shipping List has been Created.")
        ExcelApp.Application.Visible = True
        Me.Close()

Err_StartButton_Click:
        ErrNo = Err.Number

        If ErrNo <> 0 Then
            PrgName = "BtnSpeedTest"
            PriPrg = "ShipListReadBOMAutoCAD"
            ErrMsg = Err.Description
            ErrSource = Err.Source
            ErrDll = Err.LastDllError
            ErrLastLineX = Err.Erl
            ErrException = Err.GetException

            If ErrNo = -2145320885 And ErrMsg = "Problem in unloading DVB file" Then
                Resume Next
            End If

            If ErrNo = -2145320885 And InStr(ErrMsg, "Exception from HRESULT") > 0 Then
                Threading.Thread.Sleep(35)
                Resume Next
            End If

            CntDwgsNotFound = InStr(ErrMsg, "not a valid drawing")

            If ErrNo = -2145320825 And CntDwgsNotFound > 0 Then
                MsgBox("AutoCAD found a bad Drawing, " & DwgItem3 & ", Going to next drawing.")
                BadDwgFound = "Yes"
                CntDwgsNotFound = 0
                Resume Next
            End If

            If ErrNo = 91 And ErrMsg = "Problem in unloading DVB file" Then
                Resume Next
            End If

            If ErrMsg = "Cannot create ActiveX component." Then
                MsgBox("This program is having a problem opening AutoCAD, Please open AutoCAD then pick this OK button.")
                AcadApp = GetObject(, "AutoCAD.Application")
                Threading.Thread.Sleep(25)
                Resume Next
            End If



            If ErrNo = -2147418111 And Mid(ErrMsg, 1, 28) = "Call was rejected by callee." Then
                AcadApp = GetObject(, "AutoCAD.Application")
                Err.Clear()
                Threading.Thread.Sleep(25)
                Resume
            End If

            If ErrNo = 91 And Mid(ErrMsg, 1, 17) = "The RPC server is" Then
                AcadApp = CreateObject("AutoCAD.Application")
                Threading.Thread.Sleep(25)
                Resume
            End If

            If ErrNo = 91 And ErrMsg = "Problem in unloading DVB file" Then
                Resume Next
            End If

            If ErrNo = 91 And ErrMsg = "" Then
                GoTo EndPrg
            End If

            If ErrNo = 9 And InStr(ErrMsg, "Index was outside the bounds of the array.") > 0 Then
                Resume Next
            End If

            Dim st As New StackTrace(Err.GetException, True)
            CntFrames = st.FrameCount
            GetFramesSrt = st.GetFrames
            PrgLineNo = GetFramesSrt(CntFrames - 1).ToString
            PrgLineNo = PrgLineNo.Replace("@", "at")
            LenPrgLineNo = (Len(PrgLineNo))
            PrgLineNo = Mid(PrgLineNo, 1, (LenPrgLineNo - 2))

            HandleErrSQL(PrgName, ErrNo, ErrMsg, ErrSource, PriPrg, ErrDll, DwgItem3, PrgLineNo)

            If ErrNo = "462" And Mid(ErrMsg, 1, 30) = "The RPC server is unavailable." Then
                If Err.Number Then
                    Err.Clear()
                End If

                If IsNothing(AcadApp) = True Then
                    AcadApp = CreateObject("AutoCAD.Application")
                Else
                    Err.Clear()
                    MsgBox("Please Open AutoCAD then pick this OK button.")
                    AcadApp = GetObject(, "AutoCAD.Application")
                    AcadOpen = True
                End If

                If Err.Number Then
                    MsgBox(Err.Description)
                    Information.Err.Clear()
                    AcadApp = CreateObject("AutoCAD.Application")
                End If

                If ProblemAt = "ReadingTitleBlock" Then
                    ProblemAt = "ReadingTitleBlock2"
                    Resume Next
                End If

                Resume
            End If

            If ErrNo = "-2147418113" And Mid(ErrMsg, 1, 30) = "Internal application error." Then
                If Err.Number Then
                    Err.Clear()
                End If

                If IsNothing(AcadApp) = True Then
                    AcadApp = CreateObject("AutoCAD.Application")
                    Threading.Thread.Sleep(25)
                    AcadOpen = True
                    Resume
                Else
                    MsgBox("Please Open AutoCAD then pick this OK button.")
                    AcadApp = GetObject(, "AutoCAD.Application")
                    Threading.Thread.Sleep(25)
                    AcadOpen = True
                    Resume
                End If
            End If

            If ErrNo = "20" And ErrMsg = "Resume without error." Then
                Exit Sub
            End If

            If ErrNo = "-2145320900" And ErrMsg = "Failed to get the Document object" Then
                If ErrAt = "Unload Adept" Then
                    AcadApp.Documents.Add()
                    Resume
                End If
            End If

            If ErrNo = "-2147023170" And ErrMsg = "The remote procedure call failed" Then
                Information.Err.Clear()
                AcadApp = CreateObject("AutoCAD.Application")
                Threading.Thread.Sleep(50)
                AcadOpen = True
            End If

            If IsNothing(GenInfo.UserName) = True Then
                GenInfo.UserName = Environment.UserName
            End If

            If GenInfo.UserName = "dlong" Then
                ExceptPos = 0
                SearchException = "Exception"
                ExceptPos = InStr(ErrMsg, 1)
                If ExceptPos > 0 Then
                    CntExcept = (CntExcept + 1)
                    If CntExcept < 20 Then
                        If IsNothing(AcadApp) = True Then
                            AcadApp = GetObject(, "AutoCAD.Application")
                            Threading.Thread.Sleep(25)
                        End If

                        Resume
                    End If
                End If

                MsgBox(ErrMsg)
                Stop
                Resume
            Else
                ExceptPos = 0
                SearchException = "Exception"
                ExceptPos = InStr(ErrMsg, 1)
                If ExceptPos > 0 Then
                    CntExcept = (CntExcept + 1)
                    If CntExcept < 20 Then
                        If IsNothing(AcadApp) = True Then
                            AcadApp = GetObject(, "AutoCAD.Application")
                            Threading.Thread.Sleep(25)
                        End If
                        Resume
                    End If
                End If
            End If
        End If
EndPrg:
    End Sub

End Class