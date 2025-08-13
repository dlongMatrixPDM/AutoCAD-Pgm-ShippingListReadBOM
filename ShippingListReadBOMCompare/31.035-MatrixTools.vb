Imports System
Imports System.Runtime.InteropServices

Namespace MatrixTool

    Public Class MatrixTools
        'Public Structure InputType
        'Dim GottaHaveIt 'Required by .Net

        'Shared db_String, strSQL1, strSQL2, PrgName, ByName, ByNam, TableNam As String
        'Shared TaskTyp, TType, TTime, ErrMsg, ErrNo, ErrSource, FormString As String
        'Shared NewJobNo, NewSubNo, newjob, NewSub, NewTask, NewBy, NewFirst, NewLast As String
        'Shared TJob, Temp, Tsub, TTask, NLookFor, CountVal, TDate, TRecNum, TLookFor As String
        'Shared NewDate, NewTime, NewType, NewRecNo, NewDept, NewEstHrs, NewEstChk As String
        'Shared NewEst, NewAdd, NewChg, NewMemo, NewStat, strSQL3, CurrentBy As String
        'Shared CurrentTask, LastNam, FirstNam, FChrTest, CodeName, DeptName As String
        'Shared PriPrg, ProgPhase, UserName, FChr, FindUser, CodePrg As String
        'Shared dbSQL_String, WorkStationID1, OldWorkStation, FuncName As String
        'Shared RootName, PriName, SecName, FileName, Test As String
        'Shared JobNum, SubNum, TaskName, Emp, TableName, SourceName, PrevLastNam As String
        'Shared MInit, TITLE, Funct, DeptNam, HireDate, PTOPrevYr, Birth, LastNamChg As String
        'Shared Adapter2 As SqlClient.SqlDataAdapter
        'Shared ETime
        'Shared DelCnt, Cnt1, Cnt2, Cnt3, ErrMsgCnt, rs1Cnt, rs2Cnt, LenSecName As Integer
        'Shared OldRecNum, NewRecNum, LenFileName, LenPriName, LenRootName As Integer
        'Shared FormName
        'Shared db As New ADODB.Connection
        'Shared dbHandleErr As New ADODB.Connection
        'Shared db2 As New ADODB.Connection
        'Shared dbSQL As New SqlClient.SqlConnection
        'Shared dbSQL2 As New SqlClient.SqlConnection
        'Shared dbSQLState As Integer
        'Shared rs1SQL As New SqlClient.SqlCommand
        'Shared rs1 As New ADODB.Recordset
        'Shared rs2 As New ADODB.Recordset
        'Shared rs3 As New ADODB.Recordset
        'Shared DsTLengStaff As New DataSet
        'Shared DsReportsUsed As New DataSet

        'Shared ErrDll As String
        'Shared ErrLastLineX As Integer
        'Shared ErrException As System.Exception
        'Shared AcadOpen As Boolean
        'Shared rs As ADODB.Recordset

        'Shared Mospace, Paspace As Object
        'Shared UtilObj As Object
        'Public Shared AcadApp As AutoCAD.AcadApplication
        'Shared XdataType, XdataValue As Object
        'Shared Sset As AutoCAD.AcadSelectionSet
        'Shared SaveAction As Boolean
        'Shared ThisDrawing As AutoCAD.AcadDocument

        '        Public Shared Function Activate_Autocad(ByVal AcadApp) As AutoCAD.AcadApplication
        '            PrgName = "Activate_Autocad"

        '            On Error GoTo Err_Activate_Autocad

        '            On Error Resume Next
        '            If Err.Number Then
        '                Err.Clear()
        '            End If

        '            AcadApp = GetObject(, "AutoCAD.Application.")     'AcadApp = GetObject(, "Autocad.Application")
        '            AcadOpen = True

        '            If Err.Number Then
        '                'MsgBox(Err.Description)
        '                Information.Err.Clear()
        '                AcadApp = CreateObject("AutoCAD.Application.")    'AcadApp = CreateObject("Autocad.Application")
        '                AcadOpen = False             'AcadOpen = False
        '                If Err.Number Then
        '                    Information.Err.Clear()
        '                    AcadApp = CreateObject("AutoCAD.Application.")
        '                    If Err.Number Then
        '                        'MsgBox(Err.Description)
        '                        'MsgBox("Instance of 'AutoCAD.Application' could not be created.")

        '                        AcadApp.Visible = False
        '                        MsgBox("Now running " & AcadApp.Name & " version " & AcadApp.Version)


        '                        If GenInfo3135.UserName = "dlong" Then
        '                            Stop
        '                            Resume
        '                        Else
        '                            Exit Function
        '                        End If
        '                    End If
        '                End If
        '            End If

        '            AcadApp.Visible = False
        '            ThisDrawing = AcadApp.ActiveDocument
        '            Mospace = ThisDrawing.ModelSpace
        '            Paspace = ThisDrawing.PaperSpace
        '            UtilObj = ThisDrawing.Utility

        'Err_Activate_Autocad:
        '            ErrNo = Err.Number

        '            If ErrNo <> 0 Then
        '                PriPrg = "ShipListReadBOMAutoCAD"
        '                ErrMsg = Err.Description
        '                ErrSource = Err.Source
        '                ErrDll = Err.LastDllError
        '                ErrLastLineX = Err.Erl
        '                ErrException = Err.GetException

        '                Dim st As New StackTrace(Err.GetException, True)
        '                CntFrames = st.FrameCount
        '                GetFramesSrt = st.GetFrames
        '                PrgLineNo = GetFramesSrt(CntFrames - 1).ToString
        '                PrgLineNo = PrgLineNo.Replace("@", "at")
        '                LenPrgLineNo = (Len(PrgLineNo))
        '                PrgLineNo = Mid(PrgLineNo, 1, (LenPrgLineNo - 2))

        '                ShippingList_Menu.HandleErrSQL(PrgName, ErrNo, ErrMsg, ErrSource, PriPrg, ErrDll, DwgItem, PrgLineNo)
        '                If GenInfo3135.UserName = "dlong" Then
        '                    MsgBox(ErrMsg)
        '                    Stop
        '                    Resume
        '                Else
        '                    ExceptionPos = InStr(1, ErrMsg, "Exception")
        '                    CallPos = InStr(1, ErrMsg, "Call was rejected by callee")
        '                    CntExcept = (CntExcept + 1)

        '                    If CntExcept < 20 Then
        '                        If ExceptionPos > 0 Then         'If program errors out due to use input, then repeat process.
        '                            Resume
        '                        End If
        '                        If CallPos > 0 Then         'If program errors out due to use input, then repeat process.
        '                            Resume
        '                        End If
        '                    End If
        '                End If
        '            End If

        '        End Function

        'End Structure

    End Class

End Namespace
