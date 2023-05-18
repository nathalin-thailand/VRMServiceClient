Module bk_code_notused


    '    Imports System.Data.SqlClient
    '    Imports System.IO
    '    Imports System.Net
    '    Imports System.Text
    '    Imports System.Web.Script.Serialization
    '    Imports Newtonsoft.Json
    '    Imports RestSharp

    '    Public Class VRMService


    '        Dim GetDOAOnline As Boolean = False
    '        Dim StatusRun As Boolean = False


    '        Dim tbOfflineInteface As New DataTable()

    '        Public Sub OnDebug()
    '            Me.OnStart(Nothing)
    '        End Sub


    '        Protected Overrides Sub OnStart(ByVal args() As String)
    '            On Error GoTo ErrPoint

    '            Dim lines() As String = {"Starting time : " + DateTime.Now.ToString()}
    '            System.IO.File.AppendAllLines(LogPath, lines)

    '            If System.IO.File.Exists(ConfigFile) = False Then
    '                Dim strWrite As String = "TimeMin=" & TimeInteval.ToString() '& vbCrLf & "database=" & vbCrLf & "user=" & vbCrLf & "password="
    '                System.IO.File.WriteAllText(ConfigFile, strWrite)

    '            Else
    '                lines = System.IO.File.ReadAllLines(ConfigFile)
    '                For zi = 0 To lines.Length - 1
    '                    Dim linStr = lines(zi)
    '                    Dim arrStr = lines(zi).Split("=")
    '                    If (arrStr.Length = 2) Then
    '                        If LCase(arrStr(0)) = LCase("TimeMin") Then TimeInteval = Convert.ToDouble(arrStr(1))

    '                    End If
    '                Next
    '            End If

    '            Timer1.Interval = Convert.ToInt32(TimeInteval * (1000 * 60))
    '            Timer1.Enabled = True
    '            StatusRun = False
    '            Exit Sub

    'ErrPoint:
    '            lines = {Err.Description + " : " + DateTime.Now.ToString()}
    '            System.IO.File.AppendAllLines(LogPath, lines)
    '        End Sub

    '        Protected Overrides Sub OnStop()
    '            Dim lines() As String = {"Stop time : " + DateTime.Now.ToString()}
    '            System.IO.File.AppendAllLines(LogPath, lines)
    '        End Sub

    '        Private Sub Timer1_Elapsed(sender As Object, e As Timers.ElapsedEventArgs) Handles Timer1.Elapsed
    '            If StatusRun Then Exit Sub

    '            On Error GoTo ErrPoint

    '            StatusRun = True

    '            Dim ErrStr As String
    '            Dim lines() As String

    '            'Dim lines() As String = {"Run time : " + DateTime.Now.ToString()}
    '            'System.IO.File.AppendAllLines(LogPath, lines)

    '            If System.IO.File.Exists(SQLConfigFile) = False Then
    '                Dim strWrite As String = "server=" & vbCrLf & "database=" & vbCrLf & "user=" & vbCrLf & "password="
    '                System.IO.File.WriteAllText(SQLConfigFile, strWrite)
    '            Else
    '                lines = System.IO.File.ReadAllLines(SQLConfigFile)
    '                For zi = 0 To lines.Length - 1
    '                    Dim linStr = lines(zi)
    '                    Dim arrStr = lines(zi).Split("=")
    '                    If (arrStr.Length = 2) Then
    '                        If LCase(arrStr(0)) = "server" Then locServer = arrStr(1)
    '                        If LCase(arrStr(0)) = "database" Then locDB = arrStr(1)
    '                        If LCase(arrStr(0)) = "user" Then locUID = arrStr(1)
    '                        If LCase(arrStr(0)) = "password" Then locPWD = arrStr(1)
    '                    End If
    '                Next
    '            End If


    '            Dim ConnLoc As New ADODB.Connection

    '            Dim locConStr = "Driver={SQL Server};Server=" & locServer & ";Database=" & locDB & ";UID=" & locUID & ";PWD=" & locPWD & ";"
    '            ConnLoc.ConnectionString = locConStr
    '            ConnLoc.CursorLocation = ADODB.CursorLocationEnum.adUseClient
    '            ConnLoc.CommandTimeout = 1500
    '            ConnLoc.Open()


    '            tbConfig.Columns.Clear()

    '            tbConfig.Columns.AddRange(New DataColumn() {
    '            New DataColumn("keyType", GetType(String)),
    '            New DataColumn("keyWord", GetType(String)),
    '            New DataColumn("config", GetType(String))
    '         })

    '            tbConfig.Rows.Clear()

    '            Dim Rs As New ADODB.Recordset
    '            Rs = ReadSQL("select * from tbSystemOfflineConfig", ConnLoc)

    '            Dim zStrItem(Rs.Fields.Count - 1) As String
    '            Do Until Rs.EOF
    '                'tbConfig.Rows.Add(Rs)

    '                zStrItem(0) = Rs.Fields("keyType").Value
    '                zStrItem(1) = Rs.Fields("keyWord").Value
    '                zStrItem(2) = Rs.Fields("config").Value
    '                tbConfig.Rows.Add(zStrItem)
    '                Rs.MoveNext()
    '            Loop
    '            RsClose(Rs)



    '            'WriteLog(ConnLoc, "Run", "Connect", "OK")
    '            Dim ConnSrv As New ADODB.Connection
    '            If Not IsNothing(ConnLoc.ConnectionString) Then
    '                'ConnSrv = ConnectServer(ConnLoc)
    '                'If IsNothing(ConnSrv) Then GoTo JmpOut
    '                If ConnectServerAPI() Then
    '                    SycData(ConnLoc, ConnSrv)
    '                End If

    '                Call triggerLocalRun(ConnLoc)

    '            End If



    '            GoTo JmpOut

    'ErrPoint:
    '            MsgBox(Err.Description)

    '            Dim line1s() As String = {Err.Description + " : " + DateTime.Now.ToString()}
    '            System.IO.File.AppendAllLines(LogPath, line1s)
    '            Err.Clear()

    'JmpOut:
    '            StatusRun = False
    '        End Sub

    '        Public Sub SycData(ConnLoc As ADODB.Connection, ConnSrv As ADODB.Connection)

    '            System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.CurrentCulture
    '            WriteLog(ConnLoc, "Run", "SycData", "")

    '            entityId = "0"
    '            Dim rows = tbConfig.[Select](" keyType = 'Machine' and keyWord = 'Entity' ")
    '            For Each row As DataRow In rows
    '                entityId = strNull(row.Item("config"))
    '            Next

    '            If entityId = "" Then entityId = "0"
    '            If entityId = "0" Then Exit Sub

    '            Sql = "select * from vw_org_entities where entityid = " & entityId
    '            Dim Rs As New ADODB.Recordset
    '            Rs = ReadSQL(Sql, ConnLoc)
    '            If Rs.RecordCount <> 0 Then
    '                corpId = strNull(Rs.Fields("corpId").Value)
    '                grpId = strNull(Rs.Fields("grpId").Value)
    '                buId = strNull(Rs.Fields("buId").Value)
    '                companyId = strNull(Rs.Fields("companyId").Value)
    '                entityId = strNull(Rs.Fields("entityId").Value)
    '                corpCode = strNull(Rs.Fields("corpCode").Value)
    '                grpCode = strNull(Rs.Fields("grpCode").Value)
    '                buCode = strNull(Rs.Fields("buCode").Value)
    '                companyCode = strNull(Rs.Fields("companyCode").Value)
    '                entityCode = strNull(Rs.Fields("entityCode").Value)
    '            End If
    '            RsClose(Rs)

    '            rows = tbConfig.[Select](" keyType = 'DOA' and keyWord = 'GetOnline' ")
    '            For Each row As DataRow In rows
    '                GetDOAOnline = LCase(strNull(row.Item("config"))) = "yes"
    '            Next


    '            If GetDOAOnline Then UpdateDOA(ConnLoc, ConnSrv)


    '            Call InterfaceClient2Server(ConnLoc, ConnSrv)
    '            Call InterfaceServer2Client(ConnLoc, ConnSrv)
    '        End Sub


    '        Public Function InterfaceServer2Client(ConnLoc As ADODB.Connection, ConnSrv As ADODB.Connection)



    '            Dim RunDateTime As String = Now.ToString("yyyy-MMM-dd  HH:mm:ss")

    '            Dim Url = ""
    '            Dim irow As DataRow = tbConfig.[Select](" keyType = 'URL'and keyWord = 'APIGetServer' ").FirstOrDefault()
    '            If IsDBNull(irow) = False Then Url = strNull(irow.Item("config"))

    '            Dim pram = "?getEntity=" & entityId & "&keyDateTime=" & RunDateTime

    '            Dim hostName As String = GetURLHostName(Url)
    '            Dim uriName As String = Replace(Url, hostName, "")

    '            Dim data_client = New RestClient(hostName)
    '            Dim request = New RestRequest(uriName & pram, Method.Get)
    '            request.AddHeader("Accept", "text/xml")

    '            'request.AddHeader("Accept", "application/json")
    '            'request.RequestFormat = DataFormat.
    '            'request.AddJsonBody(listTB)
    '            'request.RequestFormat = DataFormat.Json
    '            'request.AddJsonBody(listTB)

    '            'request.AddHeader("Authorization", "bearer " + bearer)
    '            Dim data_response = data_client.Execute(request)
    '            Dim data_response_raw = data_response.Content

    '            If (data_response.IsSuccessful) Then

    '                Dim strJson As String = data_response_raw
    '                If (strJson <> "[]") Then
    '                    Dim RsJsUP As New ADODB.Recordset
    '                    Dim JSql As String = "exec spInterfaceClient2ServerJSonInsert '" & strJson & "'"
    '                    RsJsUP = ReadSQL(JSql, ConnLoc)
    '                    RsClose(RsJsUP)
    '                    Return True
    '                End If


    '                'Dim Rs As New ADODB.Recordset
    '                'Rs = ReadSQL("select * from tbInterfaceOfflineRecord where Id is null", ConnLoc)
    '                'Dim dt As New System.Data.DataTable
    '                '    dt = JsonConvert.DeserializeObject(Of DataTable)(strJson)
    '                '    Dim aaa As New tbInterfaceOffline()
    '                '    Dim aa1a = JsonConvert.DeserializeObject(strJson)

    '                '    aaa = JsonConvert.DeserializeObject(Of tbInterfaceOffline)(JsonConvert.DeserializeObject(aa1a))
    '                '    dt.DefaultView.Sort = "keySort ASC"
    '                '    Dim rows = dt.[Select]  '(" keyType = 'Machine' and keyWord = 'Entity' ")
    '                '    For Each row As DataRow In rows
    '                '        Dim fNameAll = ""
    '                '        Dim fValAll = ""
    '                '        Dim InserSQL = ""
    '                '        For zi = 0 To Rs.Fields.Count - 1
    '                '            Dim fName As String = Rs.Fields(zi).Name
    '                '            Dim fVal As String = RsAsInsert(Rs.Fields(fName), row.Item(fName))

    '                '            If fVal <> "null" Then
    '                '                If fNameAll <> "" Then fNameAll = fNameAll & ","
    '                '                fNameAll = fNameAll & fName

    '                '                If fValAll <> "" Then fValAll = fValAll & ","
    '                '                fValAll = fValAll & fVal
    '                '            End If

    '                '        Next

    '                '        InserSQL = "insert into tbInterfaceOfflineRecord ( " & fNameAll & " ) values (" & fValAll & ")"

    '                '        Dim RsUp As New ADODB.Recordset
    '                '        RsUp = WriteSQL(InserSQL, ConnLoc)
    '                '        RsClose(RsUp)

    '                '    Next
    '                'Return True
    '                '    RsClose(Rs)
    '            Else

    '                Dim lines() As String = {"Error asynchronous Server2Client : " + DateTime.Now.ToString()}
    '                System.IO.File.AppendAllLines(LogPath, lines)
    '                Return False
    '            End If

    '        End Function





    '        Public Function triggerLocalRun(ConnLoc As ADODB.Connection)

    '            Dim RunDateTime As String = Now.ToString("yyyy-MMM-dd  HH:mm:ss")
    '            Dim Rs As New ADODB.Recordset
    '            Rs = WriteSQL("exec spInterfaceClientTrigger '" & entityId & "','" & RunDateTime & "'", ConnLoc)
    '            RsClose(Rs)
    '        End Function










    '        Public Function InterfaceClient2Server(ConnLoc As ADODB.Connection, ConnSrv As ADODB.Connection)
    '            Dim Rs As New ADODB.Recordset

    '            Dim RunDateTime As String = Now.ToString("yyyy-MMM-dd  HH:mm:ss")

    '            Rs = ReadSQL("exec spInterfaceClient2Server '" & entityId & "', '" & RunDateTime & "' ", ConnLoc)
    '            If Rs.RecordCount = 0 Then
    '                Return True
    '            End If

    '            Dim listTB As New List(Of Object) 'List(Of tbInterfaceOffline)
    '            Dim listTB2 As New List(Of tbOfflineInteface) 'List(Of tbInterfaceOffline)


    '            Do Until Rs.EOF

    '                Dim iOffline2 = New tbOfflineInteface
    '                With iOffline2
    '                    .Id = RsNullDB(Rs.Fields("Id"))
    '                    .KeyEntity = RsNullDB(Rs.Fields("KeyEntity"))
    '                    .KeyAction = RsNullDB(Rs.Fields("KeyAction"))
    '                    .KeyRecord = RsNullDB(Rs.Fields("KeyRecord"))
    '                    .KeySort = RsNullDB(Rs.Fields("KeySort"))
    '                    .corpId = RsNullDB(Rs.Fields("corpId"))
    '                    .grpId = RsNullDB(Rs.Fields("grpId"))
    '                    .buId = RsNullDB(Rs.Fields("buId"))
    '                    .companyId = RsNullDB(Rs.Fields("companyId"))
    '                    .entityId = RsNullDB(Rs.Fields("entityId"))
    '                    .RecId = RsNullDB(Rs.Fields("RecId"))
    '                    .Id_01 = RsNullDB(Rs.Fields("Id_01"))
    '                    .Id_02 = RsNullDB(Rs.Fields("Id_02"))
    '                    .Id_03 = RsNullDB(Rs.Fields("Id_03"))
    '                    .Id_04 = RsNullDB(Rs.Fields("Id_04"))
    '                    .Id_05 = RsNullDB(Rs.Fields("Id_05"))
    '                    .Id_06 = RsNullDB(Rs.Fields("Id_06"))
    '                    .Id_07 = RsNullDB(Rs.Fields("Id_07"))
    '                    .Id_08 = RsNullDB(Rs.Fields("Id_08"))
    '                    .Id_09 = RsNullDB(Rs.Fields("Id_09"))
    '                    .Id_10 = RsNullDB(Rs.Fields("Id_10"))
    '                    .Id_11 = RsNullDB(Rs.Fields("Id_11"))
    '                    .Id_12 = RsNullDB(Rs.Fields("Id_12"))
    '                    .Id_13 = RsNullDB(Rs.Fields("Id_13"))
    '                    .Id_14 = RsNullDB(Rs.Fields("Id_14"))
    '                    .Id_15 = RsNullDB(Rs.Fields("Id_15"))

    '                    .Str_5_01 = RsNullDB(Rs.Fields("Str_5_01"))
    '                    .Str_5_02 = RsNullDB(Rs.Fields("Str_5_02"))
    '                    .Str_5_03 = RsNullDB(Rs.Fields("Str_5_03"))
    '                    .Str_5_04 = RsNullDB(Rs.Fields("Str_5_04"))
    '                    .Str_5_05 = RsNullDB(Rs.Fields("Str_5_05"))


    '                    .Str_25_01 = RsNullDB(Rs.Fields("Str_25_01"))
    '                    .Str_25_02 = RsNullDB(Rs.Fields("Str_25_02"))
    '                    .Str_25_03 = RsNullDB(Rs.Fields("Str_25_03"))
    '                    .Str_25_04 = RsNullDB(Rs.Fields("Str_25_04"))
    '                    .Str_25_05 = RsNullDB(Rs.Fields("Str_25_05"))
    '                    .Str_25_06 = RsNullDB(Rs.Fields("Str_25_06"))
    '                    .Str_25_07 = RsNullDB(Rs.Fields("Str_25_07"))
    '                    .Str_25_08 = RsNullDB(Rs.Fields("Str_25_08"))
    '                    .Str_25_09 = RsNullDB(Rs.Fields("Str_25_09"))
    '                    .Str_25_10 = RsNullDB(Rs.Fields("Str_25_10"))

    '                    .Str_50_01 = RsNullDB(Rs.Fields("Str_50_01"))
    '                    .Str_50_02 = RsNullDB(Rs.Fields("Str_50_02"))
    '                    .Str_50_03 = RsNullDB(Rs.Fields("Str_50_03"))
    '                    .Str_50_04 = RsNullDB(Rs.Fields("Str_50_04"))
    '                    .Str_50_05 = RsNullDB(Rs.Fields("Str_50_05"))
    '                    .Str_50_06 = RsNullDB(Rs.Fields("Str_50_06"))
    '                    .Str_50_07 = RsNullDB(Rs.Fields("Str_50_07"))
    '                    .Str_50_08 = RsNullDB(Rs.Fields("Str_50_08"))
    '                    .Str_50_09 = RsNullDB(Rs.Fields("Str_50_09"))
    '                    .Str_50_10 = RsNullDB(Rs.Fields("Str_50_10"))
    '                    .Str_150_01 = RsNullDB(Rs.Fields("Str_150_01"))
    '                    .Str_150_02 = RsNullDB(Rs.Fields("Str_150_02"))
    '                    .Str_150_03 = RsNullDB(Rs.Fields("Str_150_03"))
    '                    .Str_150_04 = RsNullDB(Rs.Fields("Str_150_04"))
    '                    .Str_150_05 = RsNullDB(Rs.Fields("Str_150_05"))
    '                    .Str_250_01 = RsNullDB(Rs.Fields("Str_250_01"))
    '                    .Str_250_02 = RsNullDB(Rs.Fields("Str_250_02"))
    '                    .Str_250_03 = RsNullDB(Rs.Fields("Str_250_03"))
    '                    .Str_250_04 = RsNullDB(Rs.Fields("Str_250_04"))
    '                    .Str_250_05 = RsNullDB(Rs.Fields("Str_250_05"))
    '                    .Str_500_01 = RsNullDB(Rs.Fields("Str_500_01"))
    '                    .Str_500_02 = RsNullDB(Rs.Fields("Str_500_02"))
    '                    .Str_500_03 = RsNullDB(Rs.Fields("Str_500_03"))
    '                    .Str_500_04 = RsNullDB(Rs.Fields("Str_500_04"))
    '                    .Str_500_05 = RsNullDB(Rs.Fields("Str_500_05"))
    '                    .Str_Max_01 = RsNullDB(Rs.Fields("Str_Max_01"))
    '                    .Str_Max_02 = RsNullDB(Rs.Fields("Str_Max_02"))
    '                    .Str_Max_03 = RsNullDB(Rs.Fields("Str_Max_03"))
    '                    .Str_Max_04 = RsNullDB(Rs.Fields("Str_Max_04"))
    '                    .Str_Max_05 = RsNullDB(Rs.Fields("Str_Max_05"))
    '                    .DateTime_01 = RsNullDT(Rs.Fields("DateTime_01"))
    '                    .DateTime_02 = RsNullDT(Rs.Fields("DateTime_02"))
    '                    .DateTime_03 = RsNullDT(Rs.Fields("DateTime_03"))
    '                    .DateTime_04 = RsNullDT(Rs.Fields("DateTime_04"))
    '                    .DateTime_05 = RsNullDT(Rs.Fields("DateTime_05"))

    '                    .Float_01 = RsNullDB(Rs.Fields("Float_01"))
    '                    .Float_02 = RsNullDB(Rs.Fields("Float_02"))
    '                    .Float_03 = RsNullDB(Rs.Fields("Float_03"))
    '                    .Float_04 = RsNullDB(Rs.Fields("Float_04"))
    '                    .Float_05 = RsNullDB(Rs.Fields("Float_05"))
    '                    .Float_06 = RsNullDB(Rs.Fields("Float_06"))
    '                    .Float_07 = RsNullDB(Rs.Fields("Float_07"))
    '                    .Float_08 = RsNullDB(Rs.Fields("Float_08"))
    '                    .Float_09 = RsNullDB(Rs.Fields("Float_09"))
    '                    .Float_10 = RsNullDB(Rs.Fields("Float_10"))
    '                    .Float_11 = RsNullDB(Rs.Fields("Float_11"))
    '                    .Float_12 = RsNullDB(Rs.Fields("Float_12"))
    '                    .Float_13 = RsNullDB(Rs.Fields("Float_13"))
    '                    .Float_14 = RsNullDB(Rs.Fields("Float_14"))
    '                    .Float_15 = RsNullDB(Rs.Fields("Float_15"))
    '                    .Float_16 = RsNullDB(Rs.Fields("Float_16"))
    '                    .Float_17 = RsNullDB(Rs.Fields("Float_17"))
    '                    .Float_18 = RsNullDB(Rs.Fields("Float_18"))
    '                    .Float_19 = RsNullDB(Rs.Fields("Float_19"))
    '                    .Float_20 = RsNullDB(Rs.Fields("Float_20"))

    '                    .Int_01 = RsNullDB(Rs.Fields("Int_01"))
    '                    .Int_02 = RsNullDB(Rs.Fields("Int_02"))
    '                    .Int_03 = RsNullDB(Rs.Fields("Int_03"))
    '                    .Int_04 = RsNullDB(Rs.Fields("Int_04"))
    '                    .Int_05 = RsNullDB(Rs.Fields("Int_05"))
    '                    .Int_06 = RsNullDB(Rs.Fields("Int_06"))
    '                    .Int_07 = RsNullDB(Rs.Fields("Int_07"))
    '                    .Int_08 = RsNullDB(Rs.Fields("Int_08"))
    '                    .Int_09 = RsNullDB(Rs.Fields("Int_09"))
    '                    .Int_10 = RsNullDB(Rs.Fields("Int_10"))
    '                    .Int_11 = RsNullDB(Rs.Fields("Int_11"))
    '                    .Int_12 = RsNullDB(Rs.Fields("Int_12"))
    '                    .Int_13 = RsNullDB(Rs.Fields("Int_13"))
    '                    .Int_14 = RsNullDB(Rs.Fields("Int_14"))
    '                    .Int_15 = RsNullDB(Rs.Fields("Int_15"))


    '                    .num_180_01 = RsNullDB(Rs.Fields("num_180_01"))
    '                    .num_180_02 = RsNullDB(Rs.Fields("num_180_02"))
    '                    .num_180_03 = RsNullDB(Rs.Fields("num_180_03"))
    '                    .num_180_04 = RsNullDB(Rs.Fields("num_180_04"))
    '                    .num_180_05 = RsNullDB(Rs.Fields("num_180_05"))
    '                    .bit_01 = RsNullDB(Rs.Fields("bit_01"))
    '                    .bit_02 = RsNullDB(Rs.Fields("bit_02"))
    '                    .bit_03 = RsNullDB(Rs.Fields("bit_03"))
    '                    .bit_04 = RsNullDB(Rs.Fields("bit_04"))
    '                    .bit_05 = RsNullDB(Rs.Fields("bit_05"))
    '                    .FileType = RsNullDB(Rs.Fields("FileType"))
    '                    .FileDisposition = RsNullDB(Rs.Fields("FileDisposition"))
    '                    .FileLength = RsNullDB(Rs.Fields("FileLength"))
    '                    .FileName = RsNullDB(Rs.Fields("FileName"))
    '                    .createBy = RsNullDB(Rs.Fields("createBy"))
    '                    .createDate = RsNullDT(Rs.Fields("createDate"))
    '                    .modifyBy = RsNullDB(Rs.Fields("modifyBy"))
    '                    .modifyDate = RsNullDT(Rs.Fields("modifyDate"))
    '                    .Inf_Status = RsNullDB(Rs.Fields("Inf_Status"))
    '                    .Inf_Datetime = RsNullDT(Rs.Fields("Inf_Datetime"))
    '                End With
    '                listTB2.Add(iOffline2)



    '                Dim iOffline = New With {
    '                .Id = RsNullDB(Rs.Fields("Id")),
    '                .KeyAction = RsNullDB(Rs.Fields("KeyAction")),
    '                .KeyRecord = RsNullDB(Rs.Fields("KeyRecord")),
    '                .corpId = RsNullDB(Rs.Fields("corpId")),
    '                .grpId = RsNullDB(Rs.Fields("grpId")),
    '                .buId = RsNullDB(Rs.Fields("buId")),
    '                .companyId = RsNullDB(Rs.Fields("companyId")),
    '                .entityId = RsNullDB(Rs.Fields("entityId")),
    '                .RecId = RsNullDB(Rs.Fields("RecId")),
    '                .Id_01 = RsNullDB(Rs.Fields("Id_01")),
    '                .Id_02 = RsNullDB(Rs.Fields("Id_02")),
    '                .Id_03 = RsNullDB(Rs.Fields("Id_03")),
    '                .Id_04 = RsNullDB(Rs.Fields("Id_04")),
    '                .Id_05 = RsNullDB(Rs.Fields("Id_05")),
    '                .Id_06 = RsNullDB(Rs.Fields("Id_06")),
    '                .Id_07 = RsNullDB(Rs.Fields("Id_07")),
    '                .Id_08 = RsNullDB(Rs.Fields("Id_08")),
    '                .Id_09 = RsNullDB(Rs.Fields("Id_09")),
    '                .Id_10 = RsNullDB(Rs.Fields("Id_10")),
    '                .Str_5_01 = RsNullDB(Rs.Fields("Str_5_01")),
    '                .Str_5_02 = RsNullDB(Rs.Fields("Str_5_02")),
    '                .Str_5_03 = RsNullDB(Rs.Fields("Str_5_03")),
    '                .Str_5_04 = RsNullDB(Rs.Fields("Str_5_04")),
    '                .Str_5_05 = RsNullDB(Rs.Fields("Str_5_05")),
    '                .Str_25_01 = RsNullDB(Rs.Fields("Str_25_01")),
    '                .Str_25_02 = RsNullDB(Rs.Fields("Str_25_02")),
    '                .Str_25_03 = RsNullDB(Rs.Fields("Str_25_03")),
    '                .Str_25_04 = RsNullDB(Rs.Fields("Str_25_04")),
    '                .Str_25_05 = RsNullDB(Rs.Fields("Str_25_05")),
    '                .Str_50_01 = RsNullDB(Rs.Fields("Str_50_01")),
    '                .Str_50_02 = RsNullDB(Rs.Fields("Str_50_02")),
    '                .Str_50_03 = RsNullDB(Rs.Fields("Str_50_03")),
    '                .Str_50_04 = RsNullDB(Rs.Fields("Str_50_04")),
    '                .Str_50_05 = RsNullDB(Rs.Fields("Str_50_05")),
    '                .Str_50_06 = RsNullDB(Rs.Fields("Str_50_06")),
    '                .Str_50_07 = RsNullDB(Rs.Fields("Str_50_07")),
    '                .Str_50_08 = RsNullDB(Rs.Fields("Str_50_08")),
    '                .Str_50_09 = RsNullDB(Rs.Fields("Str_50_09")),
    '                .Str_50_10 = RsNullDB(Rs.Fields("Str_50_10")),
    '                .Str_150_01 = RsNullDB(Rs.Fields("Str_150_01")),
    '                .Str_150_02 = RsNullDB(Rs.Fields("Str_150_02")),
    '                .Str_150_03 = RsNullDB(Rs.Fields("Str_150_03")),
    '                .Str_150_04 = RsNullDB(Rs.Fields("Str_150_04")),
    '                .Str_150_05 = RsNullDB(Rs.Fields("Str_150_05")),
    '                .Str_250_01 = RsNullDB(Rs.Fields("Str_250_01")),
    '                .Str_250_02 = RsNullDB(Rs.Fields("Str_250_02")),
    '                .Str_250_03 = RsNullDB(Rs.Fields("Str_250_03")),
    '                .Str_250_04 = RsNullDB(Rs.Fields("Str_250_04")),
    '                .Str_250_05 = RsNullDB(Rs.Fields("Str_250_05")),
    '                .Str_500_01 = RsNullDB(Rs.Fields("Str_500_01")),
    '                .Str_500_02 = RsNullDB(Rs.Fields("Str_500_02")),
    '                .Str_500_03 = RsNullDB(Rs.Fields("Str_500_03")),
    '                .Str_500_04 = RsNullDB(Rs.Fields("Str_500_04")),
    '                .Str_500_05 = RsNullDB(Rs.Fields("Str_500_05")),
    '                .Str_Max_01 = RsNullDB(Rs.Fields("Str_Max_01")),
    '                .Str_Max_02 = RsNullDB(Rs.Fields("Str_Max_02")),
    '                .Str_Max_03 = RsNullDB(Rs.Fields("Str_Max_03")),
    '                .Str_Max_04 = RsNullDB(Rs.Fields("Str_Max_04")),
    '                .Str_Max_05 = RsNullDB(Rs.Fields("Str_Max_05")),
    '                .DateTime_01 = RsNullDT(Rs.Fields("DateTime_01")),
    '                .DateTime_02 = RsNullDT(Rs.Fields("DateTime_02")),
    '                .DateTime_03 = RsNullDT(Rs.Fields("DateTime_03")),
    '                .DateTime_04 = RsNullDT(Rs.Fields("DateTime_04")),
    '                .DateTime_05 = RsNullDT(Rs.Fields("DateTime_05")),
    '                .Float_01 = RsNullDB(Rs.Fields("Float_01")),
    '                .Float_02 = RsNullDB(Rs.Fields("Float_02")),
    '                .Float_03 = RsNullDB(Rs.Fields("Float_03")),
    '                .Float_04 = RsNullDB(Rs.Fields("Float_04")),
    '                .Float_05 = RsNullDB(Rs.Fields("Float_05")),
    '                .Float_06 = RsNullDB(Rs.Fields("Float_06")),
    '                .Float_07 = RsNullDB(Rs.Fields("Float_07")),
    '                .Float_08 = RsNullDB(Rs.Fields("Float_08")),
    '                .Float_09 = RsNullDB(Rs.Fields("Float_09")),
    '                .Float_10 = RsNullDB(Rs.Fields("Float_10")),
    '                .Int_01 = RsNullDB(Rs.Fields("Int_01")),
    '                .Int_02 = RsNullDB(Rs.Fields("Int_02")),
    '                .Int_03 = RsNullDB(Rs.Fields("Int_03")),
    '                .Int_04 = RsNullDB(Rs.Fields("Int_04")),
    '                .Int_05 = RsNullDB(Rs.Fields("Int_05")),
    '                .num_180_01 = RsNullDB(Rs.Fields("num_180_01")),
    '                .num_180_02 = RsNullDB(Rs.Fields("num_180_02")),
    '                .num_180_03 = RsNullDB(Rs.Fields("num_180_03")),
    '                .num_180_04 = RsNullDB(Rs.Fields("num_180_04")),
    '                .num_180_05 = RsNullDB(Rs.Fields("num_180_05")),
    '                .bit_01 = RsNullDB(Rs.Fields("bit_01")),
    '                .bit_02 = RsNullDB(Rs.Fields("bit_02")),
    '                .bit_03 = RsNullDB(Rs.Fields("bit_03")),
    '                .bit_04 = RsNullDB(Rs.Fields("bit_04")),
    '                .bit_05 = RsNullDB(Rs.Fields("bit_05")),
    '                .FileType = RsNullDB(Rs.Fields("FileType")),
    '                .FileDisposition = RsNullDB(Rs.Fields("FileDisposition")),
    '                .FileLength = RsNullDB(Rs.Fields("FileLength")),
    '                .FileName = RsNullDB(Rs.Fields("FileName")),
    '                .createBy = RsNullDB(Rs.Fields("createBy")),
    '                .createDate = RsNullDT(Rs.Fields("createDate")),
    '                .modifyBy = RsNullDB(Rs.Fields("modifyBy")),
    '                .modifyDate = RsNullDT(Rs.Fields("modifyDate")),
    '                .Inf_Status = RsNullDB(Rs.Fields("Inf_Status")),
    '                .Inf_Datetime = RsNullDT(Rs.Fields("Inf_Datetime"))
    '            }


    '                listTB.Add(iOffline)

    '                Rs.MoveNext()

    '            Loop

    '            Dim Url = ""
    '            Dim irow As DataRow = tbConfig.[Select](" keyType = 'URL'and keyWord = 'APItoServer' ").FirstOrDefault()
    '            If IsDBNull(irow) = False Then Url = strNull(irow.Item("config"))
    '            Dim hostName As String = GetURLHostName(Url)
    '            Dim uriName As String = Replace(Url, hostName, "")



    '            Dim data_client = New RestClient(hostName)
    '            Dim request = New RestRequest(uriName, Method.Post)
    '            'request.AddHeader("Accept", "text/xml")
    '            request.AddHeader("Accept", "application/json")
    '            request.RequestFormat = DataFormat.Json
    '            request.AddJsonBody(listTB2)

    '            'request.AddHeader("Authorization", "bearer " + bearer)
    '            Dim data_response = data_client.Execute(request)
    '            Dim data_response_raw = data_response.Content



    '            If (data_response.IsSuccessful) Then
    '                If Rs.RecordCount <> 0 Then
    '                    Dim RsUp As New ADODB.Recordset
    '                    RsUp = WriteSQL("exec spInterfaceClient2Server_runComplete '" & entityId & "','" & RunDateTime & "'", ConnLoc)
    '                    RsClose(RsUp)
    '                End If

    '            Else

    '                Dim lines() As String = {"Error asynchronous Client2Server : " + DateTime.Now.ToString()}
    '                System.IO.File.AppendAllLines(LogPath, lines)

    '            End If

    '            RsClose(Rs)


    '            Return True

    '            'Dim dt As New DataTable
    '            'Dim adapter As OleDb.OleDbDataAdapter = New OleDb.OleDbDataAdapter
    '            'adapter.Fill(dt, Rs)
    '            'Public Function GetJson(ByVal dt As DataTable) As String
    '            'Dim strJSon As String = New JavaScriptSerializer().Serialize(From dr As DataRow In dt.Rows Select dt.Columns.Cast(Of DataColumn)().ToDictionary(Function(col) col.ColumnName, Function(col) dr(col)))
    '            'End Function
    '            'HTTPRestPOST(strJSon, "https://intranet.nathalin.com/VRM_UAT/api/client2Server")
    '            'SendRequest(strJSon, "https://intranet.nathalin.com/VRM_UAT/api/client2Server")
    '            'send1(strJSon, "https://localhost:44395/api/offline")


    '            'Dim RsSrv As New ADODB.Recordset
    '            'Dim RsLoc As New ADODB.Recordset
    '            'If Rs.RecordCount <> 0 Then

    '            '    RsSrv = WriteSQL("select * from tbInterfaceOfflineRecord where Id is null", ConnSrv)
    '            '    Do Until Rs.EOF

    '            '        RsSrv.AddNew()
    '            '        For li = 0 To Rs.Fields.Count - 1
    '            '            For si = 0 To RsSrv.Fields.Count
    '            '                If LCase(RsSrv.Fields(si).Name) = LCase(Rs.Fields(li).Name) Then
    '            '                    RsSrv.Fields(si).Value = Rs.Fields(li).Value
    '            '                    Exit For
    '            '                End If
    '            '            Next
    '            '        Next
    '            '        RsSrv.Update()

    '            '        Dim sqlRecId = "and RefId = " & Guid2Str(zSQLText(strNull(Rs.Fields("RecId").Value.ToString())))
    '            '        Dim sqlKeyAction = "and KeySub = " & zSQLText(strNull(Rs.Fields("KeyAction").Value))
    '            '        Dim sqlKeyRecord = "and KeyWord = " & zSQLText(strNull(Rs.Fields("KeyRecord").Value))


    '            '        Sql = "update tbInterfaceOfflineCheck set status = 'Close' where not(refId is null)  " & sqlRecId & sqlKeyAction & sqlKeyRecord
    '            '        RsLoc = WriteSQL(Sql, ConnLoc)
    '            '        RsClose(RsLoc)

    '            '        'Dim ConnSQL_Local = New ADODB.Connection
    '            '        'Dim ConnSQL_Server = New ADODB.Connection
    '            '        'rsLoc = WriteSQL("delete from tbFAAsset ", ConnSQL_Local)



    '            '        'ConnSQL_Local.BeginTrans()
    '            '        'Sql = "insert tbSystemLog select newid() , 'gab' , 'gab' , getdate()"
    '            '        'ConnSQL_Local.Execute(Sql)
    '            '        'ConnSQL_Local.CommitTrans()


    '            '        Rs.MoveNext()
    '            '    Loop
    '            '    RsClose(RsSrv)

    '            'End If

    '            'RsClose(Rs)

    '        End Function








    '        Public Function InterfaceClient2Server1(ConnLoc As ADODB.Connection, ConnSrv As ADODB.Connection)
    '            Dim Rs As New ADODB.Recordset
    '            Dim RsLoc As New ADODB.Recordset
    '            Dim RsSrv = New ADODB.Recordset

    '            Rs = ReadSQL("exec spInterfaceClient2Server '" & entityId & "'", ConnLoc)



    '            Dim dt As New DataTable
    '            Dim adapter As OleDb.OleDbDataAdapter = New OleDb.OleDbDataAdapter
    '            adapter.Fill(dt, Rs)
    '            'Public Function GetJson(ByVal dt As DataTable) As String
    '            Dim strJSon As String = New JavaScriptSerializer().Serialize(From dr As DataRow In dt.Rows Select dt.Columns.Cast(Of DataColumn)().ToDictionary(Function(col) col.ColumnName, Function(col) dr(col)))
    '            'End Function
    '            'HTTPRestPOST(strJSon, "https://intranet.nathalin.com/VRM_UAT/api/client2Server")
    '            'SendRequest(strJSon, "https://intranet.nathalin.com/VRM_UAT/api/client2Server")

    '            'send1(strJSon, "https://localhost:44395/api/offline")
    '            sendOff(strJSon, "https://localhost:44395/offline/client2Server")


    '            Return True
    '            If Rs.RecordCount <> 0 Then

    '                RsSrv = WriteSQL("select * from tbInterfaceOfflineRecord where Id is null", ConnSrv)
    '                Do Until Rs.EOF

    '                    RsSrv.AddNew()
    '                    For li = 0 To Rs.Fields.Count - 1
    '                        For si = 0 To RsSrv.Fields.Count
    '                            If LCase(RsSrv.Fields(si).Name) = LCase(Rs.Fields(li).Name) Then
    '                                RsSrv.Fields(si).Value = Rs.Fields(li).Value
    '                                Exit For
    '                            End If
    '                        Next
    '                    Next
    '                    RsSrv.Update()

    '                    Dim sqlRecId = "and RefId = " & Guid2Str(zSQLText(strNull(Rs.Fields("RecId").Value.ToString())))
    '                    Dim sqlKeyAction = "and KeySub = " & zSQLText(strNull(Rs.Fields("KeyAction").Value))
    '                    Dim sqlKeyRecord = "and KeyWord = " & zSQLText(strNull(Rs.Fields("KeyRecord").Value))


    '                    Sql = "update tbInterfaceOfflineCheck set status = 'Close' where not(refId is null)  " & sqlRecId & sqlKeyAction & sqlKeyRecord
    '                    RsLoc = WriteSQL(Sql, ConnLoc)
    '                    RsClose(RsLoc)

    '                    'Dim ConnSQL_Local = New ADODB.Connection
    '                    'Dim ConnSQL_Server = New ADODB.Connection
    '                    'rsLoc = WriteSQL("delete from tbFAAsset ", ConnSQL_Local)


    '                    'ConnSQL_Local.BeginTrans()
    '                    'Sql = "insert tbSystemLog select newid() , 'gab' , 'gab' , getdate()"
    '                    'ConnSQL_Local.Execute(Sql)
    '                    'ConnSQL_Local.CommitTrans()


    '                    Rs.MoveNext()
    '                Loop
    '                RsClose(RsSrv)

    '            End If

    '            RsClose(Rs)

    '        End Function


    '        Function sendOff(JsonInputStr As String, ByVal POSTUri As String) As String

    '            'Dim postData As String = Left(JsonInputStr, Len(JsonInputStr) - 1)
    '            'postData = Right(postData, Len(postData) - 1)


    '            'Dim sendUri As String = POSTUri & "?recCall=" & JsonInputStr

    '            'Dim postData As String = JsonInputStr

    '            'Dim request As WebRequest = WebRequest.Create(POSTUri)
    '            'request.Method = "POST"


    '            Dim data_client = New RestClient("https://localhost:44395")

    '            Dim request = New RestRequest("/api/offline", Method.Post)
    '            'request.AddHeader("Accept", "text/xml")

    '            request.AddHeader("Accept", "application/json")
    '            request.RequestFormat = DataFormat.Json

    '            'Dim aaa = New tbInterfaceOffline With {.KeyAction = "A"}

    '            Dim lst = New List(Of Object)

    '            Dim bbb = New tbInterfaceOffline


    '            Dim aaa = New With {
    '            .Id = Guid.NewGuid,
    '            .KeyAction = "aaa",
    '            .KeyRecord = "bbb"
    '            }

    '            lst.Add(aaa)

    '            'request.AddJsonBody(New With {.KeyAction = "test1"})
    '            request.AddJsonBody(lst)


    '            '

    '            'arr(1)



    '            'KeyAction
    '            'Dim dt 
    '            'set prm As New Parameter
    '            ''prm.ContentType = ""

    '            'prm.
    '            ''prm. = ""
    '            'request.AddParameter(New Parameter())

    '            'request.AddParameter("recCall", "test")


    '            'request.AddParameter("reCall", "test",, ParameterType.RequestBody)
    '            'Dim prm As New Parameter

    '            'prm.ContentType = "text/xml"
    '            'prm.ContentType = "text/xml"



    '            'request.Parameters.AddParameter()
    '            '        ( "recCall", JsonInputStr, ParameterType.RequestBody)




    '            'request.AddJsonBody(JsonInputStr)

    '            'request.AddHeader("Authorization", "bearer " + bearer)
    '            Dim data_response = data_client.Execute(request)
    '            Dim data_response_raw = data_response.Content

    '            'Dim request = New RestRequest("/offline/", Method.Post)
    '            'request.RequestFormat = DataFormat.Json
    '            'request.AddJsonBody(JsonInputStr)

    '            'Dim response = _client.ExecuteAsync(request)


    '            If (data_response.IsSuccessful) Then
    '                '            {
    '                '    Return response.Data;

    '            Else
    '                '{
    '                '    Return New ErpServiceResult
    '                '    {
    '                '        Result = Message.Error,
    '                '        Message = response.ErrorMessage
    '                '    };
    '            End If


    '            'request.Headers.Add("async", "true")
    '            'Dim byteArray As Byte() = Encoding.UTF8.GetBytes(dbTB)
    '            'request.ContentType = "application/json; charset=utf-8"
    '            'request.ContentLength = byteArray.Length
    '            'Dim dataStream As Stream = request.GetRequestStream()
    '            'dataStream.Write(byteArray, 0, byteArray.Length)
    '            'dataStream.Close()

    '            'Dim response As WebResponse = request.GetResponse()
    '            'Console.WriteLine(CType(response, HttpWebResponse).StatusDescription)
    '            'dataStream = response.GetResponseStream()
    '            'Dim reader As New StreamReader(dataStream)
    '            'Dim responseFromServer As String = reader.ReadToEnd()
    '            'Console.WriteLine(responseFromServer)
    '            'reader.Close()
    '            'dataStream.Close()
    '            'response.Close()

    '        End Function





    '        Function send1(JsonInputStr As String, ByVal POSTUri As String) As String
    '            'Dim postData As String = Left(JsonInputStr, Len(JsonInputStr) - 1)
    '            'postData = Right(postData, Len(postData) - 1)


    '            Dim sendUri As String = POSTUri & "?recCall=" & JsonInputStr

    '            Dim postData As String = JsonInputStr

    '            Dim request As WebRequest = WebRequest.Create(POSTUri)
    '            request.Method = "POST"



    '            request.Headers.Add("async", "true")
    '            Dim byteArray As Byte() = Encoding.UTF8.GetBytes(postData)
    '            request.ContentType = "application/json; charset=utf-8"
    '            request.ContentLength = byteArray.Length
    '            Dim dataStream As Stream = request.GetRequestStream()
    '            dataStream.Write(byteArray, 0, byteArray.Length)
    '            dataStream.Close()

    '            Dim response As WebResponse = request.GetResponse()
    '            Console.WriteLine(CType(response, HttpWebResponse).StatusDescription)
    '            dataStream = response.GetResponseStream()
    '            Dim reader As New StreamReader(dataStream)
    '            Dim responseFromServer As String = reader.ReadToEnd()
    '            Console.WriteLine(responseFromServer)
    '            reader.Close()
    '            dataStream.Close()
    '            response.Close()

    '        End Function


    '        '    Private Sub GetRequestStreamCallback(ByVal asynchronousResult As IAsyncResult)
    '        '        Dim request As HttpWebRequest = CType(asynchronousResult.AsyncState, HttpWebRequest)

    '        '        Dim postStream As Stream = request.EndGetRequestStream(asynchronousResult)
    '        '        Dim postData As String = "Gab"
    '        '        Dim byteArray As Byte() = System.Text.Encoding.UTF8.GetBytes(postData)

    '        '    postStream.Write(byteArray, 0, postData.Length)
    '        '    postStream.Close()

    '        '    Dim result As IAsyncResult = CType(request.BeginGetResponse(AddressOf GetResponseCallback, request),  _
    '        '        IAsyncResult)
    '        'End Sub

    '        '    Private Sub GetResponseCallback(ByVal asynchronousResult As IAsyncResult)
    '        '        Dim request As HttpWebRequest = CType(asynchronousResult.AsyncState, HttpWebRequest)
    '        '        Dim response As HttpWebResponse = CType(request.EndGetResponse(asynchronousResult), HttpWebResponse)

    '        '        Dim streamResponse As Stream = response.GetResponseStream()
    '        '        Dim streamRead As New StreamReader(streamResponse)
    '        '        Dim responseString As String = streamRead.ReadToEnd()

    '        '        '<<<<RETURING THE RECEIVED DATA TO THE FORM WHICH CALLED GETTING HTML FUNC.>>>>

    '        '        streamResponse.Close()
    '        '        streamRead.Close()

    '        '        response.Close()

    '        '    End Sub

    '        'Public Sub GetHTMLAsync(ByVal POSTDATA As String,  <<<<FUNCTION ADDRESS Or SUCH THING TO CALL WHEN THE ASYNC PROCEDURE Is DONE>>>>)

    '        '    Req = CType(WebRequest.Create("http://hi.asdf.com/getinformation.php"), HttpWebRequest)
    '        '    Req.Method = "POST"
    '        '    Req.ContentType = "application/x-www-form-urlencoded; charset=utf-8"
    '        '    Req.CookieContainer = CookieC
    '        '    Req.Timeout = 1000 * 30

    '        '    '<DO SOMETHING To DELIVER [POSTDATA] And ["FUNCTION" ARGUMENT] To ASYNC PROCEDURE ABOVE>

    '        '    'Async
    '        '    Dim result As IAsyncResult = CType(Req.BeginGetRequestStream(AddressOf GetRequestStreamCallback, Req), IAsyncResult)

    '        'End Sub

    '        Private Function SendRequest(JsonInputStr As String, ByVal POSTUri As String) As String
    '            Dim response As String
    '            Dim request As WebRequest
    '            Dim encoding As New UTF8Encoding() 'As New System.Text.ASCIIEncoding
    '            Dim uri = POSTUri
    '            Dim Tran = "15"
    '            Dim Amount As Integer = 1000
    '            Dim ReferenceID = "015bfa15-15ec-4dc7-903c-053ffacb6688"
    '            Dim POS = "57290070"
    '            Dim Store = "RESTSIM00000001"
    '            Dim Chain = "J@P-Reg"

    '            Dim JSONString = "{""tran"":""" & Tran & """,""amount"":""" & Amount & """,""reference"":""" & ReferenceID & """,""pos"":""" & POS & """,""store"":""" & Store & """,""chain"":""" & Chain & """}"
    '            Dim jsonDataBytes() As Byte = encoding.GetBytes(JsonInputStr)


    '            request = WebRequest.Create(uri)
    '            request.ContentLength = jsonDataBytes.Length
    '            request.ContentType = "application/json"
    '            request.Method = "POST"

    '            'Using requestStream = request.GetRequestStream
    '            '    requestStream.Write(jsonDataBytes, 0, jsonDataBytes.Length)
    '            '    requestStream.Close()

    '            '    Using responseStream = request.GetResponse.GetResponseStream
    '            '        Using reader As New StreamReader(responseStream)
    '            '            response = reader.ReadToEnd()
    '            '        End Using
    '            '    End Using
    '            'End Using

    '            Return response
    '        End Function

    '        Function DataSetToJSON(ds As DataSet) As String
    '            Dim dict As New Dictionary(Of String, Object)

    '            For Each dt As DataTable In ds.Tables
    '                Dim arr(dt.Rows.Count) As Object

    '                For i As Integer = 0 To dt.Rows.Count - 1
    '                    arr(i) = dt.Rows(i).ItemArray
    '                Next

    '                dict.Add(dt.TableName, arr)
    '            Next

    '            Dim json As New JavaScriptSerializer
    '            Return json.Serialize(dict)
    '        End Function



    '        Private Sub HTTPRestPOST(ByVal JsonInputStr As String, ByVal POSTUri As String)

    '            'Make a request to the POST URI

    '            Dim RestPOSTRequest As HttpWebRequest = HttpWebRequest.Create(POSTUri)

    '            'Convert the JSON Input to Bytes through UTF8 Encoding

    '            Dim JsonEncoding As New UTF8Encoding()

    '            Dim JsonBytes As Byte() = JsonEncoding.GetBytes(JsonInputStr)

    '            'Setting the request parameters

    '            RestPOSTRequest.Method = "POST"

    '            RestPOSTRequest.ContentType = "application/json"

    '            RestPOSTRequest.ContentLength = JsonBytes.Length

    '            'Add any other Headers for the URI

    '            'RestPOSTRequest.Headers.Add("username", "kalyan_nakka")

    '            'RestPOSTRequest.Headers.Add("password", "********")

    '            'RestPOSTRequest.Headers.Add("urikey", "MAIJHDAS54ADAJQA35IJHA784R98AJN")

    '            'Create the Input Stream for the URI

    '            Using RestPOSTRequestStream As Stream = RestPOSTRequest.GetRequestStream()

    '                'Write the Input JSON data into the Stream

    '                RestPOSTRequestStream.Write(JsonBytes, 0, JsonBytes.Length)

    '                'Response from the URI

    '                Dim RestPOSTResponse = RestPOSTRequest.GetResponse()

    '                'Create Stream for the response

    '                Using RestPOSTResponseStream As Stream = RestPOSTResponse.GetResponseStream()

    '                    'Create a Reader for the Response Stream

    '                    Using RestPOSTResponseStreamReader As New StreamReader(RestPOSTResponseStream)

    '                        Dim ResponseData = RestPOSTResponseStreamReader.ReadToEnd()

    '                        'Later utilize "ResponseData" variable as per your requirement

    '                        'Close the Reader

    '                        RestPOSTResponseStreamReader.Close()

    '                    End Using

    '                    RestPOSTResponseStream.Close()

    '                End Using

    '                RestPOSTRequestStream.Close()

    '            End Using

    '        End Sub




    '        Public Function UpdateDOA(ConnLoc As ADODB.Connection, ConnSrv As ADODB.Connection) As Boolean
    '            Dim Ret = False
    '            Try

    '                Dim rows As DataRow()
    '                Dim Sql As String
    '                Dim DOAUrl = ""
    '                Dim DOACode = ""
    '                rows = tbConfig.[Select](" keyType = 'DOA'  and keyWord = 'URL'")
    '                For Each row As DataRow In rows
    '                    DOAUrl = strNull(row.Item("config"))
    '                Next

    '                rows = tbConfig.[Select](" keyType = 'DOA' and keyWord like 'WorkFlowCode%'  ")
    '                For Each row As DataRow In rows
    '                    DOACode = strNull(row.Item("config"))
    '                    If DOACode = "" Then GoTo JmpNextLoop
    '                    Dim pramSQL As String = "code=" + DOACode + "&bcGroup=" + grpCode + "&bcBusinessUnit=" + grpCode + "&bcCompanyCode=" + companyCode + "&bcEntity=" + entityCode

    '                    If Right(DOAUrl, 1) <> "?" Then DOAUrl = DOAUrl + "?"

    '                    Dim URLGetDOA As String = DOAUrl + pramSQL

    '                    Dim request2 As HttpWebRequest = HttpWebRequest.Create(URLGetDOA)

    '                    request2.Method = "GET"
    '                    request2.ContentType = "application/x-www-form-urlencoded"
    '                    request2.KeepAlive = True
    '                    request2.ContinueTimeout = 10000
    '                    request2.Host = "intranet.nathalin.com"

    '                    Dim Header2Collection As WebHeaderCollection = request2.Headers
    '                    'Dim BearerAccessToken = "Bearer " & access_token
    '                    'Header2Collection.Add("Authorization", BearerAccessToken)
    '                    Header2Collection.Add("x-myobapi-key", "MYAPIKEY")
    '                    Header2Collection.Add("x-myobapi-version", "v2")
    '                    Header2Collection.Add("Accept-Encoding", "gzip,deflate")
    '                    Header2Collection.Add("Cache-Control", "no-cache")
    '                    Dim postBytes2 As Byte() = New UTF8Encoding().GetBytes(request2.ToString())

    '                    Dim response2 As HttpWebResponse = CType(request2.GetResponse(), HttpWebResponse)
    '                    Dim receiveStream2 As Stream = response2.GetResponseStream()
    '                    Dim readStream2 As New StreamReader(receiveStream2)
    '                    Dim jsonstring = readStream2.ReadToEnd()

    '                    If InStr(jsonstring, "The condition which is not matching") = 0 Then 'no Error
    '                        If jsonstring <> "" Then
    '                            Dim RsU As New ADODB.Recordset
    '                            Sql = "update tbSystemOfflineConfig set config = '" & jsonstring & "' where keytype = 'DOA' and keyword = '" + DOACode + "'"
    '                            RsU = WriteSQL(Sql, ConnLoc)
    '                            RsClose(RsU)
    '                        End If

    '                        Ret = True
    '                    End If

    'JmpNextLoop:
    '                Next
    '                Return Ret

    '            Catch ex As Exception
    '                Return Ret

    '            End Try



    '        End Function





    '    End Class




End Module
