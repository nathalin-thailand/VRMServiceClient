Imports System.Data.SqlClient
Imports System.IO
Imports System.Net
Imports System.Text
Imports System.Web.Script.Serialization
Imports Newtonsoft.Json
Imports RestSharp

Public Class VRMService


    Dim GetDOAOnline As Boolean = False
    Dim StatusRun As Boolean = False


    Dim tbOfflineInteface As New DataTable()

    Public Sub OnDebug()
        Me.OnStart(Nothing)
    End Sub


    Protected Overrides Sub OnStart(ByVal args() As String)
        On Error GoTo ErrPoint

        Dim lines() As String = {"Starting time : " + DateTime.Now.ToString()}
        System.IO.File.AppendAllLines(LogPath, lines)

        If System.IO.File.Exists(ConfigFile) = False Then
            Dim strWrite As String = "TimeMin=" & TimeInteval.ToString() '& vbCrLf & "database=" & vbCrLf & "user=" & vbCrLf & "password="
            System.IO.File.WriteAllText(ConfigFile, strWrite)

        Else
            lines = System.IO.File.ReadAllLines(ConfigFile)
            For zi = 0 To lines.Length - 1
                Dim linStr = lines(zi)
                Dim arrStr = lines(zi).Split("=")
                If (arrStr.Length = 2) Then
                    If LCase(arrStr(0)) = LCase("TimeMin") Then TimeInteval = Convert.ToDouble(arrStr(1))

                End If
            Next
        End If

        Timer1.Interval = Convert.ToInt32(TimeInteval * (1000 * 60))
        Timer1.Enabled = True
        StatusRun = False
        Exit Sub

ErrPoint:
        lines = {Err.Description + " : " + DateTime.Now.ToString()}
        System.IO.File.AppendAllLines(LogPath, lines)
    End Sub

    Protected Overrides Sub OnStop()
        Dim lines() As String = {"Stop time : " + DateTime.Now.ToString()}
        System.IO.File.AppendAllLines(LogPath, lines)
    End Sub

    Private Sub Timer1_Elapsed(sender As Object, e As Timers.ElapsedEventArgs) Handles Timer1.Elapsed
        If StatusRun Then Exit Sub

        On Error GoTo ErrPoint

        StatusRun = True

        Dim ErrStr As String
        Dim lines() As String

        'Dim lines() As String = {"Run time : " + DateTime.Now.ToString()}
        'System.IO.File.AppendAllLines(LogPath, lines)

        If System.IO.File.Exists(SQLConfigFile) = False Then
            Dim strWrite As String = "server=" & vbCrLf & "database=" & vbCrLf & "user=" & vbCrLf & "password="
            System.IO.File.WriteAllText(SQLConfigFile, strWrite)
        Else
            lines = System.IO.File.ReadAllLines(SQLConfigFile)
            For zi = 0 To lines.Length - 1
                Dim linStr = lines(zi)
                Dim arrStr = lines(zi).Split("=")
                If (arrStr.Length = 2) Then
                    If LCase(arrStr(0)) = "server" Then locServer = arrStr(1)
                    If LCase(arrStr(0)) = "database" Then locDB = arrStr(1)
                    If LCase(arrStr(0)) = "user" Then locUID = arrStr(1)
                    If LCase(arrStr(0)) = "password" Then locPWD = arrStr(1)
                End If
            Next
        End If


        Dim ConnLoc As New ADODB.Connection

        Dim locConStr = "Driver={SQL Server};Server=" & locServer & ";Database=" & locDB & ";UID=" & locUID & ";PWD=" & locPWD & ";"
        ConnLoc.ConnectionString = locConStr
        ConnLoc.CursorLocation = ADODB.CursorLocationEnum.adUseClient
        ConnLoc.CommandTimeout = 1500
        ConnLoc.Open()


        tbConfig.Columns.Clear()

        tbConfig.Columns.AddRange(New DataColumn() {
            New DataColumn("keyType", GetType(String)),
            New DataColumn("keyWord", GetType(String)),
            New DataColumn("config", GetType(String))
         })

        tbConfig.Rows.Clear()

        Dim Rs As New ADODB.Recordset
        Rs = ReadSQL("select * from tbSystemOfflineConfig", ConnLoc)

        Dim zStrItem(Rs.Fields.Count - 1) As String
        Do Until Rs.EOF
            'tbConfig.Rows.Add(Rs)

            zStrItem(0) = Rs.Fields("keyType").Value
            zStrItem(1) = Rs.Fields("keyWord").Value
            zStrItem(2) = Rs.Fields("config").Value
            tbConfig.Rows.Add(zStrItem)
            Rs.MoveNext()
        Loop
        RsClose(Rs)


        'WriteLog(ConnLoc, "Run", "Connect", "OK")
        Dim ConnSrv As New ADODB.Connection
        If Not IsNothing(ConnLoc.ConnectionString) Then
            'ConnSrv = ConnectServer(ConnLoc)
            'If IsNothing(ConnSrv) Then GoTo JmpOut
            If ConnectServerAPI() Then
                SycData(ConnLoc, ConnSrv)
            End If

            Call triggerLocalRun(ConnLoc)

        End If

        GoTo JmpOut

ErrPoint:
        'MsgBox(Err.Description)

        Dim line1s() As String = {Err.Description + " : " + DateTime.Now.ToString()}
        System.IO.File.AppendAllLines(LogPath, line1s)
        Err.Clear()

JmpOut:
        StatusRun = False
    End Sub

    Public Sub SycData(ConnLoc As ADODB.Connection, ConnSrv As ADODB.Connection)

        System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.CurrentCulture
        WriteLog(ConnLoc, "Run", "SycData", "")

        entityId = "0"
        Dim rows = tbConfig.[Select](" keyType = 'Machine' and keyWord = 'Entity' ")
        For Each row As DataRow In rows
            entityId = strNull(row.Item("config"))
        Next

        If entityId = "" Then entityId = "0"
        If entityId = "0" Then Exit Sub

        Sql = "select * from vw_org_entities where entityid = " & entityId
        Dim Rs As New ADODB.Recordset
        Rs = ReadSQL(Sql, ConnLoc)
        If Rs.RecordCount <> 0 Then
            corpId = strNull(Rs.Fields("corpId").Value)
            grpId = strNull(Rs.Fields("grpId").Value)
            buId = strNull(Rs.Fields("buId").Value)
            companyId = strNull(Rs.Fields("companyId").Value)
            entityId = strNull(Rs.Fields("entityId").Value)
            corpCode = strNull(Rs.Fields("corpCode").Value)
            grpCode = strNull(Rs.Fields("grpCode").Value)
            buCode = strNull(Rs.Fields("buCode").Value)
            companyCode = strNull(Rs.Fields("companyCode").Value)
            entityCode = strNull(Rs.Fields("entityCode").Value)
        End If
        RsClose(Rs)

        rows = tbConfig.[Select](" keyType = 'DOA' and keyWord = 'GetOnline' ")
        For Each row As DataRow In rows
            GetDOAOnline = LCase(strNull(row.Item("config"))) = "yes"
        Next


        If GetDOAOnline Then UpdateDOA(ConnLoc, ConnSrv)


        If InterfaceClient2Server(ConnLoc, ConnSrv) Then
            Call InterfaceTriggerServer()
        End If
        Call InterfaceServer2Client(ConnLoc, ConnSrv)
    End Sub


    Public Function InterfaceServer2Client(ConnLoc As ADODB.Connection, ConnSrv As ADODB.Connection)

        Dim RunDateTime As String = Now.ToString("yyyy-MMM-dd  HH:mm:ss")

        Dim Url = ""
        Dim irow As DataRow = tbConfig.[Select](" keyType = 'URL'and keyWord = 'APIGetServer' ").FirstOrDefault()
        If IsDBNull(irow) = False Then Url = strNull(irow.Item("config"))

        Dim pram = "?getEntity=" & entityId & "&keyDateTime=" & RunDateTime

        Dim hostName As String = GetURLHostName(Url)
        Dim uriName As String = Replace(Url, hostName, "")

        Dim data_client = New RestClient(hostName)
        Dim request = New RestRequest(uriName & pram, Method.Get)
        request.AddHeader("Accept", "text/xml")

        'request.AddHeader("Accept", "application/json")
        'request.RequestFormat = DataFormat.
        'request.AddJsonBody(listTB)
        'request.RequestFormat = DataFormat.Json
        'request.AddJsonBody(listTB)

        'request.AddHeader("Authorization", "bearer " + bearer)
        Dim data_response = data_client.Execute(request)
        Dim data_response_raw = data_response.Content

        If (data_response.IsSuccessful) Then

            Dim strJson As String = data_response_raw
            If (strJson <> "[]") Then
                Dim RsJsUP As New ADODB.Recordset
                Dim JSql As String = "exec spInterfaceClient2Server_JSonInsert '" & strJson & "'"
                RsJsUP = ReadSQL(JSql, ConnLoc)
                RsClose(RsJsUP)
            End If
            Return True

        Else

                Dim lines() As String = {"Error asynchronous Server2Client : " + DateTime.Now.ToString()}
            System.IO.File.AppendAllLines(LogPath, lines)
            Return False
        End If

    End Function



    Public Function InterfaceTriggerServer()

        Dim RunDateTime As String = Now.ToString("yyyy-MMM-dd  HH:mm:ss")

        Dim Url = ""
        Dim irow As DataRow = tbConfig.[Select](" keyType = 'URL'and keyWord = 'APITriggerServer' ").FirstOrDefault()
        If IsDBNull(irow) = False Then Url = strNull(irow.Item("config"))

        Dim pram = "?getEntity=" & entityId & "&keyDateTime=" & RunDateTime

        Dim hostName As String = GetURLHostName(Url)
        Dim uriName As String = Replace(Url, hostName, "")

        Dim data_client = New RestClient(hostName)
        Dim request = New RestRequest(uriName & pram, Method.Get)
        request.AddHeader("Accept", "text/xml")

        'request.AddHeader("Accept", "application/json")
        'request.RequestFormat = DataFormat.
        'request.AddJsonBody(listTB)
        'request.RequestFormat = DataFormat.Json
        'request.AddJsonBody(listTB)

        'request.AddHeader("Authorization", "bearer " + bearer)
        Dim data_response = data_client.Execute(request)
        Dim data_response_raw = data_response.Content

        If (data_response.IsSuccessful) Then

            Return True

        Else

            Dim lines() As String = {"Error asynchronous TriggerServer : " + DateTime.Now.ToString()}
            System.IO.File.AppendAllLines(LogPath, lines)
            Return False
        End If

    End Function





    Public Function triggerLocalRun(ConnLoc As ADODB.Connection)
        entityId = "0"
        Dim rows = tbConfig.[Select](" keyType = 'Machine' and keyWord = 'Entity' ")
        For Each row As DataRow In rows
            entityId = strNull(row.Item("config"))
        Next

        If entityId = "" Then entityId = "0"
        If entityId = "0" Then Exit Function

        Dim RunDateTime As String = Now.ToString("yyyy-MMM-dd  HH:mm:ss")
        Dim Rs As New ADODB.Recordset
        Rs = WriteSQL("exec spInterfaceClientTrigger '" & entityId & "','" & RunDateTime & "'", ConnLoc)
        RsClose(Rs)
        Return True

    End Function




    Public Function InterfaceClient2Server(ConnLoc As ADODB.Connection, ConnSrv As ADODB.Connection)
        Dim Rs As New ADODB.Recordset

        Dim RunDateTime As String = Now.ToString("yyyy-MMM-dd  HH:mm:ss")

        Rs = ReadSQL("exec spInterfaceClient2Server '" & entityId & "', '" & RunDateTime & "' ", ConnLoc)
        If Rs.RecordCount = 0 Then
            Return False
        End If

        Dim listTB As New List(Of tbOfflineInteface) 'List(Of tbInterfaceOffline)


        Do Until Rs.EOF

            Dim iOffline = New tbOfflineInteface
            With iOffline
                .Id = RsNullDB(Rs.Fields("Id"))
                .KeyEntity = RsNullDB(Rs.Fields("KeyEntity"))
                .KeyAction = RsNullDB(Rs.Fields("KeyAction"))
                .KeyRecord = RsNullDB(Rs.Fields("KeyRecord"))
                .KeySort = RsNullDB(Rs.Fields("KeySort"))
                .corpId = RsNullDB(Rs.Fields("corpId"))
                .grpId = RsNullDB(Rs.Fields("grpId"))
                .buId = RsNullDB(Rs.Fields("buId"))
                .companyId = RsNullDB(Rs.Fields("companyId"))
                .entityId = RsNullDB(Rs.Fields("entityId"))
                .RecId = RsNullDB(Rs.Fields("RecId"))
                .Id_01 = RsNullDB(Rs.Fields("Id_01"))
                .Id_02 = RsNullDB(Rs.Fields("Id_02"))
                .Id_03 = RsNullDB(Rs.Fields("Id_03"))
                .Id_04 = RsNullDB(Rs.Fields("Id_04"))
                .Id_05 = RsNullDB(Rs.Fields("Id_05"))
                .Id_06 = RsNullDB(Rs.Fields("Id_06"))
                .Id_07 = RsNullDB(Rs.Fields("Id_07"))
                .Id_08 = RsNullDB(Rs.Fields("Id_08"))
                .Id_09 = RsNullDB(Rs.Fields("Id_09"))
                .Id_10 = RsNullDB(Rs.Fields("Id_10"))


                .Str_5_01 = RsNullDB(Rs.Fields("Str_5_01"))
                .Str_5_02 = RsNullDB(Rs.Fields("Str_5_02"))
                .Str_5_03 = RsNullDB(Rs.Fields("Str_5_03"))
                .Str_5_04 = RsNullDB(Rs.Fields("Str_5_04"))
                .Str_5_05 = RsNullDB(Rs.Fields("Str_5_05"))


                .Str_25_01 = RsNullDB(Rs.Fields("Str_25_01"))
                .Str_25_02 = RsNullDB(Rs.Fields("Str_25_02"))
                .Str_25_03 = RsNullDB(Rs.Fields("Str_25_03"))
                .Str_25_04 = RsNullDB(Rs.Fields("Str_25_04"))
                .Str_25_05 = RsNullDB(Rs.Fields("Str_25_05"))
                .Str_25_06 = RsNullDB(Rs.Fields("Str_25_06"))
                .Str_25_07 = RsNullDB(Rs.Fields("Str_25_07"))
                .Str_25_08 = RsNullDB(Rs.Fields("Str_25_08"))
                .Str_25_09 = RsNullDB(Rs.Fields("Str_25_09"))
                .Str_25_10 = RsNullDB(Rs.Fields("Str_25_10"))

                .Str_25_11 = RsNullDB(Rs.Fields("Str_25_11"))
                .Str_25_12 = RsNullDB(Rs.Fields("Str_25_12"))
                .Str_25_13 = RsNullDB(Rs.Fields("Str_25_13"))
                .Str_25_14 = RsNullDB(Rs.Fields("Str_25_14"))
                .Str_25_15 = RsNullDB(Rs.Fields("Str_25_15"))
                .Str_25_16 = RsNullDB(Rs.Fields("Str_25_16"))
                .Str_25_17 = RsNullDB(Rs.Fields("Str_25_17"))
                .Str_25_18 = RsNullDB(Rs.Fields("Str_25_18"))
                .Str_25_19 = RsNullDB(Rs.Fields("Str_25_19"))
                .Str_25_20 = RsNullDB(Rs.Fields("Str_25_20"))


                .Str_50_01 = RsNullDB(Rs.Fields("Str_50_01"))
                .Str_50_02 = RsNullDB(Rs.Fields("Str_50_02"))
                .Str_50_03 = RsNullDB(Rs.Fields("Str_50_03"))
                .Str_50_04 = RsNullDB(Rs.Fields("Str_50_04"))
                .Str_50_05 = RsNullDB(Rs.Fields("Str_50_05"))
                .Str_50_06 = RsNullDB(Rs.Fields("Str_50_06"))
                .Str_50_07 = RsNullDB(Rs.Fields("Str_50_07"))
                .Str_50_08 = RsNullDB(Rs.Fields("Str_50_08"))
                .Str_50_09 = RsNullDB(Rs.Fields("Str_50_09"))
                .Str_50_10 = RsNullDB(Rs.Fields("Str_50_10"))
                .Str_150_01 = RsNullDB(Rs.Fields("Str_150_01"))
                .Str_150_02 = RsNullDB(Rs.Fields("Str_150_02"))
                .Str_150_03 = RsNullDB(Rs.Fields("Str_150_03"))
                .Str_150_04 = RsNullDB(Rs.Fields("Str_150_04"))
                .Str_150_05 = RsNullDB(Rs.Fields("Str_150_05"))
                .Str_250_01 = RsNullDB(Rs.Fields("Str_250_01"))
                .Str_250_02 = RsNullDB(Rs.Fields("Str_250_02"))
                .Str_250_03 = RsNullDB(Rs.Fields("Str_250_03"))
                .Str_250_04 = RsNullDB(Rs.Fields("Str_250_04"))
                .Str_250_05 = RsNullDB(Rs.Fields("Str_250_05"))
                .Str_500_01 = RsNullDB(Rs.Fields("Str_500_01"))
                .Str_500_02 = RsNullDB(Rs.Fields("Str_500_02"))
                .Str_500_03 = RsNullDB(Rs.Fields("Str_500_03"))
                .Str_500_04 = RsNullDB(Rs.Fields("Str_500_04"))
                .Str_500_05 = RsNullDB(Rs.Fields("Str_500_05"))
                .Str_Max_01 = RsNullDB(Rs.Fields("Str_Max_01"))
                .Str_Max_02 = RsNullDB(Rs.Fields("Str_Max_02"))
                .Str_Max_03 = RsNullDB(Rs.Fields("Str_Max_03"))
                .Str_Max_04 = RsNullDB(Rs.Fields("Str_Max_04"))
                .Str_Max_05 = RsNullDB(Rs.Fields("Str_Max_05"))
                .DateTime_01 = RsNullDT(Rs.Fields("DateTime_01"))
                .DateTime_02 = RsNullDT(Rs.Fields("DateTime_02"))
                .DateTime_03 = RsNullDT(Rs.Fields("DateTime_03"))
                .DateTime_04 = RsNullDT(Rs.Fields("DateTime_04"))
                .DateTime_05 = RsNullDT(Rs.Fields("DateTime_05"))

                .Float_01 = RsNullDB(Rs.Fields("Float_01"))
                .Float_02 = RsNullDB(Rs.Fields("Float_02"))
                .Float_03 = RsNullDB(Rs.Fields("Float_03"))
                .Float_04 = RsNullDB(Rs.Fields("Float_04"))
                .Float_05 = RsNullDB(Rs.Fields("Float_05"))
                .Float_06 = RsNullDB(Rs.Fields("Float_06"))
                .Float_07 = RsNullDB(Rs.Fields("Float_07"))
                .Float_08 = RsNullDB(Rs.Fields("Float_08"))
                .Float_09 = RsNullDB(Rs.Fields("Float_09"))
                .Float_10 = RsNullDB(Rs.Fields("Float_10"))
                .Float_11 = RsNullDB(Rs.Fields("Float_11"))
                .Float_12 = RsNullDB(Rs.Fields("Float_12"))
                .Float_13 = RsNullDB(Rs.Fields("Float_13"))
                .Float_14 = RsNullDB(Rs.Fields("Float_14"))
                .Float_15 = RsNullDB(Rs.Fields("Float_15"))
                .Float_16 = RsNullDB(Rs.Fields("Float_16"))
                .Float_17 = RsNullDB(Rs.Fields("Float_17"))
                .Float_18 = RsNullDB(Rs.Fields("Float_18"))
                .Float_19 = RsNullDB(Rs.Fields("Float_19"))
                .Float_20 = RsNullDB(Rs.Fields("Float_20"))

                .Int_01 = RsNullDB(Rs.Fields("Int_01"))
                .Int_02 = RsNullDB(Rs.Fields("Int_02"))
                .Int_03 = RsNullDB(Rs.Fields("Int_03"))
                .Int_04 = RsNullDB(Rs.Fields("Int_04"))
                .Int_05 = RsNullDB(Rs.Fields("Int_05"))
                .Int_06 = RsNullDB(Rs.Fields("Int_06"))
                .Int_07 = RsNullDB(Rs.Fields("Int_07"))
                .Int_08 = RsNullDB(Rs.Fields("Int_08"))
                .Int_09 = RsNullDB(Rs.Fields("Int_09"))
                .Int_10 = RsNullDB(Rs.Fields("Int_10"))
                .Int_11 = RsNullDB(Rs.Fields("Int_11"))
                .Int_12 = RsNullDB(Rs.Fields("Int_12"))
                .Int_13 = RsNullDB(Rs.Fields("Int_13"))
                .Int_14 = RsNullDB(Rs.Fields("Int_14"))
                .Int_15 = RsNullDB(Rs.Fields("Int_15"))


                .num_180_01 = RsNullDB(Rs.Fields("num_180_01"))
                .num_180_02 = RsNullDB(Rs.Fields("num_180_02"))
                .num_180_03 = RsNullDB(Rs.Fields("num_180_03"))
                .num_180_04 = RsNullDB(Rs.Fields("num_180_04"))
                .num_180_05 = RsNullDB(Rs.Fields("num_180_05"))
                .bit_01 = RsNullDB(Rs.Fields("bit_01"))
                .bit_02 = RsNullDB(Rs.Fields("bit_02"))
                .bit_03 = RsNullDB(Rs.Fields("bit_03"))
                .bit_04 = RsNullDB(Rs.Fields("bit_04"))
                .bit_05 = RsNullDB(Rs.Fields("bit_05"))
                .FileType = RsNullDB(Rs.Fields("FileType"))
                .FileDisposition = RsNullDB(Rs.Fields("FileDisposition"))
                .FileLength = RsNullDB(Rs.Fields("FileLength"))
                .FileName = RsNullDB(Rs.Fields("FileName"))
                .createBy = RsNullDB(Rs.Fields("createBy"))
                .createDate = RsNullDT(Rs.Fields("createDate"))
                .modifyBy = RsNullDB(Rs.Fields("modifyBy"))
                .modifyDate = RsNullDT(Rs.Fields("modifyDate"))
                .Inf_Status = RsNullDB(Rs.Fields("Inf_Status"))
                .Inf_Datetime = RsNullDT(Rs.Fields("Inf_Datetime"))
            End With
            listTB.Add(iOffline)
            Rs.MoveNext()

        Loop

        Dim Url = ""
        Dim irow As DataRow = tbConfig.[Select](" keyType = 'URL'and keyWord = 'APItoServer' ").FirstOrDefault()
        If IsDBNull(irow) = False Then Url = strNull(irow.Item("config"))
        Dim hostName As String = GetURLHostName(Url)
        Dim uriName As String = Replace(Url, hostName, "")



        Dim data_client = New RestClient(hostName)
        Dim request = New RestRequest(uriName, Method.Post)
        'request.AddHeader("Accept", "text/xml")
        request.AddHeader("Accept", "application/json")
        request.RequestFormat = DataFormat.Json
        request.AddJsonBody(listTB)

        'request.AddHeader("Authorization", "bearer " + bearer)
        Dim data_response = data_client.Execute(request)
        Dim data_response_raw = data_response.Content



        If (data_response.IsSuccessful) Then
            If Rs.RecordCount <> 0 Then
                Dim RsUp As New ADODB.Recordset
                RsUp = WriteSQL("exec spInterfaceClient2Server_runComplete '" & entityId & "','" & RunDateTime & "'", ConnLoc)
                RsClose(RsUp)
            End If

        Else

            Dim lines() As String = {"Error asynchronous Client2Server : " + DateTime.Now.ToString()}
            System.IO.File.AppendAllLines(LogPath, lines)

        End If

        RsClose(Rs)


        Return True


    End Function




    Public Function UpdateDOA(ConnLoc As ADODB.Connection, ConnSrv As ADODB.Connection) As Boolean
        Dim Ret = False
        Try

            Dim rows As DataRow()
            Dim Sql As String
            Dim DOAUrl = ""
            Dim DOACode = ""
            rows = tbConfig.[Select](" keyType = 'DOA'  and keyWord = 'URL'")
            For Each row As DataRow In rows
                DOAUrl = strNull(row.Item("config"))
            Next

            rows = tbConfig.[Select](" keyType = 'DOA' and keyWord like 'WorkFlowCode%'  ")
            For Each row As DataRow In rows
                DOACode = strNull(row.Item("config"))
                If DOACode = "" Then GoTo JmpNextLoop
                Dim pramSQL As String = "code=" + DOACode + "&bcGroup=" + grpCode + "&bcBusinessUnit=" + grpCode + "&bcCompanyCode=" + companyCode + "&bcEntity=" + entityCode

                If Right(DOAUrl, 1) <> "?" Then DOAUrl = DOAUrl + "?"

                Dim URLGetDOA As String = DOAUrl + pramSQL

                Dim request2 As HttpWebRequest = HttpWebRequest.Create(URLGetDOA)

                request2.Method = "GET"
                request2.ContentType = "application/x-www-form-urlencoded"
                request2.KeepAlive = True
                request2.ContinueTimeout = 10000
                request2.Host = "intranet.nathalin.com"

                Dim Header2Collection As WebHeaderCollection = request2.Headers
                'Dim BearerAccessToken = "Bearer " & access_token
                'Header2Collection.Add("Authorization", BearerAccessToken)
                Header2Collection.Add("x-myobapi-key", "MYAPIKEY")
                Header2Collection.Add("x-myobapi-version", "v2")
                Header2Collection.Add("Accept-Encoding", "gzip,deflate")
                Header2Collection.Add("Cache-Control", "no-cache")
                Dim postBytes2 As Byte() = New UTF8Encoding().GetBytes(request2.ToString())

                Dim response2 As HttpWebResponse = CType(request2.GetResponse(), HttpWebResponse)
                Dim receiveStream2 As Stream = response2.GetResponseStream()
                Dim readStream2 As New StreamReader(receiveStream2)
                Dim jsonstring = readStream2.ReadToEnd()

                If InStr(jsonstring, "The condition which is not matching") = 0 Then 'no Error
                    If jsonstring <> "" Then
                        Dim RsU As New ADODB.Recordset
                        Sql = "update tbSystemOfflineConfig set config = '" & jsonstring & "' where keytype = 'DOA' and keyword = '" + DOACode + "'"
                        RsU = WriteSQL(Sql, ConnLoc)
                        RsClose(RsU)
                    End If

                    Ret = True
                End If

JmpNextLoop:
            Next
            Return Ret

        Catch ex As Exception
            Return Ret

        End Try



    End Function





End Class
