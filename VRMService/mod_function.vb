Imports System.IO
Imports System.Net
Imports System.Text
Imports RestSharp

Module mod_function

    Public Function GetURLHostName(URL As String) As String
        If (URL = "") Then Return URL
        Dim tHost = URL.Split("/"c)(0)
        Dim dHost = URL.Split("/"c)(2)
        'Dim strUrl = Replace(Replace(URL, "http://", ""), "https://", "")
        Return tHost & "//" & dHost ' strUrl.Split("/"c)(3)

    End Function

    Public Function GetAPI(URL As String) As String
        Dim Retstring As String = ""
        Try
            If (URL = "") Then Return ""

            Dim hostName As String = GetURLHostName(URL)
            Dim uriName As String = Replace(URL, hostName, "")

            Dim data_client = New RestClient(hostName)
            Dim request = New RestRequest(uriName, Method.Get)
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
                Retstring = data_response_raw

            Else
                Retstring = "ERR"
            End If


            'Dim request2 As HttpWebRequest = HttpWebRequest.Create(URL)
            'Dim hostName As String = GetURLHostName(URL)
            'request2.Method = "GET"
            'request2.ContentType = "application/x-www-form-urlencoded"
            'request2.KeepAlive = True
            'request2.ContinueTimeout = 10000
            'request2.Host = hostName

            'Dim Header2Collection As WebHeaderCollection = request2.Headers
            ''Dim BearerAccessToken = "Bearer " & access_token
            ''Header2Collection.Add("Authorization", BearerAccessToken)
            'Header2Collection.Add("x-myobapi-key", "MYAPIKEY")
            'Header2Collection.Add("x-myobapi-version", "v2")
            'Header2Collection.Add("Accept-Encoding", "gzip,deflate")
            'Header2Collection.Add("Cache-Control", "no-cache")
            'Dim postBytes2 As Byte() = New UTF8Encoding().GetBytes(request2.ToString())

            'Dim response2 As HttpWebResponse = CType(request2.GetResponse(), HttpWebResponse)
            'Dim receiveStream2 As Stream = response2.GetResponseStream()
            'Dim readStream2 As New StreamReader(receiveStream2)
            'Retstring = readStream2.ReadToEnd()

            Return Retstring

        Catch ex As Exception

            Return ""

        End Try

    End Function




    Public Function ConnectServerAPI() As Boolean
        Dim Ret As Boolean = False
        Dim Url = ""
        Dim rows As DataRow() = tbConfig.[Select](" keyType = 'URL'")
        For Each row As DataRow In rows
            If LCase(strNull(row.Item("keyWord"))) = LCase("CheckOnline") Then Url = strNull(row.Item("config"))
        Next
        Dim apiResult As String = GetAPI(Url)
        If apiResult <> "" Then
            Dim arrApi = Split(apiResult, "|")
            If LCase(arrApi(0)) = "ok" Then Ret = True
        End If
        Return Ret
    End Function



    Public Function ConnectServer(ConnLoc As ADODB.Connection) As ADODB.Connection

        If Not IsNothing(ConnLoc.ConnectionString) Then
            On Error GoTo JmpErr
            Dim ConnSrv = New ADODB.Connection

            Dim serverName As String = ""
            Dim dbName As String = ""
            Dim dbUser As String = ""
            Dim dbPwd As String = ""


            Dim rows As DataRow() = tbConfig.[Select](" keyType = 'server'")
            For Each row As DataRow In rows
                If LCase(strNull(row.Item("keyWord"))) = LCase("Server") Then serverName = strNull(row.Item("config"))
                If LCase(strNull(row.Item("keyWord"))) = LCase("UserID") Then dbUser = strNull(row.Item("config"))
                If LCase(strNull(row.Item("keyWord"))) = LCase("PWD") Then dbPwd = strNull(row.Item("config"))
                If LCase(strNull(row.Item("keyWord"))) = LCase("Database") Then dbName = strNull(row.Item("config"))
            Next


            'Dim fRs As New ADODB.Recordset
            'fRs = ReadSQL("select * from tbSystemAdepterConfig where keyType = 'SQL'", ConnLoc)
            'Do Until fRs.EOF
            '    If LCase(strNull(fRs.Fields("keyWord").Value)) = LCase("Server") Then serverName = strNull(fRs.Fields("value01").Value)
            '    If LCase(strNull(fRs.Fields("keyWord").Value)) = LCase("UserID") Then dbUser = strNull(fRs.Fields("value01").Value)
            '    If LCase(strNull(fRs.Fields("keyWord").Value)) = LCase("PWD") Then dbPwd = strNull(fRs.Fields("value01").Value)
            '    If LCase(strNull(fRs.Fields("keyWord").Value)) = LCase("Database") Then dbName = strNull(fRs.Fields("value01").Value)
            '    fRs.MoveNext()
            'Loop
            'fRs = Nothing

            Dim srvConnStr = "Provider=sqloledb;Data Source=" & serverName & ";Initial Catalog=" & dbName & ";User Id=" & dbUser & ";Password=" & dbPwd & ";Trusted_Connection=False;"
            ConnSrv.ConnectionString = srvConnStr
            ConnSrv.CursorLocation = ADODB.CursorLocationEnum.adUseClient
            ConnSrv.CommandTimeout = 1500
            ConnSrv.Open()

            ConnectServer = ConnSrv
            WriteLog(ConnLoc, "OK", "connect server", "")
            Exit Function

JmpErr:
            WriteLog(ConnLoc, "Error", "fail connect server", Err.Description)
            Err.Clear()
        End If
    End Function

    Public Function SQLText() As String

    End Function


    Public Sub WriteLog(ConnLocal As ADODB.Connection, Optional logTyp As String = "", Optional logMsg As String = "", Optional logVal As String = "")

        If Not IsNothing(ConnLocal.ConnectionString) Then
            'ConnLocal.BeginTrans()
            Dim Sql As String
            Sql = "insert tbSystemLog (Id , logType , logMessage ,  logAction , logDate) select newid() ," & zSQLText(logTyp) & " ," & zSQLText(logMsg) & " , " & zSQLText(logVal) & " , getdate()"

            Dim fRs As New ADODB.Recordset
            fRs = WriteSQL(Sql, ConnLocal)


            'ConnLocal.Execute(Sql)
            'ConnLocal.CommitTrans()
        End If
    End Sub


    Public Function ConvDate(strDate As String)
        Dim dateString, format As String
        Dim result As Date

        Dim provider As Globalization.CultureInfo = Globalization.CultureInfo.CurrentCulture

        ' Parse date and time with custom specifier.
        dateString = strDate
        format = "yyyy-MMM-dd HH:mm:ss"
        Try
            Dim dtStart As DateTime = Convert.ToDateTime(dateString)
            result = Date.ParseExact(dateString, format, provider)
            Console.WriteLine("{0} converts to {1}.", dateString, result.ToString())
            Console.ReadLine()
        Catch e As FormatException
            Console.WriteLine("{0} is not in the correct format.", dateString)
        End Try
    End Function

    Public Function RsAsInsert(rsField As ADODB.Field, strVal As Object)

        If IsDBNull(strVal) Then Return "null"

        Dim retStr As String = Replace(strVal, """", """""")

        If rsField.Type = ADODB.DataTypeEnum.adGUID Then Return "'" & retStr & "'"

        If InStr(LCase(rsField.Name), LCase("datetime")) > 0 Then
            ConvDate(retStr)
            Return "convert(datetime2(7) , '" & Convert.ToDateTime(retStr).ToString("yyyy-MMM-dd HH:mm:ss") & "')"
        End If
        If InStr(LCase(rsField.Name), LCase("modifyDate")) > 0 Then Return "convert(datetime2(7) ,'" & Convert.ToDateTime(retStr).ToString("yyyy-MMM-dd HH:mm:ss") & "')"
        If InStr(LCase(rsField.Name), LCase("createDate")) > 0 Then Return "convert(datetime2(7) , '" & Convert.ToDateTime(retStr).ToString("yyyy-MMM-dd HH:mm:ss") & "')"

        If rsField.Type = ADODB.DataTypeEnum.adDBDate Then Return "convert(datetime2(7) , '" & Convert.ToDateTime(retStr).ToString("yyyy-MMM-dd HH:mm:ss") & "')"

        If rsField.Type = ADODB.DataTypeEnum.adDouble Then Return Convert.ToDouble(retStr)
        If rsField.Type = ADODB.DataTypeEnum.adInteger Then Return Convert.ToInt32(retStr)
        If rsField.Type = ADODB.DataTypeEnum.adNumeric Then Return Convert.ToDecimal(retStr)
        If rsField.Type = ADODB.DataTypeEnum.adDecimal Then Return Convert.ToDecimal(retStr)
        If rsField.Type = ADODB.DataTypeEnum.adSingle Then Return Convert.ToSingle(retStr)
        If rsField.Type = ADODB.DataTypeEnum.adBigInt Then Return Convert.ToInt64(retStr)

        Return "'" & retStr & "'"


    End Function



    Public Function RsNullDT(rsField As ADODB.Field)
        If IsDBNull(rsField.Value) Then
            Return Nothing 'CType(Nothing, DateTime?)
        Else
            Return Convert.ToDateTime(rsField.Value)
        End If

    End Function



    Public Function RsNullDB(rsField As ADODB.Field)
        If IsDBNull(rsField.Value) Then
            Return Nothing

            If rsField.Type = ADODB.DataTypeEnum.adGUID Then Return CType(Nothing, Guid?)


            If rsField.Type = ADODB.DataTypeEnum.adLongVarBinary Then Return Nothing
            If rsField.Type = ADODB.DataTypeEnum.adVarWChar Then Return Nothing
            If rsField.Type = ADODB.DataTypeEnum.adLongVarWChar Then Return Nothing
            If rsField.Type = ADODB.DataTypeEnum.adWChar Then Return Nothing

            If rsField.Type = ADODB.DataTypeEnum.adDouble Then Return Convert.ToDouble(0)
            If rsField.Type = ADODB.DataTypeEnum.adInteger Then Return Convert.ToInt32(0)
            If rsField.Type = ADODB.DataTypeEnum.adNumeric Then Return Convert.ToDecimal(0)
            If rsField.Type = ADODB.DataTypeEnum.adDecimal Then Return Convert.ToDecimal(0)
            If rsField.Type = ADODB.DataTypeEnum.adSingle Then Return Convert.ToSingle(0)
            If rsField.Type = ADODB.DataTypeEnum.adBigInt Then Return Convert.ToInt64(0)
            If rsField.Type = ADODB.DataTypeEnum.adBoolean Then Return Nothing 'CType(Nothing, Boolean)
            If rsField.Type = ADODB.DataTypeEnum.adDBDate Then Return CType(Nothing, DateTime?)
            'ElseIf IsNothing(rsField.Value) Then
            '    If rsField.Type = ADODB.DataTypeEnum.adGUID Then Return CType(Nothing, Guid?)

            '    If rsField.Type = ADODB.DataTypeEnum.adLongVarBinary Then Return vbNull

            '    If rsField.Type = ADODB.DataTypeEnum.adVarWChar Then Return vbNull
            '    If rsField.Type = ADODB.DataTypeEnum.adLongVarWChar Then Return vbNull
            '    If rsField.Type = ADODB.DataTypeEnum.adWChar Then Return vbNull

            '    If rsField.Type = ADODB.DataTypeEnum.adDouble Then Return Convert.ToDouble(0)
            '    If rsField.Type = ADODB.DataTypeEnum.adInteger Then Return Convert.ToInt32(0)
            '    If rsField.Type = ADODB.DataTypeEnum.adNumeric Then Return Convert.ToDecimal(0)
            '    If rsField.Type = ADODB.DataTypeEnum.adDecimal Then Return Convert.ToDecimal(0)
            '    If rsField.Type = ADODB.DataTypeEnum.adSingle Then Return Convert.ToSingle(0)
            '    If rsField.Type = ADODB.DataTypeEnum.adBigInt Then Return Convert.ToInt64(0)

            '    If rsField.Type = ADODB.DataTypeEnum.adBoolean Then Return CType(Nothing, Boolean)
            '    If rsField.Type = ADODB.DataTypeEnum.adDBDate Then Return CType(Nothing, DateTime?)

        Else

            If rsField.Type = ADODB.DataTypeEnum.adGUID Then Return toGuid(rsField.Value)
            If rsField.Type = ADODB.DataTypeEnum.adDBDate Then Return Convert.ToDateTime(rsField.Value)

            If rsField.Type = ADODB.DataTypeEnum.adDouble Then Return Convert.ToDouble(rsField.Value)
            If rsField.Type = ADODB.DataTypeEnum.adInteger Then Return Convert.ToInt32(rsField.Value)
            If rsField.Type = ADODB.DataTypeEnum.adNumeric Then Return Convert.ToDecimal(rsField.Value)
            If rsField.Type = ADODB.DataTypeEnum.adDecimal Then Return Convert.ToDecimal(rsField.Value)
            If rsField.Type = ADODB.DataTypeEnum.adSingle Then Return Convert.ToSingle(rsField.Value)
            If rsField.Type = ADODB.DataTypeEnum.adBigInt Then Return Convert.ToInt64(rsField.Value)

            Return rsField.Value
        End If

    End Function


    Public Function strNull(str As String, Optional ByRef Def As String = "")
        Dim strCall As String = str
        If String.IsNullOrEmpty(strCall) Then strCall = Def
        strNull = strCall
    End Function



    Public Function ReadSQL(zfSQL, zfConn) As ADODB.Recordset
        Dim zfRs As New ADODB.Recordset
        zfRs.Open(zfSQL, zfConn, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        ReadSQL = zfRs
    End Function

    Public Function WriteSQL(zfSQL, zfConn) As ADODB.Recordset
        Dim zfRs As New ADODB.Recordset
        zfRs.Open(zfSQL, zfConn, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic)
        WriteSQL = zfRs
    End Function


    Public Sub RsClose(fRs As ADODB.Recordset)
        On Error GoTo Jmp1
        fRs.Close()
Jmp1:
        fRs = Nothing
    End Sub


    '////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////Database
    Public Function zSQLFindRange(ByVal zfield As Object, ByVal zFrText As Object, ByVal zToText As Object) As String
        If zFrText <> "" Or zToText <> "" Then
            If (zFrText = zToText) Or (zFrText <> "" And zToText = "") Then
                zSQLFindRange = " and " & zSQLFind(zfield, zFrText, True)
            Else
                zFrText = zSQLText(zFrText, True)
                zToText = zSQLText(zToText, True)
                If InStr(zFrText, "%") > 0 Then zFrText = Replace(zFrText, "%", "0000000000000000000000000")
                If InStr(zToText, "%") > 0 Then zToText = Replace(zToText, "%", "ฮฮฮฮฮฮฮฮฮฮฮฮฮฮฮฮฮฮฮฮฮฮฮฮฮ")
                zSQLFindRange = " and (" & zfield & " between " & zFrText & " and " & zToText & " ) "
            End If
        End If
    End Function

    Public Function Guid2Str(ByVal zText As Object) As String

        If zText = Nothing Then zText = ""

        Guid2Str = Replace(Replace(zText, "{", ""), "}", "")
    End Function
    Public Function toGuid(ByVal zText As Object) As Guid
        If IsDBNull(zText) Then
            toGuid = Guid.Empty
        ElseIf IsNothing(zText) Then
            toGuid = Guid.Empty
        Else
            toGuid = Guid.Parse(zText)

        End If
    End Function


    Public Function zSQLText(ByVal zText As Object, Optional ByVal zSearch As Boolean = False) As String
        Dim SQLText As String
        Dim i As Integer, c As String

        If zText = Nothing Then zText = ""

        zText = Trim(zText)
        For i = 1 To Len(zText)
            c = Mid$(zText, i, 1)
            Select Case c
                Case "'"
                    SQLText = SQLText + "''"
                Case Else
                    SQLText = SQLText + c
            End Select
        Next
        zSQLText = " N'" & SQLText & "' "
        If zSearch Then zSQLText = Replace(zSQLText, "*", "%")
    End Function

    Public Function zSQLFind(ByVal zfield As Object, ByVal zText As Object, Optional ByVal zSameSearch As Boolean = False) As String
        Dim zSQLTemp As String
        If zSameSearch = False Then zText = Replace(Replace(zText, "%", "*"), "*", "")
        zSQLTemp = zSQLText(zText, True)
        If InStr(zSQLTemp, "%") = 0 Then
            zSQLFind = " " & zfield & " = " & zSQLTemp
        Else
            zSQLFind = " " & zfield & " like " & zSQLTemp
        End If
    End Function


End Module
