Module mod_variable


    '------------------------------------------------------------------------------------------

    Public LogPath As String = AppDomain.CurrentDomain.BaseDirectory + "Log.txt"
    Public ConfigFile As String = AppDomain.CurrentDomain.BaseDirectory + "conifg.txt"

    Public SQLConfigFile As String = AppDomain.CurrentDomain.BaseDirectory + "sql.cfg"
    Public TimeInteval As Double = 0.5

    '------------------------------------------------------------------------------------------
    Public locUID As String = ""
    Public locPWD As String = ""
    Public locDB As String = ""
    Public locServer As String = ""

    Public srvUID As String = ""
    Public srvPWD As String = ""
    Public srvDB As String = ""
    Public srvServer As String = ""
    '------------------------------------------------------------------------------------------


    '------------------------------------------------------------------------------------------
    Public corpId As String = ""
    Public grpId As String = ""
    Public buId As String = ""
    Public companyId As String = ""
    Public entityId As String = ""

    Public corpCode As String = ""
    Public grpCode As String = ""
    Public buCode As String = ""
    Public companyCode As String = ""
    Public entityCode As String = ""
    '------------------------------------------------------------------------------------------

    Public Sql As String = ""


    Public tbConfig As New DataTable()


	Public Structure tbInterfaceOffline


		Public Id As Guid

		Public KeyAction As String

		Public KeyRecord As String

		'Public corpId As Long
		'Public grpId As Long
		'Public buId As Long
		'Public companyId As Long
		'Public entityId As Long

		'Public RecId As Guid

		'Public Id_01? As Guid
		'Public Id_02? As Guid
		'Public Id_03? As Guid
		'Public Id_04? As Guid
		'Public Id_05? As Guid

		'Public Id_06? As Guid
		'Public Id_07? As Guid
		'Public Id_08? As Guid
		'Public Id_09? As Guid
		'Public Id_10? As Guid

		'Public Str_5_01 As String
		'Public Str_5_02 As String
		'Public Str_5_03 As String
		'Public Str_5_04 As String
		'Public Str_5_05 As String


		'Public Str_25_01 As String
		'Public Str_25_02 As String
		'Public Str_25_03 As String
		'Public Str_25_04 As String
		'Public Str_25_05 As String


		'Public Str_50_01 As String
		'Public Str_50_02 As String
		'Public Str_50_03 As String
		'Public Str_50_04 As String
		'Public Str_50_05 As String
		'Public Str_50_06 As String
		'Public Str_50_07 As String
		'Public Str_50_08 As String
		'Public Str_50_09 As String
		'Public Str_50_10 As String


		'Public Str_150_01 As String
		'Public Str_150_02 As String
		'Public Str_150_03 As String
		'Public Str_150_04 As String
		'Public Str_150_05 As String

		'Public Str_250_01 As String
		'Public Str_250_02 As String
		'Public Str_250_03 As String
		'Public Str_250_04 As String
		'Public Str_250_05 As String


		'Public Str_500_01 As String
		'Public Str_500_02 As String
		'Public Str_500_03 As String
		'Public Str_500_04 As String
		'Public Str_500_05 As String


		'Public Str_Max_01 As String
		'Public Str_Max_02 As String
		'Public Str_Max_03 As String
		'Public Str_Max_04 As String
		'Public Str_Max_05 As String


		'Public DateTime_01? As DateTime
		'Public DateTime_02? As DateTime
		'Public DateTime_03? As DateTime
		'Public DateTime_04? As DateTime
		'Public DateTime_05? As DateTime


		'Public Float_01? As Double
		'Public Float_02? As Double
		'Public Float_03? As Double
		'Public Float_04? As Double
		'Public Float_05? As Double
		'Public Float_06? As Double
		'Public Float_07? As Double
		'Public Float_08? As Double
		'Public Float_09? As Double
		'Public Float_10? As Double


		'Public Int_01? As Long
		'Public Int_02? As Long
		'Public Int_03? As Long
		'Public Int_04? As Long
		'Public Int_05? As Long


		'Public num_180_01? As Decimal
		'Public num_180_02? As Decimal
		'Public num_180_03? As Decimal
		'Public num_180_04? As Decimal
		'Public num_180_05? As Decimal

		'Public bit_01? As Boolean
		'Public bit_02? As Boolean
		'Public bit_03? As Boolean
		'Public bit_04? As Boolean
		'Public bit_05? As Boolean



		'Public FileType As String

		'Public FileDisposition As Byte()

		'Public FileLength As Long

		'Public FileName As String

		'Public createBy As String

		'Public createDate? As DateTime

		'Public modifyBy As String


		'Public modifyDate? As DateTime



		'Public Inf_Status As String

		'Public Inf_Datetime? As DateTime





	End Structure


End Module
