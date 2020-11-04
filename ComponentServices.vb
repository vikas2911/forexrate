Option Explicit On 

Imports System.Data
Imports Microsoft
Imports System.Text
Imports Interop
Imports System.Globalization
Imports System.EnterpriseServices

#Region "Public Enums"

Public Enum ErrorNumbers
    TheElementIsMandatory = 1
    TheInstanceHasNotBeenFound = 2
    TheParameterIsMandatory = 3
    TheAttributeIsMandatory = 4
    TheTimestampIsIoutOfDate = 5
    TheClassDoesNotSupportHistory = 6
    TheXMLDocumentISInvalid = 7
    InvalidISODate = 8
    SQLObjectMissing = 208
    OracleObjectMissing = 942
    InvalidDate = 15726
    InstanceHasNotBeenFound = 17448
End Enum

Public Enum ObjectStatus
    Active = 0
    Deleted = 1
End Enum

Public Enum eFrequencyType
    [Single] = 0
    Daily = 1
    Weekly = 2
    Monthly = 3
    Yearly = 4
    HalfYearly = 5
    Quarterly = 6
    NumberOfDays = 7
    NumberOfMonths = 8
    NumberOfYears = 9
    NumberOfHours = 10
    NumberOfMinutes = 11
    Minutes = 12
    Hours = 13
    Days = 14
    Months = 15
    Years = 16
    LegalNaturalQuarterDates = 17
    Weeks = 18
    None = 99
End Enum

Public Enum eFrequencyType_e
    Weekly = 1
    Monthly = 2
    Quarterly = 3
    HalfYearly = 4
    Daily = 5
    Yearly = 6
    None = 99
End Enum

Public Enum ePrecisionType
    None = 0
    [End] = 1
    Start = 2
    Specific = 3
End Enum

#End Region

#Region "IABComponent Interface"

Public Interface IABComponent

    Function DeleteCheck( _
            ByVal sComponentName As String, _
            ByVal sClassName As String, _
            ByVal sToken As String) As ComponentServices.DeleteResponse

    Sub SetComponentServices( _
            ByRef oComponentServices As ComponentServices)

    Sub ExecutePeriodProcessing( _
            ByVal dteDate As Date, _
            ByVal iProcessID As Integer, _
            ByVal sSessionID As String, _
            ByVal sBranchID As String)

End Interface

#End Region

#Region "IABSupportsClassicComponent Interface"

Public Interface IABSupportsClassicComponent

    Sub Initialize( _
            ByVal sSessionID As String, _
            ByVal sBranchID As String, _
            ByVal oComponentServices As abComponentServices.ComponentServices)

End Interface

#End Region

#Region "Database Adapters"

Friend Interface IABDatabaseAdapter

    Sub OpenConnection( _
            ByVal sServer As String, _
            ByVal sDatabaseName As String, _
            ByVal sUserName As String, _
            ByVal sPassword As String, _
            ByVal iTimeout As Integer)

    Sub CloseConnection()

    Sub StartTransaction()

    Sub CommitTransaction()

    Sub AbortTransaction()

    Sub ExecuteNonQuery( _
            ByVal sQuery As String, _
            ByRef iRowsAffected As Integer)

    Function ExecuteQueryParameterCollection( _
            ByVal sQuery As String, _
            ByVal bSingleRow As Boolean, _
            ByVal bIDOrdinal As Boolean) As ParameterCollection

    Function ExecuteQueryDataTable( _
            ByVal sQuery As String) As Data.DataTable

    Function ExecuteStoredProcedure( _
            ByVal sName As String, _
            ByVal oParameters As ParameterCollection, _
            ByVal oAlternativeParameters As ParameterCollection) As ParameterCollection

    Function ExecuteStoredProcedureParameters( _
            ByVal sName As String, _
            ByVal alParameterValues As ArrayList) As Data.DataTable

    Function ExecuteStoredProcedureDataTable( _
            ByVal sName As String, _
            ByVal oParameters As ParameterCollection) As Data.DataTable

    Function GetStoredProcedureDataTypes( _
            ByVal sName As String) As ParameterCollection

    Function ConvertTimestampToString( _
            ByVal byteTimestamp() As Byte) As String

    ReadOnly Property DBProcessId() As Integer

End Interface

Friend Class SQLServerAdapter

    Implements IABDatabaseAdapter

    Private m_oConnection As SqlClient.SqlConnection
    Private m_oTransaction As SqlClient.SqlTransaction
    Private m_sConnectionString As String
    Private m_iSPID As Integer
    Private m_iCommandTimeout As Integer

    Public ReadOnly Property DBProcessId() As Integer Implements IABDatabaseAdapter.DBProcessId
        Get
            Return m_iSPID
        End Get
    End Property

    Public Sub OpenConnection( _
            ByVal sServer As String, _
            ByVal sDatabaseName As String, _
            ByVal sUserName As String, _
            ByVal sPassword As String, _
            ByVal iTimeout As Integer) Implements IABDatabaseAdapter.OpenConnection

        Dim sConnectionString As String
        Dim sbConnectionString As New StringBuilder("Data Source=")
        Dim pcSPID As ParameterCollection
        Dim sConnectionStringForLog As String


        sbConnectionString.Append(sServer)
        sbConnectionString.Append(";Initial Catalog=")
        sbConnectionString.Append(sDatabaseName)
        sbConnectionString.Append(";User ID=")
        sbConnectionString.Append(sUserName)
        sbConnectionString.Append(";Enlist=False;Connection Timeout=")
        sbConnectionString.Append(iTimeout)

        sConnectionStringForLog = sbConnectionString.ToString()

        sbConnectionString.Append(";Password=")
        sbConnectionString.Append(sPassword)

        sConnectionString = sbConnectionString.ToString()

        m_sConnectionString = sConnectionString
        m_iCommandTimeout = iTimeout

        Try
            m_oConnection = New SqlClient.SqlConnection(sConnectionString)
            m_oConnection.Open()

            'If ComponentServices.m_structConfigSettings.eLoggingLevel = ComponentServices.LoggingLevel.Full Then
            pcSPID = ExecuteQueryParameterCollection("SELECT @@SPID AS SPID", True, False)
            m_iSPID = pcSPID.GetIntegerValue("SPID", 0)

            'ComponentServices.LogInformation("SQL Server connection open - " & m_iSPID.ToString(), "abComponentServices", "OpenConnection")
            'End If

        Catch oError As SqlClient.SqlException
            ComponentServices.LogError( _
                    True, _
                    oError.Number, _
                    oError.Message, _
                    "abComponentServices", _
                    oError.StackTrace, _
                    ComponentServices.ErrorSeverity.ES_Error, _
                    sConnectionStringForLog)
        End Try

    End Sub

    Public Sub CloseConnection() Implements IABDatabaseAdapter.CloseConnection

        Dim sLogInformation As StringBuilder

        Try
            If m_oConnection.State <> ConnectionState.Closed Then
                m_oConnection.Close()
            End If

            If ComponentServices.m_structConfigSettings.eLoggingLevel = ComponentServices.LoggingLevel.Full Then
                sLogInformation = New StringBuilder("SQL Server connection closed - SPID : ")
                sLogInformation.Append(m_iSPID.ToString(System.Globalization.CultureInfo.InvariantCulture))

                ComponentServices.LogInformation(sLogInformation.ToString, "abComponentServices", "ConnectionClosed")
            End If

        Catch ex As System.Exception
            ComponentServices.LogError(True, "ComponentServices", ex, "CloseConnection")

        Finally
            m_oConnection.Dispose()
            m_oConnection = Nothing
        End Try

    End Sub

    Public Sub StartTransaction() Implements IABDatabaseAdapter.StartTransaction
        'Default Isolation level should be ReadCommitted. As per the discussion with Jim, it is changed from ReadUncommitted
        If m_oTransaction Is Nothing Then
            m_oTransaction = m_oConnection.BeginTransaction(IsolationLevel.ReadCommitted)
        End If

    End Sub

    Private Function ExecuteQuery( _
            ByVal sQuery As String, _
            ByRef iRowsAffected As Integer) As SqlClient.SqlDataReader

        Dim oCommand As SqlClient.SqlCommand
        Dim oDataReader As SqlClient.SqlDataReader
        Dim sQueryDescription As New StringBuilder

        Try
            If m_iSPID <> 0 Then
                sQueryDescription.Append("SPID: ")
                sQueryDescription.Append(m_iSPID)
                sQueryDescription.Append(" : ")
                sQueryDescription.Append(sQuery)
            Else
                sQueryDescription.Append(sQuery)
            End If

            oCommand = New SqlClient.SqlCommand( _
                    sQuery, _
                    m_oConnection)
            oCommand.Transaction = m_oTransaction
            oCommand.CommandTimeout = m_iCommandTimeout
            oDataReader = oCommand.ExecuteReader()
            iRowsAffected = oDataReader.RecordsAffected

            Call ComponentServices.LogInformation( _
                    sQueryDescription.ToString, _
                    "abComponentServices", _
                    "ExecuteQuery")

        Catch oError As SqlClient.SqlException
            If Not (oDataReader Is Nothing) Then
                oDataReader.Close()
                oDataReader = Nothing
            End If
            ComponentServices.LogError( _
                    True, _
                    oError.Number, _
                    oError.Message, _
                    "abComponentServices", _
                    oError.StackTrace, _
                    ComponentServices.ErrorSeverity.ES_Error, _
                    sQueryDescription.ToString)
        Finally
            If Not oCommand Is Nothing Then
                'oCommand.Transaction = Nothing
                'oCommand.Connection = Nothing
                oCommand.Dispose()
                oCommand = Nothing
            End If
        End Try

        Return oDataReader

    End Function

    Public Sub ExecuteNonQuery( _
            ByVal sQuery As String, _
            ByRef iRowsAffected As Integer) Implements IABDatabaseAdapter.ExecuteNonQuery

        Dim oDataReader As SqlClient.SqlDataReader

        Try
            oDataReader = ExecuteQuery(sQuery, iRowsAffected)

        Catch ex As ActiveBankException
            ComponentServices.LogError(True, ex)
        Catch ex As System.Exception
            ComponentServices.LogError(True, "abComponentServices", ex)
        Finally
            If Not oDataReader Is Nothing Then
                oDataReader.Close()
                oDataReader = Nothing
            End If
        End Try

    End Sub

    Public Function ExecuteQueryParameterCollection( _
            ByVal strQuery As String, _
            ByVal bSingleRow As Boolean, _
            ByVal bIDOrdinal As Boolean) As ParameterCollection Implements IABDatabaseAdapter.ExecuteQueryParameterCollection

        Dim oDataReader As SqlClient.SqlDataReader
        Dim pcResult As ParameterCollection
        Dim pcResultRow As ParameterCollection
        Dim iIDOrdinal As Integer
        Dim iRow As Integer
        Dim iRowsAffected As Integer

        Try
            oDataReader = ExecuteQuery(strQuery, iRowsAffected)

            If oDataReader Is Nothing Then
                pcResult = Nothing
            Else
                pcResult = New ParameterCollection

                If bSingleRow Then
                    If oDataReader.Read() Then
                        AddDataReaderToParameterCollection( _
                                oDataReader, _
                                pcResult)
                    Else
                        pcResult = Nothing
                    End If
                Else
                    iRow = 1
                    ' Determine if we want to index by ID or row number
                    If bIDOrdinal Then
                        iIDOrdinal = GetColumnOrdinal("ID", oDataReader)
                        If iIDOrdinal = -1 Then
                            ' No ID column found
                            bIDOrdinal = False
                        End If
                    End If

                    While oDataReader.Read()
                        pcResultRow = New ParameterCollection

                        AddDataReaderToParameterCollection( _
                                oDataReader, _
                                pcResultRow)

                        If bIDOrdinal Then
                            pcResult.Add(CStr(oDataReader.GetDecimal(iIDOrdinal)), pcResultRow)
                        Else
                            pcResult.Add(CStr(iRow), pcResultRow)
                        End If

                        iRow += 1
                    End While

                    pcResult.Reset()
                End If
            End If

        Catch ex As ActiveBankException
            ComponentServices.LogError(True, ex)
        Catch ex As Exception
            ComponentServices.LogError(True, "abComponentServices", ex)
        Finally
            If Not oDataReader Is Nothing Then
                oDataReader.Close()
                oDataReader = Nothing
            End If
        End Try

        Return pcResult

    End Function

    Private Sub AddDataReaderToParameterCollection( _
            ByRef oDataReader As SqlClient.SqlDataReader, _
            ByRef oResult As ParameterCollection)

        Dim iFields As Integer
        Dim iField As Integer

        iFields = oDataReader.FieldCount - 1
        For iField = 0 To iFields
            If Not oDataReader.IsDBNull(iField) Then
                Select Case oDataReader.GetDataTypeName(iField)
                    Case "integer", "int"
                        Call oResult.Add(oDataReader.GetName(iField), oDataReader.GetInt32(iField))

                    Case "tinyint"
                        Call oResult.Add(oDataReader.GetName(iField), oDataReader.GetByte(iField))

                    Case "smallint"
                        Call oResult.Add(oDataReader.GetName(iField), oDataReader.GetInt16(iField))

                    Case "bit"
                        Call oResult.Add(oDataReader.GetName(iField), oDataReader.GetBoolean(iField))

                    Case "datetime"
                        Call oResult.Add(oDataReader.GetName(iField), oDataReader.GetDateTime(iField))

                    Case "decimal"
                        Call oResult.Add(oDataReader.GetName(iField), oDataReader.GetDecimal(iField))

                    Case "double"
                        Call oResult.Add(oDataReader.GetName(iField), oDataReader.GetDouble(iField))

                    Case "varchar", "nvarchar", "nchar", "char", "text"
                        Call oResult.Add(oDataReader.GetName(iField), oDataReader.GetString(iField))

                    Case "timestamp"
                        Call oResult.Add(oDataReader.GetName(iField), ConvertTimestampToString(CType(oDataReader.GetValue(iField), Byte())))

                    Case "bigint"  'Must be used only for Timestamp
                        Call oResult.Add(oDataReader.GetName(iField), Convert.ToString(oDataReader.GetInt64(iField)))

                    Case Else
                        ComponentServices.LogInformation("The following sql server data type string has not been recognized by activebank: " + oDataReader.GetDataTypeName(iField).ToLower(), "abComponentServices", "AddDataReaderToParameterCollection")

                End Select
            End If
        Next iField

    End Sub

    Public Function ExecuteQueryDataTable( _
            ByVal sQuery As String) As Data.DataTable Implements IABDatabaseAdapter.ExecuteQueryDataTable

        Dim oDataAdapter As SqlClient.SqlDataAdapter
        Dim oDataTable As Data.DataTable
        Dim sQueryDescription As New StringBuilder

        Try
            If m_iSPID <> 0 Then
                sQueryDescription.Append("SPID: ")
                sQueryDescription.Append(m_iSPID)
                sQueryDescription.Append(" : ")
                sQueryDescription.Append(sQuery)
            Else
                sQueryDescription.Append(sQuery)
            End If

            oDataAdapter = New SqlClient.SqlDataAdapter(sQuery, m_oConnection)
            oDataAdapter.SelectCommand.Transaction = m_oTransaction
            oDataAdapter.SelectCommand.CommandTimeout = m_iCommandTimeout

            oDataTable = New Data.DataTable
            oDataAdapter.Fill(oDataTable)

            Call ComponentServices.LogInformation( _
                    sQueryDescription.ToString, _
                    "abComponentServices", _
                    "ExecuteQueryDataTable")

        Catch oError As SqlClient.SqlException
            oDataTable = Nothing

            Throw New ActiveBankException(oError.Number, oError.Message)

        Catch oError As System.Exception
            oDataTable = Nothing

            ComponentServices.LogError(True, "abComponentServices", oError, sQueryDescription.ToString)
        Finally
            If Not oDataAdapter Is Nothing Then
                oDataAdapter.Dispose()
                oDataAdapter = Nothing
            End If
        End Try

        Return oDataTable

    End Function

    Public Sub AbortTransaction() Implements IABDatabaseAdapter.AbortTransaction

        If m_oTransaction Is Nothing Then
            Throw New ActiveBankException( _
                    "There is not currently a transaction started.", _
                    ActiveBankException.ExceptionType.System)
        Else
            Try
                m_oTransaction.Rollback()
            Catch
            Finally
                m_oTransaction.Dispose()
                m_oTransaction = Nothing
            End Try
        End If

    End Sub

    Public Sub CommitTransaction() Implements IABDatabaseAdapter.CommitTransaction

        If Not (m_oTransaction Is Nothing) Then
            Try
                m_oTransaction.Commit()

            Catch ex As System.Exception
                ComponentServices.LogError(True, "ComponentServices", ex, "CommitTransaction")
            Finally
                m_oTransaction.Dispose()
                m_oTransaction = Nothing
            End Try
        End If

    End Sub

    Public Function ExecuteStoredProcedure( _
            ByVal sName As String, _
            ByVal oParameterValues As ParameterCollection, _
            ByVal oAlternativeParameterValues As ParameterCollection) As ParameterCollection Implements IABDatabaseAdapter.ExecuteStoredProcedure

        Dim oResult As ParameterCollection
        Dim oCommand As SqlClient.SqlCommand
        Dim iParameter As Integer
        Dim iParameters As Integer
        Dim oStoredProcedureParameters As SqlClient.SqlParameterCollection
        Dim oStoredProcedureParameter As SqlClient.SqlParameter
        Dim strSQLServerParameterName As String
        Dim oParameterValue As Object
        Dim sbLog As StringBuilder
        Dim oResultParameterNames As ParameterCollection
        Dim sLogString As String
        Dim oDBNull As System.DBNull

        Try
            If ComponentServices.m_structConfigSettings.eLoggingLevel = ComponentServices.LoggingLevel.Full Then
                sbLog = New StringBuilder("Execute strored procedure: ")
                sbLog.Append(sName)
            Else
                sbLog = Nothing
            End If

            oCommand = GetStoredProcedureParameters(sName)
            oCommand.Connection = m_oConnection
            oCommand.Transaction = m_oTransaction

            oResultParameterNames = New ParameterCollection

            'Now need to populate values.
            oStoredProcedureParameters = oCommand.Parameters
            iParameters = oStoredProcedureParameters.Count - 1
            For iParameter = 0 To iParameters
                oStoredProcedureParameter = oStoredProcedureParameters.Item(iParameter)
                Select Case oStoredProcedureParameter.Direction
                    Case ParameterDirection.Input
                        oParameterValue = oParameterValues.GetValue(oStoredProcedureParameter.ParameterName.ToString)
                        'Check if this parameter has been supplied.
                        If oParameterValue Is Nothing Then
                            If oAlternativeParameterValues Is Nothing Then
                                oStoredProcedureParameter.Value = System.DBNull.Value
                            Else
                                oParameterValue = oAlternativeParameterValues.GetValue(oStoredProcedureParameter.ParameterName.ToString)
                                If oParameterValue Is Nothing Then
                                    oStoredProcedureParameter.Value = System.DBNull.Value
                                Else
                                    Select Case oStoredProcedureParameter.SqlDbType
                                        Case SqlDbType.BigInt
                                            oStoredProcedureParameter.Value = Convert.ToInt64(oParameterValue)
                                        Case Else
                                            oStoredProcedureParameter.Value = oParameterValue
                                    End Select
                                End If
                            End If
                        Else
                            Select Case oStoredProcedureParameter.SqlDbType
                                Case SqlDbType.BigInt
                                    oStoredProcedureParameter.Value = Convert.ToInt64(oParameterValue)
                                Case Else
                                    Select Case oParameterValue.GetType().FullName()
                                        Case "System.String"
                                            'vbback is used as a substitute for an empty string (to indicate an empty 
                                            'string is required, as opposed to the parameter has not been specified).
                                            If oParameterValue Is vbBack Then
                                                oParameterValue = ""
                                            End If
                                            'FromOADate(0) is used as a substitute for an NULL date value 
                                        Case "System.DateTime", "System.Date"
                                            If oParameterValue = System.DateTime.FromOADate(0) Then
                                                oParameterValue = System.DBNull.Value
                                            End If
                                    End Select
                                    oStoredProcedureParameter.Value = oParameterValue
                            End Select
                        End If

                        If Not (sbLog Is Nothing) Then
                            sbLog.Append(vbNewLine)
                            sbLog.Append(oStoredProcedureParameter.ParameterName)
                            sbLog.Append(": ")
                            If oStoredProcedureParameter.Value Is Nothing Then
                                sbLog.Append("Null")
                            Else
                                If Not oStoredProcedureParameter.Value Is DBNull.Value Then
                                    sLogString = CStr(oStoredProcedureParameter.Value)
                                    If sLogString.Length > 30000 Then
                                        sLogString = sLogString.Substring(1, 30000)
                                    End If
                                    sbLog.Append(sLogString)
                                Else
                                    sbLog.Append("Null")
                                End If
                            End If
                        End If

                    Case ParameterDirection.InputOutput, ParameterDirection.Output
                        AddSQLServerParameterToParameterCollection( _
                                oStoredProcedureParameter, _
                                oResultParameterNames)

                End Select
            Next iParameter

            If Not (sbLog Is Nothing) Then
                ComponentServices.LogInformation( _
                        sbLog.ToString(), _
                        "abComponentServices", _
                        "ExecuteStoredProcedure")
            End If

            oCommand.CommandTimeout = m_iCommandTimeout
            oCommand.ExecuteNonQuery()

            'AFS Now need to populate the result parameter collection.
            oResult = New ParameterCollection
            oResultParameterNames.Reset()
            While oResultParameterNames.MoveNext()
                strSQLServerParameterName = "@" + oResultParameterNames.GetName()
                oStoredProcedureParameter = oStoredProcedureParameters.Item(strSQLServerParameterName)
                AddSQLServerParameterToParameterCollection(oStoredProcedureParameter, oResult)
            End While

        Catch oError As ActiveBankException
            ComponentServices.LogError(True, oError)

        Catch oError As SqlClient.SqlException
            ComponentServices.LogError( _
                    True, _
                    oError.Number, _
                    oError.Message, _
                    "abComponentServices", _
                    oError.StackTrace, _
                    ComponentServices.ErrorSeverity.ES_Error)

            oResult = Nothing
        Finally
            If Not oCommand Is Nothing Then
                oCommand.Transaction = Nothing
                oCommand.Connection = Nothing
                oCommand.Dispose()
                oCommand = Nothing
            End If
        End Try

        Return oResult

    End Function

    Private Function GetStoredProcedureParameters( _
            ByVal sName As String) As SqlClient.SqlCommand

        Dim oCommand As SqlClient.SqlCommand
        Dim oConnection As SqlClient.SqlConnection
        Dim iRetries As Integer
        Dim oParameters As ParameterCollection
        Dim oParameter As SqlClient.SqlParameter

        oParameters = CType(ComponentServices.m_oStoredProcedureDefinitions.Item(sName), abComponentServices.ParameterCollection)

        If oParameters Is Nothing Then
            Try
                oConnection = New SqlClient.SqlConnection(m_sConnectionString)
                oConnection.Open()

                oCommand = New SqlClient.SqlCommand(sName, oConnection)
                oCommand.CommandType = CommandType.StoredProcedure

                SqlClient.SqlCommandBuilder.DeriveParameters(oCommand)

                oParameters = New ParameterCollection
                For Each oParameter In oCommand.Parameters
                    oParameters.Add(oParameter.ParameterName, oParameter)
                Next oParameter

                Try
                    ComponentServices.m_oStoredProcedureDefinitions.Add(sName, oParameters)
                Catch
                End Try

            Catch oError As SqlClient.SqlException
                ComponentServices.LogError( _
                        True, _
                        oError.Number, _
                        oError.Message, _
                        "abComponentServices", _
                        oError.StackTrace, _
                        ComponentServices.ErrorSeverity.ES_Error)

            Finally
                If Not oCommand Is Nothing Then
                    oCommand.Connection = Nothing
                End If
                If Not oConnection Is Nothing Then
                    If oConnection.State <> ConnectionState.Closed Then
                        oConnection.Close()
                    End If
                    oConnection.Dispose()
                    oConnection = Nothing
                End If
            End Try
        Else
            oCommand = PopulateCommandObjectFromCache(sName, oParameters)
        End If

        Return oCommand

    End Function

    Private Function PopulateCommandObjectFromCache(ByVal sName As String, ByRef oParameters As ParameterCollection) As SqlClient.SqlCommand

        Dim oCommand As SqlClient.SqlCommand
        Dim oParameter As SqlClient.SqlParameter
        Dim oCachedParameter As SqlClient.SqlParameter

        oCommand = New SqlClient.SqlCommand(sName)
        oCommand.CommandType = CommandType.StoredProcedure
        oParameters.Reset()
        While oParameters.MoveNext()
            oCachedParameter = oParameters.GetSQLCommandParameter()
            oParameter = New SqlClient.SqlParameter(oCachedParameter.ParameterName, CType(oCachedParameter.DbType, System.Data.SqlDbType), oCachedParameter.Size)
            oParameter.Precision = oCachedParameter.Precision
            oParameter.Scale = oCachedParameter.Scale
            oParameter.Direction = oCachedParameter.Direction
            oParameter.SqlDbType = oCachedParameter.SqlDbType

            oCommand.Parameters.Add(oParameter)
        End While

        Return oCommand

    End Function

    Private Sub AddSQLServerParameterToParameterCollection( _
            ByRef oParameter As SqlClient.SqlParameter, _
            ByRef oParameterCollection As ParameterCollection)

        Dim strName As String

        strName = oParameter.ParameterName.Substring(1)

        Select Case oParameter.SqlDbType
            Case SqlDbType.Bit
                oParameterCollection.Add(strName, CBool(oParameter.Value))

            Case SqlDbType.Char, SqlDbType.NVarChar, SqlDbType.VarChar, SqlDbType.NChar, SqlDbType.Text
                oParameterCollection.Add(strName, CStr(oParameter.Value))

            Case SqlDbType.DateTime
                oParameterCollection.Add(strName, CDate(oParameter.Value))

            Case SqlDbType.Decimal
                oParameterCollection.Add(strName, CDec(oParameter.Value))

            Case SqlDbType.Int
                oParameterCollection.Add(strName, CInt(oParameter.Value))

            Case SqlDbType.Float, SqlDbType.Money
                oParameterCollection.Add(strName, CDbl(oParameter.Value))

            Case SqlDbType.TinyInt
                oParameterCollection.Add(strName, CByte(oParameter.Value))

            Case SqlDbType.SmallInt
                oParameterCollection.Add(strName, Convert.ToInt16(oParameter.Value))

        End Select

    End Sub

    Private Sub SetSQLParameterToNull( _
            ByRef oParameter As SqlClient.SqlParameter)

        Select Case oParameter.SqlDbType
            Case SqlDbType.Bit
                oParameter.Value = SqlTypes.SqlBoolean.Null
            Case SqlDbType.Char, SqlDbType.NChar, SqlDbType.VarChar, SqlDbType.NVarChar
                oParameter.Value = SqlTypes.SqlString.Null
            Case SqlDbType.DateTime
                oParameter.Value = SqlTypes.SqlDateTime.Null
            Case SqlDbType.Decimal
                oParameter.Value = SqlTypes.SqlDecimal.Null
            Case SqlDbType.Int
                oParameter.Value = SqlTypes.SqlInt32.Null
            Case SqlDbType.TinyInt
                oParameter.Value = SqlTypes.SqlByte.Null
            Case Else
                oParameter.Value = Nothing

        End Select

    End Sub

    Private Function GetColumnOrdinal( _
            ByVal sName As String, _
            ByRef oDataReader As SqlClient.SqlDataReader) As Integer

        Dim iOrdinal As Integer

        Try
            iOrdinal = oDataReader.GetOrdinal(sName)
        Catch
            iOrdinal = -1
        End Try

        Return iOrdinal

    End Function

    Public Function ExecuteStoredProcedureDataTable( _
            ByVal sName As String, _
            ByVal pcParameterValues As ParameterCollection) As Data.DataTable Implements abComponentServices.IABDatabaseAdapter.ExecuteStoredProcedureDataTable

        Dim oResult As Data.DataTable
        Dim oCommand As SqlClient.SqlCommand
        Dim iParameter As Integer
        Dim iParameters As Integer
        Dim strSQLServerParameterName As String
        Dim oParameterValue As Object
        Dim sbLog As StringBuilder
        Dim oResultParameterNames As ParameterCollection
        Dim sbCommand As New StringBuilder("exec ")
        Dim oStoredProcedureParameters As SqlClient.SqlParameterCollection
        Dim oStoredProcedureParameter As SqlClient.SqlParameter
        Dim sLogString As String

        Try
            If ComponentServices.m_structConfigSettings.eLoggingLevel = ComponentServices.LoggingLevel.Full Then
                sbLog = New StringBuilder("Execute strored procedure: ")
                sbLog.Append(sName)
            Else
                sbLog = Nothing
            End If

            oCommand = GetStoredProcedureParameters(sName)
            oCommand.Connection = m_oConnection
            oCommand.Transaction = m_oTransaction
            oCommand.CommandTimeout = m_iCommandTimeout

            sbCommand.Append(sName)
            sbCommand.Append(" ")

            'AFS Now need to populate values.
            oStoredProcedureParameters = oCommand.Parameters
            iParameters = oStoredProcedureParameters.Count - 1
            For iParameter = 0 To iParameters
                oStoredProcedureParameter = oStoredProcedureParameters.Item(iParameter)
                Select Case oStoredProcedureParameter.Direction
                    Case ParameterDirection.Input
                        oParameterValue = pcParameterValues.GetValue(oStoredProcedureParameter.ParameterName.Substring(1))
                        'AFS Check if this parameter has been supplied.
                        If oParameterValue Is Nothing Then
                            oStoredProcedureParameter.Value = System.DBNull.Value
                        Else
                            oStoredProcedureParameter.Value = oParameterValue
                        End If

                        sbCommand.Append(oStoredProcedureParameter.ParameterName)
                        sbCommand.Append("=")

                        AppendStoredProcedureCommandString(oParameterValue, sbCommand)

                        If Not (sbLog Is Nothing) Then
                            sbLog.Append(vbNewLine)
                            sbLog.Append(oStoredProcedureParameter.ParameterName)
                            sbLog.Append(": ")
                            If oStoredProcedureParameter.Value Is Nothing Then
                                sbLog.Append("Null")
                            Else
                                If Not oStoredProcedureParameter.Value Is DBNull.Value Then
                                    sLogString = CStr(oStoredProcedureParameter.Value)
                                    If sLogString.Length > 30000 Then
                                        sLogString = sLogString.Substring(1, 30000)
                                    End If
                                    sbLog.Append(sLogString)
                                Else
                                    sbLog.Append("Null")
                                End If
                            End If
                        End If

                        sbCommand.Append(", ")
                End Select
            Next iParameter

            If iParameters > 0 Then
                sbCommand.Remove(sbCommand.Length - 2, 2)
            End If

            If Not (sbLog Is Nothing) Then
                ComponentServices.LogInformation( _
                        sbLog.ToString(), _
                        "abComponentServices", _
                        "ExecuteStoredProcedure")
            End If

            oResult = ExecuteQueryDataTable(sbCommand.ToString())

        Catch oError As ActiveBankException
            ComponentServices.LogError(True, oError)

        Catch oError As SqlClient.SqlException
            ComponentServices.LogError( _
                    True, _
                    oError.Number, _
                    oError.Message, _
                    "abComponentServices", _
                    oError.StackTrace, _
                    ComponentServices.ErrorSeverity.ES_Error)
        Finally
            If Not oCommand Is Nothing Then
                oCommand.Transaction = Nothing
                oCommand.Connection = Nothing
                oCommand.Dispose()
                oCommand = Nothing
            End If
        End Try

        Return oResult

    End Function

    Public Function GetStoredProcedureDataTypes( _
            ByVal sName As String) As abComponentServices.ParameterCollection Implements abComponentServices.IABDatabaseAdapter.GetStoredProcedureDataTypes

        Dim pcDataTypes As ParameterCollection
        Dim oCommand As SqlClient.SqlCommand
        Dim oParameters As SqlClient.SqlParameterCollection
        Dim iParameters As Integer
        Dim iParameter As Integer
        Dim oParameter As SqlClient.SqlParameter
        Dim sParameterName As String

        Try
            oCommand = GetStoredProcedureParameters(sName)

            pcDataTypes = New ParameterCollection

            oParameters = oCommand.Parameters
            iParameters = oParameters.Count - 1
            For iParameter = 0 To iParameters
                oParameter = oParameters.Item(iParameter)
                sParameterName = oParameter.ParameterName.Substring(1)

                Select Case oParameter.Direction
                    Case ParameterDirection.Input
                        Select Case oParameter.SqlDbType
                            Case SqlDbType.Bit
                                pcDataTypes.Add(sParameterName, VariantType.Boolean)

                            Case SqlDbType.Char, SqlDbType.NVarChar, SqlDbType.VarChar, SqlDbType.NChar
                                pcDataTypes.Add(sParameterName, VariantType.String)

                            Case SqlDbType.DateTime
                                pcDataTypes.Add(sParameterName, VariantType.Date)

                            Case SqlDbType.Decimal
                                pcDataTypes.Add(sParameterName, VariantType.Decimal)

                            Case SqlDbType.Int, SqlDbType.SmallInt
                                pcDataTypes.Add(sParameterName, VariantType.Integer)

                            Case SqlDbType.Float, SqlDbType.Money
                                pcDataTypes.Add(sParameterName, VariantType.Double)

                            Case SqlDbType.TinyInt
                                pcDataTypes.Add(sParameterName, VariantType.Byte)

                        End Select
                End Select
            Next iParameter

        Catch ex As System.Exception
            ComponentServices.LogError(True, "abComponentServices", ex)
        Finally
            If Not oCommand Is Nothing Then
                oCommand.Dispose()
                oCommand = Nothing
            End If
        End Try

        Return pcDataTypes

    End Function

    Public Function ConvertTimestampToString( _
            ByVal byteTimestamp() As Byte) As String Implements abComponentServices.IABDatabaseAdapter.ConvertTimestampToString

        Dim sbTimestamp As New StringBuilder
        Dim iByte As Integer

        For iByte = 0 To 7
            sbTimestamp.Append(Hex(byteTimestamp(iByte)).PadLeft(2, CChar("0")))
        Next iByte

        Return sbTimestamp.ToString()

    End Function

    Private Sub AppendStoredProcedureCommandString( _
            ByVal oValue As Object, _
            ByRef sbCommand As StringBuilder)

        Select Case oValue.GetType().FullName()
            Case "System.String"
                sbCommand.Append("'")
                sbCommand.Append(oValue)
                sbCommand.Append("'")

            Case "System.Int16", "System.Int32", "System.Int64"
                sbCommand.Append(oValue)

            Case "System.DateTime"
                If oValue Is Nothing Then
                    sbCommand.Append("NULL")
                Else
                    If CDate(oValue) = System.DateTime.FromOADate(0) Then
                        sbCommand.Append("NULL")
                    Else
                        sbCommand.Append("'")
                        sbCommand.Append(Format(oValue, "dd MMM yyyy"))
                        sbCommand.Append("'")
                    End If
                End If

            Case "System.Double"
                sbCommand.Append(oValue)

            Case "System.Decimal"
                sbCommand.Append(oValue)

            Case "System.Boolean"
                If CBool(oValue) Then
                    sbCommand.Append("1")
                Else
                    sbCommand.Append("0")
                End If

            Case Else
                sbCommand.Append("'")
                sbCommand.Append(oValue)
                sbCommand.Append("'")
        End Select

    End Sub

    Private Sub SetStoredProcedureParameterValues( _
        ByRef oCommand As SqlClient.SqlCommand, _
        ByVal alParameterValues As ArrayList, _
        ByVal sStoredProcedureName As String)

        '==================================================================================================
        ' Author    : Darryn Clerihew/Jatinder Virdee
        '--------------------------------------------------------------------------------------------------
        ' About...  : Sets the parameter values for a stored procedure using elements held in a list
        '             strucure.
        '
        '             Note: The implementation of this method needs to be refactored. The code was taken
        '             from the overloaded function GetStoredProcedureParameters which was causing major 
        '             performance issues with Oracle. Due to time constraints and other issues, a
        '             decision was made not to refactor the code at this time. See the TODO directive below
        '==================================================================================================

        'TODO: Use a different list structure (for e.g. a hash table) to hold the parameter values. This must contain the name (the key) and the value 
        'of the parameter (not just the value). The loop condition below must be changed to lookup the parameter name based on the key and then to
        'set the value accordingly.

        Dim oParameter As SqlClient.SqlParameter
        Dim oParameters As ParameterCollection
        Dim iParameterCount As Integer = 0
        Dim iCount As Integer
        Dim sValue As String
        Dim dtmDate As Date
        Dim iPos As Integer
        Dim sParams As String()
        Dim oDBNull As Object = System.DBNull.Value
        Dim sParameterName As String = String.Empty
        Dim sParameterValue As String = String.Empty
        Dim sSQLDataType As String = String.Empty

        Try
            If Not alParameterValues Is Nothing Then
                For iCount = 0 To oCommand.Parameters.Count - 1
                    oParameter = oCommand.Parameters(iCount)

                    Select Case oParameter.Direction
                        Case ParameterDirection.Input, ParameterDirection.InputOutput
                            sParameterName = oParameter.ParameterName
                            sParameterValue = String.Empty

                            If TypeOf alParameterValues(iParameterCount) Is System.DBNull Then
                                sParameterValue = "NULL"
                            Else
                                sParameterValue = CType(alParameterValues(iParameterCount), String)
                            End If
                            sSQLDataType = oParameter.SqlDbType.ToString

                            If ComponentServices.m_structConfigSettings.eLoggingLevel = ComponentServices.LoggingLevel.Full Then
                                sStoredProcedureName &= vbCrLf & sParameterName & ":" & sParameterValue & ":" & sSQLDataType
                            End If

                            Select Case oParameter.SqlDbType

                                Case SqlDbType.BigInt, SqlDbType.Int, SqlDbType.SmallInt, SqlDbType.TinyInt
                                    If alParameterValues(iParameterCount) Is Nothing Or TypeOf alParameterValues(iParameterCount) Is System.DBNull Then
                                        oParameter.Value = oDBNull
                                    Else
                                        oParameter.Value = CType(alParameterValues(iParameterCount), Integer)
                                    End If

                                Case SqlDbType.Decimal, SqlDbType.Float, SqlDbType.Money, SqlDbType.Real, SqlDbType.SmallMoney
                                    If alParameterValues(iParameterCount) Is Nothing Or TypeOf alParameterValues(iParameterCount) Is System.DBNull Then
                                        oParameter.Value = oDBNull
                                    Else
                                        oParameter.Value = CType(alParameterValues(iParameterCount), Decimal)
                                    End If

                                Case SqlDbType.Char, SqlDbType.NChar, SqlDbType.VarChar, SqlDbType.NText, SqlDbType.NVarChar
                                    If alParameterValues(iParameterCount) Is Nothing Or TypeOf alParameterValues(iParameterCount) Is System.DBNull Then
                                        oParameter.Value = oDBNull
                                    Else
                                        oParameter.Value = CType(alParameterValues(iParameterCount), String)
                                    End If

                                Case SqlDbType.DateTime, SqlDbType.SmallDateTime
                                    If alParameterValues(iParameterCount) Is Nothing Or TypeOf alParameterValues(iParameterCount) Is System.DBNull Then
                                        oParameter.Value = oDBNull
                                    Else
                                        sValue = alParameterValues(iParameterCount)
                                        If sValue <> String.Empty Then
                                            'handle the occrrance TO_DATE
                                            If sValue.IndexOf("TO_DATE") >= 0 Then
                                                sValue = sValue.Substring(sValue.IndexOf("(") + 1)
                                                sValue = sValue.Replace("'", "")
                                                sValue = sValue.Replace(")", "")
                                                sParams = sValue.Split(","c)

                                                If sParams(0) <> "00:00:00" Then
                                                    oParameter.Value = ComponentServices.GetISODataValue(sParams(0), sParams(1))
                                                Else
                                                    oParameter.Value = oDBNull
                                                End If
                                            Else
                                                If sValue <> "00:00:00" Then
                                                    oParameter.Value = ComponentServices.GetISODataValue(sValue)
                                                Else
                                                    oParameter.Value = oDBNull
                                                End If
                                            End If
                                        Else
                                            oParameter.Value = oDBNull
                                        End If
                                    End If

                            End Select

                            iParameterCount += 1

                            If iParameterCount > alParameterValues.Count - 1 Then
                                Exit For    'No more parameters provided in sql query so additional stored procedure parameters must have default values
                            End If

                        Case ParameterDirection.Output

                            oParameter.Direction = ParameterDirection.Output
                    End Select
                Next
            End If

        Catch oError As SqlClient.SqlException
            ComponentServices.LogError( _
                    True, _
                    oError.Number, _
                    oError.Message, _
                    "abComponentServices", _
                    oError.StackTrace, _
                    ComponentServices.ErrorSeverity.ES_Error)
        Catch oError As System.Exception
            ComponentServices.LogError(True, "abComponentServices", oError, "Error processing value for Parameter (" & sParameterName & ":" & sParameterValue & ":" & sSQLDataType & ")")
        Finally
            If ComponentServices.m_structConfigSettings.eLoggingLevel = ComponentServices.LoggingLevel.Full Then
                ComponentServices.LogInformation("SPID: " & m_iSPID & " Executing Stored Procedure : " & sStoredProcedureName, "abComponentServices", "SetStoredProcedureParameterValues")
            End If
        End Try
    End Sub

    Public Function ExecuteStoredProcedureParameters(ByVal sName As String, ByVal alParameterValues As System.Collections.ArrayList) As System.Data.DataTable Implements IABDatabaseAdapter.ExecuteStoredProcedureParameters
        '======================================================================
        ' Date Created : 27 February 2008 
        '----------------------------------------------------------------------
        ' About...     : Execute stored procedure using array list of parameters
        '======================================================================
        ' Change Log
        '
        ' Date       Author         Ref                         Comments
        ' ---------- ------------   ---------                   --------
        ' 27/02/2008 Jat Virdee     TD24359, TD24357, TD24426   Created
        '======================================================================
        Dim oDataAdapter As SqlClient.SqlDataAdapter
        Dim oDataTable As Data.DataTable
        Dim oCommand As SqlClient.SqlCommand
        Dim sQueryDescription As New StringBuilder

        Try
            If m_iSPID <> 0 Then
                sQueryDescription.Append("SPID: ")
                sQueryDescription.Append(m_iSPID)
                sQueryDescription.Append(" : ")
                sQueryDescription.Append(sName)
            Else
                sQueryDescription.Append(sName)
            End If

            'Get the stored procedure parameters (the cache will be used if possible)
            oCommand = GetStoredProcedureParameters(sName.Trim)

            'Set the parameter values based on the parsed values from the supplied SQL statement
            SetStoredProcedureParameterValues(oCommand, alParameterValues, sName)

            'Set the data adapter properties
            oCommand.Connection = m_oConnection

            oDataAdapter = New SqlClient.SqlDataAdapter(oCommand)
            oDataAdapter.SelectCommand.Transaction = m_oTransaction

            oDataTable = New Data.DataTable
            oDataAdapter.Fill(oDataTable)

            Call ComponentServices.LogInformation( _
                    sQueryDescription.ToString, _
                    "abComponentServices", _
                    "ExecuteStoredProcedureParameters")

        Catch oError As SqlClient.SqlException
            If oError.Number <> 972 Then
                ComponentServices.LogError( _
                        True, _
                        oError.Number, _
                        oError.Message, _
                        "abComponentServices", _
                        oError.StackTrace, _
                        ComponentServices.ErrorSeverity.ES_Error, _
                        sQueryDescription.ToString)
            End If

            oDataTable = Nothing
        Catch oError As System.Exception
            ComponentServices.LogError(True, "abComponentServices", oError, sName)

            oDataTable = Nothing
        Finally
            If Not oDataAdapter Is Nothing Then
                oDataAdapter.Dispose()
                oDataAdapter = Nothing
            End If

            If Not oCommand Is Nothing Then
                oCommand.Transaction = Nothing
                oCommand.Connection = Nothing
                oCommand.Dispose()
                oCommand = Nothing
            End If
        End Try

        Return oDataTable
    End Function
End Class

Friend Class OracleAdapter

    Implements IABDatabaseAdapter

    Private m_oConnection As OracleClient.OracleConnection
    Private m_oTransaction As OracleClient.OracleTransaction
    Private m_sConnectionString As String
    Private m_iSID As Integer
    Private m_iCommandTimeout As Integer

    Public Sub OpenConnection( _
            ByVal sServer As String, _
            ByVal sDatabaseName As String, _
            ByVal sUserName As String, _
            ByVal sPassword As String, _
            ByVal iTimeout As Integer) Implements IABDatabaseAdapter.OpenConnection

        Dim sConnectionString As String
        Dim sConnectionStringForLog As String
        Dim sbConnectionString As New StringBuilder("Data Source=")
        Dim iRowsAffected As Integer
        Dim pcSID As ParameterCollection

        sbConnectionString.Append(sServer)
        sbConnectionString.Append(";User=")
        sbConnectionString.Append(sDatabaseName)

        sConnectionStringForLog = sbConnectionString.ToString()

        sbConnectionString.Append(";Password=")
        sbConnectionString.Append(sPassword)
        sConnectionString = sbConnectionString.ToString()

        m_sConnectionString = sConnectionString
        m_iCommandTimeout = iTimeout

        Try
            m_oConnection = New OracleClient.OracleConnection(sConnectionString)
            m_oConnection.Open()

            If ComponentServices.m_structConfigSettings.eLoggingLevel = ComponentServices.LoggingLevel.Full Then
                pcSID = ExecuteQueryParameterCollection("select sid from v$session where audsid=userenv('SESSIONID')", True, False)
                m_iSID = pcSID.GetIntegerValue("sid", 0)

                ComponentServices.LogInformation("Oracle connection open - " & m_iSID.ToString(), "abComponentServices", "OpenConnection")
            End If

            ExecuteQuery("ALTER SESSION SET NLS_DATE_FORMAT = 'YYYY-MM-DD'", iRowsAffected)

        Catch oError As OracleClient.OracleException
            ComponentServices.LogError( _
                    True, _
                    oError.Code, _
                    oError.Message, _
                    "abComponentServices", _
                    oError.StackTrace, _
                    ComponentServices.ErrorSeverity.ES_Error, _
                    sConnectionStringForLog)
        End Try

    End Sub

    Public Sub CloseConnection() Implements IABDatabaseAdapter.CloseConnection

        Dim sLogInformation As StringBuilder

        Try
            If Not m_oConnection Is Nothing Then
                If Not m_oConnection.State = ConnectionState.Closed Then
                    m_oConnection.Close()
                End If
                m_oConnection.Dispose()
            End If

            If ComponentServices.m_structConfigSettings.eLoggingLevel = ComponentServices.LoggingLevel.Full Then
                sLogInformation = New StringBuilder("Oracle connection closed - SID : ")
                sLogInformation.Append(m_iSID.ToString(System.Globalization.CultureInfo.InvariantCulture))

                ComponentServices.LogInformation(sLogInformation.ToString, "abComponentServices", "ConnectionClosed")
            End If

        Catch ex As System.Exception
            ComponentServices.LogError(True, "abComponentServices", ex, "CloseConnection")

        Finally
            m_oConnection = Nothing
        End Try

    End Sub

    Public Sub StartTransaction() Implements IABDatabaseAdapter.StartTransaction
        'The default isolation level for oracle is read committed.
        If m_oTransaction Is Nothing Then
            m_oTransaction = m_oConnection.BeginTransaction(IsolationLevel.ReadCommitted)
        End If
    End Sub

    Public Sub AbortTransaction() Implements IABDatabaseAdapter.AbortTransaction

        If m_oTransaction Is Nothing Then
            Throw New ActiveBankException( _
                    "There is not currently a transaction started.", _
                    ActiveBankException.ExceptionType.System)
        Else
            Try
                m_oTransaction.Rollback()
            Catch
            Finally
                m_oTransaction.Dispose()
                m_oTransaction = Nothing
            End Try
        End If

    End Sub

    Public Sub CommitTransaction() Implements IABDatabaseAdapter.CommitTransaction

        If Not (m_oTransaction Is Nothing) Then
            Try
                m_oTransaction.Commit()

            Catch ex As System.Exception
                ComponentServices.LogError(True, "ComponentServices", ex, "CommitTransaction")
            Finally
                m_oTransaction.Dispose()
                m_oTransaction = Nothing
            End Try
        End If

    End Sub

    Private Function ExecuteQuery( _
            ByVal sQuery As String, _
            ByRef iRowsAffected As Integer) As OracleClient.OracleDataReader

        Dim oCommand As OracleClient.OracleCommand
        Dim oDataReader As OracleClient.OracleDataReader
        Dim sQueryDescription As New StringBuilder

        Try
            If m_iSID <> 0 Then
                sQueryDescription.Append("SID: ")
                sQueryDescription.Append(m_iSID)
                sQueryDescription.Append(" : ")
                sQueryDescription.Append(sQuery)
            Else
                sQueryDescription.Append(sQuery)
            End If

            oCommand = New OracleClient.OracleCommand( _
                    sQuery, _
                    m_oConnection)
            oCommand.Transaction = m_oTransaction
            oDataReader = oCommand.ExecuteReader()
            iRowsAffected = oDataReader.RecordsAffected

            Call ComponentServices.LogInformation( _
                    sQueryDescription.ToString, _
                    "abComponentServices", _
                    "ExecuteQuery")

        Catch oError As OracleClient.OracleException
            If oError.Code <> 972 Then
                ComponentServices.LogError( _
                        True, _
                        oError.Code, _
                        oError.Message, _
                        "abComponentServices", _
                        oError.StackTrace, _
                        ComponentServices.ErrorSeverity.ES_Error, _
                        sQueryDescription.ToString)
            Else
                If Not (oDataReader Is Nothing) Then
                    oDataReader.Close()
                    oDataReader.Dispose()
                    oDataReader = Nothing
                End If
            End If
        Finally
            If Not oCommand Is Nothing Then
                oCommand.Transaction = Nothing
                oCommand.Connection = Nothing
                oCommand.Dispose()
                oCommand = Nothing
            End If
        End Try

        Return oDataReader

    End Function

    Public Sub ExecuteNonQuery( _
            ByVal sQuery As String, _
            ByRef iRowsAffected As Integer) Implements IABDatabaseAdapter.ExecuteNonQuery

        Dim oDataReader As OracleClient.OracleDataReader

        Try
            oDataReader = ExecuteQuery(sQuery, iRowsAffected)

        Catch ex As ActiveBankException
            ComponentServices.LogError(True, ex)
        Catch ex As System.Exception
            ComponentServices.LogError(True, "abComponentServices", ex)
        Finally
            If Not oDataReader Is Nothing Then
                oDataReader.Close()
                oDataReader.Dispose()
                oDataReader = Nothing
            End If
        End Try

    End Sub

    Public Function ExecuteQueryParameterCollection( _
            ByVal sQuery As String, _
            ByVal bSingleRow As Boolean, _
            ByVal bIDOrdinal As Boolean) As ParameterCollection Implements IABDatabaseAdapter.ExecuteQueryParameterCollection

        Dim oDataReader As OracleClient.OracleDataReader
        Dim pcResult As ParameterCollection
        Dim pcResultRow As ParameterCollection
        Dim iIDOrdinal As Integer
        Dim iRow As Integer
        Dim iRowsAffected As Integer

        Try
            oDataReader = ExecuteQuery(sQuery, iRowsAffected)

            If oDataReader Is Nothing Then
                pcResult = Nothing
            Else
                pcResult = New ParameterCollection

                If bSingleRow Then
                    If oDataReader.Read() Then
                        AddDataReaderToParameterCollection( _
                                oDataReader, _
                                pcResult)
                    Else
                        pcResult = Nothing
                    End If
                Else
                    iRow = 1
                    ' Determine if we want to index by ID or row number
                    If bIDOrdinal Then
                        iIDOrdinal = GetColumnOrdinal("ID", oDataReader)
                        If iIDOrdinal = -1 Then
                            ' No ID column found
                            bIDOrdinal = False
                        End If
                    End If

                    While oDataReader.Read()
                        pcResultRow = New ParameterCollection

                        AddDataReaderToParameterCollection( _
                                oDataReader, _
                                pcResultRow)

                        If bIDOrdinal Then
                            pcResult.Add(CStr(oDataReader.GetDecimal(iIDOrdinal)), pcResultRow)
                        Else
                            pcResult.Add(CStr(iRow), pcResultRow)
                        End If

                        iRow += 1
                    End While

                    pcResult.Reset()
                End If
            End If

        Catch ex As ActiveBankException
            ComponentServices.LogError(True, ex)
        Catch ex As Exception
            ComponentServices.LogError(True, "abComponentServices", ex)
        Finally
            If Not oDataReader Is Nothing Then
                oDataReader.Close()
                oDataReader.Dispose()
                oDataReader = Nothing
            End If
        End Try

        Return pcResult

    End Function

    Private Sub AddDataReaderToParameterCollection( _
            ByRef oDataReader As OracleClient.OracleDataReader, _
            ByRef oResult As ParameterCollection)

        Dim iFields As Integer
        Dim iField As Integer

        iFields = oDataReader.FieldCount - 1
        For iField = 0 To iFields
            If Not oDataReader.IsDBNull(iField) Then
                Select Case oDataReader.GetDataTypeName(iField).ToLower()
                    Case "number"
                        Call oResult.Add(oDataReader.GetName(iField), oDataReader.GetDouble(iField))

                    Case "byte"
                        Call oResult.Add(oDataReader.GetName(iField), oDataReader.GetByte(iField))

                    Case "bit"
                        Call oResult.Add(oDataReader.GetName(iField), oDataReader.GetBoolean(iField))

                    Case "datetime"
                        Call oResult.Add(oDataReader.GetName(iField), oDataReader.GetDateTime(iField))

                    Case "decimal"
                        Call oResult.Add(oDataReader.GetName(iField), oDataReader.GetDecimal(iField))

                    Case "double"
                        Call oResult.Add(oDataReader.GetName(iField), oDataReader.GetDouble(iField))

                    Case "varchar2", "nvarchar", "nchar", "char", "nvarchar2"
                        Call oResult.Add(oDataReader.GetName(iField), oDataReader.GetString(iField))

                    Case "date"
                        Call oResult.Add(oDataReader.GetName(iField), oDataReader.GetDateTime(iField))

                    Case "rowid"
                        Call oResult.Add(oDataReader.GetName(iField), oDataReader.GetString(iField))

                    Case Else
                        ComponentServices.LogInformation("The following oracle data type string has not been recognized by activebank: " + oDataReader.GetDataTypeName(iField).ToLower(), "abComponentServices", "AddDataReaderToParameterCollection")

                End Select
            End If
        Next iField

    End Sub

    Private Function GetColumnOrdinal( _
        ByVal sName As String, _
        ByRef oDataReader As OracleClient.OracleDataReader) As Integer

        Dim iOrdinal As Integer

        Try
            iOrdinal = oDataReader.GetOrdinal(sName)
        Catch
            iOrdinal = -1
        End Try

        Return iOrdinal

    End Function

    Public Function ExecuteQueryDataTable( _
            ByVal strQuery As String) As Data.DataTable Implements IABDatabaseAdapter.ExecuteQueryDataTable

        Dim oDataAdapter As OracleClient.OracleDataAdapter
        Dim oDataTable As Data.DataTable
        Dim oCommand As OracleClient.OracleCommand
        Dim sQueryDescription As New StringBuilder
        Dim sName As String                     'The stored procedure name
        Dim alParameterValues As ArrayList      'An array list of the stored procedure's parameter values

        Try
            strQuery = strQuery.Trim()

            If m_iSID <> 0 Then
                sQueryDescription.Append("SID: ")
                sQueryDescription.Append(m_iSID)
                sQueryDescription.Append(" : ")
                sQueryDescription.Append(strQuery)
            Else
                sQueryDescription.Append(strQuery)
            End If

            If strQuery.ToUpper.Substring(0, 7) = "SELECT " _
            OrElse strQuery.ToUpper.Substring(0, 7) = "INSERT " _
            OrElse strQuery.ToUpper.Substring(0, 7) = "UPDATE " _
            OrElse strQuery.ToUpper.Substring(0, 7) = "DELETE " _
            OrElse strQuery.ToUpper.Substring(0, 9) = "TRUNCATE " Then

                oDataAdapter = New OracleClient.OracleDataAdapter(strQuery, m_oConnection)
                oDataAdapter.SelectCommand.CommandType = CommandType.Text
                oDataAdapter.SelectCommand.Transaction = m_oTransaction
            Else

                'Get the stored procedure name and the parameter values from the supplied SQL statement
                alParameterValues = DeriveParameterValues(strQuery, sName)

                'Get the stored procedure parameters (the cache will be used if possible)
                oCommand = GetStoredProcedureParameters(sName.Trim)

                'Set the parameter values based on the parsed values from the supplied SQL statement
                SetStoredProcedureParameterValues(oCommand, alParameterValues, sName)

                'Set the data adapter properties
                oCommand.Connection = m_oConnection

                oDataAdapter = New OracleClient.OracleDataAdapter(oCommand)
                oDataAdapter.SelectCommand.Transaction = m_oTransaction
            End If

            oDataTable = New Data.DataTable
            oDataAdapter.Fill(oDataTable)

            Call ComponentServices.LogInformation( _
                    sQueryDescription.ToString, _
                    "abComponentServices", _
                    "ExecuteQueryDataTable")

        Catch oError As OracleClient.OracleException
            If oError.Code <> 972 Then
                ComponentServices.LogError( _
                        True, _
                        oError.Code, _
                        oError.Message, _
                        "abComponentServices", _
                        oError.StackTrace, _
                        ComponentServices.ErrorSeverity.ES_Error, _
                        sQueryDescription.ToString)
            End If

            oDataTable = Nothing
        Catch oError As System.Exception
            ComponentServices.LogError(True, "abComponentServices", oError, sQueryDescription.ToString)

            oDataTable = Nothing
        Finally
            If Not oDataAdapter Is Nothing Then
                oDataAdapter.Dispose()
                oDataAdapter = Nothing
            End If
        End Try

        Return oDataTable

    End Function

    Public Function ExecuteStoredProcedure( _
            ByVal strName As String, _
            ByVal oParameters As ParameterCollection, _
            ByVal oAlternativeParameters As ParameterCollection) As ParameterCollection Implements IABDatabaseAdapter.ExecuteStoredProcedure

        Dim oResult As ParameterCollection
        Dim oCommand As OracleClient.OracleCommand
        Dim iParameter As Integer
        Dim iParameters As Integer
        Dim oStoredProcedureParameters As OracleClient.OracleParameterCollection
        Dim oStoredProcedureParameter As OracleClient.OracleParameter
        Dim strOracleParameterName As String
        Dim oResultParameterNames As ParameterCollection
        Dim oParameterValue As Object
        Dim sbLog As StringBuilder

        Try
            If ComponentServices.m_structConfigSettings.eLoggingLevel = ComponentServices.LoggingLevel.Full Then
                sbLog = New StringBuilder("Execute strored procedure: ")
                sbLog.Append(strName)
            Else
                sbLog = Nothing
            End If

            oCommand = GetStoredProcedureParameters(strName)
            oCommand.Connection = m_oConnection
            oCommand.Transaction = m_oTransaction

            oResultParameterNames = New ParameterCollection

            'AFS Now need to populate values.
            oStoredProcedureParameters = oCommand.Parameters
            iParameters = oStoredProcedureParameters.Count - 1
            For iParameter = 0 To iParameters
                oStoredProcedureParameter = oStoredProcedureParameters.Item(iParameter)
                Select Case oStoredProcedureParameter.Direction
                    Case ParameterDirection.Input
                        oParameterValue = oParameters.GetValue(oStoredProcedureParameter.ParameterName.ToString)
                        'AFS Check if this parameter has been supplied.
                        If oParameterValue Is Nothing Then
                            If oAlternativeParameters Is Nothing Then
                                oStoredProcedureParameter.Value = System.DBNull.Value
                            Else
                                oParameterValue = oAlternativeParameters.GetValue(oStoredProcedureParameter.ParameterName.ToString)
                                If oParameterValue Is Nothing Then
                                    Select Case oStoredProcedureParameter.OracleType
                                        Case OracleClient.OracleType.Int16, OracleClient.OracleType.Int32, OracleClient.OracleType.Number, OracleClient.OracleType.UInt16, OracleClient.OracleType.UInt32
                                            oStoredProcedureParameter.Value = 0
                                        Case Else
                                            oStoredProcedureParameter.Value = String.Empty
                                    End Select
                                Else
                                    oStoredProcedureParameter.Value = oParameterValue
                                End If
                            End If
                        Else
                            oStoredProcedureParameter.Value = oParameterValue
                        End If

                        If Not (sbLog Is Nothing) Then
                            sbLog.Append(vbNewLine)
                            sbLog.Append(oStoredProcedureParameter.ParameterName)
                            sbLog.Append(": ")
                            If oStoredProcedureParameter.Value Is Nothing Then
                                sbLog.Append("Null")
                            Else
                                sbLog.Append(oStoredProcedureParameter.Value)
                            End If
                        End If

                    Case ParameterDirection.InputOutput, ParameterDirection.Output
                        AddOracleParameterToParameterCollection( _
                                oStoredProcedureParameter, _
                                oResultParameterNames)

                End Select
            Next iParameter

            If Not (sbLog Is Nothing) Then
                ComponentServices.LogInformation( _
                        sbLog.ToString(), _
                        "abComponentServices", _
                        "ExecuteStoredProcedure")
            End If

            oCommand.ExecuteNonQuery()

            'AFS Now need to populate the result parameter collection.
            oResult = New ParameterCollection
            oResultParameterNames.Reset()
            While oResultParameterNames.MoveNext()
                strOracleParameterName = oResultParameterNames.GetName()
                oStoredProcedureParameter = oStoredProcedureParameters.Item(strOracleParameterName)
                AddOracleParameterToParameterCollection(oStoredProcedureParameter, oResult)
            End While

        Catch oError As ActiveBankException
            ComponentServices.LogError(True, oError)

        Catch oError As OracleClient.OracleException
            ComponentServices.LogError( _
                    True, _
                    oError.Code, _
                    oError.Message, _
                    "abComponentServices", _
                    oError.StackTrace, _
                    ComponentServices.ErrorSeverity.ES_Error)

            oResult = Nothing
        Finally
            If Not oCommand Is Nothing Then
                oCommand.Transaction = Nothing
                oCommand.Connection = Nothing
                oCommand.Dispose()
                oCommand = Nothing
            End If
        End Try

        Return oResult

    End Function

    Private Function GetStoredProcedureParameters( _
        ByVal sName As String) As OracleClient.OracleCommand

        Dim oCommand As OracleClient.OracleCommand
        Dim oConnection As OracleClient.OracleConnection

        Dim oParameters As ParameterCollection
        Dim oParameter As OracleClient.OracleParameter
        Dim iKey As Integer = 0

        oParameters = CType(ComponentServices.m_oStoredProcedureDefinitions.Item(sName), abComponentServices.ParameterCollection)

        If oParameters Is Nothing Then

            Debug.Write(sName + " - Parameters not cached")

            Try
                oConnection = New OracleClient.OracleConnection(m_sConnectionString)
                oConnection.Open()

                oCommand = New OracleClient.OracleCommand(sName, oConnection)
                oCommand.CommandType = CommandType.StoredProcedure

                OracleClient.OracleCommandBuilder.DeriveParameters(oCommand)

                oParameters = New ParameterCollection
                For Each oParameter In oCommand.Parameters
                    oParameters.Add(iKey.ToString, oParameter)
                    iKey += 1
                Next oParameter

                Try
                    ComponentServices.m_oStoredProcedureDefinitions.Add(sName, oParameters)
                Catch
                End Try

            Catch oError As OracleClient.OracleException
                ComponentServices.LogError( _
                        True, _
                        oError.Code, _
                        oError.Message, _
                        "abComponentServices", _
                        oError.StackTrace, _
                        ComponentServices.ErrorSeverity.ES_Error)

            Finally
                If Not oCommand Is Nothing Then
                    oCommand.Connection = Nothing
                End If
                If Not oConnection Is Nothing Then
                    If oConnection.State <> ConnectionState.Closed Then
                        oConnection.Close()
                    End If
                    oConnection.Dispose()
                    oConnection = Nothing
                End If
            End Try
        Else

            Debug.Write(sName + " - Parameters cached")

            oCommand = PopulateCommandObjectFromCache(sName, oParameters)
        End If

        Return oCommand

    End Function

    Private Sub GetStoredProcedureParameters( _
            ByVal strQuery As String, _
            ByRef oQueryCommand As OracleClient.OracleCommand)

        Dim oCommand As OracleClient.OracleCommand
        Dim oConnection As OracleClient.OracleConnection
        Dim oParameters As ParameterCollection
        Dim oParameter As OracleClient.OracleParameter
        Dim sSQL As String
        Dim sTempArray As ArrayList
        Dim iCount As Integer
        Dim iValue As Integer
        Dim sValue As String
        Dim dtmDate As Date
        Dim dValue As Decimal

        strQuery = strQuery.Trim()
        sTempArray = DeriveParameterValues(strQuery, sSQL)
        oQueryCommand.CommandText = sSQL
        oQueryCommand.CommandType = CommandType.StoredProcedure

        If oParameters Is Nothing Then
            Try
                oConnection = New OracleClient.OracleConnection(m_sConnectionString)
                oConnection.Open()

                oCommand = New OracleClient.OracleCommand(sSQL, oConnection)
                oCommand.CommandType = CommandType.StoredProcedure

                OracleClient.OracleCommandBuilder.DeriveParameters(oCommand)
                oCommand.Connection = Nothing

                For iCount = 0 To oCommand.Parameters.Count - 1
                    oParameter = oCommand.Parameters(iCount)
                    Select Case oParameter.Direction
                        Case ParameterDirection.Input, ParameterDirection.InputOutput
                            Select Case oParameter.OracleType
                                Case OracleClient.OracleType.Int16, OracleClient.OracleType.Int32, OracleClient.OracleType.UInt16, OracleClient.OracleType.UInt32
                                    iValue = CType(sTempArray(iCount), Integer)
                                    oQueryCommand.Parameters.Add(oParameter.ParameterName, oParameter.OracleType, oParameter.Size).Value = iValue
                                Case OracleClient.OracleType.Number
                                    dValue = CType(sTempArray(iCount), Decimal)
                                    oQueryCommand.Parameters.Add(oParameter.ParameterName, oParameter.OracleType, oParameter.Size).Value = dValue
                                Case OracleClient.OracleType.Char, OracleClient.OracleType.NChar, OracleClient.OracleType.VarChar, OracleClient.OracleType.NVarChar
                                    sValue = sTempArray(iCount)
                                    oQueryCommand.Parameters.Add(oParameter.ParameterName, oParameter.OracleType, oParameter.Size).Value = sValue
                                Case OracleClient.OracleType.DateTime
                                    sValue = sTempArray(iCount)
                                    If sValue <> String.Empty Then
                                        dtmDate = ComponentServices.GetISODataValue(sValue)
                                    Else
                                        dtmDate = New Date(1899, 12, 31)
                                    End If
                                    oQueryCommand.Parameters.Add(oParameter.ParameterName, oParameter.OracleType, oParameter.Size).Value = dtmDate
                            End Select

                        Case ParameterDirection.Output
                            oQueryCommand.Parameters.Add(oParameter.ParameterName, oParameter.OracleType).Direction = ParameterDirection.Output
                    End Select

                Next

            Catch oError As OracleClient.OracleException
                ComponentServices.LogError( _
                        True, _
                        oError.Code, _
                        oError.Message, _
                        "abComponentServices", _
                        oError.StackTrace, _
                        ComponentServices.ErrorSeverity.ES_Error)

            Finally
                If Not oCommand Is Nothing Then
                    oCommand.Connection = Nothing
                    oCommand.Dispose()
                    oCommand = Nothing
                End If
                If Not oConnection Is Nothing Then
                    If oConnection.State <> ConnectionState.Closed Then
                        oConnection.Close()
                    End If
                    oConnection.Dispose()
                    oConnection = Nothing
                End If
            End Try
        Else
            oQueryCommand = PopulateCommandObjectFromCache(sSQL, oParameters)
        End If

    End Sub

    Private Function DeriveParameterValues(ByVal strQuery As String, _
     ByRef sSQL As String) As ArrayList

        Dim iPrevPos As Integer
        Dim iPos As Integer
        Dim sValues As String = String.Empty
        Dim sArrayList As ArrayList
        Dim sTemp As String = String.Empty
        Dim sPrevTemp As String = String.Empty
        Dim oDBNull As Object = System.DBNull.Value
        Dim bNullValue As Boolean
        Dim bClosingQuoteFound As Boolean
        Dim iCurrentPosition As Integer
        Dim iClosingQuotePosition As Integer

        If strQuery.ToUpper.Substring(0, 7) = "EXECUTE" Then
            strQuery = Replace(strQuery, "execute ", "", , , CompareMethod.Text)
        ElseIf strQuery.ToUpper.Substring(0, 4) = "EXEC" Then
            strQuery = Replace(strQuery, "exec ", "", , , CompareMethod.Text)
        ElseIf strQuery.ToUpper.Substring(0, 4) = "CALL" Then
            strQuery = Replace(strQuery, "call ", "", , , CompareMethod.Text)
        End If

        iPos = strQuery.IndexOf(" ")

        If iPos = -1 Then
            iPos = strQuery.IndexOf("(")

            If iPos > 0 Then
                sSQL = strQuery.Substring(0, iPos).Trim
                sValues = strQuery.Substring(iPos, strQuery.Length - iPos).Trim
            Else
                sSQL = strQuery
            End If
        ElseIf iPos > 0 Then
            sSQL = strQuery.Substring(0, iPos).Trim
            sValues = strQuery.Substring(iPos, strQuery.Length - iPos).Trim
            iPos = strQuery.IndexOf("(")
            If iPos > 0 Then
                sSQL = strQuery.Substring(0, iPos).Trim
                sValues = strQuery.Substring(iPos, strQuery.Length - iPos).Trim
            End If
        End If

        If sSQL.IndexOf("(") > 0 Then
            sSQL = sSQL.Substring(1, sSQL.Length - 1).Trim
        End If

        If sValues.Length > 0 Then
            If sValues.Substring(0, 1) = "(" Then
                sValues = sValues.Substring(1, sValues.Length - 2).Trim
            ElseIf sValues.Substring(sValues.Length - 1, 1) = ")" Then
                sValues = sValues.Substring(0, sValues.Length - 1).Trim
            End If
        End If

        'Now parse the values
        iPos = 1
        iPrevPos = 0
        If sValues.Length > 0 Then
            sArrayList = New ArrayList

            While iPos <= sValues.Length
                iPos = sValues.IndexOf(",", iPrevPos)

                If iPos <= 0 Then
                    If iPrevPos = 0 Then
                        If sValues.Substring(0, 1) = "'" Then
                            sArrayList.Add(sValues.Substring(1, sValues.Length - 2))
                        Else
                            sArrayList.Add(sValues)
                        End If
                    Else
                        sTemp = sValues.Substring(iPrevPos).Trim

                        If sTemp.ToUpper = "NULL" Then
                            bNullValue = True
                        Else
                            bNullValue = False
                        End If

                        'To handle TO_DATE
                        If sPrevTemp <> String.Empty Then
                            sTemp = sPrevTemp + "," + sTemp
                            sPrevTemp = String.Empty
                        End If

                        'Handle begining and ending quotes
                        If sTemp.Length > 0 Then    'Substring will have problem if string is empty (null case)
                            If sTemp.Substring(0, 1) = "'" Then
                                sArrayList.Add(sTemp.Substring(1, sTemp.Length - 2))
                            Else
                                If bNullValue Then
                                    sArrayList.Add(oDBNull)
                                Else
                                    sArrayList.Add(sTemp)
                                End If
                            End If
                        Else
                            sArrayList.Add(sTemp)
                        End If
                    End If
                    Exit While
                Else
                    sTemp = sValues.Substring(iPrevPos, iPos - iPrevPos).Trim

                    If sTemp.Length > 0 Then
                        'handle value within single quotes
                        If sTemp.Substring(0, 1) = "'" AndAlso sPrevTemp = String.Empty Then
                            iPos = sValues.IndexOf("'", iPrevPos) + 2
                            bClosingQuoteFound = False
                            For iCurrentPosition = iPos To sValues.Length
                                If bClosingQuoteFound Then
                                    Select Case Mid(sValues, iCurrentPosition, 1)
                                        Case " "
                                        Case ","
                                            Exit For
                                        Case Else
                                            bClosingQuoteFound = False
                                    End Select
                                Else
                                    If Mid(sValues, iCurrentPosition, 1) = "'" Then
                                        bClosingQuoteFound = True
                                        iClosingQuotePosition = iCurrentPosition - 1
                                    End If
                                End If
                            Next iCurrentPosition
                            If (iPos - 1) <> iClosingQuotePosition Then
                                sArrayList.Add(sValues.Substring(iPos - 1, iClosingQuotePosition - (iPos - 1)))
                            Else
                                sArrayList.Add(String.Empty)
                            End If

                            iPos = iClosingQuotePosition
                        Else
                            If sTemp.ToUpper = "NULL" Then
                                bNullValue = True
                            Else
                                bNullValue = False
                            End If

                            'handle the occrrance TO_DATE
                            If sTemp.IndexOf("TO_DATE") >= 0 Then
                                sPrevTemp = sTemp
                            Else
                                'To handle TO_DATE
                                If sPrevTemp <> String.Empty Then
                                    sTemp = sPrevTemp + "," + sTemp
                                    sPrevTemp = String.Empty
                                End If

                                'Handle begining and ending quotes
                                If sTemp.Length > 0 Then    'Substring will have problem if string is empty (null case)
                                    If sTemp.Substring(0, 1) = "'" Then
                                        sArrayList.Add(sTemp.Substring(1, sTemp.Length - 2))
                                    Else
                                        If bNullValue Then
                                            sArrayList.Add(oDBNull)
                                        Else
                                            sArrayList.Add(sTemp)
                                        End If
                                    End If
                                Else
                                    sArrayList.Add(sTemp)
                                End If
                            End If
                        End If
                    End If

                End If

                iPrevPos = iPos + 1
            End While
        End If

        Return sArrayList
    End Function

    Private Function PopulateCommandObjectFromCache( _
            ByVal sName As String, _
            ByRef oParameters As ParameterCollection) As OracleClient.OracleCommand

        Dim oCommand As OracleClient.OracleCommand
        Dim oParameter As OracleClient.OracleParameter
        Dim oCachedParameter As OracleClient.OracleParameter
        Dim iCount As Integer = 0

        oCommand = New OracleClient.OracleCommand(sName)
        oCommand.CommandType = CommandType.StoredProcedure
        oParameters.Reset()
        For iCount = 0 To oParameters.Length - 1
            oCachedParameter = oParameters.GetOracleCommandParameter(iCount.ToString)
            oParameter = New OracleClient.OracleParameter(oCachedParameter.ParameterName, oCachedParameter.OracleType, oCachedParameter.Size)
            oParameter.Precision = oCachedParameter.Precision
            oParameter.Scale = oCachedParameter.Scale
            oParameter.Direction = oCachedParameter.Direction
            oParameter.DbType = oCachedParameter.DbType
            oCommand.Parameters.Add(oParameter)
        Next

        Return oCommand

    End Function

    Private Sub AddOracleParameterToParameterCollection( _
            ByRef oParameter As OracleClient.OracleParameter, _
            ByRef oParameterCollection As ParameterCollection)

        Dim strName As String

        strName = oParameter.ParameterName

        Select Case oParameter.OracleType
            Case Data.OracleClient.OracleType.NChar, Data.OracleClient.OracleType.NVarChar, OracleClient.OracleType.Char, OracleClient.OracleType.VarChar, OracleClient.OracleType.Clob, OracleClient.OracleType.LongRaw, OracleClient.OracleType.LongVarChar, OracleClient.OracleType.NClob
                oParameterCollection.Add(strName, CStr(oParameter.Value))

            Case Data.OracleClient.OracleType.DateTime
                oParameterCollection.Add(strName, CDate(oParameter.Value))

            Case Data.OracleClient.OracleType.Int16, Data.OracleClient.OracleType.Int32
                oParameterCollection.Add(strName, CInt(oParameter.Value))

            Case Data.OracleClient.OracleType.Double, Data.OracleClient.OracleType.Float
                oParameterCollection.Add(strName, CDbl(oParameter.Value))

            Case Data.OracleClient.OracleType.Byte
                oParameterCollection.Add(strName, CByte(oParameter.Value))

            Case Data.OracleClient.OracleType.Number
                oParameterCollection.Add(strName, CDec(oParameter.Value))

        End Select

    End Sub

    Public Function ExecuteStoredProcedureDataTable( _
            ByVal sName As String, _
            ByVal pcParameterValues As ParameterCollection) As Data.DataTable Implements abComponentServices.IABDatabaseAdapter.ExecuteStoredProcedureDataTable

        Dim oResult As System.Data.DataTable
        Dim oCommand As OracleClient.OracleCommand
        Dim iParameter As Integer
        Dim iParameters As Integer
        Dim oParameterValue As Object
        Dim sbLog As StringBuilder
        Dim sbCommand As New StringBuilder("exec ")
        Dim oStoredProcedureParameters As OracleClient.OracleParameterCollection
        Dim oStoredProcedureParameter As OracleClient.OracleParameter
        Dim sLogString As String

        Try
            If ComponentServices.m_structConfigSettings.eLoggingLevel = ComponentServices.LoggingLevel.Full Then
                sbLog = New StringBuilder("Execute strored procedure: ")
                sbLog.Append(sName)
            Else
                sbLog = Nothing
            End If

            oCommand = GetStoredProcedureParameters(sName)
            oCommand.Connection = m_oConnection
            oCommand.Transaction = m_oTransaction

            sbCommand.Append(sName)
            sbCommand.Append(" ")

            oStoredProcedureParameters = oCommand.Parameters
            iParameters = oStoredProcedureParameters.Count - 1

            For iParameter = 0 To iParameters
                oStoredProcedureParameter = oStoredProcedureParameters.Item(iParameter)
                Select Case oStoredProcedureParameter.Direction
                    Case ParameterDirection.Input
                        oParameterValue = pcParameterValues.GetValue(oStoredProcedureParameter.ParameterName.Substring(0))
                        If oParameterValue Is Nothing Then
                            oStoredProcedureParameter.Value = System.DBNull.Value
                            sbCommand.Append("''")
                        Else
                            oStoredProcedureParameter.Value = oParameterValue
                            AppendStoredProcedureCommandString(oParameterValue, sbCommand)
                        End If

                        'sbCommand.Append(oStoredProcedureParameter.ParameterName)
                        'sbCommand.Append("=")

                        If Not (sbLog Is Nothing) Then
                            sbLog.Append(vbNewLine)
                            sbLog.Append(oStoredProcedureParameter.ParameterName)
                            sbLog.Append(": ")
                            If oStoredProcedureParameter.Value Is Nothing Then
                                sbLog.Append("Null")
                            Else
                                If Not oStoredProcedureParameter.Value Is DBNull.Value Then
                                    sLogString = CStr(oStoredProcedureParameter.Value)
                                    If sLogString.Length > 30000 Then
                                        sLogString = sLogString.Substring(1, 30000)
                                    End If
                                    sbLog.Append(sLogString)
                                Else
                                    sbLog.Append("Null")
                                End If
                            End If
                        End If

                        sbCommand.Append(", ")
                End Select

            Next iParameter

            If iParameters > 0 Then
                sbCommand.Remove(sbCommand.Length - 2, 2)
            End If

            If Not (sbLog Is Nothing) Then
                ComponentServices.LogInformation( _
                        sbLog.ToString(), _
                        "abComponentServices", _
                        "ExecuteStoredProcedure")
            End If

            oResult = ExecuteQueryDataTable(sbCommand.ToString())

        Catch oError As ActiveBankException
            ComponentServices.LogError(True, oError)

        Catch oError As SqlClient.SqlException
            ComponentServices.LogError( _
                    True, _
                    oError.Number, _
                    oError.Message, _
                    "abComponentServices", _
                    oError.StackTrace, _
                    ComponentServices.ErrorSeverity.ES_Error)
        Finally
            If Not oCommand Is Nothing Then
                oCommand.Transaction = Nothing
                oCommand.Connection = Nothing
                oCommand.Dispose()
                oCommand = Nothing
            End If
        End Try

        Return oResult

    End Function

    Public Function GetStoredProcedureDataTypes( _
            ByVal sName As String) As abComponentServices.ParameterCollection Implements abComponentServices.IABDatabaseAdapter.GetStoredProcedureDataTypes

        Dim pcDataTypes As ParameterCollection
        Dim oCommand As OracleClient.OracleCommand
        Dim oParameters As OracleClient.OracleParameterCollection
        Dim iParameters As Integer
        Dim iParameter As Integer
        Dim oParameter As OracleClient.OracleParameter
        Dim sParameterName As String

        Try
            oCommand = GetStoredProcedureParameters(sName)
            pcDataTypes = New ParameterCollection

            oParameters = oCommand.Parameters
            iParameters = oParameters.Count - 1

            For iParameter = 0 To iParameters
                oParameter = oParameters.Item(iParameter)
                sParameterName = oParameter.ParameterName.Substring(1)

                Select Case oParameter.Direction
                    Case ParameterDirection.Input
                        Select Case oParameter.OracleType
                            Case Data.OracleClient.OracleType.NChar, Data.OracleClient.OracleType.NVarChar
                                pcDataTypes.Add(sParameterName, CStr(oParameter.Value))

                            Case Data.OracleClient.OracleType.DateTime
                                pcDataTypes.Add(sParameterName, CDate(oParameter.Value))

                            Case Data.OracleClient.OracleType.Int16, Data.OracleClient.OracleType.Int32
                                pcDataTypes.Add(sParameterName, CInt(oParameter.Value))

                            Case Data.OracleClient.OracleType.Double, Data.OracleClient.OracleType.Float
                                pcDataTypes.Add(sParameterName, CDbl(oParameter.Value))

                            Case Data.OracleClient.OracleType.Byte
                                pcDataTypes.Add(sParameterName, CByte(oParameter.Value))

                            Case Data.OracleClient.OracleType.Number
                                pcDataTypes.Add(sParameterName, CDec(oParameter.Value))
                        End Select

                End Select
            Next iParameter

        Catch ex As System.Exception
            ComponentServices.LogError(True, "abComponentServices", ex)
        Finally
            If Not oCommand Is Nothing Then
                oCommand.Dispose()
                oCommand = Nothing
            End If
        End Try

        Return pcDataTypes

    End Function

    Public Function ConvertTimestampToString( _
            ByVal byteTimestamp() As Byte) As String Implements abComponentServices.IABDatabaseAdapter.ConvertTimestampToString

        Dim sbTimestamp As New StringBuilder
        Dim iByte As Integer

        For iByte = 0 To 7
            sbTimestamp.Append(Hex(byteTimestamp(iByte)).PadLeft(2, CChar("0")))
        Next iByte

        Return sbTimestamp.ToString()

    End Function

    Private Sub AppendStoredProcedureCommandString( _
        ByVal oValue As Object, _
        ByRef sbCommand As StringBuilder)

        Select Case oValue.GetType().FullName()
            Case "System.String"
                sbCommand.Append("'")
                sbCommand.Append(oValue)
                sbCommand.Append("'")

            Case "System.Int16", "System.Int32", "System.Int64"
                sbCommand.Append(oValue)

            Case "System.DateTime"
                If oValue Is Nothing Then
                    sbCommand.Append("NULL")
                Else
                    If CDate(oValue) = System.DateTime.FromOADate(0) Then
                        sbCommand.Append("NULL")
                    Else
                        sbCommand.Append("'")
                        sbCommand.Append(Format(oValue, "YYYY-MM-DD"))
                        sbCommand.Append("'")
                    End If
                End If

            Case "System.Double"
                sbCommand.Append(oValue)

            Case "System.Decimal"
                sbCommand.Append(oValue)

            Case "System.Boolean"
                If CBool(oValue) Then
                    sbCommand.Append("1")
                Else
                    sbCommand.Append("0")
                End If
            Case Else
                sbCommand.Append("'")
                sbCommand.Append(oValue)
                sbCommand.Append("'")
        End Select

    End Sub

    Private Sub SetStoredProcedureParameterValues( _
        ByRef oCommand As OracleClient.OracleCommand, _
        ByVal alParameterValues As ArrayList, _
        ByVal sStoredProcedureName As String)

        '==================================================================================================
        ' Author    : Darryn Clerihew
        '--------------------------------------------------------------------------------------------------
        ' About...  : Sets the parameter values for a stored procedure using elements held in a list
        '             strucure.
        '
        '             Note: The implementation of this method needs to be refactored. The code was taken
        '             from the overloaded function GetStoredProcedureParameters which was causing major 
        '             performance issues with Oracle. Due to time constraints and other issues, a
        '             decision was made not to refactor the code at this time. See the TODO directive below
        '==================================================================================================

        'TODO: Use a different list structure (for e.g. a hash table) to hold the parameter values. This must contain the name (the key) and the value 
        'of the parameter (not just the value). The loop condition below must be changed to lookup the parameter name based on the key and then to
        'set the value accordingly.

        Dim oParameter As OracleClient.OracleParameter
        Dim oParameters As ParameterCollection
        Dim iParameterCount As Integer = 0
        Dim iCount As Integer
        'Dim iValue As Integer
        Dim sValue As String
        Dim dtmDate As Date
        'Dim dValue As Decimal
        Dim iPos As Integer
        Dim sParams As String()
        Dim oDBNull As Object = System.DBNull.Value
        Dim sParameterName As String = String.Empty
        Dim sParameterValue As String = String.Empty
        Dim sOracleDataType As String = String.Empty

        Try
            If Not alParameterValues Is Nothing Then
                For iCount = 0 To oCommand.Parameters.Count - 1
                    oParameter = oCommand.Parameters(iCount)

                    Select Case oParameter.Direction
                        Case ParameterDirection.Input, ParameterDirection.InputOutput
                            sParameterName = oParameter.ParameterName
                            sParameterValue = String.Empty

                            If TypeOf alParameterValues(iParameterCount) Is System.DBNull Then
                                sParameterValue = "NULL"
                            Else
                                sParameterValue = CType(alParameterValues(iParameterCount), String)
                            End If
                            sOracleDataType = oParameter.OracleType.ToString

                            If ComponentServices.m_structConfigSettings.eLoggingLevel = ComponentServices.LoggingLevel.Full Then
                                sStoredProcedureName &= vbCrLf & sParameterName & ":" & sParameterValue & ":" & sOracleDataType
                            End If

                            Select Case oParameter.OracleType

                                Case OracleClient.OracleType.Int16, OracleClient.OracleType.Int32, OracleClient.OracleType.UInt16, OracleClient.OracleType.UInt32
                                    If alParameterValues(iParameterCount) Is Nothing Or TypeOf alParameterValues(iParameterCount) Is System.DBNull Then
                                        oParameter.Value = oDBNull
                                    Else
                                        oParameter.Value = CType(alParameterValues(iParameterCount), Integer)
                                    End If

                                Case OracleClient.OracleType.Number
                                    If alParameterValues(iParameterCount) Is Nothing Or TypeOf alParameterValues(iParameterCount) Is System.DBNull Then
                                        oParameter.Value = oDBNull
                                    Else
                                        oParameter.Value = CType(alParameterValues(iParameterCount), Decimal)
                                    End If

                                Case OracleClient.OracleType.Char, OracleClient.OracleType.NChar, OracleClient.OracleType.VarChar, OracleClient.OracleType.NVarChar
                                    If alParameterValues(iParameterCount) Is Nothing Or TypeOf alParameterValues(iParameterCount) Is System.DBNull Then
                                        oParameter.Value = oDBNull
                                    Else
                                        oParameter.Value = CType(alParameterValues(iParameterCount), String)
                                    End If

                                Case OracleClient.OracleType.DateTime
                                    If alParameterValues(iParameterCount) Is Nothing Or TypeOf alParameterValues(iParameterCount) Is System.DBNull Then
                                        oParameter.Value = oDBNull
                                    Else
                                        sValue = alParameterValues(iParameterCount)
                                        If sValue <> String.Empty Then
                                            'handle the occrrance TO_DATE
                                            If sValue.IndexOf("TO_DATE") >= 0 Then
                                                sValue = sValue.Substring(sValue.IndexOf("(") + 1)
                                                sValue = sValue.Replace("'", "")
                                                sValue = sValue.Replace(")", "")
                                                sParams = sValue.Split(","c)

                                                If sParams(0) <> "00:00:00" Then
                                                    oParameter.Value = ComponentServices.GetISODataValue(sParams(0), sParams(1))
                                                Else
                                                    oParameter.Value = oDBNull
                                                End If
                                            Else
                                                If sValue <> "00:00:00" Then
                                                    oParameter.Value = ComponentServices.GetISODataValue(sValue)
                                                Else
                                                    oParameter.Value = oDBNull
                                                End If
                                            End If
                                        Else
                                            oParameter.Value = oDBNull
                                        End If
                                    End If

                            End Select

                            iParameterCount += 1

                            If iParameterCount > alParameterValues.Count - 1 Then
                                Exit For    'No more parameters provided in sql query so additional stored procedure parameters must have default values
                            End If

                        Case ParameterDirection.Output

                            oParameter.Direction = ParameterDirection.Output
                    End Select
                Next
            End If

        Catch oError As OracleClient.OracleException
            ComponentServices.LogError( _
                    True, _
                    oError.Code, _
                    oError.Message, _
                    "abComponentServices", _
                    oError.StackTrace, _
                    ComponentServices.ErrorSeverity.ES_Error)
        Catch oError As System.Exception
            ComponentServices.LogError(True, "abComponentServices", oError, "Error processing value for Parameter (" & sParameterName & ":" & sParameterValue & ":" & sOracleDataType & ")")
        Finally
            If ComponentServices.m_structConfigSettings.eLoggingLevel = ComponentServices.LoggingLevel.Full Then
                ComponentServices.LogInformation("SID: " & m_iSID & " Executing Stored Procedure : " & sStoredProcedureName, "abComponentServices", "SetStoredProcedureParameterValues")
            End If
        End Try
    End Sub

    Public Function ExecuteStoredProcedureParameters(ByVal sName As String, ByVal alParameterValues As System.Collections.ArrayList) As System.Data.DataTable Implements IABDatabaseAdapter.ExecuteStoredProcedureParameters
        '======================================================================
        ' Date Created : 27 February 2008 
        '----------------------------------------------------------------------
        ' About...     : Execute stored procedure using array list of parameters
        '======================================================================
        ' Change Log
        '
        ' Date       Author         Ref                         Comments
        ' ---------- ------------   ---------                   --------
        ' 27/02/2008 Jat Virdee     TD24359, TD24357, TD24426   Created
        '======================================================================
        Dim oDataAdapter As OracleClient.OracleDataAdapter
        Dim oDataTable As Data.DataTable
        Dim oCommand As OracleClient.OracleCommand
        Dim sQueryDescription As New StringBuilder

        Try
            If m_iSID <> 0 Then
                sQueryDescription.Append("SID: ")
                sQueryDescription.Append(m_iSID)
                sQueryDescription.Append(" : ")
                sQueryDescription.Append(sName)
            Else
                sQueryDescription.Append(sName)
            End If

            'Get the stored procedure parameters (the cache will be used if possible)
            oCommand = GetStoredProcedureParameters(sName.Trim)

            'Set the parameter values based on the parsed values from the supplied SQL statement
            SetStoredProcedureParameterValues(oCommand, alParameterValues, sName)

            'Set the data adapter properties
            oCommand.Connection = m_oConnection

            oDataAdapter = New OracleClient.OracleDataAdapter(oCommand)
            oDataAdapter.SelectCommand.Transaction = m_oTransaction

            oDataTable = New Data.DataTable
            oDataAdapter.Fill(oDataTable)

            Call ComponentServices.LogInformation( _
                    sQueryDescription.ToString, _
                    "abComponentServices", _
                    "ExecuteStoredProcedureParameters")

        Catch oError As OracleClient.OracleException
            If oError.Code <> 972 Then
                ComponentServices.LogError( _
                        True, _
                        oError.Code, _
                        oError.Message, _
                        "abComponentServices", _
                        oError.StackTrace, _
                        ComponentServices.ErrorSeverity.ES_Error, _
                        sQueryDescription.ToString)
            End If

            oDataTable = Nothing
        Catch oError As System.Exception
            ComponentServices.LogError(True, "abComponentServices", oError, sName)

            oDataTable = Nothing
        Finally
            If Not oDataAdapter Is Nothing Then
                oDataAdapter.Dispose()
                oDataAdapter = Nothing
            End If

            If Not oCommand Is Nothing Then
                oCommand.Transaction = Nothing
                oCommand.Connection = Nothing
                oCommand.Dispose()
                oCommand = Nothing
            End If
        End Try

        Return oDataTable
    End Function

    Public ReadOnly Property DBProcessId() As Integer Implements IABDatabaseAdapter.DBProcessId
        Get
            Return m_iSID
        End Get
    End Property

End Class

Friend Class ADSAdapter

    Implements IABDatabaseAdapter

    Private m_oConnection As OleDb.OleDbConnection
    Private m_sConnectionString As String

    Public ReadOnly Property DBProcessId() As Integer Implements IABDatabaseAdapter.DBProcessId
        Get
            Return 0 ' Meaningless here
        End Get
    End Property

    Public Sub OpenConnection( _
            ByVal sProvider As String, _
            ByVal sDatabaseName As String, _
            ByVal sUserName As String, _
            ByVal sPassword As String, _
            ByVal iTimeout As Integer) Implements IABDatabaseAdapter.OpenConnection


        Dim sConnectionString As String
        Dim sbConnectionString As StringBuilder

        sProvider = "ADsDSOObject"
        sbConnectionString = New StringBuilder
        sbConnectionString.Append("Provider=").Append(sProvider)
        sConnectionString = sbConnectionString.ToString()

        m_sConnectionString = sConnectionString

        Try
            m_oConnection = New OleDb.OleDbConnection(sConnectionString)
            m_oConnection.Open()
        Catch oError As OleDb.OleDbException
            ComponentServices.LogError( _
                    True, _
                    oError.ErrorCode, _
                    oError.Message, _
                    "abComponentServices", _
                    oError.StackTrace, _
                    ComponentServices.ErrorSeverity.ES_Error, _
                    sConnectionString)
        End Try


    End Sub

    Public Sub CloseConnection() Implements IABDatabaseAdapter.CloseConnection

        Try
            If m_oConnection.State <> ConnectionState.Closed Then
                m_oConnection.Close()
            End If

            If ComponentServices.m_structConfigSettings.eLoggingLevel = ComponentServices.LoggingLevel.Full Then
                ComponentServices.LogInformation("ADS adapter connection closed ", "abComponentServices", "ConnectionClosed")
            End If

        Catch ex As System.Exception
            ComponentServices.LogError(True, "abComponentServices", ex, "CloseConnection")

        Finally
            m_oConnection.Dispose()
            m_oConnection = Nothing
        End Try

    End Sub
    Private Function ExecuteQuery( _
          ByVal sQuery As String, _
          ByRef iRowsAffected As Integer) As OleDb.OleDbDataReader

        Dim oCommand As OleDb.OleDbCommand
        Dim oDataReader As OleDb.OleDbDataReader

        Try
            oCommand = New OleDb.OleDbCommand( _
                    sQuery, _
                    m_oConnection)
            oDataReader = oCommand.ExecuteReader()
            iRowsAffected = oDataReader.RecordsAffected

            Call ComponentServices.LogInformation( _
                    sQuery, _
                    "abComponentServices", _
                    "ExecuteQuery")

        Catch oError As OleDb.OleDbException   'OracleClient.OracleException
            If oError.ErrorCode <> 972 Then
                ComponentServices.LogError( _
                        True, _
                        oError.ErrorCode, _
                        oError.Message, _
                        "abComponentServices", _
                        oError.StackTrace, _
                        ComponentServices.ErrorSeverity.ES_Error, _
                        sQuery)
            Else
                oDataReader = Nothing
            End If
        End Try

        Return oDataReader

    End Function
    Private Sub AddDataReaderToParameterCollection( _
            ByRef oDataReader As OleDb.OleDbDataReader, _
            ByRef oResult As ParameterCollection)

        Dim iFields As Integer
        Dim iField As Integer

        iFields = oDataReader.FieldCount - 1
        For iField = 0 To iFields
            If Not oDataReader.IsDBNull(iField) Then
                Select Case oDataReader.GetDataTypeName(iField).ToLower()
                    Case "number"
                        Call oResult.Add(oDataReader.GetName(iField), oDataReader.GetInt32(iField))

                    Case "byte"
                        Call oResult.Add(oDataReader.GetName(iField), oDataReader.GetByte(iField))

                    Case "bit"
                        Call oResult.Add(oDataReader.GetName(iField), oDataReader.GetBoolean(iField))

                    Case "datetime"
                        Call oResult.Add(oDataReader.GetName(iField), oDataReader.GetDateTime(iField))

                    Case "decimal"
                        Call oResult.Add(oDataReader.GetName(iField), oDataReader.GetDecimal(iField))

                    Case "double"
                        Call oResult.Add(oDataReader.GetName(iField), oDataReader.GetDouble(iField))

                    Case "varchar2", "nvarchar", "nchar", "char", "dbtype_variant"
                        Call oResult.Add(oDataReader.GetName(iField), oDataReader.GetString(iField))

                    Case "date"
                        Call oResult.Add(oDataReader.GetName(iField), oDataReader.GetDateTime(iField))

                    Case Else
                        ComponentServices.LogInformation("The following oracle data type string has not been recoginzed by ActiveBank: " + oDataReader.GetDataTypeName(iField).ToLower(), "abComponentServices", "AddDataReaderToParameterCollection")

                End Select
            End If
        Next iField

    End Sub

    Public Sub AbortTransaction() Implements IABDatabaseAdapter.AbortTransaction

    End Sub

    Public Sub CommitTransaction() Implements IABDatabaseAdapter.CommitTransaction

    End Sub

    Public Function ConvertTimestampToString(ByVal byteTimestamp() As Byte) As String Implements IABDatabaseAdapter.ConvertTimestampToString

    End Function



    Public Function ExecuteQueryDataTable(ByVal strQuery As String) As System.Data.DataTable Implements IABDatabaseAdapter.ExecuteQueryDataTable
        Dim oDataAdapter As OleDb.OleDbDataAdapter
        Dim oDataTable As Data.DataTable

        Try
            oDataAdapter = New OleDb.OleDbDataAdapter(strQuery, m_oConnection)
            oDataTable = New Data.DataTable
            oDataAdapter.Fill(oDataTable)
            oDataAdapter.Dispose()
            oDataAdapter = Nothing

            Call ComponentServices.LogInformation( _
                    strQuery, _
                    "abComponentServices", _
                    "ExecuteQueryDataTable")

        Catch oError As OleDb.OleDbException
            If oError.ErrorCode <> 972 Then
                ComponentServices.LogError( _
                        True, _
                        oError.ErrorCode, _
                        oError.Message, _
                        "abComponentServices", _
                        oError.StackTrace, _
                        ComponentServices.ErrorSeverity.ES_Error, _
                        strQuery)
            End If

            oDataTable = Nothing
        End Try

        Return oDataTable
    End Function

    Public Function ExecuteStoredProcedure(ByVal sName As String, ByVal oParameters As ParameterCollection, ByVal oAlternativeParameters As ParameterCollection) As ParameterCollection Implements IABDatabaseAdapter.ExecuteStoredProcedure

    End Function

    Public Function ExecuteStoredProcedureDataTable(ByVal sName As String, ByVal oParameters As ParameterCollection) As System.Data.DataTable Implements IABDatabaseAdapter.ExecuteStoredProcedureDataTable

    End Function

    Public Function GetStoredProcedureDataTypes(ByVal sName As String) As ParameterCollection Implements IABDatabaseAdapter.GetStoredProcedureDataTypes

    End Function

    Public Sub StartTransaction() Implements IABDatabaseAdapter.StartTransaction

    End Sub

    Public Sub ExecuteNonQuery(ByVal sQuery As String, ByRef iRowsAffected As Integer) Implements IABDatabaseAdapter.ExecuteNonQuery

    End Sub

    Public Function ExecuteQueryParameterCollection(ByVal sQuery As String, ByVal bSingleRow As Boolean, ByVal bIDOrdinal As Boolean) As ParameterCollection Implements IABDatabaseAdapter.ExecuteQueryParameterCollection
        Dim pcResult As ParameterCollection
        Dim oDataReader As OleDb.OleDbDataReader
        Dim pcResultRow As ParameterCollection
        Dim iRowsAffected As Integer

        oDataReader = ExecuteQuery(sQuery, iRowsAffected)

        If oDataReader Is Nothing Then
            pcResult = Nothing
        Else
            pcResult = New ParameterCollection

            If bSingleRow Then
                If oDataReader.Read() Then
                    AddDataReaderToParameterCollection( _
                            oDataReader, _
                            pcResult)
                Else
                    pcResult = Nothing
                End If
            Else
                While oDataReader.Read()
                    pcResultRow = New ParameterCollection

                    AddDataReaderToParameterCollection( _
                            oDataReader, _
                            pcResultRow)
                    If Not oDataReader.GetSchemaTable.Columns("ID") Is Nothing Then
                        pcResult.Add(CStr(oDataReader.GetDecimal(oDataReader.GetOrdinal("ID"))), pcResultRow)
                    Else
                        pcResult.Add("", pcResultRow)
                    End If


                End While
            End If

            oDataReader.Close()
        End If

        Return pcResult

    End Function

    Public Function ExecuteStoredProcedureParameters(ByVal sName As String, ByVal alParameterValues As System.Collections.ArrayList) As System.Data.DataTable Implements IABDatabaseAdapter.ExecuteStoredProcedureParameters

    End Function
End Class

#End Region

#Region "ComponentServices Class"

< _
    ComClass(ComponentServices.ClassId, ComponentServices.InterfaceId, ComponentServices.EventsId) _
> _
Public Class ComponentServices

#Region "COM GUIDs"
    Public Const ClassId As String = "218789E6-6BA5-46DE-8E9B-BD3FD7DE397E"
    Public Const InterfaceId As String = "15B87297-0063-435D-B216-0896D8BC451D"
    Public Const EventsId As String = "923918E0-A92C-4C0F-83A5-B4F3C361C248"
#End Region

#Region "Class Level Declarations"

    Public Enum ErrorSeverity
        ES_Error = 1
        ES_Warning = 2
    End Enum

    Public Enum ListScrollDirection
        PreviousPage = 0
        NextPage = 1
    End Enum

    Public Enum ListFilterComparisonType
        Equals = 0
        NotEquals = 1
        GreaterThan = 2
        LessThan = 3
        StartsWith = 4
        EndsWith = 5
        Contains = 6
        GreaterThanOrEqualTo = 7
        LessThanOrEqualTo = 8
        BelongsTo = 9
        ExcludedFrom = 10
        Exists = 11
        NotExists = 12
        BitwiseEquals = 13
        BitwiseNotEquals = 14
    End Enum

    Public Enum ListFilterConjunction
        LFC_And = 0
        LFC_Or = 1
    End Enum

    Public Enum DeleteResponse
        Vetoed = 0
        Allowed = 1
    End Enum

    Public Enum LoggingLevel
        None = 0
        ErrorsOnly = 1
        WarningsAndErrors = 2
        Full = 3
    End Enum

    Friend Enum DatabaseVendor
        SQLServer = 1
        Oracle = 2
        ADS = 3
    End Enum

    Private Enum ConfigFields
        LoggingLevel = 0
        DatabaseTimeout = 1
        CachingEnabled = 2
    End Enum

    Friend Structure ConfigSettings
        Public eLoggingLevel As LoggingLevel
        Public sDatabaseServer As String
        Public sDatabaseName As String
        Public sDatabaseUserName As String
        Public sDatabasePassword As String
        Public iDatabaseTimeout As Integer
        Public bCachingEnabled As Boolean
        Public eDatabaseVendor As DatabaseVendor
        Public sApplicationPath As String
        Public sApplicationServer As String
        Public bDBSettingsEncrypted As Boolean
        Public bConfigSettingsSet As Boolean
        Public sDatabaseIdentifier As String

        'Public bAutoCreateEnumerator As Boolean
    End Structure

    Friend Shared m_structConfigSettings As ConfigSettings

    Private Shared m_oCachedParameterCollections As Collections.Specialized.HybridDictionary
    Private Shared m_oCachedDataTables As Collections.Specialized.HybridDictionary
    Private Shared m_oCachedXMLDocuments As Collections.Specialized.HybridDictionary

    Friend Shared m_oStoredProcedureDefinitions As Collections.Specialized.HybridDictionary

    Public m_iReferenceCount As Integer

    Private m_oDatabaseAdapter As IABDatabaseAdapter
    Private m_bIsInTransaction As Boolean
    Private m_bTransactionAborted As Boolean

    Private m_oCOMComponentServices As comComponentServices.ComponentServices
    Private m_iCOMReferenceCount As Integer

    'AFS Used for interop purposes.
    Private Shared m_oLogClient As Object
    Private m_oUser As Interop.abUser.User

    'Private WithEvents m_oRDOInterop As abRDOInterop.RDOClass
    Private m_oRDOInterop As abRDOInterop.RDOClass

    Private m_oAuditTrail As Object
    Private m_oBroker As Object

    'User Information
    Public m_sSessionID As String = String.Empty
    Public m_sBranchID As String = String.Empty
    Public m_decUserID As Decimal = 1D 'Default to 1 (For Auditing XML requests we will not have a session)

    'RequestID for auditing the Requests or Responses
    Private m_sRequestId As String = String.Empty

    'A dictionary that stores any database object aliases (short names)
    Private Shared m_htAliasCache As Hashtable

    Private Const DEFAULT_DB_IDENTIFIER As String = "DatabaseInfo"


#End Region

#Region "Session Properties"

    Private ReadOnly Property User() As Interop.abUser.User
        Get            
            Dim oSession As Object

            GetRDOInterop()

            If m_oUser Is Nothing Then
                m_oUser = CreateObject("DiamUser.User")
                CallByName(m_oUser, "DBHandle", CallType.Let, m_oRDOInterop)

                oSession = CreateObject("DiamSession.clsSession")
                CallByName(oSession, "DBHandle", CallType.Let, m_oRDOInterop)
                CallByName(oSession, "User", CallType.Let, m_oUser)
                CallByName(m_oUser, "SessionHandle", CallType.Let, oSession)
                CallByName(m_oRDOInterop, "User", CallType.Let, m_oUser)
            End If

            Return m_oUser

        End Get
    End Property
    Public Property SessionState_CustomerNumber() As String
        Get
            Return User.SessionHandle.CustomerNumber
        End Get
        Set(ByVal Value As String)
            User.SessionHandle.CustomerNumber = Value
        End Set
    End Property
    Public Property SessionState_CustomerName() As String
        Get
            Return User.SessionHandle.CustomerName
        End Get
        Set(ByVal Value As String)
            User.SessionHandle.CustomerName = Value
        End Set
    End Property
    Public Property SessionState_CustomerObjectKey() As String
        Get
            Return User.SessionHandle.CustomerObjectKey
        End Get
        Set(ByVal Value As String)
            User.SessionHandle.CustomerObjectKey = Value
        End Set
    End Property
    Public Property SessionState_SessionId() As String
        Get
            Return User.SessionHandle.sSessionID
        End Get
        Set(ByVal Value As String)
            User.SessionHandle.sSessionID = Value
            m_sSessionID = Value
        End Set
    End Property
    Public Property SessionState_Channel() As String
        Get
            Return User.SessionHandle.sChannel
        End Get
        Set(ByVal Value As String)
            User.SessionHandle.sChannel = Value
        End Set
    End Property

    Public Function bTerminateSession(ByVal sSessionID As String) As Boolean
        Dim bReturn As Boolean = True
        Try
            bReturn = User.SessionHandle.bEndSession(sSessionID)
        Catch ex As Exception
            bReturn = False
        End Try
        Return bReturn
    End Function
#End Region
#Region "Static Utility Methods"

    Public Shared Sub LogError( _
        ByVal bRaiseError As Boolean, _
           ByVal oError As ActiveBankException)

        LogError( _
                bRaiseError, _
                oError.Number, _
                oError.Message, _
                oError.ComponentName, _
                oError.StackTrace, _
                ErrorSeverity.ES_Error, _
                String.Empty, _
                oError.Source)

    End Sub

    Public Shared Sub LogError( _
            ByVal bRaiseError As Boolean, _
            ByVal sComponentName As String, _
            ByVal oError As System.Exception)

        LogError( _
                bRaiseError, _
                0, _
                oError.Message, _
                sComponentName, _
                oError.StackTrace, _
                ErrorSeverity.ES_Error, _
                String.Empty, _
                oError.Source)

    End Sub

    Public Shared Sub LogError( _
        ByVal bRaiseError As Boolean, _
        ByVal sDescription As String, _
        ByVal sComponentName As String, _
        ByVal sStackTrace As String)

        LogError( _
                bRaiseError, _
                0, _
                sDescription, _
                sComponentName, _
                sStackTrace, _
                ErrorSeverity.ES_Error, _
                String.Empty, _
                String.Empty)
    End Sub

    Public Shared Sub LogError( _
            ByVal bRaiseError As Boolean, _
            ByVal sDescription As String, _
            ByVal sComponentName As String, _
            ByVal sStackTrace As String, _
            ByVal eSeverity As ErrorSeverity)

        LogError( _
                bRaiseError, _
                0, _
                sDescription, _
                sComponentName, _
                sStackTrace, _
                eSeverity, _
                String.Empty, _
                String.Empty)
    End Sub

    Public Shared Sub LogError( _
            ByVal bRaiseError As Boolean, _
            ByVal iNumber As Integer, _
            ByVal sDescription As String, _
            ByVal sComponentName As String, _
            ByVal sStackTrace As String)

        LogError( _
                bRaiseError, _
                iNumber, _
                sDescription, _
                sComponentName, _
                sStackTrace, _
                ErrorSeverity.ES_Error, _
                String.Empty, _
                sComponentName)
    End Sub

    Public Shared Sub LogError( _
            ByVal bRaiseError As Boolean, _
            ByVal iNumber As Integer, _
            ByVal sDescription As String, _
            ByVal sComponentName As String, _
            ByVal sStackTrace As String, _
            ByVal eSeverity As ErrorSeverity)

        LogError( _
                bRaiseError, _
                iNumber, _
                sDescription, _
                sComponentName, _
                sStackTrace, _
                eSeverity, _
                String.Empty, _
                sComponentName)

    End Sub

    Public Shared Sub LogError( _
            ByVal bRaiseError As Boolean, _
            ByVal sComponentName As String, _
            ByVal oError As System.Exception, _
            ByVal sComment As String)

        LogError( _
                bRaiseError, _
                0, _
                oError.Message, _
                sComponentName, _
                oError.StackTrace, _
                ErrorSeverity.ES_Error, _
                sComment, _
                oError.Source)

    End Sub

    Public Shared Sub LogError( _
            ByVal bRaiseError As Boolean, _
            ByVal iNumber As Integer, _
            ByVal sDescription As String, _
            ByVal sComponentName As String, _
            ByVal sStackTrace As String, _
            ByVal eSeverity As ErrorSeverity, _
            ByVal sComment As String)

        LogError( _
                bRaiseError, _
                iNumber, _
                sDescription, _
                sComponentName, _
                sStackTrace, _
                ErrorSeverity.ES_Error, _
                sComment, _
                sComponentName)

    End Sub

    Public Shared Sub LogError( _
            ByVal bRaiseError As Boolean, _
            ByVal iNumber As Integer, _
            ByVal sDescription As String, _
            ByVal sComponentName As String, _
            ByVal sStackTrace As String, _
            ByVal eSeverity As ErrorSeverity, _
            ByVal sComment As String, _
            ByVal sSource As String)

        Dim oError As System.Exception
        Dim eLevel As LoggingLevel

        Try
            If m_structConfigSettings.sDatabaseName Is Nothing Then
                eLevel = LoggingLevel.Full
            Else
                eLevel = m_structConfigSettings.eLoggingLevel
            End If
        Catch
            eLevel = LoggingLevel.Full
        End Try

        Select Case eSeverity
            Case ErrorSeverity.ES_Error
                If eLevel >= LoggingLevel.ErrorsOnly Then
                    WriteToLog( _
                            sComponentName, _
                            "", _
                            EventLogEntryType.Error, _
                            sStackTrace, _
                            iNumber, _
                            sDescription, _
                            sComment, _
                            sSource)
                End If

            Case ErrorSeverity.ES_Warning
                If eLevel >= LoggingLevel.WarningsAndErrors Then
                    WriteToLog( _
                            sComponentName, _
                            "", _
                            EventLogEntryType.Warning, _
                            sStackTrace, _
                            iNumber, _
                            sDescription, _
                            sComment, _
                            sSource)
                End If

        End Select

        If bRaiseError Then
            Throw New ActiveBankException(iNumber, sDescription, sComponentName)
        End If

    End Sub

    Public Shared Sub LogInformation( _
            ByVal sDescription As String, _
            ByVal sComponentName As String, _
            ByVal sMethodName As String)

        LogInformation( _
                0, _
                sDescription, _
                sComponentName, _
                sMethodName)

    End Sub

    Public Shared Sub LogInformation( _
            ByVal iNumber As Integer, _
            ByVal sDescription As String, _
            ByVal sComponentName As String, _
            ByVal sMethodName As String)
        '================================================================================================
        ' About...  : Logs the information messages
        '================================================================================================
        ' Change Log
        'Date           Author       Ref            Comments
        '----           ------      ---             --------
        ' 17/11/2008    Chudamanie  TD25504         Set Logging Level to None
        '=================================================================================================
        Dim eLevel As LoggingLevel

        If m_structConfigSettings.sDatabaseName Is Nothing Then
            eLevel = LoggingLevel.None
        Else
            eLevel = m_structConfigSettings.eLoggingLevel
        End If

        If eLevel = LoggingLevel.Full Then
            WriteToLog( _
                    sComponentName, _
                    sMethodName, _
                    EventLogEntryType.Information, _
                    "", _
                    iNumber, _
                    sDescription, _
                    String.Empty, _
                    sComponentName)
        End If

    End Sub

    'AFS Need to make use of database or WMI for logging in future.
    Private Shared Sub WriteToLog( _
            ByVal sComponentName As String, _
            ByVal sMethodName As String, _
            ByVal eLevel As System.Diagnostics.EventLogEntryType, _
            ByVal sStackTrace As String, _
            ByVal iNumber As Integer, _
            ByVal sDescription As String, _
            ByVal sComment As String)

        WriteToLog( _
                sComponentName, _
                sMethodName, _
                EventLogEntryType.Information, _
                sStackTrace, _
                iNumber, _
                sDescription, _
                sComment, _
                "")

    End Sub

    Private Shared Sub WriteToLog( _
            ByVal sComponentName As String, _
            ByVal sMethodName As String, _
            ByVal eLevel As System.Diagnostics.EventLogEntryType, _
            ByVal sStackTrace As String, _
            ByVal iNumber As Integer, _
            ByVal sDescription As String, _
            ByVal sComment As String, _
            ByVal sSource As String)

        Dim sbMessage As StringBuilder
        Dim bWriteToLog As Boolean = False
        Dim eDatabaseLoggingLevel As LoggingLevel
        Dim sErrorMessageType As String = String.Empty

        Try
            If m_structConfigSettings.sDatabaseName Is Nothing Then
                eDatabaseLoggingLevel = LoggingLevel.None
            Else
                eDatabaseLoggingLevel = m_structConfigSettings.eLoggingLevel
            End If
        Catch
            eDatabaseLoggingLevel = LoggingLevel.Full
        End Try

        Select Case eLevel
            Case EventLogEntryType.Error, EventLogEntryType.FailureAudit
                If eDatabaseLoggingLevel >= LoggingLevel.ErrorsOnly Then
                    sErrorMessageType = "E"
                    bWriteToLog = True
                End If
            Case EventLogEntryType.Information
                If eDatabaseLoggingLevel >= LoggingLevel.Full Then
                    sErrorMessageType = "I"
                    bWriteToLog = True
                End If
            Case EventLogEntryType.Warning
                If eDatabaseLoggingLevel >= LoggingLevel.WarningsAndErrors Then
                    sErrorMessageType = "W"
                    bWriteToLog = True
                End If
            Case EventLogEntryType.SuccessAudit
        End Select

        If bWriteToLog Then
            sbMessage = New StringBuilder("Description: ")
            sbMessage.Append(sDescription)

            If sStackTrace.Length <> 0 Then
                sbMessage.Append(vbNewLine)
                sbMessage.Append("Call stack: ")
                sbMessage.Append(sStackTrace)
            End If

            If sComment.Length <> 0 Then
                sbMessage.Append(vbNewLine)
                sbMessage.Append("Comment: ")
                sbMessage.Append(sComment)
            End If

            WriteToNTEventLog(sComponentName, sMethodName, sbMessage.ToString(), eLevel, iNumber)

            Try
                If m_oLogClient Is Nothing Then
                    m_oLogClient = CreateObject("GFSLogServer.clsLogClient")
                End If

                Call m_oLogClient.LogMessage(sSource, IIf(sMethodName = String.Empty, sSource, sMethodName), m_structConfigSettings.sApplicationServer, CInt(eLevel), sErrorMessageType & iNumber.ToString, sbMessage.ToString())
            Catch ex As System.Exception
                EventLog.WriteEntry("activebank - Logging", ex.Message, eLevel)
            End Try
        End If

    End Sub

    Private Shared Sub WriteToNTEventLog( _
            ByVal sComponentName As String, _
            ByVal sMethodName As String, _
            ByVal sMessage As String, _
            ByVal eLevel As System.Diagnostics.EventLogEntryType, _
            ByVal iNumber As Integer)

        Dim sbMessage As StringBuilder

        sbMessage = New StringBuilder(sMessage)
        sbMessage.Append(vbNewLine)
        sbMessage.Append("Method: ")
        sbMessage.Append(sMethodName)

        sbMessage.Append(vbNewLine)
        sbMessage.Append("Number: ")
        sbMessage.Append(iNumber)

        sMessage = sbMessage.ToString()

        EventLog.WriteEntry("activebank - " + sComponentName, sMessage, eLevel)

    End Sub

    Public Shared Function CreateXMLDocument( _
            ByVal strRootName As String, _
            ByVal ParamArray strElements() As String) As Xml.XmlDocument

        Dim docDocument As Xml.XmlDocument
        Dim elmRoot As Xml.XmlElement
        Dim iEqualsPosition As Integer
        Dim strElementName As String
        Dim strElementValue As String
        Dim iElements As Integer
        Dim iElement As Integer
        Dim strElementDefinition As String
        Dim strElementPath As String
        Dim strElementPaths() As String
        Dim iElementPaths As Integer
        Dim iElementPath As Integer
        Dim elmElement As Xml.XmlElement
        Dim iAttribPos As Integer
        Dim sAttribName As String
        Dim elmSubNode As Xml.XmlElement

        Try
            docDocument = CreateXMLDocument()

            elmRoot = docDocument.CreateElement(strRootName)
            docDocument.AppendChild(elmRoot)

            If Not (strElements Is Nothing) Then
                iElements = strElements.Length - 1
                For iElement = 0 To iElements
                    strElementDefinition = strElements(iElement)

                    iEqualsPosition = strElementDefinition.IndexOf("=")
                    If iEqualsPosition = -1 Then
                        strElementName = strElementDefinition
                        strElementValue = ""
                    Else
                        strElementName = strElementDefinition.Substring(0, iEqualsPosition)
                        strElementValue = strElementDefinition.Substring(iEqualsPosition + 1)
                    End If

                    'Check if an attribute is required
                    iAttribPos = strElementName.LastIndexOf("@")
                    If iAttribPos = -1 Then
                        strElementPath = strElementName
                        strElementPaths = strElementName.Split(CChar("/"))
                        iElementPaths = strElementPaths.Length

                        strElementName = strElementPaths(iElementPaths - 1)
                        elmElement = docDocument.CreateElement(strElementName)

                        If strElementValue.Length > 0 Then
                            elmElement.InnerText = strElementValue
                        End If

                        If iElementPaths = 1 Then
                            elmRoot.AppendChild(elmElement)
                        Else
                            elmSubNode = elmRoot.SelectSingleNode(strElementPath.Substring(0, strElementPath.Length - strElementName.Length - 1))
                            If elmSubNode Is Nothing Then
                                Throw New ActiveBankException(ErrorNumbers.TheXMLDocumentISInvalid, "The XML document being created is invalid.", ActiveBankException.ExceptionType.System, "abComponentServices")
                            End If
                            elmSubNode.AppendChild(elmElement)
                        End If
                    Else
                        sAttribName = strElementName.Substring(iAttribPos + 1)
                        strElementName = strElementName.Substring(0, iAttribPos)
                        If strElementName.Length = 0 Then
                            elmElement = elmRoot
                        Else
                            elmElement = elmRoot.SelectSingleNode(strElementName)
                        End If

                        If elmElement Is Nothing Then
                            Throw New ActiveBankException(ErrorNumbers.TheXMLDocumentISInvalid, "The XML document being created is invalid.", ActiveBankException.ExceptionType.System, "abComponentServices")
                        End If

                        Call elmElement.SetAttribute(sAttribName, strElementValue)
                    End If
                Next iElement
            End If

        Catch oError As Exception
            LogError(True, oError.Message, "abComponentServices", oError.StackTrace, ErrorSeverity.ES_Error)

            Return Nothing

        End Try

        Return docDocument

    End Function

    Public Shared Function CreateXMLDocument( _
            ByVal sXML As String) As Xml.XmlDocument

        Dim docDocument As Xml.XmlDocument

        docDocument = CreateXMLDocument()
        docDocument.LoadXml(sXML)

        Return docDocument

    End Function

    Public Shared Function CreateXMLDocument() As Xml.XmlDocument

        Return New Xml.XmlDocument

    End Function

    Public Shared Function GetISODataValue( _
            ByVal bVBDataValue As Boolean) As String

        Dim sISODataValue As String

        If bVBDataValue Then
            sISODataValue = "1"
        Else
            sISODataValue = "0"
        End If

        Return sISODataValue

    End Function

    Public Shared Function GetISODataValue( _
            ByVal iVBDataValue As Integer) As String

        Dim sISODataValue As String

        sISODataValue = iVBDataValue.ToString(System.Globalization.CultureInfo.InvariantCulture)

        Return sISODataValue

    End Function

    Public Shared Function GetISODataValue( _
            ByVal dblVBDataValue As Double) As String
        '==================================================================================================
        ' Author    : Jim Hollingsworth - 21/04/2005
        '--------------------------------------------------------------------------------------------------
        ' About...  : Returns an ISO format number. i.e.
        '               1234.56
        '             regardless of the current culture / regional options.
        '==================================================================================================
        Dim strISODataValue As String

        strISODataValue = Str(dblVBDataValue).Trim()

        Return strISODataValue

    End Function

    Public Shared Function GetISODataValue( _
            ByVal dteVBDataValue As Date) As String
        '==================================================================================================
        ' Author    : Jim Hollingsworth - 21/04/2005
        '--------------------------------------------------------------------------------------------------
        ' About...  : Returns a date as a string in ISO format. i.e.
        '               CCYYMMDD
        '             regardless of the current culture / regional options.
        '==================================================================================================
        Dim sISODataValue As String

        If dteVBDataValue <> Nothing Then
            If dteVBDataValue.Date = dteVBDataValue Then
                sISODataValue = dteVBDataValue.ToString("yyyyMMdd")
            Else
                sISODataValue = dteVBDataValue.ToString("yyyyMMddTHH:mm:ss")
            End If
        Else
            sISODataValue = ""
        End If

        Return sISODataValue

    End Function

    Public Shared Function GetISODataValue( _
    ByVal sISODateValue As String) As Date

        Return GetISODataValue(sISODateValue, "yyyyMMdd")

    End Function

    Public Shared Function GetISODataValue( _
        ByVal sISODateValue As String, ByVal sDateFormat As String) As Date
        '==================================================================================================
        ' Author    : Jim Hollingsworth - 21/04/2005
        '--------------------------------------------------------------------------------------------------
        ' About...  : Returns a date variable from a string
        '             sDateFormat MUST contain:
        '               yyyy or yy
        '               mm
        '               dd
        '             and optionally, date separators
        '             Note: this function does not allow for Time in sISODateValue
        '==================================================================================================
        Dim iYear As Integer
        Dim iMonth As Integer
        Dim iDay As Integer

        If sISODateValue.Length = 0 Then
            Throw New ActiveBankException(ErrorNumbers.InvalidISODate, _
                    "Invalid ISO Date string: " & sISODateValue, _
                    ActiveBankException.ExceptionType.Business, _
                    "abComponentServices")
        Else
            Try
                ' Make sure format is ok
                sDateFormat = sDateFormat.ToLower().Replace("cc", "yy")

                If sISODateValue.Length = 6 Then
                    sDateFormat = sDateFormat.ToLower().Replace("yyyy", "yy")
                End If

                If sDateFormat.IndexOf("yyyy") > -1 Then
                    iYear = CInt(sISODateValue.Substring(sDateFormat.IndexOf("yyyy"), 4))
                Else
                    ' Must be a 2 digit year.
                    iYear = CInt(sISODateValue.Substring(sDateFormat.IndexOf("yy"), 2))
                    iYear += 2000   ' This should be changed to use the Windows date cutoff info & current century
                End If
                iMonth = CInt(sISODateValue.Substring(sDateFormat.IndexOf("mm"), 2))
                iDay = CInt(sISODateValue.Substring(sDateFormat.IndexOf("dd"), 2))

                Return DateSerial(iYear, iMonth, iDay)

            Catch ex As Exception
                Throw New ActiveBankException(ErrorNumbers.InvalidISODate, _
                        "Invalid ISO Date string: " & sISODateValue, _
                        ActiveBankException.ExceptionType.Business, _
                        "abComponentServices")
            End Try
        End If

    End Function

    Public Shared Function GetIDFromToken(ByVal sToken As String) As String

        Dim sTokenParts() As String

        If sToken.Length > 0 Then
            sTokenParts = sToken.Split("_")
            Return sTokenParts(0)
        End If

    End Function

    Public Shared Function GetByteXMLElementData( _
        ByVal sName As String, _
        ByRef ndParent As Xml.XmlNode) As Byte

        Dim sData As String
        Dim bytData As Byte
        Dim ndDataNode As Xml.XmlNode

        If sName Is Nothing Then
            ndDataNode = ndParent
        Else
            If sName.Length = 0 Then
                ndDataNode = ndParent
            Else
                ndDataNode = ndParent.SelectSingleNode(sName)
            End If
        End If

        If ndDataNode Is Nothing Then
            If sName Is Nothing Then
                Throw New ActiveBankException(ErrorNumbers.TheElementIsMandatory, "The element is mandatory.", ActiveBankException.ExceptionType.Business, "abComponentServices")
            Else
                Throw (New ActiveBankException(ErrorNumbers.TheElementIsMandatory, "The " + sName + " element is mandatory.", ActiveBankException.ExceptionType.Business, "abComponentServices"))
            End If
        Else
            sData = ndDataNode.InnerText
            If sData.Length = 0 Then
                Throw New ActiveBankException(ErrorNumbers.TheElementIsMandatory, "The " + ndDataNode.Name + " element is mandatory.", ActiveBankException.ExceptionType.Business, "abComponentServices")
            Else
                bytData = CByte(sData)
            End If
        End If

        Return bytData

    End Function

    Public Shared Function GetByteXMLElementData( _
            ByVal sName As String, _
            ByRef ndParent As Xml.XmlNode, _
            ByVal bytDefault As Byte) As Byte

        Dim bytData As Byte
        Dim sData As String
        Dim ndDataNode As Xml.XmlNode

        If sName Is Nothing Then
            ndDataNode = ndParent
        Else
            If sName.Length = 0 Then
                ndDataNode = ndParent
            Else
                ndDataNode = ndParent.SelectSingleNode(sName)
            End If
        End If

        If ndDataNode Is Nothing Then
            bytData = bytDefault
        Else
            sData = ndDataNode.InnerText
            If sData.Length = 0 Then
                bytData = bytDefault
            Else
                bytData = CByte(sData)
            End If
        End If

        Return bytData

    End Function

    Public Shared Function GetStringXMLElementData( _
            ByVal sName As String, _
            ByRef ndParent As Xml.XmlNode) As String

        Dim sData As String
        Dim ndDataNode As Xml.XmlNode

        If sName Is Nothing Then
            ndDataNode = ndParent
        Else
            If sName.Length = 0 Then
                ndDataNode = ndParent
            Else
                ndDataNode = ndParent.SelectSingleNode(sName)
            End If
        End If

        If ndDataNode Is Nothing Then
            If sName Is Nothing Then
                Throw New ActiveBankException(ErrorNumbers.TheElementIsMandatory, "The element is mandatory.", ActiveBankException.ExceptionType.Business, "abComponentServices")
            Else
                Throw (New ActiveBankException(ErrorNumbers.TheElementIsMandatory, "The " + sName + " element is mandatory.", ActiveBankException.ExceptionType.Business, "abComponentServices"))
            End If
        Else
            sData = ndDataNode.InnerText
            If sData.Length = 0 Then
                Throw (New ActiveBankException(ErrorNumbers.TheElementIsMandatory, "The " + ndDataNode.Name + " element is mandatory.", ActiveBankException.ExceptionType.Business, "abComponentServices"))
            End If
        End If

        Return sData

    End Function

    Public Shared Function GetStringXMLElementData( _
            ByVal sName As String, _
            ByRef ndParent As Xml.XmlNode, _
            ByVal sDefault As String) As String

        Dim sData As String
        Dim ndDataNode As Xml.XmlNode

        If sName Is Nothing Then
            ndDataNode = ndParent
        Else
            If sName.Length = 0 Then
                ndDataNode = ndParent
            Else
                ndDataNode = ndParent.SelectSingleNode(sName)
            End If
        End If

        If ndDataNode Is Nothing Then
            sData = sDefault
        Else
            sData = ndDataNode.InnerText
            If sData.Length = 0 Then
                sData = sDefault
            End If
        End If

        Return sData

    End Function

    Public Shared Function GetDecimalXMLElementData( _
            ByVal sName As String, _
            ByRef ndParent As Xml.XmlNode) As Decimal

        Dim sData As String
        Dim decData As Decimal
        Dim ndDataNode As Xml.XmlNode

        If sName Is Nothing Then
            ndDataNode = ndParent
        Else
            If sName.Length = 0 Then
                ndDataNode = ndParent
            Else
                ndDataNode = ndParent.SelectSingleNode(sName)
            End If
        End If

        If ndDataNode Is Nothing Then
            If sName Is Nothing Then
                Throw New ActiveBankException(ErrorNumbers.TheElementIsMandatory, "The element is mandatory.", ActiveBankException.ExceptionType.Business, "abComponentServices")
            Else
                Throw (New ActiveBankException(ErrorNumbers.TheElementIsMandatory, "The " + sName + " element is mandatory.", ActiveBankException.ExceptionType.Business, "abComponentServices"))
            End If
        Else
            sData = ndDataNode.InnerText
            If sData.Length = 0 Then
                Throw New ActiveBankException(ErrorNumbers.TheElementIsMandatory, "The " + ndDataNode.Name + " element is mandatory.", ActiveBankException.ExceptionType.Business, "abComponentServices")
            Else
                decData = CDec(sData)
            End If
        End If

        Return decData

    End Function

    Public Shared Function GetDecimalXMLElementData( _
            ByVal sName As String, _
            ByRef ndParent As Xml.XmlNode, _
            ByVal decDefault As Decimal) As Decimal

        Dim decData As Decimal
        Dim sData As String
        Dim ndDataNode As Xml.XmlNode

        If sName Is Nothing Then
            ndDataNode = ndParent
        Else
            If sName.Length = 0 Then
                ndDataNode = ndParent
            Else
                ndDataNode = ndParent.SelectSingleNode(sName)
            End If
        End If

        If ndDataNode Is Nothing Then
            decData = decDefault
        Else
            sData = ndDataNode.InnerText
            If sData.Length = 0 Then
                decData = decDefault
            Else
                decData = CDec(sData)
            End If
        End If

        Return decData

    End Function

    Public Shared Function GetDoubleXMLElementData( _
        ByVal sName As String, _
        ByRef ndParent As Xml.XmlNode) As Double

        Dim sData As String
        Dim dblData As Double
        Dim ndDataNode As Xml.XmlNode

        If sName Is Nothing Then
            ndDataNode = ndParent
        Else
            If sName.Length = 0 Then
                ndDataNode = ndParent
            Else
                ndDataNode = ndParent.SelectSingleNode(sName)
            End If
        End If

        If ndDataNode Is Nothing Then
            If sName Is Nothing Then
                Throw New ActiveBankException( _
                        ErrorNumbers.TheElementIsMandatory, _
                        "The element is mandatory.", _
                        ActiveBankException.ExceptionType.Business, _
                        "abComponentServices")
            Else
                Throw New ActiveBankException( _
                        ErrorNumbers.TheElementIsMandatory, _
                        "The " + sName + " element is mandatory.", _
                        ActiveBankException.ExceptionType.Business, _
                        "abComponentServices")
            End If
        Else
            sData = ndDataNode.InnerText
            If sData.Length = 0 Then
                Throw New ActiveBankException( _
                        ErrorNumbers.TheElementIsMandatory, _
                        "The " + ndDataNode.Name + " element is mandatory.", _
                        ActiveBankException.ExceptionType.Business, _
                        "abComponentServices")
            Else
                dblData = CType(sData, System.Double)
            End If
        End If

        Return dblData

    End Function

    Public Shared Function GetDoubleXMLElementData( _
            ByVal sName As String, _
            ByRef ndParent As Xml.XmlNode, _
            ByVal dblDefault As Double) As Double

        Dim dblData As Double
        Dim sData As String
        Dim ndDataNode As Xml.XmlNode

        If sName Is Nothing Then
            ndDataNode = ndParent
        Else
            If sName.Length = 0 Then
                ndDataNode = ndParent
            Else
                ndDataNode = ndParent.SelectSingleNode(sName)
            End If
        End If

        If ndDataNode Is Nothing Then
            dblData = dblDefault
        Else
            sData = ndDataNode.InnerText
            If sData.Length = 0 Then
                dblData = dblDefault
            Else
                dblData = CType(sData, System.Double)
            End If
        End If

        Return dblData

    End Function

    Public Shared Function GetDateXMLElementData( _
            ByVal sName As String, _
            ByRef ndParent As Xml.XmlNode, _
            ByVal dteDefault As Date) As Date

        Dim sData As String
        Dim ndDataNode As Xml.XmlNode
        Dim dteData As Date
        Dim iYear As Integer
        Dim iMonth As Integer
        Dim iDay As Integer
        Dim iHour As Integer
        Dim iMinute As Integer
        Dim iSecond As Integer

        If sName Is Nothing Then
            ndDataNode = ndParent
        Else
            If sName.Length = 0 Then
                ndDataNode = ndParent
            Else
                ndDataNode = ndParent.SelectSingleNode(sName)
            End If
        End If

        If ndDataNode Is Nothing Then
            dteData = dteDefault
        Else
            iYear = GetIntegerXMLElementData("Year", ndDataNode, 0)
            If iYear = 0 Then
                dteData = dteDefault
            Else
                iMonth = GetIntegerXMLElementData("Month", ndDataNode, 0)
                iDay = GetIntegerXMLElementData("Day", ndDataNode, 0)
                iHour = GetIntegerXMLElementData("Hour", ndDataNode, 0)
                iMinute = GetIntegerXMLElementData("Minute", ndDataNode, 0)
                iSecond = GetIntegerXMLElementData("Second", ndDataNode, 0)

                Try
                    dteData = New Date(iYear, iMonth, iDay, iHour, iMinute, iSecond)
                Catch oError As Exception
                    Throw New ActiveBankException(ErrorNumbers.InvalidDate, "Invalid Date : @@1@" & sName & "@@")
                End Try

            End If
        End If

        Return dteData

    End Function

    Public Shared Function GetBooleanXMLElementData( _
            ByVal sName As String, _
            ByRef ndParent As Xml.XmlNode, _
            ByVal bDefault As Boolean) As Boolean

        Dim bData As Boolean
        Dim sData As String
        Dim ndDataNode As Xml.XmlNode

        If sName Is Nothing Then
            ndDataNode = ndParent
        Else
            If sName.Length = 0 Then
                ndDataNode = ndParent
            Else
                ndDataNode = ndParent.SelectSingleNode(sName)
            End If
        End If

        If ndDataNode Is Nothing Then
            bData = bDefault
        Else
            sData = ndDataNode.InnerText
            If sData.Length = 0 Then
                bData = bDefault
            Else
                Select Case sData
                    Case "1", "True"
                        bData = True
                    Case Else
                        bData = False
                End Select
            End If
        End If

        Return bData

    End Function

    Public Shared Function GetBooleanXMLElementData( _
            ByVal sName As String, _
            ByRef ndParent As Xml.XmlNode) As Boolean

        Dim bData As Boolean
        Dim sData As String
        Dim ndDataNode As Xml.XmlNode

        If sName Is Nothing Then
            ndDataNode = ndParent
        Else
            If sName.Length = 0 Then
                ndDataNode = ndParent
            Else
                ndDataNode = ndParent.SelectSingleNode(sName)
            End If
        End If

        If ndDataNode Is Nothing Then
            If sName Is Nothing Then
                Throw New ActiveBankException(ErrorNumbers.TheElementIsMandatory, "The element is mandatory.", ActiveBankException.ExceptionType.Business, "abComponentServices")
            Else
                Throw New ActiveBankException(ErrorNumbers.TheElementIsMandatory, "The " + sName + " element is mandatory.", ActiveBankException.ExceptionType.Business, "abComponentServices")
            End If
        Else
            sData = ndDataNode.InnerText
            If sData.Length = 0 Then
                Throw New ActiveBankException(ErrorNumbers.TheElementIsMandatory, "The " + ndDataNode.Name + " element is mandatory.", ActiveBankException.ExceptionType.Business, "abComponentServices")
            Else
                Select Case sData
                    Case "1", "True"
                        bData = True
                    Case Else
                        bData = False
                End Select
            End If
        End If

        Return bData

    End Function

    Public Shared Function GetIntegerXMLElementData( _
            ByVal sName As String, _
            ByRef ndParent As Xml.XmlNode, _
            ByVal iDefault As Integer) As Integer

        Dim iData As Integer
        Dim sData As String
        Dim ndDataNode As Xml.XmlNode

        If sName Is Nothing Then
            ndDataNode = ndParent
        Else
            If sName.Length = 0 Then
                ndDataNode = ndParent
            Else
                ndDataNode = ndParent.SelectSingleNode(sName)
            End If
        End If

        If ndDataNode Is Nothing Then
            iData = iDefault
        Else
            sData = ndDataNode.InnerXml
            If sData.Length = 0 Then
                iData = iDefault
            Else
                iData = CType(sData, System.Int32)
            End If
        End If

        Return iData

    End Function

    Public Shared Function GetLongXMLElementData( _
            ByVal sName As String, _
            ByRef ndParent As Xml.XmlNode) As Long

        Dim lData As Long
        Dim sData As String
        Dim ndDataNode As Xml.XmlNode

        If sName Is Nothing Then
            ndDataNode = ndParent
        Else
            If sName.Length = 0 Then
                ndDataNode = ndParent
            Else
                ndDataNode = ndParent.SelectSingleNode(sName)
            End If
        End If

        If ndDataNode Is Nothing Then
            If sName Is Nothing Then
                Throw New ActiveBankException( _
                        ErrorNumbers.TheElementIsMandatory, _
                        "The element is mandatory.", _
                        ActiveBankException.ExceptionType.Business, _
                        "abComponentServices")
            Else
                Throw New ActiveBankException( _
                        ErrorNumbers.TheElementIsMandatory, _
                        "The " + sName + " element is mandatory.", _
                        ActiveBankException.ExceptionType.Business, _
                        "abComponentServices")
            End If
        Else
            sData = ndDataNode.InnerXml
            If sData.Length = 0 Then
                Throw New ActiveBankException( _
                        ErrorNumbers.TheElementIsMandatory, _
                        "The " + ndDataNode.Name + " element is mandatory.", _
                        ActiveBankException.ExceptionType.Business, _
                        "abComponentServices")
            Else
                lData = CType(sData, System.Int64)
            End If
        End If

        Return lData

    End Function

    Public Shared Function GetLongXMLElementData( _
            ByVal sName As String, _
            ByRef ndParent As Xml.XmlNode, _
            ByVal lDefault As Long) As Long

        Dim lData As Long
        Dim sData As String
        Dim ndDataNode As Xml.XmlNode

        If sName Is Nothing Then
            ndDataNode = ndParent
        Else
            If sName.Length = 0 Then
                ndDataNode = ndParent
            Else
                ndDataNode = ndParent.SelectSingleNode(sName)
            End If
        End If

        If ndDataNode Is Nothing Then
            lData = lDefault
        Else
            sData = ndDataNode.InnerXml
            If sData.Length = 0 Then
                lData = lDefault
            Else
                lData = CType(sData, System.Int64)
            End If
        End If

        Return lData

    End Function

    Public Shared Function GetIntegerXMLElementData( _
            ByVal sName As String, _
            ByRef ndParent As Xml.XmlNode) As Integer

        Dim iData As Integer
        Dim sData As String
        Dim ndDataNode As Xml.XmlNode

        If sName Is Nothing Then
            ndDataNode = ndParent
        Else
            If sName.Length = 0 Then
                ndDataNode = ndParent
            Else
                ndDataNode = ndParent.SelectSingleNode(sName)
            End If
        End If

        If ndDataNode Is Nothing Then
            If sName Is Nothing Then
                Throw New ActiveBankException( _
                        ErrorNumbers.TheElementIsMandatory, _
                        "The element is mandatory.", _
                        ActiveBankException.ExceptionType.Business, _
                        "abComponentServices")
            Else
                Throw New ActiveBankException( _
                        ErrorNumbers.TheElementIsMandatory, _
                        "The " + sName + " element is mandatory.", _
                        ActiveBankException.ExceptionType.Business, _
                        "abComponentServices")
            End If
        Else
            sData = ndDataNode.InnerXml
            If sData.Length = 0 Then
                Throw New ActiveBankException( _
                        ErrorNumbers.TheElementIsMandatory, _
                        "The " + ndDataNode.Name + " element is mandatory.", _
                        ActiveBankException.ExceptionType.Business, _
                        "abComponentServices")
            Else
                iData = CType(sData, System.Int32)
            End If
        End If

        Return iData

    End Function

    Public Shared Function GetStringXMLAttributeData( _
            ByVal sName As String, _
            ByRef ndParent As Xml.XmlNode) As String

        Return GetStringXMLAttributeData(sName, CType(ndParent, Xml.XmlElement))

    End Function

    Public Shared Function GetStringXMLAttributeData( _
            ByVal sName As String, _
            ByRef ndParent As Xml.XmlNode, _
            ByVal sDefault As String) As String

        Return GetStringXMLAttributeData(sName, CType(ndParent, Xml.XmlElement), sDefault)

    End Function

    Public Shared Function GetStringXMLAttributeData( _
            ByVal sName As String, _
            ByRef ndParent As Xml.XmlElement) As String

        Dim sValue As String

        sValue = GetStringXMLAttributeData(sName, ndParent, "")
        If sValue.Length = 0 Then
            Throw New ActiveBankException( _
                    ErrorNumbers.TheAttributeIsMandatory, _
                    "The XML Attribute [" + sName + "] is mandatory.", _
                    ActiveBankException.ExceptionType.Business, _
                    "abComponentServices")
        Else
            Return sValue
        End If

    End Function

    Public Shared Function GetStringXMLAttributeData( _
            ByVal sName As String, _
            ByRef ndParent As Xml.XmlElement, _
            ByVal sDefault As String) As String

        Dim sValue As String

        sValue = ndParent.GetAttribute(sName)
        If sValue.Length = 0 Then
            sValue = sDefault
        End If

        Return sValue

    End Function

    Public Shared Function GetIntegerXMLAttributeData( _
            ByVal sName As String, _
            ByRef ndParent As Xml.XmlElement, _
            ByVal iDefault As Integer) As Integer

        Dim sValue As String
        Dim iValue As Integer

        sValue = ndParent.GetAttribute(sName)
        If sValue.Length = 0 Then
            iValue = iDefault
        Else
            iValue = CType(sValue, System.Int32)
        End If

        Return iValue

    End Function

    Public Shared Function GetDoubleXMLAttributeData( _
            ByVal sName As String, _
            ByRef ndParent As Xml.XmlElement, _
            ByVal dblDefault As Double) As Double

        Dim sValue As String
        Dim dblValue As Double

        sValue = ndParent.GetAttribute(sName)
        If sValue.Length = 0 Then
            dblValue = dblDefault
        Else
            dblValue = CType(sValue, System.Double)
        End If

        Return dblValue

    End Function

    Public Shared Function GetDecimalXMLAttributeData( _
            ByVal sName As String, _
            ByRef ndParent As Xml.XmlElement, _
            ByVal decDefault As Decimal) As Decimal

        Dim sValue As String
        Dim decValue As Decimal

        sValue = ndParent.GetAttribute(sName)
        If sValue.Length = 0 Then
            decValue = decDefault
        Else
            decValue = CDec(sValue)
        End If

        Return decValue

    End Function

    Public Shared Function GetBooleanXMLAttributeData( _
            ByVal sName As String, _
            ByRef ndParent As Xml.XmlElement, _
            ByVal bDefault As Boolean) As Boolean

        Dim sData As String
        Dim bData As Boolean

        sData = ndParent.GetAttribute(sName)
        If sData.Length = 0 Then
            bData = bDefault
        Else
            Select Case sData
                Case "1", "True"
                    bData = True
                Case Else
                    bData = False
            End Select
        End If

        Return bData

    End Function

    Public Shared Function GetByteXMLAttributeDate( _
            ByVal sName As String, _
            ByRef ndParent As Xml.XmlElement, _
            ByVal bytDefault As Byte) As Byte

        Dim sValue As String
        Dim bytValue As Byte

        sValue = ndParent.GetAttribute(sName)
        If sValue.Length = 0 Then
            bytValue = bytDefault
        Else
            bytValue = CByte(sValue)
        End If

        Return bytValue



    End Function

    Public Shared Sub RenameNode( _
            ByVal sNewName As String, _
            ByVal sCurrentName As String, _
            ByRef ndParent As Xml.XmlNode)

        Dim ndCurrent As Xml.XmlNode

        ndCurrent = ndParent.SelectSingleNode(sCurrentName)
        If Not (ndCurrent Is Nothing) Then
            RenameNode(sNewName, ndCurrent)
        End If

    End Sub

    Public Shared Sub RenameNode( _
            ByVal sNewName As String, _
            ByRef ndCurrent As Xml.XmlNode)

        Dim elmNew As Xml.XmlElement
        Dim docParent As Xml.XmlDocument
        Dim ndParent As Xml.XmlNode
        Dim ndmAttributes As Xml.XmlNamedNodeMap
        Dim iAttributeCount As Int32
        Dim iAttributeIndex As Int32
        Dim ndAttribute As Xml.XmlNode
        Dim sAttributeName As String
        Dim sAttributeValue As String

        If Not (ndCurrent Is Nothing) Then
            ndParent = ndCurrent.ParentNode
            docParent = ndCurrent.OwnerDocument 'NB ndParent.OwnerDocument gives error if ndParent is root node of the doc

            elmNew = docParent.CreateElement(sNewName)

            'copy attributes
            ndmAttributes = ndCurrent.Attributes
            iAttributeCount = ndmAttributes.Count - 1
            For iAttributeIndex = 0 To iAttributeCount
                ndAttribute = ndmAttributes.Item(iAttributeIndex)
                sAttributeName = ndAttribute.Name
                sAttributeValue = ndAttribute.InnerText
                Call elmNew.SetAttribute(sAttributeName, sAttributeValue)
            Next

            'copy sub-nodes
            elmNew.InnerXml = ndCurrent.InnerXml
            If ndParent.ParentNode Is Nothing Then
                Call docParent.RemoveChild(docParent.FirstChild)
                ndParent.AppendChild(elmNew)
                ndCurrent = elmNew
            Else
                ndParent.AppendChild(elmNew)
                ndParent.ReplaceChild(elmNew, ndCurrent)
            End If
        End If

    End Sub

    Public Shared Function AppendXMLElement( _
            ByVal sNewName As String, _
            ByRef ndParent As Xml.XmlNode, _
            ByVal sData As String) As Xml.XmlNode

        Dim ndNewNode As Xml.XmlNode
        Dim docParent As Xml.XmlDocument

        If Not (ndParent Is Nothing) Then
            docParent = ndParent.OwnerDocument

            ndNewNode = docParent.CreateElement(sNewName)
            ndNewNode.InnerText = sData
            ndParent.AppendChild(ndNewNode)
        End If

        Return ndNewNode

    End Function

    Public Shared Function AppendXMLElement( _
            ByVal sNewName As String, _
            ByRef ndParent As Xml.XmlNode, _
            ByVal dteData As Date) As Xml.XmlNode

        AppendDateNode(ndParent, sNewName, dteData)

    End Function


    Public Shared Function AppendXMLElement( _
            ByVal sNewName As String, _
            ByRef ndParent As Xml.XmlNode, _
            ByVal bData As Boolean) As Xml.XmlNode

        Dim ndNewNode As Xml.XmlNode
        Dim docParent As Xml.XmlDocument

        If Not (ndParent Is Nothing) Then
            docParent = ndParent.OwnerDocument

            ndNewNode = docParent.CreateElement(sNewName)
            ndNewNode.InnerText = GetISODataValue(bData)
            ndParent.AppendChild(ndNewNode)
        End If

        Return ndNewNode

    End Function

    Public Shared Function AppendXMLElement( _
            ByVal sNewName As String, _
            ByRef ndParent As Xml.XmlNode, _
            ByVal iData As Integer) As Xml.XmlNode

        Dim ndNewNode As Xml.XmlNode
        Dim docParent As Xml.XmlDocument

        If Not (ndParent Is Nothing) Then
            docParent = ndParent.OwnerDocument

            ndNewNode = docParent.CreateElement(sNewName)
            ndNewNode.InnerXml = iData.ToString(System.Globalization.CultureInfo.InvariantCulture)

            ndParent.AppendChild(ndNewNode)
        End If

        Return ndNewNode

    End Function

    Public Shared Function AppendXMLElement( _
            ByVal sNewName As String, _
            ByRef ndParent As Xml.XmlNode, _
            ByVal decData As Decimal) As Xml.XmlNode

        Dim ndNewNode As Xml.XmlNode
        Dim docParent As Xml.XmlDocument

        If Not (ndParent Is Nothing) Then
            docParent = ndParent.OwnerDocument

            ndNewNode = docParent.CreateElement(sNewName)
            ndNewNode.InnerXml = decData.ToString(System.Globalization.CultureInfo.InvariantCulture)
            ndParent.AppendChild(ndNewNode)
        End If

        Return ndNewNode

    End Function

    Public Shared Function AppendXMLElement( _
            ByVal sNewName As String, _
            ByRef ndParent As Xml.XmlNode, _
            ByVal dblData As Double) As Xml.XmlNode

        Dim ndNewNode As Xml.XmlNode
        Dim docParent As Xml.XmlDocument

        If Not (ndParent Is Nothing) Then
            docParent = ndParent.OwnerDocument

            ndNewNode = docParent.CreateElement(sNewName)
            ndNewNode.InnerXml = dblData.ToString(System.Globalization.CultureInfo.InvariantCulture)
            ndParent.AppendChild(ndNewNode)
        End If

        Return ndNewNode

    End Function

    Public Shared Function AppendXMLElement( _
        ByVal sNewName As String, _
        ByRef ndParent As Xml.XmlNode) As Xml.XmlNode

        Dim ndNewNode As Xml.XmlNode
        Dim docParent As Xml.XmlDocument

        If Not (ndParent Is Nothing) Then
            docParent = ndParent.OwnerDocument

            ndNewNode = docParent.CreateElement(sNewName)
            ndParent.AppendChild(ndNewNode)
        End If

        Return ndNewNode

    End Function

    Public Sub COMInitializeComponent()
        InitializeComponent(Me, False)
    End Sub

    Public Shared Function InitializeComponent( _
            ByRef oComponentServices As ComponentServices) As ComponentServices

        Return InitializeComponent( _
                oComponentServices, _
                True)

    End Function

    Public Shared Function InitializeComponent( _
            ByRef oComponentServices As ComponentServices, _
            ByVal sSessionID As String, _
            ByVal sBranchID As String) As ComponentServices

        Return InitializeComponent( _
                oComponentServices, _
                True, _
                sSessionID, _
                sBranchID)

    End Function

    Public Shared Function InitializeComponent( _
            ByRef oComponentServices As ComponentServices, _
            ByVal bRequiresTransaction As Boolean) As ComponentServices

        Return InitializeComponent( _
                oComponentServices, _
                bRequiresTransaction, _
                "")

    End Function

    Public Shared Function InitializeComponent( _
            ByRef oComponentServices As ComponentServices, _
            ByVal bRequiresTransaction As Boolean, _
            ByVal sSessionID As String, _
            ByVal sBranchID As String) As ComponentServices

        Return InitializeComponent( _
                oComponentServices, _
                bRequiresTransaction, _
                "", _
                sSessionID, _
                sBranchID)

    End Function

    Public Shared Function InitializeComponent( _
            ByRef oComponentServices As ComponentServices, _
            ByVal sDatabaseIdentifier As String) As ComponentServices

        Return InitializeComponent( _
                oComponentServices, _
                False, _
                sDatabaseIdentifier)

    End Function

    Public Shared Function InitializeComponent( _
            ByRef oComponentServices As ComponentServices, _
            ByVal sDatabaseIdentifier As String, _
            ByVal sSessionID As String, _
            ByVal sBranchID As String) As ComponentServices

        Return InitializeComponent( _
                oComponentServices, _
                False, _
                sDatabaseIdentifier, _
                sSessionID, _
                sBranchID)

    End Function

    Public Shared Function InitializeComponent( _
            ByRef oComponentServices As ComponentServices, _
            ByVal bRequiresTransaction As Boolean, _
            ByVal sDatabaseIdentifier As String) As ComponentServices


        Return InitializeComponent( _
                oComponentServices, _
                bRequiresTransaction, _
                sDatabaseIdentifier, _
                "", _
                "")

    End Function

    Public Shared Function InitializeComponent( _
            ByRef oComponentServices As ComponentServices, _
            ByVal bRequiresTransaction As Boolean, _
            ByVal sDatabaseIdentifier As String, _
            ByVal sSessionID As String, _
            ByVal sBranchID As String) As ComponentServices

        Dim pcSession As ParameterCollection

        If oComponentServices Is Nothing Then
            oComponentServices = New ComponentServices

            If sSessionID.Length > 0 Then
                oComponentServices.m_sSessionID = sSessionID
            End If
            If sBranchID.Length > 0 Then
                oComponentServices.m_sBranchID = sBranchID
            End If

            oComponentServices.ActivateComponentServices(sDatabaseIdentifier, bRequiresTransaction)

            'Decrement the reference count (this is done when calling from a COM component)
            oComponentServices.m_iReferenceCount -= 1
        Else

            If sSessionID.Length > 0 Then
                oComponentServices.m_sSessionID = sSessionID
            End If
            If sBranchID.Length > 0 Then
                oComponentServices.m_sBranchID = sBranchID
            End If

            If bRequiresTransaction Then
                oComponentServices.StartTransaction()
            End If
        End If

        oComponentServices.m_iReferenceCount += 1
        If sSessionID.Length > 0 Then
            oComponentServices.m_sSessionID = sSessionID
        End If
        If sBranchID.Length > 0 Then
            oComponentServices.m_sBranchID = sBranchID
        End If

        If sSessionID.Length > 0 Then
            pcSession = oComponentServices.GetInstanceParameterCollection( _
                "abUser", _
                "Session", _
                sSessionID, _
                "SessionID", _
                True)

            If pcSession Is Nothing Then
                Throw New ActiveBankException( _
                        ErrorNumbers.InstanceHasNotBeenFound, _
                        "The specified instance cannot be found.", _
                        ActiveBankException.ExceptionType.Business, _
                        "abComponentServices")
            Else

                oComponentServices.m_decUserID = pcSession.GetDecimalValue("UserID", 0D)
            End If
        End If

        LogInformation("ReferenceCount : " & oComponentServices.m_iReferenceCount.ToString, "abComponentServices", "InitializeComponent")

        Return oComponentServices

    End Function

    '================================================================================================
    ' About...  : Function returns the Unique RequestId & generates one if one has not been set
    '================================================================================================

    Public ReadOnly Property RequestId() As String
        Get
            If m_sRequestId = String.Empty Then
                m_sRequestId = System.Guid.NewGuid().ToString()
            End If

            RequestId = m_sRequestId
        End Get
    End Property

    Public Sub COMFinalizeComponent()

        LogInformation("ReferenceCount : " & m_iReferenceCount.ToString, "abComponentServices", "COMFinalizeComponent")

        m_iReferenceCount = m_iReferenceCount - 1

        If m_iReferenceCount = 0 Then
            DisposeComponentServices()
        End If

    End Sub

    Public Shared Sub FinalizeComponent(ByRef oComponentServices As ComponentServices)

        FinalizeComponent(oComponentServices, False)

    End Sub

    Public Shared Sub FinalizeComponent( _
            ByRef oComponentServices As ComponentServices, _
            ByVal bAbortTransaction As Boolean)

        If Not (oComponentServices Is Nothing) Then
            Try
                If bAbortTransaction Then
                    oComponentServices.AbortTransaction()
                End If

                LogInformation("ReferenceCount : " & oComponentServices.m_iReferenceCount.ToString, "abComponentServices", "FinalizeComponent")

                oComponentServices.m_iReferenceCount -= 1

                If oComponentServices.m_iReferenceCount = 0 Then
                    oComponentServices.DisposeComponentServices()
                End If

            Catch ex As System.Exception
                LogError(True, "ComponentServices", ex)
            Finally
                If oComponentServices.m_iReferenceCount = 0 Then
                    oComponentServices = Nothing
                End If
            End Try

        End If

    End Sub

    Public Shared Function CreateComponentInstance( _
            ByVal sComponentName As String, _
            ByVal sClassName As String) As Object

        Return CreateComponentInstance( _
                sComponentName, _
                sClassName, _
                Nothing)

    End Function

    Public Shared Function CreateComponentInstance( _
            ByVal sComponentName As String, _
            ByVal sClassName As String, _
            ByRef oComponentServices As ComponentServices) As Object

        Dim oComponent As Object
        Dim sbAssemblyName As New StringBuilder(GetPathFromFileNamePath(System.Reflection.Assembly.GetExecutingAssembly().Location))
        Dim sbTypeName As New StringBuilder("FinancialObjects.ActiveBank.")

        Dim sAssemblyName As String
        Dim sTypeName As String
        Dim oComponentObjectHandle As System.Runtime.Remoting.ObjectHandle
        Dim oComponentIABComponent As abComponentServices.IABComponent

        sbAssemblyName.Append("\")
        sbAssemblyName.Append(sComponentName)
        sbAssemblyName.Append(".dll")
        sAssemblyName = sbAssemblyName.ToString()

        sbTypeName.Append(sComponentName)
        sbTypeName.Append(".")
        sbTypeName.Append(sClassName)
        sTypeName = sbTypeName.ToString()

        Try
            oComponentObjectHandle = Activator.CreateInstanceFrom(sAssemblyName, sTypeName)
        Catch oWarning As System.Exception
            LogInformation( _
                    "Cannot create [" + sComponentName + "]" + oWarning.Message, _
                    "abComponentServices", _
                    "CreateComponentInstance")

            'Might be called .2.dll.
            'This convention is used for new components with same name as old component.
            sbAssemblyName.Remove(sbAssemblyName.Length - 3, 3)
            sbAssemblyName.Append("2.dll")
            sAssemblyName = sbAssemblyName.ToString()
            oComponentObjectHandle = Activator.CreateInstanceFrom(sAssemblyName, sTypeName)
        End Try

        oComponent = oComponentObjectHandle.Unwrap()

        If Not (oComponentServices Is Nothing) Then
            Try
                oComponentIABComponent = CType(oComponent, abComponentServices.IABComponent)
            Catch oError As System.Exception
                LogError( _
                        True, _
                        "Component: " + sComponentName + ", class: " + sClassName + " does not have an IABComponent interface.", _
                        "abComponentServices", _
                        oError.StackTrace)
            End Try

            oComponentIABComponent.SetComponentServices(oComponentServices)
        End If

        Return oComponent

    End Function

    Public Sub DisposeCOMComponentInstance(ByRef oCOMObject As Object)

        System.Runtime.InteropServices.Marshal.ReleaseComObject(oCOMObject)
        oCOMObject = Nothing

        LogInformation("COMReferenceCount : " & m_iCOMReferenceCount.ToString, "abComponentServices", "DisposeCOMComponentInstance")

        m_iCOMReferenceCount -= 1

        If m_iCOMReferenceCount = 0 Then
            System.Runtime.InteropServices.Marshal.ReleaseComObject(m_oCOMComponentServices)
            m_oCOMComponentServices = Nothing
        End If

    End Sub

    Public Function CreateCOMComponentInstance(ByVal sProgID As String) As Object

        Dim oComponent As Object
        Dim oCOMComponentControl As comComponentServices.IABComponent

        If (m_oCOMComponentServices Is Nothing) Then
            Try
                m_oCOMComponentServices = New comComponentServices.ComponentServices

            Catch oWarning As System.Exception
                LogInformation( _
                        "Cannot create [comComponentServices.ComponentServices]" + oWarning.Message, _
                        "abComponentServices", _
                        "CreateCOMComponentInstance")
            End Try

            If Not (m_oCOMComponentServices Is Nothing) Then

                Try
                    m_oCOMComponentServices.Initialize(Me)
                    m_oCOMComponentServices.UserID = m_decUserID
                Catch oWarning As System.Exception
                    LogInformation( _
                            "Cannot initialize [comComponentServices.ComponentServices]" + oWarning.Message, _
                            "abComponentServices", _
                            "CreateCOMComponentInstance")
                End Try
            End If
        End If

        If Not (m_oCOMComponentServices Is Nothing) Then
            Try
                oComponent = CreateObject(sProgID)
                oCOMComponentControl = oComponent
                Call oCOMComponentControl.SetComponentServices(m_oCOMComponentServices)

            Catch oWarning As System.Exception
                LogInformation( _
                        "Cannot create [" + sProgID + "]" + oWarning.Message, _
                        "abComponentServices", _
                        "CreateCOMComponentInstance")
            End Try

        End If

        m_iCOMReferenceCount += 1

        LogInformation("COMReferenceCount : " & m_iCOMReferenceCount.ToString, "abComponentServices", "CreateCOMComponentInstance")

        Return oComponent

    End Function

    Public Shared Function CreateComponentInstance( _
            ByVal sProgID As String, _
            ByRef oComponentServices As abComponentServices.ComponentServices, _
            ByVal sSessionID As String, _
            ByVal sBranchID As String) As Object

        Dim oComponent As Object

        Try
            oComponent = CreateObject(sProgID)

            If sSessionID = String.Empty Then
                sSessionID = oComponentServices.m_sSessionID
            End If
            If sBranchID = String.Empty Then
                sBranchID = oComponentServices.m_sBranchID
            End If

            oComponentServices.SetupClassicComponent( _
                    oComponent, _
                    sSessionID, _
                    sBranchID)
        Catch oWarning As System.Exception
            LogError( _
                    False, _
                    "Classic ActiveBank Component (" + sProgID + ") cannot be instantiated." & " " & oWarning.Message, _
                    "abComponentServices", _
                    "", _
                    ErrorSeverity.ES_Warning)

            oComponent = Nothing
        End Try

        Return oComponent

    End Function

    Public Shared Sub AppendFrequencyNode( _
            ByRef ndParent As Xml.XmlNode, _
            ByVal strName As String, _
            ByVal iType As Integer, _
            ByVal iPrecision As Integer, _
            ByVal iPrecisionType As Integer, _
            ByVal iPeriod As Integer)

        Dim ndFrequency As Xml.XmlNode
        Dim ndNode As Xml.XmlNode
        Dim docParent As Xml.XmlDocument

        docParent = ndParent.OwnerDocument

        ndFrequency = docParent.CreateElement(strName)
        ndParent.AppendChild(ndFrequency)

        ndNode = docParent.CreateElement("Type")
        ndNode.InnerXml = iType.ToString(System.Globalization.CultureInfo.InvariantCulture)
        ndFrequency.AppendChild(ndNode)

        ndNode = docParent.CreateElement("Prec")
        ndNode.InnerXml = iPrecision.ToString(System.Globalization.CultureInfo.InvariantCulture)
        ndFrequency.AppendChild(ndNode)

        ndNode = docParent.CreateElement("PrecType")
        ndNode.InnerXml = iPrecisionType.ToString(System.Globalization.CultureInfo.InvariantCulture)
        ndFrequency.AppendChild(ndNode)

        ndNode = docParent.CreateElement("Prd")
        ndNode.InnerXml = iPeriod.ToString(System.Globalization.CultureInfo.InvariantCulture)
        ndFrequency.AppendChild(ndNode)

    End Sub

    Public Shared Sub AppendDateNode( _
            ByRef ndParent As Xml.XmlNode, _
            ByVal sName As String, _
            ByVal dteDate As DateTime)

        Dim ndDate As Xml.XmlNode
        Dim docParent As Xml.XmlDocument

        docParent = ndParent.OwnerDocument

        ndDate = docParent.CreateNode(Xml.XmlNodeType.Element, sName, "")
        ndParent.AppendChild(ndDate)

        AppendDateNode(ndDate, dteDate)

    End Sub

    Public Shared Sub AppendDateNode( _
            ByRef ndParent As Xml.XmlNode, _
            ByVal dteDate As DateTime)

        Dim docParent As Xml.XmlDocument
        Dim ndDateElement As Xml.XmlNode
        Dim iDateElement As Integer

        docParent = ndParent.OwnerDocument

        iDateElement = dteDate.Year()
        If iDateElement > 0 Then
            ndDateElement = docParent.CreateNode(Xml.XmlNodeType.Element, "Year", String.Empty)
            ndDateElement.InnerXml = iDateElement.ToString(System.Globalization.CultureInfo.InvariantCulture)
            ndParent.AppendChild(ndDateElement)
        End If

        iDateElement = dteDate.Month()
        If iDateElement > 0 Then
            ndDateElement = docParent.CreateNode(Xml.XmlNodeType.Element, "Month", String.Empty)
            ndDateElement.InnerXml = iDateElement.ToString(System.Globalization.CultureInfo.InvariantCulture)
            ndParent.AppendChild(ndDateElement)
        End If

        iDateElement = dteDate.Day()
        If iDateElement > 0 Then
            ndDateElement = docParent.CreateNode(Xml.XmlNodeType.Element, "Day", String.Empty)
            ndDateElement.InnerXml = iDateElement.ToString(System.Globalization.CultureInfo.InvariantCulture)
            ndParent.AppendChild(ndDateElement)
        End If

        iDateElement = dteDate.Hour()
        If iDateElement > 0 Then
            ndDateElement = docParent.CreateNode(Xml.XmlNodeType.Element, "Hour", String.Empty)
            ndDateElement.InnerXml = iDateElement.ToString(System.Globalization.CultureInfo.InvariantCulture)
            ndParent.AppendChild(ndDateElement)
        End If

        iDateElement = dteDate.Minute()
        If iDateElement > 0 Then
            ndDateElement = docParent.CreateNode(Xml.XmlNodeType.Element, "Minute", String.Empty)
            ndDateElement.InnerXml = iDateElement.ToString(System.Globalization.CultureInfo.InvariantCulture)
            ndParent.AppendChild(ndDateElement)
        End If

        iDateElement = dteDate.Second()
        If iDateElement > 0 Then
            ndDateElement = docParent.CreateNode(Xml.XmlNodeType.Element, "Second", String.Empty)
            ndDateElement.InnerXml = iDateElement.ToString(System.Globalization.CultureInfo.InvariantCulture)
            ndParent.AppendChild(ndDateElement)
        End If

    End Sub

    Public Shared Sub AppendAmountNode( _
            ByRef ndParent As Xml.XmlNode, _
            ByVal curAmount As Double, _
            ByVal strCurrencyISOCode As String)

        AppendAmountNode( _
                ndParent, _
                curAmount, _
                strCurrencyISOCode, _
                0)

    End Sub

    Public Shared Sub AppendAmountNode( _
            ByRef ndParent As Xml.XmlNode, _
            ByVal curAmount As Double, _
            ByVal sCurrencyISOCode As String, _
            ByVal dblConversionRate As Double)

        Dim docParent As Xml.XmlDocument
        Dim ndAmountElement As Xml.XmlNode

        docParent = ndParent.OwnerDocument

        ndAmountElement = docParent.CreateNode(Xml.XmlNodeType.Element, "Amt", "")
        ndAmountElement.InnerXml = curAmount.ToString(System.Globalization.CultureInfo.InvariantCulture)
        ndParent.AppendChild(ndAmountElement)

        If Not (sCurrencyISOCode Is Nothing) Then
            If sCurrencyISOCode.Length <> 0 Then
                ndAmountElement = docParent.CreateNode(Xml.XmlNodeType.Element, "CurCode", "")
                ndAmountElement.InnerXml = sCurrencyISOCode
                ndParent.AppendChild(ndAmountElement)
            End If
        End If

        If dblConversionRate > 0 Then
            ndAmountElement = docParent.CreateNode(Xml.XmlNodeType.Element, "CurRate", "")
            ndAmountElement.InnerXml = CStr(dblConversionRate)
            ndParent.AppendChild(ndAmountElement)
        End If

    End Sub

    Public Shared Sub AppendAmountNode( _
            ByRef ndParent As Xml.XmlNode, _
            ByVal strName As String, _
            ByVal curAmount As Double, _
            ByVal strCurrencyISOCode As String)

        AppendAmountNode( _
                ndParent, _
                strName, _
                curAmount, _
                strCurrencyISOCode, _
                0)

    End Sub

    Public Shared Sub AppendAmountNode( _
            ByRef ndParent As Xml.XmlNode, _
            ByVal strName As String, _
            ByVal curAmount As Double, _
            ByVal strCurrencyISOCode As String, _
            ByVal dblConversionRate As Double)

        Dim ndAmountParent As Xml.XmlNode

        ndAmountParent = ndParent.OwnerDocument.CreateNode(Xml.XmlNodeType.Element, strName, "")
        ndParent.AppendChild(ndAmountParent)

        AppendAmountNode(ndAmountParent, curAmount, strCurrencyISOCode, dblConversionRate)

    End Sub

    Public Shared Sub GetValuesFromIDNode( _
            ByRef ndID As Xml.XmlElement, _
            ByRef decID As Decimal, _
            ByRef sTimestamp As String)

        decID = CDec(ndID.InnerText)
        sTimestamp = GetStringXMLAttributeData("Timestamp", ndID, "")

    End Sub

    Public Shared Sub GetValuesFromAmountNode( _
            ByRef ndAmount As Xml.XmlNode, _
            ByRef dblAmount As Double, _
            ByRef sCurrencyISOCode As String)

        dblAmount = GetDoubleXMLElementData("Amt", ndAmount, 0)
        sCurrencyISOCode = GetStringXMLElementData("CurCode", ndAmount, "")

    End Sub

    Public Shared Sub GetValuesFromFrequencyNode( _
            ByRef ndFrequency As Xml.XmlNode, _
            ByRef iFreqType As Integer, _
            ByRef iPrec As Integer, _
            ByRef iPrecType As Integer, _
            ByRef iPrd As Integer)


        iFreqType = GetIntegerXMLElementData("Type", ndFrequency)
        iPrec = GetIntegerXMLElementData("Prec", ndFrequency, 0)
        iPrecType = GetIntegerXMLElementData("PrecType", ndFrequency, 0)
        iPrd = GetIntegerXMLElementData("Prd", ndFrequency, 0)

    End Sub

    Public Shared Function GetPathFromFileNamePath( _
            ByVal sFileNamePath As String) As String

        Dim sPath As String
        Dim iBackSlashPosition As Integer
        Dim bBackSlashFound As Boolean

        bBackSlashFound = False
        iBackSlashPosition = sFileNamePath.Length - 1
        While (Not bBackSlashFound) And (iBackSlashPosition >= 0)
            If sFileNamePath.Substring(iBackSlashPosition, 1) = "\" Then
                bBackSlashFound = True
            Else
                iBackSlashPosition = iBackSlashPosition - 1
            End If
        End While

        If bBackSlashFound Then
            sPath = sFileNamePath.Substring(0, iBackSlashPosition)
        Else
            sPath = ""
        End If

        Return sPath

    End Function

    Public Shared Function CreateErrorDocument( _
            ByRef oError As abComponentServices.ActiveBankException) As Xml.XmlDocument

        If oError.Type = ActiveBankException.ExceptionType.System Then
            LogError(False, oError)
        End If

        Return CreateErrorDocument(oError.ComponentName, oError.Number, oError.Message)

    End Function

    Public Shared Function CreateErrorDocument( _
            ByVal sComponentName As String, _
            ByRef oError As System.Exception) As Xml.XmlDocument

        LogError(False, sComponentName, oError)

        Return CreateErrorDocument(sComponentName, 0, oError.Message)

    End Function

    Public Shared Function CreateErrorDocument( _
            ByVal iNumber As Long, _
            ByVal sDescription As String) As Xml.XmlDocument

        Return CreateErrorDocument("", iNumber, sDescription)

    End Function

    Public Shared Function CreateErrorDocument( _
            ByVal sComponentName As String, _
            ByVal iNumber As Long, _
            ByVal sDescription As String) As Xml.XmlDocument

        Dim docError As Xml.XmlDocument

        docError = CreateXMLDocument( _
                "abError", _
                "Number=" + iNumber.ToString(), _
                "Description=" + sDescription)

        If sComponentName.Length <> 0 Then
            AppendXMLElement("Component", docError.SelectSingleNode("abError"), sComponentName)
        End If

        Return docError

    End Function

    Public Shared Function IsErrorDocument( _
            ByRef docError As Xml.XmlDocument) As Boolean

        Dim bError As Boolean

        If docError Is Nothing Then
            bError = False
        Else
            If docError.SelectSingleNode("abError") Is Nothing Then
                bError = False
            Else
                bError = True
            End If
        End If

        Return bError

    End Function

    Public Shared Function IsErrorDocument( _
            ByRef docError As Xml.XmlDocument, _
            ByRef iError As Integer, _
            ByRef sError As String) As Boolean

        Dim bError As Boolean
        Dim ndabError As Xml.XmlNode

        bError = IsErrorDocument(docError)
        If bError Then
            ndabError = docError.SelectSingleNode("abError")
            iError = GetIntegerXMLElementData("Number", ndabError, 0)
            sError = GetStringXMLElementData("Description", ndabError, "")
        End If

        Return bError

    End Function

    Public Shared Sub SetXMLElementData( _
            ByVal sName As String, _
            ByVal decValue As Integer, _
            ByRef ndParent As Xml.XmlNode)

        Dim ndElement As Xml.XmlNode

        If Not (ndParent Is Nothing) Then
            ndElement = ndParent.SelectSingleNode(sName)
            If ndElement Is Nothing Then
                AppendXMLElement(sName, ndParent, decValue)
            Else
                ndElement.InnerText = decValue.ToString()
            End If
        End If

    End Sub

    Public Shared Sub SetXMLElementData( _
            ByVal sName As String, _
            ByVal iValue As Decimal, _
            ByRef ndParent As Xml.XmlNode)

        Dim ndElement As Xml.XmlNode

        If Not (ndParent Is Nothing) Then
            ndElement = ndParent.SelectSingleNode(sName)
            If ndElement Is Nothing Then
                AppendXMLElement(sName, ndParent, iValue)
            Else
                ndElement.InnerText = iValue.ToString()
            End If
        End If

    End Sub

    Public Shared Sub SetXMLElementData( _
            ByVal sName As String, _
            ByVal sValue As String, _
            ByRef ndParent As Xml.XmlNode)

        Dim ndElement As Xml.XmlNode

        If Not (ndParent Is Nothing) Then
            ndElement = ndParent.SelectSingleNode(sName)
            If ndElement Is Nothing Then
                AppendXMLElement(sName, ndParent, sValue)
            Else
                ndElement.InnerText = sValue
            End If
        End If

    End Sub

    Public Shared Sub SetXMLElementData( _
            ByVal sName As String, _
            ByVal bValue As Boolean, _
            ByRef ndParent As Xml.XmlNode)

        Dim ndElement As Xml.XmlNode
        Dim sValue As String

        If Not (ndParent Is Nothing) Then
            sValue = GetISODataValue(bValue)

            ndElement = ndParent.SelectSingleNode(sName)
            If ndElement Is Nothing Then
                AppendXMLElement(sName, ndParent, sValue)
            Else
                ndElement.InnerText = sValue
            End If
        End If

    End Sub

    Public Shared Sub SetXMLElementData( _
            ByVal sName As String, _
            ByVal dteValue As Date, _
            ByRef ndParent As Xml.XmlNode)

        Dim ndElement As Xml.XmlNode

        If Not (ndParent Is Nothing) Then
            ndElement = ndParent.SelectSingleNode(sName)
            If ndElement Is Nothing Then
                AppendDateNode(ndParent, sName, dteValue)
            Else
                ndElement.RemoveAll()
                AppendDateNode(ndElement, dteValue)
            End If
        End If

    End Sub

    Public Shared Sub SetXMLAttribute( _
            ByVal sName As String, _
            ByVal ndParent As Xml.XmlNode, _
            ByVal sValue As String)

        SetXMLAttribute(sName, CType(ndParent, Xml.XmlElement), sValue)

    End Sub

    Public Shared Sub SetXMLAttribute( _
            ByVal sName As String, _
            ByVal ndParent As Xml.XmlElement, _
            ByVal sValue As String)

        If Not (ndParent Is Nothing) Then
            ndParent.SetAttribute(sName, sValue)
        End If

    End Sub

    Public Shared Sub SetXMLAttribute( _
            ByVal sName As String, _
            ByVal ndParent As Xml.XmlNode, _
            ByVal iValue As Integer)

        SetXMLAttribute(sName, CType(ndParent, Xml.XmlElement), iValue)

    End Sub

    Public Shared Sub SetXMLAttribute( _
            ByVal sName As String, _
            ByVal ndParent As Xml.XmlElement, _
            ByVal iValue As Integer)

        If Not (ndParent Is Nothing) Then
            ndParent.SetAttribute(sName, iValue.ToString(System.Globalization.CultureInfo.InvariantCulture))
        End If

    End Sub

    Public Shared Sub SetXMLAttribute( _
            ByVal sName As String, _
            ByVal ndParent As Xml.XmlNode, _
            ByVal decValue As Decimal)

        SetXMLAttribute(sName, CType(ndParent, Xml.XmlElement), decValue)

    End Sub

    Public Shared Sub SetXMLAttribute( _
            ByVal sName As String, _
            ByVal ndParent As Xml.XmlElement, _
            ByVal decValue As Decimal)

        If Not (ndParent Is Nothing) Then
            ndParent.SetAttribute(sName, decValue.ToString(System.Globalization.CultureInfo.InvariantCulture))
        End If

    End Sub

    Public Shared Sub SetXMLAttribute( _
            ByVal sName As String, _
            ByRef ndParent As Xml.XmlNode, _
            ByVal bValue As Boolean)

        SetXMLAttribute(sName, CType(ndParent, Xml.XmlElement), bValue)

    End Sub

    Public Shared Sub SetXMLAttribute( _
            ByVal sName As String, _
            ByRef ndParent As Xml.XmlElement, _
            ByVal bValue As Boolean)

        Dim sValue As String

        If Not (ndParent Is Nothing) Then
            If bValue Then
                sValue = "1"
            Else
                sValue = "0"
            End If
            ndParent.SetAttribute(sName, sValue)
        End If

    End Sub

    Public Shared Function GetLoggingLevel() As abComponentServices.ComponentServices.LoggingLevel

        Return m_structConfigSettings.eLoggingLevel

    End Function

    Public Shared Sub RemoveNode(ByVal sName As String, ByRef ndParent As Xml.XmlNode)

        Dim ndNode As Xml.XmlNode

        ndNode = ndParent.SelectSingleNode(sName)
        If Not (ndNode Is Nothing) Then
            ndParent.RemoveChild(ndNode)
        End If

    End Sub

    Public Shared Function GetNullDate() As Date

        Dim dteNullDate As Date

        dteNullDate = Date.FromOADate(0)

        Return dteNullDate

    End Function

    Public Shared Function IsDateNull(ByVal dteDate As Date) As Boolean

        Dim bNull As Boolean

        If dteDate.ToOADate() = 0 Then
            bNull = True
        Else
            bNull = False
        End If

        Return bNull

    End Function

#End Region

#Region "Interop Methods"

    Public Function SetupClassicComponent( _
            ByRef oComponent As Object, _
            ByVal sSessionID As String, _
            ByVal sBranchID As String) As Object
        '================================================================================================
        ' About...  : Setup the Classic Component
        '================================================================================================
        ' Change Log
        'Date           Author       Ref            Comments
        '----           ------      ---             --------
        ' 26/02/2008    GirishV     TD24446         Set Branch property of User component along with 
        '                                           oSession.Branch on creating m_oUser object
        '=================================================================================================
        Dim oSession As Object
        Dim oBranch As Object

        GetRDOInterop()

        CallByName(oComponent, "DBHandle", CallType.Let, m_oRDOInterop)

        If m_oUser Is Nothing Then
            m_oUser = CreateObject("DiamUser.User")
            CallByName(m_oUser, "DBHandle", CallType.Let, m_oRDOInterop)

            oSession = CreateObject("DiamSession.clsSession")
            CallByName(oSession, "DBHandle", CallType.Let, m_oRDOInterop)
            CallByName(oSession, "User", CallType.Let, m_oUser)

            If sSessionID Is Nothing OrElse sSessionID = String.Empty Then
                CallByName(oSession, "bStartSession", CallType.Method, Nothing)
            Else
                Try
                    oSession.sSessionID = sSessionID
                Catch ex As System.Exception
                    CallByName(m_oRDOInterop, "CSHandle", CallType.Let, Me)

                    oSession.sSessionID = sSessionID
                End Try
            End If

            oBranch = CreateObject("DiamBranch.Branch")
            CallByName(oBranch, "DBHandle", CallType.Let, m_oRDOInterop)

            If sBranchID Is Nothing Then
                sBranchID = oSession.Branch
            Else
                If sBranchID.Length = 0 Then
                    sBranchID = oSession.Branch
                Else
                    oSession.Branch = sBranchID
                    m_oUser.Branch = sBranchID
                End If
            End If
            oBranch.LoadRemoteBranchData(sBranchID)
            oSession.sLegalEntity = oBranch.sLegalEntity

            CallByName(m_oUser, "SessionHandle", CallType.Let, oSession)
            CallByName(m_oRDOInterop, "User", CallType.Let, m_oUser)
        End If

        CallByName(oComponent, "User", CallType.Let, m_oUser)

        Try
            If m_oBroker Is Nothing Then
                m_oBroker = CreateObject("DiamBroker.ObjectResource")
            End If
        Catch
            LogError(True, _
                    "Unable to obtain broker.", _
                    "abComponentServices", _
                    "", _
                    ErrorSeverity.ES_Error)
        End Try



        Try

            CallByName(oComponent, "Broker", CallType.Let, m_oBroker)
        Catch
            LogError( _
                    False, _
                    "Classic ActiveBank Component does have have broker interface.", _
                    "abComponentServices", _
                    "", _
                    ErrorSeverity.ES_Warning)
        End Try

        If m_oAuditTrail Is Nothing Then
            m_oAuditTrail = CreateObject("ABRAuditTrail.AuditTrail")
            CallByName(m_oAuditTrail, "DBHandle", CallType.Let, m_oRDOInterop)
            'Changed by RDT - 07/10/2004
            CallByName(m_oAuditTrail, "User", CallType.Let, m_oUser)
            CallByName(m_oAuditTrail, "Broker", CallType.Let, m_oBroker)
            m_oAuditTrail.ObjectInitiate("", Nothing)
        End If
        Try
            CallByName(oComponent, "AuditTrail", CallType.Let, m_oAuditTrail)
        Catch
            LogError( _
                    False, _
                    "Classic ActiveBank Component does have an audit trail interface.", _
                    "abComponentServices", _
                    "", _
                    ErrorSeverity.ES_Warning)
        End Try

        'Added by RDT - 07/10/2004
        Try
            ReleaseCOMObject(oSession)
        Catch
            LogError(False, "ReleaseCOMObject(oSession) failed", "ComponentServices", "SetupClassicComponent")
        End Try
        If Not (oBranch Is Nothing) Then
            Try
                CallByName(oBranch, "DBHandle", CallType.Let, CType(Nothing, abRDOInterop.RDOClassClass))
            Catch
                LogError(False, "Set oBranch.DBHandle to Nothing failed", "ComponentServices", "SetupClassicComponent")
            End Try

            Try
                ReleaseCOMObject(oBranch)
            Catch
                LogError(False, "ReleaseCOMObject(oBranch) failed", "ComponentServices", "SetupClassicComponent")
            End Try
        End If

    End Function

    Private Sub GetRDOInterop()
        If m_oRDOInterop Is Nothing Then
            Try
                m_oRDOInterop = New abRDOInterop.RDOClassClass
                ' Pass a reference to this instance of ComponentServices
                'm_oRDOInterop.CSHandle = Me
                CallByName(m_oRDOInterop, "CSHandle", CallType.Let, Me)
            Catch oError As Exception
                LogError(True, "ComponentServices", oError, "GetRDOInterop failed")
            End Try

        End If

    End Sub

    Public Sub m_oRDOInterop_ExecuteQuery( _
            ByVal sQuery As String, _
            ByRef rsRecords As ADODB.Recordset, _
            ByRef lRows As Integer, _
            ByVal bAction As Boolean, _
            ByVal sComponentName As String, _
            ByVal sClassName As String, _
            ByVal bClassCached As Boolean) 'Handles m_oRDOInterop.ExecuteQuery

        Dim oRecords As Data.DataTable

        Try
            If bAction Then
                m_oDatabaseAdapter.ExecuteNonQuery(sQuery, lRows)
                If bClassCached Then 'RDO Caching Implementation
                    InformApplicationServersOfDataChange(sComponentName, sClassName)
                End If
                rsRecords = Nothing
            Else
                If bClassCached Then 'RDO Caching Implementation
                    If (sComponentName Is Nothing) Or (sClassName Is Nothing) Then
                        bClassCached = False
                    Else
                        If (sComponentName.Length = 0) And (sClassName.Length = 0) Then
                            bClassCached = False
                        Else
                            oRecords = RetrieveDataTableFromCache(sComponentName, sClassName, sQuery)
                        End If
                    End If
                End If
                If oRecords Is Nothing Then
                    oRecords = m_oDatabaseAdapter.ExecuteQueryDataTable(sQuery)
                    If bClassCached Then 'RDO Caching Implementation
                        'LogError(False, "Adding data to the cache for " & sComponentName & "." & sClassName, "ExecuteQuery", "")
                        AddDataToCache(sComponentName, sClassName, sQuery, oRecords)
                    End If
                End If
                If Not (oRecords Is Nothing) Then
                    rsRecords = ConvertDataTableToADORecordset(oRecords)
                    If rsRecords Is Nothing Then
                        lRows = 0
                    Else
                        lRows = rsRecords.RecordCount
                    End If

                Else
                    rsRecords = Nothing
                    lRows = 0
                End If
            End If

            If lRows = -1 Then
                lRows = 0
            End If

        Catch oError As ActiveBankException
            LogError(True, oError)
        Catch oError As System.Exception
            LogError(True, "abComponentServices", oError, sQuery)
        End Try

    End Sub

    Private Function ConvertDataTableToADORecordset( _
            ByRef oDataTable As Data.DataTable) As ADODB.RecordsetClass

        Dim rsRecords As ADODB.RecordsetClass
        Dim iFields As Integer
        Dim iField As Integer
        Dim iLength As Integer
        Dim eADODataType As ADODB.DataTypeEnum
        Dim sFieldName As String
        Dim iScale As Integer
        Dim iPrecision As Integer
        Dim iRows As Integer
        Dim iRow As Integer
        Dim oDataRow As Data.DataRow

        Try
            rsRecords = New ADODB.RecordsetClass
            iFields = oDataTable.Columns.Count - 1
            If iFields > -1 Then
                For iField = 0 To iFields
                    sFieldName = oDataTable.Columns(iField).ColumnName
                    eADODataType = GetADODataTypeFromDataColumn(oDataTable.Columns(iField), iLength, iPrecision, iScale)
                    If sFieldName.Length = 0 Then
                        sFieldName = "Field" & iField.ToString(System.Globalization.CultureInfo.InvariantCulture)
                    End If

                    Try
                        rsRecords.Fields.Append( _
                                sFieldName, _
                                eADODataType, _
                                iLength, _
                                ADODB.FieldAttributeEnum.adFldIsNullable)
                        rsRecords.Fields(sFieldName).NumericScale = CByte(iScale)
                        rsRecords.Fields(sFieldName).NumericScale = CByte(iPrecision)
                    Catch oWarning As System.Exception
                        LogError( _
                                False, _
                                oWarning.Message, _
                                "abComponentServices", _
                                oWarning.StackTrace, _
                                ErrorSeverity.ES_Warning)
                    End Try

                Next iField
                rsRecords.Open()

                iFields = rsRecords.Fields.Count - 1
                iRows = oDataTable.Rows.Count - 1
                For iRow = 0 To iRows
                    oDataRow = oDataTable.Rows(iRow)

                    rsRecords.AddNew()
                    For iField = 0 To iFields
                        sFieldName = rsRecords.Fields(iField).Name
                        If oDataRow.Item(iField) Is System.DBNull.Value Then
                            rsRecords.Fields(iField).Value() = System.DBNull.Value
                        ElseIf rsRecords.Fields(iField).Type = ADODB.DataTypeEnum.adVarWChar Then
                            rsRecords.Fields(iField).Value = oDataRow(iField).ToString
                        Else
                            If sFieldName.StartsWith("Field") Then
                                rsRecords.Fields(iField).Value = oDataRow(iField)
                            Else
                                rsRecords.Fields(iField).Value = oDataRow(sFieldName)
                            End If
                        End If
                    Next iField
                    rsRecords.Update()
                Next iRow
                If rsRecords.RecordCount > 0 Then
                    rsRecords.MoveFirst()
                End If
            Else
                rsRecords = Nothing
            End If

        Catch oError As System.Exception
            rsRecords = Nothing
            LogError(True, "abComponentServices", oError, "Fieldname: " + sFieldName)

        End Try

        Return rsRecords

    End Function

    Private Function ConvertDataTableToParameterCollection( _
        ByVal oDataTable As DataTable) As ParameterCollection

        Dim sFieldName As String
        Dim oParameterCollection As ParameterCollection
        Dim oResultRow As New ParameterCollection
        Dim iRow As Integer = 0
        Dim iColumn As Integer = 0
        Dim oColumn As DataColumn
        Dim oRow As DataRow

        Try
            For Each oRow In oDataTable.Rows

                oParameterCollection = New ParameterCollection
                iColumn = 0

                For Each oColumn In oDataTable.Columns
                    sFieldName = oColumn.ColumnName
                    oParameterCollection.Add(sFieldName, IIf(IsDBNull(oRow.Item(iColumn)), "", oRow.Item(iColumn)))
                    iColumn += 1
                Next
                oResultRow.Add(iRow, oParameterCollection)
                iRow += 1
            Next
        Catch oError As Exception
            LogError(True, "abComponentServices", oError, "Fieldname: " + sFieldName)
        End Try

        Return oResultRow

    End Function

    Private Function GetADODataTypeFromDataColumn( _
            ByRef oDataColumn As Data.DataColumn, _
            ByRef iLength As Integer, _
            ByRef iPrecision As Integer, _
            ByRef iScale As Integer) As ADODB.DataTypeEnum

        Dim eADODataType As ADODB.DataTypeEnum
        Dim sColumnName As String

        sColumnName = oDataColumn.DataType.FullName()
        Select Case sColumnName
            Case "System.Int32", "System.Int16"
                eADODataType = ADODB.DataTypeEnum.adInteger
                iLength = -1
                iScale = 0
                iPrecision = 0

            Case "System.Byte"
                eADODataType = ADODB.DataTypeEnum.adTinyInt
                iLength = -1
                iScale = 0
                iPrecision = 0

            Case "System.Boolean"
                eADODataType = ADODB.DataTypeEnum.adBoolean
                iLength = -1
                iScale = 0
                iPrecision = 0

            Case "System.DateTime"
                eADODataType = ADODB.DataTypeEnum.adDate
                iLength = -1
                iScale = 0
                iPrecision = 0

            Case "System.Double", "System.Decimal"
                eADODataType = ADODB.DataTypeEnum.adDouble
                iLength = -1
                iScale = 0
                iPrecision = 0

            Case "System.String"
                eADODataType = ADODB.DataTypeEnum.adVarWChar
                iLength = 255
                iScale = 0
                iPrecision = 0

            Case "System.Byte[]"
                eADODataType = ADODB.DataTypeEnum.adBinary
                iLength = -1
                iScale = 0
                iPrecision = 0

            Case "System.Guid"
                iLength = 50
                eADODataType = ADODB.DataTypeEnum.adVarWChar
                iLength = 50
                iScale = 0
                iPrecision = 0

            Case Else
                eADODataType = ADODB.DataTypeEnum.adVarWChar
                iLength = 255
                iScale = 0
                iPrecision = 0

        End Select

        Return eADODataType

    End Function

    Public Sub m_oRDOInterop_AbortTransaction(ByRef bSuccess As Boolean) 'Handles m_oRDOInterop.AbortTransaction

        Me.AbortTransaction()

        bSuccess = True

    End Sub

    Public Sub m_oRDOInterop_CommitTransaction(ByRef bSuccess As Boolean) 'Handles m_oRDOInterop.CommitTransaction

        Try
            If m_bIsInTransaction Then
                m_oDatabaseAdapter.CommitTransaction()
            End If
            bSuccess = True
        Catch oError As ActiveBankException
            LogError(True, oError)
        Catch oError As System.Exception
            LogError(True, "abComponentServices", oError, "CommitTransaction")
        End Try
    End Sub

    Public Sub m_oRDOInterop_StartTransaction(ByRef bSuccess As Boolean) 'Handles m_oRDOInterop.StartTransaction
        Try
            Me.StartTransaction()
            bSuccess = True
        Catch oError As ActiveBankException
            LogError(True, oError)
        Catch oError As System.Exception
            LogError(True, "abComponentServices", oError, "StartTransaction")
        End Try
    End Sub

    Public Sub m_oRDOInterop_GetErrorDetails(ByRef lNumber As Integer, ByRef sDescription As String) 'Handles m_oRDOInterop.GetErrorDetails

        lNumber = 0
        sDescription = ""

    End Sub

    Public Sub m_oRDOInterop_ExecuteProcedure(ByVal sCommand As String, ByRef rsRecords As ADODB.Recordset, ByRef bSuccess As Boolean) 'Handles m_oRDOInterop.ExecuteProcedure
        Dim oDataTable As Data.DataTable

        Try
            oDataTable = m_oDatabaseAdapter.ExecuteQueryDataTable(sCommand)

            rsRecords = ConvertDataTableToADORecordset(oDataTable)

            bSuccess = True

        Catch oError As ActiveBankException
            LogError(True, oError)
        Catch oError As System.Exception
            LogError(True, "abComponentServices", oError, sCommand)
        End Try
    End Sub


    Public Sub m_oRDOInterop_ExecuteStoredProcedure(ByVal sStoredProcedureName As String, ByRef rsRecords As ADODB.Recordset, ByRef bSuccess As Boolean, ByRef vParameters As Object)
        '======================================================================
        ' Date Created : 27 February 2008 
        '----------------------------------------------------------------------
        ' About...     : Execute stored procedure using variant array of parameters
        '======================================================================
        ' Change Log
        '
        ' Date       Author         Ref                         Comments
        ' ---------- ------------   ---------                   --------
        ' 27/02/2008 Jat Virdee     TD24359, TD24357, TD24426   Created
        '======================================================================
        Dim oDataTable As Data.DataTable
        Dim lCounter As Long
        Dim alParameterValues As New System.Collections.ArrayList
        Try
            If Not vParameters Is Nothing Then
                For lCounter = 0 To UBound(vParameters)
                    alParameterValues.Add(vParameters(lCounter))
                Next
            End If

            oDataTable = m_oDatabaseAdapter.ExecuteStoredProcedureParameters(sStoredProcedureName, alParameterValues)

            rsRecords = ConvertDataTableToADORecordset(oDataTable)
            bSuccess = True

        Catch oError As ActiveBankException
            LogError(True, oError)
        Catch oError As System.Exception
            LogError(True, "abComponentServices", oError, sStoredProcedureName)
        End Try

    End Sub

    Public Sub m_oRDOInterop_GetDatabaseType(ByRef eDatabaseType As abRDOInterop.DatabaseType) 'Handles m_oRDOInterop.GetDatabaseType

        Select Case m_structConfigSettings.eDatabaseVendor
            Case DatabaseVendor.SQLServer
                eDatabaseType = abRDOInterop.DatabaseType.SQLServer

            Case DatabaseVendor.Oracle
                eDatabaseType = abRDOInterop.DatabaseType.Oracle

            Case DatabaseVendor.ADS
                eDatabaseType = abRDOInterop.DatabaseType.ADS

        End Select

    End Sub
    'Manoj Meged From DBS - Create .NET component instance from vb 
    Public Sub m_oRDOInterop_CreateComponentInstance(ByVal sComponentName As String, ByVal sClassName As String, ByRef oComponent As Object, ByVal sSession As String, ByVal sBranch As String, ByRef bSuccess As Boolean) 'Handles m_oRDOInterop.CreateComponentInstance
        Dim oClassicComponent As IABSupportsClassicComponent
        Try
            oComponent = CreateComponentInstance(sComponentName, sClassName)
        m_sSessionID = sSession
        m_sBranchID = sBranch
            oClassicComponent = oComponent
            oClassicComponent.Initialize(sSession, sBranch, Me)
            bSuccess = True
        Catch oError As ActiveBankException
            LogError(True, oError)
        Catch oError As System.Exception
            LogError(True, "abComponentServices", oError, sComponentName & "_" & sClassName)
        End Try

    End Sub

    Private Shared Sub ReleaseCOMObject(ByVal oCOMObject As Object)

        Dim iReferenceCount As Integer

        'Check added by RDT - 07/10/2004
        If Not (oCOMObject Is Nothing) Then
            Try
                iReferenceCount = -1
                If Not (oCOMObject Is Nothing) Then
                    Do
                        iReferenceCount = System.Runtime.InteropServices.Marshal.ReleaseComObject(oCOMObject)
                    Loop Until iReferenceCount <= 0
                End If
            Catch oError As System.Exception
                LogError(False, "abComponentServices", oError)
            End Try
        End If

    End Sub

    Private Sub ReleaseClassicInteropHandles()

        Dim oSession As Object

        If Not (m_oUser Is Nothing) Then
            Try
                oSession = CallByName(m_oUser, "SessionHandle", CallType.Get)
            Catch oError As System.Exception
                WriteToLog("abComponentServices", "ReleaseClassicInteropHandles", EventLogEntryType.Warning, oError.StackTrace, 0, oError.Message, _
                    "Get User.SessionHandle - Thread:" & Threading.Thread.CurrentThread.GetHashCode().ToString)
            End Try

            If Not oSession Is Nothing Then
                Try
                    CallByName(oSession, "DBHandle", CallType.Let, CType(Nothing, abRDOInterop.RDOClassClass))
                Catch oError As System.Exception
                    WriteToLog("abComponentServices", "ReleaseClassicInteropHandles", EventLogEntryType.Warning, oError.StackTrace, 0, oError.Message, _
                        "Set Session.DBHandle to Nothing - Thread:" & Threading.Thread.CurrentThread.GetHashCode().ToString)
                End Try

                Try
                    CallByName(oSession, "User", CallType.Let, CType(Nothing, abRDOInterop.RDOClassClass))
                Catch oError As Exception
                    WriteToLog("abComponentServices", "ReleaseClassicInteropHandles", EventLogEntryType.Warning, oError.StackTrace, 0, oError.Message, _
                        "Set Session.User to Nothing - Thread:" & Threading.Thread.CurrentThread.GetHashCode().ToString)
                End Try
            End If

            Try
                CallByName(m_oUser, "SessionHandle", CallType.Let, CType(Nothing, abRDOInterop.RDOClassClass))
            Catch oError As Exception
                WriteToLog("abComponentServices", "ReleaseClassicInteropHandles", EventLogEntryType.Warning, oError.StackTrace, 0, oError.Message, _
                    "Set User.SessionHandle to Nothing - Thread:" & Threading.Thread.CurrentThread.GetHashCode().ToString)
            End Try

            'Changed by RDT - 07/10/2004
            ReleaseClassicComponent(oSession)

            Try
                CallByName(m_oUser, "DBHandle", CallType.Let, CType(Nothing, abRDOInterop.RDOClassClass))
            Catch oError As Exception
                WriteToLog("abComponentServices", "ReleaseClassicInteropHandles", EventLogEntryType.Warning, oError.StackTrace, 0, oError.Message, _
                    "Set User.DBHandle to Nothing - Thread:" & Threading.Thread.CurrentThread.GetHashCode().ToString)
            End Try
        End If
        'Changed by RDT - 07/10/2004
        ReleaseClassicComponent(m_oAuditTrail, "Internal.m_oAuditTrail")
        ReleaseClassicComponent(m_oRDOInterop, "Internal.m_oRDOInterop")
        ReleaseClassicComponent(m_oUser, "Internal.m_oUser")
        ReleaseClassicComponent(m_oBroker, "Internal.m_oBroker")

    End Sub

    Public Shared Sub ReleaseClassicComponent(ByVal oComponent As Object)
        ReleaseClassicComponent(oComponent, "")
    End Sub

    Public Shared Sub ReleaseClassicComponent(ByVal oComponent As Object, ByVal sObjName As String)
        Dim oDummy As Object = Nothing

        If sObjName.Length > 0 Then
            sObjName = sObjName & "::ReleaseClassicComponent."
        End If

        'Check added by RDT - 07/10/2004
        If Not (oComponent Is Nothing) Then

            Try
                CallByName(oComponent, "Term", CallType.Method, oDummy)
            Catch oError As System.Exception
                WriteToLog("abComponentServices", "ReleaseClassicComponent", EventLogEntryType.Warning, oError.StackTrace, 0, oError.Message, _
                    sObjName & "Term - Thread:" & Threading.Thread.CurrentThread.GetHashCode().ToString)
            End Try

            Try
                CallByName(oComponent, "Broker", CallType.Let, oDummy)
            Catch oError As System.Exception
                WriteToLog("abComponentServices", "ReleaseClassicComponent", EventLogEntryType.Warning, oError.StackTrace, 0, oError.Message, _
                    sObjName & "Broker - Thread:" & Threading.Thread.CurrentThread.GetHashCode().ToString)
            End Try

            Try
                CallByName(oComponent, "DBHandle", CallType.Let, oDummy)
            Catch oError As System.Exception
                WriteToLog("abComponentServices", "ReleaseClassicComponent", EventLogEntryType.Warning, oError.StackTrace, 0, oError.Message, _
                    sObjName & "DBHandle - Thread:" & Threading.Thread.CurrentThread.GetHashCode().ToString)
            End Try

            Try
                CallByName(oComponent, "User", CallType.Let, oDummy)
            Catch oError As System.Exception
                WriteToLog("abComponentServices", "ReleaseClassicComponent", EventLogEntryType.Warning, oError.StackTrace, 0, oError.Message, _
                    sObjName & "User - Thread:" & Threading.Thread.CurrentThread.GetHashCode().ToString)
            End Try

            Try
                CallByName(oComponent, "AuditTrail", CallType.Let, oDummy)
            Catch oError As System.Exception
                WriteToLog("abComponentServices", "ReleaseClassicComponent", EventLogEntryType.Warning, oError.StackTrace, 0, oError.Message, _
                    sObjName & "AuditTrail - Thread:" & Threading.Thread.CurrentThread.GetHashCode().ToString)
            End Try

            ReleaseCOMObject(oComponent)

        End If

        oComponent = Nothing


    End Sub

    Public Function comGetConfigSettings() As String

        Dim sSettings As String
        Dim docSettings As Xml.XmlDocument

        docSettings = CreateXMLDocument("ConfigSettings", _
                "LoggingLevel=" & GetISODataValue(m_structConfigSettings.eLoggingLevel), _
                "CachingEnabled=" & GetISODataValue(m_structConfigSettings.bCachingEnabled), _
                "Timeout=" & GetISODataValue(m_structConfigSettings.iDatabaseTimeout), _
                "Name=" & m_structConfigSettings.sDatabaseName, _
                "Password=" & m_structConfigSettings.sDatabasePassword, _
                "Server=" & m_structConfigSettings.sDatabaseServer, _
                "UserID=" & m_structConfigSettings.sDatabaseUserName)

        Return docSettings.OuterXml

    End Function

    Public Sub CallWebServiceFromCOM( _
            ByVal sInput As String, _
            ByRef sOutput As String, _
            ByVal sComponentName As String, _
            ByVal sClassName As String, _
            ByVal sWebService As String, _
            ByRef oComponentServices As abComponentServices.ComponentServices)

        Dim ndInput As Xml.XmlNode
        Dim docOutput As Xml.XmlDocument
        Dim oComponent As Object

        ndInput = CreateXMLDocument(sInput).FirstChild
        oComponent = CreateComponentInstance(sComponentName, sClassName, oComponentServices)
        docOutput = CallByName(oComponent, sWebService, CallType.Method, ndInput)
        oComponent = Nothing

        sOutput = docOutput.OuterXml

    End Sub

#End Region

#Region "Constructors And Destructors"

    Public Sub New()

        'Necessary for COM interop.

        If (m_oStoredProcedureDefinitions Is Nothing) Then
            m_oStoredProcedureDefinitions = New Collections.Specialized.HybridDictionary
            Debug.Write("CS - Stored Procedure Cache Initialised")
        End If

    End Sub

    'Private Sub New( _
    '        ByVal sDatabaseIdentifier As String)

    '    Me.New(sDatabaseIdentifier, False)

    'End Sub

    'Private Sub New( _
    '        ByVal bRequiresTransaction As Boolean)

    '    Me.New("", bRequiresTransaction)

    'End Sub

    'Private Sub New( _
    '        ByVal sDatabaseIdentifier As String, _
    '        ByVal bRequiresTransaction As Boolean)

    '    Me.ActivateComponentServices(sDatabaseIdentifier, bRequiresTransaction)

    'End Sub

    Public Sub ActivateComponentServices( _
            ByVal sDatabaseIdentifier As String, _
            ByVal bRequiresTransaction As Boolean)

        Me.ActivateComponentServices(sDatabaseIdentifier, bRequiresTransaction, False)

    End Sub

    Public Sub ActivateComponentServices( _
        ByVal sDatabaseIdentifier As String, _
        ByVal bRequiresTransaction As Boolean, _
        ByVal bZeroiseRefCount As Boolean)

        Dim oClassesRoot As Win32.RegistryKey
        Dim oDatabaseInfo As Win32.RegistryKey
        Dim oConfigSettings As ParameterCollection
        Dim oActiveBank As Win32.RegistryKey
        Dim oEncryption As abComponentServices.Encryption

        Try

            If Not m_structConfigSettings.bConfigSettingsSet OrElse (UCase(sDatabaseIdentifier) <> m_structConfigSettings.sDatabaseIdentifier) Then
                oClassesRoot = Win32.RegistryKey.OpenRemoteBaseKey(Win32.RegistryHive.LocalMachine, "")
                oActiveBank = oClassesRoot.OpenSubKey("SOFTWARE\Financial Objects\ActiveBank")

                If sDatabaseIdentifier Is Nothing Then
                    sDatabaseIdentifier = DEFAULT_DB_IDENTIFIER
                Else
                    If sDatabaseIdentifier.Length = 0 Then
                        sDatabaseIdentifier = DEFAULT_DB_IDENTIFIER
                    Else
                        sDatabaseIdentifier = DEFAULT_DB_IDENTIFIER & "\" & sDatabaseIdentifier
                    End If
                End If

                oDatabaseInfo = oActiveBank.OpenSubKey(sDatabaseIdentifier)
                If oDatabaseInfo Is Nothing Then
                    Throw New ActiveBankException(0, "The database setting [" & sDatabaseIdentifier & "] cannot be found.")
                End If

                If CStr(oDatabaseInfo.GetValue("Encrypted", "0")) = "1" Then
                    m_structConfigSettings.bDBSettingsEncrypted = True
                Else
                    m_structConfigSettings.bDBSettingsEncrypted = False
                End If

                m_structConfigSettings.sDatabaseName = DecryptDBSetting(CStr(oDatabaseInfo.GetValue("Name", "activebank")))
                m_structConfigSettings.eDatabaseVendor = CType(DecryptDBSetting(CStr(oDatabaseInfo.GetValue("Vendor", ComponentServices.DatabaseVendor.SQLServer))), abComponentServices.ComponentServices.DatabaseVendor)
                m_structConfigSettings.sDatabaseServer = DecryptDBSetting(CStr(oDatabaseInfo.GetValue("Server", ".")))
                m_structConfigSettings.sDatabaseUserName = DecryptDBSetting(CStr(oDatabaseInfo.GetValue("UserID", "sa")))
                m_structConfigSettings.sDatabasePassword = DecryptDBSetting(CStr(oDatabaseInfo.GetValue("Password", "")))
                m_structConfigSettings.sDatabaseIdentifier = sDatabaseIdentifier
                'm_structConfigSettings.bAutoCreateEnumerator = CBool(oDatabaseInfo.GetValue("AutoCreateEnumerator", "0"))

                Select Case m_structConfigSettings.eDatabaseVendor
                    Case DatabaseVendor.Oracle
                        m_oDatabaseAdapter = New OracleAdapter

                    Case DatabaseVendor.SQLServer
                        m_oDatabaseAdapter = New SQLServerAdapter

                    Case DatabaseVendor.ADS
                        m_oDatabaseAdapter = New ADSAdapter

                End Select

                If m_structConfigSettings.eDatabaseVendor = DatabaseVendor.ADS Then

                    m_oDatabaseAdapter.OpenConnection( _
                            m_structConfigSettings.sDatabaseServer, _
                            m_structConfigSettings.sDatabaseName, _
                            m_structConfigSettings.sDatabaseUserName, _
                            m_structConfigSettings.sDatabasePassword, _
                            30)
                    m_structConfigSettings.bConfigSettingsSet = False
                Else

                    m_oDatabaseAdapter.OpenConnection( _
                            m_structConfigSettings.sDatabaseServer, _
                            m_structConfigSettings.sDatabaseName, _
                            m_structConfigSettings.sDatabaseUserName, _
                            m_structConfigSettings.sDatabasePassword, _
                            30)

                    CacheObjectAliases()

                    'Connection is Successful for passed Database Identifier, so use default one to get Config Settings
                    GetConfigSettings(DEFAULT_DB_IDENTIFIER)

                    m_oDatabaseAdapter.OpenConnection( _
                            m_structConfigSettings.sDatabaseServer, _
                            m_structConfigSettings.sDatabaseName, _
                            m_structConfigSettings.sDatabaseUserName, _
                            m_structConfigSettings.sDatabasePassword, _
                            m_structConfigSettings.iDatabaseTimeout)
                End If

            Else

                Select Case m_structConfigSettings.eDatabaseVendor
                    Case DatabaseVendor.Oracle
                        m_oDatabaseAdapter = New OracleAdapter

                    Case DatabaseVendor.SQLServer
                        m_oDatabaseAdapter = New SQLServerAdapter

                    Case DatabaseVendor.ADS
                        m_oDatabaseAdapter = New ADSAdapter

                End Select

                If m_structConfigSettings.eDatabaseVendor = DatabaseVendor.ADS Then
                    m_oDatabaseAdapter.OpenConnection( _
                            m_structConfigSettings.sDatabaseServer, _
                            m_structConfigSettings.sDatabaseName, _
                            m_structConfigSettings.sDatabaseUserName, _
                            m_structConfigSettings.sDatabasePassword, _
                            30)
                Else
                    m_oDatabaseAdapter.OpenConnection( _
                            m_structConfigSettings.sDatabaseServer, _
                            m_structConfigSettings.sDatabaseName, _
                            m_structConfigSettings.sDatabaseUserName, _
                            m_structConfigSettings.sDatabasePassword, _
                            m_structConfigSettings.iDatabaseTimeout)
                End If

                CacheObjectAliases()

            End If

            If bZeroiseRefCount Then
                m_iReferenceCount = 0
            Else
                m_iReferenceCount = 1
            End If

            m_iCOMReferenceCount = 0

            If bRequiresTransaction Then
                StartTransaction()
            Else
                m_bIsInTransaction = False
            End If

        Catch oError As ActiveBankException
            LogError(True, oError)
        Catch oError As System.Exception
            LogError(True, "abComponentServices", oError)
        Finally

        End Try

    End Sub

    Private Function CacheObjectAliases() _
        As Boolean
        '==================================================================================================
        ' Author    : Darryn Clerihew
        '--------------------------------------------------------------------------------------------------
        ' About...  : Loads any database object aliases (short names) to a cache for performance reasons. 
        '             The purpose of this cache is to return the short name of any given object when a query
        '             is executed.
        '
        '             Currently this cache will only be used by the Oracle database adapter.
        '
        '             The cache is implemented using a hashtable where:
        '             Key = The existing (long) object name
        '             Value = The alias (short) object name
        '==================================================================================================

        Dim sQuery As String
        Dim pcAlias As ParameterCollection
        Dim pcWorkElement As ParameterCollection
        Dim sObjectName As String
        Dim sObjectAlias As String
        Dim bTableMissing As Boolean

        'Caching the object aliases only applies to Oracle
        If m_structConfigSettings.eDatabaseVendor = DatabaseVendor.Oracle Then

            'Only load the object aliases if the cache is empty
            If (m_htAliasCache Is Nothing) Then

                m_htAliasCache = New Hashtable

                sQuery = "SELECT ObjectName, ObjectAlias FROM CS_AliasTbl"

                Try

                    pcAlias = m_oDatabaseAdapter.ExecuteQueryParameterCollection( _
                        sQuery, False, False)

                Catch oError As ActiveBankException
                    If oError.Number = ErrorNumbers.OracleObjectMissing OrElse oError.Number = ErrorNumbers.SQLObjectMissing Then
                        bTableMissing = True
                    End If

                End Try

                If Not bTableMissing Then

                    pcAlias.Reset()

                    Do While pcAlias.MoveNext
                        pcWorkElement = pcAlias.GetParameterCollection()
                        sObjectName = pcWorkElement.GetStringValue("ObjectName")
                        sObjectAlias = pcWorkElement.GetStringValue("ObjectAlias")

                        m_htAliasCache.Add(sObjectName.Trim, sObjectAlias.Trim)
                    Loop

                End If
            End If
        End If

        Return True

    End Function

    Private Function GetObjectAlias(ByVal sObjectName As String) _
        As String
        '==================================================================================================
        ' Author    : Darryn Clerihew
        '--------------------------------------------------------------------------------------------------
        ' About...  : Retrieves the specified alias (short name) from the cache based on the supplied object
        '             name.
        ' Parameters: sObjectName - The identifier (long name) for the cache item to retrieve.
        '==================================================================================================

        Dim sAlias As String = String.Empty

        'Cached aliases currently only apply to Oracle
        If m_structConfigSettings.eDatabaseVendor = DatabaseVendor.Oracle Then

            'Determine whether the cache contains the supplied object name before attempting to retrieve
            'the alias
            If m_htAliasCache.ContainsKey(sObjectName.ToUpper.Trim) Then
                sAlias = m_htAliasCache.Item(sObjectName.ToUpper.Trim)
            Else
                sAlias = sObjectName.Trim
            End If

        Else
            sAlias = sObjectName.Trim
        End If

        Return sAlias

    End Function

    Public Sub DisposeComponentServices()

        If Not m_oDatabaseAdapter Is Nothing Then
            If m_bIsInTransaction Then
                If m_bTransactionAborted Then
                    m_oDatabaseAdapter.AbortTransaction()
                Else
                    m_oDatabaseAdapter.CommitTransaction()
                End If
            End If

            m_oDatabaseAdapter.CloseConnection()
            m_oDatabaseAdapter = Nothing
            'Merged from DBS 
            ReleaseClassicInteropHandles()
        End If

    End Sub

    Private Function DecryptDBSetting(ByVal sEncryptedSetting As String) As String

        Dim sDecryptedSetting As String
        Dim oEncryption As abComponentServices.Encryption

        If m_structConfigSettings.bDBSettingsEncrypted Then
            oEncryption = New abComponentServices.Encryption
            oEncryption.sKey = "baddog"

            oEncryption.sInBufferHex = sEncryptedSetting

            If oEncryption.bCryptoDecrypt Then
                sDecryptedSetting = oEncryption.sOutBuffer
            Else
                sDecryptedSetting = sEncryptedSetting
            End If
        Else
            sDecryptedSetting = sEncryptedSetting
        End If

        Return sDecryptedSetting

    End Function

    Private Function GetConfigSettings(ByVal sDatabaseIdentifier As String) As Boolean
        '==================================================================================================
        ' Author    : Arvind Tulsiram
        '--------------------------------------------------------------------------------------------------
        ' About...  : Loads/gets the component services config settings for the given db identifier from 
        '             the abComponentServices_ConfigTbl.         
        '==================================================================================================
        Dim sbQuery As StringBuilder
        Dim oConfigSettings As ParameterCollection
        Dim bTableMissing As Boolean

        If m_structConfigSettings.eLoggingLevel = LoggingLevel.None Then
            m_structConfigSettings.eLoggingLevel = LoggingLevel.ErrorsOnly
        End If

        sbQuery = New StringBuilder
        sbQuery.Append("SELECT * FROM ")
        sbQuery.Append(GetObjectAlias("abComponentServices_ConfigTbl"))
        sbQuery.Append(" WHERE Name = '")
        sbQuery.Append(sDatabaseIdentifier)
        sbQuery.Append("'")

        Try

            oConfigSettings = m_oDatabaseAdapter.ExecuteQueryParameterCollection(sbQuery.ToString, True, False)

        Catch oError As ActiveBankException
            If oError.Number = ErrorNumbers.OracleObjectMissing OrElse oError.Number = ErrorNumbers.SQLObjectMissing Then
                bTableMissing = True
            Else
                LogError(False, oError)
            End If
        Catch oError As System.Exception
            LogError(False, "abComponentServices:GetConfigSettings", oError)
        End Try

        m_oDatabaseAdapter.CloseConnection()

        'Reload the config settings only if the CS config table exists in the target db
        If Not bTableMissing Then
            If oConfigSettings Is Nothing Then
                Throw New ActiveBankException(0, "The config settings for database setting [" & sDatabaseIdentifier & "] cannot be found.")
            End If

            m_structConfigSettings.iDatabaseTimeout = oConfigSettings.GetIntegerValue("DatabaseTimeout")
            m_structConfigSettings.bCachingEnabled = oConfigSettings.GetBooleanValue("CachingEnabled", False)
            m_structConfigSettings.eLoggingLevel = CType(oConfigSettings.GetIntegerValue("LoggingLevel"), abComponentServices.ComponentServices.LoggingLevel)

            m_structConfigSettings.sApplicationPath = GetPathFromFileNamePath(System.Reflection.Assembly.GetExecutingAssembly().Location)
            m_structConfigSettings.sApplicationServer = System.Windows.Forms.SystemInformation.ComputerName

            If m_structConfigSettings.bCachingEnabled And Not m_structConfigSettings.bConfigSettingsSet Then
                m_oCachedParameterCollections = New Collections.Specialized.HybridDictionary
                m_oCachedDataTables = New Collections.Specialized.HybridDictionary
                m_oCachedXMLDocuments = New Collections.Specialized.HybridDictionary
            End If

            m_structConfigSettings.bConfigSettingsSet = True
        End If

        Return True

    End Function

#End Region

#Region "Create Methods"

    Public Function Create( _
            ByVal sComponentName As String, _
            ByVal sClassName As String, _
            ByRef oAttributes As ParameterCollection) As Decimal

        Return Create( _
                sComponentName, _
                sClassName, _
                oAttributes, _
                False)

    End Function

    Public Function Create( _
            ByVal sComponentName As String, _
            ByVal sClassName As String, _
            ByRef oAttributes As ParameterCollection, _
            ByVal bClassCached As Boolean) As Decimal

        Dim sStoredProcedureName As String
        Dim pcResult As ParameterCollection
        Dim decID As Decimal

        sStoredProcedureName = GetCreateStoredProcedureName(sComponentName, sClassName)

        Call AppendMaintenanceAttributes(oAttributes, True)

        pcResult = m_oDatabaseAdapter.ExecuteStoredProcedure( _
                sStoredProcedureName, _
                oAttributes, _
                Nothing)

        If pcResult Is Nothing Then
            decID = Nothing
        Else
            decID = pcResult.GetDecimalValue("ID", 0)

            If bClassCached Then
                InformApplicationServersOfDataChange( _
                        sComponentName, _
                        sClassName)
            End If
        End If

        Return decID

    End Function

    Public Function Create( _
            ByVal sComponentName As String, _
            ByVal sClassName As String, _
            ByRef ndAttributes As Xml.XmlNode) As Xml.XmlDocument

        Return Create( _
                sComponentName, _
                sClassName, _
                ndAttributes, _
                False)

    End Function

    Public Function Create( _
            ByVal sComponentName As String, _
            ByVal sClassName As String, _
            ByRef ndAttributes As Xml.XmlNode, _
            ByVal bClassCached As Boolean) As Xml.XmlDocument

        Return Create( _
                sComponentName, _
                sClassName, _
                ndAttributes, _
                bClassCached, _
                Nothing)

    End Function

    Public Function Create( _
            ByVal sComponentName As String, _
            ByVal sClassName As String, _
            ByRef ndAttributes As Xml.XmlNode, _
            ByRef ndClassHierachy As Xml.XmlNode) As Xml.XmlDocument

        Return Create( _
                sComponentName, _
                sClassName, _
                ndAttributes, _
                False, _
                ndClassHierachy)

    End Function

    Public Function Create( _
            ByVal sComponentName As String, _
            ByVal sClassName As String, _
            ByRef ndAttributes As Xml.XmlNode, _
            ByVal bClassCached As Boolean, _
            ByRef ndClassHierachy As Xml.XmlNode) As Xml.XmlDocument

        Return Create( _
                sComponentName, _
                sClassName, _
                ndAttributes, _
                bClassCached, _
                ndClassHierachy, _
                "")

    End Function

    Public Function Create( _
            ByVal sComponentName As String, _
            ByVal sClassName As String, _
            ByRef ndAttributes As Xml.XmlNode, _
            ByVal bClassCached As Boolean, _
            ByRef ndClassHierachy As Xml.XmlNode, _
            ByRef sResultRootName As String) As Xml.XmlDocument

        Dim docResultParent As Xml.XmlDocument
        Dim ndResultParent As Xml.XmlNode

        'AFS Create result document.
        If sResultRootName Is Nothing Then
            sResultRootName = GetDefaultXMLResponseRootName(sClassName, "Create")
        Else
            If sResultRootName.Length = 0 Then
                sResultRootName = GetDefaultXMLResponseRootName(sClassName, "Create")
            End If
        End If
        docResultParent = CreateXMLDocument(sResultRootName, Nothing)
        ndResultParent = docResultParent.SelectSingleNode(sResultRootName)

        Me.Create( _
                sComponentName, _
                sClassName, _
                ndAttributes, _
                bClassCached, _
                ndClassHierachy, _
                ndResultParent)

        Return docResultParent

    End Function

    Public Sub Create( _
            ByVal sComponentName As String, _
            ByVal sClassName As String, _
            ByRef ndAttributes As Xml.XmlNode, _
            ByVal bClassCached As Boolean, _
            ByRef ndClassHierachy As Xml.XmlNode, _
            ByRef ndResultParent As Xml.XmlNode)

        Dim sStoredProcedureName As String
        Dim pcDataTypes As ParameterCollection
        Dim decID As Decimal
        Dim pcParameters As ParameterCollection
        Dim iSubClassDefinition As Integer
        Dim iSubClass As Integer
        Dim iSubClassDefinitions As Integer
        Dim iSubClasses As Integer
        Dim ndlSubClasses As Xml.XmlNodeList
        Dim ndlSubClassDefinitions As Xml.XmlNodeList
        Dim ndSubClassDefinition As Xml.XmlElement
        Dim ndSubClass As Xml.XmlNode
        Dim sSubClassDefinitionName As String
        Dim sSubClassName As String
        Dim sRelationalIDName As String
        Dim sbRelationalIDName As StringBuilder

        sStoredProcedureName = GetCreateStoredProcedureName(sComponentName, sClassName)
        pcDataTypes = m_oDatabaseAdapter.GetStoredProcedureDataTypes(sStoredProcedureName)
        pcParameters = PopulateParameterCollectionFromXMLNode(ndAttributes, sComponentName, sClassName, pcDataTypes)

        decID = Me.Create(sComponentName, sClassName, pcParameters, bClassCached)

        AppendIDNode(ndResultParent, "ID", decID, Nothing)

        Try
            'AFS Process sub classes.
            If Not (ndClassHierachy Is Nothing) Then
                ndlSubClassDefinitions = ndClassHierachy.SelectNodes("*")
                iSubClassDefinitions = ndlSubClassDefinitions.Count - 1
                For iSubClassDefinition = 0 To iSubClassDefinitions
                    ndSubClassDefinition = CType(ndlSubClassDefinitions.Item(iSubClassDefinition), Xml.XmlElement)
                    sSubClassDefinitionName = ndSubClassDefinition.Name
                    sSubClassName = GetStringXMLAttributeData("TagName", ndSubClassDefinition, sSubClassDefinitionName)
                    sbRelationalIDName = New StringBuilder(sClassName)
                    sbRelationalIDName.Append("ID")
                    sRelationalIDName = sbRelationalIDName.ToString()

                    'AFS Get all subclasses from ndAttributes.
                    ndlSubClasses = ndAttributes.SelectNodes(sSubClassName)
                    iSubClasses = ndlSubClasses.Count - 1
                    For iSubClass = 0 To iSubClasses
                        ndSubClass = ndlSubClasses.Item(iSubClass)
                        SetXMLElementData(sRelationalIDName, decID, ndSubClass)
                        Me.Create( _
                                sComponentName, _
                                sSubClassDefinitionName, _
                                ndSubClass, _
                                bClassCached, _
                                CType(ndSubClassDefinition, Xml.XmlNode))
                    Next iSubClass
                Next iSubClassDefinition
            End If
        Catch oError As System.Exception
            LogError(True, "abComponentServices", oError)
        End Try

    End Sub

    Private Function GetCreateStoredProcedureName( _
            ByVal sComponentName As String, _
            ByVal sClassName As String) As String

        Dim sbStoredProcedureName As New StringBuilder(sComponentName)

        sbStoredProcedureName.Append("_")
        sbStoredProcedureName.Append(sClassName)
        sbStoredProcedureName.Append("_C")

        Return GetObjectAlias(sbStoredProcedureName.ToString())

    End Function

    Private Sub AppendMaintenanceAttributes( _
            ByRef pcAttributes As ParameterCollection, _
            ByVal bCreate As Boolean)

        Dim oAttribute As Object

        If bCreate Then
            oAttribute = pcAttributes.GetValue("CreateUserID")
            If oAttribute Is Nothing Then
                Call pcAttributes.Add("CreateUserID", m_decUserID)
            End If

            oAttribute = pcAttributes.GetValue("DateCreated")
            If oAttribute Is Nothing Then
                Call pcAttributes.Add("DateCreated", Now)
            End If
        Else
            oAttribute = pcAttributes.GetValue("UpdateUserID")
            If oAttribute Is Nothing Then
                Call pcAttributes.Add("UpdateUserID", m_decUserID)
            End If

            oAttribute = pcAttributes.GetValue("DateUpdated")
            If oAttribute Is Nothing Then
                Call pcAttributes.Add("DateUpdated", Now)
            End If
        End If

        Call pcAttributes.Remove("ObjectStatus")
        Call pcAttributes.Add("ObjectStatus", ObjectStatus.Active)

    End Sub

#End Region

#Region "Update Methods"

    Public Function Update( _
            ByVal sComponentName As String, _
            ByVal sClassName As String, _
            ByVal decID As Decimal, _
            ByVal sTimestamp As String, _
            ByVal oAttributes As ParameterCollection) As String

        Return Update( _
                sComponentName, _
                sClassName, _
                decID, _
                sTimestamp, _
                oAttributes, _
                False).ToString(System.Globalization.CultureInfo.InvariantCulture)

    End Function

    Public Function Update( _
            ByVal sComponentName As String, _
            ByVal sClassName As String, _
            ByVal decID As Decimal, _
            ByVal sTimestamp As String, _
            ByVal pcAttributes As ParameterCollection, _
            ByVal bClassCached As Boolean) As Decimal

        Dim pcCurrentValues As ParameterCollection
        Dim sStoredProcedureName As String
        Dim decResult As Decimal

        sStoredProcedureName = GetUpdateStoredProcedureName(sComponentName, sClassName)

        pcCurrentValues = GetInstanceParameterCollection( _
                sComponentName, _
                sClassName, _
                CStr(decID))

        If pcCurrentValues Is Nothing Then
            Throw New ActiveBankException( _
                    ErrorNumbers.TheInstanceHasNotBeenFound, _
                    "The specified instance cannot be found.", _
                    ActiveBankException.ExceptionType.Business, _
                    "abComponentServices")
        Else
            If CheckTimestamp(pcCurrentValues.GetStringValue("Timestamp"), sTimestamp) Then
                pcAttributes.Add("ID", decID)

                Call AppendMaintenanceAttributes(pcAttributes, False)

                m_oDatabaseAdapter.ExecuteStoredProcedure( _
                        sStoredProcedureName, _
                        pcAttributes, _
                        pcCurrentValues)

                decResult = decID

                If bClassCached Then
                    InformApplicationServersOfDataChange(sComponentName, sClassName)
                End If
            Else
                Throw New ActiveBankException( _
                        ErrorNumbers.TheTimestampIsIoutOfDate, _
                        "The timestamp supplied does not match the current timestamp.", _
                        ActiveBankException.ExceptionType.Business, _
                        "abComponentServices")
            End If
        End If

        Return decID

    End Function

    Public Function Update( _
            ByVal sComponentName As String, _
            ByVal sClassName As String, _
            ByRef ndAttributes As Xml.XmlNode) As Xml.XmlDocument

        Return Update(sComponentName, sClassName, ndAttributes, False)

    End Function

    Public Function Update( _
            ByVal sComponentName As String, _
            ByVal sClassName As String, _
            ByRef ndAttributes As Xml.XmlNode, _
            ByVal bClassCached As Boolean) As Xml.XmlDocument

        Return Update( _
                sComponentName, _
                sClassName, _
                ndAttributes, _
                bClassCached, _
                Nothing)

    End Function

    Public Function Update( _
            ByVal sComponentName As String, _
            ByVal sClassName As String, _
            ByRef ndAttributes As Xml.XmlNode, _
            ByVal bClassCached As Boolean, _
            ByRef ndClassHierachy As Xml.XmlNode) As Xml.XmlDocument

        Return Update( _
                sComponentName, _
                sClassName, _
                ndAttributes, _
                bClassCached, _
                ndClassHierachy, _
                False)
    End Function

    Public Function Update( _
            ByVal sComponentName As String, _
            ByVal sClassName As String, _
            ByRef ndAttributes As Xml.XmlNode, _
            ByVal bClassCached As Boolean, _
            ByRef ndClassHierachy As Xml.XmlNode, _
            ByVal bKeepHistory As Boolean) As Xml.XmlDocument

        Dim ndID As Xml.XmlNode
        Dim decID As Decimal
        Dim sTimestamp As String
        Dim docOutput As Xml.XmlDocument

        ndID = ndAttributes.SelectSingleNode("ID")
        If ndID Is Nothing Then
            docOutput = CreateErrorDocument(-1, "The ID node is missing.")
        Else
            GetValuesFromIDNode( _
                    CType(ndID, Xml.XmlElement), _
                    decID, _
                    sTimestamp)

            docOutput = Update( _
                    sComponentName, _
                    sClassName, _
                    decID, _
                    sTimestamp, _
                    ndAttributes, _
                    bClassCached, _
                    ndClassHierachy, _
                    Nothing, bKeepHistory)
        End If

        Return docOutput

    End Function

    Public Function Update( _
            ByVal sComponentName As String, _
            ByVal sClassName As String, _
            ByRef ndID As Xml.XmlNode, _
            ByRef ndAttributes As Xml.XmlNode) As Xml.XmlDocument

        Dim decToken As Decimal
        Dim sTimestamp As String

        GetValuesFromIDNode( _
                CType(ndID, Xml.XmlElement), _
                decToken, _
                sTimestamp)

        Return Update( _
                sComponentName, _
                sClassName, _
                decToken, _
                sTimestamp, _
                ndAttributes)

    End Function

    Public Function Update( _
            ByVal sComponentName As String, _
            ByVal sClassName As String, _
            ByVal decID As Decimal, _
            ByVal sTimestamp As String, _
            ByRef ndAttributes As Xml.XmlNode) As Xml.XmlDocument

        Return Update( _
                sComponentName, _
                sClassName, _
                decID, _
                sTimestamp, _
                ndAttributes, _
                False, _
                Nothing, _
                Nothing)

    End Function


    Public Function Update( _
            ByVal sComponentName As String, _
            ByVal sClassName As String, _
            ByRef ndID As Xml.XmlNode, _
            ByRef ndAttributes As Xml.XmlNode, _
            ByVal bClassCached As Boolean, _
            ByVal ndClassHierachy As Xml.XmlNode, _
            ByVal sResultRootName As String) As Xml.XmlDocument

        Return Update( _
            sComponentName, _
            sClassName, _
            ndID, _
            ndAttributes, _
            bClassCached, _
            ndClassHierachy, _
            sResultRootName, _
            False)

    End Function

    Public Function Update( _
            ByVal sComponentName As String, _
            ByVal sClassName As String, _
            ByRef ndID As Xml.XmlNode, _
            ByRef ndAttributes As Xml.XmlNode, _
            ByVal bClassCached As Boolean, _
            ByVal ndClassHierachy As Xml.XmlNode, _
            ByVal sResultRootName As String, _
            ByVal bKeepHistory As Boolean) As Xml.XmlDocument

        Dim decID As Decimal
        Dim sTimestamp As String

        GetValuesFromIDNode( _
                CType(ndID, Xml.XmlElement), _
                decID, _
                sTimestamp)

        Return Update( _
                sComponentName, _
                sClassName, _
                decID, _
                sTimestamp, _
                ndAttributes, _
                bClassCached, _
                ndClassHierachy, _
                sResultRootName, _
                bKeepHistory)

    End Function

    Public Function Update( _
            ByVal sComponentName As String, _
            ByVal sClassName As String, _
            ByVal decID As Decimal, _
            ByVal sTimestamp As String, _
            ByRef ndAttributes As Xml.XmlNode, _
            ByVal bClassCached As Boolean, _
            ByRef ndClassHierachy As Xml.XmlNode, _
            ByVal sResultRootName As String) As Xml.XmlDocument

        Return Update( _
                sComponentName, _
                sClassName, _
                decID, _
                sTimestamp, _
                ndAttributes, _
                bClassCached, _
                ndClassHierachy, _
                sResultRootName, _
                False)

    End Function

    Public Function Update( _
            ByVal sComponentName As String, _
            ByVal sClassName As String, _
            ByVal decID As Decimal, _
            ByVal sTimestamp As String, _
            ByRef ndAttributes As Xml.XmlNode, _
            ByVal bClassCached As Boolean, _
            ByRef ndClassHierachy As Xml.XmlNode, _
            ByVal sResultRootName As String, _
            ByVal bKeepHistory As Boolean) As Xml.XmlDocument

        Dim pcDataTypes As ParameterCollection
        Dim pcParameters As ParameterCollection
        Dim sStoredProcedureName As String
        Dim pcTimestamp As ParameterCollection
        Dim docResult As Xml.XmlDocument
        Dim ndResult As Xml.XmlNode
        Dim iSubClassDefinition As Integer
        Dim iSubClass As Integer
        Dim iSubClassDefinitions As Integer
        Dim iSubClasses As Integer
        Dim ndlSubClasses As Xml.XmlNodeList
        Dim ndlSubClassDefinitions As Xml.XmlNodeList
        Dim ndSubClassDefinition As Xml.XmlElement
        Dim ndSubClass As Xml.XmlNode
        Dim sSubClassDefinitionName As String
        Dim sSubClassName As String
        Dim sRelationalIDName As String
        Dim sbRelationalIDName As StringBuilder
        Dim ndID As Xml.XmlNode
        Dim bDelete As Boolean
        Dim decSubclassID As Decimal
        Dim sSubClassTimestamp As String
        Dim bHasMany As Boolean
        Dim pcID As ParameterCollection
        Dim sbSubClassFilter As StringBuilder
        Dim bSubKeepHistory As Boolean

        sStoredProcedureName = GetUpdateStoredProcedureName(sComponentName, sClassName)
        pcDataTypes = m_oDatabaseAdapter.GetStoredProcedureDataTypes(sStoredProcedureName)
        pcParameters = PopulateParameterCollectionFromXMLNode( _
                ndAttributes, _
                sComponentName, _
                sClassName, _
                pcDataTypes)

        If bKeepHistory Then
            CreateHistory(sComponentName, sClassName, ndAttributes)
        End If

        decID = Me.Update( _
                sComponentName, _
                sClassName, _
                decID, _
                sTimestamp, _
                pcParameters, _
                bClassCached)

        'Get new timestamp.
        pcTimestamp = GetInstanceParameterCollection( _
                sComponentName, _
                sClassName, _
                CStr(decID), _
                "ID", _
                False, _
                "Timestamp")
        sTimestamp = pcTimestamp.GetStringValue("Timestamp")

        'Process subclasses.
        If Not (ndClassHierachy Is Nothing) Then
            ndlSubClassDefinitions = ndClassHierachy.SelectNodes("*")
            iSubClassDefinitions = ndlSubClassDefinitions.Count - 1
            For iSubClassDefinition = 0 To iSubClassDefinitions
                ndSubClassDefinition = CType(ndlSubClassDefinitions.Item(iSubClassDefinition), Xml.XmlElement)
                sSubClassDefinitionName = ndSubClassDefinition.Name
                sSubClassName = GetStringXMLAttributeData("TagName", ndSubClassDefinition, sSubClassDefinitionName)
                bHasMany = GetBooleanXMLAttributeData("HasMany", ndSubClassDefinition, True)
                bSubKeepHistory = GetBooleanXMLAttributeData("KeepHistory", ndSubClassDefinition, False)
                sbRelationalIDName = New StringBuilder(sClassName)
                sbRelationalIDName.Append("ID")
                sRelationalIDName = sbRelationalIDName.ToString()

                'Get all subclasses from ndAttributes.
                ndlSubClasses = ndAttributes.SelectNodes(sSubClassName)
                iSubClasses = ndlSubClasses.Count - 1
                For iSubClass = 0 To iSubClasses
                    ndSubClass = ndlSubClasses.Item(iSubClass)
                    SetXMLElementData(sRelationalIDName, decID, ndSubClass)

                    ndID = ndSubClass.SelectSingleNode("ID")
                    If ndID Is Nothing Or (Not bHasMany) Then
                        If bHasMany Then
                            'Subclass doesn't currently exist, need to create it.
                            Me.Create( _
                                    sComponentName, _
                                    sSubClassDefinitionName, _
                                    ndSubClass, _
                                    bClassCached, _
                                    CType(ndSubClassDefinition, Xml.XmlNode))
                        Else
                            'Only one subclass will be present.
                            sbSubClassFilter = New StringBuilder
                            sbSubClassFilter.Append(sRelationalIDName)
                            sbSubClassFilter.Append(" = ")
                            sbSubClassFilter.Append(decID)
                            pcID = ListParameterCollection( _
                                    sComponentName, _
                                    sSubClassDefinitionName, _
                                    sbSubClassFilter.ToString(), _
                                    "ID", _
                                    bClassCached)
                            If pcID.MoveNext() Then
                                'Subclass found.
                                pcID = pcID.GetParameterCollection()
                                decSubclassID = pcID.GetDecimalValue("ID")
                                SetXMLElementData("ID", decSubclassID, ndSubClass)

                                Me.Update( _
                                        sComponentName, _
                                        sSubClassDefinitionName, _
                                        decSubclassID, _
                                        "", _
                                        ndSubClass, _
                                        bClassCached, _
                                        CType(ndSubClassDefinition, Xml.XmlNode), _
                                        "", _
                                        (bKeepHistory Or bSubKeepHistory))
                            Else
                                'Subclass does not exist yet.
                                Me.Create( _
                                        sComponentName, _
                                        sSubClassDefinitionName, _
                                        ndSubClass, _
                                        bClassCached, _
                                        CType(ndSubClassDefinition, Xml.XmlNode))
                            End If
                        End If
                    Else
                        If GetBooleanXMLAttributeData("Delete", CType(ndSubClass, Xml.XmlElement), False) Then
                            'Delete sub class as it is marked for deletion.
                            GetValuesFromIDNode(CType(ndID, Xml.XmlElement), decSubclassID, sSubClassTimestamp)
                            Me.Delete( _
                                    sComponentName, _
                                    sSubClassDefinitionName, _
                                    decSubclassID, _
                                    bClassCached, _
                                    CType(ndSubClassDefinition, Xml.XmlNode))
                        Else
                            'The only other option is to update the subclass.
                            SetXMLAttribute("Timestamp", ndSubClass.SelectSingleNode("ID"), "")

                            Me.Update( _
                                    sComponentName, _
                                    sSubClassDefinitionName, _
                                    ndSubClass, _
                                    bClassCached, _
                                    CType(ndSubClassDefinition, Xml.XmlNode), _
                                    (bKeepHistory Or bSubKeepHistory))
                        End If
                    End If
                Next iSubClass
            Next iSubClassDefinition
        End If

        'Create result document.
        If sResultRootName Is Nothing Then
            sResultRootName = GetDefaultXMLResponseRootName(sClassName, "Update")
        Else
            If sResultRootName.Length = 0 Then
                sResultRootName = GetDefaultXMLResponseRootName(sClassName, "Update")
            End If
        End If
        docResult = CreateXMLDocument(sResultRootName, Nothing)
        ndResult = docResult.SelectSingleNode(sResultRootName)
        AppendIDNode(ndResult, "ID", decID, sTimestamp)

        Return docResult

    End Function

    Private Function CheckTimestamp( _
            ByVal sCurrentTimestamp As String, _
            ByVal sNewTimestamp As String) As Boolean

        Dim bMatches As Boolean

        If sNewTimestamp Is Nothing Then
            bMatches = True
        Else
            If sNewTimestamp.Length = 0 Then
                bMatches = True
            Else
                If sCurrentTimestamp Is Nothing Then
                    Throw New ActiveBankException(0, "The timestamp column cannot be found.", ActiveBankException.ExceptionType.System)
                Else
                    If sCurrentTimestamp.Length = 0 Then
                        Throw New ActiveBankException(0, "The timestamp column cannot be found.", ActiveBankException.ExceptionType.System)
                    Else
                        If sCurrentTimestamp = sNewTimestamp Then
                            bMatches = True
                        Else
                            bMatches = False
                        End If
                    End If
                End If
            End If
        End If

        Return bMatches

    End Function

    Private Function GetUpdateStoredProcedureName( _
            ByVal sComponentName As String, _
            ByVal sClassName As String) As String

        Dim sbStoredProcedureName As New StringBuilder(sComponentName)

        sbStoredProcedureName.Append("_")
        sbStoredProcedureName.Append(sClassName)
        sbStoredProcedureName.Append("_U")

        Return GetObjectAlias(sbStoredProcedureName.ToString())

    End Function

    Public Function ExecuteStoredProcedure( _
            ByVal sComponentName As String, _
            ByVal sName As String, _
            ByRef pcParameters As ParameterCollection) As ParameterCollection

        Dim sbFullName As New StringBuilder(sComponentName)

        sbFullName.Append("_")
        sbFullName.Append(sName)

        Return ExecuteStoredProcedure( _
                sbFullName.ToString(), _
                pcParameters)

    End Function

    Public Function ExecuteStoredProcedure( _
            ByVal sName As String, _
            ByRef pcParameters As ParameterCollection) As ParameterCollection

        sName = GetObjectAlias(sName)
        Return m_oDatabaseAdapter.ExecuteStoredProcedure( _
                 sName, _
                 pcParameters, _
                 Nothing)

    End Function

    Public Function ExecuteStoredProcedureParameterCollection( _
        ByVal sName As String, _
        ByVal oParameters As ParameterCollection) As ParameterCollection

        Dim oDataTable As DataTable
        Dim oRecords As ADODB.RecordsetClass
        Dim oParameterCollection As ParameterCollection

        sName = GetObjectAlias(sName)
        oDataTable = m_oDatabaseAdapter.ExecuteStoredProcedureDataTable(sName, oParameters)
        oParameterCollection = ConvertDataTableToParameterCollection(oDataTable)
        Return oParameterCollection

    End Function

    Public Function ExecuteStoredProcedureParameterCollection( _
        ByVal sComponentName As String, _
        ByVal sClassName As String, _
        ByRef pcParameters As ParameterCollection) As ParameterCollection

        Dim sbFullName As StringBuilder

        sbFullName = New StringBuilder(sComponentName)
        sbFullName.Append("_")
        sbFullName.Append(sClassName)

        Return ExecuteStoredProcedureParameterCollection( _
                sbFullName.ToString(), _
                pcParameters)

    End Function

    Public Function ExecuteStoredProcedureDataTable( _
        ByVal sName As String, _
        ByVal oParameters As ParameterCollection) As DataTable

        sName = GetObjectAlias(sName)
        Return m_oDatabaseAdapter.ExecuteStoredProcedureDataTable(sName, oParameters)

    End Function

    Public Function ExecuteStoredProcedureXML( _
            ByVal sComponentName As String, _
            ByVal sClassName As String, _
            ByRef pcParameters As ParameterCollection, _
            ByRef ndParent As Xml.XmlNode, _
            ByVal sRowDescription As String) As Xml.XmlDocument

        Dim oOutput As Data.DataTable
        Dim sbCommand As StringBuilder
        Dim docOutput As Xml.XmlDocument
        Dim sName As String
        Dim sbName As New StringBuilder(sComponentName)

        sbName.Append("_")
        sbName.Append(sClassName)
        sName = GetObjectAlias(sbName.ToString())

        oOutput = m_oDatabaseAdapter.ExecuteStoredProcedureDataTable(sName, pcParameters)
        AppendDataToXMLDocument(oOutput, sComponentName, sClassName, ndParent, sRowDescription)

        Return docOutput

    End Function
#End Region

#Region "History Methods"
    Private Function CreateHistory(ByVal sComponentName As String, ByVal sClassName As String, ByVal ndAttributes As Xml.XmlNode) As Xml.XmlDocument

        Dim sHistoryStordProcName As String
        Dim pcHistoryDataTypes As ParameterCollection
        Dim ndHistory As Xml.XmlNode
        Dim docHistory As New Xml.XmlDocument
        Dim docCurrent As New Xml.XmlDocument
        Dim ndID As Xml.XmlNode
        Dim decID As Decimal
        Dim sClassHistoryName As String

        sClassHistoryName = sClassName & "History"
        sHistoryStordProcName = GetCreateStoredProcedureName(sComponentName, sClassHistoryName)
        pcHistoryDataTypes = m_oDatabaseAdapter.GetStoredProcedureDataTypes(sHistoryStordProcName)

        'Make sure the Stored proc exists
        If Not pcHistoryDataTypes Is Nothing Then
            If pcHistoryDataTypes.Length > 0 Then
                'JimP - Get the current data, and add some extra nodes for history
                docCurrent = Me.GetInstance(sComponentName, sClassName, GetStringXMLElementData("ID", ndAttributes), "ID", Nothing, sClassName)
                ndHistory = docCurrent.SelectSingleNode(sClassName)

                decID = GetDecimalXMLElementData("ID", ndHistory)
                AppendIDNode(ndHistory, "OriginalID", decID, String.Empty)
                AppendDateNode(ndHistory, "Date", Now)   'TODO: change to system run date?

                ndID = ndHistory.SelectSingleNode("ID")
                If Not (ndID Is Nothing) Then
                    ndHistory.RemoveChild(ndID)
                End If

                Return Me.Create(sComponentName, sClassHistoryName, ndHistory)
            Else
                Return Me.CreateErrorDocument(ErrorNumbers.TheClassDoesNotSupportHistory, sClassName & " does not support history.")
            End If
        Else
            Return Me.CreateErrorDocument(ErrorNumbers.TheClassDoesNotSupportHistory, sClassName & " does not support history.")
        End If

    End Function

#End Region

#Region "Delete Methods"

    Public Sub Delete( _
            ByVal sComponentName As String, _
            ByVal sClassName As String, _
            ByRef ndInput As Xml.XmlNode)

        Dim decID As Decimal

        decID = GetDecimalXMLElementData("ID", ndInput)

        Delete(sComponentName, sClassName, decID)

    End Sub

    Public Sub Delete( _
            ByVal sComponentName As String, _
            ByVal sClassName As String, _
            ByVal decID As Decimal)

        Delete(sComponentName, sClassName, decID, False)

    End Sub

    Public Sub Delete( _
            ByVal sComponentName As String, _
            ByVal sClassName As String, _
            ByVal ndInput As Xml.XmlNode, _
            ByVal bClassCached As Boolean)

        Dim decID As Decimal

        decID = GetDecimalXMLElementData("ID", ndInput)

        Delete(sComponentName, sClassName, decID, bClassCached, False)

    End Sub

    Public Sub Delete( _
            ByVal sComponentName As String, _
            ByVal sClassName As String, _
            ByVal decID As Decimal, _
            ByVal bClassCached As Boolean)

        Delete(sComponentName, sClassName, decID, bClassCached, False)

    End Sub

    Public Sub Delete( _
            ByVal sComponentName As String, _
            ByVal sClassName As String, _
            ByVal decID As String, _
            ByVal bClassCached As Boolean, _
            ByVal bMarkOnly As Boolean)

        Dim pcDependencies As ParameterCollection
        Dim pcDependency As ParameterCollection
        Dim pcCurrentRecord As ParameterCollection
        Dim sbDependencyFilter As New StringBuilder("SELECT ComponentName, ClassName FROM ")
        Dim oDependencyComponent As abComponentServices.IABComponent
        Dim eDeleteCheckResponse As abComponentServices.ComponentServices.DeleteResponse
        Dim sbStoredProcedureName As New StringBuilder(sComponentName)
        Dim pcDeleteParameters As ParameterCollection
        Dim sObjectName As String

        'AFS Get the current record. Also validates for existence.
        pcCurrentRecord = GetInstanceParameterCollection( _
                sComponentName, _
                sClassName, _
                CStr(decID))

        sObjectName = GetObjectAlias(sComponentName & "_DependenciesTbl")

        'AFS Check any dependency components.
        sbDependencyFilter.Append(sObjectName)
        pcDependencies = m_oDatabaseAdapter.ExecuteQueryParameterCollection(sbDependencyFilter.ToString(), False, False)
        pcDependencies.Reset()
        While pcDependencies.MoveNext()
            pcDependency = pcDependencies.GetParameterCollection()
            oDependencyComponent = CType(CreateComponentInstance( _
                    pcDependency.GetStringValue("ComponentName"), _
                    pcDependency.GetStringValue("ClassName")), abComponentServices.IABComponent)

            eDeleteCheckResponse = oDependencyComponent.DeleteCheck( _
                    sComponentName, _
                    sClassName, _
                    decID)
            If eDeleteCheckResponse = DeleteResponse.Vetoed Then
                Throw New ActiveBankException("The component: " & pcDependency.GetStringValue("ComponentName") & " has vetoed the deletion.", ActiveBankException.ExceptionType.Business)
            End If
        End While

        'AFS Delete checks have been made on other components. Now perform the deletion.
        sbStoredProcedureName.Append("_")
        sbStoredProcedureName.Append(sClassName)
        If bMarkOnly Then
            sbStoredProcedureName.Append("_US")
        Else
            sbStoredProcedureName.Append("_D")
        End If

        sObjectName = GetObjectAlias(sbStoredProcedureName.ToString())

        pcDeleteParameters = New ParameterCollection
        pcDeleteParameters.Add("ID", decID)

        If bMarkOnly Then
            pcDeleteParameters.Add("ObjectStatus", ObjectStatus.Deleted)
        End If

        m_oDatabaseAdapter.ExecuteStoredProcedure( _
                sObjectName, _
                pcDeleteParameters, _
                Nothing)

        If bClassCached Then
            InformApplicationServersOfDataChange( _
                    sComponentName, _
                    sClassName)
        End If

    End Sub

    Public Function Delete( _
            ByVal sComponentName As String, _
            ByVal sClassName As String, _
            ByVal decID As Decimal, _
            ByVal bClassCached As Boolean, _
            ByRef ndClassHierachy As Xml.XmlNode) As Xml.XmlDocument

        Return Delete(sComponentName, sClassName, decID, bClassCached, ndClassHierachy, "")

    End Function

    Public Function Delete( _
            ByVal sComponentName As String, _
            ByVal sClassName As String, _
            ByVal decID As Decimal, _
            ByVal bClassCached As Boolean, _
            ByRef ndClassHierachy As Xml.XmlNode, _
            ByVal sResultRootName As String) As Xml.XmlDocument

        Delete(sComponentName, sClassName, decID, bClassCached, ndClassHierachy, sResultRootName, False)

    End Function

    Public Function Delete( _
            ByVal sComponentName As String, _
            ByVal sClassName As String, _
            ByVal decID As Decimal, _
            ByVal bClassCached As Boolean, _
            ByRef ndClassHierachy As Xml.XmlNode, _
            ByVal sResultRootName As String, _
            ByVal bMarkOnly As Boolean) As Xml.XmlDocument

        Dim docResultParent As Xml.XmlDocument
        Dim ndResultParent As Xml.XmlNode

        'AFS Create result document.
        If sResultRootName Is Nothing Then
            sResultRootName = GetDefaultXMLResponseRootName(sClassName, "Delete")
        Else
            If sResultRootName.Length = 0 Then
                sResultRootName = GetDefaultXMLResponseRootName(sClassName, "Delete")
            End If
        End If
        docResultParent = CreateXMLDocument(sResultRootName, Nothing)
        ndResultParent = docResultParent.SelectSingleNode(sResultRootName)

        Delete(sComponentName, sClassName, decID, bClassCached, ndClassHierachy, ndResultParent, bMarkOnly)

        Return docResultParent

    End Function

    Public Sub Delete( _
            ByVal sComponentName As String, _
            ByVal sClassName As String, _
            ByVal decID As Decimal, _
            ByVal bClassCached As Boolean, _
            ByRef ndClassHierachy As Xml.XmlNode, _
            ByRef ndResultParent As Xml.XmlNode)

        Delete(sComponentName, sClassName, decID, bClassCached, ndClassHierachy, ndResultParent, False)

    End Sub

    Public Sub Delete( _
            ByVal sComponentName As String, _
            ByVal sClassName As String, _
            ByVal decID As Decimal, _
            ByVal bClassCached As Boolean, _
            ByRef ndClassHierachy As Xml.XmlNode, _
            ByRef ndResultParent As Xml.XmlNode, _
            ByVal bMarkOnly As Boolean)

        Dim iSubClassDefinition As Integer
        Dim iSubClassDefinitions As Integer
        Dim ndlSubClassDefinitions As Xml.XmlNodeList
        Dim ndSubClassDefinition As Xml.XmlElement
        Dim sSubClassDefinitionName As String
        Dim sSubClassName As String
        Dim sSubClassFilter As String
        Dim sbSubClassFilter As StringBuilder
        Dim pcSubClasses As ParameterCollection
        Dim pcSubClass As ParameterCollection
        Dim decSubClassID As Decimal

        Delete(sComponentName, sClassName, CStr(decID), bClassCached, bMarkOnly)

        'AFS Process sub classes.
        If Not (ndClassHierachy Is Nothing) Then
            ndlSubClassDefinitions = ndClassHierachy.SelectNodes("*")
            iSubClassDefinitions = ndlSubClassDefinitions.Count - 1
            For iSubClassDefinition = 0 To iSubClassDefinitions
                ndSubClassDefinition = CType(ndlSubClassDefinitions.Item(iSubClassDefinition), Xml.XmlElement)
                sSubClassDefinitionName = ndSubClassDefinition.Name
                sSubClassName = GetStringXMLAttributeData("TagName", ndSubClassDefinition, sSubClassDefinitionName)

                'AFS Get all subclasses from database.
                sbSubClassFilter = New StringBuilder(sClassName)
                sbSubClassFilter.Append("ID")
                sbSubClassFilter.Append(" = ")
                sbSubClassFilter.Append(decID)
                sSubClassFilter = sbSubClassFilter.ToString()
                pcSubClasses = Me.ListParameterCollection( _
                        sComponentName, _
                        sSubClassDefinitionName, _
                        sSubClassFilter, _
                        "ID")
                pcSubClasses.Reset()
                While pcSubClasses.MoveNext()
                    pcSubClass = pcSubClasses.GetParameterCollection()
                    decSubClassID = pcSubClass.GetDecimalValue("ID")

                    Delete(sComponentName, sSubClassDefinitionName, CStr(decSubClassID), bClassCached, bMarkOnly)
                End While
            Next iSubClassDefinition
        End If

        AppendXMLElement("ID", ndResultParent, decID)

    End Sub

#End Region

#Region "GetInstance Methods"

    Public Function GetInstanceParameterCollection( _
            ByVal sComponentName As String, _
            ByVal sClassName As String, _
            ByVal sToken As String) As ParameterCollection

        Return GetInstanceParameterCollection( _
                sComponentName, _
                sClassName, _
                sToken, _
                "ID", _
                False)

    End Function

    Public Function GetInstanceParameterCollection( _
            ByVal sComponentName As String, _
            ByVal sClassName As String, _
            ByVal sToken As String, _
            ByVal bCache As Boolean) As ParameterCollection

        Return GetInstanceParameterCollection( _
                sComponentName, _
                sClassName, _
                sToken, _
                "ID", _
                bCache)

    End Function

    Public Function GetInstanceParameterCollection( _
            ByVal sComponentName As String, _
            ByVal sClassName As String, _
            ByVal sToken As String, _
            ByVal sTokenField As String, _
            ByVal bCache As Boolean) As ParameterCollection

        Return GetInstanceParameterCollection( _
                sComponentName, _
                sClassName, _
                sToken, _
                sTokenField, _
                bCache, _
                "*")

    End Function

    Public Function GetInstanceParameterCollection( _
            ByVal sComponentName As String, _
            ByVal sClassName As String, _
            ByVal sToken As String, _
            ByVal sTokenField As String, _
            ByVal sFields As String) As ParameterCollection

        Return GetInstanceParameterCollection( _
                sComponentName, _
                sClassName, _
                sToken, _
                sTokenField, _
                False, _
                sFields)

    End Function

    Public Function GetInstanceParameterCollection( _
            ByVal sComponentName As String, _
            ByVal sClassName As String, _
            ByVal sToken As String, _
            ByVal sTokenField As String, _
            ByVal bCache As Boolean, _
            ByVal sFields As String) As ParameterCollection

        Dim sbQuery As New StringBuilder("SELECT ")
        Dim sbView As New StringBuilder(sComponentName)
        Dim sObjectName As String
        Dim sQuery As String
        Dim oResult As ParameterCollection

        If sFields.Length = 0 Then
            sFields = "*"
        End If

        sbView.Append("_")
        sbView.Append(sClassName)
        sbView.Append("Vw")

        sObjectName = GetObjectAlias(sbView.ToString)

        sbQuery.Append(sFields)
        sbQuery.Append(" FROM ")
        sbQuery.Append(sObjectName)
        sbQuery.Append(" WHERE ")
        sbQuery.Append(sTokenField)

        If sTokenField.Trim.ToUpper = "ID" Then
            sbQuery.Append(" = ")
            sbQuery.Append(sToken)
        Else
            sbQuery.Append(" = '")
            sbQuery.Append(sToken)
            sbQuery.Append("'")
        End If

        sQuery = sbQuery.ToString()

        If bCache Then
            oResult = RetrieveParameterCollectionFromCache(sComponentName, sClassName, sQuery)
            If oResult Is Nothing Then
                oResult = m_oDatabaseAdapter.ExecuteQueryParameterCollection(sQuery, True, False)
                AddDataToCache(sComponentName, sClassName, sQuery, oResult)
            End If
        Else
            oResult = m_oDatabaseAdapter.ExecuteQueryParameterCollection(sQuery, True, False)
        End If

        If oResult Is Nothing Then
            Throw New ActiveBankException(ErrorNumbers.TheInstanceHasNotBeenFound, "The specified " + sClassName + " instance has not been found.", ActiveBankException.ExceptionType.Business)
        End If

        Return oResult

    End Function

    Public Function GetInstance( _
            ByVal sComponentName As String, _
            ByVal sClassName As String, _
            ByVal sToken As String) As Xml.XmlDocument

        Return GetInstance( _
                sComponentName, _
                sClassName, _
                sToken, _
                False)

    End Function

    Public Function GetInstance( _
            ByVal sComponentName As String, _
            ByVal sClassName As String, _
            ByVal sToken As String, _
            ByVal bCache As Boolean) As Xml.XmlDocument

        Return Me.GetInstance( _
                sComponentName, _
                sClassName, _
                sToken, _
                bCache, _
                Nothing)

    End Function

    Public Function GetInstance( _
            ByVal sComponentName As String, _
            ByVal sClassName As String, _
            ByVal sToken As String, _
            ByVal sTokenField As String, _
            ByVal bCache As Boolean) As Xml.XmlDocument

        Return Me.GetInstance( _
                sComponentName, _
                sClassName, _
                sToken, _
                bCache, _
                Nothing, _
                sTokenField, _
                "", _
                "")

    End Function

    Public Function GetInstance( _
            ByVal sComponentName As String, _
            ByVal sClassName As String, _
            ByVal sToken As String, _
            ByVal sTokenField As String, _
            ByVal sRequiredAttributes As String) As Xml.XmlDocument

        Return Me.GetInstance( _
                sComponentName, _
                sClassName, _
                sToken, _
                sTokenField, _
                sRequiredAttributes, _
                "")

    End Function

    Public Function GetInstance( _
            ByVal sComponentName As String, _
            ByVal sClassName As String, _
            ByVal sToken As String, _
            ByVal sTokenField As String, _
            ByVal sRequiredAttributes As String, _
            ByVal sRootName As String) As Xml.XmlDocument

        Return GetInstance( _
                sComponentName, _
                sClassName, _
                sToken, _
                False, _
                Nothing, _
                sTokenField, _
                sRequiredAttributes, _
                sRootName)

    End Function

    Public Function GetInstance( _
            ByVal sComponentName As String, _
            ByVal sClassName As String, _
            ByVal sToken As String, _
            ByVal bCache As Boolean, _
            ByRef ndClassHierachy As Xml.XmlNode) As Xml.XmlDocument

        Return GetInstance( _
                sComponentName, _
                sClassName, _
                sToken, _
                bCache, _
                ndClassHierachy, _
                "ID", _
                "", _
                "")

    End Function

    Public Function GetInstance( _
            ByVal sComponentName As String, _
            ByVal sClassName As String, _
            ByVal sToken As String, _
            ByVal bCache As Boolean, _
            ByRef ndClassHierachy As Xml.XmlNode, _
            ByVal sTokenField As String, _
            ByVal sRequiredAttrbutes As String, _
            ByVal sRootName As String) As Xml.XmlDocument

        Dim ndResultParent As Xml.XmlNode
        Dim sResultParentName As String
        Dim docResult As Xml.XmlDocument
        Dim sResult As String

        If sRootName.Length = 0 Then
            sResultParentName = GetDefaultXMLResponseRootName(sClassName, "Get")
        Else
            sResultParentName = sRootName
        End If

        docResult = CreateXMLDocument(sResultParentName, Nothing)
        ndResultParent = docResult.SelectSingleNode(sResultParentName)

        If bCache Then
            sResult = RetrieveXMLDocumentFromCache(sComponentName, sClassName, sToken)
            If sResult Is Nothing Then
                GetInstance( _
                        sComponentName, _
                        sClassName, _
                        sToken, _
                        ndClassHierachy, _
                        ndResultParent, _
                        "", _
                        sTokenField, _
                        sRequiredAttrbutes)

                AddDataToCache(sComponentName, sClassName, sToken, ndResultParent.OuterXml)
            Else
                docResult.LoadXml(sResult)
            End If
        Else
            GetInstance( _
                    sComponentName, _
                    sClassName, _
                    sToken, _
                    ndClassHierachy, _
                    ndResultParent, _
                    "", _
                    sTokenField, _
                    sRequiredAttrbutes)
        End If

        Return docResult

    End Function

    Public Sub GetInstance( _
            ByVal strComponentName As String, _
            ByVal strClassName As String, _
            ByVal strToken As String, _
            ByRef ndResultParent As Xml.XmlNode)

        GetInstance( _
                strComponentName, _
                strClassName, _
                strToken, _
                CType(Nothing, Xml.XmlNode), _
                ndResultParent)

    End Sub

    Public Sub GetInstance( _
            ByVal strComponentName As String, _
            ByVal strClassName As String, _
            ByVal strToken As String, _
            ByRef ndClassHierachy As Xml.XmlNode, _
            ByRef ndResultParent As Xml.XmlNode)

        GetInstance( _
                strComponentName, _
                strClassName, _
                strToken, _
                ndClassHierachy, _
                ndResultParent, _
                "", _
                "ID")

    End Sub

    Public Sub GetInstance( _
            ByVal sComponentName As String, _
            ByVal sClassName As String, _
            ByVal sToken As String, _
            ByRef ndClassHierachy As Xml.XmlNode, _
            ByRef ndResultParent As Xml.XmlNode, _
            ByVal sDataDescription As String, _
            ByVal sTokenField As String)

        GetInstance( _
                sComponentName, _
                sClassName, _
                sToken, _
                ndClassHierachy, _
                ndResultParent, _
                sDataDescription, _
                sTokenField, _
                "*")

    End Sub
    Public Sub GetInstance( _
            ByVal sComponentName As String, _
            ByVal sClassName As String, _
            ByVal sToken As String, _
            ByRef ndClassHierachy As Xml.XmlNode, _
            ByRef ndResultParent As Xml.XmlNode, _
            ByVal sDataDescription As String, _
            ByVal sTokenField As String, _
            ByVal sRequiredAttrbutes As String)

        Dim sbQuery As New StringBuilder("SELECT ")
        Dim sbView As New StringBuilder(sComponentName)
        Dim strQuery As String
        Dim iSubClassDefinition As Integer
        Dim iSubClassDefinitions As Integer
        Dim ndlSubClassDefinitions As Xml.XmlNodeList
        Dim ndSubClassDefinition As Xml.XmlElement
        Dim sSubClassDefinitionName As String
        Dim sSubClassName As String
        Dim sSubClassFilter As String
        Dim sbSubClassFilter As StringBuilder
        Dim pcSubClasses As ParameterCollection
        Dim pcSubClass As ParameterCollection
        Dim ndSubClassResultParent As Xml.XmlNode
        Dim decSubClassID As Decimal
        Dim sbResultParentSearchPath As StringBuilder
        Dim sResultParentSearchPath As String

        sbView.Append("_")
        sbView.Append(sClassName)
        sbView.Append("Vw ")


        If sRequiredAttrbutes = String.Empty Then
            sbQuery.Append("*")
        Else
            If Not sRequiredAttrbutes.Equals("*") Then
                sbQuery.Append("ID, Timestamp, ")
            End If
            sbQuery.Append(sRequiredAttrbutes)
        End If
        sbQuery.Append(" FROM ")
        sbQuery.Append(GetObjectAlias(sbView.ToString()))
        sbQuery.Append(" WHERE ")
        sbQuery.Append(sTokenField)

        If sTokenField.Trim.ToUpper = "ID" Then
            sbQuery.Append(" = ")
            sbQuery.Append(sToken)
        Else
            sbQuery.Append(" = '")
            sbQuery.Append(sToken)
            sbQuery.Append("'")
        End If

        strQuery = sbQuery.ToString()

        If Not AppendDataToXMLDocument(strQuery, sComponentName, sClassName, ndResultParent, sDataDescription) Then
            Throw New ActiveBankException( _
                    ErrorNumbers.TheInstanceHasNotBeenFound, _
                    "The specified " + sClassName + " instance has not been found.", _
                    ActiveBankException.ExceptionType.Business, _
                    "abComponentServices")
        End If

        'AFS Process sub classes.
        If Not (ndClassHierachy Is Nothing) Then
            ndlSubClassDefinitions = ndClassHierachy.SelectNodes("*")
            iSubClassDefinitions = ndlSubClassDefinitions.Count - 1
            For iSubClassDefinition = 0 To iSubClassDefinitions
                ndSubClassDefinition = CType(ndlSubClassDefinitions.Item(iSubClassDefinition), Xml.XmlElement)
                sSubClassDefinitionName = ndSubClassDefinition.Name
                sSubClassName = GetStringXMLAttributeData("TagName", ndSubClassDefinition, sSubClassDefinitionName)

                'AFS Get all subclasses from database.
                sbSubClassFilter = New StringBuilder(sClassName)
                sbSubClassFilter.Append("ID")
                sbSubClassFilter.Append(" = ")
                sbSubClassFilter.Append(sToken)
                sSubClassFilter = sbSubClassFilter.ToString()
                pcSubClasses = Me.ListParameterCollection( _
                        sComponentName, _
                        sSubClassDefinitionName, _
                        sSubClassFilter, _
                        "ID")
                pcSubClasses.Reset()
                While pcSubClasses.MoveNext()
                    pcSubClass = pcSubClasses.GetParameterCollection()
                    decSubClassID = pcSubClass.GetDecimalValue("ID")

                    If sDataDescription.Length = 0 Then
                        ndSubClassResultParent = ndResultParent
                    Else
                        sbResultParentSearchPath = New StringBuilder(sDataDescription)
                        sbResultParentSearchPath.Append("[ID=")
                        sbResultParentSearchPath.Append(sToken)
                        sbResultParentSearchPath.Append("]")
                        sResultParentSearchPath = sbResultParentSearchPath.ToString()
                        ndSubClassResultParent = ndResultParent.SelectSingleNode(sResultParentSearchPath)
                    End If

                    Me.GetInstance( _
                            sComponentName, _
                            sSubClassDefinitionName, _
                            decSubClassID.ToString(System.Globalization.CultureInfo.InvariantCulture), _
                            CType(ndSubClassDefinition, Xml.XmlNode), _
                            ndSubClassResultParent, _
                            sSubClassName, _
                            "ID")
                End While
            Next iSubClassDefinition
        End If
    End Sub

#End Region

#Region "GetInstanceHistory Methods"
    Public Function GetInstanceHistoryParameterCollection( _
        ByVal sComponentName As String, _
        ByVal sClassName As String, _
        ByVal sToken As String) As ParameterCollection

        Return GetInstanceParameterCollection(sComponentName, sClassName & "History", sToken)
    End Function

    Public Function GetInstanceHistoryParameterCollection( _
        ByVal sComponentName As String, _
        ByVal sClassName As String, _
        ByVal sToken As String, _
        ByVal bCache As Boolean) As ParameterCollection

        Return GetInstanceParameterCollection(sComponentName, sClassName & "History", sToken, bCache)
    End Function

    Public Function GetInstanceHistoryParameterCollection( _
        ByVal sComponentName As String, _
        ByVal sClassName As String, _
        ByVal sToken As String, _
        ByVal sTokenField As String, _
        ByVal bCache As Boolean) As ParameterCollection

        Return GetInstanceParameterCollection(sComponentName, sClassName & "History", sToken, sTokenField, bCache)
    End Function

    Public Function GetInstanceHistoryParameterCollection( _
        ByVal sComponentName As String, _
        ByVal sClassName As String, _
        ByVal sToken As String, _
        ByVal sTokenField As String, _
        ByVal sFields As String) As ParameterCollection

        Return GetInstanceParameterCollection(sComponentName, sClassName & "History", sToken, sTokenField, sFields)
    End Function

    Public Function GetInstanceHistoryParameterCollection( _
        ByVal sComponentName As String, _
        ByVal sClassName As String, _
        ByVal sToken As String, _
        ByVal sTokenField As String, _
        ByVal bCache As Boolean, _
        ByVal sFields As String) As ParameterCollection

        Return GetInstanceParameterCollection(sComponentName, sClassName & "History", sToken, sTokenField, bCache, sFields)
    End Function

    Public Function GetInstanceHistory( _
        ByVal sComponentName As String, _
        ByVal sClassName As String, _
        ByVal sToken As String) As Xml.XmlDocument

        Return GetInstance(sComponentName, sClassName & "History", sToken)
    End Function

    Public Function GetInstanceHistory( _
        ByVal sComponentName As String, _
        ByVal sClassName As String, _
        ByVal sToken As String, _
        ByVal bCache As Boolean) As Xml.XmlDocument

        Return GetInstance(sComponentName, sClassName & "History", sToken, bCache)
    End Function

    Public Function GetInstanceHistory( _
        ByVal sComponentName As String, _
        ByVal sClassName As String, _
        ByVal sToken As String, _
        ByVal sTokenField As String, _
        ByVal sRequiredAttributes As String) As Xml.XmlDocument

        Return GetInstance(sComponentName, sClassName & "History", sToken, sTokenField, sRequiredAttributes)
    End Function

    Public Function GetInstanceHistory( _
        ByVal sComponentName As String, _
        ByVal sClassName As String, _
        ByVal sToken As String, _
        ByVal sTokenField As String, _
        ByVal sRequiredAttributes As String, _
        ByVal sRootName As String) As Xml.XmlDocument

        Return GetInstance(sComponentName, sClassName & "History", sToken, sTokenField, sRequiredAttributes, sRootName)
    End Function

    Public Function GetInstanceHistory( _
        ByVal sComponentName As String, _
        ByVal sClassName As String, _
        ByVal sToken As String, _
        ByVal bCache As Boolean, _
        ByRef ndClassHierachy As Xml.XmlNode) As Xml.XmlDocument

        Return GetInstance(sComponentName, sClassName & "History", sToken, bCache, ndClassHierachy)
    End Function

    Public Function GetInstanceHistory( _
        ByVal sComponentName As String, _
        ByVal sClassName As String, _
        ByVal sToken As String, _
        ByVal bCache As Boolean, _
        ByRef ndClassHierachy As Xml.XmlNode, _
        ByVal sTokenField As String, _
        ByVal sRequiredAttrbutes As String, _
        ByVal sRootName As String) As Xml.XmlDocument

        Return GetInstance(sComponentName, sClassName & "History", sToken, bCache, ndClassHierachy, sTokenField, sRequiredAttrbutes, sRootName)
    End Function

    Public Sub GetInstanceHistory( _
        ByVal strComponentName As String, _
        ByVal strClassName As String, _
        ByVal strToken As String, _
        ByRef ndResultParent As Xml.XmlNode)

        GetInstance(strComponentName, strClassName & "History", strToken, ndResultParent)
    End Sub

    Public Sub GetInstanceHistory( _
        ByVal sComponentName As String, _
        ByVal sClassName As String, _
        ByVal sToken As String, _
        ByRef ndClassHierachy As Xml.XmlNode, _
        ByRef ndResultParent As Xml.XmlNode)

        GetInstance(sComponentName, sClassName & "History", sToken, ndClassHierachy, ndResultParent)
    End Sub

    Public Sub GetInstanceHistory( _
        ByVal sComponentName As String, _
        ByVal sClassName As String, _
        ByVal sToken As String, _
        ByRef ndClassHierachy As Xml.XmlNode, _
        ByRef ndResultParent As Xml.XmlNode, _
        ByVal sDataDescription As String, _
        ByVal sTokenField As String)

        GetInstance(sComponentName, sClassName & "History", sToken, ndClassHierachy, ndResultParent, sDataDescription, sTokenField)
    End Sub
#End Region

#Region "List Methods"

    Public Function ListParameterCollection( _
            ByVal sComponentName As String, _
            ByVal sClassName As String, _
            ByVal sFilter As String) As ParameterCollection

        Return ListParameterCollection( _
                sComponentName, _
                sClassName, _
                sFilter, _
                "*")

    End Function

    Public Function ListParameterCollection( _
            ByVal sComponentName As String, _
            ByVal sClassName As String, _
            ByVal bCache As Boolean) As ParameterCollection

        Return ListParameterCollection( _
                sComponentName, _
                sClassName, _
                "", _
                bCache)

    End Function

    Public Function ListParameterCollection( _
        ByVal sComponentName As String, _
        ByVal sClassName As String, _
        ByVal sFilter As String, _
        ByVal bCache As Boolean) As ParameterCollection

        Return ListParameterCollection( _
                sComponentName, _
                sClassName, _
                sFilter, _
                "*", _
                bCache)

    End Function

    Public Function ListParameterCollection( _
            ByVal sComponentName As String, _
            ByVal sClassName As String, _
            ByVal sFilter As String, _
            ByVal sFields As String) As ParameterCollection

        Return ListParameterCollection( _
                sComponentName, _
                sClassName, _
                sFilter, _
                sFields, _
                False)

    End Function

    Public Function ListParameterCollection( _
            ByVal sComponentName As String, _
            ByVal sClassName As String, _
            ByVal sFilter As String, _
            ByVal sFields As String, _
            ByVal bCache As Boolean) As ParameterCollection

        Return ListParameterCollection( _
                sComponentName, _
                sClassName, _
                sFilter, _
                sFields, _
                "", _
                bCache)

    End Function
    Public Function ListParameterCollection( _
                ByVal sComponentName As String, _
                ByVal sClassName As String, _
                ByVal sFilter As String, _
                ByVal sFields As String, _
                ByVal sSort As String, _
                ByVal bCache As Boolean) As ParameterCollection

        Return ListParameterCollection( _
                sComponentName, _
                sClassName, _
                sFilter, _
                sFields, _
                sSort, _
                bCache, _
                False)
    End Function
    Public Function ListParameterCollection( _
            ByVal sComponentName As String, _
            ByVal sClassName As String, _
            ByVal sFilter As String, _
            ByVal sFields As String, _
            ByVal sSort As String, _
            ByVal bCache As Boolean, _
            ByVal bIDOrdinal As Boolean) As ParameterCollection


        Dim sbQuery As New StringBuilder("SELECT ")
        Dim sbView As New StringBuilder(sComponentName)
        Dim oResult As ParameterCollection
        Dim sQuery As String

        If sFields.Length = 0 Then
            sFields = "*"
        End If

        sbQuery.Append(sFields)
        sbQuery.Append(" FROM ")

        sbView.Append("_")
        sbView.Append(sClassName)
        sbView.Append("Vw")

        sbQuery.Append(GetObjectAlias(sbView.ToString))

        If sFilter.Length > 0 Then
            sbQuery.Append(" WHERE ")
            sbQuery.Append(sFilter)
            sbQuery.Append(" AND ObjectStatus <> ")
            sbQuery.Append(ObjectStatus.Deleted)
        Else
            sbQuery.Append(" WHERE ")
            sbQuery.Append(" ObjectStatus <> ")
            sbQuery.Append(ObjectStatus.Deleted)
        End If

        If sSort.Length > 0 Then
            sbQuery.Append(" ORDER BY ")
            sbQuery.Append(sSort)
        End If

        sQuery = sbQuery.ToString()

        If bCache Then
            oResult = RetrieveParameterCollectionFromCache( _
                    sComponentName, _
                    sClassName, _
                    sQuery)
            If oResult Is Nothing Then
                oResult = m_oDatabaseAdapter.ExecuteQueryParameterCollection( _
                        sQuery, _
                        False, bIDOrdinal)

                AddDataToCache( _
                        sComponentName, _
                        sClassName, _
                        sQuery, _
                        oResult)
            End If
        Else
            oResult = m_oDatabaseAdapter.ExecuteQueryParameterCollection( _
                    sbQuery.ToString(), _
                    False, bIDOrdinal)
        End If

        If Not oResult Is Nothing Then
            oResult.Reset()
        End If

        Return oResult

    End Function

    Public Function List( _
            ByVal sComponentName As String, _
            ByVal sClassName As String, _
            ByVal sFilter As String) As Xml.XmlDocument

        Return List( _
                sComponentName, _
                sClassName, _
                sFilter, _
                "*")

    End Function

    Public Function List( _
            ByVal sComponentName As String, _
            ByVal sClassName As String, _
            ByVal sFilter As String, _
            ByVal sFields As String) As Xml.XmlDocument

        Return List( _
                sComponentName, _
                sClassName, _
                sFilter, _
                sFields, _
                False)

    End Function

    Public Function List( _
            ByVal sComponentName As String, _
            ByVal sClassName As String, _
            ByVal sFilter As String, _
            ByVal sFields As String, _
            ByVal bCache As Boolean) As Xml.XmlDocument

        Return List( _
                sComponentName, _
                sClassName, _
                sFilter, _
                sFields, _
                Nothing, _
                bCache)

    End Function

    Public Function List( _
            ByVal sComponentName As String, _
            ByVal sClassName As String, _
            ByVal sFilter As String, _
            ByVal sFields As String, _
            ByRef ndFilter As Xml.XmlNode, _
            ByVal bCache As Boolean) As Xml.XmlDocument

        Return List( _
                sComponentName, _
                sClassName, _
                sFilter, _
                sFields, _
                ndFilter, _
                Nothing, _
                bCache)

    End Function

    Public Function List( _
            ByVal sComponentName As String, _
            ByVal sClassName As String, _
            ByVal sFilter As String, _
            ByVal sFields As String, _
            ByRef ndFilter As Xml.XmlNode, _
            ByRef ndListHeader As Xml.XmlNode, _
            ByVal bCache As Boolean) As Xml.XmlDocument

        Dim sbRootName As New StringBuilder("List")
        Dim sRootName As String

        sbRootName.Append(sClassName)
        sbRootName.Append("s")
        sbRootName.Append("Rs")
        sRootName = sbRootName.ToString()

        Return Me.List( _
                sComponentName, _
                sClassName, _
                sFilter, _
                sFields, _
                ndFilter, _
                ndListHeader, _
                bCache, _
                sRootName)

    End Function

    Public Function List( _
            ByVal sComponentName As String, _
            ByVal sClassName As String, _
            ByVal sFilter As String, _
            ByVal sFields As String, _
            ByRef ndFilter As Xml.XmlNode, _
            ByRef ndListHeader As Xml.XmlNode, _
            ByVal bCache As Boolean, _
            ByVal sRootName As String) As Xml.XmlDocument

        Return List( _
                sComponentName, _
                sClassName, _
                sFilter, _
                sFields, _
                ndFilter, _
                ndListHeader, _
                bCache, _
                sRootName, _
                sClassName)

    End Function

    Public Function List( _
            ByVal sComponentName As String, _
            ByVal sClassName As String, _
            ByVal sFilter As String, _
            ByVal sFields As String, _
            ByRef ndFilter As Xml.XmlNode, _
            ByRef ndListHeader As Xml.XmlNode, _
            ByVal bCache As Boolean, _
            ByVal sRootName As String, _
            ByVal sRowDescription As String) As Xml.XmlDocument

        Dim docParent As Xml.XmlDocument
        Dim ndParent As Xml.XmlNode
        Dim strCachingKey As String
        Dim strResult As String
        Dim sbCachingKey As StringBuilder

        docParent = CreateXMLDocument(sRootName, Nothing)
        ndParent = docParent.SelectSingleNode(sRootName)

        If bCache Then
            sbCachingKey = New StringBuilder(sFields)
            If sFilter.Length > 0 Then
                sbCachingKey.Append(sFilter)
            End If
            If Not (ndFilter Is Nothing) Then
                sbCachingKey.Append(ndFilter.OuterXml)
            End If
            strCachingKey = sbCachingKey.ToString()
            strResult = RetrieveXMLDocumentFromCache( _
                    sComponentName, _
                    sClassName, _
                    strCachingKey)
            If strResult Is Nothing Then
                List( _
                        sComponentName, _
                        sClassName, _
                        sFilter, _
                        sFields, _
                        ndFilter, _
                        ndListHeader, _
                        ndParent, _
                        sRowDescription)

                AddDataToCache(sComponentName, sClassName, strCachingKey, docParent.OuterXml)
            Else
                docParent.LoadXml(strResult)
            End If
        Else
            List( _
                    sComponentName, _
                    sClassName, _
                    sFilter, _
                    sFields, _
                    ndFilter, _
                    ndListHeader, _
                    ndParent, _
                    sRowDescription)
        End If

        Return docParent

    End Function

    Public Sub List( _
            ByVal sComponentName As String, _
            ByVal sClassName As String, _
            ByVal sFilter As String, _
            ByVal sFields As String, _
            ByRef ndFilter As Xml.XmlNode, _
            ByRef ndListHeader As Xml.XmlNode, _
            ByVal ndParent As Xml.XmlNode, _
            ByVal sRowDescription As String)

        Dim sbQuery As StringBuilder
        Dim sbFields As StringBuilder
        Dim strQuery As String
        Dim sbFilter As StringBuilder
        Dim iRequiredPage As Integer
        Dim iMaximumRows As Integer
        Dim sbViewName As StringBuilder
        Dim sViewName As String
        Dim sSort As String
        Dim ndSort As Xml.XmlNode
        Dim sbSort As StringBuilder
        Dim ndlSorts As Xml.XmlNodeList
        Dim ndlFilters As Xml.XmlNodeList


        If Not sFields.Equals("*") Then
            sbFields = New StringBuilder(sFields)
            sbFields.Append(", ID, Timestamp")
            sFields = sbFields.ToString()
        End If

        If sFilter.Length > 0 Then
            sbFilter = New StringBuilder(" WHERE ")
            sbFilter.Append(sFilter)
            sFilter = sbFilter.ToString()
        End If

        If sFilter.Length > 0 Then
            sbFilter.Append(" AND ")
        Else
            sbFilter = New StringBuilder(" Where ")
        End If
        sbFilter.Append(" ObjectStatus <> ")
        sbFilter.Append(ObjectStatus.Deleted)
        sFilter = sbFilter.ToString()

        AppendNamedFilterAttributes(ndFilter, sbFilter)

        If Not (ndListHeader Is Nothing) Then
            iMaximumRows = GetIntegerXMLElementData("PageSize", ndListHeader, 0)
            iRequiredPage = GetIntegerXMLElementData("PageNumber", ndListHeader, 0)
            iRequiredPage = iRequiredPage - 1
            If iRequiredPage < 0 Then
                iRequiredPage = 0
            End If

            ndlSorts = ndListHeader.SelectNodes("Sort")
            If ndlSorts.Count > 0 Then
                sbSort = New StringBuilder
                For Each ndSort In ndlSorts
                    If sbSort.Length > 0 Then
                        sbSort.Append(", ")
                    Else
                        sbSort.Append(" ORDER BY ")
                    End If

                    sbSort.Append(GetStringXMLElementData("Name", ndSort))
                    If GetBooleanXMLElementData("Ascending", ndSort, True) Then
                        sbSort.Append(" ASC")
                    Else
                        sbSort.Append(" DESC")
                    End If
                Next
                sSort = sbSort.ToString()
            Else
                sSort = " ORDER BY ID"
            End If

            ndlFilters = ndListHeader.SelectNodes("Filter")
            AppendFilterAttributes(ndlFilters, sbFilter)
        Else
            sSort = ""
        End If

        If sbFilter Is Nothing Then
            sFilter = ""
        Else
            sFilter = sbFilter.ToString()
        End If

        sbQuery = New StringBuilder("SELECT ")

        If iMaximumRows > 0 Then
            sbQuery.Append(" TOP ")
            sbQuery.Append(CStr(iMaximumRows))
            sbQuery.Append(" ")
        End If

        sbQuery.Append(sFields)
        sbQuery.Append(" FROM ")

        sbViewName = New StringBuilder(sComponentName)
        sbViewName.Append("_")
        sbViewName.Append(sClassName)
        sbViewName.Append("Vw")
        sViewName = GetObjectAlias(sbViewName.ToString())

        sbQuery.Append(sViewName)

        If iRequiredPage > 0 And iMaximumRows > 0 Then
            If sFilter.Length = 0 Then
                sbQuery.Append(" WHERE ID NOT IN ")
            Else
                sbQuery.Append(sFilter)
                sbQuery.Append(" AND ID NOT IN ")
            End If
            sbQuery.Append("(SELECT TOP ")
            sbQuery.Append(iMaximumRows * iRequiredPage)
            sbQuery.Append(" ID FROM ")
            sbQuery.Append(sViewName)
            sbQuery.Append(sFilter)
            sbQuery.Append(sSort)
            sbQuery.Append(")")
        Else
            sbQuery.Append(sFilter)
        End If

        sbQuery.Append(sSort)

        strQuery = sbQuery.ToString()

        AppendDataToXMLDocument( _
                strQuery, _
                sComponentName, _
                sClassName, _
                ndParent, _
                sRowDescription)

    End Sub

    Public Function ListString( _
            ByVal sComponentName As String, _
            ByVal sClassName As String, _
            ByVal sFilter As String) As String

        Dim docResult As Xml.XmlDocument

        docResult = List( _
                sComponentName, _
                sClassName, _
                sFilter, _
                "*")

        Return docResult.OuterXml

    End Function

    Private Sub AppendNamedFilterAttributes( _
            ByRef ndNamedFilterAttributes As Xml.XmlNode, _
            ByRef sbFilter As StringBuilder)

        Dim sAttributeValue As String
        Dim ndComparisonType As Xml.XmlAttribute
        Dim ndConjunction As Xml.XmlAttribute
        Dim ndlFilterAttributesList As Xml.XmlNodeList
        Dim iFilterAttribute As Integer
        Dim iFilterAttributeCount As Integer
        Dim ndFilterAttribute As Xml.XmlNode
        Dim bFirstAttribute As Boolean = True
        Dim eComparisonType As abComponentServices.ComponentServices.ListFilterComparisonType
        Dim iOpenParentheses As Int32
        Dim iCloseParentheses As Int32

        If Not (ndNamedFilterAttributes Is Nothing) Then
            ndlFilterAttributesList = ndNamedFilterAttributes.SelectNodes("*")
            iFilterAttributeCount = ndlFilterAttributesList.Count - 1

            If iFilterAttributeCount > -1 Then
                For iFilterAttribute = 0 To iFilterAttributeCount
                    ndFilterAttribute = ndlFilterAttributesList(iFilterAttribute)

                    If Not IsFilterAttributeExcluded(ndFilterAttribute.Name) Then
                        If sbFilter Is Nothing Then
                            sbFilter = New StringBuilder(" WHERE ")
                        Else
                            ndConjunction = CType(ndFilterAttribute.Attributes.GetNamedItem("Conjunction"), Xml.XmlAttribute)
                            If ndConjunction Is Nothing Then
                                sbFilter.Append(" AND ")
                            Else
                                If ndConjunction.Value = CStr(ListFilterConjunction.LFC_Or) Then
                                    sbFilter.Append(" OR ")
                                Else
                                    sbFilter.Append(" AND ")
                                End If
                            End If
                        End If

                        If bFirstAttribute Then
                            bFirstAttribute = False
                            sbFilter.Append(" (")
                        End If

                        sAttributeValue = ndFilterAttribute.InnerXml
                        sAttributeValue = sAttributeValue.Replace("'", "''")

                        iOpenParentheses = GetIntegerXMLElementData("OpenParentheses", ndFilterAttribute, 0)
                        iCloseParentheses = GetIntegerXMLElementData("CloseParentheses", ndFilterAttribute, 0)
                        If iOpenParentheses > 0 Then
                            sbFilter.Append(CChar("("), iOpenParentheses)
                        End If

                        sbFilter.Append(ndFilterAttribute.Name)

                        If sAttributeValue.Length = 0 Then
                            sbFilter.Append(" IS NULL")
                        Else
                            ndComparisonType = CType(ndFilterAttribute.Attributes.GetNamedItem("ComparisonType"), Xml.XmlAttribute)
                            If ndComparisonType Is Nothing Then
                                eComparisonType = ListFilterComparisonType.Equals
                            Else
                                eComparisonType = CType(ndComparisonType.Value, abComponentServices.ComponentServices.ListFilterComparisonType)
                            End If
                            AppendComparisonType(sbFilter, eComparisonType, sAttributeValue)

                            If iCloseParentheses > 0 Then
                                sbFilter.Append(CChar(")"), iCloseParentheses)
                            End If
                        End If
                    End If
                Next iFilterAttribute

                sbFilter.Append(") ")
            End If
        End If

    End Sub

    Private Sub AppendFilterAttributes( _
            ByRef ndlFilters As Xml.XmlNodeList, _
            ByRef sbFilter As StringBuilder)

        Dim sAttributeValue As String
        Dim bFirstAttribute As Boolean = True
        Dim eComparisonType As abComponentServices.ComponentServices.ListFilterComparisonType
        Dim ndFilter As Xml.XmlNode
        Dim eConjunction As abComponentServices.ComponentServices.ListFilterConjunction
        Dim iOpenParentheses As Int32
        Dim iCloseParentheses As Int32

        If ndlFilters.Count > 0 Then
            For Each ndFilter In ndlFilters
                If sbFilter Is Nothing Then
                    sbFilter = New StringBuilder(" WHERE ")
                Else
                    eConjunction = CType(GetIntegerXMLElementData("Conjunction", ndFilter, abComponentServices.ComponentServices.ListFilterConjunction.LFC_And), abComponentServices.ComponentServices.ListFilterConjunction)
                    Select Case eConjunction
                        Case ListFilterConjunction.LFC_And
                            sbFilter.Append(" AND ")

                        Case ListFilterConjunction.LFC_Or
                            sbFilter.Append(" OR ")

                    End Select
                End If

                If bFirstAttribute Then
                    bFirstAttribute = False
                    sbFilter.Append(" (")
                End If

                sAttributeValue = GetStringXMLElementData("ComparisonData", ndFilter, "")
                sAttributeValue = sAttributeValue.Replace("'", "''")

                iOpenParentheses = GetIntegerXMLElementData("OpenParentheses", ndFilter, 0)
                iCloseParentheses = GetIntegerXMLElementData("CloseParentheses", ndFilter, 0)
                If iOpenParentheses > 0 Then
                    sbFilter.Append(CChar("("), iOpenParentheses)
                End If

                sbFilter.Append(GetStringXMLElementData("Name", ndFilter))

                If sAttributeValue.Length = 0 Then
                    sbFilter.Append(" IS NULL")
                Else
                    eComparisonType = CType(GetIntegerXMLElementData("ComparisonType", ndFilter, abComponentServices.ComponentServices.ListFilterComparisonType.Equals), abComponentServices.ComponentServices.ListFilterComparisonType)
                    AppendComparisonType(sbFilter, eComparisonType, sAttributeValue)
                End If

                If iCloseParentheses > 0 Then
                    sbFilter.Append(CChar(")"), iCloseParentheses)
                End If
            Next ndFilter

            sbFilter.Append(") ")
        End If

    End Sub

    Private Sub AppendComparisonType( _
            ByVal sbFilter As StringBuilder, _
            ByVal eComparisonType As abComponentServices.ComponentServices.ListFilterComparisonType, _
            ByVal sComparisonData As String)

        Dim sArray As String()
        Dim iCount As Int32
        Dim iTop As Int32
        Dim sbList As StringBuilder

        Select Case eComparisonType
            Case ListFilterComparisonType.Equals
                sbFilter.Append(" = ")
                sbFilter.Append("'")
                sbFilter.Append(sComparisonData)

            Case ListFilterComparisonType.GreaterThanOrEqualTo
                sbFilter.Append(" >= ")
                sbFilter.Append("'")
                sbFilter.Append(sComparisonData)

            Case ListFilterComparisonType.LessThanOrEqualTo
                sbFilter.Append(" <= ")
                sbFilter.Append("'")
                sbFilter.Append(sComparisonData)

            Case ListFilterComparisonType.NotEquals
                sbFilter.Append(" <> ")
                sbFilter.Append("'")
                sbFilter.Append(sComparisonData)

            Case ListFilterComparisonType.StartsWith
                sbFilter.Append(" LIKE '")
                sbFilter.Append(sComparisonData)
                sbFilter.Append("%")

            Case ListFilterComparisonType.Contains
                sbFilter.Append(" LIKE '%")
                sbFilter.Append(sComparisonData)
                sbFilter.Append("%")

            Case ListFilterComparisonType.EndsWith
                sbFilter.Append(" LIKE '%")
                sbFilter.Append(sComparisonData)

            Case ListFilterComparisonType.GreaterThan
                sbFilter.Append(" > ")
                sbFilter.Append("'")
                sbFilter.Append(sComparisonData)

            Case ListFilterComparisonType.LessThan
                sbFilter.Append(" < ")
                sbFilter.Append("'")
                sbFilter.Append(sComparisonData)

            Case ListFilterComparisonType.BelongsTo, ListFilterComparisonType.ExcludedFrom
                'wrap each element in the list in single quotes thus: WHERE Field IN ('1', '2', '3')
                'NB problems may arise if a list element itself contains a comma
                sArray = sComparisonData.Split(CChar(","))
                iTop = sArray.Length - 1
                sbList = New StringBuilder
                For iCount = 0 To iTop
                    If iCount = 0 Then
                        sbList.Append("'")
                    Else
                        sbList.Append(", '")
                    End If

                    sbList.Append(sArray(iCount))
                    sbList.Append("'")
                Next

                If eComparisonType = ListFilterComparisonType.BelongsTo Then
                    sbFilter.Append(" IN (")
                Else
                    sbFilter.Append(" NOT IN (")
                End If

                sbFilter.Append(sbList.ToString)
                sbFilter.Append(")")

            Case ListFilterComparisonType.Exists
                sbFilter.Append(" IS NOT NULL")

            Case ListFilterComparisonType.NotExists
                sbFilter.Append(" IS NULL")

            Case ListFilterComparisonType.BitwiseEquals
                sbFilter.Append(" & ")
                sbFilter.Append(sComparisonData)
                sbFilter.Append(" = ")
                sbFilter.Append(sComparisonData)

            Case ListFilterComparisonType.BitwiseNotEquals
                sbFilter.Append(" & ")
                sbFilter.Append(sComparisonData)
                sbFilter.Append(" <> ")
                sbFilter.Append(sComparisonData)

            Case Else
                sbFilter.Append(" = ")
                sbFilter.Append("'")
                sbFilter.Append(sComparisonData)

        End Select

        If eComparisonType < ListFilterComparisonType.BelongsTo Or eComparisonType > ListFilterComparisonType.NotExists Then
            sbFilter.Append("'")
        End If

    End Sub

    Private Function IsFilterAttributeExcluded( _
            ByVal sName As String) As Boolean

        Dim bExcluded As Boolean

        Select Case sName
            Case "RecordControl"
                bExcluded = True

            Case Else
                bExcluded = False

        End Select

        Return bExcluded

    End Function

    Public Shared Function AppendNode( _
            ByRef ndTarget As Xml.XmlNode, _
            ByRef ndSource As Xml.XmlNode) As Xml.XmlNode

        Return AppendNode(ndTarget, ndSource, True)

    End Function

    Public Shared Function AppendNode( _
            ByRef ndTarget As Xml.XmlNode, _
            ByRef ndSource As Xml.XmlNode, _
            ByVal bDeep As Boolean) As Xml.XmlNode

        Dim docParent As Xml.XmlDocument
        Dim ndImport As Xml.XmlNode

        If ndSource Is Nothing OrElse ndTarget Is Nothing Then
            Return Nothing
        End If

        docParent = ndTarget.OwnerDocument
        ndImport = docParent.ImportNode(ndSource, bDeep)
        Call ndTarget.AppendChild(ndImport)

        Return ndImport
    End Function


#End Region

#Region "ListHistory Methods"
    Public Function ListHistoryParameterCollection( _
        ByVal sComponentName As String, _
        ByVal sClassName As String, _
        ByVal sFilter As String) As ParameterCollection

        Return ListParameterCollection(sComponentName, sClassName & "History", sFilter)
    End Function

    Public Function ListHistoryParameterCollection( _
        ByVal sComponentName As String, _
        ByVal sClassName As String, _
        ByVal bCache As Boolean) As ParameterCollection

        Return ListParameterCollection(sComponentName, sClassName & "History", bCache)
    End Function

    Public Function ListHistoryParameterCollection( _
        ByVal sComponentName As String, _
        ByVal sClassName As String, _
        ByVal sFilter As String, _
        ByVal bCache As Boolean) As ParameterCollection

        Return ListParameterCollection(sComponentName, sClassName & "History", sFilter, bCache)
    End Function

    Public Function ListHistoryParameterCollection( _
        ByVal sComponentName As String, _
        ByVal sClassName As String, _
        ByVal sFilter As String, _
        ByVal sFields As String) As ParameterCollection

        Return ListParameterCollection(sComponentName, sClassName & "History", sFilter, sFields)
    End Function

    Public Function ListHistoryParameterCollection( _
        ByVal sComponentName As String, _
        ByVal sClassName As String, _
        ByVal sFilter As String, _
        ByVal sFields As String, _
        ByVal bCache As Boolean) As ParameterCollection

        Return ListParameterCollection(sComponentName, sClassName & "History", sFilter, sFields, bCache)
    End Function

    Public Function ListHistoryParameterCollection( _
            ByVal sComponentName As String, _
            ByVal sClassName As String, _
            ByVal sFilter As String, _
            ByVal sFields As String, _
            ByVal sSort As String, _
            ByVal bCache As Boolean) As ParameterCollection

        Return ListParameterCollection(sComponentName, sClassName & "History", sFilter, sFields, sSort, bCache)
    End Function

    Public Function ListHistory( _
        ByVal sComponentName As String, _
        ByVal sClassName As String, _
        ByVal sFilter As String) As Xml.XmlDocument

        Return List(sComponentName, sClassName & "History", sFilter)
    End Function

    Public Function ListHistory( _
        ByVal sComponentName As String, _
        ByVal sClassName As String, _
        ByVal sFilter As String, _
        ByVal sFields As String) As Xml.XmlDocument

        Return List(sComponentName, sClassName & "History", sFilter, sFields)
    End Function

    Public Function ListHistory( _
        ByVal sComponentName As String, _
        ByVal sClassName As String, _
        ByVal sFilter As String, _
        ByVal sFields As String, _
        ByVal bCache As Boolean) As Xml.XmlDocument

        Return List(sComponentName, sClassName & "History", sFilter, sFields, bCache)
    End Function

    Public Function ListHistory( _
        ByVal sComponentName As String, _
        ByVal sClassName As String, _
        ByVal sFilter As String, _
        ByVal sFields As String, _
        ByRef ndFilter As Xml.XmlNode, _
        ByVal bCache As Boolean) As Xml.XmlDocument

        Return List(sComponentName, sClassName & "History", sFilter, sFields, ndFilter, bCache)
    End Function

    Public Function ListHistory( _
        ByVal sComponentName As String, _
        ByVal sClassName As String, _
        ByVal sFilter As String, _
        ByVal sFields As String, _
        ByRef ndFilter As Xml.XmlNode, _
        ByRef ndListHeader As Xml.XmlNode, _
        ByVal bCache As Boolean) As Xml.XmlDocument

        Return List(sComponentName, sClassName & "History", sFilter, sFields, ndFilter, ndListHeader, bCache)
    End Function

    Public Function ListHistory( _
        ByVal sComponentName As String, _
        ByVal sClassName As String, _
        ByVal sFilter As String, _
        ByVal sFields As String, _
        ByRef ndFilter As Xml.XmlNode, _
        ByRef ndListHeader As Xml.XmlNode, _
        ByVal bCache As Boolean, _
        ByVal sRootName As String) As Xml.XmlDocument

        Return List(sComponentName, sClassName & "History", sFilter, sFields, ndFilter, ndListHeader, bCache, sRootName)
    End Function

    Public Function ListHistory( _
        ByVal sComponentName As String, _
        ByVal sClassName As String, _
        ByVal sFilter As String, _
        ByVal sFields As String, _
        ByRef ndFilter As Xml.XmlNode, _
        ByRef ndListHeader As Xml.XmlNode, _
        ByVal bCache As Boolean, _
        ByVal sRootName As String, _
        ByVal sRowDescription As String) As Xml.XmlDocument

        Return List(sComponentName, sClassName & "History", sFilter, sFields, ndFilter, ndListHeader, bCache, sRootName, sRowDescription)
    End Function

    Public Sub ListHistory( _
        ByVal sComponentName As String, _
        ByVal sClassName As String, _
        ByVal sFilter As String, _
        ByVal sFields As String, _
        ByRef ndFilter As Xml.XmlNode, _
        ByRef ndListHeader As Xml.XmlNode, _
        ByVal ndParent As Xml.XmlNode, _
        ByVal sRowDescription As String)

        List(sComponentName, sClassName & "History", sFilter, sFields, ndFilter, ndListHeader, ndParent, sRowDescription)
    End Sub
#End Region

#Region "Transaction Handling"

    Public Sub StartTransaction()

        m_oDatabaseAdapter.StartTransaction()

        m_bIsInTransaction = True
        m_bTransactionAborted = False

    End Sub

    Public Overloads Sub AbortTransaction()

        m_bTransactionAborted = True

    End Sub

    Public Overloads Sub AbortTransaction(ByVal force As Boolean)

        m_bTransactionAborted = True

        If force = True Then
            If m_bIsInTransaction Then
                m_oDatabaseAdapter.AbortTransaction()
                m_bTransactionAborted = False
                m_bIsInTransaction = False
            End If
        End If

    End Sub

    Public Sub CommitTransaction()

        'Can only commit transaction if transaction is open 
        If m_bIsInTransaction Then
            'Ensure that only successful transactions are committed
            If m_bTransactionAborted = False Then
                m_oDatabaseAdapter.CommitTransaction()

                ' 'Commit' old-style GL writes (done to Forward Posting table)
                GetRDOInterop()
                If Not m_oRDOInterop Is Nothing Then
                    Try
                        CallByName(m_oRDOInterop, "GLWriteCommit", CallType.Method)
                    Catch ex As System.Exception
                        ComponentServices.LogError(False, "ComponentServices", ex, "Calling m_oRDOInterop.GLWriteCommit")
                    End Try
                End If

                m_bIsInTransaction = False
            End If
        End If

    End Sub

#End Region

#Region "XML Response Build Methods"

    Private Function AppendDataToXMLDocument( _
            ByVal oData As Data.DataTable, _
            ByVal sComponentName As String, _
            ByVal sClassName As String, _
            ByRef ndParent As Xml.XmlNode, _
            ByVal sRowDescription As String) As Boolean

        Dim oAttributeNameMappings As Data.DataTable
        Dim iMappingAttributes As Integer
        Dim iMappingAttribute As Integer
        Dim pcMappingExclusions As ParameterCollection
        Dim oAttributeNameMapping As Data.DataRow
        Dim ndDataDescription As Xml.XmlNode
        Dim ndRowParent As Xml.XmlNode
        Dim iFields As Integer
        Dim iField As Integer
        Dim ndAttribute As Xml.XmlNode
        Dim sAttributeName As String
        Dim bResult As Boolean = False
        Dim iRow As Integer
        Dim iRows As Integer
        Dim oRow As Data.DataRow
        Dim oRows As Data.DataRowCollection
        Dim oColumn As DataColumn

        'AFS Get any specifically mapped fields first.
        pcMappingExclusions = New ParameterCollection
        oAttributeNameMappings = GetAttributeNameMappings(sComponentName, sClassName)
        If oAttributeNameMappings Is Nothing Then
            iMappingAttributes = -1
        Else
            iMappingAttributes = oAttributeNameMappings.Rows.Count - 1
        End If
        oRows = oData.Rows
        iRows = oRows.Count - 1
        For iRow = 0 To iRows
            oRow = oRows.Item(iRow)

            bResult = True

            If sRowDescription.Length > 0 Then
                ndDataDescription = ndParent.OwnerDocument.CreateElement(sRowDescription)
                ndParent.AppendChild(ndDataDescription)
                ndRowParent = ndDataDescription
            Else
                ndRowParent = ndParent
            End If

            For iMappingAttribute = 0 To iMappingAttributes
                oAttributeNameMapping = oAttributeNameMappings.Rows.Item(iMappingAttribute)
                Select Case CStr(oAttributeNameMapping.Item("DataType")).ToUpper()
                    Case "FREQUENCY", "PERIOD"
                        AppendFrequencyNodeFromDatabaseFields( _
                                ndRowParent, _
                                CStr(oAttributeNameMapping.Item("AttributeName")), _
                                CStr(oAttributeNameMapping.Item("DatabaseFields")), _
                                pcMappingExclusions, _
                                oRow)

                    Case "AMOUNT"
                        AppendAmountNodeFromDatabaseFields( _
                                ndRowParent, _
                                CStr(oAttributeNameMapping.Item("AttributeName")), _
                                CStr(oAttributeNameMapping.Item("DatabaseFields")), _
                                pcMappingExclusions, _
                                oRow)

                    Case Else
                        If Not oRow.IsNull(iField) Then
                            ndAttribute = ndParent.OwnerDocument.CreateElement(CStr(oAttributeNameMapping.Item("AttributeName")))
                            ndRowParent.AppendChild(ndAttribute)

                            SetXMLNodeTypeFromDataRow( _
                                    ndAttribute, _
                                    oRow, _
                                    CInt(oRow.Item(CStr(oAttributeNameMapping.Item("DatabaseFields")))))
                        End If

                End Select
            Next iMappingAttribute

            'AFS Process all other fields.
            iFields = oData.Columns.Count - 1
            For iField = 0 To iFields
                oColumn = oData.Columns.Item(iField)
                sAttributeName = oColumn.ColumnName
                If pcMappingExclusions.GetValue(sAttributeName) Is Nothing Then
                    If Not oRow.IsNull(iField) Then
                        If sAttributeName.ToUpper() <> "TIMESTAMP" Then
                            ndAttribute = ndParent.OwnerDocument.CreateElement(sAttributeName)
                            ndRowParent.AppendChild(ndAttribute)

                            SetXMLNodeTypeFromDataRow( _
                                    ndAttribute, _
                                    oRow, _
                                    iField)
                        End If
                    End If
                End If
            Next
        Next iRow

        Return bResult

    End Function

    Private Function AppendDataToXMLDocument( _
            ByVal sQuery As String, _
            ByVal sComponentName As String, _
            ByVal sClassName As String, _
            ByRef ndParent As Xml.XmlNode, _
            ByVal sRowDescription As String) As Boolean

        Dim oData As Data.DataTable

        oData = m_oDatabaseAdapter.ExecuteQueryDataTable(sQuery)

        Return AppendDataToXMLDocument( _
                oData, _
                sComponentName, _
                sClassName, _
                ndParent, _
                sRowDescription)

    End Function

    Private Sub SetXMLNodeTypeFromDataRow( _
            ByRef ndAttribute As Xml.XmlNode, _
            ByVal oDataRow As Data.DataRow, _
            ByVal iField As Integer)

        Dim sTimestamp As String = String.Empty
        Dim iTimestampOrdinal As Integer
        Dim decValue As Decimal
        Dim ndCData As System.Xml.XmlCDataSection
        Dim ndNode As Xml.XmlNode

        Select Case oDataRow.Item(iField).GetType().FullName()
            Case "System.Int32", "System.Int16"
                ndAttribute.InnerXml = CStr(oDataRow.Item(iField))

            Case "System.Byte"
                ndAttribute.InnerXml = CStr(oDataRow.Item(iField))

            Case "System.Boolean"
                ndAttribute.InnerXml = GetISODataValue(CBool(oDataRow.Item(iField)))

            Case "System.DateTime"
                Call AppendDateNode(ndAttribute, CDate(oDataRow.Item(iField)))

            Case "System.Decimal"
                If ndAttribute.Name = "ID" Then
                    If oDataRow.Item("Timestamp") Is System.DBNull.Value Then
                        sTimestamp = ""
                    Else
                        If m_structConfigSettings.eDatabaseVendor = DatabaseVendor.SQLServer Then
                            sTimestamp = m_oDatabaseAdapter.ConvertTimestampToString(CType(oDataRow.Item("Timestamp"), Byte()))
                        Else
                            sTimestamp = ""
                        End If
                    End If
                    AppendIDNode(ndAttribute, Nothing, CDec(oDataRow.Item(iField)), sTimestamp)
                Else
                    ndAttribute.InnerText = CStr(CDbl(CDec(oDataRow.Item(iField))))
                End If

            Case "System.Double"
                ndAttribute.InnerText = CStr(CDbl(oDataRow.Item(iField)))

            Case "System.String"
                ndCData = ndAttribute.OwnerDocument.CreateCDataSection(CStr(oDataRow.Item(iField)))
                ndAttribute.AppendChild(ndCData)

            Case "System.Guid"
                ndAttribute.InnerText = oDataRow.Item(iField).ToString

            Case Else
                ndCData = ndAttribute.OwnerDocument.CreateCDataSection(CStr(oDataRow.Item(iField)))
                ndAttribute.AppendChild(ndCData)

        End Select

    End Sub

    Private Sub AppendFrequencyNodeFromDatabaseFields( _
            ByRef ndParent As Xml.XmlNode, _
            ByVal strAttributeName As String, _
            ByVal strDatabaseFields As String, _
            ByRef oMappingExclusions As ParameterCollection, _
            ByRef oData As Data.DataRow)

        Dim strArrayDatabaseFields() As String
        Dim iDatabaseFields As Integer
        Dim iDatabaseField As Integer

        strArrayDatabaseFields = strDatabaseFields.Split(CChar(","))

        Try
            AppendFrequencyNode( _
                    ndParent, _
                    strAttributeName, _
                    CInt(oData.Item(strArrayDatabaseFields(0).Trim())), _
                    CInt(oData.Item(strArrayDatabaseFields(1).Trim())), _
                    CInt(oData.Item(strArrayDatabaseFields(2).Trim())), _
                    CInt(oData.Item(strArrayDatabaseFields(3).Trim())))
        Catch
            'AFS Attribute might not have been returned from database.
        End Try

        AddDatabaseFieldsToMappingExclusions( _
                strArrayDatabaseFields, _
                oMappingExclusions)

    End Sub

    Private Sub AppendAmountNodeFromDatabaseFields( _
            ByRef ndParent As Xml.XmlNode, _
            ByVal strAttributeName As String, _
            ByVal strDatabaseFields As String, _
            ByRef oMappingExclusions As ParameterCollection, _
            ByRef oData As Data.DataRow)

        Dim strArrayDatabaseFields() As String

        strArrayDatabaseFields = strDatabaseFields.Split(CChar(","))

        If strArrayDatabaseFields.Length = 1 Then
            Try
                AppendAmountNode( _
                        ndParent, _
                        strAttributeName, _
                        CDbl(oData.Item(strArrayDatabaseFields(0).Trim())), _
                        "")
            Catch
                'AFS Attribute might not have been returned from database.
            End Try
        ElseIf strArrayDatabaseFields.Length = 2 Then
            Try
                AppendAmountNode( _
                        ndParent, _
                        strAttributeName, _
                        CDbl(oData.Item(strArrayDatabaseFields(0).Trim())), _
                        CStr(oData.Item(strArrayDatabaseFields(1).Trim())))
            Catch
                'AFS Attribute might not have been returned from database.
            End Try
        Else
            Try
                AppendAmountNode( _
                        ndParent, _
                        strAttributeName, _
                        CDbl(oData.Item(strArrayDatabaseFields(0).Trim())), _
                        CStr(oData.Item(strArrayDatabaseFields(1).Trim())), _
                        CDbl(oData.Item(strArrayDatabaseFields(2).Trim())))
            Catch
                'AFS Attribute might not have been returned from database.
            End Try
        End If

        AddDatabaseFieldsToMappingExclusions( _
            strArrayDatabaseFields, _
            oMappingExclusions)

    End Sub

    Private Sub AddDatabaseFieldsToMappingExclusions( _
            ByVal strArrayDatabaseFields() As String, _
            ByRef oMappingExclusions As ParameterCollection)

        Dim iDatabaseFields As Integer
        Dim iDatabaseField As Integer

        iDatabaseFields = strArrayDatabaseFields.GetUpperBound(0)
        For iDatabaseField = 0 To iDatabaseFields
            oMappingExclusions.Add(strArrayDatabaseFields(iDatabaseField).Trim(), strArrayDatabaseFields(iDatabaseField))
        Next

    End Sub

    Private Function AttributeIsID( _
            ByVal strAttributeName As String) As Boolean

        Dim bIsID As Boolean

        strAttributeName = strAttributeName.ToUpper()
        If strAttributeName.EndsWith("ID") Then
            bIsID = True
        Else
            bIsID = False
        End If

        Return bIsID

    End Function

    Private Function GetAttributeNameMappings( _
            ByVal sComponentName As String, _
            ByVal sClassName As String) As Data.DataTable

        Dim sbMappingTableQuery As StringBuilder
        Dim oAttributeNameMappings As Data.DataTable
        Dim sMappingTableQuery As String
        Dim sObjectName As String

        sObjectName = GetObjectAlias(sComponentName & "_AttributeNameMappingsTbl")

        sbMappingTableQuery = New StringBuilder("SELECT * FROM ")
        sbMappingTableQuery.Append(sObjectName)
        sbMappingTableQuery.Append(" WHERE DatabaseClassName = '")
        sbMappingTableQuery.Append(sClassName)
        sbMappingTableQuery.Append("'")
        sMappingTableQuery = sbMappingTableQuery.ToString()

        Try
            oAttributeNameMappings = RetrieveDataTableFromCache(sComponentName, sClassName, sMappingTableQuery)
            If oAttributeNameMappings Is Nothing Then
                oAttributeNameMappings = m_oDatabaseAdapter.ExecuteQueryDataTable(sMappingTableQuery)
                If Not (oAttributeNameMappings Is Nothing) Then
                    AddDataToCache(sComponentName, sClassName, sMappingTableQuery, oAttributeNameMappings)
                End If
            End If
        Catch
            oAttributeNameMappings = Nothing
        End Try

        Return oAttributeNameMappings

    End Function

    Private Function GetDefaultXMLResponseRootName( _
            ByVal sClassName As String, _
            ByVal sAction As String) As String

        Dim sbDefaultResponseRootName As New StringBuilder(sAction)

        sbDefaultResponseRootName.Append(sClassName)
        sbDefaultResponseRootName.Append("Rs")

        Return sbDefaultResponseRootName.ToString()

    End Function

    Private Sub AppendIDNode( _
            ByRef ndParent As Xml.XmlNode, _
            ByVal sName As String, _
            ByVal decID As Decimal, _
            ByVal sTimestamp As String)

        Dim ndID As Xml.XmlElement
        Dim docParent As Xml.XmlDocument

        docParent = ndParent.OwnerDocument

        If Not (sName Is Nothing) Then
            If sName.Length <> 0 Then
                ndID = docParent.CreateElement(sName)
                ndParent.AppendChild(ndID)
            End If
        Else
            ndID = CType(ndParent, Xml.XmlElement)
        End If

        ndID.InnerText = CStr(decID)
        If Not (sTimestamp Is Nothing) Then
            If sTimestamp.Length <> 0 Then
                ndID.SetAttribute("Timestamp", sTimestamp)
            End If
        End If

    End Sub

    Public Shared Function CopyXMLNode(ByRef DestinationParentNode As Xml.XmlNode, ByRef SourceNode As Xml.XmlNode, ByVal Deep As Boolean, ByVal CopyChildNodesOnly As Boolean) As Xml.XmlNode

        Dim docDestination As Xml.XmlDocument
        Dim docSource As Xml.XmlDocument
        Dim ndCopy As Xml.XmlNode
        Dim attSource As Xml.XmlAttribute
        Dim attDestination As Xml.XmlAttribute
        Dim ndlChildren As Xml.XmlNodeList
        Dim i As Integer
        Dim ndChild As Xml.XmlNode

        docDestination = DestinationParentNode.OwnerDocument
        docSource = SourceNode.OwnerDocument

        If Not CopyChildNodesOnly Then
            ndCopy = docDestination.CreateElement(SourceNode.Name)

            For i = 0 To SourceNode.Attributes.Count - 1
                attSource = SourceNode.Attributes(i)

                attDestination = docDestination.CreateAttribute(attSource.Name)
                attDestination.InnerText = attSource.InnerText
                ndCopy.Attributes.Append(attDestination)
            Next

            DestinationParentNode.AppendChild(ndCopy)
        Else
            ndCopy = DestinationParentNode
        End If

        ndlChildren = SourceNode.SelectNodes("*")
        If Deep And ndlChildren.Count > 0 Then
            For Each ndChild In ndlChildren
                CopyXMLNode(ndCopy, ndChild, Deep, False)
            Next
        Else
            ndCopy.InnerText = SourceNode.InnerText
        End If

        Return ndCopy

    End Function

    Private Function CompareXMLNodes( _
            ByRef ndSource As Xml.XmlNode, _
            ByRef ndTarget As Xml.XmlNode) As Boolean

        Dim bMatch As Boolean = True
        Dim ndlNodes As Xml.XmlNodeList
        Dim ndNode As Xml.XmlNode
        Dim sSourceValue As String
        Dim iNodes As Integer
        Dim iNode As Integer
        Dim sTargetValue As String
        Dim ndTargetNode As Xml.XmlNode
        Dim sNodeName As String
        Dim decSourceNodeID As Decimal

        ndlNodes = ndSource.SelectNodes("*")
        iNodes = ndlNodes.Count - 1
        iNode = 0
        While (iNode <= iNodes) And bMatch
            ndNode = ndlNodes.Item(iNode)
            sNodeName = ndNode.Name

            'TODO Remove ProductDefaultID.
            If sNodeName <> "ParentID" And sNodeName <> "ProductDefaultID" And sNodeName <> "ProductSelectionCriteriaID" And sNodeName <> "ProductDefaultAccountID" And sNodeName <> "SchemeID" And sNodeName <> "ID" And sNodeName <> "LoanCreditValidationID" Then
                ndTargetNode = ndTarget.SelectSingleNode(sNodeName)
                If ndTargetNode Is Nothing Then
                    If ndNode.InnerText.Length = 0 Then
                        bMatch = True
                    Else
                        'AFS This is nasty (the select case). Temporary fix 6.3 until NUI fixed.
                        Select Case sNodeName
                            Case "Hour", "Minute", "Second", "Fraction"
                                If ndNode.InnerText = "0" Then
                                    bMatch = True
                                Else
                                    bMatch = False
                                End If

                            Case Else
                                bMatch = False

                        End Select
                    End If
                Else
                    If ndNode.SelectNodes("*").Count > 0 Then
                        decSourceNodeID = GetDecimalXMLElementData("ID", ndNode, 0)
                        If decSourceNodeID > 0 Then
                            If sNodeName <> "Loan" And sNodeName <> "Account" Then
                                ndTargetNode = ndTarget.SelectSingleNode(sNodeName + "[ID=" + decSourceNodeID.ToString() + "]")
                            End If
                            If ndTargetNode Is Nothing Then
                                bMatch = False
                            Else
                                bMatch = CompareXMLNodes(ndNode, ndTargetNode)
                            End If
                        Else
                            If ndSource.SelectNodes(sNodeName).Count > 1 Or ndTarget.SelectNodes(sNodeName).Count > 1 Then
                                If ndSource.SelectNodes(sNodeName).Count <> ndTarget.SelectNodes(sNodeName).Count Then
                                    bMatch = False
                                End If
                            Else
                                bMatch = CompareXMLNodes(ndNode, ndTargetNode)
                            End If
                        End If
                    Else
                        sSourceValue = ndNode.InnerText.Trim()
                        sTargetValue = ndTargetNode.InnerText.Trim()
                        If sSourceValue <> sTargetValue Then
                            bMatch = False
                        End If
                    End If
                End If
            End If

            iNode = iNode + 1
        End While

        Return bMatch

    End Function

#End Region

#Region "XML Input Translation Methods"

    Private Function PopulateParameterCollectionFromXMLNode( _
            ByRef ndInput As Xml.XmlNode, _
            ByVal sComponentName As String, _
            ByVal sClassName As String, _
            ByRef pcDataTypes As ParameterCollection) As ParameterCollection

        Dim pcOutput As ParameterCollection
        Dim iParameters As Integer
        Dim iParameter As Integer
        Dim ndlParameters As Xml.XmlNodeList
        Dim oParameter As Object
        Dim sParameterName As String
        Dim ndParameter As Xml.XmlNode
        Dim oAttributeNameMappings As Data.DataTable
        Dim oAttributeNameMapping As Data.DataRow
        Dim pcMappingExclusions As ParameterCollection
        Dim iAttributeNameMappings As Integer
        Dim iAttributeNameMapping As Integer
        Dim sDataType As String
        Dim pcParameters As ParameterCollection
        Dim iDataType As Integer

        pcParameters = New ParameterCollection

        'Set any specifically mapped fields first.
        pcMappingExclusions = New ParameterCollection
        oAttributeNameMappings = GetAttributeNameMappings(sComponentName, sClassName)
        If oAttributeNameMappings Is Nothing Then
            iAttributeNameMappings = -1
        Else
            iAttributeNameMappings = oAttributeNameMappings.Rows.Count - 1
        End If
        For iAttributeNameMapping = 0 To iAttributeNameMappings
            oAttributeNameMapping = oAttributeNameMappings.Rows(iAttributeNameMapping)

            sParameterName = CStr(oAttributeNameMapping.Item("AttributeName"))
            ndParameter = ndInput.SelectSingleNode(sParameterName)
            If Not (ndParameter Is Nothing) Then
                If sParameterName = "ID" Then
                    If GetStringXMLAttributeData("Timestamp", ndParameter, String.Empty) <> String.Empty Then
                        pcParameters.Add("Timestamp", GetStringXMLAttributeData("Timestamp", ndParameter))
                    End If
                End If
                sDataType = CStr(oAttributeNameMapping.Item("DataType"))
                Select Case sDataType.ToUpper()
                    Case "AMOUNT"
                        SetDatabaseFieldsFromAmountNode( _
                                ndParameter, _
                                pcParameters, _
                                CStr(oAttributeNameMapping.Item("DatabaseFields")))

                    Case "FREQUENCY", "PERIOD"
                        SetDatabaseFieldsFromFrequencyNode( _
                                ndParameter, _
                                pcParameters, _
                                CStr(oAttributeNameMapping.Item("DatabaseFields")))

                    Case Else
                        pcParameters.Add(sParameterName, ndParameter, pcDataTypes.GetIntegerValue(sParameterName))

                End Select
            End If

            pcMappingExclusions.Add(sParameterName, sParameterName)
        Next iAttributeNameMapping

        ndlParameters = ndInput.SelectNodes("*")
        iParameters = ndlParameters.Count - 1
        For iParameter = 0 To iParameters
            ndParameter = ndlParameters(iParameter)
            sParameterName = ndParameter.Name

            If pcMappingExclusions.GetValue(sParameterName) Is Nothing Then
                iDataType = pcDataTypes.GetIntegerValue(sParameterName, 0)
                If iDataType <> 0 Then
                    pcParameters.Add(sParameterName, ndParameter, iDataType, True)
                End If
                If sParameterName = "ID" Then
                    If GetStringXMLAttributeData("Timestamp", ndParameter, String.Empty) <> String.Empty Then
                        pcParameters.Add("Timestamp", GetStringXMLAttributeData("Timestamp", ndParameter))
                    End If
                End If
            End If
        Next iParameter

        Return pcParameters

    End Function

    Private Sub SetDatabaseFieldsFromAmountNode( _
            ByRef ndAmount As Xml.XmlNode, _
            ByRef pcParameters As ParameterCollection, _
            ByVal sDatabaseFields As String)

        Dim dblAmt As Double
        Dim sCurCode As String
        Dim dblCurRate As Double
        Dim arrDatabaseFields() As String

        dblAmt = GetDoubleXMLElementData("Amt", ndAmount, Nothing)
        sCurCode = GetStringXMLElementData("CurCode", ndAmount, Nothing)
        dblCurRate = GetDoubleXMLElementData("CurRate", ndAmount, Nothing)

        arrDatabaseFields = sDatabaseFields.Split(CChar(","))

        pcParameters.Add(arrDatabaseFields(0), dblAmt)
        If arrDatabaseFields.Length > 1 Then
            pcParameters.Add(arrDatabaseFields(1), sCurCode)
        End If
        If arrDatabaseFields.Length > 2 Then
            pcParameters.Add(arrDatabaseFields(2), dblCurRate)
        End If

    End Sub

    Private Sub SetDatabaseFieldsFromFrequencyNode( _
            ByRef ndFrequency As Xml.XmlNode, _
            ByRef pcParameters As ParameterCollection, _
            ByVal sDatabaseFields As String)

        Dim iType As Integer
        Dim iPrec As Integer
        Dim iPrecType As Integer
        Dim iPrd As Integer

        Dim arrDatabaseFields() As String

        iType = GetIntegerXMLElementData("Type", ndFrequency, Nothing)
        iPrec = GetIntegerXMLElementData("Prec", ndFrequency, Nothing)
        iPrecType = GetIntegerXMLElementData("PrecType", ndFrequency, Nothing)
        iPrd = GetIntegerXMLElementData("Prd", ndFrequency, Nothing)

        arrDatabaseFields = sDatabaseFields.Split(CChar(","))

        pcParameters.Add(arrDatabaseFields(0), iType)
        pcParameters.Add(arrDatabaseFields(1), iPrec)
        pcParameters.Add(arrDatabaseFields(2), iPrecType)
        pcParameters.Add(arrDatabaseFields(3), iPrd)

    End Sub

    Public Shared Sub ConvertClassicFrequency(ByVal sFrequency As String, ByRef ndTarget As Xml.XmlNode)

        'TODO
        Dim sClassicFrequency() As String
        Dim sFrequencyType As String
        Dim sDay As String
        Dim sMonth As String
        Dim sRollType As String
        Dim sType As String
        Dim sPeriod As String = "0"
        Dim sPrecision As String = "0"
        Dim sPrecisionType As String = "0"

        If Not (ndTarget Is Nothing) Then
            'extract the Classic attributes
            sClassicFrequency = Split(sFrequency, ",")
            sFrequencyType = sClassicFrequency(0)
            sDay = sClassicFrequency(1)
            sMonth = sClassicFrequency(2)
            sRollType = sClassicFrequency(3)

            'convert them to new values
            sType = ConvertFrequencyType(sFrequencyType, True)
            Select Case sType
                Case eFrequencyType.Daily, eFrequencyType.Weekly, eFrequencyType.Monthly
                    sPeriod = "0"
                    Select Case sDay
                        Case "0"
                            sPrecisionType = ePrecisionType.None
                        Case "1"
                            sPrecisionType = ePrecisionType.Start
                        Case "99"
                            sPrecisionType = ePrecisionType.End
                        Case Else
                            sPrecisionType = ePrecisionType.Specific
                            sPeriod = sDay
                    End Select

                Case eFrequencyType.Quarterly, eFrequencyType.HalfYearly, eFrequencyType.Yearly
                    sPeriod = sMonth
                    sPrecision = "0"
                    Select Case sDay
                        Case "0"
                            sPrecisionType = ePrecisionType.None
                        Case "1"
                            sPrecisionType = ePrecisionType.Start
                        Case "99"
                            sPrecisionType = ePrecisionType.End
                        Case Else
                            sPrecisionType = ePrecisionType.Specific
                            sPrecision = sDay
                    End Select
            End Select

            'save them to the xml node
            AppendXMLElement("Type", ndTarget, sType)
            AppendXMLElement("Prd", ndTarget, sPeriod)
            AppendXMLElement("Prec", ndTarget, sPrecision)
            AppendXMLElement("PrecType", ndTarget, sPrecisionType)
        End If

    End Sub

    Public Shared Function ConvertClassicFrequency(ByVal sFrequency As String, ByRef ndParent As Xml.XmlNode, ByVal sNewFrequencyNodeName As String) As Xml.XmlNode

        Dim ndFrequency As Xml.XmlNode

        ndFrequency = AppendXMLElement(sNewFrequencyNodeName, ndParent)
        ConvertClassicFrequency(sFrequency, ndFrequency)

        Return ndFrequency

    End Function

    Public Shared Function ConvertClassicFrequency(ByVal ndFrequency As Xml.XmlNode) As String

        Dim sFreqLocal As String
        Dim sPeriod As String
        Dim sPrecision As String
        Dim sPrecisionType As String
        Dim sFrequencyType As String
        Dim sDay As String
        Dim sMonth As String
        Dim sRollType As String = "3"   'always "Actual"

        sFreqLocal = GetStringXMLElementData("Type", ndFrequency, "0")
        sPeriod = GetStringXMLElementData("Prd", ndFrequency, "0")
        sPrecision = GetStringXMLElementData("Prec", ndFrequency, "0")
        sPrecisionType = GetStringXMLElementData("PrecType", ndFrequency, "0")
        sFrequencyType = ConvertFrequencyType(sFreqLocal, False)

        Select Case sFreqLocal
            Case eFrequencyType.Daily, eFrequencyType.Weekly, eFrequencyType.Monthly
                sMonth = "0"
                Select Case sPrecisionType
                    Case ePrecisionType.Specific
                        sDay = sPeriod
                    Case ePrecisionType.Start
                        sDay = "1"
                    Case ePrecisionType.End
                        sDay = "99"
                    Case Else
                        sDay = "0"
                End Select

            Case eFrequencyType.Quarterly, eFrequencyType.HalfYearly, eFrequencyType.Yearly
                sMonth = sPeriod
                Select Case sPrecisionType
                    Case ePrecisionType.Specific
                        sDay = sPrecision
                    Case ePrecisionType.Start
                        sDay = "1"
                    Case ePrecisionType.End
                        sDay = "99"
                    Case Else
                        sDay = "0"
                End Select
            Case Else
                sMonth = "0"
                sDay = "0"
        End Select

        Return sFrequencyType & "," & sDay & "," & sMonth & "," & sRollType

    End Function

    Private Shared Function ConvertFrequencyType(ByVal sValue As String, ByVal bClassicToNew As Boolean) As String

        If bClassicToNew Then
            Select Case sValue
                Case eFrequencyType_e.None
                    Return eFrequencyType.None
                Case eFrequencyType_e.Daily
                    Return eFrequencyType.Daily
                Case eFrequencyType_e.Weekly
                    Return eFrequencyType.Weekly
                Case eFrequencyType_e.Monthly
                    Return eFrequencyType.Monthly
                Case eFrequencyType_e.Quarterly
                    Return eFrequencyType.Quarterly
                Case eFrequencyType_e.HalfYearly
                    Return eFrequencyType.HalfYearly
                Case eFrequencyType_e.Yearly
                    Return eFrequencyType.Yearly
                Case Else
                    Return eFrequencyType.None
            End Select
        Else
            Select Case sValue
                Case eFrequencyType.None
                    Return eFrequencyType_e.None
                Case eFrequencyType.Daily
                    Return eFrequencyType_e.Daily
                Case eFrequencyType.Weekly
                    Return eFrequencyType_e.Weekly
                Case eFrequencyType.Monthly
                    Return eFrequencyType_e.Monthly
                Case eFrequencyType.Quarterly
                    Return eFrequencyType_e.Quarterly
                Case eFrequencyType.HalfYearly
                    Return eFrequencyType_e.HalfYearly
                Case eFrequencyType.Yearly
                    Return eFrequencyType_e.Yearly
                Case Else
                    Return eFrequencyType_e.None
            End Select
        End If

    End Function


#End Region

#Region "Caching Methods"

    Public Sub InformApplicationServersOfDataChange( _
            ByVal strComponentName As String, _
            ByVal strClassName As String)

        Dim oApplicationServers As DataTable
        Dim bMoreServersToInform As Boolean
        Dim oService As New abService.InvokeService
        Dim sServiceRequest As String
        Dim bAsync As Boolean = True
        Dim oRequestDoc As Xml.XmlDocument
        Dim sServerTable As String


        Dim iFields As Integer
        Dim iField As Integer
        Dim iRow As Integer
        Dim iRows As Integer
        Dim oRow As Data.DataRow
        Dim oRows As Data.DataRowCollection


        Try
            If m_structConfigSettings.bCachingEnabled Then

                If m_structConfigSettings.eDatabaseVendor = DatabaseVendor.Oracle Then
                    sServerTable = GetObjectAlias("abComponentServices_ServersTbl")
                Else
                    sServerTable = "abComponentServices_ServersTbl"
                End If

                oApplicationServers = m_oDatabaseAdapter.ExecuteQueryDataTable("SELECT Name FROM " & sServerTable)

                If oApplicationServers Is Nothing Then
                    Call Me.RemoveDataFromCache(strComponentName, strClassName)
                Else
                    oRows = oApplicationServers.Rows
                    iRows = oRows.Count - 1
                    If oRows.Count > 0 Then
                        oRequestDoc = CreateXMLDocument("activebank", "Header", "Header/SessionId=" & m_sSessionID, _
                                "Header/BranchId=" & m_sBranchID, _
                                "Header/Timestamp", "Header/Timestamp/Year=" & Now.Year, "Header/Timestamp/Month=" & Now.Month, _
                                "Header/Timestamp/Day=" & Now.Day, "Header/Timestamp/Hour=" & Now.Hour, _
                                "Header/Timestamp/Minute=" & Now.Minute, "Header/Timestamp/Second=" & Now.Second, _
                                "Request", "Request/ManageDataCacheRq", "Request/ManageDataCacheRq/Action=RetireObjects", "Request/ManageDataCacheRq/ComponentName=" & strComponentName, "Request/ManageDataCacheRq/ClassName=" & strClassName)

                        sServiceRequest = oRequestDoc.OuterXml

                        For iRow = 0 To iRows
                            oRow = oRows.Item(iRow)

                            iFields = oApplicationServers.Columns.Count - 1
                            For iField = 0 To iFields
                                If oApplicationServers.Columns.Item(iField).ColumnName.ToUpper = "NAME" Then
                                    If Not oRow.IsNull(iField) Then
                                        oService.InvokeServiceRequest(sServiceRequest, oRow.Item(iField).ToString(), bAsync)
                                    End If
                                End If
                            Next
                        Next iRow
                    Else
                        Call Me.RemoveDataFromCache(strComponentName, strClassName)
                    End If
                End If
            End If

        Catch ex As ActiveBankException
            LogError(True, ex)
        Catch ex As System.Exception
            LogError(True, "abComponentServices", ex)
        Finally
            oService = Nothing
        End Try

    End Sub

    Public Function ManageDataCache(ByRef ndInput As Xml.XmlNode) As Xml.XmlDocument
        Dim docOutput As Xml.XmlDocument
        Dim sComponentName As String
        Dim sClassName As String
        Dim sActionName As String

        Try
            If Not ndInput.SelectSingleNode("Action") Is Nothing Then
                sActionName = GetStringXMLElementData("Action", ndInput, String.Empty)

                If sActionName = "RetireObjects" Then
                    sComponentName = GetStringXMLElementData("ComponentName", ndInput, String.Empty)
                    sClassName = GetStringXMLElementData("ClassName", ndInput, String.Empty)

                    Me.RemoveDataFromCache(sComponentName, sClassName)
                End If
            End If

            docOutput = CreateXMLDocument("ManageDataCacheRs", Nothing)

        Catch ex As ActiveBankException
            docOutput = CreateErrorDocument(ex)
        Catch ex As System.Exception
            docOutput = CreateErrorDocument("abComponentServices", ex)
        End Try

        Return docOutput
    End Function

    Public Sub RemoveDataFromCache( _
            ByVal strComponentName As String, _
            ByVal strClassName As String)

        Dim oClassData As Collections.Specialized.HybridDictionary
        Dim strClassKey As String

        If strComponentName.Length = 0 And strClassName.Length = 0 Then
            m_oCachedXMLDocuments.Clear()
            m_oCachedParameterCollections.Clear()
            m_oCachedDataTables.Clear()
        Else
            strClassKey = GetClassCachingKey( _
                    strComponentName, _
                    strClassName)

            m_oCachedXMLDocuments.Remove(strClassKey)
            m_oCachedParameterCollections.Remove(strClassKey)
            m_oCachedDataTables.Remove(strClassKey)
        End If
    End Sub

    Private Sub AddDataToCache( _
            ByVal strComponentName As String, _
            ByVal strClassName As String, _
            ByVal strKey As String, _
            ByRef oData As ParameterCollection)

        Dim oClassData As Collections.Specialized.HybridDictionary
        Dim strClassKey As String

        If m_structConfigSettings.bCachingEnabled Then
            strClassKey = GetClassCachingKey( _
                    strComponentName, _
                    strClassName)

            Try
                oClassData = CType(m_oCachedParameterCollections.Item(strClassKey), Collections.Specialized.HybridDictionary)
                If oClassData Is Nothing Then
                    oClassData = New Collections.Specialized.HybridDictionary
                    m_oCachedParameterCollections.Add(strClassKey, oClassData)
                End If

                oClassData.Add(strKey, oData)
            Catch
            End Try
        End If

    End Sub

    Private Sub AddDataToCache( _
            ByVal strComponentName As String, _
            ByVal strClassName As String, _
            ByVal strKey As String, _
            ByRef oData As Data.DataTable)

        Dim oClassData As Collections.Specialized.HybridDictionary
        Dim strClassKey As String

        If m_structConfigSettings.bCachingEnabled Then
            strClassKey = GetClassCachingKey( _
                    strComponentName, _
                    strClassName)

            Try
                oClassData = CType(m_oCachedDataTables.Item(strClassKey), Collections.Specialized.HybridDictionary)
                If oClassData Is Nothing Then
                    oClassData = New Collections.Specialized.HybridDictionary
                    m_oCachedDataTables.Add(strClassKey, oClassData)
                End If

                oClassData.Add(strKey, oData)
            Catch
            End Try
        End If

    End Sub

    Private Sub AddDataToCache( _
            ByVal sComponentName As String, _
            ByVal sClassName As String, _
            ByVal sKey As String, _
            ByRef sData As String)

        Dim oClassData As Collections.Specialized.HybridDictionary
        Dim sClassKey As String

        If m_structConfigSettings.bCachingEnabled Then
            sClassKey = GetClassCachingKey( _
                    sComponentName, _
                    sClassName)

            Try
                oClassData = CType(m_oCachedXMLDocuments.Item(sClassKey), Collections.Specialized.HybridDictionary)
                If oClassData Is Nothing Then
                    oClassData = New Collections.Specialized.HybridDictionary
                    m_oCachedXMLDocuments.Add(sClassKey, oClassData)
                End If

                oClassData.Add(sKey, sData)
            Catch
            End Try
        End If

    End Sub

    Private Function RetrieveXMLDocumentFromCache( _
            ByVal strComponentName As String, _
            ByVal strClassName As String, _
            ByVal strKey As String) As String

        Dim oClassData As Collections.Specialized.HybridDictionary
        Dim strClassKey As String
        Dim strXMLDocument As String

        If m_structConfigSettings.bCachingEnabled Then
            strClassKey = GetClassCachingKey( _
                    strComponentName, _
                    strClassName)

            oClassData = CType(m_oCachedXMLDocuments.Item(strClassKey), Collections.Specialized.HybridDictionary)
            If oClassData Is Nothing Then
                strXMLDocument = Nothing
            Else
                strXMLDocument = CStr(oClassData.Item(strKey))
            End If
        Else
            strXMLDocument = Nothing
        End If

        Return strXMLDocument

    End Function

    Private Function RetrieveDataTableFromCache( _
        ByVal strComponentName As String, _
        ByVal strClassName As String, _
        ByVal strKey As String) As Data.DataTable

        Dim oClassData As Collections.Specialized.HybridDictionary
        Dim strClassKey As String
        Dim oDataTable As Data.DataTable

        Try
            If m_structConfigSettings.bCachingEnabled Then
                strClassKey = GetClassCachingKey( _
                        strComponentName, _
                        strClassName)

                oClassData = CType(m_oCachedDataTables.Item(strClassKey), Collections.Specialized.HybridDictionary)
                If oClassData Is Nothing Then
                    oDataTable = Nothing
                Else
                    oDataTable = CType(oClassData.Item(strKey), Data.DataTable)
                End If
            Else
                oDataTable = Nothing
            End If
        Catch ex As System.Exception
            oDataTable = Nothing
        End Try

        Return oDataTable

    End Function

    Private Function RetrieveParameterCollectionFromCache( _
            ByVal strComponentName As String, _
            ByVal strClassName As String, _
            ByVal strKey As String) As ParameterCollection

        Dim oClassData As Collections.Specialized.HybridDictionary
        Dim strClassKey As String
        Dim oParameterCollection As ParameterCollection

        If m_structConfigSettings.bCachingEnabled Then
            strClassKey = GetClassCachingKey( _
                    strComponentName, _
                    strClassName)

            oClassData = CType(m_oCachedParameterCollections.Item(strClassKey), Collections.Specialized.HybridDictionary)
            If oClassData Is Nothing Then
                oParameterCollection = Nothing
            Else
                oParameterCollection = CType(oClassData.Item(strKey), ParameterCollection)
                If Not (oParameterCollection Is Nothing) Then
                    oParameterCollection.Reset()
                End If
            End If
        Else
            oParameterCollection = Nothing
        End If

        Return oParameterCollection

    End Function

    Private Function GetClassCachingKey( _
            ByVal strComponentName As String, _
            ByVal strClassName As String) As String

        Dim sbClassKey As New StringBuilder(strComponentName)

        sbClassKey.Append(":")
        sbClassKey.Append(strClassName)

        Return sbClassKey.ToString()

    End Function

#End Region

#Region "Private Utility Methods"

#End Region

#Region "activebank Methods"

    Public Function GetAppConfigSetting(ByVal sSettingName As String) As Object

        Dim docResult As Xml.XmlDocument
        Dim ndListSetting As Xml.XmlNode
        Dim ndSetting As Xml.XmlNode

        docResult = List("abAppConfig", "Setting", "ID = '" & sSettingName.Trim & "'", "SETTING_TYPE, SETTING_VALUE", True)

        If Not docResult Is Nothing Then
            ndListSetting = docResult.SelectSingleNode("ListSettingsRs")
            ndSetting = ndListSetting.SelectSingleNode("Setting")

            'Check to see if Loan Account information retrieved.
            If Not ndSetting Is Nothing Then
                Select Case ndSetting.SelectSingleNode("SETTING_TYPE").InnerText.ToUpper
                    Case "BOOLEAN"
                        Return CType(ndSetting.SelectSingleNode("SETTING_VALUE").InnerText, Boolean)

                    Case "CHAR", "STRING", "VARCHAR"
                        Return CType(ndSetting.SelectSingleNode("SETTING_VALUE").InnerText, String)

                    Case "LONG"
                        Return CType(ndSetting.SelectSingleNode("SETTING_VALUE").InnerText, Long)

                    Case "NUMBER", "INTEGER", "INT", "NUMERIC"
                        Return CType(ndSetting.SelectSingleNode("SETTING_VALUE").InnerText, Integer)

                    Case "DECIMAL", "FLOAT"
                        Return CType(ndSetting.SelectSingleNode("SETTING_VALUE").InnerText, Decimal)

                    Case Else
                        Return CType(ndSetting.SelectSingleNode("SETTING_VALUE").InnerText, String)
                End Select
            Else
                Return CType(Nothing, Object)
            End If
        Else
            Return CType(Nothing, Object)
        End If

    End Function

    Public Function GetAppConfigSetting(ByVal sSettingName As String, ByVal bDefaultValue As Boolean) As Boolean
        '----------------------------------------------------------------------------------
        ' About     : Get Boolean App Config Setting
        ' Created   : Jat Virdee, 10th June 2003
        '----------------------------------------------------------------------------------

        Dim docResult As Xml.XmlDocument
        Dim ndListSetting As Xml.XmlNode
        Dim ndSetting As Xml.XmlNode

        docResult = List("abAppConfig", "Setting", "ID = '" & sSettingName.Trim & "'", "SETTING_VALUE", True)

        If Not docResult Is Nothing Then
            ndListSetting = docResult.SelectSingleNode("ListSettingsRs")
            ndSetting = ndListSetting.SelectSingleNode("Setting")

            'Check to see if Loan Account information retrieved.
            If Not ndSetting Is Nothing Then
                Return CType(ndSetting.SelectSingleNode("SETTING_VALUE").InnerText, Boolean)
            Else
                Return bDefaultValue
            End If
        Else
            Return bDefaultValue
        End If
    End Function

    Public Function GetAppConfigSetting(ByVal sSettingName As String, ByVal cDefaultValue As Char) As Char
        '----------------------------------------------------------------------------------
        ' About     : Get Char App Config Setting
        ' Created   : Jat Virdee, 10th June 2003
        '----------------------------------------------------------------------------------

        Dim docResult As Xml.XmlDocument
        Dim ndListSetting As Xml.XmlNode
        Dim ndSetting As Xml.XmlNode

        docResult = List("abAppConfig", "Setting", "ID = '" & sSettingName.Trim & "'", "SETTING_VALUE", True)

        If Not docResult Is Nothing Then
            ndListSetting = docResult.SelectSingleNode("ListSettingsRs")
            ndSetting = ndListSetting.SelectSingleNode("Setting")

            'Check to see if Loan Account information retrieved.
            If Not ndSetting Is Nothing Then
                Return CType(ndSetting.SelectSingleNode("SETTING_VALUE").InnerText, Char)
            Else
                Return cDefaultValue
            End If
        Else
            Return cDefaultValue
        End If
    End Function

    Public Function GetAppConfigSetting(ByVal sSettingName As String, ByVal sDefaultValue As String) As String
        '----------------------------------------------------------------------------------
        ' About     : Get String App Config Setting
        ' Created   : Jat Virdee, 10th June 2003
        '----------------------------------------------------------------------------------

        Dim docResult As Xml.XmlDocument
        Dim ndListSetting As Xml.XmlNode
        Dim ndSetting As Xml.XmlNode

        docResult = List("abAppConfig", "Setting", "ID = '" & sSettingName.Trim & "'", "SETTING_VALUE", True)

        If Not docResult Is Nothing Then
            ndListSetting = docResult.SelectSingleNode("ListSettingsRs")
            ndSetting = ndListSetting.SelectSingleNode("Setting")

            'Check to see if Loan Account information retrieved.
            If Not ndSetting Is Nothing Then
                Return CType(ndSetting.SelectSingleNode("SETTING_VALUE").InnerText, String)
            Else
                Return sDefaultValue
            End If
        Else
            Return sDefaultValue
        End If
    End Function

    Public Function GetAppConfigSetting(ByVal sSettingName As String, ByVal dDefaultValue As Decimal) As Decimal
        '----------------------------------------------------------------------------------
        ' About     : Get Decimal App Config Setting
        ' Created   : Jat Virdee, 10th June 2003
        '----------------------------------------------------------------------------------

        Dim docResult As Xml.XmlDocument
        Dim ndListSetting As Xml.XmlNode
        Dim ndSetting As Xml.XmlNode

        docResult = List("abAppConfig", "Setting", "ID = '" & sSettingName.Trim & "'", "SETTING_VALUE", True)

        If Not docResult Is Nothing Then
            ndListSetting = docResult.SelectSingleNode("ListSettingsRs")
            ndSetting = ndListSetting.SelectSingleNode("Setting")

            'Check to see if Loan Account information retrieved.
            If Not ndSetting Is Nothing Then
                Return CType(ndSetting.SelectSingleNode("SETTING_VALUE").InnerText, Decimal)
            Else
                Return dDefaultValue
            End If
        Else
            Return dDefaultValue
        End If
    End Function

    Public Function GetAppConfigSetting(ByVal sSettingName As String, ByVal lDefaultValue As Long) As Long
        '----------------------------------------------------------------------------------
        ' About     : Get Long App Config Setting
        ' Created   : Jat Virdee, 10th June 2003
        '----------------------------------------------------------------------------------

        Dim docResult As Xml.XmlDocument
        Dim ndListSetting As Xml.XmlNode
        Dim ndSetting As Xml.XmlNode

        docResult = List("abAppConfig", "Setting", "ID = '" & sSettingName.Trim & "'", "SETTING_VALUE", True)

        If Not docResult Is Nothing Then
            ndListSetting = docResult.SelectSingleNode("ListSettingsRs")
            ndSetting = ndListSetting.SelectSingleNode("Setting")

            'Check to see if Loan Account information retrieved.
            If Not ndSetting Is Nothing Then
                Return CType(ndSetting.SelectSingleNode("SETTING_VALUE").InnerText, Long)
            Else
                Return lDefaultValue
            End If
        Else
            Return lDefaultValue
        End If
    End Function

    Public Function GetAppConfigSetting(ByVal sSettingName As String, ByVal iDefaultValue As Integer) As Integer
        '----------------------------------------------------------------------------------
        ' About     : Get integer App Config Setting
        ' Created   : Jat Virdee, 10th June 2003
        '----------------------------------------------------------------------------------

        Dim docResult As Xml.XmlDocument
        Dim ndListSetting As Xml.XmlNode
        Dim ndSetting As Xml.XmlNode

        docResult = List("abAppConfig", "Setting", "ID = '" & sSettingName.Trim & "'", "SETTING_VALUE", True)

        If Not docResult Is Nothing Then
            ndListSetting = docResult.SelectSingleNode("ListSettingsRs")
            ndSetting = ndListSetting.SelectSingleNode("Setting")

            'Check to see if Loan Account information retrieved.
            If Not ndSetting Is Nothing Then
                Return CType(ndSetting.SelectSingleNode("SETTING_VALUE").InnerText, Integer)
            Else
                Return iDefaultValue
            End If
        Else
            Return iDefaultValue
        End If
    End Function

    Public Function ToDatabaseDate(ByVal dtDate As Date, ByVal bDateOnly As Boolean) As String
        If bDateOnly Then
            Return ToDatabaseDate(dtDate.Date)
        Else
            Return ToDatabaseDate(dtDate)
        End If
    End Function

    Public Function ToDatabaseDate(ByVal dtDate As Date) As String
        '================================================================================================
        ' Author    :   Jim Hollingsworth 2006-07-19
        '------------------------------------------------------------------------------------------------
        ' About...  :   Formats the passed date to a format that can be interpretted by the 
        '               current Database vendor.
        '
        '               Currently all supported Database Vendors use 'yyyy-MM-dd'.
        '               (We set the NLS_DATE FORMAT setting in Oracle to 'yyyy-MM-dd')
        '               In the future we may need to check the DB Vendor. e.g.:
        '                   If m_structConfigSettings.eDatabaseVendor = DatabaseVendor.Oracle Then
        '
        '               NOTE: This function puts the enclosing single quotes around the date,
        '               if the DB Vendor requires it.
        '================================================================================================
        Dim sbReturn As New StringBuilder("'")

        If dtDate = dtDate.Date Then
            ' No Time part
            sbReturn.Append(dtDate.ToString("yyyy-MM-dd"))
        Else
            ' Include time
            sbReturn.Append(dtDate.ToString("yyyy-MM-dd HH:mm:ss"))
        End If

        sbReturn.Append("'")
        Return sbReturn.ToString()

    End Function

    Public ReadOnly Property DBProcessId() As Integer
        '================================================================================================
        ' Author    :   Jim Hollingsworth 2009-05-03
        '------------------------------------------------------------------------------------------------
        ' About...  :   Returns the DB Process Id (SPID) of the database adapters current connection
        '================================================================================================
        Get
            Return m_oDatabaseAdapter.DBProcessId()
        End Get
    End Property

#End Region

#Region "Application Server Methods"

    Public Function GetApplicationServers( _
            ByVal decRoleID As Decimal) As ParameterCollection

        Dim pcApplicationServers As ParameterCollection

        pcApplicationServers = ListParameterCollection( _
                "abComponentServices", _
                "ServerRole", _
                "RoleID = " & decRoleID.ToString(System.Globalization.CultureInfo.InvariantCulture), _
                True)

        GetApplicationServers = pcApplicationServers

    End Function

    Public Function IsApplicationServerInRole( _
            ByVal decRoleID As Decimal) As Boolean

        Dim pcApplicationServer As ParameterCollection
        Dim sbFilter As New StringBuilder("Name = '")
        Dim bApplicationServerInRole As Boolean
        Dim decApplicationServerID As Decimal

        sbFilter.Append(m_structConfigSettings.sApplicationServer)
        sbFilter.Append("'")

        pcApplicationServer = ListParameterCollection( _
                "abComponentServices", _
                "ApplicationServers", _
                sbFilter.ToString(), _
                True)

        If pcApplicationServer Is Nothing Then
            bApplicationServerInRole = True
        Else
            If pcApplicationServer.MoveNext() Then
                'Application Server found.
                pcApplicationServer = pcApplicationServer.GetParameterCollection()
                decApplicationServerID = pcApplicationServer.GetDecimalValue("ID")

                'Check to see if in Role
                sbFilter = New StringBuilder("RoleID = ")
                sbFilter.Append(decRoleID)
                sbFilter.Append(" AND ServerID = ")
                sbFilter.Append(decApplicationServerID)

                pcApplicationServer = ListParameterCollection( _
                        "abComponentServices", _
                        "ServerRole", _
                        sbFilter.ToString(), _
                        True)

                If pcApplicationServer Is Nothing Then
                    bApplicationServerInRole = True
                Else
                    If pcApplicationServer.Length > 0 Then
                        bApplicationServerInRole = True
                    Else
                        bApplicationServerInRole = False
                    End If
                End If
            Else
                bApplicationServerInRole = True
            End If
        End If

        IsApplicationServerInRole = bApplicationServerInRole

    End Function

#End Region

    Protected Overrides Sub Finalize()
        ReleaseClassicInteropHandles()

        m_oLogClient = Nothing

        MyBase.Finalize()
    End Sub

End Class

#End Region

#Region "ParameterCollection Class"

Public Class ParameterCollection

    Implements Collections.IEnumerator
    Implements Collections.IEnumerable

    Private m_oParameters As Collections.Specialized.HybridDictionary
    Private m_oEnumerator As Collections.IDictionaryEnumerator
    'Private m_oDataReader As SqlClient.SqlDataReader

    Public Sub New()

        m_oParameters = New System.Collections.Specialized.HybridDictionary(True)

    End Sub

    Public Sub Add( _
            ByVal sName As String, _
            ByVal sValue As String)

        Try
            m_oParameters.Add(sName, sValue)
        Catch
        End Try

    End Sub

    Public Sub Add( _
            ByVal sName As String, _
            ByVal iValue As Integer)

        Try
            m_oParameters.Add(sName, iValue)
        Catch
        End Try

    End Sub

    Public Sub Add( _
            ByVal sName As String, _
            ByVal decValue As Decimal)

        Try
            m_oParameters.Add(sName, decValue)
        Catch
        End Try

    End Sub

    Public Sub Add( _
            ByVal sName As String, _
            ByVal dteValue As Date)

        Try
            m_oParameters.Add(sName, dteValue)
        Catch
        End Try

    End Sub

    Public Sub Add( _
            ByVal sName As String, _
            ByVal dblValue As Double)

        Try
            m_oParameters.Add(sName.ToUpper(), dblValue)
        Catch
        End Try

    End Sub

    Public Sub Add( _
            ByVal sName As String, _
            ByVal bValue As Boolean)

        Try
            m_oParameters.Add(sName.ToUpper(), bValue)
        Catch
        End Try

    End Sub

    Public Sub Add( _
            ByVal sName As String, _
            ByVal bytValue As Byte)

        Try
            m_oParameters.Add(sName.ToUpper(), bytValue)
        Catch
        End Try

    End Sub

    Public Sub Add( _
            ByVal sName As String, _
            ByRef oParameterCollection As ParameterCollection)

        m_oParameters.Add(sName.ToUpper(), oParameterCollection)

    End Sub

    Public Sub Add( _
            ByVal sName As String, _
            ByVal bytValue() As Byte)

        m_oParameters.Add(sName.ToUpper(), bytValue)

    End Sub

    Public Sub Add( _
            ByVal sName As String, _
            ByRef ndParameter As Xml.XmlNode, _
            ByVal iDataType As Integer)

        Add(sName.ToUpper(), ndParameter, iDataType, False)

    End Sub

    Public Sub Add( _
            ByVal sName As String, _
            ByRef ndParameter As Xml.XmlNode, _
            ByVal iDataType As Integer, _
            ByVal bUseSpecialCharacterForBlank As Boolean)

        Select Case iDataType
            Case VariantType.Boolean
                m_oParameters.Add(sName.ToUpper(), ComponentServices.GetBooleanXMLElementData("", ndParameter, Nothing))
            Case VariantType.Byte
                If Not (ndParameter Is Nothing) Then
                    If ndParameter.InnerText.Length = 0 Then
                        m_oParameters.Add(sName.ToUpper(), System.DBNull.Value)
                    Else
                        m_oParameters.Add(sName.ToUpper(), ComponentServices.GetIntegerXMLElementData("", ndParameter, Nothing))
                    End If
                End If
            Case VariantType.Date
                If Not (ndParameter Is Nothing) Then
                    If ndParameter.InnerText.Length = 0 Then
                        m_oParameters.Add(sName.ToUpper(), System.DBNull.Value)
                    Else
                        m_oParameters.Add(sName.ToUpper(), ComponentServices.GetDateXMLElementData("", ndParameter, Nothing))
                    End If
                End If
            Case VariantType.Decimal
                If Not (ndParameter Is Nothing) Then
                    If ndParameter.InnerText.Length = 0 Then
                        m_oParameters.Add(sName.ToUpper(), System.DBNull.Value)
                    Else
                        m_oParameters.Add(sName.ToUpper(), ComponentServices.GetDecimalXMLElementData("", ndParameter, Nothing))
                    End If
                End If
            Case VariantType.Double
                If Not (ndParameter Is Nothing) Then
                    If ndParameter.InnerText.Length = 0 Then
                        m_oParameters.Add(sName.ToUpper(), System.DBNull.Value)
                    Else
                        m_oParameters.Add(sName.ToUpper(), ComponentServices.GetDoubleXMLElementData("", ndParameter, Nothing))
                    End If
                End If
            Case VariantType.Integer
                If Not (ndParameter Is Nothing) Then
                    If ndParameter.InnerText.Length = 0 Then
                        m_oParameters.Add(sName.ToUpper(), System.DBNull.Value)
                    Else
                        m_oParameters.Add(sName.ToUpper(), ComponentServices.GetIntegerXMLElementData("", ndParameter, Nothing))
                    End If
                End If
            Case VariantType.String
                'vbback is used as a substitute for an empty string (to indicate an empty 
                'string is required, as opposed to the parameter has not been specified).
                If bUseSpecialCharacterForBlank _
                AndAlso Not ndParameter Is Nothing _
                AndAlso ndParameter.InnerText = Nothing Then
                    m_oParameters.Add(sName.ToUpper(), vbBack)
                Else
                    m_oParameters.Add(sName.ToUpper(), ComponentServices.GetStringXMLElementData("", ndParameter, Nothing))
                End If

            Case Else
                ComponentServices.LogError( _
                        False, _
                        0, _
                        "Parameter data type not recognized.", _
                        "abComponentServices", _
                        "ParameterCollection.Add(ndParameter As Xml.XmlNode, iDataType As Integer", _
                        ComponentServices.ErrorSeverity.ES_Warning, _
                        sName)

        End Select

    End Sub

    Public Sub Add( _
            ByVal sName As String, _
            ByRef oSQLCommandParameter As SqlClient.SqlParameter)

        m_oParameters.Add(sName.ToUpper(), oSQLCommandParameter)

    End Sub

    Public Sub Add( _
            ByVal sName As String, _
            ByRef oOracleCommandParameter As OracleClient.OracleParameter)

        m_oParameters.Add(sName.ToUpper(), oOracleCommandParameter)

    End Sub

    Public Function GetValue( _
            ByVal sName As String) As Object

        Dim oValue As Object

        oValue = m_oParameters.Item(sName)

        Return oValue

    End Function

    Public Function GetValue() As Object

        Dim oValue As Object

        oValue = m_oEnumerator.Value

        Return oValue

    End Function

    Public Function GetStringValue(ByVal sName As String) As String

        Dim sValue As String

        sValue = GetStringValue(sName, Nothing)
        If sValue = Nothing Then
            Throw New ActiveBankException( _
                    ErrorNumbers.TheParameterIsMandatory, _
                    "The parameter [" + sName + "] is mandatory.", _
                    ActiveBankException.ExceptionType.Business, _
                    "abComponentServices")
        Else
            Return sValue
        End If

    End Function

    Public Function GetStringValue( _
        ByVal sName As String, _
        ByVal sDefault As String) As String

        Dim sValue As String

        If TypeOf m_oParameters.Item(sName.ToUpper()) Is String Then
            sValue = CStr(m_oParameters.Item(sName.ToUpper()))
        Else
            sValue = sDefault
        End If

        Return sValue

    End Function

    Public Function GetBooleanValue( _
            ByVal sName As String, _
            ByVal bDefault As Boolean) As Boolean

        Dim bValue As Boolean

        If TypeOf m_oParameters.Item(sName.ToUpper()) Is Boolean _
        OrElse TypeOf m_oParameters.Item(sName.ToUpper()) Is Integer _
        OrElse TypeOf m_oParameters.Item(sName.ToUpper()) Is Double Then
            bValue = CBool(m_oParameters.Item(sName.ToUpper()))
        Else
            bValue = bDefault
        End If

        Return bValue

    End Function

    Public Function GetDateValue(ByVal sName As String) As Date

        Dim dteValue As Date

        dteValue = GetDateValue(sName, Nothing)
        If dteValue = Nothing Then
            Throw New ActiveBankException( _
                    ErrorNumbers.TheParameterIsMandatory, _
                    "The parameter [" + sName + "] is mandatory.", _
                    ActiveBankException.ExceptionType.Business, _
                    "abComponentServices")
        Else
            Return dteValue
        End If

    End Function

    Public Function GetDateValue( _
            ByVal sName As String, _
            ByVal dteDefault As Date) As Date

        Dim dteValue As Date

        dteValue = CDate(m_oParameters.Item(sName.ToUpper()))
        If dteValue = Nothing Then
            dteValue = dteDefault
        End If

        Return dteValue

    End Function

    Public Function GetDecimalValue(ByVal sName As String) As Decimal

        Dim decValue As Decimal

        decValue = GetDecimalValue(sName, Nothing)
        If decValue = Nothing Then
            Throw New ActiveBankException( _
                    ErrorNumbers.TheParameterIsMandatory, _
                    "The parameter [" + sName + "] is mandatory.", _
                    ActiveBankException.ExceptionType.Business, _
                    "abComponentServices")
        Else
            Return decValue
        End If

    End Function

    Public Function GetDecimalValue( _
            ByVal sName As String, _
            ByVal decDefault As Decimal) As Decimal

        Dim decValue As Decimal

        decValue = CDec(m_oParameters.Item(sName.ToUpper()))
        If decValue = Nothing Then
            decValue = decDefault
        End If

        Return decValue

    End Function

    Public Function GetIntegerValue(ByVal sName As String) As Integer

        Dim iValue As Integer

        iValue = GetIntegerValue(sName, Nothing)
        If iValue = Nothing Then
            Throw New ActiveBankException( _
                    ErrorNumbers.TheParameterIsMandatory, _
                    "The parameter [" + sName + "] is mandatory.", _
                    ActiveBankException.ExceptionType.Business, _
                    "abComponentServices")
        Else
            Return iValue
        End If

    End Function

    Public Function GetIntegerValue( _
            ByVal sName As String, _
            ByVal iDefault As Integer) As Integer

        Dim iValue As Integer

        iValue = CInt(m_oParameters.Item(sName.ToUpper()))
        If iValue = Nothing Then
            iValue = iDefault
        End If

        Return iValue

    End Function

    Public Function GetDoubleValue(ByVal sName As String) As Double

        Dim dblValue As Double

        dblValue = GetDoubleValue(sName, Nothing)
        If dblValue = Nothing Then
            Throw New ActiveBankException( _
                    ErrorNumbers.TheParameterIsMandatory, _
                    "The parameter [@@1@" + sName + "@@] is mandatory.", _
                    ActiveBankException.ExceptionType.Business, _
                    "abComponentServices")
        Else
            Return dblValue
        End If

    End Function

    Public Function GetDoubleValue( _
            ByVal sName As String, _
            ByVal dblDefault As Double) As Double

        Dim dblValue As Decimal

        dblValue = CDbl(m_oParameters.Item(sName.ToUpper()))
        If dblValue = Nothing Then
            dblValue = dblDefault
        End If

        Return dblValue

    End Function

    Public Function GetParameterCollection() As ParameterCollection

        Dim pcParameterCollection As ParameterCollection

        pcParameterCollection = CType(m_oEnumerator.Value, ParameterCollection)

        Return pcParameterCollection

    End Function

    Public Function GetParameterCollection(ByVal sName As String) As ParameterCollection

        Dim pcParameterCollection As ParameterCollection

        pcParameterCollection = CType(m_oParameters.Item(sName), ParameterCollection)

        Return pcParameterCollection

    End Function

    Public Function GetParameterCollection(ByVal iIndex As Integer) As ParameterCollection

        Dim pcParameterCollection As ParameterCollection

        pcParameterCollection = CType(m_oParameters.Item(iIndex), ParameterCollection)

        Return pcParameterCollection

    End Function

    Public Function GetSQLCommandParameter() As SqlClient.SqlParameter

        Dim oSQLCommandParameter As SqlClient.SqlParameter

        oSQLCommandParameter = CType(m_oEnumerator.Value, SqlClient.SqlParameter)

        Return oSQLCommandParameter

    End Function

    Public Function GetSQLCommandParameter(ByVal sName As String) As SqlClient.SqlParameter

        Dim oSQLCommandParameter As SqlClient.SqlParameter

        oSQLCommandParameter = CType(m_oParameters.Item(sName), SqlClient.SqlParameter)

        Return oSQLCommandParameter

    End Function

    Public Function GetSQLCommandParameter(ByVal iIndex As Integer) As SqlClient.SqlParameter

        Dim oSQLCommandParameter As SqlClient.SqlParameter

        oSQLCommandParameter = CType(m_oParameters.Item(iIndex), SqlClient.SqlParameter)

        Return oSQLCommandParameter

    End Function

    Public Function GetOracleCommandParameter() As OracleClient.OracleParameter

        Dim oOracleCommandParameter As OracleClient.OracleParameter

        oOracleCommandParameter = CType(m_oEnumerator.Value, OracleClient.OracleParameter)

        Return oOracleCommandParameter

    End Function

    Public Function GetOracleCommandParameter(ByVal sName As String) As OracleClient.OracleParameter

        Dim oOracleCommandParameter As OracleClient.OracleParameter

        oOracleCommandParameter = CType(m_oParameters.Item(sName), OracleClient.OracleParameter)

        Return oOracleCommandParameter

    End Function

    Public Function GetOracleCommandParameter(ByVal iIndex As Integer) As OracleClient.OracleParameter

        Dim oOracleCommandParameter As OracleClient.OracleParameter

        oOracleCommandParameter = CType(m_oParameters.Item(iIndex), OracleClient.OracleParameter)

        Return oOracleCommandParameter

    End Function

    Public Sub AppendToXMLDocument( _
            ByRef ndParent As Xml.XmlNode)

        Dim oValue As Object
        Dim ndParameter As Xml.XmlNode
        Dim docParent As Xml.XmlDocument

        docParent = ndParent.OwnerDocument

        Me.Reset()
        While Me.MoveNext()
            ndParameter = docParent.CreateElement(Me.GetName())

            oValue = Me.GetValue()
            Select Case oValue.GetType().FullName()
                Case "System.String"
                    ndParameter.InnerText = CStr(oValue)

                Case "System.Int16", "System.Int32", "System.Int64"
                    ndParameter.InnerText = CStr(oValue)

                Case "System.DateTime"
                    ComponentServices.AppendDateNode(ndParameter, CDate(oValue))

                Case "System.Double"
                    ndParameter.InnerText = CStr(oValue)

                Case "System.Decimal"
                    ndParameter.InnerText = CStr(oValue)

                Case "System.Boolean"
                    If CBool(oValue) Then
                        ndParameter.InnerText = "1"
                    Else
                        ndParameter.InnerText = "0"
                    End If

            End Select

            ndParent.AppendChild(ndParameter)
        End While

    End Sub

    Public Function GetName() As String

        Dim sName As String

        sName = CStr(m_oEnumerator.Key)

        Return sName

    End Function

    Public Sub SetValue( _
            ByVal oValue As Object)

        m_oParameters.Item(m_oEnumerator.Key) = oValue

    End Sub

    Public Sub SetValue( _
            ByVal sName As String, _
            ByVal oValue As Object)

        m_oParameters.Item(sName.ToUpper()) = oValue

    End Sub

    Public Sub Reset() Implements System.Collections.IEnumerator.Reset

        If m_oEnumerator Is Nothing Then
            m_oEnumerator = m_oParameters.GetEnumerator()
        Else
            m_oEnumerator.Reset()
        End If

    End Sub

    Public Function MoveNext() As Boolean Implements System.Collections.IEnumerator.MoveNext
        Dim bResult As Boolean

        'If bAutoCreateEnumerator Then
        '    If m_oEnumerator Is Nothing Then
        '        ' Someone has called MoveNext without doing a reset...
        '        Reset()
        '    End If
        'End If

        Try
            bResult = m_oEnumerator.MoveNext()
        Catch oError As Exception
            'EventLog.WriteEntry("[ComponentServices]", "Thread : " & Threading.Thread.CurrentThread.GetHashCode().ToString & vbCrLf & "MoveNext failed (Record Count:" & m_oParameters.Count.ToString & ") :" & oError.Message & vbCrLf & oError.StackTrace, EventLogEntryType.Error)
            bResult = False
        End Try

        Return bResult

    End Function

    Public Sub Remove(ByVal sName As String)

        Try
            m_oParameters.Remove(sName.ToUpper())
        Catch
        End Try

    End Sub

    Public ReadOnly Property Length() As Integer

        Get
            Length = m_oParameters.Count
        End Get

    End Property

    Public ReadOnly Property Current() As Object Implements System.Collections.IEnumerator.Current

        Get
            Current = Me.m_oEnumerator.Current
        End Get

    End Property

    Public Function GetEnumerator() As System.Collections.IEnumerator Implements System.Collections.IEnumerable.GetEnumerator

        If m_oEnumerator Is Nothing Then
            m_oEnumerator = m_oParameters.GetEnumerator()
        Else
            m_oEnumerator.Reset()
        End If

        'EventLog.WriteEntry( "1", CStr(m_oEnumerator Is Nothing))
        Return m_oEnumerator

    End Function
    Public Function Item(ByVal Index As Integer) As ParameterCollection
        Return m_oParameters.Item(Index.ToString(System.Globalization.CultureInfo.InvariantCulture))
    End Function
End Class

#End Region

#Region "ActiveBankException Class"

Public Class ActiveBankException : Inherits System.Exception

    Public Enum ExceptionType
        System = 1
        Business = 2
    End Enum

    Private m_eType As ExceptionType
    Private m_iNumber As Integer
    Private m_sComponentName As String

    Public Sub New( _
            ByVal iNumber As Integer, _
            ByVal sDescription As String)

        MyBase.New(sDescription)

        m_iNumber = iNumber
        m_eType = ExceptionType.Business
        m_sComponentName = "abComponentServices"

    End Sub

    Public Sub New( _
            ByVal sDescription As String, _
            ByVal eType As ActiveBankException.ExceptionType)

        MyBase.New(sDescription)

        m_iNumber = 0
        m_eType = eType
        m_sComponentName = "abComponentServices"

    End Sub

    Public Sub New( _
        ByVal iNumber As Integer, _
        ByVal sMessage As String, _
        ByVal sComponentName As String)

        MyBase.New(sMessage)

        m_iNumber = iNumber
        m_eType = ExceptionType.System
        m_sComponentName = sComponentName

    End Sub

    Public Sub New( _
            ByVal iNumber As Integer, _
            ByVal sDescription As String, _
            ByVal eType As ExceptionType)

        MyBase.New(sDescription)

        m_iNumber = iNumber
        m_eType = eType
        m_sComponentName = "abComponentServices"

    End Sub

    Public Sub New( _
            ByVal iNumber As Integer, _
            ByVal sDescription As String, _
            ByVal eType As ExceptionType, _
            ByVal sComponentName As String)

        MyBase.New(sDescription)

        m_iNumber = iNumber
        m_eType = eType
        m_sComponentName = sComponentName

    End Sub

    Public Sub New( _
            ByRef docError As Xml.XmlDocument)

        MyBase.New(ComponentServices.GetStringXMLElementData("Description", docError.SelectSingleNode("abError"), ""))

        m_iNumber = ComponentServices.GetIntegerXMLElementData("Number", docError.SelectSingleNode("abError"), 0)
        m_eType = ExceptionType.Business
        m_sComponentName = ComponentServices.GetStringXMLElementData("Component", docError.SelectSingleNode("abError"), "")

    End Sub

    Public Property Number() As Integer
        Get
            Number = m_iNumber
        End Get
        Set(ByVal iNumber As Integer)
            m_iNumber = iNumber
        End Set
    End Property

    Public Property Type() As ActiveBankException.ExceptionType
        Get
            Type = m_eType
        End Get
        Set(ByVal eType As ActiveBankException.ExceptionType)
            m_eType = eType
        End Set
    End Property

    Public Property ComponentName() As String
        Get
            ComponentName = m_sComponentName
        End Get
        Set(ByVal sComponentName As String)
            m_sComponentName = sComponentName
        End Set
    End Property

End Class

#End Region

#Region "ActiveBankXMLDocument Class"

Public Class ActiveBankXMLDocument

    Private m_oXPathDocument As Xml.XPath.XPathDocument
    Private m_oDocument As Xml.XmlDocument
    Private m_bReadOnly As Boolean
    Private m_oNavigator As Xml.XPath.XPathNavigator

    Friend Sub New(ByRef oReader As Xml.XmlReader)

        m_bReadOnly = True

        m_oXPathDocument = New System.Xml.XPath.XPathDocument(oReader)
        m_oNavigator = m_oXPathDocument.CreateNavigator()

    End Sub

    Friend Sub New(ByRef oReader As Xml.XmlReader, ByVal bReadOnly As Boolean)

        m_bReadOnly = bReadOnly

        If bReadOnly Then
            m_oXPathDocument = New System.Xml.XPath.XPathDocument(oReader)
            m_oNavigator = m_oXPathDocument.CreateNavigator()
        Else
            m_oDocument = New Xml.XmlDocument
            m_oDocument.LoadXml(oReader.ReadOuterXml())
        End If

    End Sub

    Friend Sub New(ByVal sPathAndFileName As String)

        m_bReadOnly = True

        m_oXPathDocument = New System.Xml.XPath.XPathDocument(sPathAndFileName)
        m_oNavigator = m_oXPathDocument.CreateNavigator()

    End Sub

    Friend Sub New(ByVal sPathAndFileName As String, ByVal bReadOnly As Boolean)

        m_bReadOnly = bReadOnly

        If bReadOnly Then
            m_oXPathDocument = New System.Xml.XPath.XPathDocument(sPathAndFileName)
            m_oNavigator = m_oXPathDocument.CreateNavigator()
        Else
            m_oDocument = New Xml.XmlDocument
            m_oDocument.Load(sPathAndFileName)
        End If

    End Sub

    Friend Sub New(ByVal oDocument As Xml.XmlDocument)

        m_bReadOnly = False

        m_oDocument = oDocument

    End Sub

    Public Function GetStringElement() As String

        Dim sValue As String

        If m_bReadOnly Then
        End If

        Return sValue

    End Function

End Class

#End Region

#Region "activebankWarningCollection Class"

Public Class activebankWarningCollection

    Private m_oWarnings As Collections.Specialized.HybridDictionary

    Public Sub New()

        m_oWarnings = New System.Collections.Specialized.HybridDictionary

    End Sub

    Public Sub Add(ByVal iNumber, ByVal sDescription)

        m_oWarnings.Add(iNumber, sDescription)

    End Sub

    Public Sub AppendToXMLDocument(ByRef docInput As Xml.XmlNode)

        Dim ndWarnings As Xml.XmlNode
        Dim oEnumerator As Collections.IDictionaryEnumerator
        Dim ndWarning As Xml.XmlNode
        Dim ndRoot As Xml.XmlNode

        If docInput Is Nothing Then
            docInput = ComponentServices.CreateXMLDocument("<abWarnings/>")
            ndWarnings = docInput.SelectSingleNode("abWarnings")
        Else
            ndRoot = docInput.FirstChild
            If ndRoot Is Nothing Then
                docInput = ComponentServices.CreateXMLDocument("<abWarnings/>")
                ndWarnings = docInput.SelectSingleNode("abWarnings")
            Else
                ndWarnings = ComponentServices.AppendXMLElement("abWarnings", ndRoot)
            End If
        End If

        oEnumerator = m_oWarnings.GetEnumerator()
        While oEnumerator.MoveNext()
            ndWarning = ComponentServices.AppendXMLElement("Warning", ndWarnings)

            ComponentServices.AppendXMLElement("Number", ndWarning, oEnumerator.Key)
            ComponentServices.AppendXMLElement("Description", ndWarning, oEnumerator.Value)
        End While

    End Sub

End Class

#End Region
