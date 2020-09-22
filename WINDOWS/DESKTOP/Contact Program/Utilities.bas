Attribute VB_Name = "Utilities"

Public Function OpenEmployeeDB(rs As ADODB.Recordset, conn As ADODB.Connection, sql As String)
    Dim strConnect As String
    Dim strProvider As String
    Dim strDataSource As String
    Dim strDatabaseName As String

    strProvider = "Provider=Microsoft.Jet.OLEDB.4.0;"
    strDataSource = App.Path
    strDatabaseName = "\Employee.mdb"
    strDataSource = "Data Source=" & strDataSource & _
        strDatabaseName
    
    strConnect = strProvider & strDataSource



    conn.CursorLocation = adUseClient
    conn.Open strConnect


    rs.CursorType = adOpenStatic

    rs.LockType = adLockPessimistic

    rs.Source = sql

    rs.ActiveConnection = conn

    rs.Open
End Function

