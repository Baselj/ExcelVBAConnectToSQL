Attribute VB_Name = "ConnectToSQL"
'Your connection string
Const SqlCon As String = "Provider=sqloledb.1;Data Source=pri-trans-sql;Initial Catalog=YOURDATABASE;User ID=YOURUSERNAME;Password=YOURPASSWORD;"

Sub ConnectToDB()
    
    Dim Conn As ADODB.Connection
    Dim Data As ADODB.Recordset
    Dim Field As ADODB.Field
    
    Set Conn = New ADODB.Connection
    Set Data = New ADODB.Recordset
    Conn.ConnectionString = SqlCon
    Conn.Open
    
    On Error GoTo CloseConnection
    
    With Data
    .ActiveConnection = SqlCon
    .Source = GetSQLString()
    .LockType = adLockReadOnly
    .CursorType = adOpenDynamic
    .Open
    End With
    
    On Error GoTo CloseRecordSet
    
    Sheet1.Activate
    Sheet1.Cells.Clear
    Sheet1.Range("A1").Select
    
    
    For Each Field In Data.Fields
        ActiveCell.Value = Field.Name
        ActiveCell.Offset(0, 1).Select
    Next Field
    
    
    Sheet1.Range("A1").Select
    Sheet1.Range("A2").CopyFromRecordset Data

    Conn.Close
    Data.Close

    On Error GoTo 0
    Exit Sub
    
CloseRecordSet:
    Data.Close
    
CloseConnection:
    Conn.Close
    MsgBox Err.Description

End Sub

'Your SQL Query
Function GetSQLString() As String
    Dim SqlString As String
    
    SqlString = "SELECT * FROM YOURTABLE" & _
                "WHERE CONDITION 1" & _
                "AND CONDITION 2"
                
    GetSQLString = SqlString
    
End Function




















