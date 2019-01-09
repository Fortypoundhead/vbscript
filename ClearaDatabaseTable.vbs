' Clear a Database Table

Const adOpenStatic = 3
Const adLockOptimistic = 3
Const adUseClient = 3
Set objConnection = CreateObject("ADODB.Connection")
Set objRecordset = CreateObject("ADODB.Recordset")
objConnection.Open "DSN=Inventory;"
objRecordset.CursorLocation = adUseClient
objRecordset.Open "Delete * FROM Hardware" , objConnection, adOpenStatic, adLockOptimistic
objConnection.Close