Attribute VB_Name = "OtherDatabase"
Public Function GetDatabaseConStr(Product_Info As String, Optional Emp_No) As String

Dim SQLQ
Dim xVersion
Dim xDatabaseName
Dim xDatabaseServer
Dim xUsername
Dim xPassword
Dim xDatabasePath
Dim rsDataSetup As New ADODB.Recordset
On Error GoTo err_GetDatabaseConStr
GetDatabaseConStr = ""

SQLQ = "SELECT * FROM APPLICATION_PARAMETER WHERE PARA_TYPE='Integration' AND PARA_CATEGORY='" & Product_Info & "' AND PARA_CATEGORY2='Database Setup' "

rsDataSetup.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly

Do Until rsDataSetup.EOF

    If rsDataSetup("PARA_NAME") = "Version_Info" Then
       xVersion = rsDataSetup("PARA_VALUE")
    Else
       If xVersion = "MS SQL Server" Then
            If rsDataSetup("PARA_NAME") = "Database_Name" Then
                xDatabaseName = rsDataSetup("PARA_VALUE")
            End If
            If rsDataSetup("PARA_NAME") = "Database_Server" Then
                xDatabaseServer = rsDataSetup("PARA_VALUE")
            End If

            If rsDataSetup("PARA_NAME") = "User_Name" Then
                xUsername = rsDataSetup("PARA_VALUE")
            End If
            If rsDataSetup("PARA_NAME") = "Password" Then
                xPassword = rsDataSetup("PARA_VALUE")
            End If

        Else
            If rsDataSetup("PARA_NAME") = "Database_Path" Then
                xDatabasePath = rsDataSetup("PARA_VALUE")
                xDatabasePath = xDatabasePath & IIf(Right(xDatabasePath, 1) = "\", "", "\")
            End If
            If rsDataSetup("PARA_NAME") = "Database_Name" Then
                xDatabaseName = rsDataSetup("PARA_VALUE")
            End If
        End If
    End If
    rsDataSetup.MoveNext
Loop

rsDataSetup.Close

If xVersion = "MS SQL Server" Then
    GetDatabaseConStr = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & xUsername & ";Password=" & xPassword & ";Initial Catalog=" & xDatabaseName & ";Data Source=" & xDatabaseServer
ElseIf xVersion = "MS Access" Then
    GetDatabaseConStr = "Provider=Microsoft.Jet.OLEDB.4.0;Jet OLEDB:Database Password=petman;Data Source=" & xDatabasePath & xDatabaseName
End If

Exit Function

err_GetDatabaseConStr:

    GetDatabaseConStr = ""

End Function
