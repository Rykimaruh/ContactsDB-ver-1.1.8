Imports System.Security.Cryptography
Imports System.Text
Module Functions
    Dim con As New OleDb.OleDbConnection
    Dim provider, dbsource, dbmydocs, thedatabase, fulldbpath As String
    Dim dbset As New DataSet 'copy of the info from a database
    Dim dbadapter As OleDb.OleDbDataAdapter
    Dim sql As String = "SELECT * FROM tblContacts"
    Dim inc, maxrow As Integer
    Dim dbName As String = Trim(Setup.TextBox2.Text)

    Public Function customerid()
        Dim i As Integer
        Dim str As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789"
        Dim rand As New Random
        Dim sb As New StringBuilder
        For i = 0 To 4
            Dim inx As Integer = rand.Next(0, 35)
            sb.Append(str.Substring(inx, 1))
        Next
        Return sb.ToString
    End Function
    Public Sub emailtolowercase()
        Form1.TextBox8.Text = Form1.TextBox8.Text.ToLower
    End Sub
    Public Function checkemail()
        Dim email As String = Form1.TextBox8.Text
        Dim isemail As String = "@"
        Dim dotcom As String = ".com"
        If InStr(email, isemail) And InStr(email, dotcom) Then
            Return True
        Else
            MessageBox.Show("Invalid email!")
            Form1.TextBox8.Text = ""
            Return False

        End If
    End Function
    Public Sub autonumber()
        Dim newdbrow As DataRow
        newdbrow = dbset.Tables(dbName).NewRow()
        newdbrow.Item(0) = newdbrow
    End Sub
End Module