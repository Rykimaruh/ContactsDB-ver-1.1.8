Imports System.Data.OleDb
Public Class SearchForm
    Dim dbset As New DataSet
    Dim con As New OleDbConnection
    Dim provider As String = "PROVIDER = Microsoft.Jet.OLEDB.4.0;"
    Dim thedatabase As String = "/AddressBook.mdb" 'name of the database
    Dim dbmydocs As String = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) 'location of the database
    Dim fulldbpath As String = dbmydocs & thedatabase 'combine the fullpath and name of the database into one path
    Dim dbsource As String = "Data Source = " & fulldbpath 'combine data name and data path to make datasource
    Dim sql, sqlsearch As String ' Our SQL Statement

    Private Sub SearchForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Size = New System.Drawing.Size(291, 161)
        Label2.Visible = False
        Label3.Visible = False
        Label4.Visible = False
        TextBox1.Visible = False
        TextBox2.Visible = False
        TextBox3.Visible = False
        Me.DataGridView1.Visible = False
        con.ConnectionString = provider & dbsource
        ' Our SQL Statement
        Dim sql As String
        sql = "SELECT * FROM tblContacts"
        Dim adapter As New OleDbDataAdapter(sql, con) 'This is our DataAdapter. This executes our SQL Statement above against the Database we defined in the Connection String
        Dim dt As New DataTable("tblContacts") ' Gets the records from the table and fills our adapter with those.
        adapter.Fill(dt)
        DataGridView1.DataSource = dt ' Assigns our DataSource on the DataGridView
        Dim sql1 As String
        sql1 = "SELECT * FROM tblContacts"
        Dim adapter1 As New OleDbDataAdapter(sql1, con)
        Dim cmd1 As New OleDbCommand(sql1, con)
        con.Open()
        Dim myreader As OleDbDataReader = cmd1.ExecuteReader
        myreader.Read()
        con.Close()
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        If ComboBox1.SelectedItem = "Last Name" Then
            Label2.Visible = True
            TextBox1.Visible = True

            Label3.Visible = False
            TextBox2.Visible = False
            Label4.Visible = False
            TextBox3.Visible = False
        End If
        If ComboBox1.SelectedItem = "Customer ID" Then
            Label3.Visible = True
            TextBox2.Visible = True

            Label2.Visible = False
            TextBox1.Visible = False
            Label4.Visible = False
            TextBox3.Visible = False
        End If
        If ComboBox1.SelectedItem = "City" Then
            Label4.Visible = True
            TextBox3.Visible = True

            Label2.Visible = False
            TextBox1.Visible = False
            Label3.Visible = False
            TextBox2.Visible = False
        End If
    End Sub

    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        Me.Close()
    End Sub
    'SEARCH BUTTON
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Size = New System.Drawing.Size(627, 443)
        DataGridView1.Visible = True
        Label2.Visible = False
        TextBox1.Visible = False
        ComboBox1.Visible = False
        Label1.Visible = False
        Button1.Visible = False
        Label3.Visible = False
        TextBox2.Visible = False
        Label4.Visible = False
        TextBox3.Visible = False
        con.ConnectionString = provider & dbsource

        ' method to retrieve a column name from database
        con.ConnectionString = provider & dbsource
        sql = "Select * FROM tblContacts"
        Dim dbadapter As New OleDbDataAdapter(sql, con)
        Dim dt As New DataTable("tblContacts")
        dbadapter.Fill(dt)
        Me.Label5.Text = dt.Columns(1).ColumnName

        If ComboBox1.SelectedItem = "City" Then
            con.ConnectionString = provider & dbsource
            ' Dim dt As New DataTable
            Dim sqlsearch As String   'SQL Statement so our User can search for City
            Dim citycol As String = dt.Columns(5).ColumnName
            sqlsearch = "SELECT * FROM tblContacts WHERE [" & citycol & "] LIKE '%" & TextBox3.Text & "%'" '& " OR LastName LIKE '%" & TextBox3.Text & "%'"
            Dim adapter As New OleDbDataAdapter(sqlsearch, con) ' Once again we execute the SQL statements against our DataBase
            'Shows the records and updates the DataGridView
            adapter.Fill(dt)
            DataGridView1.DataSource = dt
        End If
    End Sub

    Private Sub ReturnToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ReturnToolStripMenuItem.Click
        DataGridView1.Visible = False
        Me.Size = New System.Drawing.Size(297, 166)
        Label2.Visible = False
        TextBox1.Visible = False
        ComboBox1.Visible = False
        Label1.Visible = False
        Button1.Visible = False
        ComboBox1.Visible = True
        Button1.Visible = True
        Label1.Visible = True
    End Sub

    Private Sub MenuStrip1_ItemClicked(sender As Object, e As ToolStripItemClickedEventArgs) Handles MenuStrip1.ItemClicked

    End Sub
End Class