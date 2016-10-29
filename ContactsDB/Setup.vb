Imports System.IO
Public Class Setup
    Dim con As New OleDb.OleDbConnection
    Dim provider, dbsource, dbmydocs, thedatabase, fulldbpath As String
    Dim dbset As New DataSet 'copy of the info from a database
    Dim dbadapter As OleDb.OleDbDataAdapter
    Dim inc As Integer
    Dim maxrow As Integer
    Dim strFileName As String


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'OpenFileDialog1.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
        ' OpenFileDialog1.Title = "Open Text File"
        'OpenFileDialog1.Filter = "SQL(*.sql)|*.sql|Access (*.mdb)|*.mdb|All|*.*"
        'OpenFD.ShowDialog() had OpenFD.ShowDialog() twice in this line and in line 43
        ' If OpenFileDialog1.ShowDialog = Windows.Forms.DialogResult.Cancel Then
        'OpenFileDialog1.Dispose()
        ' Else
        ' strFileName = OpenFileDialog1.FileName
        ' TextBox1.Text = OpenFileDialog1.FileName
        ' End If
        FolderBrowserDialog1.ShowNewFolderButton = False
        If (FolderBrowserDialog1.ShowDialog() = DialogResult.OK) Then
            TextBox1.Text = FolderBrowserDialog1.SelectedPath
        End If
    End Sub
    Public Sub exists()
        Dim path As String = TextBox1.Text
        If Directory.Exists(path) Then
            'System.IO.File.Exists(path + "\" & foldernames(i))
            'MessageBox.Show("One or more directories exist already")
            MessageBox.Show("Confirmed:" & IO.Path.GetFullPath(path) + "\" & TextBox2.Text)
            Exit Sub
        End If
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs)
        exists()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        'If ComboBox1.SelectedItem = "Microsoft Access" Then
        '    extension = ".mdb"
        'End If
        'If ComboBox1.SelectedItem = "SQL" Then
        '    extension = ".sql"
        'End If
        'If ComboBox1.SelectedItem = "Oracle" Then
        '    extension = ".dbf"
        'End If

        If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Then
            MessageBox.Show("No empty fields allowed!")
            Exit Sub
        Else
            Form1.Show()
        End If

    End Sub

   
    Private Sub ExxitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExxitToolStripMenuItem.Click
        Me.Close()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If con.State = ConnectionState.Open Then
            MessageBox.Show("Connected")
        End If

        If con.State = ConnectionState.Closed Then
            MessageBox.Show("Not Connected")
        End If
    End Sub

    Private Sub Button2_Click_1(sender As Object, e As EventArgs) Handles Button2.Click
        CreateTbl.Show()
    End Sub
End Class