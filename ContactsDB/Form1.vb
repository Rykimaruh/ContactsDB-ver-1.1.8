Public Class Form1
    Dim con As New OleDb.OleDbConnection
    Dim provider, dbsource, dbmydocs, thedatabase, fulldbpath As String
    Dim dbset As New DataSet 'copy of the info from a database
    'Dim extension As String = Trim(".mdb")
    Dim tblName As String = Trim(Setup.TextBox3.Text)
    Dim dbName As String = Trim(Setup.TextBox4.Text)
    Dim dbadapter As OleDb.OleDbDataAdapter
    Dim sql As String = "SELECT * FROM " + tblName
    Dim inc As Integer
    Dim maxrow As Integer
    Dim dt As New DataTable
    'commandbuilder builds sql string for you. Necessary to update database
     
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles Me.Load
        TextBox1.ReadOnly = True
        TextBox2.ReadOnly = True
        TextBox3.ReadOnly = True
        TextBox4.ReadOnly = True
        TextBox5.ReadOnly = True
        TextBox6.ReadOnly = True
        TextBox7.ReadOnly = True
        TextBox8.ReadOnly = True
        TextBox9.ReadOnly = True
        TextBox10.ReadOnly = True
        Button6.Enabled = False
        Button7.Visible = False
        Button11.Visible = False
        ComboBox1.Visible = False


        'from here
        '!!!!Need to find a way to transfer the data in Setup textboxes to this form's load event.
        'Temporarily placed the current database directory path to textbox1 and database name to textbox2 until i can load Setup form first

        provider = "PROVIDER = Microsoft.Jet.OLEDB.4.0;" 'provider technology used to connect to the Database
        ' thedatabase = "/AddressBook.mdb" 'name of the database
        thedatabase = "/" & dbName
        'dbmydocs = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) 'location of the database
        dbmydocs = Setup.TextBox1.Text 'fulldoc path
        fulldbpath = dbmydocs & thedatabase 'combine the fullpath and name of the database into one path
        dbsource = "Data Source = " & fulldbpath 'combine data name and data path to make datasource

        con.ConnectionString = provider & dbsource 'sets up connection string (IMPORTANT)
        con.Open()
        dbadapter = New OleDb.OleDbDataAdapter(sql, con) 'pass connection string and sql string to dbadapter
        dbadapter.Fill(dbset, dbName) 'fill dataset with records from table. Name of table will be used to display information onto VB. NET forms.
        con.Close()

        maxrow = dbset.Tables(dbName).Rows.Count 'counts the amount of rows are in a dataset
        inc = -1 'makes sure counters are set correctly

        'starts at first record?
        If inc <> maxrow - 1 Then
            inc += 1
            navigaterecords()
        End If
        navigaterecords()


        '---------------------------------------------------------------
        con.ConnectionString = provider & dbsource
        sql = "Select * FROM " + tblName

        Dim dt As New DataTable(tblName)
        dbadapter.Fill(dt)
        Me.Label12.Text = dt.Columns(0).ColumnName
        Me.Label13.Text = dt.Columns(1).ColumnName
        Me.Label14.Text = dt.Columns(2).ColumnName 'because i chose 2 columns COlumn 2 does not appear
        Me.Label15.Text = dt.Columns(3).ColumnName
        Me.Label16.Text = dt.Columns(5).ColumnName
        Me.Label17.Text = dt.Columns(6).ColumnName
        Me.Label18.Text = dt.Columns(7).ColumnName
        Me.Label19.Text = dt.Columns(8).ColumnName
        Me.Label20.Text = dt.Columns(9).ColumnName
        Me.Label21.Text = dt.Columns(10).ColumnName
        Me.Label22.Text = dt.Columns(11).ColumnName
    End Sub
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Try
            If inc <> maxrow - 1 Then
                inc += 1
                navigaterecords()
            Else
                MessageBox.Show("End of records")
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
    Public Sub navigaterecords()
        Try
            'what is in AddressBook table from fill, add infromation from rows and items set.
            TextBox1.Text = dbset.Tables(dbName).Rows(inc).Item(1) 'Fname
            TextBox2.Text = dbset.Tables(dbName).Rows(inc).Item(2) 'Lname
            TextBox3.Text = dbset.Tables(dbName).Rows(inc).Item(3) 'Address1
            TextBox4.Text = dbset.Tables(dbName).Rows(inc).Item(4) 'Address2
            TextBox5.Text = dbset.Tables(dbName).Rows(inc).Item(5) 'City
            ComboBox1.Text = dbset.Tables(dbName).Rows(inc).Item(6) 'Country
            TextBox11.Text = dbset.Tables(dbName).Rows(inc).Item(6)
            TextBox6.Text = dbset.Tables(dbName).Rows(inc).Item(7) 'Zip Code
            TextBox7.Text = dbset.Tables(dbName).Rows(inc).Item(8) 'Phone
            TextBox8.Text = dbset.Tables(dbName).Rows(inc).Item(9) 'Email
            TextBox9.Text = dbset.Tables(dbName).Rows(inc).Item(10) 'CustomerID
            TextBox10.Text = dbset.Tables(dbName).Rows(inc).Item(11) 'Notes


            If dbset.Tables(dbName).Rows(inc).Item(10) = "" Then
                MessageBox.Show("Note field empty")
            End If
            'add .ToString in order to avoid the "dbnull to type string not valid" error message. This is becasue the some of these fields can't have null, therefore to string voncersion is better
            'however it won't record into the database for some reason so for now do not add .ToString submethod 
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If inc > 0 Then
            inc -= 1
            navigaterecords()
        ElseIf inc = -1 Then 'since it doesn't load to 0 and starts at -1
            MessageBox.Show("No records")
        Else
            MessageBox.Show("First Record")
        End If

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If inc <> 0 Then
            inc = 0
            navigaterecords()
        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If inc <> maxrow - 1 Then
            inc = maxrow - 1
            navigaterecords()
        End If
    End Sub
    'UPDATE BUTTON
    'Error:Dynamic sql generation for the update command is not supported against a select command that does not return any key column information. Need primary key
    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Dim cb As New OleDb.OleDbCommandBuilder(dbadapter) 'this has to go after dbadapter has been called in previous code lines
        Try
            If inc = -1 Then
                MessageBox.Show("No empty fields allowed!")
                TextBox1.ReadOnly = False
                TextBox2.ReadOnly = False
                TextBox3.ReadOnly = False
                TextBox4.ReadOnly = False
                TextBox5.ReadOnly = False
                TextBox6.ReadOnly = False
                TextBox7.ReadOnly = False
                TextBox8.ReadOnly = False
                TextBox9.ReadOnly = False
                TextBox10.ReadOnly = False
            Else
                TextBox1.Text = StrConv(TextBox1.Text, VbStrConv.ProperCase) 'the first letter of every word is uppercase
                TextBox2.Text = StrConv(TextBox2.Text, VbStrConv.ProperCase)
                TextBox3.Text = StrConv(TextBox3.Text, VbStrConv.ProperCase)
                TextBox4.Text = StrConv(TextBox4.Text, VbStrConv.ProperCase)
                TextBox5.Text = StrConv(TextBox5.Text, VbStrConv.ProperCase)
                emailtolowercase()
                dbset.Tables(dbName).Rows(inc).Item(1) = TextBox1.Text
                dbset.Tables(dbName).Rows(inc).Item(2) = TextBox2.Text
                dbset.Tables(dbName).Rows(inc).Item(3) = TextBox3.Text
                dbset.Tables(dbName).Rows(inc).Item(4) = TextBox4.Text
                dbset.Tables(dbName).Rows(inc).Item(5) = TextBox5.Text
                dbset.Tables(dbName).Rows(inc).Item(6) = ComboBox1.Text
                dbset.Tables(dbName).Rows(inc).Item(7) = TextBox6.Text

                dbset.Tables(dbName).Rows(inc).Item(8) = TextBox7.Text

                dbset.Tables(dbName).Rows(inc).Item(10) = TextBox9.Text 'CustomerID
                dbset.Tables(dbName).Rows(inc).Item(11) = TextBox10.Text


                If checkemail() = False Then
                    dbset.Tables(dbName).Rows(inc).Item(9) = ""
                    dbadapter.Update(dbset, dbName) 'Update is a method of the data adapter. You need to set the info on the dataset of the named database
                ElseIf checkemail() = True Then
                    dbset.Tables(dbName).Rows(inc).Item(9) = TextBox8.Text
                    dbadapter.Update(dbset, dbName) 'Update is a method of the data adapter. You need to set the info on the dataset of the named database
                    MessageBox.Show("Data Updated")
                End If

                If TextBox9.Text = "" Then
                    TextBox9.Text = customerid()
                End If
                TextBox1.ReadOnly = True
                TextBox2.ReadOnly = True
                TextBox3.ReadOnly = True
                TextBox4.ReadOnly = True
                TextBox5.ReadOnly = True
                TextBox6.ReadOnly = True
                TextBox7.ReadOnly = True
                TextBox8.ReadOnly = True
                TextBox9.ReadOnly = True
                TextBox10.ReadOnly = True
                TextBox11.Visible = True
                ComboBox1.Visible = False
                Button6.Enabled = False
                Button7.Visible = False
                Button11.Visible = False

                'put textboxes back to readonly once again
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try


    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Button6.Enabled = True
        Button5.Enabled = False
        Button7.Enabled = False
        Button8.Enabled = False
        Button11.Visible = True
        TextBox11.Visible = False
        ComboBox1.Visible = True

        TextBox1.Clear()
        TextBox2.Clear()
        TextBox3.Clear()
        TextBox4.Clear()
        TextBox5.Clear()
        TextBox6.Clear()
        TextBox7.Clear()
        TextBox8.Clear()
        TextBox9.Clear()
        TextBox10.Clear()

        TextBox1.ReadOnly = False
        TextBox2.ReadOnly = False
        TextBox3.ReadOnly = False
        TextBox4.ReadOnly = False
        TextBox5.ReadOnly = False
        TextBox6.ReadOnly = False
        TextBox7.ReadOnly = False
        TextBox8.ReadOnly = False
        TextBox9.ReadOnly = True
        TextBox10.ReadOnly = False
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        Button5.Enabled = True
        Button7.Enabled = True
        Button8.Enabled = True
        Button6.Enabled = False

        TextBox1.ReadOnly = True
        TextBox2.ReadOnly = True
        TextBox3.ReadOnly = True
        TextBox4.ReadOnly = True
        TextBox5.ReadOnly = True
        TextBox6.ReadOnly = True
        TextBox7.ReadOnly = True
        TextBox8.ReadOnly = True
        TextBox9.ReadOnly = True
        TextBox10.ReadOnly = True
        TextBox11.Visible = True
        ComboBox1.Visible = False

        inc = 0
        navigaterecords()
    End Sub
    'COMMIT BUTTON 
    'WARNING: COMMITTING DATA SEEMS TO ONLY WORK WHEN YOU ARE ON AN EXISTING RECORD! IF YOU ADD IT FROM THE GET GO IT WONT WORK. FIX THIS WHEN YOU CAN!
    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        'If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Or TextBox4.Text = "" Or TextBox5.Text = "" Or TextBox6.Text = "" Or TextBox8.Text = "" Or TextBox9.Text = "" Then
        'MessageBox.Show("No empty fields allowed")
        'Exit Sub
        ' End If

        If inc <> -1 Then
            Dim cb As New OleDb.OleDbCommandBuilder(dbadapter) 'always create a new commandbuilder object when trying to write to or update a database
            Dim newdbrow As DataRow 'new row on dataset then you need datarow
            TextBox9.Text = customerid()
            emailtolowercase()
            autonumber()
            'object reference not set to an instance object here
            newdbrow = dbset.Tables(dbName).NewRow() 'create new row in the AddressBook database. Create new row in the dataset AddressBook. NewRow Method

            TextBox1.Text = StrConv(TextBox1.Text, VbStrConv.ProperCase) 'the first letter of every word is uppercase
            TextBox2.Text = StrConv(TextBox2.Text, VbStrConv.ProperCase)
            TextBox3.Text = StrConv(TextBox3.Text, VbStrConv.ProperCase)
            TextBox4.Text = StrConv(TextBox4.Text, VbStrConv.ProperCase)
            TextBox5.Text = StrConv(TextBox5.Text, VbStrConv.ProperCase)


            newdbrow.Item(1) = TextBox1.Text 'FName
            newdbrow.Item(2) = TextBox2.Text 'LName
            newdbrow.Item(3) = TextBox3.Text 'Address1
            newdbrow.Item(4) = TextBox4.Text 'Address2
            newdbrow.Item(5) = TextBox5.Text 'City
            newdbrow.Item(6) = ComboBox1.Text 'Country
            newdbrow.Item(7) = TextBox6.Text 'Zipcode
            newdbrow.Item(8) = TextBox7.Text 'Phone
            newdbrow.Item(9) = TextBox8.Text 'Email
            checkemail()
            newdbrow.Item(10) = TextBox9.Text 'CusomterID
            newdbrow.Item(11) = TextBox10.Text 'Notes


            dbset.Tables(dbName).Rows.Add(newdbrow) 'this is the method that actually adds the row to the dataset. Add method of Row prop. Datarow info as paramater
            dbadapter.Update(dbset, dbName) 'insert to statement error
            MessageBox.Show("New record added to Database!")

            Button6.Enabled = False
            Button5.Enabled = True
            Button7.Enabled = True >
            Button8.Enabled = True

            navigaterecords()

        End If
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        If MessageBox.Show("Do you really want to Delete this Record?", "Delete", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) = DialogResult.No Then
            MessageBox.Show("Operation Cancelled")
            Exit Sub
        Else
            If inc <> -1 Then
                Dim cb As New OleDb.OleDbCommandBuilder(dbadapter)
                dbset.Tables(dbName).Rows(inc).Delete() 'just as there is an add method, there is adelete method 
                dbadapter.Update(dbset, dbName) 'basically this needs to be called everytime you want to make any changes directly to the database and not Data set
                maxrow -= 1
                inc = 0
                navigaterecords()
            Else
                MessageBox.Show("No Records!")
            End If
        End If
    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        TextBox1.ReadOnly = False
        TextBox2.ReadOnly = False
        TextBox3.ReadOnly = False
        TextBox4.ReadOnly = False
        TextBox5.ReadOnly = False
        TextBox6.ReadOnly = False
        TextBox7.ReadOnly = False
        TextBox8.ReadOnly = False
        TextBox10.ReadOnly = False
        TextBox11.Visible = False
        ComboBox1.Visible = True
        Button7.Visible = True
        Button11.Visible = True
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        TextBox1.Clear()
        TextBox2.Clear()
        TextBox3.Clear()
        TextBox4.Clear()
        TextBox5.Clear()
        TextBox6.Clear()
        TextBox7.Clear()
        TextBox8.Clear()
        TextBox10.Clear()
        TextBox10.Clear()

        TextBox1.ReadOnly = True
        TextBox2.ReadOnly = True
        TextBox3.ReadOnly = True
        TextBox4.ReadOnly = True
        TextBox5.ReadOnly = True
        TextBox6.ReadOnly = True
        TextBox7.ReadOnly = True
        TextBox8.ReadOnly = True
        TextBox10.ReadOnly = True
        TextBox11.Visible = True
        ComboBox1.Visible = False
        Button7.Visible = False
        Button11.Visible = False
        Button5.Enabled = True
        Button6.Enabled = False
        Button8.Enabled = True

    End Sub

    Private Sub SearchToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SearchToolStripMenuItem.Click
        SearchForm.Show()
    End Sub

    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        Me.Close()
    End Sub

    Private Sub SetupToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SetupToolStripMenuItem.Click
        Try
            Setup.Show()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub CreateToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CreateToolStripMenuItem.Click
        CreateTbl.Show()
    End Sub
End Class
