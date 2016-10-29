Public Class CreateTbl
    Dim con As New OleDb.OleDbConnection
    Dim provider, dbsource, dbmydocs, thedatabase, fulldbpath As String
    Dim dbset As New DataSet 'copy of the info from a database
    Dim dbadapter As OleDb.OleDbDataAdapter
    Dim inc, maxrow As Integer

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        provider = "PROVIDER = Microsoft.Jet.OLEDB.4.0;" 'provider technology used to connect to the Database
        thedatabase = "/" & Setup.TextBox2.Text
        dbmydocs = Setup.TextBox1.Text 'fulldoc path
        fulldbpath = dbmydocs & thedatabase 'combine the fullpath and name of the database into one path
        dbsource = "Data Source = " & fulldbpath 'combine data name and data path to make datasource
        con.ConnectionString = provider & dbsource 'sets up connection string (IMPORTANT)
        'convert local data types to ms access
        Dim currency, number, autonumber As String
        Dim usertbl As DataTable = Nothing
        Dim connection As System.Data.OleDb.OleDbConnection = New System.Data.OleDb.OleDbConnection(con.ConnectionString) 'creating new object OLEDB with con

        'one column
        'find a way to improve this later on
        If ComboBox1.SelectedItem = 1 Then
            Try
                Dim sqlquery As String = "CREATE TABLE " & TextBox1.Text & _
           " (" & "id UNIQUE " & "," & Trim(TextBox2.Text) & " " & ComboBox2.SelectedItem & ");"

                If ComboBox2.SelectedItem = "AutoNumber" Then
                    autonumber = New Integer
                    ComboBox2.SelectedItem = autonumber
                End If
                If ComboBox2.SelectedItem = "Number" Then
                    number = "Integer"
                    ComboBox2.SelectedItem = number
                End If
                If ComboBox2.SelectedItem = "Currency" Then
                    currency = "Double"
                    ComboBox2.SelectedItem = currency
                End If

                Dim my_dbConnection As New System.Data.OleDb.OleDbConnection(con.ConnectionString)
                'create a command
                Dim my_Command As New System.Data.OleDb.OleDbCommand(sqlquery, my_dbConnection)
                'connection open
                connection.Open()
                my_dbConnection.Open()
                'command execute
                my_Command.ExecuteNonQuery()
                'close connection
                Dim restrictions() As String = New String(3) {}
                restrictions(3) = "Table"
                'Get list of user tables
                usertbl = connection.GetSchema("Tables", restrictions)
                connection.Close()
                my_dbConnection.Close()
                MessageBox.Show("Table Created!")
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
        End If

        'two columns
        If ComboBox1.SelectedItem = 2 Then
            Try
                'Start CREATE TABLE process code here
                Dim sqlquery As String = "CREATE TABLE " & TextBox1.Text & _
            " (" & Trim(TextBox2.Text) & " " & ComboBox2.SelectedItem & ", " & Trim(TextBox3.Text) & " " & ComboBox3.SelectedItem & ");"


                If ComboBox2.SelectedItem = "AutoNumber" Or ComboBox3.SelectedItem = "AutoNumber"  Then
                    autonumber = New Integer
                    ComboBox2.SelectedItem = autonumber
                    ComboBox3.SelectedItem = autonumber
                End If
                If ComboBox2.SelectedItem = "Number" Or ComboBox3.SelectedItem = "Number" Then
                    number = "Integer"
                    ComboBox2.SelectedItem = number
                    ComboBox3.SelectedItem = number
                End If
                If ComboBox2.SelectedItem = "Currency" Or ComboBox3.SelectedItem = "Currency" Then
                    currency = "Double"
                    ComboBox2.SelectedItem = currency
                    ComboBox3.SelectedItem = currency
                End If

                Dim my_dbConnection As New System.Data.OleDb.OleDbConnection(con.ConnectionString)
                'create a command
                Dim my_Command As New System.Data.OleDb.OleDbCommand(sqlquery, my_dbConnection)
                'connection open
                connection.Open()
                my_dbConnection.Open()
                'command execute
                my_Command.ExecuteNonQuery()
                'close connection
                Dim restrictions() As String = New String(3) {}
                restrictions(3) = "Table"
                'Get list of user tables
                usertbl = connection.GetSchema("Tables", restrictions)
                connection.Close()
                my_dbConnection.Close()
                MessageBox.Show("Table Created!")
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
        End If

        If ComboBox1.SelectedItem = 3 Then
            Dim sqlquery As String = "CREATE TABLE " & TextBox1.Text & _
        " (" & Trim(TextBox2.Text) & " " & ComboBox2.SelectedItem & ", " & Trim(TextBox3.Text) & " " & ComboBox3.SelectedItem & "," _
             & Trim(TextBox4.Text) & " " & ComboBox4.SelectedItem & ");"
            Dim my_dbConnection As New System.Data.OleDb.OleDbConnection(con.ConnectionString)
            'create a command
            Dim my_Command As New System.Data.OleDb.OleDbCommand(sqlquery, my_dbConnection)
            'connection open
            connection.Open()
            my_dbConnection.Open()
            'command execute
            my_Command.ExecuteNonQuery()
            'close connection
            Dim restrictions() As String = New String(3) {}
            restrictions(3) = "Table"
            'Get list of user tables
            usertbl = connection.GetSchema("Tables", restrictions)
            connection.Close()
            my_dbConnection.Close()
            MessageBox.Show("Table Created!")
        End If

        If ComboBox1.SelectedItem = 4 Then
            Dim sqlquery As String = "CREATE TABLE " & TextBox1.Text & _
        " (" & Trim(TextBox2.Text) & " " & ComboBox2.SelectedItem & ", " & Trim(TextBox3.Text) & " " & ComboBox3.SelectedItem & "," _
             & Trim(TextBox4.Text) & " " & ComboBox4.SelectedItem & ", " & Trim(TextBox5.Text) & " " & ComboBox5.SelectedItem & ");"
            Dim my_dbConnection As New System.Data.OleDb.OleDbConnection(con.ConnectionString)
            'create a command
            Dim my_Command As New System.Data.OleDb.OleDbCommand(sqlquery, my_dbConnection)
            'connection open
            connection.Open()
            my_dbConnection.Open()
            'command execute
            my_Command.ExecuteNonQuery()
            'close connection
            Dim restrictions() As String = New String(3) {}
            restrictions(3) = "Table"
            'Get list of user tables
            usertbl = connection.GetSchema("Tables", restrictions)
            connection.Close()
            my_dbConnection.Close()
            MessageBox.Show("Table Created!")
        End If

        If ComboBox1.SelectedItem = 5 Then
            Dim sqlquery As String = "CREATE TABLE " & TextBox1.Text & _
        " (" & Trim(TextBox2.Text) & " " & ComboBox2.SelectedItem & ", " & Trim(TextBox3.Text) & " " & ComboBox3.SelectedItem & "," _
             & Trim(TextBox4.Text) & " " & ComboBox4.SelectedItem & ", " & Trim(TextBox5.Text) & " " & ComboBox5.SelectedItem & "," _
             & Trim(TextBox6.Text) & " " & ComboBox6.SelectedItem & ");"
            Dim my_dbConnection As New System.Data.OleDb.OleDbConnection(con.ConnectionString)
            'create a command
            Dim my_Command As New System.Data.OleDb.OleDbCommand(sqlquery, my_dbConnection)
            'connection open
            connection.Open()
            my_dbConnection.Open()
            'command execute
            my_Command.ExecuteNonQuery()
            'close connection
            Dim restrictions() As String = New String(3) {}
            restrictions(3) = "Table"
            'Get list of user tables
            usertbl = connection.GetSchema("Tables", restrictions)
            connection.Close()
            my_dbConnection.Close()
            MessageBox.Show("Table Created!")
        End If

        If ComboBox1.SelectedItem = 6 Then
            Dim sqlquery As String = "CREATE TABLE " & TextBox1.Text & _
        " (" & Trim(TextBox2.Text) & " " & ComboBox2.SelectedItem & ", " & Trim(TextBox3.Text) & " " & ComboBox3.SelectedItem & "," _
             & Trim(TextBox4.Text) & " " & ComboBox4.SelectedItem & ", " & Trim(TextBox5.Text) & " " & ComboBox5.SelectedItem & "," _
             & Trim(TextBox6.Text) & " " & ComboBox6.SelectedItem & "," & Trim(TextBox7.Text) & " " & ComboBox7.SelectedItem & ");"
            Dim my_dbConnection As New System.Data.OleDb.OleDbConnection(con.ConnectionString)
            'create a command
            Dim my_Command As New System.Data.OleDb.OleDbCommand(sqlquery, my_dbConnection)
            'connection open
            connection.Open()
            my_dbConnection.Open()
            'command execute
            my_Command.ExecuteNonQuery()
            'close connection
            Dim restrictions() As String = New String(3) {}
            restrictions(3) = "Table"
            'Get list of user tables
            usertbl = connection.GetSchema("Tables", restrictions)
            connection.Close()
            my_dbConnection.Close()
            MessageBox.Show("Table Created!")
        End If

        If ComboBox1.SelectedItem = 7 Then
            Dim sqlquery As String = "CREATE TABLE " & TextBox1.Text & _
        " (" & Trim(TextBox2.Text) & " " & ComboBox2.SelectedItem & ", " & Trim(TextBox3.Text) & " " & ComboBox3.SelectedItem & "," _
             & Trim(TextBox4.Text) & " " & ComboBox4.SelectedItem & ", " & Trim(TextBox5.Text) & " " & ComboBox5.SelectedItem & "," _
             & Trim(TextBox6.Text) & " " & ComboBox6.SelectedItem & "," & Trim(TextBox7.Text) & " " & ComboBox7.SelectedItem & "," _
             & Trim(TextBox8.Text) & " " & ComboBox8.SelectedItem & ");"
            Dim my_dbConnection As New System.Data.OleDb.OleDbConnection(con.ConnectionString)
            'create a command
            Dim my_Command As New System.Data.OleDb.OleDbCommand(sqlquery, my_dbConnection)
            'connection open
            connection.Open()
            my_dbConnection.Open()
            'command execute
            my_Command.ExecuteNonQuery()
            'close connection
            Dim restrictions() As String = New String(3) {}
            restrictions(3) = "Table"
            'Get list of user tables
            usertbl = connection.GetSchema("Tables", restrictions)
            connection.Close()
            my_dbConnection.Close()
            MessageBox.Show("Table Created!")
        End If

        If ComboBox1.SelectedItem = 8 Then
            Dim sqlquery As String = "CREATE TABLE " & TextBox1.Text & _
        " (" & Trim(TextBox2.Text) & " " & ComboBox2.SelectedItem & ", " & Trim(TextBox3.Text) & " " & ComboBox3.SelectedItem & "," _
             & Trim(TextBox4.Text) & " " & ComboBox4.SelectedItem & ", " & Trim(TextBox5.Text) & " " & ComboBox5.SelectedItem & "," _
             & Trim(TextBox6.Text) & " " & ComboBox6.SelectedItem & "," & Trim(TextBox7.Text) & " " & ComboBox7.SelectedItem & "," _
             & Trim(TextBox8.Text) & " " & ComboBox8.SelectedItem & "," & Trim(TextBox9.Text) & " " & ComboBox9.SelectedItem & ");"
            Dim my_dbConnection As New System.Data.OleDb.OleDbConnection(con.ConnectionString)
            'create a command
            Dim my_Command As New System.Data.OleDb.OleDbCommand(sqlquery, my_dbConnection)
            'connection open
            connection.Open()
            my_dbConnection.Open()
            'command execute
            my_Command.ExecuteNonQuery()
            'close connection
            Dim restrictions() As String = New String(3) {}
            restrictions(3) = "Table"
            'Get list of user tables
            usertbl = connection.GetSchema("Tables", restrictions)
            connection.Close()
            my_dbConnection.Close()
            MessageBox.Show("Table Created!")
        End If

        If ComboBox1.SelectedItem = 9 Then
            Dim sqlquery As String = "CREATE TABLE " & TextBox1.Text & _
        " (" & Trim(TextBox2.Text) & " " & ComboBox2.SelectedItem & ", " & Trim(TextBox3.Text) & " " & ComboBox3.SelectedItem & "," _
             & Trim(TextBox4.Text) & " " & ComboBox4.SelectedItem & ", " & Trim(TextBox5.Text) & " " & ComboBox5.SelectedItem & "," _
             & Trim(TextBox6.Text) & " " & ComboBox6.SelectedItem & "," & Trim(TextBox7.Text) & " " & ComboBox7.SelectedItem & "," _
             & Trim(TextBox8.Text) & " " & ComboBox8.SelectedItem & "," & Trim(TextBox9.Text) & " " & ComboBox9.SelectedItem & "," _
             & Trim(TextBox10.Text) & " " & ComboBox10.SelectedItem & ");"
            Dim my_dbConnection As New System.Data.OleDb.OleDbConnection(con.ConnectionString)
            'create a command
            Dim my_Command As New System.Data.OleDb.OleDbCommand(sqlquery, my_dbConnection)
            'connection open
            connection.Open()
            my_dbConnection.Open()
            'command execute
            my_Command.ExecuteNonQuery()
            'close connection
            Dim restrictions() As String = New String(3) {}
            restrictions(3) = "Table"
            'Get list of user tables
            usertbl = connection.GetSchema("Tables", restrictions)
            connection.Close()
            my_dbConnection.Close()
            MessageBox.Show("Table Created!")
        End If

        If ComboBox1.SelectedItem = 10 Then
            Dim sqlquery As String = "CREATE TABLE " & TextBox1.Text & _
        " (" & Trim(TextBox2.Text) & " " & ComboBox2.SelectedItem & ", " & Trim(TextBox3.Text) & " " & ComboBox3.SelectedItem & "," _
             & Trim(TextBox4.Text) & " " & ComboBox4.SelectedItem & ", " & Trim(TextBox5.Text) & " " & ComboBox5.SelectedItem & "," _
             & Trim(TextBox6.Text) & " " & ComboBox6.SelectedItem & "," & Trim(TextBox7.Text) & " " & ComboBox7.SelectedItem & "," _
             & Trim(TextBox8.Text) & " " & ComboBox8.SelectedItem & "," & Trim(TextBox9.Text) & " " & ComboBox9.SelectedItem & "," _
             & Trim(TextBox10.Text) & " " & ComboBox10.SelectedItem & "," & Trim(TextBox11.Text) & " " & ComboBox11.SelectedItem & "," _
             & ");"
            Dim my_dbConnection As New System.Data.OleDb.OleDbConnection(con.ConnectionString)
            'create a command
            Dim my_Command As New System.Data.OleDb.OleDbCommand(sqlquery, my_dbConnection)
            'connection open
            connection.Open()
            my_dbConnection.Open()
            'command execute
            my_Command.ExecuteNonQuery()
            'close connection
            Dim restrictions() As String = New String(3) {}
            restrictions(3) = "Table"
            'Get list of user tables
            usertbl = connection.GetSchema("Tables", restrictions)
            connection.Close()
            my_dbConnection.Close()
            MessageBox.Show("Table Created!")
        End If

        If ComboBox1.SelectedItem = 11 Then
            Dim sqlquery As String = "CREATE TABLE " & TextBox1.Text & _
        " (" & Trim(TextBox2.Text) & " " & ComboBox2.SelectedItem & ", " & Trim(TextBox3.Text) & " " & ComboBox3.SelectedItem & "," _
             & Trim(TextBox4.Text) & " " & ComboBox4.SelectedItem & ", " & Trim(TextBox5.Text) & " " & ComboBox5.SelectedItem & "," _
             & Trim(TextBox6.Text) & " " & ComboBox6.SelectedItem & "," & Trim(TextBox7.Text) & " " & ComboBox7.SelectedItem & "," _
             & Trim(TextBox8.Text) & " " & ComboBox8.SelectedItem & "," & Trim(TextBox9.Text) & " " & ComboBox9.SelectedItem & "," _
             & Trim(TextBox10.Text) & " " & ComboBox10.SelectedItem & "," & Trim(TextBox11.Text) & " " & ComboBox11.SelectedItem & "," _
             & Trim(TextBox12.Text) & " " & ComboBox12.SelectedItem & ");"
            Dim my_dbConnection As New System.Data.OleDb.OleDbConnection(con.ConnectionString)
            'create a command
            Dim my_Command As New System.Data.OleDb.OleDbCommand(sqlquery, my_dbConnection)
            'connection open
            connection.Open()
            my_dbConnection.Open()
            'command execute
            my_Command.ExecuteNonQuery()
            'close connection
            Dim restrictions() As String = New String(3) {}
            restrictions(3) = "Table"
            'Get list of user tables
            usertbl = connection.GetSchema("Tables", restrictions)
            connection.Close()
            my_dbConnection.Close()
            MessageBox.Show("Table Created!")
        End If

        If ComboBox1.SelectedItem = 12 Then
            Dim sqlquery As String = "CREATE TABLE " & TextBox1.Text & _
        " (" & Trim(TextBox2.Text) & " " & ComboBox2.SelectedItem & ", " & Trim(TextBox3.Text) & " " & ComboBox3.SelectedItem & "," _
             & Trim(TextBox4.Text) & " " & ComboBox4.SelectedItem & ", " & Trim(TextBox5.Text) & " " & ComboBox5.SelectedItem & "," _
             & Trim(TextBox6.Text) & " " & ComboBox6.SelectedItem & "," & Trim(TextBox7.Text) & " " & ComboBox7.SelectedItem & "," _
             & Trim(TextBox8.Text) & " " & ComboBox8.SelectedItem & "," & Trim(TextBox9.Text) & " " & ComboBox9.SelectedItem & "," _
             & Trim(TextBox10.Text) & " " & ComboBox10.SelectedItem & "," & Trim(TextBox11.Text) & " " & ComboBox11.SelectedItem & "," _
             & Trim(TextBox12.Text) & " " & ComboBox12.SelectedItem & "," & Trim(TextBox13.Text) & " " & ComboBox13.SelectedItem & "," _
             & ");"
            Dim my_dbConnection As New System.Data.OleDb.OleDbConnection(con.ConnectionString)
            'create a command
            Dim my_Command As New System.Data.OleDb.OleDbCommand(sqlquery, my_dbConnection)
            'connection open
            connection.Open()
            my_dbConnection.Open()
            'command execute
            my_Command.ExecuteNonQuery()
            'close connection
            Dim restrictions() As String = New String(3) {}
            restrictions(3) = "Table"
            'Get list of user tables
            usertbl = connection.GetSchema("Tables", restrictions)
            connection.Close()
            my_dbConnection.Close()
            MessageBox.Show("Table Created!")
        End If

        If ComboBox1.SelectedItem = 13 Then
            Dim sqlquery As String = "CREATE TABLE " & TextBox1.Text & _
        " (" & "ID CONSTRAINT PRIMARY KEY " & "," & Trim(TextBox2.Text) & " " & ComboBox2.SelectedItem & ", " & Trim(TextBox3.Text) & " " & ComboBox3.SelectedItem & "," _
             & Trim(TextBox4.Text) & " " & ComboBox4.SelectedItem & ", " & Trim(TextBox5.Text) & " " & ComboBox5.SelectedItem & "," _
             & Trim(TextBox6.Text) & " " & ComboBox6.SelectedItem & "," & Trim(TextBox7.Text) & " " & ComboBox7.SelectedItem & "," _
             & Trim(TextBox8.Text) & " " & ComboBox8.SelectedItem & "," & Trim(TextBox9.Text) & " " & ComboBox9.SelectedItem & "," _
             & Trim(TextBox10.Text) & " " & ComboBox10.SelectedItem & "," & Trim(TextBox11.Text) & " " & ComboBox11.SelectedItem & "," _
             & Trim(TextBox12.Text) & " " & ComboBox12.SelectedItem & "," & Trim(TextBox13.Text) & " " & ComboBox13.SelectedItem & "," _
             & Trim(TextBox14.Text) & " " & ComboBox14.SelectedItem & ");"

            Dim my_dbConnection As New System.Data.OleDb.OleDbConnection(con.ConnectionString)
            'create a command
            Dim my_Command As New System.Data.OleDb.OleDbCommand(sqlquery, my_dbConnection)
            'connection open
            connection.Open()
            my_dbConnection.Open()
            'command execute
            my_Command.ExecuteNonQuery()
            'close connection
            Dim restrictions() As String = New String(3) {}
            restrictions(3) = "Table"
            'Get list of user tables
            usertbl = connection.GetSchema("Tables", restrictions)
            connection.Close()
            my_dbConnection.Close()
            MessageBox.Show("Table Created!")
        End If

        If ComboBox1.SelectedItem = 14 Then
            Dim sqlquery As String = "CREATE TABLE " & TextBox1.Text & _
        " (" & Trim(TextBox2.Text) & " " & ComboBox2.SelectedItem & ", " & Trim(TextBox3.Text) & " " & ComboBox3.SelectedItem & "," _
             & Trim(TextBox4.Text) & " " & ComboBox4.SelectedItem & ", " & Trim(TextBox5.Text) & " " & ComboBox5.SelectedItem & "," _
             & Trim(TextBox6.Text) & " " & ComboBox6.SelectedItem & "," & Trim(TextBox7.Text) & " " & ComboBox7.SelectedItem & "," _
             & Trim(TextBox8.Text) & " " & ComboBox8.SelectedItem & "," & Trim(TextBox9.Text) & " " & ComboBox9.SelectedItem & "," _
             & Trim(TextBox10.Text) & " " & ComboBox10.SelectedItem & "," & Trim(TextBox11.Text) & " " & ComboBox11.SelectedItem & "," _
             & Trim(TextBox12.Text) & " " & ComboBox12.SelectedItem & "," & Trim(TextBox13.Text) & " " & ComboBox13.SelectedItem & "," _
             & Trim(TextBox14.Text) & " " & ComboBox14.SelectedItem & "," & Trim(TextBox15.Text) & " " & ComboBox15.SelectedItem & "," _
             & ");"

            Dim my_dbConnection As New System.Data.OleDb.OleDbConnection(con.ConnectionString)
            'create a command
            Dim my_Command As New System.Data.OleDb.OleDbCommand(sqlquery, my_dbConnection)
            'connection open
            connection.Open()
            my_dbConnection.Open()
            'command execute
            my_Command.ExecuteNonQuery()
            'close connection
            Dim restrictions() As String = New String(3) {}
            restrictions(3) = "Table"
            'Get list of user tables
            usertbl = connection.GetSchema("Tables", restrictions)
            connection.Close()
            my_dbConnection.Close()
            MessageBox.Show("Table Created!")
        End If

        If ComboBox1.SelectedItem = 15 Then
            Dim sqlquery As String = "CREATE TABLE " & TextBox1.Text & _
        " (" & Trim(TextBox2.Text) & " " & ComboBox2.SelectedItem & ", " & Trim(TextBox3.Text) & " " & ComboBox3.SelectedItem & "," _
             & Trim(TextBox4.Text) & " " & ComboBox4.SelectedItem & ", " & Trim(TextBox5.Text) & " " & ComboBox5.SelectedItem & "," _
             & Trim(TextBox6.Text) & " " & ComboBox6.SelectedItem & "," & Trim(TextBox7.Text) & " " & ComboBox7.SelectedItem & "," _
             & Trim(TextBox8.Text) & " " & ComboBox8.SelectedItem & "," & Trim(TextBox9.Text) & " " & ComboBox9.SelectedItem & "," _
             & Trim(TextBox10.Text) & " " & ComboBox10.SelectedItem & "," & Trim(TextBox11.Text) & " " & ComboBox11.SelectedItem & "," _
             & Trim(TextBox12.Text) & " " & ComboBox12.SelectedItem & "," & Trim(TextBox13.Text) & " " & ComboBox13.SelectedItem & "," _
             & Trim(TextBox14.Text) & " " & ComboBox14.SelectedItem & "," & Trim(TextBox15.Text) & " " & ComboBox15.SelectedItem & "," _
             & Trim(TextBox15.Text) & " " & ComboBox15.SelectedItem & ");"

            Dim my_dbConnection As New System.Data.OleDb.OleDbConnection(con.ConnectionString)
            'create a command
            Dim my_Command As New System.Data.OleDb.OleDbCommand(sqlquery, my_dbConnection)
            'connection open
            connection.Open()
            my_dbConnection.Open()
            'command execute
            my_Command.ExecuteNonQuery()
            'close connection
            Dim restrictions() As String = New String(3) {}
            restrictions(3) = "Table"
            'Get list of user tables
            usertbl = connection.GetSchema("Tables", restrictions)
            connection.Close()
            my_dbConnection.Close()
            MessageBox.Show("Table Created!")
        End If

        Try
            '    'CHECK TEXTBOXES!!!
            '    ' Start CREATE TABLE process code here
            '    Dim sqlquery As String = "CREATE TABLE " & TextBox1.Text & _
            '" (" & Trim(TextBox2.Text) & " " & ComboBox2.SelectedItem & ", " & Trim(TextBox3.Text) & " " & ComboBox3.SelectedItem & "," _
            '     & Trim(TextBox4.Text) & " " & ComboBox4.SelectedItem & ", " & Trim(TextBox5.Text) & " " & ComboBox5.SelectedItem & "," _
            '     & Trim(TextBox6.Text) & " " & ComboBox6.SelectedItem & "," & Trim(TextBox7.Text) & " " & ComboBox7.SelectedItem & "," _
            '     & Trim(TextBox8.Text) & " " & ComboBox8.SelectedItem & "," & Trim(TextBox9.Text) & " " & ComboBox9.SelectedItem & "," _
            '     & Trim(TextBox10.Text) & " " & ComboBox10.SelectedItem & "," & Trim(TextBox11.Text) & " " & ComboBox11.SelectedItem & "," _
            '     & Trim(TextBox12.Text) & " " & ComboBox12.SelectedItem & "," & Trim(TextBox13.Text) & " " & ComboBox13.SelectedItem & "," _
            '     & Trim(TextBox14.Text) & " " & ComboBox14.SelectedItem & "," & Trim(TextBox15.Text) & " " & ComboBox15.SelectedItem & "," _
            '     & ");"


            If ComboBox2.SelectedItem = "AutoNumber" Or ComboBox3.SelectedItem = "AutoNumber" Or ComboBox4.SelectedItem = "AutoNumber" _
                Or ComboBox5.SelectedItem = "AutoNumber" Or ComboBox6.SelectedItem = "AutoNumber" Or ComboBox7.SelectedItem = "AutoNumber" _
               Or ComboBox8.SelectedItem = "AutoNumber" Or ComboBox9.SelectedItem = "AutoNumber" Or ComboBox10.SelectedItem = "AutoNumber" _
               Or ComboBox11.SelectedItem = "AutoNumber" Or ComboBox12.SelectedItem = "AutoNumber" Or ComboBox13.SelectedItem = "AutoNumber" _
               Or ComboBox14.SelectedItem = "AutoNumber" Or ComboBox15.SelectedItem = "AutoNumber" Then
                autonumber = New Integer
                ComboBox2.SelectedItem = autonumber
            End If
            If ComboBox2.SelectedItem = "Number" Or ComboBox3.SelectedItem = "Number" Or ComboBox4.SelectedItem = "Number" _
                Or ComboBox5.SelectedItem = "Number" Or ComboBox6.SelectedItem = "Number" Or ComboBox7.SelectedItem = "Number" _
               Or ComboBox8.SelectedItem = "Number" Or ComboBox9.SelectedItem = "Number" Or ComboBox10.SelectedItem = "Number" _
               Or ComboBox11.SelectedItem = "Number" Or ComboBox12.SelectedItem = "Number" Or ComboBox13.SelectedItem = "Number" _
               Or ComboBox14.SelectedItem = "Number" Or ComboBox15.SelectedItem = "Number" Then
                number = "Integer"
                ComboBox2.SelectedItem = number
            End If
            If ComboBox2.SelectedItem = "Currency" Or ComboBox3.SelectedItem = "Currency" Or ComboBox4.SelectedItem = "Currency" _
                Or ComboBox5.SelectedItem = "Currency" Or ComboBox6.SelectedItem = "Currency" Or ComboBox7.SelectedItem = "Currency" _
               Or ComboBox8.SelectedItem = "Currency" Or ComboBox9.SelectedItem = "Currency" Or ComboBox10.SelectedItem = "Currency" _
               Or ComboBox11.SelectedItem = "Currency" Or ComboBox12.SelectedItem = "Currency" Or ComboBox13.SelectedItem = "Currency" _
               Or ComboBox14.SelectedItem = "Currency" Or ComboBox15.SelectedItem = "Currency" Then
                currency = "Double"
                ComboBox2.SelectedItem = currency
            End If


            'Dim query As String = "CREATE TABLE NewTable " _
            ' & "(FirstName CHAR, LastName CHAR, " _
            '' & "SSN INTEGER CONSTRAINT MyFieldConstraint " _
            ' & "PRIMARY KEY);"

            'Dim my_dbConnection As New System.Data.OleDb.OleDbConnection(con.ConnectionString)
            ''create a command
            'Dim my_Command As New System.Data.OleDb.OleDbCommand(sqlquery, my_dbConnection)
            ''connection open
            'connection.Open()
            'my_dbConnection.Open()
            ''command execute
            'my_Command.ExecuteNonQuery()
            ''close connection
            'Dim restrictions() As String = New String(3) {}
            'restrictions(3) = "Table"
            ''Get list of user tables
            'usertbl = connection.GetSchema("Tables", restrictions)
            'connection.Close()
            'my_dbConnection.Close()
            'MessageBox.Show("Table Created!")
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

End Class