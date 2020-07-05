Imports System.Data.OleDb
Imports System.IO
Imports System.Threading
Imports jtl_update_ZahlungsArt.shared_variables
Imports System.Net

Imports System.Data.SqlClient
Imports System.Xml

Imports System.Text

Imports System.Web
Imports System.Text.RegularExpressions
Imports System.Security.Cryptography

Imports Microsoft.Win32
Imports Microsoft.Win32.Registry


Public Class Form1
    Dim oconn, oconn2 As Data.OleDb.OleDbConnection
    Dim oconn3, oconn4 As Data.SqlClient.SqlConnection
    Dim ocmd, ocmd2 As Data.OleDb.OleDbCommand
    Dim ocmd3, ocmd4 As Data.SqlClient.SqlCommand
    Dim odr, odr2 As Data.OleDb.OleDbDataReader
    Dim odr3, odr4 As Data.SqlClient.SqlDataReader
    Dim dt, dt2 As New DataTable
    Dim start_before, start_before2 As Boolean
    Dim regKey As RegistryKey

    Public Function check_db_connection() As Boolean
        'check db connection
        Try
            connect()

            oconn3 = New Data.SqlClient.SqlConnection(conect2)
            oconn3.Open()
            oconn3.Close()


            '''''  MsgBox("Datenbank Verbindung erfolgreich hergestellt!")

            Return True
        Catch ex As Exception
            MsgBox("Test der Datenbankverbindung fehlgeschlagen, bitte Verbindungsdaten korrigieren!")

            Return False
        End Try

    End Function

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim count, kZahlungsArt As Integer

        count = 0


        If Not IsNumeric(TextBox1.Text) Then
            MsgBox("Bitte geben Sie die Nummer ein ?")
        Else

            kZahlungsArt = Int(TextBox1.Text)

            connect()

            oconn3 = New Data.SqlClient.SqlConnection(conect2)
            oconn3.Open()
            ocmd3 = New Data.SqlClient.SqlCommand()


            With ocmd3
                .Connection = oconn3
                .CommandType = Data.CommandType.Text

                Try

                    .CommandText = "select count(*) from  [eazybusiness].[dbo].[tBestellung]  where kZahlungsArt='" & kZahlungsArt & "'"


                    odr3 = .ExecuteReader()
                    While odr3.Read
                        Try
                            count = odr3.Item(0)
                        Catch ex As Exception

                        End Try
                    End While
                    odr3.Close()

                Catch ex As Exception
                    MsgBox(ex.ToString)
                End Try


            End With

            oconn3.Close()


        End If

        Label7.Text = "die Anzahl der Datensätze für dieses kZahlungsArt : " + Str(count)
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim count, kZahlungsArt, kZahlungsArt_new, b As Integer

        count = 0


        If Not IsNumeric(TextBox1.Text) Then
            MsgBox("Bitte geben Sie die Nummer ein ?")
        Else


            b = MsgBox("Aktualisieren Sie diese kZahlungsArt-Werte unbedingt ?", MsgBoxStyle.YesNo)
            If b = 6 Then
                'continue in this process
            Else
                'exit the debugging
                Exit Sub
            End If

            Button1.Text = "Warten Sie mal"
            Button1.Enabled = False
            Button2.Enabled = False
            Button4.Enabled = False
            Button5.Enabled = False


            kZahlungsArt = Int(TextBox1.Text)

            Try
                kZahlungsArt_new = Int(DataGridView1.SelectedRows.Item(0).Cells(0).Value)
            Catch ex As Exception
                kZahlungsArt_new = 0
            End Try

            connect()

            oconn3 = New Data.SqlClient.SqlConnection(conect2)
            oconn3.Open()
            ocmd3 = New Data.SqlClient.SqlCommand()


            With ocmd3
                .Connection = oconn3
                .CommandType = Data.CommandType.Text
                .CommandTimeout = 0

                '    Label7.Text = "Warten Sie mal....."

                Try
                    ' update the records in db
                    .CommandText = "update  [eazybusiness].[dbo].[tBestellung]  set kZahlungsArt='" & kZahlungsArt_new & "'  where kZahlungsArt='" & kZahlungsArt & "'"


                    .ExecuteNonQuery()


                Catch ex As Exception
                    MsgBox(ex.ToString)
                End Try


            End With

            oconn3.Close()

            Label7.Text = "Getan....."

            Button1.Text = "Aktualisieren"
            Button1.Enabled = True
            Button2.Enabled = True
            Button4.Enabled = True
            Button5.Enabled = True
        End If


    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim count, kZahlungsArt, kZahlungsArt_new, b As Integer
        Dim cName As String


        count = 0


        If Not IsNumeric(TextBox1.Text) Then
            MsgBox("Bitte geben Sie die Nummer ein ?")
        Else


            b = MsgBox("Aktualisieren Sie diese kZahlungsArt-Werte unbedingt ?", MsgBoxStyle.YesNo)
            If b = 6 Then
                'continue in this process
            Else
                'exit the debugging
                Exit Sub
            End If

            Button1.Text = "Warten Sie mal"
            Button1.Enabled = False
            Button2.Enabled = False
            Button3.Enabled = False
            Button4.Enabled = False
            Button5.Enabled = False


            kZahlungsArt = Int(TextBox1.Text)

            Try
                kZahlungsArt_new = Int(DataGridView1.SelectedRows.Item(0).Cells(0).Value)
            Catch ex As Exception
                kZahlungsArt_new = 0
            End Try
            Try
                cName = DataGridView1.SelectedRows.Item(0).Cells(1).Value
            Catch ex As Exception
                cName = ""
            End Try

            connect()

            oconn3 = New Data.SqlClient.SqlConnection(conect2)
            oconn3.Open()
            ocmd3 = New Data.SqlClient.SqlCommand()


            With ocmd3
                .Connection = oconn3
                .CommandType = Data.CommandType.Text
                .CommandTimeout = 0

                '    Label7.Text = "Warten Sie mal....."

                Try
                    ' update the records in db for tBestellung table
                    .CommandText = "update  [eazybusiness].[dbo].[tBestellung]  set kZahlungsArt='" & kZahlungsArt_new & "'  where kZahlungsArt='" & kZahlungsArt & "'"

                    .ExecuteNonQuery()


                    ' update the records in db for tZahlung table
                    .CommandText = "update  [eazybusiness].[dbo].[tZahlung]  set kZahlungsart='" & kZahlungsArt_new & "',cName='" + cName + "'  where kZahlungsart='" & kZahlungsArt & "'"

                    .ExecuteNonQuery()


                Catch ex As Exception
                    MsgBox(ex.ToString)
                End Try


            End With

            oconn3.Close()

            Label7.Text = "Getan....."

            Button1.Text = "Aktualisieren"
            Button1.Enabled = False
            Button2.Enabled = True
            Button3.Enabled = True
            Button4.Enabled = True
            Button5.Enabled = True
        End If
    End Sub

    Public Sub fill_datagridview1_fromdb()
        connect()

        'show values for [tZahlungsart] in datagridview
        Dim myConnection As OleDbConnection = New OleDbConnection
        myConnection.ConnectionString = conect1


        Dim da As OleDbDataAdapter

        'set my query
        da = New OleDbDataAdapter("select [kZahlungsart]
      ,[cName]
      from  [eazybusiness].[dbo].[tZahlungsart]   order by [kZahlungsart]", myConnection)

        'where [fSteuersatz]<>0



        'create a new dataset 
        Dim ds As DataSet = New DataSet
        ' fill dataset 
        Try
            da.Fill(ds, "[eazybusiness].[dbo].[tZahlungsart]")

            DataGridView1.DataSource = ds.Tables(0)
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try



        DataGridView1.Columns(1).Width = 200
    End Sub


    Private Sub saveData()
        Dim server, db, user, pass, date1, date2 As String
        Dim tax_value, tax_value2 As Integer

        Dim obj__dx_settings As New dx_settings
        Dim ret As String = ""
        Dim tmp_var As String = ""

        Try


            server = TextBox2.Text
            db = TextBox3.Text
            user = TextBox4.Text
            pass = TextBox5.Text





            obj__dx_settings.deleteSettingsFile() ' datei loeschen

            ret = obj__dx_settings.initSettingsFile()

            obj__dx_settings.addSetting("mssql_server_database", server)
            obj__dx_settings.addSetting("mssql_username", user)
            obj__dx_settings.addSetting("mssql_password", pass)

            obj__dx_settings.addSetting("tb_mssql_database", db)




        Catch ex As Exception

        End Try
    End Sub

    Public Sub writeErrorLog(ByVal p_txt As String)

        Try



            Dim file As System.IO.StreamWriter
                file = My.Computer.FileSystem.OpenTextFileWriter(Application.StartupPath() + "\errorlog.txt", True)
                file.WriteLine(Date.Now & " | " & p_txt)
                file.Close()



        Catch ex As Exception
            ' nix
        End Try

    End Sub



    Public Sub show_msg()

        MsgBox("Bitte warten Sie, bis die Suche abgeschlossen ist .....")

    End Sub





    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        saveData()

    End Sub



    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim d1 As String
        Dim check_db_status As Boolean

        check_db_status = False

        'call connect method to connect to sql server in shared_variables file
        connect()

        Try
            TextBox2.Text = server
            TextBox3.Text = db
            TextBox4.Text = user
            TextBox5.Text = pass


        Catch ex As Exception

        End Try
        'check db connection
        check_db_status = check_db_connection()

        If check_db_status = False Then
            Exit Sub
        End If



        DataGridView1.Width = Me.Width - DataGridView1.Left - 50
        RichTextBox1.Width = Me.Width - RichTextBox1.Left - 50

        Label9.Left = Me.Right - 200


        'fill datagridview1
        fill_datagridview1_fromdb()



        'create columns of roles datagrid view
        'dt.Columns.Add("kSteuersatz", GetType(Integer))
        'dt.Columns.Add("fSteuersatz_old", GetType(Decimal))
        'dt.Columns.Add("fSteuersatz_new", GetType(Decimal))






    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim path As String

        If TextBox2.Text = "" Or TextBox3.Text = "" Or TextBox4.Text = "" Then
            MsgBox("Bitte füllen Sie alle Felder für Verbindungsdaten mit db aus ??")
            Exit Sub
        End If


        server = TextBox2.Text
        db = TextBox3.Text
        user = TextBox4.Text
        pass = TextBox5.Text

        saveData()

        'check db connection
        Try
            connect()

            oconn3 = New Data.SqlClient.SqlConnection(conect2)
            oconn3.Open()
            oconn3.Close()
            MsgBox("Datenbank Verbindung erfolgreich hergestellt!")
        Catch ex As Exception
            MsgBox("Test der Datenbankverbindung fehlgeschlagen, bitte Verbindungsdaten korrigieren!")
        End Try






    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Application.Exit()

    End Sub
End Class
