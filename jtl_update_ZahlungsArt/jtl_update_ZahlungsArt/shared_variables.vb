Public Class shared_variables
    Public Shared conect1, conect2, conectionstring, server, db, user, pass, sql_qry, sql_qry2, date1, date2 As String
    Public Shared tax_value, tax_value2, start_with_windows_status As Integer
    Public Shared Sub connect()

        Dim obj__dx_settings As New dx_settings
        Dim tmp_ret As String = ""
        Dim tmp_var As String = ""

        tmp_ret = obj__dx_settings.initSettingsFile()

        tmp_var = obj__dx_settings.getSetting("mssql_server_database")
        If tmp_var <> obj__dx_settings.ret_notFound Then
            server = tmp_var
        End If

        tmp_var = obj__dx_settings.getSetting("mssql_username")
        If tmp_var <> obj__dx_settings.ret_notFound Then
            user = tmp_var
        End If


        tmp_var = obj__dx_settings.getSetting("mssql_password")
        If tmp_var <> obj__dx_settings.ret_notFound Then
            pass = tmp_var
        End If

        tmp_var = obj__dx_settings.getSetting("tb_mssql_database")
        If tmp_var <> obj__dx_settings.ret_notFound Then
            db = tmp_var
        End If



        'tmp_var = obj__dx_settings.getSetting("tax_value")
        'If tmp_var <> obj__dx_settings.ret_notFound Then
        '    tax_value = tmp_var
        'End If

        'tmp_var = obj__dx_settings.getSetting("tax_value2")
        'If tmp_var <> obj__dx_settings.ret_notFound Then
        '    tax_value2 = tmp_var
        'End If

        tmp_var = obj__dx_settings.getSetting("date1")
        If tmp_var <> obj__dx_settings.ret_notFound Then
            date1 = tmp_var
        End If

        tmp_var = obj__dx_settings.getSetting("date2")
        If tmp_var <> obj__dx_settings.ret_notFound Then
            date2 = tmp_var
        End If


        tmp_var = obj__dx_settings.getSetting("start_with_windows_status")
        If tmp_var <> obj__dx_settings.ret_notFound Then
            start_with_windows_status = tmp_var
        End If


        'prepair connection string to sql db
        'oledb connection
        conect1 = "provider=SQLOLEDB.1;data source=" & server
        conect1 = conect1 & ";database=" & db
        conect1 = conect1 & " ;user id=" & user
        conect1 = conect1 & " ;password= " & pass
        conect1 = conect1 & ";Initial Catalog=" & db & ";"



        'sqldbconnection
        conect2 = "data source=" & server
        conect2 = conect2 & ";database=" & db
        conect2 = conect2 & " ;user id=" & user
        conect2 = conect2 & " ;password= " & pass
        conect2 = conect2 & ";"



        '  conectionstring = "Server=khaled-pc\sqlexpress;database=Erasito;user id=sa;password=flah2019;Initial Catalog=Erasito;"


    End Sub

End Class
