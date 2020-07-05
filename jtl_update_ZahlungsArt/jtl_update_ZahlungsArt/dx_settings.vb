Imports System
Imports System.Management
Imports System.Xml

Public Class dx_settings

    'retval-overhaul, general-settings-autocreate

    Public ReadOnly ret_fileNameInvalid As String = "fileNameInvalid"
    Public ReadOnly ret_fileCreated As String = "fileCreated"
    Public ReadOnly ret_fileFound As String = "fileFound"
    Public ReadOnly ret_fileError As String = "fileError"
    Public ReadOnly ret_success As String = "success"
    Public ReadOnly ret_notFound As String = "notFound"
    Public ReadOnly ret_alreadyExists As String = "alreadyExists"
    Public ReadOnly ret_notInitialized As String = "notInitialized"

    Public ReadOnly standard_filepath As String = Application.StartupPath() + "\settings.xml"

    Private settingsFilePath As String = ""
    Private lastError As String = ""

    Private encoding As New System.Text.UTF8Encoding

    Public Sub deleteSettingsFile()
        If System.IO.File.Exists(standard_filepath) = True Then
            Try
                System.IO.File.Delete(standard_filepath)
            Catch ex As Exception

            End Try
        End If
    End Sub

    Public Function initSettingsFile(Optional ByVal filepath As String = "standard") As String

        Dim fileStream As System.IO.FileStream
        Dim xmlWriter As XmlTextWriter

        If filepath = "standard" Then
            filepath = Me.standard_filepath
        End If

        If validateFilePath(filepath) = False Then
            Return ret_fileNameInvalid
        End If

        If System.IO.File.Exists(filepath) = False Then
            Try
                fileStream = System.IO.File.Create(filepath)
                fileStream.Close()

                xmlWriter = New XmlTextWriter(filepath, Me.encoding)
                xmlWriter.Formatting = Formatting.Indented
                xmlWriter.WriteStartDocument()
                xmlWriter.WriteStartElement("root")
                xmlWriter.WriteStartElement("general_settings")
                xmlWriter.WriteFullEndElement()
                xmlWriter.WriteFullEndElement()
                xmlWriter.WriteEndDocument()
                xmlWriter.Close()

                Me.settingsFilePath = filepath
                Return ret_fileCreated
            Catch ex As Exception
                Me.lastError = ex.ToString()
                Return ret_fileError
            End Try
        Else
            Me.settingsFilePath = filepath
            Return ret_fileFound
        End If

    End Function

    Public Function addSetting(ByVal key As String, ByVal value As String, Optional ByVal branch As String = "general_settings", Optional autoCreateBranch As Boolean = False) As String

        Dim xmlDocument As XmlDocument = New XmlDocument()
        Dim branchNode As XmlNode
        Dim settingNode As XmlNode
        Dim rootNode As XmlNode
        Dim xpathBranch As String = "root/" + branch
        Dim xpathSetting As String = "root/" + branch + "/" + key

        If Me.settingsFilePath = "" Then
            Return ret_notInitialized
        End If

        Try
            xmlDocument.Load(Me.settingsFilePath)

            If xmlDocument.SelectSingleNode(xpathBranch) Is Nothing Then
                If autoCreateBranch = True Then
                    rootNode = xmlDocument.SelectSingleNode("root")
                    branchNode = xmlDocument.CreateNode(XmlNodeType.Element, branch, "")
                    rootNode.AppendChild(branchNode)
                Else
                    Return ret_notFound
                End If
            End If

            If xmlDocument.SelectSingleNode(xpathSetting) Is Nothing Then

                settingNode = xmlDocument.CreateNode(XmlNodeType.Element, key, "")
                settingNode.InnerText = value

                branchNode = xmlDocument.SelectSingleNode(xpathBranch)
                branchNode.AppendChild(settingNode)

                xmlDocument.Save(Me.settingsFilePath)
                Return ret_success

            Else
                Return ret_alreadyExists
            End If
        Catch ex As Exception
            Me.lastError = ex.ToString()
            Return ret_fileError
        End Try

    End Function

    Public Function getSetting(ByVal key As String, Optional ByVal branch As String = "general_settings") As String

        Dim xmlDocument As XmlDocument = New XmlDocument()
        Dim settingNode As XmlNode
        Dim xpathSetting As String = "root/" + branch + "/" + key
        Dim settingValue As String = ""

        If Me.settingsFilePath = "" Then
            Return ret_notInitialized
        End If

        Try
            xmlDocument.Load(Me.settingsFilePath)
            settingNode = xmlDocument.SelectSingleNode(xpathSetting)

            If Not settingNode Is Nothing Then
                settingValue = settingNode.InnerText
                Return settingValue
            Else
                Return ret_notFound
            End If
        Catch ex As Exception
            Me.lastError = ex.ToString()
            Return ret_fileError
        End Try

    End Function

    Public Function editSetting(ByVal key As String, ByVal value As String, Optional ByVal branch As String = "general_settings") As String

        Dim xmlDocument As XmlDocument = New XmlDocument()
        Dim settingNode As XmlNode
        Dim xpathSetting As String = "root/" + branch + "/" + key
        Dim settingValue As String = ""

        If Me.settingsFilePath = "" Then
            Return ret_notInitialized
        End If

        Try
            xmlDocument.Load(Me.settingsFilePath)
            settingNode = xmlDocument.SelectSingleNode(xpathSetting)

            If Not settingNode Is Nothing Then
                settingNode.InnerText = value
                xmlDocument.Save(Me.settingsFilePath)
                Return ret_success
            Else
                Return ret_notFound
            End If
        Catch ex As Exception
            Me.lastError = ex.ToString()
            Return ret_fileError
        End Try

    End Function

    Public Function deleteSetting(ByVal key As String, Optional ByVal branch As String = "general_settings") As String

        Dim xmlDocument As XmlDocument = New XmlDocument()
        Dim settingNode As XmlNode
        Dim xpathSetting As String = "root/" + branch + "/" + key
        Dim settingValue As String = ""

        If Me.settingsFilePath = "" Then
            Return ret_notInitialized
        End If

        Try
            xmlDocument.Load(Me.settingsFilePath)
            settingNode = xmlDocument.SelectSingleNode(xpathSetting)

            If Not settingNode Is Nothing Then
                xmlDocument.RemoveChild(settingNode)
                xmlDocument.Save(Me.settingsFilePath)
                Return ret_success
            Else
                Return ret_notFound
            End If
        Catch ex As Exception
            Me.lastError = ex.ToString()
            Return ret_fileError
        End Try

    End Function

    Public Function addBranch(ByVal name As String) As String

        Dim xmlDocument As XmlDocument = New XmlDocument()
        Dim rootNode As XmlNode
        Dim branchNode As XmlNode
        Dim xpathBranch As String = "root/" + name

        If Me.settingsFilePath = "" Then
            Return ret_notInitialized
        End If

        Try
            xmlDocument.Load(Me.settingsFilePath)

            If xmlDocument.SelectSingleNode(xpathBranch) Is Nothing Then
                branchNode = xmlDocument.CreateNode(XmlNodeType.Element, name, "")
                rootNode = xmlDocument.SelectSingleNode("root")
                rootNode.AppendChild(branchNode)
                xmlDocument.Save(Me.settingsFilePath)
                Return ret_success
            Else
                Return ret_alreadyExists
            End If
        Catch ex As Exception
            Me.lastError = ex.ToString()
            Return ret_fileError
        End Try

    End Function

    Public Function deleteBranch(ByVal name As String) As String

        Dim xmlDocument As XmlDocument = New XmlDocument()
        Dim branchNode As XmlNode
        Dim rootnode As XmlNode
        Dim xpathBranch As String = "root/" + name

        If Me.settingsFilePath = "" Then
            Return ret_notInitialized
        End If

        Try
            xmlDocument.Load(Me.settingsFilePath)
            rootnode = xmlDocument.SelectSingleNode("root")
            branchNode = xmlDocument.SelectSingleNode(xpathBranch)

            If Not branchNode Is Nothing Then
                rootnode.RemoveChild(branchNode)
                xmlDocument.Save(Me.settingsFilePath)
                Return ret_success
            Else
                Return ret_notFound
            End If
        Catch ex As Exception
            Me.lastError = ex.ToString()
            Return ret_fileError
        End Try

    End Function

    Private Function validateFilePath(ByVal filePath As String)

        If filePath Is Nothing Then
            Return False
        End If

        For Each badChar As Char In System.IO.Path.GetInvalidPathChars
            If InStr(filePath, badChar) > 0 Then
                Return False
            End If
        Next

        Return True

    End Function

End Class

