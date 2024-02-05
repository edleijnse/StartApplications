Imports System.IO
Imports IniParser
Imports IniParser.Model

Public Class MainWindow
    Private ReadOnly parser As New FileIniDataParser()

    Public Sub New()
        InitializeComponent()
        AddHandler Me.Loaded, AddressOf Window_Loaded
        LoadButtonContentsFromIniFile()
    End Sub

    Private Sub LoadButtonContentsFromIniFile()
        Dim filePath As String = GetIniFilePath()
        Try
            If File.Exists(filePath) Then
                Dim data = parser.ReadFile(filePath)
                Dim buttonSection = data("Buttons")
                AssignContentToButtons(buttonSection)
            Else
                MessageBox.Show($"File: {filePath} does not exist.")
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub AssignContentToButtons(buttonSection As KeyDataCollection)
        Button1.Content = Path.GetFileName(buttonSection("Button1"))
        Button2.Content = Path.GetFileName(buttonSection("Button2"))
        Button3.Content = Path.GetFileName(buttonSection("Button3"))
        Button4.Content = Path.GetFileName(buttonSection("Button4"))
    End Sub

    Private Function GetIniFilePath() As String
        Dim rootPath As String = New DirectoryInfo(AppDomain.CurrentDomain.BaseDirectory).Parent.Parent.Parent.FullName
        Return Path.Combine(rootPath, "startapplications.ini")
    End Function

    Public Sub ChangeButtonContent(myButton As String, buttonContent As String)
        Dim filePath As String = GetIniFilePath()
        Try
            If File.Exists(filePath) Then
                Dim data = parser.ReadFile(filePath)
                data("Buttons")(myButton) = buttonContent
                parser.WriteFile(filePath, data)
            Else
                MessageBox.Show($"File: {filePath} does not exist.")
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub Window_Loaded(sender As Object, e As RoutedEventArgs)
        Me.Top = SystemParameters.WorkArea.Height - Me.Height
        Me.Left = (SystemParameters.WorkArea.Width - Me.Width) / 2
    End Sub

    Private Sub Window_KeyDown(sender As Object, e As KeyEventArgs)
        Select Case e.Key
            Case Key.Left
                Me.Left -= 200
            Case Key.Right
                Me.Left += 200
            Case Key.Up
                Me.Top -= 200
            Case Key.Down
                Me.Top += 200
        End Select
    End Sub

    Private Sub BtnDropFile_Drop(sender As Object, e As DragEventArgs)
        Dim button = CType(sender, Button)
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            Dim files = DirectCast(e.Data.GetData(DataFormats.FileDrop), String())
            If files IsNot Nothing AndAlso files.Length <> 0 Then
                Dim lastPart As String = Path.GetFileName(files(0))
                button.Content = lastPart
                ChangeButtonContent(button.Name, files(0))
            End If
        End If
    End Sub

    Private Sub ButtonClickHandler(sender As Object, e As RoutedEventArgs)
        Dim button = CType(sender, Button)
        Dim myApplicationHandler As New ApplicationHandler()
        Select Case button.Content
            Case "Start Word"
                myApplicationHandler.startWord()
            Case "Start Excel"
                myApplicationHandler.startExcel()
            Case "Start OneNote"
                myApplicationHandler.StartOneNote()
            Case "Start Outlook"
                myApplicationHandler.StartOutlook()
            Case Else
                StartCustomApplication(button.Name, myApplicationHandler)
        End Select
    End Sub

    Private Sub StartCustomApplication(buttonName As String, applicationHandler As ApplicationHandler)
        Dim filePath As String = GetIniFilePath()
        Try
            If File.Exists(filePath) Then
                Dim data = parser.ReadFile(filePath)
                Dim myApplication = data("Buttons")(buttonName)
                applicationHandler.StartApplication(myApplication)
            Else
                MessageBox.Show($"File: {filePath} does not exist.")
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Public Sub Button_Click(sender As Object, e As RoutedEventArgs)
        ' Your event handling logic goes here.
    End Sub
End Class