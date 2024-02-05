Imports System.IO
Imports IniParser
Imports IniParser.Model

Public Class MainWindow
    Dim ReadOnly parser = new FileIniDataParser()

    Public Sub New()
        InitializeComponent()
        AddHandler Me.Loaded, AddressOf Window_Loaded
        Dim data As IniData
        Try
            Dim rootPath As String =
                    New DirectoryInfo(AppDomain.CurrentDomain.BaseDirectory).Parent.Parent.Parent.FullName
            Dim filePath As String = Path.Combine(rootPath, "startapplications.ini")
            data = parser.ReadFile(filePath)

            ' If your ini file has a section named "Buttons"
            Dim buttonSection = data("Buttons")

            ' Assigns the content from your ini file to your buttons
            Dim lastPartButton1 As String = Path.GetFileName(buttonSection("Button1"))
            Dim lastPartButton2 As String = Path.GetFileName(buttonSection("Button2"))
            Dim lastPartButton3 As String = Path.GetFileName(buttonSection("Button3"))
            Dim lastPartButton4 As String = Path.GetFileName(buttonSection("Button4"))
            Button1.Content = lastPartButton1
            Button2.Content = lastPartButton2
            Button3.Content = lastPartButton3
            Button4.Content = lastPartButton4
            ' ... repeat for every Button

        Catch ex As FileNotFoundException
            ' Handle the case when the ini file could not be found
            MessageBox.Show(ex.Message)
        Catch ex As Exception
            ' Handle the case when the ini file could not be parsed correctly
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Sub ChangeButtonContent(myButton as string, buttonContent as String)

        Dim parser = New FileIniDataParser()

        Try
            ' Read file
            Dim rootPath As String =
                    New DirectoryInfo(AppDomain.CurrentDomain.BaseDirectory).Parent.Parent.Parent.FullName
            Dim filePath As String = Path.Combine(rootPath, "startapplications.ini")

            If File.Exists(filePath) Then
                Dim data = parser.ReadFile(filePath)

                ' Modify content for Button
                data("Buttons")(myButton) = buttonContent

                ' Save file
                parser.WriteFile(filePath, data)
            Else
                MessageBox.Show($"File: {filePath} does not exist.")
            End If

        Catch ex As FileNotFoundException
            MessageBox.Show(ex.Message)
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub Window_Loaded(sender As Object, e As RoutedEventArgs)
        ' Position the window at the bottom of the working area (just above the taskbar )
        Me.Top = SystemParameters.WorkArea.Height - Me.Height
        Me.Left = (SystemParameters.WorkArea.Width - Me.Width)/2
    End Sub

    Private Sub Window_KeyDown(sender As Object, e As KeyEventArgs)
        Select Case e.Key
            Case Key.Left
                Me.Left -= 200 'Moves the Window 10 units to the left.
            Case Key.Right
                Me.Left += 200 'Moves the Window 10 units to the right.
            Case Key.Up
                Me.Top -= 200 'Moves the Window 10 units up.
            Case Key.Down
                Me.Top += 200 'Moves the Window 10 units down.
        End Select
    End Sub

    Private Sub BtnDropFile_Drop(sender As Object, e As DragEventArgs)
        Dim button = CType(sender, Button)
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            ' Note that you can have more than one file.
            Dim files = DirectCast(e.Data.GetData(DataFormats.FileDrop), String())
            If files IsNot Nothing AndAlso files.Length <> 0 Then
                ' For this example, suppose we're only interested in the first file.
                Dim lastPart As String = Path.GetFileName(files(0))
                button.Content = lastPart
                ChangeButtonContent(button.Name, files(0))
                ' MessageBox.Show("You dropped: " & files(0))
            End If
        End If
    End Sub

    Private Sub ButtonClickHandler(sender As Object, e As RoutedEventArgs)
        Dim button = CType(sender, Button)
        Dim saveContent = button.Content
        if (InStr("Start Word", saveContent)) Then
            Dim myApplicationHandler as New ApplicationHandler
            myApplicationHandler.startWord()
        Else if (InStr("Start Excel", saveContent)) Then
            Dim myApplicationHandler as New ApplicationHandler
            myApplicationHandler.startExcel()
        Else if (InStr("Start OneNote", saveContent)) Then
            Dim myApplicationHandler as New ApplicationHandler
            myApplicationHandler.StartOneNote()
        Else if (InStr("Start Outlook", saveContent)) Then
            Dim myApplicationHandler as New ApplicationHandler
            myApplicationHandler.StartOutlook()
        Else
            Dim myApplicationHandler as New ApplicationHandler
            Dim parser = New FileIniDataParser()
            Dim myApplication as string

            Try
                ' Read file
                Dim rootPath As String =
                        New DirectoryInfo(AppDomain.CurrentDomain.BaseDirectory).Parent.Parent.Parent.FullName
                Dim filePath As String = Path.Combine(rootPath, "startapplications.ini")

                If File.Exists(filePath) Then
                    Dim data = parser.ReadFile(filePath)

                    ' Modify content for Button
                    myApplication = data("Buttons")(button.Name)
                Else
                    MessageBox.Show($"File: {filePath} does not exist.")
                End If

            Catch ex As FileNotFoundException
                MessageBox.Show(ex.Message)
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try

            myApplicationHandler.StartApplication(myApplication)
        End If
    End Sub

    Public Sub Button_Click(sender As Object, e As RoutedEventArgs)
        ' Your event handling logic goes here.
    End Sub
End Class
