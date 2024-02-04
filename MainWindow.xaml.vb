Public Class MainWindow
    Public Sub New()
        InitializeComponent()
        AddHandler Me.Loaded, AddressOf Window_Loaded
    End Sub
    Private Sub Window_Loaded(sender As Object, e As RoutedEventArgs)
        ' Position the window at the bottom of the working area (just above the taskbar )
        Me.Top = SystemParameters.WorkArea.Height - Me.Height
        Me.Left = (SystemParameters.WorkArea.Width - Me.Width) / 2
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
        Dim button As Button = CType(sender, Button)
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            ' Note that you can have more than one file.
            Dim files() As String = DirectCast(e.Data.GetData(DataFormats.FileDrop), String())
            If files IsNot Nothing AndAlso files.Length <> 0 Then
                ' For this example, suppose we're only interested in the first file.
                chosenText.Text=files(0)
                button.Content=files(0)
                ' MessageBox.Show("You dropped: " & files(0))
            End If
        End If
    End Sub

    Private Sub ButtonClickHandler(sender As Object, e As RoutedEventArgs)
        Dim button As Button = CType(sender, Button)
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
            button.Content = "Clicked: " + saveContent
        End If
    End Sub

    Public Sub Button_Click(sender As Object, e As RoutedEventArgs)
        ' Your event handling logic goes here.
    End Sub
End Class
