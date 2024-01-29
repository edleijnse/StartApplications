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
