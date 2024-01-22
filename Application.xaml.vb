Class Application
    Public Sub New()
        InitializeComponent()
    End Sub

    Private Sub InitializeComponent()
        ' Throw New NotImplementedException
    End Sub

    Private Sub ButtonClickHandler(sender As Object, e As RoutedEventArgs)
        Dim button As Button = CType(sender, Button)
        button.Content = "Clicked!"
    End Sub
End Class
