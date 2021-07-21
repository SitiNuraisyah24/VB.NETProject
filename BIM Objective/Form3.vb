Public Class Form3
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles Me.Load
        Dim scoreArchived As Integer
        Dim scoreAvailable As Integer
        Dim percentScoreArchived As Integer
        Dim sectionWeight As Double
        Dim sectionScore As Double


        Label40.Text = Label36.Text
    End Sub

    Private Sub btnBack_Click(sender As Object, e As EventArgs) Handles btnBack.Click
        Me.Hide()
        Form2.ShowDialog()
    End Sub

    Private Sub btnPrint_Click(sender As Object, e As EventArgs) Handles btnPrint.Click
    End Sub
End Class
