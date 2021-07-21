Imports System.Data.OleDb
Public Class Form1

    Dim con As New OleDbConnection
    Dim pro As String
    Dim connString As String
    Dim command As String

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'TODO: This line of code loads data into the 'ProjectdbDataSet.BIMObjective' table. You can move, or remove it, as needed.
        Me.BIMObjectiveTableAdapter.Fill(Me.ProjectdbDataSet.BIMObjective)
        'Refresh DataGridView after new input/update data
        Me.BIMObjectiveDataGridView.RefreshEdit()
    End Sub

    Private Sub BIMObjectiveBindingNavigatorSaveItem_Click(sender As Object, e As EventArgs)
        Me.Validate()
        Me.BIMObjectiveBindingSource.EndEdit()
        Me.TableAdapterManager.UpdateAll(Me.ProjectdbDataSet)

    End Sub

    Private Sub btn_Search_Click(sender As Object, e As EventArgs) Handles btn_Search.Click
        Me.BIMObjectiveTableAdapter.FillBy1(Me.ProjectdbDataSet.BIMObjective, txtSearch.Text)
    End Sub

    Private Sub BindingNavigatorAddNewItem_Click(sender As Object, e As EventArgs) Handles BindingNavigatorAddNewItem.Click
        Form2.Show()
    End Sub

    Private Sub BIMObjectiveDataGridView_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles BIMObjectiveDataGridView.CellContentClick
        Dim dgv As DataGridView

        dgv = BIMObjectiveDataGridView

        If e.ColumnIndex = 10 Then
            Form2.ID = CInt(dgv.CurrentRow.Cells(0).Value)
            Form2.OrganizationName = CStr(dgv.CurrentRow.Cells(1).Value)
            Form2.ProjectName = CStr(dgv.CurrentRow.Cells(2).Value)
            Form2.AssessorName = CStr(dgv.CurrentRow.Cells(3).Value)
            Form2.AssessorDesignation = CStr(dgv.CurrentRow.Cells(3).Value)
            Form2.scoreAchieved1 = CInt(dgv.CurrentRow.Cells(4).Value)
            Form2.scoreAchieved2 = CInt(dgv.CurrentRow.Cells(5).Value)
            Form2.scoreAchieved3 = CInt(dgv.CurrentRow.Cells(6).Value)
            Form2.scoreAchieved4 = CInt(dgv.CurrentRow.Cells(7).Value)
            'Open Form2.
            Form2.Show()
        End If

    End Sub

    Private Sub BIMObjectiveDataGridView_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles BIMObjectiveDataGridView.MouseDoubleClick
        If BIMObjectiveDataGridView.Rows.GetRowCount(DataGridViewElementStates.Selected) > 0 Then
            Form2.tb_ID.Text = BIMObjectiveDataGridView.CurrentRow.Cells(0).Value.ToString
            Form2.tb_NameOfOrganisation.Text = BIMObjectiveDataGridView.CurrentRow.Cells(1).Value.ToString
            Form2.tb_NameOfProject.Text = BIMObjectiveDataGridView.CurrentRow.Cells(2).Value.ToString
            Form2.tb_AssessorName.Text = BIMObjectiveDataGridView.CurrentRow.Cells(3).Value.ToString
            Form2.tb_AssessorDesignation.Text = BIMObjectiveDataGridView.CurrentRow.Cells(4).Value.ToString
            Form2.scoreAchieved1 = BIMObjectiveDataGridView.CurrentRow.Cells(5).Value.ToString
            Form2.scoreAchieved2 = BIMObjectiveDataGridView.CurrentRow.Cells(6).Value.ToString
            Form2.scoreAchieved3 = BIMObjectiveDataGridView.CurrentRow.Cells(7).Value.ToString
            Form2.scoreAchieved4 = BIMObjectiveDataGridView.CurrentRow.Cells(8).Value.ToString
            Form2.ShowDialog()
        End If
    End Sub

    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        Dim exitDialog As DialogResult
        exitDialog = MessageBox.Show("Do you want to exit?", "Exit", MessageBoxButtons.YesNo)

        If exitDialog = DialogResult.No Then
            e.Cancel = True
        Else
            Application.ExitThread()
        End If
    End Sub

End Class
