Imports System.Data.OleDb
Public Class Form2

    Public ID As Integer
    Public OrganizationName As String
    Public ProjectName As String
    Public AssessorName As String
    Public AssessorDesignation As String
    Public scoreAchieved1 As Integer
    Public scoreAchieved2 As Integer
    Public scoreAchieved3 As Integer
    Public scoreAchieved4 As Integer
    Dim valSectionScore1 As Double
    Dim valSectionScore2 As Double
    Dim valSectionScore3 As Double
    Dim valSectionScore4 As Double
    Dim currentRecord As Integer = 0
    Dim con As New OleDbConnection
    Dim pro As String
    Dim connString As String
    Dim command As String
    Dim cmd As OleDbCommand


    'To handle the error = The Rows cannot be programmatically added to the DataGridView's rows collection 
    'when the control is data-bound.
    Public newRow As DataRow

    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'TODO: This line of code loads data into the 'ProjectdbDataSet.BIMObjective' table. You can move, or remove it, as needed.
        Form1.BIMObjectiveTableAdapter.Fill(Form1.ProjectdbDataSet.BIMObjective)

        Dim parseID = Val(tb_ID.Text)
        parseID = ID
        OrganizationName = tb_NameOfOrganisation.Text
        ProjectName = tb_NameOfProject.Text
        AssessorName = tb_AssessorName.Text
        AssessorDesignation = tb_AssessorDesignation.Text

        If scoreAchieved1 = 0 Then
            Q1_0.Checked = True
        ElseIf scoreAchieved1 = 1 Then
            Q1_1.Checked = True
        ElseIf scoreAchieved1 = 2 Then
            Q1_2.Checked = True
        ElseIf scoreAchieved1 = 3 Then
            Q1_3.Checked = True
        ElseIf scoreAchieved1 = 4 Then
            Q1_4.Checked = True
        ElseIf scoreAchieved1 = 5 Then
            Q1_5.Checked = True
        End If


        If scoreAchieved2 = 0 Then
            Q2_0.Checked = True
        ElseIf scoreAchieved2 = 1 Then
            Q2_1.Checked = True
        ElseIf scoreAchieved2 = 2 Then
            Q2_2.Checked = True
        ElseIf scoreAchieved2 = 3 Then
            Q2_3.Checked = True
        ElseIf scoreAchieved2 = 4 Then
            Q2_4.Checked = True
        ElseIf scoreAchieved2 = 5 Then
            Q2_5.Checked = True
        End If


        If scoreAchieved3 = 0 Then
            Q3_0.Checked = True
        ElseIf scoreAchieved3 = 1 Then
            Q3_1.Checked = True
        ElseIf scoreAchieved3 = 2 Then
            Q3_2.Checked = True
        ElseIf scoreAchieved3 = 3 Then
            Q3_3.Checked = True
        ElseIf scoreAchieved3 = 4 Then
            Q3_4.Checked = True
        ElseIf scoreAchieved3 = 5 Then
            Q3_5.Checked = True
        End If


        If scoreAchieved4 = 0 Then
            Q4_0.Checked = True
        ElseIf scoreAchieved4 = 1 Then
            Q4_1.Checked = True
        ElseIf scoreAchieved4 = 2 Then
            Q4_2.Checked = True
        ElseIf scoreAchieved4 = 3 Then
            Q4_3.Checked = True
        ElseIf scoreAchieved4 = 4 Then
            Q4_4.Checked = True
        ElseIf scoreAchieved4 = 5 Then
            Q4_5.Checked = True
        End If
    End Sub

    Private Function calculatetotalSectionScore() As Double
        Dim totalSectionScore As Double = 0.0
        totalSectionScore = totalSectionScore + valSectionScore1 + valSectionScore2 + valSectionScore3 + valSectionScore4
        Return totalSectionScore
    End Function

    Private Function percentageScoreAchieved1(ByVal scoreAchieved1 As Integer) As Integer
        Dim totalPercentage As Integer
        totalPercentage = (scoreAchieved1 / 5) * 100
        Return totalPercentage
    End Function

    Private Function percentageScoreAchieved2(ByVal scoreAchieved2 As Integer) As Integer
        Dim totalPercentage As Integer
        totalPercentage = (scoreAchieved2 / 5) * 100
        Return totalPercentage
    End Function

    Private Function percentageScoreAchieved3(ByVal scoreAchieved3 As Integer) As Integer
        Dim totalPercentage As Integer
        totalPercentage = (scoreAchieved3 / 5) * 100
        Return totalPercentage
    End Function

    Private Function percentageScoreAchieved4(ByVal scoreAchieved4 As Integer) As Integer
        Dim totalPercentage As Integer
        totalPercentage = (scoreAchieved4 / 5) * 100
        Return totalPercentage
    End Function

    Private Function calculateSectionScore1(ByVal scoreAchieved1 As Integer) As Double
        Dim sectionScore1 As Double
        sectionScore1 = Math.Round(Val(scoreAchieved1 / 5 * 5.374), 3)
        Return sectionScore1
    End Function

    Private Function calculateSectionScore2(ByVal scoreAchieved2 As Integer) As Double
        Dim sectionScore2 As Double
        sectionScore2 = Math.Round(Val(scoreAchieved2 / 5 * 5.374), 3)
        Return sectionScore2
    End Function
    Private Function calculateSectionScore3(ByVal scoreAchieved3 As Integer) As Double
        Dim sectionScore3 As Double
        sectionScore3 = Math.Round(Val(scoreAchieved3 / 5 * 5.374), 3)
        Return sectionScore3
    End Function

    Private Function calculateSectionScore4(ByVal scoreAchieved4 As Integer) As Double
        Dim sectionScore4 As Double
        sectionScore4 = Math.Round(Val(scoreAchieved4 / 5 * 5.374), 3)
        Return sectionScore4
    End Function

    Private Sub nextRecord(AddVal As Integer)
        currentRecord += AddVal
        If currentRecord > Form1.BIMObjectiveDataGridView.Rows.Count - 1 Then currentRecord = 0 'Loop to first record
        If currentRecord < 0 Then currentRecord = Form1.BIMObjectiveDataGridView.Rows.Count - 1 'Loop to last record


    End Sub

    'Sub procedure when user wants to update the data.
    Public Sub bt_Update_Click(sender As Object, e As EventArgs) Handles bt_Update.Click
        Form1.Validate()
        Form1.BIMObjectiveBindingSource.EndEdit()

        'Update input into MS Access root folder (not in the Debug folder)
        con.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source = C:\Users\user\Desktop\CSC301\CSC301 GROUP PROJECT\BIM Objective\projectdb.accdb"
        con.Open()
        command = "UPDATE BimObjective SET [Name of Organization] = '" & tb_NameOfOrganisation.Text & "', [Name of Project] = '" & tb_NameOfProject.Text & "', [Assessor Name] = '" & tb_AssessorName.Text & "',
                                             [Assessor Designation] = '" & tb_AssessorDesignation.Text & "', [Score Achieved 1] = '" & scoreAchieved1 & "', [Score Achieved 2] = '" & scoreAchieved2 & "',
                                             [Score Achieved 3] = '" & scoreAchieved3 & "', [Score Achieved 4] = '" & scoreAchieved1 & "' WHERE [ID] = " & tb_ID.Text & ""
        Dim cmd As OleDbCommand = New OleDbCommand(command, con)

        Try
            cmd.ExecuteNonQuery()
            cmd.Dispose()
            con.Close()
            tb_ID.Clear()
            tb_NameOfOrganisation.Clear()
            tb_NameOfProject.Clear()
            tb_AssessorName.Clear()
            tb_AssessorDesignation.Clear()
            MessageBox.Show("Data updated.")
            Me.Close()
            Form1.Show()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Me.Hide()
            Form1.Show()
        End Try
    End Sub

    'Sub procedure when user wants to add new data.
    Private Sub bt_Add_Click(sender As Object, e As EventArgs) Handles bt_Add.Click
        'Replacement of Form1.BIMObjectiveTableAdapter.Rows.Add() method as the DataGridView is a DataBound type.
        newRow = Form1.ProjectdbDataSet.Tables(0).NewRow

        'Adding input row by row into DataGridView
        newRow.Item(0) = ID
        newRow.Item(1) = OrganizationName
        newRow.Item(2) = ProjectName
        newRow.Item(3) = AssessorName
        newRow.Item(4) = AssessorDesignation
        newRow.Item(5) = scoreAchieved1
        newRow.Item(6) = scoreAchieved2
        newRow.Item(7) = scoreAchieved3
        newRow.Item(8) = scoreAchieved4
        Form1.BIMObjectiveTableAdapter.Insert(OrganizationName, ProjectName, AssessorName, AssessorDesignation,
                                              scoreAchieved1, scoreAchieved2, scoreAchieved3, scoreAchieved4)
        Form1.BIMObjectiveTableAdapter.Update(Form1.ProjectdbDataSet.Tables(0))

        'Insert input into MS Access root folder (not in the Debug folder)
        con.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source = C:\Users\user\Desktop\CSC301\CSC301 GROUP PROJECT\BIM Objective\projectdb.accdb"
        con.Open()
        command = "INSERT INTO BimObjective ([ID], [Name of Organization], [Name of Project], [Assessor Name],
                                             [Assessor Designation], [Score Achieved 1], [Score Achieved 2],
                                             [Score Achieved 3], [Score Achieved 4]) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)"
        cmd = New OleDbCommand(command, con)
        cmd.Parameters.Add(New OleDbParameter("ID", CType(tb_ID.Text, String)))
        cmd.Parameters.Add(New OleDbParameter("Name of Organization", CType(tb_NameOfOrganisation.Text, String)))
        cmd.Parameters.Add(New OleDbParameter("Name of Project", CType(tb_NameOfProject.Text, String)))
        cmd.Parameters.Add(New OleDbParameter("Assessor Name", CType(tb_AssessorName.Text, String)))
        cmd.Parameters.Add(New OleDbParameter("Assessor Designation", CType(tb_AssessorDesignation.Text, String)))
        cmd.Parameters.Add(New OleDbParameter("Score Achieved 1", CType(scoreAchieved1, Integer)))
        cmd.Parameters.Add(New OleDbParameter("Score Achieved 2", CType(scoreAchieved2, Integer)))
        cmd.Parameters.Add(New OleDbParameter("Score Achieved 3", CType(scoreAchieved3, Integer)))
        cmd.Parameters.Add(New OleDbParameter("Score Achieved 4", CType(scoreAchieved4, Integer)))

        Try
            cmd.ExecuteNonQuery()
            cmd.Dispose()
            con.Close()
            tb_ID.Clear()
            tb_NameOfOrganisation.Clear()
            tb_NameOfProject.Clear()
            tb_AssessorName.Clear()
            tb_AssessorDesignation.Clear()
            MessageBox.Show("Data saved.")
            Me.Close()
            Form1.Show()
            Form1.BIMObjectiveTableAdapter.Fill(Form1.ProjectdbDataSet.BIMObjective)
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Form1.Show()
        End Try

    End Sub

    'If a radiobutton (0, 1, 2, 3, 4, 5) is checked for each question, the sub will process the coding inside it.
    '-------------------------------------------------------------------------------------------------------------
    'If radiobutton (0) for Question 1 is checked.
    Private Sub Q1_0_CheckedChanged(sender As Object, e As EventArgs) Handles Q1_0.CheckedChanged
        scoreAchieved1 = 0
        Q1_ScoreAchieved.Text = scoreAchieved1
        Q1_PercentageScoreAchieved.Text = percentageScoreAchieved1(scoreAchieved1)
        Q1_SectionScore.Text = calculateSectionScore1(scoreAchieved1)

    End Sub

    'If radiobutton (1) for Question 1 is checked.
    Private Sub Q1_1_CheckedChanged(sender As Object, e As EventArgs) Handles Q1_1.CheckedChanged
        scoreAchieved1 = 1
        Q1_ScoreAchieved.Text = scoreAchieved1
        Q1_PercentageScoreAchieved.Text = percentageScoreAchieved1(scoreAchieved1)
        Q1_SectionScore.Text = calculateSectionScore1(scoreAchieved1)
        valSectionScore1 = calculateSectionScore1(scoreAchieved1)
    End Sub

    'If radiobutton (2) for Question 1 is checked.
    Private Sub Q1_2_CheckedChanged(sender As Object, e As EventArgs) Handles Q1_2.CheckedChanged
        scoreAchieved1 = 2
        Q1_ScoreAchieved.Text = scoreAchieved1
        Q1_PercentageScoreAchieved.Text = percentageScoreAchieved1(scoreAchieved1)
        Q1_SectionScore.Text = calculateSectionScore1(scoreAchieved1)
        valSectionScore1 = calculateSectionScore1(scoreAchieved1)
    End Sub

    'If radiobutton (3) for Question 1 is checked.
    Private Sub Q1_3_CheckedChanged(sender As Object, e As EventArgs) Handles Q1_3.CheckedChanged
        scoreAchieved1 = 3
        Q1_ScoreAchieved.Text = scoreAchieved1
        Q1_PercentageScoreAchieved.Text = percentageScoreAchieved1(scoreAchieved1)
        Q1_SectionScore.Text = calculateSectionScore1(scoreAchieved1)
        valSectionScore1 = calculateSectionScore1(scoreAchieved1)
    End Sub

    'If radiobutton (4) for Question 1 is checked.
    Private Sub Q1_4_CheckedChanged(sender As Object, e As EventArgs) Handles Q1_4.CheckedChanged
        scoreAchieved1 = 4
        Q1_ScoreAchieved.Text = scoreAchieved1
        Q1_PercentageScoreAchieved.Text = percentageScoreAchieved1(scoreAchieved1)
        Q1_SectionScore.Text = calculateSectionScore1(scoreAchieved1)
        valSectionScore1 = calculateSectionScore1(scoreAchieved1)
    End Sub

    'If radiobutton (5) for Question 1 is checked.
    Private Sub Q1_5_CheckedChanged(sender As Object, e As EventArgs) Handles Q1_5.CheckedChanged
        scoreAchieved1 = 5
        Q1_ScoreAchieved.Text = scoreAchieved1
        Q1_PercentageScoreAchieved.Text = percentageScoreAchieved1(scoreAchieved1)
        Q1_SectionScore.Text = calculateSectionScore1(scoreAchieved1)
        valSectionScore1 = calculateSectionScore1(scoreAchieved1)
    End Sub

    'If radiobutton (0) for Question 2 is checked.
    Private Sub Q2_0_CheckedChanged(sender As Object, e As EventArgs) Handles Q2_0.CheckedChanged
        scoreAchieved2 = 0
        Q2_ScoreAchieved.Text = scoreAchieved2
        Q2_PercentageScoreAchieved.Text = percentageScoreAchieved2(scoreAchieved2)
        Q2_SectionScore.Text = calculateSectionScore2(scoreAchieved2)
        valSectionScore2 = calculateSectionScore2(scoreAchieved2)
    End Sub

    'If radiobutton (1) for Question 2 is checked.
    Private Sub Q2_1_CheckedChanged(sender As Object, e As EventArgs) Handles Q2_1.CheckedChanged
        scoreAchieved2 = 2
        Q2_ScoreAchieved.Text = scoreAchieved2
        Q2_PercentageScoreAchieved.Text = percentageScoreAchieved2(scoreAchieved2)
        Q2_SectionScore.Text = calculateSectionScore2(scoreAchieved2)
        valSectionScore2 = calculateSectionScore2(scoreAchieved2)
    End Sub

    'If radiobutton (2) for Question 2 is checked.
    Private Sub Q2_2_CheckedChanged(sender As Object, e As EventArgs) Handles Q2_2.CheckedChanged
        scoreAchieved2 = 2
        Q2_ScoreAchieved.Text = scoreAchieved2
        Q2_PercentageScoreAchieved.Text = percentageScoreAchieved2(scoreAchieved2)
        Q2_SectionScore.Text = calculateSectionScore2(scoreAchieved2)
        valSectionScore2 = calculateSectionScore2(scoreAchieved2)
    End Sub

    'If radiobutton (3) for Question 2 is checked.
    Private Sub Q2_3_CheckedChanged(sender As Object, e As EventArgs) Handles Q2_3.CheckedChanged
        scoreAchieved2 = 3
        Q2_ScoreAchieved.Text = scoreAchieved2
        Q2_PercentageScoreAchieved.Text = percentageScoreAchieved2(scoreAchieved2)
        Q2_SectionScore.Text = calculateSectionScore2(scoreAchieved2)
        valSectionScore2 = calculateSectionScore2(scoreAchieved2)
    End Sub

    'If radiobutton (4) for Question 2 is checked.
    Private Sub Q2_4_CheckedChanged(sender As Object, e As EventArgs) Handles Q2_4.CheckedChanged
        scoreAchieved2 = 4
        Q2_ScoreAchieved.Text = scoreAchieved2
        Q2_PercentageScoreAchieved.Text = percentageScoreAchieved2(scoreAchieved2)
        Q2_SectionScore.Text = calculateSectionScore2(scoreAchieved2)
        valSectionScore2 = calculateSectionScore2(scoreAchieved2)
    End Sub

    'If radiobutton (5) for Question 2 is checked.
    Private Sub Q2_5_CheckedChanged(sender As Object, e As EventArgs) Handles Q2_5.CheckedChanged
        scoreAchieved2 = 5
        Q2_ScoreAchieved.Text = scoreAchieved2
        Q2_PercentageScoreAchieved.Text = percentageScoreAchieved2(scoreAchieved2)
        Q2_SectionScore.Text = calculateSectionScore2(scoreAchieved2)
        valSectionScore2 = calculateSectionScore2(scoreAchieved2)
    End Sub

    'If radiobutton (0) for Question 3 is checked.
    Private Sub Q3_0_CheckedChanged(sender As Object, e As EventArgs) Handles Q3_0.CheckedChanged
        scoreAchieved3 = 0
        Q3_ScoreAchieved.Text = scoreAchieved3
        Q3_PercentageScoreAchieved.Text = percentageScoreAchieved3(scoreAchieved3)
        Q3_SectionScore.Text = calculateSectionScore3(scoreAchieved3)
        valSectionScore3 = calculateSectionScore3(scoreAchieved3)
    End Sub

    'If radiobutton (1) for Question 3 is checked.
    Private Sub Q3_1_CheckedChanged(sender As Object, e As EventArgs) Handles Q3_1.CheckedChanged
        scoreAchieved3 = 1
        Q3_ScoreAchieved.Text = scoreAchieved3
        Q3_PercentageScoreAchieved.Text = percentageScoreAchieved3(scoreAchieved3)
        Q3_SectionScore.Text = calculateSectionScore3(scoreAchieved3)
        valSectionScore3 = calculateSectionScore3(scoreAchieved3)
    End Sub

    'If radiobutton (2) for Question 3 is checked.
    Private Sub Q3_2_CheckedChanged(sender As Object, e As EventArgs) Handles Q3_2.CheckedChanged
        scoreAchieved3 = 2
        Q3_ScoreAchieved.Text = scoreAchieved3
        Q3_PercentageScoreAchieved.Text = percentageScoreAchieved3(scoreAchieved3)
        Q3_SectionScore.Text = calculateSectionScore3(scoreAchieved3)
        valSectionScore3 = calculateSectionScore3(scoreAchieved3)
    End Sub

    'If radiobutton (3) for Question 3 is checked.
    Private Sub Q3_3_CheckedChanged(sender As Object, e As EventArgs) Handles Q3_3.CheckedChanged
        scoreAchieved3 = 3
        Q3_ScoreAchieved.Text = scoreAchieved3
        Q3_PercentageScoreAchieved.Text = percentageScoreAchieved3(scoreAchieved3)
        Q3_SectionScore.Text = calculateSectionScore3(scoreAchieved3)
        valSectionScore3 = calculateSectionScore3(scoreAchieved3)
    End Sub

    'If radiobutton (4) for Question 3 is checked.
    Private Sub Q3_4_CheckedChanged(sender As Object, e As EventArgs) Handles Q3_4.CheckedChanged
        scoreAchieved3 = 4
        Q3_ScoreAchieved.Text = scoreAchieved3
        Q3_PercentageScoreAchieved.Text = percentageScoreAchieved3(scoreAchieved3)
        Q3_SectionScore.Text = calculateSectionScore3(scoreAchieved3)
        valSectionScore3 = calculateSectionScore3(scoreAchieved3)
    End Sub

    'If radiobutton (5) for Question 3 is checked.
    Private Sub Q3_5_CheckedChanged(sender As Object, e As EventArgs) Handles Q3_5.CheckedChanged
        scoreAchieved3 = 5
        Q3_ScoreAchieved.Text = scoreAchieved3
        Q3_PercentageScoreAchieved.Text = percentageScoreAchieved3(scoreAchieved3)
        Q3_SectionScore.Text = calculateSectionScore3(scoreAchieved3)
        valSectionScore3 = calculateSectionScore3(scoreAchieved3)
    End Sub

    'If radiobutton (0) for Question 4 is checked.
    Private Sub Q4_0_CheckedChanged(sender As Object, e As EventArgs) Handles Q4_0.CheckedChanged
        scoreAchieved4 = 0
        Q4_ScoreAchieved.Text = scoreAchieved4
        Q4_PercentageScoreAchieved.Text = percentageScoreAchieved4(scoreAchieved4)
        Q4_SectionScore.Text = calculateSectionScore4(scoreAchieved4)
        valSectionScore4 = calculateSectionScore4(scoreAchieved4)
    End Sub

    'If radiobutton (1) for Question 4 is checked.
    Private Sub Q4_1_CheckedChanged(sender As Object, e As EventArgs) Handles Q4_1.CheckedChanged
        scoreAchieved4 = 1
        Q4_ScoreAchieved.Text = scoreAchieved4
        Q4_PercentageScoreAchieved.Text = percentageScoreAchieved4(scoreAchieved4)
        Q4_SectionScore.Text = calculateSectionScore4(scoreAchieved4)
        valSectionScore4 = calculateSectionScore4(scoreAchieved4)
    End Sub

    'If radiobutton (2) for Question 4 is checked.
    Private Sub Q4_2_CheckedChanged(sender As Object, e As EventArgs) Handles Q4_2.CheckedChanged
        scoreAchieved4 = 2
        Q4_ScoreAchieved.Text = scoreAchieved4
        Q4_PercentageScoreAchieved.Text = percentageScoreAchieved4(scoreAchieved4)
        Q4_SectionScore.Text = calculateSectionScore4(scoreAchieved4)
        valSectionScore4 = calculateSectionScore4(scoreAchieved4)
    End Sub

    'If radiobutton (3) for Question 4 is checked.
    Private Sub Q4_3_CheckedChanged(sender As Object, e As EventArgs) Handles Q4_3.CheckedChanged
        scoreAchieved4 = 3
        Q4_ScoreAchieved.Text = scoreAchieved4
        Q4_PercentageScoreAchieved.Text = percentageScoreAchieved4(scoreAchieved4)
        Q4_SectionScore.Text = calculateSectionScore4(scoreAchieved4)
        valSectionScore4 = calculateSectionScore4(scoreAchieved4)
    End Sub

    'If radiobutton (4) for Question 4 is checked.
    Private Sub Q4_4_CheckedChanged(sender As Object, e As EventArgs) Handles Q4_4.CheckedChanged
        scoreAchieved4 = 4
        Q4_ScoreAchieved.Text = scoreAchieved4
        Q4_PercentageScoreAchieved.Text = percentageScoreAchieved4(scoreAchieved4)
        Q4_SectionScore.Text = calculateSectionScore4(scoreAchieved4)
        valSectionScore4 = calculateSectionScore4(scoreAchieved4)
    End Sub

    'If radiobutton (5) for Question 4 is checked.
    Private Sub Q4_5_CheckedChanged(sender As Object, e As EventArgs) Handles Q4_5.CheckedChanged
        scoreAchieved4 = 5
        Q4_ScoreAchieved.Text = scoreAchieved4
        Q4_PercentageScoreAchieved.Text = percentageScoreAchieved4(scoreAchieved4)
        Q4_SectionScore.Text = calculateSectionScore4(scoreAchieved4)
        valSectionScore4 = calculateSectionScore4(scoreAchieved4)
    End Sub

    Private Sub Form2_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        Form1.BIMObjectiveTableAdapter.Fill(Form1.ProjectdbDataSet.BIMObjective)
        Form1.Show()
    End Sub

    Private Sub bt_TotalSectionScore_Click(sender As Object, e As EventArgs) Handles bt_TotalSectionScore.Click
        tb_TotalSectionScore.Text = calculatetotalSectionScore()
    End Sub

    Private Sub bt_Cancel_Click(sender As Object, e As EventArgs) Handles bt_Cancel.Click
        Dim cancelMessage = MessageBox.Show("Do you want to cancel?", "Cancel", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

        If cancelMessage = Windows.Forms.DialogResult.Yes Then
            Form1.Show()
        Else
            Exit Sub
        End If
    End Sub

    Private Sub bt_Print_Click(sender As Object, e As EventArgs) Handles bt_Print.Click
        Form3.Show()
    End Sub

    Private Sub btn_Delete_Click(sender As Object, e As EventArgs) Handles btn_Delete.Click
        'Delete data from MS Access root folder (not in the Debug folder)
        con.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source = C:\Users\user\Desktop\CSC301\CSC301 GROUP PROJECT\BIM Objective\projectdb.accdb"
        con.Open()
        command = "DELETE FROM BimObjective WHERE [ID] = " & tb_ID.Text & ""
        Dim cmd As OleDbCommand = New OleDbCommand(command, con)

        Try
            cmd.ExecuteNonQuery()
            cmd.Dispose()
            con.Close()
            MessageBox.Show("Data deleted.")
            Me.Close()
            Form1.Show()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Me.Hide()
            Form1.Show()
        End Try

    End Sub
End Class
