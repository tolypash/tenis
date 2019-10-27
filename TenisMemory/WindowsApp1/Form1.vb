Imports VB = Microsoft.VisualBasic
Public Class Tenis
    Dim C_Th(8) As Boolean
    Dim BinaryV(8) As Boolean
    Dim SelectedBit As Integer
    Dim VenusTemp As Integer = 733 'Venus temperature
    Dim Tintersection As Integer = 760
    Dim temp As Integer 'temporary value 2

    Dim Tinc1B As Boolean = False 'checks if it has reached new equilibrium so head temperature does not change
    Private Sub Tinc1_Scroll(sender As Object, e As EventArgs) Handles Tinc1.Scroll
        ClearChart() 'clearing chart

        Th1.Text = VenusTemp + Tinc1.Value & "K"

        If Tinc1B = False Then
            THeadT1.Text = Th1.Text
        End If

        temp = CInt(VB.Left(THeadT1.Text, 3))
        If temp = Tintersection Then
            Tinc1B = True
            Bin1.Text = "Binary Value: 1"
            BinaryV(1) = True
        Else
            Bin1.Text = "Binary Value: 0"
        End If

        Image() 'checks each bit in the byte and turns the corresponding cell to bg color = black

        DrawChart() 'REDRAWING

        PictureBox1.Size = New Size(10, 73 + Tinc1.Value)
        TB1.Location = New Point(6, 139 + Tinc1.Value)
        Label2.Location = New Point(95, 145 + Tinc1.Value)

        CalculateDec()
        CalculateHex()

        CheckQ1()
    End Sub

    Dim Tinc2B As Boolean = False
    Private Sub Tinc2_Scroll(sender As Object, e As EventArgs) Handles Tinc2.Scroll
        ClearChart()

        Th2.Text = VenusTemp + Tinc2.Value & "K"

        If Tinc2B = False Then
            THeadT2.Text = Th2.Text
        End If

        temp = CInt(VB.Left(THeadT2.Text, 3))
        If temp = Tintersection Then
            Tinc2B = True
            Bin2.Text = "Binary Value: 1"
            BinaryV(2) = True
        Else
            Bin2.Text = "Binary Value: 0"
        End If

        Image()

        DrawChart()

        PictureBox2.Size = New Size(10, 73 + Tinc2.Value)
        TB2.Location = New Point(6, 139 + Tinc2.Value)
        Label3.Location = New Point(95, 145 + Tinc2.Value)

        CalculateDec()
        CalculateHex()

        CheckQ2()
    End Sub

    Dim Tinc3B As Boolean = False
    Private Sub Tinc3_Scroll(sender As Object, e As EventArgs) Handles Tinc3.Scroll
        ClearChart()

        Th3.Text = VenusTemp + Tinc3.Value & "K"

        If Tinc3B = False Then
            THeadT3.Text = Th3.Text
        End If

        temp = CInt(VB.Left(THeadT3.Text, 3))
        If temp = Tintersection Then
            Tinc3B = True
            Bin3.Text = "Binary Value: 1"
            BinaryV(3) = True
        Else
            Bin3.Text = "Binary Value: 0"
        End If

        Image()

        DrawChart()

        PictureBox3.Size = New Size(10, 73 + Tinc3.Value)
        TB3.Location = New Point(6, 139 + Tinc3.Value)
        Label4.Location = New Point(95, 145 + Tinc3.Value)

        CalculateDec()
        CalculateHex()

        CheckQ3()
    End Sub

    Dim Tinc4B As Boolean = False
    Private Sub Tinc4_Scroll(sender As Object, e As EventArgs) Handles Tinc4.Scroll
        ClearChart()

        Th4.Text = VenusTemp + Tinc4.Value & "K"

        If Tinc4B = False Then
            THeadT4.Text = Th4.Text
        End If

        temp = CInt(VB.Left(THeadT4.Text, 3))
        If temp = Tintersection Then
            Tinc4B = True
            Bin4.Text = "Binary Value: 1"
            BinaryV(4) = True
        Else
            Bin4.Text = "Binary Value: 0"
        End If

        Image()

        DrawChart()

        PictureBox4.Size = New Size(10, 73 + Tinc4.Value)
        TB4.Location = New Point(6, 139 + Tinc4.Value)
        Label5.Location = New Point(95, 145 + Tinc4.Value)

        CalculateDec()
        CalculateHex()

        CheckQ4()
    End Sub

    Dim Tinc5B As Boolean = False
    Private Sub Tinc5_Scroll(sender As Object, e As EventArgs) Handles Tinc5.Scroll
        ClearChart()

        Th5.Text = VenusTemp + Tinc5.Value & "K"

        If Tinc5B = False Then
            THeadT5.Text = Th5.Text
        End If

        temp = CInt(VB.Left(THeadT5.Text, 3))
        If temp = Tintersection Then
            Tinc5B = True
            Bin5.Text = "Binary Value: 1"
            BinaryV(5) = True
        Else
            Bin5.Text = "Binary Value: 0"
        End If

        Image()

        DrawChart()

        PictureBox5.Size = New Size(10, 73 + Tinc5.Value)
        TB5.Location = New Point(6, 139 + Tinc5.Value)
        Label6.Location = New Point(95, 145 + Tinc5.Value)

        CalculateDec()
        CalculateHex()

        CheckQ5()
    End Sub

    Dim Tinc6B As Boolean = False
    Private Sub Tinc6_Scroll(sender As Object, e As EventArgs) Handles Tinc6.Scroll
        ClearChart()

        Th6.Text = VenusTemp + Tinc6.Value & "K"

        If Tinc6B = False Then
            THeadT6.Text = Th6.Text
        End If

        temp = CInt(VB.Left(THeadT6.Text, 3))
        If temp = Tintersection Then
            Tinc6B = True
            Bin6.Text = "Binary Value: 1"
            BinaryV(6) = True
        Else
            Bin6.Text = "Binary Value: 0"
        End If

        Image()

        DrawChart()

        PictureBox6.Size = New Size(10, 73 + Tinc6.Value)
        TB6.Location = New Point(6, 139 + Tinc6.Value)
        Label7.Location = New Point(95, 145 + Tinc6.Value)

        CalculateDec()
        CalculateHex()

        CheckQ6()
    End Sub

    Dim Tinc7B As Boolean = False
    Private Sub Tinc7_Scroll(sender As Object, e As EventArgs) Handles Tinc7.Scroll
        ClearChart()

        If Tinc7B = False Then
            Th7.Text = VenusTemp + Tinc7.Value & "K"
        End If

        THeadT7.Text = Th7.Text

        temp = CInt(VB.Left(THeadT7.Text, 3))
        If temp = Tintersection Then
            Tinc7B = True
            Bin7.Text = "Binary Value: 1"
            BinaryV(7) = True
        Else
            Bin7.Text = "Binary Value: 0"
        End If

        Image()

        DrawChart()

        PictureBox7.Size = New Size(10, 73 + Tinc7.Value)
        TB7.Location = New Point(6, 139 + Tinc7.Value)
        Label8.Location = New Point(95, 145 + Tinc7.Value)

        CalculateDec()
        CalculateHex()

        CheckQ7()
    End Sub

    Dim Tinc8B As Boolean = False
    Private Sub Tinc8_Scroll(sender As Object, e As EventArgs) Handles Tinc8.Scroll
        ClearChart()

        Th8.Text = VenusTemp + Tinc8.Value & "K"

        If Tinc8B = False Then
            THeadT8.Text = Th8.Text
        End If

        temp = CInt(VB.Left(THeadT8.Text, 3))
        If temp = Tintersection Then
            Tinc8B = True
            Bin8.Text = "Binary Value: 1"
            BinaryV(8) = True
        Else
            Bin8.Text = "Binary Value: 0"
        End If

        Image()

        DrawChart()

        PictureBox8.Size = New Size(10, 73 + Tinc8.Value)
        TB8.Location = New Point(6, 139 + Tinc8.Value)
        Label9.Location = New Point(95, 145 + Tinc8.Value)

        CalculateDec()
        CalculateHex()

        CheckQ8()
    End Sub
    Sub Image() 'draws the 8 bit black and white image
        For c = 1 To 8
            If BinaryV(c) = True Then
                DataGridView1.Columns("Column" & c).DefaultCellStyle.BackColor = Color.Black
            End If
        Next

    End Sub

    Sub DrawChart() 'drawing chart and marking the equilibrium points of the corresponding bit
        Me.Chart1.ChartAreas("ChartArea1").AxisX.Minimum = VenusTemp - 33
        Me.Chart1.ChartAreas("ChartArea1").AxisX.Maximum = VenusTemp + 99
        Me.Chart1.ChartAreas("ChartArea1").AxisY.Minimum = 0
        Me.Chart1.ChartAreas("ChartArea1").AxisY.Maximum = 1

        Me.Chart1.Series("Q Cond Out").Points.AddXY(VenusTemp - 33, 0)
        Me.Chart1.Series("Q Cond Out").Points.AddXY(VenusTemp + 99, 0.48)

        Me.Chart1.Series("Q NF,in").Points.AddXY(VenusTemp - 33, 0.25) 'POINT 0
        Me.Chart1.Series("Q NF,in").Points.AddXY(VenusTemp, 0.12) 'POINT 1
        Me.Chart1.Series("Q NF,in").Points.AddXY(VenusTemp + 17, 0.1)
        Me.Chart1.Series("Q NF,in").Points.AddXY(VenusTemp + 47, 0.45)
        Me.Chart1.Series("Q NF,in").Points.AddXY(VenusTemp + 66, 0.36) 'POINT 4
        Me.Chart1.Series("Q NF,in").Points.AddXY(VenusTemp + 99, 0)

        If BinaryV(SelectedBit) = False Then
            Me.Chart1.Series("Q NF,in").Points(1).MarkerStyle = DataVisualization.Charting.MarkerStyle.Circle
            Me.Chart1.Series("Q NF,in").Points(1).MarkerSize = 10
            Me.Chart1.Series("Q NF,in").Points(1).MarkerColor = Color.Green
        Else
            Me.Chart1.Series("Q NF,in").Points(4).MarkerStyle = DataVisualization.Charting.MarkerStyle.None
            Me.Chart1.Series("Q NF,in").Points(4).MarkerStyle = DataVisualization.Charting.MarkerStyle.Circle
            Me.Chart1.Series("Q NF,in").Points(4).MarkerSize = 10
            Me.Chart1.Series("Q NF,in").Points(4).MarkerColor = Color.Green
        End If
    End Sub

    Dim DecimalV As Integer
    Sub CalculateDec() 'Calculates decimal value of Binary set BinaryV
        DecimalV = 0
        For c = 1 To 8
            If BinaryV(c) = True Then
                DecimalV = DecimalV + 2 ^ (c - 1)
            End If
        Next

        BinaryDecimal.Text = "Decimal Value: " & DecimalV
    End Sub

    Sub CalculateHex() 'Calculates hexadecimal value of Binary set BinaryV
        Dim Hex() As Char = {"0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "A", "B", "C", "D", "E", "F"}
        Dim p1, p2 As Integer 'part1,part2
        Dim f As String 'final part

        For i = 1 To 4
            If BinaryV(i) = True Then
                p1 = p1 + 2 ^ (i - 1)
            End If
        Next

        For j = 1 To 4
            If BinaryV(j + 4) = True Then
                p2 = p2 + 2 ^ (j - 1)
            End If
        Next

        f = Hex(p1) & Hex(p2)

        lblHex.Text = "Hexadecimal Value: " & f
    End Sub

    'CheckQ1-Q8 checks heat energy conducted outside and heat energy conducted inside of each cell and forms an inequality for user to visualise 
    Sub CheckQ1()
        Dim t As Integer
        t = CInt(VB.Left(THeadT1.Text, 3))
        If t = 733 Then
            lblQ.Text = "Q Cond Out = Q NF,in"
        ElseIf t > 733 And t < 760 Then
            lblQ.Text = "Q Cond Out > Q NF,in"
        ElseIf t >= 760 Then
            lblQ.Text = "Q Cond Out < Q NF,in"
        End If
    End Sub

    Sub CheckQ2()
        Dim t As Integer
        t = CInt(VB.Left(THeadT2.Text, 3))
        If t = 733 Then
            lblQ.Text = "Q Cond Out = Q NF,in"
        ElseIf t > 733 And t < 760 Then
            lblQ.Text = "Q Cond Out > Q NF,in"
        ElseIf t >= 760 Then
            lblQ.Text = "Q Cond Out < Q NF,in"
        End If
    End Sub

    Sub CheckQ3()
        Dim t As Integer
        t = CInt(VB.Left(THeadT3.Text, 3))
        If t = 733 Then
            lblQ.Text = "Q Cond Out = Q NF,in"
        ElseIf t > 733 And t < 760 Then
            lblQ.Text = "Q Cond Out > Q NF,in"
        ElseIf t >= 760 Then
            lblQ.Text = "Q Cond Out < Q NF,in"
        End If
    End Sub

    Sub CheckQ4()
        Dim t As Integer
        t = CInt(VB.Left(THeadT4.Text, 3))
        If t = 733 Then
            lblQ.Text = "Q Cond Out = Q NF,in"
        ElseIf t > 733 And t < 760 Then
            lblQ.Text = "Q Cond Out > Q NF,in"
        ElseIf t >= 760 Then
            lblQ.Text = "Q Cond Out < Q NF,in"
        End If
    End Sub

    Sub CheckQ5()
        Dim t As Integer
        t = CInt(VB.Left(THeadT5.Text, 3))
        If t = 733 Then
            lblQ.Text = "Q Cond Out = Q NF,in"
        ElseIf t > 733 And t < 760 Then
            lblQ.Text = "Q Cond Out > Q NF,in"
        ElseIf t >= 760 Then
            lblQ.Text = "Q Cond Out < Q NF,in"
        End If
    End Sub

    Sub CheckQ6()
        Dim t As Integer
        t = CInt(VB.Left(THeadT6.Text, 3))
        If t = 733 Then
            lblQ.Text = "Q Cond Out = Q NF,in"
        ElseIf t > 733 And t < 760 Then
            lblQ.Text = "Q Cond Out > Q NF,in"
        ElseIf t >= 760 Then
            lblQ.Text = "Q Cond Out < Q NF,in"
        End If
    End Sub

    Sub CheckQ7()
        Dim t As Integer
        t = CInt(VB.Left(THeadT7.Text, 3))
        If t = 733 Then
            lblQ.Text = "Q Cond Out = Q NF,in"
        ElseIf t > 733 And t < 760 Then
            lblQ.Text = "Q Cond Out > Q NF,in"
        ElseIf t >= 760 Then
            lblQ.Text = "Q Cond Out < Q NF,in"
        End If
    End Sub

    Sub CheckQ8()
        Dim t As Integer
        t = CInt(VB.Left(THeadT8.Text, 3))
        If t = 733 Then
            lblQ.Text = "Q Cond Out = Q NF,in"
        ElseIf t > 733 And t < 760 Then
            lblQ.Text = "Q Cond Out > Q NF,in"
        ElseIf t >= 760 Then
            lblQ.Text = "Q Cond Out < Q NF,in"
        End If
    End Sub

    Sub ClearChart() 'Clearing points to redraw in order to check if new stable state has been reached
        Me.Chart1.Series(0).Points.Clear()
        Me.Chart1.Series(1).Points.Clear()
    End Sub

    Private Sub HELP_initialize_Click(sender As Object, e As EventArgs) Handles HELP_initialize.Click
        MsgBox("This overwrites all bits written prior with 0's by using peltier effect to cool one side of a plate and use that to cool down the top of the system so that all values stored will be 0.")
    End Sub

    '<---INITIALIZING AND BEAUTIFYING OF GUI STARTS HERE--->
    Sub CheckButt() 'disables and enables appropriate objects for a better user experience
        If C_Th(1) = False Then
            C1Th.Enabled = False
            Tinc1.Enabled = True
        Else
            C1Th.Enabled = True
            Tinc1.Enabled = False
        End If

        If C_Th(2) = False Then
            C2Th.Enabled = False
            Tinc2.Enabled = True
        Else
            C2Th.Enabled = True
            Tinc2.Enabled = False
        End If

        If C_Th(3) = False Then
            C3Th.Enabled = False
            Tinc3.Enabled = True
        Else
            C3Th.Enabled = True
            Tinc3.Enabled = False
        End If

        If C_Th(4) = False Then
            C4Th.Enabled = False
            Tinc4.Enabled = True
        Else
            C4Th.Enabled = True
            Tinc4.Enabled = False
        End If

        If C_Th(5) = False Then
            C5Th.Enabled = False
            Tinc5.Enabled = True
        Else
            C5Th.Enabled = True
            Tinc5.Enabled = False
        End If

        If C_Th(6) = False Then
            C6Th.Enabled = False
            Tinc6.Enabled = True
        Else
            C6Th.Enabled = True
            Tinc6.Enabled = False
        End If

        If C_Th(7) = False Then
            C7Th.Enabled = False
            Tinc7.Enabled = True
        Else
            C7Th.Enabled = True
            Tinc7.Enabled = False
        End If

        If C_Th(8) = False Then
            C8Th.Enabled = False
            Tinc8.Enabled = True
        Else
            C8Th.Enabled = True
            Tinc8.Enabled = False
        End If
    End Sub

    Private Sub C1Th_Click(sender As Object, e As EventArgs) Handles C1Th.Click
        ClearChart()
        SelectedBit = 1

        C_Th(1) = False
        For c = 1 To 8
            If c <> 1 Then
                C_Th(c) = True
            End If
        Next

        CheckButt()
        DrawChart()

        CheckQ1()
    End Sub

    Private Sub C2Th_Click(sender As Object, e As EventArgs) Handles C2Th.Click
        ClearChart()
        SelectedBit = 2

        C_Th(2) = False
        For c = 1 To 8
            If c <> 2 Then
                C_Th(c) = True
            End If
        Next

        CheckButt()
        DrawChart()

        CheckQ2()
    End Sub

    Private Sub C3Th_Click(sender As Object, e As EventArgs) Handles C3Th.Click
        ClearChart()
        SelectedBit = 3

        C_Th(3) = False
        For c = 1 To 8
            If c <> 3 Then
                C_Th(c) = True
            End If
        Next

        CheckButt()
        DrawChart()

        CheckQ3()
    End Sub

    Private Sub C4Th_Click(sender As Object, e As EventArgs) Handles C4Th.Click
        ClearChart()
        SelectedBit = 4

        C_Th(4) = False
        For c = 1 To 8
            If c <> 4 Then
                C_Th(c) = True
            End If
        Next

        CheckButt()
        DrawChart()

        CheckQ4()
    End Sub

    Private Sub C5Th_Click(sender As Object, e As EventArgs) Handles C5Th.Click
        ClearChart()
        SelectedBit = 5

        C_Th(5) = False
        For c = 1 To 8
            If c <> 5 Then
                C_Th(c) = True
            End If
        Next

        CheckButt()
        DrawChart()

        CheckQ5()
    End Sub

    Private Sub C6Th_Click(sender As Object, e As EventArgs) Handles C6Th.Click
        ClearChart()
        SelectedBit = 6

        C_Th(6) = False
        For c = 1 To 8
            If c <> 6 Then
                C_Th(c) = True
            End If
        Next

        CheckButt()
        DrawChart()

        CheckQ6()
    End Sub

    Private Sub C7Th_Click(sender As Object, e As EventArgs) Handles C7Th.Click
        ClearChart()
        SelectedBit = 7

        C_Th(7) = False
        For c = 1 To 8
            If c <> 7 Then
                C_Th(c) = True
            End If
        Next

        CheckButt()
        DrawChart()

        CheckQ7()
    End Sub

    Private Sub C8Th_Click(sender As Object, e As EventArgs) Handles C8Th.Click
        ClearChart()
        SelectedBit = 8

        C_Th(8) = False
        For c = 1 To 8
            If c <> 8 Then
                C_Th(c) = True
            End If
        Next

        CheckButt()
        DrawChart()

        CheckQ8()
    End Sub

    'INITIALIZING
    Private Sub btnClear_Click(sender As Object, e As EventArgs) Handles btnClear.Click
        VenusTemp = 733

        Dim p As String = "Binary Value: 0"
        Dim k As String = "733K"
        For c = 1 To 8
            BinaryV(c) = False
        Next
        For c = 1 To 8
            DataGridView1.Columns("Column" & c).DefaultCellStyle.BackColor = Color.White
        Next
        ClearChart()

        Bin1.Text = p
        Bin2.Text = p
        Bin3.Text = p
        Bin4.Text = p
        Bin5.Text = p
        Bin6.Text = p
        Bin7.Text = p
        Bin8.Text = p

        Th1.Text = k
        Th2.Text = k
        Th3.Text = k
        Th4.Text = k
        Th5.Text = k
        Th6.Text = k
        Th7.Text = k
        Th8.Text = k

        THeadT1.Text = k
        THeadT2.Text = k
        THeadT3.Text = k
        THeadT4.Text = k
        THeadT5.Text = k
        THeadT6.Text = k
        THeadT7.Text = k
        THeadT8.Text = k

        Tinc1B = False
        Tinc2B = False
        Tinc3B = False
        Tinc4B = False
        Tinc5B = False
        Tinc6B = False
        Tinc7B = False
        Tinc8B = False

        Tinc1.Value = 0
        Tinc2.Value = 0
        Tinc3.Value = 0
        Tinc4.Value = 0
        Tinc5.Value = 0
        Tinc6.Value = 0
        Tinc7.Value = 0
        Tinc8.Value = 0

        PictureBox1.Size = New Size(10, 73)
        PictureBox2.Size = New Size(10, 73)
        PictureBox3.Size = New Size(10, 73)
        PictureBox4.Size = New Size(10, 73)
        PictureBox5.Size = New Size(10, 73)
        PictureBox6.Size = New Size(10, 73)
        PictureBox7.Size = New Size(10, 73)
        PictureBox8.Size = New Size(10, 73)

        TB1.Location = New Point(6, 139)
        TB2.Location = New Point(6, 139)
        TB3.Location = New Point(6, 139)
        TB4.Location = New Point(6, 139)
        TB5.Location = New Point(6, 139)
        TB6.Location = New Point(6, 139)
        TB7.Location = New Point(6, 139)
        TB8.Location = New Point(6, 139)

        Label2.Location = New Point(95, 145)
        Label3.Location = New Point(95, 145)
        Label4.Location = New Point(95, 145)
        Label5.Location = New Point(95, 145)
        Label6.Location = New Point(95, 145)
        Label7.Location = New Point(95, 145)
        Label8.Location = New Point(95, 145)
        Label9.Location = New Point(95, 145)

        BinaryDecimal.Text = "Decimal Value: 0"
        lblHex.Text = "Hexadecimal Value: 00"
    End Sub
End Class
