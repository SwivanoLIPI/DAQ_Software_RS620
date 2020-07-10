Public Class Form2


    Private Sub Form2_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
       
        'membuat grafik
        Dim k As Integer

        For k = 1 To Form1.TextBox2.Text
            Chart1.Series(0).Points.AddXY(Form1.ListView1.Items(k - 1).SubItems(0).Text.ToString, Val(CDec(Form1.ListView1.Items(k - 1).SubItems(1).Text)))
            Chart1.ChartAreas(0).AxisY.Minimum = Val(CDec(Form1.TextBox7.Text) - Val(CDec(Form1.TextBox8.Text) / 2))
            Chart1.ChartAreas(0).AxisY.Maximum = Val(CDec(Form1.TextBox10.Text) + Val(CDec(Form1.TextBox8.Text) / 2))
            Chart1.ChartAreas("ChartArea1").AxisX.Title = "Waktu Pengambilan Data"
            Chart1.ChartAreas("ChartArea1").AxisY.Title = "frekuensi (Hz)"
            Chart1.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Spline

        Next k
        

    End Sub
    Public Sub wait(ByVal Dt As Double)
        Dim IDay As Double = Date.Now.DayOfYear
        Dim CDay As Double
        Dim ITime As Double = Date.Now.TimeOfDay.TotalSeconds
        Dim CTime As Double
        Dim DiffDay As Double

        Do
            Application.DoEvents()
            CDay = Date.Now.DayOfYear
            CTime = Date.Now.TimeOfDay.TotalSeconds
            DiffDay = CDay - IDay
            CTime = CTime + 86400 * DiffDay
            If CTime >= ITime + Dt Then Exit Do

        Loop
    End Sub


    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        Dim zoomin As String
        zoomin = Val(Form1.TextBox2.Text.ToString) / 5
        Chart1.ChartAreas(0).AxisX.ScaleView.Size = zoomin
        Form1.TextBox2.Text = zoomin
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        Dim zoomout As String
        zoomout = Val(Form1.TextBox2.Text.ToString) * 5
        Chart1.ChartAreas(0).AxisX.ScaleView.Size = zoomout
        Form1.TextBox2.Text = zoomout
    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        With Chart1.ChartAreas(0)
            .AxisX.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.NotSet
            .AxisY.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.NotSet


        End With
    End Sub

    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
        With Chart1.ChartAreas(0)
            .AxisX.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Dash
            .AxisY.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Dash

        End With
    End Sub

    Private Sub Button13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button13.Click
        With Chart1.ChartAreas(0)
            .AxisX.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Solid
            .AxisY.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Solid
            '.AxisY2.MajorGrid.LineDashStyle = ChartDashStyle.NotSet
        End With
    End Sub

 
End Class