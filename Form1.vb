Public Class Form1
    Dim meter As NationalInstruments.NI4882.Device
    Dim tipeA As Integer = 2
    Dim l As ListViewItem
    Dim i As Integer
    Dim exc As New Excel.Application
    Dim wbk As Excel.Workbook
    Dim sht As Excel.Worksheet

    Dim InstRead(100) As Double 'tentukan jumlah pengambilan data (tipe A)
    Dim abort As Integer
    Dim reading As String
    Dim iterasi As Integer
    Dim baris As Integer
    Dim channel As String
    Private Property SaveFileDialog1 As Object


    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If TextBox1.Text = "" Then
            MsgBox("Harap isikan GPIB Address!")
            'reset()
            ' Exit Sub
        Else if TextBox3.Text = "" Then
            MsgBox("Harap isikan Gate Time!")
            'reset()
            'Exit Sub
        ElseIf TextBox2.Text = "" Then
            MsgBox("Jumlah Pengukuran!")
            ' reset()
            'Exit Sub
        Else
            RadioButton1.Enabled = True
            RadioButton2.Enabled = True
            RadioButton3.Enabled = True
            'ComboBox1.Enabled = True
            'ComboBox2.Enabled = True
            ' ComboBox3.Enabled = True
            ' ComboBox4.Enabled = True
            meter = New NationalInstruments.NI4882.Device(0, TextBox1.Text, 0, TimeoutValue.T100s) 'masukan alamat GPIB untuk meter
            'combobox1
            Label4.Enabled = True
            Label4.Text = "Jml. data terambil"
            TextBox4.Enabled = True
            TextBox4.Show()
            'iterasi

            For Me.baris = 1 To Me.TextBox2.Text
                l = Me.ListView1.Items.Add("")
                For j As Integer = 1 To Me.ListView1.Columns.Count
                    l.SubItems.Add("")
                Next
                For Me.iterasi = 2 To tipeA
                    If Button2.Enabled = False Then
                        Exit Sub
                    End If
                    meter.Write(":func 'FREQ" + channel + "'")
                    meter.Write(":FREQ:ARM:STAR:SOUR IMM") 'These 3 lines enable using
                    meter.Write(":FREQ:ARM:STOP:SOUR TIM") 'time arming with a 0.1 second
                    meter.Write(":FREQ:ARM:STOP:TIM" + TextBox3.Text)
                    Button4.Enabled = False
                    Button3.Enabled = False

                    ' meter.Write("Enter Cnt;A$")
                    meter.Write(":READ:FREQ?") '--->mendisplaykan angka
                    ' meter.Write("GOTO 1040")
                    'meter.Write("GT5")

                    'meter.Write("GOTO 1060")                 
                    'meter.Write("END")
                    InstRead(0) = Val(meter.ReadString)
                    'reading = Val(meter.ReadString)
                    'reading = InstRead(0) 'Format(InstRead(0), "##0.000 000" + " Hz")
                    reading = Format(InstRead(0), "")
                    'reading = Format(reading, "000 000 000 000.000 000")
                    'TextBox3.Text = reading
                    'reading = reading.Replace("F   ", "")
                    'reading = Val(meter.ReadString)
                    'meter.Write("s5")
                    wait(1)
                    ' If Button2.Enabled = True Then
                    'Exit Sub
                    'Else
                    Me.ListView1.Items(baris - 1).SubItems(iterasi - 1).Text() = reading
                    'wait(ComboBox3.SelectedItem)
                    ' meter.Write("SR5")
                    'End If
                    'Dim i As String = 1
                    Me.ListView1.Items(baris - 1).SubItems(0).Text = baris

                    Me.ListView1.Items(baris - 1).SubItems(2).Text = Date.Now.ToString("HH:mm:ss")

                    Me.ListView1.Items(baris - 1).SubItems(3).Text = Date.Now.ToString("dd/MM/yyyy")
                    TextBox4.Text = baris
                Next
            Next
            Button3.Enabled = True
            Button4.Enabled = True
            'meter.Write("A$")

        End If
        MsgBox("Pengukuran Selesai!")
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

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        ListView1.Items.Clear()
        Button2.Enabled = True

    End Sub




    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click

        Button2.Enabled = False
        Button4.Enabled = True
        Button3.Enabled = True
        Button1.Enabled = True
        RadioButton1.Enabled = True
        RadioButton2.Enabled = True
        RadioButton3.Enabled = True
        'ComboBox1.Enabled = True
        '  ComboBox2.Enabled = True
        'ComboBox3.Enabled = True
        ' ComboBox4.Enabled = True
        TextBox1.Enabled = True
        TextBox2.Enabled = True

        Button1.Text = "Return"
        Exit Sub

    End Sub

    Private Sub Button1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

    End Sub


    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim xlApp As Excel.Application
        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkSheet As Excel.Worksheet
        Dim misValue As Object = System.Reflection.Missing.Value
        Dim i As Integer

        xlApp = New Excel.Application
        xlWorkBook = xlApp.Workbooks.Add(misValue)
        xlWorkSheet = xlWorkBook.Sheets("sheet1")
        Dim col As Integer = 1
        For j As Integer = 0 To ListView1.Columns.Count - 1
            xlWorkSheet.Cells(1, col) = ListView1.Columns(j).Text.ToString
            col = col + 1
        Next


        For i = 0 To ListView1.Items.Count - 1
            xlWorkSheet.Cells(i + 2, 1) = ListView1.Items.Item(i).Text.ToString
            xlWorkSheet.Cells(i + 2, 2) = ListView1.Items.Item(i).SubItems(1).Text
            'xlWorkSheet.Cells(i + 2, 3) = ListView1.Items.Item(i).SubItems(2).Text
            'xlWorkSheet.Cells(i + 2, 4) = ListView1.Items.Item(i).SubItems(3).Text

        Next
        Dim dlg As New SaveFileDialog
        dlg.Filter = "Excel Files (*.xlsx)|*.xlsx"
        dlg.FilterIndex = 1
        dlg.InitialDirectory = My.Application.Info.DirectoryPath & "\EXCEL\\EICHER\BILLS\"
        dlg.FileName = " "
        Dim ExcelFile As String = ""
        If dlg.ShowDialog = Windows.Forms.DialogResult.OK Then
            ExcelFile = dlg.FileName
            xlWorkSheet.SaveAs(ExcelFile)
        End If
        xlWorkBook.Close()

        xlApp.Quit()


    End Sub




    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        TextBox4.Hide()
        Label4.Text = ""
        TextBox6.Text = Date.Now.ToString("dd/MM/yyyy")
        'RadioButton3.Checked = False
        ' RadioButton1.Checked = False
        'RadioButton2.Checked = False
    End Sub

    Private Sub reset()
        Button2.Enabled = False
        Button4.Enabled = True
        Button3.Enabled = True
        Button1.Enabled = True
        RadioButton1.Enabled = True
        RadioButton2.Enabled = True
        RadioButton3.Enabled = True
        'ComboBox1.Enabled = True
        ' ComboBox2.Enabled = True
        '  ComboBox3.Enabled = True
        ' ComboBox4.Enabled = True
        TextBox1.Enabled = True
        TextBox2.Enabled = True
        TextBox4.Hide()
        Label4.Text = ""
    End Sub

    Private Sub RadioButton1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton1.CheckedChanged
        RadioButton1.Enabled = True
        channel = "1"
    End Sub

    Private Sub RadioButton2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton2.CheckedChanged
        channel = "2"
    End Sub

    Private Sub RadioButton3_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton3.CheckedChanged
        channel = "3"
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        RadioButton1.PerformClick()
        TextBox3.Text = "1.0"
        TextBox1.Text = "3"
        TextBox2.Text = "10"
        TextBox4.Text = ""
    End Sub

   
    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        TextBox5.Text = Date.Now.ToString("HH:mm:ss")
    End Sub
End Class
