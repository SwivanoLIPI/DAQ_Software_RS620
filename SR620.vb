Imports System.Windows.Forms.DataVisualization.Charting

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
    Dim reading1 As String
    Dim iterasi As Integer
    Dim baris As Integer
    Dim channel As String
    Dim md As String
    Dim arm As String
    Dim sz As String
    Dim disp As String
    Dim clk As String
    Dim fsrc As String
    Dim src As String
    Dim imp As String
    Dim refo As String
    Dim cpl As String
    Dim slp As String
    Dim alv As String
    Dim scr As String
    Dim gt As String
    'Dim idn As String
    Dim chn As String
    Private Property SaveFileDialog1 As Object



    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        'kondisi jika mode tidak dipilih

        If RadioButton3.Checked Or RadioButton4.Checked Or RadioButton5.Checked Or RadioButton6.Checked Or RadioButton7.Checked Or RadioButton8.Checked Or RadioButton9.Checked Or RadioButton10.Checked Or RadioButton13.Checked Then

            md = md
        Else
            MsgBox("Please choose measurement mode")
            Exit Sub
        End If
        'jika cuztomized mode belum dipilih 
        ' If TextBox4.Enabled = True And TextBox4.Text = "" Then
        'MsgBox("Please fill the syntax in cuztomized mode box !")
        'Label11.ForeColor = Color.Red

        'Exit Sub
        'End If

        'If TextBox1.Enabled = True And TextBox1.Text = "" Then
        'MsgBox("Please fill the function in cuztomized mode box !")
        'Label10.ForeColor = Color.Red
        'Exit Sub
        'End If

        ''Label18.Text = ""
        TextBox7.Text = ""
        TextBox8.Text = ""
        TextBox9.Text = ""
        TextBox9.Text = ""
        TextBox10.Text = ""
        TextBox12.Text = ""
        Label4.Text = "Recent data taken"

        Button1.Enabled = False
        Button3.Enabled = False
        Button4.Enabled = False
        Button5.Enabled = False
        Button6.Enabled = False
        Button7.Enabled = False
        Button8.Enabled = False
        '  Button4.Enabled = False
        '  Button2.Enabled = True

        'If RadioButton1.Checked Then

        '  If TextBox1.Text = "" Then
        ' MsgBox("Harap isikan GPIB Address!")
        '
        ' Exit Sub
        '   ElseIf ComboBox1.Text = "" Then
        '  MsgBox("Harap isikan Gate / ARM!")

        '   Exit Sub
        '  ElseIf TextBox2.Text = "" Then
        '  MsgBox("Jumlah Pengukuran!")

        ' Exit Sub
        '  End If
        ' Button1.Enabled = False
        '  RadioButton1.Enabled = False
        '  RadioButton2.Enabled = False
        ' RadioButton3.Enabled = False
        '  TextBox6.Enabled = False
        '  ComboBox1.Enabled = False
        '  TextBox1.Enabled = False
        '  TextBox2.Enabled = False

        ' Button6.Enabled = False
        '  Button3.Enabled = False

        'menentukan level on/off

        If RadioButton12.Checked Then
            alv = "1"
        Else
            alv = "0"
        End If

        'menentukan slope

        If ComboBox8.Text = "Pos +" Then
            slp = "0"
        Else
            slp = "1"
        End If

        'menentukan coupling

        If ComboBox11.Text = "AC" Then
            cpl = "1"
        Else
            cpl = "0"
        End If



        'MENENTUKAN source

        If RadioButton1.Checked Then
            src = "SRCE0"
            chn = "1"
        ElseIf RadioButton2.Checked Then
            src = "SRCE1"
            chn = "2"
        ElseIf RadioButton3.Checked Then
            src = "SRCE2"
            chn = "0"
        Else
            src = "SRC3"
            'chn = "3"
        End If

        'MENENTUKAN refference output level

        If ComboBox4.Text = "TTL" Then
            refo = "1"
        Else
            refo = "0"
        End If

        'MENENTUKAN MODE
        If RadioButton6.Checked Then
            md = "MODE0"
            Form2.Label21.Text = "Time Interval"
            ColumnHeader2.Text = "Time(s)"
        ElseIf RadioButton5.Checked Then
            md = "MODE1"
            ColumnHeader2.Text = "Width"
            Form2.Label21.Text = ColumnHeader2.Text
        ElseIf RadioButton4.Checked Then
            md = "MODE2"
            ColumnHeader2.Text = "Tr/Tf"
            Form2.Label21.Text = ColumnHeader2.Text
        ElseIf RadioButton7.Checked Then
            md = "MODE3"
            ColumnHeader2.Text = "Frequency(Hz)"
            Form2.Label21.Text = ColumnHeader2.Text
        ElseIf RadioButton8.Checked Then
            md = "MODE4"
            ColumnHeader2.Text = "Period(s)"
            Form2.Label21.Text = ColumnHeader2.Text
        ElseIf RadioButton9.Checked Then
            md = "MODE5"
            ColumnHeader2.Text = "Phase"
            Form2.Label21.Text = ColumnHeader2.Text
        ElseIf md = "MODE6" Then
            ColumnHeader2.Text = "Count"
            Form2.Label21.Text = ColumnHeader2.Text
        Else
            md = TextBox4.Text
            ColumnHeader2.Text = TextBox1.Text
            Form2.Label21.Text = TextBox1.Text
        End If


        'MENENTUKAN ARM

        If ComboBox1.Text = "+/- Time" Then
            arm = "ARMM0"
        ElseIf ComboBox1.Text = "+Time" Then
            arm = "ARMM1"
        ElseIf ComboBox1.Text = "0.01 s" Then
            arm = "ARMM3"
        ElseIf ComboBox1.Text = "0.1 s" Then
            arm = "ARMM4"
        ElseIf ComboBox1.Text = "1 s" Then
            arm = "ARMM5"
        ElseIf ComboBox1.Text = "1 Period" Then
            arm = "ARMM2"
        ElseIf ComboBox1.Text = "Ext trig +/- time" Then
            arm = "ARMM6"
        ElseIf ComboBox1.Text = "Ext trig + time" Then
            arm = "ARMM7"
        ElseIf ComboBox1.Text = "Ext gate/trig holdoff" Then
            arm = "ARMM8"
        ElseIf ComboBox1.Text = "Ext 0.01 s" Then
            arm = "ARMM10"
        ElseIf ComboBox1.Text = "Ext 0.1 s" Then
            arm = "ARMM11"
        ElseIf ComboBox1.Text = "Ext 1 s" Then
            arm = "ARMM12"
        Else
            arm = "ARMM9"
        End If

        'menentukan gate

        gt = Val(ComboBox10.Text * ComboBox12.Text)

        'MENENTUKAN SAMPLE SIZE

        sz = CDec(ComboBox3.Text * ComboBox2.Text)

        'MENENTUKAN MODE DISPLAY

        If ComboBox5.Text = "MEAN" Then
            disp = "XAVG?"
            scr = "DISP0"
            'ColumnHeader2.Text = ""
        ElseIf ComboBox5.Text = "REL" Then
            disp = "XREL" + 1
            scr = "DISP1"
        ElseIf ComboBox5.Text = "JITTER" Then
            disp = "XJIT?"
            scr = "DISP2"
        ElseIf ComboBox5.Text = "MAX" Then
            disp = "XMAX?"
            scr = "DISP3"
        ElseIf ComboBox5.Text = "MIN" Then
            disp = "XMIN?"
            scr = "DISP4"
        ElseIf ComboBox5.Text = "TRIG" Then
            scr = "DISP5"

        Else
            disp = "VOLT?" + chn
            scr = "DISP6"

        End If

        'MENENTUKAN MODE SOURCE

        If ComboBox6.Text = "Internal" Then
            clk = "CLCK0"
            ComboBox7.Enabled = False
        End If

        If clk = "CLCK1" Then
            ComboBox7.Enabled = True
        End If

        'MENENTUKAN FREQUENCY SOURCE
        If ComboBox7.Text = "5 MHz" Then
            fsrc = "CLKF1"
        ElseIf ComboBox7.Text = "10 MHz" Then
            fsrc = "CLKF0"
        Else
            ComboBox6.Text = "Internal"
            'fsrc = "CLKF0"
            ComboBox7.Enabled = False
        End If

        'MENENTUKAN IMPEDANCE

        If ComboBox9.Text = "50 Ohm" Then
            imp = "0"
        ElseIf ComboBox9.Text = "1 MOhm" Then
            imp = "1"
        Else
            imp = "2"
        End If


        meter = New NationalInstruments.NI4882.Device(0, 16, 0, TimeoutValue.T100s) 'masukan alamat GPIB untuk meter
        'combobox1

        'menentukan gate/arm
        '  If ComboBox1.Text = "1.0 s" Then
        'gate = "ARMM5"
        ' End If
        ' meter.Write(cpl)
        meter.Write(TextBox5.Text)
        'meter.write(TextBox7.Text)
        'meter.write(TextBox9.Text)
        meter.Write("LEVL" + chn + "," + TextBox11.Text)
        meter.Write("TMOD" + chn + "," + alv)
        meter.Write("TSLP" + chn + "," + slp)
        meter.Write("TCPL" + chn + "," + cpl)
        meter.Write("LOCL1")
        meter.Write("AUTM1")
        meter.Write("TERM" + chn + "," + imp)
        meter.Write("RLVL" + refo)
        meter.Write(scr)
        meter.Write(gt)
        meter.Write("MEAS0")
        meter.Write(src)
        meter.Write(fsrc)
        meter.Write(clk)
        meter.Write("SIZE" + sz)
        meter.Write(md)
        meter.Write(arm)
        wait(Val(sz) + 1)
        wait(Val(TextBox3.Text.ToString))
        For Me.baris = 1 To Me.TextBox2.Text
            l = Me.ListView1.Items.Add("")
            For j As Integer = 1 To Me.ListView1.Columns.Count
                l.SubItems.Add("")
            Next
            For Me.iterasi = 2 To tipeA

                GroupBox3.Enabled = False
                GroupBox5.Enabled = False
                GroupBox1.Enabled = False
                GroupBox4.Enabled = False
                GroupBox9.Enabled = False
                GroupBox7.Enabled = False
                GroupBox2.Enabled = False
                GroupBox6.Enabled = False
                ' 
                '  meter.Write("SIZE1")
                ' meter.Write(":FREQ:ARM:STAR:SOUR IMM") 'These 3 lines enable using
                ' meter.Write(":FREQ:ARM:STOP:SOUR TIM") 'time arming with a 0.1 second
                ' meter.Write("ARMM5")
                '  Button4.Enabled = False
                ' Button3.Enabled = False

                meter.Write(disp) '--->mendisplaykan angka

                InstRead(0) = Val(meter.ReadString)

                reading1 = CDec(InstRead(0))

                'Val(ComboBox1.Text) - 0.36) 'Val(TextBox3.Text))


                Me.ListView1.Items(baris - 1).SubItems(1).Text = reading1
                Me.ListView1.Items(baris - 1).SubItems(0).Text = Date.Now.ToString("HH:mm:ss")
                Label6.Text = baris
                '  If Button2.Enabled = False Then
                'Exit Sub

                ' End If

                wait(1)

            Next
        Next

        'menentukan allan variance
        Dim p As Integer

        For p = 2 To TextBox2.Text
            If TextBox12.Text = "" Then
                TextBox12.Text = "0"
            End If

            TextBox12.Text = CDec(CDec(TextBox12.Text) + (CDec(ListView1.Items.Item(p - 1).SubItems(1).Text) - CDec(ListView1.Items.Item(p - 2).SubItems(1).Text)) ^ 2)

        Next p

        TextBox12.Text = Math.Sqrt(CDec((TextBox12.Text)) / CDec(2 * CDec(TextBox2.Text - 1)))

        'menentukan titik minimum


        TextBox7.Text = Val(CDec(ListView1.Items(0).SubItems(1).Text))

        For p = 2 To Val(TextBox2.Text)
            If TextBox7.Text = "" Then
                TextBox7.Text = "0"
            End If

            If Val(CDec(ListView1.Items(p - 1).SubItems(1).Text)) > Val(CDec(ListView1.Items(p - 2).SubItems(1).Text)) And Val(CDec(TextBox7.Text)) > Val(CDec(ListView1.Items(p - 2).SubItems(1).Text)) And Val(CDec(TextBox7.Text)) > 0 Then
                TextBox7.Text = Val(CDec(ListView1.Items(p - 2).SubItems(1).Text))
            Else
                TextBox7.Text = Val(CDec(TextBox7.Text))
            End If
        Next p

        'menentukan titik maksimum
        Dim n As Integer
        For n = 1 To TextBox2.Text

            If TextBox10.Text = "" Then
                TextBox10.Text = "0"
            End If

            If CDec(ListView1.Items(n - 1).SubItems(1).Text) > CDec(TextBox10.Text) Then
                TextBox10.Text = CDec(ListView1.Items(n - 1).SubItems(1).Text)
                TextBox10.Text = CDec(TextBox10.Text)
            End If
        Next n


        'menentukan discreapancy

        TextBox8.Text = Val(CDec(TextBox10.Text) - CDec(TextBox7.Text))

        'menentukan rata-rata
        Dim o As Integer
        Dim div As Decimal
        For o = 1 To TextBox2.Text
            If TextBox9.Text = "" Then
                TextBox9.Text = "0"
            End If
            TextBox9.Text = Val(CDec(TextBox9.Text)) + Val(CDec(ListView1.Items.Item(o - 1).SubItems(1).Text))
        Next o
        div = Val(Val(CDec(TextBox9.Text)) / Val(CDec(TextBox2.Text)))
        TextBox9.Text = div

        '  Label4.Enabled = True
        '  Label4.Text = "Jml. data terambil"
        '  Button3.Enabled = True
        '  Button4.Enabled = True

        GroupBox3.Enabled = True
        GroupBox5.Enabled = True
        GroupBox1.Enabled = True
        GroupBox4.Enabled = True
        GroupBox9.Enabled = True
        GroupBox7.Enabled = True
        Button1.Enabled = True
        Button3.Enabled = True
        Button4.Enabled = True
        Button5.Enabled = True
        Button6.Enabled = True
        Button7.Enabled = True
        Button8.Enabled = True


        MsgBox("Pengukuran Selesai!")
        meter.Write("locL0")
        ' reset()
        Exit Sub


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
        Label6.Text = ""
        Button6.Enabled = True
        TextBox7.Text = ""
        TextBox8.Text = ""
        TextBox9.Text = ""
        TextBox9.Text = ""
        TextBox10.Text = ""
        TextBox12.Text = ""
        'TextBox4.Text = ""
        'TextBox5.Text = ""
        'TextBox7.Text = ""
        'TextBox8.Text = ""
        'TextBox9.Text = ""
        Exit Sub
    End Sub




    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        
        Button6.Enabled = True
        Button7.Enabled = True
        Button8.Enabled = True
        Label4.Text = "Jml. Data terambil"
        Button6.Enabled = True
        'TextBox6.Enabled = True
        'TextBox3.Enabled = True
        RadioButton1.Enabled = True
        RadioButton2.Enabled = True
        RadioButton3.Enabled = True
        'TextBox1.Enabled = True
        TextBox2.Enabled = True
        Label6.Text = baris - 1
        Button2.Enabled = False
        Button4.Enabled = True
        Button3.Enabled = True
        Button1.Enabled = True
        RadioButton1.Enabled = True
        RadioButton2.Enabled = True
        RadioButton3.Enabled = True

        'TextBox1.Enabled = True
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
            xlWorkSheet.Cells(i + 2, 2) = CDec(ListView1.Items.Item(i).SubItems(1).Text)

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
        'Dim idn As String
        If Label6.Text = "Label6" Then
            Label6.Text = 0
        End If
        meter = New NationalInstruments.NI4882.Device(0, 16, 0, TimeoutValue.T100s) 'masukan alamat GPIB untuk meter
        meter.Write("*IDN?") '--->mendisplay identitas

        Label23.Text = meter.ReadString
        Label10.Enabled = False
        Label11.Enabled = False
        TextBox1.Enabled = False
        TextBox4.Enabled = False
        TextBox6.Text = ""
        TextBox6.Enabled = False
        TextBox5.Text = ""
        TextBox5.Enabled = False
        GroupBox6.Enabled = False
        GroupBox2.Enabled = False
        'Label23 = idn
        ' Label4.Text = ""
        Label2.Text = Date.Now.ToString("dd/MM/yyyy")
        'Label6.Text = ""
        RadioButton1.Checked = False
        RadioButton4.Checked = False
        Label24.Enabled = False
        TextBox11.Enabled = False
        meter.Write("locL0")
        
        Exit Sub
       

    End Sub

    Private Sub reset()
        Button2.Enabled = False
        Button4.Enabled = True
        Button3.Enabled = True
        Button1.Enabled = True
        RadioButton1.Enabled = True
        RadioButton2.Enabled = True
        RadioButton3.Enabled = True
        'TextBox1.Enabled = True
        TextBox2.Enabled = True
        'TextBox6.Enabled = True
        ' TextBox3.Enabled = True
        Label4.Text = ""

    End Sub




    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        ComboBox10.Text = "1"
        ComboBox9.Text = "50 Ohm"
        ComboBox4.Text = "TTL"
        ComboBox12.Text = "1"
        TextBox3.Text = "0.001"
        ComboBox6.Text = "External"
        ComboBox8.Text = "Pos +"
        ComboBox7.Enabled = True
        ComboBox7.Text = "10 MHz"
        RadioButton7.PerformClick()
        RadioButton1.PerformClick()
        RadioButton12.PerformClick()
        ComboBox5.Text = "MEAN"
        ComboBox3.Text = "1"
        ComboBox2.Text = "1"
        ComboBox1.Text = "1 s"
        ComboBox11.Text = "DC"
        'TextBox1.Text = "16"
        TextBox2.Text = "10"
        ''TextBox4.Text = ""
    End Sub



   
    Private Sub ComboBox6_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox6.SelectedIndexChanged
        If ComboBox6.Text = "Internal" Then
            ComboBox7.Enabled = False
            Label16.Enabled = False
        Else
            ComboBox7.Enabled = True
            Label16.Enabled = True

        End If
    End Sub

    Private Sub GroupBox4_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GroupBox4.Enter

    End Sub

    Private Sub Button5_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click

        Form2.Chart1.Series(0).Points.Clear()
  
        Form2.Show()
    End Sub






    Private Sub RadioButton14_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton14.CheckedChanged
        Label24.Enabled = True
        TextBox11.Enabled = True
    End Sub

    Private Sub RadioButton12_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton12.CheckedChanged
        Label24.Enabled = False
        TextBox11.Enabled = False
    End Sub


    Private Sub ComboBox5_SelectedIndexChanged_1(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If ComboBox5.Text = "TRIG" Then
            MsgBox("This mode has no output in output table. Output just in instrument display")

        End If
    End Sub


    Private Sub RadioButton13_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton13.CheckedChanged
        Label10.Enabled = True
        Label11.Enabled = True
        TextBox1.Enabled = True
        TextBox4.Enabled = True
        GroupBox2.Enabled = True
    End Sub


    Private Sub Label6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label6.Click

    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        TextBox6.Enabled = True
        TextBox5.Enabled = True
        GroupBox6.Enabled = True
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        TextBox6.Text = ""
        TextBox6.Enabled = False
        TextBox5.Text = ""
        TextBox5.Enabled = False
        GroupBox6.Enabled = False
    End Sub

    Private Sub RadioButton9_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton9.CheckedChanged
        GroupBox2.Enabled = False
    End Sub

    Private Sub RadioButton8_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton8.CheckedChanged
        GroupBox2.Enabled = False
    End Sub

    Private Sub RadioButton7_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton7.CheckedChanged
        GroupBox2.Enabled = False
    End Sub

    Private Sub RadioButton10_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton10.CheckedChanged
        GroupBox2.Enabled = False
    End Sub

    Private Sub RadioButton4_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton4.CheckedChanged
        GroupBox2.Enabled = False
    End Sub

    Private Sub RadioButton5_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton5.CheckedChanged
        GroupBox2.Enabled = False
    End Sub

    Private Sub RadioButton6_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton6.CheckedChanged
        GroupBox2.Enabled = False
    End Sub

    Private Sub TextBox4_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox4.TextChanged
        Label11.ForeColor = Color.Black
    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged
        Label10.ForeColor = Color.Black
    End Sub

   
End Class
