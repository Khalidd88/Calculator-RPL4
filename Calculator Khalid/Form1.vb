Imports System.Data

Public Class Form1
    Dim firstNum As Double
    Dim secondNum As Double
    Dim operation As String
    Dim result As Double


    Private KalkulatorAktif As Boolean = False
    Private ekspresi As String = ""
    Private runningText As String = "Welcome to the Calculator!    " ' Add extra spaces for smoother movement

    ' Button ON
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        MessageBox.Show("Kalkulator ON", "Notifikasi", MessageBoxButtons.OK, MessageBoxIcon.Information)
        ActivateCalculator(True)
    End Sub

    ' Button OFF
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        MessageBox.Show("Kalkulator OFF", "Notifikasi", MessageBoxButtons.OK, MessageBoxIcon.Information)
        ActivateCalculator(False)
    End Sub

    ' Method untuk mengaktifkan/mematikan kalkulator
    Private Sub ActivateCalculator(ByVal isActive As Boolean)
        KalkulatorAktif = isActive
        TextBox2.Clear()
        ekspresi = ""

        If isActive Then
            Label9.Visible = True   ' Tampilkan label ON
            Label10.Visible = False  ' Sembunyikan label OFF
            Me.BackColor = Color.White ' Ubah latar belakang menjadi putih

            ' Enable all buttons when ON
            For i As Integer = 1 To 20
                Dim btn As Control = Me.Controls.Find("Button" & i, True).FirstOrDefault()
                If btn IsNot Nothing Then btn.Enabled = True
            Next
        Else
            Label9.Visible = False   ' Sembunyikan label ON
            Label10.Visible = True    ' Tampilkan label OFF
            Me.BackColor = Color.LightGray ' Ubah latar belakang menjadi abu-abu

            ' Disable all buttons when OFF
            For i As Integer = 2 To 20
                Dim btn As Control = Me.Controls.Find("Button" & i, True).FirstOrDefault()
                If btn IsNot Nothing Then btn.Enabled = False
            Next
        End If
    End Sub

    ' Method untuk menambahkan angka/operator ke ekspresi di akhir
    Private Sub AppendToExpression(ByVal value As String)
        ekspresi &= value
        TextBox2.Text = ekspresi
    End Sub

    ' Tombol angka dan operator
    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        If KalkulatorAktif Then AppendToExpression("1")
    End Sub
    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        If KalkulatorAktif Then AppendToExpression("2")
    End Sub
    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        If KalkulatorAktif Then AppendToExpression("3")
    End Sub
    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        If KalkulatorAktif Then AppendToExpression("4")
    End Sub
    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
        If KalkulatorAktif Then AppendToExpression("5")
    End Sub
    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click
        If KalkulatorAktif Then AppendToExpression("6")
    End Sub
    Private Sub Button17_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button17.Click
        If KalkulatorAktif Then AppendToExpression("7")
    End Sub
    Private Sub Button16_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button16.Click
        If KalkulatorAktif Then AppendToExpression("8")
    End Sub
    Private Sub Button15_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button15.Click
        If KalkulatorAktif Then AppendToExpression("9")
    End Sub
    Private Sub Button18_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button18.Click
        If KalkulatorAktif Then AppendToExpression("0")
    End Sub
    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        If KalkulatorAktif Then AppendToExpression("*")
    End Sub
    Private Sub Button20_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button20.Click
        If KalkulatorAktif Then AppendToExpression("+")
    End Sub
    Private Sub Button12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button12.Click
        If KalkulatorAktif Then AppendToExpression("-")
    End Sub
    Private Sub Button14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button14.Click
        If KalkulatorAktif Then AppendToExpression("/")
    End Sub

    ' Tombol =
    Private Sub Button13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button13.Click
        If KalkulatorAktif Then
            Try
                Dim hasil As Double = EvaluateExpression(ekspresi)
                TextBox2.Text = hasil.ToString()
                ekspresi = hasil.ToString()
            Catch ex As Exception
                MessageBox.Show("Ekspresi tidak valid", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                TextBox2.Clear()
                ekspresi = ""
            End Try
        End If
    End Sub

    ' Method untuk mengevaluasi ekspresi menggunakan DataTable untuk menangani urutan operator
    Private Function EvaluateExpression(ByVal exp As String) As Double
        Dim table As New DataTable()
        Return Convert.ToDouble(table.Compute(exp, String.Empty))
    End Function

    ' Tombol C (clear)
    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        TextBox2.Clear()
        ekspresi = ""
    End Sub

    ' Tombol Backspace
    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        If TextBox2.Text.Length > 0 Then
            TextBox2.Text = TextBox2.Text.Substring(0, TextBox2.Text.Length - 1)
            ekspresi = ekspresi.Substring(0, ekspresi.Length - 1)
        End If
    End Sub

    ' Button untuk desimal
    Private Sub Button19_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button19.Click
        If KalkulatorAktif AndAlso Not ekspresi.EndsWith(".") Then
            AppendToExpression(".")
        End If
    End Sub

    ' Button untuk menutup aplikasi
    Private Sub Button21_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button21.Click
        Me.Close()
    End Sub

    ' Timer untuk jam, menit, detik
    Private Sub Timer2_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer2.Tick
        Label5.Text = DateTime.Now.ToString("HH:mm:ss") ' Format jam, menit, detik
    End Sub

    ' Timer untuk running text
    Private Sub Timer3_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer3.Tick
        If Label6.Left >= Me.Width Then
            Label6.Left = -Label6.Width
        Else
            Label6.Left = Label6.Left + 20
        End If
    End Sub

    ' Form Load event
    Private Sub Form1_Load(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Load
        ' Set interval untuk Timer1 agar memperbarui setiap detik
        Timer1.Interval = 1000 ' Set interval 1 detik
        Timer1.Start()
        Label6.Text = runningText ' Set the label text to running text
        Label6.Left = Me.Width ' Set initial position of the label to the right edge of the form
        Timer3.Interval = 50 ' Set timer interval to 50ms for smoother movement
        Timer3.Start() ' Start the timer for running text

        Timer2.Interval = 1000 ' Set interval for Timer2 to 1 second
        Timer2.Start() ' Start the timer for real-time clock
    End Sub

    ' Timer untuk tahun, tanggal, waktu (YY/MM/DD)
    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        ' Tampilkan tanggal dalam format YY/MM/DD dengan jam dan menit
        Label3.Text = DateTime.Now.ToString("yy/MM/dd HH:mm:ss")
    End Sub


End Class
