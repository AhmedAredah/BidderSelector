Public Class OpenNew 

'Dim OpenedDatabaseAddress As String 
    Sub New()
        InitSkins()
        InitializeComponent()
    End Sub
    Sub InitSkins()
        DevExpress.Skins.SkinManager.EnableFormSkins()
        DevExpress.UserSkins.BonusSkins.Register()
    End Sub

'    Function GetCommandLineArgs() As String()
'   ' Declare variables.
'       Dim separators As String = " "
'       Dim commands As String = Microsoft.VisualBasic.Command()
'       Dim args() As String = commands.Split(separators.ToCharArray)
'       Return args
'End Function
Private Sub SimpleButton1_Click( sender As Object,  e As EventArgs) Handles SimpleButton1.Click
        
        If OpenFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK then
            'System.IO.File.Copy(OpenFileDialog1.FileName,Application.StartupPath & XtraForm1.DatabasePath,True)
            XtraForm1.DatabasePath = OpenFileDialog1.FileName
            XtraForm1.SavedLocation = OpenFileDialog1.FileName
            XtraForm1.Show
            XtraForm1.LoadInformation

            'OpenedDatabaseAddress = OpenFileDialog1.FileName
            'StartupForm.Show()
            'StartupForm.SpinEdit1.Enabled = False
            'StartupForm.ComboBoxEdit1.Enabled = False
            'StartupForm.ComboBoxEdit2.Enabled = False
            Me.hide
        Else

        End If

End Sub

Private Sub SimpleButton2_Click( sender As Object,  e As EventArgs) Handles SimpleButton2.Click
        XtraForm1.DatabasePath = Setting.GetApplicationLocation(Application.ExecutablePath) & "\MainDatabase.BEV"
        XtraForm1.SavedLocation = ""
        Me.Hide
        StartupForm.Show
        StartupForm.SpinEdit1.Enabled = True
        'StartupForm.ComboBoxEdit1.Enabled = True
        'StartupForm.ComboBoxEdit2.Enabled = True
End Sub

    Private Sub OpenNew_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        Me.Dispose
        Application.ExitThread
    End Sub

    Private Sub OpenNew_Load(sender As Object, e As EventArgs) Handles Me.Load
        If (Command <> "") Then
            Me.Opacity = 0
            'MsgBox(My.Application.CommandLineArgs(0).ToString)
            Dim OpenFileName As String = My.Application.CommandLineArgs(0).ToString
            XtraForm1.DatabasePath = OpenFileName
            XtraForm1.SavedLocation = OpenFileName
            XtraForm1.Show
            XtraForm1.LoadInformation
            'Me.Hide
        Else
            Opacity = 100
        End If
    End Sub


    Private Sub OpenNew_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        If (Command <> "") Then
            Me.Hide
        End If
    End Sub
End Class