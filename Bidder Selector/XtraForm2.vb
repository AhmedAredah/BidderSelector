Public Class XtraForm2 

    Sub New()
        InitSkins()
        InitializeComponent()
    End Sub
    Sub InitSkins()
        DevExpress.Skins.SkinManager.EnableFormSkins()
        DevExpress.UserSkins.BonusSkins.Register()
    End Sub

'Private Sub XtraForm2_Shown( sender As Object,  e As EventArgs) Handles MyBase.Shown
'        PrintPreviewBarItem47.Visibility = DevExpress.XtraBars.BarItemVisibility.Never
'        PrintPreviewBarItem5.Visibility = DevExpress.XtraBars.BarItemVisibility.Never
'        PrintPreviewBarItem2.Visibility = DevExpress.XtraBars.BarItemVisibility.Never
'        PrintPreviewBarItem1.Visibility = DevExpress.XtraBars.BarItemVisibility.Never
'End Sub

End Class