Option Compare Database
Option Explicit

Private Sub cmdBrincando_Click()
    With Assistant
        .Visible = True
        .Animation = msoAnimationAppear
        .Animation = msoAnimationGestureUp
        .Visible = False
        
    End With
End Sub

Private Sub cmdChegando_Click()
    With Assistant
        .Visible = True
        .Animation = msoAnimationAppear
    End With
End Sub

Private Sub cmdSaindo_Click()

    With Assistant
        .Visible = True
        .Animation = msoAnimationGetAttentionMinor
        .Animation = msoAnimationDisappear
        .Visible = False
    End With

End Sub


Private Sub Comando3_Click()
 DoCmd.SendObject acSendReport, "rptOrcamentos", acFormatHTML, "ailtonzsilva@yahoo.com.br", , , "Teste[1]", "Olá", False
End Sub
