Attribute VB_Name = "MD_0002_CampoComoDATA"
'Option Explicit

'==========================================================================================================================================
'------------   FORMATAÇÃO DE CAMPOS EM FORMATO DE DATA -----------------------------------------------------------------------------------
'==========================================================================================================================================

    Private Sub TextBox0_6_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
      TextBox0_6.MaxLength = 10 'quantidade de digitos a 10  00/00/0000

        Select Case KeyAscii
        Case 48 To 57
           If TextBox0_6.SelStart = 2 Then TextBox0_6.SelText = "/"
           If TextBox0_6.SelStart = 5 Then TextBox0_6.SelText = "/"
        Case Else: KeyAscii = 0     'permitir que apenas números sejam digitados
        End Select
     
    End Sub
   
    Private Sub TextBox0_7_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
      TextBox0_7.MaxLength = 10 'quantidade de digitos a 10  00/00/0000

        Select Case KeyAscii
        Case 48 To 57
           If TextBox0_7.SelStart = 2 Then TextBox0_7.SelText = "/"
           If TextBox0_7.SelStart = 5 Then TextBox0_7.SelText = "/"
        Case Else: KeyAscii = 0     'permitir que apenas números sejam digitados
        End Select
     
    End Sub
