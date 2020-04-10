Attribute VB_Name = "Módulo2"
Option Explicit

'================================================================================================================  ALS-2017 =====
'------------   GUARDAR O NOME DO FORM ATIVO - CADA UM TEM UM   -----------------------------------------------------------------
'================================================================================================================  ALS-2017 =====
Private Sub UserForm_Activate()
'ao ativar a tela é necessario redefinir a variável:
'---------------------------------------------------------------------------------
'vai capturar qual tela esta ativa e guardar em uma variável publica
'---------------------------------------------------------------------------------
Call M0000_Acao_Definir_Nome_Form_Ativo_VND
End Sub


Sub M0000_Acao_Definir_Nome_Form_Ativo_VND()

'---------------------------------------------------------------------------------
'vai capturar qual tela esta ativa e guardar em uma variável publica
'---------------------------------------------------------------------------------
       Dim UForm As Object
       Dim UFName As String

       UFName = Me.Name

       'zerar a avaiavel:
        Set NOME_FormCarregado = Nothing

       For Each UForm In VBA.UserForms
           If UForm.Name = UFName Then
               Set NOME_FormCarregado = UForm '.Name 'caso o form apontado estiver aberto, retorna true
               Exit For
           End If
       Next
       'MsgBox NOME_FormCarregado.Name

End Sub
