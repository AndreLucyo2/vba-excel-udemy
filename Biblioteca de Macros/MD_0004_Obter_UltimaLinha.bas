Attribute VB_Name = "Módulo2"
Option Explicit

Sub Macro1()
Attribute Macro1.VB_ProcData.VB_Invoke_Func = " \n14"
Dim V_ULTLinha As Long
sht_GERAR_TXT.Activate
V_ULTLinha = Range("N" & Rows.Count).End(xlUp).Row + 1
MsgBox V_ULTLinha
End Sub
