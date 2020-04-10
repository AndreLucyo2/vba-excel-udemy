Attribute VB_Name = "MD_0001_ConverterDataValor"
Option Explicit

Sub COVERTERDATAEMVALOR()

 Dim V_DT As Date
 
 V_DT = sht_BD_CHAVES.Range("DT_V1").Value 'DATA NO FORMATA 00/00/0000
 
 MsgBox CDbl(DateValue(V_DT)) 'RETORNA DATA EM VALOR : EX.: 42751

End Sub


Sub COVERTERDATAEMTEXTO()

 Dim V_DT As Date
 
 V_DT = sht_BD_CHAVES.Range("DT_V1").Value 'DATA entrada
 
 MsgBox Format(V_DT, "dd/mm/yyyy") 'RETORNA DATA EM TEXTO : EX.: 00/00/0000

End Sub
