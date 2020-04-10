# Banco de codigos VBA Excel

Banco de codigos e macros VBA

~~~VBA
```
	Sub cidades()

		Range("A1", "D28").Copy
		
		Sheets.Add
		
		Range("A1").PasteSpecial
		
		Selection.EntireColumn.AutoFit

	End Sub
```
~~~
