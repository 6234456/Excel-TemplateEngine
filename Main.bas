Option Explicit

Sub main()

    Dim d As New Dicts
    
    Dim i
    
    For i = 2 To d.x("data", 1)
        Call d.load("data", 1, i, 1)
        
        addShtWithName d.item("name")
        Call fillTheTemplate(Worksheets(d.item("name")).UsedRange, d)
        
        Set d = Nothing
    Next i
    
    

End Sub

Private Function addShtWithName(shtName As String)
    On Error Resume Next

    Application.ScreenUpdating = False

    Worksheets("template").Copy , Worksheets(Worksheets.Count)
    Worksheets(Worksheets.Count).Name = shtName
    
    Application.ScreenUpdating = True
    
End Function


'''''''''''''''''
'@desc  template engine
'       normal variable with {}
'       {=} for formulas in R1C1-Form
'       within the formula, the variables are wrapped with  {}
'@param rng     the target Range to be filled in
'       data    Dicts Object with data. Keys are corresponding to the variables in the template
'''''''''''''''''
Private Sub fillTheTemplate(ByRef rng As Range, ByRef data As Dicts)
    Dim reg_formula As Object
    Set reg_formula = CreateObject("vbscript.regexp")
    
    With reg_formula
        .pattern = "^{(=\S*{\S+}\S*)}$"
    End With
    
    Dim reg As Object
    Set reg = CreateObject("vbscript.regexp")
    
    With reg
        .pattern = "{(\S+)}"
    End With
    
    Dim c
    Dim tmpval As String

    For Each c In rng.Cells
        tmpval = c.Value
        If Not IsEmpty(tmpval) And Not c.HasFormula Then
            If reg.test(tmpval) Then
                If reg_formula.test(tmpval) Then
                    tmpval = reg_formula.Execute(tmpval)(0).submatches(0)
                    
                    c.FormulaR1C1 = processTemplateStr(data, tmpval)
                Else

                    c.Value = processTemplateStr(data, tmpval)
                End If
                
                
                If reg.test(tmpval) Then
                    Debug.Print reg.Execute(tmpval)(0).submatches(0)
                    Err.Raise 9999, , "variable '" & reg.Execute(tmpval)(0).submatches(0) & "' in template not found"
                End If

            End If
        End If
    Next c
End Sub

Private Function changeDecimalPoint(ByVal n) As Variant
    
    changeDecimalPoint = IIf(IsNumeric(n), Replace("" & n, ",", "."), n)
    
End Function

Private Function processTemplateStr(ByRef d As Dicts, ByRef tmpl As String) As String
    
    Dim k
          
    For Each k In d.dict.Keys
        tmpl = Replace(tmpl, "{" & k & "}", changeDecimalPoint(d.dict(k)))
    Next k
    
    processTemplateStr = tmpl
End Function

