'''''''''''''''''
'@desc  template engine
'       normal variable with {}
'       {=} for formulas in R1C1-Form
'       within the formula, the variables are wrapped with  {}
'@param rng     the target Range to be filled in
'       data    Dicts Object with data. Keys are corresponding to the variables in the template
'''''''''''''''''
Private Sub fillTheTemplate(rng As Range, data As Dicts)
    Dim reg_formula As Object
    Set reg_formula = CreateObject("vbscript.regexp")
    
    With reg_formula
        .pattern = "^{(=\S*{\S+}\S*)}$"
    End With
    
    Dim reg As Object
    Set reg = CreateObject("vbscript.regexp")
    
    With reg
        .pattern = "({\S+})"
    End With

    For Each c In rng.Cells
        tmpval = c.Value
        If Not IsEmpty(tmpval) And Not c.HasFormula Then
            If reg.Test(tmpval) Then
                If reg_formula.Test(tmpval) Then
                    tmpval = reg_formula.Execute(tmpval)(0).submatches(0)
                    
                    For Each k In data.dict.keys
                        tmpval = Replace(tmpval, k, data.dict(k))
                    Next k
                    
                    c.FormulaR1C1 = tmpval
                    
                Else
                    For Each k In data.dict.keys
                        tmpval = Replace(tmpval, k, data.dict(k))
                    Next k
                    
                    c.Value = tmpval
                End If
                
                
                If reg.Test(tmpval) Then
                '    Debug.Print reg.Execute(tmpval)(0).submatches(0)
                    Err.Raise 9999, , "variable" & reg.Execute(tmpval)(0).submatches(0) & "in template not found"
                End If

            End If
        End If
    Next c
End Sub
