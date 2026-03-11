Attribute VB_Name = "AMCEqnNum"
Sub LaTeX转公式()
    Dim cleanText As String
    If Selection.Type = wdSelectionIP Then Exit Sub
    
    cleanText = Selection.Text
    
    ' 1. 脱去外壳与清理多行的换行符
    cleanText = Replace(cleanText, "$$", "")
    cleanText = Replace(cleanText, "$", "")
    cleanText = Replace(cleanText, "\[", "")
    cleanText = Replace(cleanText, "\]", "")
    cleanText = Replace(cleanText, vbCr, "")
    cleanText = Replace(cleanText, vbLf, "")
    cleanText = Trim(cleanText)
    
    ' 2. 原地替换为干净的代码
    Selection.Text = cleanText
    
    ' 3. 划定公式区域，Word 会自动用你刚设置的全局 LaTeX 基因去解析它！
    Selection.OMaths.Add Range:=Selection.Range
    Selection.OMaths(1).BuildUp
    
    ' 4. 光标移出公式，继续码字
    Selection.Collapse Direction:=wdCollapseEnd
End Sub
