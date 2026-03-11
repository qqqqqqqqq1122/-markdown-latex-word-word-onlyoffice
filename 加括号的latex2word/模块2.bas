Attribute VB_Name = "模块2"
Sub LaTeX原生编号转换()
    Dim rng As Range
    Set rng = Selection.Range
    
    ' 1. 获取选中的文本
    If rng.End = rng.Start Then Exit Sub
    
    Dim rawLatex As String
    rawLatex = rng.Text
    
    ' 2. 净化代码：去掉已有的 $$, $, 换行符
    rawLatex = Replace(rawLatex, "$$", "")
    rawLatex = Replace(rawLatex, "$", "")
    rawLatex = Replace(rawLatex, vbCr, "")
    rawLatex = Replace(rawLatex, vbLf, "")
    rawLatex = Trim(rawLatex)
    
    ' 3. 自动获取下一个编号（基于书签或简单的计数，这里示例手动输入或默认）
    ' 为了最快协作，我们直接在后面挂上 #()，光标最后会停在括号里
    Dim finalCode As String
    finalCode = rawLatex & "#()"
    
    ' 4. 原地替换并转化为公式
    rng.Text = finalCode
    rng.OMaths.Add Range:=rng
    
    Dim objEq As OMath
    Set objEq = rng.OMaths(1)
    
    ' 5. 强制设为 LaTeX 模式并渲染
    ' 这一步会让 # 后的括号自动飞到右边
    objEq.BuildUp
    
    ' 6. 把光标精准定位到公式末尾编号的括号中间，方便你填数字
    rng.Select
    Selection.MoveLeft Unit:=wdCharacter, Count:=2
End Sub
