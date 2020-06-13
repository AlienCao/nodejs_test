# nodejs_test

> 一个 nodejs 的小练习而已

##### 练习文章

[NodeJS_test][1]

##### 安装

	npm install nodejs_test
##### 运行
	hehe
##### 运行结果
	Hello World
mome

Sub Macro1()
'
' Macro1 Macro
'
' Keyboard Shortcut: Ctrl+r
'
    With Selection.Font
        .Color = -16776961
        .TintAndShade = 0
    End With
End Sub

Sub Macro2()
'
' Macro2 Macro
' 赤枠
'
' Keyboard Shortcut: Ctrl+e
'
    ActiveSheet.Shapes.AddShape(msoShapeRectangle, ActiveCell.Left, ActiveCell.Top, 95.2940944882, 40.5881889764).Select
    Selection.ShapeRange.Fill.Visible = msoFalse
    With Selection.ShapeRange.Line
        .Visible = msoTrue
        .ForeColor.RGB = RGB(255, 0, 0)
        .Transparency = 0
    End With
End Sub
Sub Macro3()
'
' Macro3 Macro
'
' Keyboard Shortcut: Ctrl+y
'
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End Sub

Sub Macro4()
  Dim tshape As Shape
    
    With ActiveSheet.Range("D14:E20")
        Set tshape = ActiveSheet.Shapes.AddShape(msoShapeDownArrow, ActiveCell.Left, ActiveCell.Top, 95.2940944882, 50.5881889764)
        tshape.Name = "シェイプDownArrow"
    End With
    Set tshape = Nothing
End Sub
