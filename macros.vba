Public height As Integer
Public width As Integer
Public stylePicture As String
Public stylePictureText As String
Public textUnderPicture As String
Sub InitGlobalVars()
    height = 400
    width = 400
    stylePicture = "Рисунок"
    stylePictureText = "Рисунок текст"
    textUnderPicture = "Рисунок"
End Sub
' -----------------------------------------------------------------------------------------------------------------
' ---------Is exists style. Fast search----------------------------------------------------------------------------
Function StyleExists(styleName As Variant) As Boolean
    On Error Resume Next
    StyleExists = Not ActiveDocument.Styles(styleName) Is Nothing
    On Error GoTo 0
End Function
' -----------------------------------------------------------------------------------------------------------------
' -----------Is it printed-----------------------------------------------------------------------------------------
Function PrintTryAgain(isFirstTry As Boolean) As Boolean
    If Not isFirstTry Then
        MsgBox "Try again", vbInformation
    End If
    PrintTryAgain = False
End Function
' -----------------------------------------------------------------------------------------------------------------
' ---InputBox---------------------------[0] - height; [1] - width--------------------------------------------------
Function GetSizeImage() As Variant
    Dim parts() As String
    
    Dim isFirstTry As Boolean
    isFirstTry = True
    Do
        isFirstTry = PrintTryAgain(isFirstTry)
        
        InputHW = InputBox("Enter {width}:", "width", width)
    Loop Until (Len(InputHW) > 0)

    GetSizeImage = Array(Trim(InputHW), Trim(InputHW))
End Function
' -----------------------------------------------------------------------------------------------------------------
' -----------------------------------------------------------------------------------------------------------------
Function isInitSizeImage(sizeImage As Variant) As Boolean
    If Not (IsNumeric(sizeImage(0)) And IsNumeric(sizeImage(1))) Then
        If sizeImage(0) = "" And sizeImage(1) = "" Then
            isInitSizeImage False
            Exit Function
        End If
        sizeImage(0) = height
        sizeImage(1) = width
        MsgBox "Error, but height & width: " & height & ";" & width
    End If
    isInitSizeImage = True
End Function
' -----------------------------------------------------------------------------------------------------------------
' -InputBox styles, text--[0] - stylePicture; [1] - stylePictureText; [2] - textUnderPicture-----------------------
Function GetStylesAndText() As Variant
    Dim parts() As String
    
    Dim isFirstTry As Boolean
    isFirstTry = True
    Do
        isFirstTry = PrintTryAgain(isFirstTry)

        p_pt_ut = InputBox("Enter {stylePicture};{stylePictureText};{textUnderPicture}:", "stylePicture;stylePictureText;textUnderPicture", stylePicture & ";" & stylePictureText & ";" & textUnderPicture)
        If p_pt_ut = "" Then
            p_pt_ut = ";;"
        End If
        parts = Split(p_pt_ut, ";")
    Loop Until (Len(p_pt_ut) > 0 And UBound(parts) = 2)

    GetStylesAndText = Array(Trim(parts(0)), Trim(parts(1)), Trim(parts(2)))
End Function
' -----------------------------------------------------------------------------------------------------------------
' -----------------------------------------------------------------------------------------------------------------
Function isInitStylesAndText(stylesAndText As Variant) As Boolean
    If Not (Len(stylesAndText(0)) > 0 And Len(stylesAndText(1)) > 0 And Len(stylesAndText(2)) > 0) Then
        If stylesAndText(0) = "" And stylesAndText(1) = "" And stylesAndText(2) = "" Then
            isInitStylesAndText = False
            Exit Function
        End If
        stylesAndText(0) = stylePicture
        stylesAndText(1) = stylePictureText
        stylesAndText(2) = textUnderPicture
        MsgBox "Error, but texts styles and underText; stylePictureText; textUnderPicture: " & stylePicture & ";" & stylePictureText & ";" & textUnderPicture
    End If
    If Not (StyleExists(stylesAndText(0)) And StyleExists(stylesAndText(1))) Then
        MsgBox "Error, Style not exists"
        isInitStylesAndText = False
        Exit Function
    End If
    isInitStylesAndText = True
End Function
' -----------------------------------------------------------------------------------------------------------------
' -----Can-cange-name-Sub------------------------------------------------------------------------------------------
Sub Select_Image_scale_text()
    InitGlobalVars
    Dim i As Long
    Dim shp As InlineShape
    Dim rng As Range
    
    ' Have Selection
    If Selection.Type = wdSelectionNormal Then
        Dim countPic As Long
        ' ------------------------------------------------------------------------------------
        sizeImage = GetSizeImage()
        stylesAndText = GetStylesAndText()
        
        If Not (isInitStylesAndText(stylesAndText) And isInitSizeImage(sizeImage)) Then
            Exit Sub
        End If
        ' ------------------------------------------------------------------------------------

        ' All picture in Selection.InlineShapes
        For Each shp In Selection.InlineShapes
            With shp
                countPic = countPic + 1

                .height = sizeImage(0)
                .width = sizeImage(1)
                
                .Range.Style = ActiveDocument.Styles(stylesAndText(0))

                Set rng = .Range
                rng.Collapse Direction:=wdCollapseEnd

                rng.InsertParagraphAfter
                rng.InsertParagraphAfter

                rng.MoveStart wdParagraph, 1

                rng.Text = stylesAndText(2) & " "

                rng.Collapse Direction:=wdCollapseEnd
                rng.Fields.Add Range:=rng, Type:=wdFieldEmpty, Text:="SEQ Рисунок \* ARABIC ", PreserveFormatting:=False

                rng.Paragraphs(1).Range.Style = ActiveDocument.Styles(stylesAndText(1))
            End With
        Next shp
        
        If countPic > 0 Then
            MsgBox "Count pictures: " & countPic & vbCrLf & "width: " & sizeImage(1)
        Else
            MsgBox "Pictures were not found"
        End If
    Else
        MsgBox "Highlight the marks of the paragraph!", vbInformation
    End If
End Sub

