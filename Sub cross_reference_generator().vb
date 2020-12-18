Sub cross_reference_generator()
'
' cross_reference_generator Macro
' source: https://stackoverflow.com/questions/47559316/macro-to-insert-a-cross-reference-based-on-selection?rq=1
'
    Dim RefList As Variant
    Dim Ref As String
    Dim i As Integer

    With ActiveDocument
        .ActiveWindow.View.ShowFieldCodes = True
        Selection.HomeKey Unit:=wdStory
        RefList = .GetCrossReferenceItems(wdRefTypeNumberedItem)
        For i = UBound(RefList) To 1 Step -1
            Selection.HomeKey Unit:=wdStory
            Ref = Trim(RefList(i))
            RefNum = Split(Ref, " ")(0)

            ' loop through document, ref: https://stackoverflow.com/questions/43284752/vba-word-i-would-like-to-find-a-phrase-select-the-words-before-it-and-italici/43289801
            With Selection.Find

                 .Forward = True
                 .Wrap = wdFindStop
                 .Text = RefNum
                ' Fix from Tim Williams https://stackoverflow.com/questions/65335123/selection-find-through-the-document-doesnt-match-the-text-its-supposed-to-find
                Do While Selection.Find.Execute = True
                    Selection.InsertCrossReference ReferenceType:="Numbered item", _
                                                   ReferenceKind:=wdNumberFullContext, _
                                                   ReferenceItem:=CStr(i), _
                                                   InsertAsHyperlink:=True, _
                                                   IncludePosition:=False, _
                                                   SeparateNumbers:=False, _
                                                   SeparatorString:=" "


                Loop

            End With
        Next i

        .ActiveWindow.View.ShowFieldCodes = False
    End With

End Sub
