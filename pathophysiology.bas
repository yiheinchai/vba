Sub DeletePathophysiologyBlock()
    ' Macro to find a top-level bullet starting with "Pathophysiology"
    ' and delete it along with all its sub-bullets.
    ' Looping backwards is safer when deleting items.

    Dim i As Long
    Dim targetPara As Paragraph
    Dim nextPara As Paragraph
    Dim blockStart As Long
    Dim blockEnd As Long
    Dim FoundBlock As Boolean
    Dim targetText As String
    Dim targetLevel As Integer

    ' --- Configuration ---
    targetText = "pathophysiology" ' Text to find (lowercase for case-insensitive search)
    targetLevel = 1               ' The list level to search for (1 = top level)
    ' --- End Configuration ---

    Application.ScreenUpdating = False ' Speed up macro and prevent screen flicker

    ' Loop backwards through all paragraphs in the active document
    For i = ActiveDocument.Paragraphs.Count To 1 Step -1
        Set targetPara = ActiveDocument.Paragraphs(i)
        FoundBlock = False ' Reset flag for each potential target

        ' Check if the paragraph is part of a list AND is at the target level
        If targetPara.Range.ListFormat.ListType <> wdListNoNumbering And _
           targetPara.Range.ListFormat.ListLevelNumber = targetLevel Then

            ' Check if the paragraph text (trimmed, lowercased) starts with the target text.
            ' Trim removes leading/trailing spaces and the paragraph mark itself.
            If LCase(Trim(targetPara.Range.Text)) Like targetText & "*" Then
                FoundBlock = True
            End If

            If FoundBlock Then
                ' Found the target top-level paragraph. Now define the range to delete.
                blockStart = targetPara.Range.Start
                blockEnd = targetPara.Range.End ' Initialize end with the target para itself

                ' Look ahead (downwards in the document) to find subsequent paragraphs
                ' that are sub-items (higher list level)
                Dim j As Long
                For j = i + 1 To ActiveDocument.Paragraphs.Count
                    ' Defensive check in case paragraph index becomes invalid after deletion
                    On Error Resume Next
                    Set nextPara = ActiveDocument.Paragraphs(j)
                    If Err.Number <> 0 Then
                        On Error GoTo 0 ' Reset error handling
                        Exit For ' Stop looking ahead if paragraph doesn't exist
                    End If
                    On Error GoTo 0 ' Reset error handling

                    ' Check if the next paragraph is a sub-item (higher level)
                    If nextPara.Range.ListFormat.ListType <> wdListNoNumbering And _
                       nextPara.Range.ListFormat.ListLevelNumber > targetLevel Then
                        ' It's a sub-bullet belonging to the block, extend the deletion range
                        blockEnd = nextPara.Range.End
                    Else
                        ' It's NOT a sub-bullet (it's another level 1, or not a list item, or end of doc)
                        ' The block ends *before* this paragraph 'j'.
                        Exit For ' Stop extending the block
                    End If
                Next j ' Check next paragraph downwards

                ' Define the complete range and delete it
                If blockStart < blockEnd Then ' Ensure range is valid
                    Dim blockToDelete As Range
                    Set blockToDelete = ActiveDocument.Range(Start:=blockStart, End:=blockEnd)
                    blockToDelete.Delete

                    ' Optional: If you only want to delete the *first* matching block found
                    ' when searching from the end of the document, uncomment the next line:
                    ' Exit For ' Exits the main loop (i) after the first deletion

                End If ' End If blockStart < blockEnd

            End If ' End If FoundBlock
        End If ' End If Is targetLevel List Item
    Next i ' Check previous paragraph upwards

    Application.ScreenUpdating = True ' Restore screen updates
    MsgBox "Deletion process complete. Check document.", vbInformation

End SubSub DeletePathophysiologyBlock()
    ' Macro to find a top-level bullet starting with "Pathophysiology"
    ' and delete it along with all its sub-bullets.
    ' Looping backwards is safer when deleting items.

    Dim i As Long
    Dim targetPara As Paragraph
    Dim nextPara As Paragraph
    Dim blockStart As Long
    Dim blockEnd As Long
    Dim FoundBlock As Boolean
    Dim targetText As String
    Dim targetLevel As Integer

    ' --- Configuration ---
    targetText = "pathophysiology" ' Text to find (lowercase for case-insensitive search)
    targetLevel = 1               ' The list level to search for (1 = top level)
    ' --- End Configuration ---

    Application.ScreenUpdating = False ' Speed up macro and prevent screen flicker

    ' Loop backwards through all paragraphs in the active document
    For i = ActiveDocument.Paragraphs.Count To 1 Step -1
        Set targetPara = ActiveDocument.Paragraphs(i)
        FoundBlock = False ' Reset flag for each potential target

        ' Check if the paragraph is part of a list AND is at the target level
        If targetPara.Range.ListFormat.ListType <> wdListNoNumbering And _
           targetPara.Range.ListFormat.ListLevelNumber = targetLevel Then

            ' Check if the paragraph text (trimmed, lowercased) starts with the target text.
            ' Trim removes leading/trailing spaces and the paragraph mark itself.
            If LCase(Trim(targetPara.Range.Text)) Like targetText & "*" Then
                FoundBlock = True
            End If

            If FoundBlock Then
                ' Found the target top-level paragraph. Now define the range to delete.
                blockStart = targetPara.Range.Start
                blockEnd = targetPara.Range.End ' Initialize end with the target para itself

                ' Look ahead (downwards in the document) to find subsequent paragraphs
                ' that are sub-items (higher list level)
                Dim j As Long
                For j = i + 1 To ActiveDocument.Paragraphs.Count
                    ' Defensive check in case paragraph index becomes invalid after deletion
                    On Error Resume Next
                    Set nextPara = ActiveDocument.Paragraphs(j)
                    If Err.Number <> 0 Then
                        On Error GoTo 0 ' Reset error handling
                        Exit For ' Stop looking ahead if paragraph doesn't exist
                    End If
                    On Error GoTo 0 ' Reset error handling

                    ' Check if the next paragraph is a sub-item (higher level)
                    If nextPara.Range.ListFormat.ListType <> wdListNoNumbering And _
                       nextPara.Range.ListFormat.ListLevelNumber > targetLevel Then
                        ' It's a sub-bullet belonging to the block, extend the deletion range
                        blockEnd = nextPara.Range.End
                    Else
                        ' It's NOT a sub-bullet (it's another level 1, or not a list item, or end of doc)
                        ' The block ends *before* this paragraph 'j'.
                        Exit For ' Stop extending the block
                    End If
                Next j ' Check next paragraph downwards

                ' Define the complete range and delete it
                If blockStart < blockEnd Then ' Ensure range is valid
                    Dim blockToDelete As Range
                    Set blockToDelete = ActiveDocument.Range(Start:=blockStart, End:=blockEnd)
                    blockToDelete.Delete

                    ' Optional: If you only want to delete the *first* matching block found
                    ' when searching from the end of the document, uncomment the next line:
                    ' Exit For ' Exits the main loop (i) after the first deletion

                End If ' End If blockStart < blockEnd

            End If ' End If FoundBlock
        End If ' End If Is targetLevel List Item
    Next i ' Check previous paragraph upwards

    Application.ScreenUpdating = True ' Restore screen updates
    MsgBox "Deletion process complete. Check document.", vbInformation

End Sub
