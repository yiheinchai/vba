Sub CollapseNonManagementHeading3s()

    Dim para As Paragraph
    Dim headingText As String

    ' Turn off screen updating to speed up the process
    Application.ScreenUpdating = False

    ' Loop through all paragraphs in the active document
    For Each para In ActiveDocument.Paragraphs
        ' Check if the paragraph's style is "Heading 3"
        ' Using wdStyleHeading3 constant is more robust than "Heading 3" string
        ' as it's language-independent.
        If para.Style = ActiveDocument.Styles(wdStyleHeading3) Then
            ' Get the text of the heading, remove the trailing paragraph mark
            headingText = para.Range.Text
            If Len(headingText) > 0 Then ' Check if there's any text
                If Right(headingText, 1) = vbCr Then
                    headingText = Left(headingText, Len(headingText) - 1)
                End If
            End If
            ' Trim any leading/trailing spaces
            headingText = Trim(headingText)

            ' Compare the heading text (case-insensitive)
            If LCase(headingText) = "management" Then
                ' If it IS "Management", ensure it's EXPANDED
                para.CollapsedState = False
            Else
                ' If it's any OTHER "Heading 3", COLLAPSE it
                para.CollapsedState = True
            End If
        End If
    Next para

    ' Restore screen updating
    Application.ScreenUpdating = True

    MsgBox "Processing complete. Non-'Management' Heading 3 sections have been collapsed.", vbInformation

End Sub
