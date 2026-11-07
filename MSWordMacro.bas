
Sub ConvertBanglaEnglishFonts()

    Dim doc As Document
    Dim ch As Range
    Dim i As Long
    Dim code As Long

    Set doc = ActiveDocument

    Application.ScreenUpdating = False

    For i = 1 To doc.Characters.Count

        Set ch = doc.Characters(i)

        If Len(ch.Text) > 0 Then

            code = AscW(Left(ch.Text, 1))

            If code < 0 Then code = code + 65536

            Select Case code

                ' Bangla Unicode characters
                Case &H980 To &H9FF

                    With ch.Font
                        .Name = "NikoshBAN"
                        .NameBi = "NikoshBAN"
                    End With

                ' Bangla danda (?)
                Case &H964

                    With ch.Font
                        .Name = "NikoshBAN"
                        .NameBi = "NikoshBAN"
                    End With

                ' English letters, numbers and punctuation
                Case 0 To 127

                    With ch.Font
                        .Name = "Times New Roman"
                        .NameAscii = "Times New Roman"
                    End With

            End Select

        End If

    Next i

    Application.ScreenUpdating = True

    MsgBox "Font conversion completed successfully!", vbInformation

End Sub

