Namespace ImageFileRename

    Public Class FileNameStyle

        Public Shared Function FileNameConforms(ByVal strFileName As String) As Boolean
            Return True
        End Function

        Public Shared Function FileTitle(ByVal strFileName As String) As String
            Dim strToken() As String
            Dim strExtension As String
            Dim strReturn As String

            strReturn = FileName(strFileName)
            strToken = strReturn.Split(".")
            strExtension = strToken(UBound(strToken))
            strReturn = Left(strReturn, Len(strReturn) - (Len(strExtension) + 1))

            Return strReturn

        End Function

        Public Overridable Function FileDate(ByVal strFileName As String) As Date
            Return Date.MinValue
        End Function

        Public Overridable Function FileNumber(ByVal strFileName As String) As String
            Return ""
        End Function

        Public Overridable Function FileDescription(ByVal strFileName As String) As String
            Dim strReturn As String

            strReturn = FileTitle(strFileName)

            Return strReturn
        End Function

        Public Overridable Function Extension(ByVal strFileName As String) As String

            Return modCommonFunctions.Extension(strFileName)

        End Function

    End Class

    Public Class Filename_Canon
        Inherits ImageFileRename.FileNameStyle

        Public Overloads Shared Function FileNameConforms(ByVal strFileName As String) As Boolean
            Dim blnReturn As Boolean
            Dim strFileTitle As String

            blnReturn = False

            strFileTitle = FileTitle(strFileName)

            Select Case True
                Case Len(strFileTitle) <> 8
                Case Not (Left(strFileTitle, 4) = "IMG_" Or Left(strFileTitle, 2) = "ST")
                Case Not IsNumeric(Right(strFileTitle, 4))
                Case Else
                    blnReturn = True
            End Select

            Return blnReturn

        End Function

        Public Overrides Function FileDescription(ByVal strFileName As String) As String
            Dim strReturn As String

            strReturn = ""

            Return strReturn

        End Function

        Public Overrides Function FileNumber(ByVal strFileName As String) As String

            Dim strReturn As String

            strReturn = Split(FileTitle(strFileName), "_")(1)

            If Left(FileTitle(strFileName), 2) = "ST" Then
                strReturn &= Right(Left(FileTitle(strFileName), 3), 1)
            End If

            Return strReturn

        End Function
    End Class

    Public Class FileName_Minolta
        Inherits ImageFileRename.FileNameStyle

        Public Overloads Shared Function FileNameConforms(ByVal strFileName As String) As Boolean
            Dim blnReturn As Boolean
            Dim strFileTitle As String

            blnReturn = False

            strFileTitle = FileTitle(strFileName)

            Select Case True
                Case Len(strFileTitle) <> 8
                Case Left(strFileTitle, 4) <> "PICT"
                Case Not IsNumeric(Right(strFileTitle, 4))
                Case Else
                    blnReturn = True
            End Select

            Return blnReturn

        End Function

        Public Overrides Function FileNumber(ByVal strFileName As String) As String

            Return Right(FileTitle(strFileName), 4)

        End Function

    End Class

    Public Class FileName_Sony
        Inherits ImageFileRename.FileNameStyle

        Public Overloads Shared Function FileNameConforms(ByVal strFileName As String) As Boolean
            Dim blnReturn As Boolean
            Dim strFileTitle As String

            blnReturn = False

            strFileTitle = FileTitle(strFileName)

            Select Case True
                Case Len(strFileTitle) <> 8
                Case Left(strFileTitle, 3) <> "DSC"
                Case Not IsNumeric(Right(strFileTitle, 4))
                Case Else
                    blnReturn = True
            End Select

            Return blnReturn

        End Function

        Public Overrides Function FileNumber(ByVal strFileName As String) As String

            Return Right(FileTitle(strFileName), 4)

        End Function
    End Class

    Public Class FileName_Konika
        Inherits ImageFileRename.FileNameStyle

        Public Overloads Shared Function FileNameConforms(ByVal strFileName As String) As Boolean
            Dim blnReturn As Boolean
            Dim strFileTitle As String

            blnReturn = False

            strFileTitle = FileTitle(strFileName)

            Select Case True
                Case Len(strFileTitle) <> 8
                Case Left(strFileTitle, 1) <> "P"
                Case Not IsNumeric(Right(strFileTitle, 7))
                Case Else
                    blnReturn = True
            End Select

            Return blnReturn

        End Function

        Public Overrides Function FileDescription(ByVal strFileName As String) As String
            Dim strReturn As String

            strReturn = ""

            Return strReturn

        End Function

        Public Overrides Function FileNumber(ByVal strFileName As String) As String

            Return Right(FileTitle(strFileName), 4)

        End Function

    End Class

    Public Class FileName_Film_Frame
        Inherits ImageFileRename.FileNameStyle

        Public Overloads Shared Function FileNameConforms(ByVal strFileName As String) As Boolean
            Dim blnReturn As Boolean
            Dim strFileTitle As String

            strFileTitle = FileTitle(strFileName)

            blnReturn = False

            Select Case True
                Case Len(strFileTitle) < 6
                Case UBound(Split(strFileTitle, "_")) <> 1
                Case Not IsNumeric(Split(strFileTitle, "_")(0))
                Case Not IsNumeric(Split(strFileTitle, "_")(1))
                Case Else
                    blnReturn = True
            End Select

            Return blnReturn

        End Function

        Public Overrides Function FileDescription(ByVal strFileName As String) As String
            Dim strReturn As String

            strReturn = ""

            Return strReturn

        End Function

        Public Overrides Function FileNumber(ByVal strFileName As String) As String
            Dim strReturn As String

            strReturn = FileTitle(strFileName)

            Return strReturn

        End Function

    End Class

    Public Class FileName_Film_FrameDescr
        Inherits ImageFileRename.FileNameStyle

        Public Overloads Shared Function FileNameConforms(ByVal strFileName As String) As Boolean
            Dim blnReturn As Boolean
            Dim strFileTitle As String

            strFileTitle = FileTitle(strFileName)

            blnReturn = False

            Select Case True
                Case Len(strFileTitle) < 6
                Case UBound(Split(strFileTitle, "_")) <> 1
                Case Not IsNumeric(Split(strFileTitle, "_")(0))
                Case IsNumeric(Split(strFileTitle, "_")(1))
                Case Not IsNumeric(Left(Split(strFileTitle, "_")(1), 1))
                Case Else
                    blnReturn = True
            End Select

            Return blnReturn

        End Function

        Public Overrides Function FileDescription(ByVal strFileName As String) As String
            Dim strReturn As String
            Dim strTest As String
            Dim intCount As Integer

            strReturn = FileTitle(strFileName)
            strReturn = Split(strReturn, "_")(1)
            For intCount = 1 To Len(strReturn)
                strTest = Left(strReturn, intCount)
                If Not IsNumeric(strTest) Then
                    strReturn = Right(strReturn, Len(strReturn) - (intCount - 1))
                    Exit For
                End If
            Next

            Return strReturn

        End Function

        Public Overrides Function FileNumber(ByVal strFileName As String) As String
            Dim strReturn As String

            strReturn = FileTitle(strFileName)
            strReturn = Left(strReturn, Len(strReturn) - Len(FileDescription(strFileName)))

            Return strReturn

        End Function

    End Class

    Public Class FileName_Film_Frame_Descr
        Inherits ImageFileRename.FileNameStyle

        Public Overloads Shared Function FileNameConforms(ByVal strFileName As String) As Boolean
            Dim blnReturn As Boolean
            Dim strFileTitle As String

            strFileTitle = FileTitle(strFileName)

            blnReturn = False

            Select Case True
                Case Len(strFileTitle) < 6
                Case UBound(Split(strFileTitle, "_")) <> 2
                Case Not IsNumeric(Split(strFileTitle, "_")(0))
                Case Not Len(Split(strFileTitle, "_")(0)) = 4
                Case Not IsNumeric(Split(strFileTitle, "_")(1))
                Case Else
                    blnReturn = True
            End Select

            Return blnReturn

        End Function

        Public Overrides Function FileDescription(ByVal strFileName As String) As String
            Dim strReturn As String

            strReturn = FileTitle(strFileName)
            strReturn = Split(strReturn, "_")(2)

            Return strReturn

        End Function

        Public Overrides Function FileNumber(ByVal strFileName As String) As String
            Dim strReturn As String

            strReturn = FileTitle(strFileName)
            strReturn = Split(strReturn, "_")(0) & "_" & Split(strReturn, "_")(1)

            Return strReturn

        End Function

    End Class

    Public Class FileName_Date_Film_Frame_Descr
        Inherits ImageFileRename.FileNameStyle

        Public Overloads Shared Function FileNameConforms(ByVal strFileName As String) As Boolean
            Dim blnReturn As Boolean
            Dim strFileTitle As String

            strFileTitle = FileTitle(strFileName)

            blnReturn = False

            Select Case True
                Case Len(strFileTitle) < 6
                Case UBound(Split(strFileTitle, "_")) <> 3
                Case Not IsNumeric(Split(strFileTitle, "_")(0))
                Case Not Len(Split(strFileTitle, "_")(0)) = 8
                Case Not IsNumeric(Split(strFileTitle, "_")(1))
                Case Not IsNumeric(Split(strFileTitle, "_")(2))
                Case Else
                    blnReturn = True
            End Select

            Return blnReturn

        End Function

        Public Overrides Function FileDate(ByVal strFileName As String) As Date
            Dim dteReturn As Date
            Dim strYear As String
            Dim strMonth As String
            Dim strDay As String
            Dim strFileTitle As String

            dteReturn = Date.MinValue

            strFileTitle = FileTitle(strFileName)

            Try
                strYear = Left(strFileTitle, 4)
                strMonth = Right(Left(strFileTitle, 6), 2)
                strDay = Right(Left(strFileTitle, 8), 2)

                dteReturn.AddYears(CInt(strYear) - 1)
                If Not strMonth = "00" Then
                    dteReturn.AddMonths(CInt(strMonth - 1))
                End If
                If Not strDay = "00" Then
                    dteReturn.AddDays(CInt(strDay - 1))
                End If
            Catch ex As Exception
                dteReturn = Date.MinValue
            End Try

            Return dteReturn

        End Function

        Public Overrides Function FileDescription(ByVal strFileName As String) As String
            Dim strReturn As String

            strReturn = FileTitle(strFileName)
            strReturn = Split(strReturn, "_")(3)

            Return strReturn

        End Function

        Public Overrides Function FileNumber(ByVal strFileName As String) As String
            Dim strReturn As String

            strReturn = FileTitle(strFileName)
            strReturn = Split(strReturn, "_")(1) & "_" & Split(strReturn, "_")(2)

            Return strReturn

        End Function

    End Class

    Public Class FileName_Date_ImageNumber_Descr
        Inherits ImageFileRename.FileNameStyle

        Public Overloads Shared Function FileNameConforms(ByVal strFileName As String) As Boolean
            Dim blnReturn As Boolean
            Dim strFileTitle As String

            strFileTitle = FileTitle(strFileName)

            blnReturn = False

            Select Case True
                Case Len(strFileTitle) < 6
                Case UBound(Split(strFileTitle, "_")) <> 2
                Case Not IsNumeric(Split(strFileTitle, "_")(0))
                Case Not Len(Split(strFileTitle, "_")(0)) = 8
                Case Not IsNumeric(Split(strFileTitle, "_")(1))
                Case Else
                    blnReturn = True
            End Select

            Return blnReturn

        End Function

        Public Overrides Function FileDate(ByVal strFileName As String) As Date
            Dim dteReturn As Date
            Dim strYear As String
            Dim strMonth As String
            Dim strDay As String
            Dim strFileTitle As String

            dteReturn = Date.MinValue

            strFileTitle = FileTitle(strFileName)

            Try
                strYear = Left(strFileTitle, 4)
                strMonth = Right(Left(strFileTitle, 6), 2)
                strDay = Right(Left(strFileTitle, 8), 2)

                dteReturn.AddYears(CInt(strYear) - 1)
                If Not strMonth = "00" Then
                    dteReturn.AddMonths(CInt(strMonth - 1))
                End If
                If Not strDay = "00" Then
                    dteReturn.AddDays(CInt(strDay - 1))
                End If
            Catch ex As Exception
                dteReturn = Date.MinValue
            End Try

            Return dteReturn

        End Function

        Public Overrides Function FileDescription(ByVal strFileName As String) As String
            Dim strReturn As String

            strReturn = FileTitle(strFileName)
            strReturn = Split(strReturn, "_")(2)

            Return strReturn

        End Function

        Public Overrides Function FileNumber(ByVal strFileName As String) As String
            Dim strReturn As String

            strReturn = FileTitle(strFileName)
            strReturn = Split(strReturn, "_")(1)

            Return strReturn

        End Function

    End Class

    Public Class FileName_Date_Time_ImageNumber
        Inherits ImageFileRename.FileNameStyle

        Public Overloads Shared Function FileNameConforms(ByVal strFileName As String) As Boolean
            Dim blnReturn As Boolean
            Dim strFileTitle As String

            strFileTitle = FileTitle(strFileName)

            blnReturn = False

            Select Case True
                Case Len(strFileTitle) < 21
                Case UBound(Split(strFileTitle, "_")) <> 2
                Case Not UBound(Split(Split(strFileTitle, "_")(0), "-")) <> 3
                Case Not UBound(Split(Split(strFileTitle, "_")(1), "-")) <> 3
                Case Not IsNumeric(Split(strFileTitle, "_")(2))
                Case Else
                    blnReturn = True
            End Select

            Return blnReturn

        End Function

        Public Overrides Function FileDate(ByVal strFileName As String) As Date
            Dim dteReturn As Date
            Dim strYear As String
            Dim intYear As Integer
            Dim strMonth As String
            Dim strDay As String
            Dim strFileTitle As String

            dteReturn = Date.MinValue

            strFileTitle = FileTitle(strFileName)

            Try
                strYear = strFileTitle.Substring(6, 2)
                strMonth = strFileTitle.Substring(3, 2)
                strDay = Left(strFileTitle, 2)

                intYear = CInt(strYear)
                If intYear > 30 Then
                    intYear += 1900
                Else
                    intYear += 2000
                End If
                dteReturn.AddYears(strYear)

                If Not strMonth = "00" Then
                    dteReturn.AddMonths(CInt(strMonth - 1))
                End If
                If Not strDay = "00" Then
                    dteReturn.AddDays(CInt(strDay - 1))
                End If
            Catch ex As Exception
                dteReturn = Date.MinValue
            End Try

            Return dteReturn

        End Function

        Public Overrides Function FileDescription(ByVal strFileName As String) As String
            
            Return ""

        End Function

        Public Overrides Function FileNumber(ByVal strFileName As String) As String
            Dim strReturn As String

            strReturn = FileTitle(strFileName)
            strReturn = Split(strReturn, "_")(2)

            Return strReturn

        End Function

    End Class

    Public Class Filename_Olympus
        Inherits ImageFileRename.FileNameStyle

        Public Overloads Shared Function FileNameConforms(ByVal strFileName As String) As Boolean
            Dim blnReturn As Boolean
            Dim strFileTitle As String

            blnReturn = False

            strFileTitle = FileTitle(strFileName)

            Select Case True
                Case Len(strFileTitle) <> 8
                Case Not (Left(strFileTitle, 2) = "PC")
                Case Not IsNumeric(Right(strFileTitle, 6))
                Case Else
                    blnReturn = True
            End Select

            Return blnReturn

        End Function

        Public Overrides Function FileNumber(ByVal strFileName As String) As String

            Dim strReturn As String

            strReturn = Right(FileTitle(strFileName), 4)

            Return strReturn

        End Function
    End Class

    Public Class Filename_Nokia
        Inherits ImageFileRename.FileNameStyle

        Public Overloads Shared Function FileNameConforms(ByVal strFileName As String) As Boolean
            Dim blnReturn As Boolean
            Dim strFileTitle As String

            blnReturn = False

            strFileTitle = FileTitle(strFileName)

            Select Case True
                Case Len(strFileTitle) <> 9
                Case Not (Left(strFileTitle, 5) = "Image")
                Case Not IsNumeric(Right(strFileTitle, 4))
                Case Else
                    blnReturn = True
            End Select

            Return blnReturn

        End Function

        Public Overrides Function FileNumber(ByVal strFileName As String) As String

            Dim strReturn As String

            strReturn = Right(strFileName, 4)

            Return strReturn

        End Function
    End Class

    Public Class FileName_ImageNumber
        Inherits ImageFileRename.FileNameStyle

        Public Overloads Shared Function FileNameConforms(ByVal strFileName As String) As Boolean
            Dim blnReturn As Boolean
            Dim strFileTitle As String

            strFileTitle = FileTitle(strFileName)

                blnreturn= IsNumeric(strFileTitle)

            Return blnReturn

        End Function

        Public Overrides Function FileNumber(ByVal strFileName As String) As String
            Dim strReturn As String

            strReturn = FileTitle(strFileName)

            Return strReturn

        End Function

        Public Overrides Function FileDescription(ByVal strFileName As String) As String
            Dim strReturn As String

            strReturn = ""

            Return strReturn

        End Function

    End Class

End Namespace