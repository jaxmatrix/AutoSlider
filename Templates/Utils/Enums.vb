Imports System.Diagnostics

Namespace SlideTemplates
    Module Enums
        Enum SlideComponents
            TextHeader
            TextSubHeader
            TextHighlight
            Text
            PointsOrdered
            PointsUnordered
            PointsSymbol
            Image
            Cosmetic
        End Enum

        Enum LayoutComponents
            Title
            Description
            Points
            Image
            Cosmetic
        End Enum
    End Module

    Namespace Processor
        Module General
            Function TextToEnum(Of T As Structure)(enumString As String) As T
                Try
                    ' Ensure the type is an Enum
                    If Not GetType(T).IsEnum Then
                        Throw New ArgumentException("T must be an enumerated type")
                    End If

                    ' Parse the string into the corresponding enum value
                    Dim enumValue As T = CType([Enum].Parse(GetType(T), enumString, True), T)
                    Return enumValue
                Catch ex As Exception
                    Throw New ArgumentException($"Invalid enum string: '{enumString}' for enum type '{GetType(T).Name}'", ex)
                End Try
            End Function
        End Module

        Module SlideComponents
            Function EnumToString(componentEnum As Enums.SlideComponents)
                Return componentEnum.ToString()
            End Function

            Function StringToEnum(textHint As String)
                Return CType([Enum].Parse(GetType(Enums.SlideComponents), textHint), Enums.SlideComponents)
            End Function

        End Module

        Module LayoutComponents
            Function EnumToString(componentEnum As Enums.LayoutComponents)
                Return componentEnum.ToString()
            End Function

            Function StringToEnum(textHint As String)
                Return CType([Enum].Parse(GetType(Enums.LayoutComponents), textHint), Enums.SlideComponents)
            End Function

            Function GetContentRequriement(layoutEnum As Enums.LayoutComponents)
                Select Case layoutEnum
                    Case Enums.LayoutComponents.Title
                        Return 1
                    Case Enums.LayoutComponents.Description
                        Return 1
                    Case Enums.LayoutComponents.Points
                        Return 5
                    Case Else
                        Return 0
                End Select
            End Function
            Function GetContentRequriement(layoutEnum As String)
                Dim enumValue As Enums.LayoutComponents = Processor.LayoutComponents.StringToEnum(layoutEnum)
                Return GetContentRequriement(enumValue)
            End Function
        End Module
    End Namespace

End Namespace
