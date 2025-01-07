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
        End Module
    End Namespace

End Namespace
