' The following statement enables all SolidEdgeCommunity extension methods.
Imports SolidEdgeCommunity.Extensions

Module Module1

    Sub Main()
        Dim objApplication As SolidEdgeFramework.Application

        SolidEdgeCommunity.Extensions.ApplicationExtensions.StartCommand(objApplication, SolidEdgeConstants.PartCommandConstants.PartEditCopy)
    End Sub

End Module
