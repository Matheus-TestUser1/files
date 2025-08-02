Attribute VB_Name = "mdlCalendário"
Option Explicit


Dim Rótulos() As New cCalendário

Function GetCalendário() As Date
        
          
    Dim lTotalRótulos As Long
    Dim ctrl As control
    Dim FRM As frmCalendário
    
    Set FRM = New frmCalendário
    
       For Each ctrl In FRM.Controls
        If ctrl.Name Like "l?c?" Then
            lTotalRótulos = lTotalRótulos + 1
            ReDim Preserve Rótulos(1 To lTotalRótulos)
            Set Rótulos(lTotalRótulos).lblGrupo = ctrl
        End If
    Next ctrl

    FRM.Show
    
        If IsDate(FRM.Tag) Then
        GetCalendário = FRM.Tag
    Else
        GetCalendário = Date
    End If
        
    Unload FRM



End Function
