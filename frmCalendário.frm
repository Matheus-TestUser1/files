VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCalendário 
   Caption         =   "Calendário"
   ClientHeight    =   2730
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   2820
   OleObjectBlob   =   "frmCalendário.frx":0000
   StartUpPosition =   1  'CenterOwner
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
End
Attribute VB_Name = "frmCalendário"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit
Dim vDateSelectedVar As Date

Public Property Get SelectDate() As Date
    SelectDate = vDateSelectedVar
End Property
Private Sub UserForm_Initialize()
       lblHoje = "Hoje: " & Format(Date, "dd/mm/yyyy")
    sb = Year(Date) * 12 + Month(Date)
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
   
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        Hide
    End If
End Sub
Private Sub lblHoje_Click()
       sb = Year(Date) * 12 + Month(Date)
End Sub

Private Sub sb_Change()
       Atualizar DateSerial(sb \ 12, sb Mod 12, 1)
End Sub

Private Sub Atualizar(dt As Date)
      
    Dim L As Long
    Dim C As Long
    Dim cInício As Long
    Dim dtDia As Date
    Dim ctrl As control
    
    lblMêsAno = Format(dt, "mmmm yyyy")
    
    For L = 1 To 6
        For C = 1 To 7
            Set ctrl = Controls("l" & L & "c" & C)
            
            dtDia = DateSerial(Year(dt), Month(dt), (L - 1) * 7 + C - Weekday(dt) + 1)
            ctrl.Caption = Format(Day(dtDia), "00")
            ctrl.Tag = dtDia
           
            If Month(dtDia) <> Month(dt) Then
                ctrl.ForeColor = &H808080
            Else
                ctrl.ForeColor = &H0
            End If
            
            If dtDia = Date Then
                ctrl.ForeColor = &HFF&
            End If
        Next C
    Next L

End Sub
