VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Logon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

    Dim DefaultTime     As Date

Private Sub Form_Error(DataErr As Integer, Response As Integer)
  
    Const TimeHourMaximum   As Integer = 24
    Const TimeHourDefault   As Integer = 20
    Const TimeMinuteTenMax  As Integer = 5
    
    Dim ctl                 As Control
    
    Dim Text                As String
    Dim SelStart            As Integer
    
    On Error Resume Next
    
    Set ctl = Screen.ActiveControl
    
    Select Case ctl.Name
        Case "Logon"
            Text = ctl.Text
            SelStart = ctl.SelStart
            If Not IsDate(Text) Then
                DoCmd.Beep
                If Val(Left(Text, 2)) > TimeHourMaximum Then
                    Mid(Text, 1) = CStr(TimeHourDefault)
                ElseIf Len(Text) > 3 Then
                    ' Length of Text is larger than two hour digits and the kolon.
                    Mid(Text, 1 + 3) = CStr(TimeMinuteTenMax)
                End If
            End If
            ctl.Text = Text
            ctl.SelStart = SelStart
            ctl.SelLength = 1
            Response = acDataErrContinue
    End Select

    Set ctl = Nothing

End Sub

Private Sub Form_Load()
    
    Const InitialTime   As Date = #6:00:00 AM#
    
    Me!Logon.ShowDatePicker = False
    Me!Logon.InputMask = "90:00;1;0"
    Me!Logon.Format = "hh:nn"
    SetDefaultTime InitialTime

End Sub

Private Sub Logon_AfterUpdate()

    With Me!Logon
        If IsNull(.Value) Then
            ' Rem this line out to allow the textbox to be cleared.
            .Value = #12:00:00 AM#
        Else
            SetDefaultTime DateAdd("n", 1, .Value)
        End If
    End With
    
End Sub

Private Sub Logon_Click()

    With Me!Logon
        If .SelStart = 2 Then
            ' Move the cursor off the separator (colon)
            ' to the first digit of minutes.
            .SelStart = 3
            .SelLength = 1
        End If
    End With

End Sub

Private Sub Logon_Enter()
  
    With Me!Logon
        If IsNull(.Value) Then
            .Value = DefaultTime
        End If
    End With

End Sub

Private Sub Logon_KeyPress(KeyAscii As Integer)
  
    Const AsciiZero As Integer = 48
    
    Dim Text        As String
    Dim Char        As String
    Dim Length      As Integer
    Dim SelStart    As Integer
    
    With Me!Logon
        Select Case KeyAscii
            Case 8, 9, 10, 13, 27, 127
                ' Allow navigation etc. with
                ' Tab, (Ctrl+)BackSpace, (Ctrl+)Enter, and Escape.
            Case Is > 0
                If Not IsNumeric(Chr(KeyAscii)) Then
                    ' Replace any invalid entry with a zero.
                    KeyAscii = AsciiZero
                End If
            
                Text = .Text
                Length = Len(Text)
                SelStart = .SelStart
            
                Char = Mid(Text, 1 + SelStart, 1)
                If SelStart < Length Then
                    If KeyAscii <= AsciiZero + 2 Then
                        ' Always accept 0, 1, 2.
                    Else
                        ' Check if the text will represent a valid time.
                        ' If not, restore the overwritten digit.
                        Mid(Text, 1 + SelStart, 1) = Chr(KeyAscii)
                        If Not IsDate(Text) Then
                            DoCmd.Beep
                            KeyAscii = Asc(Char)
                        End If
                    End If
                End If
        End Select
    End With

End Sub

Private Sub SetDefaultTime(ThisTime As Date)

    DefaultTime = ThisTime
    Me!Logon.DefaultValue = Format(ThisTime, "\#hh:nn:00 AM/PM\#")

End Sub
