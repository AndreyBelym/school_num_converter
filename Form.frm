VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Converter"
   ClientHeight    =   6255
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5655
   Icon            =   "Form.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6255
   ScaleWidth      =   5655
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Перевод чисел из родственных систем счисления"
      Height          =   1215
      Left            =   120
      TabIndex        =   10
      Top             =   4680
      Width           =   5415
      Begin VB.CommandButton Command2 
         Caption         =   "Посчитать"
         Height          =   375
         Left            =   1800
         TabIndex        =   14
         Top             =   600
         Width           =   1575
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   4320
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   240
         Width           =   975
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   3360
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   3255
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   $"Form.frx":0442
         Height          =   495
         Left            =   4320
         TabIndex        =   17
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   $"Form.frx":0455
         Height          =   375
         Left            =   3360
         TabIndex        =   16
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Число"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   600
         Width           =   1695
      End
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Из Р в 10"
      Height          =   375
      Left            =   2280
      TabIndex        =   9
      Top             =   4320
      Width           =   1935
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Из 10 в Р"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   4320
      Value           =   -1  'True
      Width           =   2055
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   300
      Left            =   5280
      Picture         =   "Form.frx":0466
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   7
      Top             =   4320
      Width           =   300
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   0
      Top             =   120
   End
   Begin MSComctlLib.StatusBar stsStatus 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Top             =   5880
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1588
            MinWidth        =   1588
            Text            =   "00.00.00"
            TextSave        =   "00.00.00"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   2117
            MinWidth        =   2117
            TextSave        =   "27.09.2008"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   3440
            MinWidth        =   2734
            Text            =   "Переведено чисел: 0000"
            TextSave        =   "Переведено чисел: 0000"
         EndProperty
      EndProperty
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      ItemData        =   "Form.frx":0568
      Left            =   120
      List            =   "Form.frx":056A
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   3840
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   2895
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   120
      Width           =   5415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Посчитать"
      Height          =   375
      Left            =   4320
      TabIndex        =   1
      Top             =   3840
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Top             =   3840
      Width           =   2895
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Ваше число"
      Height          =   195
      Left            =   1320
      TabIndex        =   5
      Top             =   3120
      Width           =   2925
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Основание системы счисления"
      Height          =   585
      Left            =   120
      TabIndex        =   4
      Top             =   3120
      Width           =   1200
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Dim p As Byte, rc As String, oldtime As String, kol As Byte

Private Sub Combo2_Click()
Combo3.Clear
Select Case Combo2.List(Combo2.ListIndex)
Case Is = "2"
    Combo3.AddItem "4"
    Combo3.AddItem "8"
    Combo3.AddItem "16"
Case Is = 3
    Combo3.AddItem "9"
Case Is = "4"
    Combo3.AddItem "2"
    Combo3.AddItem "16"
Case Is = "8"
    Combo3.AddItem "2"
Case Is = "9"
    Combo3.AddItem "3"
Case Is = "16"
    Combo3.AddItem "2"
    Combo3.AddItem "4"
End Select
Combo3.ListIndex = 0

End Sub

Private Sub Command1_Click()
On Error GoTo 10
Text2.Text = ""
Dim chast As Double, Celoe As String, KolCel As Long, koldrob As Long, KolPer As Long, DrobChisl As String, DrobZnam As Double, PerChisl As String, PerZnam As Double
KolPer = 0: koldrob = 0
p = Val(Combo1.List(Combo1.ListIndex))
rc = Text1.Text
If Len(rc) <> 0 Then
Do
    Celoe = Celoe + Left(rc, 1)
    rc = Right(rc, Len(rc) - 1)
    KolCel = KolCel + 1
Loop Until Len(rc) = 0 Or Left(rc, 1) = "." Or Left(rc, 1) = ","
End If

If Len(rc) <> 0 Then
    rc = Right(rc, Len(rc) - 1)
    Do
        Select Case Left(rc, 1)
        Case Is = "("
            rc = Right(rc, Len(rc) - 1)
            Do
                KolPer = KolPer + 1
                PerChisl = PerChisl + Left(rc, 1)
                rc = Right(rc, Len(rc) - 1)
            Loop Until Left(rc, 1) = ")"
            rc = ""
        Case Else
            koldrob = koldrob + 1
            
            DrobChisl = DrobChisl + Left(rc, 1)
            rc = Right(rc, Len(rc) - 1)
        End Select
    Loop Until Len(rc) = 0
    Dim pDrob As Byte
    If Option1.Value = True Then
    pDrob = 10
    Else
    pDrob = p
    If koldrob <> 0 Then DrobChisl = Trim(Str(Gorner(DrobChisl, koldrob)))
    If KolPer <> 0 Then PerChisl = Trim(Str(Gorner(PerChisl, KolPer)))
    End If
    If koldrob <> 0 Then DrobZnam = pDrob ^ koldrob
    If KolPer <> 0 Then PerZnam = (pDrob ^ KolPer) - 1
    If koldrob = 0 Then
        If KolPer <> 0 Then
            
            DrobChisl = PerChisl
            DrobZnam = PerZnam
            koldrob = 1
        End If
    Else
        If KolPer <> 0 Then
            DrobChisl = Trim(Str(Val(DrobChisl) * PerZnam + Val(PerChisl)))
            DrobZnam = DrobZnam * PerZnam
            koldrob = 1
        Else
            koldrob = 1
        End If
    End If
End If
If Option1.Value = True Then
If Val(Celoe) <> 0 Then
    Dim ost As Byte
    chast = Val(Celoe)
    Do
        Select Case chast Mod p
        Case Is = 10
            Text2.Text = "A" + Text2.Text
        Case Is = 11
            Text2.Text = "B" + Text2.Text
        Case Is = 12
            Text2.Text = "C" + Text2.Text
        Case Is = 13
            Text2.Text = "D" + Text2.Text
        Case Is = 14
            Text2.Text = "E" + Text2.Text
        Case Is = 15
            Text2.Text = "F" + Text2.Text
        Case Else
            Text2.Text = Trim(Str(chast Mod p)) + Text2.Text
        End Select
        chast = chast \ p
    Loop Until chast < p
    Select Case chast
    Case Is = 10
        Text2.Text = "A" + Text2.Text
    Case Is = 11
        Text2.Text = "B" + Text2.Text
    Case Is = 12
        Text2.Text = "C" + Text2.Text
    Case Is = 13
        Text2.Text = "D" + Text2.Text
    Case Is = 14
        Text2.Text = "E" + Text2.Text
    Case Is = 15
        Text2.Text = "F" + Text2.Text
    Case Else
        If chast <> 0 Then Text2.Text = Trim(Str(chast)) + Text2.Text
    End Select
Else
    Text2.Text = "0"
End If

If koldrob <> 0 Then
 GetPeriod p, DrobZnam, Val(DrobChisl)
End If
Else
If Val(Celoe) <> 0 Then
Text2.Text = Trim(Str(Gorner(Celoe, KolCel)))
Else
Text2.Text = "0"
End If
If koldrob <> 0 Then
GetPeriod 10, DrobZnam, Val(DrobChisl)
End If
End If
kol = kol + 1
stsStatus.Panels(3).Text = "Переведено чисел:" + Str(kol)
Text1.SetFocus
GoTo 20
10
MsgBox "Error#" + Str(Err.Number) + Chr(13) + "Description: " + CStr(Err.Description) + Chr(13) + "Source: " + CStr(Err.Source), vbCritical, "Error!"
Text1.SetFocus
20 End Sub



Private Sub Command2_Click()
On Error GoTo 10
Dim ada As Byte, rct As String, Cel As String, KCel As Long, Drob As String, KDrob As Long, Per As String, KPer As Long
rct = Text3.Text
If Len(rct) <> 0 Then
Do
    Cel = Cel + Left(rct, 1)
    rct = Right(rct, Len(rct) - 1)
    KCel = KCel + 1
Loop Until Len(rct) = 0 Or Left(rct, 1) = "." Or Left(rct, 1) = ","
End If

If Len(rct) <> 0 Then
    rct = Right(rct, Len(rct) - 1)
    Do
        Select Case Left(rct, 1)
        Case Is = "("
            rct = Right(rct, Len(rct) - 1)
            Do
                KPer = KPer + 1
                Per = Per + Left(rct, 1)
                rct = Right(rct, Len(rct) - 1)
            Loop Until Left(rct, 1) = ")"
            rct = ""
        Case Else
            KDrob = KDrob + 1
            
            Drob = Drob + Left(rct, 1)
            rct = Right(rct, Len(rct) - 1)
        End Select
    Loop Until Len(rct) = 0
End If
ada = 1
If Val(Combo2.List(Combo2.ListIndex)) < Val(Combo3.List(Combo3.ListIndex)) Then
Select Case Combo2.List(Combo2.ListIndex)
    Case Is = "2"
        Select Case Combo3.List(Combo3.ListIndex)
        Case Is = "4"
        ada = 2
        Case Is = "8"
        ada = 3
        Case Is = "16"
        ada = 4
        End Select
    Case Is = "3"
        ada = 2
    Case Is = "4"
        ada = 2
End Select
If Cel <> "" Then
If KCel Mod ada <> 0 Then
For i = 1 To ada - KCel Mod ada
Cel = "0" + Cel
Next
KCel = KCel + (ada - KCel Mod ada)
End If
End If
End If
If Cel <> "" Then
Cel = GetRod(Val(Combo2.List(Combo2.ListIndex)), Val(Combo3.List(Combo3.ListIndex)), Cel, KCel, ada)
Do Until Left(Cel, 1) <> "0" Or Len(Cel) = 1
    Cel = Right(Cel, Len(Cel) - 1)
Loop
Text2.Text = Cel
End If
If Drob <> "" Then
    Text2.Text = Text2.Text + "."
    If Val(Combo2.List(Combo2.ListIndex)) <> 16 Then
    
    If KDrob Mod ada = 0 Then
        Drob = GetRod(Val(Combo2.List(Combo2.ListIndex)), Val(Combo3.List(Combo3.ListIndex)), Drob, KDrob, ada)
    Else
        If Per <> "" Then
            Do
            Drob = Drob + Per
            KDrob = KDrob + KPer
            Loop Until KDrob Mod ada = 0
            Drob = GetRod(Val(Combo2.List(Combo2.ListIndex)), Val(Combo3.List(Combo3.ListIndex)), Drob, KDrob, ada)
        Else
            Do
            Drob = Drob + "0"
            KDrob = KDrob + 1
            Loop Until KDrob Mod ada = 0
            Drob = GetRod(Val(Combo2.List(Combo2.ListIndex)), Val(Combo3.List(Combo3.ListIndex)), Drob, KDrob, ada)
        End If
    End If
    Else
    Drob = GetRod(Val(Combo2.List(Combo2.ListIndex)), Val(Combo3.List(Combo3.ListIndex)), Drob, KDrob, ada)
    End If
Else
If Per <> "" Then Text2.Text = Text2.Text + "."
End If
If Per <> "" Then
    If Val(Combo2.List(Combo2.ListIndex)) <> 16 Then
    If KPer Mod ada = 0 Then
        Per = GetRod(Val(Combo2.List(Combo2.ListIndex)), Val(Combo3.List(Combo3.ListIndex)), Per, KPer, ada)
    Else
        Dim perbeg As String, kperbeg As Long
        perbeg = Per
        kperbeg = KPer
        Do
            Per = Per + perbeg
            KPer = KPer + kperbeg
        Loop Until KPer Mod ada = 0
        Per = GetRod(Val(Combo2.List(Combo2.ListIndex)), Val(Combo3.List(Combo3.ListIndex)), Per, KPer, ada)
    End If
    Else
    Per = GetRod(Val(Combo2.List(Combo2.ListIndex)), Val(Combo3.List(Combo3.ListIndex)), Per, KPer, ada)
   End If
     If Per = Drob Then Text2.Text = Text2.Text + "(" + Per + ")" Else Text2.Text = Text2.Text + Drob + "(" + Per + ")"
Else
    If Drob <> "" Then Text2.Text = Text2.Text + Drob
End If
kol = kol + 1
stsStatus.Panels(3).Text = "Переведено чисел:" + Str(kol)
Text3.SetFocus
GoTo 20
10
MsgBox "Error#" + Str(Err.Number) + Chr(13) + "Description: " + CStr(Err.Description) + Chr(13) + "Source: " + CStr(Err.Source), vbCritical, "Error!"
Text3.SetFocus

20 End Sub


Private Sub Form_Load()
oldtime = Time$
stsStatus.Panels(1).Text = Time$
stsStatus.Panels(3).Text = "Переведено чисел:" + Str(kol)
For x = 2 To 16
Combo1.AddItem (Trim(Str(x)))
Next
Combo1.ListIndex = 0
Combo2.AddItem "2"
Combo2.AddItem "3"
Combo2.AddItem "4"
Combo2.AddItem "8"
Combo2.AddItem "9"
Combo2.AddItem "16"
Combo2.ListIndex = 0
Label2.Caption = "Число для перевода:" + Chr(13)
Open App.Path + "\2to4" For Random As #1 Len = 8
Open App.Path + "\2to8" For Random As #2 Len = 8
Open App.Path + "\2to16" For Random As #3 Len = 8
Open App.Path + "\3to9" For Random As #4 Len = 8
Open App.Path + "\4to16" For Random As #5 Len = 8
End Sub

Private Sub Picture1_Click()
frmAbout.Show vbModal
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
Case 48 To 57
Case 40
Case 41
Case 44
Case 65 To 70
Case 46
Case 8
Case 13
Command1_Click
Case Else
KeyAscii = 0
End Select
End Sub

Private Sub Timer1_Timer()


If Time$ <> oldtime Then stsStatus.Panels(1).Text = Time$: oldtime = Time$

End Sub

Public Sub GetPeriod(p1 As Byte, DrZnam As Double, DrChisl As Double)
    Dim rc As String, DrobChisl() As Double, DrobZnam As Double, p As Byte, koldrob As Double
    p = p1
    koldrob = 1
    ReDim Preserve DrobChisl(1)
    DrobChisl(1) = DrChisl
    rc = ""
    DrobZnam = DrZnam
    Dim period As Boolean, periodbeg As Double
    Do
        koldrob = koldrob + 1
        ReDim Preserve DrobChisl(koldrob)
        DrobChisl(koldrob) = DrobChisl(koldrob - 1) * p
        Debug.Print DrobChisl(koldrob) \ DrobZnam
        Select Case DrobChisl(koldrob) \ DrobZnam
        Case Is = 10
            rc = rc + "A"
        Case Is = 11
            rc = rc + "B"
        Case Is = 12
            rc = rc + "C"
        Case Is = 13
            rc = rc + "D"
        Case Is = 14
            rc = rc + "E"
        Case Is = 15
            rc = rc + "F"
        Case Else
            rc = rc + Trim(Str(DrobChisl(koldrob) \ DrobZnam))
        End Select
        DrobChisl(koldrob) = DrobChisl(koldrob) Mod DrobZnam
        If DrobChisl(koldrob) <> 0 Then
            Dim i As Double
            For i = 1 To koldrob - 1
            If DrobChisl(koldrob) = DrobChisl(i) Then
            period = True
            perbeg = i - 1
            End If
            Next
        End If
    Loop Until period = True Or DrobChisl(koldrob) = 0
    If period = True Then
        Text2.Text = Text2.Text + "." + Left(rc, perbeg) + "(" + Right(rc, Len(rc) - perbeg) + ")"
    Else
        Text2.Text = Text2.Text + "." + rc
    End If
End Sub

Function Gorner(Cel As String, KCel As Long)
Dim Celoe As String, KolCel As Long, chast As Double
Dim x As Long
Celoe = Cel
KolCel = KCel
x = 1

Select Case Mid(Celoe, x, 1)
Case Is = "A"
chast = 10
Case Is = "B"
chast = 11
Case Is = "C"
chast = 12
Case Is = "D"
chast = 13
Case Is = "E"
chast = 14
Case Is = "F"
chast = 15
Case Else
chast = Val(Mid(Celoe, x, 1))
End Select

x = 2
If KolCel <> 1 Then

Do
Select Case Mid(Celoe, x, 1)
Case Is = "A"
chast = chast * p + 10
Case Is = "B"
chast = chast * p + 11
Case Is = "C"
chast = chast * p + 12
Case Is = "D"
chast = chast * p + 13
Case Is = "E"
chast = chast * p + 14
Case Is = "F"
chast = chast * p + 15
Case Else
chast = chast * p + Val(Mid(Celoe, x, 1))
End Select
x = x + 1
Loop Until x >= KolCel + 1
End If
Gorner = Trim(Str(chast))

End Function


Private Sub Text3_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
Case 48 To 57
Case 40
Case 41
Case 44
Case 65 To 70
Case 46
Case 8
Case 13
Command1_Click
Case Else
KeyAscii = 0
End Select
End Sub
Function GetRod(p1 As Byte, p2 As Byte, ChisRod As String, ChisKol As Long, ad As Byte)
Dim i As Byte, fileno As Byte, pos As Long, temp As String

pos = 1
Select Case p1
    Case Is = 2
        Select Case p2
            Case Is = 4
                fileno = 1
            Case Is = 8
                fileno = 2
            Case Is = 16
                fileno = 3
        End Select
    Case Is = 3
        fileno = 4
    Case Is = 4
        If p2 = 2 Then fileno = 1 Else fileno = 5
    Case Is = 8
        fileno = 2
    Case Is = 9
        fileno = 4
    Case Is = 16
        If p2 = 2 Then fileno = 3 Else fileno = 5
End Select
Dim PRod As rod
i = 1
Do
    Do Until EOF(fileno)
    Get #fileno, i, PRod
    If Mid(ChisRod, pos, ad) = Trim(PRod.px) Then temp = temp + Trim(PRod.py) Else If Mid(ChisRod, pos, ad) = Trim(PRod.py) Then temp = temp + Trim(PRod.px)
    i = i + 1
    Loop
    pos = pos + ad
    i = 1
        Get #fileno, i, PRod
Loop Until pos > ChisKol
GetRod = temp
End Function
