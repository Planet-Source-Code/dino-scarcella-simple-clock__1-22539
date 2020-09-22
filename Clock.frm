VERSION 5.00
Begin VB.Form Clock 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   8535
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10080
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   8535
   ScaleWidth      =   10080
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer TimerEverySecond 
      Interval        =   999
      Left            =   120
      Top             =   135
   End
   Begin VB.Label lblDate 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   420
      Left            =   5190
      TabIndex        =   1
      Top             =   7200
      Width           =   1530
   End
   Begin VB.Label lblDigitime 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   540
      Left            =   4620
      TabIndex        =   0
      Top             =   1260
      Width           =   4215
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      Index           =   19
      X1              =   7950
      X2              =   7860
      Y1              =   5505
      Y2              =   5475
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      Index           =   18
      X1              =   8055
      X2              =   7950
      Y1              =   5265
      Y2              =   5250
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      Index           =   21
      X1              =   7725
      X2              =   7650
      Y1              =   5910
      Y2              =   5850
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      Index           =   22
      X1              =   7545
      X2              =   7470
      Y1              =   6135
      Y2              =   6060
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      Index           =   23
      X1              =   7350
      X2              =   7275
      Y1              =   6330
      Y2              =   6255
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      Index           =   17
      X1              =   8130
      X2              =   8025
      Y1              =   4995
      Y2              =   4980
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      Index           =   24
      X1              =   7125
      X2              =   7050
      Y1              =   6510
      Y2              =   6435
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      Index           =   16
      X1              =   8175
      X2              =   8055
      Y1              =   4695
      Y2              =   4680
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      Index           =   26
      X1              =   6735
      X2              =   6675
      Y1              =   6735
      Y2              =   6630
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      Index           =   27
      X1              =   6495
      X2              =   6465
      Y1              =   6825
      Y2              =   6735
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      Index           =   28
      X1              =   6240
      X2              =   6225
      Y1              =   6900
      Y2              =   6810
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      Index           =   29
      X1              =   5940
      X2              =   5925
      Y1              =   6960
      Y2              =   6855
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      Index           =   31
      X1              =   5355
      X2              =   5385
      Y1              =   6960
      Y2              =   6855
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      Index           =   32
      X1              =   5085
      X2              =   5115
      Y1              =   6915
      Y2              =   6795
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      Index           =   33
      X1              =   4845
      X2              =   4875
      Y1              =   6825
      Y2              =   6720
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      Index           =   34
      X1              =   4620
      X2              =   4665
      Y1              =   6735
      Y2              =   6630
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      Index           =   36
      X1              =   4170
      X2              =   4245
      Y1              =   6480
      Y2              =   6405
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      Index           =   37
      X1              =   3945
      X2              =   4020
      Y1              =   6300
      Y2              =   6225
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      Index           =   38
      X1              =   3750
      X2              =   3840
      Y1              =   6090
      Y2              =   6015
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      Index           =   39
      X1              =   3600
      X2              =   3690
      Y1              =   5880
      Y2              =   5835
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      Index           =   41
      X1              =   3375
      X2              =   3465
      Y1              =   5490
      Y2              =   5445
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      Index           =   42
      X1              =   3285
      X2              =   3375
      Y1              =   5250
      Y2              =   5220
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      Index           =   43
      X1              =   3195
      X2              =   3315
      Y1              =   4995
      Y2              =   4965
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      Index           =   44
      X1              =   3165
      X2              =   3255
      Y1              =   4710
      Y2              =   4695
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      Index           =   46
      X1              =   3165
      X2              =   3255
      Y1              =   4215
      Y2              =   4230
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      Index           =   47
      X1              =   3180
      X2              =   3285
      Y1              =   3960
      Y2              =   3990
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      Index           =   48
      X1              =   3270
      X2              =   3360
      Y1              =   3690
      Y2              =   3720
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      Index           =   49
      X1              =   3360
      X2              =   3435
      Y1              =   3435
      Y2              =   3480
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      Index           =   51
      X1              =   3585
      X2              =   3645
      Y1              =   3030
      Y2              =   3090
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      Index           =   52
      X1              =   3765
      X2              =   3840
      Y1              =   2805
      Y2              =   2880
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      Index           =   53
      X1              =   3975
      X2              =   4050
      Y1              =   2595
      Y2              =   2670
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      Index           =   54
      X1              =   4185
      X2              =   4245
      Y1              =   2430
      Y2              =   2505
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      Index           =   56
      X1              =   4605
      X2              =   4635
      Y1              =   2175
      Y2              =   2250
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      Index           =   57
      X1              =   4845
      X2              =   4875
      Y1              =   2085
      Y2              =   2175
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      Index           =   58
      X1              =   5115
      X2              =   5130
      Y1              =   1995
      Y2              =   2100
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      Index           =   59
      X1              =   5370
      X2              =   5385
      Y1              =   1950
      Y2              =   2055
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      Index           =   35
      X1              =   4395
      X2              =   4470
      Y1              =   6645
      Y2              =   6525
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      Index           =   40
      X1              =   3480
      X2              =   3600
      Y1              =   5700
      Y2              =   5625
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      Index           =   50
      X1              =   3465
      X2              =   3570
      Y1              =   3255
      Y2              =   3330
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      Index           =   55
      X1              =   4395
      X2              =   4470
      Y1              =   2295
      Y2              =   2400
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      Index           =   25
      X1              =   6945
      X2              =   6855
      Y1              =   6630
      Y2              =   6525
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      Index           =   20
      X1              =   7845
      X2              =   7710
      Y1              =   5700
      Y2              =   5625
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      Index           =   14
      X1              =   8175
      X2              =   8070
      Y1              =   4215
      Y2              =   4230
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      Index           =   13
      X1              =   8130
      X2              =   8040
      Y1              =   3975
      Y2              =   3990
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      Index           =   12
      X1              =   8055
      X2              =   7965
      Y1              =   3705
      Y2              =   3720
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      Index           =   11
      X1              =   7965
      X2              =   7860
      Y1              =   3465
      Y2              =   3495
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      Index           =   9
      X1              =   7755
      X2              =   7680
      Y1              =   3045
      Y2              =   3075
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      Index           =   8
      X1              =   7575
      X2              =   7500
      Y1              =   2805
      Y2              =   2880
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      Index           =   7
      X1              =   7365
      X2              =   7290
      Y1              =   2595
      Y2              =   2670
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      Index           =   6
      X1              =   7155
      X2              =   7095
      Y1              =   2415
      Y2              =   2505
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      Index           =   4
      X1              =   6720
      X2              =   6675
      Y1              =   2190
      Y2              =   2280
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      Index           =   3
      X1              =   6480
      X2              =   6435
      Y1              =   2085
      Y2              =   2190
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      Index           =   2
      X1              =   6225
      X2              =   6195
      Y1              =   2010
      Y2              =   2115
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      Index           =   1
      X1              =   5955
      X2              =   5940
      Y1              =   1965
      Y2              =   2070
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      Index           =   5
      X1              =   6930
      X2              =   6870
      Y1              =   2295
      Y2              =   2415
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      Index           =   10
      X1              =   7890
      X2              =   7770
      Y1              =   3255
      Y2              =   3315
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      Index           =   30
      X1              =   5670
      X2              =   5670
      Y1              =   6975
      Y2              =   6810
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      Index           =   45
      X1              =   3150
      X2              =   3300
      Y1              =   4455
      Y2              =   4455
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      Index           =   15
      X1              =   8190
      X2              =   8055
      Y1              =   4455
      Y2              =   4455
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      Index           =   0
      X1              =   5670
      X2              =   5670
      Y1              =   1935
      Y2              =   2115
   End
   Begin VB.Line MinuteH 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   6
      X1              =   5670
      X2              =   5670
      Y1              =   4395
      Y2              =   6780
   End
   Begin VB.Line HourH 
      BorderColor     =   &H008080FF&
      BorderWidth     =   6
      X1              =   5670
      X2              =   5670
      Y1              =   4395
      Y2              =   3195
   End
   Begin VB.Line SecondH 
      BorderColor     =   &H000000FF&
      X1              =   5670
      X2              =   5670
      Y1              =   4395
      Y2              =   2115
   End
End
Attribute VB_Name = "Clock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub TimerEverySecond_Timer()
 lblDate.Caption = Date
 lblDigitime.Caption = Time
 SecondH.X2 = Line1(Second(Time)).X2
 SecondH.Y2 = Line1(Second(Time)).Y2
 MinuteH.X2 = Line1(Minute(Time)).X2
 MinuteH.Y2 = Line1(Minute(Time)).Y2
 
 Select Case Minute(Time)
  Case 0 To 12
   HourH.X2 = Line1((Hour(Time) Mod 12) * 5).X2
   HourH.Y2 = Line1((Hour(Time) Mod 12) * 5).Y2
   HourH.X2 = (HourH.X1 + HourH.X2) / 2
   HourH.Y2 = (HourH.Y1 + HourH.Y2) / 2
  Case 13 To 24
   HourH.X2 = Line1(((Hour(Time) Mod 12) * 5) + 1).X2
   HourH.Y2 = Line1(((Hour(Time) Mod 12) * 5) + 1).Y2
   HourH.X2 = (HourH.X1 + HourH.X2) / 2
   HourH.Y2 = (HourH.Y1 + HourH.Y2) / 2
  Case 25 To 36
   HourH.X2 = Line1(((Hour(Time) Mod 12) * 5) + 2).X2
   HourH.Y2 = Line1(((Hour(Time) Mod 12) * 5) + 2).Y2
   HourH.X2 = (HourH.X1 + HourH.X2) / 2
   HourH.Y2 = (HourH.Y1 + HourH.Y2) / 2
  Case 37 To 48
   HourH.X2 = Line1(((Hour(Time) Mod 12) * 5) + 3).X2
   HourH.Y2 = Line1(((Hour(Time) Mod 12) * 5) + 3).Y2
   HourH.X2 = (HourH.X1 + HourH.X2) / 2
   HourH.Y2 = (HourH.Y1 + HourH.Y2) / 2
  Case 49 To 59
   HourH.X2 = Line1(((Hour(Time) Mod 12) * 5) + 4).X2
   HourH.Y2 = Line1(((Hour(Time) Mod 12) * 5) + 4).Y2
   HourH.X2 = (HourH.X1 + HourH.X2) / 2
   HourH.Y2 = (HourH.Y1 + HourH.Y2) / 2
  End Select
 End Sub

Private Sub Form_MouseMove(Button As Integer, shift As Integer, X As Single, y As Single)
Static counter As Integer
 
 counter = counter + 1
 If counter = 8 Then Unload Me

End Sub


