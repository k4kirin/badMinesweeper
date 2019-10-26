VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Minesweeper!"
   ClientHeight    =   4395
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6375
   LinkTopic       =   "Form1"
   ScaleHeight     =   4395
   ScaleWidth      =   6375
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CommandRestart 
      Caption         =   "Restart"
      Height          =   615
      Left            =   4440
      TabIndex        =   101
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton CommandFlag 
      Caption         =   "Flag Mode Off"
      Height          =   615
      Left            =   4440
      TabIndex        =   100
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton Command100 
      Height          =   255
      Left            =   3600
      TabIndex        =   99
      Top             =   2880
      Width           =   255
   End
   Begin VB.CommandButton Command99 
      Height          =   255
      Left            =   3360
      TabIndex        =   98
      Top             =   2880
      Width           =   255
   End
   Begin VB.CommandButton Command98 
      Height          =   255
      Left            =   3120
      TabIndex        =   97
      Top             =   2880
      Width           =   255
   End
   Begin VB.CommandButton Command97 
      Height          =   255
      Left            =   2880
      TabIndex        =   96
      Top             =   2880
      Width           =   255
   End
   Begin VB.CommandButton Command96 
      Height          =   255
      Left            =   2640
      TabIndex        =   95
      Top             =   2880
      Width           =   255
   End
   Begin VB.CommandButton Command95 
      Height          =   255
      Left            =   2400
      TabIndex        =   94
      Top             =   2880
      Width           =   255
   End
   Begin VB.CommandButton Command94 
      Height          =   255
      Left            =   2160
      TabIndex        =   93
      Top             =   2880
      Width           =   255
   End
   Begin VB.CommandButton Command93 
      Height          =   255
      Left            =   1920
      TabIndex        =   92
      Top             =   2880
      Width           =   255
   End
   Begin VB.CommandButton Command92 
      Height          =   255
      Left            =   1680
      TabIndex        =   91
      Top             =   2880
      Width           =   255
   End
   Begin VB.CommandButton Command91 
      Height          =   255
      Left            =   1440
      TabIndex        =   90
      Top             =   2880
      Width           =   255
   End
   Begin VB.CommandButton Command90 
      Height          =   255
      Left            =   3600
      TabIndex        =   89
      Top             =   2640
      Width           =   255
   End
   Begin VB.CommandButton Command89 
      Height          =   255
      Left            =   3360
      TabIndex        =   88
      Top             =   2640
      Width           =   255
   End
   Begin VB.CommandButton Command88 
      Height          =   255
      Left            =   3120
      TabIndex        =   87
      Top             =   2640
      Width           =   255
   End
   Begin VB.CommandButton Command87 
      Height          =   255
      Left            =   2880
      TabIndex        =   86
      Top             =   2640
      Width           =   255
   End
   Begin VB.CommandButton Command86 
      Height          =   255
      Left            =   2640
      TabIndex        =   85
      Top             =   2640
      Width           =   255
   End
   Begin VB.CommandButton Command85 
      Height          =   255
      Left            =   2400
      TabIndex        =   84
      Top             =   2640
      Width           =   255
   End
   Begin VB.CommandButton Command84 
      Height          =   255
      Left            =   2160
      TabIndex        =   83
      Top             =   2640
      Width           =   255
   End
   Begin VB.CommandButton Command83 
      Height          =   255
      Left            =   1920
      TabIndex        =   82
      Top             =   2640
      Width           =   255
   End
   Begin VB.CommandButton Command82 
      Height          =   255
      Left            =   1680
      TabIndex        =   81
      Top             =   2640
      Width           =   255
   End
   Begin VB.CommandButton Command81 
      Height          =   255
      Left            =   1440
      TabIndex        =   80
      Top             =   2640
      Width           =   255
   End
   Begin VB.CommandButton Command80 
      Height          =   255
      Left            =   3600
      TabIndex        =   79
      Top             =   2400
      Width           =   255
   End
   Begin VB.CommandButton Command79 
      Height          =   255
      Left            =   3360
      TabIndex        =   78
      Top             =   2400
      Width           =   255
   End
   Begin VB.CommandButton Command78 
      Height          =   255
      Left            =   3120
      TabIndex        =   77
      Top             =   2400
      Width           =   255
   End
   Begin VB.CommandButton Command77 
      Height          =   255
      Left            =   2880
      TabIndex        =   76
      Top             =   2400
      Width           =   255
   End
   Begin VB.CommandButton Command76 
      Height          =   255
      Left            =   2640
      TabIndex        =   75
      Top             =   2400
      Width           =   255
   End
   Begin VB.CommandButton Command75 
      Height          =   255
      Left            =   2400
      TabIndex        =   74
      Top             =   2400
      Width           =   255
   End
   Begin VB.CommandButton Command74 
      Height          =   255
      Left            =   2160
      TabIndex        =   73
      Top             =   2400
      Width           =   255
   End
   Begin VB.CommandButton Command73 
      Height          =   255
      Left            =   1920
      TabIndex        =   72
      Top             =   2400
      Width           =   255
   End
   Begin VB.CommandButton Command72 
      Height          =   255
      Left            =   1680
      TabIndex        =   71
      Top             =   2400
      Width           =   255
   End
   Begin VB.CommandButton Command71 
      Height          =   255
      Left            =   1440
      TabIndex        =   70
      Top             =   2400
      Width           =   255
   End
   Begin VB.CommandButton Command70 
      Height          =   255
      Left            =   3600
      TabIndex        =   69
      Top             =   2160
      Width           =   255
   End
   Begin VB.CommandButton Command69 
      Height          =   255
      Left            =   3360
      TabIndex        =   68
      Top             =   2160
      Width           =   255
   End
   Begin VB.CommandButton Command68 
      Height          =   255
      Left            =   3120
      TabIndex        =   67
      Top             =   2160
      Width           =   255
   End
   Begin VB.CommandButton Command67 
      Height          =   255
      Left            =   2880
      TabIndex        =   66
      Top             =   2160
      Width           =   255
   End
   Begin VB.CommandButton Command66 
      Height          =   255
      Left            =   2640
      TabIndex        =   65
      Top             =   2160
      Width           =   255
   End
   Begin VB.CommandButton Command65 
      Height          =   255
      Left            =   2400
      TabIndex        =   64
      Top             =   2160
      Width           =   255
   End
   Begin VB.CommandButton Command64 
      Height          =   255
      Left            =   2160
      TabIndex        =   63
      Top             =   2160
      Width           =   255
   End
   Begin VB.CommandButton Command63 
      Height          =   255
      Left            =   1920
      TabIndex        =   62
      Top             =   2160
      Width           =   255
   End
   Begin VB.CommandButton Command62 
      Height          =   255
      Left            =   1680
      TabIndex        =   61
      Top             =   2160
      Width           =   255
   End
   Begin VB.CommandButton Command61 
      Height          =   255
      Left            =   1440
      TabIndex        =   60
      Top             =   2160
      Width           =   255
   End
   Begin VB.CommandButton Command60 
      Height          =   255
      Left            =   3600
      TabIndex        =   59
      Top             =   1920
      Width           =   255
   End
   Begin VB.CommandButton Command59 
      Height          =   255
      Left            =   3360
      TabIndex        =   58
      Top             =   1920
      Width           =   255
   End
   Begin VB.CommandButton Command58 
      Height          =   255
      Left            =   3120
      TabIndex        =   57
      Top             =   1920
      Width           =   255
   End
   Begin VB.CommandButton Command57 
      Height          =   255
      Left            =   2880
      TabIndex        =   56
      Top             =   1920
      Width           =   255
   End
   Begin VB.CommandButton Command56 
      Height          =   255
      Left            =   2640
      TabIndex        =   55
      Top             =   1920
      Width           =   255
   End
   Begin VB.CommandButton Command55 
      Height          =   255
      Left            =   2400
      TabIndex        =   54
      Top             =   1920
      Width           =   255
   End
   Begin VB.CommandButton Command54 
      Height          =   255
      Left            =   2160
      TabIndex        =   53
      Top             =   1920
      Width           =   255
   End
   Begin VB.CommandButton Command53 
      Height          =   255
      Left            =   1920
      TabIndex        =   52
      Top             =   1920
      Width           =   255
   End
   Begin VB.CommandButton Command52 
      Height          =   255
      Left            =   1680
      TabIndex        =   51
      Top             =   1920
      Width           =   255
   End
   Begin VB.CommandButton Command51 
      Height          =   255
      Left            =   1440
      TabIndex        =   50
      Top             =   1920
      Width           =   255
   End
   Begin VB.CommandButton Command50 
      Height          =   255
      Left            =   3600
      TabIndex        =   49
      Top             =   1680
      Width           =   255
   End
   Begin VB.CommandButton Command49 
      Height          =   255
      Left            =   3360
      TabIndex        =   48
      Top             =   1680
      Width           =   255
   End
   Begin VB.CommandButton Command48 
      Height          =   255
      Left            =   3120
      TabIndex        =   47
      Top             =   1680
      Width           =   255
   End
   Begin VB.CommandButton Command47 
      Height          =   255
      Left            =   2880
      TabIndex        =   46
      Top             =   1680
      Width           =   255
   End
   Begin VB.CommandButton Command46 
      Height          =   255
      Left            =   2640
      TabIndex        =   45
      Top             =   1680
      Width           =   255
   End
   Begin VB.CommandButton Command45 
      Height          =   255
      Left            =   2400
      TabIndex        =   44
      Top             =   1680
      Width           =   255
   End
   Begin VB.CommandButton Command44 
      Height          =   255
      Left            =   2160
      TabIndex        =   43
      Top             =   1680
      Width           =   255
   End
   Begin VB.CommandButton Command43 
      Height          =   255
      Left            =   1920
      TabIndex        =   42
      Top             =   1680
      Width           =   255
   End
   Begin VB.CommandButton Command42 
      Height          =   255
      Left            =   1680
      TabIndex        =   41
      Top             =   1680
      Width           =   255
   End
   Begin VB.CommandButton Command41 
      Height          =   255
      Left            =   1440
      TabIndex        =   40
      Top             =   1680
      Width           =   255
   End
   Begin VB.CommandButton Command40 
      Height          =   255
      Left            =   3600
      TabIndex        =   39
      Top             =   1440
      Width           =   255
   End
   Begin VB.CommandButton Command39 
      Height          =   255
      Left            =   3360
      TabIndex        =   38
      Top             =   1440
      Width           =   255
   End
   Begin VB.CommandButton Command38 
      Height          =   255
      Left            =   3120
      TabIndex        =   37
      Top             =   1440
      Width           =   255
   End
   Begin VB.CommandButton Command37 
      Height          =   255
      Left            =   2880
      TabIndex        =   36
      Top             =   1440
      Width           =   255
   End
   Begin VB.CommandButton Command36 
      Height          =   255
      Left            =   2640
      TabIndex        =   35
      Top             =   1440
      Width           =   255
   End
   Begin VB.CommandButton Command35 
      Height          =   255
      Left            =   2400
      TabIndex        =   34
      Top             =   1440
      Width           =   255
   End
   Begin VB.CommandButton Command34 
      Height          =   255
      Left            =   2160
      TabIndex        =   33
      Top             =   1440
      Width           =   255
   End
   Begin VB.CommandButton Command33 
      Height          =   255
      Left            =   1920
      TabIndex        =   32
      Top             =   1440
      Width           =   255
   End
   Begin VB.CommandButton Command32 
      Height          =   255
      Left            =   1680
      TabIndex        =   31
      Top             =   1440
      Width           =   255
   End
   Begin VB.CommandButton Command31 
      Height          =   255
      Left            =   1440
      TabIndex        =   30
      Top             =   1440
      Width           =   255
   End
   Begin VB.CommandButton Command30 
      Height          =   255
      Left            =   3600
      TabIndex        =   29
      Top             =   1200
      Width           =   255
   End
   Begin VB.CommandButton Command29 
      Height          =   255
      Left            =   3360
      TabIndex        =   28
      Top             =   1200
      Width           =   255
   End
   Begin VB.CommandButton Command28 
      Height          =   255
      Left            =   3120
      TabIndex        =   27
      Top             =   1200
      Width           =   255
   End
   Begin VB.CommandButton Command27 
      Height          =   255
      Left            =   2880
      TabIndex        =   26
      Top             =   1200
      Width           =   255
   End
   Begin VB.CommandButton Command26 
      Height          =   255
      Left            =   2640
      TabIndex        =   25
      Top             =   1200
      Width           =   255
   End
   Begin VB.CommandButton Command25 
      Height          =   255
      Left            =   2400
      TabIndex        =   24
      Top             =   1200
      Width           =   255
   End
   Begin VB.CommandButton Command24 
      Height          =   255
      Left            =   2160
      TabIndex        =   23
      Top             =   1200
      Width           =   255
   End
   Begin VB.CommandButton Command23 
      Height          =   255
      Left            =   1920
      TabIndex        =   22
      Top             =   1200
      Width           =   255
   End
   Begin VB.CommandButton Command22 
      Height          =   255
      Left            =   1680
      TabIndex        =   21
      Top             =   1200
      Width           =   255
   End
   Begin VB.CommandButton Command21 
      Height          =   255
      Left            =   1440
      TabIndex        =   20
      Top             =   1200
      Width           =   255
   End
   Begin VB.CommandButton Command20 
      Height          =   255
      Left            =   3600
      TabIndex        =   19
      Top             =   960
      Width           =   255
   End
   Begin VB.CommandButton Command19 
      Height          =   255
      Left            =   3360
      TabIndex        =   18
      Top             =   960
      Width           =   255
   End
   Begin VB.CommandButton Command18 
      Height          =   255
      Left            =   3120
      TabIndex        =   17
      Top             =   960
      Width           =   255
   End
   Begin VB.CommandButton Command17 
      Height          =   255
      Left            =   2880
      TabIndex        =   16
      Top             =   960
      Width           =   255
   End
   Begin VB.CommandButton Command16 
      Height          =   255
      Left            =   2640
      TabIndex        =   15
      Top             =   960
      Width           =   255
   End
   Begin VB.CommandButton Command15 
      Height          =   255
      Left            =   2400
      TabIndex        =   14
      Top             =   960
      Width           =   255
   End
   Begin VB.CommandButton Command14 
      Height          =   255
      Left            =   2160
      TabIndex        =   13
      Top             =   960
      Width           =   255
   End
   Begin VB.CommandButton Command13 
      Height          =   255
      Left            =   1920
      TabIndex        =   12
      Top             =   960
      Width           =   255
   End
   Begin VB.CommandButton Command12 
      Height          =   255
      Left            =   1680
      TabIndex        =   11
      Top             =   960
      Width           =   255
   End
   Begin VB.CommandButton Command11 
      Height          =   255
      Left            =   1440
      TabIndex        =   10
      Top             =   960
      Width           =   255
   End
   Begin VB.CommandButton Command10 
      Height          =   255
      Left            =   3600
      TabIndex        =   9
      Top             =   720
      Width           =   255
   End
   Begin VB.CommandButton Command9 
      Height          =   255
      Left            =   3360
      TabIndex        =   8
      Top             =   720
      Width           =   255
   End
   Begin VB.CommandButton Command8 
      Height          =   255
      Left            =   3120
      TabIndex        =   7
      Top             =   720
      Width           =   255
   End
   Begin VB.CommandButton Command7 
      Height          =   255
      Left            =   2880
      TabIndex        =   6
      Top             =   720
      Width           =   255
   End
   Begin VB.CommandButton Command6 
      Height          =   255
      Left            =   2640
      TabIndex        =   5
      Top             =   720
      Width           =   255
   End
   Begin VB.CommandButton Command5 
      Height          =   255
      Left            =   2400
      TabIndex        =   4
      Top             =   720
      Width           =   255
   End
   Begin VB.CommandButton Command4 
      Height          =   255
      Left            =   2160
      TabIndex        =   3
      Top             =   720
      Width           =   255
   End
   Begin VB.CommandButton Command3 
      Height          =   255
      Left            =   1920
      TabIndex        =   2
      Top             =   720
      Width           =   255
   End
   Begin VB.CommandButton Command2 
      Height          =   255
      Left            =   1680
      TabIndex        =   1
      Top             =   720
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000010&
      Height          =   255
      Left            =   1440
      TabIndex        =   0
      Top             =   720
      Width           =   255
   End
   Begin VB.Label Label4 
      Caption         =   "You Win!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      TabIndex        =   106
      Top             =   3840
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label Label3 
      Caption         =   "28/1/2019"
      Height          =   255
      Left            =   3720
      TabIndex        =   105
      Top             =   4080
      Width           =   2655
   End
   Begin VB.Label Label2 
      Caption         =   "By Kirin Sarangkasiri 612 #12 34964"
      Height          =   495
      Left            =   3720
      TabIndex        =   104
      Top             =   3840
      Width           =   2655
   End
   Begin VB.Label LabelGO 
      Caption         =   "Game Over"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   103
      Top             =   3360
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "Minesweeper!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1200
      TabIndex        =   102
      Top             =   0
      Width           =   3255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim f(9, 9) As Integer
Dim r(9, 9) As Integer
Dim p As Integer
Dim fl As Integer
Private Sub CommandFlag_Click()
    If fl = 0 Then
        fl = 1
        CommandFlag.Caption = "Flag Mode On"
    Else
        fl = 0
        CommandFlag.Caption = "Flag Mode Off"
    End If
End Sub

Private Sub CommandRestart_Click()
    LabelGO.Visible = False
    For i = 0 To 9
        For j = 0 To 9
            f(i, j) = 0
            r(i, j) = 0
        Next
    Next
    setUp
    Command1.Caption = ""
    Command2.Caption = ""
    Command3.Caption = ""
    Command4.Caption = ""
    Command5.Caption = ""
    Command6.Caption = ""
    Command7.Caption = ""
    Command8.Caption = ""
    Command9.Caption = ""
    Command10.Caption = ""
    Command11.Caption = ""
    Command12.Caption = ""
    Command13.Caption = ""
    Command14.Caption = ""
    Command15.Caption = ""
    Command16.Caption = ""
    Command17.Caption = ""
    Command18.Caption = ""
    Command19.Caption = ""
    Command20.Caption = ""
    Command21.Caption = ""
    Command22.Caption = ""
    Command23.Caption = ""
    Command24.Caption = ""
    Command25.Caption = ""
    Command26.Caption = ""
    Command27.Caption = ""
    Command28.Caption = ""
    Command29.Caption = ""
    Command30.Caption = ""
    Command31.Caption = ""
    Command32.Caption = ""
    Command33.Caption = ""
    Command34.Caption = ""
    Command35.Caption = ""
    Command36.Caption = ""
    Command37.Caption = ""
    Command38.Caption = ""
    Command39.Caption = ""
    Command40.Caption = ""
    Command41.Caption = ""
    Command42.Caption = ""
    Command43.Caption = ""
    Command44.Caption = ""
    Command45.Caption = ""
    Command46.Caption = ""
    Command47.Caption = ""
    Command48.Caption = ""
    Command49.Caption = ""
    Command50.Caption = ""
    Command51.Caption = ""
    Command52.Caption = ""
    Command53.Caption = ""
    Command54.Caption = ""
    Command55.Caption = ""
    Command56.Caption = ""
    Command57.Caption = ""
    Command58.Caption = ""
    Command59.Caption = ""
    Command60.Caption = ""
    Command61.Caption = ""
    Command62.Caption = ""
    Command63.Caption = ""
    Command64.Caption = ""
    Command65.Caption = ""
    Command66.Caption = ""
    Command67.Caption = ""
    Command68.Caption = ""
    Command69.Caption = ""
    Command70.Caption = ""
    Command71.Caption = ""
    Command72.Caption = ""
    Command73.Caption = ""
    Command74.Caption = ""
    Command75.Caption = ""
    Command76.Caption = ""
    Command77.Caption = ""
    Command78.Caption = ""
    Command79.Caption = ""
    Command80.Caption = ""
    Command81.Caption = ""
    Command82.Caption = ""
    Command83.Caption = ""
    Command84.Caption = ""
    Command85.Caption = ""
    Command86.Caption = ""
    Command87.Caption = ""
    Command88.Caption = ""
    Command89.Caption = ""
    Command90.Caption = ""
    Command91.Caption = ""
    Command92.Caption = ""
    Command93.Caption = ""
    Command94.Caption = ""
    Command95.Caption = ""
    Command96.Caption = ""
    Command97.Caption = ""
    Command98.Caption = ""
    Command99.Caption = ""
    Command100.Caption = ""
    For Each ctrl In Me.Controls
         If TypeOf ctrl Is CommandButton Then
             ctrl.Enabled = True
         End If
    Next
    
End Sub

Private Sub Form_Load()
    setUp
End Sub
Function setUp()
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Randomize
    For k = 1 To 20
        p = Int(100 * Rnd)
        If f(Int(p / 10), p - 10 * (Int(p / 10))) = 0 Then
            f(Int(p / 10), p - 10 * (Int(p / 10))) = 9
        Else
            i = i - 1
        End If
        Next
    For i = 0 To 9
        For j = 0 To 9
            If f(i, j) <> 9 Then
                If i <> 0 Then
                    If j <> 0 Then If f(i - 1, j - 1) = 9 Then f(i, j) = f(i, j) + 1
                    If j <> 9 Then If f(i - 1, j + 1) = 9 Then f(i, j) = f(i, j) + 1
                    If f(i - 1, j) = 9 Then f(i, j) = f(i, j) + 1
                End If
                If i <> 9 Then
                    If j <> 0 Then If f(i + 1, j - 1) = 9 Then f(i, j) = f(i, j) + 1
                    If j <> 9 Then If f(i + 1, j + 1) = 9 Then f(i, j) = f(i, j) + 1
                    If f(i + 1, j) = 9 Then f(i, j) = f(i, j) + 1
                End If
                If j <> 0 Then If f(i, j - 1) = 9 Then f(i, j) = f(i, j) + 1
                If j <> 9 Then If f(i, j + 1) = 9 Then f(i, j) = f(i, j) + 1
            End If
        Next
    Next
End Function
Function endGame()
    Dim ctrl As Control
    p = 100
    Command1_Click
    Command2_Click
    Command3_Click
    Command4_Click
    Command5_Click
    Command6_Click
    Command7_Click
    Command8_Click
    Command9_Click
    Command10_Click
    Command11_Click
    Command12_Click
    Command13_Click
    Command14_Click
    Command15_Click
    Command16_Click
    Command17_Click
    Command18_Click
    Command19_Click
    Command20_Click
    Command21_Click
    Command22_Click
    Command23_Click
    Command24_Click
    Command25_Click
    Command26_Click
    Command27_Click
    Command28_Click
    Command29_Click
    Command30_Click
    Command31_Click
    Command32_Click
    Command33_Click
    Command34_Click
    Command35_Click
    Command36_Click
    Command37_Click
    Command38_Click
    Command39_Click
    Command40_Click
    Command41_Click
    Command42_Click
    Command43_Click
    Command44_Click
    Command45_Click
    Command46_Click
    Command47_Click
    Command48_Click
    Command49_Click
    Command50_Click
    Command51_Click
    Command52_Click
    Command53_Click
    Command54_Click
    Command55_Click
    Command56_Click
    Command57_Click
    Command58_Click
    Command59_Click
    Command60_Click
    Command61_Click
    Command62_Click
    Command63_Click
    Command64_Click
    Command65_Click
    Command66_Click
    Command67_Click
    Command68_Click
    Command69_Click
    Command70_Click
    Command71_Click
    Command72_Click
    Command73_Click
    Command74_Click
    Command75_Click
    Command76_Click
    Command77_Click
    Command78_Click
    Command79_Click
    Command80_Click
    Command81_Click
    Command82_Click
    Command83_Click
    Command84_Click
    Command85_Click
    Command86_Click
    Command87_Click
    Command88_Click
    Command89_Click
    Command90_Click
    Command91_Click
    Command92_Click
    Command93_Click
    Command94_Click
    Command95_Click
    Command96_Click
    Command97_Click
    Command98_Click
    Command99_Click
    Command100_Click
    For Each ctrl In Me.Controls
         If TypeOf ctrl Is CommandButton Then
             ctrl.Enabled = False
         End If
    Next
    LabelGO.Visible = True
    CommandRestart.Enabled = True
End Function
Private Sub Command1_Click()
    If fl = 1 Then
        If Command1.Caption <> "F" Then
            Command1.Caption = "F"
        Else
            Command1.Caption = ""
        End If
    ElseIf Command1.Caption <> "F" Then
        r(0, 0) = 1
        If f(0, 0) = 9 Then
            Command1.Caption = "B"
            If p <> 100 Then endGame
        Else
            Command1.Caption = f(0, 0)
        End If
        If f(0, 0) = 0 Then
            If r(1, 0) = 0 And f(1, 0) < 9 Then Command11_Click
            If r(0, 1) = 0 And f(0, 1) < 9 Then Command2_Click
        End If
    End If
End Sub
Private Sub Command2_Click()
    If fl = 1 Then
        If Command2.Caption <> "F" Then
            Command2.Caption = "F"
        Else
            Command2.Caption = ""
        End If
    ElseIf Command2.Caption <> "F" Then
        r(0, 1) = 1
        If f(0, 1) = 9 Then
            Command2.Caption = "B"
            If p <> 100 Then endGame
        Else
            Command2.Caption = f(0, 1)
        End If
        If f(0, 1) = 0 Then
            If r(1, 1) = 0 And f(1, 1) < 9 Then Command12_Click
            If r(0, 2) = 0 And f(0, 2) < 9 Then Command3_Click
            If r(0, 0) = 0 And f(0, 0) < 9 Then Command1_Click
        End If
    End If
End Sub
Private Sub Command3_Click()
    If fl = 1 Then
        If Command3.Caption <> "F" Then
            Command3.Caption = "F"
        Else
            Command3.Caption = ""
        End If
    ElseIf Command3.Caption <> "F" Then
        r(0, 2) = 1
        If f(0, 2) = 9 Then
            Command3.Caption = "B"
            If p <> 100 Then endGame
        Else
            Command3.Caption = f(0, 2)
        End If
        If f(0, 2) = 0 Then
            If r(1, 2) = 0 And f(1, 2) < 9 Then Command13_Click
            If r(0, 3) = 0 And f(0, 3) < 9 Then Command4_Click
            If r(0, 1) = 0 And f(0, 1) < 9 Then Command2_Click
        End If
    End If
End Sub
Private Sub Command4_Click()
    If fl = 1 Then
        If Command4.Caption <> "F" Then
            Command4.Caption = "F"
        Else
            Command4.Caption = ""
        End If
    ElseIf Command4.Caption <> "F" Then
        r(0, 3) = 1
        If f(0, 3) = 9 Then
            Command4.Caption = "B"
            If p <> 100 Then endGame
        Else
            Command4.Caption = f(0, 3)
        End If
        If f(0, 3) = 0 Then
            If r(1, 3) = 0 And f(1, 3) < 9 Then Command14_Click
            If r(0, 4) = 0 And f(0, 4) < 9 Then Command5_Click
            If r(0, 2) = 0 And f(0, 2) < 9 Then Command3_Click
        End If
    End If
End Sub
Private Sub Command5_Click()
    If fl = 1 Then
        If Command5.Caption <> "F" Then
            Command5.Caption = "F"
        Else
            Command5.Caption = ""
        End If
    ElseIf Command5.Caption <> "F" Then
        r(0, 4) = 1
        If f(0, 4) = 9 Then
            Command5.Caption = "B"
            If p <> 100 Then endGame
        Else
            Command5.Caption = f(0, 4)
        End If
        If f(0, 4) = 0 Then
            If r(1, 4) = 0 And f(1, 4) < 9 Then Command15_Click
            If r(0, 5) = 0 And f(0, 5) < 9 Then Command6_Click
            If r(0, 3) = 0 And f(0, 3) < 9 Then Command4_Click
        End If
    End If
End Sub
Private Sub Command6_Click()
    If fl = 1 Then
        If Command6.Caption <> "F" Then
            Command6.Caption = "F"
        Else
            Command6.Caption = ""
        End If
    ElseIf Command6.Caption <> "F" Then
        r(0, 5) = 1
        If f(0, 5) = 9 Then
            Command6.Caption = "B"
            If p <> 100 Then endGame
        Else
            Command6.Caption = f(0, 5)
        End If
        If f(0, 5) = 0 Then
            If r(1, 5) = 0 And f(1, 5) < 9 Then Command16_Click
            If r(0, 6) = 0 And f(0, 6) < 9 Then Command7_Click
            If r(0, 4) = 0 And f(0, 4) < 9 Then Command5_Click
        End If
    End If
End Sub
Private Sub Command7_Click()
    If fl = 1 Then
        If Command7.Caption <> "F" Then
            Command7.Caption = "F"
        Else
            Command7.Caption = ""
        End If
    ElseIf Command7.Caption <> "F" Then
        r(0, 6) = 1
        If f(0, 6) = 9 Then
            Command7.Caption = "B"
            If p <> 100 Then endGame
        Else
            Command7.Caption = f(0, 6)
        End If
        If f(0, 6) = 0 Then
            If r(1, 6) = 0 And f(1, 6) < 9 Then Command17_Click
            If r(0, 7) = 0 And f(0, 7) < 9 Then Command8_Click
            If r(0, 5) = 0 And f(0, 5) < 9 Then Command6_Click
        End If
    End If
End Sub
Private Sub Command8_Click()
    If fl = 1 Then
        If Command8.Caption <> "F" Then
            Command8.Caption = "F"
        Else
            Command8.Caption = ""
        End If
    ElseIf Command8.Caption <> "F" Then
        r(0, 7) = 1
        If f(0, 7) = 9 Then
            Command8.Caption = "B"
            If p <> 100 Then endGame
        Else
            Command8.Caption = f(0, 7)
        End If
        If f(0, 7) = 0 Then
            If r(1, 7) = 0 And f(1, 7) < 9 Then Command18_Click
            If r(0, 8) = 0 And f(0, 8) < 9 Then Command9_Click
            If r(0, 6) = 0 And f(0, 6) < 9 Then Command7_Click
        End If
    End If
End Sub
Private Sub Command9_Click()
    If fl = 1 Then
        If Command9.Caption <> "F" Then
            Command9.Caption = "F"
        Else
            Command9.Caption = ""
        End If
    ElseIf Command9.Caption <> "F" Then
        r(0, 8) = 1
        If f(0, 8) = 9 Then
            Command9.Caption = "B"
            If p <> 100 Then endGame
        Else
            Command9.Caption = f(0, 8)
        End If
        If f(0, 8) = 0 Then
            If r(1, 8) = 0 And f(1, 8) < 9 Then Command19_Click
            If r(0, 9) = 0 And f(0, 9) < 9 Then Command10_Click
            If r(0, 7) = 0 And f(0, 7) < 9 Then Command8_Click
        End If
    End If
End Sub
Private Sub Command10_Click()
    If fl = 1 Then
        If Command10.Caption <> "F" Then
            Command10.Caption = "F"
        Else
            Command10.Caption = ""
        End If
    ElseIf Command10.Caption <> "F" Then
        r(0, 9) = 1
        If f(0, 9) = 9 Then
            Command10.Caption = "B"
            If p <> 100 Then endGame
        Else
            Command10.Caption = f(0, 9)
        End If
        If f(0, 9) = 0 Then
            If r(1, 9) = 0 And f(1, 9) < 9 Then Command20_Click
            If r(0, 8) = 0 And f(0, 8) < 9 Then Command9_Click
        End If
    End If
End Sub
Private Sub Command11_Click()
    If fl = 1 Then
        If Command11.Caption <> "F" Then
            Command11.Caption = "F"
        Else
            Command11.Caption = ""
        End If
    ElseIf Command11.Caption <> "F" Then
        r(1, 0) = 1
        If f(1, 0) = 9 Then
            Command11.Caption = "B"
            If p <> 100 Then endGame
        Else
            Command11.Caption = f(1, 0)
        End If
        If f(1, 0) = 0 Then
            If r(0, 0) = 0 And f(0, 0) < 9 Then Command1_Click
            If r(2, 0) = 0 And f(2, 0) < 9 Then Command21_Click
            If r(1, 1) = 0 And f(1, 1) < 9 Then Command12_Click
        End If
    End If
End Sub
Private Sub Command12_Click()
    If fl = 1 Then
        If Command12.Caption <> "F" Then
            Command12.Caption = "F"
        Else
            Command12.Caption = ""
        End If
    ElseIf Command12.Caption <> "F" Then
        r(1, 1) = 1
        If f(1, 1) = 9 Then
            Command12.Caption = "B"
            If p <> 100 Then endGame
        Else
            Command12.Caption = f(1, 1)
        End If
        If f(1, 1) = 0 Then
            If r(0, 1) = 0 And f(0, 1) < 9 Then Command2_Click
            If r(2, 1) = 0 And f(2, 1) < 9 Then Command22_Click
            If r(1, 2) = 0 And f(1, 2) < 9 Then Command13_Click
            If r(1, 0) = 0 And f(1, 0) < 9 Then Command11_Click
        End If
    End If
End Sub
Private Sub Command13_Click()
    If fl = 1 Then
        If Command13.Caption <> "F" Then
            Command13.Caption = "F"
        Else
            Command13.Caption = ""
        End If
    ElseIf Command13.Caption <> "F" Then
        r(1, 2) = 1
        If f(1, 2) = 9 Then
            Command13.Caption = "B"
            If p <> 100 Then endGame
        Else
            Command13.Caption = f(1, 2)
        End If
        If f(1, 2) = 0 Then
            If r(0, 2) = 0 And f(0, 2) < 9 Then Command3_Click
            If r(2, 2) = 0 And f(2, 2) < 9 Then Command23_Click
            If r(1, 3) = 0 And f(1, 3) < 9 Then Command14_Click
            If r(1, 1) = 0 And f(1, 1) < 9 Then Command12_Click
        End If
    End If
End Sub
Private Sub Command14_Click()
    If fl = 1 Then
        If Command14.Caption <> "F" Then
            Command14.Caption = "F"
        Else
            Command14.Caption = ""
        End If
    ElseIf Command14.Caption <> "F" Then
        r(1, 3) = 1
        If f(1, 3) = 9 Then
            Command14.Caption = "B"
            If p <> 100 Then endGame
        Else
            Command14.Caption = f(1, 3)
        End If
        If f(1, 3) = 0 Then
            If r(0, 3) = 0 And f(0, 3) < 9 Then Command4_Click
            If r(2, 3) = 0 And f(2, 3) < 9 Then Command24_Click
            If r(1, 4) = 0 And f(1, 4) < 9 Then Command15_Click
            If r(1, 2) = 0 And f(1, 2) < 9 Then Command13_Click
        End If
    End If
End Sub
Private Sub Command15_Click()
    If fl = 1 Then
        If Command15.Caption <> "F" Then
            Command15.Caption = "F"
        Else
            Command15.Caption = ""
        End If
    ElseIf Command15.Caption <> "F" Then
        r(1, 4) = 1
        If f(1, 4) = 9 Then
            Command15.Caption = "B"
            If p <> 100 Then endGame
        Else
            Command15.Caption = f(1, 4)
        End If
        If f(1, 4) = 0 Then
            If r(0, 4) = 0 And f(0, 4) < 9 Then Command5_Click
            If r(2, 4) = 0 And f(2, 4) < 9 Then Command25_Click
            If r(1, 5) = 0 And f(1, 5) < 9 Then Command16_Click
            If r(1, 3) = 0 And f(1, 3) < 9 Then Command14_Click
        End If
    End If
End Sub
Private Sub Command16_Click()
    If fl = 1 Then
        If Command16.Caption <> "F" Then
            Command16.Caption = "F"
        Else
            Command16.Caption = ""
        End If
    ElseIf Command16.Caption <> "F" Then
        r(1, 5) = 1
        If f(1, 5) = 9 Then
            Command16.Caption = "B"
            If p <> 100 Then endGame
        Else
            Command16.Caption = f(1, 5)
        End If
        If f(1, 5) = 0 Then
            If r(0, 5) = 0 And f(0, 5) < 9 Then Command6_Click
            If r(2, 5) = 0 And f(2, 5) < 9 Then Command26_Click
            If r(1, 6) = 0 And f(1, 6) < 9 Then Command17_Click
            If r(1, 4) = 0 And f(1, 4) < 9 Then Command15_Click
        End If
    End If
End Sub
Private Sub Command17_Click()
    If fl = 1 Then
        If Command17.Caption <> "F" Then
            Command17.Caption = "F"
        Else
            Command17.Caption = ""
        End If
    ElseIf Command17.Caption <> "F" Then
        r(1, 6) = 1
        If f(1, 6) = 9 Then
            Command17.Caption = "B"
            If p <> 100 Then endGame
        Else
            Command17.Caption = f(1, 6)
        End If
        If f(1, 6) = 0 Then
            If r(0, 6) = 0 And f(0, 6) < 9 Then Command7_Click
            If r(2, 6) = 0 And f(2, 6) < 9 Then Command27_Click
            If r(1, 7) = 0 And f(1, 7) < 9 Then Command18_Click
            If r(1, 5) = 0 And f(1, 5) < 9 Then Command16_Click
        End If
    End If
End Sub
Private Sub Command18_Click()
    If fl = 1 Then
        If Command18.Caption <> "F" Then
            Command18.Caption = "F"
        Else
            Command18.Caption = ""
        End If
    ElseIf Command18.Caption <> "F" Then
        r(1, 7) = 1
        If f(1, 7) = 9 Then
            Command18.Caption = "B"
            If p <> 100 Then endGame
        Else
            Command18.Caption = f(1, 7)
        End If
        If f(1, 7) = 0 Then
            If r(0, 7) = 0 And f(0, 7) < 9 Then Command8_Click
            If r(2, 7) = 0 And f(2, 7) < 9 Then Command28_Click
            If r(1, 8) = 0 And f(1, 8) < 9 Then Command19_Click
            If r(1, 6) = 0 And f(1, 6) < 9 Then Command17_Click
        End If
    End If
End Sub
Private Sub Command19_Click()
    If fl = 1 Then
        If Command19.Caption <> "F" Then
            Command19.Caption = "F"
        Else
            Command19.Caption = ""
        End If
    ElseIf Command19.Caption <> "F" Then
        r(1, 8) = 1
        If f(1, 8) = 9 Then
            Command19.Caption = "B"
            If p <> 100 Then endGame
        Else
            Command19.Caption = f(1, 8)
        End If
        If f(1, 8) = 0 Then
            If r(0, 8) = 0 And f(0, 8) < 9 Then Command9_Click
            If r(2, 8) = 0 And f(2, 8) < 9 Then Command29_Click
            If r(1, 9) = 0 And f(1, 9) < 9 Then Command20_Click
            If r(1, 7) = 0 And f(1, 7) < 9 Then Command18_Click
        End If
    End If
End Sub
Private Sub Command20_Click()
    If fl = 1 Then
        If Command20.Caption <> "F" Then
            Command20.Caption = "F"
        Else
            Command20.Caption = ""
        End If
    ElseIf Command20.Caption <> "F" Then
        r(1, 9) = 1
        If f(1, 9) = 9 Then
            Command20.Caption = "B"
            If p <> 100 Then endGame
        Else
            Command20.Caption = f(1, 9)
        End If
        If f(1, 9) = 0 Then
            If r(0, 9) = 0 And f(0, 9) < 9 Then Command10_Click
            If r(2, 9) = 0 And f(2, 9) < 9 Then Command30_Click
            If r(1, 8) = 0 And f(1, 8) < 9 Then Command19_Click
        End If
    End If
End Sub
Private Sub Command21_Click()
    If fl = 1 Then
        If Command21.Caption <> "F" Then
            Command21.Caption = "F"
        Else
            Command21.Caption = ""
        End If
    ElseIf Command21.Caption <> "F" Then
        r(2, 0) = 1
        If f(2, 0) = 9 Then
            Command21.Caption = "B"
            If p <> 100 Then endGame
        Else
            Command21.Caption = f(2, 0)
        End If
        If f(2, 0) = 0 Then
            If r(1, 0) = 0 And f(1, 0) < 9 Then Command11_Click
            If r(3, 0) = 0 And f(3, 0) < 9 Then Command31_Click
            If r(2, 1) = 0 And f(2, 1) < 9 Then Command22_Click
        End If
    End If
End Sub
Private Sub Command22_Click()
    If fl = 1 Then
        If Command22.Caption <> "F" Then
            Command22.Caption = "F"
        Else
            Command22.Caption = ""
        End If
    ElseIf Command22.Caption <> "F" Then
        r(2, 1) = 1
        If f(2, 1) = 9 Then
            Command22.Caption = "B"
            If p <> 100 Then endGame
        Else
            Command22.Caption = f(2, 1)
        End If
        If f(2, 1) = 0 Then
            If r(1, 1) = 0 And f(1, 1) < 9 Then Command12_Click
            If r(3, 1) = 0 And f(3, 1) < 9 Then Command32_Click
            If r(2, 2) = 0 And f(2, 2) < 9 Then Command23_Click
            If r(2, 0) = 0 And f(2, 0) < 9 Then Command21_Click
        End If
    End If
End Sub
Private Sub Command23_Click()
    If fl = 1 Then
        If Command23.Caption <> "F" Then
            Command23.Caption = "F"
        Else
            Command23.Caption = ""
        End If
    ElseIf Command23.Caption <> "F" Then
        r(2, 2) = 1
        If f(2, 2) = 9 Then
            Command23.Caption = "B"
            If p <> 100 Then endGame
        Else
            Command23.Caption = f(2, 2)
        End If
        If f(2, 2) = 0 Then
            If r(1, 2) = 0 And f(1, 2) < 9 Then Command13_Click
            If r(3, 2) = 0 And f(3, 2) < 9 Then Command33_Click
            If r(2, 3) = 0 And f(2, 3) < 9 Then Command24_Click
            If r(2, 1) = 0 And f(2, 1) < 9 Then Command22_Click
        End If
    End If
End Sub
Private Sub Command24_Click()
    If fl = 1 Then
        If Command24.Caption <> "F" Then
            Command24.Caption = "F"
        Else
            Command24.Caption = ""
        End If
    ElseIf Command24.Caption <> "F" Then
        r(2, 3) = 1
        If f(2, 3) = 9 Then
            Command24.Caption = "B"
            If p <> 100 Then endGame
        Else
            Command24.Caption = f(2, 3)
        End If
        If f(2, 3) = 0 Then
            If r(1, 3) = 0 And f(1, 3) < 9 Then Command14_Click
            If r(3, 3) = 0 And f(3, 3) < 9 Then Command34_Click
            If r(2, 4) = 0 And f(2, 4) < 9 Then Command25_Click
            If r(2, 2) = 0 And f(2, 2) < 9 Then Command23_Click
        End If
    End If
End Sub
Private Sub Command25_Click()
    If fl = 1 Then
        If Command25.Caption <> "F" Then
            Command25.Caption = "F"
        Else
            Command25.Caption = ""
        End If
    ElseIf Command25.Caption <> "F" Then
        r(2, 4) = 1
        If f(2, 4) = 9 Then
            Command25.Caption = "B"
            If p <> 100 Then endGame
        Else
            Command25.Caption = f(2, 4)
        End If
        If f(2, 4) = 0 Then
            If r(1, 4) = 0 And f(1, 4) < 9 Then Command15_Click
            If r(3, 4) = 0 And f(3, 4) < 9 Then Command35_Click
            If r(2, 5) = 0 And f(2, 5) < 9 Then Command26_Click
            If r(2, 3) = 0 And f(2, 3) < 9 Then Command24_Click
        End If
    End If
End Sub
Private Sub Command26_Click()
    If fl = 1 Then
        If Command26.Caption <> "F" Then
            Command26.Caption = "F"
        Else
            Command26.Caption = ""
        End If
    ElseIf Command26.Caption <> "F" Then
        r(2, 5) = 1
        If f(2, 5) = 9 Then
            Command26.Caption = "B"
            If p <> 100 Then endGame
        Else
            Command26.Caption = f(2, 5)
        End If
        If f(2, 5) = 0 Then
            If r(1, 5) = 0 And f(1, 5) < 9 Then Command16_Click
            If r(3, 5) = 0 And f(3, 5) < 9 Then Command36_Click
            If r(2, 6) = 0 And f(2, 6) < 9 Then Command27_Click
            If r(2, 4) = 0 And f(2, 4) < 9 Then Command25_Click
        End If
    End If
End Sub
Private Sub Command27_Click()
    If fl = 1 Then
        If Command27.Caption <> "F" Then
            Command27.Caption = "F"
        Else
            Command27.Caption = ""
        End If
    ElseIf Command27.Caption <> "F" Then
        r(2, 6) = 1
        If f(2, 6) = 9 Then
            Command27.Caption = "B"
            If p <> 100 Then endGame
        Else
            Command27.Caption = f(2, 6)
        End If
        If f(2, 6) = 0 Then
            If r(1, 6) = 0 And f(1, 6) < 9 Then Command17_Click
            If r(3, 6) = 0 And f(3, 6) < 9 Then Command37_Click
            If r(2, 7) = 0 And f(2, 7) < 9 Then Command28_Click
            If r(2, 5) = 0 And f(2, 5) < 9 Then Command26_Click
        End If
    End If
End Sub
Private Sub Command28_Click()
    If fl = 1 Then
        If Command28.Caption <> "F" Then
            Command28.Caption = "F"
        Else
            Command28.Caption = ""
        End If
    ElseIf Command28.Caption <> "F" Then
        r(2, 7) = 1
        If f(2, 7) = 9 Then
            Command28.Caption = "B"
            If p <> 100 Then endGame
        Else
            Command28.Caption = f(2, 7)
        End If
        If f(2, 7) = 0 Then
            If r(1, 7) = 0 And f(1, 7) < 9 Then Command18_Click
            If r(3, 7) = 0 And f(3, 7) < 9 Then Command38_Click
            If r(2, 8) = 0 And f(2, 8) < 9 Then Command29_Click
            If r(2, 6) = 0 And f(2, 6) < 9 Then Command27_Click
        End If
    End If
End Sub
Private Sub Command29_Click()
    If fl = 1 Then
        If Command29.Caption <> "F" Then
            Command29.Caption = "F"
        Else
            Command29.Caption = ""
        End If
    ElseIf Command29.Caption <> "F" Then
        r(2, 8) = 1
        If f(2, 8) = 9 Then
            Command29.Caption = "B"
            If p <> 100 Then endGame
        Else
            Command29.Caption = f(2, 8)
        End If
        If f(2, 8) = 0 Then
            If r(1, 8) = 0 And f(1, 8) < 9 Then Command19_Click
            If r(3, 8) = 0 And f(3, 8) < 9 Then Command39_Click
            If r(2, 9) = 0 And f(2, 9) < 9 Then Command30_Click
            If r(2, 7) = 0 And f(2, 7) < 9 Then Command28_Click
        End If
    End If
End Sub
Private Sub Command30_Click()
    If fl = 1 Then
        If Command30.Caption <> "F" Then
            Command30.Caption = "F"
        Else
            Command30.Caption = ""
        End If
    ElseIf Command30.Caption <> "F" Then
        r(2, 9) = 1
        If f(2, 9) = 9 Then
            Command30.Caption = "B"
            If p <> 100 Then endGame
        Else
            Command30.Caption = f(2, 9)
        End If
        If f(2, 9) = 0 Then
            If r(1, 9) = 0 And f(1, 9) < 9 Then Command20_Click
            If r(3, 9) = 0 And f(3, 9) < 9 Then Command40_Click
            If r(2, 8) = 0 And f(2, 8) < 9 Then Command29_Click
        End If
    End If
End Sub
Private Sub Command31_Click()
    If fl = 1 Then
        If Command31.Caption <> "F" Then
            Command31.Caption = "F"
        Else
            Command31.Caption = ""
        End If
    ElseIf Command31.Caption <> "F" Then
        r(3, 0) = 1
        If f(3, 0) = 9 Then
            Command31.Caption = "B"
            If p <> 100 Then endGame
        Else
            Command31.Caption = f(3, 0)
        End If
        If f(3, 0) = 0 Then
            If r(2, 0) = 0 And f(2, 0) < 9 Then Command21_Click
            If r(4, 0) = 0 And f(4, 0) < 9 Then Command41_Click
            If r(3, 1) = 0 And f(3, 1) < 9 Then Command32_Click
        End If
    End If
End Sub
Private Sub Command32_Click()
    If fl = 1 Then
        If Command32.Caption <> "F" Then
            Command32.Caption = "F"
        Else
            Command32.Caption = ""
        End If
    ElseIf Command32.Caption <> "F" Then
        r(3, 1) = 1
        If f(3, 1) = 9 Then
            Command32.Caption = "B"
            If p <> 100 Then endGame
        Else
            Command32.Caption = f(3, 1)
        End If
        If f(3, 1) = 0 Then
            If r(2, 1) = 0 And f(2, 1) < 9 Then Command22_Click
            If r(4, 1) = 0 And f(4, 1) < 9 Then Command42_Click
            If r(3, 2) = 0 And f(3, 2) < 9 Then Command33_Click
            If r(3, 0) = 0 And f(3, 0) < 9 Then Command31_Click
        End If
    End If
End Sub
Private Sub Command33_Click()
    If fl = 1 Then
        If Command33.Caption <> "F" Then
            Command33.Caption = "F"
        Else
            Command33.Caption = ""
        End If
    ElseIf Command33.Caption <> "F" Then
        r(3, 2) = 1
        If f(3, 2) = 9 Then
            Command33.Caption = "B"
            If p <> 100 Then endGame
        Else
            Command33.Caption = f(3, 2)
        End If
        If f(3, 2) = 0 Then
            If r(2, 2) = 0 And f(2, 2) < 9 Then Command23_Click
            If r(4, 2) = 0 And f(4, 2) < 9 Then Command43_Click
            If r(3, 3) = 0 And f(3, 3) < 9 Then Command34_Click
            If r(3, 1) = 0 And f(3, 1) < 9 Then Command32_Click
        End If
    End If
End Sub
Private Sub Command34_Click()
    If fl = 1 Then
        If Command34.Caption <> "F" Then
            Command34.Caption = "F"
        Else
            Command34.Caption = ""
        End If
    ElseIf Command34.Caption <> "F" Then
        r(3, 3) = 1
        If f(3, 3) = 9 Then
            Command34.Caption = "B"
            If p <> 100 Then endGame
        Else
            Command34.Caption = f(3, 3)
        End If
        If f(3, 3) = 0 Then
            If r(2, 3) = 0 And f(2, 3) < 9 Then Command24_Click
            If r(4, 3) = 0 And f(4, 3) < 9 Then Command44_Click
            If r(3, 4) = 0 And f(3, 4) < 9 Then Command35_Click
            If r(3, 2) = 0 And f(3, 2) < 9 Then Command33_Click
        End If
    End If
End Sub
Private Sub Command35_Click()
    If fl = 1 Then
        If Command35.Caption <> "F" Then
            Command35.Caption = "F"
        Else
            Command35.Caption = ""
        End If
    ElseIf Command35.Caption <> "F" Then
        r(3, 4) = 1
        If f(3, 4) = 9 Then
            Command35.Caption = "B"
            If p <> 100 Then endGame
        Else
            Command35.Caption = f(3, 4)
        End If
        If f(3, 4) = 0 Then
            If r(2, 4) = 0 And f(2, 4) < 9 Then Command25_Click
            If r(4, 4) = 0 And f(4, 4) < 9 Then Command45_Click
            If r(3, 5) = 0 And f(3, 5) < 9 Then Command36_Click
            If r(3, 3) = 0 And f(3, 3) < 9 Then Command34_Click
        End If
    End If
End Sub
Private Sub Command36_Click()
    If fl = 1 Then
        If Command36.Caption <> "F" Then
            Command36.Caption = "F"
        Else
            Command36.Caption = ""
        End If
    ElseIf Command36.Caption <> "F" Then
        r(3, 5) = 1
        If f(3, 5) = 9 Then
            Command36.Caption = "B"
            If p <> 100 Then endGame
        Else
            Command36.Caption = f(3, 5)
        End If
        If f(3, 5) = 0 Then
            If r(2, 5) = 0 And f(2, 5) < 9 Then Command26_Click
            If r(4, 5) = 0 And f(4, 5) < 9 Then Command46_Click
            If r(3, 6) = 0 And f(3, 6) < 9 Then Command37_Click
            If r(3, 4) = 0 And f(3, 4) < 9 Then Command35_Click
        End If
    End If
End Sub
Private Sub Command37_Click()
    If fl = 1 Then
        If Command37.Caption <> "F" Then
            Command37.Caption = "F"
        Else
            Command37.Caption = ""
        End If
    ElseIf Command37.Caption <> "F" Then
        r(3, 6) = 1
        If f(3, 6) = 9 Then
            Command37.Caption = "B"
            If p <> 100 Then endGame
        Else
            Command37.Caption = f(3, 6)
        End If
        If f(3, 6) = 0 Then
            If r(2, 6) = 0 And f(2, 6) < 9 Then Command27_Click
            If r(4, 6) = 0 And f(4, 6) < 9 Then Command47_Click
            If r(3, 7) = 0 And f(3, 7) < 9 Then Command38_Click
            If r(3, 5) = 0 And f(3, 5) < 9 Then Command36_Click
        End If
    End If
End Sub
Private Sub Command38_Click()
    If fl = 1 Then
        If Command38.Caption <> "F" Then
            Command38.Caption = "F"
        Else
            Command38.Caption = ""
        End If
    ElseIf Command38.Caption <> "F" Then
        r(3, 7) = 1
        If f(3, 7) = 9 Then
            Command38.Caption = "B"
            If p <> 100 Then endGame
        Else
            Command38.Caption = f(3, 7)
        End If
        If f(3, 7) = 0 Then
            If r(2, 7) = 0 And f(2, 7) < 9 Then Command28_Click
            If r(4, 7) = 0 And f(4, 7) < 9 Then Command48_Click
            If r(3, 8) = 0 And f(3, 8) < 9 Then Command39_Click
            If r(3, 6) = 0 And f(3, 6) < 9 Then Command37_Click
        End If
    End If
End Sub
Private Sub Command39_Click()
    If fl = 1 Then
        If Command39.Caption <> "F" Then
            Command39.Caption = "F"
        Else
            Command39.Caption = ""
        End If
    ElseIf Command39.Caption <> "F" Then
        r(3, 8) = 1
        If f(3, 8) = 9 Then
            Command39.Caption = "B"
            If p <> 100 Then endGame
        Else
            Command39.Caption = f(3, 8)
        End If
        If f(3, 8) = 0 Then
            If r(2, 8) = 0 And f(2, 8) < 9 Then Command29_Click
            If r(4, 8) = 0 And f(4, 8) < 9 Then Command49_Click
            If r(3, 9) = 0 And f(3, 9) < 9 Then Command40_Click
            If r(3, 7) = 0 And f(3, 7) < 9 Then Command38_Click
        End If
    End If
End Sub
Private Sub Command40_Click()
    If fl = 1 Then
        If Command40.Caption <> "F" Then
            Command40.Caption = "F"
        Else
            Command40.Caption = ""
        End If
    ElseIf Command40.Caption <> "F" Then
        r(3, 9) = 1
        If f(3, 9) = 9 Then
            Command40.Caption = "B"
            If p <> 100 Then endGame
        Else
            Command40.Caption = f(3, 9)
        End If
        If f(3, 9) = 0 Then
            If r(2, 9) = 0 And f(2, 9) < 9 Then Command30_Click
            If r(4, 9) = 0 And f(4, 9) < 9 Then Command50_Click
            If r(3, 8) = 0 And f(3, 8) < 9 Then Command39_Click
        End If
    End If
End Sub
Private Sub Command41_Click()
    If fl = 1 Then
        If Command41.Caption <> "F" Then
            Command41.Caption = "F"
        Else
            Command41.Caption = ""
        End If
    ElseIf Command41.Caption <> "F" Then
        r(4, 0) = 1
        If f(4, 0) = 9 Then
            Command41.Caption = "B"
            If p <> 100 Then endGame
        Else
            Command41.Caption = f(4, 0)
        End If
        If f(4, 0) = 0 Then
            If r(3, 0) = 0 And f(3, 0) < 9 Then Command31_Click
            If r(5, 0) = 0 And f(5, 0) < 9 Then Command51_Click
            If r(4, 1) = 0 And f(4, 1) < 9 Then Command42_Click
        End If
    End If
End Sub
Private Sub Command42_Click()
    If fl = 1 Then
        If Command42.Caption <> "F" Then
            Command42.Caption = "F"
        Else
            Command42.Caption = ""
        End If
    ElseIf Command42.Caption <> "F" Then
        r(4, 1) = 1
        If f(4, 1) = 9 Then
            Command42.Caption = "B"
            If p <> 100 Then endGame
        Else
            Command42.Caption = f(4, 1)
        End If
        If f(4, 1) = 0 Then
            If r(3, 1) = 0 And f(3, 1) < 9 Then Command32_Click
            If r(5, 1) = 0 And f(5, 1) < 9 Then Command52_Click
            If r(4, 2) = 0 And f(4, 2) < 9 Then Command43_Click
            If r(4, 0) = 0 And f(4, 0) < 9 Then Command41_Click
        End If
    End If
End Sub
Private Sub Command43_Click()
    If fl = 1 Then
        If Command43.Caption <> "F" Then
            Command43.Caption = "F"
        Else
            Command43.Caption = ""
        End If
    ElseIf Command43.Caption <> "F" Then
        r(4, 2) = 1
        If f(4, 2) = 9 Then
            Command43.Caption = "B"
            If p <> 100 Then endGame
        Else
            Command43.Caption = f(4, 2)
        End If
        If f(4, 2) = 0 Then
            If r(3, 2) = 0 And f(3, 2) < 9 Then Command33_Click
            If r(5, 2) = 0 And f(5, 2) < 9 Then Command53_Click
            If r(4, 3) = 0 And f(4, 3) < 9 Then Command44_Click
            If r(4, 1) = 0 And f(4, 1) < 9 Then Command42_Click
        End If
    End If
End Sub
Private Sub Command44_Click()
    If fl = 1 Then
        If Command44.Caption <> "F" Then
            Command44.Caption = "F"
        Else
            Command44.Caption = ""
        End If
    ElseIf Command44.Caption <> "F" Then
        r(4, 3) = 1
        If f(4, 3) = 9 Then
            Command44.Caption = "B"
            If p <> 100 Then endGame
        Else
            Command44.Caption = f(4, 3)
        End If
        If f(4, 3) = 0 Then
            If r(3, 3) = 0 And f(3, 3) < 9 Then Command34_Click
            If r(5, 3) = 0 And f(5, 3) < 9 Then Command54_Click
            If r(4, 4) = 0 And f(4, 4) < 9 Then Command45_Click
            If r(4, 2) = 0 And f(4, 2) < 9 Then Command43_Click
        End If
    End If
End Sub
Private Sub Command45_Click()
    If fl = 1 Then
        If Command45.Caption <> "F" Then
            Command45.Caption = "F"
        Else
            Command45.Caption = ""
        End If
    ElseIf Command45.Caption <> "F" Then
        r(4, 4) = 1
        If f(4, 4) = 9 Then
            Command45.Caption = "B"
            If p <> 100 Then endGame
        Else
            Command45.Caption = f(4, 4)
        End If
        If f(4, 4) = 0 Then
            If r(3, 4) = 0 And f(3, 4) < 9 Then Command35_Click
            If r(5, 4) = 0 And f(5, 4) < 9 Then Command55_Click
            If r(4, 5) = 0 And f(4, 5) < 9 Then Command46_Click
            If r(4, 3) = 0 And f(4, 3) < 9 Then Command44_Click
        End If
    End If
End Sub
Private Sub Command46_Click()
    If fl = 1 Then
        If Command46.Caption <> "F" Then
            Command46.Caption = "F"
        Else
            Command46.Caption = ""
        End If
    ElseIf Command46.Caption <> "F" Then
        r(4, 5) = 1
        If f(4, 5) = 9 Then
            Command46.Caption = "B"
            If p <> 100 Then endGame
        Else
            Command46.Caption = f(4, 5)
        End If
        If f(4, 5) = 0 Then
            If r(3, 5) = 0 And f(3, 5) < 9 Then Command36_Click
            If r(5, 5) = 0 And f(5, 5) < 9 Then Command56_Click
            If r(4, 6) = 0 And f(4, 6) < 9 Then Command47_Click
            If r(4, 4) = 0 And f(4, 4) < 9 Then Command45_Click
        End If
    End If
End Sub
Private Sub Command47_Click()
    If fl = 1 Then
        If Command47.Caption <> "F" Then
            Command47.Caption = "F"
        Else
            Command47.Caption = ""
        End If
    ElseIf Command47.Caption <> "F" Then
        r(4, 6) = 1
        If f(4, 6) = 9 Then
            Command47.Caption = "B"
            If p <> 100 Then endGame
        Else
            Command47.Caption = f(4, 6)
        End If
        If f(4, 6) = 0 Then
            If r(3, 6) = 0 And f(3, 6) < 9 Then Command37_Click
            If r(5, 6) = 0 And f(5, 6) < 9 Then Command57_Click
            If r(4, 7) = 0 And f(4, 7) < 9 Then Command48_Click
            If r(4, 5) = 0 And f(4, 5) < 9 Then Command46_Click
        End If
    End If
End Sub
Private Sub Command48_Click()
    If fl = 1 Then
        If Command48.Caption <> "F" Then
            Command48.Caption = "F"
        Else
            Command48.Caption = ""
        End If
    ElseIf Command48.Caption <> "F" Then
        r(4, 7) = 1
        If f(4, 7) = 9 Then
            Command48.Caption = "B"
            If p <> 100 Then endGame
        Else
            Command48.Caption = f(4, 7)
        End If
        If f(4, 7) = 0 Then
            If r(3, 7) = 0 And f(3, 7) < 9 Then Command38_Click
            If r(5, 7) = 0 And f(5, 7) < 9 Then Command58_Click
            If r(4, 8) = 0 And f(4, 8) < 9 Then Command49_Click
            If r(4, 6) = 0 And f(4, 6) < 9 Then Command47_Click
        End If
    End If
End Sub
Private Sub Command49_Click()
    If fl = 1 Then
        If Command49.Caption <> "F" Then
            Command49.Caption = "F"
        Else
            Command49.Caption = ""
        End If
    ElseIf Command49.Caption <> "F" Then
        r(4, 8) = 1
        If f(4, 8) = 9 Then
            Command49.Caption = "B"
            If p <> 100 Then endGame
        Else
            Command49.Caption = f(4, 8)
        End If
        If f(4, 8) = 0 Then
            If r(3, 8) = 0 And f(3, 8) < 9 Then Command39_Click
            If r(5, 8) = 0 And f(5, 8) < 9 Then Command59_Click
            If r(4, 9) = 0 And f(4, 9) < 9 Then Command50_Click
            If r(4, 7) = 0 And f(4, 7) < 9 Then Command48_Click
        End If
    End If
End Sub
Private Sub Command50_Click()
    If fl = 1 Then
        If Command50.Caption <> "F" Then
            Command50.Caption = "F"
        Else
            Command50.Caption = ""
        End If
    ElseIf Command50.Caption <> "F" Then
        r(4, 9) = 1
        If f(4, 9) = 9 Then
            Command50.Caption = "B"
            If p <> 100 Then endGame
        Else
            Command50.Caption = f(4, 9)
        End If
        If f(4, 9) = 0 Then
            If r(3, 9) = 0 And f(3, 9) < 9 Then Command40_Click
            If r(5, 9) = 0 And f(5, 9) < 9 Then Command60_Click
            If r(4, 8) = 0 And f(4, 8) < 9 Then Command49_Click
        End If
    End If
End Sub
Private Sub Command51_Click()
    If fl = 1 Then
        If Command51.Caption <> "F" Then
            Command51.Caption = "F"
        Else
            Command51.Caption = ""
        End If
    ElseIf Command51.Caption <> "F" Then
        r(5, 0) = 1
        If f(5, 0) = 9 Then
            Command51.Caption = "B"
            If p <> 100 Then endGame
        Else
            Command51.Caption = f(5, 0)
        End If
        If f(5, 0) = 0 Then
            If r(4, 0) = 0 And f(4, 0) < 9 Then Command41_Click
            If r(6, 0) = 0 And f(6, 0) < 9 Then Command61_Click
            If r(5, 1) = 0 And f(5, 1) < 9 Then Command52_Click
        End If
    End If
End Sub
Private Sub Command52_Click()
    If fl = 1 Then
        If Command52.Caption <> "F" Then
            Command52.Caption = "F"
        Else
            Command52.Caption = ""
        End If
    ElseIf Command52.Caption <> "F" Then
        r(5, 1) = 1
        If f(5, 1) = 9 Then
            Command52.Caption = "B"
            If p <> 100 Then endGame
        Else
            Command52.Caption = f(5, 1)
        End If
        If f(5, 1) = 0 Then
            If r(4, 1) = 0 And f(4, 1) < 9 Then Command42_Click
            If r(6, 1) = 0 And f(6, 1) < 9 Then Command62_Click
            If r(5, 2) = 0 And f(5, 2) < 9 Then Command53_Click
            If r(5, 0) = 0 And f(5, 0) < 9 Then Command51_Click
        End If
    End If
End Sub
Private Sub Command53_Click()
    If fl = 1 Then
        If Command53.Caption <> "F" Then
            Command53.Caption = "F"
        Else
            Command53.Caption = ""
        End If
    ElseIf Command53.Caption <> "F" Then
        r(5, 2) = 1
        If f(5, 2) = 9 Then
            Command53.Caption = "B"
            If p <> 100 Then endGame
        Else
            Command53.Caption = f(5, 2)
        End If
        If f(5, 2) = 0 Then
            If r(4, 2) = 0 And f(4, 2) < 9 Then Command43_Click
            If r(6, 2) = 0 And f(6, 2) < 9 Then Command63_Click
            If r(5, 3) = 0 And f(5, 3) < 9 Then Command54_Click
            If r(5, 1) = 0 And f(5, 1) < 9 Then Command52_Click
        End If
    End If
End Sub
Private Sub Command54_Click()
    If fl = 1 Then
        If Command54.Caption <> "F" Then
            Command54.Caption = "F"
        Else
            Command54.Caption = ""
        End If
    ElseIf Command54.Caption <> "F" Then
        r(5, 3) = 1
        If f(5, 3) = 9 Then
            Command54.Caption = "B"
            If p <> 100 Then endGame
        Else
            Command54.Caption = f(5, 3)
        End If
        If f(5, 3) = 0 Then
            If r(4, 3) = 0 And f(4, 3) < 9 Then Command44_Click
            If r(6, 3) = 0 And f(6, 3) < 9 Then Command64_Click
            If r(5, 4) = 0 And f(5, 4) < 9 Then Command55_Click
            If r(5, 2) = 0 And f(5, 2) < 9 Then Command53_Click
        End If
    End If
End Sub
Private Sub Command55_Click()
    If fl = 1 Then
        If Command55.Caption <> "F" Then
            Command55.Caption = "F"
        Else
            Command55.Caption = ""
        End If
    ElseIf Command55.Caption <> "F" Then
        r(5, 4) = 1
        If f(5, 4) = 9 Then
            Command55.Caption = "B"
            If p <> 100 Then endGame
        Else
            Command55.Caption = f(5, 4)
        End If
        If f(5, 4) = 0 Then
            If r(4, 4) = 0 And f(4, 4) < 9 Then Command45_Click
            If r(6, 4) = 0 And f(6, 4) < 9 Then Command65_Click
            If r(5, 5) = 0 And f(5, 5) < 9 Then Command56_Click
            If r(5, 3) = 0 And f(5, 3) < 9 Then Command54_Click
        End If
    End If
End Sub
Private Sub Command56_Click()
    If fl = 1 Then
        If Command56.Caption <> "F" Then
            Command56.Caption = "F"
        Else
            Command56.Caption = ""
        End If
    ElseIf Command56.Caption <> "F" Then
        r(5, 5) = 1
        If f(5, 5) = 9 Then
            Command56.Caption = "B"
            If p <> 100 Then endGame
        Else
            Command56.Caption = f(5, 5)
        End If
        If f(5, 5) = 0 Then
            If r(4, 5) = 0 And f(4, 5) < 9 Then Command46_Click
            If r(6, 5) = 0 And f(6, 5) < 9 Then Command66_Click
            If r(5, 6) = 0 And f(5, 6) < 9 Then Command57_Click
            If r(5, 4) = 0 And f(5, 4) < 9 Then Command55_Click
        End If
    End If
End Sub
Private Sub Command57_Click()
    If fl = 1 Then
        If Command57.Caption <> "F" Then
            Command57.Caption = "F"
        Else
            Command57.Caption = ""
        End If
    ElseIf Command57.Caption <> "F" Then
        r(5, 6) = 1
        If f(5, 6) = 9 Then
            Command57.Caption = "B"
            If p <> 100 Then endGame
        Else
            Command57.Caption = f(5, 6)
        End If
        If f(5, 6) = 0 Then
            If r(4, 6) = 0 And f(4, 6) < 9 Then Command47_Click
            If r(6, 6) = 0 And f(6, 6) < 9 Then Command67_Click
            If r(5, 7) = 0 And f(5, 7) < 9 Then Command58_Click
            If r(5, 5) = 0 And f(5, 5) < 9 Then Command56_Click
        End If
    End If
End Sub
Private Sub Command58_Click()
    If fl = 1 Then
        If Command58.Caption <> "F" Then
            Command58.Caption = "F"
        Else
            Command58.Caption = ""
        End If
    ElseIf Command58.Caption <> "F" Then
        r(5, 7) = 1
        If f(5, 7) = 9 Then
            Command58.Caption = "B"
            If p <> 100 Then endGame
        Else
            Command58.Caption = f(5, 7)
        End If
        If f(5, 7) = 0 Then
            If r(4, 7) = 0 And f(4, 7) < 9 Then Command48_Click
            If r(6, 7) = 0 And f(6, 7) < 9 Then Command68_Click
            If r(5, 8) = 0 And f(5, 8) < 9 Then Command59_Click
            If r(5, 6) = 0 And f(5, 6) < 9 Then Command57_Click
        End If
    End If
End Sub
Private Sub Command59_Click()
    If fl = 1 Then
        If Command59.Caption <> "F" Then
            Command59.Caption = "F"
        Else
            Command59.Caption = ""
        End If
    ElseIf Command59.Caption <> "F" Then
        r(5, 8) = 1
        If f(5, 8) = 9 Then
            Command59.Caption = "B"
            If p <> 100 Then endGame
        Else
            Command59.Caption = f(5, 8)
        End If
        If f(5, 8) = 0 Then
            If r(4, 8) = 0 And f(4, 8) < 9 Then Command49_Click
            If r(6, 8) = 0 And f(6, 8) < 9 Then Command69_Click
            If r(5, 9) = 0 And f(5, 9) < 9 Then Command60_Click
            If r(5, 7) = 0 And f(5, 7) < 9 Then Command58_Click
        End If
    End If
End Sub
Private Sub Command60_Click()
    If fl = 1 Then
        If Command60.Caption <> "F" Then
            Command60.Caption = "F"
        Else
            Command60.Caption = ""
        End If
    ElseIf Command60.Caption <> "F" Then
        r(5, 9) = 1
        If f(5, 9) = 9 Then
            Command60.Caption = "B"
            If p <> 100 Then endGame
        Else
            Command60.Caption = f(5, 9)
        End If
        If f(5, 9) = 0 Then
            If r(4, 9) = 0 And f(4, 9) < 9 Then Command50_Click
            If r(6, 9) = 0 And f(6, 9) < 9 Then Command70_Click
            If r(5, 8) = 0 And f(5, 8) < 9 Then Command59_Click
        End If
    End If
End Sub
Private Sub Command61_Click()
    If fl = 1 Then
        If Command61.Caption <> "F" Then
            Command61.Caption = "F"
        Else
            Command61.Caption = ""
        End If
    ElseIf Command61.Caption <> "F" Then
        r(6, 0) = 1
        If f(6, 0) = 9 Then
            Command61.Caption = "B"
            If p <> 100 Then endGame
        Else
            Command61.Caption = f(6, 0)
        End If
        If f(6, 0) = 0 Then
            If r(5, 0) = 0 And f(5, 0) < 9 Then Command51_Click
            If r(7, 0) = 0 And f(7, 0) < 9 Then Command71_Click
            If r(6, 1) = 0 And f(6, 1) < 9 Then Command62_Click
        End If
    End If
End Sub
Private Sub Command62_Click()
    If fl = 1 Then
        If Command62.Caption <> "F" Then
            Command62.Caption = "F"
        Else
            Command62.Caption = ""
        End If
    ElseIf Command62.Caption <> "F" Then
        r(6, 1) = 1
        If f(6, 1) = 9 Then
            Command62.Caption = "B"
            If p <> 100 Then endGame
        Else
            Command62.Caption = f(6, 1)
        End If
        If f(6, 1) = 0 Then
            If r(5, 1) = 0 And f(5, 1) < 9 Then Command52_Click
            If r(7, 1) = 0 And f(7, 1) < 9 Then Command72_Click
            If r(6, 2) = 0 And f(6, 2) < 9 Then Command63_Click
            If r(6, 0) = 0 And f(6, 0) < 9 Then Command61_Click
        End If
    End If
End Sub
Private Sub Command63_Click()
    If fl = 1 Then
        If Command63.Caption <> "F" Then
            Command63.Caption = "F"
        Else
            Command63.Caption = ""
        End If
    ElseIf Command63.Caption <> "F" Then
        r(6, 2) = 1
        If f(6, 2) = 9 Then
            Command63.Caption = "B"
            If p <> 100 Then endGame
        Else
            Command63.Caption = f(6, 2)
        End If
        If f(6, 2) = 0 Then
            If r(5, 2) = 0 And f(5, 2) < 9 Then Command53_Click
            If r(7, 2) = 0 And f(7, 2) < 9 Then Command73_Click
            If r(6, 3) = 0 And f(6, 3) < 9 Then Command64_Click
            If r(6, 1) = 0 And f(6, 1) < 9 Then Command62_Click
        End If
    End If
End Sub
Private Sub Command64_Click()
    If fl = 1 Then
        If Command64.Caption <> "F" Then
            Command64.Caption = "F"
        Else
            Command64.Caption = ""
        End If
    ElseIf Command64.Caption <> "F" Then
        r(6, 3) = 1
        If f(6, 3) = 9 Then
            Command64.Caption = "B"
            If p <> 100 Then endGame
        Else
            Command64.Caption = f(6, 3)
        End If
        If f(6, 3) = 0 Then
            If r(5, 3) = 0 And f(5, 3) < 9 Then Command54_Click
            If r(7, 3) = 0 And f(7, 3) < 9 Then Command74_Click
            If r(6, 4) = 0 And f(6, 4) < 9 Then Command65_Click
            If r(6, 2) = 0 And f(6, 2) < 9 Then Command63_Click
        End If
    End If
End Sub
Private Sub Command65_Click()
    If fl = 1 Then
        If Command65.Caption <> "F" Then
            Command65.Caption = "F"
        Else
            Command65.Caption = ""
        End If
    ElseIf Command65.Caption <> "F" Then
        r(6, 4) = 1
        If f(6, 4) = 9 Then
            Command65.Caption = "B"
            If p <> 100 Then endGame
        Else
            Command65.Caption = f(6, 4)
        End If
        If f(6, 4) = 0 Then
            If r(5, 4) = 0 And f(5, 4) < 9 Then Command55_Click
            If r(7, 4) = 0 And f(7, 4) < 9 Then Command75_Click
            If r(6, 5) = 0 And f(6, 5) < 9 Then Command66_Click
            If r(6, 3) = 0 And f(6, 3) < 9 Then Command64_Click
        End If
    End If
End Sub
Private Sub Command66_Click()
    If fl = 1 Then
        If Command66.Caption <> "F" Then
            Command66.Caption = "F"
        Else
            Command66.Caption = ""
        End If
    ElseIf Command66.Caption <> "F" Then
        r(6, 5) = 1
        If f(6, 5) = 9 Then
            Command66.Caption = "B"
            If p <> 100 Then endGame
        Else
            Command66.Caption = f(6, 5)
        End If
        If f(6, 5) = 0 Then
            If r(5, 5) = 0 And f(5, 5) < 9 Then Command56_Click
            If r(7, 5) = 0 And f(7, 5) < 9 Then Command76_Click
            If r(6, 6) = 0 And f(6, 6) < 9 Then Command67_Click
            If r(6, 4) = 0 And f(6, 4) < 9 Then Command65_Click
        End If
    End If
End Sub
Private Sub Command67_Click()
    If fl = 1 Then
        If Command67.Caption <> "F" Then
            Command67.Caption = "F"
        Else
            Command67.Caption = ""
        End If
    ElseIf Command67.Caption <> "F" Then
        r(6, 6) = 1
        If f(6, 6) = 9 Then
            Command67.Caption = "B"
            If p <> 100 Then endGame
        Else
            Command67.Caption = f(6, 6)
        End If
        If f(6, 6) = 0 Then
            If r(5, 6) = 0 And f(5, 6) < 9 Then Command57_Click
            If r(7, 6) = 0 And f(7, 6) < 9 Then Command77_Click
            If r(6, 7) = 0 And f(6, 7) < 9 Then Command68_Click
            If r(6, 5) = 0 And f(6, 5) < 9 Then Command66_Click
        End If
    End If
End Sub
Private Sub Command68_Click()
    If fl = 1 Then
        If Command68.Caption <> "F" Then
            Command68.Caption = "F"
        Else
            Command68.Caption = ""
        End If
    ElseIf Command68.Caption <> "F" Then
        r(6, 7) = 1
        If f(6, 7) = 9 Then
            Command68.Caption = "B"
            If p <> 100 Then endGame
        Else
            Command68.Caption = f(6, 7)
        End If
        If f(6, 7) = 0 Then
            If r(5, 7) = 0 And f(5, 7) < 9 Then Command58_Click
            If r(7, 7) = 0 And f(7, 7) < 9 Then Command78_Click
            If r(6, 8) = 0 And f(6, 8) < 9 Then Command69_Click
            If r(6, 6) = 0 And f(6, 6) < 9 Then Command67_Click
        End If
    End If
End Sub
Private Sub Command69_Click()
    If fl = 1 Then
        If Command69.Caption <> "F" Then
            Command69.Caption = "F"
        Else
            Command69.Caption = ""
        End If
    ElseIf Command69.Caption <> "F" Then
        r(6, 8) = 1
        If f(6, 8) = 9 Then
            Command69.Caption = "B"
            If p <> 100 Then endGame
        Else
            Command69.Caption = f(6, 8)
        End If
        If f(6, 8) = 0 Then
            If r(5, 8) = 0 And f(5, 8) < 9 Then Command59_Click
            If r(7, 8) = 0 And f(7, 8) < 9 Then Command79_Click
            If r(6, 9) = 0 And f(6, 9) < 9 Then Command70_Click
            If r(6, 7) = 0 And f(6, 7) < 9 Then Command68_Click
        End If
    End If
End Sub
Private Sub Command70_Click()
    If fl = 1 Then
        If Command70.Caption <> "F" Then
            Command70.Caption = "F"
        Else
            Command70.Caption = ""
        End If
    ElseIf Command70.Caption <> "F" Then
        r(6, 9) = 1
        If f(6, 9) = 9 Then
            Command70.Caption = "B"
            If p <> 100 Then endGame
        Else
            Command70.Caption = f(6, 9)
        End If
        If f(6, 9) = 0 Then
            If r(5, 9) = 0 And f(5, 9) < 9 Then Command60_Click
            If r(7, 9) = 0 And f(7, 9) < 9 Then Command80_Click
            If r(6, 8) = 0 And f(6, 8) < 9 Then Command69_Click
        End If
    End If
End Sub
Private Sub Command71_Click()
    If fl = 1 Then
        If Command71.Caption <> "F" Then
            Command71.Caption = "F"
        Else
            Command71.Caption = ""
        End If
    ElseIf Command71.Caption <> "F" Then
        r(7, 0) = 1
        If f(7, 0) = 9 Then
            Command71.Caption = "B"
            If p <> 100 Then endGame
        Else
            Command71.Caption = f(7, 0)
        End If
        If f(7, 0) = 0 Then
            If r(6, 0) = 0 And f(6, 0) < 9 Then Command61_Click
            If r(8, 0) = 0 And f(8, 0) < 9 Then Command81_Click
            If r(7, 1) = 0 And f(7, 1) < 9 Then Command72_Click
        End If
    End If
End Sub
Private Sub Command72_Click()
    If fl = 1 Then
        If Command72.Caption <> "F" Then
            Command72.Caption = "F"
        Else
            Command72.Caption = ""
        End If
    ElseIf Command72.Caption <> "F" Then
        r(7, 1) = 1
        If f(7, 1) = 9 Then
            Command72.Caption = "B"
            If p <> 100 Then endGame
        Else
            Command72.Caption = f(7, 1)
        End If
        If f(7, 1) = 0 Then
            If r(6, 1) = 0 And f(6, 1) < 9 Then Command62_Click
            If r(8, 1) = 0 And f(8, 1) < 9 Then Command82_Click
            If r(7, 2) = 0 And f(7, 2) < 9 Then Command73_Click
            If r(7, 0) = 0 And f(7, 0) < 9 Then Command71_Click
        End If
    End If
End Sub
Private Sub Command73_Click()
    If fl = 1 Then
        If Command73.Caption <> "F" Then
            Command73.Caption = "F"
        Else
            Command73.Caption = ""
        End If
    ElseIf Command73.Caption <> "F" Then
        r(7, 2) = 1
        If f(7, 2) = 9 Then
            Command73.Caption = "B"
            If p <> 100 Then endGame
        Else
            Command73.Caption = f(7, 2)
        End If
        If f(7, 2) = 0 Then
            If r(6, 2) = 0 And f(6, 2) < 9 Then Command63_Click
            If r(8, 2) = 0 And f(8, 2) < 9 Then Command83_Click
            If r(7, 3) = 0 And f(7, 3) < 9 Then Command74_Click
            If r(7, 1) = 0 And f(7, 1) < 9 Then Command72_Click
        End If
    End If
End Sub
Private Sub Command74_Click()
    If fl = 1 Then
        If Command74.Caption <> "F" Then
            Command74.Caption = "F"
        Else
            Command74.Caption = ""
        End If
    ElseIf Command74.Caption <> "F" Then
        r(7, 3) = 1
        If f(7, 3) = 9 Then
            Command74.Caption = "B"
            If p <> 100 Then endGame
        Else
            Command74.Caption = f(7, 3)
        End If
        If f(7, 3) = 0 Then
            If r(6, 3) = 0 And f(6, 3) < 9 Then Command64_Click
            If r(8, 3) = 0 And f(8, 3) < 9 Then Command84_Click
            If r(7, 4) = 0 And f(7, 4) < 9 Then Command75_Click
            If r(7, 2) = 0 And f(7, 2) < 9 Then Command73_Click
        End If
    End If
End Sub
Private Sub Command75_Click()
    If fl = 1 Then
        If Command75.Caption <> "F" Then
            Command75.Caption = "F"
        Else
            Command75.Caption = ""
        End If
    ElseIf Command75.Caption <> "F" Then
        r(7, 4) = 1
        If f(7, 4) = 9 Then
            Command75.Caption = "B"
            If p <> 100 Then endGame
        Else
            Command75.Caption = f(7, 4)
        End If
        If f(7, 4) = 0 Then
            If r(6, 4) = 0 And f(6, 4) < 9 Then Command65_Click
            If r(8, 4) = 0 And f(8, 4) < 9 Then Command85_Click
            If r(7, 5) = 0 And f(7, 5) < 9 Then Command76_Click
            If r(7, 3) = 0 And f(7, 3) < 9 Then Command74_Click
        End If
    End If
End Sub
Private Sub Command76_Click()
    If fl = 1 Then
        If Command76.Caption <> "F" Then
            Command76.Caption = "F"
        Else
            Command76.Caption = ""
        End If
    ElseIf Command76.Caption <> "F" Then
        r(7, 5) = 1
        If f(7, 5) = 9 Then
            Command76.Caption = "B"
            If p <> 100 Then endGame
        Else
            Command76.Caption = f(7, 5)
        End If
        If f(7, 5) = 0 Then
            If r(6, 5) = 0 And f(6, 5) < 9 Then Command66_Click
            If r(8, 5) = 0 And f(8, 5) < 9 Then Command86_Click
            If r(7, 6) = 0 And f(7, 6) < 9 Then Command77_Click
            If r(7, 4) = 0 And f(7, 4) < 9 Then Command75_Click
        End If
    End If
End Sub
Private Sub Command77_Click()
    If fl = 1 Then
        If Command77.Caption <> "F" Then
            Command77.Caption = "F"
        Else
            Command77.Caption = ""
        End If
    ElseIf Command77.Caption <> "F" Then
        r(7, 6) = 1
        If f(7, 6) = 9 Then
            Command77.Caption = "B"
            If p <> 100 Then endGame
        Else
            Command77.Caption = f(7, 6)
        End If
        If f(7, 6) = 0 Then
            If r(6, 6) = 0 And f(6, 6) < 9 Then Command67_Click
            If r(8, 6) = 0 And f(8, 6) < 9 Then Command87_Click
            If r(7, 7) = 0 And f(7, 7) < 9 Then Command78_Click
            If r(7, 5) = 0 And f(7, 5) < 9 Then Command76_Click
        End If
    End If
End Sub
Private Sub Command78_Click()
    If fl = 1 Then
        If Command78.Caption <> "F" Then
            Command78.Caption = "F"
        Else
            Command78.Caption = ""
        End If
    ElseIf Command78.Caption <> "F" Then
        r(7, 7) = 1
        If f(7, 7) = 9 Then
            Command78.Caption = "B"
            If p <> 100 Then endGame
        Else
            Command78.Caption = f(7, 7)
        End If
        If f(7, 7) = 0 Then
            If r(6, 7) = 0 And f(6, 7) < 9 Then Command68_Click
            If r(8, 7) = 0 And f(8, 7) < 9 Then Command88_Click
            If r(7, 8) = 0 And f(7, 8) < 9 Then Command79_Click
            If r(7, 6) = 0 And f(7, 6) < 9 Then Command77_Click
        End If
    End If
End Sub
Private Sub Command79_Click()
    If fl = 1 Then
        If Command79.Caption <> "F" Then
            Command79.Caption = "F"
        Else
            Command79.Caption = ""
        End If
    ElseIf Command79.Caption <> "F" Then
        r(7, 8) = 1
        If f(7, 8) = 9 Then
            Command79.Caption = "B"
            If p <> 100 Then endGame
        Else
            Command79.Caption = f(7, 8)
        End If
        If f(7, 8) = 0 Then
            If r(6, 8) = 0 And f(6, 8) < 9 Then Command69_Click
            If r(8, 8) = 0 And f(8, 8) < 9 Then Command89_Click
            If r(7, 9) = 0 And f(7, 9) < 9 Then Command80_Click
            If r(7, 7) = 0 And f(7, 7) < 9 Then Command78_Click
        End If
    End If
End Sub
Private Sub Command80_Click()
    If fl = 1 Then
        If Command80.Caption <> "F" Then
            Command80.Caption = "F"
        Else
            Command80.Caption = ""
        End If
    ElseIf Command80.Caption <> "F" Then
        r(7, 9) = 1
        If f(7, 9) = 9 Then
            Command80.Caption = "B"
            If p <> 100 Then endGame
        Else
            Command80.Caption = f(7, 9)
        End If
        If f(7, 9) = 0 Then
            If r(6, 9) = 0 And f(6, 9) < 9 Then Command70_Click
            If r(8, 9) = 0 And f(8, 9) < 9 Then Command90_Click
            If r(7, 8) = 0 And f(7, 8) < 9 Then Command79_Click
        End If
    End If
End Sub
Private Sub Command81_Click()
    If fl = 1 Then
        If Command81.Caption <> "F" Then
            Command81.Caption = "F"
        Else
            Command81.Caption = ""
        End If
    ElseIf Command81.Caption <> "F" Then
        r(8, 0) = 1
        If f(8, 0) = 9 Then
            Command81.Caption = "B"
            If p <> 100 Then endGame
        Else
            Command81.Caption = f(8, 0)
        End If
        If f(8, 0) = 0 Then
            If r(7, 0) = 0 And f(7, 0) < 9 Then Command71_Click
            If r(9, 0) = 0 And f(9, 0) < 9 Then Command91_Click
            If r(8, 1) = 0 And f(8, 1) < 9 Then Command82_Click
        End If
    End If
End Sub
Private Sub Command82_Click()
    If fl = 1 Then
        If Command82.Caption <> "F" Then
            Command82.Caption = "F"
        Else
            Command82.Caption = ""
        End If
    ElseIf Command82.Caption <> "F" Then
        r(8, 1) = 1
        If f(8, 1) = 9 Then
            Command82.Caption = "B"
            If p <> 100 Then endGame
        Else
            Command82.Caption = f(8, 1)
        End If
        If f(8, 1) = 0 Then
            If r(7, 1) = 0 And f(7, 1) < 9 Then Command72_Click
            If r(9, 1) = 0 And f(9, 1) < 9 Then Command92_Click
            If r(8, 2) = 0 And f(8, 2) < 9 Then Command83_Click
            If r(8, 0) = 0 And f(8, 0) < 9 Then Command81_Click
        End If
    End If
End Sub
Private Sub Command83_Click()
    If fl = 1 Then
        If Command83.Caption <> "F" Then
            Command83.Caption = "F"
        Else
            Command83.Caption = ""
        End If
    ElseIf Command83.Caption <> "F" Then
        r(8, 2) = 1
        If f(8, 2) = 9 Then
            Command83.Caption = "B"
            If p <> 100 Then endGame
        Else
            Command83.Caption = f(8, 2)
        End If
        If f(8, 2) = 0 Then
            If r(7, 2) = 0 And f(7, 2) < 9 Then Command73_Click
            If r(9, 2) = 0 And f(9, 2) < 9 Then Command93_Click
            If r(8, 3) = 0 And f(8, 3) < 9 Then Command84_Click
            If r(8, 1) = 0 And f(8, 1) < 9 Then Command82_Click
        End If
    End If
End Sub
Private Sub Command84_Click()
    If fl = 1 Then
        If Command84.Caption <> "F" Then
            Command84.Caption = "F"
        Else
            Command84.Caption = ""
        End If
    ElseIf Command84.Caption <> "F" Then
        r(8, 3) = 1
        If f(8, 3) = 9 Then
            Command84.Caption = "B"
            If p <> 100 Then endGame
        Else
            Command84.Caption = f(8, 3)
        End If
        If f(8, 3) = 0 Then
            If r(7, 3) = 0 And f(7, 3) < 9 Then Command74_Click
            If r(9, 3) = 0 And f(9, 3) < 9 Then Command94_Click
            If r(8, 4) = 0 And f(8, 4) < 9 Then Command85_Click
            If r(8, 2) = 0 And f(8, 2) < 9 Then Command83_Click
        End If
    End If
End Sub
Private Sub Command85_Click()
    If fl = 1 Then
        If Command85.Caption <> "F" Then
            Command85.Caption = "F"
        Else
            Command85.Caption = ""
        End If
    ElseIf Command85.Caption <> "F" Then
        r(8, 4) = 1
        If f(8, 4) = 9 Then
            Command85.Caption = "B"
            If p <> 100 Then endGame
        Else
            Command85.Caption = f(8, 4)
        End If
        If f(8, 4) = 0 Then
            If r(7, 4) = 0 And f(7, 4) < 9 Then Command75_Click
            If r(9, 4) = 0 And f(9, 4) < 9 Then Command95_Click
            If r(8, 5) = 0 And f(8, 5) < 9 Then Command86_Click
            If r(8, 3) = 0 And f(8, 3) < 9 Then Command84_Click
        End If
    End If
End Sub
Private Sub Command86_Click()
    If fl = 1 Then
        If Command86.Caption <> "F" Then
            Command86.Caption = "F"
        Else
            Command86.Caption = ""
        End If
    ElseIf Command86.Caption <> "F" Then
        r(8, 5) = 1
        If f(8, 5) = 9 Then
            Command86.Caption = "B"
            If p <> 100 Then endGame
        Else
            Command86.Caption = f(8, 5)
        End If
        If f(8, 5) = 0 Then
            If r(7, 5) = 0 And f(7, 5) < 9 Then Command76_Click
            If r(9, 5) = 0 And f(9, 5) < 9 Then Command96_Click
            If r(8, 6) = 0 And f(8, 6) < 9 Then Command87_Click
            If r(8, 4) = 0 And f(8, 4) < 9 Then Command85_Click
        End If
    End If
End Sub
Private Sub Command87_Click()
    If fl = 1 Then
        If Command87.Caption <> "F" Then
            Command87.Caption = "F"
        Else
            Command87.Caption = ""
        End If
    ElseIf Command87.Caption <> "F" Then
        r(8, 6) = 1
        If f(8, 6) = 9 Then
            Command87.Caption = "B"
            If p <> 100 Then endGame
        Else
            Command87.Caption = f(8, 6)
        End If
        If f(8, 6) = 0 Then
            If r(7, 6) = 0 And f(7, 6) < 9 Then Command77_Click
            If r(9, 6) = 0 And f(9, 6) < 9 Then Command97_Click
            If r(8, 7) = 0 And f(8, 7) < 9 Then Command88_Click
            If r(8, 5) = 0 And f(8, 5) < 9 Then Command86_Click
        End If
    End If
End Sub
Private Sub Command88_Click()
    If fl = 1 Then
        If Command88.Caption <> "F" Then
            Command88.Caption = "F"
        Else
            Command88.Caption = ""
        End If
    ElseIf Command88.Caption <> "F" Then
        r(8, 7) = 1
        If f(8, 7) = 9 Then
            Command88.Caption = "B"
            If p <> 100 Then endGame
        Else
            Command88.Caption = f(8, 7)
        End If
        If f(8, 7) = 0 Then
            If r(7, 7) = 0 And f(7, 7) < 9 Then Command78_Click
            If r(9, 7) = 0 And f(9, 7) < 9 Then Command98_Click
            If r(8, 8) = 0 And f(8, 8) < 9 Then Command89_Click
            If r(8, 6) = 0 And f(8, 6) < 9 Then Command87_Click
        End If
    End If
End Sub
Private Sub Command89_Click()
    If fl = 1 Then
        If Command89.Caption <> "F" Then
            Command89.Caption = "F"
        Else
            Command89.Caption = ""
        End If
    ElseIf Command89.Caption <> "F" Then
        r(8, 8) = 1
        If f(8, 8) = 9 Then
            Command89.Caption = "B"
            If p <> 100 Then endGame
        Else
            Command89.Caption = f(8, 8)
        End If
        If f(8, 8) = 0 Then
            If r(7, 8) = 0 And f(7, 8) < 9 Then Command79_Click
            If r(9, 8) = 0 And f(9, 8) < 9 Then Command99_Click
            If r(8, 9) = 0 And f(8, 9) < 9 Then Command90_Click
            If r(8, 7) = 0 And f(8, 7) < 9 Then Command88_Click
        End If
    End If
End Sub
Private Sub Command90_Click()
    If fl = 1 Then
        If Command90.Caption <> "F" Then
            Command90.Caption = "F"
        Else
            Command90.Caption = ""
        End If
    ElseIf Command90.Caption <> "F" Then
        r(8, 9) = 1
        If f(8, 9) = 9 Then
            Command90.Caption = "B"
            If p <> 100 Then endGame
        Else
            Command90.Caption = f(8, 9)
        End If
        If f(8, 9) = 0 Then
            If r(7, 9) = 0 And f(7, 9) < 9 Then Command80_Click
            If r(9, 9) = 0 And f(9, 9) < 9 Then Command100_Click
            If r(8, 8) = 0 And f(8, 8) < 9 Then Command89_Click
        End If
    End If
End Sub
Private Sub Command91_Click()
    If fl = 1 Then
        If Command91.Caption <> "F" Then
            Command91.Caption = "F"
        Else
            Command91.Caption = ""
        End If
    ElseIf Command91.Caption <> "F" Then
        r(9, 0) = 1
        If f(9, 0) = 9 Then
            Command91.Caption = "B"
            If p <> 100 Then endGame
        Else
            Command91.Caption = f(9, 0)
        End If
        If f(9, 0) = 0 Then
            If r(8, 0) = 0 And f(8, 0) < 9 Then Command81_Click
            If r(9, 1) = 0 And f(9, 1) < 9 Then Command92_Click
        End If
    End If
End Sub
Private Sub Command92_Click()
    If fl = 1 Then
        If Command92.Caption <> "F" Then
            Command92.Caption = "F"
        Else
            Command92.Caption = ""
        End If
    ElseIf Command92.Caption <> "F" Then
        r(9, 1) = 1
        If f(9, 1) = 9 Then
            Command92.Caption = "B"
            If p <> 100 Then endGame
        Else
            Command92.Caption = f(9, 1)
        End If
        If f(9, 1) = 0 Then
            If r(8, 1) = 0 And f(8, 1) < 9 Then Command82_Click
            If r(9, 2) = 0 And f(9, 2) < 9 Then Command93_Click
            If r(9, 0) = 0 And f(9, 0) < 9 Then Command91_Click
        End If
    End If
End Sub
Private Sub Command93_Click()
    If fl = 1 Then
        If Command93.Caption <> "F" Then
            Command93.Caption = "F"
        Else
            Command93.Caption = ""
        End If
    ElseIf Command93.Caption <> "F" Then
        r(9, 2) = 1
        If f(9, 2) = 9 Then
            Command93.Caption = "B"
            If p <> 100 Then endGame
        Else
            Command93.Caption = f(9, 2)
        End If
        If f(9, 2) = 0 Then
            If r(8, 2) = 0 And f(8, 2) < 9 Then Command83_Click
            If r(9, 3) = 0 And f(9, 3) < 9 Then Command94_Click
            If r(9, 1) = 0 And f(9, 1) < 9 Then Command92_Click
        End If
    End If
End Sub
Private Sub Command94_Click()
    If fl = 1 Then
        If Command94.Caption <> "F" Then
            Command94.Caption = "F"
        Else
            Command94.Caption = ""
        End If
    ElseIf Command94.Caption <> "F" Then
        r(9, 3) = 1
        If f(9, 3) = 9 Then
            Command94.Caption = "B"
            If p <> 100 Then endGame
        Else
            Command94.Caption = f(9, 3)
        End If
        If f(9, 3) = 0 Then
            If r(8, 3) = 0 And f(8, 3) < 9 Then Command84_Click
            If r(9, 4) = 0 And f(9, 4) < 9 Then Command95_Click
            If r(9, 2) = 0 And f(9, 2) < 9 Then Command93_Click
        End If
    End If
End Sub
Private Sub Command95_Click()
    If fl = 1 Then
        If Command95.Caption <> "F" Then
            Command95.Caption = "F"
        Else
            Command95.Caption = ""
        End If
    ElseIf Command95.Caption <> "F" Then
        r(9, 4) = 1
        If f(9, 4) = 9 Then
            Command95.Caption = "B"
            If p <> 100 Then endGame
        Else
            Command95.Caption = f(9, 4)
        End If
        If f(9, 4) = 0 Then
            If r(8, 4) = 0 And f(8, 4) < 9 Then Command85_Click
            If r(9, 5) = 0 And f(9, 5) < 9 Then Command96_Click
            If r(9, 3) = 0 And f(9, 3) < 9 Then Command94_Click
        End If
    End If
End Sub
Private Sub Command96_Click()
    If fl = 1 Then
        If Command96.Caption <> "F" Then
            Command96.Caption = "F"
        Else
            Command96.Caption = ""
        End If
    ElseIf Command96.Caption <> "F" Then
        r(9, 5) = 1
        If f(9, 5) = 9 Then
            Command96.Caption = "B"
            If p <> 100 Then endGame
        Else
            Command96.Caption = f(9, 5)
        End If
        If f(9, 5) = 0 Then
            If r(8, 5) = 0 And f(8, 5) < 9 Then Command86_Click
            If r(9, 6) = 0 And f(9, 6) < 9 Then Command97_Click
            If r(9, 4) = 0 And f(9, 4) < 9 Then Command95_Click
        End If
    End If
End Sub
Private Sub Command97_Click()
    If fl = 1 Then
        If Command97.Caption <> "F" Then
            Command97.Caption = "F"
        Else
            Command97.Caption = ""
        End If
    ElseIf Command97.Caption <> "F" Then
        r(9, 6) = 1
        If f(9, 6) = 9 Then
            Command97.Caption = "B"
            If p <> 100 Then endGame
        Else
            Command97.Caption = f(9, 6)
        End If
        If f(9, 6) = 0 Then
            If r(8, 6) = 0 And f(8, 6) < 9 Then Command87_Click
            If r(9, 7) = 0 And f(9, 7) < 9 Then Command98_Click
            If r(9, 5) = 0 And f(9, 5) < 9 Then Command96_Click
        End If
    End If
End Sub
Private Sub Command98_Click()
    If fl = 1 Then
        If Command98.Caption <> "F" Then
            Command98.Caption = "F"
        Else
            Command98.Caption = ""
        End If
    ElseIf Command98.Caption <> "F" Then
        r(9, 7) = 1
        If f(9, 7) = 9 Then
            Command98.Caption = "B"
            If p <> 100 Then endGame
        Else
            Command98.Caption = f(9, 7)
        End If
        If f(9, 7) = 0 Then
            If r(8, 7) = 0 And f(8, 7) < 9 Then Command88_Click
            If r(9, 8) = 0 And f(9, 8) < 9 Then Command99_Click
            If r(9, 6) = 0 And f(9, 6) < 9 Then Command97_Click
        End If
    End If
End Sub
Private Sub Command99_Click()
    If fl = 1 Then
        If Command99.Caption <> "F" Then
            Command99.Caption = "F"
        Else
            Command99.Caption = ""
        End If
    ElseIf Command99.Caption <> "F" Then
        r(9, 8) = 1
        If f(9, 8) = 9 Then
            Command99.Caption = "B"
            If p <> 100 Then endGame
        Else
            Command99.Caption = f(9, 8)
        End If
        If f(9, 8) = 0 Then
            If r(8, 8) = 0 And f(8, 8) < 9 Then Command89_Click
            If r(9, 9) = 0 And f(9, 9) < 9 Then Command100_Click
            If r(9, 7) = 0 And f(9, 7) < 9 Then Command98_Click
        End If
    End If
End Sub
Private Sub Command100_Click()
    If fl = 1 Then
        If Command100.Caption <> "F" Then
            Command100.Caption = "F"
        Else
            Command100.Caption = ""
        End If
    ElseIf Command100.Caption <> "F" Then
        r(9, 9) = 1
        If f(9, 9) = 9 Then
            Command100.Caption = "B"
            If p <> 100 Then endGame
        Else
            Command100.Caption = f(9, 9)
        End If
        If f(9, 9) = 0 Then
            If r(8, 9) = 0 And f(8, 9) < 9 Then Command90_Click
            If r(9, 8) = 0 And f(9, 8) < 9 Then Command99_Click
        End If
    End If
End Sub

