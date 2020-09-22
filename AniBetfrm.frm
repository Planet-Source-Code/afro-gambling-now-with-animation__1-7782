VERSION 5.00
Begin VB.Form Gamble 
   Caption         =   "Gamble"
   ClientHeight    =   7590
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7005
   FillStyle       =   2  'Horizontal Line
   Icon            =   "AniBetfrm.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7590
   ScaleWidth      =   7005
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   5520
      Top             =   1080
   End
   Begin VB.Timer tmrBet 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   5520
      Top             =   720
   End
   Begin VB.Frame Frame1 
      Height          =   4095
      Left            =   0
      TabIndex        =   28
      Top             =   3360
      Width           =   6975
      Begin VB.Image Image8 
         Height          =   375
         Left            =   2400
         Picture         =   "AniBetfrm.frx":0442
         Stretch         =   -1  'True
         Top             =   3570
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Image Image7 
         Height          =   375
         Left            =   2400
         Picture         =   "AniBetfrm.frx":0884
         Stretch         =   -1  'True
         Top             =   3120
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Image Image6 
         Height          =   375
         Left            =   2400
         Picture         =   "AniBetfrm.frx":0CC6
         Stretch         =   -1  'True
         Top             =   2610
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Image Image5 
         Height          =   375
         Left            =   2400
         Picture         =   "AniBetfrm.frx":1108
         Stretch         =   -1  'True
         Top             =   2160
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Image Image4 
         Height          =   375
         Left            =   2400
         Picture         =   "AniBetfrm.frx":154A
         Stretch         =   -1  'True
         Top             =   1680
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Image Image3 
         Height          =   375
         Left            =   2400
         Picture         =   "AniBetfrm.frx":198C
         Stretch         =   -1  'True
         Top             =   1200
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Image Image2 
         Height          =   375
         Left            =   2400
         Picture         =   "AniBetfrm.frx":1DCE
         Stretch         =   -1  'True
         Top             =   720
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Image Image1 
         Height          =   375
         Left            =   2400
         Picture         =   "AniBetfrm.frx":2210
         Stretch         =   -1  'True
         Top             =   150
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Shape Shape8 
         FillColor       =   &H00FFC0FF&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   120
         Top             =   3600
         Width           =   375
      End
      Begin VB.Shape Shape7 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   120
         Top             =   3120
         Width           =   375
      End
      Begin VB.Shape Shape6 
         FillColor       =   &H00FFFF00&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   120
         Top             =   2640
         Width           =   375
      End
      Begin VB.Shape Shape5 
         FillColor       =   &H0000FF00&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   120
         Top             =   2160
         Width           =   375
      End
      Begin VB.Shape Shape4 
         FillColor       =   &H00FF00FF&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   120
         Top             =   1680
         Width           =   375
      End
      Begin VB.Shape Shape3 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   120
         Top             =   1200
         Width           =   375
      End
      Begin VB.Shape Shape2 
         FillColor       =   &H00C00000&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   120
         Top             =   720
         Width           =   375
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H0000FFFF&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   120
         Top             =   240
         Width           =   375
      End
      Begin VB.Line Line48 
         BorderColor     =   &H80000005&
         BorderWidth     =   5
         X1              =   6000
         X2              =   6720
         Y1              =   3960
         Y2              =   3960
      End
      Begin VB.Line Line47 
         BorderColor     =   &H80000005&
         BorderWidth     =   5
         X1              =   6000
         X2              =   6720
         Y1              =   3480
         Y2              =   3480
      End
      Begin VB.Line Line46 
         BorderColor     =   &H80000005&
         BorderWidth     =   5
         X1              =   6000
         X2              =   6720
         Y1              =   3000
         Y2              =   3000
      End
      Begin VB.Line Line45 
         BorderColor     =   &H80000005&
         BorderWidth     =   5
         X1              =   6000
         X2              =   6720
         Y1              =   2520
         Y2              =   2520
      End
      Begin VB.Line Line44 
         BorderColor     =   &H80000005&
         BorderWidth     =   5
         X1              =   6000
         X2              =   6720
         Y1              =   2040
         Y2              =   2040
      End
      Begin VB.Line Line43 
         BorderColor     =   &H80000005&
         BorderWidth     =   5
         X1              =   6000
         X2              =   6720
         Y1              =   1560
         Y2              =   1560
      End
      Begin VB.Line Line42 
         BorderColor     =   &H80000005&
         BorderWidth     =   5
         X1              =   6000
         X2              =   6720
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Line Line41 
         BorderColor     =   &H80000005&
         BorderWidth     =   5
         X1              =   6000
         X2              =   6720
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Line Line40 
         BorderColor     =   &H80000005&
         BorderWidth     =   5
         X1              =   4680
         X2              =   5400
         Y1              =   3960
         Y2              =   3960
      End
      Begin VB.Line Line39 
         BorderColor     =   &H80000005&
         BorderWidth     =   5
         X1              =   4680
         X2              =   5400
         Y1              =   3480
         Y2              =   3480
      End
      Begin VB.Line Line38 
         BorderColor     =   &H80000005&
         BorderWidth     =   5
         X1              =   4680
         X2              =   5400
         Y1              =   3000
         Y2              =   3000
      End
      Begin VB.Line Line37 
         BorderColor     =   &H80000005&
         BorderWidth     =   5
         X1              =   4680
         X2              =   5400
         Y1              =   2520
         Y2              =   2520
      End
      Begin VB.Line Line36 
         BorderColor     =   &H80000005&
         BorderWidth     =   5
         X1              =   4680
         X2              =   5400
         Y1              =   2040
         Y2              =   2040
      End
      Begin VB.Line Line35 
         BorderColor     =   &H80000005&
         BorderWidth     =   5
         X1              =   4680
         X2              =   5400
         Y1              =   1560
         Y2              =   1560
      End
      Begin VB.Line Line34 
         BorderColor     =   &H80000005&
         BorderWidth     =   5
         X1              =   4680
         X2              =   5400
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Line Line33 
         BorderColor     =   &H80000005&
         BorderWidth     =   5
         X1              =   4680
         X2              =   5400
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Line Line32 
         BorderColor     =   &H80000005&
         BorderWidth     =   5
         X1              =   3480
         X2              =   4200
         Y1              =   3960
         Y2              =   3960
      End
      Begin VB.Line Line31 
         BorderColor     =   &H80000005&
         BorderWidth     =   5
         X1              =   2400
         X2              =   3120
         Y1              =   3960
         Y2              =   3960
      End
      Begin VB.Line Line30 
         BorderColor     =   &H80000005&
         BorderWidth     =   5
         X1              =   1320
         X2              =   2040
         Y1              =   3960
         Y2              =   3960
      End
      Begin VB.Line Line29 
         BorderColor     =   &H80000005&
         BorderWidth     =   5
         X1              =   240
         X2              =   960
         Y1              =   3960
         Y2              =   3960
      End
      Begin VB.Line Line28 
         BorderColor     =   &H80000005&
         BorderWidth     =   5
         X1              =   3480
         X2              =   4200
         Y1              =   3480
         Y2              =   3480
      End
      Begin VB.Line Line27 
         BorderColor     =   &H80000005&
         BorderWidth     =   5
         X1              =   2400
         X2              =   3120
         Y1              =   3480
         Y2              =   3480
      End
      Begin VB.Line Line26 
         BorderColor     =   &H80000005&
         BorderWidth     =   5
         X1              =   1320
         X2              =   2040
         Y1              =   3480
         Y2              =   3480
      End
      Begin VB.Line Line25 
         BorderColor     =   &H80000005&
         BorderWidth     =   5
         X1              =   240
         X2              =   960
         Y1              =   3480
         Y2              =   3480
      End
      Begin VB.Line Line24 
         BorderColor     =   &H80000005&
         BorderWidth     =   5
         X1              =   3480
         X2              =   4200
         Y1              =   3000
         Y2              =   3000
      End
      Begin VB.Line Line23 
         BorderColor     =   &H80000005&
         BorderWidth     =   5
         X1              =   2400
         X2              =   3120
         Y1              =   3000
         Y2              =   3000
      End
      Begin VB.Line Line22 
         BorderColor     =   &H80000005&
         BorderWidth     =   5
         X1              =   1320
         X2              =   2040
         Y1              =   3000
         Y2              =   3000
      End
      Begin VB.Line Line21 
         BorderColor     =   &H80000005&
         BorderWidth     =   5
         X1              =   240
         X2              =   960
         Y1              =   3000
         Y2              =   3000
      End
      Begin VB.Line Line20 
         BorderColor     =   &H80000005&
         BorderWidth     =   5
         X1              =   3480
         X2              =   4200
         Y1              =   2520
         Y2              =   2520
      End
      Begin VB.Line Line19 
         BorderColor     =   &H80000005&
         BorderWidth     =   5
         X1              =   2400
         X2              =   3120
         Y1              =   2520
         Y2              =   2520
      End
      Begin VB.Line Line18 
         BorderColor     =   &H80000005&
         BorderWidth     =   5
         X1              =   1320
         X2              =   2040
         Y1              =   2520
         Y2              =   2520
      End
      Begin VB.Line Line17 
         BorderColor     =   &H80000005&
         BorderWidth     =   5
         X1              =   240
         X2              =   960
         Y1              =   2520
         Y2              =   2520
      End
      Begin VB.Line Line16 
         BorderColor     =   &H80000005&
         BorderWidth     =   5
         X1              =   3480
         X2              =   4200
         Y1              =   2040
         Y2              =   2040
      End
      Begin VB.Line Line15 
         BorderColor     =   &H80000005&
         BorderWidth     =   5
         X1              =   2400
         X2              =   3120
         Y1              =   2040
         Y2              =   2040
      End
      Begin VB.Line Line14 
         BorderColor     =   &H80000005&
         BorderWidth     =   5
         X1              =   1320
         X2              =   2040
         Y1              =   2040
         Y2              =   2040
      End
      Begin VB.Line Line13 
         BorderColor     =   &H80000005&
         BorderWidth     =   5
         X1              =   240
         X2              =   960
         Y1              =   2040
         Y2              =   2040
      End
      Begin VB.Line Line12 
         BorderColor     =   &H80000005&
         BorderWidth     =   5
         X1              =   3480
         X2              =   4200
         Y1              =   1560
         Y2              =   1560
      End
      Begin VB.Line Line11 
         BorderColor     =   &H80000005&
         BorderWidth     =   5
         X1              =   2400
         X2              =   3120
         Y1              =   1560
         Y2              =   1560
      End
      Begin VB.Line Line10 
         BorderColor     =   &H80000005&
         BorderWidth     =   5
         X1              =   1320
         X2              =   2040
         Y1              =   1560
         Y2              =   1560
      End
      Begin VB.Line Line9 
         BorderColor     =   &H80000005&
         BorderWidth     =   5
         X1              =   240
         X2              =   960
         Y1              =   1560
         Y2              =   1560
      End
      Begin VB.Line Line8 
         BorderColor     =   &H80000005&
         BorderWidth     =   5
         X1              =   3480
         X2              =   4200
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Line Line7 
         BorderColor     =   &H80000005&
         BorderWidth     =   5
         X1              =   2400
         X2              =   3120
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Line Line6 
         BorderColor     =   &H80000005&
         BorderWidth     =   5
         X1              =   1320
         X2              =   2040
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Line Line5 
         BorderColor     =   &H80000005&
         BorderWidth     =   5
         X1              =   240
         X2              =   960
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Line Line4 
         BorderColor     =   &H80000005&
         BorderWidth     =   5
         X1              =   3480
         X2              =   4200
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Line Line3 
         BorderColor     =   &H80000005&
         BorderWidth     =   5
         X1              =   2400
         X2              =   3120
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000005&
         BorderWidth     =   5
         X1              =   1320
         X2              =   2040
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000005&
         BorderWidth     =   5
         X1              =   240
         X2              =   960
         Y1              =   600
         Y2              =   600
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Random Car"
      Height          =   375
      Left            =   2640
      TabIndex        =   27
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton Lifeline 
      Caption         =   "Lifeline"
      Enabled         =   0   'False
      Height          =   375
      Left            =   960
      TabIndex        =   26
      Top             =   2160
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Random"
      Height          =   375
      Left            =   960
      TabIndex        =   25
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton Place 
      Caption         =   "Placings"
      Height          =   375
      Left            =   960
      TabIndex        =   16
      Top             =   1800
      Width           =   855
   End
   Begin VB.CommandButton Quit 
      Caption         =   "Quit"
      Height          =   375
      Left            =   960
      TabIndex        =   15
      Top             =   1440
      Width           =   855
   End
   Begin VB.OptionButton GT1 
      Caption         =   "GT-ONE"
      Height          =   195
      Left            =   2640
      TabIndex        =   12
      Top             =   2280
      Width           =   975
   End
   Begin VB.OptionButton XJR 
      Caption         =   "XJR-15"
      Height          =   195
      Left            =   2640
      TabIndex        =   11
      Top             =   2040
      Width           =   975
   End
   Begin VB.OptionButton Storm 
      Caption         =   "Storm V12"
      Height          =   195
      Left            =   2640
      TabIndex        =   10
      Top             =   1800
      Width           =   1215
   End
   Begin VB.OptionButton XJ220 
      Caption         =   "XJ220"
      Height          =   195
      Left            =   2640
      TabIndex        =   9
      Top             =   1560
      Width           =   975
   End
   Begin VB.OptionButton Speed12 
      Caption         =   "Speed 12"
      Height          =   195
      Left            =   2640
      TabIndex        =   8
      Top             =   1320
      Width           =   1215
   End
   Begin VB.OptionButton F1 
      Caption         =   "McLaren F1"
      Height          =   195
      Left            =   2640
      TabIndex        =   7
      Top             =   1080
      Width           =   1215
   End
   Begin VB.OptionButton F50 
      Caption         =   "F50"
      Height          =   195
      Left            =   2640
      TabIndex        =   6
      Top             =   840
      Width           =   975
   End
   Begin VB.OptionButton Corvette 
      Caption         =   "Corvette"
      Height          =   195
      Left            =   2640
      TabIndex        =   5
      Top             =   600
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Bet"
      Height          =   375
      Left            =   960
      TabIndex        =   4
      Top             =   1080
      Width           =   855
   End
   Begin VB.TextBox Bet 
      Height          =   375
      Left            =   840
      TabIndex        =   2
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Prediction"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3840
      TabIndex        =   30
      Top             =   240
      Width           =   885
   End
   Begin VB.Label lblWin 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "No-one is racing!!!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3960
      TabIndex        =   29
      Top             =   2640
      Width           =   2775
   End
   Begin VB.Shape Shape16 
      FillColor       =   &H00FFC0FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   4440
      Top             =   2280
      Width           =   375
   End
   Begin VB.Shape Shape15 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   4440
      Top             =   2040
      Width           =   375
   End
   Begin VB.Shape Shape14 
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   4440
      Top             =   1800
      Width           =   375
   End
   Begin VB.Shape Shape13 
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   4440
      Top             =   1560
      Width           =   375
   End
   Begin VB.Shape Shape12 
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   4440
      Top             =   1320
      Width           =   375
   End
   Begin VB.Shape Shape11 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   4440
      Top             =   1080
      Width           =   375
   End
   Begin VB.Shape Shape10 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   4440
      Top             =   840
      Width           =   375
   End
   Begin VB.Shape Shape9 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   4440
      Top             =   600
      Width           =   375
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4170
      TabIndex        =   24
      Top             =   2280
      Width           =   135
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4170
      TabIndex        =   23
      Top             =   2040
      Width           =   135
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4170
      TabIndex        =   22
      Top             =   1800
      Width           =   135
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4170
      TabIndex        =   21
      Top             =   1560
      Width           =   135
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4170
      TabIndex        =   20
      Top             =   1320
      Width           =   135
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4170
      TabIndex        =   19
      Top             =   1080
      Width           =   135
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4170
      TabIndex        =   18
      Top             =   840
      Width           =   135
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4170
      TabIndex        =   17
      Top             =   600
      Width           =   135
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Days left"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2160
      TabIndex        =   14
      Top             =   2640
      Width           =   1185
   End
   Begin VB.Label Days 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "30"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1680
      TabIndex        =   13
      Top             =   2640
      Width           =   345
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "$"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   600
      TabIndex        =   3
      Top             =   600
      Width           =   180
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "$"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   720
      TabIndex        =   1
      Top             =   2520
      Width           =   180
   End
   Begin VB.Label Money 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   840
      TabIndex        =   0
      Top             =   2520
      Width           =   510
   End
End
Attribute VB_Name = "Gamble"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub BOMB()
Dim DeadCar
DeadCar = Int((18 * Rnd) + 1)

If DeadCar = 3 Or DeadCar = 5 Then
    Image3.Visible = True
ElseIf DeadCar = 9 Then
    Image2.Visible = True
ElseIf DeadCar = 13 Then
    Image1.Visible = True
ElseIf DeadCar = 15 Then
    Image7.Visible = True
ElseIf DeadCar = 11 Then
    Image4.Visible = True
ElseIf DeadCar = 18 Then
    Image5.Visible = True
ElseIf DeadCar = 1 Then
    Image6.Visible = True
ElseIf DeadCar = 14 Then
    Image8.Visible = True
End If
End Sub


Private Sub CorvetteSub()

    MsgBox "Corvette Won"
    pLACINGS.Corvette.Caption = pLACINGS.Corvette.Caption + 1
    Days.Caption = Days.Caption - 1
If Corvette = False Then
    Money.Caption = Money.Caption - Bet.Text
ElseIf Corvette = True Then
    Money.Caption = Money.Caption + Val(Bet.Text) * 8
    pLACINGS.BetsWon = pLACINGS.BetsWon + 1
    pLACINGS.VetteBet = pLACINGS.VetteBet + 1
    
If Money.Caption = 0 Then
    MsgBox "You are bankrupt after " & 30 - Days.Caption & " Days. Better luck next time"
Unload pLACINGS
Unload Gamble
End
Else
End If
End If

Corvette.Enabled = True
F50.Enabled = True
F1.Enabled = True
GT1.Enabled = True
Storm.Enabled = True
Speed12.Enabled = True
XJR.Enabled = True
XJ220.Enabled = True

End Sub



Private Sub F1Sub()

    MsgBox "The McLaren F1 Won the race!"
    pLACINGS.F1.Caption = pLACINGS.F1.Caption + 1
    Days.Caption = Days.Caption - 1
If F1 = False Then
    Money.Caption = Money.Caption - Bet.Text
ElseIf F1 = True Then
    Money.Caption = Money.Caption + Val(Bet.Text)
    pLACINGS.BetsWon = pLACINGS.BetsWon + 1
    pLACINGS.F1Bet = pLACINGS.F1Bet + 1
   If Money.Caption = 0 Then
    MsgBox "You are bankrupt after " & 30 - Days.Caption & " Days. Better luck next time"
Unload pLACINGS
Unload Gamble
End
Else
End If
End If
Corvette.Enabled = True
F50.Enabled = True
F1.Enabled = True
GT1.Enabled = True
Storm.Enabled = True
Speed12.Enabled = True
XJR.Enabled = True
XJ220.Enabled = True
End Sub

Private Sub F50Sub()

    MsgBox "F50 Won"
    pLACINGS.F50.Caption = pLACINGS.F50.Caption + 1
    Days.Caption = Days.Caption - 1

If F50 = False Then
    Money.Caption = Money.Caption - Bet
ElseIf F50 = True Then
    Money.Caption = Money.Caption + Val(Bet.Text) * 3
    pLACINGS.BetsWon = pLACINGS.BetsWon + 1
    pLACINGS.F50Bet = pLACINGS.F50Bet + 1
If Money.Caption = 0 Then
    MsgBox "You are bankrupt after " & 30 - Days.Caption & " Days. Better luck next time"
Unload pLACINGS
Unload Gamble
End
Else
End If
End If
Corvette.Enabled = True
F50.Enabled = True
F1.Enabled = True
GT1.Enabled = True
Storm.Enabled = True
Speed12.Enabled = True
XJR.Enabled = True
XJ220.Enabled = True
End Sub

Private Sub GT1Sub()

    MsgBox "The Toyota GT-ONE won the race!"
    pLACINGS.GT1.Caption = pLACINGS.GT1.Caption + 1
    Days.Caption = Days.Caption - 1
If GT1 = False Then
    Money.Caption = Money.Caption - Bet
ElseIf GT1 = True Then
    Money.Caption = Money.Caption + Val(Bet.Text) * 7
    pLACINGS.BetsWon = pLACINGS.BetsWon + 1
    pLACINGS.GT1Bet = pLACINGS.GT1Bet + 1
If Money.Caption = 0 Then
    MsgBox "You are bankrupt after " & 30 - Days.Caption & " Days. Better luck next time"
Unload pLACINGS
Unload Gamble
End
Else
End If
End If
Corvette.Enabled = True
F50.Enabled = True
F1.Enabled = True
GT1.Enabled = True
Storm.Enabled = True
Speed12.Enabled = True
XJR.Enabled = True
XJ220.Enabled = True
End Sub

Private Sub TVRSub()
    
    MsgBox "The TVR Speed 12 Won"
    pLACINGS.Speed12.Caption = pLACINGS.Speed12.Caption + 1
    Days.Caption = Days.Caption - 1
If Speed12 = False Then
    Money.Caption = Money.Caption - Bet
ElseIf Speed12 = True Then
    Money.Caption = Money.Caption + Val(Bet.Text) * 2
    pLACINGS.BetsWon = pLACINGS.BetsWon + 1
    pLACINGS.TVRBet = pLACINGS.TVRBet + 1
    
If Money.Caption = 0 Then
    MsgBox "You are bankrupt after " & 30 - Days.Caption & " Days. Better luck next time"
Unload pLACINGS
Unload Gamble
End
Else
End If
End If
Corvette.Enabled = True
F50.Enabled = True
F1.Enabled = True
GT1.Enabled = True
Storm.Enabled = True
Speed12.Enabled = True
XJR.Enabled = True
XJ220.Enabled = True
End Sub

Private Sub V12Sub()

    MsgBox "The Lister Storm V12 Won"
    pLACINGS.Storm.Caption = pLACINGS.Storm.Caption + 1
    Days.Caption = Days.Caption - 1
    
If Storm = False Then
    Money.Caption = Money.Caption - Bet
ElseIf Storm = True Then
    Money.Caption = Money.Caption + Val(Bet.Text) * 5
    pLACINGS.BetsWon = pLACINGS.BetsWon + 1
    pLACINGS.V12Bet = pLACINGS.V12Bet + 1
    
    If Money.Caption = 0 Then
    MsgBox "You are bankrupt after " & 30 - Days.Caption & " Days. Better luck next time"
Unload pLACINGS
Unload Gamble
End
Else
End If
End If
Corvette.Enabled = True
F50.Enabled = True
F1.Enabled = True
GT1.Enabled = True
Storm.Enabled = True
Speed12.Enabled = True
XJR.Enabled = True
XJ220.Enabled = True
End Sub

Private Sub XJ220Sub()
    MsgBox "The Jaguar XJ220 Won"
    pLACINGS.XJ220.Caption = pLACINGS.XJ220.Caption + 1
    Days.Caption = Days.Caption - 1

If XJ220 = False Then
    Money.Caption = Money.Caption - Bet
ElseIf XJ220 = True Then
    Money.Caption = Money.Caption + Val(Bet.Text) * 4
    pLACINGS.BetsWon = pLACINGS.BetsWon + 1
    pLACINGS.JagBet = pLACINGS.JagBet + 1
    If Money.Caption = 0 Then
    MsgBox "You are bankrupt after " & 30 - Days.Caption & " Days. Better luck next time"
Unload pLACINGS
Unload Gamble
End
Else
End If
End If

Corvette.Enabled = True
F50.Enabled = True
F1.Enabled = True
GT1.Enabled = True
Storm.Enabled = True
Speed12.Enabled = True
XJR.Enabled = True
XJ220.Enabled = True
End Sub

Private Sub XJR15Sub()

    MsgBox "The Jaguar XJR-15 Won"
    pLACINGS.XJR.Caption = pLACINGS.XJR.Caption + 1
    Days.Caption = Days.Caption - 1
If XJR = False Then
    Money.Caption = Money.Caption - Bet
ElseIf XJR = True Then
    Money.Caption = Money.Caption + Val(Bet.Text) * 6
    pLACINGS.BetsWon = pLACINGS.BetsWon + 1
    pLACINGS.XJRBet = pLACINGS.XJRBet + 1
If Money.Caption = 0 Then
    MsgBox "You are bankrupt after " & 30 - Days.Caption & " Days. Better luck next time"
Unload pLACINGS
Unload Gamble
End
Else

End If
End If
Corvette.Enabled = True
F50.Enabled = True
F1.Enabled = True
GT1.Enabled = True
Storm.Enabled = True
Speed12.Enabled = True
XJR.Enabled = True
XJ220.Enabled = True
End Sub

Private Sub Command1_Click()
Dim MyBomb
MyBomb = Int((45 * Rnd) + 1)

If MyBomb = 13 Then
    BOMB

End If

Command3.Enabled = False

If Bet.Text = "" Then
    Bet.Text = 5
ElseIf Bet.Text <= 0 Then
    Bet.Text = 5
End If

If Bet.Text > Val(Money.Caption) Then
    Beep
    Exit Sub
Else
End If
Bet = Bet.Text
tmrBet.Enabled = True

If Money.Caption < 0 Then
    Money = 0
Else
    Money = Money
End If

If Days.Caption = 0 And Money.Caption > 100 Then
    MsgBox "You have made $" & Money.Caption - 100 & " in 30 days of racing"
pLACINGS.Show
Gamble.Hide
pLACINGS.Command1.Enabled = False
pLACINGS.Quit.Enabled = False
pLACINGS.Quit.Visible = False
pLACINGS.Quit2.Enabled = True
pLACINGS.Quit2.Visible = True
ElseIf Days.Caption = 0 And Money.Caption < 100 Then
    MsgBox "You have lost $" & 100 - Money.Caption & " in 30 days of racing"
pLACINGS.Show
Gamble.Hide
pLACINGS.Command1.Enabled = False
pLACINGS.Quit.Enabled = False
pLACINGS.Quit.Visible = False
pLACINGS.Quit2.Enabled = True
pLACINGS.Quit2.Visible = True
ElseIf Days.Caption = 0 And Money.Caption = 100 Then
    MsgBox "You have neither gained or lost any money! Nothing to complain about really!!!"
pLACINGS.Show
Gamble.Hide
pLACINGS.Command1.Enabled = False
pLACINGS.Quit.Enabled = False
pLACINGS.Quit.Visible = False
pLACINGS.Quit2.Enabled = True
pLACINGS.Quit2.Visible = True
End If

If Money.Caption = 0 Then
    MsgBox "You are bankrupt after " & 30 - Days.Caption & " Days. Better luck next time"
Unload pLACINGS
Unload Gamble
End
Else
End If

If Days.Caption = 10 Then
    Lifeline.Enabled = True
    Lifeline.Visible = True
End If

End Sub




Private Sub Command2_Click()

Command3.Enabled = False

Dim RndCar
RndCar = Int((8 * Rnd) + 1)
If RndCar = 1 Then
    Corvette = True
ElseIf RndCar = 2 Then
    F50 = True
ElseIf RndCar = 3 Then
    F1 = True
ElseIf RndCar = 4 Then
    Speed12 = True
ElseIf RndCar = 5 Then
    Storm = True
ElseIf RndCar = 6 Then
    GT1 = True
ElseIf RndCar = 7 Then
    XJR = True
ElseIf RndCar = 8 Then
    XJ220 = True
End If

Dim RndMoney
RndMoney = Int((Money * Rnd) + 1)
Bet.Text = RndMoney

Bet = Bet.Text
tmrBet.Enabled = True

If Money.Caption < 0 Then
    Money = 0
Else
    Money = Money
End If

If Days.Caption = 0 And Money.Caption > 100 Then
    MsgBox "You have made $" & Money.Caption - 100 & " in 30 days of racing"
pLACINGS.Show
Gamble.Hide
pLACINGS.Command1.Enabled = False
pLACINGS.Quit.Enabled = False
pLACINGS.Quit.Visible = False
pLACINGS.Quit2.Enabled = True
pLACINGS.Quit2.Visible = True
ElseIf Days.Caption = 0 And Money.Caption < 100 Then
    MsgBox "You have lost $" & 100 - Money.Caption & " in 30 days of racing"
pLACINGS.Show
Gamble.Hide
pLACINGS.Command1.Enabled = False
pLACINGS.Quit.Enabled = False
pLACINGS.Quit.Visible = False
pLACINGS.Quit2.Enabled = True
pLACINGS.Quit2.Visible = True
ElseIf Days.Caption = 0 And Money.Caption = 100 Then
    MsgBox "You have neither gained or lost any money! Nothing to complain about really!!!"
pLACINGS.Show
Gamble.Hide
pLACINGS.Command1.Enabled = False
pLACINGS.Quit.Enabled = False
pLACINGS.Quit.Visible = False
pLACINGS.Quit2.Enabled = True
pLACINGS.Quit2.Visible = True
End If

If Money.Caption = 0 Then
    MsgBox "You are bankrupt after " & 30 - Days.Caption & " Days. Better luck next time"
Unload pLACINGS
Unload Gamble
End
Else
End If

If Days.Caption = 10 Then
    Lifeline.Enabled = True
    Lifeline.Visible = True
End If


End Sub

Private Sub Command3_Click()
Command3.Enabled = False
If Bet = "" Then
    Beep
    Bet = 5
ElseIf Bet < 1 Then
    Beep
    Bet = 5
ElseIf Bet > Money Then
    Beep
    Bet = 5
End If

Dim RndCar
RndCar = Int((8 * Rnd) + 1)
If RndCar = 1 Then
    Corvette = True
ElseIf RndCar = 2 Then
    F50 = True
ElseIf RndCar = 3 Then
    F1 = True
ElseIf RndCar = 4 Then
    Speed12 = True
ElseIf RndCar = 5 Then
    Storm = True
ElseIf RndCar = 6 Then
    GT1 = True
ElseIf RndCar = 7 Then
    XJR = True
ElseIf RndCar = 8 Then
    XJ220 = True
End If

tmrBet.Enabled = True
Bet = Bet.Text


If Money.Caption < 0 Then
    Money = 0
Else
    Money = Money
End If

If Days.Caption = 0 And Money.Caption > 100 Then
    MsgBox "You have made $" & Money.Caption - 100 & " in 30 days of racing"
pLACINGS.Show
Gamble.Hide
pLACINGS.Command1.Enabled = False
pLACINGS.Quit.Enabled = False
pLACINGS.Quit.Visible = False
pLACINGS.Quit2.Enabled = True
pLACINGS.Quit2.Visible = True
ElseIf Days.Caption = 0 And Money.Caption < 100 Then
    MsgBox "You have lost $" & 100 - Money.Caption & " in 30 days of racing"
pLACINGS.Show
Gamble.Hide
pLACINGS.Command1.Enabled = False
pLACINGS.Quit.Enabled = False
pLACINGS.Quit.Visible = False
pLACINGS.Quit2.Enabled = True
pLACINGS.Quit2.Visible = True
ElseIf Days.Caption = 0 And Money.Caption = 100 Then
    MsgBox "You have neither gained or lost any money! Nothing to complain about really!!!"
pLACINGS.Show
Gamble.Hide
pLACINGS.Command1.Enabled = False
pLACINGS.Quit.Enabled = False
pLACINGS.Quit.Visible = False
pLACINGS.Quit2.Enabled = True
pLACINGS.Quit2.Visible = True
End If

If Money.Caption = 0 Then
    MsgBox "You are bankrupt after " & 30 - Days.Caption & " Days. Better luck next time"
Unload pLACINGS
Unload Gamble
End
Else
End If

If Days.Caption = 10 Then
    Lifeline.Enabled = True
    Lifeline.Visible = True
End If



End Sub


Private Sub Corvette_Click()
Corvette = True
End Sub

Private Sub F1_Click()
F1 = True
End Sub


Private Sub F50_Click()
F50 = True
End Sub


Private Sub Form_Load()
Randomize
Lifeline.Enabled = False
Lifeline.Visible = False
MsgBox "You have 30 days to try and win as much money as possible betting on one of eight cars. Good luck!"
End Sub

Private Sub GT1_Click()
GT1 = True
End Sub

Private Sub Lifeline_Click()
Money = Money + 25
Lifeline.Enabled = False
Lifeline.Visible = False
End Sub

Private Sub Place_Click()
Gamble.Hide
pLACINGS.Show
End Sub


Private Sub Quit_Click()
Dim Msg, Style, Title, Response, MyString
Msg = "Are you sure you want to quit?"
Style = vbYesNo + vbQuestion + vbDefaultButton2
Title = "So you wanna quit?"  ' Define title.
Response = MsgBox(Msg, Style, Title)
If Response = vbYes Then
MyString = "Yes"    'end

If Money.Caption > 100 Then
    MsgBox "Loser! You quit after " & 30 - Days.Caption & " days and you had earned $" & Money.Caption - 100
Unload pLACINGS
Unload Gamble
End

ElseIf Money.Caption < 100 Then
    MsgBox "So you couldn't stand losing any more money? You lost $" & 100 - Money.Caption & " after " & 30 - Days.Caption & " days"
Unload pLACINGS
Unload Gamble
End

ElseIf Money.Caption = 100 And Days.Caption = 30 Then
    MsgBox "Um, yeah! Don't you want to at least try gambling first!?!"
Unload pLACINGS
Unload Gamble
End

ElseIf Money.Caption = 100 Then
    MsgBox "Can't complain! You still have $100 after " & 30 - Days.Caption & " days!"
Unload pLACINGS
Unload Gamble
End
End If

Else
    MyString = No 'Carry on
End If
End Sub


Private Sub Speed12_Click()
Speed12 = True
End Sub


Private Sub Storm_Click()
Storm = True
End Sub


Private Sub Timer1_Timer()
If Money.Caption < 0 Then
    Money = 0
Else
    Money = Money
End If

If Days.Caption = 0 And Money.Caption > 100 Then
    MsgBox "You have made $" & Money.Caption - 100 & " in 30 days of racing"
pLACINGS.Show
Gamble.Hide
pLACINGS.Command1.Enabled = False
pLACINGS.Quit.Enabled = False
pLACINGS.Quit.Visible = False
pLACINGS.Quit2.Enabled = True
pLACINGS.Quit2.Visible = True
tmrBet.Enabled = False
ElseIf Days.Caption = 0 And Money.Caption < 100 Then
    MsgBox "You have lost $" & 100 - Money.Caption & " in 30 days of racing"
pLACINGS.Show
Gamble.Hide
pLACINGS.Command1.Enabled = False
pLACINGS.Quit.Enabled = False
pLACINGS.Quit.Visible = False
pLACINGS.Quit2.Enabled = True
pLACINGS.Quit2.Visible = True
tmrBet.Enabled = False
ElseIf Days.Caption = 0 And Money.Caption = 100 Then
    MsgBox "You have neither gained or lost any money! Nothing to complain about really!!!"
pLACINGS.Show
Gamble.Hide
pLACINGS.Command1.Enabled = False
pLACINGS.Quit.Enabled = False
pLACINGS.Quit.Visible = False
pLACINGS.Quit2.Enabled = True
pLACINGS.Quit2.Visible = True
tmrBet.Enabled = False
End If

If Money.Caption = 0 Then
    MsgBox "You are bankrupt after " & 30 - Days.Caption & " Days. Better luck next time"
    tmrBet.Enabled = False
Unload pLACINGS
Unload Gamble
End
Else
End If

End Sub

Private Sub tmrBet_Timer()

Corvette.Enabled = False
F50.Enabled = False
F1.Enabled = False
GT1.Enabled = False
Storm.Enabled = False
Speed12.Enabled = False
XJR.Enabled = False
XJ220.Enabled = False

Dim MyF1
Dim MyGT1
Dim MyV12
Dim MyXJ220
Dim MyXJR
Dim MyVette
Dim MyF50
Dim MyTVR


MyF1 = Int((205 * Rnd) + 1)
MyTVR = Int((199 * Rnd) + 1)
MyF50 = Int((198 * Rnd) + 1)
MyXJ220 = Int((197 * Rnd) + 1)
MyV12 = Int((196 * Rnd) + 1)
MyXJR = Int((195 * Rnd) + 1)
MyGT1 = Int((194 * Rnd) + 1)
MyVette = Int((193 * Rnd) + 1)



Shape1.Left = Shape1.Left + MyVette
Shape2.Left = Shape2.Left + MyF50
Shape3.Left = Shape3.Left + MyF1
Shape4.Left = Shape4.Left + MyTVR
Shape5.Left = Shape5.Left + MyXJ220
Shape6.Left = Shape6.Left + MyV12
Shape7.Left = Shape7.Left + MyXJR
Shape8.Left = Shape8.Left + MyGT1

If Shape1.Left >= 6340 Then
    Shape1.Left = 120
    Shape2.Left = 120
    Shape3.Left = 120
    Shape4.Left = 120
    Shape5.Left = 120
    Shape6.Left = 120
    Shape7.Left = 120
    Shape8.Left = 120
    Command3.Enabled = True
    tmrBet.Enabled = False
    CorvetteSub
ElseIf Shape2.Left >= 6340 Then
    F50Sub
    Shape1.Left = 120
    Shape2.Left = 120
    Shape3.Left = 120
    Shape4.Left = 120
    Shape5.Left = 120
    Shape6.Left = 120
    Shape7.Left = 120
    Shape8.Left = 120
    Command3.Enabled = True
    tmrBet.Enabled = False
ElseIf Shape3.Left >= 6340 Then
    F1Sub
    Shape1.Left = 120
    Shape2.Left = 120
    Shape3.Left = 120
    Shape4.Left = 120
    Shape5.Left = 120
    Shape6.Left = 120
    Shape7.Left = 120
    Shape8.Left = 120
    Command3.Enabled = True
    tmrBet.Enabled = False
ElseIf Shape4.Left >= 6340 Then
    TVRSub
    Shape1.Left = 120
    Shape2.Left = 120
    Shape3.Left = 120
    Shape4.Left = 120
    Shape5.Left = 120
    Shape6.Left = 120
    Shape7.Left = 120
    Shape8.Left = 120
    Command3.Enabled = True
    tmrBet.Enabled = False
ElseIf Shape5.Left >= 6340 Then
    XJ220Sub
    Shape1.Left = 120
    Shape2.Left = 120
    Shape3.Left = 120
    Shape4.Left = 120
    Shape5.Left = 120
    Shape6.Left = 120
    Shape7.Left = 120
    Shape8.Left = 120
    Command3.Enabled = True
    tmrBet.Enabled = False
ElseIf Shape6.Left >= 6340 Then
    V12Sub
    Shape1.Left = 120
    Shape2.Left = 120
    Shape3.Left = 120
    Shape4.Left = 120
    Shape5.Left = 120
    Shape6.Left = 120
    Shape7.Left = 120
    Shape8.Left = 120
    Command3.Enabled = True
    tmrBet.Enabled = False
ElseIf Shape7.Left >= 6340 Then
    XJR15Sub
    Shape1.Left = 120
    Shape2.Left = 120
    Shape3.Left = 120
    Shape4.Left = 120
    Shape5.Left = 120
    Shape6.Left = 120
    Shape7.Left = 120
    Shape8.Left = 120
    Command3.Enabled = True
    tmrBet.Enabled = False
ElseIf Shape8.Left >= 6340 Then
    GT1Sub
    Shape1.Left = 120
    Shape2.Left = 120
    Shape3.Left = 120
    Shape4.Left = 120
    Shape5.Left = 120
    Shape6.Left = 120
    Shape7.Left = 120
    Shape8.Left = 120
    Command3.Enabled = True
    tmrBet.Enabled = False
End If

If Image1.Visible = True And Shape1.Left >= 2300 Then
    Image1.Visible = False
    Shape1.Left = 0
ElseIf Image2.Visible = True And Shape2.Left >= 2300 Then
    Image2.Visible = False
    Shape2.Left = 0
ElseIf Image3.Visible = True And Shape3.Left >= 2300 Then
    Image3.Visible = False
    Shape3.Left = 0
ElseIf Image4.Visible = True And Shape4.Left >= 2300 Then
    Image4.Visible = False
    Shape4.Left = 0
ElseIf Image5.Visible = True And Shape5.Left >= 2300 Then
    Image5.Visible = False
    Shape5.Left = 0
ElseIf Image6.Visible = True And Shape6.Left >= 2300 Then
    Image6.Visible = False
    Shape6.Left = 0
ElseIf Image7.Visible = True And Shape7.Left >= 2300 Then
    Image7.Visible = False
    Shape7.Left = 0
ElseIf Image8.Visible = True And Shape8.Left >= 2300 Then
    Image8.Visible = False
    Shape8.Left = 0
End If

If Shape1.Left > Shape2.Left And Shape1.Left > Shape3.Left And Shape1.Left > Shape4.Left And Shape1.Left > Shape5.Left And Shape1.Left > Shape6.Left And Shape1.Left > Shape7.Left And Shape1.Left > Shape8.Left Then
    lblWin.Caption = "The Chevrolet Corvette is leading"
ElseIf Shape2.Left > Shape1.Left And Shape2.Left > Shape3.Left And Shape2.Left > Shape4.Left And Shape2.Left > Shape5.Left And Shape2.Left > Shape6.Left And Shape2.Left > Shape7.Left And Shape2.Left > Shape8.Left Then
    lblWin.Caption = "The Ferrari F50 is in the lead"
ElseIf Shape3.Left > Shape1.Left And Shape3.Left > Shape2.Left And Shape3.Left > Shape4.Left And Shape3.Left > Shape5.Left And Shape3.Left > Shape6.Left And Shape3.Left > Shape7.Left And Shape3.Left > Shape8.Left Then
    lblWin.Caption = "The McLaren F1 is winning"
ElseIf Shape4.Left > Shape1.Left And Shape4.Left > Shape3.Left And Shape4.Left > Shape2.Left And Shape4.Left > Shape5.Left And Shape4.Left > Shape6.Left And Shape4.Left > Shape7.Left And Shape4.Left > Shape8.Left Then
    lblWin.Caption = "The TVR Speed 12 has taken the lead"
ElseIf Shape5.Left > Shape1.Left And Shape5.Left > Shape2.Left And Shape5.Left > Shape4.Left And Shape5.Left > Shape3.Left And Shape5.Left > Shape6.Left And Shape5.Left > Shape7.Left And Shape5.Left > Shape8.Left Then
    lblWin.Caption = "The Jaguar XJ220 is storming ahead"
ElseIf Shape6.Left > Shape1.Left And Shape6.Left > Shape2.Left And Shape6.Left > Shape4.Left And Shape6.Left > Shape3.Left And Shape6.Left > Shape5.Left And Shape6.Left > Shape7.Left And Shape6.Left > Shape8.Left Then
    lblWin.Caption = "The Lister Storm V12 is drinking fuel like it's water"
ElseIf Shape7.Left > Shape1.Left And Shape7.Left > Shape2.Left And Shape5.Left > Shape4.Left And Shape7.Left > Shape3.Left And Shape7.Left > Shape6.Left And Shape7.Left > Shape5.Left And Shape7.Left > Shape8.Left Then
    lblWin.Caption = "The Jaguar XJR-15 is really running on full steam now"
ElseIf Shape8.Left > Shape1.Left And Shape8.Left > Shape2.Left And Shape5.Left > Shape4.Left And Shape8.Left > Shape3.Left And Shape8.Left > Shape6.Left And Shape8.Left > Shape7.Left And Shape8.Left > Shape5.Left Then
    lblWin.Caption = "The Toyota GT-ONE has pinched the lead"
End If

End Sub

Private Sub XJ220_Click()
XJ220 = True
End Sub


Private Sub XJR_Click()
XJR = True
End Sub


