VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   Caption         =   " TERMODINAMICA"
   ClientHeight    =   9900
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10515
   LinkTopic       =   "Form1"
   ScaleHeight     =   9900
   ScaleWidth      =   10515
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command9 
      Caption         =   "BORRAR TODO"
      Height          =   855
      Left            =   7440
      TabIndex        =   69
      Top             =   2520
      Width           =   1935
   End
   Begin VB.CommandButton Command8 
      Caption         =   "INICIO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   63
      Top             =   960
      Width           =   1935
   End
   Begin VB.CommandButton Command7 
      Caption         =   "SALIR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7440
      Style           =   1  'Graphical
      TabIndex        =   62
      Top             =   960
      Width           =   1935
   End
   Begin VB.CommandButton Command6 
      Caption         =   "CALCULO FINAL"
      Height          =   615
      Left            =   8520
      TabIndex        =   61
      Top             =   9000
      Width           =   1335
   End
   Begin VB.CommandButton Command5 
      Caption         =   "CALCULO 3"
      Height          =   375
      Left            =   8520
      TabIndex        =   56
      Top             =   7560
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "CALCULO 2"
      Height          =   375
      Left            =   8520
      TabIndex        =   51
      Top             =   5880
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "CALCULO 1"
      Height          =   375
      Left            =   8520
      TabIndex        =   50
      Top             =   4200
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "HUMEDAD ESPECIFICA"
      Height          =   855
      Left            =   360
      TabIndex        =   41
      Top             =   7800
      Width           =   1695
   End
   Begin VB.TextBox Text28 
      Height          =   495
      Left            =   15840
      TabIndex        =   35
      Top             =   9120
      Width           =   1335
   End
   Begin VB.TextBox Text27 
      Height          =   495
      Left            =   14040
      TabIndex        =   34
      Top             =   9120
      Width           =   1455
   End
   Begin VB.TextBox Text26 
      Height          =   495
      Left            =   12360
      TabIndex        =   33
      Top             =   9120
      Width           =   1335
   End
   Begin VB.TextBox Text25 
      Height          =   495
      Left            =   10680
      TabIndex        =   32
      Top             =   9120
      Width           =   1335
   End
   Begin VB.TextBox Text24 
      Height          =   495
      Left            =   15840
      TabIndex        =   31
      Top             =   7440
      Width           =   1335
   End
   Begin VB.TextBox Text23 
      Height          =   495
      Left            =   14040
      TabIndex        =   30
      Top             =   7440
      Width           =   1455
   End
   Begin VB.TextBox Text22 
      Height          =   495
      Left            =   12360
      TabIndex        =   29
      Top             =   7440
      Width           =   1335
   End
   Begin VB.TextBox Text21 
      Height          =   495
      Left            =   10680
      TabIndex        =   28
      Top             =   7440
      Width           =   1335
   End
   Begin VB.TextBox Text20 
      Height          =   495
      Left            =   15840
      TabIndex        =   27
      Top             =   5760
      Width           =   1335
   End
   Begin VB.TextBox Text19 
      Height          =   495
      Left            =   14040
      TabIndex        =   26
      Top             =   5760
      Width           =   1455
   End
   Begin VB.TextBox Text18 
      Height          =   495
      Left            =   12360
      TabIndex        =   25
      Top             =   5760
      Width           =   1335
   End
   Begin VB.TextBox Text17 
      Height          =   495
      Left            =   10680
      TabIndex        =   24
      Top             =   5760
      Width           =   1335
   End
   Begin VB.TextBox Text16 
      Height          =   495
      Left            =   15840
      TabIndex        =   23
      Top             =   4080
      Width           =   1335
   End
   Begin VB.TextBox Text15 
      Height          =   495
      Left            =   14040
      TabIndex        =   22
      Top             =   4080
      Width           =   1455
   End
   Begin VB.TextBox Text14 
      Height          =   495
      Left            =   12360
      TabIndex        =   21
      Top             =   4080
      Width           =   1335
   End
   Begin VB.TextBox Text13 
      Height          =   495
      Left            =   10680
      TabIndex        =   20
      Top             =   4080
      Width           =   1335
   End
   Begin VB.TextBox Text12 
      Height          =   495
      Left            =   2040
      TabIndex        =   19
      Top             =   9480
      Width           =   1455
   End
   Begin VB.TextBox Text11 
      Height          =   495
      Left            =   360
      TabIndex        =   18
      Top             =   9480
      Width           =   1335
   End
   Begin VB.TextBox Text10 
      Height          =   495
      Left            =   3840
      TabIndex        =   17
      Top             =   6840
      Width           =   1335
   End
   Begin VB.TextBox Text9 
      Height          =   495
      Left            =   2160
      TabIndex        =   16
      Top             =   6840
      Width           =   1335
   End
   Begin VB.TextBox Text8 
      Height          =   495
      Left            =   360
      TabIndex        =   15
      Top             =   6840
      Width           =   1455
   End
   Begin VB.TextBox Text7 
      Height          =   495
      Left            =   4200
      TabIndex        =   10
      Top             =   4440
      Width           =   1455
   End
   Begin VB.TextBox Text6 
      Height          =   525
      Left            =   2280
      TabIndex        =   9
      Top             =   4440
      Width           =   1455
   End
   Begin VB.TextBox Text5 
      Height          =   495
      Left            =   360
      TabIndex        =   8
      Top             =   4440
      Width           =   1575
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   4080
      TabIndex        =   7
      Top             =   3000
      Width           =   1575
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   2160
      TabIndex        =   5
      Top             =   3000
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Height          =   525
      Left            =   360
      TabIndex        =   4
      Top             =   3000
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   360
      TabIndex        =   3
      Top             =   2040
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CALCULO DE Pt ,Ps1, Ps2"
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   5280
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   1485
      Left            =   2280
      Picture         =   "programa de termo.frx":0000
      Stretch         =   -1  'True
      Top             =   240
      Width           =   1800
   End
   Begin VB.Label Label33 
      BackStyle       =   0  'Transparent
      Caption         =   "m3, T3, W3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   17400
      TabIndex        =   68
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Label Label32 
      BackStyle       =   0  'Transparent
      Caption         =   "m2. T2. W2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   10560
      TabIndex        =   67
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label Label31 
      BackStyle       =   0  'Transparent
      Caption         =   "m1, T1 , W1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10680
      TabIndex        =   66
      Top             =   720
      Width           =   1935
   End
   Begin VB.Label Label30 
      BackStyle       =   0  'Transparent
      Caption         =   "SALIDA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   17280
      TabIndex        =   65
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label Label29 
      BackStyle       =   0  'Transparent
      Caption         =   "ENTRADA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10320
      TabIndex        =   64
      Top             =   240
      Width           =   1815
   End
   Begin VB.Line Line11 
      X1              =   14280
      X2              =   15000
      Y1              =   1680
      Y2              =   1440
   End
   Begin VB.Line Line10 
      X1              =   14280
      X2              =   15000
      Y1              =   1200
      Y2              =   1440
   End
   Begin VB.Line Line9 
      X1              =   14280
      X2              =   15360
      Y1              =   2160
      Y2              =   1800
   End
   Begin VB.Line Line8 
      X1              =   14280
      X2              =   15360
      Y1              =   720
      Y2              =   1080
   End
   Begin VB.Line Line7 
      X1              =   15360
      X2              =   17160
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Line Line6 
      X1              =   15360
      X2              =   17160
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Line Line5 
      X1              =   12240
      X2              =   14280
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Line Line4 
      X1              =   12240
      X2              =   14280
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Line Line3 
      X1              =   12240
      X2              =   14280
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   2880
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line1 
      X1              =   12240
      X2              =   14280
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Humedad relativa 3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   15840
      TabIndex        =   60
      Top             =   8400
      Width           =   1455
   End
   Begin VB.Label Label27 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Psat 3 [Pa]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   14280
      TabIndex        =   59
      Top             =   8400
      Width           =   1095
   End
   Begin VB.Label Label26 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Temperatura 3 [ªC]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12360
      TabIndex        =   58
      Top             =   8400
      Width           =   1335
   End
   Begin VB.Label Label25 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "W3 [kg H2O/kg as]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10680
      TabIndex        =   57
      Top             =   8280
      Width           =   1455
   End
   Begin VB.Label Label24 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "m3aire seco [kg/min]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   15720
      TabIndex        =   55
      Top             =   6840
      Width           =   1695
   End
   Begin VB.Label Label23 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "m3H2O [kg/min]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   14160
      TabIndex        =   54
      Top             =   6720
      Width           =   1215
   End
   Begin VB.Label Label22 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "P3aire seco [Pa]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12360
      TabIndex        =   53
      Top             =   6720
      Width           =   1215
   End
   Begin VB.Label Label21 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "P3 H2O[Pa]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10680
      TabIndex        =   52
      Top             =   6960
      Width           =   1335
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "m2aire seco [kg/min]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   15720
      TabIndex        =   49
      Top             =   5040
      Width           =   1695
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "m2H2O [kg/min]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   14160
      TabIndex        =   48
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "P2aire seco [Pa]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12360
      TabIndex        =   47
      Top             =   5040
      Width           =   1455
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "P2 H2O [Pa]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10800
      TabIndex        =   46
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "m1aire seco[Kg/min]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   15840
      TabIndex        =   45
      Top             =   3360
      Width           =   1455
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "m1H2O[Kg/min]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   13920
      TabIndex        =   44
      Top             =   3360
      Width           =   1695
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "P1aire seco [Pa]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12360
      TabIndex        =   43
      Top             =   3360
      Width           =   1455
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "P1 H2O [Pa]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10800
      TabIndex        =   42
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "W2 [kg H2O/kg as]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2040
      TabIndex        =   40
      Top             =   8880
      Width           =   1575
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "W1 [kg H2O/kg as]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   39
      Top             =   8880
      Width           =   1455
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Psat 2 [Pa]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   38
      Top             =   6480
      Width           =   1335
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Psat 1 [Pa]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   37
      Top             =   6480
      Width           =   1215
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Ptotal [Pa]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   36
      Top             =   6480
      Width           =   1215
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Humedad relativa 2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4200
      TabIndex        =   14
      Top             =   3720
      Width           =   1575
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Humedad relativa 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      TabIndex        =   13
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Temperatura 2 [ºC]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   12
      Top             =   3600
      Width           =   1455
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Temperatura 1 [ºC]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   11
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "masa 2 [kg/min]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   6
      Top             =   3600
      Width           =   1575
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "masa 1 [kg/min]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   2520
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Altitud [m]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   1680
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()

z = Val(Text1.Text)

Pt = 101.325 * (1 - 0.0000225577 * z) ^ 5.2559
Text8.Text = Format(Pt, "0.00")

C8 = -5800.2206
C9 = 1.3914993
C10 = -0.048640239
C11 = 0.000041764768
C12 = -0.000000014452093
C13 = 6.5459673
T1 = Val(Text3.Text) + 273
k = (C8 / T1) + C9 + (C10 * T1) + (C11 * T1 ^ 2) + (C12 * T1 ^ 3) + (C13 * Log(T1))
e = 2.718281828
p1 = e ^ k
Text9.Text = Format(p1)

T2 = Val(Text6.Text) + 273
k1 = (C8 / T2) + C9 + (C10 * T2) + (C11 * T2 ^ 2) + (C12 * T2 ^ 3) + (C13 * Log(T2))
e = 2.718281828
p2 = e ^ k1
Text10.Text = Format(p2)



Text8.Visible = True
Text9.Visible = True
Text10.Visible = True

Label8.Visible = True
Label9.Visible = True
Label10.Visible = True
Command2.Visible = True

End Sub

Private Sub Command2_Click()
PQ = 72000
q = Val(Text4.Text) / 100
PH2O = q * Val(Text9.Text)
Pas = PQ - PH2O
w1 = 0.622 * (PH2O / Pas)
Text11.Text = Format(w1)

Text11.Visible = True
Label11.Visible = True


q1 = Val(Text7.Text) / 100
P2H2O = q1 * Val(Text10.Text)
P2As = PQ - P2H2O
w2 = 0.622 * (P2H2O / P2As)
Text12.Text = Format(w2)

Text12.Visible = True
Label12.Visible = True

Command3.Visible = True



End Sub

Private Sub Command3_Click()
PQ = 72000
q = Val(Text4.Text) / 100

PH2O = q * Val(Text9.Text)
Pas = PQ - PH2O
Text13.Text = Val(PH2O)
Text14.Text = Val(Pas)

Pas = PQ - PH2O
w1 = 0.622 * (PH2O / Pas)

m1 = Val(Text2.Text)
m1H2O = (w1 * m1) / (1 + w1)
Text15.Text = Format(m1H2O)
mas = m1H2O / w1
Text16.Text = Format(mas)

Text13.Visible = True
Text14.Visible = True
Text15.Visible = True
Text16.Visible = True

Label13.Visible = True
Label14.Visible = True
Label15.Visible = True
Label16.Visible = True

Command4.Visible = True
End Sub

Private Sub Command4_Click()
PQ = 72000
q2 = Val(Text7.Text) / 100
P2H2O = q2 * Val(Text10.Text)
P2As = PQ - P2H2O
Text17.Text = Val(P2H2O)
Text18.Text = Val(P2As)

P2As = PQ - P2H2O
w2 = 0.622 * (P2H2O / P2As)


m2 = Val(Text5.Text)
m2H2O = (w2 * m2) / (1 + w2)
Text19.Text = Format(m2H2O)
m2as = m2H2O / w2
Text20.Text = Format(m2as)


Text17.Visible = True
Text18.Visible = True
Text19.Visible = True
Text20.Visible = True

Label17.Visible = True
Label18.Visible = True
Label19.Visible = True
Label20.Visible = True

Command5.Visible = True




End Sub

Private Sub Command5_Click()
PQ = 72000
q = Val(Text4.Text) / 100

PH2O = q * Val(Text9.Text)
Pas = PQ - PH2O
Text13.Text = Val(PH2O)
Text14.Text = Val(Pas)

Pas = PQ - PH2O
w1 = 0.622 * (PH2O / Pas)

m1 = Val(Text2.Text)
m1H2O = (w1 * m1) / (1 + w1)
Text15.Text = Format(m1H2O)
mas = m1H2O / w1
Text16.Text = Format(mas)


PQ = 72000
q2 = Val(Text7.Text) / 100
P2H2O = q2 * Val(Text10.Text)
P2As = PQ - P2H2O
Text17.Text = Val(P2H2O)
Text18.Text = Val(P2As)

P2As = PQ - P2H2O
w2 = 0.622 * (P2H2O / P2As)

m2 = Val(Text5.Text)
m2H2O = (w2 * m2) / (1 + w2)
Text19.Text = Format(m2H2O)
m2as = m2H2O / w2
Text20.Text = Format(m2as)

m3H2O = m1H2O + m2H2O
Text23.Text = Format(m3H2O)

m3as = mas + m2as
Text24.Text = Format(m3as)


W3 = m3H2O / m3as
P3H2O = (W3 * PQ) / (0.622 + W3)
Text21.Text = Format(P3H2O)

Pas = PQ - P3H2O
Text22.Text = Format(Pas)

Text21.Visible = True
Text22.Visible = True
Text23.Visible = True
Text24.Visible = True
Label21.Visible = True
Label22.Visible = True
Label23.Visible = True
Label24.Visible = True

Command6.Visible = True
End Sub

Private Sub Command6_Click()
PQ = 72000
q = Val(Text4.Text) / 100

PH2O = q * Val(Text9.Text)
Pas = PQ - PH2O
Text13.Text = Val(PH2O)
Text14.Text = Val(Pas)

Pas = PQ - PH2O
w1 = 0.622 * (PH2O / Pas)

m1 = Val(Text2.Text)
m1H2O = (w1 * m1) / (1 + w1)
Text15.Text = Format(m1H2O)
mas = m1H2O / w1
Text16.Text = Format(mas)


PQ = 72000
q2 = Val(Text7.Text) / 100
P2H2O = q2 * Val(Text10.Text)
P2As = PQ - P2H2O
Text17.Text = Val(P2H2O)
Text18.Text = Val(P2As)

P2As = PQ - P2H2O
w2 = 0.622 * (P2H2O / P2As)

m2 = Val(Text5.Text)
m2H2O = (w2 * m2) / (1 + w2)
Text19.Text = Format(m2H2O)
m2as = m2H2O / w2
Text20.Text = Format(m2as)

m3H2O = m1H2O + m2H2O
Text23.Text = Format(m3H2O)

m3as = mas + m2as
Text24.Text = Format(m3as)


W3 = m3H2O / m3as
Text25.Text = Format(W3)

T1 = Val(Text3.Text)
T2 = Val(Text6.Text)

T3 = (m2 * T2 + m1 * T1) / (m1 + m2)
Text26.Text = Format(T3)

P1sat = Val(Text9.Text)
P2sat = Val(Text10.Text)

P3sat = P1sat + (P2sat - P1sat) * ((T3 - T1) / (T2 - T1))
Text27.Text = Format(P3sat)


PQ = 72000
q2 = Val(Text7.Text) / 100
P2H2O = q2 * Val(Text10.Text)
P2As = PQ - P2H2O
Text17.Text = Val(P2H2O)
Text18.Text = Val(P2As)

P2As = PQ - P2H2O
w2 = 0.622 * (P2H2O / P2As)

m2 = Val(Text5.Text)
m2H2O = (w2 * m2) / (1 + w2)
Text19.Text = Format(m2H2O)
m2as = m2H2O / w2
Text20.Text = Format(m2as)

m3H2O = m1H2O + m2H2O
Text23.Text = Format(m3H2O)

m3as = mas + m2as
Text24.Text = Format(m3as)


W3 = m3H2O / m3as
P3H2O = (W3 * PQ) / (0.622 + W3)
Text21.Text = Format(P3H2O)

O3 = (P3H2O / P3sat) * 100
Text28.Text = Format(O3)

Text25.Visible = True
Text26.Visible = True
Text27.Visible = True
Text28.Visible = True
Label25.Visible = True
Label26.Visible = True
Label27.Visible = True
Label28.Visible = True


End Sub

Private Sub Command7_Click()
End
End Sub

Private Sub Command8_Click()
A = MsgBox("Desea iniciar el programa ", vbYesNo + 32, "SOLUCIONES")
If A = vbYes Then
z = MsgBox("Introduzca los valores iniciales....Tenga en cuenta que los valores de presion hacen referencia a la ciudad de QUITO ", vbOKOnly, "INICIANDO")
Command8.Visible = False
Text1.Visible = True
Text2.Visible = True
Text3.Visible = True
Text4.Visible = True
Text5.Visible = True
Text6.Visible = True
Text7.Visible = True
Label1.Visible = True
Label2.Visible = True
Label3.Visible = True
Label4.Visible = True
Label5.Visible = True
Label6.Visible = True
Command1.Visible = True
Else
End
End If


End Sub

Private Sub Command9_Click()
Text1.Text = CLEAN
Text2.Text = CLEAN
Text3.Text = CLEAN
Text4.Text = CLEAN
Text5.Text = CLEAN
Text6.Text = CLEAN
Text7.Text = CLEAN
Text8.Text = CLEAN
Text9.Text = CLEAN
Text10.Text = CLEAN
Text11.Text = CLEAN
Text12.Text = CLEAN
Text13.Text = CLEAN
Text14.Text = CLEAN
Text15.Text = CLEAN
Text16.Text = CLEAN
Text17.Text = CLEAN
Text18.Text = CLEAN
Text19.Text = CLEAN
Text20.Text = CLEAN
Text21.Text = CLEAN
Text22.Text = CLEAN
Text23.Text = CLEAN
Text24.Text = CLEAN
Text25.Text = CLEAN
Text26.Text = CLEAN
Text27.Text = CLEAN
Text28.Text = CLEAN
End Sub


Private Sub Form_Load()

Text1.Visible = False
Text2.Visible = False
Text3.Visible = False
Text4.Visible = False
Text5.Visible = False
Text6.Visible = False
Text7.Visible = False
Text8.Visible = False
Text9.Visible = False
Text10.Visible = False
Text11.Visible = False
Text12.Visible = False
Text13.Visible = False
Text14.Visible = False
Text15.Visible = False
Text16.Visible = False
Text17.Visible = False
Text18.Visible = False
Text19.Visible = False
Text20.Visible = False
Text21.Visible = False
Text22.Visible = False
Text23.Visible = False
Text24.Visible = False
Text25.Visible = False
Text26.Visible = False
Text27.Visible = False
Text28.Visible = False

Label1.Visible = False
Label2.Visible = False
Label3.Visible = False
Label4.Visible = False
Label5.Visible = False
Label6.Visible = False
Label7.Visible = False
Label8.Visible = False
Label9.Visible = False
Label10.Visible = False
Label11.Visible = False
Label12.Visible = False
Label13.Visible = False
Label14.Visible = False
Label15.Visible = False
Label16.Visible = False
Label17.Visible = False
Label18.Visible = False
Label19.Visible = False
Label20.Visible = False
Label21.Visible = False
Label22.Visible = False
Label23.Visible = False
Label24.Visible = False
Label25.Visible = False
Label26.Visible = False
Label27.Visible = False
Label28.Visible = False

Command1.Visible = False
Command2.Visible = False
Command3.Visible = False
Command4.Visible = False
Command5.Visible = False
Command6.Visible = False


End Sub

