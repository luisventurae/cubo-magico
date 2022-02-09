VERSION 5.00
Begin VB.Form frmjuego 
   BackColor       =   &H00C0E0FF&
   Caption         =   "Cubo mágico"
   ClientHeight    =   8235
   ClientLeft      =   3480
   ClientTop       =   1680
   ClientWidth     =   10860
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8235
   ScaleWidth      =   10860
   Begin VB.Timer tecla 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   6960
      Top             =   3480
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ver puntuaciones"
      Height          =   375
      Left            =   2160
      TabIndex        =   56
      Top             =   7560
      Width           =   1455
   End
   Begin VB.CommandButton cmdlimpiar 
      BackColor       =   &H00FFFFC0&
      Caption         =   "limpiar"
      Enabled         =   0   'False
      Height          =   315
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   49
      Top             =   5640
      Width           =   615
   End
   Begin VB.CommandButton cmdborrar 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Borrar una casilla"
      Enabled         =   0   'False
      Height          =   255
      Left            =   9360
      Style           =   1  'Graphical
      TabIndex        =   48
      Top             =   5640
      Width           =   1335
   End
   Begin VB.CommandButton cmdre 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Volver a jugar"
      Height          =   735
      Left            =   9480
      Style           =   1  'Graphical
      TabIndex        =   47
      Top             =   7440
      Width           =   1335
   End
   Begin VB.Frame fra123 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Boton de juego"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   1695
      Left            =   7320
      TabIndex        =   43
      Top             =   3840
      Width           =   1575
      Begin VB.OptionButton opt9 
         BackColor       =   &H00C0E0FF&
         Height          =   255
         Left            =   1080
         TabIndex        =   55
         Top             =   1440
         Width           =   255
      End
      Begin VB.OptionButton opt8 
         BackColor       =   &H00C0E0FF&
         Height          =   375
         Left            =   600
         TabIndex        =   54
         Top             =   1320
         Width           =   495
      End
      Begin VB.OptionButton opt7 
         BackColor       =   &H00C0E0FF&
         Height          =   255
         Left            =   120
         TabIndex        =   53
         Top             =   1320
         Width           =   375
      End
      Begin VB.OptionButton opt6 
         BackColor       =   &H00C0E0FF&
         Height          =   255
         Left            =   1080
         TabIndex        =   52
         Top             =   840
         Width           =   375
      End
      Begin VB.OptionButton opt5 
         BackColor       =   &H00C0E0FF&
         Height          =   375
         Left            =   600
         TabIndex        =   51
         Top             =   720
         Width           =   495
      End
      Begin VB.OptionButton opt4 
         BackColor       =   &H00C0E0FF&
         Height          =   375
         Left            =   120
         TabIndex        =   50
         Top             =   720
         Width           =   495
      End
      Begin VB.OptionButton opt3 
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   375
         Left            =   1080
         TabIndex        =   46
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton opt2 
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   375
         Left            =   600
         TabIndex        =   45
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton opt1 
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   375
         Left            =   120
         TabIndex        =   44
         Top             =   240
         Value           =   -1  'True
         Width           =   735
      End
   End
   Begin VB.CommandButton cmd9 
      Caption         =   "9"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10200
      TabIndex        =   42
      Top             =   5040
      Width           =   495
   End
   Begin VB.CommandButton cmd6 
      Caption         =   "6"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10200
      TabIndex        =   41
      Top             =   4440
      Width           =   495
   End
   Begin VB.CommandButton cmd3 
      Caption         =   "3"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10200
      TabIndex        =   40
      Top             =   3840
      Width           =   495
   End
   Begin VB.CommandButton cmd8 
      Caption         =   "8"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9600
      TabIndex        =   39
      Top             =   5040
      Width           =   495
   End
   Begin VB.CommandButton cmd5 
      Caption         =   "5"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9600
      TabIndex        =   38
      Top             =   4440
      Width           =   495
   End
   Begin VB.CommandButton cmd2 
      Caption         =   "2"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9600
      TabIndex        =   37
      Top             =   3840
      Width           =   495
   End
   Begin VB.CommandButton cmd1 
      Caption         =   "1"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9000
      TabIndex        =   36
      Top             =   3840
      Width           =   495
   End
   Begin VB.CommandButton cmd4 
      Caption         =   "4"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9000
      TabIndex        =   35
      Top             =   4440
      Width           =   495
   End
   Begin VB.CommandButton cmd7 
      Caption         =   "7"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9000
      TabIndex        =   34
      Top             =   5040
      Width           =   495
   End
   Begin VB.Timer Timerverificador 
      Interval        =   3
      Left            =   7560
      Top             =   6000
   End
   Begin VB.CommandButton cmdsalir 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Salir"
      Height          =   375
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   7680
      Width           =   1215
   End
   Begin VB.CommandButton cmdayuda 
      Caption         =   "Ayuda"
      Enabled         =   0   'False
      Height          =   375
      Left            =   360
      TabIndex        =   26
      Top             =   7560
      Width           =   1695
   End
   Begin VB.CommandButton cmdfinal 
      Caption         =   "Terminar la cuenta y ver tu puntaje"
      Enabled         =   0   'False
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
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   7440
      Width           =   3615
   End
   Begin VB.CommandButton cmdinicio 
      Caption         =   "Iniciar el juego"
      Default         =   -1  'True
      Height          =   495
      Left            =   8160
      TabIndex        =   19
      Top             =   1320
      Width           =   1575
   End
   Begin VB.TextBox txtnombre 
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
      Left            =   7080
      TabIndex        =   18
      Top             =   840
      Width           =   3615
   End
   Begin VB.Timer Tiempo 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   9960
      Top             =   1920
   End
   Begin VB.Timer Timerpta 
      Interval        =   2
      Left            =   600
      Top             =   3720
   End
   Begin VB.TextBox txt9 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   72
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   4560
      TabIndex        =   8
      Top             =   5040
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txt8 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   72
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   3000
      TabIndex        =   7
      Top             =   5040
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txt7 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   72
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   1440
      TabIndex        =   6
      Top             =   5040
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txt6 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   72
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   4560
      TabIndex        =   5
      Top             =   3480
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txt5 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   72
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   3000
      TabIndex        =   4
      Top             =   3480
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txt4 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   72
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   1440
      TabIndex        =   3
      Top             =   3480
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txt3 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   72
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   4560
      TabIndex        =   2
      Top             =   1920
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txt2 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   72
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   3000
      TabIndex        =   1
      Top             =   1920
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txt1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   72
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   1440
      TabIndex        =   0
      Top             =   1920
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Line punta 
      BorderColor     =   &H000000FF&
      BorderWidth     =   15
      X1              =   5640
      X2              =   5880
      Y1              =   6840
      Y2              =   7320
   End
   Begin VB.Line punta1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   15
      X1              =   6000
      X2              =   6480
      Y1              =   7320
      Y2              =   6960
   End
   Begin VB.Line punta2 
      BorderColor     =   &H000000FF&
      BorderWidth     =   15
      X1              =   5640
      X2              =   6480
      Y1              =   6720
      Y2              =   6960
   End
   Begin VB.Line linea 
      BorderColor     =   &H000000FF&
      BorderWidth     =   10
      Visible         =   0   'False
      X1              =   7320
      X2              =   5880
      Y1              =   1440
      Y2              =   7440
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Puntuación:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   7320
      TabIndex        =   33
      Top             =   5760
      Width           =   1575
   End
   Begin VB.Label lblmal 
      BackStyle       =   0  'Transparent
      Caption         =   "Ya fuiste!: Mayor de 12 minutos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   7920
      TabIndex        =   32
      Top             =   7200
      Width           =   2775
   End
   Begin VB.Label lblmaomenos 
      BackStyle       =   0  'Transparent
      Caption         =   "Mas o menos!: Hasta 12 minutos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   375
      Left            =   7920
      TabIndex        =   31
      Top             =   6840
      Width           =   2775
   End
   Begin VB.Label lblbien 
      BackStyle       =   0  'Transparent
      Caption         =   "Bien!: Hasta 7 minutos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   7920
      TabIndex        =   30
      Top             =   6480
      Width           =   2775
   End
   Begin VB.Label lblexcelente 
      BackStyle       =   0  'Transparent
      Caption         =   "Excelente!: Hasta 2 minutos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   375
      Index           =   0
      Left            =   7920
      TabIndex        =   29
      Top             =   6120
      Width           =   2775
   End
   Begin VB.Shape shapeoculto 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00808000&
      FillStyle       =   0  'Solid
      Height          =   5895
      Left            =   120
      Top             =   1320
      Width           =   6735
   End
   Begin VB.Label lblnombredeljugador 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   6840
      TabIndex        =   28
      Top             =   2400
      Width           =   4095
   End
   Begin VB.Label lbl3 
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   735
      Left            =   7680
      TabIndex        =   24
      Top             =   1800
      Width           =   735
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   735
      Left            =   9360
      TabIndex        =   22
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label lbl2 
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   735
      Left            =   8520
      TabIndex        =   23
      Top             =   1800
      Width           =   735
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   ":     :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   495
      Left            =   8280
      TabIndex        =   25
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Pon tu nombre:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   4800
      TabIndex        =   20
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmjuego.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   855
      Left            =   120
      TabIndex        =   17
      Top             =   120
      Width           =   10695
   End
   Begin VB.Label h1 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   480
      TabIndex        =   16
      Top             =   6600
      Width           =   1095
   End
   Begin VB.Label diagonal 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6120
      TabIndex        =   15
      Top             =   6600
      Width           =   975
   End
   Begin VB.Label h4 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4920
      TabIndex        =   14
      Top             =   6600
      Width           =   975
   End
   Begin VB.Label h3 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3360
      TabIndex        =   13
      Top             =   6600
      Width           =   975
   End
   Begin VB.Label h2 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1800
      TabIndex        =   12
      Top             =   6600
      Width           =   975
   End
   Begin VB.Label v3 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6120
      TabIndex        =   11
      Top             =   5520
      Width           =   975
   End
   Begin VB.Label v2 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6120
      TabIndex        =   10
      Top             =   3960
      Width           =   975
   End
   Begin VB.Label v1 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   6120
      TabIndex        =   9
      Top             =   2280
      Width           =   975
   End
   Begin VB.Shape Shape1 
      Height          =   4815
      Left            =   1320
      Top             =   1800
      Width           =   4815
   End
End
Attribute VB_Name = "frmjuego"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Declare Function Getkeypress Lib "user32" Alias "GetAsyncKeyState" (ByVal key As Long) As Integer

Private Sub tecla_Timer()

If Getkeypress(vbKey1) Then
cmd1.Value = True
End If

If Getkeypress(vbKey2) Then
cmd2.Value = True
End If

If Getkeypress(vbKey3) Then
cmd3.Value = True
End If

If Getkeypress(vbKey4) Then
cmd4.Value = True
End If

If Getkeypress(vbKey5) Then
cmd5.Value = True
End If

If Getkeypress(vbKey6) Then
cmd6.Value = True
End If

If Getkeypress(vbKey7) Then
cmd7.Value = True
End If

If Getkeypress(vbKey8) Then
cmd8.Value = True
End If

If Getkeypress(vbKey9) Then
cmd9.Value = True
End If
End Sub

Private Sub Tiempo_Timer()
lbl1.Caption = Val(lbl1.Caption) + 1
If lbl1.Caption = "100" Then
lbl1.Caption = "00"
lbl2.Caption = "0" & Val(lbl2.Caption) + 1
  If lbl2.Caption > 9 Then
  lbl2.Caption = Val(lbl2.Caption)
  End If
End If
If lbl2.Caption = "60" Then
lbl2.Caption = "00"
lbl3.Caption = "0" & Val(lbl3.Caption) + 1
  If lbl3.Caption > 9 Then
  lbl3.Caption = Val(lbl3.Caption)
  End If
End If
End Sub

Private Sub cmd1_Click()
If opt1.Value = True Then
txt1.Text = " 1"
ElseIf opt2.Value = True Then
txt2.Text = " 1"
ElseIf opt3.Value = True Then
txt3.Text = " 1"
ElseIf opt4.Value = True Then
txt4.Text = " 1"
ElseIf opt5.Value = True Then
txt5.Text = " 1"
ElseIf opt6.Value = True Then
txt6.Text = " 1"
ElseIf opt7.Value = True Then
txt7.Text = " 1"
ElseIf opt8.Value = True Then
txt8.Text = " 1"
ElseIf opt9.Value = True Then
txt9.Text = " 1"
End If
End Sub

Private Sub cmd2_Click()
If opt1.Value = True Then
txt1.Text = " 2"
ElseIf opt2.Value = True Then
txt2.Text = " 2"
ElseIf opt3.Value = True Then
txt3.Text = " 2"
ElseIf opt4.Value = True Then
txt4.Text = " 2"
ElseIf opt5.Value = True Then
txt5.Text = " 2"
ElseIf opt6.Value = True Then
txt6.Text = " 2"
ElseIf opt7.Value = True Then
txt7.Text = " 2"
ElseIf opt8.Value = True Then
txt8.Text = " 2"
ElseIf opt9.Value = True Then
txt9.Text = " 2"
End If
End Sub

Private Sub cmd3_Click()
If opt1.Value = True Then
txt1.Text = " 3"
ElseIf opt2.Value = True Then
txt2.Text = " 3"
ElseIf opt3.Value = True Then
txt3.Text = " 3"
ElseIf opt4.Value = True Then
txt4.Text = " 3"
ElseIf opt5.Value = True Then
txt5.Text = " 3"
ElseIf opt6.Value = True Then
txt6.Text = " 3"
ElseIf opt7.Value = True Then
txt7.Text = " 3"
ElseIf opt8.Value = True Then
txt8.Text = " 3"
ElseIf opt9.Value = True Then
txt9.Text = " 3"
End If
End Sub

Private Sub cmd4_Click()
If opt1.Value = True Then
txt1.Text = " 4"
ElseIf opt2.Value = True Then
txt2.Text = " 4"
ElseIf opt3.Value = True Then
txt3.Text = " 4"
ElseIf opt4.Value = True Then
txt4.Text = " 4"
ElseIf opt5.Value = True Then
txt5.Text = " 4"
ElseIf opt6.Value = True Then
txt6.Text = " 4"
ElseIf opt7.Value = True Then
txt7.Text = " 4"
ElseIf opt8.Value = True Then
txt8.Text = " 4"
ElseIf opt9.Value = True Then
txt9.Text = " 4"
End If
End Sub

Private Sub cmd5_Click()
If opt1.Value = True Then
txt1.Text = " 5"
ElseIf opt2.Value = True Then
txt2.Text = " 5"
ElseIf opt3.Value = True Then
txt3.Text = " 5"
ElseIf opt4.Value = True Then
txt4.Text = " 5"
ElseIf opt5.Value = True Then
txt5.Text = " 5"
ElseIf opt6.Value = True Then
txt6.Text = " 5"
ElseIf opt7.Value = True Then
txt7.Text = " 5"
ElseIf opt8.Value = True Then
txt8.Text = " 5"
ElseIf opt9.Value = True Then
txt9.Text = " 5"
End If
End Sub

Private Sub cmd6_Click()
If opt1.Value = True Then
txt1.Text = " 6"
ElseIf opt2.Value = True Then
txt2.Text = " 6"
ElseIf opt3.Value = True Then
txt3.Text = " 6"
ElseIf opt4.Value = True Then
txt4.Text = " 6"
ElseIf opt5.Value = True Then
txt5.Text = " 6"
ElseIf opt6.Value = True Then
txt6.Text = " 6"
ElseIf opt7.Value = True Then
txt7.Text = " 6"
ElseIf opt8.Value = True Then
txt8.Text = " 6"
ElseIf opt9.Value = True Then
txt9.Text = " 6"
End If
End Sub

Private Sub cmd7_Click()
If opt1.Value = True Then
txt1.Text = " 7"
ElseIf opt2.Value = True Then
txt2.Text = " 7"
ElseIf opt3.Value = True Then
txt3.Text = " 7"
ElseIf opt4.Value = True Then
txt4.Text = " 7"
ElseIf opt5.Value = True Then
txt5.Text = " 7"
ElseIf opt6.Value = True Then
txt6.Text = " 7"
ElseIf opt7.Value = True Then
txt7.Text = " 7"
ElseIf opt8.Value = True Then
txt8.Text = " 7"
ElseIf opt9.Value = True Then
txt9.Text = " 7"
End If
End Sub

Private Sub cmd8_Click()
If opt1.Value = True Then
txt1.Text = " 8"
ElseIf opt2.Value = True Then
txt2.Text = " 8"
ElseIf opt3.Value = True Then
txt3.Text = " 8"
ElseIf opt4.Value = True Then
txt4.Text = " 8"
ElseIf opt5.Value = True Then
txt5.Text = " 8"
ElseIf opt6.Value = True Then
txt6.Text = " 8"
ElseIf opt7.Value = True Then
txt7.Text = " 8"
ElseIf opt8.Value = True Then
txt8.Text = " 8"
ElseIf opt9.Value = True Then
txt9.Text = " 8"
End If
End Sub

Private Sub cmd9_Click()
If opt1.Value = True Then
txt1.Text = " 9"
ElseIf opt2.Value = True Then
txt2.Text = " 9"
ElseIf opt3.Value = True Then
txt3.Text = " 9"
ElseIf opt4.Value = True Then
txt4.Text = " 9"
ElseIf opt5.Value = True Then
txt5.Text = " 9"
ElseIf opt6.Value = True Then
txt6.Text = " 9"
ElseIf opt7.Value = True Then
txt7.Text = " 9"
ElseIf opt8.Value = True Then
txt8.Text = " 9"
ElseIf opt9.Value = True Then
txt9.Text = " 9"
End If
End Sub

Private Sub cmdayuda_Click()
MsgBox "Piensa !!! ps jaja!", vbInformation, "Ayuda"
End Sub

Private Sub cmdborrar_Click()
If opt1.Value = True Then
txt1.Text = ""
ElseIf opt2.Value = True Then
txt2.Text = ""
ElseIf opt3.Value = True Then
txt3.Text = ""
ElseIf opt4.Value = True Then
txt4.Text = ""
ElseIf opt5.Value = True Then
txt5.Text = ""
ElseIf opt6.Value = True Then
txt6.Text = ""
ElseIf opt7.Value = True Then
txt7.Text = ""
ElseIf opt8.Value = True Then
txt8.Text = ""
ElseIf opt9.Value = True Then
txt9.Text = ""
End If
End Sub

Private Sub cmdfinal_Click()
Tiempo.Enabled = False
frmpuntaje.Show
cmdlimpiar.Enabled = False
cmdborrar.Enabled = False
cmd1.Enabled = False
cmd2.Enabled = False
cmd3.Enabled = False
cmd4.Enabled = False
cmd5.Enabled = False
cmd6.Enabled = False
cmd7.Enabled = False
cmd8.Enabled = False
cmd9.Enabled = False
fra123.Enabled = False
If txt1.Text = " 5" And txt2.Text = " 5" And txt3.Text = " 5" And txt4.Text = " 5" And txt5.Text = " 5" And txt6.Text = " 5" And txt7.Text = " 5" And txt8.Text = " 5" And txt9.Text = " 5" Then
frmpuntaje.lblpuntaje.Caption = "Ja ja, que picón(a) que eres " & txtnombre.Text & ", pierdes por hacer trampa"
ElseIf lbl3.Caption < "02" Then
frmpuntaje.lblpuntaje.Caption = "Execelente " & txtnombre.Text & ", bien Hecho"
ElseIf lbl3.Caption < "07" Then
frmpuntaje.lblpuntaje.Caption = "Bien " & txtnombre.Text & ", para la proxima vez puede ser mejor"
ElseIf lbl3.Caption < "12" Then
frmpuntaje.lblpuntaje.Caption = "Mas o menos " & txtnombre.Text & ", pero pudiste hacerlo mucho mejor"
ElseIf lbl3.Caption > "12" Then
frmpuntaje.lblpuntaje.Caption = "Ya fuiste " & txtnombre.Text & ", sobrepasaste el tiempo, que pena"
End If
'al puntaje
If frmpuntaje.lblpuntaje.Caption = "Ja ja, que picón(a) que eres " & txtnombre.Text & ", pierdes por hacer trampa" Then
frmpuntuacion.lstpuntaje.AddItem txtnombre.Text & "          " & "Tramposo!!" & "    " & lbl3.Caption & ":" & lbl2.Caption & ":" & lbl1.Caption
ElseIf lbl3.Caption < "01" Then
frmpuntuacion.lstpuntaje.AddItem txtnombre.Text & "          " & "10pts" & "         " & lbl3.Caption & ":" & lbl2.Caption & ":" & lbl1.Caption
ElseIf lbl3.Caption < "01" Then
frmpuntuacion.lstpuntaje.AddItem txtnombre.Text & "          " & "9pts" & "          " & lbl3.Caption & ":" & lbl2.Caption & ":" & lbl1.Caption
ElseIf lbl3.Caption < "02" Then
frmpuntuacion.lstpuntaje.AddItem txtnombre.Text & "          " & "8pts" & "          " & lbl3.Caption & ":" & lbl2.Caption & ":" & lbl1.Caption
ElseIf lbl3.Caption < "04" Then
frmpuntuacion.lstpuntaje.AddItem txtnombre.Text & "          " & "7pts" & "          " & lbl3.Caption & ":" & lbl2.Caption & ":" & lbl1.Caption
ElseIf lbl3.Caption < "05" Then
frmpuntuacion.lstpuntaje.AddItem txtnombre.Text & "          " & "6pts" & "          " & lbl3.Caption & ":" & lbl2.Caption & ":" & lbl1.Caption
ElseIf lbl3.Caption < "07" Then
frmpuntuacion.lstpuntaje.AddItem txtnombre.Text & "          " & "5pts" & "          " & lbl3.Caption & ":" & lbl2.Caption & ":" & lbl1.Caption
ElseIf lbl3.Caption < "08" Then
frmpuntuacion.lstpuntaje.AddItem txtnombre.Text & "          " & "4pts" & "          " & lbl3.Caption & ":" & lbl2.Caption & ":" & lbl1.Caption
ElseIf lbl3.Caption < "10" Then
frmpuntuacion.lstpuntaje.AddItem txtnombre.Text & "          " & "3pts" & "          " & lbl3.Caption & ":" & lbl2.Caption & ":" & lbl1.Caption
ElseIf lbl3.Caption < "12" Then
frmpuntuacion.lstpuntaje.AddItem txtnombre.Text & "          " & "2pts" & "          " & lbl3.Caption & ":" & lbl2.Caption & ":" & lbl1.Caption
ElseIf lbl3.Caption < "12" Then
frmpuntuacion.lstpuntaje.AddItem txtnombre.Text & "          " & "1pt" & "           " & lbl3.Caption & ":" & lbl2.Caption & ":" & lbl1.Caption
End If

End Sub

Private Sub cmdinicio_Click()
If txtnombre.Text = "" Then
MsgBox "Coloque su nombre, Por favor...", vbExclamation
Else
txtnombre.Enabled = False
Tiempo.Enabled = True
cmdinicio.Enabled = False
tecla.Enabled = True
lblnombredeljugador.Caption = txtnombre.Text
txt1.Visible = True
txt2.Visible = True
txt3.Visible = True
txt4.Visible = True
txt5.Visible = True
txt6.Visible = True
txt7.Visible = True
txt8.Visible = True
txt9.Visible = True
shapeoculto.Visible = False
cmdayuda.Enabled = True
fra123.Enabled = True
cmd1.Enabled = True
cmd2.Enabled = True
cmd3.Enabled = True
cmd4.Enabled = True
cmd5.Enabled = True
cmd6.Enabled = True
cmd7.Enabled = True
cmd8.Enabled = True
cmd9.Enabled = True
cmdborrar.Enabled = True
cmdlimpiar.Enabled = True
End If

End Sub

Private Sub cmdlimpiar_Click()
txt1.Text = ""
txt2.Text = ""
txt3.Text = ""
txt4.Text = ""
txt5.Text = ""
txt6.Text = ""
txt7.Text = ""
txt8.Text = ""
txt9.Text = ""
End Sub

Private Sub cmdre_Click()
mensaje = "¿Seguro que desea volver a comenzar?"
estilo = vbQuestion & vbYesNo
titulo = "Volver a jugar"
rpta = MsgBox(mensaje, estilo, titulo)
If rpta = vbYes Then
txtnombre = ""
txtnombre.Enabled = True
shapeoculto.Visible = True
fra123.Enabled = False
Tiempo.Enabled = False
lbl1.Caption = "00"
lbl2.Caption = "00"
lbl3.Caption = "00"
cmdlimpiar.Value = True
tecla.Enabled = False
linea.Visible = True
punta.Visible = True
punta1.Visible = True
punta2.Visible = True
lblnombredeljugador.Caption = ""
txt1.Visible = False
txt2.Visible = False
txt3.Visible = False
txt4.Visible = False
txt5.Visible = False
txt6.Visible = False
txt7.Visible = False
txt8.Visible = False
txt9.Visible = False
cmdayuda.Enabled = False
cmdfinal.Enabled = False
cmd1.Enabled = False
cmd2.Enabled = False
cmd3.Enabled = False
cmd4.Enabled = False
cmd5.Enabled = False
cmd6.Enabled = False
cmd7.Enabled = False
cmd8.Enabled = False
cmd9.Enabled = False
cmdlimpiar.Enabled = False
cmdborrar.Enabled = False
cmdinicio.Enabled = True
ElseIf rpta = vbNo Then
cmdre.SetFocus
End If
End Sub

Private Sub cmdsalir_Click()

'salida en juego
If Tiempo.Enabled = True Then
  mensaje = "¿Seguro que desea Salir sin haber terminado de jugar? eso quiere decir que se rinde"
  estilo = vbQuestion & vbYesNo
  titulo = "Se cachorrea????"
  rpta = MsgBox(mensaje, estilo, titulo)
  If rpta = vbYes Then
  MsgBox "Qué mal!!! se perdieron los datos guardados por no seguir el juego!"
  End
  ElseIf rpta = vbNo Then
  cmdsalir.SetFocus
  End If
'salida normal
ElseIf Tiempo.Enabled = False Then
  mensaje = "¿Seguro que desea Salir? si sale se perderán todos los datos"
  estilo = vbQuestion & vbYesNo
  titulo = "Desea Salir??"
  rpta = MsgBox(mensaje, estilo, titulo)
  If rpta = vbYes Then
  End
  ElseIf rpta = vbNo Then
  cmdsalir.SetFocus
  End If
End If
End Sub


Private Sub Command1_Click()
frmpuntuacion.Show
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
End
End Sub


Private Sub lblnombredeljugador_Click()
lblbien.ForeColor = RGB(204, 255, 0)
lblmaomenos.ForeColor = RGB(255, 204, 0)
lblmal.ForeColor = RGB(255, 51, 51)
End Sub


Private Sub Timerpta_Timer()
v1.Caption = Val(txt1.Text) + Val(txt2.Text) + Val(txt3.Text)
v2.Caption = Val(txt4.Text) + Val(txt5.Text) + Val(txt6.Text)
v3.Caption = Val(txt7.Text) + Val(txt8.Text) + Val(txt9.Text)
h1.Caption = Val(txt7.Text) + Val(txt5.Text) + Val(txt3.Text)
h2.Caption = Val(txt1.Text) + Val(txt4.Text) + Val(txt7.Text)
h3.Caption = Val(txt2.Text) + Val(txt5.Text) + Val(txt8.Text)
h4.Caption = Val(txt3.Text) + Val(txt6.Text) + Val(txt9.Text)
diagonal.Caption = Val(txt9.Text) + Val(txt5.Text) + Val(txt1.Text)

If v1.Caption = "15" And v2.Caption = "15" And v3.Caption = "15" And h1.Caption = "15" And h2.Caption = "15" And h3.Caption = "15" And h4.Caption = "15" And diagonal.Caption = "15" Then
cmdfinal.Enabled = True
cmdfinal.BackColor = RGB(204, 255, 255)
linea.Visible = True
punta.Visible = True
punta1.Visible = True
punta2.Visible = True
Else
cmdfinal.Enabled = False
cmdfinal.BackColor = -2147483633
linea.Visible = False
punta.Visible = False
punta1.Visible = False
punta2.Visible = False
End If

If Tiempo.Enabled = False Then
cmdre.Caption = "Volver a jugar"
ElseIf Tiempo.Enabled = True Then
cmdre.Caption = "Me Rindo!"
End If

'tiempo detenido
If Tiempo.Enabled = False Then
linea.Visible = False
punta.Visible = False
punta1.Visible = False
punta2.Visible = False
End If

End Sub

Private Sub Timerverificador_Timer()
If txt1.Text > "9" And txt1.Text < "1" Then
MsgBox "Colocar solo numeros del 1 al 9 sin repetirlos", vbInformation
txt1.Text = " "
Else
txt1.Text = txt1.Text
End If
If txt2.Text > "9" And txt2.Text < "1" Then
MsgBox "Colocar solo numeros del 1 al 9 sin repetirlos", vbInformation
txt2.Text = " "
Else
txt2.Text = txt2.Text
End If
If txt3.Text > "9" And txt3.Text < "1" Then
MsgBox "Colocar solo numeros del 1 al 9 sin repetirlos", vbInformation
txt3.Text = " "
Else
txt3.Text = txt3.Text
End If
If txt4.Text > "9" And txt4.Text < "1" Then
MsgBox "Colocar solo numeros del 1 al 9 sin repetirlos", vbInformation
txt4.Text = " "
Else
txt4.Text = txt4.Text
End If
If txt5.Text > "9" And txt5.Text < "1" Then
MsgBox "Colocar solo numeros del 1 al 9 sin repetirlos", vbInformation
txt5.Text = " "
Else
txt5.Text = txt5.Text
End If
If txt6.Text > "9" And txt6.Text < "1" Then
MsgBox "Colocar solo numeros del 1 al 9 sin repetirlos", vbInformation
txt6.Text = " "
Else
txt6.Text = txt6.Text
End If
If txt7.Text > "9" And txt7.Text < "1" Then
MsgBox "Colocar solo numeros del 1 al 9 sin repetirlos", vbInformation
txt7.Text = " "
Else
txt7.Text = txt7.Text
End If
If txt8.Text > "9" And txt8.Text < "1" Then
MsgBox "Colocar solo numeros del 1 al 9 sin repetirlos", vbInformation
txt8.Text = " "
Else
txt8.Text = txt8.Text
End If
If txt9.Text > "9" And txt9.Text < "1" Then
MsgBox "Colocar solo numeros del 1 al 9 sin repetirlos", vbInformation
txt9.Text = " "
Else
txt9.Text = txt9.Text
End If

End Sub
