VERSION 5.00
Begin VB.Form frmpuntuacion 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Puntuaciones"
   ClientHeight    =   7095
   ClientLeft      =   4365
   ClientTop       =   2025
   ClientWidth     =   9090
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7095
   ScaleWidth      =   9090
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdaceptar 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Aceptar"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6000
      Width           =   2055
   End
   Begin VB.CommandButton cmdlimpiar 
      Caption         =   "Limpiar toda la lista de puntuación"
      Height          =   495
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6240
      Width           =   2415
   End
   Begin VB.CommandButton cmdborrar 
      Caption         =   "Borrar el puntaje"
      Height          =   495
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6240
      Width           =   2295
   End
   Begin VB.ListBox lstpuntaje 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Left            =   360
      TabIndex        =   1
      Top             =   1800
      Width           =   8295
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Tiempo:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   855
      Left            =   6120
      TabIndex        =   4
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Puntaje:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   735
      Left            =   3600
      TabIndex        =   3
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Nombres:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   855
      Left            =   600
      TabIndex        =   2
      Top             =   1200
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Puntuaciones:"
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
      Height          =   855
      Left            =   2760
      TabIndex        =   0
      Top             =   360
      Width           =   3375
   End
End
Attribute VB_Name = "frmpuntuacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdaceptar_Click()
frmpuntuacion.Hide
End Sub

Private Sub cmdborrar_Click()
If lstpuntaje.ListIndex <> -1 Then
lstpuntaje.RemoveItem lstpuntaje.ListIndex
End If
End Sub

Private Sub cmdlimpiar_Click()
mensaje = "¿Seguro que desea borrar todas las puntuaciones?"
estilo = vbQuestion & vbYesNo
titulo = "Limpiar"
rpta = MsgBox(mensaje, estilo, titulo)
If rpta = vbYes Then
lstpuntaje.Clear
ElseIf rpta = vbNo Then
cmdlimpiar.SetFocus
End If
End Sub

Private Sub Form_Load()
cmdborrar.BackColor = RGB(255, 255, 204)
cmdlimpiar.BackColor = RGB(153, 204, 255)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
frmpuntuacion.Hide
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmpuntuacion.Hide
End Sub
