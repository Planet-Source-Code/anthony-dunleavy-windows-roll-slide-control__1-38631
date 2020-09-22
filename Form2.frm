VERSION 5.00
Begin VB.Form Form2 
   AutoRedraw      =   -1  'True
   Caption         =   "CoolAnimate Example"
   ClientHeight    =   3975
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5505
   LinkTopic       =   "Form2"
   ScaleHeight     =   3975
   ScaleWidth      =   5505
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Animate Now"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3975
      TabIndex        =   11
      Top             =   2760
      Width           =   1245
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "Form2.frx":0000
      Left            =   2160
      List            =   "Form2.frx":0034
      TabIndex        =   10
      Text            =   "1000"
      Top             =   2880
      Width           =   825
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Form2.frx":008F
      Left            =   1665
      List            =   "Form2.frx":00B4
      TabIndex        =   9
      Text            =   "None"
      Top             =   2490
      Width           =   2025
   End
   Begin Project1.CoolAnimate CoolAnimate1 
      Left            =   4080
      Top             =   3555
      _ExtentX        =   2461
      _ExtentY        =   529
      CloseWindowEffect=   4
   End
   Begin VB.Label Label2 
      Caption         =   "Delay (in Milliseconds)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   165
      TabIndex        =   8
      Top             =   2940
      Width           =   2235
   End
   Begin VB.Label Label2 
      Caption         =   "Animation Style"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   165
      TabIndex        =   7
      Top             =   2520
      Width           =   1485
   End
   Begin VB.Label Label2 
      Caption         =   "For This Example"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   165
      TabIndex        =   6
      Top             =   1455
      Width           =   4335
   End
   Begin VB.Label Label2 
      Caption         =   "Using the CoolAnimate Control"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   165
      TabIndex        =   5
      Top             =   135
      Width           =   4335
   End
   Begin VB.Label Label2 
      Caption         =   "Step 2) Set the Animation Style and Delay of the CoolAnimate Control"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   165
      TabIndex        =   4
      Top             =   585
      Width           =   5325
   End
   Begin VB.Label Label2 
      Caption         =   "CoolAnimate1.AnimateClose"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   645
      TabIndex        =   3
      Top             =   1035
      Width           =   4500
   End
   Begin VB.Label Label2 
      Caption         =   "Step 3) Place the following code inside the Form_Unload Event"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   165
      TabIndex        =   2
      Top             =   810
      Width           =   4875
   End
   Begin VB.Label Label2 
      Caption         =   "Step 1) Drop the CoolAnimate Control onto a form"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   165
      TabIndex        =   1
      Top             =   375
      Width           =   4335
   End
   Begin VB.Label Label1 
      Caption         =   $"Form2.frx":0125
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   165
      TabIndex        =   0
      Top             =   1680
      Width           =   5295
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

    Unload Me
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    CoolAnimate1.CloseWindowEffect = Combo1.ListIndex
    CoolAnimate1.EffectDelay = Val(Combo2.Text)
    CoolAnimate1.AnimateClose
    
End Sub

