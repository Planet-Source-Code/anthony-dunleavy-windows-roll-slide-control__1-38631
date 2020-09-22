VERSION 5.00
Begin VB.UserControl CoolAnimate 
   Appearance      =   0  'Flat
   BackColor       =   &H00B49800&
   BorderStyle     =   1  'Fixed Single
   CanGetFocus     =   0   'False
   ClientHeight    =   405
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1920
   HasDC           =   0   'False
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   405
   ScaleWidth      =   1920
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cool Animate"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   30
      Width           =   1125
   End
End
Attribute VB_Name = "CoolAnimate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Enum Animation

    None
    Fade
    SlideRight
    SlideLeft
    SlideUp
    SlideDown
    RollRight
    RollLeft
    RollUp
    RollDown
    ZoomToNothing

End Enum
    

Private Declare Function AnimateWindow Lib "user32" _
                    (ByVal hwnd As Long, ByVal dwTime As Long, _
                     ByVal dwFlags As Long) As Boolean

'Property Variables:
Dim m_EffectDelay As Variant
Dim m_CloseWindowEffect As Animation

'Default Property Values:
Const m_def_EffectDelay = 1000
Const m_def_CloseWindowEffect = RollUp


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get EffectDelay() As Variant

    EffectDelay = m_EffectDelay
    
End Property

Public Property Let EffectDelay(ByVal New_EffectDelay As Variant)

    m_EffectDelay = New_EffectDelay
    PropertyChanged "EffectDelay"
    
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get CloseWindowEffect() As Animation

    CloseWindowEffect = m_CloseWindowEffect
    
End Property

Public Property Let CloseWindowEffect(ByVal New_CloseWindowEffect As Animation)
    
    m_CloseWindowEffect = New_CloseWindowEffect
    PropertyChanged "CloseWindowEffect"
    
End Property


'Initialize Properties for User Control
Private Sub UserControl_InitProperties()

    m_EffectDelay = m_def_EffectDelay
    m_CloseWindowEffect = m_def_CloseWindowEffect

    UserControl.Width = 1400
    UserControl.Height = 300
    
End Sub


'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_EffectDelay = PropBag.ReadProperty("EffectDelay", m_def_EffectDelay)
    m_CloseWindowEffect = PropBag.ReadProperty("CloseWindowEffect", m_def_CloseWindowEffect)

End Sub

Private Sub UserControl_Resize()

    UserControl.Width = 1400
    UserControl.Height = 300
    
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("EffectDelay", m_EffectDelay, m_def_EffectDelay)
    Call PropBag.WriteProperty("CloseWindowEffect", m_CloseWindowEffect, m_def_CloseWindowEffect)

End Sub

Private Sub CoolAnimate(ByVal WindowEffect)

    Dim Effect As Long

    Select Case WindowEffect

    Case Fade: Effect = &H80000 Or &H10000
    Case SlideRight: Effect = &H1 Or &H40000 Or &H10000
    Case SlideLeft: Effect = &H2 Or &H40000 Or &H10000
    Case SlideDown: Effect = &H4 Or &H40000 Or &H10000
    Case SlideUp: Effect = &H8 Or &H40000 Or &H10000
    Case RollRight: Effect = &H1 Or &H10000
    Case RollLeft: Effect = &H2 Or &H10000
    Case RollDown: Effect = &H4 Or &H10000
    Case RollUp: Effect = &H8 Or &H10000
    Case ZoomToNothing: Effect = &H10 Or &H10000

    End Select
    
    AnimateWindow UserControl.ContainerHwnd, _
                  EffectDelay, _
                  Effect
    
    
End Sub


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14
Public Function AnimateClose() As Variant

    CoolAnimate CloseWindowEffect
    
End Function

