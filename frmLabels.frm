VERSION 5.00
Begin VB.Form frmLabels 
   Caption         =   "Label Effects"
   ClientHeight    =   4890
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6870
   LinkTopic       =   "Form1"
   ScaleHeight     =   4890
   ScaleWidth      =   6870
   StartUpPosition =   3  'Windows Default
   Begin VB.Label lblShadow 
      BackStyle       =   0  'Transparent
      Caption         =   "Shadow"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Index           =   0
      Left            =   720
      TabIndex        =   1
      Top             =   2640
      Width           =   5535
   End
   Begin VB.Label lblBorder 
      BackStyle       =   0  'Transparent
      Caption         =   "Border"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Index           =   0
      Left            =   1080
      TabIndex        =   0
      Top             =   240
      Width           =   4815
   End
End
Attribute VB_Name = "frmLabels"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Enum LabelEffects
    lblEffectShadow = 0
    lblEffectBorder = 1
End Enum

Sub AddLabelEffect(lbl As Object, _
    iEffectSize As Integer, _
    cEffectColor As ColorConstants, _
    cTextColor As ColorConstants, _
    Optional LabelEffect As LabelEffects = lblEffectShadow)
    
    Dim iMoveRight As Integer
    Dim iMoveDown As Integer
    Dim iCounter As Integer
    Dim iIterator As Integer
    Dim iStep As Integer
    
    'Unload all existing occurences of the Label
    For iIterator = 1 To lbl.Count - 1
        Unload lbl(iIterator)
    Next iIterator
    
    If LabelEffect = lblEffectShadow Then
        'Add Shadow (i.e. New labels to the top and left
        For iIterator = 1 To iEffectSize
            Load lbl(iIterator)
            lbl(iIterator).ForeColor = cEffectColor
            lbl(iIterator).Left = lbl(iIterator - 1).Left - 1
            lbl(iIterator).Top = lbl(iIterator - 1).Top - 1
            lbl(iIterator).Visible = True
        Next iIterator
    ElseIf LabelEffect = lblEffectBorder Then
        'Add a border (i.e. new labels all around the existing label.
        iCounter = 1
        Me.Show
        For iMoveRight = -1 To 1
            For iMoveDown = -1 To 1
                For iIterator = 1 To iEffectSize
                    Load lbl(iCounter)
                    lbl(iCounter).Left = lbl(0).Left + (iIterator * iMoveRight)
                    x = lbl(0).Left
                    x = lbl(iCounter).Left
                    lbl(iCounter).Top = lbl(0).Top + (iIterator * iMoveDown)
                    lbl(iCounter).ForeColor = cEffectColor
                    lbl(iCounter).Visible = True
                    iCounter = iCounter + 1
                Next iIterator
            Next iMoveDown
        Next iMoveRight
    End If
    lbl(0).ForeColor = cTextColor
End Sub

Private Sub Form_Load()
    AddLabelEffect lblShadow, 40, vbBlack, vbYellow, lblEffectShadow
    AddLabelEffect lblBorder, 20, vbBlack, vbYellow, lblEffectBorder
End Sub

