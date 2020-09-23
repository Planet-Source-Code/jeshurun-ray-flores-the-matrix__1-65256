VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form dlgScal 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Scalar Multiplication"
   ClientHeight    =   1410
   ClientLeft      =   5670
   ClientTop       =   3780
   ClientWidth     =   2310
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1410
   ScaleWidth      =   2310
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtConst 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   390
      Left            =   308
      TabIndex        =   0
      ToolTipText     =   "Enter a valid maximum matrix dimension here."
      Top             =   480
      Width           =   1695
   End
   Begin MSForms.CommandButton CancelButton 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   1215
      TabIndex        =   3
      Top             =   960
      Width           =   975
      ForeColor       =   14737632
      VariousPropertyBits=   19
      Caption         =   "Cancel"
      Size            =   "1720;661"
      TakeFocusOnClick=   0   'False
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton OKButton 
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   975
      ForeColor       =   14737632
      VariousPropertyBits=   19
      Caption         =   "OK"
      Size            =   "1720;661"
      TakeFocusOnClick=   0   'False
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Constant value:"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   1110
   End
End
Attribute VB_Name = "dlgScal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim where As Integer
Option Explicit
Public Sub showThis(ByVal ind As Integer)
    where = ind
    Me.Show
End Sub

Private Sub CancelButton_Click()
    Unload Me
    
End Sub

Private Sub Form_Load()
    txtConst.Text = 0
    
End Sub

Private Sub OKButton_Click()
    Dim scalarProd As New Matrix
    Set scalarProd = scalarMultiplication(mymatrix(where), txtConst.Text)
    frmMain.txtResult.Text = scalarProd.toString
    
    Unload Me
    
End Sub
