VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form dlgSettings 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Matrices Settings"
   ClientHeight    =   1440
   ClientLeft      =   1215
   ClientTop       =   2055
   ClientWidth     =   2325
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1440
   ScaleWidth      =   2325
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtmax 
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
      Left            =   315
      TabIndex        =   1
      ToolTipText     =   "Enter a valid maximum matrix dimension here."
      Top             =   480
      Width           =   1695
   End
   Begin MSForms.CommandButton CancelButton 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   1200
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
      Caption         =   "Maximum Matrix Dimension:"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1950
   End
End
Attribute VB_Name = "dlgSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CancelButton_Click()
    Unload Me
    
End Sub

Private Sub Form_Load()
    txtmax.Text = max
    
End Sub

Private Sub OKButton_Click()
    If IsNumeric(txtmax.Text) = False Or CInt(txtmax.Text) = 0 Then
        MsgBox "Invalid input!", vbCritical, "ERROR"
        Exit Sub
    Else
        Dim choice As Integer
        Dim tmpmax As Integer
        tmpmax = txtmax.Text
        
        If tmpmax < mymatrix(1).getColumn Or tmpmax < mymatrix(1).getRow Or _
            tmpmax < mymatrix(2).getColumn Or tmpmax < mymatrix(2).getRow Then
            Dim str As String
            str = "The program has detected that there exists a matrix having a " & _
                "dimension greater than the newly specified dimension limit. " & vbCrLf & _
                "Continuing will reset the matrix with the dimension greater " & _
                "than the specified maximum dimension." & vbCrLf & vbCrLf & _
                "Do you want to continue?"
            choice = MsgBox(str, vbYesNoCancel + vbQuestion, "Verification")
            Select Case choice
                Case 6
                    If tmpmax < mymatrix(1).getColumn Or tmpmax < mymatrix(1).getRow Then
                        mymatrix(1).Initialize 1, 1
                        frmMain.txtMat1.Text = mymatrix(1).toString
                    End If
                    If tmpmax < mymatrix(2).getColumn Or tmpmax < mymatrix(2).getRow Then
                        mymatrix(2).Initialize 1, 1
                        frmMain.txtMat2.Text = mymatrix(2).toString
                    End If
                    
                    max = txtmax.Text
                    frmMain.statbar.SimpleText = "Current maximum square matrix dimension is " & max & "."
                    Unload Me
                    
                Case 7
                    Unload Me
                Case 2
                    Exit Sub
            End Select
        End If
        Unload Me
    End If
End Sub
