VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form dlgProp 
   BackColor       =   &H00000000&
   Caption         =   "Matrix Properties"
   ClientHeight    =   675
   ClientLeft      =   2385
   ClientTop       =   1500
   ClientWidth     =   8115
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   675
   ScaleMode       =   0  'User
   ScaleWidth      =   7955.881
   Begin MSMask.MaskEdBox txtR 
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   120
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   0
      PromptInclude   =   0   'False
      MaxLength       =   12
      Mask            =   "############"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtC 
      Height          =   375
      Left            =   2280
      TabIndex        =   3
      Top             =   120
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   0
      PromptInclude   =   0   'False
      MaxLength       =   12
      Mask            =   "############"
      PromptChar      =   "_"
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      X1              =   117.647
      X2              =   7882.351
      Y1              =   600
      Y2              =   600
   End
   Begin MSForms.CommandButton CancelButton 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   6840
      TabIndex        =   6
      Top             =   120
      Width           =   1215
      ForeColor       =   14737632
      VariousPropertyBits=   19
      Caption         =   "Cancel"
      Size            =   "2143;661"
      TakeFocusOnClick=   0   'False
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton OKButton 
      Default         =   -1  'True
      Height          =   375
      Left            =   5520
      TabIndex        =   5
      Top             =   120
      Width           =   1215
      ForeColor       =   14737632
      VariousPropertyBits=   19
      Caption         =   "OK"
      Size            =   "2143;661"
      TakeFocusOnClick=   0   'False
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdChange 
      Height          =   375
      Left            =   3120
      TabIndex        =   4
      Top             =   120
      Width           =   2295
      ForeColor       =   14737632
      VariousPropertyBits=   19
      Caption         =   "Change Matrix Size"
      Size            =   "4048;661"
      TakeFocusOnClick=   0   'False
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "COLUMN:"
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   1440
      TabIndex        =   1
      Top             =   210
      Width           =   735
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ROW:"
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   450
   End
End
Attribute VB_Name = "dlgProp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'ht = 980
'width = 8250

Dim where As Integer
Dim newmatrix As New Matrix

Option Explicit


Public Sub showThis(ByVal ind As Integer)
    
    Me.Caption = "Properties of Matrix " & ind

    showMatrixFields mymatrix(ind).getRow, mymatrix(ind).getColumn
    openMatrix mymatrix(ind)
    newmatrix.clone mymatrix(ind)
    txtR.Text = mymatrix(ind).getRow
    txtC.Text = mymatrix(ind).getColumn
    where = ind
    Me.Show vbModal
    
End Sub

Private Sub CancelButton_Click()
    Unload Me
    
End Sub

Private Sub cmdChange_Click()
    Dim row As Integer, col As Integer
    If CInt(txtR.Text) > max Or CInt(txtC.Text) > max Then
        MsgBox "Invalid input!" & vbCrLf & vbCrLf & _
               "Please enter within the range" & vbCrLf & _
               " of the maximum dimension.", vbExclamation, _
               "Matrix dimension error."
    Else
        hideAll
        showMatrixFields CInt(txtR.Text), CInt(txtC.Text)
        newmatrix.Initialize CInt(txtR.Text), CInt(txtC.Text)
        openMatrix newmatrix
        openMatrix mymatrix(where)
        
    End If
    
    
End Sub

Private Sub Form_Load()
    ReDim elem(1 To max, 1 To max) As TextBox
    addTillMax
    
End Sub

Private Sub OKButton_Click()
    hideAll
    showMatrixFields txtR.Text, txtC.Text
    On Error GoTo onError
    initializeThis newmatrix
    openMatrix newmatrix
    newmatrix.Initialize txtR.Text, txtC.Text
    initializeThis newmatrix
    mymatrix(where).clone newmatrix
    frmMain.txtMat1.Text = mymatrix(1).toString
    frmMain.txtMat2.Text = mymatrix(2).toString
    
    Unload Me
    Exit Sub
    
onError:
    MsgBox "Some inputs are invalid." & _
            "Please review inputs.", vbExclamation
    Exit Sub
    
End Sub
