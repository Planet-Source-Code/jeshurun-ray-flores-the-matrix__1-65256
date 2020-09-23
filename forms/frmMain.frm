VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "The Matrix"
   ClientHeight    =   7410
   ClientLeft      =   825
   ClientTop       =   1185
   ClientWidth     =   11220
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7410
   ScaleWidth      =   11220
   Begin MSComctlLib.StatusBar statbar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   7155
      Width           =   11220
      _ExtentX        =   19791
      _ExtentY        =   450
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.ListBox opsList 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   1500
      ItemData        =   "frmMain.frx":0000
      Left            =   4350
      List            =   "frmMain.frx":0019
      TabIndex        =   4
      Top             =   720
      Width           =   2535
   End
   Begin MSForms.TextBox txtResult 
      Height          =   2415
      Left            =   3720
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   4080
      Width           =   3885
      VariousPropertyBits=   -1610594281
      BackColor       =   0
      ForeColor       =   12632256
      BorderStyle     =   1
      ScrollBars      =   3
      Size            =   "6853;4260"
      BorderColor     =   8421504
      SpecialEffect   =   0
      FontName        =   "Lucida Console"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   34
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdPerform 
      Default         =   -1  'True
      Height          =   615
      Left            =   4350
      TabIndex        =   8
      Top             =   2280
      Width           =   2535
      ForeColor       =   12632256
      BackColor       =   0
      Caption         =   "Perform Selected Operation"
      Size            =   "4471;1085"
      TakeFocusOnClick=   0   'False
      FontName        =   "Arial Narrow"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.TextBox txtMat1 
      Height          =   2415
      Left            =   400
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   480
      Width           =   3375
      VariousPropertyBits=   -1610594281
      BackColor       =   14737632
      ForeColor       =   12632256
      BorderStyle     =   1
      ScrollBars      =   3
      Size            =   "5953;4260"
      BorderColor     =   8421504
      SpecialEffect   =   0
      FontName        =   "Lucida Console"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   34
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdProp2 
      Height          =   495
      Left            =   8040
      TabIndex        =   7
      Top             =   3000
      Width           =   2295
      ForeColor       =   14737632
      BackColor       =   -2147483630
      Caption         =   "Properties of Matrix B"
      Size            =   "4048;873"
      TakeFocusOnClick=   0   'False
      FontName        =   "Arial Narrow"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdProp1 
      Height          =   495
      Left            =   1080
      TabIndex        =   6
      Top             =   3000
      Width           =   2295
      ForeColor       =   14737632
      BackColor       =   -2147483630
      Caption         =   "Properties of Matrix A"
      Size            =   "4048;873"
      TakeFocusOnClick=   0   'False
      FontName        =   "Arial Narrow"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.TextBox txtMat2 
      Height          =   2415
      Left            =   7440
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   480
      Width           =   3375
      VariousPropertyBits=   -1610594281
      BackColor       =   0
      ForeColor       =   12632256
      BorderStyle     =   1
      ScrollBars      =   3
      Size            =   "5953;4260"
      BorderColor     =   8421504
      SpecialEffect   =   0
      FontName        =   "Lucida Console"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   34
      ParagraphAlign  =   3
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select an operation:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   4320
      TabIndex        =   0
      Top             =   480
      Width           =   1935
   End
   Begin MSForms.Image Image3 
      Height          =   7980
      Left            =   -50
      Top             =   -570
      Width           =   11340
      Size            =   "20002;14076"
      Picture         =   "frmMain.frx":00E1
   End
   Begin VB.Menu mufile 
      Caption         =   "File"
      Index           =   0
      Begin VB.Menu misett 
         Caption         =   "Settings"
      End
      Begin VB.Menu miexit 
         Caption         =   "EXIT"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdPerform_Click()
    Dim index As Integer
    Dim result As New Matrix
    index = opsList.ListIndex
    Dim com As String
    com = check(index)
    If com <> "" Then
        MsgBox com, vbExclamation, "ERROR"
        Exit Sub
    End If
    
    Select Case index
        Case 0, 1, 2, 3, 4
            Select Case index
            Case 0
                Set result = Addition(mymatrix(1), mymatrix(2))
            Case 1
                Set result = Subtraction(mymatrix(1), mymatrix(2))
            Case 2
                Set result = Subtraction(mymatrix(2), mymatrix(1))
            Case 3
                Set result = Multiplication(mymatrix(1), mymatrix(2))
            Case 4
                Set result = Multiplication(mymatrix(2), mymatrix(1))
            End Select
            txtResult.Text = result.toString
            
        Case 5, 6
            dlgScal.showThis index - 4
            
    End Select
    
End Sub

Private Sub cmdProp1_Click()
    dlgProp.showThis 1
    
End Sub

Private Sub cmdProp2_Click()
    dlgProp.showThis 2
    
End Sub

Private Sub Form_Load()
    opsList.ListIndex = 0
    max = 10
   
    mymatrix(1).Initialize 1, 1
    mymatrix(2).Initialize 1, 1
    
    txtMat1.Text = mymatrix(1).toString
    txtMat2.Text = mymatrix(2).toString
    statbar.SimpleText = "Current maximum square matrix dimension is " & max & "."
End Sub

Private Sub miexit_Click()
    Unload Me
    
End Sub

Private Sub misett_Click()
    dlgSettings.Show vbModal
    
End Sub

