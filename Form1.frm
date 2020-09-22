VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Stenography Sample"
   ClientHeight    =   6330
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8085
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6330
   ScaleWidth      =   8085
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   90
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Top             =   1125
      Width           =   7860
   End
   Begin VB.PictureBox Picture2 
      Height          =   3915
      Left            =   4050
      ScaleHeight     =   3855
      ScaleWidth      =   3840
      TabIndex        =   7
      Top             =   1935
      Width           =   3900
   End
   Begin VB.PictureBox Picture1 
      Height          =   3915
      Left            =   75
      ScaleHeight     =   3855
      ScaleWidth      =   3840
      TabIndex        =   5
      Top             =   1950
      Width           =   3900
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Read"
      Height          =   360
      Left            =   6930
      TabIndex        =   3
      Top             =   5910
      Width           =   1035
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   315
      Left            =   45
      TabIndex        =   2
      Top             =   5940
      Width           =   3960
      _ExtentX        =   6985
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Blend"
      Height          =   360
      Left            =   5850
      TabIndex        =   1
      Top             =   5910
      Width           =   1035
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   90
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   315
      Width           =   7860
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Message retrieved:"
      Height          =   195
      Left            =   105
      TabIndex        =   10
      Top             =   855
      Width           =   1350
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "After:"
      Height          =   195
      Left            =   4050
      TabIndex        =   8
      Top             =   1665
      Width           =   375
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Before:"
      Height          =   195
      Left            =   75
      TabIndex        =   6
      Top             =   1680
      Width           =   510
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Message to be stored:"
      Height          =   195
      Left            =   105
      TabIndex        =   4
      Top             =   45
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents cBlend As clsBlend
Attribute cBlend.VB_VarHelpID = -1

Private Sub cBlend_Progress(Value As Integer)
    ProgressBar1.Value = Value
End Sub

Private Sub Command1_Click()
    Set cBlend = New clsBlend
        
    Picture1.Picture = LoadPicture(App.Path & "\Soap Bubbles.bmp")
    Call cBlend.Blend(App.Path & "\Soap Bubbles.bmp", App.Path & "\message.bmp", App.Path & "\blend.key", Text1.Text)
    ProgressBar1.Value = 0
    Picture2.Picture = LoadPicture(App.Path & "\message.bmp")
    MsgBox "Done."
    Set cBlend = Nothing
End Sub

Private Sub Command2_Click()
    Dim sTemp As String
    
    Set cBlend = New clsBlend
    
    Text2.Text = ""
    Call cBlend.Read(App.Path & "\message.bmp", App.Path & "\blend.key", sTemp)
    Text2.Text = sTemp
    ProgressBar1.Value = 0
    MsgBox "Done."
    Set cBlend = Nothing
End Sub

Private Sub Form_Load()
    Text1.Text = "This is a test message..."
End Sub
