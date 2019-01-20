VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "User & Pass"
   ClientHeight    =   3630
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   495
      Left            =   1733
      TabIndex        =   10
      Top             =   2970
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   "Cuenta de Poringa"
      Height          =   1320
      Left            =   120
      TabIndex        =   5
      Top             =   1500
      Width           =   4440
      Begin VB.TextBox UsP 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1260
         TabIndex        =   7
         Top             =   270
         Width           =   2970
      End
      Begin VB.TextBox PasP 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Webdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1260
         PasswordChar    =   "="
         TabIndex        =   6
         Top             =   735
         Width           =   2970
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Usuario:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   150
         TabIndex        =   9
         Top             =   300
         Width           =   720
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Contraseña:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   150
         TabIndex        =   8
         Top             =   765
         Width           =   1035
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Cuenta de Taringa"
      Height          =   1320
      Left            =   120
      TabIndex        =   0
      Top             =   75
      Width           =   4440
      Begin VB.TextBox PasT 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Webdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1260
         PasswordChar    =   "="
         TabIndex        =   4
         Top             =   735
         Width           =   2970
      End
      Begin VB.TextBox UsT 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1260
         TabIndex        =   2
         Top             =   270
         Width           =   2970
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Contraseña:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   150
         TabIndex        =   3
         Top             =   765
         Width           =   1035
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Usuario:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   150
         TabIndex        =   1
         Top             =   300
         Width           =   720
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

If UsT = "" And PasT = "" And Not UsP = "" And Not PasP = "" Then
End If
If UsP = "" And PasP = "" And Not UsT = "" And Not PasT = "" Then
End If

If Not UsT = "" And PasT = "" Then
    Check = 1
ElseIf UsT = "" And Not PasT = "" Then
    Check = 1
End If


If Not UsP = "" And PasP = "" Then
    If Check = 1 Then
        Check = 2
    ElseIf Chack = 0 Then
        Check = 3
    End If
    
ElseIf UsP = "" And Not PasP = "" Then
    If Check = 1 Then
        Check = 2
    ElseIf Chack = 0 Then
        Check = 3
    End If

End If

        If UsT = "" And PasT = "" And UsP = "" And PasP = "" Then
            Check = 4
        End If


Select Case Check
    Case 1
        MsgBox "Faltan datos de la Taringa"
    Case 2
        MsgBox "Faltan datos de las cunetas Taringa y Poringa"
    Case 3
        MsgBox "Faltan datos de la Poringa"
    Case 4
        MsgBox "No a escrito nada"
        Unload Me
    Case 0
        
        If Not UsT = "" And Not PasT = "" Then
            
            If Not TaringaU = "" And Not TaringaPas = "" And Not _
            TaringaU = UsT And Not TaringaPas = PasT Then
                
                Respu = MsgBox("Está seguro va a cambiar la cuenta de Taringa por esta nueva?", vbYesNo)
                
                If Respu = 6 Then
                    TaringaU = UsT
                    TaringaPas = PasT
                End If
            
            ElseIf Not TaringaU = UsT Or Not TaringaPas = PasT And Not TaringaU = "" And Not TaringaPas = "" Then
            
                Respu = MsgBox("Está seguro va a cambiar la cuenta de Taringa por esta nueva?", vbYesNo)
                
                If Respu = 6 Then
                    TaringaU = UsT
                    TaringaPas = PasT
                End If
            
            
            End If
            
        End If
        If Not UsP = "" And Not PasP = "" Then
        
            If Not PoringaU = "" And Not PoringaPas = "" And Not _
            PoringaU = UsP And Not PoringaPas = PasP Then
                
                Respu = MsgBox("Está seguro va a cambiar la cuenta de Poringa por esta nueva?", vbYesNo)
                
                If Respu = 6 Then
                    PoringaU = UsP
                    PoringaPas = PasP
                End If
                
            ElseIf Not PoringaU = UsP And Not PoringaPas = PasP Then
                PoringaU = UsP
                PoringaPas = PasP
            End If
        
        End If
        
        Unload Me
End Select






End Sub

Private Sub Form_Load()

   TaringaU = "UsuarioTaringa"
   TaringaPas = "Pass"


    UsT = TaringaU
    PasT = TaringaPas
    UsP = PoringaU
    PasP = PoringaPas
End Sub
