VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form Form1 
   BackColor       =   &H004C4C4C&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Robate POST de Taringa"
   ClientHeight    =   7155
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9285
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "Form1.frx":038A
   ScaleHeight     =   7155
   ScaleWidth      =   9285
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton User_pass_buton 
      Caption         =   "User Pass"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4725
      TabIndex        =   28
      Top             =   3585
      Width           =   1215
   End
   Begin VB.CheckBox Tagcheck 
      Caption         =   "Descargar Tag"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   495
      Left            =   6300
      TabIndex        =   26
      Top             =   3495
      Width           =   1905
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H80000005&
      Caption         =   "STOP"
      Height          =   675
      Left            =   4395
      TabIndex        =   25
      Top             =   195
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Ataque de un solo Post"
      Height          =   570
      Index           =   1
      Left            =   1830
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   1920
      Width           =   1560
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Ataque masivo"
      Height          =   570
      Index           =   0
      Left            =   540
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   1920
      Value           =   -1  'True
      Width           =   1290
   End
   Begin VB.Frame Frame 
      Height          =   1665
      Index           =   1
      Left            =   547
      TabIndex        =   21
      Top             =   2370
      Visible         =   0   'False
      Width           =   8190
      Begin VB.TextBox Text6 
         Height          =   315
         Left            =   195
         TabIndex        =   22
         Text            =   "/posts/downloads/1905629/lo-que-sea.html                          Esto es un Ejemplo"
         Top             =   780
         Width           =   7380
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Dirección URL"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   5250
         TabIndex        =   24
         Top             =   270
         Width           =   1050
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00C00000&
         X1              =   3795
         X2              =   3795
         Y1              =   375
         Y2              =   540
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00C00000&
         X1              =   7635
         X2              =   7635
         Y1              =   375
         Y2              =   540
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C00000&
         X1              =   3795
         X2              =   7650
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Especifique la dirección del Post Ejemplo: /posts/downloads/1905629/lo-que-sea.html"
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
         Left            =   195
         TabIndex        =   23
         Top             =   540
         Width           =   7380
      End
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Form1.frx":68CC
      Left            =   7020
      List            =   "Form1.frx":68D6
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   1905
      Width           =   1725
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   150
      TabIndex        =   11
      Text            =   "60"
      Top             =   735
      Width           =   2130
   End
   Begin VB.PictureBox Picture2 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   765
      Left            =   0
      ScaleHeight     =   765
      ScaleWidth      =   9285
      TabIndex        =   7
      Top             =   6390
      Width           =   9285
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "INFO: -NONE-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   45
         TabIndex        =   8
         Top             =   45
         Width           =   8700
         WordWrap        =   -1  'True
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   9285
      TabIndex        =   5
      Top             =   5835
      Width           =   9285
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "POST: -NONE-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   45
         TabIndex        =   6
         Top             =   45
         Width           =   8700
         WordWrap        =   -1  'True
      End
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000005&
      Caption         =   "Atacar a TARINGA CARAJOS"
      Height          =   675
      Left            =   2685
      TabIndex        =   0
      Top             =   195
      Width           =   1695
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1875
      Top             =   6150
   End
   Begin InetCtlsObjects.Inet Inet2 
      Left            =   1080
      Top             =   6015
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      AccessType      =   1
      URL             =   "http://"
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   360
      Top             =   6030
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      AccessType      =   1
      URL             =   "http://"
   End
   Begin VB.TextBox Text4 
      Height          =   2475
      Left            =   2775
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   4
      Top             =   6120
      Visible         =   0   'False
      Width           =   2640
   End
   Begin VB.TextBox Text3 
      Height          =   2475
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Top             =   6120
      Visible         =   0   'False
      Width           =   2640
   End
   Begin VB.TextBox Text2 
      Height          =   2475
      Left            =   2715
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      Top             =   6000
      Visible         =   0   'False
      Width           =   2640
   End
   Begin VB.TextBox Text1 
      Height          =   2475
      Left            =   5685
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   6165
      Visible         =   0   'False
      Width           =   2640
   End
   Begin VB.Frame Frame 
      Height          =   1665
      Index           =   0
      Left            =   547
      TabIndex        =   17
      Top             =   2370
      Width           =   8190
      Begin VB.TextBox Cate 
         Height          =   315
         Left            =   195
         TabIndex        =   19
         Text            =   "/posts/downloads/"
         Top             =   780
         Width           =   7380
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Especifique la categoria Ejemplo: /posts/downloads/ Exactamente igual con sus /"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   195
         TabIndex        =   20
         Top             =   540
         Width           =   7020
      End
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Una Venezuela Libre"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3210
      TabIndex        =   27
      Top             =   4860
      Width           =   2850
   End
   Begin VB.Image Image2 
      Height          =   1500
      Left            =   4125
      Picture         =   "Form1.frx":68EC
      Top             =   975
      Width           =   1500
   End
   Begin VB.Image Image1 
      Height          =   1560
      Left            =   0
      Picture         =   "Form1.frx":DE5E
      Top             =   4215
      Width           =   9285
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Especifique el dominio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   390
      Left            =   5895
      TabIndex        =   15
      Top             =   1860
      Width           =   1020
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "tec.web44.net"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   555
      Left            =   5295
      TabIndex        =   13
      Top             =   105
      Width           =   3165
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Especifique el numero de paginas de siguiente de la lista de Post"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   585
      Left            =   135
      TabIndex        =   12
      Top             =   105
      Width           =   2250
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PAGINA: 0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   5625
      TabIndex        =   9
      Top             =   705
      Width           =   1095
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "POST REPETIDOS: 0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   5625
      TabIndex        =   10
      Top             =   990
      Width           =   1860
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Function UTF8_Decode(ByVal sStr As String)
    
    Dim l As Long, sUTF8 As String, iChar As Integer, iChar2 As Integer
    
    For l = 1 To Len(sStr)
        iChar = Asc(Mid(sStr, l, 1))
        
        If iChar > 127 Then
            
            If Not iChar And 32 Then ' 2 chars
                iChar2 = Asc(Mid(sStr, l + 1, 1))
                sUTF8 = sUTF8 & ChrW$(((31 And iChar) * 64 + (63 And iChar2)))
                l = l + 1
            Else
                Dim iChar3 As Integer
                iChar2 = Asc(Mid(sStr, l + 1, 1))
                iChar3 = Asc(Mid(sStr, l + 2, 1))
                sUTF8 = sUTF8 & ChrW$(((iChar And 15) * 16 * 256) + ((iChar2 And 63) * 64) + (iChar3 And 63))
                l = l + 2
            End If
        
        Else
            sUTF8 = sUTF8 & Chr$(iChar)
        End If
    
    Next l
    
    UTF8_Decode = sUTF8

End Function

Function FindIt_NOMBRE(TEXTO As String) As String

'Dim Fso As New Scripting.FileSystemObject
'Dim f As File
'<title>Pagina nueva 1</title>
Esperate = True

On Error Resume Next

    gFindStringput = "<h1 property=" & Chr(34) & "dc:title" & Chr(34) & ">"
    Dim start, pos, findstring, sourcestring, Msg, Response, Offset

Static ya As Boolean
Static ya2 As Boolean
    
InI:

    If (gCurPos = Text4.SelStart) Then
        Offset = 1
    Else
        Offset = 0
    End If

    If gFirstTime Then Offset = 0
    start = 0 + Offset
        
    
    If gFindCase Then
        findstring = gFindStringput
        sourcestring = Text4.Text
    Else
        
        findstring = gFindStringput
        sourcestring = TEXTO
    
    End If
            
    If gFindDirection = 0 Then
        pos = InStr(start + 1, sourcestring, findstring)
    Else
        For pos = start - 1 To 0 Step -1
            If pos = 0 Then Exit For
            ddd = Mid(sourcestring, pos, Len(findstring))
            If Mid(sourcestring, pos, Len(findstring)) = findstring Then Exit For
        Next
    End If
  
    If ya = True And Not ya2 = True Then
    gFindDirection = 0
    gFindString = gFindStringput
    ya2 = True
    
        GoTo InI
    End If
    ya2 = False
    gFirstTime = False
    posti = pos
    pos = pos + 24
        
    For I = 0 To Len(TEXTO) - pos
                ddd = Mid(sourcestring, pos + I, 5)
                
            If Mid(sourcestring, pos + I, 5) = "</h1>" Then
                    
                    
                    ggg = Mid(sourcestring, pos, I)
                    ggg = Replace(ggg, "/", "")
                    ggg = Replace(ggg, ":", "")
                    ggg = Replace(ggg, "*", "x")
                    ggg = Replace(ggg, "?", "")
                    ggg = Replace(ggg, "\", "")
                    ggg = Replace(ggg, "|", "")
                    ggg = Replace(ggg, ">", "")
                    ggg = Replace(ggg, "<", "")
                    'ggg = Replace(ggg, "+", " mas ")
                    
                    ggg = Replace(ggg, Chr(34), "")
                    ggg = Replace(ggg, vbNewLine, "")
                    ggg = Replace(ggg, vbCr, "")
                    
                    
                    TITLES = ggg
                   
                    FindIt_NOMBRE = ggg
                    Text4.Text = Replace(Text4.Text, "<h1 property=" & Chr(34) & "dc:title" & Chr(34) & ">", "<h1>")
                    Esperate = False
                    
                Exit For
            End If
    Next I
    
End Function
Function Limpiar(TEXTO As String) As String

Esperate = True

On Error Resume Next

    gFindStringput = "<span>"
    Dim start, pos, findstring, sourcestring, Msg, Response, Offset

Static ya As Boolean
Static ya2 As Boolean
    
InI:

    If (gCurPos = Text4.SelStart) Then
        Offset = 1
    Else
        Offset = 0
    End If

    If gFirstTime Then Offset = 0
    start = 0 + Offset
        
    
    If gFindCase Then
        findstring = gFindStringput
        sourcestring = Text4.Text
    Else
        
        findstring = gFindStringput
        sourcestring = TEXTO
    
    End If
            
    If gFindDirection = 0 Then
        pos = InStr(start + 1, sourcestring, findstring)
    Else
        For pos = start - 1 To 0 Step -1
            If pos = 0 Then Exit For
            ddd = Mid(sourcestring, pos, Len(findstring))
            If Mid(sourcestring, pos, Len(findstring)) = findstring Then Exit For
        Next
    End If
  
    If ya = True And Not ya2 = True Then
    gFindDirection = 0
    gFindString = gFindStringput
    ya2 = True
    
        GoTo InI
    End If
    ya2 = False
    gFirstTime = False
        
        Limpiar = Right(sourcestring, Len(sourcestring) - pos + 1)
    
End Function

Sub FindIt(TEXTO As String)
    gFindStringput = "href=" & Chr(34)
    Dim start, pos, findstring, sourcestring, Msg, Response, Offset

Static ya As Boolean
Static ya2 As Boolean
    
InI:

    If (gCurPos = Text1.SelStart) Then
        Offset = 1
    Else
        Offset = 0
    End If

    If gFirstTime Then Offset = 0
    start = 0 + Offset
        
    
    If gFindCase Then
        findstring = gFindStringput
        sourcestring = Text1.Text
    Else
        
        findstring = gFindStringput
        sourcestring = TEXTO
    
    End If
            
    If gFindDirection = 0 Then
        pos = InStr(start + 1, sourcestring, findstring)
    Else
        For pos = start - 1 To 0 Step -1
            If pos = 0 Then Exit For
            ddd = Mid(sourcestring, pos, Len(findstring))
            If Mid(sourcestring, pos, Len(findstring)) = findstring Then Exit For
        Next
    End If
  
    If ya = True And Not ya2 = True Then
    gFindDirection = 0
    gFindString = gFindStringput
    ya2 = True
    
        GoTo InI
    End If
    ya2 = False
    gFirstTime = False
    pos = pos + 6
    
    For I = 0 To Len(TEXTO) - pos
                
            If Mid(sourcestring, pos + I, 1) = Chr(34) Then
                    II = II + 1
                    If Combo1.ListIndex = 0 Then
                        linea = "http://www.taringa.net" & Mid(sourcestring, pos, I)
                    Else
                        linea = "http://www.poringa.net" & Mid(sourcestring, pos, I)
                    End If
                    ReDim Preserve LinesdePost(II - 1)
                    LinesdePost(II - 1) = linea
                Exit For
            End If
    Next I
    
End Sub

Private Sub Matriz_de_Titulos_de_los_post()
On Error GoTo FIN
'Dim MiMatriz() As Integer
'ReDim Preserve MiMatriz(15)
'Text2.SelStart = 2
'MsgBox Text2.Text
Dim lineas() As String

Dim I%
lineas() = Split(Text2.Text, vbCrLf)
For I = 0 To UBound(lineas)
    If Not InStr(1, lineas(I), "href=") = 0 Then
        FindIt (lineas(I))
    End If
Next

For I = 0 To UBound(LinesdePost)
    
    Do While Inet2.StillExecuting
        DoEvents
    Loop
    
    If StopB = False Then
        Call Inet2.Execute(LinesdePost(I), "GET")
    Else
        StopB = False: Text5.Enabled = True: Label5.Enabled = True: Label4.Caption = "POST REPETIDOS: 0": Label3.Caption = "PAGINA: 0"
        Option1(0).Value = True: Option1(1).Enabled = True: Option1(0).Enabled = True: Frame(0).Enabled = True: Frame(1).Enabled = True
        Combo1.Enabled = True: Cate.Enabled = True: Command1.Caption = "Atacar a TARINGA CARAJOS": Command1.Enabled = True
        Command2.Visible = False: Command2.Enabled = True: Esperate = False: numerodenagina = 0: Label3.Caption = "PAGINA: "
        Text1.Text = "": Text2.Text = "": Text3.Text = "": Text4.Text = "": II = 0: existen = 0: ReDim LinesdePost(0)
        Siguiente = False: Tagcheck.Enabled = True: User_pass_buton.Enabled = True: Exit Sub
    End If
    
Next I
If Not numerodenagina >= Text5.Text Then
    Siguiente = True
Else
    numerodenagina = 0
    Siguiente = True
End If

GoTo OK

FIN:
    End
OK:

End Sub


Private Sub Sacar_post()

    Esperate = True
Text4.Text = ""
Dim t() As String
Dim inicio As Boolean

t = Split(Text3.Text, vbCrLf)
For cuenta = 0 To UBound(t)

If Combo1.ListIndex = 0 Then
    'inipost = "<hr />"
    inipost = "<span property=" & Chr(34) & "dc:content"
    
Else
    'inipost = "<br />"
    inipost = "<span property=" & Chr(34) & "dc:content"
End If
        
        
If Not InStr(1, t(cuenta), "<!-- Cuerpo -->", vbTextCompare) = 0 Or inicio = True Or inicio2 = True Then
    
If InStr(1, t(cuenta), "<script", vbTextCompare) = 0 And Not inicio3 = True Then
    
    
    If InStr(1, t(cuenta), "<div class=" & Chr(34) & "post-contenido" & Chr(34) & ">", vbTextCompare) = 0 And Not inicio2 = True Then
            If InStr(1, t(cuenta), "<div class=" & Chr(34) & "compartir-mov" & Chr(34), vbTextCompare) = 0 Then
                If InStr(1, t(cuenta), "next.php?id", vbTextCompare) = 0 And InStr(1, t(cuenta), "<a href=" & Chr(34) & "/prev.php?id=", vbTextCompare) = 0 And InStr(1, t(cuenta), "<!-- Cuerpo -->", vbTextCompare) = 0 Then
                    Text4.Text = Text4.Text & t(cuenta) & vbCrLf
                End If
            Else
                Exit For
            End If
            inicio = True

            
    ElseIf Not InStr(1, t(cuenta), inipost, vbTextCompare) = 0 Or Not InStr(1, t(cuenta), "<!-- Cuerpo -->", vbTextCompare) = 0 Then
        inicio2 = False
        Text4.Text = Text4.Text & "<span>" & vbCrLf
    Else
        If Not inicio2 = True Then
            Text4.Text = Text4.Text & "<div>" & vbCrLf
            inicio2 = True
        End If
    End If
Else
    If Not InStr(1, t(cuenta), "</script", vbTextCompare) = 0 Then
    inicio3 = False
    Else
    inicio3 = True
    End If

End If
    
    End If



Next cuenta


Text4.Text = Replace(Text4.Text, "Ã¡", "á")
Text4.Text = Replace(Text4.Text, "Ã©", "é")
Text4.Text = Replace(Text4.Text, "Ã" & Chr(173), "í")
Text4.Text = Replace(Text4.Text, "Ã³", "ó")
Text4.Text = Replace(Text4.Text, "Ãº", "ú")
Text4.Text = Replace(Text4.Text, "Ã" & Chr(129), "Á")
Text4.Text = Replace(Text4.Text, "Ã‰", "É")
Text4.Text = Replace(Text4.Text, "Ã" & Chr(141), "Í")
Text4.Text = Replace(Text4.Text, "Ã“", "Ó")
Text4.Text = Replace(Text4.Text, "Ãš", "Ú")
Text4.Text = Replace(Text4.Text, "Ã±", "ñ")
Text4.Text = Replace(Text4.Text, "Ã‘", "Ñ")
Text4.Text = Replace(Text4.Text, "Ãº", "º")
Text4.Text = Replace(Text4.Text, "Ãª", "ª")
Text4.Text = Replace(Text4.Text, "Â¿", "¿")
Text4.Text = Replace(Text4.Text, "Â®", "®")
Text4.Text = Replace(Text4.Text, "Â¡", "¡")
Text4.Text = Replace(Text4.Text, "â€“", "-")
Text4.Text = Replace(Text4.Text, "&amp;", " y ")



Text4.Text = Replace(Text4.Text, "<div class=" & Chr(34) & "post-contenedor" & Chr(34) & ">", "<div>")
Text4.Text = Replace(Text4.Text, "<div class=" & Chr(34) & "post-title" & Chr(34) & ">", "<div>")
'Text4.Text = Replace(Text4.Text, "<div class=" & Chr(34) & "post-contenido" & Chr(34) & ">", "<div>")
Text4.Text = Replace(Text4.Text, "<!-- Cuerpo -->", "")


Text4.Text = Replace(Text4.Text, vbLf, vbCrLf)
Text4.Text = Replace(Text4.Text, "http://links.itaringa.net/out?", "")

If Tagcheck = 1 Then
    Sacar_Tag_de_los_post
    Text4.Text = Text4.Text & vbCrLf & Tag
Else
    Text4.Text = Text4.Text
End If

Dim canalLibre As Integer

Nombre = FindIt_NOMBRE(Text4.Text)
Text4.Text = Limpiar(Text4.Text)

canalLibre = FreeFile

Label1.Caption = "POST: " & Nombre



If Len(Nombre) >= 226 Then
    Nombre = Left(Nombre, 221)
End If

existeono = Dir$(App.Path & "\Taringa POST\" & Nombre & ".html")

If Not existeono = "" Then
    existen = existen + 1
    Label2.Caption = "INFO: " & "Ya fue bajado el POST: " & Chr(34) & Nombre & Chr(34)
    Label4.Caption = "POST REPETIDOS:" & existen
    no = True
End If

If Not no = True And Not Nombre = "" Then
Open App.Path & "\Taringa POST\" & Nombre & ".html" For Output As #canalLibre
dd = dd + 1
'Escribimos el contenido del TextBox al fichero

'Print #canalLibre, "<TITLE>" & Nombre & "</TITLE>" & vbCrLf & Text4 & vbCrLf

Print #canalLibre, Text4 & vbCrLf


Close #canalLibre
    Esperate = False
Me.Caption = "Robate POST de Taringa -- ATACANDO... POS DESCARGADO: " & dd
Label2.Caption = "INFO: "
End If

If Option1(1).Value = True Then
    MsgBox "Listo", , "Robate POST de Taringa"
    
        Combo1.Enabled = True
        Cate.Enabled = True
        Option1(0).Value = False
        Option1(0).Enabled = True
        Option1(1).Value = True
        Option1(1).Enabled = True
        Command1.Enabled = True
        Command1.Caption = "Atacar a TARINGA CARAJOS"
        Label1.Caption = "POST: -NONE-"
        Label2.Caption = "INFO: -NONE-"
        Frame(0).Enabled = True
        Frame(1).Enabled = True
        Tagcheck.Enabled = True
        User_pass_buton.Enabled = True
End If


End Sub

Private Sub Command1_Click()
        
        Combo1.Enabled = False
        Text5.Enabled = False
        Cate.Enabled = False
        Option1(0).Enabled = False
        Option1(1).Enabled = False
        Frame(0).Enabled = False
        Frame(1).Enabled = False
        Tagcheck.Enabled = False
        User_pass_buton.Enabled = False
        
        If Option1(0).Value = True Then
        
            If Combo1.ListIndex = 0 Then
veriuserT1:

checkcue = checkcue + 1
                
                If TaringaU = "" Or TaringaPas = "" Then
                
If checkcue = 2 Then
    Text5.Enabled = True: Label5.Enabled = True: Label4.Caption = "POST REPETIDOS: 0": Label3.Caption = "PAGINA: 0"
    Option1(0).Value = True: Option1(1).Enabled = True: Option1(0).Enabled = True: Frame(0).Enabled = True
    Frame(1).Enabled = True: Combo1.Enabled = True: Cate.Enabled = True: Command1.Caption = "Atacar a TARINGA CARAJOS"
    Command1.Enabled = True: Command2.Visible = False: Command2.Enabled = True: Esperate = False: numerodenagina = 0
    Label3.Caption = "PAGINA: ": Text1.Text = "": Text2.Text = "": Text3.Text = "": Text4.Text = "": II = 0: existen = 0
    ReDim LinesdePost(0): Siguiente = False: Tagcheck.Enabled = True: Exit Sub
End If
                
                    MsgBox "No a especificado la cuanta de Taringa"
                    Form2.Show 1
                    GoTo veriuserT1
                Else
                
                If Not login = True Then Call Inet2.Execute("http://www.taringa.net/login.php", "POST", "nick=" & TaringaU & "&pass=" & TaringaPas, "Content-Type: application/x-www-form-urlencoded")
                    Do While Inet2.StillExecuting
                        DoEvents
                    Loop
                End If

                
                Call Inet1.Execute("http://www.taringa.net" & Cate, "GET")
                Command2.Visible = True
            Else
            
veriuserP1:
                
                
checkcue = checkcue + 1
                
                If PoringaU = "" Or PoringaPas = "" Then
                
If checkcue = 2 Then
    Text5.Enabled = True: Label5.Enabled = True: Label4.Caption = "POST REPETIDOS: 0": Label3.Caption = "PAGINA: 0"
    Option1(0).Value = True: Option1(1).Enabled = True: Option1(0).Enabled = True: Frame(0).Enabled = True
    Frame(1).Enabled = True: Combo1.Enabled = True: Cate.Enabled = True: Command1.Caption = "Atacar a TARINGA CARAJOS"
    Command1.Enabled = True: Command2.Visible = False: Command2.Enabled = True: Esperate = False: numerodenagina = 0
    Label3.Caption = "PAGINA: ": Text1.Text = "": Text2.Text = "": Text3.Text = "": Text4.Text = "": II = 0: existen = 0
    ReDim LinesdePost(0): Siguiente = False: Tagcheck.Enabled = True: Exit Sub
End If
                
                    MsgBox "No a especificado la cuanta de Poringa"
                    Form2.Show 1
                    GoTo veriuserP1
                Else
                
                    If Not login = True Then Call Inet2.Execute("http://www.poringa.net/login.php", "POST", "nick=" & PoringaU & "&pass=" & PoringaPas, "Content-Type: application/x-www-form-urlencoded")
                    Do While Inet2.StillExecuting
                        DoEvents
                    Loop

                
                End If
            
                Call Inet1.Execute("http://www.poringa.net" & Cate, "GET")
                Command2.Visible = True
            End If
                Command1.Enabled = False
                Command1.Caption = "PROCESANDO"
        
        Else
        
        
            If Combo1.ListIndex = 0 Then
                
veriuserT2:
                
checkcue = checkcue + 1
                
                If TaringaU = "" Or TaringaPas = "" Then
                
If checkcue = 2 Then
    Text5.Enabled = True: Label5.Enabled = True: Label4.Caption = "POST REPETIDOS: 0": Label3.Caption = "PAGINA: 0"
    Option1(0).Value = True: Option1(1).Enabled = True: Option1(0).Enabled = True: Frame(0).Enabled = True
    Frame(1).Enabled = True: Combo1.Enabled = True: Cate.Enabled = True: Command1.Caption = "Atacar a TARINGA CARAJOS"
    Command1.Enabled = True: Command2.Visible = False: Command2.Enabled = True: Esperate = False: numerodenagina = 0
    Label3.Caption = "PAGINA: ": Text1.Text = "": Text2.Text = "": Text3.Text = "": Text4.Text = "": II = 0: existen = 0
    ReDim LinesdePost(0): Siguiente = False: Tagcheck.Enabled = True: Exit Sub
End If
                
                    MsgBox "No a especificado la cuanta de Taringa"
                    Form2.Show 1
                    GoTo veriuserT2
                Else
                
                If Not login = True Then Call Inet2.Execute("http://www.taringa.net/login.php", "POST", "nick=" & TaringaU & "&pass=" & TaringaPas, "Content-Type: application/x-www-form-urlencoded")
                    Do While Inet2.StillExecuting
                        DoEvents
                    Loop
                End If
                
                Call Inet2.Execute("http://www.taringa.net" & Text6, "GET")
            Else
            
                
veriuserP2:
                
checkcue = checkcue + 1
                
                If PoringaU = "" Or PoringaPas = "" Then
                
If checkcue = 2 Then
    Text5.Enabled = True: Label5.Enabled = True: Label4.Caption = "POST REPETIDOS: 0": Label3.Caption = "PAGINA: 0"
    Option1(0).Value = True: Option1(1).Enabled = True: Option1(0).Enabled = True: Frame(0).Enabled = True
    Frame(1).Enabled = True: Combo1.Enabled = True: Cate.Enabled = True: Command1.Caption = "Atacar a TARINGA CARAJOS"
    Command1.Enabled = True: Command2.Visible = False: Command2.Enabled = True: Esperate = False: numerodenagina = 0
    Label3.Caption = "PAGINA: ": Text1.Text = "": Text2.Text = "": Text3.Text = "": Text4.Text = "": II = 0: existen = 0
    ReDim LinesdePost(0): Siguiente = False: Tagcheck.Enabled = True: Exit Sub
End If
                
                    MsgBox "No a especificado la cuanta de Poringa"
                    Form2.Show 1
                    GoTo veriuserP2
                Else
                
                
                If Not login = True Then Call Inet2.Execute("http://www.poringa.net/login.php", "POST", "nick=" & PoringaU & "&pass=" & PoringaPas, "Content-Type: application/x-www-form-urlencoded")
                    Do While Inet2.StillExecuting
                        DoEvents
                    Loop
                End If
                
                Call Inet2.Execute("http://www.poringa.net" & Text6, "GET")
            End If
                Command1.Enabled = False
                Command1.Caption = "PROCESANDO"
        
        
        End If
        
         
End Sub

Private Sub Sacar_encabezados_de_los_post()
Text2.Text = ""
Dim t() As String
Dim inicio As Boolean
t = Split(Text1.Text, vbCrLf)
For cuenta = 0 To UBound(t)
If Not t(cuenta) = "" Then

           
        If Not InStr(1, t(cuenta), "<!-- inicio posts -->", vbTextCompare) = 0 Or sacalospost = True Then
            sacalospost = True
            If Not InStr(1, t(cuenta), "<a href=" & Chr(34) & Cate, vbTextCompare) = 0 Then
                If InStr(1, t(cuenta), "<a href=" & Chr(34) & Cate & "pagina", vbTextCompare) = 0 Then
                    Text2.Text = Text2.Text & t(cuenta) & vbCrLf & t(cuenta + 1) & "<br>" & vbCrLf & vbCrLf
                Else
                    Exit For
                End If
            End If
        End If


End If
    
    

Next cuenta

Text2.Text = Replace(Text2.Text, "Ã¡", "á")
Text2.Text = Replace(Text2.Text, "Ã©", "é")
Text2.Text = Replace(Text2.Text, "Ã" & Chr(173), "í")
Text2.Text = Replace(Text2.Text, "Ã³", "ó")
Text2.Text = Replace(Text2.Text, "Ãº", "ú")
Text2.Text = Replace(Text2.Text, "Ã" & Chr(129), "Á")
Text2.Text = Replace(Text2.Text, "Ã‰", "É")
Text2.Text = Replace(Text2.Text, "Ã" & Chr(141), "Í")
Text2.Text = Replace(Text2.Text, "Ã“", "Ó")
Text2.Text = Replace(Text2.Text, "Ãš", "Ú")
Text2.Text = Replace(Text2.Text, "Ã±", "ñ")
Text2.Text = Replace(Text2.Text, "Ã‘", "Ñ")
Text2.Text = Replace(Text2.Text, "Ãº", "º")
Text2.Text = Replace(Text2.Text, "Ãª", "ª")
Text2.Text = Replace(Text2.Text, "Â¿", "¿")
Text2.Text = Replace(Text2.Text, "Â®", "®")
Text2.Text = Replace(Text2.Text, "Â¡", "¡")

Matriz_de_Titulos_de_los_post
    
End Sub


Private Sub Sacar_Tag_de_los_post()
Tag = ""
Dim t() As String
Dim inicio As Boolean
t = Split(Text3.Text, vbCrLf)
For cuenta = 0 To UBound(t)
If Not t(cuenta) = "" Then

           
        If Not InStr(1, t(cuenta), "<div class=" & Chr(34) & "tags-block" & Chr(34) & ">", vbTextCompare) = 0 Or sacalospost = True Then
            sacalospost = True
           
           If Not InStr(1, t(cuenta), "<a href=" & Chr(34), vbTextCompare) = 0 Then
           
                
            If InStr(1, t(cuenta), "<li>", vbTextCompare) = 0 Then
            
                Tag = Tag & Replace(t(cuenta), Chr(9), "")
                If InStr(1, t(cuenta + 1), "</div>", vbTextCompare) = 0 Then
                    Tag = Tag & vbCrLf
                End If
            Else
                Exit For
            End If
           End If
        End If
    End If


    

Next cuenta
Tag = Replace(Tag, "- <a ", "<a ")
Tag = Replace(Tag, " <a ", "<a ")
Tag = "<!-- INI TAG -->" & vbCrLf & Tag & vbCrLf & "<!-- FIN TAG -->"

    
End Sub

Private Sub Command2_Click()
    StopB = True
    Command2.Enabled = False
End Sub

Private Sub Form_Load()

Combo1.ListIndex = 0


'numerodenagina = 797
If Not Dir$(App.Path & "\Taringa POST\", vbDirectory) <> "" Then
    MkDir (App.Path & "\Taringa POST\")
End If

'Me.Caption = "Pagina de timos posts: 0"
    ReDim LinesdePost(0)
'    dd = 3253
End Sub

Private Sub Inet1_StateChanged(ByVal State As Integer)
    Select Case State
       
       Case 12
          
          Dim bDone As Boolean: bDone = False: Dim strData As String: strData = ""
          DoEvents: vtData = Inet1.GetChunk(1024, icString)
          Do While Not bDone
          DoEvents
          Text1.DataChanged = True
          strData = strData & vtData
          vtData = Inet1.GetChunk(1024, icString)
          DoEvents
             If Len(vtData) <= 0 Then
                bDone = True
              End If
          Loop
        
          Text1.Text = UTF8_Decode(Replace(strData, vbLf, vbCrLf))
          Sacar_encabezados_de_los_post
    End Select
End Sub

Private Sub Inet2_StateChanged(ByVal State As Integer)
On Error Resume Next
Select Case State
   
   Case 12
      
      Dim bDone As Boolean: bDone = False: Dim strData As String: strData = ""
      DoEvents: vtData = Inet2.GetChunk(1024, icString)
      Do While Not bDone
      DoEvents
      Text3.DataChanged = True
      strData = strData & vtData
      vtData = Inet2.GetChunk(1024, icString)
      DoEvents
         If Len(vtData) <= 0 Then
            bDone = True
          End If
      Loop
      
      
       Text3.Text = UTF8_Decode(Replace(strData, vbLf, vbCrLf))
        If Text3.Text = "1: Home" Then
            Text3.Text = ""
            login = True
            Exit Sub
        ElseIf Text3.Text = "0: Datos no v&aacute;lidos" Then
            MsgBox "Cuenta Mala Datos no válidos"
            Form2.Show 1
            StopB = True
            Command2.Enabled = False
       End If
       Sacar_post
End Select

End Sub


Private Sub Option1_Click(Index As Integer)

Select Case Index
    Case 0
        Frame(1).Visible = False
        Frame(0).Visible = True
        Text5.Enabled = True
        Label5.Enabled = True
        Text5.BackColor = &H80000005
    Case 1
        Frame(0).Visible = False
        Frame(1).Visible = True
        Text5.Enabled = False
        Text5.BackColor = &H8000000B
        Label5.Enabled = False
    End Select

End Sub

Private Sub Text6_GotFocus()
If Text6.Text = "/posts/downloads/1905629/lo-que-sea.html                          Esto es un Ejemplo" Then Text6.Text = ""

End Sub

Private Sub Text6_LostFocus()
If Text6.Text = "" Then Text6.Text = "/posts/downloads/1905629/lo-que-sea.html                          Esto es un Ejemplo"
End Sub

Private Sub Timer1_Timer()

If Siguiente = True And Inet2.StillExecuting = False Then
    Esperate = False
    numerodenagina = numerodenagina + 1
    Label3.Caption = "PAGINA: " & numerodenagina
    
    Text1.Text = "": Text2.Text = "": Text3.Text = "": Text4.Text = ""
    II = 0
    existen = 0
    ReDim LinesdePost(0)
    If Combo1.ListIndex = 0 Then
        Call Inet1.Execute("http://www.taringa.net" & Cate & "pagina" & numerodenagina & ".html", "GET")
    Else
        Call Inet1.Execute("http://www.poringa.net" & Cate & "pagina" & numerodenagina & ".html", "GET")
    End If
    Siguiente = False
End If

End Sub


Private Sub User_pass_buton_Click()
    Form2.Show 1
End Sub


