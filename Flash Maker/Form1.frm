VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash.ocx"
Begin VB.Form Form1 
   Caption         =   "Ricardo Arias SWF CREATOR - DONT FORGET TO VOTE !!!"
   ClientHeight    =   6975
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7605
   LinkTopic       =   "Form1"
   ScaleHeight     =   6975
   ScaleWidth      =   7605
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   375
      Left            =   240
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   315
      ScaleWidth      =   1290
      TabIndex        =   3
      Top             =   4920
      Width           =   1350
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   3855
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   7095
      _cx             =   12515
      _cy             =   6800
      FlashVars       =   ""
      Movie           =   ""
      Src             =   ""
      WMode           =   "Window"
      Play            =   -1  'True
      Loop            =   -1  'True
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   -1  'True
      Base            =   ""
      AllowScriptAccess=   "always"
      Scale           =   "ShowAll"
      DeviceFont      =   0   'False
      EmbedMovie      =   0   'False
      BGColor         =   ""
      SWRemote        =   ""
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Text            =   "Dont forget to vote for my code!!!!!!!!"
      Top             =   6000
      Width           =   6255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Make SWF"
      Height          =   375
      Left            =   5640
      TabIndex        =   0
      Top             =   6600
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "See the Image that is falling down!!!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   240
      TabIndex        =   6
      Top             =   120
      Width           =   6855
   End
   Begin VB.Label Label2 
      Caption         =   "Write ANY text here and it will anymated with an awesome effect in a Flash Movie!!"
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
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   5520
      Width           =   7215
   End
   Begin VB.Label Label1 
      Caption         =   $"Form1.frx":1596
      Height          =   615
      Left            =   1920
      TabIndex        =   4
      Top             =   4800
      Width           =   4815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'the concept is easy:
'you need to create each 'object' (text, image, blocks) and then add it to EVERY frame of the swf
'In case you want to change this object (rotation, traslation) need to do it frame by frame
'see the code and you will see
'
'A very important thing:
'You can add something to the swf, lets say in the first 50 frames
'Then you can add other object to the first 20
'Every time you write something to the frames its independient of what you already add
'
'Sorry if the code are a little mess, but i was just experimenting, when i start i DONT KNOW how to acomplish nothing
'then i was discovering everything
'Thats why maybe its a little mess : )
'
'ricardoarias@yahoo.com
'
'DONT FORGET TO VOTE
'This is one of the best codes in PSC!!!! Im sure.
'
'I hope that some of you help me to control better the text effects...
'----------------------------------------------------------------------------

Private Sub Command1_Click()

    Dim mv As Object, obj As Object, txt As Object, font As Object
    Dim btn As Object, act As Object, txt1(100) As Object, font1(100) As Object
    Dim Fname As String
    Fname = App.Path + "\sample.swf" ' The file that will be created
    Set Movie = CreateObject("swfobjs.swfMovie")
    Set obj = CreateObject("swfobjs.swfObject")
    Set obj1 = CreateObject("swfobjs.swfObject")
    Set btn = CreateObject("swfobjs.swfObject")
    Set act = CreateObject("swfobjs.swfAction")

    ' Set moive attribute
    With Movie
        .SetSize 6000 * 1.3, 4000 * 1.3
        .SetFrameBkColor 255, 255, 255
        .SetFrameRate 20
    End With
'Need to find a way that the outfile has the desire size
'Flash1.Width = 6000 * 1.3
'Flash1.Height = 4000 * 1.3

Set pic = CreateObject("swfobjs.swfObject")
pic.MakePicture 0, 0, Picture1.Width, Picture1.Height, Picture1.Width / Screen.TwipsPerPixelX, Picture1.Height / Screen.TwipsPerPixelY, App.Path + "\btndecode.JPG"
'top, left, width,height (in twips inside the swf), the 2 last in pixels
'Los 2 primeros son top y left, los 2 siguientes son widht y height en twips dentro del swf y los 2 ultimos en pixeles el rectangulo que tomara del bmp o jpg
Movie.AddObject pic ' Add picture

'Making the blocks
With obj
        .MakePolygon 500, 500
        .AddStraightLine 0, 3000
        .AddStraightLine 3000, 0
        .AddStraightLine 0, -3000
        .AddStraightLine -3000, 0
        .SetSolidFill 128, 0, 128, 70
End With
obj.SetDepth 1
Movie.AddObject obj

With obj1
        .MakePolygon 3000, 500
        .AddStraightLine 0, 2500
        .AddStraightLine 2500, 0
        .AddStraightLine 0, -2500
        .AddStraightLine -2500, 0
        .SetSolidFill 18, 250, 18, 20
End With

Movie.AddObject obj1

For i = 0 To 320
Movie.GotoFrame i
Movie.RemoveObject pic
Movie.RemoveObject obj
'Here we make the image falling down, a very nice effect it fall very far away!!!
pic.Translate i * 20, i * 20
pic.scaleEx (65536 * 10) / (i + 1), (65536 * 10) / (i + 1)
pic.Rotate 65536 * (i * 10)
obj.Rotate -65536 * (i * 1)
Movie.AddObject pic
Movie.AddObject obj
Next i

For i = 0 To 320
Movie.GotoFrame i
Movie.RemoveObject obj1
obj1.Rotate 65536 * (i * 1)
Movie.AddObject obj1
Next i

    Set font = CreateObject("swfobjs.swfObject")
    With font
        .MakeFont "Arial"
        .AddGlyph "Arial", "Flash", AscW("Hp")
        .AddGlyph "Arial", " Example", AscW("i")
        .AddGlyph "Arial", "http://bukoo.sourceforge.net", AscW("A")
    End With
    
    Set txt = CreateObject("swfobjs.swfObject")
    With txt
        .MakeTextEx "Hpi", font, 1270, 870, 1000
        .SetSolidFill 128, 128, 128, 100
    End With
    Movie.GotoFrame 0
    Movie.AddObject txt
    
    With txt
        .MakeTextEx "Hpi", font, 1200, 800, 1000
        .SetSolidFill 0, 0, 255, 255
    End With
    Movie.GotoFrame 0
    Movie.AddObject txt
    
For i = 0 To 100
    'On Error Resume Next
Movie.GotoFrame i
Movie.RemoveObject txt
If i * 5 < 255 Then
    txt.SetSolidFill 0, 0, 255, 255 - (i * 5)
    End If
Movie.AddObject txt
Next i

final:
'////////////////////////////////////
'---VAMOS A INTENTAR ANIMAR TEXTO
'We will try to animate some text!!!!!!!!!


Dim Cuantas As Integer: Dim Contador As Integer: Dim Letra As String: Dim Donde As Long
Dim Transparencia As Long: Dim LastLetra As String: Dim Espacio As Integer
Dim LetterSize As Integer: Dim FactorEscala As Integer
LetterSize = 300
Cuantas = Len(Text1.Text)

For i = 1 To Cuantas
Set font1(i) = CreateObject("swfobjs.swfObject")
    With font1(i)
        .MakeFont "Arial"
    End With
Set txt1(i) = CreateObject("swfobjs.swfObject")
DoEvents
Letra = Mid(Text1.Text, i, 1) 'Pick one letter at time
font1(i).AddGlyph "Arial", Letra, AscW(Letra) 'Add letter to the object
With txt1(i)
        'Trying to mantain similar distance betwen different letters.. not all use the same space
        If LastLetra = "l" Or LastLetra = "i" Or LastLetra = "j" Then
            Espacio = Espacio + (LetterSize * 0.21)
        ElseIf LastLetra = "f" Or LastLetra = "r" Then
            'estas letras tienen menos espaciado
            Espacio = Espacio + (LetterSize * 0.3)
        ElseIf LastLetra = "t" Then
            Espacio = Espacio + (LetterSize * 0.25)
        ElseIf LastLetra = "s" Then
            Espacio = Espacio + (LetterSize * 0.45)
        ElseIf LastLetra = "o" Or LastLetra = "p" Or LastLetra = "q" Or LastLetra = "n" Then
            Espacio = Espacio + (LetterSize * 0.5)
        ElseIf LastLetra = "w" Then
            Espacio = Espacio + (LetterSize * 0.65)
        ElseIf LastLetra = "m" Then
            Espacio = Espacio + (LetterSize * 0.75)
        Else
            'espaciado Normal
            'Normal space
            Espacio = Espacio + (LetterSize * 0.49)
        End If
        LastLetra = Letra
        .MakeTextEx Letra, font1(i), Espacio, 2500, LetterSize
        .SetSolidFill 255, 128, 128, 1
End With
Movie.GotoFrame Donde
Movie.AddObject txt1(i)
Transparencia = 1
FactorEscala = 155 '75 '155'310
'ADD THE TEXT TO THE FRAME WITH FX
'Wowwwwwwww here we go !!!!!!!!
For ii = Donde + 1 To Donde + 30
Transparencia = Transparencia + 7
Movie.GotoFrame ii
Movie.RemoveObject txt1(i)
txt1(i).scaleEx (65535 * 10) / FactorEscala + 1, (65535 * 10) / FactorEscala + 1
FactorEscala = FactorEscala - 10
txt1(i).SetSolidFill Transparencia, 0, Transparencia, Transparencia
txt1(i).Rotate -65536 * (ii * 50)
Movie.AddObject txt1(i)
Next ii
Movie.RemoveObject txt1(i)
txt1(i).scaleEx (LetterSize * 35.565) * 2, (LetterSize * 35.565) * 2
txt1(i).Rotate 0
Movie.AddObject txt1(i)
Donde = ii - 20
Next i


'---FIN ANIMACION DE TEXTO
' END OF TEXT ANIMATION... WE DONE IT !!!! HURRA !!!
    
'//////////////////////////////////////////////////////////////// CREA LA PELICULA
'CREATE THE SWF
 Movie.WriteMovie Fname

    Set Movie = Nothing
    Set obj = Nothing
    Set txt = Nothing
    Set font = Nothing
    Set btn = Nothing
    Set act = Nothing

    ' A Shockwave Flash OCX is needed for next 2 lines
    Flash1.Movie = Fname
    Flash1.Play

'THE END

End Sub


