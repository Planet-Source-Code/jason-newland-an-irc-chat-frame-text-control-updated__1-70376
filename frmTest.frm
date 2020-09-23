VERSION 5.00
Begin VB.Form frmTest 
   Caption         =   "IRC Output Text Display with UTF-8 Font Encoding"
   ClientHeight    =   3975
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9195
   LinkTopic       =   "Form1"
   ScaleHeight     =   3975
   ScaleWidth      =   9195
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pBackground 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   900
      Left            =   5880
      Picture         =   "frmTest.frx":0000
      ScaleHeight     =   60
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   80
      TabIndex        =   1
      Top             =   705
      Visible         =   0   'False
      Width           =   1200
   End
   Begin Project1.ucChatFrame pChatText 
      Height          =   3435
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5340
      _extentx        =   9419
      _extenty        =   6059
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'IRC CHAT BOX WINDOW (ala mIRC)
'©2007-2008 Jason James Newland, KangaSoft Software
'   UPDATE: 28 April, 2008 - Bug fixes, plus better organized demonstration code and the
'   inclusion of UTF-8 font encoding

Private Sub Form_Load()
    On Error Resume Next
    Dim strTemp As String
    'Set the font
    Me.pChatText.Font.Name = "Tahoma"
    Me.pChatText.Font.Size = 10
    'Set background tag for image alignment style - 1 = Centered, 2 = Fill/Stretch
    '   3 = Tiled & 4 = Photo (top right hand corner)
    Me.pBackground.Tag = "4"
    'Set the background image bound to a picture box (must be set to pixels
    Set Me.pChatText.BGSource = Me.pBackground
    Me.pChatText.Refresh
    'We would load our textfile on join before anything else
    Me.pChatText.LoadFile App.Path & "\xaim.log", 50
    'Format a UTF-8 line (normally we would either do this during add, or at the socket
    'end (speed) - we convert the wide character set to ANSI then reconvert it to wide
    'with the UTF8 code page (wont work in all fonts unless unicode is supported and
    'then it MUST have the correct script embedded in the font)
    OutText LineSep & vbCrLf & "* Attempting to rejoin #r00t" & vbCrLf & LineSep, 3
    strTemp = WToA("* Topic is '4 à¹‘Û©ÛžÛ©à¹‘ Beyond your reach.... à¹‘Û©ÛžÛ©à¹‘  '", CP_ACP)
    strTemp = AToW(strTemp, CP_UTF8)
    'Output some IRC based text
    OutText "* Now talking in #r00t", 4
    OutText strTemp, 4
    OutText "* Set by hdawgy_ on Mon Apr 21 13:00:02", 4
    OutText "<+wikid> O_O", 13
    OutText "<Jay> same", 13
    OutText "<Jay> im ps7 god", 13
    OutText "<+wikid> lol thats not good that the power went out", 13
    OutText "<+wikid> night steve", 13
    OutText "<Jay> i mean cs2", 13
    OutText "<Jay> :S", 13
    OutText "<Jay> night steve", 13
    OutText "<~xalixcatx[a]> 12slappin afr0-kl0wn to.... Green Day - Working Class Hero 12- (4:05/192kbps) 12", 13
    OutText "<+wikid> imma head to bed as well", 13
    OutText "<Jay> night wikid", 13
    OutText "<+sh0rt1e> i mean ive been doing since i was like 14, but i can do alot with like photoshop and html", 13
    OutText "* ~nitsirk 11· tanya tucker - what do i do with me 11· 2:5711/256kbps", 1
    OutText "<Jay> bah", 13
    OutText "<+Afr0-Kl0wn> 12«10Kmp312»4¨1008 I See It Now4¨12«1003:3712»", 13
    OutText "<hdawgy_> what about04 with complete li06nes of text that continue on to something a little longer than what most people would account for in the given instance and would generally consider annoying, not to mention that usin03g such long text 15is often disregarded 14as whole, for the lack of ability to09 consume large amount07s of bullshit", 13
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Me.pChatText.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub

'Text output
Private Sub OutText(sText As String, iCol As Integer, Optional IsTimeStamp As Boolean = True)
    On Error Resume Next
    Dim tmpArr() As Byte, s As Long, W As Long
    tmpArr = sText
    '
    Me.pChatText.AddLine tmpArr, iCol, IsTimeStamp
    '
    Erase tmpArr()
End Sub
