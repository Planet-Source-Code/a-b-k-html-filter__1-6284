VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8550
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11655
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8550
   ScaleWidth      =   11655
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox tf 
      Height          =   375
      Left            =   2160
      TabIndex        =   20
      Text            =   "Tag Type (i.e. head, or script)"
      Top             =   1680
      Width           =   1575
   End
   Begin VB.CommandButton Command17 
      Caption         =   "Convert 2 single line"
      Height          =   375
      Left            =   2160
      TabIndex        =   15
      ToolTipText     =   "converts the whole text into a single lined string"
      Top             =   3600
      Width           =   1575
   End
   Begin VB.CommandButton Command15 
      Caption         =   "Filter Spaces"
      Height          =   375
      Left            =   2160
      TabIndex        =   14
      ToolTipText     =   "Filter out spaces in the resulting text"
      Top             =   3120
      Width           =   1575
   End
   Begin VB.CommandButton Command14 
      Caption         =   "HTML 2 Text"
      Height          =   375
      Left            =   2160
      TabIndex        =   13
      ToolTipText     =   "Tries to filter out the HTML related part from the source code"
      Top             =   1200
      Width           =   1575
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Get page title"
      Height          =   375
      Left            =   2160
      TabIndex        =   12
      ToolTipText     =   "Speaks for itself huh?"
      Top             =   720
      Width           =   1575
   End
   Begin SHDocVwCtl.WebBrowser browser 
      Height          =   4215
      Left            =   120
      TabIndex        =   11
      Top             =   4200
      Width           =   11415
      ExtentX         =   20135
      ExtentY         =   7435
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Navigate"
      Default         =   -1  'True
      Height          =   255
      Left            =   5520
      TabIndex        =   10
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Find HTML"
      Height          =   375
      Left            =   240
      TabIndex        =   9
      ToolTipText     =   "Finds the next HTML Tag"
      Top             =   2160
      Width           =   1575
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Find Text"
      Height          =   375
      Left            =   240
      TabIndex        =   8
      ToolTipText     =   "Finds the text string typed in the below box in the source code"
      Top             =   720
      Width           =   1575
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Filter SHTML"
      Height          =   375
      Left            =   2160
      TabIndex        =   7
      ToolTipText     =   $"Form1.frx":324A
      Top             =   2160
      Width           =   1575
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Find SHTML"
      Height          =   375
      Left            =   240
      TabIndex        =   6
      ToolTipText     =   "Finds the HTML tag type as given below, from start to end"
      Top             =   1200
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Find between keys"
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   2640
      Width           =   1575
   End
   Begin VB.TextBox s2 
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Text            =   "Key 2"
      Top             =   3600
      Width           =   1575
   End
   Begin VB.TextBox s1 
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Text            =   "Key 1"
      Top             =   3120
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Filter HTML"
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      ToolTipText     =   "Filter out all HTML tags"
      Top             =   2640
      Width           =   1575
   End
   Begin RichTextLib.RichTextBox rt 
      Height          =   3495
      Left            =   3960
      TabIndex        =   1
      Top             =   600
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   6165
      _Version        =   393217
      HideSelection   =   0   'False
      ScrollBars      =   2
      TextRTF         =   $"Form1.frx":32E0
   End
   Begin VB.TextBox tURL 
      Height          =   285
      Left            =   840
      TabIndex        =   0
      Text            =   "www.weather.com/weather/us/zips/37235.html"
      Top             =   90
      Width           =   4575
   End
   Begin VB.Frame Frame1 
      Caption         =   "Search Tools"
      Height          =   3615
      Left            =   120
      TabIndex        =   17
      Top             =   480
      Width           =   1815
      Begin VB.TextBox ts 
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Text            =   "Text or Tag Type"
         Top             =   1200
         Width           =   1575
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Filtering Tools"
      Height          =   3615
      Left            =   2040
      TabIndex        =   18
      Top             =   480
      Width           =   1815
   End
   Begin InetCtlsObjects.Inet Inet 
      Left            =   6720
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "http://"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command10_Click()
    Dim url As String
    
    url = "http://" + tURL.Text
    rt.Text = Inet.OpenURL(url)
    Form1.Caption = "Filtering URL: " + tURL.Text
    browser.Navigate url
End Sub

Private Sub Command11_Click()
    tURL.Text = GetTitle(rt.Text)
End Sub
Private Sub Command14_Click()
    rt.Text = HTML2Text(rt.Text)
End Sub

Private Sub Command15_Click()
    rt.Text = Convert2Space(rt.Text)
End Sub

Private Sub Command17_Click()
    rt.Text = Convert2SingleLine(rt.Text)
End Sub

Private Sub Command2_Click()
    rt.Text = CleanHTMLTags(rt.Text)
End Sub


Private Sub Command4_Click()
    rt.Text = GetStringBetween(rt.Text, s1.Text, s2.Text)
End Sub

Private Sub Command6_Click()
    Dim s, l As Long
    FindSHTMLTag rt.Text, CStr(ts.Text), s, l
    If l > 0 Then
        rt.SelStart = s
        rt.SelLength = l
    End If
End Sub

Private Sub Command7_Click()
    rt.Text = CleanSHTMLTags(rt.Text, s2.Text)
End Sub

Private Sub Command8_Click()
    Dim s, l As Long
    
    l = Len(ts.Text)
    s = InStr(1, rt.Text, ts.Text, vbTextCompare)
    If s <= 0 Then
        Exit Sub
    Else
        rt.SelStart = s - 1
        rt.SelLength = l
    End If
End Sub

Private Sub Command9_Click()
    Dim s, l As Long
    s = 0
    l = 0
    FindHTMLTag rt.Text, s, l, CInt(rt.SelStart + rt.SelLength)
    tURL.Text = "s= " & s & " l= " & l
    rt.SelStart = s
    If l >= 0 Then
        rt.SelLength = l
    Else
        rt.SelLength = 0
    End If
End Sub

Private Sub Form_Load()
    Form1.Caption = "HTML File Filtering Tools"
    PrintMessage
End Sub

Private Sub PrintMessage()
    Dim str As String
    
    str = "This Program is developed to assist in filtering information from web pages." _
        + Chr(10) _
        + "For this purpose 'textfilter' module is written." _
        + Chr(10) _
        + "This is not the most perfect thing, but it may really help you to visualize how the information you want can be filtered from a web page." _
        + Chr(10) _
        + "Play around with the commands to get used to them. Once you get used to the functionality, then you can read the ' textfilter' module and start using the extra functions that are not reachable from this interface. " _
        + "By then, I hope you will be able to filter any information from any page..." _
        + Chr(10) _
        + "This box will hold the source code of the URL entered above and the actual web page will be displayed below..." _
        + Chr(10) _
        + "------------------------> A. B. K. <---------------------" _
        + Chr(10) _
        + "Please send your comments to" _
        + Chr(10) + Chr(10) + Chr(9) _
        + "cyber_dude@engineer.com. " _
        + Chr(10) + Chr(10) _
        + "For additional code submissions check out " _
        + Chr(10) + Chr(10) + Chr(9) _
        + "http://www.members.tripod.com/prog-dude/"

    
    rt.Text = str

End Sub
