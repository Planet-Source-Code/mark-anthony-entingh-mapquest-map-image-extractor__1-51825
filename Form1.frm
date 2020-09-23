VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6345
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9960
   LinkTopic       =   "Form1"
   ScaleHeight     =   6345
   ScaleWidth      =   9960
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Form1.frx":0000
      Left            =   60
      List            =   "Form1.frx":0002
      TabIndex        =   12
      Text            =   "8"
      Top             =   3480
      Width           =   855
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   60
      TabIndex        =   9
      Text            =   "US"
      Top             =   2820
      Width           =   1755
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   60
      TabIndex        =   7
      Text            =   "33901"
      Top             =   2220
      Width           =   1755
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   60
      TabIndex        =   5
      Text            =   "1415 Dean St"
      Top             =   1620
      Width           =   1755
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   60
      MaxLength       =   2
      TabIndex        =   3
      Text            =   "FL"
      Top             =   1020
      Width           =   1755
   End
   Begin Project1.InternetFile Inet1 
      Left            =   8100
      Top             =   5760
      _ExtentX        =   1005
      _ExtentY        =   1005
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Go"
      Height          =   435
      Left            =   8700
      TabIndex        =   1
      Top             =   5820
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   60
      TabIndex        =   0
      Text            =   "Fort Myers"
      Top             =   420
      Width           =   1755
   End
   Begin VB.Label Label6 
      Caption         =   "Zoom Level"
      Height          =   195
      Left            =   60
      TabIndex        =   11
      Top             =   3240
      Width           =   1035
   End
   Begin VB.Label Label5 
      Caption         =   "Country"
      Height          =   195
      Left            =   60
      TabIndex        =   10
      Top             =   2580
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "Zip Code"
      Height          =   195
      Left            =   60
      TabIndex        =   8
      Top             =   1980
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Address"
      Height          =   195
      Left            =   60
      TabIndex        =   6
      Top             =   1380
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "State"
      Height          =   195
      Left            =   60
      TabIndex        =   4
      Top             =   780
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "City"
      Height          =   195
      Left            =   60
      TabIndex        =   2
      Top             =   180
      Width           =   615
   End
   Begin VB.Image Image1 
      Height          =   1815
      Left            =   3600
      Top             =   60
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private DownloadStatus As Integer
Private HTMLSourceCode As String

Private Sub Command1_Click()
    If GetMap(Text1.Text, Text2.Text, Text3.Text, Text4.Text, Text5.Text, Combo1.Text, "c:\tmp.gif") = True Then Image1.Picture = LoadPicture("c:\tmp.gif")
End Sub


Function GetMap(city, state, address, zip, country, zoom, savepath) As Boolean
    Inet1.LocalFile = App.Path & "\tmp.htm"
    
    'generate url string
    
    url = "http://www.mapquest.com/maps/map.adp?city="
    url = url & city & "&state=" & state & "&address=" & address
    url = url & "&zip=" & zip & "&country=" & country & "&zoom=" & zoom
    
    url = Replace(url, " ", "%20")
    
    
    Inet1.url = url
    DownloadStatus = 0
    Inet1.StartDownload
    
    Do Until DownloadStatus > 0
        DoEvents
    Loop
    If DownloadStatus = 2 Then
        GetMap = False
        Exit Function
    End If
    
    HTMLSourceCode = ""
    f = FreeFile
    Open App.Path & "\tmp.htm" For Input As #f
        HTMLSourceCode = Input(LOF(f), f)
    Close #f
    
    
    
    xstart = 1
    xend = 1
    xstart = InStr(1, HTMLSourceCode, "mqmapgend?MQMapGenRequest=")
    If xstart > 0 Then
        For X = xstart To 1 Step -1
            If LCase(Mid(HTMLSourceCode, X, 7)) = "http://" Then
                xstart = X
                Exit For
            End If
        Next
        
        For X = xstart To Len(HTMLSourceCode)
            If Mid(HTMLSourceCode, X, 1) = "'" Or Mid(HTMLSourceCode, X, 1) = Chr(34) Or Mid(HTMLSourceCode, X, 1) = ">" Or Mid(HTMLSourceCode, X, 1) = "<" Or Mid(HTMLSourceCode, X, 1) = " " Then
                xend = X
                Exit For
            End If
        Next
    Else
        GetMap = False
    End If
    
    xlink = Mid(HTMLSourceCode, xstart, xend - xstart)
    
    
    '-----------------------------------
    'DOWNLOAD IMAGE
    '-----------------------------------
    Inet1.LocalFile = savepath
    Inet1.url = xlink
    DownloadStatus = 0
    Inet1.StartDownload
    
    Do Until DownloadStatus > 0
        DoEvents
    Loop
    If DownloadStatus = 2 Then
        GetMap = False
        Exit Function
    End If
    
    GetMap = True
End Function

Private Sub Form_Load()
    Combo1.AddItem "9"
    Combo1.AddItem "8"
    Combo1.AddItem "7"
    Combo1.AddItem "6"
    Combo1.AddItem "5"
    Combo1.AddItem "4"
    Combo1.AddItem "3"
    Combo1.AddItem "2"
    Combo1.AddItem "1"
    Combo1.AddItem "0"
End Sub

Private Sub Inet1_DownloadCancelled(lPosition As Long)
    DownloadStatus = 3
End Sub

Private Sub Inet1_DownloadComplete()
    DownloadStatus = 1
End Sub

Private Sub Inet1_DownloadError(sErrorDescription As String)
    DownloadStatus = 2
End Sub
