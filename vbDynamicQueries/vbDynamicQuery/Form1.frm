VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5700
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10155
   LinkTopic       =   "Form1"
   ScaleHeight     =   5700
   ScaleWidth      =   10155
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkOutdoorOnly 
      Caption         =   "Show only outdoor centers"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6960
      TabIndex        =   5
      Top             =   780
      Width           =   3015
   End
   Begin MSComctlLib.Slider sldrPopulation 
      Height          =   435
      Left            =   6780
      TabIndex        =   3
      Top             =   120
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   767
      _Version        =   393216
      LargeChange     =   5000
      Max             =   1000000
      SelStart        =   1000000
      TickFrequency   =   100000
      Value           =   1000000
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   5325
      Width           =   10155
      _ExtentX        =   17912
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   8678
            Object.ToolTipText     =   "The city the user clicked"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   8678
            Object.ToolTipText     =   "Information about the selected city"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picCanvas 
      Height          =   5295
      Left            =   0
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   5235
      ScaleWidth      =   6675
      TabIndex        =   0
      Top             =   0
      Width           =   6735
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Population"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8940
      TabIndex        =   4
      Top             =   180
      Width           =   1035
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3915
      Left            =   6900
      TabIndex        =   1
      Top             =   1380
      Width           =   3255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'A  simple dynamic query demonstration
'We have a class that represents cities and some info about them
'We then filter the view depending on whether slider values match
'the city attributes.

'Note that the map was loaded into picCanvas via its picture property
'at design time. Also, I found the coordinates of the city by creating a
'label, and putting to coordinates into it on a mouse click event. After,
'I removed the label and that call.

'We could be much more sophisticated than this. FOr example, instead of
'showing or not showing cities, we could mute the cities that don't match
'(perhaps by making them grey) and emphasise the cities that do (perhaps
'by making them red and larger). We could also pop up information about
'the cities directly above it, or even show a picture of that city.
'There are *LOTS* of things you can do with this simple idea. Be creative!

Option Explicit

Dim colCitys As New Collection  'Our list of all cities.

Private Sub Form_Load()
    
    Form1.Caption = "A Simple Dynamic Query Demonstration"
    Form1.ScaleMode = vbPixels
    
    'The canvas look and drawing behaviour
    picCanvas.AutoRedraw = True
    picCanvas.ScaleMode = vbPixels
    picCanvas.ForeColor = vbRed  'All cities will be drawn in blue
    
    'Make the cities. Note that the statistics are completely made up!
    MakeCity "Calgary", "Commercial Center of Alberta", 238, 284, 800000, True
    MakeCity "Edmonton", "Capitol City of Alberta", 254, 199, 500000, False
    MakeCity "Red Deer", "Rural city between Calgary and Edmonton", 248, 251, 100000, False
    MakeCity "Banff", "Tourist Mecca, but busy", 215, 273, 10000, True
    MakeCity "Canmore", "The best place in the world", 226, 275, 10000, True
    MakeCity "Jasper", "A low-key tourist center", 173, 213, 5000, True
    MakeCity "Lethbridge", "Southernmost city in Alberta", 273, 324, 30000, True
    MakeCity "Medicine Hat", "A great name for a city", 303, 310, 20000, False
    MakeCity "LoydMinster", "A rural city", 306, 214, 15000, False
 
    Label1.Caption = "Use the controls above to show only those cities matching the population and whether its an outdoor center. Left-click the city for city information."
    CanvasFilterAndDraw
End Sub

Private Sub MakeCity(Name As String, Info As String, X As Integer, Y As Integer, Population As Long, OutdoorCenter As Boolean)
    Dim iCity As New City
    Const CityWidth As Integer = 10 'The Extent of the square representing the Citys
    
    iCity.Name = Name
    iCity.Information = Info
    iCity.XLeft = X
    iCity.YTop = Y
    iCity.Extent = CityWidth
    iCity.Population = Population
    iCity.OutdoorCenter = OutdoorCenter
    Set iCity.Canvas = picCanvas
    colCitys.Add iCity
End Sub
'Clear the canvas, and redraw only the Citys in the collection that match the current filters
Private Sub CanvasFilterAndDraw()
    Dim iCity As City
    picCanvas.Cls
    For Each iCity In colCitys
        If iCity.Population <= sldrPopulation.Value Then
            If chkOutdoorOnly.Value = 0 Or (chkOutdoorOnly.Value = 1 And iCity.OutdoorCenter = True) Then
                iCity.Draw
            End If
        End If
    Next iCity
End Sub

' if we click on a City, put its name and info in the status bar
Private Sub picCanvas_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim iCity As City

    'Find out if there is a city under the mouse pointer. If there is,
    'show its name and info in the status bar on the bottom.
    'Otherwise clear the statusbar.
    If Button <> 1 Then Exit Sub
    StatusBar1.Panels(1).Text = ""
    StatusBar1.Panels(2).Text = ""
    For Each iCity In colCitys
        If iCity.Inside(X, Y) Then
            StatusBar1.Panels(1).Text = iCity.Name
            StatusBar1.Panels(2).Text = iCity.Information
            Exit Sub
        End If
    Next iCity
End Sub


'Update the map to show (or not show) outdoor centers
Private Sub chkOutdoorOnly_Click()
    CanvasFilterAndDraw
End Sub

'Update the map to show cities matching the requested population
Private Sub sldrPopulation_Scroll()
    CanvasFilterAndDraw
End Sub
