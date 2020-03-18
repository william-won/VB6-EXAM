VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "City"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'This simple class defines a city, which gets drawn as a square.
'Various statistical attributes of that city are kept here, as
'well as how to draw it.

'Note that you can use a helper wizard to create a class:
'Select Menu Project -> Add Class Module -> VB Class Builder

'Properties:
'   Name - The name of the city
'   Description - a brief descriptive phrase describing the city
'   Population - the number of people in the city
'   OutdoorCenter - whether its a locus of outdoor activity.
'
'   XLeft   - the left side of the square
'   YTop    - the top side of the square
'   Extent  - the width of the square
'   Canvas  - the control where we will draw the square onto
'Methods
'   Draw    - Draw the square onto the Canvas
'   Inside  - Check if an X, Y point is inside the square
    
   
'local variable(s) to hold property value(s).
'This is a good coding style to follow when using classes in VB
Option Explicit

Private mvarXLeft As Long 'local copy
Private mvarYTop As Long 'local copy
Private mvarExtent As Long 'local copy
Private mvarCanvas As Control 'local copy

Private mvarName As String 'local copy
Private mvarInformation As String 'local copy
Private mvarDescription As Boolean 'local copy
Private mvarPopulation As Long 'local copy
Private mvarOutdoorCenter As Boolean 'local copy

'These are just locally defined variables that hold the bottom right X Y coordinates
Private XRight As Long
Private YBottom As Long

'''
''' Properties: These define the square location and size.
''' Note that each Let calls the Recalculate routine, which
''' calculates and locally stores the coordinates of the bottom right X Y coordinates
'''

'Statistical information about the city
Public Property Let Name(ByVal vData As String)
    mvarName = vData
End Property
Public Property Get Name() As String
    Name = mvarName
End Property

Public Property Let Information(ByVal vData As String)
    mvarInformation = vData
End Property
Public Property Get Information() As String
    Information = mvarInformation
End Property

Public Property Let Population(ByVal vData As Long)
    mvarPopulation = vData
End Property
Public Property Get Population() As Long
    Population = mvarPopulation
End Property

Public Property Let OutdoorCenter(ByVal vData As Boolean)
    mvarOutdoorCenter = vData
End Property
Public Property Get OutdoorCenter() As Boolean
    OutdoorCenter = mvarOutdoorCenter
End Property

'Drawing information about the city
Public Property Let XLeft(ByVal vData As Long)
    mvarXLeft = vData
    Recalculate
End Property
Public Property Get XLeft() As Long
    XLeft = mvarXLeft
End Property

Public Property Let YTop(ByVal vData As Long)
    mvarYTop = vData
    Recalculate
End Property
Public Property Get YTop() As Long
    YTop = mvarYTop
End Property

Public Property Let Extent(ByVal vData As Long)
    mvarExtent = vData
    Recalculate
End Property

Public Property Get Extent() As Long
    Extent = mvarExtent
End Property


'Note that a control is an object, so it uses Set instead of Let
Public Property Set Canvas(ByVal vData As Control)
    Set mvarCanvas = vData
End Property
Public Property Get Canvas() As Control
    If IsObject(mvarCanvas) Then
        Set Canvas = mvarCanvas
    Else
        Canvas = mvarCanvas
    End If
End Property

'''
''' METHODS
'''
Public Sub Draw()
    If mvarExtent = 0 Then Exit Sub
    mvarCanvas.Line (mvarXLeft, mvarYTop)-(XRight, YBottom), , BF
End Sub
Public Function Inside(ByVal X As Long, ByVal Y As Long) As Boolean
    If X < mvarXLeft Or X > XRight Or Y < mvarYTop Or Y > YBottom Then
        Inside = False
    Else
        Inside = True
    End If
End Function

'When the object is first created, this is called automatically.
'We use it to initialize everything to 0
'(in VB, everything is automatically set to 0 so this isn't really necessary)
Private Sub Class_Initialize()
    Me.XLeft = 0
    Me.YTop = 0
    Me.Extent = 0
End Sub

'''
''' Other Private Procedures used by the above
'''

' Recalculate the coordinates of the bottom right corner
Private Sub Recalculate()
    XRight = mvarXLeft + mvarExtent
    YBottom = mvarYTop + mvarExtent
End Sub
