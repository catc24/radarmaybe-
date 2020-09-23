VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "NWS Ridge Radar Basic"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   270
      Top             =   2340
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   1830
      Left            =   180
      TabIndex        =   0
      Top             =   90
      Width           =   2040
      ExtentX         =   3598
      ExtentY         =   3228
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   0
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.Menu MnuSpecial 
      Caption         =   "Special Intrest Map's"
      Begin VB.Menu MnuUSRadLoop 
         Caption         =   "U.S. Radar Loop"
      End
      Begin VB.Menu MnuSerfaceAnalisis 
         Caption         =   "Serface Analisis"
      End
   End
   Begin VB.Menu MnuLocation 
      Caption         =   "Radar By Location"
      Begin VB.Menu MnuSelectLocation 
         Caption         =   "Select Location"
      End
      Begin VB.Menu MnuMapType 
         Caption         =   "Radar Map Type"
         Begin VB.Menu MnuBaseReflectivity 
            Caption         =   "Base Reflectivity"
         End
         Begin VB.Menu MnuBaseVelocity 
            Caption         =   "Base Velocity"
         End
         Begin VB.Menu MnuCompositeReflectivity 
            Caption         =   "Composite Reflectivity"
         End
         Begin VB.Menu MnuStormRelativeMotion 
            Caption         =   "Storm Relative Motion"
         End
         Begin VB.Menu MnuOneHourPrecipitation 
            Caption         =   "One-Hour Precipitation"
         End
         Begin VB.Menu MnuStormTotalPrecipitation 
            Caption         =   "Storm Total Precipitation"
         End
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public MapType As String
Public NwsOffice As String
Private RefreshTimer As Integer
Private Sub Form_Load()
WebBrowser1.Navigate "about:blank"
MapType = "N0R"
MnuBaseReflectivity.Checked = True
LoadOptionalMap "http://radar.weather.gov/Conus/Loop/NatLoop_Small.gif"
End Sub
Public Function LoadOptionalMap(dMapURL As String)
RefreshTimer = 0
Call WebBrowser1.Document.Script.Document.Clear
Call WebBrowser1.Document.Script.Document.write("<html><head></head><body><center><p>")
Call WebBrowser1.Document.Script.Document.write("<img src='" & dMapURL & "'>")
Call WebBrowser1.Document.Script.Document.write("</center></body></html>")
Call WebBrowser1.Refresh
Call WebBrowser1.Refresh
End Function
Public Sub LoadMap()
RefreshTimer = 0
Call WebBrowser1.Document.Script.Document.Clear
Call WebBrowser1.Document.Script.Document.write("<html><head></head><body><center><p>")
Call WebBrowser1.Document.Script.Document.write("<img src='http://www.srh.noaa.gov/ridge/lite/" & MapType & "/" & NwsOffice & "_loop.gif'>")
Call WebBrowser1.Document.Script.Document.write("</center></body></html>")
Call WebBrowser1.Refresh
Call WebBrowser1.Refresh
End Sub
Private Sub Form_Resize()
WebBrowser1.Top = Me.ScaleTop
WebBrowser1.Left = Me.ScaleLeft
WebBrowser1.Height = Me.ScaleHeight
WebBrowser1.Width = Me.ScaleWidth
End Sub
Private Sub UncheckAllMapTypes()
MnuBaseReflectivity.Checked = False
MnuStormRelativeMotion.Checked = False
MnuBaseVelocity.Checked = False
MnuOneHourPrecipitation.Checked = False
MnuCompositeReflectivity.Checked = False
MnuStormTotalPrecipitation.Checked = False
End Sub
Private Sub MnuBaseReflectivity_Click()
UncheckAllMapTypes
MnuBaseReflectivity.Checked = True
MapType = "N0R"
If NwsOffice <> "" Then LoadMap
End Sub
Private Sub MnuBaseVelocity_Click()
UncheckAllMapTypes
MnuBaseVelocity.Checked = True
MapType = "N0V"
If NwsOffice <> "" Then LoadMap
End Sub
Private Sub MnuCompositeReflectivity_Click()
UncheckAllMapTypes
MnuCompositeReflectivity.Checked = True
MapType = "NCR"
If NwsOffice <> "" Then LoadMap
End Sub
Private Sub MnuOneHourPrecipitation_Click()
UncheckAllMapTypes
MnuOneHourPrecipitation.Checked = True
MapType = "N1P"
If NwsOffice <> "" Then LoadMap
End Sub
Private Sub MnuSelectLocation_Click()
WebBrowser1.Navigate2 App.Path & "\RadLoc.htm"
End Sub
Private Sub MnuSerfaceAnalisis_Click()
LoadOptionalMap "http://adds.aviationweather.gov/data/progs/hpc_sfc_analysis.gif"
End Sub
Private Sub MnuStormRelativeMotion_Click()
UncheckAllMapTypes
MnuStormRelativeMotion.Checked = True
MapType = "N0S"
If NwsOffice <> "" Then LoadMap
End Sub
Private Sub MnuStormTotalPrecipitation_Click()
UncheckAllMapTypes
MnuStormTotalPrecipitation.Checked = True
MapType = "NTP"
If NwsOffice <> "" Then LoadMap
End Sub
Private Sub MnuUSRadLoop_Click()
LoadOptionalMap "http://radar.weather.gov/Conus/Loop/NatLoop_Small.gif"
End Sub
Private Sub Timer1_Timer()
RefreshTimer = RefreshTimer + 1
If RefreshTimer = 256 Then
RefreshTimer = 0
If MapType <> "" And NwsOffice <> "" Then
WebBrowser1.Refresh
WebBrowser1.Refresh
End If
End If
End Sub
Private Sub WebBrowser1_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
If LCase(URL) = LCase(App.Path & "\RadLoc.htm") Or URL = "about:blank" Then
Cancel = False
Else
Cancel = True
If MapType <> "" Then
NwsOffice = Right(URL, 3)
LoadMap
End If
End If
End Sub

