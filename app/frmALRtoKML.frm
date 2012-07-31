VERSION 5.00
Begin VB.Form frmALRtoKML 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ALR to KML convertor"
   ClientHeight    =   5280
   ClientLeft      =   12165
   ClientTop       =   4455
   ClientWidth     =   3750
   Icon            =   "frmALRtoKML.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmALRtoKML.frx":030A
   ScaleHeight     =   5280
   ScaleWidth      =   3750
   Begin VB.Label lblMail 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   1110
      TabIndex        =   3
      Top             =   4980
      Width           =   1275
   End
   Begin VB.Label lblNotify 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "[No file loaded]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   735
      Left            =   0
      TabIndex        =   2
      Top             =   3300
      Width           =   3735
   End
   Begin VB.Label lblCONVERTBUTTON 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000FF&
      Height          =   855
      Left            =   360
      TabIndex        =   1
      Top             =   4080
      Width           =   3075
   End
   Begin VB.Label lblDummy 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "[No file loaded]"
      ForeColor       =   &H80000004&
      Height          =   315
      Left            =   4200
      TabIndex        =   0
      Top             =   240
      Width           =   1215
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmALRtoKML"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub Form_Load()
    EnableDragDrop Me.hwnd
End Sub

Public Sub GotADrop(ByVal strfile As String)
    lblDummy.Caption = strfile
    lblNotify.ForeColor = &H80FF&
    lblNotify.Caption = "File Loaded"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    DisableDragDrop Me.hwnd
End Sub

Private Sub lblMail_Click()
ShellExecute Me.hwnd, "", "mailto:vandael.marc@gmail.com", "", "", 1
End Sub


Private Sub lblCONVERTBUTTON_Click()
On Error GoTo ProcError
Dim sALRpath As String
Dim iFnum As Integer
Dim sALRwholefile As String
Dim lines As Variant
Dim one_line As Variant
Dim num_rows As Long
Dim num_cols As Long
Dim ALRtable() As String
Dim Row As Long
Dim Column As Long

    sALRpath = lblDummy.Caption
    
    iFnum = FreeFile
    Open sALRpath For Input As iFnum
    sALRwholefile = Input$(LOF(iFnum), #iFnum)
    Close iFnum

    lines = Split(sALRwholefile, vbCrLf)

    num_rows = UBound(lines)
    num_rows = num_rows - 1
    one_line = Split(lines(11), ",")
    num_cols = 2
    ReDim ALRtable(num_rows, num_cols)

    For Row = 11 To num_rows
        If Len(lines(Row)) > 0 Then
            one_line = Split(lines(Row), ",")
            For Column = 1 To num_cols
                ALRtable(Row, Column) = one_line(Column)
            Next Column
        End If
    Next Row
    

    Dim sFileText As String
    Dim iFileNo As Integer
    iFileNo = FreeFile
    
    Dim sKMLpath As String
    sKMLpath = lblDummy.Caption
    sKMLpath = Left(sKMLpath, Len(sKMLpath) - 3)
    sKMLpath = sKMLpath & "kml"
    
    
    'Open sKMLpath For Output As #iFileNo
    '    Print #iFileNo, "<kml><Document><Placemark><LineString><coordinates>"
  '
   '     For Row = 11 To num_rows
    '        Print #iFileNo, ALRtable(Row, 2) & "," & ALRtable(Row, 1) & ",0" & vbNewLine;
     '   Next Row
      '
       ' Print #iFileNo, "</coordinates></LineString></Placemark></Document></kml>"
'
 '   Close #iFileNo
  
  
    lblNotify.ForeColor = &HFF00&
    lblNotify.Caption = "File converted OK"
  
  
    'Dim KML As String
    'KML = "<?xml version=""1.0"" encoding=""UTF-8""?>" & ControlChars.Newline
    'KML &= "<kml xmlns=""http://earth.google.com/kml/2.0"">"

    

    'Bestandsnaam uit volledig pad halen
        Dim sKMLpathSplit() As String
        Dim sKMLfilename As String
        Dim iKMLfilenameArrayLength As Integer

        sKMLpathSplit = Split(sKMLpath, "\")
        iKMLfilenameArrayLength = UBound(sKMLpathSplit)
        sKMLfilename = IIf(iKMLfilenameArrayLength = 0, "", sKMLpathSplit(iKMLfilenameArrayLength))
    '-----------------------------------
  

      Open sKMLpath For Output As #iFileNo2
        Print #iFileNo, "<kml><Document><Placemark><LineString><coordinates>"
  
        For Row = 11 To num_rows
            Print #iFileNo, ALRtable(Row, 2) & "," & ALRtable(Row, 1) & ",0" & vbNewLine;
        Next Row
        
        Print #iFileNo, "</coordinates></LineString></Placemark></Document></kml>"

    Close #iFileNo
  
  
  
  
  
  
  
    Exit Sub
  
ProcError:
    lblNotify.ForeColor = &HFF&
    lblNotify.Caption = "Error!"
    Exit Sub
End Sub

