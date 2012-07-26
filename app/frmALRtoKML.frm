VERSION 5.00
Begin VB.Form frmALRtoKML 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ALR to KML convertor"
   ClientHeight    =   5280
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3750
   Icon            =   "frmALRtoKML.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmALRtoKML.frx":030A
   ScaleHeight     =   5280
   ScaleWidth      =   3750
   StartUpPosition =   3  'Windows Default
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
Dim filenamemail As String
ShellExecute Me.hwnd, "", "mailto:vandael.marc@gmail.com", "", "", 1
End Sub


Private Sub lblCONVERTBUTTON_Click()
On Error GoTo ProcError
Dim file_name As String
Dim fnum As Integer
Dim whole_file As String
Dim lines As Variant
Dim one_line As Variant
Dim num_rows As Long
Dim num_cols As Long
Dim the_array() As String
Dim R As Long
Dim C As Long

    file_name = lblDummy.Caption
    
    fnum = FreeFile
    Open file_name For Input As fnum
    whole_file = Input$(LOF(fnum), #fnum)
    Close fnum

    lines = Split(whole_file, vbCrLf)

    num_rows = UBound(lines)
    num_rows = num_rows - 1
    one_line = Split(lines(11), ",")
    num_cols = 2
    ReDim the_array(num_rows, num_cols)

    For R = 11 To num_rows
        If Len(lines(R)) > 0 Then
            one_line = Split(lines(R), ",")
            For C = 1 To num_cols
                the_array(R, C) = one_line(C)
            Next C
        End If
    Next R
    

    Dim sFileText As String
    Dim iFileNo As Integer
    iFileNo = FreeFile
    
    Dim newfilename As String
    newfilename = lblDummy.Caption
    newfilename = Left(newfilename, Len(newfilename) - 3)
    newfilename = newfilename & "kml"
    
    Open newfilename For Output As #iFileNo
        Print #iFileNo, "<kml>" & vbNewLine & "<Document>" & vbNewLine & "<Placemark>" & vbNewLine & "<LineString>" & vbNewLine & "<coordinates>"
  
        For R = 11 To num_rows
            Print #iFileNo, the_array(R, 2) & "," & the_array(R, 1) & ",0" & vbNewLine;
        Next R
        
        Print #iFileNo, "</coordinates>" & vbNewLine & "</LineString>" & vbNewLine & "</Placemark>" & vbNewLine & "</Document>" & vbNewLine & "</kml>"

    Close #iFileNo
  
  
    lblNotify.ForeColor = &HFF00&
    lblNotify.Caption = "File converted OK"
  
    Exit Sub
  
ProcError:
    lblNotify.ForeColor = &HFF&
    lblNotify.Caption = "Error!"
    Exit Sub
End Sub

