﻿Imports System.IO

Public Class frmALRtoKML
    Dim strooien As Boolean
    Dim endtagset As Boolean
    Dim wstrooien As Boolean
    Dim wendtagset As Boolean
    Dim Instruction As Boolean


    Private Sub frmALRtoKML_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        'Enable dropping of files on this form
        Me.AllowDrop = True

        'Add handlers
        AddHandler Me.DragDrop, AddressOf Form_DragDrop
        AddHandler Me.DragEnter, AddressOf Form_DragEnter

        strooien = True
        endtagset = False
        wstrooien = True
        wendtagset = False
        Instruction = False
    End Sub

    Private Sub Form_DragEnter(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DragEventArgs)

        'Display different mouse pointer if file is dragged
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.All
        End If

    End Sub

    Private Sub Form_DragDrop(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DragEventArgs)
        Dim strDroppedFiles As String()
        Dim rtbLength2 As Long
        Dim rtbMessage2 As String
        strDroppedFiles = DirectCast(e.Data.GetData(DataFormats.FileDrop), String())
        For Each strDroppedFile As String In strDroppedFiles



            Dim num_rows As Long
            Dim num_cols As Long
            Dim x As Integer
            Dim y As Integer
            Dim strarray(1, 1) As String




            ' Load the file.
            Dim tmpstream As StreamReader = File.OpenText(strDroppedFile)
            Dim strlines() As String
            Dim strline() As String

            Dim strDroppedFileName As String

            strDroppedFileName = Path.GetFileName(strDroppedFile)

            'Load content of file to strLines array
            strlines = tmpstream.ReadToEnd().Split(Environment.NewLine)

            ' Redimension the array.
            num_rows = UBound(strlines)
            num_rows = num_rows - 1

            Try
                strline = strlines(12).Split(",")
            Catch ex As Exception
                Dim rtbLength As Long
                Dim rtbMessage As String
                rtbMessage = vbNewLine & strDroppedFileName & vbTab & vbTab & "INVALID FILE"
                rtbLength = RichTextBox1.TextLength
                RichTextBox1.AppendText(rtbMessage)
                RichTextBox1.SelectionStart = rtbLength
                RichTextBox1.SelectionLength = rtbMessage.Length
                RichTextBox1.SelectionColor = Color.DarkOrange
                GoTo PrematureLoopEnd
            End Try

            num_cols = UBound(strline)



            ReDim strarray(num_rows, num_cols)
            Dim startx As Integer


            Try
                For x = 12 To num_rows
                    strline = strlines(x).Split(","c, ":"c)
                    For y = 0 To num_cols
                        strarray(x, y) = strline(y)
                    Next
                Next
                startx = 12
            Catch ex As IndexOutOfRangeException
                For x = 12 To num_rows
                    strline = strlines(x).Split(","c, ":"c)
                    For y = 0 To num_cols
                        strarray(x, y) = strline(y)
                    Next
                Next
                startx = 13
            End Try


            Dim newfilepath As String
            newfilepath = Microsoft.VisualBasic.Left(strDroppedFile, Len(strDroppedFile) - 4)
            newfilepath = newfilepath & ".kml"

            Dim newfilename As String
            newfilename = Path.GetFileName(newfilepath)

            Using writer As StreamWriter = New StreamWriter(newfilepath)
                writer.WriteLine("<?xml version=""1.0"" encoding=""UTF-8""?>")
                writer.WriteLine("<kml xmlns=""http://www.opengis.net/kml/2.2"">")
                writer.WriteLine("<Folder>")
                writer.WriteLine("<name>" & newfilename & "</name><open>0</open>")
                writer.WriteLine("<description>")
                writer.WriteLine("Generated by AL3-to-KML beta. Coded by Vandael Marc")
                writer.WriteLine("</description>")

                writer.WriteLine(vbTab & "<Document>")
                writer.WriteLine(vbTab & "<name>" & newfilename & " - Traject</name>")
                writer.WriteLine(vbTab & "<Style><ListStyle><listItemType>checkHideChildren</listItemType><bgColor>00ffffff</bgColor><maxSnippetLines>2</maxSnippetLines></ListStyle></Style>")
                writer.WriteLine(vbTab & "<Style id=""VolledigTraject""><LineStyle><width>1</width><color>FFDC78F0</color></LineStyle></Style>")

                writer.WriteLine(vbTab & vbTab & "<Placemark>")
                writer.WriteLine(vbTab & vbTab & vbTab & "<styleUrl>#VolledigTraject</styleUrl>")
                writer.WriteLine(vbTab & vbTab & vbTab & "<LineString>")
                writer.WriteLine(vbTab & vbTab & vbTab & "<gx:drawOrder>5</gx:drawOrder>")
                writer.WriteLine(vbTab & vbTab & vbTab & vbTab & "<coordinates>")

                For x = startx To num_rows
                    'If strarray(x, 5) = "0" Then
                    'strooien = False
                    'endtagset = False
                    'End If

                    'System.Diagnostics.Debug.Print(strarray(x, 0).ToString)
                    'MsgBox(strarray(x, 0))

                    If strarray(x, 0) = vbLf & "Instructions" Then


                        writer.WriteLine(vbTab & vbTab & vbTab & vbTab & "</coordinates>")
                        writer.WriteLine(vbTab & vbTab & vbTab & "</LineString>")
                        writer.WriteLine(vbTab & vbTab & "</Placemark>")
                        writer.WriteLine(vbTab & "</Document>")
                        writer.WriteLine("</Folder>")
                        writer.WriteLine("</kml>")

                        rtbMessage2 = vbNewLine & strDroppedFileName & vbTab & vbTab & "converted ok"
                        rtbLength2 = RichTextBox1.TextLength
                        RichTextBox1.AppendText(rtbMessage2)
                        RichTextBox1.SelectionStart = rtbLength2
                        RichTextBox1.SelectionLength = rtbMessage2.Length
                        RichTextBox1.SelectionColor = Color.DarkGray

                        Exit Sub
                    End If


                    'If Instruction = False Then
                    writer.WriteLine(vbTab & vbTab & vbTab & vbTab & vbTab & strarray(x, 1) & "," & strarray(x, 2) & ",0")
                    'End If

                    'If strarray(x, 5) = "1" Then
                    '    strooien = True
                    '    If endtagset = False Then
                    '        writer.WriteLine(vbTab & vbTab & vbTab & vbTab & "</coordinates>")
                    '        writer.WriteLine(vbTab & vbTab & vbTab & "</LineString>")
                    '        writer.WriteLine(vbTab & vbTab & "</Placemark>")
                    '        writer.WriteLine(vbTab & vbTab & "<Placemark>")
                    '        writer.WriteLine(vbTab & vbTab & vbTab & "<styleUrl>#NietStrooien</styleUrl>")
                    '        writer.WriteLine(vbTab & vbTab & vbTab & "<LineString>")
                    '        writer.WriteLine(vbTab & vbTab & vbTab & "<gx:drawOrder>5</gx:drawOrder>")
                    '        writer.WriteLine(vbTab & vbTab & vbTab & vbTab & "<coordinates>")

                    '        endtagset = True
                    '    End If
                    'End If
                Next



                writer.WriteLine(vbTab & vbTab & vbTab & vbTab & "</coordinates>")
                writer.WriteLine(vbTab & vbTab & vbTab & "</LineString>")
                writer.WriteLine(vbTab & vbTab & "</Placemark>")
                writer.WriteLine(vbTab & "</Document>")


                'writer.WriteLine(vbTab & "<Document>")
                'writer.WriteLine(vbTab & "<name>" & newfilename & " - Strooien</name>")
                'writer.WriteLine(vbTab & "<Style><ListStyle><listItemType>checkHideChildren</listItemType><bgColor>00ffffff</bgColor><maxSnippetLines>2</maxSnippetLines></ListStyle></Style>")
                'writer.WriteLine(vbTab & "<Style id=""Strooien""><LineStyle><width>6</width><color>FF14F000</color></LineStyle></Style>")

                'writer.WriteLine(vbTab & vbTab & "<Placemark>")
                'writer.WriteLine(vbTab & vbTab & vbTab & "<styleUrl>#Strooien</styleUrl>")
                'writer.WriteLine(vbTab & vbTab & vbTab & "<LineString>")
                'writer.WriteLine(vbTab & vbTab & vbTab & "<gx:drawOrder>5</gx:drawOrder>")
                'writer.WriteLine(vbTab & vbTab & vbTab & vbTab & "<coordinates>")

                'For x = startx To num_rows

                '    If strarray(x, 5) = "1" Then
                '        wstrooien = True
                '    End If

                '    If wstrooien = True Then
                '        writer.WriteLine(vbTab & vbTab & vbTab & vbTab & vbTab & strarray(x, 2) & "," & strarray(x, 1) & ",0")
                '    End If

                '    If strarray(x, 5) = "0" Then
                '        wstrooien = False

                '        writer.WriteLine(vbTab & vbTab & vbTab & vbTab & "</coordinates>")
                '        writer.WriteLine(vbTab & vbTab & vbTab & "</LineString>")
                '        writer.WriteLine(vbTab & vbTab & "</Placemark>")
                '        writer.WriteLine(vbTab & vbTab & "<Placemark>")
                '        writer.WriteLine(vbTab & vbTab & vbTab & "<styleUrl>#Strooien</styleUrl>")
                '        writer.WriteLine(vbTab & vbTab & vbTab & "<LineString>")
                '        writer.WriteLine(vbTab & vbTab & vbTab & "<gx:drawOrder>5</gx:drawOrder>")
                '        writer.WriteLine(vbTab & vbTab & vbTab & vbTab & "<coordinates>")



                '    End If

                'Next



                'writer.WriteLine(vbTab & vbTab & vbTab & vbTab & "</coordinates>")
                'writer.WriteLine(vbTab & vbTab & vbTab & "</LineString>")
                'writer.WriteLine(vbTab & vbTab & "</Placemark>")
                'writer.WriteLine(vbTab & "</Document>")


                'writer.WriteLine(vbTab & "<Document>")
                'writer.WriteLine(vbTab & "<name>" & newfilename & " - Traject</name>")
                'writer.WriteLine(vbTab & "<visibility>0</visibility>")
                'writer.WriteLine(vbTab & "<Style><ListStyle><listItemType>checkHideChildren</listItemType><bgColor>00ffffff</bgColor><maxSnippetLines>2</maxSnippetLines></ListStyle></Style>")
                'writer.WriteLine(vbTab & "<Style id=""VolledigTraject""><LineStyle><width>1</width><color>FFDC78F0</color></LineStyle></Style>")

                'writer.WriteLine(vbTab & vbTab & "<Placemark>")
                'writer.WriteLine(vbTab & vbTab & vbTab & "<styleUrl>#VolledigTraject</styleUrl>")
                'writer.WriteLine(vbTab & vbTab & vbTab & "<LineString>")
                'writer.WriteLine(vbTab & vbTab & vbTab & "<gx:drawOrder>0</gx:drawOrder>")
                'writer.WriteLine(vbTab & vbTab & vbTab & vbTab & "<coordinates>")

                'For x = startx To num_rows
                '    writer.WriteLine(vbTab & vbTab & vbTab & vbTab & vbTab & strarray(x, 2) & "," & strarray(x, 1) & ",0")
                'Next

                'writer.WriteLine(vbTab & vbTab & vbTab & vbTab & "</coordinates>")
                'writer.WriteLine(vbTab & vbTab & vbTab & "</LineString>")
                'writer.WriteLine(vbTab & vbTab & "</Placemark>")
                'writer.WriteLine(vbTab & "</Document>")
                writer.WriteLine("</Folder>")
                writer.WriteLine("</kml>")

            End Using


            rtbMessage2 = vbNewLine & strDroppedFileName & vbTab & vbTab & "converted ok"
            rtbLength2 = RichTextBox1.TextLength
            RichTextBox1.AppendText(rtbMessage2)
            RichTextBox1.SelectionStart = rtbLength2
            RichTextBox1.SelectionLength = rtbMessage2.Length
            RichTextBox1.SelectionColor = Color.DarkGray


PrematureLoopEnd:
        Next strDroppedFile

    End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub
End Class
