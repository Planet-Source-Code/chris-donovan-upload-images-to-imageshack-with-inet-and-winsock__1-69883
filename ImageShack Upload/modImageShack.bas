Attribute VB_Name = "modImageShack"
Option Explicit

Public Enum UpType
    m_Inet = 0
    m_Winsock = 1
End Enum

'// Prepare Body and Headers for image upload
'// Instead of making two functions for Inet and Winsock, both are combined in one function
'// Returns one item in a String Array for Winsock (header + body)
'// Returns two items in a String Array for Inet (header and body separately)
Public Function PrepareImageUpload(ByVal ImagePath As String, UploadType As UpType, _
                            Optional RandomName As Boolean = False) As String()
    
    Dim intFile     As Integer  '// Next available number for the Open statement
    Dim BodyLength  As Long     '// Length of the body
    Dim ImageData   As String   '// Image contents
    Dim FileName    As String   '// File name to send to the server
    Dim Boundary    As String   '// Body boundary
    Dim Body        As String   '// Body contents
    Dim Header      As String   '// Header contents
    Dim TempArray() As String   '// Temporary array to hold Header and Body
    
    On Error GoTo ErrHandler
    
    '// Get the image contents
    intFile = FreeFile
    Open ImagePath For Binary As #intFile
        ImageData = String(LOF(intFile), Chr(0))
        Get #intFile, , ImageData
    Close #intFile

    '// Use original image name or use a random name in online ImageShack url
    '// ImageShack always adds 3 extra numbers/letters to the image name
    '// Random image name (i.e: http://img248.imageshack.us/img248/2103/c2h8q4bk3m7.jpg)
    '// Original image name (i.e: http://img248.imageshack.us/img248/2103/Screenshot3m7.jpg)
    If RandomName Then
        FileName = RandomString(8) & FileExtensionFromPath(ImagePath) '// Random
    Else
        FileName = FileNameFromPath(ImagePath) '// Original
    End If
    
    Boundary = RandomString(32) '// Create Boundary

    '// Create Body contents
    Body = "--" & Boundary & vbCrLf
    Body = Body & "Content-Disposition: form-data; name=""fileupload""; filename=""" & FileName & """" & vbCrLf
    Body = Body & "Content-Type: multipart/form-data" & vbCrLf
    Body = Body & vbCrLf & ImageData
    Body = Body & vbCrLf & "--" & Boundary & "--"

    BodyLength = Len(Body)
    
    '// Create Header contents
    If UploadType = m_Winsock Then
      Header = "POST /? HTTP/1.0" & vbCrLf '// Only add this if uploading image with Winsock
    End If
    
    Header = Header & "Host: imageshack.us" & vbCrLf
    Header = Header & "Content-Type: multipart/form-data, boundary=" & Boundary & vbCrLf
    Header = Header & "Content-Length: " & BodyLength & vbCrLf & vbCrLf
    
    If UploadType = m_Winsock Then
      Header = Header & Body '// Only add this if uploading image with Winsock
    End If

    
    '// Winsock 'SendData' sends Header + Body in one piece
    '// Inet 'Execute' sends Header and Body separately
    If UploadType = m_Winsock Then
        ReDim TempArray(0) '// One piece
        TempArray(0) = Header '// If Winsock is used, then Header + Body are one piece
    Else
        ReDim TempArray(1) '// Two pieces
        TempArray(0) = Body '// If Inet is used, then we need Header and Body separately
        TempArray(1) = Header
    End If

    PrepareImageUpload = TempArray '// Copy TempArray to Public PrepareImageUpload array
    Erase TempArray

Exit Function
ErrHandler:
    MsgBox Err.Number & " - " & Err.Description
    Err.Clear
End Function

'// Extract all the image links from the html source code and return them in a String Array
Public Function GrabLinks(strHTML As String) As String()
    Dim TempArray(7)    As String
    Dim tmp             As String
    Dim pos1            As Long
    Dim pos2            As Long
    
    On Error GoTo ErrHandler
    
    '// If an image is small, then no 'Thumbnail' links are returned
    If InStr(strHTML, "Please use clickable thumbnail") Then
    
        '// Thumbnail for Websites
        pos1 = InStr(strHTML, "value=""&lt;a href=&quot;http://")
        If pos1 Then
          pos2 = InStr(strHTML, """ /> Thumbnail for Websites")
            If pos2 Then
                tmp = Mid$(strHTML, pos1 + 7, pos2 - (pos1 + 7))
                tmp = Replace(tmp, "&lt;", "<")
                tmp = Replace(tmp, "&quot;", Chr$(34))
                tmp = Replace(tmp, "&gt;", ">")
                TempArray(0) = tmp
            End If
        End If
    
        '// Thumbnail for forums (1)
        pos1 = InStr(pos2, strHTML, "value=""")
        If pos1 Then
          pos2 = InStr(pos1, strHTML, """ /> Thumbnail for forums (1)")
            If pos2 Then
                tmp = Mid$(strHTML, pos1 + 7, pos2 - (pos1 + 7))
                tmp = Replace(tmp, vbCrLf, vbNullString)
                TempArray(1) = tmp
            End If
        End If
    
        '// Thumbnail for forums (2)
        pos1 = InStr(pos2, strHTML, "value=""")
        If pos1 Then
          pos2 = InStr(pos1, strHTML, """ /> Thumbnail for forums (2)")
            If pos2 Then
                tmp = Mid$(strHTML, pos1 + 7, pos2 - (pos1 + 7))
                TempArray(2) = tmp
            End If
        End If
    
    Else
        TempArray(0) = "n.a"
        TempArray(1) = "n.a"
        TempArray(2) = "n.a"
    End If
    
    
    '// Search the position of the 'Hotlink' links
    pos2 = InStr(strHTML, "Include details")
    If pos2 Then
    
        '// Hotlink for forums (1)
        pos1 = InStr(pos2, strHTML, "width: 500px"" size=""70"" value=""")
        If pos1 Then
          pos2 = InStr(pos1, strHTML, """/>")
            If pos2 Then
                tmp = Mid$(strHTML, pos1 + 31, pos2 - (pos1 + 31))
                TempArray(3) = tmp
            End If
        End If
    
        '// Hotlink for forums (2)
        pos1 = InStr(pos2, strHTML, "width: 500px"" size=""70"" value=""")
        If pos1 Then
          pos2 = InStr(pos1, strHTML, """/>")
            If pos2 Then
                tmp = Mid$(strHTML, pos1 + 31, pos2 - (pos1 + 31))
                TempArray(4) = tmp
            End If
        End If
    
        '// Hotlink for Websites
        pos1 = InStr(pos2, strHTML, "width: 500px"" size=""70"" value=""")
        If pos1 Then
          pos2 = InStr(pos1, strHTML, """/>")
            If pos2 Then
                tmp = Mid$(strHTML, pos1 + 31, pos2 - (pos1 + 31))
                tmp = Replace(tmp, "&lt;", "<")
                tmp = Replace(tmp, "&quot;", Chr$(34))
                tmp = Replace(tmp, "&gt;", ">")
                TempArray(5) = tmp
            End If
        End If
    
        '// Show image to friends
        pos1 = InStr(pos2, strHTML, "<a href=""")
        If pos1 Then
          pos2 = InStr(pos1, strHTML, """><b>Show</b>")
            If pos2 Then
                tmp = Mid$(strHTML, pos1 + 9, pos2 - (pos1 + 9))
                TempArray(6) = tmp
            End If
        End If

        '// Direct link to image
        pos1 = InStr(1, strHTML, "background-color: #DDDDAA;"" size=""70"" value=""")
        If pos1 Then
          pos2 = InStr(pos1, strHTML, """/>")
            If pos2 Then
                tmp = Mid$(strHTML, pos1 + 45, pos2 - (pos1 + 45))
                TempArray(7) = tmp
            End If
        End If
    
    Else
        TempArray(3) = "n.a"
        TempArray(4) = "n.a"
        TempArray(5) = "n.a"
        TempArray(6) = "n.a"
        TempArray(7) = "n.a"
    End If
    
    GrabLinks = TempArray '// Copy TempArray to Public GrabLinks array
    Erase TempArray

Exit Function
ErrHandler:
    MsgBox Err.Number & " - " & Err.Description
    Err.Clear
End Function

'// Returns the file name from a full path
Private Function FileNameFromPath(strPath As String) As String
    FileNameFromPath = Right$(strPath, Len(strPath) - InStrRev(strPath, "\"))
End Function

'// Returns file extension (including dot) from a full path or file name
Public Function FileExtensionFromPath(strPath As String) As String
    FileExtensionFromPath = Right$(strPath, (Len(strPath) - InStrRev(strPath, ".")) + 1)
End Function

'// Random string for the boundary and random image name
Private Function RandomString(ByVal HowMany As Integer)
    Dim i       As Integer
    Dim btByte  As Byte
    
    Randomize
    For i = 1 To HowMany
        btByte = Int(Rnd() * 127)
        If (btByte >= Asc("0") And btByte <= Asc("9")) Or _
           (btByte >= Asc("A") And btByte <= Asc("Z")) Or _
           (btByte >= Asc("a") And btByte <= Asc("z")) Then
            RandomString = RandomString & Chr(btByte)
        Else
            i = i - 1
        End If
    Next i
End Function

