<div align="center">

## cAspImage


</div>

### Description

cAspImage is a vbscript class that lets you read various properties from image files, including width, height and color depth. This cannot be done directly with ASP/vbscript because this type of information has to be parsed out from the image file itself. Supported file formats are PNG, GIF, BMP and JPG.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Terje Hauger](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/terje-hauger.md)
**Level**          |Advanced
**User Rating**    |4.8 (24 globes from 5 users)
**Compatibility**  |ASP \(Active Server Pages\)
**Category**       |[Complete Applications](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/complete-applications__4-7.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/terje-hauger-caspimage__4-9209/archive/master.zip)





### Source Code

```
============================================================
STEP 1: SAVE THE FOLLOWING CODE AS cAspImage.asp:
============================================================
<%
'============================================================
' MODULE:  cAspImage.asp
' AUTHOR:  © www.u229.no
' CREATED: May 2005
'============================================================
' COMMENT:
' Read Image Properties from BMP, GIF, PNG and JPG files.
' Requirements: Microsoft Data Access Components installed on Web Server.
' PLEASE NOTE: Some JPEG files contain Thumbnails. In those cases this code will fail because
' it will think that the thumbnail's width/height are the "real" values.
' If this is a concern see more info on line 151.
'============================================================
' TODO:
'============================================================
' ROUTINES:
' - Private Sub Class_Initialize()
' - Private Sub Class_Terminate()
' - Public Function ReadImage(sFullPath)
' - Private Function ReadByteArray(sFullPath)
' - Private Sub EmptyVariables()
'============================================================
Class cAspImage
'// MODULE VARIABLES
Private m_arrBytes        '// Byte array holding the image file
Private m_lWidth          '// Width in pixels
Private m_lHeight         '// Height in pixels
Private m_iColorDepth      '// Color Depth (BitsPerPixel)
Private m_lImageSize      '// # Bytes in image
Private m_sDateCreated    '// Date Created
Private m_sLastModified     '// Date last saved
Private m_sImageType     '// PNG, JPG, GIF87a/GIF89a, BMP
Private m_sErrorMsg       '// Error message: Check this if ReadImage returns false
'// PROPERTIES
Public Property Get Width()
  Width = m_lWidth
End Property
Public Property Get Height()
  Height = m_lHeight
End Property
Public Property Get ColorDepth()
  ColorDepth = m_iColorDepth
End Property
Public Property Get ImageSize()
  ImageSize = m_lImageSize
End Property
Public Property Get DateCreated()
  DateCreated = m_sDateCreated
End Property
Public Property Get DateLastModified()
  DateLastModified = m_sLastModified
End Property
Public Property Get ImageType()
  ImageType = m_sImageType
End Property
Public Property Get ErrorMessage()
  ErrorMessage = m_sErrorMsg
End Property
'------------------------------------------------------------------------------------------------------------
' Comment: Init module variables.
'------------------------------------------------------------------------------------------------------------
Private Sub Class_Initialize()
  On Error Resume Next
  Call EmptyVariables
End Sub
'------------------------------------------------------------------------------------------------------------
' Comment: Clean up.
'------------------------------------------------------------------------------------------------------------
Private Sub Class_Terminate()
End Sub
'------------------------------------------------------------------------------------------------------------
' Comment: Main routine returning the image properties.
'------------------------------------------------------------------------------------------------------------
Public Function ReadImage(sFullPath)
'  On Error Resume Next
  Dim oFSO
  Dim oFile
  Dim i
  Dim bStop
	Dim lTmpHeight
	Dim lTmpWidth
	Dim iTmpDepth
  '// These 3 are created to speed up the looping.
  Dim i4
  Dim byteTmp
  Dim lSafeSize
  Call EmptyVariables
  bStop = False
  If IsEmpty(oFSO) Then Set oFSO = Server.CreateObject("Scripting.FileSystemObject")
  If oFSO.FileExists(sFullPath) Then
    Set oFile = oFSO.GetFile(sFullPath)
    m_lImageSize = oFile.Size
    m_sDateCreated = FormatDateTime(oFile.DateCreated, 2)
    m_sLastModified = FormatDateTime(oFile.DateLastModified, 2)
    If Not ReadByteArray(sFullPath) Then m_sErrorMsg = "Error Reading Image File"
'---------------------------- GIF
    If AscB(MidB(m_arrBytes, 1, 1)) = 71 And AscB(MidB(m_arrBytes, 2, 1)) = 73 And AscB(MidB( _
            m_arrBytes, 3, 1)) = 70 Then
      m_sImageType = "GIF89a"
      If AscB(MidB(m_arrBytes, 5, 1)) = 55 Then m_sImageType = "GIF87a"
      m_lWidth = CLng(AscB(MidB(m_arrBytes, 7, 1)) + (AscB(MidB(m_arrBytes, 8, 1)) * 256))
      m_lHeight = CLng(AscB(MidB(m_arrBytes, 9, 1)) + (AscB(MidB(m_arrBytes, 10, 1)) * 256))
      m_iColorDepth = 2 ^ ((Asc(CStr(AscB(MidB(m_arrBytes, 11, 1)))) And 7) + 1)
      bStop = True
    End If
'---------------------------- JPG
		If Not bStop Then
			If AscB(MidB(m_arrBytes, 1, 1)) = 255 And AscB(MidB(m_arrBytes, 2, 1)) = 216 And AscB(MidB( _
							m_arrBytes, 3, 1)) = 255 And AscB(MidB(m_arrBytes, 4, 1)) = 224 Then
				m_sImageType = "JPG"
				lSafeSize = (m_lImageSize - 1)
				For i = 5 To lSafeSize
					If AscB(MidB(m_arrBytes, i, 1)) = 255 Then
						byteTmp = AscB(MidB(m_arrBytes, i + 1, 1))
						If (byteTmp > 191) And (byteTmp < 208) Then
						  i4 = AscB(MidB(m_arrBytes, i + 4, 1))
'=============================================================================================
'// Some JPEG files contain Thumbnails. In those cases this code will fail because it will think that the thumbnail's width/height are the "real" values.
'// If you care about the "thumbnail problem" you may comment existing code/uncomment the other lines below.
'// Be aware that this will dramatically slow down the looping time because we then will have to loop through the whole file(s)
							m_lHeight = CLng(AscB(MidB(m_arrBytes, i + 6, 1)) + (AscB(MidB(m_arrBytes, i + 5, 1)) * 256))
							m_lWidth = CLng(AscB(MidB(m_arrBytes, i + 8, 1)) + (AscB(MidB(m_arrBytes, i + 7, 1)) * 256))
							m_iColorDepth = CInt(i4) * CInt(AscB(MidB(m_arrBytes, i + 9, 1)))
'							lTmpHeight = CLng(AscB(MidB(m_arrBytes, i + 6, 1)) + (AscB(MidB(m_arrBytes, i + 5, 1)) * 256))
'							lTmpWidth = CLng(AscB(MidB(m_arrBytes, i + 8, 1)) + (AscB(MidB(m_arrBytes, i + 7, 1)) * 256))
'							iTmpDepth = CInt(i4) * CInt(AscB(MidB(m_arrBytes, i + 9, 1)))
'
							If m_iColorDepth > 0 And (i4 > 1 And i4 < 17) Then
'							If iTmpDepth > 0 And (i4 > 1 And i4 < 17) Then
'								If (lTmpHeight > m_lHeight) Or (lTmpWidth > m_lWidth) Then
'									m_lHeight = lTmpHeight
'									m_lWidth = lTmpWidth
'									m_iColorDepth = iTmpDepth
                                    Exit For
'								End If
							End If
'=============================================================================================
						End If
					End If
				Next
				bStop = True
			End If
		End If
'---------------------------- PNG
    If Not bStop Then
      If AscB(MidB(m_arrBytes, 1, 1)) = 137 And AscB(MidB(m_arrBytes, 2, 1)) = 80 And AscB( _
              MidB(m_arrBytes, 3, 1)) = 78 And AscB(MidB(m_arrBytes, 4, 1)) = 71 _
              And AscB(MidB(m_arrBytes, 5, 1)) = 13 And AscB(MidB(m_arrBytes, 6, _
              1)) = 10 And AscB(MidB(m_arrBytes, 7, 1)) = 26 And AscB(MidB(m_arrBytes, 8, 1)) = 10 Then
        m_sImageType = "PNG"
        m_lWidth = CLng(AscB(MidB(m_arrBytes, 20, 1)) + (AscB(MidB(m_arrBytes, 19, 1)) * 256))
        m_lHeight = CLng(AscB(MidB(m_arrBytes, 24, 1)) + (AscB(MidB(m_arrBytes, 23, 1)) * 256))
        Select Case CInt(AscB(MidB(m_arrBytes, 26, 1)))                 '// Get Bit Depth
          Case 0
            m_iColorDepth = CInt(AscB(MidB(m_arrBytes, 25, 1)))         '// Grayscale
          Case 2
            m_iColorDepth = CInt(AscB(MidB(m_arrBytes, 25, 1))) * 3      '// RGB encoded
          Case 3
            m_iColorDepth = 8                                   '// Palette based, 8 bpp
            Case 4
            m_iColorDepth = CInt(AscB(MidB(m_arrBytes, 25, 1))) * 2      '// greyscale with alpha
          Case 6
            m_iColorDepth = CInt(AscB(MidB(m_arrBytes, 25, 1))) * 4      '// RGB encoded with alpha
          Case Else
        End Select
        bStop = True
      End If
    End If
'---------------------------- BMP
    If Not bStop Then
      If AscB(MidB(m_arrBytes, 1, 1)) = 66 And AscB(MidB(m_arrBytes, 2, 1)) = 77 Then
        m_sImageType = "BMP"
        m_lWidth = CLng(AscB(MidB(m_arrBytes, 19, 1)) + (AscB(MidB(m_arrBytes, 20, 1)) * 256))
        m_lHeight = CLng(AscB(MidB(m_arrBytes, 23, 1)) + (AscB(MidB(m_arrBytes, 24, 1)) * 256))
        m_iColorDepth = CInt(AscB(MidB(m_arrBytes, 29, 1)))
        bStop = True
      End If
    End If
'----------------------------
  Else
    m_sErrorMsg = "Error in File Path: " & sFullPath
  End If
  Set oFile = Nothing
  Set oFSO = Nothing
  ReadImage = (Err.Number = 0)
End Function
'------------------------------------------------------------------------------------------------------------
' Comment: Read image into byte array.
'------------------------------------------------------------------------------------------------------------
Private Function ReadByteArray(sFullPath)
  On Error Resume Next
  Dim oStream
  If IsEmpty(oStream) Then Set oStream = Server.CreateObject("ADODB.Stream")
  With oStream
    .Type = 1           '// adTypeBinary
    .Open
    .LoadFromFile sFullPath
    m_arrBytes = .Read
  End With
  oStream.Close
  Set oStream = Nothing
  ReadByteArray = (Err.Number = 0)
End Function
'------------------------------------------------------------------------------------------------------------
' Comment: Set module variables empty.
'------------------------------------------------------------------------------------------------------------
Private Sub EmptyVariables()
  On Error Resume Next
  m_lWidth = 0
  m_lHeight = 0
  m_iColorDepth = 0
  m_lImageSize = 0
  m_sDateCreated = ""
  m_sLastModified = ""
  m_sImageType = "Unknown"
  m_sErrorMsg = ""
End Sub
End Class
%>
===========================================
STEP 2: SAVE THE FOLLOWING AS start.asp IN THE SAME FOLDER AS ABOVE.
ALSO PUT A GIF FILE INTO THE SAME FOLDER AND NAME IT test.gif.
THEN POINT YOUR BROWSER TO start.asp.
===========================================
<% @Language="VBScript" %>
<%
Option Explicit
'On Error Resume Next
%>
<!--#include file="cAspImage.asp"-->
<%
'// HOW TO USE THIS CODE:
Dim oAspImg
Set oAspImg = New cAspImage
With oAspImg
  .ReadImage(Server.MapPath("test.gif"))
	Response.Write "ImageSize: " & .ImageSize & "<br />"
	Response.Write "Date Created: " & .DateCreated & "<br />"
	Response.Write "Date Last Modified: " & .DateLastModified & "<br />"
	Response.Write "ColorDepth: " & .ColorDepth & "<br />"
	Response.Write "Width: " & .Width & "<br />"
	Response.Write "Height: " & .Height & "<br />"
	Response.Write "ImageType: " & .ImageType & "<br />"
	Response.Write "Error Message: " & .ErrorMessage & "<br />"
End With
Set oAspImg = Nothing
%>
```

