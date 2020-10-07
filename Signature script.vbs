'Option Explicit
On Error Resume Next

   Dim strSigName, strLogo
   Dim strFullName, strPhoneNumber, strTitle 
   Dim strFax, strTel, strCompany, strCorpEmail
   Dim strStreetAddress, strCityAddress, strSiteAddress, strSecondarySiteAddress
   Dim boolUpdateStyle

'==========================================================================
' Some script variables
'==========================================================================

'  Name signature
   strSigName  = "SignatureName"
'  If signature exists, overwrite (true) or leave alone (false)?
   boolUpdateStyle = true

'==========================================================================
' Set some static information
'==========================================================================

'  Company information
   strFullName = "Full name" ' Full name
   strTitle    = "title" ' Title
   strCompany  = "CompanyName" ' Company name
   strStreetAddress = "Street name and number" ' Company street address
   strCityAddress = "City zip code and state" ' Company city address
   strTel      = "Phone number" ' Company telphone number
   strPhoneNumber = "Cell number" ' Company cellphone number
   strFax      = "Fax number" ' Company fax number
   strSiteAddress = "https://www.example.com" ' Company site address
   strSecondarySiteAddress = "" ' Secondary company address

'  Fallback email address when no address is found
   strCorpEmail = "conatct@domain.com"

'  Logo name to be saved in %appdata%\Microsoft\Signatures\
   strLogo = "logo.jpg" ' Must apply format


'==========================================================================
' Get Signature Folder
'==========================================================================
   Dim objShell
   Set objShell = CreateObject("WScript.Shell")
   strSigFolder = ObjShell.ExpandEnvironmentStrings("%appdata%") & "\Microsoft\Signatures\"
   Set objShell = Nothing

   
'==========================================================================
' Get Logo from web
'==========================================================================
   dim xHttpLogo: Set xHttpLogo = createobject("Microsoft.XMLHTTP")
   dim bStrmLogo: Set bStrmLogo = createobject("Adodb.Stream")
   xHttpLogo.Open "GET", "https://www.example.com/logo.jpg", False ' Change the address
   xHttpLogo.Send

   with bStrmLogo
      .type = 1 '//binary
      .open
      .write xHttpLogo.responseBody
      .savetofile strSigFolder+"\"+strLogo, 2
   end with

'==========================================================================
' Get Signature Folder
'==========================================================================
   Dim objFSO, objFile
   Set objFSO   = CreateObject("Scripting.FileSystemObject")

   If Not (objFSO.FolderExists(strSigFolder)) Then
      Call objFSO.CreateFolder(strSigFolder)
   End If

   strHTMFile = strSigFolder & strSigName & ".htm"
   strRTFFile = strSigFolder & strSigName & ".rtf"
   strTXTFile = strSigFolder & strSigName & ".txt"


'==========================================================================
' Create HTM File
'==========================================================================


   Err.Clear
   Set objFile = objFSO.CreateTextFile(strHTMFile, boolUpdateStyle, False)

   If Err.Number = 0 Then
      objFile.Write "<!DOCTYPE html><html lang='en'><head> <meta charset='UTF-8'> <meta name='viewport' content='width=device-width, initial-scale=1.0'></head>"&vbCrLf
      objFile.Write "<body> <table width='100%' style='font-size: 8.0pt;'> <tr style='font-weight: bold;'> <td style='font-family: Verdana;'>" & strFullName & "</td></tr><tr style='font-weight: bold;'> <td style='font-family: Verdana;'>" & strTitle & "</tr><tr style='font-weight: bold;'><td style='font-family: Verdana;'>" & strCompany & "</tr>"&vbCrLf
      objFile.Write "<tr> <td style='font-family: Verdana;'>&nbsp;&nbsp;" & strStreetAddress & ",</td></tr><tr> <td style='font-family: Verdana;'>&nbsp;&nbsp;" & strCityAddress & "</td></tr><tr> <td style='font-family: Verdana;'>&nbsp;&nbsp;Tel " & strTel & "</td></tr><tr> <td style='font-family: Verdana;'>&nbsp;&nbsp;Cell " & strPhoneNumber & "</td></tr><tr> <td style='font-family: Verdana;'>&nbsp;&nbsp;Fax " & strFax & "</td></tr><tr> <td style='font-family: Verdana;'><span>&nbsp;&nbsp;&#9993;</span>&nbsp;<a href='mailto:" & strCorpEmail & "'>" & strCorpEmail & "</a></td></tr>"&vbCrLf
      objFile.Write "<tr> <td style='font-family: Verdana;'>&nbsp;</td></tr><tr> <td style='font-family: Verdana;'><a href=" & strSiteAddress & ">" & strSiteAddress & "</a></td></tr><tr> <td style='font-family: Verdana;'><a href=" & strSecondarySiteAddress & ">" & strSecondarySiteAddress & "</a></td></tr><tr> <td style='font-family: Verdana;'>&nbsp;</td></tr><tr> <td style='font-family: Verdana;'><img src=" &  strSigFolder + strLogo & "></td></tr></table></body></html>"&vbCrLf
      objFile.close
   End If


'==========================================================================
' Create TXT File
'==========================================================================
   Err.Clear
   Set objFile = objFSO.CreateTextFile(strTXTFile, boolUpdateStyle, False)
   If Err.Number = 0 Then
      objFile.Write "Kind Regards" & vbCrLf
      objFile.Write strFullName & vbCrLf
      objFile.Write "Tel " & strPhoneNumber & " | Fax " & strFax & " | Cell " & strPhoneNumber &"" & vbCrLf
      objFile.close
   End If


'==========================================================================
' Create RTF File
'==========================================================================
   Err.Clear
   Set objFile = objFSO.CreateTextFile(strRTFFile, boolUpdateStyle, False)
   If Err.Number = 0 Then
      objfile.write "{\rtf1\ansi\ansicpg1252\deff0\deflang1033{\fonttbl{\f0\fswiss\fprq2\fcharset0 Calibri;}{\f1\froman\fprq2\fcharset2 Webdings;}}" & vbCrLF
      objfile.write "{\colortbl;\red031\green073\blue125;\red0\green0\blue255;\red0\green128\blue0;}" & vbCrLF
      objfile.write "{\*\generator Msftedit 5.41.15.1507;}\viewkind4\uc1\pard\sb100\sa100\cf1\lang2057\f0\fs20 " & strFullName & "\line " & vbCrLF
      objfile.write strTitle & "\line " & strCompany & "\line T: " & strTel & "   F: " & strFax & "\line E: " & vbCrLF
      objfile.write "{\field{\*\fldinst{HYPERLINK ""mailto:" & strCorpEmail & """}}{\fldrslt{\ul " & strCorpEmail & "}}}\ulnone\f0\fs20    " & vbCrLF
      objfile.write "{\field{\*\fldinst{HYPERLINK """ & strSiteAddress & """}}{\fldrslt{\ul " & strSiteAddress & "}}}\ulnone \f0\fs20\par" & vbCrLF
      objfile.write "{\field{\*\fldinst{HYPERLINK """ & strSecondarySiteAddress & """}}{\fldrslt{\ul " & strSecondarySiteAddress & "}}}\ulnone\f0\fs20\par" & vbCrLF
      objfile.write "\pard\cf1\lang1033\par" & vbCrLF
      objfile.write "}" & vbCrLF
      objFile.close
   End If


'==========================================================================
' Tidy-up
'==========================================================================
   set objFile = Nothing
   set objFSO  = Nothing

