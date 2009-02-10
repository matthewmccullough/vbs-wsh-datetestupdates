Option Explicit 

Rem Enable Error Handling
'On Error Resume Next

Dim fso
Dim shell
Dim clArgs, clArgsNum
Rem Dim fileLocal, fileRemote

Dim dateOrig, dateRemote
Dim diffMinutes
Dim sysPropFolderRemote
Dim folderOrig, folderOrigTemp, folderLocalTemp, folderRemote

'Set up Global Objects
Set shell = WScript.CreateObject( "WScript.Shell" )
Set fso=CreateObject("Scripting.FileSystemObject")

'Get Command Line Arguments
Set clArgs = WScript.Arguments
clArgsNum = clArgs.Count
Wscript.StdOut.WriteLine "clArgsNum = " + clArgsNum
If clArgsNum <> 2 Then
   WScript.Echo "Improper number of command line args detected! Usage: autoupdate.vbs <.\app\> <remoteFolderUNCPath>"
   WScript.Quit 1
End If
'Load Command line args into vars
folderOrig = clArgs.Item(0)
folderOrigTemp = folderOrig + ".orig"
folderLocalTemp = folderOrig + ".new"
folderRemote = clArgs.Item(1)

Rem Alternate way to get params from environment vars
Rem http://cwashington.netreach.net/depo/view.asp?Index=665
'sysPropFolderRemote = shell.ExpandEnvironmentStrings("%C_ONE_APP_UPDATE_PATH%")
Rem Test if system property is not null.  If not copy it into folderRemote, otherwise default it
Rem IsNull -> http://www.w3schools.com/VBscript/vbscript_ref_functions.asp


dateOrig=fso.getFolder(folderOrig).DateLastModified
dateRemote=fso.getFolder(folderRemote).DateLastModified
If Err.Number <> 0 Then
  	WScript.Echo "Error while retriving " + folderRemote
End If


Rem Should we test each file for an update, or just one known file?

diffMinutes=CLng(DateDiff("n", dateOrig, dateRemote))

wscript.echo diffMinutes,"Differing Minutes"

If diffMinutes > 0 Then
  WScript.Echo "Update Needed. Copying Files."
  'Copy remote folder to local lib.update
  'http://msdn.microsoft.com/en-us/library/xbfwysex(VS.85).aspx
  fso.CopyFolder folderRemote, folderLocalTemp
  If Err.Number <> 0 Then
  	WScript.Echo "Error while copying " + folderRemote + " to " + folderLocalTemp
  End If
  
  'if copy success, then rename local lib folder to lib.prev
  fso.MoveFolder folderOrig, folderOrigTemp
  'rename lib.update to lib
  fso.MoveFolder folderLocalTemp, folderOrig
  'start app
Else
  WScript.Echo "Up To Date. Skipping Copying Files."
End If

WScript.Quit




Rem http://blog.netnerds.net/2007/01/vbscript-download-and-save-a-binary-file/

'Download File From HTTP - Approach 1
Function DownloadFile(DownloadUrl)
  'generic file downloader, saves to temp
  'Get name of file from url (whatever follows the final forwardslash "/")
  Dim arURL, FileName, FileSaveLocation
  arURL = Split(DownloadUrl,"/",-1,1)
  If arURL(UBound(arURL)) = "" Then 'if there is a trailing forwardslash
  FileName = arURL(UBound(arURL) -1)
  Else
  filename = arURL(UBound(arURL))
  End If
  'Get temp folder location
  Dim oFS, TempDir
  Set oFS = CreateObject("Scripting.FileSystemObject")
  Set TempDir = oFS.getSpecialFolder(2)
  Wscript.Echo TempDir & "\" & FileName
  FileSaveLocation = TempDir & "\" & FileName
  ' Fetch the file
  Dim oXMLHTTP, oADOStream
  Set oXMLHTTP = CreateObject("MSXML2.XMLHTTP")
  oXMLHTTP.open "GET", DownloadUrl, false
  oXMLHTTP.send()
  If oXMLHTTP.Status = 200 Then
  Set oADOStream = CreateObject("ADODB.Stream")
  oADOStream.Open
  oADOStream.Type = 1 'adTypeBinary
  oADOStream.Write oXMLHTTP.ResponseBody
  oADOStream.Position = 0 'Set the stream position to the start
  If oFS.Fileexists(FileSaveLocation) Then oFS.DeleteFile FileSaveLocation
  Set oFS = Nothing
  oADOStream.SaveToFile FileSaveLocation
  oADOStream.Close
  Set oADOStream = Nothing
  End if
  Set oXMLHTTP = Nothing
End Function

'Download File From HTTP - Approach 2
Function DownloadFile2()
  'Set your settings
  strFileURL = "http://www.domain.com/file.zip"
  strHDLocation = "D:\file.zip"

  ' Fetch the file
  Set objXMLHTTP = CreateObject("MSXML2.XMLHTTP")

  objXMLHTTP.open "GET", strFileURL, false
  objXMLHTTP.send()

  If objXMLHTTP.Status = 200 Then
    Set objADOStream = CreateObject("ADODB.Stream")
    objADOStream.Open
    objADOStream.Type = 1 'adTypeBinary

    objADOStream.Write objXMLHTTP.ResponseBody
    objADOStream.Position = 0    'Set the stream position to the start

    Set objFSO = Createobject("Scripting.FileSystemObject")
      If objFSO.Fileexists(strHDLocation) Then objFSO.DeleteFile strHDLocation
    Set objFSO = Nothing

    objADOStream.SaveToFile strHDLocation
    objADOStream.Close
    Set objADOStream = Nothing
  End If
  Set objXMLHTTP = Nothing
End Function

'Download File From HTTP - Approach 3
Function DownloadFile3()
  Dim ie
  Set ie=CreateObject("internetexplorer.application")
  ie.visible=false
  ie.navigateto("http://www.vbsedit.com/samples.gif")
  Do while ie.busy=true
  WScript.sleep 60
  Loop
  ie.Document.execwb "saveas", 2
  ie.quit
End Function