This is a Windows Scripting Host, Visual Basic Script intended to be run by cscript.exe.  CScript.exe is a command console runner for WSH scripts.  WScript.exe is a windowed runner for WSH scripts.

This script performs a rudimentary function, similar to JNLP, in that it checks for updates on a remote server and copies the folder locally if any remote files are newer than the local ones.

It is invoked like so:
cscript autoupdate.vbs ".\dest" "\\someserver\source"