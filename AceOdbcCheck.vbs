Option Explicit

' Copyright 2017 Gordon D. Thompson
'
' Licensed under the Apache License, Version 2.0 (the "License");
' you may not use this file except in compliance with the License.
' You may obtain a copy of the License at
'
'     http://www.apache.org/licenses/LICENSE-2.0
'
' Unless required by applicable law or agreed to in writing, software
' distributed under the License is distributed on an "AS IS" BASIS,
' WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
' See the License for the specific language governing permissions and
' limitations under the License.

Const appTitle = "Microsoft Access ""ACE"" ODBC Status v1.0.0"
Const driverName = "Microsoft Access Driver (*.mdb, *.accdb)"

Dim bits(1)
bits(0) = "32-bit"
bits(1) = "64-bit"
Dim wShell
Set wShell = CreateObject("WScript.Shell")
Dim s
s = wShell.ExpandEnvironmentStrings("%SystemRoot%")
Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")
Dim bitnessIndex
If fso.FileExists(s & "\System32\msjet40.dll") Then
    bitnessIndex = 0
Else
    bitnessIndex = 1
End If

Dim dbFileSpec
dbFileSpec = CreateEmptyDb(fso)

Dim con
Set con = CreateObject("ADODB.Connection")
Dim conSuccess
On Error Resume Next
con.Open "Driver={" & driverName & "};Dbq=" & dbFileSpec
conSuccess = (Err.Number = 0)
On Error Goto 0
If conSuccess Then
    con.Close
    Set con = Nothing
End If

' clean up .mdb file and temp folder
Dim f
Set f = fso.GetFile(dbFileSpec)
Dim fldr
Set fldr = f.ParentFolder
f.Delete True
fldr.Delete True
Set f = Nothing
Set fldr = Nothing

If conSuccess Then
    MsgBox "Your " & bits(bitnessIndex) & " ODBC setup is working properly.", vbInformation + vbSystemModal, appTitle
Else
    Set con = Nothing
    s = "Connection failed. The driver" & vbCrLf & vbCrlf & _ 
            driverName & vbCrlf & vbCrLf & _
            "is not properly installed for " & bits(bitnessIndex) & " applications."
    If bitnessIndex = 0 Then
        MsgBox s, vbExclamation + vbSystemModal, appTitle
    Else
        If vbYes = MsgBox(s & vbCrlf & vbCrlf & "Would you like to try running this script in 32-bit mode?", vbExclamation + vbYesNo + vbSystemModal, appTitle) Then
            s = wShell.ExpandEnvironmentStrings("%SystemRoot%") & "\SysWOW64\WSCRIPT.EXE """ & WScript.ScriptFullName & """"
            wShell.Run s
            WScript.Quit
        End If
    End If
End If

Set fso = Nothing
Set wShell = Nothing


Private Function CreateEmptyDb(fso)
    Const TemporaryFolder = 2
    Const dbFileName = "empty.mdb"
    Dim zipFileSpec
    zipFileSpec = fso.GetSpecialFolder(TemporaryFolder) & "\" & fso.GetTempName & ".zip"
    Dim strm
    Set strm = CreateObject("ADODB.Stream")
    strm.Type = 1  ' adTypeBinary
    strm.Open
    strm.Write DecodeBase64(GetZipAsBase64())
    strm.SaveToFile zipFileSpec
    strm.Close
    Set strm = Nothing
    
    ' unzip
    ' ref - http://www.rondebruin.nl/win/s7/win002.htm
    Dim destination
    destination = fso.GetSpecialFolder(TemporaryFolder) & "\" & fso.GetTempName
    fso.CreateFolder destination
    Dim oApp
    Set oApp = CreateObject("Shell.Application")
    oApp.Namespace(destination).CopyHere oApp.Namespace(zipFileSpec).Items.Item(dbFileName)
    Set oApp = Nothing
    
    fso.DeleteFile zipFileSpec
    CreateEmptyDb = destination & "\" & dbFileName
End Function


Private Function DecodeBase64(strData)
    'ref - http://web.archive.org/web/20060527094535/http://www.nonhostile.com/howto-encode-decode-base64-vb6.asp
    Dim objXML
    Set objXML = CreateObject("MSXML2.DOMDocument")
    Dim objNode
    Set objNode = objXML.createElement("b64")
    objNode.DataType = "bin.base64"
    objNode.Text = strData
    DecodeBase64 = objNode.nodeTypedValue
   
    Set objNode = Nothing
    Set objXML = Nothing
End Function


Private Function GetZipAsBase64()
    Dim s
    s = ""
    s = s & "UEsDBBQAAAAIANE0VkruOcVQtAgAAAAwAQAJAAAAZW1wdHkubWRi7NxZ6ExRHAfw37n3zph9xlgT" & vbCrLf
    s = s & "kYQU2UJe7IrsS7JkHfsyYuyyK8sDJYVI8iAv4sELHmR9oEiIB0uRyPJEWWL8zjn3zDX3GsTYv58x" & vbCrLf
    s = s & "d35zf/ecc2fm/v/63XP/Q4JoeGHSvNykBbkm/acWmvTuKVfR8Xn25InRMyOfHJ6+oHs32rtq//bi" & vbCrLf
    s = s & "xj3tzi878bztqXuX9025/fDs/r7vtpwatH/cs869H+x9e+9Ny0exmlsmDuywotOzG/eOzRteq9fN" & vbCrLf
    s = s & "yLEOrS8sG7s1u7LVwfsrd0ca73p/pNERMebOhmkTdr6+urnZxukXGrfofe3FlanPizUu9j2wtG7i" & vbCrLf
    s = s & "+ZPWk6+PCBN1bNOWAAAAAAAAoDqEINz+45sQubRwjwWLTqYXpwkAAAAAAAAA/jiviz/CEnfjspcG" & vbCrLf
    s = s & "DtFonnOPmPNCvlGG16Q41SSLQ3mnMG8SloFpUh4nZF9RsomX3HO2tEFCD2IFEhGZcHhcnailVsZ4" & vbCrLf
    s = s & "U5UI+RJZTjgyUUN15SUacB8xmYhxUD7GByFbuLwWctS4TCS5M99e6USGO/tsIk3RzydSFPEn1F7J" & vbCrLf
    s = s & "nQ3sVVQmwqUX2L58d4Vvdx2zuwlu4x9DdRWnUGAMWybssjdRfQy8p70oT/P4NpWmUIFfa2+aRAW+" & vbCrLf
    s = s & "T+b7Ql5b210zlbdcwEsde2tH0nzKuXGM+tIcjqdzy7oc51WLmfx8Hg3i9XP5mUP9KMfLAbSYR5fL" & vbCrLf
    s = s & "PrSUWy/gfEY9H0h53mIRzeGtE2rNEM7meZyI7kWNNJiWyL3mTIbzk+RI/Lwge+exh/FWHPOaadxy" & vbCrLf
    s = s & "gHyNvBf1feuH0wy1jwXueQQto/ncx/oa8k0XFolikT77L6OOO22tWq5XT70GwX813Z8tr43+cN0P" & vbCrLf
    s = s & "WP4oqk/XUU9Lx2l5Ti7d96++/zWX3pkoOfIHRv+YyU9dHdoR+fOojuWUPHDVwZuQGXW0xuRho48p" & vbCrLf
    s = s & "+dr4zkMCAAAAAMA/xhKHU/JxsNBVR71K9b8szfluUlF+jJbX/I4ucEWgXBW+4rM0RMV6POoWR4GS" & vbCrLf
    s = s & "OEw9uAwdyOVPX1XEzVAF4ExdsKqCMUODOZoly1m3MArTcH7s/Q1FmsWvKOaOtYm+UogFR0LRBAAA" & vbCrLf
    s = s & "AAAAAH8qSzxNysekVV7silItruc6h0c4jHj1f5wf476uPj9dLfScf8h3MYAcy9Yz384nLeSGlju1" & vbCrLf
    s = s & "X5bQz2PuqYTABLdMcN769jn/aKX5+CyfYyjwTZ5XmEyL9HS3mqSeryZcF/JtpppQjpjp7sBEdqw0" & vbCrLf
    s = s & "Hdvuk7h98JyBmshewI9TeemeofjC7LPD73nCP43sCk4jN/WPFnxl8u3mPu04pnwBAAAAAAD+aZa4" & vbCrLf
    s = s & "lJCPre1A/e9VlLr+D3G9y1upOyV5kyR9k0rXBeiE+DRhUp+9Wt2U/yFflf86asr/z54X8M4kBBNh" & vbCrLf
    s = s & "30mJLnGdcHmJ+OkPIkVT+Jbnin8RV/PzuHKfzpX0ZHX1QYqXn+YyfI5gOfXyrymrx5upNcP42TS+" & vbCrLf
    s = s & "60u5p/Ay92m7yluV9dXQbKXORxT0+QleN4Oj+d6VD4+LxU7F4vxi0fbFFn+WKfnZzu/0eF3pknbH" & vbCrLf
    s = s & "a9NXnX+YzstSG5vbpANtwl8cx+E2GdPmS5fBm5z3Fyde7huuzPj29/rb30WcHwEAAAAAgL+aEC9j" & vbCrLf
    s = s & "Js6q7/9rl36WOpSanGqUupXckRyUjCbPJVYnuiZexY/GZ8ebxwkAAAAA4L8g6MetIQAANa9ZBT/W" & vbCrLf
    s = s & "iSPyyYo9NCHqLn/tZdQf+Tv82+tHh4NfbPU2uRjcf27/Af1nD50t/wurKfTKmUPH9587Ojd59qjR" & vbCrLf
    s = s & "03TC0gnedry7qWovVk+cvXh2/4HuupC3Ljdg7NCBc/Vqx1s9Iz905mh367C3OjhgDZm0VHLwALXG" & vbCrLf
    s = s & "JgAAAAAAAPgZHHGrYv1vOZFMk+5rqnLeE36P1fqUTU1HBras3mXgyMJcBiFThAtT7FvmBIGtTghs" & vbCrLf
    s = s & "J0JZDgAAAAAA8A8QolHaxLZ7/T8BAAAAAAAAgP6OumqoSi8/1okjhibtr83/+wL8DdPfw53/z5aC" & vbCrLf
    s = s & "hAmy7hUBWWGCpHtpQNYyQcoEafdigaxtgowJarqXD2QdE4RMUMsEtTlQg0Y4UGOFTVDDBHVMUNe9" & vbCrLf
    s = s & "6CAbNUHMBPjyCQAAAAAAgO8lxLaUiUOffP8fAQAAAAAAAPwuGYIqc0ST1I9dQQAAAAAAAAAAfzrM" & vbCrLf
    s = s & "/wMAAAAAAMCfpxrfRF+Vv17vTv8IOf8fIgAAAAAAAICP7d09aBNhGAfw/32kVnN3bwVBBAcdqkto" & vbCrLf
    s = s & "tZYoxkpziVDFIDX1A1xstbXaoqFN/agIpeAggoKoq2jbQRRxqB+TOBntIlLIoItgNulSELfGN9fG" & vbCrLf
    s = s & "Izk9IkIC4f97uVx4cvyf57hkeKdQPeP+n4iIiIiIiKj+cf9PREREREREVP8UpctUsawRL8RpERRv" & vbCrLf
    s = s & "rSFro5UxB0wQERERERHRP1mLgHxVAKFg8XjuXvhbrrN4lh9MJJDEFYziMPoxjF6kcRYXcF5WBuW7" & vbCrLf
    s = s & "lDzbn+LwLhubEES+0ECvqEE3xmSLERna74TugneF3VCtotAoYtjnxLXBu7a5cWpFcYfQh3NyvlNI" & vbCrLf
    s = s & "VzKj1KTKUNU3NC5DY/Mt8K5Qye0KyEP7Y5Tfw9kD79pdcuM+wfFCqDz65DG68mDa4V1tbqDiG9jj" & vbCrLf
    s = s & "hA2vRPnetKIkgxqWrceC+CwyYlY8FLfEVTEojghbtIrXVtSaM5Nmp/nBeGlMGSFjg/EoeJt/xk5E" & vbCrLf
    s = s & "RERE9H90Zw+SzwvY8wbWQIFwS9HVjc+P3bT0VCIydKDn/Y6JsROL23N3njT8fNqxX81+zL5ZiMwE" & vbCrLf
    s = s & "w81f7O+z7ZtTLVOZ+9cu5h93t04al6YTe9OR9N1X5o3Lk0e35q//CI13aR3vZvTegYPPHoyf/Lqz" & vbCrLf
    s = s & "eTpjzZ0Z2ZIqtlUKbZecSUpLVZ9EBbAO+D3JKh2Bv5UaqjqcbKt7J9FrM4kGqfpt1drcrdRk1aqt" & vbCrLf
    s = s & "87jLJ4m5JWfDX1rSZEnxlpbkhXbW/YV5r1K8VwVQYBWvKn7nyjt6S2p56RdQSwECPwAUAAAACADR" & vbCrLf
    s = s & "NFZK7jnFULQIAAAAMAEACQAkAAAAAAAAACAAAAAAAAAAZW1wdHkubWRiCgAgAAAAAAABABgA2cFS" & vbCrLf
    s = s & "9RCN0gHZwVL1EI3SAdnBUvUQjdIBUEsFBgAAAAABAAEAWwAAANsIAAAAAA=="
    GetZipAsBase64 = s
End Function
