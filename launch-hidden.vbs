' تشغيل اداة تسجيل البوستة في الخلفية (بيفتح مع تسجيل الدخول، من غير نافذة).
' بيدور على pythonw بالظبط، وبعدين بيشغل postal_ocr.py في نفس الفولدر.
Option Explicit
Dim sh, fso, base, py, cmd, f, tries, i

Set sh = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
base = Left(WScript.ScriptFullName, InStrRev(WScript.ScriptFullName, "\") - 1)

py = ""

' 1) المسار المكتوب وقت التثبيت (الادق)
f = base & "\pythonw.txt"
If fso.FileExists(f) Then
    On Error Resume Next
    py = Trim(fso.OpenTextFile(f, 1).ReadAll)
    On Error GoTo 0
    If Len(py) > 0 And Not fso.FileExists(py) Then py = ""
End If

' 2) المسارات المعتادة
If py = "" Then
    tries = Array( _
        sh.ExpandEnvironmentStrings("%LOCALAPPDATA%") & "\Programs\Python\Python312\pythonw.exe", _
        sh.ExpandEnvironmentStrings("%LOCALAPPDATA%") & "\Programs\Python\Python313\pythonw.exe", _
        sh.ExpandEnvironmentStrings("%LOCALAPPDATA%") & "\Programs\Python\Python311\pythonw.exe", _
        sh.ExpandEnvironmentStrings("%ProgramFiles%") & "\Python312\pythonw.exe", _
        sh.ExpandEnvironmentStrings("%ProgramFiles%") & "\Python313\pythonw.exe")
    For i = 0 To UBound(tries)
        If fso.FileExists(tries(i)) Then
            py = tries(i)
            Exit For
        End If
    Next
End If

' 3) اي بايثون موجود في المسار
If py = "" Then py = "pythonw.exe"

cmd = """" & py & """ """ & base & "\postal_ocr.py"""
sh.CurrentDirectory = base
sh.Run cmd, 0, False
