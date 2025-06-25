Attribute VB_Name = "Time"
Declare Sub GetSystemTime Lib "Kernel32" (lpSystemTime As SYSTEMTIME)
Declare Function SetSystemTime Lib "Kernel32" (lpSystemTime As SYSTEMTIME) As Long
Global pathstr
Type SYSTEMTIME
        wYear As Long
        wMonth As Long
        wDayOfWeek As Long
        wDay As Long
        wHour As Long
        wMinute As Long
        wSecond As Long
        wMilliseconds As Long
End Type
Function gettimedifference(h As Long, m As Long, s As Long) As Long
Dim GMTime As SYSTEMTIME
Dim h1 As Long
Dim m1 As Long
Dim s1 As Long
Dim mseconds1 As Long
Dim mseconds2 As Long
On Error GoTo terror
    mseconds1 = (h * 1000 * 3600) + (m * 1000 * 60) + (s * 1000)
    GetSystemTime GMTime
    
    h1 = GMTime.wHour
    m1 = GMTime.wMinute
    s1 = GMTime.wSecond
    mseconds2 = (h1 * 1000 * 3600) + (m1 * 1000 * 60) + (s1 * 1000)
    gettimedifference = (mseconds2 - mseconds1) / 1000
Exit Function
terror:
    Open pathstr For Append As #1
    Write #1, "Error In gettimedifference in Timeclass", Now
    Write #1, "Error is ", Err.Description
    Close #1
    
End Function



