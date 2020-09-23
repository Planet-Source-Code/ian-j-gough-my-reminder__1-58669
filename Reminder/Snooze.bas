Attribute VB_Name = "Snooze"
'Taken from http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=3552&lngWId=1
'Please vote for the above code if you make use of this function
'
'
'
Option Explicit
Public exitPause As Boolean
'
Public Function timedPause(secs As Long)
    Dim secStart As Variant
    Dim secNow As Variant
    Dim secDiff As Variant
    Dim Temp%
'
    exitPause = False 'this is our early way out out of the pause
'
    secStart = Format(Now(), "mm/dd/yyyy hh:nn:ss AM/PM") 'get the starting seconds
    
'
'
    Do While secDiff < secs
        If exitPause = True Then Exit Do
        secNow = Format(Now(), "mm/dd/yyyy hh:nn:ss AM/PM") 'this is the current time and date at any itteration of the loop
        secDiff = DateDiff("s", secStart, secNow) 'this compares the start time with the current time
        Temp% = DoEvents
    Loop
End Function
'

