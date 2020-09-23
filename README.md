<div align="center">

## DateFunctions


</div>

### Description

Use this module to do many calculations concerning dates. I will maybe add a few more later on.

This module inlcudes the following Functions:

1. DayOfWeek (Returns the day of the week of a certain date)

2. DayOfYear (Returns the day of the year, eg. 31 December 1999 will be 365)

3. DaysBetween (Returns the amount of days between two dates)

4. DaysInMonth (Retruns the days in a specified month, eg. 29 in February 2000)

5. DaysInYear (Returns the days in a specific year, eg. 365 in 1999)

6. IsLeapYear (Returns whether a year is a leap year)

Come on in, and take a look!!
 
### More Info
 
All the functions have different input parameters.

Beginners will have to know how functions work. Basic knowledge, actually.

All the functions return something different


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Jan Botha](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/jan-botha.md)
**Level**          |Unknown
**User Rating**    |3.7 (11 globes from 3 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Math/ Dates](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/math-dates__1-37.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/jan-botha-datefunctions__1-3579/archive/master.zip)





### Source Code

```
'*************************************************
'*DATEFUNCTIONS                 *
'*                        *
'*By: Jan Botha                 *
'*eMail: c03jabot@prg.wcape.school.za      *
'*Date: Sunday, 19 September 1999        *
'*Inspired by David I Schneider's book,     *
'*  "An Introduction to Programming using   *
'*  Visual Basic 5.0 - Third Edition"     *
'*I only got one of the formulas out from his  *
'*book as well as the idea. As I programmed on  *
'*I got ideas for other functions too.      *
'*So here they are!               *
'*************************************************
Option Explicit
'This returns the day of the week of a certain date.
'It will only work with dates after 1582, because
'the calendar we use today was introduced then
Public Function DayOfWeek(ByVal Day As Integer, ByVal Month As Integer, ByVal Year As Integer) As String
  Dim w As Integer, wQuotient, wRemainder, int6
  If Month = 1 Then
    Month = 13
    Year = Year - 1
   ElseIf Month = 2 Then
    Month = 14
    Year = Year - 1
  End If
  int6 = 0.6 * (Month + 1)
  int6 = Int(int6)
  'I got this formula from David I Schneider's book
  '"An Introduction to Programming using Visual Basic 5.0 - Third Edition"
  w = Day + 2 * Month + int6 + Year + Int(Year / 4) - Int(Year / 100) + Int(Year / 400) + 2
  wQuotient = Int(w / 7)
  DayOfWeek = DayString(w - (wQuotient * 7))
End Function
'See what day of the year it is
Public Function DayOfYear(ByVal Day As Integer, ByVal Month As Integer, ByVal LeapYear As Boolean) As Integer
  Dim i As Integer, fDay As Integer
  For i = 1 To Month - 1
    fDay = fDay + DaysInMonth(i, LeapYear)
  Next
  fDay = fDay + Day
  DayOfYear = fDay
End Function
'This function check how many days there are between
'two certain dates
Public Function DaysBetween(ByVal startDay As Integer, ByVal startMonth As Integer, ByVal startYear As Integer, ByVal endDay As Integer, ByVal endMonth As Integer, ByVal endYear As Integer) As Long
  Dim startIsLeap As Boolean, endIsLeap As Boolean
  Dim daysToEnd As Integer, fDays As Integer
  startIsLeap = IsLeapYear(startYear)
  endIsLeap = IsLeapYear(endYear)
  startDay = DayOfYear(startDay, startMonth, startIsLeap)
  endDay = DayOfYear(endDay, endMonth, endIsLeap)
  If startYear = endYear Then
    DaysBetween = endDay - startDay
    Exit Function
  End If
  daysToEnd = DaysInYear(startYear) - startDay
  For i = startYear + 1 To endYear - 1
    fDays = fDays + DaysInYear(i)
  Next
  fDays = fDays + daysToEnd + endDay
  DaysBetween = fDays
End Function
Public Function DaysInMonth(ByVal Month As Integer, ByVal LeapYear As Boolean) As Integer
  Select Case Month
    Case 1, 3, 5, 7, 8, 10, 12: DaysInMonth = 31
    Case 2
      If LeapYear Then
        DaysInMonth = 29
       Else
        DaysInMonth = 28
      End If
    Case 4, 6, 9, 11: DaysInMonth = 30
  End Select
End Function
'Use this function to determine how many days there are in a year
Public Function DaysInYear(ByVal Year As Integer) As Integer
  'leap years have 366 days and other years have
  '365. simple
  If IsLeapYear(Year) Then
    DaysInYear = 366
   Else
    DaysInYear = 365
  End If
End Function
Private Function DayString(ByVal Weekday As Integer)
  'this function is used by the DayOfWeek function only
  Select Case Weekday
    Case 0: DayString = "Saturday"
    Case 1: DayString = "Sunday"
    Case 2: DayString = "Monday"
    Case 3: DayString = "Tuesday"
    Case 4: DayString = "Wednesday"
    Case 5: DayString = "Thursday"
    Case 6: DayString = "Friday"
  End Select
End Function
' Use this function to determine if a certain year is a leap year.
Public Function IsLeapYear(ByVal Year As Integer) As Boolean
  If Year Mod 4 = 0 Then
    IsLeapYear = True
    If Year Mod 100 = 0 And Year Mod 400 <> 0 Then
      IsLeapYear = False
    End If
  End If
  'all years divisible by 4 are leap years with the exception
  'of years that are divisible by 100 and not by 400
End Function
Please email me comments, suggestions and especially BUGS!
c03jabot@prg.wcape.school.za
```

