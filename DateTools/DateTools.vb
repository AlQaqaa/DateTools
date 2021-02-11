
''' <summary>
''' Date tools: https://github.com/AlQaqaa/DateTools
''' Created by AlQaqaa benghuzi che.ben.arg@gmail.com, on 11-2-2021
''' Contact me on: 
''' Github: https://github.com/AlQaqaa
''' Facebook: https://fb.com/alqaqaaben
''' Twitter: https://twitter.com/alqaqaaben
''' </summary>

Public Class DateTools

    ' Get number of holidays between two dates
    Public Function GetNumberOfHolidays(startDay As DateTime, endDay As DateTime) As Integer

        Dim numberOfDays As Integer = 0

        While (startDay <= endDay)

            Select Case startDay.DayOfWeek
                Case DayOfWeek.Friday
                    numberOfDays += 1
                Case DayOfWeek.Saturday
                    numberOfDays += 1
            End Select

            startDay = startDay.AddDays(1)
        End While

        Return numberOfDays
    End Function

    'Get Last Day In Month
    Public Function LastDayInMonth(ByVal dayDate As Date) As Date

        Dim numberDaysOfMonth As Integer = System.DateTime.DaysInMonth(dayDate.Year, dayDate.Month)
        Dim lastDay As Date = New DateTime(Year(dayDate), Month(dayDate), numberDaysOfMonth)

        Return lastDay
    End Function

    ' Get number of business days between two dates
    Public Function GetBusinessDays(startDay As DateTime, endDay As DateTime) As Integer

        Dim numberOfDays As Integer = (endDay.Date - startDay.Date).Days + 1 - GetNumberOfHolidays(startDay, endDay)

        Return numberOfDays
    End Function

    'Get First Business Day of Month
    Public Function FirstBusinessDayOfMonth(ByVal dayDate As Date) As Date

        Dim firstDay As Date = New DateTime(dayDate.Year, dayDate.Month, 1)

        Select Case dayDate.DayOfWeek
            Case DayOfWeek.Friday
                firstDay = firstDay.AddDays(2)
            Case DayOfWeek.Saturday
                firstDay = firstDay.AddDays(1)
        End Select

        Return firstDay
    End Function

    'Get Last Business Day of Month
    Public Function LastBusinessDayOfMonth(ByVal dayDate As Date) As Date

        Dim lastDay As Date = New DateTime(dayDate.Year, dayDate.Month, System.DateTime.DaysInMonth(dayDate.Year, dayDate.Month))

        Select Case dayDate.DayOfWeek
            Case DayOfWeek.Friday
                lastDay = lastDay.AddDays(-1)
            Case DayOfWeek.Saturday
                lastDay = lastDay.AddDays(-2)
        End Select

        Return lastDay
    End Function

    'Returns True if Date passed is a Holiday
    Public Function IsHoliday(ByVal dayDate As Date) As Boolean

        Dim isHolidya As Boolean = False

        Select Case dayDate.DayOfWeek
            Case DayOfWeek.Friday
                isHolidya = True
            Case DayOfWeek.Saturday
                isHolidya = True
        End Select

        Return isHolidya
    End Function

    ' Convert gregorian date to Hijri date
    Public Function ToHijri(ByVal gDate As Date, Optional ByVal format As String = Nothing) As String
        Return gDate.ToString(format, New Globalization.CultureInfo("ar-SA"))
    End Function

End Class
