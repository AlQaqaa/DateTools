Module Module1

    Sub Main()

        Dim dateTools As New DateTools

        Dim numberOfHolidays As Integer = dateTools.GetNumberOfHolidays("02/01/2021", Date.Now)
        Dim lastDayInMonth As Date = dateTools.LastDayInMonth(Date.Now)
        Dim numberOfBusinessDays As Integer = dateTools.GetBusinessDays("02/01/2021", Date.Now)
        Dim firstBusinessDayOfMonth As Date = dateTools.FirstBusinessDayOfMonth("5/11/2021")
        Dim lastBusinessDayOfMonth As Date = dateTools.LastBusinessDayOfMonth("7/11/2021")
        Dim isHoliday As Boolean = dateTools.IsHoliday(Date.Now)
        Dim hijriDate As String = dateTools.ToHijri(Date.Now.Date, "yyyy MMM dd")

        Console.WriteLine("Number Of Holidays: " & numberOfHolidays)
        Console.WriteLine("Last Day In This Month: " & lastDayInMonth)
        Console.WriteLine("Number Of Business Days: " & numberOfBusinessDays)
        Console.WriteLine("First Business Day Of Month: " & firstBusinessDayOfMonth)
        Console.WriteLine("Last Business Day Of Month: " & lastBusinessDayOfMonth)
        Console.WriteLine("Is Today a Holiday: " & isHoliday)
        Console.WriteLine("Today Is: " & hijriDate)
        Console.ReadLine()

    End Sub

End Module
