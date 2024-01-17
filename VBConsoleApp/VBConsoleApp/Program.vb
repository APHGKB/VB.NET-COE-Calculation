Imports System
Imports System.Net.Http
Imports Astronomy
Imports SunCalcNet

Module Program
    'Holidays 2024
    'New DateTime([Year],[Month],[Day])
    'This holiday refers from office calendar
    ' <---- Edit is to Modify date after we arrived new year
    Dim customHolidays As New List(Of DateTime) From {
        New DateTime(DateTime.Now.Year - 1, 10, 28), 'Thadingyut Holiday <---- Edit [Previous Year]
        New DateTime(DateTime.Now.Year - 1, 10, 29), 'Thadingyut Holiday <---- Edit [Previous Year]
        New DateTime(DateTime.Now.Year - 1, 10, 30), 'Thadingyut Holiday <---- Edit [Previous Year]
        New DateTime(DateTime.Now.Year - 1, 11, 12),  'Deepavali [Previous Year]
        New DateTime(DateTime.Now.Year - 1, 11, 27), 'Full Moon of  Tasaungmone   <----- Edit [Previous Year]
        New DateTime(DateTime.Now.Year - 1, 11, 7), 'National Day [Previous Year]
        New DateTime(DateTime.Now.Year - 1, 12, 25), 'Chirstmas Day [Previous Year]
        New DateTime(DateTime.Now.Year - 1, 12, 31),  'New Year [Previous Year]
        New DateTime(DateTime.Now.Year, 1, 1),   'First day of new year
        New DateTime(DateTime.Now.Year, 1, 4),   'Independent Day
        New DateTime(DateTime.Now.Year, 2, 12),  'Union Day
        New DateTime(DateTime.Now.Year, 3, 2),   'Farmers' Day
        New DateTime(DateTime.Now.Year, 3, 25),  'Full Moon Day of Tabaung <---- Edit
        New DateTime(DateTime.Now.Year, 3, 27),  'Armed Forces Day
        New DateTime(DateTime.Now.Year, 4, 13),  'Thingyun Holiday 
        New DateTime(DateTime.Now.Year, 4, 14),  'Thingyun Holiday 
        New DateTime(DateTime.Now.Year, 4, 15),  'Thingyun Holiday 
        New DateTime(DateTime.Now.Year, 4, 16),  'Thingyun Holiday 
        New DateTime(DateTime.Now.Year, 5, 1),   'Labour Day
        New DateTime(DateTime.Now.Year, 7, 17),  'Eid ul-Adha Day
        New DateTime(DateTime.Now.Year, 7, 19),  'Martyrs' Day
        New DateTime(DateTime.Now.Year, 7, 22),
        New DateTime(DateTime.Now.Year, 10, 16), 'Thadingyut Holiday <---- Edit
        New DateTime(DateTime.Now.Year, 10, 17), 'Thadingyut Holiday <---- Edit
        New DateTime(DateTime.Now.Year, 10, 18), 'Thadingyut Holiday <---- Edit
        New DateTime(DateTime.Now.Year, 11, 1),  'Deepavali
        New DateTime(DateTime.Now.Year, 11, 15), 'Full Moon of  Tasaungmone   <----- Edit
        New DateTime(DateTime.Now.Year, 11, 16),
        New DateTime(DateTime.Now.Year, 12, 25), 'Chirstmas Day
        New DateTime(DateTime.Now.Year, 12, 31), 'New Year
        New DateTime(DateTime.Now.Year + 1, 1, 1),   'First day of new year [New Year]
        New DateTime(DateTime.Now.Year + 1, 1, 4),   'Independent Day [New Year]
        New DateTime(DateTime.Now.Year + 1, 2, 12),  'Union Day [New Year]
        New DateTime(DateTime.Now.Year + 1, 3, 2),   'Farmers' Day [New Year]
        New DateTime(DateTime.Now.Year + 1, 3, 14),  'Full Moon Day of Tabaung <---- Edit [New Year]
        New DateTime(DateTime.Now.Year + 1, 3, 27),  'Armed Forces Day [New Year]
        New DateTime(DateTime.Now.Year + 1, 4, 13),  'Thingyun Holiday [New Year]
        New DateTime(DateTime.Now.Year + 1, 4, 14),  'Thingyun Holiday [New Year]
        New DateTime(DateTime.Now.Year + 1, 4, 15),  'Thingyun Holiday [New Year]
        New DateTime(DateTime.Now.Year + 1, 4, 16),  'Thingyun Holiday [New Year]
        New DateTime(DateTime.Now.Year + 1, 5, 1),   'Labour Day [New Year]
        New DateTime(DateTime.Now.Year + 1, 5, 11)  'Full Moon Day of Kasone <---- Edit [New Year]
    }
    Dim nyugokudate As DateTime
    Dim calculateDate As Integer
    Dim intOption As Integer
    Dim nyukokuEndDate As DateTime
    Dim nyukokuYear As Integer
    Dim nyukokuMonth As Integer
    Dim nyukokuDay As Integer
    Dim myKey As ConsoleKeyInfo
    Dim ECdate() As String
    Dim interviewDate() As String
    Dim sakuseiDate() As String
    Sub Main(args As String())
        showCOEOptions()

        Try
            intOption = CInt(Console.ReadLine())
        Catch ex As Exception
            Console.WriteLine(" You must type only number ")
            Environment.Exit(0)
        End Try
        If Not checkOption(intOption) Then
            Console.WriteLine(" You must type only number between 1 and 4")
            Environment.Exit(0)
        End If
        Console.Write("Total Calculation Days [ 60 or 66 ] :")
        Try
            calculateDate = CInt(Console.ReadLine())
            If Not checkCalculateDate(calculateDate) Then
                Console.WriteLine(" Your number is invalid. ")
                Environment.Exit(0)
            End If
        Catch ex As Exception
            Console.WriteLine("You must enter only numbers ")
            Environment.Exit(0)
        End Try
        Console.WriteLine("    Example of EC date :: 2024/1/1")
        Console.Write("Enter EC Date ::")
        Try
            ECdate = Split(Console.ReadLine, "/")
            If Not checkECOrInterviewDate(ECdate) Then
                Console.WriteLine("EC Date Format is invalid ")
                Environment.Exit(0)
            End If
        Catch ex As Exception
            Console.WriteLine(" EC Date is Invalid or Empty")
            Environment.Exit(0)
        End Try
        Console.WriteLine("    Example of Interview date :: 2024/1/1")
        Console.Write("Enter Interview Date ::")
        Try
            interviewDate = Split(Console.ReadLine, "/")
            If Not checkECOrInterviewDate(interviewDate) Then
                Console.WriteLine("EC Date Format is invalid ")
                Environment.Exit(0)
            End If
        Catch ex As Exception
            Console.WriteLine(" Interview Date is Invalid or Empty")
            Environment.Exit(0)
        End Try
        If Not compareECAndInterview(ECdate, interviewDate) Then
            Console.WriteLine(" EC date and Interview date are so far. ")
            Environment.Exit(0)
        End If
        Console.Write("Enter Nyukoku Year ::")
        Try
            nyukokuYear = Console.ReadLine()
        Catch ex As Exception
            Console.WriteLine(" Nyukoku Year is Empty.")
            Environment.Exit(0)
        End Try
        ' Check NyuKoKu Year Input
        If Not checkYear(nyukokuYear) Then
            Console.WriteLine(" Nyukoku Year is invalid.")
            Environment.Exit(0)
        End If
        Console.Write("Enter Nyukoku Month ::")
        Try
            nyukokuMonth = Console.ReadLine()
        Catch ex As Exception
            Console.WriteLine(" Nyukoku Month is Empty.")
            Environment.Exit(0)
        End Try
        ' Check NyuKoKu Month Input
        If Not checkMonth(nyukokuMonth) Then
            Console.WriteLine(" Nyukoku Month is invalid.")
            Environment.Exit(0)
        End If
        Console.Write("Enter Nyukoku Day ::")
        Try
            nyukokuDay = Console.ReadLine()
        Catch ex As Exception
            Console.WriteLine(" Nyukoku Day is Empty.")
            Environment.Exit(0)
        End Try
        If Not checkDay(nyukokuDay, nyukokuMonth) Then
            Console.WriteLine(" Nyukoku Day is invalid.")
            Environment.Exit(0)
        End If
        Console.Write("Do you get sakusei date from Viber Or Mail : ")
        myKey = Console.ReadKey()
        If myKey.Key = ConsoleKey.Y Then
            Console.WriteLine("    Example of Sakusei date :: 2024/1/1")
            Console.Write("Enter your sakusei Date ::")
            Try
                sakuseiDate = Split(Console.ReadLine, "/")
                If Not checkECOrInterviewDate(sakuseiDate) Then
                    Console.WriteLine("Sakusei Date Format is invalid ")
                    Environment.Exit(0)
                End If
            Catch ex As Exception
                Console.WriteLine(" EC Date is Invalid or Empty")
                Environment.Exit(0)
            End Try
        ElseIf myKey.Key = ConsoleKey.N Then
            sakuseiDate = calculateSakuSeiDate()
        Else
            Console.WriteLine(" You clicked wrong key ")
            Environment.Exit(0)
        End If
        Console.WriteLine()
        nyugokudate = New DateTime(nyukokuYear, nyukokuMonth, nyukokuDay)
        Console.WriteLine("NyuGokuDate : " + nyugokudate.ToShortDateString)
        nyugokudate = getNihongoStartDate(nyugokudate)
        'loop start point as 1
        'daycount 60 or 66 60 is nogyou 66 is course
        Dim startpoint As Integer = 1
        Dim dayCount As Integer = 1
        Console.WriteLine("Nihongo " + calculateDate.ToString + " days")
        Console.WriteLine("----------------")
        'Check
        Console.WriteLine(nyugokudate.Year.ToString + "/ " + nyugokudate.Month.ToString + "/ " + nyugokudate.Day.ToString + "    " + calculateDate.ToString)
        While dayCount <> calculateDate
            nyukokuEndDate = nyugokudate.AddDays(-startpoint)
            startpoint += 1
            'Console.WriteLine(dayCount)
            'Console.WriteLine(startpoint)
            If isHoliday(nyukokuEndDate) Then
                Continue While
            End If
            dayCount += 1
            Console.WriteLine(nyukokuEndDate.Year.ToString + "/ " + nyukokuEndDate.Month.ToString + "/ " + nyukokuEndDate.Day.ToString + "        " + ((calculateDate + 1) - dayCount).ToString)
        End While
        Console.WriteLine("----------------")
        outputResult()
        Console.ReadKey()
    End Sub

    'COE Options Display
    Function showCOEOptions()
        Console.WriteLine("Please type a number that you want to do")
        Console.WriteLine("   --> [1]  COURSE ")
        Console.WriteLine("   --> [2]  NOUGYOU ")
        Console.WriteLine("   --> [3]  KAISHA ")
        Console.WriteLine("   --> [4]  KAIGO ")
        Console.Write("Number ::")
    End Function

    'Check Option
    Function checkOption(coeOption As Integer) As Boolean
        If coeOption >= 1 And coeOption < 5 Then
            Return True
        End If
        Return False
    End Function

    'Check Holiday 
    Function isHoliday(dateCheck As DateTime) As Boolean
        For Each currentDate As DateTime In customHolidays
            If dateCheck.Date = currentDate.Date Then
                Return True
                Exit For
            End If
        Next
        Return isWeekend(dateCheck)
    End Function

    'Get NihonGoStartDate with DateTime DataType
    Function getNihongoStartDate(dateCheck As DateTime) As DateTime
        If dateCheck.Day = 1 Then
            dateCheck = dateCheck.AddMonths(-1)
            'Console.Write(dateCheck)
            dateCheck = dateCheck.AddDays(-16)
        Else
            dateCheck = dateCheck.AddMonths(-1)
            dateCheck = New DateTime(dateCheck.Year, dateCheck.Month, 1)
            dateCheck = dateCheck.AddDays(-11)
        End If
        'Console.WriteLine(isHoliday(dateCheck))
        While isHoliday(dateCheck)
            'Console.Write(dateCheck)
            dateCheck = dateCheck.AddDays(-1)
            'Console.WriteLine(isHoliday(dateCheck))
        End While
        Return dateCheck
    End Function

    'Check Weekend Day
    Function isWeekend(dateCheck As DateTime) As Boolean
        Return dateCheck.DayOfWeek = DayOfWeek.Saturday Or dateCheck.DayOfWeek = DayOfWeek.Sunday
    End Function

    'Check year
    Function checkYear(input_year As String) As Boolean
        Dim now_year As Integer = DateTime.Now.Year
        If input_year.Length < 5 Then
            If CInt(input_year) = now_year Or CInt(input_year) = (now_year - 1) Or CInt(input_year) = (now_year + 1) Then
                Return True
            End If
        End If
        Return False
    End Function

    Function checkMonth(input_month As String) As Boolean
        'Console.WriteLine(CInt(input_month) > 1 And CInt(input_month) < 12)
        If CInt(input_month) >= 1 And CInt(input_month) <= 12 Then
            Return True
        End If
        Return False
    End Function

    Function checkDay(input_day As String, input_month As String) As Boolean
        Dim thtyO() As Integer = {1, 3, 5, 7, 8, 10, 12}
        Dim thty() As Integer = {4, 6, 9, 11}
        Dim chk As Boolean
        For Each toMonth As Integer In thtyO
            If input_month = toMonth Then
                If (input_day >= 1 And input_day <= 31) Then
                    Return True
                End If
                Exit For
            End If
        Next
        For Each tMonth As Integer In thtyO
            If input_month = tMonth Then
                If (input_day >= 1 And input_day <= 30) Then
                    Return True
                End If
                Exit For
            End If
        Next
        If (input_day >= 1 And input_day <= 28) Then
            Return True
        End If
        Return False
        'Edit
    End Function
    Function checkCalculateDate(caldate As Integer) As Boolean
        If caldate = 60 Or caldate = 66 Then
            Return True
        End If
        Return False
    End Function

    'EC date and Interview Date formats are valid or not
    Function checkECOrInterviewDate(eiDate() As String) As Boolean
        If checkYear(eiDate(0)) And checkMonth(eiDate(1)) And checkDay(eiDate(2), eiDate(1)) Then
            'Console.WriteLine("True 1")
            If Not isHoliday(New DateTime(eiDate(0), eiDate(1), eiDate(2))) Then
                Return True
            End If
        End If
        Return False
    End Function

    'Year must be same both interview date and ec date
    'Month and Day are not same
    Function compareECAndInterview(ecDate() As String, interviewDate() As String)
        If ecDate(0).Equals(interviewDate(0)) And ecDate(2).Equals(interviewDate(2)) = False Then
            Return True
        End If
        Return False
    End Function

    Function customFormatString(label As String, myYear As Integer, myMonth As Integer, myDay As Integer) As String
        Return String.Format("{0} : {1}/ {2}/ {3}", {label, myYear.ToString, myMonth.ToString, myDay.ToString})
    End Function

    Function calculateSakuSeiDate() As String()
        Dim tempDate() As String
        If CInt(ECdate(1)) > CInt(interviewDate(1)) Then
            tempDate = ECdate
            tempDate(2) = tempDate(2) + 2
            While Not checkECOrInterviewDate(tempDate)
                tempDate(2) = tempDate(2) + 1
            End While
            Return tempDate
        ElseIf CInt(ECdate(1)) < CInt(interviewDate(1)) Then
            tempDate = interviewDate
            tempDate(2) = tempDate(2) + 2
            While Not checkECOrInterviewDate(tempDate)
                tempDate(2) = tempDate(2) + 1
            End While
            Return tempDate
        Else
            If CInt(ECdate(2)) < CInt(interviewDate(2)) Then
                tempDate = interviewDate
                tempDate(2) = tempDate(2) + 2
                While Not checkECOrInterviewDate(tempDate)
                    tempDate(2) = tempDate(2) + 1
                End While
                Return tempDate
            Else
                tempDate = ECdate
                tempDate(2) = tempDate(2) + 2
                While Not checkECOrInterviewDate(tempDate)
                    tempDate(2) = tempDate(2) + 1

                End While
                Return tempDate
            End If
        End If
    End Function
    Function outputResult()
        Dim course_end_date As DateTime = nyukokuEndDate.AddMonths(-1)
        Dim course_start_date As DateTime = course_end_date.AddMonths(-3)
        If intOption = 1 Or intOption = 3 Then
            Console.WriteLine(customFormatString("Nihongo Start Date", nyukokuEndDate.Year, nyukokuEndDate.Month, nyukokuEndDate.Day))
            Console.WriteLine(customFormatString("Nihongo End Date", nyugokudate.Year, nyugokudate.Month, nyugokudate.Day))
            Console.WriteLine(customFormatString("Course Start Date", course_start_date.Year, course_start_date.Month, course_start_date.Day))
            Console.WriteLine(customFormatString("Course End Date", course_end_date.Year, course_end_date.Month, course_end_date.Day))
            Console.WriteLine(customFormatString("Sakusei Date", sakuseiDate(0), sakuseiDate(1), sakuseiDate(2)))
        Else
            Console.WriteLine(customFormatString("Nihongo Start Date", nyukokuEndDate.Year, nyukokuEndDate.Month, nyukokuEndDate.Day))
            Console.WriteLine(customFormatString("Nihongo End Date", nyugokudate.Year, nyugokudate.Month, nyugokudate.Day))
            Console.WriteLine(customFormatString("Sakusei Date", sakuseiDate(0), sakuseiDate(1), sakuseiDate(2)))
        End If
    End Function

End Module
