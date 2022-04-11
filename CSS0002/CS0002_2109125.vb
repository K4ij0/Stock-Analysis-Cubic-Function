Imports System.Math

Module StockAnalysis_CubicFunction

    'Giving an array of 252 
    Dim arr_TrDates(251) As String
    Dim arr_OpenPrices(251) As Double
    Dim arr_HighPrices(251) As Double
    Dim arr_LowPrices(251) As Double
    Dim arr_ClosePrices(251) As Double
    Dim arr_Volume(251) As Integer
    Sub Main(args As String())


        menu()

    End Sub


    Sub menu()
        ' Displays Menu Data
        Console.Title = " CS0002 Assignment"
        Threading.Thread.Sleep(50)
        ' Menu Display
        Console.WriteLine("╔═══════════════════╗")
        Console.WriteLine("   Welcome To Menu ")
        Console.WriteLine("╚═══════════════════╝")

        'Spawns after 2 second
        Console.ForegroundColor = ConsoleColor.White

        Console.WriteLine("────────────")
        Console.WriteLine("└>1: Find the max and min")
        Console.WriteLine("└>2: Stock Analysis")
        Console.WriteLine("└>3: Exit")
        Console.WriteLine("────────────")
        Console.ResetColor()
        Console.BackgroundColor = ConsoleColor.DarkRed
        Threading.Thread.Sleep(500)
        Console.Write("Enter your number:")
        Console.ResetColor()
        Console.Write(" ")


        Dim input As Integer = 0



        ' Checks if selects a right number between 1-3, if not it loops again to make the user choose a right number.

        ' Checks the value of accessing number
        If Integer.TryParse(Console.ReadLine(), input) Then
            Select Case input
                Case 1
                    'Appendix A'
                    AppendixA()

                Case 2
                    'Appendix B'
                    Console.Clear()
                    AppendixB()

                Case 3
                    'Exit'           
                    Console.Clear()

                    Console.WriteLine("════════════════════════════ - - - - - - - - >")
                    Console.Write("Terminating.... ")
                    End

                Case Else
                    ' If he inputs the wrong number in the case, it will say wrong number
                    Console.Clear()
                    Console.WriteLine("---------------------------------------")
                    Console.Write("/* Wrong number, use only number 1 to 3 *\ ")
                    Console.WriteLine()
                    Console.WriteLine("---------------------------------------")
                    Threading.Thread.Sleep(600)
                    Console.Clear()

                    menu()
            End Select
        Else
            ' Now, if you have a letter/symbol, a message will show that the user has inputted a wrong character, and then it will stay for 6 second, and restart menu
            Console.Clear()
            Console.WriteLine("---------------------------------------")
            Console.Write("/* Wrong character, use only numbers *\ ")
            Console.WriteLine()
            Console.WriteLine("---------------------------------------")
            Threading.Thread.Sleep(600)
            Console.Clear()
            menu()
        End If

        'Restarting Menu'

    End Sub



    Sub AppendixA()

        Dim a, b, c, d As Integer
        Console.WriteLine("────────────")
        Console.Write("Please enter your a,b,c,d number for the cubic function aX^3+bX^2+cX+d")
        Console.WriteLine()

        ' I have to make sure it only takes numbers, not letters

        ' The reason I am not just declaring the value to be in a straight away 
        ' is because we first have to check, if the user input is a number 
        ' if it is it will store into it, else it wont' in this case we just need to check
        ' if is not a intenger, it will show this... 
        ' To do this I used Integer.TryParse.
        '


        Console.BackgroundColor = ConsoleColor.DarkRed
        Console.Write("Enter A :")
        Console.ResetColor()
        Console.Write(" ")

        If Not Integer.TryParse(Console.ReadLine(), a) Then

            Console.Clear()
            Console.WriteLine("---------------------------------------")
            Console.Write("/* Wrong input for [A]... use only numbers *\ ")
            Console.WriteLine()
            Console.WriteLine("---------------------------------------")
            Threading.Thread.Sleep(600)
            Console.Clear()
            AppendixA()
        ElseIf a = 0 Then
            Console.Clear()
            Console.WriteLine("---------------------------------------")
            Console.Write("/* [A] cannot be = 0  *\ ")
            Console.WriteLine()
            Console.WriteLine("---------------------------------------")
            Threading.Thread.Sleep(600)
            Console.Clear()
            AppendixA()
        End If

        Console.BackgroundColor = ConsoleColor.DarkRed
        Console.Write("Enter B :")
        Console.ResetColor()
        Console.Write(" ")

        If Not Integer.TryParse(Console.ReadLine(), b) Then
            Console.Clear()
            Console.WriteLine("----------------------------------------------------")
            Console.Write("/* Wrong input for [B]... Restarting Application... *\ ")
            Console.WriteLine()
            Console.WriteLine("----------------------------------------------------")
            Threading.Thread.Sleep(500)
            Console.Clear()
            AppendixA()
        End If

        Console.BackgroundColor = ConsoleColor.DarkRed
        Console.Write("Enter C :")
        Console.ResetColor()
        Console.Write(" ")

        If Not Integer.TryParse(Console.ReadLine(), c) Then
            Console.Clear()

            Console.WriteLine("----------------------------------------------------")
            Console.Write("/* Wrong input for [C]... Restarting Application...*\ ")
            Console.WriteLine()
            Console.WriteLine("----------------------------------------------------")
            Console.ResetColor()
            Threading.Thread.Sleep(500)
            Console.Clear()
            AppendixA()

        End If

        Console.BackgroundColor = ConsoleColor.DarkRed
        Console.Write("Enter D :")
        Console.ResetColor()
        Console.Write(" ")

        If Not Integer.TryParse(Console.ReadLine(), d) Then
            Console.Clear()
            Console.WriteLine("----------------------------------------------------")
            Console.Write("/* Wrong number for [D]... Restarting Application...* ")
            Console.WriteLine()
            Console.WriteLine("----------------------------------------------------")
            Threading.Thread.Sleep(500)
            Console.Clear()
            AppendixA()
        End If




        ' Getting the roots from the formula 𝑥𝑟𝑜𝑜𝑡 = −2𝑏±√(2𝑏) 2−4∗(3𝑎)∗𝑐 2*(3a)
        Dim x_root, x2_root As Double
        x_root = ((-2 * b) + Sqrt((2 * b) ^ 2 - (4 * 3 * a) * c)) / (2 * 3 * a)

        x2_root = ((-2 * b) - Sqrt((2 * b) ^ 2 - (4 * 3 * a) * c)) / (2 * 3 * a)



        ' Console.WriteLine("The positve Root is : " & x_root)
        '  Console.WriteLine("The negative Root is : " & x2_root)

        ' Plugging into the f''x values
        ' One for the negative and positive (sign)
        Dim second_plus_derfx, second_minus_derfx As String

        second_plus_derfx = (6 * a * x_root) + (2 * b)

        second_minus_derfx = (6 * a * x2_root) + (2 * b)

        ' Calculating the cubic function inside the if statement 𝑓(𝑥) = 𝑎𝑥 + 3 +𝑏𝑥2 + 𝑐𝑥 + d
        Dim cubic_fuc As Decimal

        ' Implementing the First  derivative

        Dim first_plus_derfx, first_minus_derfx As String
        first_plus_derfx = 3 * a * (x_root) ^ 2 + (2 * b) + c
        first_minus_derfx = 3 * a * (x2_root) ^ 2 + (2 * b) + c

        ' I'm gonna check if its greater than 0 to find max and min
        If (second_plus_derfx > 0 And second_minus_derfx < 0) Then
            cubic_fuc = a * (x_root) ^ 3 + b * (x_root) ^ 2 + c * (x_root) + d
            Console.WriteLine("────────────")

            Console.Write("The local min is: ")
            Console.ForegroundColor = ConsoleColor.Green
            Console.Write(Math.Round(x_root, 2), " ")
            Console.ResetColor()

            Console.Write(" f(x) is: ")
            Console.ForegroundColor = ConsoleColor.Green
            Console.Write(Math.Round(cubic_fuc, 2))
            Console.ResetColor()
            Console.WriteLine()



            cubic_fuc = a * (x2_root) ^ 3 + b * (x2_root) ^ 2 + c * (x2_root) + d

            Console.WriteLine("────────────")
            Console.Write("The local max is: ")
            Console.ForegroundColor = ConsoleColor.Green
            Console.Write(Math.Round(x2_root, 2), " ")
            Console.ResetColor()
            Console.Write(" f(x) is: ")
            Console.ForegroundColor = ConsoleColor.Green
            Console.Write(Math.Round(cubic_fuc, 2))
            Console.ResetColor()
            ' Does the opposite
        ElseIf second_plus_derfx < 0 And second_minus_derfx > 0 Then

            cubic_fuc = a * (x_root) ^ 3 + b * (x_root) ^ 2 + c * (x_root) + d

            Console.WriteLine("────────────")
            Console.Write("The local max is :")
            Console.ForegroundColor = ConsoleColor.Green
            Console.Write(Math.Round(x_root, 2))
            Console.ResetColor()
            Console.Write("f(x) is: ")
            Console.ForegroundColor = ConsoleColor.Green
            Console.Write(Math.Round(cubic_fuc, 2))
            Console.ResetColor()


            cubic_fuc = a * (x2_root) ^ 3 + b * (x2_root) ^ 2 + c * (x2_root) + d
            Console.WriteLine("────────────")

            Console.Write("The local min is : ")
            Console.ForegroundColor = ConsoleColor.Green
            Console.Write(Math.Round(x2_root, 2))
            Console.ResetColor()
            Console.Write(" f(x) is :")
            Console.ForegroundColor = ConsoleColor.Green
            Console.Write(Math.Round(cubic_fuc, 2))
            Console.ResetColor()

            'If both der = 0 
        ElseIf first_plus_derfx And first_minus_derfx = 0 And second_plus_derfx And second_minus_derfx = 0 Then

            Console.WriteLine("────────────")
            Console.ForegroundColor = ConsoleColor.Red
            Console.WriteLine("The function might have a inflection point")
            Console.ResetColor()

        Else
            ' if no root are found
            Console.WriteLine("────────────")
            Console.ForegroundColor = ConsoleColor.Red
            Console.WriteLine("No real root found for f(x) therefore no maximum or minimum")
            Console.ResetColor()
        End If


        'Now, the function came to an end
        ' So I am going to allow the person to terminate the program or go back to the menu
        Console.WriteLine()
        Terminating()

    End Sub

    Sub AppendixB()

        Dim input2 As Integer = 0


        'Display menu of Data Analysis
        Console.ForegroundColor = ConsoleColor.White
        Console.WriteLine("╔═══════════════════╗")
        Console.WriteLine("    Stock Analysis")
        Console.WriteLine("╚═══════════════════╝")

        Console.WriteLine("────────────")
        Console.WriteLine("└>1: Yearly Data Analysis")
        Console.WriteLine("└>2: Monthly data Analysis")
        Console.WriteLine("└>3: Exit to main menu")
        Console.WriteLine("────────────")
        Console.ResetColor()
        Console.BackgroundColor = ConsoleColor.DarkRed
        Console.Write("Enter a number :")
        Console.ResetColor()
        Console.Write(" ")

        ' Checks Which option to choose
        If Integer.TryParse(Console.ReadLine(), input2) Then
            Dim fh As Long, line As String
            'Gets the directory of the file
            Dim file_name As String = "AMD.csv"


            'Opens the file
            fh = FreeFile()
            FileOpen(fh, file_name, OpenMode.Input)

            Dim counter As Integer = 0

            'EOF to avoid the error generated by attempting to get input past the end of a file.
            While Not EOF(fh)

                line = LineInput(fh)
                Dim arr_File() As String = Split(line, ",")
                'Storing all the rows of the dates
                arr_TrDates(counter) = arr_File(0)
                'Storing all the rows of Open Prices
                arr_OpenPrices(counter) = arr_File(1)
                'Storing all the rows of High prices
                arr_HighPrices(counter) = arr_File(2)
                'Storing all the rows of Low pRice
                arr_LowPrices(counter) = arr_File(3)
                'Storing all the rows of Close Price
                arr_ClosePrices(counter) = arr_File(4)
                'Storing all the rows of Volume
                arr_Volume(counter) = arr_File(5)



                ' Runs 252 times to check each lines, which will also be used to count how many days there are
                counter += 1
            End While
            'after reading it will close down
            FileClose(fh)

            Select Case input2
                Case 1

                    Dim lowest As Double = 10000.0
                    Dim highest As Double = 0
                    Dim highest_Vol As Integer = 0
                    Dim index_highestTradingDate As Double
                    Dim percChange As Double

                    Console.WriteLine()

                    ' Loops around the 252 Lines
                    For i = 0 To arr_TrDates.Length - 1

                        ' We need to check if the lowerprices is less than lowest if it is, we store in the lowest

                        ' To do this we declare lowest outside of the while loop, because
                        ' then we have to call the lowest outside to get the final answer
                        ' this way it doesn't print 252 times and all the value but just one
                        ' we set lowest as 10000.0 which is the highest value that they can store if it is
                        ' whichever is the lowest in lowprices it will store in the lowest.
                        If arr_LowPrices(i) < lowest Then
                            lowest = arr_LowPrices(i)

                        End If


                        'Similarly here we do the opposite, we set highest as 0
                        ' and we say whichever number is higher in arr_HighPrices will be the highest be stored as the highest.
                        If arr_HighPrices(i) > highest Then
                            highest = arr_HighPrices(i)

                        End If

                        'Now i have to find the trading volume date
                        ' I will make a new variable (highest_Vol) to find the highest volume within the arr_vol(index)
                        ' Check if the volume is higher than the (highest_Vol) if its it will store there

                        If arr_Volume(i) > highest_Vol Then
                            highest_Vol = arr_Volume(i)
                            ' Then we have to store the index of the volume in a new variable, which then we can call this index to find
                            ' the specific date with the highest volume we found,
                            'In this case we store it in index_highestTradingDate and call it in the consolwriteline arr_TrDates(index_highestTradingDate)
                            ' this goes through the arr_TrDates of the dates and find the index of the counter.
                            index_highestTradingDate = i

                        End If


                        ' Runs 252 times to check each lines, which will also be used to count how many days there are
                    Next

                    'Outputs data for each Trades.
                    Console.WriteLine("────────────")
                    Console.Write(“Total number of trading days: ")
                    Console.ForegroundColor = ConsoleColor.Green
                    Console.Write(counter)
                    Console.ResetColor()
                    Console.WriteLine()

                    Console.Write(“Opening price of the first trading day: ")
                    Console.ForegroundColor = ConsoleColor.Green
                    Console.Write(arr_OpenPrices(0))
                    Console.ResetColor()
                    Console.WriteLine()

                    Console.Write(“Closing price of the last trading day: ")
                    Console.ForegroundColor = ConsoleColor.Green
                    Console.Write(arr_ClosePrices.Last)
                    Console.ResetColor()
                    Console.WriteLine()


                    Console.Write(“Highest trading price of year:")
                    Console.ForegroundColor = ConsoleColor.Green
                    Console.Write(highest)
                    Console.ResetColor()
                    Console.WriteLine()

                    Console.Write(“Lowest trading price of year: ")
                    Console.ForegroundColor = ConsoleColor.Green
                    Console.Write(lowest)
                    Console.ResetColor()
                    Console.WriteLine()

                    Console.Write(“Date with the highest trading volume: ")
                    Console.ForegroundColor = ConsoleColor.Green
                    Console.Write(arr_TrDates(index_highestTradingDate))
                    Console.ResetColor()
                    Console.WriteLine()

                    ' Changing in percentage
                    'Now we have to find the percentage change from the highest volume date index for the closing prices
                    'So to do this, I will just get the index of the highest volume trading date and plug in the array of closeprices to find the close price for that day
                    ' and then to find the previous closingprices i can just - 1 index to find the previous close
                    ' and then we just do price changem to find the price dif, which means
                    ' ((trading date - previoustrading date) / previous trading date) * 100 to find the percentage
                    ' we also have to round up to 2 decimal places so we use math.round to 2 decimal places
                    ' And we store it in percChange 
                    percChange = (arr_ClosePrices(index_highestTradingDate) - arr_ClosePrices(index_highestTradingDate - 1)) / arr_ClosePrices(index_highestTradingDate - 1) * 100

                    Console.Write(“Price change from last closing price: ")
                    Console.ForegroundColor = ConsoleColor.Green
                    Console.Write(Str(Math.Round(percChange, 2)) & "%")
                    Console.ResetColor()
                    Console.WriteLine()


                    ' Once information give, ask the  user to go back to stock analysis
                    Console.WriteLine("────────────")
                    Console.Write("Would you like to go back to Stock Analysis? ")
                    Console.ForegroundColor = ConsoleColor.Red
                    Console.Write("(Y/N)")
                    Console.ResetColor()
                    Console.WriteLine()

                    KeyIput()

                Case 2
                    'Will show the Monthly data analysis
                    monthlyData()

                Case 3
                    'Exit'
                    Console.Clear()
                    menu()

            End Select
        Else
            Console.Clear()
            Console.WriteLine("---------------------------------------")
            Console.Write("/* Input a number between 1 to 3: *\ ")
            Console.WriteLine()
            Console.WriteLine("---------------------------------------")
            Threading.Thread.Sleep(600)
            Console.Clear()
            AppendixB()
        End If
    End Sub

    Sub monthlyData()

        'Display the option given to user
        Console.WriteLine("Jan, Feb, Mar, Apr, May, Jun, Jul, Aug, Sept, Oct, Nov, or Dec")
        Console.WriteLine("────────────")
        Console.BackgroundColor = ConsoleColor.DarkRed
        Console.Write("Type a month name in lowercase ")
        Console.ResetColor()
        Console.Write(":")



        Dim month As String = Console.ReadLine()
        month = month.ToLower()
        Dim monthNumber As String = ""
        Select Case month
            Case "jan"
                'Stores the string for the month, which will then be used to check the months
                monthNumber = "01"
            Case "feb"

                monthNumber = "02"

            Case "mar"

                monthNumber = "03"

            Case "apr"

                monthNumber = "04"

            Case "may"

                monthNumber = "05"

            Case "jun"

                monthNumber = "06"

            Case "jul"

                monthNumber = "07"

            Case "aug"

                monthNumber = "08"

            Case "sept"

                monthNumber = "09"

            Case "oct"

                monthNumber = "10"

            Case "nov"

                monthNumber = "11"

            Case "dec"

                monthNumber = "12"

            Case Else

                Console.Write("Type ERROR. Try Again : ")

                monthlyData()
                'Prevent futher exceution
                Exit Sub
        End Select
        Dim OpeningPrice As Double = 0
        Dim ClosingPrice As Double = 0
        Dim HighestTradingPrice As Double = 0
        ' Sets the doubel highest value that can be taken which is compared in the if statement
        Dim LowestTradingPrice As Double = Double.MaxValue
        Dim TotalTradingVolume As Double
        'Check if month was selected
        If monthNumber <> "" Then
            'Loop through array
            For i = 0 To arr_TrDates.Length - 1
                'Get month from date
                ' EG : 01/02/2005 it, removes the / and gets only 01  02  2005
                Dim arr_Date() As String = arr_TrDates(i).Split("/")
                'Check if the month is equal to the selected month
                ' The position of the months is 1, since we start from 0, which is the days stored, 1 is the month
                ' We check if the input that person types is the same as the array dates
                ' For example if someone types January, we store monthNumber = 01, which then find the arr_Date(1) string value of 01 
                If arr_Date(1) = monthNumber Then
                    'Set Opening Price once
                    If OpeningPrice = 0 Then
                        OpeningPrice = arr_OpenPrices(i)
                    End If
                    'Set LowestTradingPrice if initial value is higher than new value
                    If LowestTradingPrice > arr_LowPrices(i) Then
                        LowestTradingPrice = arr_LowPrices(i)
                    End If
                    'Set HighestTradingPrice if initial value is smaller than new value
                    If HighestTradingPrice < arr_HighPrices(i) Then
                        HighestTradingPrice = arr_HighPrices(i)
                    End If
                    'Set Opening Price always to make it take the value of the last one
                    ClosingPrice = arr_ClosePrices(i)
                    TotalTradingVolume += arr_Volume(i)
                End If
            Next

            'Display results
            Console.WriteLine("────────────")
            Console.Write("Opening price: ")
            Console.ForegroundColor = ConsoleColor.Green
            Console.Write(OpeningPrice)
            Console.ResetColor()
            Console.WriteLine()

            Console.Write("Closing price: ")
            Console.ForegroundColor = ConsoleColor.Green
            Console.Write(ClosingPrice)
            Console.ResetColor()
            Console.WriteLine()


            Console.Write("Highest trading price: ")
            Console.ForegroundColor = ConsoleColor.Green
            Console.Write(HighestTradingPrice)
            Console.ResetColor()
            Console.WriteLine()


            Console.Write("Lowest trading price: ")
            Console.ForegroundColor = ConsoleColor.Green
            Console.Write(LowestTradingPrice)
            Console.ResetColor()
            Console.WriteLine()


            Console.Write("Total trading volume: ")
            Console.ForegroundColor = ConsoleColor.Green
            Console.Write(TotalTradingVolume)
            Console.ResetColor()
            Console.WriteLine()


            Console.WriteLine("────────────")
            Console.Write("Would you like to go back to Stock Analysis?")
            Console.ForegroundColor = ConsoleColor.DarkRed
            Console.Write(" (Y/N)")
            Console.ResetColor()
            KeyIput()

        End If


    End Sub





    Sub Terminating()

        ' Display if the user wants to quit
        Console.WriteLine("────────────")
        Console.Write("Do you want to QUIT?")
        Console.ForegroundColor = ConsoleColor.Red
        Console.Write(" (Y/N)")
        Console.ResetColor()
        Console.WriteLine()


        ' stores the key input
        Dim usr_key As String
        usr_key = Console.ReadKey.Key

        ' Checks if the user input is either Y or N, and then takes and action
        ' If yes, it will close the interface
        ' If no, it will clear the console and restart menu
        ' Else, it will loop the function

        If usr_key = ConsoleKey.Y Then
            Console.Clear()

            Console.WriteLine("════════════════════════════ - - - - - - - - >")
            Console.Write("Terminating : ")

            End
            Console.ReadKey(True)



        ElseIf usr_key = ConsoleKey.N Then
            Console.Clear()
            menu()
            Console.ReadKey(True)

        Else
            ' Else it will loop again 
            ' To do this, we going to return the function

            Console.Write(" ──» is a wrong key. Try Again. ")

            Terminating()

        End If

    End Sub

    Sub KeyIput()
        ' If user key press will let go back to stock analysis 
        ' if not it will quit
        Dim usr_key As String
        usr_key = Console.ReadKey.Key

        If usr_key = ConsoleKey.Y Then
            Console.Clear()
            AppendixB()
            Console.ReadKey(False)
        ElseIf usr_key = ConsoleKey.N Then
            Console.Clear()
            Console.WriteLine("════════════════════════════ - - - - ->")
            Console.Write("Terminating : ")
            End
            Console.ReadKey(True)
        Else
            Console.WriteLine()
            Console.WriteLine("Error, it will Quit in 3 second")
            Threading.Thread.Sleep(3)
            Console.Clear()
            Console.WriteLine("════════════════════════════ - - - - - - - - >")
            Console.Write("Terminating.... ")

        End If

    End Sub

End Module

