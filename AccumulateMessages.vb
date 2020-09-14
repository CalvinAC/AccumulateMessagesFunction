'Calvin Coxen
'RCET0265
'Fall 2020
'Accumulate Messages Function
'https://github.com/CalvinAC/AccumulateMessagesFunction/blob/master/AccumulateMessages.vb

Option Explicit On
Option Strict On
Option Compare Text




Module AccumulateMessages

    Sub Main()
        'This code prompts the user to enter a message and how to call or clear it,
        'it also creates the variables for the users message and to clear them
        Dim userInput As String
        Dim message As String
        Dim clearMemory As Boolean
        Console.WriteLine("Type in a message and the program will save it. 
To see the written messages displayed, type 'call' 
To clear all messages, type 'clear'
To stop the program, type 'q' ")

        'This while loop allows the user to see all the messages they wrote,
        'clear them from memory, or exit the program. 
        Do

            userInput = Console.ReadLine()
            If userInput = "call" Then
                MsgBox(message)
            ElseIf userInput = "clear" Then
                clearMemory = True
            ElseIf userInput = "q" Then
                Exit Sub
            End If

            message = AccumulateMessageFunction(userInput, clearMemory)

            clearMemory = False

        Loop
    End Sub

    'This function stores each message from the user and is used to recall 
    'every message written or clears eveyrthing when prompted to do so
    Function AccumulateMessageFunction(ByVal newMessage As String, ByVal clear As Boolean) As String
        Static savedMessage As String

        If newMessage = "call" Then
            'This is left blank so the message box doesnt display the message "call"
        ElseIf clear Then
            savedMessage = " "
        Else savedMessage &= newMessage & vbNewLine

        End If

        Return savedMessage
    End Function

End Module
