Option Strict On
Option Explicit On

'Taylor Herndon
'RCET0265
'Spring 2021
'Accumulate Message Function
'

Module AccumulateMessages

    Sub Main()

        Dim ReturnString As String = Nothing
        Dim ReturnNumber As Integer = Nothing

        'Test A, store results in message storage
        Console.WriteLine("[Test A]: Please press a letter key.")
        ReturnString = Actions.ConsoleKeyLetterToString()

        If ReturnString = Nothing Then

            AccumulateMessages.StoreMessage("[Test A Result]:Error Letter key not found.", False)

        Else

            AccumulateMessages.StoreMessage("[Test A Result]: The key you pressed was " & ReturnString & ".", False)

        End If
        ReturnString = Nothing

        'Test B, store result in message storage
        Console.Clear()
        Console.WriteLine("[Test B]: Please press a number key.")
        ReturnNumber = Actions.ConsoleKeyToNumber()

        If ReturnNumber = Nothing Then

            AccumulateMessages.StoreMessage("[Test B Result]: Error - Number key not found.", False)

        Else

            AccumulateMessages.StoreMessage("[Test B Result]: The number you pressed was " & Convert.ToString(ReturnNumber) & ".", False)

        End If
        ReturnNumber = Nothing

        'Test C, store result in message storage
        Console.Clear()
        Console.WriteLine("[Test C]: Please press any letter or number key.")
        ReturnString = Actions.ConsoleKeyToString()

        If ReturnString = Nothing Then

            AccumulateMessages.StoreMessage("[Test C Result]: Error - Key not found.", False)

        Else

            AccumulateMessages.StoreMessage("[Test C Result]: The key you pressed was " & ReturnString & ".", False)

        End If
        ReturnString = Nothing

        Actions.LoadingBar()

        'Write out all stored messages
        Console.WriteLine(AccumulateMessages.StoreMessage("", False))

        'Clear all messages and write all messages to confirm the clear function works.
        AccumulateMessages.StoreMessage("", True)
        Console.WriteLine(AccumulateMessages.StoreMessage("", False))

        Console.ReadKey()

    End Sub

    Function StoreMessage(Message As String, Clear As Boolean) As String

        Static _MessageStorage As String

        If Clear = True Then

            'Reset string if clear is true
            _MessageStorage = ""

        Else

            If Message = "" Then

                'If no message is given do not modify the string.

            Else

                'Add new message to string
                Message = Message & vbNewLine & vbNewLine
                _MessageStorage = _MessageStorage & Message

            End If

        End If

        Return _MessageStorage

    End Function

End Module
