I had used "chat.openai.com" to refine and debug the script

Following code was taken from chat.openai.com for formatting the j column to have the colors red for negatives and green for positives
 
For Each cell In ws.Range("J2:J" & LastRow)
                If cell.Value > 0 Then
                    cell.Interior.Color = RGB(0, 255, 0)
                ElseIf cell.Value < 0 Then
                    cell.Interior.Color = RGB(255, 0, 0)