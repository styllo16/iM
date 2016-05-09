
Module Base32
    Const cBase32Alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ234567"
    Const cBase32Pad = "="

    Public Function ToBase32String(ByVal Data As Byte(), Optional ByVal IncludePadding As Boolean = True) As String
        Dim RetVal As String = ""
        Dim Segments As New List(Of Long)
        ' Divide the input data into 5 byte segments  
        Dim Index As Integer = 0

        While Index < Data.Length
            Dim CurrentSegment As Long = 0
            Dim SegmentSize As Integer = 0

            While Index < Data.Length And SegmentSize < 5
                CurrentSegment <<= 8
                CurrentSegment += Data(Index)
                Index += 1
                SegmentSize += 1
            End While

            ' If the size of the last segment was less than 5 bytes, pad with zeros  
            CurrentSegment <<= (8 * (5 - SegmentSize))
            Segments.Add(CurrentSegment)
        End While

        ' Convert each 5 byte segment into 8 character strings  

        For Each CurrentSegment As Long In Segments
            For x As Integer = 0 To 7
                RetVal &= cBase32Alphabet.Chars((CurrentSegment >> (7 - x) * 5) And &H1F)
            Next
        Next

        ' Correct the end of the string (where the input wasn't a multiple of 5 bytes)  

        Dim LastSegmentUsefulDataLength As Integer = Math.Ceiling((Data.Length Mod 5) * 8 / 5)
        RetVal = RetVal.Substring(0, RetVal.Length - (8 - LastSegmentUsefulDataLength))

        ' Add the padding characters   
        If IncludePadding Then
            RetVal &= New String(cBase32Pad, 8 - LastSegmentUsefulDataLength)
        End If

        Return RetVal
    End Function


    Public Function FromBase32String(ByVal Data As String) As Byte()
        Dim RetVal As New List(Of Byte)
        Dim Segments As New List(Of Long)
        Dim x As Integer

        ' Remove any trailing padding  
        Data = Data.TrimEnd(New Char() {cBase32Pad})

        ' Process the string 8 characters at a time  
        Dim Index As Integer = 0

        While Index < Data.Length
            Dim CurrentSegment As Long = 0
            Dim SegmentSize As Integer = 0

            While Index < Data.Length And SegmentSize < 8
                CurrentSegment <<= 5
                CurrentSegment = CurrentSegment Or cBase32Alphabet.IndexOf(Data.Chars(Index))
                Index += 1
                SegmentSize += 1
            End While

            ' If the size of the last segment was less than 40 bits, pad it  
            CurrentSegment <<= (5 * (8 - SegmentSize))
            Segments.Add(CurrentSegment)
        End While

        ' Break the 5 byte segments back down into individual bytes  

        For Each CurrentSegment As Long In Segments
            For x = 0 To 4
                RetVal.Add((CurrentSegment >> (4 - x) * 8) And &HFF)
            Next
        Next

        ' Remove any bytes of padding from the output  

        Dim BytesToRemove As Integer = 4
        'Dim BytesToRemove As Integer = 5 - (Math.Ceiling(Math.Ceiling(3 * 8 / 5) / 2))
        RetVal.RemoveRange(RetVal.Count - BytesToRemove, BytesToRemove)
        Return RetVal.ToArray()
    End Function

End Module
