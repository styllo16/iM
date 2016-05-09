Public Class KeyDec
    Private encryptedstr As String
    Private strUsername As String
    Private intUsers As Integer
    Private ExpDate As Date


    Public Property UserName As String
        Get
            Return strUsername
        End Get

        Set(ByVal value As String)

        End Set
    End Property
    Public Property Users As Integer
        Get
            Return intUsers
        End Get

        Set(ByVal value As Integer)

        End Set
    End Property

    Public Property ExpiryDate As Date
        Get
            Return ExpDate
        End Get

        Set(ByVal value As Date)

        End Set
    End Property
    Public Property isValidKey As Boolean
        Get
            If My.Settings.SMS_Account_Name = strUsername Then

                Return True
            Else
                Return False

            End If


        End Get

        Set(ByVal value As Boolean)

        End Set
    End Property
    Public Property isExpired As Boolean
        Get
            If ExpDate < Date.Today Then

                Return True
            Else
                Return False

            End If


        End Get

        Set(ByVal value As Boolean)

        End Set
    End Property
    Public Sub Decode(ByVal encr As String)
        encryptedstr = Decrypt(encr)

        Dim pos_of_first_dot As Integer
        Dim pos_of_Second_dot As Integer
        Dim UserNamestr As String
        Dim UserCountstr As String
        Dim datestr As String



        pos_of_first_dot = InStr(encryptedstr, ".")
        pos_of_Second_dot = InStrRev(encryptedstr, ".")

        UserNamestr = Mid(encryptedstr, 1, pos_of_first_dot - 1)
        strUsername = UserNamestr

        UserCountstr = Mid(encryptedstr, pos_of_first_dot + 1, pos_of_Second_dot - pos_of_first_dot - 1)
        intUsers = CInt(UserCountstr)

        datestr = Mid(encryptedstr, pos_of_Second_dot + 1)
        ExpDate = Date.Parse(datestr)

    End Sub




End Class
