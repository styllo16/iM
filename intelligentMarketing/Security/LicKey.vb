Imports LicenseKeyDemo.Encryption
Imports System.Runtime.InteropServices

Public Class LicKey

    Public Const BaseDate As Date = #1/1/2014#

    Public Structure SerialInf
        Dim ProdID As Integer
        Dim Serial As Integer
        Dim OptVal As UInt16
    End Structure

    Dim Serial As SerialInf
    Dim SecretKey As String
    Dim SecretSalt As String

    Public Sub New()
        Serial = New SerialInf
        Serial.Serial = 0
        Serial.ProdID = 0
        Serial.OptVal = 0
        SecretSalt = ""
        SecretKey = GenKey(SecretSalt & Application.ProductName, 256)
    End Sub

    Property SerialNo() As Integer
        Get
            Return Serial.Serial
        End Get
        Set(ByVal value As Integer)
            Serial.Serial = value
        End Set
    End Property

    Property ProductID() As Integer
        Get
            Return Serial.ProdID
        End Get
        Set(ByVal value As Integer)
            Serial.ProdID = value
        End Set
    End Property

    Property OptValue() As UInt16
        Get
            Return Serial.OptVal
        End Get
        Set(ByVal value As UInt16)
            Serial.OptVal = value
        End Set
    End Property

    Property KeyCode() As String
        Get
            Return CryptSerial()
        End Get
        Set(ByVal value As String)
            Serial = DecryptSerial(value)
        End Set
    End Property

    Property Salt() As String
        Get
            Return (SecretSalt)
        End Get
        Set(ByVal value As String)
            SecretSalt = value
            SecretKey = GenKey(SecretSalt & Application.ProductName, 256)
        End Set
    End Property

    Property ExpDate() As Date
        Get
            Return DateAdd(DateInterval.Day, Serial.OptVal, BaseDate)
        End Get
        Set(ByVal value As Date)
            Dim Days As Integer = DateDiff(DateInterval.Day, BaseDate, value)

            If Days <= 65535 Then
                Serial.OptVal = Days
            Else
                Dim A As New ArgumentException("Date out of range")
                Throw A
            End If
        End Set
    End Property

    Public Function OptionEnabled(ByVal OptNo As Integer) As Boolean

        If OptNo > 0 And OptNo <= 16 Then
            Dim Opt As UInt16 = (Serial.OptVal >> OptNo - 1) And 1

            If Opt > 0 Then
                Return True
            Else
                Return (False)
            End If
        End If
    End Function

    Public Sub SetOption(ByVal OptNo As Integer)
        If OptNo > 0 And OptNo <= 16 Then
            Dim OptBit As UInt16 = 1 << (OptNo - 1)
            Serial.OptVal = Serial.OptVal Or OptBit
        End If

    End Sub

    Public Sub UnsetOption(ByVal OptNo As Integer)
        If OptNo > 0 And OptNo <= 16 Then
            Dim OptBit As UInt16 = 1 << (OptNo - 1)
            OptBit = Not (OptBit)
            Serial.OptVal = Serial.OptVal And OptBit
        End If
    End Sub

    Private Function CryptSerial() As String
        Try

            Dim BASerial As Byte() = SerialToBA(Serial)

            Dim Sym As New Encryption.Symmetric(Encryption.Symmetric.Provider.Rijndael)
            Dim Key As New Encryption.Data(SecretKey)
            Dim CryptSer As Encryption.Data = Sym.Encrypt(New Encryption.Data(BASerial), Key)

            Dim B32Str As String = CryptSer.ToBase32

            Return B32Str

        Catch ex As Exception
            Return ""
        End Try

    End Function

    Private Function DecryptSerial(ByVal KeyCode As String) As SerialInf

        Dim Ret As New SerialInf
        Ret.OptVal = 0
        Ret.ProdID = 0
        Ret.Serial = 0

        Dim SettingsBA As Byte()

        Try
            Dim sym As New Encryption.Symmetric(Encryption.Symmetric.Provider.Rijndael)
            Dim CryptDat As New Encryption.Data
            CryptDat.Base32 = KeyCode

            Debug.Print(CryptDat.Hex)

            Dim key As New Encryption.Data(SecretKey)
            Dim DecryptedData As Encryption.Data = sym.Decrypt(CryptDat, key)

            ' Get the raw decrypted bytes
            SettingsBA = DecryptedData.Bytes

            Return BAToSerial(SettingsBA)

        Catch ex As Exception
            Return Ret
        End Try

    End Function

    Private Function GenKey(ByVal InVal As String, ByVal Bits As Integer) As String
        ' Converts any arbritrary string to a 
        ' length suitable for use as a crypto key
        Dim Ret As String = ""
        Dim InLen As Integer = InVal.Length
        Dim ByteCount As Integer = Bits / 8

        If InLen >= ByteCount Then
            Ret = Left(InVal, ByteCount)
        Else
            Ret = InVal
            Dim Cur As Integer = 0
            Dim Rand As Integer = 64
            For Cur = 1 To ByteCount - InLen
                Rand = Rand + 1
                If Rand > Asc("z") Then Rand = 65
                Dim PadChar As String = Chr(Rand)
                Ret = Ret & PadChar.Trim
            Next
        End If

        Return Ret
    End Function

    Private Function SerialToBA(ByVal SI As SerialInf) As Byte()
        Dim BA As Byte()
        Dim Ptr As IntPtr = Marshal.AllocHGlobal(Marshal.SizeOf(SI))
        ReDim BA(Marshal.SizeOf(SI) - 1)
        Marshal.StructureToPtr(SI, Ptr, False)
        Marshal.Copy(Ptr, BA, 0, Marshal.SizeOf(SI))
        Marshal.FreeHGlobal(Ptr)
        Return BA
    End Function

    Private Function BAToSerial(ByVal BA As Byte()) As SerialInf
        Dim SI As SerialInf
        Dim GC As GCHandle = GCHandle.Alloc(BA, GCHandleType.Pinned)
        Dim Obj As Object = Marshal.PtrToStructure(GC.AddrOfPinnedObject, SI.GetType)
        GC.Free()
        Return Obj
    End Function
End Class
