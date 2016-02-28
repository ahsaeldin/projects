'#If CONFIG = "Debug" Then
'Imports NUnit.Framework

'Namespace BaloraUnitTests
'    <TestFixture()> _
'    Public Class EncryptionTest

'        Private fileToBeTested As String = "E:\Rocknee\KV\Cacadu\Code\Trunk\Cacadu\bin\Debug\encriptionTests.txt"
'        Private _TargetString As String
'        Private _TargetData As Balora.Encryption.Data

'        <TestFixtureSetUp()> _
'        Public Sub SetupMethods()

'            Dim current As String = Util.GetCurrentExecutingDirectory
'            If Not current.Contains("Rocknee") Then
'                fileToBeTested = current & "encriptionTests.txt"
'            End If

'            _TargetString = "The instinct of nearly all societies is to lock up anybody who is truly free. " & _
'             "First, society begins by trying to beat you up. If this fails, they try to poison you. " & _
'             "If this fails too, they finish by loading honors on your head." & _
'             " - Jean Cocteau (1889-1963)"
'            _TargetData = New Encryption.Data(_TargetString)
'        End Sub


'        <Test(), Category("Hash")> _
'        Public Sub SaltedHashes()
'            Assert.AreEqual("6CD9DD96", _
'             SaltedHash(Encryption.Hash.Provider.CRC32, _
'              New Encryption.Data("Shazam!")).ToHex)
'            Assert.AreEqual("4F7FA9C182C5FA60F9197F4830296685", _
'             SaltedHash(Encryption.Hash.Provider.MD5, _
'              New Encryption.Data("SnapCracklePop")).ToHex)
'            Assert.AreEqual("3DC330B4E4E61C8DF039EAE93EC16412E22425FB", _
'             SaltedHash(Encryption.Hash.Provider.SHA1, _
'              New Encryption.Data("全球最大的華文新聞網站", System.Text.Encoding.Unicode)).ToHex)
'            Assert.AreEqual("EFAE307AEE511D6078FDF0D4372F4D0C8135170C5F7626CB19B04BFDBABBBDB2", _
'             SaltedHash(Encryption.Hash.Provider.SHA256, _
'              New Encryption.Data("!@#$%^&*()_-+=", System.Text.Encoding.ASCII)).ToHex)
'            Assert.AreEqual( _
'             "582B31C13EF16D706EC2514FDA08316A369DF1F130D34A0A2A16B065D82662A1101EA01110AB7C8F9022A1CEA76FD6B9", _
'             SaltedHash(Encryption.Hash.Provider.SHA384, _
'              New Encryption.Data("supercalifragilisticexpialidocious", System.Text.Encoding.ASCII)).ToHex)
'            Assert.AreEqual( _
'             "44FAA06E8E80666408304E3458621769699A76B591C6389F958C0DDA1D80A82965D169E8AA7D3C1A0637BCB7B0F45D420389C629D19E255D64A923F6C4F87FD8", _
'             SaltedHash(Encryption.Hash.Provider.SHA512, _
'              New Encryption.Data("42", System.Text.Encoding.ASCII)).ToHex)
'        End Sub

'        <Test(), Category("Hash")> _
'        Public Sub HashFile()
'            Dim h1 As New Encryption.Hash(Encryption.Hash.Provider.CRC32)
'            Dim hashHex As String
'            Dim sr As IO.StreamReader = Nothing
'            Using sr
'                sr = New IO.StreamReader(fileToBeTested)
'                hashHex = h1.Calculate(sr.BaseStream).ToHex
'            End Using
'            Assert.AreEqual(hashHex, "83D1593C")
'            Dim h2 As New Encryption.Hash(Encryption.Hash.Provider.MD5)
'            Try
'                sr = New IO.StreamReader(fileToBeTested)
'                hashHex = h2.Calculate(sr.BaseStream).ToHex
'            Finally
'                If Not sr Is Nothing Then sr.Close()
'            End Try
'            Assert.AreEqual(hashHex, "BA270983312091197E7F41DF2263EF32")
'        End Sub

'        <Test(), Category("Hash")> _
'        Public Sub Hash()
'            Assert.AreEqual("AA692113", _
'             Hash(Encryption.Hash.Provider.CRC32).ToHex)
'            Assert.AreEqual("44D36517B0CCE797FF57118ABE264FD9", _
'             Hash(Encryption.Hash.Provider.MD5).ToHex)
'            Assert.AreEqual("9E93AB42BCC8F738C7FBB6CCA27A902DC663DBE1", _
'             Hash(Encryption.Hash.Provider.SHA1).ToHex)
'            Assert.AreEqual("40AF07ABFE970590B2C313619983651B1E7B2F8C2D855C6FD4266DAFD7A5E670", _
'             Hash(Encryption.Hash.Provider.SHA256).ToHex)
'            Assert.AreEqual("9FC0AFB3DA61201937C95B133AB397FE62C329D6061A8768DA2B9D09923F07624869D01CD76826E1152DAB7BFAA30915", _
'             Hash(Encryption.Hash.Provider.SHA384).ToHex)
'            Assert.AreEqual("2E7D4B051DD528F3E9339E0927930007426F4968B5A4A08349472784272F17DA5C532EDCFFE14934988503F77DEF4AB58EB05394838C825632D04A10F42A753B", _
'             Hash(Encryption.Hash.Provider.SHA512).ToHex)
'        End Sub

'        Private Function SaltedHash(ByVal p As Encryption.Hash.Provider, ByVal salt As Encryption.Data) As Encryption.Data
'            Dim h As New Encryption.Hash(p)
'            Return h.Calculate(New Encryption.Data(_TargetString), salt)
'        End Function

'        Private Function Hash(ByVal p As Encryption.Hash.Provider) As Encryption.Data
'            Dim h As New Encryption.Hash(p)
'            Return h.Calculate(New Encryption.Data(_TargetString))
'        End Function

'        <Test(), Category("Asymmetric")> _
'        Public Sub Asymmetric()
'            'You need to add the app settings from here
'            'http://www.codeproject.com/KB/security/SimpleEncryption.aspx
'            Console.WriteLine("You must add the app settings in http://www.codeproject.com/KB/security/SimpleEncryption.aspx")
'            Dim secret As String = "Pack my box with five dozen liquor jugs."
'            Assert.AreEqual(secret, AsymmetricNewKey(secret))
'            Assert.AreEqual(secret, AsymmetricNewKey(secret, 384))
'            Assert.AreEqual(secret, AsymmetricNewKey(secret, 512))
'            Assert.AreEqual(secret, AsymmetricNewKey(secret, 1024))
'            Assert.AreEqual(secret, AsymmetricConfigKey(secret))
'            Assert.AreEqual(secret, AsymmetricXmlKey(secret))
'        End Sub

'        Private Function AsymmetricXmlKey(ByVal secret As String) As String
'            Dim publicKeyXml As String = _
'             "<RSAKeyValue>" & _
'             "<Modulus>0D59Km2Eo9oopcm7Y2wOXx0TRRXQFybl9HHe/ve47Qcf2EoKbs9nkuMmhCJlJzrq6ZJzgQSEbpVyaWn8OHq0I50rQ13dJsALEquhlfwVWw6Hit7qRvveKlOAGfj8xdkaXJLYS1tA06tKHfYxgt6ysMBZd0DIedYoE1fe3VlLZyE=</Modulus>" & _
'             "<Exponent>AQAB</Exponent>" & _
'             "</RSAKeyValue>"

'            Dim privateKeyXml As String = _
'             "<RSAKeyValue>" & _
'             "<Modulus>0D59Km2Eo9oopcm7Y2wOXx0TRRXQFybl9HHe/ve47Qcf2EoKbs9nkuMmhCJlJzrq6ZJzgQSEbpVyaWn8OHq0I50rQ13dJsALEquhlfwVWw6Hit7qRvveKlOAGfj8xdkaXJLYS1tA06tKHfYxgt6ysMBZd0DIedYoE1fe3VlLZyE=</Modulus>" & _
'             "<Exponent>AQAB</Exponent>" & _
'             "<P>/1cvDks8qlF1IXKNwcXW8tjTlhjidjGtbT9k7FCYug+P6ZBDfqhUqfvjgLFF/+dAkoofNqliv89b8DRy4gS4qQ==</P>" & _
'             "<Q>0Mgq7lyvmVPR1r197wnba1bWbJt8W2Ki8ilUN6lX6Lkk04ds9y3A0txy0ESya7dyg9NLscfU3NQMH8RRVnJtuQ==</Q>" & _
'             "<DP>+uwfRumyxSDlfSgInFqh/+YKD5+GtGXfKtO4hu4xF+8BGqJ1YXtkL+Njz2zmADOt5hOr1tigPSQ2EhhIqUnAeQ==</DP>" & _
'             "<DQ>M5Ofd28SOjCIwCHjwG+Q8v1qzz3CBNljI6uuEGoXO3ixbkggVRfKcMzg2C6AXTfeZE6Ifoy9OyhvLlHTPiXakQ==</DQ>" & _
'             "<InverseQ>yQIJMLdL6kU4VK7M5b5PqWS8XzkgxfnaowRs9mhSEDdwwWPtUXO8aQ9G3zuiDUqNq9j5jkdt77+c2stBdV97ew==</InverseQ>" & _
'             "<D>HOpQXu/OFyJXuo2EY43BgRt8bX9V4aEZFRQqrqSfHOp8VYASasiJzS+VTYupGAVqUPxw5V1HNkOyG0kIKJ+BG6BpIwLIbVKQn/ROs7E3/vBdg2+QXKhikMz/4gYx2oEsXW2kzN1GMRop2lrrJZJNGE/eG6i4lQ1/inj1Tk/sqQE=</D>" & _
'             "</RSAKeyValue>"

'            Dim encryptedData As Encryption.Data
'            Dim decryptedData As Encryption.Data
'            Dim asym As New Encryption.Asymmetric
'            Dim asym2 As New Encryption.Asymmetric

'            encryptedData = asym.Encrypt(New Encryption.Data(secret), publicKeyXml)
'            decryptedData = asym2.Decrypt(encryptedData, privateKeyXml)

'            Return decryptedData.ToString
'        End Function

'        Private Function AsymmetricConfigKey(ByVal secret As String) As String
'            Dim encryptedData As Encryption.Data
'            Dim decryptedData As Encryption.Data
'            Dim asym As New Encryption.Asymmetric
'            Dim asym2 As New Encryption.Asymmetric

'            encryptedData = asym.Encrypt(New Encryption.Data(secret))
'            decryptedData = asym2.Decrypt(encryptedData)

'            Return decryptedData.ToString
'        End Function

'        Private Function AsymmetricNewKey(ByVal secret As String, Optional ByVal keysize As Integer = 0) As String
'            Dim pubkey As New Encryption.Asymmetric.PublicKey
'            Dim privkey As New Encryption.Asymmetric.PrivateKey
'            Dim encryptedData As Encryption.Data
'            Dim decryptedData As Encryption.Data
'            Dim asym As New Encryption.Asymmetric
'            Dim asym2 As New Encryption.Asymmetric

'            If keysize = 0 Then
'                asym = New Encryption.Asymmetric
'                asym2 = New Encryption.Asymmetric
'            Else
'                asym = New Encryption.Asymmetric(keysize)
'                asym2 = New Encryption.Asymmetric(keysize)
'            End If
'            asym.GenerateNewKeyset(pubkey, privkey)

'            encryptedData = asym.Encrypt(New Encryption.Data(secret), pubkey)
'            decryptedData = asym2.Decrypt(encryptedData, privkey)

'            Return decryptedData.ToString
'        End Function

'        <Test(), Category("Symmetric")> _
'        Public Sub Symmetric()
'            Assert.AreEqual(_TargetString, SymmetricLoopback(Encryption.Symmetric.Provider.DES))
'            Assert.AreEqual(_TargetString, SymmetricWithKey(Encryption.Symmetric.Provider.DES))
'            Assert.AreEqual(_TargetString, SymmetricLoopback(Encryption.Symmetric.Provider.RC2))
'            Assert.AreEqual(_TargetString, SymmetricWithKey(Encryption.Symmetric.Provider.RC2))
'            Assert.AreEqual(_TargetString, SymmetricLoopback(Encryption.Symmetric.Provider.Rijndael))
'            Assert.AreEqual(_TargetString, SymmetricWithKey(Encryption.Symmetric.Provider.Rijndael))
'            Assert.AreEqual(_TargetString, SymmetricLoopback(Encryption.Symmetric.Provider.TripleDES))
'            Assert.AreEqual(_TargetString, SymmetricWithKey(Encryption.Symmetric.Provider.TripleDES))
'        End Sub

'        <Test(), Category("Symmetric")> _
'        Public Sub SymmetricFile()
'            '-- compare the hash of the decrypted file to what it should be after encryption/decryption
'            '-- using pure file streams
'            Dim h As New Encryption.Hash(Encryption.Hash.Provider.MD5)
'            Dim result As String = SymmetricFilePrivate(Encryption.Symmetric.Provider.TripleDES, fileToBeTested, "Password, Yo!")
'            Assert.AreEqual("BA270983312091197E7F41DF2263EF32", result)
'            result = SymmetricFilePrivate(Encryption.Symmetric.Provider.RC2, fileToBeTested, "0nTheDownLow1")
'            Assert.AreEqual("BA270983312091197E7F41DF2263EF32", result)
'        End Sub

'        Private Function SymmetricFilePrivate(ByVal p As Encryption.Symmetric.Provider, _
'         ByVal fileName As String, ByVal key As String) As String

'            Dim EncryptedFilePath As String = "../" & IO.Path.GetFileNameWithoutExtension(fileName) & "-encrypted.txt"

'            '-- encrypt the file to memory
'            Dim sym As New Encryption.Symmetric(p)
'            sym.Key = New Encryption.Data(key)
'            Dim encryptedData As New Encryption.Data
'            Dim sr As IO.StreamReader = Nothing
'            Using sr
'                sr = New IO.StreamReader(fileName)
'                encryptedData = sym.Encrypt(sr.BaseStream)
'            End Using

'            '-- write encrypted data to a new binary file
'            Dim sw As New IO.StreamWriter(EncryptedFilePath)
'            Dim bw As New IO.BinaryWriter(sw.BaseStream)
'            bw.Write(encryptedData.Bytes)
'            bw.Close()

'            '-- decrypt this binary file
'            Dim decryptedData As Encryption.Data
'            Dim sym2 As New Encryption.Symmetric(p)
'            sym2.Key = New Encryption.Data(key)
'            Try
'                sr = New IO.StreamReader(EncryptedFilePath)
'                decryptedData = sym.Decrypt(sr.BaseStream)
'            Finally
'                If Not sr Is Nothing Then sr.Close()
'                IO.File.Delete(EncryptedFilePath)
'            End Try

'            '-- get the MD5 hash of the returned data
'            Dim h As New Encryption.Hash(Encryption.Hash.Provider.MD5)
'            Return h.Calculate(decryptedData).ToHex
'        End Function

'        ''' <summary>
'        ''' test using user-provided keys and init vectors
'        ''' </summary>
'        Private Function SymmetricWithKey(ByVal p As Encryption.Symmetric.Provider) As String
'            Dim sym As New Encryption.Symmetric(p, False)
'            Dim sym2 As New Encryption.Symmetric(p, False)
'            Dim encryptedData As Encryption.Data
'            Dim decryptedData As Encryption.Data

'            Dim keyData As New Encryption.Data("MySecretPassword")
'            Dim ivData As New Encryption.Data("MyInitializationVector")

'            sym.IntializationVector = ivData
'            encryptedData = sym.Encrypt(_TargetData, keyData)
'            sym2.IntializationVector = ivData
'            decryptedData = sym2.Decrypt(encryptedData, keyData)

'            '-- the data will be padded to the encryption blocksize, so we need to trim it back down.
'            Return decryptedData.ToString.Substring(0, _TargetData.Bytes.Length)
'        End Function

'        ''' <summary>
'        ''' test using auto-generated keys
'        ''' </summary>
'        Private Function SymmetricLoopback(ByVal p As Encryption.Symmetric.Provider) As String
'            Dim sym As New Encryption.Symmetric(p)
'            Dim sym2 As New Encryption.Symmetric(p)
'            Dim encryptedData As Encryption.Data
'            Dim decryptedData As Encryption.Data

'            encryptedData = sym.Encrypt(New Encryption.Data(_TargetString))
'            decryptedData = sym2.Decrypt(encryptedData, sym.Key)

'            '-- the data will be padded to the encryption blocksize, so we need to trim it back down.
'            Return decryptedData.ToString.Substring(0, _TargetData.Bytes.Length)
'        End Function

'    End Class
'End Namespace
'#End If
