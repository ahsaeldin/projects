Imports CuteMacro.Balora
Imports System.Runtime.InteropServices


'//My internal process communications class

Public Class IntProcCom

    '//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    Public Function SendStringToAprocess(ByVal text As String) As Integer
        '//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

        Try

            Dim CMHwnd As IntPtr = FindWindow(vbNullString, "Perfect Macro Recorder")
            'get string from textbox
            Dim str As String = text
            'string to byte array
            Dim B() As Byte = System.Text.Encoding.Default.GetBytes(str)
            'allocate memory space for byte array, and get a pointer to it
            Dim lpB As IntPtr = Marshal.AllocHGlobal(B.Length)
            'copy the byte array into memory  
            Marshal.Copy(B, 0, lpB, B.Length)
            'setup a standard structure for the WM_COPYDATA message
            Dim CD As COPYDATASTRUCT
            With CD
                .dwData = 0 'can be used for custom indexing between apps
                .cbData = B.Length 'length of data
                .lpData = lpB.ToInt32 'pointer to the data
            End With
            'clean up array
            Erase B
            'allocate memory space for structure, and get a pointer to it
            Dim lpCD As IntPtr = Marshal.AllocHGlobal(CD.cbData)
            'copy structure to allocated memory place
            Marshal.StructureToPtr(CD, lpCD, False)

            'send message
            'SendMessage(CMHwnd, WM_COPYDATA, 0, lpCD.ToInt32)
            Dim Res As Integer
            Dim SendingResult As Integer
            Res = SendMessageTimeout(CMHwnd, WM_COPYDATA, 0, lpCD.ToInt32, SMTO_ABORTIFHUNG, 3000, SendingResult)

            Return Res

        Catch ex As Exception

            Util.WriteToErrorLog(ex.Message, ex.StackTrace, "Error Sending Path to the other instance")
            Return 0

        End Try

    End Function

    '//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    Public Shared Sub CopyDataToClipboard(ByVal format As String, ByVal str As String)
        '//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

        'الوظيفة دي ووظيفة
        'PasteDataFromClipboard
        'عشان النسخ واللصق من وإلي الكيب بورد بتنسيقات معينة
        'مما يعنني أولا علي ارسال رسائل بين حالات البرنامج المختلفة
        'لتفادي استخدام
        'WM_COPYDATA 
        'في وظيفة
        'MyApplication_Startup
        'التي تسخدم في إرسال مسار ملف السي ماك إلي النسخة التي تعمل
        'فعلياً من البرنامج

        ' Creates a new data format.
        Dim CDataFormat As DataFormats.Format = DataFormats.GetFormat(format)

        ' Creates a new object and store it in a DataObject using myFormat 
        ' as the type of format. 
        Dim CopiedText As String = str
        Dim DataObject As New DataObject(CDataFormat.Name, CopiedText)

        'لازم تاني بارامتر يبقي ترو عشان تقدر تنقل محتويات الكليب بوورد بين برنامجين
        Clipboard.SetDataObject(DataObject, True)

    End Sub

    Public Shared Function PasteDataFromClipboard(ByVal format As String) As String

        'الوظيفة دي ووظيفة
        'PasteDataFromClipboard
        'عشان النسخ واللصق من وإلي الكيب بورد بتنسيقات معينة
        'مما يعنني أولا علي ارسال رسائل بين حالات البرنامج المختلفة
        'لتفادي استخدام
        'WM_COPYDATA 
        'في وظيفة
        'MyApplication_Startup
        'التي تسخدم في إرسال مسار ملف السي ماك إلي النسخة التي تعمل
        'فعلياً من البرنامج

        ' Creates a new data format.
        Dim CMFormat As DataFormats.Format = DataFormats.GetFormat(format)

        ' Performs some processing steps.
        ' Retrieves the data from the clipboard.
        Dim RetrievedItemsObject As IDataObject = Clipboard.GetDataObject()

        ' Converts the IDataObject type to MyNewObject type. 
        Dim CopiedItems As String = CType(RetrievedItemsObject.GetData(CMFormat.Name), String)

        PasteDataFromClipboard = CopiedItems

    End Function
End Class
