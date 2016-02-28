VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim a As String
Dim FileData As String


Const path2 = "exemacro.exe"
Const path1 = "H:\Documents and Settings\AHMED.CPRIN\Desktop\hacked.jmf"
'Const path3 = "H:\Documents and Settings\AHMED.CPRIN\Desktop\hacked.exe"
Const path3 = "haceked.exe"

    Open path2 For Binary As #1
        
         FileData = Input(LOF(1), 1)
        
    Close #1


FileData = FileData & "cprinsosonono"
    
   ' Open path1 For Binary As #2
        
   '      a = Input(LOF(2), 2)
        
   ' Close #2

'FileData = FileData & a
    
    
     Open path3 For Binary As #3
        
         Put #3, , FileData
        
    Close #3


Dim kk

End Sub
