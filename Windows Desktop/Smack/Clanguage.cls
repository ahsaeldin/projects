VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Clanguage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Language As Integer


Public Property Get CurrentLanguage() As Integer

CurrentLanguage = Language

End Property

Public Property Let CurrentLanguage(ByVal vNewValue As Integer)

Language = vNewValue

End Property


Public Sub LangJmacroLoad(Lang As Integer)
Dim mnuweb As String
Dim Label1 As String
Dim cmdHelp As String
Dim mnuHelp2 As String
Dim mnuabout As String
Dim cmdRecord As String
Dim cmdmanage As String
Dim vismnuExit As String
Dim vismnuhelp As String
Dim vismnuHide As String
Dim cmdreviews As String
Dim cmdoptions As String
Dim cmdLanguage As String
Dim cmdPlayBack As String
Dim vismnuabout As String
Dim vismnuhelpt As String
Dim vismnusmack As String
Dim vismnumanage As String
Dim vismnuRecord As String
Dim vismnuControl As String
Dim vismnuoptions As String
Dim vismnuplayback As String
Dim vismnuLanguage As String
Dim vismnuSendReviews As String
Dim vismnuAvailbemacros As String
Dim vismnuCprinGoldonline As String

Select Case Lang
       
       Case 1 'Arabic
            
            cmdHelp = "& ⁄·Ì„« "
            cmdLanguage = "&«··€Â"
            cmdmanage = "«&œ«—Â"
            cmdoptions = "&ŒÌ«—« "
            cmdPlayBack = " &‰›Ì–"
            cmdRecord = " &”ÃÌ· „«ﬂ—Ê ÃœÌœ"
            cmdreviews = "√&—”· —√Ìﬂ"
            Label1 = "«&⁄«œ…  ‰›Ì–"
            vismnuabout = "&ÕÊ·..."
            vismnuControl = " Õ&ﬂ„"
            vismnuCprinGoldonline = "CprinGold ⁄·Ï «·ÊÌ» "
            vismnuExit = "«‰Â«¡"
            vismnuhelp = " ⁄·Ì„« "
            vismnuhelpt = " ⁄·Ì„« "
            vismnuLanguage = "·€Â"
            vismnuHide = "√Œ›«¡"
            vismnumanage = "«œ«—Â"
            vismnuoptions = "ŒÌ«—« "
            vismnuplayback = " ‰›Ì–"
            vismnuRecord = " ”ÃÌ·"
            vismnuSendReviews = "√—”· —√Ìﬂ"
            vismnusmack = "—∆Ì”Ï"
            mnuabout = "ÕÊ·"
            mnuHelp2 = " ⁄·Ì„« "
            mnuweb = "CprinGold ⁄·Ï «·ÊÌ»"
       
             
       Case 2 'English
            mnuabout = "About..."
            mnuHelp2 = "Help Topics"
            mnuweb = "CprinGold online"
            cmdHelp = "&Help"
            cmdLanguage = "Language"
            cmdmanage = "&Manage"
            cmdoptions = "&Options"
            cmdPlayBack = "&Playback"
            cmdRecord = "&Record New Macro"
            cmdreviews = "&Exit"
            Label1 = "Rebeat Count"
            vismnuabout = "About..."
            vismnuControl = "&Control"
            vismnuCprinGoldonline = "CprinGold online"
            vismnuExit = "&Exit"
            vismnuhelp = "Help"
            vismnuhelpt = "Help Topics"
            vismnuHide = "Hide Smack"
            vismnumanage = "Manage"
            vismnuoptions = "Options"
            vismnuplayback = "Playback"
            vismnuRecord = "Recored New Macro"
            vismnuSendReviews = "Send Reviews"
            vismnusmack = "Main"
       
              
       Case 3 'France
            mnuabout = "Au sujet de..."
            mnuHelp2 = "MatiËres d'aide"
            mnuweb = "CprinGold"
            cmdHelp = "Aide"
            cmdLanguage = "Langue"
            cmdmanage = "ContrÙlez"
            cmdoptions = "Options"
            cmdPlayBack = "Playback"
            cmdRecord = "Record"
            cmdreviews = "les revues"
            Label1 = "Compte de rebeat"
            vismnuabout = "Au sujet de..."
            vismnuControl = "Commande"
            vismnuCprinGoldonline = "CprinGold"
            vismnuExit = "Sortie"
            vismnuhelp = "Aide"
            vismnuhelpt = "MatiËres d'aide"
            vismnuHide = "Peau Smack"
            vismnumanage = "ContrÙlez"
            vismnuoptions = "Options"
            vismnuplayback = "Playback"
            vismnuRecord = "Nouveau macro record"
            vismnuSendReviews = "Envoyez les revues"
            vismnusmack = "Principal"
       
       Case 4 'Italian
       
       Case 5 'Spanish
       
       Case 6 'German
       
       Case 7 'Portguesesse
       
       

End Select





With frmJMacro

     .cmdHelp.caption = cmdHelp
     .cmdLanguage.caption = cmdLanguage
     .cmdmanage.caption = cmdmanage
     .cmdoptions.caption = cmdoptions
     .cmdPlayBack.caption = cmdPlayBack
     .cmdAliasRecord.caption = cmdRecord
     .cmdreviews.caption = cmdreviews
     .Label1.caption = Label1
     .mnuHelp2.caption = mnuHelp2
     .mnuweb.caption = mnuweb
     .mnuabout.caption = mnuabout
     .vismnuabout.caption = vismnuabout
     .vismnuControl.caption = vismnuControl
     .vismnuCprinGoldonline.caption = vismnuCprinGoldonline
     .vismnuExit.caption = vismnuExit
     .vismnuhelp.caption = vismnuhelp
     .vismnuhelpt.caption = vismnuhelpt
     .vismnuHide.caption = vismnuHide
     .vismnumanage.caption = vismnumanage
     .vismnuoptions.caption = vismnuoptions
     .vismnuplayback.caption = vismnuplayback
     .vismnuRecord.caption = vismnuRecord
     .vismnuSendReviews.caption = vismnuSendReviews
     .vismnusmack.caption = vismnusmack
     
End With




End Sub

Public Sub LangFrmAboutLoad(Lang As Integer)

Dim cmdOK As String
Dim Label4 As String
Dim caption As String
Dim Command1 As String
Dim Descriton As String


Select Case Lang

       Case 1
              
            Descriton = "Ì„ﬂ‰ﬂ  Ê›Ì— «·ﬂÀÌ— „‰ «·Êﬁ  »√” Œœ«„Â ⁄‰ ÿ—Ìﬁ  ”ÃÌ· ÷—»«  «·›√—Â Ê ·ÊÕ… «·„›« ÌÕ À„ «⁄«œ… «· ‰›Ì– ›Ï Êﬁ  ·«Õﬁ"
             
            cmdOK = "Õ”‰«"
             
            Command1 = "«·‘—«¡ «·√‰"
             
            caption = "ÕÊ·..."
            
            Label4 = ""
            
       Case 2
            
            Descriton = "Save your time with Smack by Record your Keyboard and Mouse actions and playback them later."
             
            cmdOK = "ok"
             
            Command1 = "order"
             
            caption = "About..."
            
            Label4 = "© Copyright 2005, CprinGold Software. All rights reserved.By using this software, you agree to be bound by the            licence agrement distributed with application."
       
       Case 3
 
            Descriton = "Ëconomiser votre temps avec Smack enregistrez prËs votre clavier et actions "
            
            cmdOK = "ok"
             
            Command1 = "ordre"
             
            caption = "au sujet de" & "..."
            
            Label4 = "© Copyright 2005, CprinGold Software. Toutes les droites reserved.By en utilisant ce logiciel, vous acceptez d'Ítre liÈs par l'accord de permis distribuÈ avec l'application."
       
       Case 4
       
       Case 5
       
       Case 6
       
       Case 7

End Select

With frmAbout
     
     .lblDescription = Descriton
      
     .cmdOK.caption = cmdOK
     
     .Command1.caption = Command1
     
     .caption = caption
     
   '  .Label4 = Label4
     

End With



End Sub

Public Sub LangfrmEnterMacroName(Lang As Integer)

Dim lblDescription, Label1, Label2, CancelButton, OKButton, caption

Select Case Lang
       
       Case 1
       
            OKButton = "&Õ›Ÿ"
            
            CancelButton = "&«·€«¡"
            
            caption = "Õ›Ÿ „«ﬂ—Ê ÃœÌœ"
             
            lblDescription = "Ê’›"
            
            Label1 = "√œŒ· √”„ ··„«ﬂ—Ê «·ÃœÌœ"
            
            Label2 = "√” Œœ„ „› «Õ √Œ ’«— · ‰›Ì– «·„«ﬂ—Ê"
               
             
       Case 2
               
            OKButton = "&Save"
            
            CancelButton = "&Cancel"
            
            caption = "Save New Macro"
             
            lblDescription = "Description"
            
            Label1 = "Enter Name for this Macro"
            
            Label2 = "Use Shortcut Key to playback this Macro"
            
            
           
       
       Case 3
            
            OKButton = "Èconomiser"
            
            CancelButton = "dÈcommandez"
            
            caption = "nouveau macro"
             
            lblDescription = "description"
            
            Label1 = "Ecrivez le nom pour ce macro"
            
            Label2 = "Employez la clef de raccourci au playback ce macro"
           
       Case 4
       
       Case 5
       
       Case 6
       
       Case 7
       


End Select

With frmEnterMacroName
            
     .caption = caption
       
     .Label1.caption = Label1
     
     .Label2.caption = Label2
     
     .OKButton.caption = OKButton
     
     .CancelButton.caption = CancelButton
      
     .lblDescription.caption = lblDescription
     
End With

End Sub

Public Sub LangFrmLanguage(Lang As Integer)

Dim cmdCancel, cmdSet, lstLanguage, caption, Defual

Select Case Lang
       
       Case 1
            lstLanguage = 1
            cmdSet = " ‰›Ì–"
            caption = "«··€Â"
            cmdCancel = "«·€«¡"
            Defual = "«› —«÷ÌÂ"
       Case 2
            lstLanguage = 2
            cmdSet = "&Set"
            caption = "Language"
            cmdCancel = "&Cancel"
            Defual = "&Defualt"
            
       Case 3
            lstLanguage = 3
            cmdSet = "placez"
            caption = "Langue"
            cmdCancel = "dÈcommandez"
            Defual = "dÈfaut"
       Case 4
       Case 5
       Case 6
       Case 7
End Select


With frmLanguage
       
     .cmdCancel.caption = cmdCancel
     
     .cmdSet.caption = cmdSet
      
     .Default.caption = Defual
    
      Select Case lstLanguage
             Case 1
                  
                  .lstLanguage.List(0) = "«·⁄—»ÌÂ"
                  .lstLanguage.List(1) = "«·√‰Ã·Ì“ÌÂ"
                  .lstLanguage.List(2) = "«·›—‰”ÌÂ"
                  .lstLanguage.List(3) = "«·√·„«‰ÌÂ"
                  .lstLanguage.List(4) = "«·√Ìÿ«·ÌÂ"
                  .lstLanguage.List(5) = "«·»— €«·ÌÂ"
                  .lstLanguage.List(6) = "«·√”»«‰ÌÂ"
             
             Case 2
                  
                  .lstLanguage.List(0) = "Arabic"
                  .lstLanguage.List(1) = "English"
                  .lstLanguage.List(2) = "France"
                  .lstLanguage.List(3) = "German"
                  .lstLanguage.List(4) = "Italian"
                  .lstLanguage.List(5) = "Portuguese"
                  .lstLanguage.List(6) = "Spanish"
             
              Case 3
                  
                  .lstLanguage.List(0) = "Arabe"
                  .lstLanguage.List(1) = "Anglais"
                  .lstLanguage.List(2) = "La France"

                
      
      End Select
     
     .caption = caption
     
     
End With


End Sub

Public Sub LangFrmManage(Lang As Integer)

Dim item1, item2, item3, item4, item5, item6
Dim caption, OKButton, cmdshortkey, cmdDeleteAll, cmddelete, cmdrename, cmdstartup


Select Case Lang
       Case 1
            item1 = "«”„ «·„«ﬂ—Ê   "
            item2 = "                      «·Ê’›"
            item3 = "«· «—ÌŒ"
            item4 = "«·Êﬁ "
            item5 = "„› «Õ √Œ ’«—"
            item6 = "„·«ÕŸ« "
            caption = "√œ«—…"
            OKButton = "«€·«ﬁ"
            cmddelete = "«·€«¡"
            cmdDeleteAll = "«·€«¡ «·ﬂ·"
            cmdrename = "√⁄«œ…  ”„ÌÂ"
            cmdshortkey = "„› «Õ √Œ ’«—"
            cmdstartup = " „«ﬂ—Ê »œ¡ «· ‘€Ì·"
       Case 2
            item1 = "   Macro Name"
            item2 = "                      Description"
            item3 = "Date"
            item4 = "Time"
            item5 = "Shortcut Key"
            item6 = "Notes"
            caption = "Manage your Macros"
            OKButton = "&Close"
            cmddelete = "&Delete"
            cmdDeleteAll = "Delete &All"
            cmdrename = "&Rename"
            cmdshortkey = "&Shortcut key"
            cmdstartup = "St&artup Macro"
       Case 3
            item1 = "   macro nom"
            item2 = "                      Description"
            item3 = "Date"
            item4 = "temps"
            item5 = "clef de raccourci"
            item6 = "Notes"
            caption = "contrÙlez Macros"
            OKButton = "fermez-vous"
            cmddelete = "effacement"
            cmdDeleteAll = "supprimez tous"
            cmdrename = "retitrez"
            cmdshortkey = "clef de raccourci"
            cmdstartup = "Macro De dÈmarrage"
       
       Case 4
            item1 = "   Makroname"
            item2 = "                      Beschreibung"
            item3 = "Datum"
            item4 = "Zeit"
            item5 = "K¸rzungschl¸ssel"
            item6 = "Notes"
            caption = "handhaben Sie"
            OKButton = "schlie?en Sie"
            cmddelete = "&Lˆschung"
            cmdDeleteAll = "l?schen Sie alle"
            cmdrename = "benennen Sie"
            cmdshortkey = "K¸rzungschl¸ssel"
            cmdstartup = "StartcMakro"
       Case 5
       Case 6
       Case 7
End Select
With frmManage
    
     .caption = caption
     .OKButton.caption = OKButton
     .cmdshortkey.caption = cmdshortkey
     .cmdDeleteAll.caption = cmdDeleteAll
     '.cmdRunStartup.Caption = cmdstartup
     .cmdDeleteMacro.caption = cmddelete
     .cmdRenameMacro.caption = cmdrename
     .lstviewMacros.ColumnHeaders.Item(1).Text = item1
     .lstviewMacros.ColumnHeaders.Item(2).Text = item2
     .lstviewMacros.ColumnHeaders.Item(3).Text = item3
     .lstviewMacros.ColumnHeaders.Item(4).Text = item4
     .lstviewMacros.ColumnHeaders.Item(5).Text = item5
     .lstviewMacros.ColumnHeaders.Item(6).Text = item6
          
End With
          

        

End Sub

Public Sub LangFrmOptions(Lang As Integer)
Dim chkStartUp, fraSample1, tab1, tab2, tab3
Dim OptHigh, optlow, optnormal, chkAlways, cmdApply
Dim cmdCancel, cmdChange, cmdDefault, cmdOK, caption

Select Case Lang
       Case 1
            chkStartUp = " ‘€Ì· «·»—‰«„Ã ⁄‰œ »œ¡  ‘€Ì· «·ﬂ„»ÌÊ —"
            fraSample1 = "√Œ — ”—⁄…  ‰›Ì– «·„«ﬂ—Ê"
            tab1 = "”—⁄…  ‰›Ì– «·„«ﬂ—Ê"
            tab2 = "»œ¡ «· ‘€Ì·"
            tab3 = "„›« ÌÕ «·√Œ ’«—"
            OptHigh = "”—⁄Â ⁄«·ÌÂ"
            optlow = "”—⁄Â »ÿÌ∆Â"
            optnormal = "”—⁄Â ⁄«œÌÂ"
            chkAlways = "√” Œœ„ Â–Â «·”—⁄Â œ«∆„«"
            cmdApply = " ÿ»Ìﬁ"
            cmdCancel = "«·€«¡"
            cmdChange = " €ÌÌ—"
            cmdDefault = "«·√› —«÷ÌÂ"
            cmdOK = "«€·«ﬁ"
            caption = "ŒÌ«—« "
       Case 2
            chkStartUp = "&Run Smack when computer Startup"
            fraSample1 = "Choose a Playback Speed"
            tab1 = "&PlayBack Speed"
            tab2 = "&Startup"
            tab3 = "Shortcut &Keys"
            OptHigh = "&High Speed"
            optlow = "&Low Speed"
            optnormal = "&Normal Speed"
            chkAlways = "Al&ways PlayBack with this Speed"
            cmdApply = "&Apply"
            cmdCancel = "&Cancel"
            cmdChange = "C&hange"
            cmdDefault = "&Default"
            cmdOK = "&Ok"
            caption = "Smack Options"
       Case 3
            chkStartUp = "Courez Smack quand dÈmarrage d'ordinateur"
            fraSample1 = "vitesse de playback"
            tab1 = "vitesse de playback"
            tab2 = "DÈmarrage"
            tab3 = "clef de raccourci"
            OptHigh = "‡ grande vitesse"
            optlow = " ‡ vitesse rÈduite"
            optnormal = "vitesse normale"
            chkAlways = "employez toujours cette vitesse"
            cmdApply = "appliquez"
            cmdCancel = "dÈcommandez"
            cmdChange = "changez"
            cmdDefault = "dÈfaut"
            cmdOK = "ok"
            caption = "options"
End Select

With frmOptions

      .chkStartUp.caption = chkStartUp
      .fraSample1.caption = fraSample1
      .SSTab1.TabCaption(0) = tab1
      .SSTab1.TabCaption(1) = tab2
      .SSTab1.TabCaption(2) = tab3
     .OptHigh.caption = OptHigh
     .optLowSpeed.caption = optlow
     .optNormalSpeed.caption = optnormal
     .chkAlways.caption = chkAlways
     .cmdApply.caption = cmdApply
     .cmdCancel.caption = cmdCancel
     .cmdChange.caption = cmdChange
     .cmdDefualt.caption = cmdDefault
     .cmdOK.caption = cmdOK
     .caption = caption
End With


End Sub

Public Sub LangFrmRenameMacro(Lang As Integer)
Dim caption, lblDescrtion, Label1, CancelButton, OKButton

Select Case Lang
       Case 1
            
            caption = "√⁄«œ…  ”„ÌÂ «·„«ﬂ—Ê"
            lblDescrtion = "Ê’›"
            Label1 = "√”„ «·„«ﬂ—Ê «·ÃœÌœ"
            CancelButton = "√·€«¡"
            OKButton = "«€·«ﬁ"
       
       Case 2
       
            caption = "Rename Macro"
            lblDescrtion = "Descrtion"
            Label1 = "New Macro Name"
            CancelButton = "&Cancel"
            OKButton = "&Ok"
       
       Case 3
       
            caption = "Retitrez Macro"
            lblDescrtion = "Description"
            Label1 = "Ècrivez le macro nom"
            CancelButton = "&DÈcommandez"
            OKButton = "&Ok"
       
End Select




With frmRenameMacro

     .caption = caption
     .lblDescription.caption = lblDescrtion
     .Label1.caption = Label1
     .CancelButton.caption = CancelButton
     .OKButton.caption = OKButton
     
End With



End Sub

Public Sub LangFrmSetShortKey(Lang As Integer)
Dim caption, Label3, cmdClose, OKButton, CancelButton

Select Case Lang
       Case 1
            caption = " ⁄ÌÌ‰ „› «Õ √Œ ’«—"
            Label3 = "√” Œœ„ „› «Õ √Œ ’«—"
            cmdClose = "√€·«ﬁ"
            OKButton = "Õ›Ÿ"
            CancelButton = "«·€«¡"
       
       Case 2
            caption = "Shortcut Key"
            Label3 = "Use Shortcut Key"
            cmdClose = "&Close"
            OKButton = "&Set"
            CancelButton = "&Cancel"
       Case 3
            caption = "Clef de raccourci"
            Label3 = "Employez la clef de raccourci"
            cmdClose = "Fermez-vous"
            OKButton = "&Placez"
            CancelButton = "&DÈcommandez"
      
End Select

With frmSetShortKey
    
     .caption = caption
     
     .Label3.caption = Label3
     
     .cmdClose.caption = cmdClose
     
     .OKButton.caption = OKButton
     
     .CancelButton.caption = CancelButton
 
End With



End Sub





Public Sub LangFrmTip(Lang As Integer)
Dim lblTip, Label1, OKButton

Select Case Lang

       Case 1
         
         Label1 = "·«  ŸÂ— À«‰Ì…"
         OKButton = "Õ”‰«"
      
       Case 2
       
         Label1 = "Don't show this again"
         OKButton = "&Ok"
       
       Case 3
          
         Label1 = "Ne montrez pas ceci encore"
         OKButton = "&Ok"
       
End Select

With frmTip

     .Label1.caption = Label1
     .OKButton.caption = OKButton
End With


End Sub

Public Sub LangSystray(Lang As Integer)
Dim menuabout, mnuExit, mnuSendReviews, CprinGold, mnuoptions
Dim mnuPlayback, mnumanage, menuHelp, mnuRecord
Select Case Lang
       Case 1
            mnuRecord = " ”ÃÌ· „«ﬂ—Ê ÃœÌœ"
            menuHelp = " ⁄·Ì„« "
            mnumanage = "√œ«—…"
            mnuPlayback = " ‰›Ì–"
            mnuoptions = "ŒÌ«—« "
            mnuExit = "«‰Â«¡"
            CprinGold = "CprinGold ⁄·Ï «·ÊÌ»"
            menuabout = "ÕÊ·..."
            mnuSendReviews = "√—”· —√Ìﬂ"
       Case 2
            mnuRecord = "Record New Macro"
            menuHelp = "Help Topics"
            mnumanage = "Manage Macros"
            mnuPlayback = "PlayBack"
            mnuoptions = "Options"
            mnuExit = "Exit"
            CprinGold = "CprinGold online"
            menuabout = "About..."
            mnuSendReviews = "Send Reviews"
       Case 3
            mnuRecord = "Nouveau macro record"
            menuHelp = "MatiËres d'aide"
            mnumanage = "ContrÙlez Macros"
            mnuPlayback = "PlayBack"
            mnuoptions = "Options"
            mnuExit = "Sortie"
            CprinGold = "CprinGold"
            menuabout = "Au sujet de..."
            mnuSendReviews = "Envoyez les revues"
End Select

With frmJMacro
     .mnuRecord.caption = mnuRecord
     .menuHelp.caption = menuHelp
     .mnumanage.caption = mnumanage
     .mnuPlayback.caption = mnuPlayback
     .mnuoptions.caption = mnuoptions
     .CprinGold.caption = CprinGold
     .mnuExit.caption = mnuExit
     .menuabout.caption = menuabout
     .mnuSendReviews.caption = mnuSendReviews
End With


End Sub
