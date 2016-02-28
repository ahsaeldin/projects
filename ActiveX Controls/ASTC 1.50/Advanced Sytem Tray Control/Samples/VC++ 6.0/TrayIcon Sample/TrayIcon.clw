; CLW file contains information for the MFC ClassWizard

[General Info]
Version=1
LastClass=CTrayIconDlg
LastTemplate=CDialog
NewFileInclude1=#include "stdafx.h"
NewFileInclude2=#include "TrayIcon.h"
ODLFile=TrayIcon.odl

ClassCount=4
Class1=CTrayIconApp
Class2=CTrayIconDlg
Class3=CAboutDlg
Class4=CTrayIconDlgAutoProxy

ResourceCount=6
Resource1=IDD_ABOUTBOX
Resource2=IDR_MAINFRAME
Resource3=IDD_TRAYICON_DIALOG
Resource4=IDD_ABOUTBOX (English (U.S.))
Resource5=IDD_TRAYICON_DIALOG (English (U.S.))
Resource6=IDR_MENU1

[CLS:CTrayIconApp]
Type=0
HeaderFile=TrayIcon.h
ImplementationFile=TrayIcon.cpp
Filter=N

[CLS:CTrayIconDlg]
Type=0
HeaderFile=TrayIconDlg.h
ImplementationFile=TrayIconDlg.cpp
Filter=D
LastObject=CTrayIconDlg
BaseClass=CDialog
VirtualFilter=dWC

[CLS:CAboutDlg]
Type=0
HeaderFile=TrayIconDlg.h
ImplementationFile=TrayIconDlg.cpp
Filter=D

[DLG:IDD_ABOUTBOX]
Type=1
ControlCount=4
Control1=IDC_STATIC,static,1342177283
Control2=IDC_STATIC,static,1342308352
Control3=IDC_STATIC,static,1342308352
Control4=IDOK,button,1342373889
Class=CAboutDlg

[CLS:CTrayIconDlgAutoProxy]
Type=0
HeaderFile=DlgProxy.h
ImplementationFile=DlgProxy.cpp
BaseClass=CCmdTarget
Filter=N

[DLG:IDD_TRAYICON_DIALOG]
Type=1
ControlCount=3
Control1=IDOK,button,1342242817
Control2=IDCANCEL,button,1342242816
Control3=IDC_STATIC,static,1342308352
Class=CTrayIconDlg

[DLG:IDD_TRAYICON_DIALOG (English (U.S.))]
Type=1
Class=CTrayIconDlg
ControlCount=9
Control1=IDC_BUTSHOW,button,1342242816
Control2=IDC_ChangeIcon,button,1342242816
Control3=IDC_Hide,button,1342242816
Control4=IDC_SBalloon,button,1342242816
Control5=IDC_TRPOP,button,1342242816
Control6=IDC_StopTrack,button,1342242816
Control7=IDC_SBalloon2,button,1342242816
Control8=IDC_STATIC,button,1342177287
Control9=IDC_TRAYICON1,{30E583A6-8E26-11DA-A1C2-807A50C10000},1342242816

[DLG:IDD_ABOUTBOX (English (U.S.))]
Type=1
Class=?
ControlCount=4
Control1=IDC_STATIC,static,1342177283
Control2=IDC_STATIC,static,1342308480
Control3=IDC_STATIC,static,1342308352
Control4=IDOK,button,1342373889

[MNU:IDR_MENU1]
Type=1
Class=CTrayIconDlg
Command1=ID_MAIN_SHOW
Command2=ID_MAIN_EXIT
CommandCount=2

