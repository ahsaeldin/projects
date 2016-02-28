; CLW file contains information for the MFC ClassWizard

[General Info]
Version=1
LastClass=CTrayDlg
LastTemplate=CDialog
NewFileInclude1=#include "stdafx.h"
NewFileInclude2=#include "Tray.h"

ClassCount=3
Class1=CTrayApp
Class2=CTrayDlg
Class3=CAboutDlg

ResourceCount=5
Resource1=IDD_ABOUTBOX
Resource2=IDR_MAINFRAME
Resource3=IDD_TRAY_DIALOG
Resource4=IDD_ABOUTBOX (English (U.S.))
Resource5=IDD_TRAY_DIALOG (English (U.S.))

[CLS:CTrayApp]
Type=0
HeaderFile=Tray.h
ImplementationFile=Tray.cpp
Filter=N

[CLS:CTrayDlg]
Type=0
HeaderFile=TrayDlg.h
ImplementationFile=TrayDlg.cpp
Filter=D
LastObject=IDC_MSFLEXGRID1
BaseClass=CDialog
VirtualFilter=dWC

[CLS:CAboutDlg]
Type=0
HeaderFile=TrayDlg.h
ImplementationFile=TrayDlg.cpp
Filter=D
LastObject=IDOK

[DLG:IDD_ABOUTBOX]
Type=1
ControlCount=4
Control1=IDC_STATIC,static,1342177283
Control2=IDC_STATIC,static,1342308352
Control3=IDC_STATIC,static,1342308352
Control4=IDOK,button,1342373889
Class=CAboutDlg


[DLG:IDD_TRAY_DIALOG]
Type=1
ControlCount=3
Control1=IDOK,button,1342242817
Control2=IDCANCEL,button,1342242816
Control3=IDC_STATIC,static,1342308352
Class=CTrayDlg

[DLG:IDD_TRAY_DIALOG (English (U.S.))]
Type=1
Class=CTrayDlg
ControlCount=5
Control1=IDC_TRAYICONS1,{CD7CA542-262E-11DA-A1C1-00E04C770EEB},1342242816
Control2=IDC_MSFLEXGRID1,{6262D3A0-531B-11CF-91F6-C2863C385E30},1342242816
Control3=IDC_BUTTON4,button,1342242816
Control4=IDC_STATIC,button,1342177287
Control5=IDC_BUTTON3,button,1342242816

[DLG:IDD_ABOUTBOX (English (U.S.))]
Type=1
Class=?
ControlCount=4
Control1=IDC_STATIC,static,1342177283
Control2=IDC_STATIC,static,1342308480
Control3=IDC_STATIC,static,1342308352
Control4=IDOK,button,1342373889

