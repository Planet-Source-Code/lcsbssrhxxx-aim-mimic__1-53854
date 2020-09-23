<div align="center">

## \[AIM Mimic\]

<img src="PIC20045171929259466.JPG">
</div>

### Description

AIM Mimic is a program that allows you to log on AIM (AOL Instant Messenger) and then mimic another user. Fetures inclues:

Fake Away, Fake Idle, Edit Profile, Send IM, Mimic, Event Log, Internet Options, Window Flash when reciving an IM, Beep when sending / reciving an IM, help, and much more, see program!

Enjoi!!!

-LCSBSSRHXXX

AIM Mimic uses TocSock.ocx

You can download TocSock.ocx at www.lenshell.com/files.htm
 
### More Info
 
AIM Mimic uses TocSock.ocx

You can download TocSock.ocx at www.lenshell.com/files.htm


<span>             |<span>
---                |---
**Submitted On**   |2004-05-17 17:10:08
**By**             |[LCSBSSRHXXX](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/lcsbssrhxxx.md)
**Level**          |Advanced
**User Rating**    |4.8 (24 globes from 5 users)
**Compatibility**  |VB 6\.0
**Category**       |[Internet/ HTML](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/internet-html__1-34.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[\[AIM\_Mimic1746875172004\.zip](https://github.com/Planet-Source-Code/lcsbssrhxxx-aim-mimic__1-53854/archive/master.zip)

### API Declarations

```
Const FLASHW_STOP = 0
Const FLASHW_CAPTION = &H1
Const FLASHW_TRAY = &H2
Const FLASHW_ALL = (FLASHW_CAPTION Or FLASHW_TRAY)
Const FLASHW_TIMER = &H4
Const FLASHW_TIMERNOFG = &HC
Private Type FLASHWINFO
  cbSize As Long
  hwnd As Long
  dwFlags As Long
  uCount As Long
  dwTimeout As Long
End Type
Private Declare Function FlashWindowEx Lib "user32" (pfwi As FLASHWINFO) As Boolean
Private Declare Function Beep Lib "kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long
Dim Running As Boolean
Dim LoggedOn As Boolean
Dim nRet As Long
Dim sCPL As String
```





