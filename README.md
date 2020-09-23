<div align="center">

## How to add an icon to the tray


</div>

### Description

One of the questions that occurs most often in the VB Q and A forum is how to add an icon to the tray area of the Windows 95 taskbar.This tip will show you how to add and delete the icon,and also trap the mouse events.
 
### More Info
 
Create two command buttons (command1 and command2) and a picture box (picture1) to the form. For the picture property of the Picture Box select the icon you want to be displayed in the tray.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[King](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/king.md)
**Level**          |Unknown
**User Rating**    |5.0 (30 globes from 6 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Windows System Services](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-system-services__1-35.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/king-how-to-add-an-icon-to-the-tray__1-3160/archive/master.zip)

### API Declarations

```

       Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias _
       "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As _
       NOTIFYICONDATA) As Long
       Public Type NOTIFYICONDATA
         cbSize As Long
         hwnd As Long
         uID As Long
         uFlags As Long
         uCallbackMessage As Long
         hIcon As Long
         szTip As String * 64
       End Type
       Public Const NIM_ADD = &H0
       Public Const NIM_MODIFY = &H1
       Public Const NIM_DELETE = &H2
       Public Const NIF_MESSAGE = &H1
       Public Const NIF_ICON = &H2
       Public Const NIF_TIP = &H4
       'Make your own constant, e.g.:
       Public Const NIF_DOALL = NIF_MESSAGE Or NIF_ICON Or NIF_TIP
       Public Const WM_MOUSEMOVE = &H200
       Public Const WM_LBUTTONDBLCLK = &H203
       Public Const WM_LBUTTONDOWN = &H201
       Public Const WM_RBUTTONDOWN = &H204
```


### Source Code

```
Public Sub CreateIcon()
       Dim Tic As NOTIFYICONDATA
       Tic.cbSize = Len(Tic)
       Tic.hwnd = Picture1.hwnd
       Tic.uID = 1&
       Tic.uFlags = NIF_DOALL
       Tic.uCallbackMessage = WM_MOUSEMOVE
       Tic.hIcon = Picture1.Picture
       Tic.szTip = "Visual Basic Demo Project" & Chr$(0)
       erg = Shell_NotifyIcon(NIM_ADD, Tic)
       End Sub
       Public Sub DeleteIcon()
       Dim Tic As NOTIFYICONDATA
       Tic.cbSize = Len(Tic)
       Tic.hwnd = Picture1.hwnd
       Tic.uID = 1&
       erg = Shell_NotifyIcon(NIM_DELETE, Tic)
       End Sub
Private Sub Command1_Click()
CreateIcon
End Sub
Private Sub Command2_Click()
DeleteIcon
End Sub
Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
X = X / Screen.TwipsPerPixelX
       Select Case X
       Case WM_LBUTTONDOWN
       Caption = "Left Click"
       Case WM_RBUTTONDOWN
       Caption = "Right Click"
       Case WM_MOUSEMOVE
       Caption = "Move"
       Case WM_LBUTTONDBLCLK
       Caption = "Double Click"
       End Select
End Sub
```

