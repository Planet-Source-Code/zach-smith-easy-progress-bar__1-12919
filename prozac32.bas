Attribute VB_Name = "prozac32"
'Prozac32.bas
'Here it is, my Aol6 bas
'All subs were tested, and designed to work with Aol 6
'I'd enjoy all feedback, questions, comments whatever...
'email me at: proxzach@yahoo.com

Option Explicit
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (destination As Any, source As Any, ByVal Length As Long)
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Const PROCESS_READ = &H10
Public Const RIGHTS_REQUIRED = &HF0000
Public Const LB_GETCOUNT = &H18B
Public Const LB_GETITEMDATA = &H199




Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function SendMessageByNum& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Declare Function SendMessageLong& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)

Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long


Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFilename As String) As Long
Public Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long

Public Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long

Public Declare Function MciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long

Public Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal y As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function SndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFilename As String) As Long
Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Const BM_GETCHECK = &HF0
Public Const BM_SETCHECK = &HF1
Public Const HWND_NOTOPMOST = -2
Public Const HWND_TOPMOST = -1

Public Const LB_GETTEXT = &H189
Public Const LB_GETTEXTLEN = &H18A
Public Const LB_SETCURSEL = &H186
Public Const LB_SETSEL = &H185
Public Const SND_ASYNC = &H1
Public Const SND_NODEFAULT = &H2
Public Const SND_FLAG = SND_ASYNC Or SND_NODEFAULT
Public Const SW_Hide = 0
Public Const SW_SHOW = 5
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const VK_DOWN = &H28
Public Const VK_LEFT = &H25
Public Const VK_MENU = &H12
Public Const VK_RETURN = &HD
Public Const VK_RIGHT = &H27
Public Const VK_SHIFT = &H10
Public Const VK_SPACE = &H20
Public Const VK_UP = &H26
Public Const WM_CHAR = &H102
Public Const WM_CLOSE = &H10
Public Const WM_COMMAND = &H111
Public Const WM_GETTEXT = &HD
Public Const WM_GETTEXTLENGTH = &HE
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_MOVE = &HF012
Public Const WM_SETTEXT = &HC
Public Const WM_SYSCOMMAND = &H112


Public Const ENTER_KEY = 13
Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Public Type POINTAPI
        X As Long
        y As Long
End Type



Public Function AddRoomtoList(list As ListBox)









Public Const PROCESS_READ = &H10
Public Const RIGHTS_REQUIRED = &HF0000
Public Const LB_GETCOUNT = &H18B
Public Const LB_GETITEMDATA = &H199

Public Sub AddAOLListToListbox(ListToGet As Long, ListToPut As ListBox)
    ' I get a lot of people asking me questions like "How do I add my
    ' buddy list to a listbox?" So I put this together. It's an edited
    ' version of the addroomtolistbox sub found in most bas files.
    ' To create this sub I looked at several addroomtolist subs (I'm
    ' telling you this because I don't like to take credit for stuff I
    ' just played around with. The sub edited for this was the
    ' addroom sub from dos32.bas). I looked at these subs because they
    ' will add almost any other aol list to a list box. Every addroom
    ' sub I looked at looked like the others (what's that tell you).
    ' Anyway, I also put in some comments so you could better
    ' understand these types of subs.
    On Error Resume Next
    Dim cProcess As Long, itmHold As Long, ListItem As String
    Dim psnHold As Long, rBytes As Long, i As Integer
    Dim sThread As Long, mThread As Long
    ' Obtain the identifiers of a thread and process that are associated
    ' with the window. A process is a running application and a thread
    ' is a task that the program is doing (like a program could be doing
    ' several things, each of these things would be a thread).
    sThread = GetWindowThreadProcessId(ListToGet, cProcess)
    ' Open the handle to the existing process
    mThread = OpenProcess(PROCESS_READ Or RIGHTS_REQUIRED, False, cProcess)
    If mThread <> 0 Then
        For i = 0 To SendMessage(ListToGet, LB_GETCOUNT, 0, 0) - 1
            ListItem = String(4, vbNullChar)
            itmHold = SendMessage(ListToGet, LB_GETITEMDATA, ByVal CLng(i), ByVal 0&)
            itmHold = itmHold + 24
            ' Read memory from the address space of the process
            Call ReadProcessMemory(mThread, itmHold, ListItem, 4, rBytes)
            Call CopyMemory(psnHold, ByVal ListItem, 4)
            psnHold = psnHold + 6
            ListItem = String(16, vbNullChar)
            Call ReadProcessMemory(mThread, psnHold, ListItem, Len(ListItem), rBytes)
            ' cut nulls off
            ListItem = Left(ListItem, InStr(ListItem, vbNullChar) - 1)
            ListToPut.AddItem ListItem
        Next i
        Call CloseHandle(mThread)
    End If
End Sub




Public Function CopyFromListtoList(ListToCopyFrom As ListBox, ListToCopyTo As ListBox)
'this will copy all the contents of one listbox to another...
'Example:
'Call CopyFromListtoList(list1, list2)
'that will copy everything from listbox 1 over to listbox 2
Dim i As Long
For i& = 0& To ListToCopyFrom.ListCount - 1
ListToCopyFrom.ListIndex = i&
ListToCopyTo.AddItem ListToCopyFrom.Text
Next i&
End Function

Public Function EnterMemberRoom(room As String)
'Enter a member chatroom, this is good for a roombust,
'example:
'Call EnterMemberRoom(Text1)
   Call Keyword("aol://2719:61-2-" + room$)
End Function
Public Function EnterPrivateRoom(room As String)
'Enter a private chatroom, this is good for a roombust,
'example:
'Call EnterPrivateRoom(Text1)
Call Keyword("aol://2719:2-2-" + room$)
End Function
Public Function FindAol() As Long
'find if the aol window is present
'example:
'If FindAol <> 0& then
'msgbox "Aol window found!"
'else
'msgbox "Aol not found!"
'end if

Dim Counter As Long
Dim AOLMMI As Long
Dim AOLToolbar As Long
Dim MDIClient As Long
Dim AOLFrame As Long
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLToolbar& = FindWindowEx(AOLFrame&, 0&, "AOL Toolbar", vbNullString)
AOLMMI& = FindWindowEx(AOLFrame&, 0&, "_AOL_MMI", vbNullString)
Do While (Counter& <> 100&) And (MDIClient& = 0& Or AOLToolbar& = 0& Or AOLMMI& = 0&): DoEvents
    AOLFrame& = FindWindowEx(AOLFrame&, AOLFrame&, "AOL Frame25", vbNullString)
    MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
    AOLToolbar& = FindWindowEx(AOLFrame&, 0&, "AOL Toolbar", vbNullString)
    AOLMMI& = FindWindowEx(AOLFrame&, 0&, "_AOL_MMI", vbNullString)
    If MDIClient& And AOLToolbar& And AOLMMI& Then Exit Do
    Counter& = Val(Counter&) + 1&
Loop
If Val(Counter&) < 100& Then
    FindAol& = AOLFrame&
    Exit Function
End If
End Function
Public Function FindChatroom() As Long
'Finds the aol chatroom
'Example:
'If Findchatroom <> 0& then
'msgbox GetWindowCaption(Findchatroom) + "window found!"
'else
'msgbox "chatroom not found!"
'end if

Dim Counter As Long
Dim AOLStatic5 As Long
Dim AOLIcon3 As Long
Dim AOLStatic4 As Long
Dim AOLListbox As Long
Dim AOLStatic3 As Long
Dim AOLImage As Long
Dim AOLIcon2 As Long
Dim RICHCNTL2 As Long
Dim AOLStatic2 As Long
Dim i As Long
Dim AOLIcon As Long
Dim AOLCombobox As Long
Dim richcntl As Long
Dim AOLStatic As Long
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", vbNullString)
AOLStatic& = FindWindowEx(AOLChild&, 0&, "_AOL_Static", vbNullString)
richcntl& = FindWindowEx(AOLChild&, 0&, "RICHCNTL", vbNullString)
AOLCombobox& = FindWindowEx(AOLChild&, 0&, "_AOL_Combobox", vbNullString)
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
For i& = 1& To 3&
    AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
Next i&
AOLStatic2& = FindWindowEx(AOLChild&, AOLStatic&, "_AOL_Static", vbNullString)
RICHCNTL2& = FindWindowEx(AOLChild&, richcntl&, "RICHCNTL", vbNullString)
AOLIcon2& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
For i& = 1& To 2&
    AOLIcon2& = FindWindowEx(AOLChild&, AOLIcon2&, "_AOL_Icon", vbNullString)
Next i&
AOLImage& = FindWindowEx(AOLChild&, 0&, "_AOL_Image", vbNullString)
AOLStatic3& = FindWindowEx(AOLChild&, AOLStatic2&, "_AOL_Static", vbNullString)
AOLStatic3& = FindWindowEx(AOLChild&, AOLStatic3&, "_AOL_Static", vbNullString)
AOLListbox& = FindWindowEx(AOLChild&, 0&, "_AOL_Listbox", vbNullString)
AOLStatic4& = FindWindowEx(AOLChild&, AOLStatic3&, "_AOL_Static", vbNullString)
AOLIcon3& = FindWindowEx(AOLChild&, AOLIcon2&, "_AOL_Icon", vbNullString)
For i& = 1& To 7&
    AOLIcon3& = FindWindowEx(AOLChild&, AOLIcon3&, "_AOL_Icon", vbNullString)
Next i&
AOLStatic5& = FindWindowEx(AOLChild&, AOLStatic4&, "_AOL_Static", vbNullString)
Do While (Counter& <> 100&) And (AOLStatic& = 0& Or richcntl& = 0& Or AOLCombobox& = 0& Or AOLIcon& = 0& Or AOLStatic2& = 0& Or RICHCNTL2& = 0& Or AOLIcon2& = 0& Or AOLImage& = 0& Or AOLStatic3& = 0& Or AOLListbox& = 0& Or AOLStatic4& = 0& Or AOLIcon3& = 0& Or AOLStatic5& = 0&): DoEvents
    AOLChild& = FindWindowEx(MDIClient&, AOLChild&, "AOL Child", vbNullString)
    AOLStatic& = FindWindowEx(AOLChild&, 0&, "_AOL_Static", vbNullString)
    richcntl& = FindWindowEx(AOLChild&, 0&, "RICHCNTL", vbNullString)
    AOLCombobox& = FindWindowEx(AOLChild&, 0&, "_AOL_Combobox", vbNullString)
    AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
    For i& = 1& To 3&
        AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
    Next i&
    AOLStatic2& = FindWindowEx(AOLChild&, AOLStatic&, "_AOL_Static", vbNullString)
    RICHCNTL2& = FindWindowEx(AOLChild&, richcntl&, "RICHCNTL", vbNullString)
    AOLIcon2& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
    For i& = 1& To 2&
        AOLIcon2& = FindWindowEx(AOLChild&, AOLIcon2&, "_AOL_Icon", vbNullString)
    Next i&
    AOLImage& = FindWindowEx(AOLChild&, 0&, "_AOL_Image", vbNullString)
    AOLStatic3& = FindWindowEx(AOLChild&, AOLStatic2&, "_AOL_Static", vbNullString)
    AOLStatic3& = FindWindowEx(AOLChild&, AOLStatic3&, "_AOL_Static", vbNullString)
    AOLListbox& = FindWindowEx(AOLChild&, 0&, "_AOL_Listbox", vbNullString)
    AOLStatic4& = FindWindowEx(AOLChild&, AOLStatic3&, "_AOL_Static", vbNullString)
    AOLIcon3& = FindWindowEx(AOLChild&, AOLIcon2&, "_AOL_Icon", vbNullString)
    For i& = 1& To 7&
        AOLIcon3& = FindWindowEx(AOLChild&, AOLIcon3&, "_AOL_Icon", vbNullString)
    Next i&
    AOLStatic5& = FindWindowEx(AOLChild&, AOLStatic4&, "_AOL_Static", vbNullString)
    If AOLStatic& And richcntl& And AOLCombobox& And AOLIcon& And AOLStatic2& And RICHCNTL2& And AOLIcon2& And AOLImage& And AOLStatic3& And AOLListbox& And AOLStatic4& And AOLIcon3& And AOLStatic5& Then Exit Do
    Counter& = Val(Counter&) + 1&
Loop
If Val(Counter&) < 100& Then
    FindChatroom& = AOLChild&
    Exit Function
End If
End Function
Public Function FindIM() As Long
'Find Instant Message window
'Example:
'If FindIM <> 0& then
'msgbox "Instant message window found!"
'else
'msgbox "instant message window not found!"

Dim Counter As Long
Dim AOLIcon2 As Long
Dim richcntl As Long
Dim i As Long
Dim AOLIcon As Long
Dim AOLEdit As Long
Dim AOLStatic As Long
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", vbNullString)
AOLStatic& = FindWindowEx(AOLChild&, 0&, "_AOL_Static", vbNullString)
AOLEdit& = FindWindowEx(AOLChild&, 0&, "_AOL_Edit", vbNullString)
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
For i& = 1& To 8&
    AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
Next i&
richcntl& = FindWindowEx(AOLChild&, 0&, "RICHCNTL", vbNullString)
AOLIcon2& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
AOLIcon2& = FindWindowEx(AOLChild&, AOLIcon2&, "_AOL_Icon", vbNullString)
Do While (Counter& <> 100&) And (AOLStatic& = 0& Or AOLEdit& = 0& Or AOLIcon& = 0& Or richcntl& = 0& Or AOLIcon2& = 0&): DoEvents
    AOLChild& = FindWindowEx(MDIClient&, AOLChild&, "AOL Child", vbNullString)
    AOLStatic& = FindWindowEx(AOLChild&, 0&, "_AOL_Static", vbNullString)
    AOLEdit& = FindWindowEx(AOLChild&, 0&, "_AOL_Edit", vbNullString)
    AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
    For i& = 1& To 8&
        AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
    Next i&
    richcntl& = FindWindowEx(AOLChild&, 0&, "RICHCNTL", vbNullString)
    AOLIcon2& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
    AOLIcon2& = FindWindowEx(AOLChild&, AOLIcon2&, "_AOL_Icon", vbNullString)
    If AOLStatic& And AOLEdit& And AOLIcon& And richcntl& And AOLIcon2& Then Exit Do
    Counter& = Val(Counter&) + 1&
Loop
If Val(Counter&) < 100& Then
    FindIM& = AOLChild&
    Exit Function
End If
End Function
Public Function FormOnTop(Form As Form)
'Make your form on top
'Example: Call FormOnTop(form1)

Call SetWindowPos(Form.hwnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, FLAGS)
End Function
Public Function FormNotOnTop(Form As Form)
'Make your form not on top
'Example: Call FormNotOnTop(form1)

Call SetWindowPos(Form.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
End Function
Public Function GetWindowCaption(WindowHandle As Long) As String
    'This will get a windows caption
    'Example: Text1.Text = GetWindowCaption(FindIM)
    'If There is an IM window open it will get the caption and display it in a textbox
    Dim buffer As String, TextLength As Long
    TextLength& = GetWindowTextLength(WindowHandle&)
    buffer$ = String(TextLength&, 0&)
    Call GetWindowText(WindowHandle&, buffer$, TextLength& + 1)
    GetWindowCaption$ = buffer$
End Function

Public Function Keyword(Word As String)
   'Aol keyword...
   'example: Call Keyword("blah")
    Dim aol As Long
    Dim blah As Long
    Dim Toolbar As Long
    Dim Combo As Long
    Dim EditWindow As Long
    aol& = FindWindow("AOL Frame25", vbNullString)
    blah& = FindWindowEx(aol&, 0&, "AOL Toolbar", vbNullString)
    Toolbar& = FindWindowEx(blah&, 0&, "_AOL_Toolbar", vbNullString)
    Combo& = FindWindowEx(Toolbar&, 0&, "_AOL_Combobox", vbNullString)
    EditWindow& = FindWindowEx(Combo&, 0&, "Edit", vbNullString)
    Call SendMessageByString(EditWindow&, WM_SETTEXT, 0&, Word$)
    Call SendMessageLong(EditWindow&, WM_CHAR, VK_SPACE, 0&)
    Call SendMessageLong(EditWindow&, WM_CHAR, VK_RETURN, 0&)
End Function
Public Function SendChat(Text As String)
'Send chat to chatroom
'Call Sendchat("i hate my job")
Dim richcntl As Long
richcntl& = FindWindowEx(FindChatroom(), 0&, "RICHCNTL", vbNullString)
richcntl& = FindWindowEx(FindChatroom(), richcntl&, "RICHCNTL", vbNullString)
Call SendMessageByString(richcntl&, WM_SETTEXT, 0&, Text$)
Call SendMessageByNum(richcntl&, WM_CHAR, 13, 0&)
End Function
Public Function SendIm(SN As String)
'Sends an IM
       Dim i As Long
Dim AOLIcon As Long
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "Send Instant Message")
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)

Call Keyword("aol://9293:" & SN$)
For i& = 1& To 9&
    AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
Next i&
Call SendMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
Call PostMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)

Call Keyword("aol://9293:" & SN$)

End Function




Public Function TimeOut(time As Long)
    'pause for a set amount of time
    'Call Timeout(1)
    '1 = 1 second....
    Dim Current As Long
    Current = Timer
    Do Until Timer - Current >= time
        DoEvents
    Loop
End Function


