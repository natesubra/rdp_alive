Imports System.Runtime.CompilerServices
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Threading

Module Module1
    ' Original Source Here: https://www.codeproject.com/Tips/1174585/RDP-Session-Keep-Alive#_comments
    ' Change PollIntervalSeconds to however many seconds you want between wakup events
    ' use arg1 for sleep time , arg2 for rdp server name, arg3 for show active rdp server names
    ' ex: rdp_alive.exe 30 server_name 1
    Const PollIntervalSeconds As Integer = 50
    Public interval As Integer = PollIntervalSeconds
    Public server As String = ""
    Public showServer As Boolean = False
    Public BoolMouseLoc As Boolean = False
    Public BoolNeedToRestore As Boolean
    Private Const WmSetfocus As Integer = &H7
    Private Const WmMousemove As Integer = &H200
    Private Const ALT As Integer = &HA4
    Private Const EXTENDEDKEY As Integer = &H1
    Private Const KEYUP As Integer = &H2

    ' ----------------------------------------------------
    ' This is all of the User32 magic that lets you do things like find the windows
    ' handles, titles, class names, dimensions etc...
    ' ----------------------------------------------------
    Delegate Function EnumWindowDelegate(hWnd As IntPtr, lparam As IntPtr) As Boolean
    Private ReadOnly Callback As New EnumWindowDelegate(AddressOf EnumWindowProc)
    Delegate Function EnumWindowDelegate2(hWnd As IntPtr, lparam As IntPtr) As Boolean
    Private ReadOnly Callback2 As New EnumWindowDelegate2(AddressOf EnumWindowProc2)
    Private Declare Function EnumWindows Lib "User32.dll" _
          (wndenumproc As EnumWindowDelegate, lparam As IntPtr) As Boolean
    Private Declare Auto Function EnumChildWindows Lib "User32.dll" _
          (windowHandle As IntPtr, wndenumproc2 As EnumWindowDelegate2, lParam As IntPtr) As Boolean
    Private Declare Auto Function GetWindowText Lib "User32.dll" _
          (hwnd As IntPtr, txt As Byte(), lng As Integer) As Integer
    Private Declare Function IsWindowVisible Lib "User32.dll" (hwnd As IntPtr) As Boolean
    Private Declare Function IsIconic Lib "User32.dll" (hwnd As IntPtr) As Boolean
    Private Declare Function GetWindowTextLengthA Lib "User32.dll" (hwnd As IntPtr) As Integer
    Private Declare Auto Function ShowWindow Lib "User32.dll" (hwnd As IntPtr, lng As Integer) _
        As Integer
    Private Declare Function GetForegroundWindow Lib "User32.dll" () As IntPtr
    Private Declare Function SetForegroundWindow Lib "User32.dll" (hwnd As IntPtr) As Integer
    Declare Auto Function SendMessage Lib "user32.dll" _
        (hWnd As IntPtr, msg As Integer, wParam As IntPtr, lParam As IntPtr) As IntPtr
    Private Declare Function GetClientRect Lib "User32.dll" _
        (hWnd As IntPtr, ByRef lpRect As Rect) As Boolean
    <DllImport("user32.dll", CharSet:=CharSet.Auto)>
    Private Sub GetClassName(hWnd As IntPtr, lpClassName As StringBuilder, nMaxCount As Integer)
    End Sub

    <DllImport("user32.dll")>
    Private Sub keybd_event(bVk As Byte, bScan As Byte, dwFlags As UInteger, dwExtraInfo As Integer)
    End Sub
    ' Makes a structure to store the windows dimensions in
    <StructLayout(LayoutKind.Sequential)>
    Public Structure Rect
        Public Left As Integer
        Public Top As Integer
        Public Right As Integer
        Public Bottom As Integer
        Public Overrides Function ToString() As String
            Return String.Format("{0},{1},{2},{3}", Left, Top, Right, Bottom)
        End Function
    End Structure

    Sub Main(ByVal sArgs() As String)
        Dim storeCurrenthWnd As IntPtr
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance)
        If sArgs.Length > 0 Then
            interval = sArgs(0)
        End If
        If sArgs.Length > 1 Then
            server = sArgs(1)
        End If
        If sArgs.Length > 2 Then
            showServer = (sArgs(2) > 0)
        End If
        Do
            BoolNeedToRestore = False
            ' Switches up the mouse move location each poll
            BoolMouseLoc = Not BoolMouseLoc

            ' Stores the handle of the current window you have in focus.
            ' It will only use this to change focus back if one of your RDP
            ' windows needed to be unminimized.
            storeCurrenthWnd = GetForegroundWindow()

            ' Look through all windows and determine which ones are
            ' RDP related and keep them awake.
            EnumWindows(Callback, IntPtr.Zero)
            ' If the program needed to unminimize one or more of your
            ' RDP windows this will set the focus back on the window
            ' you were on prior to that.
            If BoolNeedToRestore Then
                ShowWindow(storeCurrenthWnd, 5)
                ' Needed to reliably set the foreground window
                keybd_event(CByte(ALT), &H45, EXTENDEDKEY Or 0, 0)
                keybd_event(CByte(ALT), &H45, EXTENDEDKEY Or KEYUP, 0)
                SetForegroundWindow(storeCurrenthWnd)
            End If
            'Force garbage collection
            GC.Collect()
            GC.WaitForPendingFinalizers()
            ' Wait for number of seconds specified in
            ' the PollIntervalSeconds variable.
            For intInterval = 1 To PollIntervalSeconds * 10
                Thread.Sleep(100)
            Next
        Loop
    End Sub

    ' This function looks at the titles of the windows you currently
    ' have open and only looks at those that have "remote desktop" in them.
    ' It then determines if the RDP session windows are minimized, if they are
    ' it will un-minmize them.  This is needed in order to keep them awake.
    '
    ' It then sets focus (in the background without it affecting the users currently
    ' focused window) on each RDP session and then determines where the center of that
    ' window where the mouse SendMessage should be sent to (this is probably not necessary
    ' but I figured why not).  I'm also having it switch mouse locations each poll.
    ' Near the end of the function it then makes a call to all child windows to keep
    ' them awake in the same way that this function does with the exception of not
    ' needing to perform the un-minimizing steps (The child windows are for keeping
    ' each of the RDP sessions in the Remote Desktop Connection Manager active).
    Private Function EnumWindowProc(hWnd As IntPtr, lparam As IntPtr) As Boolean
        Dim myRect As Rect
        Dim arrRect(4) As String
        Dim coord As Integer

        If IsWindowVisible(hWnd) Then
            ' Gather the window titles
            Dim theLength As Integer = GetWindowTextLengthA(hWnd)
            Dim theReturn(theLength * 2) As Byte '2x the size of the Max length
            GetWindowText(hWnd, theReturn, theLength + 1)
            Dim theText = ""
            For x = 0 To (theLength - 1) * 2
                If theReturn(x) <> 0 Then
                    theText &= Chr(theReturn(x))
                End If
            Next
            ' Narrow it to RDP sessions
            If theText.ToLower Like "*" & server & "*remote desktop*" Then
                If showServer Then
                    Console.WriteLine(theText & " - " & CStr(hWnd))
                End If
                ' Un-minimize the RDP sessions that are minimized
                If IsIconic(hWnd) Then
                    ShowWindow(hWnd, 5)
                    ShowWindow(hWnd, 9)
                    ' Since it had to keep at least one RDP session awake
                    ' Make sure it knows to set focus back to the window
                    ' the user was on.
                    BoolNeedToRestore = True
                End If
                ' Set focus on the RDP session (without affecting the users currently
                ' focused window)
                SendMessage(hWnd, WmSetfocus, 0, 0)
                ' Find the RDP window dimensions and figure out the center of it
                GetClientRect(hWnd, myRect)
                arrRect = Split(myRect.ToString(), ",")
                Console.WriteLine("Client Rect   : " & myRect.ToString())
                Console.WriteLine("Click Location: " & CInt(CInt(arrRect(2)) / 2) & " X " &
        CInt(CInt(arrRect(3)) / 2))
                coord = CInt(CInt(arrRect(2)) / 2) * &H10000 + CInt(CInt(arrRect(3)) / 2)
                ' Send the mouse move on the RDP session in the background (without
                ' affecting the users currently focused window).  Swapping the location
                ' each poll.
                If BoolMouseLoc Then
                    SendMessage(hWnd, WmMousemove, 0, 0)
                Else
                    SendMessage(hWnd, WmMousemove, 0, coord)
                End If
                ' Look through all of the child windows as well in order
                ' to keep Remote Desktop Connection Manager RDP sessions
                ' active.
                EnumChildWindows(hWnd, Callback2, IntPtr.Zero)
                Console.WriteLine("")

            End If
        End If
        Return True
    End Function

    ' This does the same thing that the EnumWindowsProc function does but does
    ' it for child windows.  The only thing that is different in this one is that
    ' it doesn't un-minimize anything (that is done for the parent window) and
    ' it gathers the window class so that we can skip several window class types
    ' that that cause issues or don't need to be woken up.
    Private Function EnumWindowProc2(hWnd As IntPtr, lparam As IntPtr) As Boolean
        Dim myRect As Rect
        Dim arrRect(4) As String
        Dim coord As Integer
        Dim myClassName As String
        ' Gather the window titles
        Dim theLength As Integer = GetWindowTextLengthA(hWnd)
        Dim theReturn(theLength * 2) As Byte '2x the size of the Max length
        GetWindowText(hWnd, theReturn, theLength + 1)
        Dim theText = ""
        For x = 0 To (theLength - 1) * 2
            If theReturn(x) <> 0 Then
                theText &= Chr(theReturn(x))
            End If
        Next
        ' Get the window class name
        myClassName = GetWindowClass(hWnd)
        ' Ignore child windows that don't need to be woken up
        If Not theText.Trim().Length = 0 And Not theText.ToLower.MultiContains("disconnected") And
     Not myClassName.ToLower.MultiContains("atl:", ".button.", "uicontainerclass",
      ".systreeview32.", ".scrollbar.", "opwindowclass", "opcontainerclass") Then
            Console.WriteLine(vbTab & theText & " - Child - " & CStr(hWnd) & " - " & myClassName)
            ' Set focus on the RDP session (without affecting the users currently
            ' focused  window)
            'SendMessage(hWnd, WmSetfocus, 0, 0)
            ' Find the RDP window dimensions and figure out the center of it
            GetClientRect(hWnd, myRect)
            arrRect = Split(myRect.ToString(), ",")
            Console.WriteLine("Child Client Rect   : " & myRect.ToString())
            Console.WriteLine("Child Click Location: " & CInt(CInt(arrRect(2)) / 2) & " X " &
      CInt(CInt(arrRect(3)) / 2))
            coord = CInt(CInt(arrRect(2)) / 2) * &H10000 + CInt(CInt(arrRect(3)) / 2)
            ' Send the mouse move on the RDP session in the background (without
            ' affecting the users currently focused window).  Swapping the location
            ' each poll.
            If BoolMouseLoc Then
                SendMessage(hWnd, WmMousemove, 0, 0)
            Else
                SendMessage(hWnd, WmMousemove, 0, coord)
            End If
        End If
        Return True
    End Function
    ' Gathers the windows class names
    Public Function GetWindowClass(hwnd As Long) As String
        Dim sClassName As New StringBuilder("", 256)
        Call GetClassName(hwnd, sClassName, 256)
        Return sClassName.ToString
    End Function
    <Extension>
    Public Function MultiContains(str As String, ParamArray values() As String) As Boolean
        Return values.Any(Function(val) str.Contains(val))
    End Function
End Module
