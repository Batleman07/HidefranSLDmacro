'On 64 bit version of SolidWorks 2006-2012, the VBA form opens behind the SolidWorks _
application window. (This issue is fixed in SolidWorks 2013 and later.) This code _
prevents that from occurring. Note that the default user form name used in this code _
is UserForm1. You will need to change it to the name of your user form.

'This code will not cause problems on SolidWorks 2013 and later.

'This is a heavily modified version of the macro provided in S-013887
'Modified by Keith Rice
'CADSharp LLC
'www.cadsharp.com

Option Explicit

#If VBA7 Then

Sub main()
'Show form
UserForm1.Show vbModeless
End Sub

#Else

'Determines whether a file is an executable (.exe) file, and if so, which subsystem runs the executable file.
Declare Function GetBinaryType Lib "kernel32" Alias "GetBinaryTypeA" (ByVal sFileName As String, ByRef BinType As Long) As Long

'Retrieves a handle to the top-level window whose class name and window name match the specified strings.
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

'Changes the size, position, and Z order of a child, pop-up, or top-level window.
Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

'Used to get the placement of a window
Public Declare Function GetWindowPlacement Lib "user32" (ByVal hWnd As Long, ByRef lpwndpl As WINDOWPLACEMENT) As Long

Public Type RECT
Left As Long
Top As Long
Right As Long
Bottom As Long
End Type

Public Type POINTAPI
x As Long
y As Long
End Type

Public Type WINDOWPLACEMENT
Length As Long
FLAGS As Long
showCmd As Long
ptMinPosition As POINTAPI
ptMaxPosition As POINTAPI
rcNormalPosition As RECT
End Type

Sub main()
Dim lBinaryType As Long
Dim hWnd As Long     'window handle
Dim hWndSolidWorks As Long 'solidworks window handle
Dim swApp As SldWorks.SldWorks
Dim swFrame As SldWorks.Frame
Dim sWinPlacementSolidWorks As WINDOWPLACEMENT
Dim sWinPlacementForm As WINDOWPLACEMENT

Dim lSolidWorksXval As Long
Dim lSolidWorksYval As Long
Dim lFormXval As Long
Dim lFormYval As Long
Dim lXval As Long
Dim lYval As Long
Dim lHeightSolidWorks As Long

Set swApp = Application.SldWorks

'This function returns lBinaryType as 6 for x64 and 0 for x86 (32 bit)
GetBinaryType swApp.GetExecutablePath & "\sldworks.exe", lBinaryType

'Show form
UserForm1.Show vbModeless

'If x64 make form on top
If lBinaryType = 6 Then
'Find SolidWorks window placement
Set swFrame = swApp.Frame
hWndSolidWorks = swFrame.GetHWnd    'note that IFrame::GetHWndx64 is not supported
GetWindowPlacement hWndSolidWorks, sWinPlacementSolidWorks
lSolidWorksXval = sWinPlacementSolidWorks.rcNormalPosition.Right / 2
lSolidWorksYval = sWinPlacementSolidWorks.rcNormalPosition.Bottom / 2

'Find form window placement
hWnd = FindWindow(vbNullString, UserForm1.Caption)
GetWindowPlacement hWnd, sWinPlacementForm
lFormXval = sWinPlacementForm.rcNormalPosition.Right / 2
lFormYval = sWinPlacementForm.rcNormalPosition.Bottom / 2

'X and Y locations for form
lXval = lSolidWorksXval - lFormXval
lYval = lSolidWorksYval - lFormYval

'Move the form
'See this article for information on SetWindowPos arguments: _
http://msdn.microsoft.com/en-us/library/windows/desktop/ms633545(v=vs.85).aspx
'The flags are currently set for SWP_NOSIZE (1)
SetWindowPos hWnd, -1, lXval, lYval, 0, 0, 1
End If
End Sub

#End If