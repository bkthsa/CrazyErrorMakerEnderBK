Imports System.Runtime.InteropServices
Imports System.Windows.Forms

Public Class Form1
    ' Define sound resources
    Dim errorsound As String = "Windows_Foreground"
    Dim warninfosound As String = "Windows_Background"
    Dim notify As String = "Windows_Notify_System_Generic"
    Dim hardwareinsert As String = "Windows_Hardware_Insert"
    Dim hardwarermove As String = "Windows_Hardware_Remove"
    Dim hardwarefail As String = "Windows_Hardware_Fail"
    Dim shutdown As String = "Windows_Shutdown"

    ' API declarations for registering and handling hotkeys
    <DllImport("user32.dll", SetLastError:=True)>
    Public Shared Function RegisterHotKey(ByVal hwnd As IntPtr, ByVal id As Integer, ByVal fsModifiers As Integer, ByVal vk As Integer) As Boolean
    End Function

    <DllImport("user32.dll", SetLastError:=True)>
    Public Shared Function UnregisterHotKey(ByVal hwnd As IntPtr, ByVal id As Integer) As Boolean
    End Function

    ' Constants for modifiers
    Private Const MOD_NONE As Integer = 0
    Private Const MOD_ALT As Integer = &H1
    Private Const MOD_CONTROL As Integer = &H2
    Private Const MOD_SHIFT As Integer = &H4

    ' Hotkey IDs for "1", "2", "3", "4", "5", "6", "7", "8", "9", Spacebar, and Q, W, E
    Private Const HOTKEY_ID_1 As Integer = 1
    Private Const HOTKEY_ID_2 As Integer = 2
    Private Const HOTKEY_ID_3 As Integer = 3
    Private Const HOTKEY_ID_4 As Integer = 4
    Private Const HOTKEY_ID_5 As Integer = 5
    Private Const HOTKEY_ID_6 As Integer = 6
    Private Const HOTKEY_ID_7 As Integer = 7
    Private Const HOTKEY_ID_8 As Integer = 8
    Private Const HOTKEY_ID_9 As Integer = 9
    Private Const HOTKEY_ID_0 As Integer = 0
    Private Const HOTKEY_ID_L As Integer = 98 ' Hotkey ID for the "L" key
    Private Const HOTKEY_ID_SPACE As Integer = 10
    Private Const HOTKEY_ID_Q As Integer = 11 ' Q Key
    Private Const HOTKEY_ID_W As Integer = 12 ' W Key
    Private Const HOTKEY_ID_E As Integer = 13 ' E Key
    Private Const HOTKEY_ID_SEMICOLON As Integer = 34 ' Hotkey ID for the semicolon (;) key


    ' Virtual key codes for "1", "2", "3", "4", "5", "6", "7", "8", "9", Spacebar, Q, W, E
    Private Const VK_1 As Integer = &H31 ' Virtual key code for the "1" key
    Private Const VK_2 As Integer = &H32 ' Virtual key code for the "2" key
    Private Const VK_3 As Integer = &H33 ' Virtual key code for the "3" key
    Private Const VK_4 As Integer = &H34 ' Virtual key code for the "4" key
    Private Const VK_5 As Integer = &H35 ' Virtual key code for the "5" key
    Private Const VK_6 As Integer = &H36 ' Virtual key code for the "6" key
    Private Const VK_7 As Integer = &H37 ' Virtual key code for the "7" key
    Private Const VK_8 As Integer = &H38 ' Virtual key code for the "8" key
    Private Const VK_9 As Integer = &H39 ' Virtual key code for the "9" key
    Private Const VK_0 As Integer = &H30
    Private Const VK_SPACE As Integer = &H20 ' Virtual key code for the Spacebar
    Private Const VK_Q As Integer = &H51 ' Virtual key code for the "Q" key
    Private Const VK_W As Integer = &H57 ' Virtual key code for the "W" key
    Private Const VK_E As Integer = &H45 ' Virtual key code for the "E" key
    Private Const VK_L As Integer = &H4C ' Virtual key code for the "L" key
    Private Const VK_SEMICOLON As Integer = &HBA ' Virtual key code for the semicolon (;) key


    ' List of cloned forms
    Private clonedForms As New List(Of Form)()
    ' Track last cloned Form5
    Private lastClonedForm5 As Form = Nothing
    ' Form12 instance
    Private form12 As Form = Nothing
    Private isForm12Visible As Boolean = False

    ' Form load - Register hotkeys
    Private Sub Form1_Load(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Load
        MessageBox.Show("How to use: Error: 1,2,3,4,6,7,8,9. Notification: 5. Clear error: Spaces. Hardware Sound: Q,W,E. Shutdown: L", "CrazyErrorMaker by EnderBK", MessageBoxButtons.OK, MessageBoxIcon.Information)
        MessageBox.Show("NOTICE: This app registers key 1,2,3,4,5,6,7,8,9,SPACE,Q,W,E,L and runs in the background.", "CrazyErrorMaker by EnderBK", MessageBoxButtons.OK, MessageBoxIcon.Information)

        ' Register hotkeys for number keys, spacebar, and Q, W, E
        RegisterHotKey(Me.Handle, HOTKEY_ID_1, MOD_NONE, VK_1)
        RegisterHotKey(Me.Handle, HOTKEY_ID_2, MOD_NONE, VK_2)
        RegisterHotKey(Me.Handle, HOTKEY_ID_3, MOD_NONE, VK_3)
        RegisterHotKey(Me.Handle, HOTKEY_ID_4, MOD_NONE, VK_4)
        RegisterHotKey(Me.Handle, HOTKEY_ID_5, MOD_NONE, VK_5)
        RegisterHotKey(Me.Handle, HOTKEY_ID_6, MOD_NONE, VK_6)
        RegisterHotKey(Me.Handle, HOTKEY_ID_7, MOD_NONE, VK_7)
        RegisterHotKey(Me.Handle, HOTKEY_ID_8, MOD_NONE, VK_8)
        RegisterHotKey(Me.Handle, HOTKEY_ID_9, MOD_NONE, VK_9)
        RegisterHotKey(Me.Handle, HOTKEY_ID_0, MOD_NONE, VK_0)
        RegisterHotKey(Me.Handle, HOTKEY_ID_SPACE, MOD_NONE, VK_SPACE)


        ' Register hotkeys for Q, W, E (to show/hide Form12)
        RegisterHotKey(Me.Handle, HOTKEY_ID_Q, MOD_NONE, VK_Q)
        RegisterHotKey(Me.Handle, HOTKEY_ID_W, MOD_NONE, VK_W)
        RegisterHotKey(Me.Handle, HOTKEY_ID_E, MOD_NONE, VK_E)
        RegisterHotKey(Me.Handle, HOTKEY_ID_L, MOD_NONE, VK_L)
    End Sub

    ' Handle hotkey messages
    Protected Overrides Sub WndProc(ByRef m As Message)
        If m.Msg = &H312 Then ' Hotkey message (WM_HOTKEY)
            Select Case m.WParam.ToInt32()
                ' Handle the number and spacebar hotkeys
                Case HOTKEY_ID_1
                    CloneFormAtCursor(Me)
                    GlobalFunctions.PlaySoundFromResources(errorsound)
                Case HOTKEY_ID_2
                    GlobalFunctions.PlaySoundFromResources(warninfosound)
                    CloneFormAtCursor(Form2)
                Case HOTKEY_ID_3
                    GlobalFunctions.PlaySoundFromResources(warninfosound)
                    CloneFormAtCursor(Form3)
                Case HOTKEY_ID_4
                    GlobalFunctions.PlaySoundFromResources(warninfosound)
                    CloneFormAtCursor(Form4)
                Case HOTKEY_ID_5
                    CloneFormAtBottomRight(Form5)
                    GlobalFunctions.PlaySoundFromResources(notify)
                Case HOTKEY_ID_6
                    CloneFormAtCursor(Form6)
                    GlobalFunctions.PlaySoundFromResources(warninfosound)
                Case HOTKEY_ID_7
                    GlobalFunctions.PlaySoundFromResources(errorsound)
                    CloneFormAtCursor(Form7)
                Case HOTKEY_ID_8
                    GlobalFunctions.PlaySoundFromResources(warninfosound)
                    CloneFormAtCursor(Form8)
                Case HOTKEY_ID_9
                    GlobalFunctions.PlaySoundFromResources(errorsound)
                    CloneFormAtCursor(Form9)
                Case HOTKEY_ID_0
                    CloneFormAtCursor(Form10)
                Case HOTKEY_ID_SPACE
                    ClearAllClonedForms()
                    Form11.Hide()

                    ' Handle the Q, W, E keys to show/hide Form12
                Case HOTKEY_ID_Q
                    GlobalFunctions.PlaySoundFromResources(hardwareinsert)
                    If form12 Is Nothing Then
                        form12 = New Form12() ' Create Form12 if it doesn't exist
                    End If
                    ShowForm12()
                Case HOTKEY_ID_W
                    HideForm12()
                    GlobalFunctions.PlaySoundFromResources(hardwarermove)
                Case HOTKEY_ID_E
                    HideForm12()
                    GlobalFunctions.PlaySoundFromResources(hardwarefail)
                Case HOTKEY_ID_L
                    GlobalFunctions.PlaySoundFromResources(shutdown)
                    Form11.Show()
                Case HOTKEY_ID_SEMICOLON
                    Form11.Close()
            End Select
        End If
        MyBase.WndProc(m)
    End Sub

    ' Show Form12 at the top-right corner
    Private Sub ShowForm12()
        If form12 IsNot Nothing AndAlso Not isForm12Visible Then
            form12.Show()
            ' Get screen size (without taskbar)
            Dim screenHeight As Integer = Screen.PrimaryScreen.WorkingArea.Height
            Dim screenWidth As Integer = Screen.PrimaryScreen.WorkingArea.Width
            form12.Location = New Point(screenWidth - form12.Width, 0)
            isForm12Visible = True
        End If
    End Sub

    ' Hide Form12
    Private Sub HideForm12()
        If form12 IsNot Nothing AndAlso isForm12Visible Then
            form12.Hide()
            isForm12Visible = False
        End If
    End Sub

    ' Clone Form at cursor position
    Private Sub CloneFormAtCursor(ByVal formToClone As Form)
        If formToClone Is Nothing Then Return

        ' Create a new form (target form) to clone the controls into
        Dim targetForm As New Form()

        ' Clone the specified form (formToClone) to the target form
        FormClonerModule.CloneForm(formToClone, targetForm)

        ' Set the cloned form's location at the cursor position
        targetForm.Location = Cursor.Position
        targetForm.Show()

        ' Add the cloned form to the list
        clonedForms.Add(targetForm)
    End Sub

    ' Clone Form5 at bottom-right corner
    Private Sub CloneFormAtBottomRight(ByVal formToClone As Form)
        If formToClone Is Nothing Then Return

        ' Create a new form (target form) to clone the controls into
        Dim targetForm As New Form()

        ' Clone the specified form (formToClone) to the target form
        FormClonerModule.CloneForm(formToClone, targetForm)

        ' Check if there is already a cloned Form5 and adjust the position if needed
        If lastClonedForm5 IsNot Nothing AndAlso Not lastClonedForm5.IsDisposed Then
            ' Move the new cloned Form5 25 pixels higher
            Dim newLocation As Point = lastClonedForm5.Location
            newLocation.Y -= 130
            targetForm.Location = newLocation
        Else
            ' Position the cloned form at the bottom-right corner just above the taskbar
            Dim screenHeight As Integer = Screen.PrimaryScreen.WorkingArea.Height
            Dim screenWidth As Integer = Screen.PrimaryScreen.WorkingArea.Width
            targetForm.Location = New Point(screenWidth - targetForm.Width, screenHeight - targetForm.Height - -55)
        End If

        ' Show the cloned form
        targetForm.Show()

        ' Update the last cloned Form5 reference
        lastClonedForm5 = targetForm

        ' Add the cloned form to the list of cloned forms
        clonedForms.Add(targetForm)

        ' Set a timer to automatically close the cloned form after 5 seconds
        Dim timer As New System.Windows.Forms.Timer() ' Explicitly use System.Windows.Forms.Timer
        AddHandler timer.Tick, Sub(sender As Object, e As EventArgs)
                                   targetForm.Close()
                                   clonedForms.Remove(targetForm)
                                   timer.Stop()
                               End Sub
        timer.Interval = 2000 ' 5 seconds
        timer.Start()
    End Sub

    ' Clear all cloned forms
    Private Sub ClearAllClonedForms()
        For Each clonedForm In clonedForms
            clonedForm.Close()
        Next
        clonedForms.Clear()
    End Sub

    ' Form close - Unregister hotkeys
    Protected Overrides Sub OnFormClosing(ByVal e As FormClosingEventArgs)
        ' Unregister all hotkeys when the form is closed
        UnregisterHotKey(Me.Handle, HOTKEY_ID_1)
        UnregisterHotKey(Me.Handle, HOTKEY_ID_2)
        UnregisterHotKey(Me.Handle, HOTKEY_ID_3)
        UnregisterHotKey(Me.Handle, HOTKEY_ID_4)
        UnregisterHotKey(Me.Handle, HOTKEY_ID_5)
        UnregisterHotKey(Me.Handle, HOTKEY_ID_6)
        UnregisterHotKey(Me.Handle, HOTKEY_ID_7)
        UnregisterHotKey(Me.Handle, HOTKEY_ID_8)
        UnregisterHotKey(Me.Handle, HOTKEY_ID_9)
        UnregisterHotKey(Me.Handle, HOTKEY_ID_0)
        UnregisterHotKey(Me.Handle, HOTKEY_ID_L)
        UnregisterHotKey(Me.Handle, HOTKEY_ID_SPACE)
        UnregisterHotKey(Me.Handle, HOTKEY_ID_Q)
        UnregisterHotKey(Me.Handle, HOTKEY_ID_W)
        UnregisterHotKey(Me.Handle, HOTKEY_ID_E)
        UnregisterHotKey(Me.Handle, HOTKEY_ID_SEMICOLON)
        MyBase.OnFormClosing(e)
    End Sub

    Private Sub PictureBox1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox1.Click

    End Sub
End Class
