' Main Form: Form1.vb
Imports System.Drawing
Imports System.Runtime.InteropServices
Imports System.Threading
Imports System.Windows.Forms

Public Class Form1
    ' Windows API declarations
    <DllImport("user32.dll", CharSet:=CharSet.Auto, CallingConvention:=CallingConvention.StdCall)>
    Public Shared Sub mouse_event(ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
    End Sub

    <DllImport("user32.dll")>
    Private Shared Function GetCursorPos(ByRef lpPoint As Point) As Boolean
    End Function

    <DllImport("user32.dll")>
    Private Shared Function GetDC(ByVal hwnd As IntPtr) As IntPtr
    End Function

    <DllImport("user32.dll")>
    Private Shared Function ReleaseDC(ByVal hwnd As IntPtr, ByVal hdc As IntPtr) As Integer
    End Function

    <DllImport("gdi32.dll")>
    Private Shared Function GetPixel(ByVal hdc As IntPtr, ByVal nXPos As Integer, ByVal nYPos As Integer) As UInteger
    End Function

    <DllImport("user32.dll")>
    Private Shared Function RegisterHotKey(ByVal hwnd As IntPtr, ByVal id As Integer, ByVal fsModifiers As Integer, ByVal vk As Integer) As Boolean
    End Function

    <DllImport("user32.dll")>
    Private Shared Function UnregisterHotKey(ByVal hwnd As IntPtr, ByVal id As Integer) As Boolean
    End Function

    ' Constants for mouse events
    Private Const MOUSEEVENTF_LEFTDOWN As Long = &H2
    Private Const MOUSEEVENTF_LEFTUP As Long = &H4
    Private Const MOUSEEVENTF_RIGHTDOWN As Long = &H8
    Private Const MOUSEEVENTF_RIGHTUP As Long = &H10

    ' Hotkey constants
    Private Const WM_HOTKEY As Integer = &H312
    Private Const MOD_NONE As Integer = 0
    Private Const VK_F6 As Integer = &H75
    Private Const VK_F7 As Integer = &H76
    Private Const VK_F8 As Integer = &H77

    ' Form variables
    Private clickingEnabled As Boolean = False
    Private clickTimer As New System.Windows.Forms.Timer()
    Private colorDetectionTimer As New System.Windows.Forms.Timer()
    Private targetColor As Color = Color.Red
    Private colorTolerance As Integer = 10
    Private clickPosition As Point = New Point(100, 100)
    Private isPositionSet As Boolean = False
    Private colorDetectionEnabled As Boolean = False
    Private colorDetectionActive As Boolean = False

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        SetupForm()
        SetupTimers()
        RegisterHotkeys()
        UpdateUI()
    End Sub

    Private Sub SetupForm()
        Me.Text = "Advanced Auto-Clicker"
        Me.Size = New Size(450, 600)
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.FormBorderStyle = FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.TopMost = True ' Always stay on top like Task Manager

        ' Create controls
        CreateControls()
    End Sub

    Private Sub CreateControls()
        ' Interval GroupBox
        Dim grpInterval As New GroupBox()
        grpInterval.Text = "Click Settings"
        grpInterval.Location = New Point(10, 10)
        grpInterval.Size = New Size(410, 120)

        Dim lblInterval As New Label()
        lblInterval.Text = "Click Interval (ms):"
        lblInterval.Location = New Point(10, 25)

        Dim nudInterval As New NumericUpDown()
        nudInterval.Name = "nudInterval"
        nudInterval.Location = New Point(120, 22)
        nudInterval.Size = New Size(80, 20)
        nudInterval.Minimum = 1
        nudInterval.Maximum = 10000
        nudInterval.Value = 1000

        Dim lblClickType As New Label()
        lblClickType.Text = "Click Type:"
        lblClickType.Location = New Point(220, 25)

        Dim cmbClickType As New ComboBox()
        cmbClickType.Name = "cmbClickType"
        cmbClickType.Location = New Point(290, 22)
        cmbClickType.Size = New Size(100, 20)
        cmbClickType.DropDownStyle = ComboBoxStyle.DropDownList
        cmbClickType.Items.AddRange({"Left Click", "Right Click"})
        cmbClickType.SelectedIndex = 0

        Dim btnSetPosition As New Button()
        btnSetPosition.Name = "btnSetPosition"
        btnSetPosition.Text = "Set Click Position"
        btnSetPosition.Location = New Point(10, 55)
        btnSetPosition.Size = New Size(120, 25)
        AddHandler btnSetPosition.Click, AddressOf BtnSetPosition_Click

        Dim lblPosition As New Label()
        lblPosition.Name = "lblPosition"
        lblPosition.Text = "Position: Not Set"
        lblPosition.Location = New Point(140, 60)
        lblPosition.Size = New Size(150, 15)

        Dim btnStartStop As New Button()
        btnStartStop.Name = "btnStartStop"
        btnStartStop.Text = "Start Clicking (F6)"
        btnStartStop.Location = New Point(300, 55)
        btnStartStop.Size = New Size(100, 25)
        btnStartStop.BackColor = Color.LightGreen
        AddHandler btnStartStop.Click, AddressOf BtnStartStop_Click

        grpInterval.Controls.AddRange({lblInterval, nudInterval, lblClickType, cmbClickType, btnSetPosition, lblPosition, btnStartStop})

        ' Color Detection GroupBox
        Dim grpColor As New GroupBox()
        grpColor.Text = "Color Detection"
        grpColor.Location = New Point(10, 140)
        grpColor.Size = New Size(410, 150)

        Dim chkColorDetection As New CheckBox()
        chkColorDetection.Name = "chkColorDetection"
        chkColorDetection.Text = "Enable Color Detection"
        chkColorDetection.Location = New Point(10, 25)
        AddHandler chkColorDetection.CheckedChanged, AddressOf ChkColorDetection_CheckedChanged

        Dim lblTargetColor As New Label()
        lblTargetColor.Text = "Target Color:"
        lblTargetColor.Location = New Point(10, 55)

        Dim btnColorPicker As New Button()
        btnColorPicker.Name = "btnColorPicker"
        btnColorPicker.Text = "Pick Color"
        btnColorPicker.Location = New Point(90, 52)
        btnColorPicker.Size = New Size(80, 23)
        btnColorPicker.BackColor = targetColor
        AddHandler btnColorPicker.Click, AddressOf BtnColorPicker_Click

        Dim lblTolerance As New Label()
        lblTolerance.Text = "Tolerance:"
        lblTolerance.Location = New Point(190, 55)

        Dim nudTolerance As New NumericUpDown()
        nudTolerance.Name = "nudTolerance"
        nudTolerance.Location = New Point(250, 52)
        nudTolerance.Size = New Size(60, 20)
        nudTolerance.Minimum = 0
        nudTolerance.Maximum = 255
        nudTolerance.Value = colorTolerance

        Dim lblColorPosition As New Label()
        lblColorPosition.Text = "Detection Position:"
        lblColorPosition.Location = New Point(10, 85)

        Dim btnSetColorPosition As New Button()
        btnSetColorPosition.Name = "btnSetColorPosition"
        btnSetColorPosition.Text = "Set Position"
        btnSetColorPosition.Location = New Point(120, 82)
        btnSetColorPosition.Size = New Size(80, 23)
        AddHandler btnSetColorPosition.Click, AddressOf BtnSetColorPosition_Click

        Dim lblColorPos As New Label()
        lblColorPos.Name = "lblColorPos"
        lblColorPos.Text = "Pos: 100, 100"
        lblColorPos.Location = New Point(210, 87)

        Dim lblColorStatus As New Label()
        lblColorStatus.Name = "lblColorStatus"
        lblColorStatus.Text = "Status: Inactive"
        lblColorStatus.Location = New Point(10, 115)
        lblColorStatus.Size = New Size(200, 15)

        grpColor.Controls.AddRange({chkColorDetection, lblTargetColor, btnColorPicker, lblTolerance, nudTolerance, lblColorPosition, btnSetColorPosition, lblColorPos, lblColorStatus})

        ' Status GroupBox
        Dim grpStatus As New GroupBox()
        grpStatus.Text = "Status & Information"
        grpStatus.Location = New Point(10, 300)
        grpStatus.Size = New Size(410, 120)

        Dim lblStatus As New Label()
        lblStatus.Name = "lblStatus"
        lblStatus.Text = "Status: Ready"
        lblStatus.Location = New Point(10, 25)
        lblStatus.Size = New Size(200, 15)

        Dim lblClickCount As New Label()
        lblClickCount.Name = "lblClickCount"
        lblClickCount.Text = "Clicks Performed: 0"
        lblClickCount.Location = New Point(10, 45)

        Dim lblHotkey As New Label()
        lblHotkey.Text = "Hotkeys: F6 (Toggle), F7 (Setup), F8 (Force Stop)"
        lblHotkey.Location = New Point(10, 65)
        lblHotkey.Size = New Size(350, 15)

        Dim btnReset As New Button()
        btnReset.Name = "btnReset"
        btnReset.Text = "Reset Count"
        btnReset.Location = New Point(10, 85)
        btnReset.Size = New Size(80, 25)
        AddHandler btnReset.Click, AddressOf BtnReset_Click

        Dim lblCurrentPos As New Label()
        lblCurrentPos.Name = "lblCurrentPos"
        lblCurrentPos.Text = "Current Cursor: Press F7"
        lblCurrentPos.Location = New Point(200, 45)
        lblCurrentPos.Size = New Size(150, 15)

        grpStatus.Controls.AddRange({lblStatus, lblClickCount, lblHotkey, lblCurrentPos, btnReset})

        ' Add all groups to form
        Me.Controls.AddRange({grpInterval, grpColor, grpStatus})
    End Sub

    Private Sub SetupTimers()
        ' Click timer
        clickTimer.Interval = 1000
        AddHandler clickTimer.Tick, AddressOf ClickTimer_Tick

        ' Color detection timer
        colorDetectionTimer.Interval = 100 ' Check every 100ms
        AddHandler colorDetectionTimer.Tick, AddressOf ColorDetectionTimer_Tick
    End Sub

    Private Sub RegisterHotkeys()
        RegisterHotKey(Me.Handle, 1, MOD_NONE, VK_F6)
        RegisterHotKey(Me.Handle, 2, MOD_NONE, VK_F7)
        RegisterHotKey(Me.Handle, 3, MOD_NONE, VK_F8)
    End Sub

    Protected Overrides Sub WndProc(ByRef m As Message)
        MyBase.WndProc(m)
        If m.Msg = WM_HOTKEY Then
            If m.WParam.ToInt32() = 1 Then
                ' F6 pressed - Toggle clicking
                ToggleClicking()
            ElseIf m.WParam.ToInt32() = 2 Then
                ' F7 pressed - Get cursor coordinates
                GetCurrentCursorPosition()
            ElseIf m.WParam.ToInt32() = 3 Then
                ' F8 pressed - Force stop clicking
                ForceStopClicking()
            End If
        End If
    End Sub

    Private clickCount As Integer = 0

    Private Sub ClickTimer_Tick(sender As Object, e As EventArgs)
        If clickingEnabled AndAlso isPositionSet Then
            PerformClick()
            clickCount += 1
            UpdateClickCount()
        End If
    End Sub

    Private Sub ColorDetectionTimer_Tick(sender As Object, e As EventArgs)
        If colorDetectionEnabled Then
            Try
                Dim currentColor As Color = GetPixelColor(clickPosition.X, clickPosition.Y)
                Dim colorMatch As Boolean = ColorsMatch(currentColor, targetColor, colorTolerance)

                UpdateColorStatus(colorMatch, currentColor)

                If colorMatch AndAlso Not colorDetectionActive Then
                    ' Color detected, start clicking
                    colorDetectionActive = True
                    StartClicking()
                ElseIf Not colorMatch AndAlso colorDetectionActive Then
                    ' Color not detected, stop clicking
                    colorDetectionActive = False
                    StopClicking()
                End If
            Catch ex As Exception
                ' Handle any errors in color detection
                UpdateColorStatus(False, Color.Black)
            End Try
        Else
            ' If color detection is disabled, stop the timer
            colorDetectionTimer.Stop()
        End If
    End Sub

    Private Sub PerformClick()
        Dim cmbClickType As ComboBox = DirectCast(Me.Controls.Find("cmbClickType", True)(0).Parent.Controls.Find("cmbClickType", True)(0), ComboBox)

        If cmbClickType.SelectedIndex = 0 Then
            ' Left click
            mouse_event(MOUSEEVENTF_LEFTDOWN, clickPosition.X, clickPosition.Y, 0, 0)
            mouse_event(MOUSEEVENTF_LEFTUP, clickPosition.X, clickPosition.Y, 0, 0)
        Else
            ' Right click
            mouse_event(MOUSEEVENTF_RIGHTDOWN, clickPosition.X, clickPosition.Y, 0, 0)
            mouse_event(MOUSEEVENTF_RIGHTUP, clickPosition.X, clickPosition.Y, 0, 0)
        End If
    End Sub

    Private Function GetPixelColor(x As Integer, y As Integer) As Color
        Try
            Dim hdc As IntPtr = GetDC(IntPtr.Zero)
            Dim pixel As UInteger = GetPixel(hdc, x, y)
            ReleaseDC(IntPtr.Zero, hdc)

            ' Extract RGB components safely
            Dim r As Integer = CInt(pixel And &HFF)
            Dim g As Integer = CInt((pixel And &HFF00) >> 8)
            Dim b As Integer = CInt((pixel And &HFF0000) >> 16)

            Return Color.FromArgb(r, g, b)
        Catch ex As Exception
            ' Return a default color if pixel reading fails
            Return Color.Black
        End Try
    End Function

    Private Function ColorsMatch(color1 As Color, color2 As Color, tolerance As Integer) As Boolean
        Try
            Return Math.Abs(CInt(color1.R) - CInt(color2.R)) <= tolerance AndAlso
                   Math.Abs(CInt(color1.G) - CInt(color2.G)) <= tolerance AndAlso
                   Math.Abs(CInt(color1.B) - CInt(color2.B)) <= tolerance
        Catch ex As Exception
            Return False
        End Try
    End Function

    Private Sub GetCurrentCursorPosition()
        Dim currentPos As Point
        GetCursorPos(currentPos)

        ' Set this as the click position
        clickPosition = currentPos
        isPositionSet = True

        ' Update the current position label
        Dim lblCurrentPos As Label = DirectCast(Me.Controls.Find("lblCurrentPos", True)(0).Parent.Controls.Find("lblCurrentPos", True)(0), Label)
        lblCurrentPos.Text = $"Current Cursor: {currentPos.X}, {currentPos.Y}"

        ' Also get the color at that position for reference
        Dim currentColor As Color = GetPixelColor(currentPos.X, currentPos.Y)

        ' Store the color for color detection
        targetColor = currentColor

        ' Update the color picker button to show the new color
        Dim btnColorPicker As Button = DirectCast(Me.Controls.Find("btnColorPicker", True)(0).Parent.Controls.Find("btnColorPicker", True)(0), Button)
        btnColorPicker.BackColor = targetColor

        ' Update UI to reflect new click position
        UpdateUI()
        UpdateColorPositionLabel()

        ' Show a temporary message with detailed info
        Dim message As String = $"Position: {currentPos.X}, {currentPos.Y}" & vbCrLf &
                               $"Color captured: R{currentColor.R}, G{currentColor.G}, B{currentColor.B}" & vbCrLf &
                               $"Hex Color: #{currentColor.R:X2}{currentColor.G:X2}{currentColor.B:X2}" & vbCrLf & vbCrLf &
                               "✓ Click position set!" & vbCrLf &
                               "✓ Color saved for detection!" & vbCrLf &
                               "✓ Detection position updated!"

        MessageBox.Show(message, "Complete Setup - F7", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    Private Sub BtnStartStop_Click(sender As Object, e As EventArgs)
        ToggleClicking()
    End Sub

    Private Sub ToggleClicking()
        If clickingEnabled Then
            StopClicking()
        Else
            StartClicking()
        End If
    End Sub

    Private Sub StartClicking()
        If Not isPositionSet Then
            MessageBox.Show("Please set a click position first!", "Position Required", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        clickingEnabled = True
        Dim nudInterval As NumericUpDown = DirectCast(Me.Controls.Find("nudInterval", True)(0).Parent.Controls.Find("nudInterval", True)(0), NumericUpDown)
        clickTimer.Interval = CInt(nudInterval.Value)
        clickTimer.Start()
        UpdateUI()
    End Sub

    Private Sub StopClicking()
        clickingEnabled = False
        clickTimer.Stop()
        UpdateUI()
    End Sub

    Private Sub ForceStopClicking()
        ' Force stop all clicking activities permanently
        clickingEnabled = False
        colorDetectionActive = False
        colorDetectionEnabled = False

        ' Stop all timers
        clickTimer.Stop()
        colorDetectionTimer.Stop()

        ' Disable color detection checkbox
        Dim chkColorDetection As CheckBox = DirectCast(Me.Controls.Find("chkColorDetection", True)(0).Parent.Controls.Find("chkColorDetection", True)(0), CheckBox)
        chkColorDetection.Checked = False

        UpdateUI()

        ' Show confirmation message
        MessageBox.Show("Auto-clicker completely stopped!" & vbCrLf & "All automatic functions disabled.", "F8 - Force Stop", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    Private Sub BtnSetPosition_Click(sender As Object, e As EventArgs)
        Me.WindowState = FormWindowState.Minimized
        Thread.Sleep(1000) ' Give user time to position cursor

        GetCursorPos(clickPosition)
        isPositionSet = True

        Me.WindowState = FormWindowState.Normal
        UpdateUI()
    End Sub

    Private Sub BtnSetColorPosition_Click(sender As Object, e As EventArgs)
        Me.WindowState = FormWindowState.Minimized
        Thread.Sleep(1000)

        GetCursorPos(clickPosition)

        Me.WindowState = FormWindowState.Normal
        UpdateColorPositionLabel()
    End Sub

    Private Sub BtnColorPicker_Click(sender As Object, e As EventArgs)
        Dim colorDialog As New ColorDialog()
        colorDialog.Color = targetColor

        If colorDialog.ShowDialog() = DialogResult.OK Then
            targetColor = colorDialog.Color
            Dim btnColorPicker As Button = DirectCast(sender, Button)
            btnColorPicker.BackColor = targetColor
        End If
    End Sub

    Private Sub ChkColorDetection_CheckedChanged(sender As Object, e As EventArgs)
        Dim chk As CheckBox = DirectCast(sender, CheckBox)
        colorDetectionEnabled = chk.Checked

        If colorDetectionEnabled Then
            colorDetectionTimer.Start()
        Else
            colorDetectionTimer.Stop()
            colorDetectionActive = False
            If clickingEnabled AndAlso colorDetectionActive Then
                StopClicking()
            End If
        End If

        UpdateUI()
    End Sub

    Private Sub ChkAlwaysOnTop_CheckedChanged(sender As Object, e As EventArgs)
        Dim chk As CheckBox = DirectCast(sender, CheckBox)
        Me.TopMost = chk.Checked
    End Sub

    Private Sub BtnReset_Click(sender As Object, e As EventArgs)
        clickCount = 0
        UpdateClickCount()
    End Sub

    Private Sub UpdateUI()
        Dim btnStartStop As Button = DirectCast(Me.Controls.Find("btnStartStop", True)(0).Parent.Controls.Find("btnStartStop", True)(0), Button)
        Dim lblStatus As Label = DirectCast(Me.Controls.Find("lblStatus", True)(0).Parent.Controls.Find("lblStatus", True)(0), Label)
        Dim lblPosition As Label = DirectCast(Me.Controls.Find("lblPosition", True)(0).Parent.Controls.Find("lblPosition", True)(0), Label)

        If clickingEnabled Then
            btnStartStop.Text = "Stop Clicking (F6)"
            btnStartStop.BackColor = Color.LightCoral
            lblStatus.Text = "Status: Clicking Active"
        Else
            btnStartStop.Text = "Start Clicking (F6)"
            btnStartStop.BackColor = Color.LightGreen
            lblStatus.Text = "Status: Ready"
        End If

        If isPositionSet Then
            lblPosition.Text = $"Position: {clickPosition.X}, {clickPosition.Y}"
        Else
            lblPosition.Text = "Position: Not Set"
        End If
    End Sub

    Private Sub UpdateClickCount()
        Dim lblClickCount As Label = DirectCast(Me.Controls.Find("lblClickCount", True)(0).Parent.Controls.Find("lblClickCount", True)(0), Label)
        lblClickCount.Text = $"Clicks Performed: {clickCount}"
    End Sub

    Private Sub UpdateColorPositionLabel()
        Dim lblColorPos As Label = DirectCast(Me.Controls.Find("lblColorPos", True)(0).Parent.Controls.Find("lblColorPos", True)(0), Label)
        lblColorPos.Text = $"Pos: {clickPosition.X}, {clickPosition.Y}"
    End Sub

    Private Sub UpdateColorStatus(colorMatch As Boolean, currentColor As Color)
        Dim lblColorStatus As Label = DirectCast(Me.Controls.Find("lblColorStatus", True)(0).Parent.Controls.Find("lblColorStatus", True)(0), Label)

        If colorDetectionEnabled Then
            If colorMatch Then
                lblColorStatus.Text = "Status: Color Detected - Active"
                lblColorStatus.ForeColor = Color.Green
            Else
                lblColorStatus.Text = $"Status: Monitoring - Current: R{currentColor.R},G{currentColor.G},B{currentColor.B}"
                lblColorStatus.ForeColor = Color.Blue
            End If
        Else
            lblColorStatus.Text = "Status: Inactive"
            lblColorStatus.ForeColor = Color.Black
        End If
    End Sub

    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        UnregisterHotKey(Me.Handle, 1)
        UnregisterHotKey(Me.Handle, 2)
        UnregisterHotKey(Me.Handle, 3)
        clickTimer.Stop()
        colorDetectionTimer.Stop()
    End Sub
End Class