Imports System.Drawing
Imports System.Windows.Forms

Module FormClonerModule
    ' Method to clone all controls from the source form to the destination form
    Public Sub CloneForm(ByVal originalForm As Form, ByVal targetForm As Form)
        ' Set the start position to Manual, so that Location can be manually set
        targetForm.StartPosition = FormStartPosition.Manual

        ' Clone the form properties (title, size, etc.)
        CloneFormProperties(originalForm, targetForm)

        ' Clear any existing controls on the target form
        targetForm.Controls.Clear()

        ' Loop through each control in the original form
        For Each control As Control In originalForm.Controls
            ' Create a new control of the same type as the original control
            Dim clonedControl As Control = CloneControl(control)

            ' Set the cloned control's location and size to match the original
            clonedControl.Location = control.Location
            clonedControl.Size = control.Size
            clonedControl.Anchor = control.Anchor
            clonedControl.TabIndex = control.TabIndex

            ' Add the cloned control to the target form
            targetForm.Controls.Add(clonedControl)
        Next

        ' Set the position of the cloned form to the cursor position on screen
        Dim cursorPosition As Point = Cursor.Position
        targetForm.Location = cursorPosition
    End Sub

    ' Helper method to clone the form's properties (such as title, size, etc.)
    Private Sub CloneFormProperties(ByVal originalForm As Form, ByVal targetForm As Form)
        ' Ensure both forms are not Nothing
        If originalForm Is Nothing OrElse targetForm Is Nothing Then
            Throw New ArgumentNullException("Both originalForm and targetForm must be non-null.")
        End If

        ' Clone form's basic properties
        targetForm.Text = originalForm.Text  ' Clone the form's title
        targetForm.ClientSize = originalForm.ClientSize  ' Clone the client size (form size excluding borders)
        targetForm.StartPosition = FormStartPosition.Manual  ' Set StartPosition to Manual to avoid interference
        targetForm.WindowState = originalForm.WindowState  ' Clone the window state (Maximized, Minimized, Normal)
        targetForm.AutoScaleMode = originalForm.AutoScaleMode  ' Clone the auto scale mode
        targetForm.AutoScaleDimensions = originalForm.AutoScaleDimensions  ' Clone AutoScaleDimensions if set
        targetForm.FormBorderStyle = originalForm.FormBorderStyle  ' Clone the form border style
        targetForm.Icon = originalForm.Icon  ' Clone the form's icon if set
        targetForm.MaximizeBox = originalForm.MaximizeBox  ' Clone the maximize box state
        targetForm.MinimizeBox = originalForm.MinimizeBox  ' Clone the minimize box state
        targetForm.HelpButton = originalForm.HelpButton  ' Clone the HelpButton property
        targetForm.ShowIcon = originalForm.ShowIcon
        targetForm.ShowInTaskbar = originalForm.ShowInTaskbar

        ' Make sure the cloned form supports transparency
        targetForm.AllowTransparency = originalForm.AllowTransparency
        targetForm.BackColor = Color.Lime
        targetForm.TransparencyKey = Color.Lime ' Make the background transparent
    End Sub


    ' Helper method to clone a single control
    Private Function CloneControl(ByVal originalControl As Control) As Control
        Dim clonedControl As Control = Nothing

        ' Clone based on the type of the control
        Select Case originalControl.GetType()
            Case GetType(Button)
                ' Clone Button
                Dim originalButton As Button = DirectCast(originalControl, Button)
                clonedControl = New Button()
                clonedControl.Text = originalButton.Text
                clonedControl.BackColor = Color.Transparent ' Make button background transparent
                clonedControl.ForeColor = originalButton.ForeColor

            Case GetType(TextBox)
                ' Clone TextBox
                Dim originalTextBox As TextBox = DirectCast(originalControl, TextBox)
                clonedControl = New TextBox()
                clonedControl.Text = originalTextBox.Text
                clonedControl.BackColor = Color.Transparent ' Make textbox background transparent
                clonedControl.ForeColor = originalTextBox.ForeColor

            Case GetType(Label)
                ' Clone Label
                Dim originalLabel As Label = DirectCast(originalControl, Label)
                clonedControl = New Label()
                clonedControl.Text = originalLabel.Text
                clonedControl.BackColor = Color.Transparent ' Make label background transparent
                clonedControl.ForeColor = originalLabel.ForeColor

            Case GetType(CheckBox)
                ' Clone CheckBox
                Dim originalCheckBox As CheckBox = DirectCast(originalControl, CheckBox)
                clonedControl = New CheckBox()
                ' Clone specific properties for CheckBox
                DirectCast(clonedControl, CheckBox).Checked = originalCheckBox.Checked
                clonedControl.Text = originalCheckBox.Text
                clonedControl.BackColor = Color.Transparent ' Make checkbox background transparent
                clonedControl.ForeColor = originalCheckBox.ForeColor

            Case GetType(PictureBox)
                ' Clone PictureBox
                Dim originalPicBox As PictureBox = DirectCast(originalControl, PictureBox)
                clonedControl = New PictureBox()
                ' Clone specific properties for PictureBox
                DirectCast(clonedControl, PictureBox).SizeMode = originalPicBox.SizeMode
                clonedControl.BackColor = Color.Transparent ' Make picturebox background transparent
                ' Clone the image as a new Bitmap to avoid shared references
                DirectCast(clonedControl, PictureBox).Image = If(originalPicBox.Image IsNot Nothing, New Bitmap(originalPicBox.Image), Nothing)

            Case Else
                ' If the control type isn't handled, just copy it directly
                clonedControl = originalControl
        End Select

        ' Return the cloned control
        Return clonedControl
    End Function
End Module
