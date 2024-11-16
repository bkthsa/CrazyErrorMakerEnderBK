Public Class Form11

    ' Form Resize event to keep the PictureBox centered and resized
    Private Sub Form11_Resize(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Resize
        ' Center PictureBox1 in the middle of the form
        CenterPictureBox()

        ' Resize the PictureBox to maintain the GIF's aspect ratio and fit in the form
        ResizePictureBox()
    End Sub

    ' Method to center the PictureBox in the middle of the form
    Private Sub CenterPictureBox()
        ' Calculate the position to center the PictureBox in the middle of the form
        Dim centerX As Integer = (Me.ClientSize.Width - PictureBox1.Width) \ 2
        Dim centerY As Integer = (Me.ClientSize.Height - PictureBox1.Height) \ 2

        ' Set the PictureBox's location to the calculated center point
        PictureBox1.Location = New Point(centerX, centerY)
    End Sub

    ' Method to resize the PictureBox based on form size and aspect ratio
    Private Sub ResizePictureBox()
        ' Check if the PictureBox has an image (GIF or any image)
        If PictureBox1.Image IsNot Nothing Then
            ' Get the original dimensions of the image
            Dim gifWidth As Integer = PictureBox1.Image.Width
            Dim gifHeight As Integer = PictureBox1.Image.Height

            ' Calculate the aspect ratio (width / height)
            Dim aspectRatio As Double = gifWidth / gifHeight

            ' Resize the PictureBox (scale down by 80% of the form's width)
            ' Ensure the scaling maintains aspect ratio
            Dim newWidth As Integer = CInt(Me.ClientSize.Width * 0.5) ' Resize to 80% of the form's width
            Dim newHeight As Integer = CInt(newWidth / aspectRatio)

            ' If the calculated height exceeds the form's height, adjust it
            If newHeight > Me.ClientSize.Height Then
                newHeight = CInt(Me.ClientSize.Height * 0.5) ' Resize to 80% of the form's height
                newWidth = CInt(newHeight * aspectRatio)
            End If

            ' Apply the new size to the PictureBox
            PictureBox1.Size = New Size(newWidth, newHeight)

            ' Keep the PictureBox centered after resizing
            CenterPictureBox()
        End If
    End Sub

    ' Form Load event to initialize the PictureBox
    Private Sub Form11_Load(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Load
        ' Set the PictureBox SizeMode to Zoom to maintain the GIF's aspect ratio
        PictureBox1.SizeMode = PictureBoxSizeMode.Zoom

        ' Center the PictureBox when the form first loads
        CenterPictureBox()

        ' Resize the PictureBox to fit the GIF when the form is first loaded
        ResizePictureBox()
    End Sub
End Class
