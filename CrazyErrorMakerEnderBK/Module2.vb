Imports System.Media
Imports System.Threading

Module GlobalFunctions

    ' Play sound from resources without blocking the previous sound using separate threads
    Public Sub PlaySoundFromResources(ByVal resourceName As String)
        ' Access the embedded WAV file from resources as a stream
        Dim sound As IO.Stream = My.Resources.ResourceManager.GetStream(resourceName)

        If sound IsNot Nothing Then
            ' Start a new thread to play the sound asynchronously without blocking the previous sound
            Dim playThread As New Thread(Sub()
                                             Dim player As New SoundPlayer(sound)
                                             player.Play() ' Use Play() for non-blocking asynchronous playback
                                         End Sub)
            playThread.IsBackground = True ' This will make the thread a background thread
            playThread.Start() ' Start the thread to play sound
        Else
            MessageBox.Show("Sound file not found.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub

End Module
