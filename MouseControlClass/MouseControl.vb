Imports System.Windows.Forms
Imports System.Drawing

<System.ComponentModel.Description("A user control to mouse control")>
<System.ComponentModel.DisplayName("Mouse Control")>
Public Class MouseControl

    Property MousePosition As Point
        Get
            Return Cursor.Position
        End Get
        Set(value As Point)
            Cursor.Position = value
        End Set
    End Property

    Property MousePositionAsPointOfCursor As PointOfCursor
        Get
            Return PointOfCursor.ConvertFromPoint(Cursor.Position)
        End Get
        Set(value As PointOfCursor)
            Cursor.Position = PointOfCursor.ConvertToPoint(value)
        End Set
    End Property
    Enum Mouse_State
        Down = 1
        Up = 2
    End Enum
    Enum Button
        Left = 1
        Right = 2
        Middle = 3
    End Enum
    Private Declare Sub mouse_event Lib "user32.dll" (ByVal dwFlags As Integer, ByVal dx As Integer, ByVal dy As Integer,
                                                      ByVal cButtons As Integer, ByVal dwExtraInfo As Integer)
    Private Sub Mouse_Click(ByVal button As Button, ByVal state As Mouse_State, XY As Point)
        Cursor.Position = XY
        Select Case button
            Case 1
                If state = Mouse_State.Down Then
                    mouse_event(2, 100, 100, 0, 0)
                Else
                    mouse_event(4, 100, 100, 0, 0)
                End If
            Case 2
                If state = Mouse_State.Down Then
                    mouse_event(8, 100, 100, 0, 0)
                Else
                    mouse_event(16, 100, 100, 0, 0)
                End If
            Case 3
                If state = Mouse_State.Down Then
                    mouse_event(32, 100, 100, 0, 0)
                Else
                    mouse_event(64, 100, 100, 0, 0)
                End If
        End Select
    End Sub
    Private Sub Mouse_Click(ByVal button As Button, ByVal state As Mouse_State)
        Select Case button
            Case 1
                If state = Mouse_State.Down Then
                    mouse_event(2, 100, 100, 0, 0)
                Else
                    mouse_event(4, 100, 100, 0, 0)
                End If
            Case 2
                If state = Mouse_State.Down Then
                    mouse_event(8, 100, 100, 0, 0)
                Else
                    mouse_event(16, 100, 100, 0, 0)
                End If
            Case 3
                If state = Mouse_State.Down Then
                    mouse_event(32, 100, 100, 0, 0)
                Else
                    mouse_event(64, 100, 100, 0, 0)
                End If
        End Select
    End Sub
    Private Sub Mouse_Click(ByVal button As Button, ByVal state As Mouse_State, Optional ByVal X As Integer = 0, Optional Y As Integer = 0)
        Cursor.Position = New Point(X, Y)
        Select Case button
            Case 1
                If state = Mouse_State.Down Then
                    mouse_event(2, 100, 100, 0, 0)
                Else
                    mouse_event(4, 100, 100, 0, 0)
                End If
            Case 2
                If state = Mouse_State.Down Then
                    mouse_event(8, 100, 100, 0, 0)
                Else
                    mouse_event(16, 100, 100, 0, 0)
                End If
            Case 3
                If state = Mouse_State.Down Then
                    mouse_event(32, 100, 100, 0, 0)
                Else
                    mouse_event(64, 100, 100, 0, 0)
                End If
        End Select
    End Sub
    Public Sub CursorPointChange(ByVal X As Integer, ByVal y As Integer)
        Cursor.Position = New Point(X, y)
    End Sub
    Public Sub CursorPointChange(XY As Point)
        Cursor.Position = XY
    End Sub
    Function CursorPoint() As Point
        Return System.Windows.Forms.Cursor.Position
    End Function
    Public Sub LeftButtonClick()
        Mouse_Click(Button.Left, Mouse_State.Down)
        Mouse_Click(Button.Left, Mouse_State.Up)
    End Sub
    Public Sub LeftButtonClick(P As Point)
        Mouse_Click(Button.Left, Mouse_State.Down, P)
        Mouse_Click(Button.Left, Mouse_State.Up)
    End Sub
    Public Sub LeftButtonClick(X As Integer, Y As Integer)
        Mouse_Click(Button.Left, Mouse_State.Down, X, Y)
        Mouse_Click(Button.Left, Mouse_State.Up)
    End Sub
    Public Sub RightButtonClick()
        Mouse_Click(Button.Right, Mouse_State.Down)
        Mouse_Click(Button.Right, Mouse_State.Up)
    End Sub
    Public Sub RightButtonClick(P As Point)
        Mouse_Click(Button.Right, Mouse_State.Down, P)
        Mouse_Click(Button.Right, Mouse_State.Up)
    End Sub
    Public Sub RightButtonClick(X As Integer, Y As Integer)
        Mouse_Click(Button.Right, Mouse_State.Down, X, Y)
        Mouse_Click(Button.Right, Mouse_State.Up)
    End Sub
    Public Sub MiddleButtonClick()
        Mouse_Click(Button.Middle, Mouse_State.Down)
        Mouse_Click(Button.Middle, Mouse_State.Up)
    End Sub
    Public Sub MiddleButtonClick(P As Point)
        Mouse_Click(Button.Middle, Mouse_State.Down, P)
        Mouse_Click(Button.Middle, Mouse_State.Up)
    End Sub
    Public Sub MiddleButtonClick(X As Integer, Y As Integer)
        Mouse_Click(Button.Middle, Mouse_State.Down, X, Y)
        Mouse_Click(Button.Middle, Mouse_State.Up)
    End Sub
    Public Sub LeftDown()
        Mouse_Click(Button.Left, Mouse_State.Down)
    End Sub
    Public Sub LeftDown(P As Point)
        Mouse_Click(Button.Left, Mouse_State.Down, P)
    End Sub
    Public Sub LeftDown(X As Integer, Y As Integer)
        Mouse_Click(Button.Left, Mouse_State.Down, X, Y)
    End Sub
    Public Sub LeftUp()
        Mouse_Click(Button.Left, Mouse_State.Up)
    End Sub
    Public Sub LeftUp(P As Point)
        Mouse_Click(Button.Left, Mouse_State.Up, P)
    End Sub
    Public Sub LeftUp(X As Integer, Y As Integer)
        Mouse_Click(Button.Left, Mouse_State.Up, X, Y)
    End Sub
    Public Sub RightDown()
        Mouse_Click(Button.Right, Mouse_State.Down)
    End Sub
    Public Sub RightDown(P As Point)
        Mouse_Click(Button.Right, Mouse_State.Down, P)
    End Sub
    Public Sub RightDown(X As Integer, Y As Integer)
        Mouse_Click(Button.Right, Mouse_State.Down, X, Y)
    End Sub
    Public Sub RightUp()
        Mouse_Click(Button.Right, Mouse_State.Up)
    End Sub
    Public Sub RightUp(P As Point)
        Mouse_Click(Button.Right, Mouse_State.Up, P)
    End Sub
    Public Sub RightUp(X As Integer, Y As Integer)
        Mouse_Click(Button.Right, Mouse_State.Up, X, Y)
    End Sub
    Public Sub MiddleDown()
        Mouse_Click(Button.Middle, Mouse_State.Down)
    End Sub
    Public Sub MiddleDown(P As Point)
        Mouse_Click(Button.Middle, Mouse_State.Down, P)
    End Sub
    Public Sub MiddleDown(X As Integer, Y As Integer)
        Mouse_Click(Button.Middle, Mouse_State.Down, X, Y)
    End Sub
    Public Sub MiddleUp()
        Mouse_Click(Button.Middle, Mouse_State.Up)
    End Sub
    Public Sub MiddleUp(P As Point)
        Mouse_Click(Button.Middle, Mouse_State.Up, P)
    End Sub
    Public Sub MiddleUp(X As Integer, Y As Integer)
        Mouse_Click(Button.Middle, Mouse_State.Up, X, Y)
    End Sub
    Structure PointOfCursor
        Dim X As Integer
        Dim Y As Integer

        Friend Shared Function ConvertToPoint(X As PointOfCursor) As Point
            Return New Point(X.X, X.Y)
        End Function

        Friend Shared Function ConvertFromPoint(X As Point) As PointOfCursor
            Return New PointOfCursor(X.X, X.Y)
        End Function

        Sub New(X As Integer, Y As Integer)
            Me.X = X
            Me.Y = Y
        End Sub

    End Structure
End Class

