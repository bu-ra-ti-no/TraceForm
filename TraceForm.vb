''' <summary>
''' Trace output window.
''' </summary>
<System.ComponentModel.DesignerCategory("Code")> _
Friend Class TraceForm
    Inherits Form
    Private _MaxLineCount As Integer
    Private WithEvents LV As ListView
    Private TraceDelegate As Action(Of String) = AddressOf Trace

    Public Sub New()
        MyBase.New()
        FormBorderStyle = Windows.Forms.FormBorderStyle.SizableToolWindow
        LV = New ListView
        LV.Font = New Font("Courier", 11, FontStyle.Regular, GraphicsUnit.Pixel)
        LV.BackColor = Color.LightGray
        LV.ForeColor = Color.Black
        LV.HeaderStyle = ColumnHeaderStyle.None
        LV.Columns.Add(String.Empty)
        LV.Dock = DockStyle.Fill
        LV.View = View.Details
        MaxLineCount = 22
        Size = New Size(200, MaximumSize.Height)
        LV.GetType.GetProperty("DoubleBuffered", _
                                      Reflection.BindingFlags.NonPublic _
                                      Or Reflection.BindingFlags.Instance). _
                                      SetValue(LV, True, Nothing)
        Controls.Add(LV)
    End Sub

    ''' <summary>
    ''' The maximum number of rows that are visible in the list.
    ''' </summary>
    Public Property MaxLineCount() As Integer
        Get
            Return _MaxLineCount
        End Get
        Set(ByVal value As Integer)
            If value < 1 Then value = 1
            If value > 5000 Then value = 5000
            _MaxLineCount = value
            Dim h As Integer
            If LV.Items.Count = 0 Then
                LV.Items.Add(LV.ToString)
                h = LV.GetItemRect(0).Height
                LV.Items.Clear()
            Else
                h = LV.GetItemRect(0).Height
            End If
            MaximumSize = New Size(Screen.GetWorkingArea(Me).Width, _
                                   value * h - ClientSize.Height + _
                                   Height + LV.Height - _
                                   LV.ClientSize.Height)
        End Set
    End Property

    ''' <summary>
    ''' Adds a new row to the list.
    ''' </summary>
    Public Sub Trace(ByVal Message As String)
        LV.BeginUpdate()
        If LV.Items.Count = MaxLineCount Then
            LV.Items.RemoveAt(0)
        End If
        Dim i = LV.Items.Add(Message)
        If Message.StartsWith("*") Then i.ForeColor = Color.Crimson
        LV.EndUpdate()
    End Sub

    ''' <summary>
    ''' Adds a new row to the list (not from main thread).
    ''' </summary>
    Public Sub TraceSafe(ByVal Message As String)
        Invoke(TraceDelegate, Message)
    End Sub

    Private Sub ListView1_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles LV.Resize
        LV.Columns(0).Width = Math.Max(LV.ClientSize.Width - 10, 5)
    End Sub
End Class

