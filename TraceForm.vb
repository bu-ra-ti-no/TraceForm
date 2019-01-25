''' <summary>
''' Trace output window.
''' </summary>
<System.ComponentModel.DesignerCategory("Code")> _
Friend Class TraceForm
    Inherits Form
    Private _maxLineCount As Integer
    Private WithEvents LV As ListView
    Private _traceDelegate As Action(Of String) = AddressOf InnerTrace

    Public Sub New()
        MyBase.New()
        FormBorderStyle = FormBorderStyle.SizableToolWindow
        ShowInTaskbar = False
        LV = New ListView With {
            .Font = New Font("Courier", 11, FontStyle.Regular, GraphicsUnit.Pixel),
            .BackColor = Color.LightGray,
            .ForeColor = Color.Black,
            .HeaderStyle = ColumnHeaderStyle.None,
            .Dock = DockStyle.Fill,
            .View = View.Details
        }
        LV.Columns.Add(String.Empty)
        MaxLineCount = 22
        Size = New Size(200, MaximumSize.Height)
        LV.GetType.GetProperty("DoubleBuffered",
                                      Reflection.BindingFlags.NonPublic Or
                                      Reflection.BindingFlags.Instance).
                                      SetValue(LV, True, Nothing)
        Controls.Add(LV)
    End Sub

    ''' <summary>
    ''' The maximum number of rows that are visible in the list.
    ''' </summary>
    Public Property MaxLineCount() As Integer
        Get
            Return _maxLineCount
        End Get
        Set(value As Integer)
            If value < 1 Then value = 1
            If value > 5000 Then value = 5000
            _maxLineCount = value
            Dim h As Integer
            If LV.Items.Count = 0
                LV.Items.Add(LV.ToString)
                h = LV.GetItemRect(0).Height
                LV.Items.Clear()
            Else
                h = LV.GetItemRect(0).Height
            End If
            MaximumSize = New Size(Screen.GetWorkingArea(Me).Width,
                                   value * h - ClientSize.Height +
                                   Height + LV.Height -
                                   LV.ClientSize.Height)
        End Set
    End Property

    ''' <summary>
    ''' Adds a new row to the list.
    ''' </summary>
    Public Sub Trace(message As String)
        If InvokeRequired
            Invoke(_traceDelegate, message)
        Else
            InnerTrace(message)
        End If
    End Sub

    Private Sub InnerTrace(message)
        LV.BeginUpdate()
        If LV.Items.Count = MaxLineCount
            LV.Items.RemoveAt(0)
        End If
        Dim i = LV.Items.Add(message)
        If message.StartsWith("*") Then i.ForeColor = Color.Crimson
        LV.EndUpdate()
    End Sub

    Private Sub ListView1_Resize() Handles LV.Resize
        LV.Columns(0).Width = Math.Max(LV.ClientSize.Width - 10, 5)
    End Sub
End Class
