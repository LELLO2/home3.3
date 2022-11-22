Imports Microsoft.VisualBasic.FileIO

Public Class Form1

    Private fileName As String = ""
    Private tables As List(Of List(Of String)) = New List(Of List(Of String))()
    Private headers As List(Of String) = New List(Of String)()



    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        If Not (Me.fileName = "") Then

            Using parser As TextFieldParser = New TextFieldParser(fileName)
                parser.TextFieldType = FieldType.Delimited
                parser.SetDelimiters(",")
                Dim i As Integer = 0

                While Not parser.EndOfData
                    Dim fields As String() = parser.ReadFields()

                    If headers.Count = 0 Then

                        For Each field As String In fields
                            Me.tables.Add(New List(Of String)())
                            Me.tables(i).Add(field)
                            Me.headers.Add(field)
                            i += 1
                        Next
                    Else

                        For Each field As String In fields
                            Me.tables.Add(New List(Of String)())
                            Me.tables(i).Add(field)
                            i += 1
                        Next
                    End If

                    i = 0
                    Me.RichTextBox1.AppendText("riuscito\n" & vbCrLf)
                End While

                Me.ListBox1.DataSource = headers
            End Using
        End If
    End Sub

    Private Function allLower(ByVal temp As List(Of String)) As List(Of String)
        Dim list As List(Of String) = New List(Of String)()

        For Each field As String In temp
            list.Add(field.ToLower())
        Next

        Return list
    End Function

    Private Sub compute_row(ByVal tempList As List(Of String))
        Dim currentRow As List(Of String) = allLower(tempList)
        currentRow.RemoveAt(0)
        Dim rows As List(Of List(Of String)) = New List(Of List(Of String))()

        For Each row As String In currentRow

            If Not (rows.Exists(Function(e) e.Contains(row))) Then
                rows.Add(currentRow.FindAll(Function(s) s.Equals(row)))
            End If
        Next

        For Each elem As List(Of String) In rows

            If elem.First() <> "" Then
                Me.RichTextBox2.AppendText("""" & elem.First() & """ occurs: " + elem.Count().ToString() & " time" & vbCrLf)
            ElseIf elem.Count() > 1 Then
                Me.RichTextBox2.AppendText("There are " & elem.Count() & " blank values." & vbCrLf)
            Else
                Me.RichTextBox2.AppendText("There is " & elem.Count() & " blank value." & vbCrLf)
            End If
        Next
    End Sub

    Private Sub ListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox1.SelectedIndexChanged

        Dim i As Integer = Me.ListBox1.SelectedIndex
        Dim elem_list As List(Of String) = Me.tables.ElementAt(i)
        Me.RichTextBox1.Text = ""
        Me.RichTextBox2.Text = ""

        For Each field As String In elem_list
            Me.RichTextBox1.AppendText(field & " " & vbCrLf)
        Next

        compute_row(elem_list)

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.OpenFileDialog1.ShowDialog()
        Me.fileName = Me.OpenFileDialog1.FileName
    End Sub
End Class
