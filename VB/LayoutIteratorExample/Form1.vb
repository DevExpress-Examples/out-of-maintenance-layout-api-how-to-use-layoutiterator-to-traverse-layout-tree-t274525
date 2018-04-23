Imports DevExpress.XtraRichEdit.API.Layout
Imports DevExpress.XtraRichEdit.API.Native
Imports System
Imports System.Windows.Forms

Namespace LayoutIteratorExample
    Partial Public Class Form1
        Inherits DevExpress.XtraBars.Ribbon.RibbonForm

        Private layoutIterator As LayoutIterator
        Private doc As SubDocument
        Private coloredRange As DocumentRange

        Public Sub New()
            InitializeComponent()
            AddHandler rgLevel.EditValueChanged, AddressOf rgLevel_EditValueChanged

            repositoryItemLayoutLevel.Items.AddRange(System.Enum.GetValues(GetType(LayoutLevel)))
            cmbLayoutLevel.EditValue = LayoutLevel.Box
            cmbLayoutLevel.Enabled = False

            richEditControl1.LoadDocument("Test.docx")
            AddHandler richEditControl1.DocumentLoaded, AddressOf richEditControl1_DocumentLoaded
            doc = richEditControl1.Document
        End Sub

        Private Sub richEditControl1_DocumentLoaded(ByVal sender As Object, ByVal e As EventArgs)
            layoutIterator = New LayoutIterator(richEditControl1.DocumentLayout)
        End Sub

        Private Sub btnMoveNext_ItemClick(ByVal sender As Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles barButtonItem1.ItemClick
'            #Region "#MoveNext"
            Dim result As Boolean = False
            Dim s As String = String.Empty

            ' Create a new iterator if the document has been changed and the layout is updated.
            If Not layoutIterator.IsLayoutValid Then
                CreateNewIterator()
            End If

            Select Case barEditItemRgLevel.EditValue.ToString()
                Case "Any"
                    result = layoutIterator.MoveNext()
                Case "Level"
                    result = layoutIterator.MoveNext(CType(cmbLayoutLevel.EditValue, LayoutLevel))
                Case "LevelWithinParent"
                    result = layoutIterator.MoveNext(CType(cmbLayoutLevel.EditValue, LayoutLevel), False)
            End Select

            If Not result Then
                s = "Cannot move."
                If layoutIterator.IsStart Then
                    s &= ControlChars.Lf & "Start is reached"
                ElseIf layoutIterator.IsEnd Then
                    s &= ControlChars.Lf & "End is reached"
                End If
                MessageBox.Show(s)
            End If
'            #End Region ' #MoveNext
            UpdateInfoAndSelection()
        End Sub


        Private Sub btnMovePrev_ItemClick(ByVal sender As Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles btnMovePrev.ItemClick
'            #Region "#MovePrev"
            Dim result As Boolean = False
            Dim s As String = String.Empty
            ' Create a new iterator if the document has been changed and the layout is updated.
            If Not layoutIterator.IsLayoutValid Then
                CreateNewIterator()
            End If

            Select Case barEditItemRgLevel.EditValue.ToString()
                Case "Any"
                    result = layoutIterator.MovePrevious()
                Case "Level"
                    result = layoutIterator.MovePrevious(CType(cmbLayoutLevel.EditValue, LayoutLevel))
                Case "LevelWithinParent"
                    result = layoutIterator.MovePrevious(CType(cmbLayoutLevel.EditValue, LayoutLevel), False)
            End Select

            If Not result Then
                s = "Cannot move."
                If layoutIterator.IsStart Then
                    s &= ControlChars.Lf & "Start is reached."
                ElseIf layoutIterator.IsEnd Then
                    s &= ControlChars.Lf & "End is reached."
                End If
                    MessageBox.Show(s)
            End If
'            #End Region ' #MovePrev
            UpdateInfoAndSelection()
        End Sub

        Private Sub btnStartOver_ItemClick(ByVal sender As Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles btnStartOver.ItemClick
            If coloredRange IsNot Nothing Then
                ResetRange(coloredRange)
            End If
            CreateNewIterator()
        End Sub

        Private Sub CreateNewIterator()
            layoutIterator = New LayoutIterator(richEditControl1.DocumentLayout)
            doc = richEditControl1.Document
            UpdateInfoAndSelection()
            MessageBox.Show("Layout is modified, creating a new iterator.")
        End Sub

        Private Sub btnStartHere_ItemClick(ByVal sender As Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles btnStartHere.ItemClick
            If coloredRange IsNot Nothing Then
                ResetRange(coloredRange)
            End If

            doc = richEditControl1.Document.CaretPosition.BeginUpdateDocument()
            richEditControl1.Document.ChangeActiveDocument(doc)
            layoutIterator = New LayoutIterator(richEditControl1.DocumentLayout, doc.Range)

            Dim el As RangedLayoutElement = richEditControl1.DocumentLayout.GetElement(richEditControl1.Document.CaretPosition, LayoutType.PlainTextBox)
            Do
                Dim element As RangedLayoutElement = TryCast(layoutIterator.Current, RangedLayoutElement)
                If (element IsNot Nothing) AndAlso (element.Equals(el)) Then
                    UpdateInfoAndSelection()
                    Return
                End If
            Loop While layoutIterator.MoveNext()
        End Sub

        Private Sub btnSetRange_ItemClick(ByVal sender As Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles btnSetRange.ItemClick
            If coloredRange IsNot Nothing Then
                ResetRange(coloredRange)
            End If

            coloredRange = richEditControl1.Document.Selection
            If coloredRange.Length = 0 Then
                Return
            End If

            ' Highlight selected range.
            Dim d As SubDocument = coloredRange.BeginUpdateDocument()
            Dim cp As CharacterProperties = d.BeginUpdateCharacters(coloredRange)
            cp.BackColor = System.Drawing.Color.Yellow
            d.EndUpdateCharacters(cp)
            coloredRange.EndUpdateDocument(d)

            ' Create a new iterator limited to the specified range.
            layoutIterator = New LayoutIterator(richEditControl1.DocumentLayout, coloredRange)

            doc = coloredRange.BeginUpdateDocument()
            richEditControl1.Document.ChangeActiveDocument(doc)
            coloredRange.EndUpdateDocument(doc)

            ' Select the first element in the highlighted range.
            Dim el As RangedLayoutElement = richEditControl1.DocumentLayout.GetElement(coloredRange.Start, LayoutType.PlainTextBox)
            Do While layoutIterator.MoveNext()
                Dim element As RangedLayoutElement = TryCast(layoutIterator.Current, RangedLayoutElement)
                If (element IsNot Nothing) AndAlso (element.Equals(el)) Then
                    UpdateInfoAndSelection()
                    Return
                End If
            Loop
        End Sub

        Private Sub UpdateInfoAndSelection()
            Dim element As LayoutElement = layoutIterator.Current
            infoElement.Caption = String.Empty
            If element IsNot Nothing Then
                Dim rangedElement As RangedLayoutElement = TryCast(element, RangedLayoutElement)
                infoElement.Caption = element.Type.ToString()
                If rangedElement IsNot Nothing Then
                    Dim r As DocumentRange = doc.CreateRange(rangedElement.Range.Start, rangedElement.Range.Length)
                    richEditControl1.Document.ChangeActiveDocument(doc)
                    richEditControl1.Document.Selection = r
                End If
            End If
        End Sub

        Private Sub ResetRange(ByVal r As DocumentRange)
            Dim d As SubDocument = r.BeginUpdateDocument()
            Dim cp As CharacterProperties = d.BeginUpdateCharacters(r)
            cp.BackColor = System.Drawing.Color.White
            d.EndUpdateCharacters(cp)
            r.EndUpdateDocument(d)
        End Sub

        Private Sub rgLevel_EditValueChanged(ByVal sender As Object, ByVal e As EventArgs)
            Dim val As String = DirectCast(sender, DevExpress.XtraEditors.RadioGroup).EditValue.ToString()
            If val = "Any" Then
                cmbLayoutLevel.Enabled = False
            Else
                cmbLayoutLevel.Enabled = True
            End If
        End Sub

    End Class
End Namespace
