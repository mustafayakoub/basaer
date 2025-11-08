Attribute VB_Name = "RestoreFootnotes"
' ===============================================
' ماكرو لإعادة الحواشي السفلية إلى مستند Word
' Macro to restore footnotes to Word document
' ===============================================
'
' الاستخدام / Usage:
' 1. افتح مستند Word المعدل / Open the edited Word document
' 2. اذهب إلى View > Macros / Go to View > Macros
' 3. اختر RestoreFootnotesFromText واضغط Run / Select RestoreFootnotesFromText and click Run
' 4. سيتم تحويل جميع النصوص بين && إلى حواشي سفلية / All text between && will be converted to footnotes
'
' ===============================================

Sub RestoreFootnotesFromText()
    '
    ' إعادة الحواشي السفلية من النص بين &&
    ' Restore footnotes from text between &&
    '
    Dim doc As Document
    Dim rng As Range
    Dim searchText As String
    Dim footnoteText As String
    Dim footnoteCount As Long
    Dim startTime As Double
    Dim startPos As Long
    Dim endPos As Long

    ' بدء المؤقت / Start timer
    startTime = Timer

    ' الحصول على المستند النشط / Get active document
    Set doc = ActiveDocument

    ' إيقاف تحديث الشاشة لتحسين الأداء / Disable screen updating for better performance
    Application.ScreenUpdating = False

    ' إيقاف التراجع / Disable undo
    Application.UndoRecord.StartCustomRecord "إعادة الحواشي"

    On Error GoTo ErrorHandler

    ' البحث عن النص بين && باستخدام التعبيرات النمطية / Search for text between && using wildcards
    footnoteCount = 0

    ' البدء من بداية المستند / Start from beginning of document
    Set rng = doc.Content
    rng.Collapse Direction:=wdCollapseStart

    ' إعداد البحث / Setup find
    With rng.Find
        .ClearFormatting
        .Text = " &&*&& "
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = True
        .MatchSoundsLike = False
        .MatchAllWordForms = False

        ' البحث المتكرر / Loop through all matches
        Do While .Execute
            ' الحصول على النص المطابق / Get matched text
            Dim matchedText As String
            matchedText = rng.Text

            ' استخراج نص الحاشية (إزالة &&) / Extract footnote text (remove &&)
            footnoteText = Mid(matchedText, 4, Len(matchedText) - 6)
            footnoteText = Trim(footnoteText)

            ' التحقق من أن النص ليس فارغًا / Check that text is not empty
            If Len(footnoteText) > 0 Then
                ' حفظ الموقع / Save position
                Dim savedStart As Long
                savedStart = rng.Start

                ' حذف النص / Delete the text
                rng.Delete

                ' تعيين النطاق إلى الموقع المحفوظ / Set range to saved position
                Set rng = doc.Range(Start:=savedStart, End:=savedStart)

                ' إضافة حاشية سفلية / Add footnote
                Dim fn As Footnote
                Set fn = doc.Footnotes.Add(Range:=rng, Text:=footnoteText)

                ' الانتقال إلى الموقع بعد الحاشية / Move to position after footnote
                Set rng = doc.Range(Start:=savedStart + 1, End:=doc.Content.End)
                rng.Collapse Direction:=wdCollapseStart

                footnoteCount = footnoteCount + 1

                ' تحديث شريط الحالة / Update status bar
                If footnoteCount Mod 10 = 0 Then
                    Application.StatusBar = "جاري إعادة الحواشي... " & footnoteCount
                End If
            Else
                ' الانتقال للبحث التالي / Move to next search
                rng.Collapse Direction:=wdCollapseEnd
            End If
        Loop
    End With

    ' إنهاء التراجع / End undo record
    Application.UndoRecord.EndCustomRecord

    ' إعادة تفعيل تحديث الشاشة / Re-enable screen updating
    Application.ScreenUpdating = True

    ' حساب الوقت المستغرق / Calculate elapsed time
    Dim elapsedTime As Double
    elapsedTime = Timer - startTime

    ' عرض رسالة النجاح / Show success message
    Application.StatusBar = False
    MsgBox "تمت الإعادة بنجاح!" & vbCrLf & vbCrLf & _
           "عدد الحواشي المستعادة: " & footnoteCount & vbCrLf & _
           "الوقت المستغرق: " & Format(elapsedTime, "0.00") & " ثانية" & vbCrLf & vbCrLf & _
           "Restoration completed successfully!" & vbCrLf & _
           "Footnotes restored: " & footnoteCount & vbCrLf & _
           "Time: " & Format(elapsedTime, "0.00") & " seconds", _
           vbInformation, "إعادة الحواشي / Restore Footnotes"

    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    Application.StatusBar = False
    MsgBox "حدث خطأ: " & Err.Description & vbCrLf & vbCrLf & _
           "Error: " & Err.Description, _
           vbCritical, "خطأ / Error"
End Sub

Sub RestoreFootnotesQuick()
    '
    ' نسخة سريعة من الماكرو (بدون رسائل تأكيد)
    ' Quick version of the macro (without confirmation messages)
    '
    Call RestoreFootnotesFromText
End Sub

' ===============================================
' دوال مساعدة / Helper Functions
' ===============================================

Function CountFootnoteMarkers() As Long
    '
    ' عد عدد علامات && في المستند
    ' Count the number of && markers in the document
    '
    Dim doc As Document
    Dim rng As Range
    Dim count As Long

    Set doc = ActiveDocument
    Set rng = doc.Content

    count = 0

    With rng.Find
        .ClearFormatting
        .Text = " &&*&& "
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchWildcards = True

        Do While .Execute
            count = count + 1
            rng.Collapse Direction:=wdCollapseEnd
        Loop
    End With

    CountFootnoteMarkers = count
End Function

Sub ShowFootnoteMarkersCount()
    '
    ' عرض عدد علامات الحواشي في المستند
    ' Display count of footnote markers in document
    '
    Dim count As Long
    count = CountFootnoteMarkers()

    MsgBox "عدد علامات الحواشي في المستند: " & count & vbCrLf & vbCrLf & _
           "Number of footnote markers in document: " & count, _
           vbInformation, "عدد الحواشي / Footnote Count"
End Sub
