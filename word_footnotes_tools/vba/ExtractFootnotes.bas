Attribute VB_Name = "ExtractFootnotes"
' ===============================================
' ماكرو لاستخراج الحواشي السفلية من مستند Word
' Macro to extract footnotes from Word document
' ===============================================
'
' الاستخدام / Usage:
' 1. افتح مستند Word / Open a Word document
' 2. اذهب إلى View > Macros / Go to View > Macros
' 3. اختر ExtractFootnotesToText واضغط Run / Select ExtractFootnotesToText and click Run
' 4. سيتم تحويل جميع الحواشي إلى نص بين && / All footnotes will be converted to text between &&
'
' ===============================================

Sub ExtractFootnotesToText()
    '
    ' استخراج الحواشي السفلية وتحويلها إلى نص بين &&
    ' Extract footnotes and convert them to text between &&
    '
    Dim doc As Document
    Dim fn As Footnote
    Dim en As Endnote
    Dim rng As Range
    Dim footnoteText As String
    Dim footnoteCount As Long
    Dim endnoteCount As Long
    Dim startTime As Double

    ' بدء المؤقت / Start timer
    startTime = Timer

    ' الحصول على المستند النشط / Get active document
    Set doc = ActiveDocument

    ' إيقاف تحديث الشاشة لتحسين الأداء / Disable screen updating for better performance
    Application.ScreenUpdating = False

    ' إيقاف التراجع / Disable undo
    Application.UndoRecord.StartCustomRecord "استخراج الحواشي"

    On Error GoTo ErrorHandler

    ' معالجة الحواشي السفلية / Process footnotes
    footnoteCount = 0
    Do While doc.Footnotes.Count > 0
        Set fn = doc.Footnotes(1)

        ' الحصول على نص الحاشية / Get footnote text
        footnoteText = fn.Range.Text

        ' إزالة علامات الفقرة الزائدة / Remove extra paragraph marks
        footnoteText = Replace(footnoteText, vbCr, " ")
        footnoteText = Replace(footnoteText, vbLf, " ")
        footnoteText = Trim(footnoteText)

        ' الحصول على موقع مرجع الحاشية / Get footnote reference location
        Set rng = fn.Reference

        ' حذف الحاشية / Delete the footnote
        fn.Delete

        ' إدراج النص بين && / Insert text between &&
        rng.Text = " &&" & footnoteText & "&& "

        footnoteCount = footnoteCount + 1

        ' تحديث شريط الحالة / Update status bar
        If footnoteCount Mod 10 = 0 Then
            Application.StatusBar = "جاري معالجة الحواشي... " & footnoteCount
        End If
    Loop

    ' معالجة الهوامش النهائية / Process endnotes
    endnoteCount = 0
    Do While doc.Endnotes.Count > 0
        Set en = doc.Endnotes(1)

        ' الحصول على نص الهامش / Get endnote text
        footnoteText = en.Range.Text

        ' إزالة علامات الفقرة الزائدة / Remove extra paragraph marks
        footnoteText = Replace(footnoteText, vbCr, " ")
        footnoteText = Replace(footnoteText, vbLf, " ")
        footnoteText = Trim(footnoteText)

        ' الحصول على موقع مرجع الهامش / Get endnote reference location
        Set rng = en.Reference

        ' حذف الهامش / Delete the endnote
        en.Delete

        ' إدراج النص بين && / Insert text between &&
        rng.Text = " &&" & footnoteText & "&& "

        endnoteCount = endnoteCount + 1

        ' تحديث شريط الحالة / Update status bar
        If endnoteCount Mod 10 = 0 Then
            Application.StatusBar = "جاري معالجة الهوامش النهائية... " & endnoteCount
        End If
    Loop

    ' إنهاء التراجع / End undo record
    Application.UndoRecord.EndCustomRecord

    ' إعادة تفعيل تحديث الشاشة / Re-enable screen updating
    Application.ScreenUpdating = True

    ' حساب الوقت المستغرق / Calculate elapsed time
    Dim elapsedTime As Double
    elapsedTime = Timer - startTime

    ' عرض رسالة النجاح / Show success message
    Application.StatusBar = False
    MsgBox "تم الاستخراج بنجاح!" & vbCrLf & vbCrLf & _
           "عدد الحواشي السفلية: " & footnoteCount & vbCrLf & _
           "عدد الهوامش النهائية: " & endnoteCount & vbCrLf & _
           "الوقت المستغرق: " & Format(elapsedTime, "0.00") & " ثانية" & vbCrLf & vbCrLf & _
           "Extraction completed successfully!" & vbCrLf & _
           "Footnotes: " & footnoteCount & vbCrLf & _
           "Endnotes: " & endnoteCount & vbCrLf & _
           "Time: " & Format(elapsedTime, "0.00") & " seconds", _
           vbInformation, "استخراج الحواشي / Extract Footnotes"

    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    Application.StatusBar = False
    MsgBox "حدث خطأ: " & Err.Description & vbCrLf & vbCrLf & _
           "Error: " & Err.Description, _
           vbCritical, "خطأ / Error"
End Sub

Sub ExtractFootnotesQuick()
    '
    ' نسخة سريعة من الماكرو (بدون رسائل تأكيد)
    ' Quick version of the macro (without confirmation messages)
    '
    Call ExtractFootnotesToText
End Sub
