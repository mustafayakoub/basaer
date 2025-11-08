#!/bin/bash
# ===============================================
# سكربت تثبيت أدوات استخراج الحواشي
# Installation script for footnote extraction tools
# ===============================================

echo "=================================="
echo "تثبيت أدوات استخراج الحواشي"
echo "Installing Footnote Tools"
echo "=================================="
echo ""

# التحقق من وجود Python
if ! command -v python3 &> /dev/null; then
    echo "❌ خطأ: Python 3 غير مثبت"
    echo "❌ Error: Python 3 is not installed"
    echo "يرجى تثبيت Python 3 أولاً / Please install Python 3 first"
    echo "https://www.python.org/downloads/"
    exit 1
fi

echo "✓ تم العثور على Python 3"
echo "✓ Python 3 found"
python3 --version
echo ""

# التحقق من وجود pip
if ! command -v pip3 &> /dev/null; then
    echo "❌ خطأ: pip غير مثبت"
    echo "❌ Error: pip is not installed"
    exit 1
fi

echo "✓ تم العثور على pip"
echo "✓ pip found"
echo ""

# إنشاء بيئة افتراضية (اختياري)
echo "هل تريد إنشاء بيئة افتراضية؟ (y/n)"
echo "Do you want to create a virtual environment? (y/n)"
read -r response

if [[ "$response" =~ ^([yY][eE][sS]|[yY])$ ]]; then
    echo "جاري إنشاء البيئة الافتراضية..."
    echo "Creating virtual environment..."
    python3 -m venv venv

    echo "تفعيل البيئة الافتراضية..."
    echo "Activating virtual environment..."
    source venv/bin/activate

    echo "✓ تم إنشاء وتفعيل البيئة الافتراضية"
    echo "✓ Virtual environment created and activated"
    echo ""
fi

# تثبيت المتطلبات
echo "جاري تثبيت المتطلبات..."
echo "Installing requirements..."
pip3 install -r requirements.txt

if [ $? -eq 0 ]; then
    echo ""
    echo "=================================="
    echo "✓ تم التثبيت بنجاح!"
    echo "✓ Installation completed successfully!"
    echo "=================================="
    echo ""
    echo "الاستخدام / Usage:"
    echo "python3 extract_footnotes.py input.docx"
    echo "python3 restore_footnotes.py output.docx"
    echo ""
else
    echo ""
    echo "❌ فشل التثبيت"
    echo "❌ Installation failed"
    exit 1
fi
