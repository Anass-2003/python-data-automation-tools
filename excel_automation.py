import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


def create_master_dashboard():
    # 1. إنشاء ملف إكسيل جديد
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Master Dashboard"

    # إلغاء خطوط الشبكة الإضافية لجعل المظهر احترافياً
    ws.views.sheetView[0].showGridLines = True

    # 2. إضافة العنوان الرئيسي
    ws.merge_cells("A1:F1")
    ws["A1"] = "نظام إدارة مخازن موقع الإنشاءات - Master Dashboard (MEP & ELV)"
    ws["A1"].font = Font(name="Calibri", size=16, bold=True, color="FFFFFF")
    ws["A1"].fill = PatternFill(
        start_color="1F4E79", end_color="1F4E79", fill_type="solid"
    )
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 40

    # 3. إعداد عناوين الجدول
    headers = [
        "كود المادة",
        "اسم المادة (MEP / ELV)",
        "نوع الحركة",
        "الكمية",
        "الرصيد الحالي",
        "حالة المخزون",
    ]
    ws.append([])  # سطر فارغ
    ws.append(headers)

    header_fill = PatternFill(
        start_color="2F5597", end_color="2F5597", fill_type="solid"
    )
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")

    for col_num in range(1, 7):
        cell = ws.cell(row=3, column=col_num)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[3].height = 25

    # 4. إضافة بيانات تجريبية تلقائية
    sample_data = [
        ["ELV-001", "Cat6 Network Cable 305m", "صرف", 20, 5, "⚠️ نقص مخزون"],
        ["MEP-102", "Copper Pipe 3/4 inch", "إضافة", 100, 150, "✅ متوفر"],
        ["MEP-205", "Circuit Breaker 32A", "صرف", 15, 8, "⚠️ نقص مخزون"],
        ["ELV-009", "CCTV Camera Dome 4MP", "إضافة", 30, 45, "✅ متوفر"],
    ]

    thin_border = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9"),
    )

    alert_fill = PatternFill(
        start_color="FCE4D6", end_color="FCE4D6", fill_type="solid"
    )
    alert_font = Font(color="C00000", bold=True)

    for row_idx, row_data in enumerate(sample_data, start=4):
        ws.append(row_data)
        ws.row_dimensions[row_idx].height = 22
        for col_idx in range(1, 7):
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = thin_border

            # تمييز التنبيهات أوتوماتيكياً
            if "نقص" in str(ws.cell(row=row_idx, column=6).value):
                ws.cell(row=row_idx, column=6).fill = alert_fill
                ws.cell(row=row_idx, column=6).font = alert_font

    # 5. ضبط عرض الأعمدة أوتوماتيكياً
    for col in ws.columns:
        max_len = max(len(str(cell.value or "")) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 5, 18)

    # 6. حفظ الملف تلقائياً
    file_name = "Master_Management.xlsx"
    wb.save(file_name)
    print(
        f"✅ تم إنشاء وتنسيق ملف الإكسيل أوتوماتيكياً بنجاح! اسم الملف: {file_name}"
    )


if __name__ == "__main__":
    create_master_dashboard() 3 3" 