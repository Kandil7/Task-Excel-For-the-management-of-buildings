import openpyxl
from openpyxl.styles import PatternFill, Border, Side, Alignment, Font
from openpyxl.utils import get_column_letter
from datetime import date

# إنشاء مصنف Excel جديد
wb = openpyxl.Workbook()

#  #  #  # دوال  لتنسيق  الخلايا  #  #  #  #
def set_cell_style(cell, font_size=12, bold=False, bg_color=None, align="center"):
    """
    تطبيق تنسيقات على الخلية.

    Args:
        cell: خلية Excel المراد تنسيقها.
        font_size (int, optional): حجم الخط. Defaults to 12.
        bold (bool, optional):  جعل الخط غامقًا. Defaults to False.
        bg_color (str, optional): لون خلفية الخلية (كود سداسي عشري). Defaults to None.
        align (str, optional): محاذاة النص (left, center, right). Defaults to "center".
    """

    cell.font = Font(size=font_size, bold=bold)
    if bg_color:
        cell.fill = PatternFill(start_color=bg_color, end_color=bg_color, fill_type="solid")
    cell.border = Border(left=Side(style='thin'), 
                         right=Side(style='thin'), 
                         top=Side(style='thin'), 
                         bottom=Side(style='thin'))
    cell.alignment = Alignment(horizontal=align, vertical='center')


#  #  #  #  بيانات  نموذجية  #  #  #  # 
buildings = [
    {'name': 'عمارة أ', 'units': 5, 'notes': ''},
    {'name': 'عمارة ب', 'units': 14, 'notes': ''}
]

units = [
    {'unit_no': '101', 'building': 'عمارة أ', 'type': 'استديو', 'rent': 2000, 'status': 'مُؤجّرة', 'notes': ''},
    {'unit_no': '102', 'building': 'عمارة أ', 'type': 'شقة غرفة وصالة', 'rent': 3500, 'status': 'شاغرة', 'notes': 'تحتاج إلى صيانة'},
    {'unit_no': '201', 'building': 'عمارة ب', 'type': 'شقة غرفتين وصالة', 'rent': 4500, 'status': 'مُؤجّرة', 'notes': ''}
]

tenants = [
    {'unit_no': '101', 'name': 'محمد أحمد', 'id': '1234567890', 'mobile': '05xxxxxxxx', 'start_date': date(2024,1,1), 'end_date': date(2024,12,31), 'rent': 2000, 'email': 'mohamed@example.com', 'notes': ''},
    {'unit_no': '102', 'name': 'سارة خالد', 'id': '9876543210', 'mobile': '05xxxxxxxx', 'start_date': date(2024,2,15), 'end_date': date(2025,2,14), 'rent': 3500, 'email': 'sara@example.com', 'notes': ''}
]

rents_paid = [
    {'unit_no': '101', 'month': 'يناير', 'year': 2024, 'amount': 2000, 'date': date(2024, 1, 5), 'method': 'تحويل بنكي', 'status': 'مدفوع', 'notes': ''},
    {'unit_no': '102', 'month': 'فبراير', 'year': 2024, 'amount': 3500, 'date': date(2024, 2, 10), 'method': 'نقداً', 'status': 'مدفوع', 'notes': ''},
    {'unit_no': '101', 'month': 'فبراير', 'year': 2024, 'amount': 2000, 'date': None, 'method': '', 'status': 'غير مدفوع', 'notes': ''}
]

expenses = [
    {'building': 'عمارة أ', 'date': date(2024, 1, 1), 'type': 'فاتورة كهرباء', 'amount': 500, 'category': 'فواتير', 'notes': ''},
    {'building': 'عمارة ب', 'date': date(2024, 1, 15), 'type': 'صيانة', 'amount': 200, 'category': 'صيانة', 'notes': 'صيانة  مصعد'}
]

#  إنشاء   قاموس    يربط   بين   اسم   ورقة    العمل    و    القائمة
data_dict = {
    'العمارات': buildings,
    'الوحدات': units,
    'المستأجرين': tenants,
    'الإيجارات': rents_paid, 
    'المصروفات': expenses
}

#  #  #  # إنشاء  أوراق العمل  #  #  #  # 
sheet_names = ['العمارات', 'الوحدات', 'المستأجرين', 'الإيجارات', 'المصروفات']
for sheet_name in sheet_names:
    sheet = wb.create_sheet(sheet_name)

    #  #  #  #  كتابة  العناوين  #  #  #  # 
    if sheet_name == 'العمارات':
        headers = ['اسم العمارة', 'عدد الوحدات', 'ملاحظات']
    elif sheet_name == 'الوحدات':
        headers = ['رقم الوحدة', 'العمارة', 'التصنيف', 'الإيجار الشهري', 'الحالة',  'ملاحظات']
    elif sheet_name == 'المستأجرين':
        headers = ['رقم الوحدة', 'اسم المستأجر', 'رقم الهوية', 'رقم الجوال', 'تاريخ بداية العقد', 'تاريخ نهاية العقد', 'قيمة الإيجار', 'البريد الإلكتروني', 'ملاحظات']
    elif sheet_name == 'الإيجارات':
        headers = ['رقم الوحدة', 'الشهر', 'السنة', 'قيمة الإيجار', 'تاريخ الدفع', 'طريقة الدفع', 'الحالة', 'ملاحظات']
    elif sheet_name == 'المصروفات':
        headers = ['العمارة', 'التاريخ', 'نوع المصروفات', 'القيمة',  'الفئة', 'ملاحظات']
    else:
        continue 

    for col_num, header in enumerate(headers, 1):
        cell = sheet.cell(row=1, column=col_num)
        cell.value = header
        set_cell_style(cell, bold=True, bg_color="C0C0C0")
        sheet.column_dimensions[get_column_letter(col_num)].width = 15

    #  #  #  #  إضافة  البيانات  #  #  #  # 
    for row_num, data in enumerate(data_dict[sheet_name], 2): 
        for col_num, key in enumerate(data, 1):
            cell = sheet.cell(row=row_num, column=col_num)
            cell.value = data[key]
            set_cell_style(cell)
            
#  تلوين   صفوف    الإيجارات    غير   المدفوعة
if 'الإيجارات' in wb.sheetnames:
    sheet = wb['الإيجارات']
    for row_num in range(2, sheet.max_row + 1):
        status_cell = sheet.cell(row=row_num, column=7)  #  عمود   "الحالة"
        if status_cell.value == "غير مدفوع":
            for col_num in range(1, sheet.max_column + 1):
                set_cell_style(sheet.cell(row=row_num, column=col_num), bg_color="FFC7CE") 

# حذف  ورقة  العمل  الافتراضية
del wb['Sheet']

# حفظ  ملف Excel
wb.save('إدارة عمارات سكنية.xlsx')