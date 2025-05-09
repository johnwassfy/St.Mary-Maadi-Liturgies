import datetime
import copticDate
from commonFunctions import read_excel_cell, read_excel_cells_by_array
import os

def relative_path(relative_path):
    script_directory = os.path.dirname(os.path.abspath(__file__))
    absolute_path = os.path.join(script_directory, relative_path)
    return absolute_path

workbook = relative_path(r"Tables.xlsx")
sheetname = "المناسبات"

# Pre-calculate values from Excel
coptic = copticDate.CopticCalendar()
baramonElmilad = [coptic.current_coptic_datetime[0], read_excel_cell(workbook, sheetname, "D4"), read_excel_cell(workbook, sheetname, "F4")]
Elmilad = [coptic.current_coptic_datetime[0], read_excel_cell(workbook, sheetname, "D5"), read_excel_cell(workbook, sheetname, "F5")]
SomElmilad_start = coptic.coptic_date_before(43, Elmilad)

baramonElghetas = [read_excel_cell(workbook, sheetname, "D7"), read_excel_cell(workbook, sheetname, "F7")]

SomNynawa = [read_excel_cell(workbook, sheetname, "D11"), read_excel_cell(workbook, sheetname, "F11")]
Fes7Younan = [read_excel_cell(workbook, sheetname, "D12"), read_excel_cell(workbook, sheetname, "F12")]

ElSomElkbyr = [read_excel_cell(workbook, sheetname, "D13"), read_excel_cell(workbook, sheetname, "F13")]
SabtElrefa3 = coptic.coptic_date_before(2, [coptic.current_coptic_datetime[0], ElSomElkbyr[0], ElSomElkbyr[1]])
A7dElrefa3 = coptic.coptic_date_before(1, [coptic.current_coptic_datetime[0], ElSomElkbyr[0], ElSomElkbyr[1]])

A7dElKonoz = coptic.coptic_date_after(7, A7dElrefa3)
A7dElTagrba = coptic.coptic_date_after(7, A7dElKonoz)
A7dElEbnEldal = coptic.coptic_date_after(7, A7dElTagrba)
A7dElSamerya = coptic.coptic_date_after(7, A7dElEbnEldal)
A7dElM5l3 = coptic.coptic_date_after(7, A7dElSamerya)
A7dElMawlodA3ma = coptic.coptic_date_after(7, A7dElM5l3)

Gom3t5tamElsom = [read_excel_cell(workbook, sheetname, "D16"), read_excel_cell(workbook, sheetname, "F16")]
saturday_of_lazarus = [read_excel_cell(workbook, sheetname, "D17"), read_excel_cell(workbook, sheetname, "F17")]
Elsh3anyn = [read_excel_cell(workbook, sheetname, "D18"), read_excel_cell(workbook, sheetname, "F18")]
khamysEl3hd = [read_excel_cell(workbook, sheetname, "D19"), read_excel_cell(workbook, sheetname, "F19")]
GreatFriday = [read_excel_cell(workbook, sheetname, "D20"), read_excel_cell(workbook, sheetname, "F20")]
HolySaturday = [read_excel_cell(workbook, sheetname, "D21"), read_excel_cell(workbook, sheetname, "F21")]
El2yama = [read_excel_cell(workbook, sheetname, "D22"), read_excel_cell(workbook, sheetname, "F22")]

elso3od = [read_excel_cell(workbook, sheetname, "D24"), read_excel_cell(workbook, sheetname, "F24")]

el3nsara = [read_excel_cell(workbook, sheetname, "D25"), read_excel_cell(workbook, sheetname, "F25")]

def get_season(date):
    CD = copticDate.CopticCalendar()
    coptic_date = CD.gregorian_to_coptic(datetime.datetime(date.year, date.month, date.day, date.hour,date.minute))
    month, day = coptic_date[1], coptic_date[2]
    season = 0

    if month == 1 and day <= 16:
        season = 1  # فترة النيروز
    elif (month == 1 and (17 <= day <= 19)) or (month == 7 and day == 10):
        season = 2  # عيدي الصليب
    elif [baramonElmilad[1], baramonElmilad[2]]<= [month, day] < [Elmilad[1], Elmilad[2]]:
        season = 3 # برامون الميلاد
    elif [month, day] == [Elmilad[1], Elmilad[2]]:
        season = 4 # عيد الميلاد
    elif [month, day] == [Elmilad[1] + 1, Elmilad[2]]:
        season = 4.1 # اليوم الثاني من الميلاد
    elif month == 4:
        season = 5  # كيهك
    elif day == 29 and (month < 4 or month > 7):
        season = 32 # تذكار الاعياد السيدية
    elif [SomElmilad_start[1], SomElmilad_start[2]] <=  [month, day] < [4, 1]:
        season = 6 # صوم الميلاد
    elif [5, 6] > [month, day] > [Elmilad[1], Elmilad[2]]:
        season = 4.2 # فترة الميلاد
    elif month == 5 and day == 6:
        season = 7  #الختان
    elif baramonElghetas  <= [month, day] < [5, 11]:
        season = 8   # برامون الغطاس
    elif [month, day] == [5, 11]:
        season = 9   # عيد الغطاس
    elif [month, day] == [5, 12]:
        season = 9.1   # عيد الغطاس
    elif month == 5 and day == 13:
        season = 10 #عرس قانا الجليل
    elif month == 6 and day == 8:
        season = 11  # دخول المسيح الهيكل
    elif SomNynawa <= [month, day] < Fes7Younan:
        season = 12  # صوم نينوى
    elif [month, day] == Fes7Younan:
        season = 13 # فصح يونان
    elif month == 7 and day == 29:
        season = 14  # عيد البشارة
    elif ElSomElkbyr == [month, day]:
        season = 15.4 #الإثنين الأول من الصوم الكبير
    elif [month, day] == Gom3t5tamElsom:
        season = 15.1  # جمعة ختام الصوم
    elif [SabtElrefa3[1], SabtElrefa3[2]] == [month, day]:
        season = 15.2  # سبت الرفاع
    elif [A7dElrefa3[1], A7dElrefa3[2]] == [month, day]:
        season = 15.3  # احد الرفاع
    elif A7dElKonoz == coptic_date:
        season = 15.5 #أحد الكنوز
    elif A7dElTagrba == coptic_date:
        season = 15.6 #أحد التجربة
    elif A7dElEbnEldal == coptic_date:
        season = 15.7 #أحد الإبن الضال
    elif A7dElSamerya == coptic_date:
        season = 15.8 #أحد السامرية
    elif A7dElM5l3 == coptic_date:
        season = 15.9 #أحد المخلع
    elif A7dElMawlodA3ma == coptic_date:
        season = 15.11 #أحد المولود أعمى
    elif ElSomElkbyr < [month, day] < Gom3t5tamElsom:
        season = 15  # الصوم الكبير
    elif [month, day] == saturday_of_lazarus:
        season = 16  # سبت لعازر
    elif[month, day] == Elsh3anyn:
        season = 17  # احد الشعانين
    elif [month, day] ==  khamysEl3hd:
        season = 19  # خميس العهد
    elif [month, day] == GreatFriday:
        season = 20 #الجمعة العظيمة
    elif Elsh3anyn < [month, day] <= GreatFriday:
        season = 18 #إسبوع الالام
    elif [month, day] == HolySaturday:
        season = 21  # سبت النور
    elif [month, day] == El2yama :
        season = 22  #عيد القيامة
    elif [month, day] == elso3od:
        season = 25 #عيد الصعود
    elif [month, day] == el3nsara :
        season = 26 # عيدالعنصرة
    elif [month, day] == [9, 24]:
        season = 23 #عيد دخول المسيح أرض مصر
    elif [month, day] == [9, 24] < el3nsara :
        season = 23.1 #عيد دخول المسيح أرض مصر في الخمسين
    elif [month, day] == [9, 24] == el3nsara :
        season = 23.2 # عيد دخول المسيح أرض مصر في العنصرة
    elif [month, day] == elso3od :
        season = 23.3 # عيد دخول المسيح أرض مصر في الصعود
    elif El2yama < [month, day] < elso3od:
        season = 24 # فترة الخماسين المقدسة
    elif elso3od < [month, day] < el3nsara:
        season = 24.1 # فترة الخماسين المقدسة
    elif el3nsara < [month, day] < [11, 5] :
        season = 27 #صوم الرسل
    elif [month, day] == [11, 5] :
        season = 28  # عيد الرسل
    elif [month, day] == [12, 13]:
        season = 29 # عيد التجلي
    elif [12, 1] <= [month, day] < [12, 16]:
        season = 30 # صوم العذراء
    elif [month, day] == [12, 16]:
        season = 31 # عيد العذراء
    else:
        season = 0 # سنوي

    return season

def get_season_name(season_number):
    seasons = {
        0: "سنوي",
        1: "فترة النيروز",
        2: "عيد الصليب",
        3: "برامون الميلاد",
        4: "عيد الميلاد المجيد",
        4.1: "اليوم الثاني من الميلاد",
        4.2 : "فترة الميلاد",
        5: "شهر كيهك",
        6: "صوم الميلاد",
        7: "عيد الختان",
        8: "برامون الغطاس",
        9: "عيد الغطاس",
        9.1: "اليوم الثاني من الغطاس",
        10: "عرس قانا الجليل",
        11: "دخول المسيح الهيكل",
        12: "صوم نينوى",
        13: "فصح يونان",
        14: "عيد البشارة",
        15: "الصوم الكبير",
        15.1: "جمعة ختام الصوم",
        15.2: "سبت الرفاع",
        15.3: "احد الرفاع",
        15.4: "الإثنين الأول من الصوم الكبير",
        15.5: "الأحد الأول من الصوم الكبير",
        15.6: "الأحد الثاني من الصوم الكبير",
        15.7: "الأحد الثالث من الصوم الكبير",
        15.8: "الأحد الرابع من الصوم الكبير",
        15.9: "الأحد الخامس من الصوم الكبير",
        15.11: "الأحد السادس من الصوم الكبير",
        16: "سبت لعازر",
        17: "احد الشعانين",
        18: "أسبوع الالام",
        19: "خميس العهد",
        20: "الجمعة العظيمة",
        21: "سبت النور",
        22: "عيد القيامة",
        23: "دخول المسيح أرض مصر",
        23.1: "دخول المسيح أرض مصر",
        23.2: "عيد العنصرة ودخول المسيح أرض مصر",
        23.3: "عيد الصعود ودخول المسيح أرض مصر ",
        24: "الخمسين المقدسة",
        24.1: "الخمسين المقدسة",
        25: "عيد الصعود",
        26: "عيد العنصرة",
        27: "صوم الرسل",
        28: "عيد الرسل",
        29: "عيد التجلي",
        30: "صوم العذراء",
        31: "إظهار صعود جسد العذراء",
        32: "تذكار الاعياد السيدية"
    }
    return seasons.get(season_number, "Unknown Season")

