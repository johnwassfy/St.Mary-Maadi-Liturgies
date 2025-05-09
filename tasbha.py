from commonFunctions import *
import win32com.client
from Season import El2yama, el3nsara
from datetime import datetime

def tasbha(copticdate, Aashya, season):
    from copticDate import CopticCalendar
    cd = CopticCalendar().coptic_to_gregorian(copticdate)
    prs = relative_path(r"الإبصلمودية.pptx")
    excel = relative_path(r"بيانات القداسات.xlsx")
    sheet = "التسبحة"
    weekday = cd.weekday()
    show_array = [[1, 1]]
    hide_array = [[1, 1]]
    if Aashya == False : elzoksologyat(excel, season, "نصف الليل")
    data = find_slide_nums_arrays(excel, sheet, 
                                ["تين ثينو", "الذكصولوجيات", "ني اثنوس تيرو", "ني اثنوس تيرو",
                                 "ثيؤطوكية الأحد 7-9", "ثيؤطوكية الأحد 7-9",
                                 "ثيؤطوكية الأحد 16-18", "ثيؤطوكية الأحد 16-18", "قانون الايمان", "قدوس قدوس قدوس",
                                 "تين ناف", "تين ناف"
                                ],
                                [1, 2, 1, 2, 1, 2, 1, 2, 1, 2, 1, 2])
    tynav = data[10]
    tynav2 = data[11]
    theo79 = data[4]
    theo792 = data[5]
    theo1618 = data[6]
    theo16182 = data[7]

    if weekday == 6 or weekday == 0 or weekday == 1:
        adam_data = find_slide_nums_arrays(excel, sheet, 
                                ["ابصالية الأحد 1", "ابصالية الأحد الثانية", "ابصالية الاثنين", "ابصالية الاثنين",
                                 "ابصالية الثلاثاء", "ابصالية الثلاثاء", "مقدمة الثيؤطوكيات الأدام", "مقدمة الثيؤطوكيات الأدام",
                                 "ثيؤطوكية الأحد 1-6", "ثيؤطوكية الأحد 11-15", "ثيؤطوكية الإثنين", "ثيؤطوكية الإثنين",
                                 "ثيؤطوكية الثلاثاء", "ثيؤطوكية الثلاثاء",
                                 "لبش الإثنين", "لبش الإثنين", "لبش الثلاثاء", "لبش الثلاثاء",
                                 "ختام الثؤطوكيات الادام", "ختام الثؤطوكيات الادام"
                                ],
                                [1, 2, 1, 2, 1, 2, 1, 2, 1, 2, 1, 2, 1, 2, 1, 2, 1, 2, 1, 2])
        theoStart = adam_data[6]
        theoStart2 = adam_data[7]
        theoEnd = adam_data[18]
        theoEnd2 = adam_data[19]

        if  weekday == 0:
            ebsalya = adam_data[2]
            ebsalya2 = adam_data[3]
            theo = adam_data[10]
            theo2 = adam_data[11]
            lobsh = adam_data[14]
            lobsh2 = adam_data[15]

        elif weekday == 1:
            ebsalya = adam_data[4]
            ebsalya2 = adam_data[5]
            theo = adam_data[12]
            theo2 = adam_data[13]
            lobsh = adam_data[16]
            lobsh2 = adam_data[17]

        else :
            # sunday(prs)
            ebsalya = adam_data[0]
            ebsalya2 = adam_data[1]
            if Aashya:
                theo = find_slide_num(excel, sheet, "ثيؤطوكية الأحد 11-15", 1)
            else:
                theo = adam_data[8]
            theo2 = adam_data[9]
            lobsh = 1
            lobsh2 = 1

        if season == 1:
            ebsalyElmonasba = find_slide_num(excel, sheet, "ابصالية آدام لعيد النيروز",  1)
            ebsalyElmonasba2 = find_slide_num(excel, sheet, "ابصالية آدام لعيد النيروز",  2)

        elif season == 2:
            ebsalyElmonasba = find_slide_num(excel, sheet, "ابصالية آدام لعيد الصليب",  1)
            ebsalyElmonasba2 = find_slide_num(excel, sheet, "ابصالية آدام لعيد الصليب",  2)
            
        elif season == 29 :
            ebsalyElmonasba = find_slide_num(excel, sheet, "ابصالية آدام لعيد التجلي",  1)
            ebsalyElmonasba2 = find_slide_num(excel, sheet, "ابصالية آدام لعيد التجلي",  2)
        
        elif season == 30 :
            ebsalyElmonasba = find_slide_num(excel, sheet, "ابصالية آدام لصوم العذراء",  1)
            ebsalyElmonasba2 = find_slide_num(excel, sheet, "ابصالية آدام لصوم العذراء",  2)

        elif season == 31:
            ebsalyElmonasba = find_slide_num(excel, sheet, "ابصالية آدام لعيد العذراء", 1)
            ebsalyElmonasba2 = find_slide_num(excel, sheet, "ابصالية آدام لعيد العذراء", 2)

        if season != 0 :
            show_array.extend([[ebsalyElmonasba, ebsalyElmonasba2]])
        
    else :
        wats_data = find_slide_nums_arrays(excel, sheet, 
                                ["ابصالية الأربعاء", "ابصالية الأربعاء", "ابصالية الخميس", "ابصالية الخميس",
                                 "ابصالية الجمعة", "ابصالية الجمعة", "ابصالية السبت", "ابصالية السبت",
                                 "مقدمة الثيؤطوكيات الواطس", "مقدمة الثيؤطوكيات الواطس", "ثيؤطوكية الأربعاء",
                                 "ثيؤطوكية الأربعاء", "ثيؤطوكية الخميس", "ثيؤطوكية الخميس", "ثيؤطوكية الجمعة",
                                 "ثيؤطوكية الجمعة", "ثيؤطوكية السبت", "ثيؤطوكية السبت", "لبش الأربعاء",
                                 "لبش الأربعاء", "لبش الخميس", "لبش الخميس", "لبش الجمعة", "لبش الجمعة",
                                 "شيرات السبت 1", "شيرات السبت 2", "ختام الثيؤطوكيات الواطس", "ختام الثيؤطوكيات الواطس"
                                ],
                                [1, 2, 1, 2, 1, 2, 1, 2, 1, 2, 1, 2, 1, 2, 1, 2, 1, 2, 1, 2, 1, 2, 1, 2, 1, 2, 1, 2])
        
        theoStart = wats_data[8]
        theoStart2 = wats_data[9]
        theoEnd = wats_data[26]
        theoEnd2 = wats_data[27]
        
        if weekday == 2:
            ebsalya = wats_data[0]
            ebsalya2 = wats_data[1]
            theo = wats_data[10]
            theo2 = wats_data[11]
            lobsh = wats_data[18]
            lobsh2 = wats_data[19]

        elif weekday == 3:
            ebsalya = wats_data[2]
            ebsalya2 = wats_data[3]
            theo = wats_data[12]
            theo2 = wats_data[13]
            lobsh = wats_data[20]
            lobsh2 = wats_data[21]

        elif weekday == 4:
            ebsalya = wats_data[4]
            ebsalya2 = wats_data[5]
            theo = wats_data[14]
            theo2 = wats_data[15]
            lobsh = wats_data[22]
            lobsh2 = wats_data[23]

        else:
            ebsalya = wats_data[6]
            ebsalya2 = wats_data[7]
            theo = wats_data[16]
            theo2 = wats_data[17]
            lobsh = wats_data[24]
            lobsh2 = wats_data[25]

        if season == 1:
            ebsalyElmonasba = find_slide_num(excel, sheet, "ابصالية واطس لعيد النيروز",  1)
            ebsalyElmonasba2 = find_slide_num(excel, sheet, "ابصالية واطس لعيد النيروز",  2)

        elif season == 2:
            ebsalyElmonasba = find_slide_num(excel, sheet, "ابصالية واطس لعيد الصليب",  1)
            ebsalyElmonasba2 = find_slide_num(excel, sheet, "ابصالية واطس لعيد الصليب",  2)

        elif season == 29 :
            ebsalyElmonasba = find_slide_num(excel, sheet, "ابصالية واطس لعيد التجلي",  1)
            ebsalyElmonasba2 = find_slide_num(excel, sheet, "ابصالية واطس لعيد التجلي",  2)

        elif season == 30 :
            ebsalyElmonasba = find_slide_num(excel, sheet, "ابصالية واطس لصوم العذراء",  1)
            ebsalyElmonasba2 = find_slide_num(excel, sheet, "ابصالية واطس لصوم العذراء",  2)
        
        elif season == 31:
            ebsalyElmonasba = find_slide_num(excel, sheet, "ابصالية واطس لعيد العذراء", 1)
            ebsalyElmonasba2 = find_slide_num(excel, sheet, "ابصالية واطس لعيد العذراء", 2)

        if season != 0 :
            if Aashya or season>=30:
                show_array.append([ebsalyElmonasba, ebsalyElmonasba2])
            else:
                show_array.extend([[ebsalyElmonasba, ebsalyElmonasba2]])

    
    show_array.extend([[ebsalya, ebsalya2], [theoStart, theoStart2], [theo, theo2], [lobsh, lobsh2], [theoEnd, theoEnd2]])

    if Aashya == True :
        tasb7a = data[0]
        tasb7a2 = data[1]
        nyethnos = data[2]
        nyethnos2 = data[3]
        hide_array.extend([[tasb7a, tasb7a2]])
        show_array.append([nyethnos, nyethnos2])

    elif Aashya == False and weekday<6:
            show_array.append([theo79, theo792])

    if (El2yama <= [copticdate[1], copticdate[2]] <= el3nsara) or ((([copticdate[1], copticdate[2]] > el3nsara) or (copticdate[1] <= 3)) and weekday == 6):
        show_array.append([theo1618, theo16182])
        if Aashya == False :
            show_array.append([tynav, tynav2])

    powerpoint = win32com.client.Dispatch("PowerPoint.Application")
    powerpoint.Visible = True  # Open PowerPoint application
    presentation = open_presentation_relative_path(prs)
    
    if Aashya == False : run_vba_with_slide_id_bakr_aashya(excel, sheet, prs, presentation)

    hide_slides(presentation, hide_array)
    show_slides(presentation, show_array)

    if weekday < 6 and Aashya == False:
        sections = {presentation.SectionProperties.Name(i): i for i in range(1, presentation.SectionProperties.Count + 1)}
        move_index = sections["ثيؤطوكية الأحد 7-9"]
        target_index = sections["لبش الهوس الاول"]
        if move_index < target_index:
            target_index -= 1
        presentation.SectionProperties.Move(move_index, target_index + 1)
        if weekday > 1 :
            sections = {presentation.SectionProperties.Name(i): i for i in range(1, presentation.SectionProperties.Count + 1)}
            move_index = sections["مقدمة الدفنار"]
            target_index = sections["مقدمة الدفنار الآدام"]
            if move_index < target_index:
                target_index -= 1
            presentation.SectionProperties.Move(move_index, target_index + 1)

    presentation.SlideShowSettings.Run()

def kiahk(copticdate):
    from copticDate import CopticCalendar
    cd = CopticCalendar().coptic_to_gregorian(copticdate)
    prs = relative_path(r"الإبصلمودية الكيهكية.pptx")
    prs_new = relative_path(r"Data\CopyData\الإبصلمودية الكيهكية.pptx")
    replacefile(prs, prs_new)
    replacefile(relative_path(r"الذكصولوجيات.pptx"), relative_path(r"Data\CopyData\الذكصولوجيات.pptx"))
    excel = relative_path(r"بيانات القداسات.xlsx")
    sheet = "تسبحة كيهك"
    weekday = cd.weekday()
    if weekday==6:
        show_slide_ranges_from_sections(prs, excel, sheet, ["امدح عذراء و بتول", "ابصالية الأحد", "افتح فاي بالتسابيح",
                                                            "مقدمة الثيؤطوكيات الأدام", "ثيؤطوكية الأحد 1", "التفسير الأول",
                                                            "ثيؤطوكية الأحد 2", "التفسير الثاني", "ثيؤطوكية الأحد 3",
                                                            "التفسير الثالث", "ثيؤطوكية الأحد 4", "التفسير الرابع", 
                                                            "ثيؤطوكية الأحد 5", "التفسير الخامس", "ثيؤطوكية الأحد 6",
                                                            "التفسير السادس", "ثيؤطوكية الأحد 10", "ثيؤطوكية الأحد 11-15",
                                                            "طرح الفعلة القديسين", "لحن ختام طرح الفعلة", "ختام الثؤطوكيات الادام",
                                                            "مديح مراحمك يا إلهي"])

    elif weekday==0:
        show_slide_ranges_from_sections(prs, excel, sheet, ["ابصالية الاثنين كيهك", "ابصالية الاثنين", "مقدمة الثيؤطوكيات الأدام", 
                                                            "ثيؤطوكية الإثنين", "لبش الإثنين", "ختام الثؤطوكيات الادام",
                                                            "مديح مراحمك يا إلهي"])

    elif weekday==1:
        show_slide_ranges_from_sections(prs, excel, sheet, ["ابصالية الثلاثاء كيهك", "ابصالية الثلاثاء", "مقدمة الثيؤطوكيات الأدام", 
                                                            "ثيؤطوكية الثلاثاء", "لبش الثلاثاء", "ختام الثؤطوكيات الادام",
                                                            "مديح مراحمك يا إلهي"])

    elif weekday==2:
        show_slide_ranges_from_sections(prs, excel, sheet, ["ابصالية الأربعاء كيهك", "ابصالية الأربعاء", "مقدمة الثيؤطوكيات الواطس", 
                                                            "ثيؤطوكية الأربعاء", "لبش الأربعاء", "ختام الثيؤطوكيات الواطس"])

    elif weekday==3:
        show_slide_ranges_from_sections(prs, excel, sheet, ["ابصالية الخميس كيهك", "ابصالية الخميس", "مقدمة الثيؤطوكيات الواطس", 
                                                            "ثيؤطوكية الخميس", "لبش الخميس", "ختام الثيؤطوكيات الواطس"])

    elif weekday==4:
        show_slide_ranges_from_sections(prs, excel, sheet, ["ابصالية الجمعة كيهك", "ابصالية الجمعة", "مقدمة الثيؤطوكيات الواطس", 
                                                            "ثيؤطوكية الجمعة", "لبش الجمعة", "ختام الثيؤطوكيات الواطس"])

    elif weekday==5:
        show_slide_ranges_from_sections(prs, excel, sheet, ["ابصالية السبت كيهك", "ابصالية السبت", "مقدمة الثيؤطوكيات الواطس", 
                                                            "ثيؤطوكية السبت", "لبش السبت", "ختام الثيؤطوكيات الواطس"])

    presentation = open_presentation_relative_path(prs)
    run_vba_with_slide_id_bakr_aashya(excel, sheet, prs, presentation)
    if weekday != 6:
        move_sections_range(presentation, "ثيؤطوكية الأحد 7-أ", "امدح في البتول", "طرح آدام على الهوس الاول")

    if weekday < 6 and weekday > 1:
       move_sections(presentation, ["مقدمة الدفنار"], ["مقدمة الدفنار الآدام"])

    presentation.SlideShowSettings.Run()


