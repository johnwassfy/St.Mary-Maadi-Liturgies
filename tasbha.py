from commonFunctions import relative_path, replacefile, show_hide_insertImage_replaceText, elzoksologyat, open_presentation_relative_path, run_vba_with_slide_id_bakr_aashya, move_sections_v2
import win32com.client

def tasbha(copticdate, Aashya, season):
    from copticDate import CopticCalendar
    from Season import el3nsara
    cd = CopticCalendar().coptic_to_gregorian(copticdate)
    prs = relative_path(r"الإبصلمودية.pptx")
    excel = relative_path(r"Files Data.xlsx")
    sheet = "التسبحة"
    weekday = cd.weekday()
    show_full_sections = []
    show_full_sections_ranges = [[]]
    hide_full_sections_ranges = [[]]

    replacefile(prs, relative_path(r"Data\CopyData\الإبصلمودية.pptx"))
    if Aashya == False : elzoksologyat(excel, season, "نصف الليل")
    
    # tasbha_values =  ["تين ثينو", "الذكصولوجيات", "ني اثنوس تيرو", "ثيؤطوكية الأحد 7-9", "ثيؤطوكية الأحد 16-18", "قانون الايمان", "قدوس قدوس قدوس", "تين ناف"]
    
    tasbha_values = ['{5ADE5763-A478-4AF7-A88C-0DB7B091DF06}', '{40E9947C-C8E3-4367-92E7-64F6896E8A5F}', '{E052F467-3B39-4F35-9D72-8B11039F040B}', '{64CE0420-479F-4F12-AE0C-F1218BF21635}', '{4E843BF3-1D30-4BED-905C-E66AA3D90EC5}', '{A12368B5-4E89-4682-AF79-DC1979BA120B}', '{F2F363F3-5DD8-474B-94A0-6895758AB76D}', '{9A6F925B-011F-483C-B6B1-783155284B27}']

    if weekday == 6 or weekday == 0 or weekday == 1:
        # adam_data = ["ابصالية الأحد 1", "ابصالية الأحد الثانية", "ابصالية الاثنين", "ابصالية الثلاثاء",
        #              "مقدمة الثيؤطوكيات الأدام", "ثيؤطوكية الأحد 1-6", "ثيؤطوكية الأحد 11-15",
        #              "ثيؤطوكية الإثنين", "ثيؤطوكية الثلاثاء", "لبش الإثنين", "لبش الثلاثاء",
        #              "ختام الثؤطوكيات الادام"
        #             ]
        
        adam_data = ['{F8548FDB-8D40-484A-8D19-36EC50E838FD}', '{F263153B-C7A8-4E6B-AACA-6F05AF050F2E}', '{C966C7AC-73F9-4177-AA7F-71D0428224AF}', '{31645468-E515-4F5A-85DB-DEE662F6432A}', '{E358EDB7-F8FF-43DA-A8B6-81839E23E4C6}', '{8B50A9B8-162A-45FD-A40D-5405E501F1E6}', '{67866127-A8D5-451C-B0C2-1CE6E6FBCD1F}', '{5022D768-2E12-4BEA-8D76-E3896BD58932}', '{D1BFEE47-99F3-4046-8C36-B6397205435B}', '{B08DAA27-DC93-470F-8EE4-DBA2CDED73FF}', '{EA64C1D5-8011-4ED1-AECA-ACA0D1D96925}', '{14A3A43C-A9F7-45A8-A510-EE3F33D99572}']
        
        show_full_sections.extend([adam_data[4], adam_data[11]])

        if  weekday == 0:
            show_full_sections.extend([adam_data[2], adam_data[7], adam_data[9]])

        elif weekday == 1:
            show_full_sections.extend([adam_data[3], adam_data[8], adam_data[10]])

        else :
            show_full_sections_ranges.extend([[adam_data[0], adam_data[1]]])
            if Aashya:
                show_full_sections.append(adam_data[6])
            else:
                show_full_sections_ranges.extend([[adam_data[5], adam_data[6]]])

        match(season):
            case 1: #عيد النيروز
                ebsalyElmonasba = '{DCF0D2EF-0E5D-4349-8B22-523BE5D2C719}'
            case 2: #عيد الصليب
                ebsalyElmonasba = '{1970A997-AC32-4FF7-B7A2-DAF83BF4F40B}'
            case 24: #الخميسن
                ebsalyElmonasba = '{43AC03AD-AC75-480D-987F-66CB8CBE3883}'
            case 29: #عيد التجلي
                ebsalyElmonasba = '{EF0F739B-A8DE-419D-8D45-757AA9347AB5}'
            case 30: #صوم العذراء
                ebsalyElmonasba = '{2908EF39-9CFE-4101-AED3-B54AD30D5A78}'
            case 31: #عيد العذراء
                ebsalyElmonasba = '{CF62ACEE-48F9-4ABA-ADDC-6180BEC4873D}'
            case default:
                ebsalyElmonasba = ''
        if season != 0 :
            show_full_sections.append(ebsalyElmonasba)
        
    else :
        wats_data = ["ابصالية الأربعاء", "ابصالية الخميس", "ابصالية الجمعة", "ابصالية السبت",
                     "مقدمة الثيؤطوكيات الواطس", "ثيؤطوكية الأربعاء", "ثيؤطوكية الخميس", 
                     "ثيؤطوكية الجمعة", "ثيؤطوكية السبت", "لبش الأربعاء", "لبش الخميس", "لبش الجمعة",
                     "شيرات السبت 1", "شيرات السبت 2", "ختام الثيؤطوكيات الواطس"
                    ]
        
        wats_data = ['{8ABA75EA-D793-46A0-8AE2-5B61A6B4FD7E}', '{02352F94-02C4-4D7F-9247-697DA282E7C9}', '{E8504067-DC7B-4818-8157-B947A0A74D9A}', '{BF504610-6275-426C-A939-798A885C5C71}', '{F96B080A-3FB2-430B-9BED-E692E913A9B0}', '{5AD56D85-2906-43FB-98E9-FB96F1B37293}', '{88249BFF-471A-47A1-B7BC-E5A5093EC8D7}', '{6C9361D4-74F3-4201-B28D-7EB59C9D9A46}', '{25CBC7C4-A68C-4EBD-B127-98DA707B3413}', '{824D594F-C079-4552-882A-CC297F319D7D}', '{260E1FAC-A9F6-4E94-BAAB-EFD045CD242D}', '{27A6E4EC-9C9A-4029-8EAF-A984FA647997}', '{FA5AF629-FC64-4123-92EA-193DFE2229CC}', '{2DF7B6FE-B056-4813-B72C-DFE470371815}', '{BF439D71-64D1-4376-8E4A-812437425EBB}']
        
        show_full_sections.extend([wats_data[4], wats_data[14]])
        
        if weekday == 2:
            show_full_sections.extend([wats_data[0], wats_data[5], wats_data[9]])

        elif weekday == 3:
            show_full_sections.extend([wats_data[1], wats_data[6], wats_data[10]])

        elif weekday == 4:
            show_full_sections.extend([wats_data[2], wats_data[7], wats_data[11]])

        else:
            show_full_sections.extend([wats_data[3], wats_data[8]])
            show_full_sections_ranges.extend([[wats_data[12], wats_data[13]]])

        match(season):
            case 1: #عيد النيروز
                ebsalyElmonasba = '{EAE15FFA-C230-4B43-9FCF-316199A1C57F}'
            case 2: #عيد الصليب
                ebsalyElmonasba = '{07BB69AA-4BA1-4166-9978-AF812FA02FD7}'
            case 29: #عيد التجلي
                ebsalyElmonasba = '{95F02DE0-6540-4250-B6D4-213F4C9B73FC}'
            case 30: #صوم العذراء
                ebsalyElmonasba = '{222D1CFF-8162-4F43-A7FC-D6E04CE630E4}'
            case 31: #عيد العذراء
                ebsalyElmonasba = '{E2D40FD7-171F-428B-86DB-65B332AB25F3}'
            case default:
                ebsalyElmonasba = ''
        if season != 0 :
            show_full_sections.append(ebsalyElmonasba)

    if Aashya == True :
        hide_full_sections_ranges.extend([[tasbha_values[0], tasbha_values[1]], [tasbha_values[5], tasbha_values[6]]])
        show_full_sections.append(tasbha_values[2])

    elif Aashya == False and weekday<6:
        show_full_sections.append(tasbha_values[3])

    if (23.1 <= season <= 24.1) or (season >= 25 and season <= 26)  or ((([copticdate[1], copticdate[2]] > el3nsara) or (copticdate[1] <= 3)) and weekday == 6):
        show_full_sections.append(tasbha_values[4])
        if Aashya == False :
            show_full_sections.append(tasbha_values[7])
        show_hide_insertImage_replaceText(prs, excel, sheet, show_full_sections, None, show_full_sections_ranges, hide_full_sections_ranges, None, ["لأنك قمت","aktwnk", "آك طونك"])
    else:
        show_hide_insertImage_replaceText(prs, excel, sheet, show_full_sections, None, show_full_sections_ranges, hide_full_sections_ranges, None, None)
    
    powerpoint = win32com.client.Dispatch("PowerPoint.Application")
    powerpoint.Visible = True
    presentation = open_presentation_relative_path(prs)
    
    if Aashya == False : run_vba_with_slide_id_bakr_aashya(excel, sheet, prs, presentation, '{40E9947C-C8E3-4367-92E7-64F6896E8A5F}')

    if weekday < 6 and Aashya == False:
        move_index = ['{64CE0420-479F-4F12-AE0C-F1218BF21635}']
        target_index = ['{68B92169-4103-465B-B31B-28B5C35D1468}']
        if weekday > 1 :
            move_index.append('{DBBEB49F-3396-41D0-81FF-0A028C3CB4DA}')
            target_index.append('{0420AA0C-B21A-478D-88EA-8378E9539EDE}')
    
        move_sections_v2(presentation, move_index, target_index)

    presentation.SlideShowSettings.Run()


# def kiahk(copticdate):
#     from copticDate import CopticCalendar
#     cd = CopticCalendar().coptic_to_gregorian(copticdate)
#     prs = relative_path(r"الإبصلمودية الكيهكية.pptx")
#     prs_new = relative_path(r"Data\CopyData\الإبصلمودية الكيهكية.pptx")
#     replacefile(prs, prs_new)
#     replacefile(relative_path(r"الذكصولوجيات.pptx"), relative_path(r"Data\CopyData\الذكصولوجيات.pptx"))
#     excel = relative_path(r"Files Data.xlsx")
#     sheet = "تسبحة كيهك"
#     weekday = cd.weekday()
#     if weekday==6:
#         show_slide_ranges_from_sections(prs, excel, sheet, ["امدح عذراء و بتول", "ابصالية الأحد", "افتح فاي بالتسابيح",
#                                                             "مقدمة الثيؤطوكيات الأدام", "ثيؤطوكية الأحد 1", "التفسير الأول",
#                                                             "ثيؤطوكية الأحد 2", "التفسير الثاني", "ثيؤطوكية الأحد 3",
#                                                             "التفسير الثالث", "ثيؤطوكية الأحد 4", "التفسير الرابع", 
#                                                             "ثيؤطوكية الأحد 5", "التفسير الخامس", "ثيؤطوكية الأحد 6",
#                                                             "التفسير السادس", "ثيؤطوكية الأحد 10", "ثيؤطوكية الأحد 11-15",
#                                                             "طرح الفعلة القديسين", "لحن ختام طرح الفعلة", "ختام الثؤطوكيات الادام",
#                                                             "مديح مراحمك يا إلهي"])

#     elif weekday==0:
#         show_slide_ranges_from_sections(prs, excel, sheet, ["ابصالية الاثنين كيهك", "ابصالية الاثنين", "مقدمة الثيؤطوكيات الأدام", 
#                                                             "ثيؤطوكية الإثنين", "لبش الإثنين", "ختام الثؤطوكيات الادام",
#                                                             "مديح مراحمك يا إلهي"])

#     elif weekday==1:
#         show_slide_ranges_from_sections(prs, excel, sheet, ["ابصالية الثلاثاء كيهك", "ابصالية الثلاثاء", "مقدمة الثيؤطوكيات الأدام", 
#                                                             "ثيؤطوكية الثلاثاء", "لبش الثلاثاء", "ختام الثؤطوكيات الادام",
#                                                             "مديح مراحمك يا إلهي"])

#     elif weekday==2:
#         show_slide_ranges_from_sections(prs, excel, sheet, ["ابصالية الأربعاء كيهك", "ابصالية الأربعاء", "مقدمة الثيؤطوكيات الواطس", 
#                                                             "ثيؤطوكية الأربعاء", "لبش الأربعاء", "ختام الثيؤطوكيات الواطس"])

#     elif weekday==3:
#         show_slide_ranges_from_sections(prs, excel, sheet, ["ابصالية الخميس كيهك", "ابصالية الخميس", "مقدمة الثيؤطوكيات الواطس", 
#                                                             "ثيؤطوكية الخميس", "لبش الخميس", "ختام الثيؤطوكيات الواطس"])

#     elif weekday==4:
#         show_slide_ranges_from_sections(prs, excel, sheet, ["ابصالية الجمعة كيهك", "ابصالية الجمعة", "مقدمة الثيؤطوكيات الواطس", 
#                                                             "ثيؤطوكية الجمعة", "لبش الجمعة", "ختام الثيؤطوكيات الواطس"])

#     elif weekday==5:
#         show_slide_ranges_from_sections(prs, excel, sheet, ["ابصالية السبت كيهك", "ابصالية السبت", "مقدمة الثيؤطوكيات الواطس", 
#                                                             "ثيؤطوكية السبت", "لبش السبت", "ختام الثيؤطوكيات الواطس"])

#     presentation = open_presentation_relative_path(prs)
#     run_vba_with_slide_id_bakr_aashya(excel, sheet, prs, presentation)
#     if weekday != 6:
#         move_sections_range(presentation, "ثيؤطوكية الأحد 7-أ", "امدح في البتول", "طرح آدام على الهوس الاول")

#     if weekday < 6 and weekday > 1:
#        move_sections(presentation, ["مقدمة الدفنار"], ["مقدمة الدفنار الآدام"])

#     presentation.SlideShowSettings.Run()


