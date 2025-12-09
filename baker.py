import os
from commonFunctions import *
import win32com.client

"_____________________________________OLD CODE_DESIGN_____________________________________"

"_________________________________NO_MORE_OLD_CODE_DESIGN_________________________________"

"_____________________________________NEW_CODE_DESIGN_____________________________________"

def bakerSanawy(season, copticdate, adam = False, Bishop = False, guestBishop = 0):
    from copticDate import CopticCalendar
    cd = CopticCalendar().coptic_to_gregorian(copticdate)
    prs = relative_path(r"رفع بخور عشية و باكر.pptx")  # Using the relative path
    excel = relative_path(r"Files Data.xlsx")
    excel2 = relative_path(r"Tables.xlsx")
    des_sheet ="رفع بخور"
    replacefile(prs, relative_path(r"Data\CopyData\رفع بخور عشية و باكر.pptx"))
    
    elzoksologyat(excel, season,"باكر")

    if cd.weekday() == 6:
        katamars = relative_path(r"Data\القطمارس\القطمارس السنوي احاد.pptx")
        katamars_sheet = "القطمارس السنوي آحاد"
        km = copticdate[1]
        kd = (copticdate[2] - 1) // 7 + 1
    else: 
        katamars = relative_path(r"Data\القطمارس\القطمارس السنوي ايام.pptx")
        katamars_sheet = "القطمارس السنوي أيام"
        km, kd = find_Readings_Date(copticdate[1], copticdate[2])

    katamars_values = fetch_data_arrays(excel2, katamars_sheet, km, kd, [6, 7, 8])
    elmzmor = katamars_values[0]
    elengil = katamars_values[1]
    elengil2 = katamars_values[2]

    # baker_show_full_sections = ["فلنسبح مع الملائكة"]
    
    baker_show_full_sections = ['{2ECE1F1B-C143-4CE2-B550-348BEE185974}']
    baker_hide_full_sections = []

    # if adam:
    #     aashya_show_full_sections.append("ارباع الناقوس الادام")
    # else:
    #     aashya_show_full_sections.append("ارباع الناقوس الواطس")

    if adam:
        baker_show_full_sections.append("{D02C088A-01E0-4A8C-8D73-21E3FD3616EB}")
    else:
        baker_show_full_sections.append("{9495E38B-CE03-4E75-AED4-960DE95BA747}")

    # aashya_values = ["الانجيل", "المزمور", "تكملة على حسب المناسبة"]

    aashya_values = find_slide_nums_arrays_v2(excel, des_sheet, 
                    ['{B74DBB8C-2B2D-46E4-9508-DA46008D19A4}', '{6D1E6E7D-EECE-483C-A3AE-C135D02E717C}', '{A18EDC94-F257-4FAC-99C7-0A8EA70F0FAF}'], 
                    2, [2, 2, 1])

    #المزمور و الانجيل
    elengil3 = aashya_values[0]
    elmzmor1 = aashya_values[1]

    #الختام
    elkhetam = aashya_values[2]

    #الاواشي
    if cd.weekday() == 5: #السبت
        # baker_show_full_sections.append("اوشية الراقدين")
        baker_show_full_sections.append('{83E6BC33-A9EC-45CA-89B6-24EFBC51B654}')
    elif cd.weekday() == 6:#الاحد
        # baker_show_full_sections.extend(["اوشية المرضى", "اوشية القرابين"])
        baker_show_full_sections.extend(["{069F7A79-999B-4223-82AE-CAF356118167}", "{2C897F14-44CC-430E-9BE1-EB379FE7A9C7}"])
    else:
        # baker_show_full_sections.extend(["اوشية المرضى", "اوشية المسافرين"])
        baker_show_full_sections.extend(["{069F7A79-999B-4223-82AE-CAF356118167}", "{A059EEC9-5D25-453F-A956-A2E149F0773C}"])

    #مرد الانجيل
    if season == 27:
        # baker_show_full_sections.append("ربع يقال في صوم الرسل")
        baker_show_full_sections.append("{F5AB11D4-D7D2-4DA3-A830-32BA45BCB16D}")
    elif season == 28:
        # baker_show_full_sections.extend(["ربع قال في عيد الرسل", "ربع بطرس و بولس"])
        baker_show_full_sections.extend(["{A0B1F7D2-3C4B-4E5F-8A6B-9D0C1F2E3F4A}", "{38E5A337-7696-4261-833A-DF790456C6A8}"])
    if season == 30 | 31 :
        # baker_show_full_sections.append('مرد انجيل صوم العذراء - باكر')
        # baker_hide_full_sections.append('مرد الانجيل السنوي')
        baker_show_full_sections.append('{037D8578-7219-4388-AFC5-4753352BFA8C}')
        baker_hide_full_sections.append('{BEECCC68-2AEF-4568-91AA-98BCD14D3B92}')

    baker_season = CopticCalendar().get_coptic_date_range(copticdate)

    if Bishop == True:
        prs3 = relative_path(r"Data\حضور الأسقف.pptx")
        sheet = "في حضور الأسقف"

        # bishop_show_values = ['تكملة في حضور الاسقف', 'مارو اتشاسف', 'فليرفعوه', 'في حضور الاسقف']

        bishop_show_values = ['{F76B0D75-0474-45B5-B79F-7416F354543A}', '{62A12AF8-CB6D-4CC5-9DB0-B73A7C24E2AD}',
                              '{23533FC3-43FE-456F-9454-70C3088055E7}', '{A9183893-7B7E-459F-8547-F7A8F7D2D521}']
                
        baker_show_full_sections.extend(bishop_show_values)

        if guestBishop > 0:
            # bishop_values = ["صلاة الشكر", "صلاة الشكر", "طوبه هينا الكبيرة", "طوبه هينا الكبيرة", 
            #                  "نيم بينيوت", "نيم بينيوت", "الاسبسمس", "الاسبسمس"]

            # bishopDes_values = ["تكملة في حضور الاسقف", "طوبه هينا الكبيرة", "مارو اتشاسف"]
            
            bishop_values = find_slide_nums_arrays_v2(excel, sheet, 
                            ['{6851F163-CBEF-4014-A853-CE100557BA6A}', '{6851F163-CBEF-4014-A853-CE100557BA6A}',
                             '{B084BC40-61E1-4477-98DA-15CFB06AEE91}', '{B084BC40-61E1-4477-98DA-15CFB06AEE91}',
                             '{97203297-EECB-4D41-B2E3-AD9A4863847E}', '{97203297-EECB-4D41-B2E3-AD9A4863847E}', 
                             '{D1378DB5-29D1-4800-9D96-10F2535EEB57}', '{D1378DB5-29D1-4800-9D96-10F2535EEB57}'], 
                            2, [1, 2, 1, 2, 1, 2, 1, 2])

            bishopDes_values = find_slide_nums_arrays_v2(excel, des_sheet, 
                               ['{F76B0D75-0474-45B5-B79F-7416F354543A}', 
                                '{8DD21CDE-CB6B-4D5B-B995-D2747AB69ED1}', 
                                '{62A12AF8-CB6D-4CC5-9DB0-B73A7C24E2AD}'],                  
                               2, [2, 2, 2])

            elshokr = bishopDes_values[0]
            elshokr1 = bishop_values[0]
            elshokr2 = bishop_values[1]

            tobhyna = bishopDes_values[1]
            tobhyna1 = bishop_values[2]
            tobhyna2 = bishop_values[3]

            maro = bishopDes_values[2]
            maro1 = bishop_values[4]
            maro2 = bishop_values[5]

            if guestBishop < 2:
                elshokr2 = elshokr2-1
                tobhyna2 = tobhyna2-2
                maro2 = maro2-1
        
            start_positions = [elengil3, elmzmor1, maro, tobhyna, elshokr]
            start_slides = [elengil, elmzmor, maro1, tobhyna1, elshokr1]
            end_slides = [elengil2, elmzmor, maro2, tobhyna2, elshokr2]

    if guestBishop == 0:
        start_positions = [elengil3, elmzmor1]
        start_slides = [elengil, elmzmor]
        end_slides = [elengil2, elmzmor]

    if cd.weekday() == 6:
        show_hide_insertImage_replaceText(prs, excel, des_sheet, baker_show_full_sections, baker_hide_full_sections, new_Text=["لأنك قمت","aktwnk", "آك طونك"])
    else:
        show_hide_insertImage_replaceText(prs, excel, des_sheet, baker_show_full_sections, baker_hide_full_sections)

    powerpoint = win32com.client.Dispatch("PowerPoint.Application")
    powerpoint.Visible = True  # Open PowerPoint application
    presentation1 = open_presentation_relative_path(prs)
    presentation2 = open_presentation_relative_path(katamars)

    if guestBishop > 0:
        presentation3 = open_presentation_relative_path(prs3)

    sections = {presentation1.SectionProperties.Name(i): i for i in range(1, presentation1.SectionProperties.Count + 1)}
    target_index = sections["أوشية الموضع"]
    show_array = []
    if baker_season == "Air" :
        move_index = sections["اوشية الأهوية والثمار"]
        air = find_slide_index_by_title(presentation1, "الاهوية", elkhetam)
        show_array.append([air, air])
    elif baker_season == "Water" :
        move_index = sections["اوشية المياة"]
        water = find_slide_index_by_title(presentation1, "المياة", elkhetam)
        show_array.append([water, water])
    else:
        move_index = sections["أوشية الزروع"]
        tree = find_slide_index_by_title(presentation1, "الزروع", elkhetam)
        show_array.append([tree, tree])

    run_vba_with_slide_id_bakr_aashya(excel, des_sheet, prs, presentation1)

    show_slides(presentation1, show_array)

    if move_index < target_index:
        target_index -= 1
    presentation1.SectionProperties.Move(move_index, target_index + 1)

    # Initialize variables for current position, slide, and end index
    current_position = start_positions[0]
    current_start_slide = int(start_slides[0])
    current_end_slide = int(end_slides[0])

    # Initialize index for start position, slide, and end slide
    position_index = 1
    slide_index = 1
    end_index = 1

    while current_start_slide <= current_end_slide and slide_index <= presentation1.Slides.Count:
        if (current_position == elengil3 or current_position == elmzmor1):
            source_slide = presentation2.Slides(current_end_slide)
            source_slide.Copy()
            new_slide = presentation1.Slides.Paste(current_position).SlideShowTransition.Hidden = False
            current_end_slide -= 1
            if(current_start_slide > current_end_slide):
                current_position += 1

        elif guestBishop == True :
            source_slide = presentation3.Slides(current_end_slide)
            source_slide.Copy()
            presentation1.Slides.Paste(current_position)
            current_end_slide -= 1
            if(current_start_slide > current_end_slide):
                current_position += 1

        else:
            source_slide = presentation2.Slides(current_start_slide)
            source_slide.Copy()
            new_slide = presentation1.Slides.Paste(current_position)
            new_slide.SlideShowTransition.Hidden = False
            current_start_slide += 1
            current_position += 1

        # Move to the next round if all slides in the current range have been processed
        if current_start_slide > current_end_slide:
            # Check if there are more rounds
            if position_index < len(start_positions):
                # Update variables for the next round
                current_position = start_positions[position_index]
                current_start_slide = start_slides[slide_index]
                current_end_slide = end_slides[end_index]
                position_index += 1
                slide_index += 1
                end_index += 1

    presentation2.Close()
    if guestBishop > 0:
        presentation3.Close()

    presentation1.SlideShowSettings.Run()

def bakerKiahk(copticdate, adam = False, Bishop = False, guestBishop = 0):
    from copticDate import CopticCalendar
    cd = CopticCalendar().coptic_to_gregorian(copticdate)
    prs = relative_path(r"رفع بخور عشية و باكر.pptx")  # Using the relative path
    excel = relative_path(r"Files Data.xlsx")
    excel2 = relative_path(r"Tables.xlsx")
    des_sheet ="رفع بخور"
    replacefile(prs, relative_path(r"Data\CopyData\رفع بخور عشية و باكر.pptx"))

    elzoksologyat(excel, 5, "باكر")

    if cd.weekday() == 6:
        katamars = relative_path(r"Data\القطمارس\القطمارس السنوي احاد.pptx")
        katamars_sheet = "القطمارس السنوي آحاد"
        km = copticdate[1]
        kd = (copticdate[2] - 1) // 7 + 1
    else: 
        katamars = relative_path(r"Data\القطمارس\القطمارس السنوي ايام.pptx")
        katamars_sheet = "القطمارس السنوي أيام"
        km, kd = find_Readings_Date(copticdate[1], copticdate[2])

    katamars_values = fetch_data_arrays(excel2, katamars_sheet, km, kd, [6, 7, 8])
    elmzmor = katamars_values[0]
    elengil = katamars_values[1]
    elengil2 = katamars_values[2]

    # baker_kiahk_show_full_sections = ["تسبحة الملائكة", "تكملة مشتركة لكيهك", "ربع الملاك غبريال"]
    # baker_kiahk_hide_full_sections = ["مرد الانجيل السنوي", "اري ابرسفافين"]
    
    baker_kiahk_show_full_sections = ['{2ECE1F1B-C143-4CE2-B550-348BEE185974}', '{98BAFA87-CFC9-4304-8B00-99AA9EC72A3D}', '{E79471F5-6C95-4ECB-AD9D-CE660DCAAA70}']
    baker_kiahk_hide_full_sections = ['{BEECCC68-2AEF-4568-91AA-98BCD14D3B92}', '{49534D46-CF48-4D9B-ADBF-B776827F6670}']
    show_array = []

    # if adam:
    #     aashya_show_full_sections.append("ارباع الناقوس الادام")
    # else:
    #     aashya_show_full_sections.append("ارباع الناقوس الواطس")

    if adam:
        baker_kiahk_show_full_sections.append("{D02C088A-01E0-4A8C-8D73-21E3FD3616EB}")
    else:
        baker_kiahk_show_full_sections.append("{9495E38B-CE03-4E75-AED4-960DE95BA747}")

    #الاواشي
    if cd.weekday() == 5: #السبت
        # baker_show_full_sections.append("اوشية الراقدين")
        baker_kiahk_show_full_sections.append('{83E6BC33-A9EC-45CA-89B6-24EFBC51B654}')
    elif cd.weekday() == 6:#الاحد
        # baker_show_full_sections.extend(["اوشية المرضى", "اوشية القرابين"])
        baker_kiahk_show_full_sections.extend(["{069F7A79-999B-4223-82AE-CAF356118167}", "{2C897F14-44CC-430E-9BE1-EB379FE7A9C7}"])
    else:
        # baker_show_full_sections.extend(["اوشية المرضى", "اوشية المسافرين"])
        baker_kiahk_show_full_sections.extend(["{069F7A79-999B-4223-82AE-CAF356118167}", "{A059EEC9-5D25-453F-A956-A2E149F0773C}"])

    #مرد الانجيل
    if copticdate[2] <= 14:
        # kiahk_show_full_sections.extend(['مرد انجيل كيهك 1'])
        baker_kiahk_show_full_sections.extend(['{765066DC-37D0-48B3-9046-D4F98F5B05BC}'])
    else:
        # kiahk_show_full_sections.extend(['مرد انجيل كيهك 2'])
        baker_kiahk_show_full_sections.extend(['{D5E69BAC-0157-4B69-9255-B6775E2EE11D}'])

    # kiahk_aashya_values = ["الانجيل", "المزمور", "تكملة على حسب المناسبة"]
    kiahk_aashya_values = find_slide_nums_arrays_v2(excel, des_sheet, 
                    ['{B74DBB8C-2B2D-46E4-9508-DA46008D19A4}', '{6D1E6E7D-EECE-483C-A3AE-C135D02E717C}', '{A18EDC94-F257-4FAC-99C7-0A8EA70F0FAF}'], 
                    2, [2, 2, 1])
    
    #المزمور و الانجيل
    elengil3 = kiahk_aashya_values[0]
    elmzmor1 = kiahk_aashya_values[1]

    #الختام
    elkhetam = kiahk_aashya_values[2]

    if Bishop == True:
        prs3 = relative_path(r"Data\حضور الأسقف.pptx")
        sheet = "في حضور الأسقف"

        # bishop_show_values = ['تكملة في حضور الاسقف', 'مارو اتشاسف', 'فليرفعوه', 'في حضور الاسقف']

        bishop_show_values = ['{F76B0D75-0474-45B5-B79F-7416F354543A}', '{62A12AF8-CB6D-4CC5-9DB0-B73A7C24E2AD}',
                              '{23533FC3-43FE-456F-9454-70C3088055E7}', '{A9183893-7B7E-459F-8547-F7A8F7D2D521}']
                
        baker_kiahk_show_full_sections.extend(bishop_show_values)

        if guestBishop > 0:
            # bishop_values = ["صلاة الشكر", "صلاة الشكر", "طوبه هينا الكبيرة", "طوبه هينا الكبيرة", 
            #                  "نيم بينيوت", "نيم بينيوت", "الاسبسمس", "الاسبسمس"]

            # bishopDes_values = ["تكملة في حضور الاسقف", "طوبه هينا الكبيرة", "مارو اتشاسف"]
            
            bishop_values = find_slide_nums_arrays_v2(excel, sheet, 
                            ['{6851F163-CBEF-4014-A853-CE100557BA6A}', '{6851F163-CBEF-4014-A853-CE100557BA6A}',
                             '{B084BC40-61E1-4477-98DA-15CFB06AEE91}', '{B084BC40-61E1-4477-98DA-15CFB06AEE91}',
                             '{97203297-EECB-4D41-B2E3-AD9A4863847E}', '{97203297-EECB-4D41-B2E3-AD9A4863847E}', 
                             '{D1378DB5-29D1-4800-9D96-10F2535EEB57}', '{D1378DB5-29D1-4800-9D96-10F2535EEB57}'], 
                            2, [1, 2, 1, 2, 1, 2, 1, 2])

            bishopDes_values = find_slide_nums_arrays_v2(excel, des_sheet, 
                               ['{F76B0D75-0474-45B5-B79F-7416F354543A}', 
                                '{8DD21CDE-CB6B-4D5B-B995-D2747AB69ED1}', 
                                '{62A12AF8-CB6D-4CC5-9DB0-B73A7C24E2AD}'],                  
                               2, [2, 2, 2])

            elshokr = bishopDes_values[0]
            elshokr1 = bishop_values[0]
            elshokr2 = bishop_values[1]

            tobhyna = bishopDes_values[1]
            tobhyna1 = bishop_values[2]
            tobhyna2 = bishop_values[3]

            maro = bishopDes_values[2]
            maro1 = bishop_values[4]
            maro2 = bishop_values[5]

            if guestBishop < 2:
                elshokr2 = elshokr2-1
                tobhyna2 = tobhyna2-2
                maro2 = maro2-1
        
            start_positions = [elengil3, elmzmor1, maro, tobhyna, elshokr]
            start_slides = [elengil, elmzmor, maro1, tobhyna1, elshokr1]
            end_slides = [elengil2, elmzmor, maro2, tobhyna2, elshokr2]

    if guestBishop == 0:
        start_positions = [elengil3, elmzmor1]
        start_slides = [elengil, elmzmor]
        end_slides = [elengil2, elmzmor]
    
    if cd.weekday() == 6:
        show_hide_insertImage_replaceText(prs, excel, des_sheet, baker_kiahk_show_full_sections, baker_kiahk_hide_full_sections, new_Text=["لأنك قمت","aktwnk", "آك طونك"])
    else:
        show_hide_insertImage_replaceText(prs, excel, des_sheet, baker_kiahk_show_full_sections, baker_kiahk_hide_full_sections)

    powerpoint = win32com.client.Dispatch("PowerPoint.Application")
    powerpoint.Visible = True  # Open PowerPoint application
    presentation1 = open_presentation_relative_path(prs)
    presentation2 = open_presentation_relative_path(katamars)

    if guestBishop > 0:
        presentation3 = open_presentation_relative_path(prs3)

    sections = {presentation1.SectionProperties.Name(i): i for i in range(1, presentation1.SectionProperties.Count + 1)}
    target_index = sections["أوشية الموضع"]
    move_index = sections["أوشية الزروع"]
    kiahk = find_slide_index_by_title(presentation1, "صوم الميلاد", elkhetam)
    show_array.append([kiahk, kiahk])

    run_vba_with_slide_id_bakr_aashya(excel, des_sheet, prs, presentation1)
        
    show_slides(presentation1, show_array)
    
    if move_index < target_index:
        target_index -= 1
    presentation1.SectionProperties.Move(move_index, target_index + 1)

    # Initialize variables for current position, slide, and end index
    current_position = start_positions[0]
    current_start_slide = int(start_slides[0])
    current_end_slide = int(end_slides[0])

    # Initialize index for start position, slide, and end slide
    position_index = 1
    slide_index = 1
    end_index = 1

    while current_start_slide <= current_end_slide and slide_index <= presentation1.Slides.Count:
        if (current_position == elengil3 or current_position == elmzmor1):
            source_slide = presentation2.Slides(current_end_slide)
            source_slide.Copy()
            new_slide = presentation1.Slides.Paste(current_position).SlideShowTransition.Hidden = False
            current_end_slide -= 1
            if(current_start_slide > current_end_slide):
                current_position += 1

        elif guestBishop > 0:
            source_slide = presentation3.Slides(current_end_slide)
            source_slide.Copy()
            presentation1.Slides.Paste(current_position)
            current_end_slide -= 1
            if(current_start_slide > current_end_slide):
                current_position += 1

        else:
            source_slide = presentation2.Slides(current_start_slide)
            source_slide.Copy()
            new_slide = presentation1.Slides.Paste(current_position)
            new_slide.SlideShowTransition.Hidden = False
            current_start_slide += 1
            current_position += 1

        # Move to the next round if all slides in the current range have been processed
        if current_start_slide > current_end_slide:
            # Check if there are more rounds
            if position_index < len(start_positions):
                # Update variables for the next round
                current_position = start_positions[position_index]
                current_start_slide = start_slides[slide_index]
                current_end_slide = end_slides[end_index]
                position_index += 1
                slide_index += 1
                end_index += 1

    presentation2.Close()
    if Bishop == True:
        presentation3.Close()

    presentation1.SlideShowSettings.Run()

def bakerElmilad(season, copticdate, Bishop = False, guestBishop = 0):
    from copticDate import CopticCalendar
    cd = CopticCalendar().coptic_to_gregorian(copticdate)
    prs = relative_path(r"رفع بخور عشية و باكر.pptx")  # Using the relative path
    excel = relative_path(r"Files Data.xlsx")
    excel2 = relative_path(r"Tables.xlsx")
    des_sheet ="رفع بخور"
    img = relative_path(r"Data\Designs\الميلاد.png")
    replacefile(prs, relative_path(r"Data\CopyData\رفع بخور عشية و باكر.pptx"))

    elzoksologyat(excel, season, "باكر")

    katamars = relative_path(r"Data\القطمارس\القطمارس السنوي ايام.pptx")
    katamars_sheet = "القطمارس السنوي أيام"
    km, kd = find_Readings_Date(copticdate[1], copticdate[2])

    katamars_values = fetch_data_arrays(excel2, katamars_sheet, km, kd, [6, 7, 8])
    elmzmor = katamars_values[0]
    elengil = katamars_values[1]
    elengil2 = katamars_values[2]

    # baker_elmilad_show_full_sections = ["اوشية المرضى", "السبع طرائق", "اوشية القرابين", "تسبحة الملائكة", 
    #                                     "ارباع الناقوس الفرايحي", "ارباع عيد الميلاد", "مرد مزمور الميلاد",
    #                                     "ختام ارباع الناقوس الفرايحي", "مرد إنجيل عشية وباكر عيد الميلاد", "فاي اريه بي اوو"]
    # baker_elmilad_hide_full_sections = ["قدوس الصلب", "قدوس الصعود", "مرد الانجيل السنوي", "اري ابرسفافين", "ختام ارباع الناقوس السنوي"]

    baker_elmilad_show_full_sections = ['{B7D98377-B994-4654-B49C-DE10E0DDE4F1}', '{385BAEA4-D798-4AD0-920D-ADEC0B972186}', '{2C897F14-44CC-430E-9BE1-EB379FE7A9C7}', '{2ECE1F1B-C143-4CE2-B550-348BEE185974}', '{08277E26-4A59-441A-90BE-E448498302F7}', '{013099AC-E8DF-4B35-A585-A33DD38F2D00}', '{B87EBA1A-E0E4-4E68-87D7-3C4A798CF278}', '{DC185CA7-709B-4E23-B0E7-0543D701D531}', '{49683C84-CF98-4F29-AD91-1CF871381FD1}', '{069F7A79-999B-4223-82AE-CAF356118167}']
    baker_elmilad_hide_full_sections = ['{F5175F64-5E0E-4336-A80A-F953446C58B3}', '{AE3146ED-6425-4136-9613-0253836230D1}', '{BEECCC68-2AEF-4568-91AA-98BCD14D3B92}', '{49534D46-CF48-4D9B-ADBF-B776827F6670}', '{BE52E209-B022-45B3-8B56-2C35B7DE60D1}']

    # baker_elmilad_values = ["قدوس الميلاد", "مزمور باكر عيد الميلاد قبطي", "طواف مزمور عشية وباكر عيد الميلاد",
    #                         "المزمور", "الانجيل", "تكملة على حسب المناسبة"]

    baker_elmilad_values = find_slide_nums_arrays_v2(excel, des_sheet, ['{C361BE0A-6275-4F21-8ACA-B240ABDAEF5D}', '{215DD854-9794-4375-BC3C-48AB43CE6E80}', '{978B4431-5038-42EA-A7C5-E7190483D4CF}', '{6D1E6E7D-EECE-483C-A3AE-C135D02E717C}', '{B74DBB8C-2B2D-46E4-9508-DA46008D19A4}', '{A18EDC94-F257-4FAC-99C7-0A8EA70F0FAF}'], 
                                                     2, [1, 2, 2, 2, 2, 1])

    Elkhetam = baker_elmilad_values[5]

    elengil3 = baker_elmilad_values[4]
    elmzmor1 = baker_elmilad_values[3]

    tawaf_mazmor_elmilad = baker_elmilad_values[2]
    mazmor_baker_elmilad = baker_elmilad_values[1]

    agios = baker_elmilad_values[0]

    if Bishop == True:
        prs3 = relative_path(r"Data\حضور الأسقف.pptx")
        sheet = "في حضور الأسقف"

        # bishop_show_values = ['تكملة في حضور الاسقف', 'مارو اتشاسف', 'فليرفعوه', 'في حضور الاسقف']

        bishop_show_values = ['{F76B0D75-0474-45B5-B79F-7416F354543A}', '{62A12AF8-CB6D-4CC5-9DB0-B73A7C24E2AD}',
                              '{23533FC3-43FE-456F-9454-70C3088055E7}', '{A9183893-7B7E-459F-8547-F7A8F7D2D521}']
                
        baker_elmilad_show_full_sections.extend(bishop_show_values)

        if guestBishop > 0:
            # bishop_values = ["صلاة الشكر", "صلاة الشكر", "طوبه هينا الكبيرة", "طوبه هينا الكبيرة", 
            #                  "نيم بينيوت", "نيم بينيوت", "الاسبسمس", "الاسبسمس"]

            # bishopDes_values = ["تكملة في حضور الاسقف", "طوبه هينا الكبيرة", "مارو اتشاسف"]
            
            bishop_values = find_slide_nums_arrays_v2(excel, sheet, 
                            ['{6851F163-CBEF-4014-A853-CE100557BA6A}', '{6851F163-CBEF-4014-A853-CE100557BA6A}',
                             '{B084BC40-61E1-4477-98DA-15CFB06AEE91}', '{B084BC40-61E1-4477-98DA-15CFB06AEE91}',
                             '{97203297-EECB-4D41-B2E3-AD9A4863847E}', '{97203297-EECB-4D41-B2E3-AD9A4863847E}', 
                             '{D1378DB5-29D1-4800-9D96-10F2535EEB57}', '{D1378DB5-29D1-4800-9D96-10F2535EEB57}'], 
                            2, [1, 2, 1, 2, 1, 2, 1, 2])

            bishopDes_values = find_slide_nums_arrays_v2(excel, des_sheet, 
                               ['{F76B0D75-0474-45B5-B79F-7416F354543A}', 
                                '{8DD21CDE-CB6B-4D5B-B995-D2747AB69ED1}', 
                                '{62A12AF8-CB6D-4CC5-9DB0-B73A7C24E2AD}'],                  
                               2, [2, 2, 2])

            elshokr = bishopDes_values[0]
            elshokr1 = bishop_values[0]
            elshokr2 = bishop_values[1]

            tobhyna = bishopDes_values[1]
            tobhyna1 = bishop_values[2]
            tobhyna2 = bishop_values[3]

            maro = bishopDes_values[2]
            maro1 = bishop_values[4]
            maro2 = bishop_values[5]

            if guestBishop < 2:
                elshokr2 = elshokr2-1
                tobhyna2 = tobhyna2-2
                maro2 = maro2-1
        
            start_positions = [elengil3, elmzmor1, maro, agios, agios, tobhyna, elshokr]
            start_slides = [elengil, elmzmor, maro1, agios, agios, tobhyna1, elshokr1]
            end_slides = [elengil2, elmzmor, maro2, agios, agios, tobhyna2, elshokr2]

    if guestBishop == 0:
        start_positions = [elengil3, elmzmor1, agios, agios]
        start_slides = [elengil, elmzmor, agios, agios]
        end_slides = [elengil2, elmzmor, agios, agios]

    show_hide_insertImage_replaceText(prs, excel, des_sheet, baker_elmilad_show_full_sections, baker_elmilad_hide_full_sections, image_path=img, new_Text=["لأنك ولِدتَ","aumack","آف ماسك"])

    powerpoint = win32com.client.Dispatch("PowerPoint.Application")
    powerpoint.Visible = True  # Open PowerPoint application
    presentation1 = open_presentation_relative_path(prs)
    presentation2 = open_presentation_relative_path(katamars)

    if guestBishop > 0:
        presentation3 = open_presentation_relative_path(prs3)
            
    show_array = [[mazmor_baker_elmilad, mazmor_baker_elmilad], [tawaf_mazmor_elmilad, tawaf_mazmor_elmilad]]
    
    khetam_elmilad = find_slide_index_by_title(presentation1, "الميلاد", Elkhetam)
    show_array.append([khetam_elmilad, khetam_elmilad])

    run_vba_with_slide_id_bakr_aashya(excel, des_sheet, prs, presentation1)

    show_slides(presentation1, show_array)

    # Initialize variables for current position, slide, and end index
    current_position = start_positions[0]
    current_start_slide = int(start_slides[0])
    current_end_slide = int(end_slides[0])

    # Initialize index for start position, slide, and end slide
    position_index = 1
    slide_index = 1
    end_index = 1

    while current_start_slide <= current_end_slide and slide_index <= presentation1.Slides.Count:
        if (current_position == elengil3 or current_position == elmzmor1):
            source_slide = presentation2.Slides(current_end_slide)
            source_slide.Copy()
            new_slide = presentation1.Slides.Paste(current_position)
            pic_shape = new_slide.Shapes.AddPicture(img, LinkToFile=False, SaveWithDocument=True, Left=0, Top=144, Width=720, Height=396)
            pic_shape.ZOrder(1)
            new_slide.SlideShowTransition.Hidden = False
            current_end_slide -= 1

        elif current_position == agios :
            source_slide = presentation1.Slides(current_end_slide)
            source_slide.Copy()
            presentation1.Slides.Paste(current_position)
            current_end_slide -= 1
            if(current_start_slide > current_end_slide):
                current_position += 1

        elif Bishop and current_position in {maro, tobhyna, elshokr}:
            source_slide = presentation3.Slides(current_end_slide)
            source_slide.Copy()
            presentation1.Slides.Paste(current_position)
            current_end_slide -= 1
            if(current_start_slide > current_end_slide):
                current_position += 1

        else:
            source_slide = presentation2.Slides(current_start_slide)
            source_slide.Copy()
            new_slide = presentation1.Slides.Paste(current_position)
            new_slide.SlideShowTransition.Hidden = False
            current_start_slide += 1
            current_position += 1

        # Move to the next round if all slides in the current range have been processed
        if current_start_slide > current_end_slide:
            # Check if there are more rounds
            if position_index < len(start_positions):
                # Update variables for the next round
                current_position = start_positions[position_index]
                current_start_slide = start_slides[slide_index]
                current_end_slide = end_slides[end_index]
                position_index += 1
                slide_index += 1
                end_index += 1

    move_sections_v2(presentation1, ['{A0DD24B6-1053-42A1-8391-038649B3219B}'], ['{9DACDA4B-74BF-4F4B-A961-DA43C70DC545}'])

    presentation2.Close()
    if guestBishop > 0:
        presentation3.Close()

    presentation1.SlideShowSettings.Run()

def bakerElSomElkbyr(copticdate, Bishop = False, guestBishop = 0):
    from copticDate import CopticCalendar
    cd = CopticCalendar().coptic_to_gregorian(copticdate)
    prs = relative_path(r"رفع بخور عشية و باكر.pptx")  # Using the relative path
    excel = relative_path(r"Files Data.xlsx")
    excel2 = relative_path(r"Tables.xlsx")
    sheet ="رفع بخور"
    replacefile(prs, relative_path(r"Data\CopyData\رفع بخور عشية و باكر.pptx"))

