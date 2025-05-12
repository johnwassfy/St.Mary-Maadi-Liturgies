import os
from commonFunctions import *
import win32com.client

# def Aashya(copticdate, adam = False):
#     from copticDate import CopticCalendar
#     cd = CopticCalendar().coptic_to_gregorian(copticdate)
#     prs = relative_path(r"عشية.pptx")  # Using the relative path
#     excel = relative_path(r"بيانات القداسات.xlsx")
#     excel2 = relative_path(r"Tables.xlsx")
#     sheet ="عشية"

#     if cd.weekday() == 6:
#         sunday()
#         katamars = relative_path(r"Data\القطمارس\الاحاد\القطمارس السنوي احاد (عشية).pptx")
#         katamars_sheet = "قطمارس الاحاد للعشية"
#         km = copticdate[1]
#         kd = (copticdate[2] - 1) // 7 + 1
#     else: 
#         katamars = relative_path(r"Data\القطمارس\الايام\القطمارس السنوي ايام (عشية).pptx")
#         katamars_sheet = "القطمارس السنوي العشية"
#         km, kd = find_Readings_Date(copticdate[1], copticdate[2])

#     katamars_values = fetch_data_arrays(excel2, katamars_sheet, km, kd, [3, 4, 5])
#     elmzmor = katamars_values[0]
#     elengil = katamars_values[1]
#     elengil2 = katamars_values[2]

#     aashya_values = find_slide_nums_arrays(excel, sheet, ["الانجيل", "الانجيل","ارباع الناقوس الادام", "ارباع الناقوس الادام", 
#                                                          "أرباع الناقوس الواطس", "أرباع الناقوس الواطس"], 
#                                                         [1, 2, 1, 2, 1, 2])

#     #المزمور و الانجيل
#     elengil3 = aashya_values[1]
#     elmzmor1 = aashya_values[0] + 2

#     #ارباع الناقوس
#     arbaaAdam = aashya_values[2]
#     arbaaAdam2 = aashya_values[3]
#     arbaaWats = aashya_values[4]
#     arbaaWats2 = aashya_values[5]

#     show_array = [[1, 1]]
#     hide_array = [[1, 1]]

#     if adam :
#         show_array.append([arbaaAdam, arbaaAdam2])
#         hide_array.append([arbaaWats, arbaaWats2])

#     if (copticdate[1] == 12 and copticdate[2]<=16):
#         el3adra_values = find_slide_nums_arrays(excel, sheet,
#                                                 ["أفرحى يا مريم", "أفرحى يا مريم", "لحن شيري ماريا", "اسمعي يا ابنة",
#                                                  "طواف مزمور عشية صوم العذراء", "طواف مزمور عشية صوم العذراء",
#                                                  "مرد الانجيل عشية", "مرد الانجيل عشية", "مرد الانجيل صوم العذراء", "مرد الانجيل صوم العذراء", 
#                                                  "مرد مزمور التجلي", "مرد انجيل التجلي", "ختام ارباع الناقوس", "ختام ارباع الناقوس",
#                                                  "ختام ارباع الناقوس الفرايحي", "ختام ارباع الناقوس الفرايحي", 
#                                                  "تكملة ارباع الناقوس", "ذكصولوجية التجلي", "ذكصولوجية التجلي", "فاي اريه بي اوو"],
#                                                 [1, 2, 1, 2, 1, 2, 1, 2, 1, 2, 1, 1, 1, 2, 1, 2, 1, 1, 2, 1])
#         efra7y = el3adra_values[0]
#         efra7y2 = el3adra_values[1]

#         tamagyd = el3adra_values[2]
#         tamagyd2 = el3adra_values[3]
        
#         if copticdate[2] == 13 :
#             mrdelmazmor = el3adra_values[10]
#             mrdelmazmor2 = el3adra_values[10]

#             mrdengilnew = el3adra_values[11]
#             mrdengilnew2 = el3adra_values[19]
            
#             khetamArbaa = el3adra_values[12] + 1
#             khetamArbaa2 = el3adra_values[13]

#             khetamArbaaFaray7y = el3adra_values[14]
#             khetamArbaaFaray7y2 = el3adra_values[15]

#             zoksologya = el3adra_values[17]
#             zoksologya2 = el3adra_values[18]
            
#             hide_array.append([khetamArbaa, khetamArbaa2])
#             show_array.extend([[khetamArbaaFaray7y, khetamArbaaFaray7y2], [zoksologya, zoksologya2]])
#         else:
#             mrdelmazmor = el3adra_values[4]
#             mrdelmazmor2 = el3adra_values[5]
            
#             mrdengilnew = el3adra_values[8]
#             mrdengilnew2 = el3adra_values[9]
    
#         mrdelengil = el3adra_values[6]
#         mrdelengil2 = el3adra_values[7]

#         if copticdate == None:
#             copticdate = CopticCalendar().gregorian_to_coptic()
#             season = CopticCalendar().get_coptic_date_range(copticdate)
#         else:
#             season = CopticCalendar().get_coptic_date_range(copticdate)

#         show_array.extend([[efra7y, efra7y2], [tamagyd, tamagyd2], [mrdelmazmor, mrdelmazmor2], [mrdengilnew, mrdengilnew2]])
#         hide_array.extend([[mrdelengil, mrdelengil2]])

#     start_positions = [elengil3, elmzmor1]
#     start_slides = [elengil, elmzmor]
#     end_slides = [elengil2, elmzmor]
    
#     powerpoint = win32com.client.Dispatch("PowerPoint.Application")
#     powerpoint.Visible = True  # Open PowerPoint application
#     presentation1 = open_presentation_relative_path(prs)
#     presentation2 = open_presentation_relative_path(katamars)

#     if copticdate[1] == 12 and copticdate[2] == 13:
#         rt = find_slide_index_by_label(presentation1, "التجلي", el3adra_values[16])
#         rt2 = find_slide_index_by_label(presentation1, "التجلي 2", el3adra_values[16])
#         show_array.append([rt, rt2])
    
#     hide_slides(presentation1, hide_array)
#     show_slides(presentation1, show_array)

#     sections = {presentation1.SectionProperties.Name(i): i for i in range(1, presentation1.SectionProperties.Count + 1)}
#     target_index = sections["اوشية الموضع"]
#     if season == "Air" :
#         move_index = sections["اوشية الاهوية"]
#     elif season == "Water" :
#         move_index = sections["اوشية المياة"]
#     else:
#         move_index = sections["اوشية الزروع و العشب"]
#     if move_index < target_index:
#         target_index -= 1
#     presentation1.SectionProperties.Move(move_index, target_index + 1)

#     # Initialize variables for current position, slide, and end index
#     current_position = start_positions[0]
#     current_start_slide = int(start_slides[0])
#     current_end_slide = int(end_slides[0])

#     # Initialize index for start position, slide, and end slide
#     position_index = 1
#     slide_index = 1
#     end_index = 1

#     while current_start_slide <= current_end_slide and slide_index <= presentation1.Slides.Count:
#         if (current_position == elengil3 or current_position == elmzmor1):
#             source_slide = presentation2.Slides(current_end_slide)
#             source_slide.Copy()
#             new_slide = presentation1.Slides.Paste(current_position).SlideShowTransition.Hidden = False
#             current_end_slide -= 1
#             if(current_start_slide > current_end_slide):
#                 current_position += 1

#         else:
#             source_slide = presentation2.Slides(current_start_slide)
#             source_slide.Copy()
#             new_slide = presentation1.Slides.Paste(current_position)
#             new_slide.SlideShowTransition.Hidden = False
#             current_start_slide += 1
#             current_position += 1

#         # Move to the next round if all slides in the current range have been processed
#         if current_start_slide > current_end_slide:
#             # Check if there are more rounds
#             if position_index < len(start_positions):
#                 # Update variables for the next round
#                 current_position = start_positions[position_index]
#                 current_start_slide = start_slides[slide_index]
#                 current_end_slide = end_slides[end_index]
#                 position_index += 1
#                 slide_index += 1
#                 end_index += 1

#     presentation2.Close()

def aashyaSanawy(season, copticdate, adam = False, Bishop = False, guestBishop = 0):
    from copticDate import CopticCalendar
    cd = CopticCalendar().coptic_to_gregorian(copticdate)
    prs = relative_path(r"رفع بخور عشية و باكر.pptx")  # Using the relative path
    excel = relative_path(r"بيانات القداسات.xlsx")
    excel2 = relative_path(r"Tables.xlsx")
    des_sheet ="رفع بخور"
    replacefile(prs, relative_path(r"Data\CopyData\رفع بخور عشية و باكر.pptx"))

    elzoksologyat(excel, season, "عشية")

    if cd.weekday() == 6:
        # sunday()
        katamars = relative_path(r"Data\القطمارس\الاحاد\القطمارس السنوي احاد (عشية).pptx")
        katamars_sheet = "قطمارس الاحاد للعشية"
        km = copticdate[1]
        kd = (copticdate[2] - 1) // 7 + 1
    else: 
        katamars = relative_path(r"Data\القطمارس\الايام\القطمارس السنوي ايام (عشية).pptx")
        katamars_sheet = "القطمارس السنوي العشية"
        km, kd = find_Readings_Date(copticdate[1], copticdate[2])

    katamars_values = fetch_data_arrays(excel2, katamars_sheet, km, kd, [3, 4, 5])
    elmzmor = katamars_values[0]
    elengil = katamars_values[1]
    elengil2 = katamars_values[2]

    # aashya_show_full_sections = ["اوشية الراقدين", "تفضل يا رب"]
    
    aashya_show_full_sections = ['{83E6BC33-A9EC-45CA-89B6-24EFBC51B654}', '{D8AFA182-3999-4072-94E9-F65D50B876B9}']

    # if adam:
    #     aashya_show_full_sections.append("ارباع الناقوس الادام")
    # else:
    #     aashya_show_full_sections.append("ارباع الناقوس الواطس")

    if adam:
        aashya_show_full_sections.append("{D02C088A-01E0-4A8C-8D73-21E3FD3616EB}")
    else:
        aashya_show_full_sections.append("{9495E38B-CE03-4E75-AED4-960DE95BA747}")

    # aashya_values = ["الانجيل", "المزمور", "تكملة على حسب المناسبة"]

    aashya_values = find_slide_nums_arrays_v2(excel, des_sheet, 
                    ['{B74DBB8C-2B2D-46E4-9508-DA46008D19A4}', '{6D1E6E7D-EECE-483C-A3AE-C135D02E717C}', '{A18EDC94-F257-4FAC-99C7-0A8EA70F0FAF}'], 
                    2, [2, 2, 1])

    #المزمور و الانجيل
    elengil3 = aashya_values[0]
    elmzmor1 = aashya_values[1]

    #الختام
    elkhetam = aashya_values[2]

    awashy_season = CopticCalendar().get_coptic_date_range(copticdate)

    if Bishop == True:
        prs3 = relative_path(r"Data\حضور الأسقف.pptx")
        sheet = "في حضور الأسقف"

        # bishop_show_values = ['تكملة في حضور الاسقف', 'مارو اتشاسف', 'فليرفعوه', 'في حضور الاسقف']

        bishop_show_values = ['{F76B0D75-0474-45B5-B79F-7416F354543A}', '{62A12AF8-CB6D-4CC5-9DB0-B73A7C24E2AD}',
                              '{23533FC3-43FE-456F-9454-70C3088055E7}', '{A9183893-7B7E-459F-8547-F7A8F7D2D521}']
                
        aashya_show_full_sections.extend(bishop_show_values)

        if guestBishop > 0:
            # bishop_values = ["صلاة الشكر", "صلاة الشكر", "طوبه هينا الكبيرة", "طوبه هينا الكبيرة", 
            #                  "نيم بينيوت", "نيم بينيوت", "الاسبسمس", "الاسبسمس"]

            bishopDes_values = ["تكملة في حضور الاسقف", "طوبه هينا الكبيرة", "مارو اتشاسف"]
            
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

            maro = bishopDes_values[8]
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
        show_hide_insertImage_replaceText(prs, excel, des_sheet, aashya_show_full_sections, new_Text=["لأنك قمت","aktwnk", "آك طونك"])
    else:
        show_hide_insertImage_replaceText(prs, excel, des_sheet, aashya_show_full_sections)
    powerpoint = win32com.client.Dispatch("PowerPoint.Application")
    powerpoint.Visible = True  # Open PowerPoint application
    presentation1 = open_presentation_relative_path(prs)
    presentation2 = open_presentation_relative_path(katamars)

    if guestBishop > 0:
        presentation3 = open_presentation_relative_path(prs3)

    sections = {presentation1.SectionProperties.Name(i): i for i in range(1, presentation1.SectionProperties.Count + 1)}
    target_index = sections["أوشية الموضع"]
    show_array = []
    if awashy_season == "Air" :
        move_index = sections["اوشية الأهوية والثمار"]
        air = find_slide_index_by_title(presentation1, "الاهوية", elkhetam)
        show_array.append([air, air])
    elif awashy_season == "Water" :
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

def aashyaKiahk(copticdate, adam = False, Bishop = False, guestBishop = 0):
    from copticDate import CopticCalendar
    cd = CopticCalendar().coptic_to_gregorian(copticdate)
    prs = relative_path(r"رفع بخور عشية و باكر.pptx")  # Using the relative path
    excel = relative_path(r"بيانات القداسات.xlsx")
    excel2 = relative_path(r"Tables.xlsx")
    sheet ="رفع بخور"

    elzoksologyat(excel, 5, "عشية")

    if cd.weekday() == 6:
        katamars = relative_path(r"Data\القطمارس\الاحاد\القطمارس السنوي احاد (عشية).pptx")
        katamars_sheet = "قطمارس الاحاد للعشية"
        km = copticdate[1]
        kd = (copticdate[2] - 1) // 7 + 1
    else: 
        katamars = relative_path(r"Data\القطمارس\الايام\القطمارس السنوي ايام (عشية).pptx")
        katamars_sheet = "القطمارس السنوي العشية"
        km, kd = find_Readings_Date(copticdate[1], copticdate[2])

    katamars_values = fetch_data_arrays(excel2, katamars_sheet, km, kd, [3, 4, 5])
    elmzmor = katamars_values[0]
    elengil = katamars_values[1]
    elengil2 = katamars_values[2]

    aashya_values = find_slide_nums_arrays(excel, sheet, ["الانجيل", "الانجيل","ارباع الناقوس الادام", "ارباع الناقوس الادام", 
                                                          "أرباع الناقوس الواطس", "أرباع الناقوس الواطس", "تكملة ارباع الناقوس",
                                                          "اوشية الراقدين", "اوشية الراقدين", "تفضل يا رب", "تفضل يا رب",
                                                          "مرد انجيل كيهك 1", "مرد انجيل كيهك 1", "مرد انجيل كيهك 2", "مرد انجيل كيهك 2",
                                                          "تكملة على حسب المناسبة", "مرد الانجيل السنوي", "مرد الانجيل السنوي",
                                                          "تكملة مشتركة لكيهك", "تكملة مشتركة لكيهك"], 
                                                         [1, 2, 1, 2, 1, 2, 1, 1, 2, 1, 2, 1, 2, 1, 2, 1, 1, 2, 1, 2])

    #المزمور و الانجيل
    elengil3 = aashya_values[1]
    elmzmor1 = aashya_values[0] + 1

    #ارباع الناقوس
    arbaaAdam = aashya_values[2]
    arbaaAdam2 = aashya_values[3]
    arbaaWats = aashya_values[4]
    arbaaWats2 = aashya_values[5]
    arab3elna2os = aashya_values[6]

    #الراقدين و تفضل يا رب
    elrakdyn = aashya_values[7]
    elrakdyn2 = aashya_values[8]
    tfdlyarb = aashya_values[9]
    tfdlyarb2 = aashya_values[10]

    #مرد الانجيل    
    if copticdate[2] <= 14:
        mrdelengil = aashya_values[11]
        mrdelengil2 = aashya_values[12]
    else:
        mrdelengil = aashya_values[13]
        mrdelengil2 = aashya_values[14]

    mrdelengilSanawy = aashya_values[16]
    mrdelengilSanawy2 = aashya_values[17]
    takmelaMrdelengil = aashya_values[18]
    takmelaMrdelengil2 = aashya_values[19]

    #الختام
    elkhetam = aashya_values[15]

    show_array = [[elrakdyn, elrakdyn2], [tfdlyarb, tfdlyarb2], [mrdelengil, mrdelengil2], [takmelaMrdelengil, takmelaMrdelengil2]]
    hide_array = [[mrdelengilSanawy, mrdelengilSanawy2]]

    if adam :
        show_array.append([arbaaAdam, arbaaAdam2])
        hide_array.append([arbaaWats, arbaaWats2])

    if Bishop == True:
        prs3 = relative_path(r"Data\حضور الأسقف.pptx")
        bishopSheet = "في حضور الأسقف"

        bishop_values = find_slide_nums_arrays(excel, bishopSheet, ["صلاة الشكر", "صلاة الشكر"], 
                                                                   [1, 2])
        
        bishopDes_values = find_slide_nums_arrays(excel, sheet, ["صلاة الشكر", "اوشية الآباء", 
                                                                 "في حضور الاسقف", "في حضور الاسقف"],
                                                                [2, 2, 1, 2])
        
        elshokr1 = bishop_values[0]
        elshokrDes = bishopDes_values[0] - 1

        elaba2Des = bishopDes_values[1] - 2

        elkhetamBishop = bishopDes_values[2]
        elkhetamBishop2 = bishopDes_values[3]

        if guestBishop > 0:
            if guestBishop == 1:
                elshokr2 = bishop_values[1] - 1
                elaba2 = elshokr2
                elaba22 = elshokr2

            elif guestBishop == 2:
                elshokr2 = bishop_values[1]
                elaba2 = elshokr2 - 1
                elaba22 = elshokr2

            start_positions = [elaba2Des, elengil3, elmzmor1, elshokrDes]
            start_slides = [elaba2, elengil, elmzmor, elshokr1]
            end_slides = [elaba22, elengil2, elmzmor, elshokr2]
        else:
            elshokr2 = find_slide_num(excel, bishopSheet, "صلاة الشكر", 2) - 2
            start_positions = [elengil3, elmzmor1, elshokrDes]
            start_slides = [elengil, elmzmor, elshokr1]
            end_slides = [elengil2, elmzmor, elshokr2]

        show_array.append([elkhetamBishop, elkhetamBishop2])
    else:
        start_positions = [elengil3, elmzmor1]
        start_slides = [elengil, elmzmor]
        end_slides = [elengil2, elmzmor]

    powerpoint = win32com.client.Dispatch("PowerPoint.Application")
    powerpoint.Visible = True  # Open PowerPoint application
    presentation1 = open_presentation_relative_path(prs)
    presentation2 = open_presentation_relative_path(katamars)
    if Bishop == True:
        presentation3 = open_presentation_relative_path(prs3)

    arab3elna2os_malakGhobrial = find_slide_indices_by_ordered_labels(presentation1, ["الملاك غبريال", "الملاك غبريال 2", "الملاك غبريال المبشر"], arab3elna2os)

    rob3 = arab3elna2os_malakGhobrial[0]
    rob32 = arab3elna2os_malakGhobrial[1]
    nos_rob3 = arab3elna2os_malakGhobrial[2]

    show_array.append([rob3, rob32])
    hide_array.append([nos_rob3, nos_rob3])

    sections = {presentation1.SectionProperties.Name(i): i for i in range(1, presentation1.SectionProperties.Count + 1)}
    target_index = sections["أوشية الموضع"]
    move_index = sections["أوشية الزروع"]
    kiahk = find_slide_index_by_title(presentation1, "صوم الميلاد", elkhetam)
    show_array.append([kiahk, kiahk])

    run_vba_with_slide_id_bakr_aashya(excel, sheet, prs, presentation1)
        
    hide_slides(presentation1, hide_array)
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

        elif Bishop == True :
            source_slide = presentation3.Slides(current_end_slide)
            source_slide.Copy()
            if current_position == elaba2Des:
                presentation1.Slides.Paste(current_position).SlideShowTransition.Hidden = True
            else:
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

