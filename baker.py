import os
from commonFunctions import *
import win32com.client

# def remove_background_picture(prs):
#     presentation = Presentation(relative_path(prs))
    
#     # Check if the first slide has a picture in it
#     first_slide = presentation.slides[0]
#     shapes = first_slide.shapes
#     has_background_picture = False
    
#     # Check if there's a picture in the background
#     for shape in shapes:
#         if shape.shape_type == 13:  # Shape type 13 represents a picture
#             has_background_picture = True
#             break
    
#     # If the first slide has a picture, proceed to remove background pictures from all slides
#     if has_background_picture:
#         for slide in presentation.slides:
#             shapes = slide.shapes
#             background_picture = None
            
#             # Find the picture that's in the background
#             for shape in shapes:
#                 if shape.shape_type == 13:  # Shape type 13 represents a picture
#                     if background_picture is None or shape.left < background_picture.left:
#                         background_picture = shape
            
#             # Remove the background picture
#             if background_picture is not None:
#                 background_picture._element.getparent().remove(background_picture._element)    
#     presentation.save(relative_path(prs))

def baker3ydElrosol(adam = False):
    prs = r"باكر.pptx"
    katamars = r"Data\القطمارس\الايام\القطمارس السنوي ايام (باكر).pptx"
    excel = relative_path(r"بيانات القداسات.xlsx")
    excel2 = relative_path(r"Tables.xlsx")
    sheet ="باكر"
    katamars_sheet = "القطمارس السنوي باكر"
    km, kd = find_Readings_Date(11, 5)
    katamars_values = fetch_data_arrays(excel2, katamars_sheet, km, kd, [3, 4, 5])
    elmzmor = katamars_values[0]
    elengil = katamars_values[1]
    elengil2 = katamars_values[2]

    baker_values = find_slide_nums_arrays(excel, sheet, ["ربع يقال في صوم الرسل", "الانجيل", "الانجيل",
                                                        "بطرس و بولس","بطرس و بولس",
                                                        "ارباع الناقوس الادام", "ارباع الناقوس الادام", 
                                                        "أرباع الناقوس الواطس", "أرباع الناقوس الواطس", 
                                                        "تكملة ارباع الناقوس 2"], 
                                                        [1, 2, 1, 1, 2, 1, 2, 1, 2, 1])

    #مرد الانجيل
    mrdelengil = baker_values[0]

    #المزمور و الانجيل
    elengil3 = baker_values[1]
    elmzmor1 = baker_values[2] + 2

    #بطرس و بولس
    pnp = baker_values[3]
    pnp2 = baker_values[4]

    #ارباع الناقوس
    arbaaAdam = baker_values[5]
    arbaaAdam2 = baker_values[6]
    arbaaWats = baker_values[7]
    arbaaWats2 = baker_values[8]
    rob3pnp = baker_values[9]

    if adam == True:
        show_array = [[mrdelengil, mrdelengil], [pnp, pnp2], [arbaaAdam, arbaaAdam2]]
        hide_array = [[arbaaWats, arbaaWats2]]
    else:
        show_array = [[mrdelengil, mrdelengil], [pnp, pnp2]]

    start_positions = [elengil3, elmzmor1]
    start_slides = [elengil, elmzmor]
    end_slides = [elengil2, elmzmor]

    powerpoint = win32com.client.Dispatch("PowerPoint.Application")
    powerpoint.Visible = True  # Open PowerPoint application
    presentation1 = open_presentation_relative_path(prs)
    presentation2 = open_presentation_relative_path(katamars)
    pnpRob3 = find_slide_index_by_title(presentation1, "السلام لأبينا بطرس: ومعلمنا بولس: العمودين العظيمين: مثبتي المؤمنين.", rob3pnp)
    show_array.append([pnpRob3, pnpRob3])

    show_slides(presentation1, show_array)
    if adam == True: 
        hide_slides(presentation1, hide_array)

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

def bakerSanawy(season, copticdate, adam = False, Bishop = False, guestBishop = 0):
    from copticDate import CopticCalendar
    cd = CopticCalendar().coptic_to_gregorian(copticdate)
    prs = relative_path(r"رفع بخور عشية و باكر.pptx")  # Using the relative path
    excel = relative_path(r"بيانات القداسات.xlsx")
    excel2 = relative_path(r"Tables.xlsx")
    sheet ="رفع بخور"
    replacefile(prs, relative_path(r"Data\CopyData\رفع بخور عشية و باكر.pptx"))
    
    elzoksologyat(excel, season,"باكر")

    if cd.weekday() == 6:
        sunday(prs)
        katamars = relative_path(r"Data\القطمارس\الاحاد\القطمارس السنوي احاد (باكر).pptx")
        katamars_sheet = "قطمارس الاحاد لباكر"
        km = copticdate[1]
        kd = (copticdate[2] - 1) // 7 + 1
    else: 
        katamars = relative_path(r"Data\القطمارس\الايام\القطمارس السنوي ايام (باكر).pptx")
        katamars_sheet = "القطمارس السنوي باكر"
        km, kd = find_Readings_Date(copticdate[1], copticdate[2])

    katamars_values = fetch_data_arrays(excel2, katamars_sheet, km, kd, [3, 4, 5])
    elmzmor = katamars_values[0]
    elengil = katamars_values[1]
    elengil2 = katamars_values[2]

    baker_values = find_slide_nums_arrays(excel, sheet, ["الانجيل", "الانجيل","ارباع الناقوس الادام", "ارباع الناقوس الادام", 
                                                          "أرباع الناقوس الواطس", "أرباع الناقوس الواطس", 
                                                          "اوشية الراقدين", "اوشية الراقدين", "اوشية المرضي", "اوشية المرضي",
                                                          "اوشية المسافرين", "اوشية المسافرين", "اوشية القرابين", "اوشية القرابين",
                                                          "فلنسبح مع الملائكة", "فلنسبح مع الملائكة", "تكملة على حسب المناسبة"], 
                                                         [1, 2, 1, 2, 1, 2, 1, 2, 1, 2, 1, 2, 1, 2, 1, 2, 1])

    #المزمور و الانجيل
    elengil3 = baker_values[1]
    elmzmor1 = baker_values[0] + 1

    #ارباع الناقوس
    arbaaAdam = baker_values[2]
    arbaaAdam2 = baker_values[3]
    arbaaWats = baker_values[4]
    arbaaWats2 = baker_values[5]

    #الاواشي
    if cd.weekday() == 5:
        elawashy = baker_values[6]
        elawashy2 = baker_values[7]
        elawashy22 = 1
        elawashy222 = 1
    elif cd.weekday() == 6:
        elawashy = baker_values[8]
        elawashy2 = baker_values[9]
        elawashy22 = baker_values[12]
        elawashy222 = baker_values[13]
    else:
        elawashy = baker_values[8]
        elawashy2 = baker_values[9]
        elawashy22 = baker_values[10]
        elawashy222 = baker_values[11]

    angels = baker_values[14]
    angels2 = baker_values[15]

    #الختام
    elkhetam = baker_values[16]

    show_array = [[1, 1], [elawashy, elawashy2], [elawashy22, elawashy222], [angels, angels2]]
    hide_array = [[1, 1]]

    if adam :
        show_array.append([arbaaAdam, arbaaAdam2])
        hide_array.append([arbaaWats, arbaaWats2])

    season = CopticCalendar().get_coptic_date_range(copticdate)

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

    sections = {presentation1.SectionProperties.Name(i): i for i in range(1, presentation1.SectionProperties.Count + 1)}
    target_index = sections["أوشية الموضع"]
    if season == "Air" :
        move_index = sections["اوشية الأهوية والثمار"]
        air = find_slide_index_by_title(presentation1, "الاهوية", elkhetam)
        show_array.append([air, air])
    elif season == "Water" :
        move_index = sections["اوشية المياة"]
        water = find_slide_index_by_title(presentation1, "المياة", elkhetam)
        show_array.append([water, water])
    else:
        move_index = sections["أوشية الزروع"]
        tree = find_slide_index_by_title(presentation1, "الزروع", elkhetam)
        show_array.append([tree, tree])
        
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

def bakerKiahk(copticdate, adam = False, Bishop = False, guestBishop = 0):
    from copticDate import CopticCalendar
    cd = CopticCalendar().coptic_to_gregorian(copticdate)
    prs = relative_path(r"رفع بخور عشية و باكر.pptx")  # Using the relative path
    excel = relative_path(r"بيانات القداسات.xlsx")
    excel2 = relative_path(r"Tables.xlsx")
    sheet ="رفع بخور"

    elzoksologyat(excel, 5, "باكر")

    if cd.weekday() == 6:
        katamars = relative_path(r"Data\القطمارس\الاحاد\القطمارس السنوي احاد (باكر).pptx")
        katamars_sheet = "قطمارس الاحاد باكر"
        km = copticdate[1]
        kd = (copticdate[2] - 1) // 7 + 1
    else: 
        katamars = relative_path(r"Data\القطمارس\الايام\القطمارس السنوي ايام (باكر).pptx")
        katamars_sheet = "القطمارس السنوي باكر"
        km, kd = find_Readings_Date(copticdate[1], copticdate[2])

    katamars_values = fetch_data_arrays(excel2, katamars_sheet, km, kd, [3, 4, 5])
    elmzmor = katamars_values[0]
    elengil = katamars_values[1]
    elengil2 = katamars_values[2]

    baker_values = find_slide_nums_arrays(excel, sheet, ["الانجيل", "الانجيل","ارباع الناقوس الادام", "ارباع الناقوس الادام", 
                                                          "أرباع الناقوس الواطس", "أرباع الناقوس الواطس", "تكملة ارباع الناقوس",
                                                          "اوشية الراقدين", "اوشية الراقدين", "اوشية المرضي", "اوشية المرضي",
                                                          "اوشية المسافرين", "اوشية المسافرين", "اوشية القرابين", "اوشية القرابين",
                                                          "فلنسبح مع الملائكة", "فلنسبح مع الملائكة",
                                                          "مرد انجيل كيهك 1", "مرد انجيل كيهك 1", "مرد انجيل كيهك 2", "مرد انجيل كيهك 2",
                                                          "تكملة على حسب المناسبة", "مرد الانجيل السنوي", "مرد الانجيل السنوي",
                                                          "تكملة مشتركة لكيهك", "تكملة مشتركة لكيهك"], 
                                                         [1, 2, 1, 2, 1, 2, 1, 1, 2, 1, 2, 1, 2, 1, 2, 1, 2, 1, 2, 1, 2, 1, 1, 2, 1, 2])

    #المزمور و الانجيل
    elengil3 = baker_values[1]
    elmzmor1 = baker_values[0] + 1

    #ارباع الناقوس
    arbaaAdam = baker_values[2]
    arbaaAdam2 = baker_values[3]
    arbaaWats = baker_values[4]
    arbaaWats2 = baker_values[5]
    arab3elna2os = baker_values[6]

    #الاواشي
    if cd.weekday() == 5:
        elawashy = baker_values[7]
        elawashy2 = baker_values[8]
        elawashy22 = 1
        elawashy222 = 1
    elif cd.weekday() == 6:
        elawashy = baker_values[9]
        elawashy2 = baker_values[10]
        elawashy22 = baker_values[13]
        elawashy222 = baker_values[14]
    else:
        elawashy = baker_values[9]
        elawashy2 = baker_values[10]
        elawashy22 = baker_values[11]
        elawashy222 = baker_values[12]

    angel = baker_values[15]
    angel2 = baker_values[16]

    #مرد الانجيل    
    if copticdate[2] <= 14:
        mrdelengil = baker_values[17]
        mrdelengil2 = baker_values[18]
    else:
        mrdelengil = baker_values[19]
        mrdelengil2 = baker_values[20]

    mrdelengilSanawy = baker_values[22]
    mrdelengilSanawy2 = baker_values[23]
    takmelaMrdelengil = baker_values[24]
    takmelaMrdelengil2 = baker_values[25]

    #الختام
    elkhetam = baker_values[21]

    show_array = [[elawashy, elawashy2], [elawashy22, elawashy222], [angel, angel2], [mrdelengil, mrdelengil2], [takmelaMrdelengil, takmelaMrdelengil2]]
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



def bakerSanawy(season, copticdate, adam = False, Bishop = False, guestBishop = 0):
    from copticDate import CopticCalendar
    cd = CopticCalendar().coptic_to_gregorian(copticdate)
    prs = relative_path(r"رفع بخور عشية و باكر.pptx")  # Using the relative path
    excel = relative_path(r"بيانات القداسات.xlsx")
    excel2 = relative_path(r"Tables.xlsx")
    sheet ="رفع بخور"
    replacefile(prs, relative_path(r"Data\CopyData\رفع بخور عشية و باكر.pptx"))
    
    elzoksologyat(excel, season,"باكر")

    if cd.weekday() == 6:
        sunday(prs)
        katamars = relative_path(r"Data\القطمارس\الاحاد\القطمارس السنوي احاد (باكر).pptx")
        katamars_sheet = "قطمارس الاحاد لباكر"
        km = copticdate[1]
        kd = (copticdate[2] - 1) // 7 + 1
    else: 
        katamars = relative_path(r"Data\القطمارس\الايام\القطمارس السنوي ايام (باكر).pptx")
        katamars_sheet = "القطمارس السنوي باكر"
        km, kd = find_Readings_Date(copticdate[1], copticdate[2])

    katamars_values = fetch_data_arrays(excel2, katamars_sheet, km, kd, [3, 4, 5])
    elmzmor = katamars_values[0]
    elengil = katamars_values[1]
    elengil2 = katamars_values[2]

def bakerElSomElkbyr(copticdate, Bishop = False, guestBishop = 0):
    from copticDate import CopticCalendar
    cd = CopticCalendar().coptic_to_gregorian(copticdate)
    prs = relative_path(r"رفع بخور عشية و باكر.pptx")  # Using the relative path
    excel = relative_path(r"بيانات القداسات.xlsx")
    excel2 = relative_path(r"Tables.xlsx")
    sheet ="رفع بخور"
    replacefile(prs, relative_path(r"Data\CopyData\رفع بخور عشية و باكر.pptx"))
