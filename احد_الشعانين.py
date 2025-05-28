import os
from commonFunctions import relative_path, insert_image_to_slides_same_file, find_slide_num, find_slide_index_by_title, open_presentation_relative_path, hide_slides, show_slides
import win32com.client
from pptx import Presentation
from pptx.util import Inches

def bakerElsh3anyn ():
    prs1 = r"باكر.pptx"  
    prs2 = r"Data\القداسات\قداس احد الشعانين.pptx"
    excel = relative_path(r"Files Data.xlsx")
    source_sheet = "أحد الشعانين"
    des_sheet ="باكر"
    design = relative_path(r"Data\Designs\الشعانين.png")
    insert_image_to_slides_same_file(relative_path(prs1), design)

    #مرد الانجيل
    mrdelengil = find_slide_num(excel, des_sheet, "مرد انجيل باكر الشعانين", 1)
    mrdelengil2 = find_slide_num(excel, des_sheet, "مرد انجيل باكر الشعانين", 2)
    mrdelengil3 = find_slide_num(excel, des_sheet, "مرد الانجيل باكر", 1)
    mrdelengil4 = find_slide_num(excel, des_sheet, "مرد الانجيل باكر", 2)

    #الانجيل و المزمور
    elengil = find_slide_num(excel, source_sheet, "انجيل باكر", 1)
    elengil2 = find_slide_num(excel, source_sheet, "انجيل باكر", 2)
    elengil3 = find_slide_num(excel, des_sheet, "الانجيل", 1)+2
    elengil4 = find_slide_num(excel, des_sheet, "الانجيل", 2) - 1
    elmazmor = find_slide_num(excel, source_sheet, "مزمور باكر", 1)

    #مرد المزمور
    mrdmazor = find_slide_num(excel, des_sheet, "مرد المزمور الشعانين", 1)
    mrdmazor2 = find_slide_num(excel, des_sheet, "مرد المزمور الشعانين", 2)
    mrdmazor3 = find_slide_num(excel, des_sheet, "مرد المزمور", 1)

    #الذكصولوجيات
    elzoksologyat = find_slide_num(excel, source_sheet, "ذكصولوجية شعانيني 1", 1)
    elzoksologyat2 = find_slide_num(excel, source_sheet, "ذكصولوجية شعانيني 3", 2)
    elzoksologyat3 = find_slide_num(excel, des_sheet, "مقدمة الذوكصولجيات", 2) 

    #الاواشي
    tourists = find_slide_num(excel, des_sheet, "اوشية المسافرين", 1)
    tourists2 = find_slide_num(excel, des_sheet, "اوشية المسافرين", 2)
    bread = find_slide_num(excel, des_sheet, "اوشية القرابين", 1)
    bread2 = find_slide_num(excel, des_sheet, "اوشية القرابين", 2)

    #ارباع الناقوس
    arba3elnakos = find_slide_num(excel, des_sheet, "ارباع الناقوس الادام", 2) +1
    arba3elnakos2 = find_slide_num(excel, source_sheet, "ارباع الشعانين", 1)
    arba3elnakos3 = find_slide_num(excel, source_sheet, "ارباع الشعانين", 2)

    start_positions = [elengil4, elengil3, elzoksologyat3, arba3elnakos]
    start_slides = [elengil, elmazmor, elzoksologyat, arba3elnakos2]
    end_slides = [elengil2, elmazmor, elzoksologyat2, arba3elnakos3]
    hide_array = [[mrdelengil3, mrdelengil4], [mrdmazor3, mrdmazor3], [tourists, tourists2]]
    show_array = [[mrdelengil, mrdelengil2], [mrdmazor, mrdmazor2], [bread, bread2]]
    
    powerpoint = win32com.client.Dispatch("PowerPoint.Application")
    powerpoint.Visible = True  # Open PowerPoint application
    el5tam = find_slide_num(excel, des_sheet, "ختام الصلوات", 1)
    presentation1 = open_presentation_relative_path(prs1)
    khtamElseason = find_slide_index_by_title(presentation1, "إبن الله دخل أورشليم.", el5tam+3)
    show_array.append([khtamElseason, khtamElseason])
    presentation2 = open_presentation_relative_path(prs2)

    hide_slides(presentation1, hide_array)
    show_slides(presentation1, show_array)

    # Initialize variables for current position, slide, and end index
    current_position = start_positions[0]
    current_start_slide = int(start_slides[0])
    current_end_slide = int(end_slides[0])

    # Initialize index for start position, slide, and end slide
    position_index = 1
    slide_index = 1
    end_index = 1

    while current_start_slide <= current_end_slide and slide_index <= presentation2.Slides.Count:
        if ((current_start_slide > elengil and current_start_slide <= elengil2) or 
            (current_start_slide > elzoksologyat and current_end_slide <= elzoksologyat2) or 
            (current_start_slide > arba3elnakos2 and current_end_slide <= arba3elnakos3)):
            source_slide = presentation2.Slides(current_start_slide)
            source_slide.Copy()
            presentation1.Windows(1).Activate()
            presentation1.Application.CommandBars.ExecuteMso("PasteSourceFormatting")
            current_position += 1
            current_start_slide += 1
                
        else:
            source_slide = presentation2.Slides(current_start_slide)
            is_hidden = source_slide.SlideShowTransition.Hidden
            source_slide.Copy()
            new_slide = presentation1.Slides.Paste(current_position)
            default_master = presentation1.Designs(2).SlideMaster
            desired_layout_name = new_slide.CustomLayout.Name 
            desired_layout_index = None
            for i in range(1, default_master.CustomLayouts.Count + 1):
                if default_master.CustomLayouts(i).Name == desired_layout_name:
                    desired_layout_index = i
                    break
            
            if desired_layout_index is not None:
                new_slide.CustomLayout = default_master.CustomLayouts(desired_layout_index)
            if is_hidden:
                new_slide.SlideShowTransition.Hidden = True
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

