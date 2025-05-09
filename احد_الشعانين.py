import os
from commonFunctions import *
import win32com.client
from pptx import Presentation
from pptx.util import Inches
import time 

def open_presentation_relative_path(relative_path):
    script_directory = os.path.dirname(os.path.abspath(__file__))
    absolute_path = os.path.join(script_directory, relative_path)
    powerpoint = win32com.client.Dispatch("PowerPoint.Application")
    presentation = powerpoint.Presentations.Open(absolute_path)
    return presentation

def relative_path(relative_path):
    script_directory = os.path.dirname(os.path.abspath(__file__))
    absolute_path = os.path.join(script_directory, relative_path)
    return absolute_path

def find_slide_index_by_title(presentation, title, start_index=1):
    # Get the total number of slides in the presentation
    num_slides = presentation.Slides.Count
    
    # Iterate through each slide starting from the specified start_index
    if start_index < 1 or start_index > num_slides:
        raise ValueError("start_index is out of range")

    for i in range(start_index, num_slides + 1):
        slide = presentation.Slides(i)
        for shape in slide.Shapes:
            if shape.HasTextFrame:
                text_frame = shape.TextFrame
                if text_frame.HasText:
                    if title in text_frame.TextRange.Text:
                        return i
    return -1

def insert_image_to_slides_same_file(original_presentation_path, image_path):
    # Load the presentation
    prs = Presentation(original_presentation_path)

    # Loop through each slide in the presentation
    for slide in prs.slides:
        # Calculate the position and size of the image
        left = Inches(0)  # Horizontal position: 0.1 cm
        top = Inches(2)  # Vertical position: 5.67 cm
        width = Inches(10)  # Width: 25.14 cm
        height = Inches(5.5)  # Height: 13.53 cm

        # Add the image to the slide
        pic = slide.shapes.add_picture(image_path, left, top, width, height)

        # Send the image to the back
        slide.shapes._spTree.insert(2, pic._element)
        
    # Save the modified presentation, overwriting the original file
    prs.save(original_presentation_path)

def hide_slides(presentation, hide_array):
    for start_slide, end_slide in hide_array:
        for i in range(start_slide, end_slide + 1):
            presentation.Slides(i).SlideShowTransition.Hidden = True

def show_slides(presentation, show_array):
    for start_slide, end_slide in show_array:
        for i in range(start_slide, end_slide + 1):
            presentation.Slides(i).SlideShowTransition.Hidden = False

# def get_layouts_and_slide_indices(presentation_path, ms):
#     prs = Presentation(presentation_path)
#     master_slide = prs.slide_masters[ms]  # Index 1 corresponds to the second master slide
#     excel =  relative_path(r"بيانات القداسات.xlsx")
#     des_sheet ="سنوي"
#     kirolos = find_slide_num(excel, des_sheet, "صلاة الصلح بعنق العبودية  كيرلسي 1", 1)
#     kirolos2 = find_slide_num(excel, des_sheet, "نهاية الكيرلسي", 2)
#     el2sma = find_slide_num(excel, des_sheet, "قسمة - أيها السيد الرب إلهنا", 1)
#     el2sma2 = find_slide_num(excel, des_sheet, "قسمة للأب في الصوم الكبير - أيها السيد الرب الإله ضابط الكل", 2)
#     el2sma3 = find_slide_num(excel, des_sheet, "قسمة للأب علي ذبح اسحق - وحدث في الأيام التى أراد الله", 1)
#     el2sma4 = find_slide_num(excel, des_sheet, "تذكار البشارة والميلاد والقيامة - نسبح ونمجد إله الآلهة ورب الأرباب", 2)

#     layouts_and_indices = {}
    
#     # Find the layouts in the master slide
#     for layout in master_slide.slide_layouts:
#         layout_name = layout.name
#         layout_indices = []
        
#         # Find slides using the current layout
#         for i, slide in enumerate(prs.slides, start=1):
#             if slide.slide_layout == layout and (i < kirolos or i > kirolos2) and (i < el2sma or i > el2sma2) and (i<el2sma3 or i> el2sma4):
#                     layout_indices.append(i)
        
#         # Store layout name and indices
#         layouts_and_indices[layout_name] = layout_indices

#     return layouts_and_indices

# def update_layout_for_slides(ms_source, ms_target):
#     # Get the PowerPoint application object
#     ppt_app = win32com.client.Dispatch("PowerPoint.Application")
    
#     # Get the active presentation
#     presentation = ppt_app.ActivePresentation
    
#     # Get the target master
#     target_master = presentation.Designs(ms_target).SlideMaster
    
#     # Get layouts and slide indices from MS2
#     layouts_and_indices_ms2 = get_layouts_and_slide_indices(presentation.FullName, ms_source)
    
#     # Iterate through the layouts in the target master (MS4)
#     for target_layout in target_master.CustomLayouts:
#         # Get the layout name
#         layout_name = target_layout.Name
        
#         # Check if the layout exists in the MS2 layouts
#         if layout_name in layouts_and_indices_ms2:
#             # Get the indices of slides using this layout in MS2
#             slide_indices_to_change = layouts_and_indices_ms2[layout_name]
            
#             # Iterate through the slides and switch the layout
#             for slide_index in slide_indices_to_change:
#                 slide = presentation.Slides(slide_index)
#                 slide.CustomLayout = target_layout

def odasElsh3anyn(Bishop=False, guestBishop=0):
    prs1 = r"قداس.pptx"
    prs2 = r"Data\القداسات\قداس احد الشعانين.pptx"
    excel = relative_path(r"بيانات القداسات.xlsx")
    source_sheet = "أحد الشعانين"
    des_sheet ="سنوي"
    design = relative_path(r"Data\Designs\الشعانين.png")
    insert_image_to_slides_same_file(relative_path(prs1), design)

    #التوزيع
    int1 = find_slide_num(excel, des_sheet, "اك اسماروؤت", 1)
    int2 = find_slide_num(excel, des_sheet, "بي اويك", 2)
    int00 = find_slide_num(excel, des_sheet, "جي اف اسماروؤت", 1)
    el5tam = find_slide_num(excel, des_sheet, "ختام الصلوات", 1)
    sn = find_slide_num(excel, source_sheet, "التوزيع", 1)
    fs = find_slide_num(excel, des_sheet, "مزمور التوزيع", 1) +1
    ls = find_slide_num(excel, des_sheet, "مزمور التوزيع", 2) - 1
    elmady7 = find_slide_num(excel, source_sheet, "مدائح", 1)
    elmady72 = find_slide_num(excel, source_sheet, "مدائح", 2)

    #القسمة
    int3 = find_slide_num(excel, des_sheet, "قسمة للأب في حد الشعانين - أيها الرب ربنا لقد صار إسمك", 1)
    int4 = find_slide_num(excel, des_sheet, "قسمة للأب في حد الشعانين - أيها الرب ربنا لقد صار إسمك", 2)
    int5 = find_slide_num(excel, des_sheet, "قسمة - أيها السيد الرب إلهنا", 1)
    int6 = find_slide_num(excel, des_sheet, "قسمة - أيها السيد الرب إلهنا", 2)

    #الاواشي
    alahwyabasyly = find_slide_num(excel, des_sheet, "اوشية اهوية السماء", 1)
    alahwyabasyly2 = find_slide_num(excel, des_sheet, "اوشية اهوية السماء", 2)
    alahwya8r8ory = find_slide_num(excel, des_sheet, "اوشية اهوية السماء غ", 1)
    alahwya8r8ory2 = find_slide_num(excel, des_sheet, "اوشية اهوية السماء غ", 2)
    
    #الاسبسمس الواطس
    int7 = find_slide_num(excel, source_sheet, "الاسبسمس واطس", 1)
    int8 = find_slide_num(excel, source_sheet, "الاسبسمس واطس", 2)
    int9 = find_slide_num(excel, des_sheet, "ايها الرب ـ الاسبسمس الواطس", 1) + 1

    #اسبسمس ادام
    int10 = find_slide_num(excel, source_sheet, "الاسبسمس الادام", 1)
    int11 = find_slide_num(excel, source_sheet, "الاسبسمس الادام", 2)
    int12 = find_slide_num(excel, des_sheet, "الاسبسمس الادام", 1) +1

    #مرد الانجيل الرابع
    int13 = find_slide_num(excel, des_sheet, "مرد الانجيل الرابع الشعانين", 1)
    int14 = find_slide_num(excel, des_sheet, "مرد الانجيل الرابع الشعانين", 2)
    int15 = find_slide_num(excel, des_sheet, "مرد الانجيل", 1)
    int16 = find_slide_num(excel, des_sheet, "مرد الانجيل", 2) - 2

    #المزمور الثاني عربي و الانجيل الرابع
    int17 = find_slide_num(excel, source_sheet, "المزمور الثاني عربي", 1)
    int18 = find_slide_num(excel, source_sheet, "المزمور الثاني عربي", 2)
    int19 = find_slide_num(excel, source_sheet, "انجيل يوحنا", 1)
    int20 = find_slide_num(excel, source_sheet, "انجيل يوحنا", 2)
    int21 = find_slide_num(excel, des_sheet, "المزمور", 1)+2
    int22 = find_slide_num(excel, des_sheet, "المزمور", 2)

    #السنجاري 2
    int23 = find_slide_num(excel, source_sheet, "المزمور السنجاري 2", 1)
    int24 = find_slide_num(excel, source_sheet, "المزمور السنجاري 2", 2)
    int25 = find_slide_num(excel, des_sheet, "مارو اتشاسف", 1)

    #اوشية الانجيل و الاناجيل الثلاثه الاولى
    int26 =  find_slide_num(excel, des_sheet, "اجيوس", 2) + 1
    int27 = find_slide_num(excel, source_sheet, "اوشية الانجيل", 1)
    int28 = find_slide_num(excel, source_sheet, "مرد الانجيل الثالث", 2)

    #المحير
    elmo7yr = find_slide_num(excel, des_sheet, "محير احد الشعانين", 1)
    elmo7yr2 = find_slide_num(excel, des_sheet, "محير احد الشعانين", 2)
    elmo7yr3 = find_slide_num(excel, des_sheet, "تكملة مشتركة للمحير", 1)
    elmo7yr4 = find_slide_num(excel, des_sheet, "تكملة مشتركة للمحير", 2)

    # القراءات
    snksar = find_slide_num(excel, des_sheet, "السنكسار", 1)
    aksia = find_slide_num(excel, des_sheet, "اكسيا", 2)
    int29= find_slide_num(excel, des_sheet, "الابركسيس", 2)
    elebrksis = find_slide_num(excel, source_sheet, "الابركسيس عربي", 1)
    elebrksis2 = find_slide_num(excel, source_sheet, "الابركسيس عربي", 2)
    int30 = find_slide_num(excel, des_sheet, "مرد ابركسيس سنوي", 1)
    int301 = find_slide_num(excel, des_sheet, "مرد ابركسيس سنوي", 2)
    int31 = find_slide_num(excel, des_sheet, "مرد ابركسيس شعانيني", 1)
    int311 = find_slide_num(excel, des_sheet, "مرد ابركسيس شعانيني", 2)
    int32 = find_slide_num(excel, des_sheet, "الكاثوليكون", 2) 
    elkatholykon = find_slide_num(excel, source_sheet, "كاثوليكون عربي", 1)
    elkatholykon2 = find_slide_num(excel, source_sheet, "كاثوليكون عربي", 2)
    int33 = find_slide_num(excel, des_sheet, "البولس عربي", 1) + 1
    elbouls = find_slide_num(excel, source_sheet, "البولس عربي", 1)
    elbouls2 = find_slide_num(excel, source_sheet, "البولس عربي", 2)

    #طاي شوري و الليلويا فاي بي بي
    tayshoriy = find_slide_num(excel, des_sheet, "طاي شوري", 1)
    tayshoriy1 = find_slide_num(excel, des_sheet, "طاي شوري", 2)
    faybebi = find_slide_num(excel, des_sheet, "الليلويا فاي بيبي", 1)
    faybebi1 = find_slide_num(excel, des_sheet, "الليلويا فاي بيبي", 2)

    # ني سافيف
    nysaviv = find_slide_num(excel, des_sheet, "ني سافيف تيرو", 1)
    nysaviv2 = find_slide_num(excel, des_sheet, "ني سافيف تيرو", 2)

    #اللي القربان و افلوجيمينوس
    evlogimanos = find_slide_num(excel, des_sheet, "افلوجيمنوس", 1)
    evlogimanos2 = find_slide_num(excel, des_sheet, "افلوجيمنوس", 2)
    eksmarot2 = find_slide_num(excel, des_sheet, "لحن اك اسمارؤوت", 2)
    evlogimanos2 = find_slide_num(excel, des_sheet, "افلوجيمنوس", 2)
    el2orban = find_slide_num(excel, des_sheet, "اللي القربان", 1)
    el2orban2 = find_slide_num(excel, des_sheet, "اللي القربان", 2)
    abynav = find_slide_num(excel, des_sheet, "ابيناف شوبي", 1)
    abynav2 = find_slide_num(excel, des_sheet, "ابيناف شوبي", 2)

    if Bishop == True:
        prs3 = relative_path(r"Data\حضور الأسقف.pptx")
        sheet = "في حضور الأسقف"

        elshokr1 = find_slide_num(excel, sheet, "صلاة الشكر", 1)
        elshokrDes = find_slide_num(excel, des_sheet, "صلاة الشكر", 2) - 1

        bhmot8ar1 = find_slide_num(excel, sheet, "بهموت غار الصغيرة", 1)
        bhmot8arDes = find_slide_num(excel, des_sheet, "باهموت غار الصغيرة", 2) - 1

        embiniot = find_slide_num(excel, des_sheet, "امبين يوت اتطايوت", 1) 
        embiniot2 = find_slide_num(excel, des_sheet, "امبين يوت اتطايوت", 2)

        bishopHyten1 = find_slide_num(excel, sheet, "الهيتنيات", 1)
        bishopHytenDes = find_slide_num(excel, des_sheet, "تكملة الهيتنيات", 2) + 1

        marotshasf = find_slide_num(excel, des_sheet, "مارو اتشاسف", 1)
        marotshasf2 = find_slide_num(excel, des_sheet, "مارو اتشاسف", 2)

        elaba2basyly = find_slide_num(excel, des_sheet, "اوشية الاباء (ب)", 2) - 1
        elaba28or8ory = find_slide_num(excel, des_sheet, "اوشية الاباء غ", 1) - 1

        if guestBishop > 0:
            if guestBishop == 1:
                elshokr2 = find_slide_num(excel, sheet, "صلاة الشكر", 2) - 1
                bhmot8ar2 = find_slide_num(excel, sheet, "بهموت غار الصغيرة", 2) - 1
                bishopHyten2 = find_slide_num(excel, sheet, "الهيتنيات", 2) -3
                elaba2 = elshokr2
                elaba22 = elshokr2
            
            elif guestBishop == 2:
                elshokr2 = find_slide_num(excel, sheet, "صلاة الشكر", 2)
                bhmot8ar2 = find_slide_num(excel, sheet, "بهموت غار الصغيرة", 2)
                bishopHyten2 = find_slide_num(excel, sheet, "الهيتنيات", 2)
                elaba2 = elshokr2 - 1
                elaba22 = elshokr2
            start_positions = [int00 +2, int00+2, fs, elaba28or8ory, elaba2basyly, int9, int12, int22, int21, int25, int26, int29, int32, int33, bhmot8arDes, bishopHytenDes,embiniot2, elshokrDes]
            start_slides = [evlogimanos, elmady7, sn, elaba2, int7, int10, int19, int17, int23, int27, elebrksis, elkatholykon, elbouls, bhmot8ar1, bishopHyten1, bhmot8ar1, elshokr1]
            end_slides = [evlogimanos2, elmady72, sn, elaba22, int8, int11, int20, int18, int24, int28, elebrksis2, elkatholykon2, elbouls2, bhmot8ar2, bishopHyten2, bhmot8ar2, elshokr2]
        
        else:
            elshokr2 = find_slide_num(excel, sheet, "صلاة الشكر", 2) - 2
            start_positions = [int00+2, int00+2, fs, int9, int12, int22, int21, int25, int26, elmo7yr, int29, int32, int33, elshokrDes]
            start_slides = [evlogimanos, elmady7, sn, int7, int10, int19, int17, int23, int27, evlogimanos, elebrksis, elkatholykon, elbouls, elshokr1]
            end_slides = [evlogimanos2, elmady72, sn, int8, int11, int20, int18, int24, int28, evlogimanos2, elebrksis2, elkatholykon2, elbouls2, elshokr2]
        show_array = [[evlogimanos, eksmarot2], [el2orban, el2orban2], [faybebi, faybebi1], [tayshoriy, tayshoriy1], 
                      [marotshasf, marotshasf2], [int31, int311], [elmo7yr, elmo7yr2], [elmo7yr3, elmo7yr4], 
                      [int13, int14], [alahwyabasyly, alahwyabasyly2], [alahwya8r8ory, alahwya8r8ory2], 
                      [int3, int4], [int00, int00+2], [embiniot, embiniot2], [nysaviv, nysaviv2]]
    
    else:
        show_array = [[evlogimanos, eksmarot2], [el2orban, el2orban2], [faybebi, faybebi1], [tayshoriy, tayshoriy1], 
                      [int31, int311], [elmo7yr, elmo7yr2], [elmo7yr3, elmo7yr4], [int13, int14], 
                      [alahwyabasyly, alahwyabasyly2], [alahwya8r8ory, alahwya8r8ory2],[int3, int4], [int00, int00+2]]
        start_positions = [int00+2, int00+2, fs, int9, int12, int22, int21, int25, int26, elmo7yr, int29, int32, int33]
        start_slides = [evlogimanos, elmady7, sn, int7, int10, int19, int17, int23, int27, evlogimanos, elebrksis, elkatholykon, elbouls]
        end_slides = [evlogimanos2, elmady72, sn, int8, int11, int20, int18, int24, int28, evlogimanos2, elebrksis2, elkatholykon2, elbouls2]

    hide_array = [[abynav, abynav2], [int30, int301],[snksar, aksia], [int15, int16], [int5, int6], [int1, int2]]

    powerpoint = win32com.client.Dispatch("PowerPoint.Application")
    powerpoint.Visible = True  # Open PowerPoint application

    presentation1 = open_presentation_relative_path(prs1)
    khtamElseason = find_slide_index_by_title(presentation1, "إبن الله دخل أورشليم.", el5tam)
    show_array.append([khtamElseason, khtamElseason])
    presentation2 = open_presentation_relative_path(prs2)
    if(Bishop):
        show_slides(presentation2, [[find_slide_num(excel, source_sheet, "مارو اتشاسف", 1), 
                                     find_slide_num(excel, source_sheet, "مارو اتشاسف", 2)]])
        presentation3 = open_presentation_relative_path(prs3)

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
        if current_position == fs:
            source_slide = presentation2.Slides(current_start_slide)
            source_slide.Copy()
            slide_index1 = ls
            while slide_index1 >= current_position:
                new_slide = presentation1.Slides.Paste(slide_index1)

                pic_shape = new_slide.Shapes.AddPicture(design, LinkToFile=False, SaveWithDocument=True, Left=0, Top=144, Width=720, Height=396)
                pic_shape.ZOrder(1)

                if slide_index1 > ls - 14 and slide_index1 <= ls - 3:
                    new_slide.SlideShowTransition.Hidden = True
                
                slide_index1 -= 1
            current_start_slide += 1

        elif current_position == int9 or current_position == int12:
            for _ in range(6):
                if current_position == int9:
                    presentation1.Slides(int9).Delete()
                else:
                    presentation1.Slides(int12).Delete()
            source_slide = presentation2.Slides(current_start_slide)
            is_hidden = source_slide.SlideShowTransition.Hidden
            source_slide.Copy()
            new_slide = presentation1.Slides.Paste(current_position)
            
            pic_shape = new_slide.Shapes.AddPicture(design, LinkToFile=False, SaveWithDocument=True, Left=0, Top=144, Width=720, Height=396)
            pic_shape.ZOrder(1)

            if is_hidden:
                new_slide.SlideShowTransition.Hidden = True
            current_start_slide += 1
            current_position += 1

        elif current_start_slide >= evlogimanos and current_end_slide <= evlogimanos2:
            source_slide = presentation1.Slides(current_start_slide)
            is_hidden = source_slide.SlideShowTransition.Hidden
            source_slide.Copy()
            new_slide = presentation1.Slides.Paste(current_position)
            if is_hidden:
                new_slide.SlideShowTransition.Hidden = True
            current_start_slide += 1
            current_position += 1

        elif ((current_start_slide >= elmady7 and current_end_slide <= elmady72) or 
              (current_start_slide > int27 and current_end_slide <= int28) or
              (current_start_slide > int17 and current_end_slide <= int18) or
              (current_start_slide > elebrksis and current_end_slide <= elebrksis2) or 
              (current_start_slide > elkatholykon and current_end_slide <= elkatholykon2) or 
              (current_start_slide > elbouls and current_end_slide <= elbouls2)):
            
            source_slide = presentation2.Slides(current_start_slide)
    
            # Copy slide with source formatting
            source_slide.Copy()
            
            # Activate presentation1
            presentation1.Windows(1).Activate()
            
            # Paste with source formatting using ExecuteMso
            presentation1.Application.CommandBars.ExecuteMso("PasteSourceFormatting")

            current_start_slide += 1
            current_position += 1

        elif Bishop == True and (current_position == elaba28or8ory or current_position == elaba2basyly or 
                                 current_position == bhmot8arDes or current_position == bishopHytenDes or 
                                 current_position == embiniot2 or current_position == elshokrDes):
            source_slide = presentation3.Slides(current_end_slide)
            source_slide.Copy()
            new_slide = presentation1.Slides.Paste(current_position)
            new_slide.SlideShowTransition.Hidden = False
            pic_shape = new_slide.Shapes.AddPicture(design, LinkToFile=False, SaveWithDocument=True, Left=0, Top=144, Width=720, Height=396)
            pic_shape.ZOrder(1)
            current_end_slide -= 1
            if(current_start_slide > current_end_slide):
                current_position += 1

        else:
            source_slide = presentation2.Slides(current_start_slide)
            is_hidden = source_slide.SlideShowTransition.Hidden
            source_slide.Copy()
            new_slide = presentation1.Slides.Paste(current_position)
            default_master = presentation1.Designs(4).SlideMaster
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

# Close presentations
    presentation2.Close()
    if Bishop == True:
        presentation3.Close()

def bakerElsh3anyn ():
    prs1 = r"باكر.pptx"  
    prs2 = r"Data\القداسات\قداس احد الشعانين.pptx"
    excel = relative_path(r"بيانات القداسات.xlsx")
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

