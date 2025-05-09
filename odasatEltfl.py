from commonFunctions import *
import win32com.client

def odasEltflSomElrosol (copticdate):
    from copticDate import CopticCalendar
    prs1 = r"قداس الطفل.pptx"  # Using the relative path
    prs3 = r"Data\القطمارس\الايام\القطمارس السنوي ايام (القداس).pptx"
    excel = relative_path(r"بيانات القداسات.xlsx")
    excel2 = relative_path(r"Tables.xlsx")
    des_sheet ="قداس الطفل"
    katamars_sheet = "القطمارس السنوي القداس"
    km, kd = find_Readings_Date(copticdate[1], copticdate[2])
    katamars_values = fetch_data_arrays(excel2, katamars_sheet, km, kd, [3, 4, 5, 6, 7, 8])
    elbouls1 = katamars_values[0]
    elkatholikon1 = katamars_values[1]
    elebrksis1 = katamars_values[2]
    elmzmor = katamars_values[3]
    elengil = katamars_values[4]
    elbouls2 = elkatholikon1 - 1
    elkatholikon2 = elebrksis1 - 1
    elebrksis2 = elmzmor - 1
    elengil2 = katamars_values[5]

    des_sheet_values = find_slide_nums_arrays(excel, des_sheet, ["جي اف اسماروؤت", "جي اف اسماروؤت", "اسومين",  "اسومين",
                                                                "قسمة - أيها السيد الرب إلهنا", "قسمة - أيها السيد الرب إلهنا",
                                                                "قسمة للإبن تقال في صوم الرسل - أنت هو كلمة الآب", 
                                                                "قسمة للإبن تقال في صوم الرسل - أنت هو كلمة الآب",
                                                                "ربع يقال في صوم الرسل", "ربع يقال في صوم الرسل",
                                                                "المزمور", "المزمور", "الابركسيس عربي", "مرد ابركسيس الرسل", "مرد ابركسيس الرسل",
                                                                "كاثوليكون عربي", "البولس عربي", "تي شوري", "تي شوري", 
                                                                "الليلويا جي افمفئي", "الليلويا جي افمفئي"], 
                                                                [1, 2, 1, 2, 1, 2, 1, 2, 1, 2, 1, 2, 2, 1, 2, 2, 2, 1, 2, 1, 2])

    #التوزيع
    jefsmarot = des_sheet_values[0]
    jefsmarot2 = des_sheet_values[1]
    asomyn = des_sheet_values[2]
    asomyn2 = des_sheet_values[3]

    #القسمة
    el2smaSanawy = des_sheet_values[4]
    el2smaSanawy2 = des_sheet_values[5] - 1
    el2smaRosol = des_sheet_values[6]
    el2smaRosol2 = des_sheet_values[7]

    #الاواشي
    if copticdate == None:
        copticdate = CopticCalendar().gregorian_to_coptic()
        season = CopticCalendar().get_coptic_date_range(copticdate)
    else:
        season = CopticCalendar().get_coptic_date_range(copticdate)
    if season == "Air":
        seasonbasyly = find_slide_num(excel, des_sheet, "اوشية اهوية السماء", 1)
        seasonbasyly2 = find_slide_num(excel, des_sheet, "اوشية اهوية السماء", 2)
    else:
        seasonbasyly = find_slide_num(excel, des_sheet, "اوشية المياة", 1)
        seasonbasyly2 = find_slide_num(excel, des_sheet, "اوشية المياة", 2)

    #مرد الانجيل
    mrdelengilRosol = des_sheet_values[8]
    mrdelengilRosol2 = des_sheet_values[9]

    #المزمور و الانجيل
    elengil3 = des_sheet_values[11]
    elmzmor1 = des_sheet_values[10] + 2

    #القرائات
    elebrksis3 =  des_sheet_values[12]
    mrdelebrksisRosol = des_sheet_values[13]
    mrdelebrksisRosol2 = des_sheet_values[14]
    elkatholikon3 = des_sheet_values[15]
    elbouls3 = des_sheet_values[16]

    #تي شوري و الليلويا جي اف ميفي
    tishory1 = des_sheet_values[17]
    tishory2 = des_sheet_values[18]
    allyloya1 = des_sheet_values[19]
    allyloya2 = des_sheet_values[20]

    start_positions = [elengil3, elmzmor1, elebrksis3, elkatholikon3, elbouls3]
    start_slides = [elengil, elmzmor, elebrksis1, elkatholikon1, elbouls1]
    end_slides = [elengil2, elmzmor, elebrksis2, elkatholikon2, elbouls2]
    show_array = [[jefsmarot, jefsmarot2], [asomyn, asomyn2], [el2smaRosol, el2smaRosol2], [seasonbasyly, seasonbasyly2], 
                  [mrdelengilRosol, mrdelengilRosol2], [mrdelebrksisRosol, mrdelebrksisRosol2], 
                  [tishory1, tishory2], [allyloya1, allyloya2]]

    hide_array = [[el2smaSanawy, el2smaSanawy2]]

    powerpoint = win32com.client.Dispatch("PowerPoint.Application")
    powerpoint.Visible = True  # Open PowerPoint application
    presentation1 = open_presentation_relative_path(prs1)
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

    while current_start_slide <= current_end_slide and slide_index <= presentation1.Slides.Count:
        if (current_position == elengil3 or current_position == elmzmor1 or current_position == elebrksis3 
            or current_position == elkatholikon3 or current_position == elbouls3):
            source_slide = presentation3.Slides(current_end_slide)
            source_slide.Copy()
            new_slide = presentation1.Slides.Paste(current_position).SlideShowTransition.Hidden = False
            current_end_slide -= 1
            if(current_start_slide > current_end_slide):
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

    sections = {presentation1.SectionProperties.Name(i): i for i in range(1, presentation1.SectionProperties.Count + 1)}
        
    # Get section indices
    move_index = sections["جي اف اسماروؤت"]
    target_index = sections["مزمور التوزيع"]
    
    # Adjust target_index if move_index is greater, because the section will be removed before insertion
    if move_index < target_index:
        target_index -= 1
    
    # Move section_to_move to the position before target_section
    presentation1.SectionProperties.Move(move_index, target_index + 1)

    # Close presentations
    presentation3.Close()

def odasEltfl3ydElrosol ():
    prs1 = relative_path(r"قداس الطفل.pptx")  # Using the relative path
    excel = relative_path(r"بيانات القداسات.xlsx")
    excel2 = relative_path(r"Tables.xlsx")
    des_sheet ="قداس الطفل"

    prs3 = relative_path(r"Data\القطمارس\الايام\القطمارس السنوي ايام (القداس).pptx")
    katamars_sheet = "القطمارس السنوي القداس"
    km, kd = 11, 5

    katamars_values = fetch_data_arrays(excel2, katamars_sheet, km, kd, [3, 4, 5, 6, 7, 8])
    elbouls1 = katamars_values[0]
    elkatholikon1 = katamars_values[1]
    elebrksis1 = katamars_values[2]
    elmzmor = katamars_values[3]
    elengil = katamars_values[4]
    elbouls2 = elkatholikon1 - 1
    elkatholikon2 = elebrksis1 - 1
    elebrksis2 = elmzmor - 1
    elengil2 = katamars_values[5]

    des_sheet_values = find_slide_nums_arrays(excel, des_sheet, ["قسمة - أيها السيد الرب إلهنا", "قسمة - أيها السيد الرب إلهنا",
                                                                "قسمة للإبن تقال في صوم الرسل - أنت هو كلمة الآب", 
                                                                "قسمة للإبن تقال في صوم الرسل - أنت هو كلمة الآب",
                                                                "المزمور", "المزمور", "الابركسيس عربي","كاثوليكون عربي", "البولس عربي",
                                                                "تي شوري", "تي شوري", "الليلويا جي افمفئي", "الليلويا جي افمفئي"],
                                                                [1, 2, 1, 2, 1, 2, 2, 2, 2, 1, 2, 1, 2])

    #القسمة
    el2smaSanawy = des_sheet_values[0]
    el2smaSanawy2 = des_sheet_values[1]
    el2smaRosol = des_sheet_values[2]
    el2smaRosol2 = des_sheet_values[3]

    #الاواشي
    seasonbasyly = find_slide_num(excel, des_sheet, "اوشية المياة", 1)
    seasonbasyly2 = find_slide_num(excel, des_sheet, "اوشية المياة", 2)

    #المزمور و الانجيل
    elengil3 = des_sheet_values[4]
    elmzmor1 = des_sheet_values[5] + 2

    #القرائات
    elebrksis3 =  des_sheet_values[6]
    elkatholikon3 = des_sheet_values[7]
    elbouls3 = des_sheet_values[8]

    #تي شوري و الليلويا جي اف ميفي
    shory1 = des_sheet_values[9]
    shory2 = des_sheet_values[10]
    allyloya1 = des_sheet_values[11]
    allyloya2 = des_sheet_values[12]

    start_positions = [elengil3, elmzmor1, elebrksis3, elkatholikon3, elbouls3]
    start_slides = [elengil, elmzmor, elebrksis1, elkatholikon1, elbouls1]
    end_slides = [elengil2, elmzmor, elebrksis2, elkatholikon2, elbouls2]
    show_array = [[el2smaRosol, el2smaRosol2], [seasonbasyly, seasonbasyly2], [shory1, shory2], [allyloya1, allyloya2]]
    hide_array = [[el2smaSanawy, el2smaSanawy2]]

    powerpoint = win32com.client.Dispatch("PowerPoint.Application")
    powerpoint.Visible = True  # Open PowerPoint application
    presentation1 = open_presentation_relative_path(prs1)
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

    while current_start_slide <= current_end_slide and slide_index <= presentation1.Slides.Count:
        if (current_position == elengil3 or current_position == elmzmor1 or current_position == elebrksis3 
            or current_position == elkatholikon3 or current_position == elbouls3):
            source_slide = presentation3.Slides(current_end_slide)
            source_slide.Copy()
            new_slide = presentation1.Slides.Paste(current_position).SlideShowTransition.Hidden = False
            current_end_slide -= 1
            if(current_start_slide > current_end_slide):
                current_position += 1

        else:
            source_slide = presentation3.Slides(current_start_slide)
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

    presentation3.Close()

def odasEltflElnayrooz(copticdate):
    from copticDate import CopticCalendar
    cd = CopticCalendar().coptic_to_gregorian(copticdate)
    prs1 = relative_path(r"قداس الطفل.pptx")  # Using the relative path
    prs2 = relative_path(r"Data\القداسات\قداس عيد النيروز و الصليب.pptx")
    excel = relative_path(r"بيانات القداسات.xlsx")
    excel2 = relative_path(r"Tables.xlsx")
    des_sheet ="قداس الطفل"
 
    prs3 = relative_path(r"Data\القطمارس\الايام\القطمارس السنوي ايام (القداس).pptx")
    katamars_sheet = "القطمارس السنوي القداس"
    km, kd = find_Readings_Date(copticdate[1], copticdate[2])

    katamars_values = fetch_data_arrays(excel2, katamars_sheet, km, kd, [3, 4, 5, 6, 7, 8])
    elbouls1 = katamars_values[0]
    elkatholikon1 = katamars_values[1]
    elebrksis1 = katamars_values[2]
    elmzmor = katamars_values[3]
    elengil = katamars_values[4]
    elbouls2 = elkatholikon1 - 1
    elkatholikon2 = elebrksis1 - 1
    elebrksis2 = elmzmor - 1
    elengil2 = katamars_values[5]

    des_sheet_values = find_slide_nums_arrays(excel, des_sheet, ["مزمور التوزيع", "مزمور التوزيع",
                                                                "قسمة - أيها السيد الرب إلهنا", "قسمة - أيها السيد الرب إلهنا",
                                                                "قسمة النيروز - لنسبح الله تسبيحا جديدا", "قسمة النيروز - لنسبح الله تسبيحا جديدا",
                                                                "اوشية المياة", "اوشية المياة", "مرد انجيل النيروز", "مرد انجيل النيروز",
                                                                "فاي اريه بي اوو", "فاي اريه بي اوو","مرد الانجيل", "مرد الانجيل",
                                                                "مرد مزمور النيروز", "مرد مزمور النيروز", "مرد المزمور",
                                                                "مرد ابركسيس النيروز", "مرد ابركسيس النيروز","طاي شوري", "طاي شوري",
                                                                "الليلويا فاي بيبي", "الليلويا فاي بيبي", "المزمور", "المزمور",
                                                                "الابركسيس عربي", "كاثوليكون عربي", "البولس عربي"
                                                                ], 
                                                                [1, 2, 1, 2, 1, 2, 1, 2, 1, 2, 1, 2, 1, 2, 1, 2, 1, 
                                                                 1, 2, 1, 2, 1, 2, 1, 2, 2, 2, 2])
    
    
    #التوزيع
    sn = 1
    fs = des_sheet_values[0] + 1 
    ls = des_sheet_values[1] - 1

    #القسمة
    el2smaSanawy = des_sheet_values[2]
    el2smaSanawy2 = des_sheet_values[3] - 1
    el2smaNayrooz = des_sheet_values[4]
    el2smaNayrooz2 = des_sheet_values[5]

    #اوشية المياة
    seasonbasyly = des_sheet_values[6]
    seasonbasyly2 = des_sheet_values[7]
    season8r8ory = des_sheet_values[8]
    season8r8ory2 = des_sheet_values[9]

    #مرد الإنجيل
    mrdengilNayrooz = des_sheet_values[8]
    mrdengilNayrooz2 = des_sheet_values[9]
    fayereby = des_sheet_values[10]
    fayereby2 = des_sheet_values[11]
    mrdengil = des_sheet_values[12]
    mrdengil2 = des_sheet_values[13]

    #المزمور و الانجيل
    elengil3 = des_sheet_values[24]
    elmzmor1 = des_sheet_values[23] + 2

    #مرد المزمور
    mrdElmzmorNayrooz = des_sheet_values[14]
    mrdElmzmorNayrooz2 = des_sheet_values[15]
    mrdElmzmor = des_sheet_values[16]

    #مرد الابركسيس
    mrdelebrksis = des_sheet_values[17]
    mrdelebrksis2 = des_sheet_values[18]

    #القرائات
    elebrksis3 =  des_sheet_values[25]
    elkatholikon3 = des_sheet_values[26]
    elbouls3 = des_sheet_values[27]

    #الليلويا فاي بيبي و طاي شوري
    shory1 = des_sheet_values[19]
    shory2 = des_sheet_values[20]
    allyloya1 = des_sheet_values[21]
    allyloya2 = des_sheet_values[22]

    start_positions = [fs, elengil3, elmzmor1, elebrksis3, elkatholikon3, elbouls3]
    start_slides = [sn, elengil, elmzmor, elebrksis1, elkatholikon1, elbouls1]
    end_slides = [sn, elengil2, elmzmor, elebrksis2, elkatholikon2, elbouls2]
    show_array = [[el2smaNayrooz, el2smaNayrooz2], [season8r8ory, season8r8ory2], [seasonbasyly, seasonbasyly2], 
                    [mrdengilNayrooz, mrdengilNayrooz2], [fayereby, fayereby], [mrdElmzmorNayrooz, mrdElmzmorNayrooz2], 
                    [mrdelebrksis, mrdelebrksis2], [shory1, shory2], [allyloya1, allyloya2]]

    hide_array = [[el2smaSanawy, el2smaSanawy2], [mrdengil, mrdengil2], [mrdElmzmor, mrdElmzmor]]

    powerpoint = win32com.client.Dispatch("PowerPoint.Application")
    powerpoint.Visible = True  # Open PowerPoint application
    presentation1 = open_presentation_relative_path(prs1)
    presentation2 = open_presentation_relative_path(prs2)
    presentation3 = open_presentation_relative_path(prs3)

    hide_slides(presentation1, hide_array)
    show_slides(presentation1, show_array)

    # Initialize variables for current position, slide, and end index
    current_position = start_positions[0]
    current_start_slide = start_slides[0]
    current_end_slide = end_slides[0]

    # Initialize index for start position, slide, and end slide
    position_index = 1
    slide_index = 1
    end_index = 1

    while current_start_slide <= current_end_slide and slide_index <= presentation1.Slides.Count:
        if current_position == fs:
            source_slide = presentation2.Slides(current_start_slide)
            source_slide.Copy()
            slide_index1 = ls
            while slide_index1 >= current_position:
                new_slide = presentation1.Slides.Paste(slide_index1)

                if slide_index1 > ls - 14 and slide_index1 <= ls - 3:
                    new_slide.SlideShowTransition.Hidden = True
                
                slide_index1 -= 1
            current_start_slide += 1

        elif (current_position == elengil3 or current_position == elmzmor1 or current_position == elebrksis3 
            or current_position == elkatholikon3 or current_position == elbouls3):
            source_slide = presentation3.Slides(current_end_slide)
            source_slide.Copy()
            new_slide = presentation1.Slides.Paste(current_position).SlideShowTransition.Hidden = False
            current_end_slide -= 1
            if(current_start_slide > current_end_slide):
                current_position += 1

        else:
            source_slide = presentation2.Slides(current_start_slide)
            is_hidden = source_slide.SlideShowTransition.Hidden
            source_slide.Copy()
            new_slide = presentation1.Slides.Paste(current_position)
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

    move_section_names = [
        "قسمة النيروز - لنسبح الله تسبيحا جديدا",
    ]

    target_section_names = [
        "القسمة",
    ]

    # Call the function once for all moves
    move_sections(presentation1, move_section_names, target_section_names)

    presentation2.Close()
    presentation3.Close()

def odasEltflKiahk(copticdate):
    prs1 = relative_path(r"قداس الطفل.pptx")  # Using the relative path
    excel = relative_path(r"بيانات القداسات.xlsx")
    excel2 = relative_path(r"Tables.xlsx")
    des_sheet ="قداس الطفل"

    prs3 = relative_path(r"Data\القطمارس\الايام\القطمارس السنوي ايام (القداس).pptx")
    katamars_sheet = "القطمارس السنوي القداس"
    km, kd = find_Readings_Date(copticdate[1], copticdate[2])

    katamars_values = fetch_data_arrays(excel2, katamars_sheet, km, kd, [3, 4, 5, 6, 7, 8])
    elbouls1 = katamars_values[0]
    elkatholikon1 = katamars_values[1]
    elebrksis1 = katamars_values[2]
    elmzmor = katamars_values[3]
    elengil = katamars_values[4]
    elbouls2 = elkatholikon1 - 1
    elkatholikon2 = elebrksis1 - 1
    elebrksis2 = elmzmor - 1
    elengil2 = katamars_values[5]

    des_sheet_values = find_slide_nums_arrays(excel, des_sheet, ["المزمور", "المزمور", "الابركسيس عربي","كاثوليكون عربي", "البولس عربي", 
                                                                "تي شوري", "تي شوري", "الليلويا جي افمفئي", "الليلويا جي افمفئي", "الهيتنيات",
                                                                "قسمة - أيها السيد الرب إلهنا", "قسمة - أيها السيد الرب إلهنا",
                                                                "قسمة للأب في صوم و عيد الميلاد - أيها السيد الرب إلهنا",
                                                                "قسمة للأب في صوم و عيد الميلاد - أيها السيد الرب إلهنا", 
                                                                "مرد انجيل كيهك 1", "مرد انجيل كيهك 1", "مرد انجيل كيهك 2", "مرد انجيل كيهك 2",
                                                                "تكملة مشتركة كيهك", "تكملة مشتركة كيهك", "مرد الانجيل", "مرد الانجيل"], 
                                                                [1, 2, 2, 2, 2, 1, 2, 1, 2, 1, 1, 2, 1, 2, 1, 2, 1, 2, 1, 2, 1, 2])
        
    #القسمة
    el2smaSanawy = des_sheet_values[10]
    el2smaSanawy2 = des_sheet_values[11]-1
    el2smaElmilad = des_sheet_values[12]
    el2smaElmilad2 = des_sheet_values[13]

    #الاواشي
    seasonbasyly = find_slide_num(excel, des_sheet, "اوشية الزروع  العشب", 1)
    seasonbasyly2 = find_slide_num(excel, des_sheet, "اوشية الزروع  العشب", 2)

    #مرد الانجيل
    if copticdate[2]<=14:
        mrdengilkiahk = des_sheet_values[14]
        mrdengilkiahk2 = des_sheet_values[15]
    else:
        mrdengilkiahk = des_sheet_values[16]
        mrdengilkiahk2 = des_sheet_values[17]

    takmela = des_sheet_values[18]
    takmela2 = des_sheet_values[19]
    mrdengilsanawy = des_sheet_values[20]
    mrdengilsanawy2 = des_sheet_values[21]

    #المزمور و الانجيل
    elengil3 = des_sheet_values[1]
    elmzmor1 = des_sheet_values[0] + 2

    #القرائات
    elebrksis3 =  des_sheet_values[2]
    elkatholikon3 = des_sheet_values[3]
    elbouls3 = des_sheet_values[4]

    #تي شوري و الليلويا جي اف ميفي
    shory1 = des_sheet_values[5]
    shory2 = des_sheet_values[6]
    allyloya1 = des_sheet_values[7]
    allyloya2 = des_sheet_values[8]

    start_positions = [elengil3, elmzmor1, elebrksis3, elkatholikon3, elbouls3]
    start_slides = [elengil, elmzmor, elebrksis1, elkatholikon1, elbouls1]
    end_slides = [elengil2, elmzmor, elebrksis2, elkatholikon2, elbouls2]
    show_array = [[el2smaElmilad, el2smaElmilad2], [seasonbasyly, seasonbasyly2], [mrdengilkiahk, mrdengilkiahk2],
                  [takmela, takmela2], [shory1, shory2], [allyloya1, allyloya2]]

    powerpoint = win32com.client.Dispatch("PowerPoint.Application")
    powerpoint.Visible = True  # Open PowerPoint application
    presentation1 = open_presentation_relative_path(prs1)
    presentation3 = open_presentation_relative_path(prs3)

    hide_slides(presentation1, [[el2smaSanawy, el2smaSanawy2], [mrdengilsanawy, mrdengilsanawy2]])

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
        if (current_position == elengil3 or current_position == elmzmor1 or current_position == elebrksis3 
            or current_position == elkatholikon3 or current_position == elbouls3):
            source_slide = presentation3.Slides(current_end_slide)
            source_slide.Copy()
            new_slide = presentation1.Slides.Paste(current_position).SlideShowTransition.Hidden = False
            current_end_slide -= 1
            if(current_start_slide > current_end_slide):
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

    sections = {presentation1.SectionProperties.Name(i): i for i in range(1, presentation1.SectionProperties.Count + 1)}
    move_index = sections["جي اف اسماروؤت"]
    target_index = sections["مزمور التوزيع"]
    if move_index < target_index:
        target_index -= 1
    presentation1.SectionProperties.Move(move_index, target_index + 1)

    move_index = sections["قسمة للأب في صوم و عيد الميلاد - أيها السيد الرب إلهنا"]
    target_index = sections["القسمة"]
    if move_index < target_index:
        target_index -= 1
    presentation1.SectionProperties.Move(move_index, target_index + 1)
    presentation3.Close()

def odasEltflElmilad():
    return

def odasSanawy(copticdate, season):
    from copticDate import CopticCalendar
    cd = CopticCalendar().coptic_to_gregorian(copticdate)
    prs1 = relative_path(r"قداس الطفل.pptx")  # Using the relative path
    excel = relative_path(r"بيانات القداسات.xlsx")
    excel2 = relative_path(r"Tables.xlsx")
    des_sheet ="قداس الطفل"
    replacefile(prs1, relative_path(r"Data\CopyData\قداس الطفل.pptx"))

    prs2 = relative_path(r"Data\القطمارس\الايام\القطمارس السنوي ايام (القداس).pptx")
    katamars_sheet = "القطمارس السنوي القداس"
    km, kd = find_Readings_Date(copticdate[1], copticdate[2])
    katamars_offsets = [3, 4, 5, 6, 7, 8]

    katamars_values = fetch_data_arrays(excel2, katamars_sheet, km, kd, katamars_offsets)
    elbouls1 = katamars_values[0]
    elkatholikon1 = katamars_values[1]
    elebrksis1 = katamars_values[2]
    elmzmor = katamars_values[3]
    elengil = katamars_values[4]
    elbouls2 = elkatholikon1 - 1
    elkatholikon2 = elebrksis1 - 1
    elebrksis2 = elmzmor - 1
    elengil2 = katamars_values[5]

    sanawy_show_full_sections = []
    sanawy_hide_full_sections = []
    sanawy_show_values = []
    sanawy_hide_values = []

    sanawy_values = ["تكملة على حسب المناسبة", "الانجيل", "المزمور", "الابركسيس", "الكاثوليكون", "البولس عربي"]
    sanawy_values = find_slide_nums_arrays_v2(excel, des_sheet, 
                       ['{A18EDC94-F257-4FAC-99C7-0A8EA70F0FAF}', '{C7D4A109-F792-4661-BAD0-075FD1A1909F}', '{B74DBB8C-2B2D-46E4-9508-DA46008D19A4}', '{E234C6C7-3837-4CE4-A541-CDC9627AAAC2}', '{6D4B3F52-63BF-435F-BF0C-C9D41120C2A3}', '{D88055F5-EAA0-4C8E-8249-C364A572BF7B}'], 
                       2, [1, 2, 2, 2, 2, 2])
    
    #الختام
    elkhetam = sanawy_values[0]
    #الاواشي
    AwashySeason = CopticCalendar().get_coptic_date_range(copticdate)
    match AwashySeason:
        case "Air": sanawy_show_full_sections.extend(['{BC7E3DCD-6AA8-44CC-B8AF-BC3E2BC71B5A}', '{A20DA654-32F7-4B4C-96CB-C76232EB96E8}']),
        case "Tree": sanawy_show_full_sections.extend(['{F94B3D1F-649D-4839-BD2E-19439E173129}', '{5DD6BABA-9FE4-4D33-9F90-0C865CB95EE4}'])
        case "Water": sanawy_show_full_sections.extend(['{C7FC170A-D45F-4D4E-BD01-F17CADBFB65C}', '{3D4C118C-E6FF-4DF8-8E8F-B0CDF0FDBA54}'])
    #القرائات
    elengil3 = sanawy_values[1]
    elmazmor3 = sanawy_values[2]
    elebrksis3 = sanawy_values[3]
    elkatholikon3 = sanawy_values[4]
    elbouls3 = sanawy_values[5]

    #تي شوري و الليلويا جي اف ميفي
    if cd.weekday() == 2 or cd.weekday() == 4 or season == 6 or season == 30:
        sanawy_show_full_sections.extend(['{F8B8BD1A-9861-4FDC-A89B-06B55C0795E8}', '{43B74D4E-929B-4654-8DB7-4675EFF27370}'])
    else:
        sanawy_show_full_sections.extend(['{072F3D96-A6C8-405F-9A23-7CCA1B2F13FF}', '{20F525FD-C708-4DDD-8E40-FE502EFEBDDE}'])

    start_positions = [elengil3, elmazmor3, elebrksis3, elkatholikon3, elbouls3]
    start_slides = [elengil, elmzmor, elebrksis1, elkatholikon1, elbouls1]
    end_slides = [elengil2, elmzmor, elebrksis2, elkatholikon2, elbouls2]

    El3adraAndAngels = False 
    
    if season == 30 or season == 31 or copticdate[2]==21 or (copticdate[1]==9 and copticdate[2]==1) :
        El3adraAndAngels == True
        # el3adra_show_values = ["مرد انجيل كيهك 2 و صوم العذراء", 
        #                        "قسمة أعياد الملائكة والسيدة العذراء وسنوى (هوذا كائن معنا على هذه)",
        #                        "اطاي بارثينوس"]
        # el3adra_hide_values = ["مرد الانجيل", "قسمة القداس الباسيلي (أيها السيد الرب إلهنا)"]
        el3adra_show_values = ['{71232865-63AA-40C5-8F02-BABBFE7297D3}', '{18835C90-087E-4BAC-9D66-708BC1E04983}', '{D5E69BAC-0157-4B69-9255-B6775E2EE11D}' ]
        el3adra_hide_values = ['{06D592C8-05BF-4B7C-86F7-FDAB3FAB5FB1}', '{681FF6A7-4230-4171-8F41-83FD64E8C960}']
        sanawy_hide_full_sections.extend(el3adra_hide_values)
        sanawy_show_full_sections.extend(el3adra_show_values)

    elif copticdate[2] == 12:
        El3adraAndAngels == True
        # elmalakmikhael_values = ["مرد ابركسيس الملاك ميخائيل", "تكملة للملاك ميخائيل 2",
        #                               "ربع للملاك ميخائيل" , "هيتينية الملاك ميخائيل"],
        
        elmalakmikhael_show_values = ['{E95B1DDC-4235-4C02-91A4-DCB7A2808C33}', '{9EF543FB-A75B-4171-B358-2EB549C98411}', '{71232865-63AA-40C5-8F02-BABBFE7297D3}']
        elmalakmikhael_hide_values = ['{681FF6A7-4230-4171-8F41-83FD64E8C960}']
        sanawy_show_full_sections.extend(elmalakmikhael_show_values)
        sanawy_hide_full_sections.extend(elmalakmikhael_hide_values)
        elmalakmikhael_values = find_slide_nums_arrays_v2(excel, des_sheet, 
                                ["{4329E910-BD2C-4FBB-8FF3-A59F06EE9D45}", "{5BB65881-3E8A-4130-839D-6EB6F9D5FAFA}"],
                                2, [1, 2])
        mrdebrksis = elmalakmikhael_values[0]
        mrdebrksis3 = elmalakmikhael_values[1]
        if copticdate[1] == 3:
            mrdebrksis2 = find_slide_num_v2(excel, des_sheet, "{14A3F09D-ACCA-461F-AD67-08484F44D518}", 2, 1)
        elif copticdate[1] == 10:
            mrdebrksis2 = find_slide_num_v2(excel, des_sheet, "{02EBDDE5-1CBF-452A-A12A-A3F76FE68DDC}", 2, 1)
        else:
            mrdebrksis2 = find_slide_num_v2(excel, des_sheet, "{56E5BC0D-5FFC-4411-AC9C-78085E58A9E3}", 2, 1)
        sanawy_show_values.extend([[mrdebrksis, mrdebrksis], [mrdebrksis2, mrdebrksis2], [mrdebrksis3, mrdebrksis3]])
    
    elif season == 13:
        sanawy_show_full_sections.extend(['{EB470874-9124-4BDB-8075-189A7B264402}', '{4A03C859-F4BF-49C0-8ACD-88213CF6D13D}'])
        sanawy_hide_full_sections.extend(['{06D592C8-05BF-4B7C-86F7-FDAB3FAB5FB1}', '{ECE652ED-1345-4C6D-B92D-5996CFA27AEE}'])
    
    if cd.weekday() == 6:
        show_hide_replaceText(prs1, excel, des_sheet, sanawy_show_full_sections, sanawy_hide_full_sections, ["لأنك قمت","aktwnk", "آك طونك"])
    else:
        show_hide(prs1, excel, des_sheet, sanawy_show_full_sections, sanawy_hide_full_sections)

    powerpoint = win32com.client.Dispatch("PowerPoint.Application")
    powerpoint.Visible = True  # Open PowerPoint application
    presentation1 = open_presentation_relative_path(prs1)
    presentation2 = open_presentation_relative_path(prs2)

    khetamValue = 0
    if season == 6:
        khetamValue = find_slide_index_by_title(presentation1, "صوم الميلاد", elkhetam)
    elif season == "Air" :
        khetamValue = find_slide_index_by_title(presentation1, "الاهوية", elkhetam)
    elif season == "Water" :
        khetamValue = find_slide_index_by_title(presentation1, "المياة", elkhetam)
    else:
        khetamValue = find_slide_index_by_title(presentation1, "الزروع", elkhetam)
    
    sanawy_show_values.append([khetamValue, khetamValue])
    show_slides(presentation1, sanawy_show_values)

    # Initialize variables for current position, slide, and end index
    current_position = start_positions[0]
    current_start_slide = int(start_slides[0])
    current_end_slide = int(end_slides[0])

    # Initialize index for start position, slide, and end slide
    position_index = 1
    slide_index = 1
    end_index = 1

    while current_start_slide <= current_end_slide and slide_index <= presentation1.Slides.Count:
        if (current_position == elengil3 or current_position == elmazmor3 or current_position == elebrksis3 
            or current_position == elkatholikon3 or current_position == elbouls3):
            source_slide = presentation2.Slides(current_end_slide)
            source_slide.Copy()
            presentation1.Slides.Paste(current_position).SlideShowTransition.Hidden = False
            current_end_slide -= 1
            if current_start_slide > current_end_slide:
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

    if El3adraAndAngels :
        move_sections_v2(presentation1, ['{71232865-63AA-40C5-8F02-BABBFE7297D3}'], ['{01F61BFD-210C-4F99-A7C7-7308CFAA93F4}'])
    elif season == 6:
        move_sections_v2(presentation1, ['{C2F28915-B86E-4596-8EB2-7455EF4E91BD}'], ['{01F61BFD-210C-4F99-A7C7-7308CFAA93F4}'])
    
    presentation2.Close()
    presentation1.SlideShowSettings.Run()

def odasElSomElkbyr(copticdate, season):
    from copticDate import CopticCalendar
    cd = CopticCalendar().coptic_to_gregorian(copticdate)
    prs1 = relative_path(r"قداس الطفل.pptx")  # Using the relative path
    excel = relative_path(r"بيانات القداسات.xlsx")
    excel2 = relative_path(r"Tables.xlsx")
    des_sheet ="قداس الطفل"
    replacefile(prs1, relative_path(r"Data\CopyData\قداس الطفل.pptx"))

    prs2 = relative_path(r"Data\القطمارس\الصوم الكبير و صوم نينوى\قطمارس الصوم الكبير.pptx")
    katamars_sheet = "قطمارس الصوم الكبير"
    km = copticdate[1]
    kd = copticdate[2]
    katamars_offsets = [10, 11, 12, 13, 14, 15]

    katamars_values = fetch_data_arrays(excel2, katamars_sheet, km, kd, katamars_offsets)
    elbouls1 = katamars_values[0]
    elkatholikon1 = katamars_values[1]
    elebrksis1 = katamars_values[2]
    elmzmor = katamars_values[3]
    elengil = katamars_values[4]
    elbouls2 = elkatholikon1 - 1
    elkatholikon2 = elebrksis1 - 1
    elebrksis2 = elmzmor - 1
    elengil2 = katamars_values[5]

    # som_show_full_sections = ["مرد انجيل احاد الصوم الكبير و تكملة الايام", "مدائح الصوم الكبير", 
    #                           "قسمة للآب في الصوم الكبير المقدس (أيها السيد الرب الإله ضابط الكل)"]
    # som_hide_full_sections = ["مرد الانجيل", "ربع للعذراء", "اك اسماروؤت", "بي اويك",
    #                           "قسمة القداس الباسيلي (أيها السيد الرب إلهنا)"]

    som_show_full_sections = ['{6C210ECD-CC91-4984-B251-46939B2A0039}', '{41DD9B19-60F1-44CE-BF9E-34D6DA71F069}', '{A89B587D-E8EA-452C-8CAB-B9EE2C91743D}']
    som_hide_full_sections = ['{B9A30F5E-0C89-471B-A99A-23DBE7F58504}', '{507EFD97-98F8-4376-848B-20D72E16D2C1}', '{681FF6A7-4230-4171-8F41-83FD64E8C960}', '{ECE652ED-1345-4C6D-B92D-5996CFA27AEE}', '{06D592C8-05BF-4B7C-86F7-FDAB3FAB5FB1}']
    som_hide_full_sections_ranges = []
    som_show_values = [[]]
    som_hide_values = [[]]
    
    # som_values = ["تكملة على حسب المناسبة", "مزمور التوزيع", "مزمور التوزيع", 
    #               "مرد توزيع الصوم الكبير", "قسمة للإبن في أيام صوم الأربعين المقدس (أنت هو الله الرحوم مخلص)", 
    #               "الانجيل", "المزمور", "الابركسيس", "الكاثوليكون", "البولس عربي"]
    
    som_values = find_slide_nums_arrays_v2(excel, des_sheet, 
                 ['{A18EDC94-F257-4FAC-99C7-0A8EA70F0FAF}', '{C29E5A83-A98B-4077-8194-99A6D803EF53}', '{C29E5A83-A98B-4077-8194-99A6D803EF53}', '{F9ED982E-F6FB-4E2B-8955-C5E80C70C2D6}', '{F076B353-F7AB-4001-959A-5D482DE256DB}', '{C7D4A109-F792-4661-BAD0-075FD1A1909F}', '{B74DBB8C-2B2D-46E4-9508-DA46008D19A4}', '{E234C6C7-3837-4CE4-A541-CDC9627AAAC2}', '{6D4B3F52-63BF-435F-BF0C-C9D41120C2A3}', '{D88055F5-EAA0-4C8E-8249-C364A572BF7B}'],
                 2, [1, 1, 2, 1, 2, 2, 2, 2, 2, 2])
    
    #الختام
    elkhetam = som_values[0]
    #التوزيع
    mazmorELtawzy3 = som_values[1] +3
    mazmorELtawzy32 = som_values[2] - 1
    mrdMazmorEltawzy3 = som_values[3]
    #القسمة
    EsmaElsomElkbyr1 = som_values[4]
    #القرائات
    elengil3 = som_values[5]
    elmazmor3 = som_values[6]
    elebrksis3 = som_values[7]
    elkatholikon3 = som_values[8]
    elbouls3 = som_values[9]

    #اوشية الاهوية
    som_show_full_sections.extend(['{A20DA654-32F7-4B4C-96CB-C76232EB96E8}', '{BC7E3DCD-6AA8-44CC-B8AF-BC3E2BC71B5A}'])

    # som_hide_full_sections.extend(["سوتيس امين", "سوتيس امين 2", "مرد ابركسيس سنوي", "اكسيا", "الختام السنوي"])
    # som_show_full_sections.extend(["الليلويا اي اي ايخون", "نيف سينتى", " انثو تي تي شوري", 
    #                                "مرد الابركسيس لايام الصوم", "مرد انجيل ايام الصوم الكبير", "بي ماي رومي", 
    #                                "الختام في الصوم المقدس"])
    # som_hide_full_sections_ranges = [["الهيتنيات", "تكملة الهيتينيات"]]
    
    som_hide_full_sections.extend(['{A8A52E1F-44DD-45E1-A737-4E13E15D5F1F}', '{AF548422-5DCE-4418-8D21-7DB43CBC4C00}', '{D5DB63D0-39EE-49CE-8855-58CE02719834}', '{4D2B15D5-C978-467C-9D6C-726FE25128B8}', '{0D2A50D9-F484-4E60-922B-66FF81444E2C}'])
    som_show_full_sections.extend(['{315091E2-E367-43B7-A35E-4175DF947038}', '{229D9524-1F56-4456-A2B5-2321A4532E39}', '{456002DB-7C3A-44F7-87FE-507A15868231}', '{8DFA2B1F-C47F-42A1-A4F9-ED09CB4F6CB8}', '{9810C502-7526-4D63-96A4-F676E5AF5A5F}', '{44ABFE06-796C-4477-8C9D-E1B568FAD2FF}', '{4A3AE26D-6D71-4143-8C05-7618E08EF248}'])
    som_hide_full_sections_ranges.extend([['{79CED7F3-DA1D-467F-AA09-4187C8DE51E8}', '{646A8184-7F05-453A-A2F1-EB9A77D7F0EE}']])
    
    start_positions = [mazmorELtawzy3, elengil3, elmazmor3, elebrksis3, elkatholikon3, elbouls3]
    start_slides = [mrdMazmorEltawzy3, elengil, elmzmor, elebrksis1, elkatholikon1, elbouls1]
    end_slides = [mrdMazmorEltawzy3, elengil2, elmzmor, elebrksis2, elkatholikon2, elbouls2]

    show_hide_combined(prs1, excel, des_sheet, som_show_full_sections, som_hide_full_sections, [], som_hide_full_sections_ranges)
    
    som_show_values.extend([[EsmaElsomElkbyr1, EsmaElsomElkbyr1]])

    powerpoint = win32com.client.Dispatch("PowerPoint.Application")
    powerpoint.Visible = True  # Open PowerPoint application
    presentation1 = open_presentation_relative_path(prs1)
    presentation2 = open_presentation_relative_path(prs2)

    khetamElsom = find_slide_index_by_title(presentation1, "ايام الصوم الكبير", elkhetam)
    khetamElsom2 = find_slide_index_by_title(presentation1, "ايام الصوم الكبير 2", elkhetam)
    
    show_slides(presentation1, [[khetamElsom, khetamElsom2], [EsmaElsomElkbyr1, EsmaElsomElkbyr1]])

    # Initialize variables for current position, slide, and end index
    current_position = start_positions[0]
    current_start_slide = int(start_slides[0])
    current_end_slide = int(end_slides[0])

    # Initialize index for start position, slide, and end slide
    position_index = 1
    slide_index = 1
    end_index = 1

    while current_start_slide <= current_end_slide and slide_index <= presentation1.Slides.Count:
        if current_position == mazmorELtawzy3:
            source_slide = presentation1.Slides(current_start_slide)
            source_slide.Copy()
            slide_index1 = mazmorELtawzy32
            while slide_index1 >= current_position:
                presentation1.Slides.Paste(slide_index1).SlideShowTransition.Hidden = False
                slide_index1 -= 1
            current_start_slide += 1

        elif current_position in {elengil3, elmazmor3, elebrksis3, elkatholikon3, elbouls3}:
            source_slide = presentation2.Slides(current_end_slide)
            source_slide.Copy()
            presentation1.Slides.Paste(current_position).SlideShowTransition.Hidden = False
            current_end_slide -= 1
            if current_start_slide > current_end_slide:
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
    presentation1.SlideShowSettings.Run()

