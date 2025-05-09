import os
from commonFunctions import *
import win32com.client

def run_vba_with_slide_id(excel, sheet, prs, presentation, slide_id_pairs=None, sha3anyn = False):
    if slide_id_pairs is None:
        slide_id_pairs = []

    vba_values = find_slide_nums_arrays_v2(excel, sheet, 
                 ["{6C2572D7-3BC9-4B6E-B091-7E8D7012BBB4}", "{94BB223B-8F01-4E99-AA4F-47986D7DFFB1}", 
                  "{01F61BFD-210C-4F99-A7C7-7308CFAA93F4}"], 
                 2, [2, 2, 1])
                                        
    el2smaBasyly = vba_values[0]
    el2smaKirolosy = vba_values[1]
    el2sma = vba_values[2]

    vba_main_slides_ids = get_slide_ids_by_numbers(prs, [el2smaBasyly, el2smaKirolosy, el2sma])

    el2smaBasyly = vba_main_slides_ids[0]
    el2smaKirolosy = vba_main_slides_ids[1]
    el2sma = vba_main_slides_ids[2]

    # Access the VBA project
    vba_project = presentation.VBProject
    modules = vba_project.VBComponents

    # Add a new module to the VBA project
    new_module = modules.Add(1)  # 1 corresponds to a standard module

    # Generate the VBA code for the subroutine using SlideID
    vba_code = "Sub OnSlideShowPageChange()\n"
    vba_code += "    Dim currentSlideID As Long\n"
    vba_code += "    currentSlideID = ActivePresentation.SlideShowWindow.View.Slide.SlideID\n\n"
    vba_code += "    Select Case currentSlideID\n"
    vba_code += f"        Case {el2smaBasyly}\n"
    vba_code += f"            ActivePresentation.SlideShowWindow.View.GotoSlide GetSlideIndexByID({el2sma})\n"
    vba_code += f"        Case {el2smaKirolosy}\n"
    vba_code += f"            ActivePresentation.SlideShowWindow.View.GotoSlide GetSlideIndexByID({el2sma})\n"
    
    # Optional additional slide_id_pairs
    for target_slide_id, jump_to_slide_id in slide_id_pairs:
        vba_code += f"        Case {target_slide_id}\n"
        vba_code += f"            ActivePresentation.SlideShowWindow.View.GotoSlide GetSlideIndexByID({jump_to_slide_id})\n"
    if sha3anyn:
        sha3anyn_values = find_slide_nums_arrays_v2(excel, sheet,
                          ['{19FBAC32-77B5-4396-91CC-2127D0B8FF63}', '{19FBAC32-77B5-4396-91CC-2127D0B8FF63}', 
                           '{F9886962-A539-40A8-95D6-122DAFEC2303}', '{92F4141C-A127-4893-9542-743F62A24C83}',
                           '{0E543BA7-09D4-4CB8-8101-0BDC26CBCA40}'], 
                          2, [1, 2, 1, 2, 2])
        sha3anyn_oshya_engyl = sha3anyn_values[0]
        sha3anyn_oshya_engyl2 = sha3anyn_values[1]
        sha3anyn_oshya_engyl3 = sha3anyn_values[2]
        sha3anyn_elmazmor = sha3anyn_values[3]
        sha3anyn_elmazmor2 = sha3anyn_values[4]
        sha3anyn_oshya_engyl_slideID = get_slide_ids_by_numbers(prs, [sha3anyn_oshya_engyl, sha3anyn_oshya_engyl2, sha3anyn_oshya_engyl3, sha3anyn_elmazmor, sha3anyn_elmazmor2])
        sha3anyn_oshya_engyl = sha3anyn_oshya_engyl_slideID[0]
        sha3anyn_oshya_engyl2 = sha3anyn_oshya_engyl_slideID[1]
        sha3anyn_oshya_engyl3 = sha3anyn_oshya_engyl_slideID[2]
        sha3anyn_elmazmor = sha3anyn_oshya_engyl_slideID[3]
        sha3anyn_elmazmor2 = sha3anyn_oshya_engyl_slideID[4]
        vba_code += f"        Case {sha3anyn_oshya_engyl}\n"
        vba_code += f"            ActivePresentation.SlideShowWindow.View.GotoNamedShow " + '"temp"' + "\n"
        vba_code += f"        Case {sha3anyn_oshya_engyl2}\n"
        vba_code += f"            ActivePresentation.SlideShowWindow.View.EndNamedShow\n"
        vba_code += f"            StartSlideshow\n"
        vba_code += f"        Case {sha3anyn_elmazmor}\n"
        vba_code += f"            ActivePresentation.SlideShowWindow.View.GotoSlide GetSlideIndexByID({sha3anyn_elmazmor2})\n"
    vba_code += "    End Select\n"
    vba_code += "End Sub\n"
    vba_code += """Function GetSlideIndexByID(slideID As Long) As Long
    Dim slide As slide
    For Each slide In ActivePresentation.Slides
        If slide.slideID = slideID Then
            GetSlideIndexByID = slide.SlideIndex
            Exit Function
        End If
    Next slide
    ' Optional: add error handling if slide ID is not found
    MsgBox "Slide ID " & slideID & " not found.", vbExclamation
End Function
"""

    if sha3anyn:
        vba_code += """Sub StartSlideshow()
    With ActivePresentation.SlideShowSettings
            .ShowWithNarration = True
            .ShowWithAnimation = True
            .LoopUntilStopped = False
            .AdvanceMode = ppSlideShowUseSlideTimings
            .RangeType = ppShowAll
            .Run
        End With
        """
        vba_code += f"ActivePresentation.SlideShowWindow.View.GotoSlide GetSlideIndexByID({sha3anyn_oshya_engyl3})\n"
    vba_code += f"End Sub\n"

    
    # Add the generated code to the new module
    new_module.CodeModule.AddFromString(vba_code)

    # Set up the slideshow to call OnSlideShowPageChange on each slide change
    presentation.SlideShowSettings.Run()

    # Optionally run the macro immediately to initialize
    presentation.Application.Run("OnSlideShowPageChange")

    presentation.SlideShowWindow.View.Exit()

def agbya(presentation, slide_number, custom_show_number):
    ppMouseClick = 1
    ppActionNamedSlideShow = 7
    textbox_name = "المزامير"
    match custom_show_number:
        case 1:
            custom_show_name = "التالتة و السادسة"
        case 2:
            custom_show_name = "التالتة للتاسعة"
        case 3:
            custom_show_name = "التالتة للنوم"
        case 4:
            custom_show_name = "الثالثة فقط لعيد العنصرة"
    try:
        slide = presentation.Slides(slide_number)

        # Find the textbox by name
        shape = None
        for s in slide.Shapes:
            if s.Name == textbox_name:
                shape = s
                break

        if not shape:
            print(f"[❌] Textbox named '{textbox_name}' not found on slide {slide_number}.")
            return

        action = shape.ActionSettings(ppMouseClick)

        # Set action first
        action.Action = ppActionNamedSlideShow
        action.SlideShowName = custom_show_name
        action.ShowAndReturn = True  # Only now is it safe to set
    except Exception as e:
        print(f"[❌] Error occurred while setting action on slide {slide_number}: {e}")

"_____________________________________OLD CODE_DESIGN_____________________________________"

def odasSomElrosol (copticdate, Bishop=False, guestBishop=0):
    from copticDate import CopticCalendar
    cd = CopticCalendar().coptic_to_gregorian(copticdate)
    prs1 = relative_path(r"قداس.pptx")  # Using the relative path
    prs2 = relative_path(r"Data\القداسات\قداس صوم و عيد الرسل.pptx")
    excel = relative_path(r"بيانات القداسات.xlsx")
    excel2 = relative_path(r"Tables.xlsx")
    source_sheet = "الرسل"
    des_sheet ="سنوي"

    if cd.weekday() == 6:
        # sunday(prs1)
        prs3 = relative_path(r"Data\القطمارس\الاحاد\القطمارس السنوي احاد (القداس).pptx")
        katamars_sheet = "قطمارس الاحاد للقداس"
        km = copticdate[1]
        kd = (copticdate[2] - 1) // 7 + 1
    else: 
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

    des_sheet_values = find_slide_nums_arrays(excel, des_sheet, ["جي اف اسماروؤت", "جي اف اسماروؤت", "اسومين",  "اسومين",
                                                                "قسمة - أيها السيد الرب إلهنا", "قسمة - أيها السيد الرب إلهنا",
                                                                "قسمة للإبن تقال في صوم الرسل - أنت هو كلمة الآب", 
                                                                "قسمة للإبن تقال في صوم الرسل - أنت هو كلمة الآب",
                                                                "ايها الرب ـ الاسبسمس الواطس", "الاسبسمس الادام",
                                                                "ربع يقال في صوم الرسل", "ربع يقال في صوم الرسل",
                                                                "المزمور و الانجيل", "المزمور و الانجيل", "الابركسيس", "مرد ابركسيس الرسل", "مرد ابركسيس الرسل",
                                                                "الكاثوليكون", "لحن اندوس", "لحن اندوس", "البولس عربي", 
                                                                "طاي شوري", "طاي شوري", "الليلويا فاي بيبي", "الليلويا فاي بيبي",
                                                                "تي شوري", "تي شوري", "الليلويا جي افمفئي", "الليلويا جي افمفئي"], 
                                                                [1, 2, 1, 2, 1, 2, 1, 2, 1, 1, 1, 2, 1, 2, 2, 1, 2, 2, 1, 2, 2, 1, 2, 1, 2, 1, 2, 1, 2])

    #التوزيع
    jefsmarot = des_sheet_values[0]
    jefsmarot2 = des_sheet_values[1]
    asomyn = des_sheet_values[2]
    asomyn2 = des_sheet_values[3]

    #القسمة
    el2smaSanawy = des_sheet_values[4]
    el2smaSanawy2 = des_sheet_values[5]
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
        season8r8ory = find_slide_num(excel, des_sheet, "اوشية اهوية السماء غ", 1)
        season8r8ory2 = find_slide_num(excel, des_sheet, "اوشية اهوية السماء غ", 2)
    else:
        seasonbasyly = find_slide_num(excel, des_sheet, "اوشية المياة", 1)
        seasonbasyly2 = find_slide_num(excel, des_sheet, "اوشية المياة", 2)
        season8r8ory = find_slide_num(excel, des_sheet, "اوشية المياة غ", 1)
        season8r8ory2 = find_slide_num(excel, des_sheet, "اوشية المياة غ", 2) 

    #مرد الانجيل
    mrdelengilRosol = des_sheet_values[10]
    mrdelengilRosol2 = des_sheet_values[11]

    #المزمور و الانجيل
    elengil3 = des_sheet_values[13]
    elmzmor1 = des_sheet_values[12] + 2

    #القرائات
    elebrksis3 =  des_sheet_values[14]
    mrdelebrksisRosol = des_sheet_values[15]
    mrdelebrksisRosol2 = des_sheet_values[16]
    elkatholikon3 = des_sheet_values[17]
    ondos = des_sheet_values[18]
    ondos2 = des_sheet_values[19]
    elbouls3 = des_sheet_values[20]

    #تي شوري و الليلويا جي اف ميفي
    if cd.weekday() == 6 :
        shory1 = des_sheet_values[21]
        shory2 = des_sheet_values[22]
        allyloya1 = des_sheet_values[23]
        allyloya2 = des_sheet_values[24]
    else:
        shory1 = des_sheet_values[25]
        shory2 = des_sheet_values[26]
        allyloya1 = des_sheet_values[27]
        allyloya2 = des_sheet_values[28]

    if Bishop == True:
        prs4 = relative_path(r"Data\حضور الأسقف.pptx")
        sheet = "في حضور الأسقف"

        bishop_values = find_slide_nums_arrays(excel, sheet, ["صلاة الشكر", "صلاة الشكر", "بهموت غار الصغيرة", "بهموت غار الصغيرة"
                                                             "الهيتنيات", "الهيتنيات"], 
                                                             [1, 2, 1, 2, 1, 2])
        
        bishopDes_values = find_slide_nums_arrays(excel, des_sheet, ["صلاة الشكر", "ني سافيف تيرو", "ني سافيف تيرو", "باهموت غار الصغيرة",
                                                                    "امبين يوت اتطايوت", "امبين يوت اتطايوت", "الهيتنيات", 
                                                                    "اوشية الاباء (ب)", "اوشية الاباء غ"],
                                                                    [2, 1, 2, 2, 1, 2, 2, 2, 1])

        elshokr1 = bishop_values[0]
        elshokrDes = bishopDes_values[0] - 1

        nysaviv = bishopDes_values[1]
        nysaviv2 = bishopDes_values[2]

        bhmot8ar1 = bishop_values[2]
        bhmot8arDes = bishopDes_values[3] - 1

        embiniot = bishopDes_values[4] 
        embiniot2 = bishopDes_values[5]

        bishopHyten1 = bishop_values[4]
        bishopHytenDes = bishopDes_values[6] + 1

        elaba2basyly = bishopDes_values[7] - 1
        elaba28or8ory = bishopDes_values[8] - 1

        if guestBishop > 0:
            if guestBishop == 1:
                elshokr2 = bishop_values[1] - 1
                bhmot8ar2 = bishop_values[3] - 1
                bishopHyten2 = bishop_values[5] - 3
                elaba2 = elshokr2
                elaba22 = elshokr2
            
            elif guestBishop == 2:
                elshokr2 = bishop_values[1]
                bhmot8ar2 = bishop_values[3]
                bishopHyten2 = bishop_values[5]
                elaba2 = elshokr2 - 1
                elaba22 = elshokr2

            start_positions = [elaba28or8ory, elaba2basyly, elengil3, elmzmor1, elebrksis3, elkatholikon3, elbouls3, bhmot8arDes, bishopHytenDes, embiniot2, elshokrDes]
            start_slides = [elaba2, elaba2, elengil, elmzmor, elebrksis1, elkatholikon1, elbouls1, bhmot8ar1, bishopHyten1, bhmot8ar1, elshokr1]
            end_slides = [elaba22, elaba22, elengil2, elmzmor, elebrksis2, elkatholikon2, elbouls2, bhmot8ar2, bishopHyten2, bhmot8ar2, elshokr2]

        else:
            elshokr2 = bishop_values[1] - 2
            start_positions = [elengil3, elmzmor1, elebrksis3, elkatholikon3, elbouls3, elshokrDes]
            start_slides = [elengil, elmzmor, elebrksis1, elkatholikon1, elbouls1, elshokr1]
            end_slides = [elengil2, elmzmor, elebrksis2, elkatholikon2, elbouls2, elshokr2]
        show_array = [[jefsmarot, jefsmarot2], [asomyn, asomyn2], [el2smaRosol, el2smaRosol2], [season8r8ory, season8r8ory2], 
                      [seasonbasyly, seasonbasyly2], [mrdelengilRosol, mrdelengilRosol2], [mrdelebrksisRosol, mrdelebrksisRosol2], 
                      [ondos, ondos2], [embiniot, embiniot2], [shory1, shory2], [nysaviv, nysaviv2], [allyloya1, allyloya2]]
            
    else:
        start_positions = [elengil3, elmzmor1, elebrksis3, elkatholikon3, elbouls3]
        start_slides = [elengil, elmzmor, elebrksis1, elkatholikon1, elbouls1]
        end_slides = [elengil2, elmzmor, elebrksis2, elkatholikon2, elbouls2]
        show_array = [[jefsmarot, jefsmarot2], [asomyn, asomyn2], [el2smaRosol, el2smaRosol2], [season8r8ory, season8r8ory2], 
                      [seasonbasyly, seasonbasyly2], [mrdelengilRosol, mrdelengilRosol2], [mrdelebrksisRosol, mrdelebrksisRosol2], 
                      [ondos, ondos2], [shory1, shory2], [allyloya1, allyloya2]]

    hide_array = [[el2smaSanawy, el2smaSanawy2]]

    powerpoint = win32com.client.Dispatch("PowerPoint.Application")
    powerpoint.Visible = True  # Open PowerPoint application
    presentation1 = open_presentation_relative_path(prs1)
    presentation2 = open_presentation_relative_path(prs2)
    presentation3 = open_presentation_relative_path(prs3)
    if Bishop == True:
        presentation4 = open_presentation_relative_path(prs4)

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

        elif Bishop == True and (current_position == elaba28or8ory or current_position == elaba2basyly or 
                                 current_position == bhmot8arDes or current_position == bishopHytenDes or 
                                 current_position == embiniot2 or current_position == elshokrDes):
            source_slide = presentation4.Slides(current_end_slide)
            source_slide.Copy()
            new_slide = presentation1.Slides.Paste(current_position)
            new_slide.SlideShowTransition.Hidden = False
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

    move_section_names = [
        "جي اف اسماروؤت",
        "اسبسمس ادام لصوم الرسل",
        "ختام الاسبسمس الادام",
        "اسبسمس واطس لصوم الرسل",
        "ختام الأسبسمس الواطس"
    ]

    target_section_names = [
        "مزمور التوزيع",
        "الاسبسمس الادام",
        "اسبسمس ادام لصوم الرسل",
        "أسبسمس واطس",
        "اسبسمس واطس لصوم الرسل"
    ]

    move_sections(presentation1, move_section_names, target_section_names)


    # Close presentations
    presentation2.Close()
    presentation3.Close()
    if Bishop == True:
        presentation4.Close()

def odas3ydElrosol (Bishop=False, guestBishop=0):
    from copticDate import CopticCalendar
    cd = CopticCalendar().coptic_to_gregorian([1740, 11, 5])
    prs1 = relative_path(r"قداس.pptx")  # Using the relative path
    prs2 = relative_path(r"Data\القداسات\قداس صوم و عيد الرسل.pptx")
    excel = relative_path(r"بيانات القداسات.xlsx")
    excel2 = relative_path(r"Tables.xlsx")
    source_sheet = "الرسل"
    des_sheet ="سنوي"

    if cd.weekday() == 6:
        # sunday(prs1)
        prs3 = relative_path(r"Data\القطمارس\الاحاد\القطمارس السنوي احاد (القداس).pptx")
        katamars_sheet = "قطمارس الاحاد للقداس"
        km, kd = 11, 1
    else: 
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

    des_sheet_values = find_slide_nums_arrays(excel, des_sheet, ["جي اف اسماروؤت", "جي اف اسماروؤت",
                                                                "قسمة - أيها السيد الرب إلهنا", "قسمة - أيها السيد الرب إلهنا",
                                                                "قسمة للإبن تقال في صوم الرسل - أنت هو كلمة الآب", 
                                                                "قسمة للإبن تقال في صوم الرسل - أنت هو كلمة الآب",
                                                                "ايها الرب ـ الاسبسمس الواطس", "الاسبسمس الادام",
                                                                "ربع يقال في عيد الرسل", "ربع يقال في عيد الرسل",
                                                                "المزمور و الانجيل", "المزمور و الانجيل", "الابركسيس", "مرد ابركسيس عيد الرسل", "مرد ابركسيس عيد الرسل",
                                                                "الكاثوليكون", "لحن اندوس", "لحن اندوس", "البولس عربي", "الهيتنيات",
                                                                "تي شوري", "تي شوري", "الليلويا جي افمفئي", "الليلويا جي افمفئي", 
                                                                "طاي شوري", "طاي شوري", "الليلويا فاي بيبي", "الليلويا فاي بيبي"], 
                                                                [1, 2, 1, 2, 1, 2, 1, 1, 1, 2, 1, 2, 2, 1, 2, 2, 1, 2, 2, 1, 1, 2, 1, 2, 1, 2, 1, 2])

    
    #التوزيع
    jefsmarot = des_sheet_values[0]
    jefsmarot2 = des_sheet_values[1]

    #القسمة
    el2smaSanawy = des_sheet_values[2]
    el2smaSanawy2 = des_sheet_values[3]
    el2smaRosol = des_sheet_values[4]
    el2smaRosol2 = des_sheet_values[5]

    #الاواشي
    seasonbasyly = find_slide_num(excel, des_sheet, "اوشية المياة", 1)
    seasonbasyly2 = find_slide_num(excel, des_sheet, "اوشية المياة", 2)
    season8r8ory = find_slide_num(excel, des_sheet, "اوشية المياة غ", 1)
    season8r8ory2 = find_slide_num(excel, des_sheet, "اوشية المياة غ", 2) 

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
    ondos = des_sheet_values[16]
    ondos2 = des_sheet_values[17]
    elbouls3 = des_sheet_values[18]

    #الهيتنيات
    elhiteniat = des_sheet_values[19]

    #تي شوري و الليلويا جي اف ميفي
    if cd.weekday() == 6 :
        shory1 = des_sheet_values[24]
        shory2 = des_sheet_values[25]
        allyloya1 = des_sheet_values[26]
        allyloya2 = des_sheet_values[27]
    else:
        shory1 = des_sheet_values[20]
        shory2 = des_sheet_values[21]
        allyloya1 = des_sheet_values[22]
        allyloya2 = des_sheet_values[23]

    if Bishop == True:
        prs4 = relative_path(r"Data\حضور الأسقف.pptx")
        sheet = "في حضور الأسقف"

        bishop_values = find_slide_nums_arrays(excel, sheet, ["صلاة الشكر", "صلاة الشكر", "بهموت غار الصغيرة", "بهموت غار الصغيرة"
                                                             "الهيتنيات", "الهيتنيات"], 
                                                             [1, 2, 1, 2, 1, 2])
        
        bishopDes_values = find_slide_nums_arrays(excel, des_sheet, ["صلاة الشكر", "ني سافيف تيرو", "ني سافيف تيرو", "باهموت غار الصغيرة",
                                                                    "امبين يوت اتطايوت", "امبين يوت اتطايوت", "الهيتنيات", 
                                                                    "اوشية الاباء (ب)", "اوشية الاباء غ"],
                                                                    [2, 1, 2, 2, 1, 2, 2, 2, 1])

        elshokr1 = bishop_values[0]
        elshokrDes = bishopDes_values[0] - 1

        nysaviv = bishopDes_values[1]
        nysaviv2 = bishopDes_values[2]

        bhmot8ar1 = bishop_values[2]
        bhmot8arDes = bishopDes_values[3] - 1

        embiniot = bishopDes_values[4] 
        embiniot2 = bishopDes_values[5]

        bishopHyten1 = bishop_values[4]
        bishopHytenDes = bishopDes_values[6] + 1

        elaba2basyly = bishopDes_values[7] - 1
        elaba28or8ory = bishopDes_values[8] - 1

        if guestBishop > 0:
            if guestBishop == 1:
                elshokr2 = bishop_values[1] - 1
                bhmot8ar2 = bishop_values[3] - 1
                bishopHyten2 = bishop_values[5] - 3
                elaba2 = elshokr2
                elaba22 = elshokr2
            
            elif guestBishop == 2:
                elshokr2 = bishop_values[1]
                bhmot8ar2 = bishop_values[3]
                bishopHyten2 = bishop_values[5]
                elaba2 = elshokr2 - 1
                elaba22 = elshokr2

            start_positions = [elaba28or8ory, elaba2basyly, elengil3, elmzmor1, elebrksis3, elkatholikon3, elbouls3, bhmot8arDes, bishopHytenDes, embiniot2, elshokrDes]
            start_slides = [elaba2, elaba2, elengil, elmzmor, elebrksis1, elkatholikon1, elbouls1, bhmot8ar1, bishopHyten1, bhmot8ar1, elshokr1]
            end_slides = [elaba22, elaba22, elengil2, elmzmor, elebrksis2, elkatholikon2, elbouls2, bhmot8ar2, bishopHyten2, bhmot8ar2, elshokr2]

        else:
            elshokr2 = bishop_values[1] - 2
            start_positions = [elengil3, elmzmor1, elebrksis3, elkatholikon3, elbouls3, elshokrDes]
            start_slides = [elengil, elmzmor, elebrksis1, elkatholikon1, elbouls1, elshokr1]
            end_slides = [elengil2, elmzmor, elebrksis2, elkatholikon2, elbouls2, elshokr2]
        show_array = [[jefsmarot, jefsmarot2], [el2smaRosol, el2smaRosol2], [season8r8ory, season8r8ory2], 
                      [seasonbasyly, seasonbasyly2], [mrdelengilRosol, mrdelengilRosol2], [mrdelebrksisRosol, mrdelebrksisRosol2], 
                      [ondos, ondos2], [embiniot, embiniot2], [shory1, shory2], [nysaviv, nysaviv2], [allyloya1, allyloya2]]
            
    else:
        start_positions = [elengil3, elmzmor1, elebrksis3, elkatholikon3, elbouls3]
        start_slides = [elengil, elmzmor, elebrksis1, elkatholikon1, elbouls1]
        end_slides = [elengil2, elmzmor, elebrksis2, elkatholikon2, elbouls2]
        show_array = [[jefsmarot, jefsmarot2], [el2smaRosol, el2smaRosol2], [season8r8ory, season8r8ory2], 
                      [seasonbasyly, seasonbasyly2], [mrdelengilRosol, mrdelengilRosol2], [mrdelebrksisRosol, mrdelebrksisRosol2], 
                      [ondos, ondos2], [shory1, shory2], [allyloya1, allyloya2]]

    hide_array = [[el2smaSanawy, el2smaSanawy2]]

    powerpoint = win32com.client.Dispatch("PowerPoint.Application")
    powerpoint.Visible = True  # Open PowerPoint application
    presentation1 = open_presentation_relative_path(prs1)
    presentation2 = open_presentation_relative_path(prs2)
    presentation3 = open_presentation_relative_path(prs3)
    pnpRob3 = find_slide_index_by_title(presentation1, "بصلوات سيدي الأبوين الرسولين أبينا بطرس ومعلمنا بولس:", elhiteniat)
    show_array.append([pnpRob3, pnpRob3+1])
    if Bishop == True:
        presentation4 = open_presentation_relative_path(prs4)

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

        elif Bishop == True and (current_position == elaba28or8ory or current_position == elaba2basyly or 
                                 current_position == bhmot8arDes or current_position == bishopHytenDes or 
                                 current_position == embiniot2 or current_position == elshokrDes):
            source_slide = presentation4.Slides(current_end_slide)
            source_slide.Copy()
            new_slide = presentation1.Slides.Paste(current_position)
            new_slide.SlideShowTransition.Hidden = False
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
    presentation2.Close()
    presentation3.Close()
    if Bishop == True:
        presentation4.Close()

def odasEltagaly (Bishop=False, guestBishop=0):
    prs1 = relative_path(r"قداس.pptx")  # Using the relative path
    prs2 = relative_path(r"Data\القداسات\قداس عيد التجلي.pptx")
    excel = relative_path(r"بيانات القداسات.xlsx")
    source_sheet = "التجلي"
    des_sheet ="سنوي"

    des_sheet_values = find_slide_nums_arrays(excel, des_sheet, ["مزمور التوزيع", "مزمور التوزيع",
                                                                "قسمة - أيها السيد الرب إلهنا", "قسمة - أيها السيد الرب إلهنا",
                                                                "قسمة للإبن تقال في الأعياد السيدية وسنوى - نسبح ونمجد إله الآلهة", "قسمة للإبن تقال في الأعياد السيدية وسنوى - نسبح ونمجد إله الآلهة",
                                                                "اوشية المياة", "اوشية المياة", "اوشية المياة غ", "اوشية المياة غ",
                                                                "مرد انجيل التجلي", "فاي اريه بي اوو", "مرد الانجيل", "مرد الانجيل",
                                                                "المزمور و الانجيل", "المزمور و الانجيل", "مرد المزمور", "مرد مزمور التجلي", 
                                                                "السنكسار", "الابركسيس", "مرد ابركسيس التجلي", "مرد ابركسيس التجلي",
                                                                "الكاثوليكون", "البولس عربي", "طاي شوري", "طاي شوري", 
                                                                "الليلويا فاي بيبي", "الليلويا فاي بيبي"], 
                                                                [1, 2, 1, 2, 1, 2, 1, 2, 1, 2, 1, 1, 1, 2, 1, 2, 1, 2, 2, 2, 1, 2, 2, 2, 1, 2, 1, 2])
    
    source_sheet_values = find_slide_nums_arrays(excel, source_sheet, ["مرد التوزيع عيد التجلي", "الانجيل", "الانجيل",
                                                                       "المزمور", "السنكسار", "السنكسار",
                                                                       "الابركسيس", "الابركسيس", "الكاثوليكون", "الكاثوليكون",
                                                                       "البولس", "البولس"], 
                                                                      [1, 1, 2, 1, 1, 2, 1, 2, 1, 2, 1, 2])
    
    #التوزيع
    sn = source_sheet_values[0]
    fs = des_sheet_values[0] + 1 
    ls = des_sheet_values[1] - 1

    #القسمة
    el2smaSanawy = des_sheet_values[2]
    el2smaSanawy2 = des_sheet_values[3]
    el2smaTagaly = des_sheet_values[4]
    el2smaTagaly2 = des_sheet_values[5]

    #الاواشي
    seasonbasyly = des_sheet_values[6]
    seasonbasyly2 = des_sheet_values[7]
    season8r8ory = des_sheet_values[8]
    season8r8ory2 = des_sheet_values[9]

    #مرد الإنجيل
    mrdengilTagaly = des_sheet_values[10]
    fayereby = des_sheet_values[11]
    mrdengil = des_sheet_values[12]
    mrdengil2 = des_sheet_values[13]

    #المزمور و الانجيل
    elengil = source_sheet_values[1]
    elengil2 = source_sheet_values[2]
    elengil3 = des_sheet_values[15]
    elmzmor = source_sheet_values[3]
    elmzmor3 = des_sheet_values[14] + 1
    mrdElmzmor = des_sheet_values[16]
    mrdElmzmorTagaly = des_sheet_values[17]

    #الاقرائات
    elsnksar = source_sheet_values[4]
    elsnksar2 = source_sheet_values[5]
    elsenksar3 = des_sheet_values[18]
    elebrksis = source_sheet_values[6]
    elebrksis2 = source_sheet_values[7]
    elebrksis3 = des_sheet_values[19]
    mrdelebrksis = des_sheet_values[20]
    mrdelebrksis2 = des_sheet_values[21]
    elkatholikon = source_sheet_values[8]
    elkatholikon2 = source_sheet_values[9]
    elkatholikon3 = des_sheet_values[22]
    elbouls = source_sheet_values[10]
    elbouls2 = source_sheet_values[11]
    elbouls3 = des_sheet_values[23]

    #تي شوري و الليلويا جي اف ميفي
    shory1 = des_sheet_values[24]
    shory2 = des_sheet_values[25]
    allyloya1 = des_sheet_values[26]
    allyloya2 = des_sheet_values[27]

    if Bishop == True:
        prs3 = relative_path(r"Data\حضور الأسقف.pptx")
        sheet = "في حضور الأسقف"

        bishop_values = find_slide_nums_arrays(excel, sheet, ["صلاة الشكر", "صلاة الشكر", "بهموت غار الصغيرة", "بهموت غار الصغيرة"
                                                             "الهيتنيات", "الهيتنيات"], 
                                                             [1, 2, 1, 2, 1, 2])
        
        bishopDes_values = find_slide_nums_arrays(excel, des_sheet, ["صلاة الشكر", "ني سافيف تيرو", "ني سافيف تيرو", "باهموت غار الصغيرة",
                                                                    "امبين يوت اتطايوت", "امبين يوت اتطايوت", "الهيتنيات", 
                                                                    "اوشية الاباء (ب)", "اوشية الاباء غ"],
                                                                    [2, 1, 2, 2, 1, 2, 2, 2, 1])

        elshokr1 = bishop_values[0]
        elshokrDes = bishopDes_values[0] - 1

        nysaviv = bishopDes_values[1]
        nysaviv2 = bishopDes_values[2]

        bhmot8ar1 = bishop_values[2]
        bhmot8arDes = bishopDes_values[3] - 1

        embiniot = bishopDes_values[4] 
        embiniot2 = bishopDes_values[5]

        bishopHyten1 = bishop_values[4]
        bishopHytenDes = bishopDes_values[6] + 1

        elaba2basyly = bishopDes_values[7] - 1
        elaba28or8ory = bishopDes_values[8] - 1

        if guestBishop > 0:
            if guestBishop == 1:
                elshokr2 = bishop_values[1] - 1
                bhmot8ar2 = bishop_values[3] - 1
                bishopHyten2 = bishop_values[5] - 3
                elaba2 = elshokr2
                elaba22 = elshokr2
            
            elif guestBishop == 2:
                elshokr2 = bishop_values[1]
                bhmot8ar2 = bishop_values[3]
                bishopHyten2 = bishop_values[5]
                elaba2 = elshokr2 - 1
                elaba22 = elshokr2

            start_positions = [fs, elaba28or8ory, elaba2basyly, elengil3, elmzmor3, elebrksis3, elkatholikon3, elbouls3, bhmot8arDes, bishopHytenDes, embiniot2, elshokrDes]
            start_slides = [sn, elaba2, elaba2, elengil, elmzmor, elebrksis, elkatholikon, elbouls, bhmot8ar1, bishopHyten1, bhmot8ar1, elshokr1]
            end_slides = [sn, elaba22, elaba22, elengil2, elmzmor, elebrksis2, elkatholikon2, elbouls2, bhmot8ar2, bishopHyten2, bhmot8ar2, elshokr2]

        else:
            elshokr2 = bishop_values[1] - 2
            start_positions = [elengil3, elmzmor3, elebrksis3, elkatholikon3, elbouls3, elshokrDes]
            start_slides = [elengil, elmzmor, elebrksis, elkatholikon, elbouls, elshokr1]
            end_slides = [elengil2, elmzmor, elebrksis2, elkatholikon2, elbouls2, elshokr2]
        show_array = [[el2smaTagaly, el2smaTagaly2], [season8r8ory, season8r8ory2], [seasonbasyly, seasonbasyly2], 
                      [mrdengilTagaly, mrdengilTagaly], [fayereby, fayereby], [mrdelebrksis, mrdelebrksis2], 
                      [embiniot, embiniot2], [shory1, shory2], [nysaviv, nysaviv2], [allyloya1, allyloya2]]
            
    else:
        start_positions = [fs, elengil3, elmzmor3, elsenksar3, elebrksis3, elkatholikon3, elbouls3]
        start_slides = [sn, elengil, elmzmor, elsnksar, elebrksis, elkatholikon, elbouls]
        end_slides = [sn, elengil2, elmzmor, elsnksar2, elebrksis2, elkatholikon2, elbouls2]
        show_array = [[el2smaTagaly, el2smaTagaly2], [season8r8ory, season8r8ory2], [seasonbasyly, seasonbasyly2], 
                      [mrdengilTagaly, mrdengilTagaly], [fayereby, fayereby], [mrdElmzmorTagaly, mrdElmzmorTagaly], 
                      [mrdelebrksis, mrdelebrksis2], [shory1, shory2], [allyloya1, allyloya2]]

    hide_array = [[el2smaSanawy, el2smaSanawy2], [mrdengil, mrdengil2], [mrdElmzmor, mrdElmzmor]]

    powerpoint = win32com.client.Dispatch("PowerPoint.Application")
    powerpoint.Visible = True  # Open PowerPoint application
    script_directory = os.path.dirname(os.path.abspath(__file__))
    absolute_path = os.path.join(script_directory, prs1)
    presentation1 = powerpoint.Presentations.Open(absolute_path)
    presentation2 = open_presentation_relative_path(prs2)

    if Bishop == True:
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

                if slide_index1 > ls - 14 and slide_index1 <= ls - 3:
                    new_slide.SlideShowTransition.Hidden = True
                
                slide_index1 -= 1
            current_start_slide += 1

        elif (current_position == elengil3 or current_position == elmzmor3 or current_position == elebrksis3 
            or current_position == elkatholikon3 or current_position == elbouls3):
            source_slide = presentation2.Slides(current_end_slide)
            source_slide.Copy()
            new_slide = presentation1.Slides.Paste(current_position).SlideShowTransition.Hidden = False
            current_end_slide -= 1
            if(current_start_slide > current_end_slide):
                current_position += 1

        elif Bishop == True and (current_position == elaba28or8ory or current_position == elaba2basyly or 
                                 current_position == bhmot8arDes or current_position == bishopHytenDes or 
                                 current_position == embiniot2 or current_position == elshokrDes):
            source_slide = presentation3.Slides(current_end_slide)
            source_slide.Copy()
            new_slide = presentation1.Slides.Paste(current_position)
            new_slide.SlideShowTransition.Hidden = False
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

    presentation2.Close()
    if Bishop == True:
        presentation3.Close()

def odasElnayrooz (copticdate, Bishop=False, guestBishop=0):
    from copticDate import CopticCalendar
    cd = CopticCalendar().coptic_to_gregorian(copticdate)
    prs1 = relative_path(r"قداس.pptx")  # Using the relative path
    prs2 = relative_path(r"Data\القداسات\قداس عيد النيروز و الصليب.pptx")
    excel = relative_path(r"بيانات القداسات.xlsx")
    excel2 = relative_path(r"Tables.xlsx")
    des_sheet ="سنوي"

    if cd.weekday() == 6:
        # sunday(prs1)
        prs3 = relative_path(r"Data\القطمارس\الاحاد\القطمارس السنوي احاد (القداس).pptx")
        katamars_sheet = "قطمارس الاحاد للقداس"
        km = copticdate[1]
        kd = (copticdate[2] - 1) // 7 + 1
    else: 
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
                                                                "اوشية المياة غ", "اوشية المياة غ", "اوشية المياة", "اوشية المياة", 
                                                                "مرد انجيل النيروز", "مرد انجيل النيروز", "فاي اريه بي اوو", "فاي اريه بي اوو",
                                                                "مرد الانجيل", "مرد الانجيل", "مرد مزمور النيروز", "مرد مزمور النيروز", "مرد المزمور",
                                                                "مرد ابركسيس النيروز", "مرد ابركسيس النيروز",
                                                                "طاي شوري", "طاي شوري", "الليلويا فاي بيبي", "الليلويا فاي بيبي", "المزمور و الانجيل", "المزمور و الانجيل",
                                                                "الابركسيس", "الكاثوليكون", "البولس عربي", "مرد توزيع النيروز"
                                                                ], 
                                                                [1, 2, 1, 2, 1, 2, 1, 2, 1, 2, 1, 2, 1, 2, 1, 2, 1, 2, 1, 1, 2, 1, 2, 1, 2, 
                                                                 1, 2, 2, 2, 2, 1])
    
    
    #التوزيع
    sn = des_sheet_values[30]
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
    mrdengilNayrooz = des_sheet_values[10]
    mrdengilNayrooz2 = des_sheet_values[11]
    fayereby = des_sheet_values[12]
    fayereby2 = des_sheet_values[13]
    mrdengil = des_sheet_values[14]
    mrdengil2 = des_sheet_values[15]

    #المزمور و الانجيل
    elengil3 = des_sheet_values[26]
    elmzmor1 = des_sheet_values[25] + 1

    #مرد المزمور
    mrdElmzmorNayrooz = des_sheet_values[16]
    mrdElmzmorNayrooz2 = des_sheet_values[17]
    mrdElmzmor = des_sheet_values[18]

    #مرد الابركسيس
    mrdelebrksis = des_sheet_values[19]
    mrdelebrksis2 = des_sheet_values[20]

    #القرائات
    elebrksis3 =  des_sheet_values[27]
    elkatholikon3 = des_sheet_values[28]
    elbouls3 = des_sheet_values[29]

    #الليلويا فاي بيبي و طاي شوري
    shory1 = des_sheet_values[21]
    shory2 = des_sheet_values[22]
    allyloya1 = des_sheet_values[23]
    allyloya2 = des_sheet_values[24]

    if Bishop == True:
        prs4 = relative_path(r"Data\حضور الأسقف.pptx")
        sheet = "في حضور الأسقف"

        bishop_values = find_slide_nums_arrays(excel, sheet, ["صلاة الشكر", "صلاة الشكر", "بهموت غار الصغيرة", "بهموت غار الصغيرة"
                                                             "الهيتنيات", "الهيتنيات"], 
                                                             [1, 2, 1, 2, 1, 2])
        
        bishopDes_values = find_slide_nums_arrays(excel, des_sheet, ["صلاة الشكر", "ني سافيف تيرو", "ني سافيف تيرو", "باهموت غار الصغيرة",
                                                                    "امبين يوت اتطايوت", "امبين يوت اتطايوت", "الهيتنيات", 
                                                                    "اوشية الاباء (ب)", "اوشية الاباء غ"],
                                                                    [2, 1, 2, 2, 1, 2, 2, 2, 1])

        elshokr1 = bishop_values[0]
        elshokrDes = bishopDes_values[0] - 1

        nysaviv = bishopDes_values[1]
        nysaviv2 = bishopDes_values[2]

        bhmot8ar1 = bishop_values[2]
        bhmot8arDes = bishopDes_values[3] - 1

        embiniot = bishopDes_values[4] 
        embiniot2 = bishopDes_values[5]

        bishopHyten1 = bishop_values[4]
        bishopHytenDes = bishopDes_values[6] + 1

        elaba2basyly = bishopDes_values[7] - 1
        elaba28or8ory = bishopDes_values[8] - 1

        if guestBishop > 0:
            if guestBishop == 1:
                elshokr2 = bishop_values[1] - 1
                bhmot8ar2 = bishop_values[3] - 1
                bishopHyten2 = bishop_values[5] - 3
                elaba2 = elshokr2
                elaba22 = elshokr2
            
            elif guestBishop == 2:
                elshokr2 = bishop_values[1]
                bhmot8ar2 = bishop_values[3]
                bishopHyten2 = bishop_values[5]
                elaba2 = elshokr2 - 1
                elaba22 = elshokr2

            start_positions = [fs, elaba28or8ory, elaba2basyly, elengil3, elmzmor1, elebrksis3, elkatholikon3, elbouls3, bhmot8arDes, bishopHytenDes, embiniot2, elshokrDes]
            start_slides = [sn, elaba2, elaba2, elengil, elmzmor, elebrksis1, elkatholikon1, elbouls1, bhmot8ar1, bishopHyten1, bhmot8ar1, elshokr1]
            end_slides = [sn, elaba22, elaba22, elengil2, elmzmor, elebrksis2, elkatholikon2, elbouls2, bhmot8ar2, bishopHyten2, bhmot8ar2, elshokr2]

        else:
            elshokr2 = bishop_values[1] - 2
            start_positions = [elengil3, elmzmor1, elebrksis3, elkatholikon3, elbouls3, elshokrDes]
            start_slides = [elengil, elmzmor, elebrksis1, elkatholikon1, elbouls1, elshokr1]
            end_slides = [elengil2, elmzmor, elebrksis2, elkatholikon2, elbouls2, elshokr2]
        show_array = [[el2smaNayrooz, el2smaNayrooz2], [season8r8ory, season8r8ory2], [seasonbasyly, seasonbasyly2], 
                      [mrdengilNayrooz, mrdengilNayrooz2], [fayereby, fayereby], [mrdElmzmorNayrooz, mrdElmzmorNayrooz2],
                      [mrdelebrksis, mrdelebrksis2], [embiniot, embiniot2], [shory1, shory2], [nysaviv, nysaviv2], [allyloya1, allyloya2]]
            
    else:
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
    presentation3 = open_presentation_relative_path(prs3)

    if Bishop == True:
        presentation4 = open_presentation_relative_path(prs4)

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
            source_slide = presentation1.Slides(current_start_slide)
            source_slide.Copy()
            slide_index1 = ls
            while slide_index1 >= current_position:
                new_slide = presentation1.Slides.Paste(slide_index1).SlideShowTransition.Hidden = False
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

        elif Bishop == True and (current_position == elaba28or8ory or current_position == elaba2basyly or 
                                 current_position == bhmot8arDes or current_position == bishopHytenDes or 
                                 current_position == embiniot2 or current_position == elshokrDes):
            source_slide = presentation4.Slides(current_end_slide)
            source_slide.Copy()
            new_slide = presentation1.Slides.Paste(current_position)
            new_slide.SlideShowTransition.Hidden = False
            current_end_slide -= 1
            if(current_start_slide > current_end_slide):
                current_position += 1

        else:
            source_slide = presentation1.Slides(current_start_slide)
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
        "اسبسمس ادام للنيروز",
        "ختام الاسبسمس الادام",
        "الأسبسمس الواطس للنيروز",
        "ختام الأسبسمس الواطس"
    ]

    target_section_names = [
        "القسمة",
        "الاسبسمس الادام",
        "اسبسمس ادام للنيروز",
        "أسبسمس واطس",
        "الأسبسمس الواطس للنيروز"
    ]

    # Call the function once for all moves
    move_sections(presentation1, move_section_names, target_section_names)

    presentation3.Close()
    if Bishop == True:
        presentation4.Close()   

def odasKiahk (copticdate, Bishop=False, guestBishop=0):
    from copticDate import CopticCalendar
    cd = CopticCalendar().coptic_to_gregorian(copticdate)
    prs1 = relative_path(r"قداس.pptx") 
    excel = relative_path(r"بيانات القداسات.xlsx")
    excel2 = relative_path(r"Tables.xlsx")
    des_sheet ="سنوي"

    if cd.weekday() == 6:
        # sunday(prs1)
        prs2 = relative_path(r"Data\القطمارس\الاحاد\القطمارس السنوي احاد (القداس).pptx")
        katamars_sheet = "قطمارس الاحاد للقداس"
        km = copticdate[1]
        kd = (copticdate[2] - 1) // 7 + 1
    else: 
        prs2 = relative_path(r"Data\القطمارس\الايام\القطمارس السنوي ايام (القداس).pptx")
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

    kiahk_values = find_slide_nums_arrays(excel, des_sheet, 
                                          ["تكملة على حسب المناسبة", "مدائح كيهك", "بي اويك", "بي اويك", "اك اسماروؤت", "اك اسماروؤت", 
                                           "قسمة - أيها السيد الرب إلهنا", "قسمة - أيها السيد الرب إلهنا",
                                           "قسمة للأب في صوم و عيد الميلاد - أيها السيد الرب إلهنا", "قسمة للأب في صوم و عيد الميلاد - أيها السيد الرب إلهنا",
                                           "اوشية الزروع  العشب", "اوشية الزروع  العشب", 
                                           "اوشية الزروع  العشب غ", "اوشية الزروع  العشب غ", "مرد الانجيل", "مرد الانجيل", 
                                           "مرد انجيل كيهك 1", "مرد انجيل كيهك 1", "مرد انجيل كيهك 2 و صوم العذراء", "مرد انجيل كيهك 2 و صوم العذراء",
                                           "تكملة مشتركة لكيهك", "تكملة مشتركة لكيهك", "المزمور و الانجيل", "المزمور و الانجيل", 
                                           "مرد ابركسيس كيهك 1و3", "مرد ابركسيس كيهك 1و3", 
                                           "مرد ابركسيس كيهك 2و4 و الملاك غبريال", "مرد ابركسيس كيهك 2و4 و الملاك غبريال", 
                                           "الابركسيس", "الكاثوليكون", "البولس عربي", "الهيتنيات",
                                           "طاي شوري", "طاي شوري", "الليلويا فاي بيبي", "الليلويا فاي بيبي",
                                           "تي شوري", "تي شوري", "الليلويا جي افمفئي", "الليلويا جي افمفئي"], 
                                          [1, 1, 1, 2, 1, 2, 1, 2, 1, 2, 1, 2, 1, 2, 1, 2, 1, 2, 1, 2, 1, 2, 1, 2, 1, 2, 1, 2,
                                           2, 2, 2, 1, 1, 2, 1, 2, 1, 2, 1, 2])

    #الختام
    elkhetam = kiahk_values[0]

    #التوزيع
    elmaday7 = kiahk_values[1]
    byoyk = kiahk_values[2]
    byoyk2 = kiahk_values[3]
    ekesmarot = kiahk_values[4]
    ekesmarot2 = kiahk_values[5]

    #القسمة
    el2smaElsanawy = kiahk_values[6]
    el2smaElsanawy2 = kiahk_values[7]
    el2smaElmilad = kiahk_values[8]
    el2smaElmilad2 = kiahk_values[9]

    #الاواشي
    seasonbasyly = kiahk_values[10]
    seasonbasyly2 = kiahk_values[11]
    season8r8ory = kiahk_values[12]
    season8r8ory2 = kiahk_values[13]

    #مرد الانجيل
    mrdengil = kiahk_values[14]
    mrdengil2 = kiahk_values[15]
    
    if copticdate[2] <= 14:
        mrdengilkiahk = kiahk_values[16]
        mrdengilkiahk2 = kiahk_values[17]
    else:
        mrdengilkiahk = kiahk_values[18]
        mrdengilkiahk2 = kiahk_values[19]
    
    mrdelengiltakmla = kiahk_values[20]
    mrdelengiltakmla2 = kiahk_values[21]

    #المزمور و الانجيل
    elmazmor2 = kiahk_values[22] + 1
    elengil3 = kiahk_values[23]

    #مرد الابركسيس 
    if copticdate[2] <= 7 or copticdate[2] > 21:
        mrdelebrksis = kiahk_values[24]
        mrdelebrksis2 = kiahk_values[25]
    else:
        mrdelebrksis = kiahk_values[26]
        mrdelebrksis2 = kiahk_values[27]

    #القرائات
    elebrksis3 =  kiahk_values[28]
    elkatholikon3 = kiahk_values[29]
    elbouls3 = kiahk_values[30]

    #الليلويا فاي بيبي / جي اف ميفي  و  تي شوري / طاي شوري
    if cd.weekday() > 4 :
        shory = kiahk_values[32]
        shory2 = kiahk_values[33]
        alleluia = kiahk_values[34]
        alleluia2 = kiahk_values[35]
    else:
        shory = kiahk_values[36]
        shory2 = kiahk_values[37]
        alleluia = kiahk_values[38]
        alleluia2 = kiahk_values[39]

    if Bishop == True:
        prs3 = relative_path(r"Data\حضور الأسقف.pptx")
        sheet = "في حضور الأسقف"

        bishop_values = find_slide_nums_arrays(excel, sheet, ["صلاة الشكر", "صلاة الشكر", "بهموت غار الصغيرة", "بهموت غار الصغيرة",
                                                             "الهيتنيات", "الهيتنيات", "الاسبسمس", "الاسبسمس"], 
                                                             [1, 2, 1, 2, 1, 2, 1, 2])
        
        bishopDes_values = find_slide_nums_arrays(excel, des_sheet, ["صلاة الشكر", "ني سافيف تيرو", "ني سافيف تيرو", "باهموت غار الصغيرة",
                                                                    "امبين يوت اتطايوت", "امبين يوت اتطايوت", "الهيتنيات", 
                                                                    "اوشية الاباء (ب)", "اوشية الاباء غ", "الاسبسمس الادام",
                                                                    "في حضور الاسقف", "في حضور الاسقف"],
                                                                    [2, 1, 2, 2, 1, 2, 2, 2, 1, 2, 1, 2])

        elshokr1 = bishop_values[0]
        elshokrDes = bishopDes_values[0] - 1

        nysaviv = bishopDes_values[1]
        nysaviv2 = bishopDes_values[2]

        bhmot8ar1 = bishop_values[2]
        bhmot8arDes = bishopDes_values[3] - 1

        embiniot = bishopDes_values[4] 
        embiniot2 = bishopDes_values[5]

        bishopHyten1 = bishop_values[4]
        bishopHytenDes = bishopDes_values[6] + 1

        elaba2basyly = bishopDes_values[7] - 1
        elaba28or8ory = bishopDes_values[8] - 1

        elkhetamBishop = bishopDes_values[10]
        elkhetamBishop2 = bishopDes_values[11]

        if guestBishop > 0:
            elesbsmosDes = bishopDes_values[9]
            elesbsmos = bishop_values[6]
            if guestBishop == 1:
                elshokr2 = bishop_values[1] - 1
                bhmot8ar2 = bishop_values[3] - 1
                bishopHyten2 = bishop_values[5] - 3
                elaba2 = elshokr2
                elaba22 = elshokr2
                elesbsmos2 = bishop_values[7] - 2
            
            elif guestBishop == 2:
                elshokr2 = bishop_values[1]
                bhmot8ar2 = bishop_values[3]
                bishopHyten2 = bishop_values[5]
                elaba2 = elshokr2 - 1
                elaba22 = elshokr2
                elesbsmos2 = bishop_values[7]

            start_positions = [elaba28or8ory, elaba2basyly, elesbsmosDes, elengil3, elmazmor2, elebrksis3, elkatholikon3, elbouls3, bhmot8arDes, bishopHytenDes, embiniot2, elshokrDes]
            start_slides = [elaba2, elaba2, elesbsmos, elengil, elmzmor, elebrksis1, elkatholikon1, elbouls1, bhmot8ar1, bishopHyten1, bhmot8ar1, elshokr1]
            end_slides = [elaba22, elaba22, elesbsmos2, elengil2, elmzmor, elebrksis2, elkatholikon2, elbouls2, bhmot8ar2, bishopHyten2, bhmot8ar2, elshokr2]

        else:
            elshokr2 = bishop_values[1] - 2
            start_positions = [elengil3, elmazmor2, elebrksis3, elkatholikon3, elbouls3, elshokrDes]
            start_slides = [elengil, elmzmor, elebrksis1, elkatholikon1, elbouls1, elshokr1]
            end_slides = [elengil2, elmzmor, elebrksis2, elkatholikon2, elbouls2, elshokr2]
        show_array = [[elmaday7, elmaday7], [season8r8ory, season8r8ory2], [seasonbasyly, seasonbasyly2], 
                      [mrdengilkiahk, mrdengilkiahk2], [mrdelengiltakmla, mrdelengiltakmla2],
                      [mrdelebrksis, mrdelebrksis2], [embiniot, embiniot2], [shory, shory2], [nysaviv, nysaviv2], 
                      [alleluia, alleluia2], [elkhetamBishop, elkhetamBishop2]]

    else:
        start_positions = [elengil3, elmazmor2, elebrksis3, elkatholikon3, elbouls3]
        start_slides = [elengil, elmzmor, elebrksis1, elkatholikon1, elbouls1]
        end_slides = [elengil2, elmzmor, elebrksis2, elkatholikon2, elbouls2]
        show_array = [[elmaday7, elmaday7], [el2smaElmilad, el2smaElmilad2], [season8r8ory, season8r8ory2], 
                      [seasonbasyly, seasonbasyly2], [mrdengilkiahk, mrdengilkiahk2], [mrdelengiltakmla, mrdelengiltakmla2], 
                      [mrdelebrksis, mrdelebrksis2], [shory, shory2], [alleluia, alleluia2]]

    hide_array = [[byoyk, byoyk2], [ekesmarot, ekesmarot2], [el2smaElsanawy, el2smaElsanawy2], [mrdengil, mrdengil2]]

    powerpoint = win32com.client.Dispatch("PowerPoint.Application")
    powerpoint.Visible = True  # Open PowerPoint application
    presentation1 = open_presentation_relative_path(prs1)
    presentation2 = open_presentation_relative_path(prs2)

    if Bishop == True:
        presentation3 = open_presentation_relative_path(prs3)

    khetamSomElmilad = find_slide_index_by_title(presentation1, "صوم الميلاد", elkhetam)
    elmalakGhobrial = find_slide_index_by_label(presentation1, "الملاك غبريال", kiahk_values[31])
    elmalakGhobrial2 = find_slide_index_by_label(presentation1, "الملاك غبريال 2", kiahk_values[31])
    hitniatKiahk = find_slide_index_by_label(presentation1, "كيهك", kiahk_values[31])
    hitniatKiahk2 = find_slide_index_by_label(presentation1, "كيهك 2", kiahk_values[31])
    show_array.extend([[elmalakGhobrial, elmalakGhobrial2], [hitniatKiahk, hitniatKiahk2], [khetamSomElmilad, khetamSomElmilad]])

    match(copticdate[2]):
        case 12:
            elmalakmikhael_values = find_slide_nums_arrays(excel, des_sheet, 
                                                       ["مرد ابركسيس الملاك ميخائيل", "تكملة للملاك ميخائيل 1", 
                                                        "تكملة للملاك ميخائيل 2",
                                                       ], 
                                                       [1, 1, 2])
            hitniat = find_slide_index_by_label(presentation1, "الملاك ميخائيل", kiahk_values[31])
            hitniat2 = find_slide_index_by_label(presentation1, "الملاك ميخائيل 2", kiahk_values[31])
            mrdebrksis = elmalakmikhael_values[0]
            mrdebrksis2 = elmalakmikhael_values[1]
            mrdebrksis3 = elmalakmikhael_values[2]
            show_array.extend([[hitniat, hitniat2], [mrdebrksis, mrdebrksis2], [mrdebrksis3, mrdebrksis3]])

        case 13:
            elmalakrofaeil_values = find_slide_nums_arrays(excel, des_sheet, 
                                                       ["مرد ابركسيس الملاك رافائيل", "مرد ابركسيس الملاك رافائيل"], 
                                                       [1, 2])
            hitniat = find_slide_index_by_label(presentation1, "الملاك رافائيل", kiahk_values[31])
            hitniat2 = find_slide_index_by_label(presentation1, "الملاك رافائيل 2", kiahk_values[31])
            mrdebrksis = elmalakrofaeil_values[0]
            mrdebrksis2 = elmalakrofaeil_values[1]
            show_array.extend([[hitniat, hitniat2], [mrdebrksis, mrdebrksis2]])
    
    show_slides(presentation1, show_array)
    hide_slides(presentation1, hide_array)

    vba_values = find_slide_nums_arrays(excel, des_sheet, ["اسبسمس واطس لكيهك", "اسبسمس واطس لكيهك"], [1, 2])

    elesbasmos = vba_values[0]
    elesbasmos2 = vba_values[1]
    
    endElesbasmoselwats1 = find_slide_index_by_title(presentation1, "الذهاب لختام اسبسمس كيهك الواطس 1", elesbasmos)
    endElesbasmoselwats2 = find_slide_index_by_title(presentation1, "الذهاب لختام اسبسمس كيهك الواطس 2", elesbasmos)

    vba_array = get_slide_ids_by_number_pairs(prs1, [(endElesbasmoselwats1, elesbasmos2), (endElesbasmoselwats2, elesbasmos2)])
    run_vba_with_slide_id(excel, des_sheet, prs1, presentation1, vba_array)
    
    # Initialize variables for current position, slide, and end index
    current_position = start_positions[0]
    current_start_slide = int(start_slides[0])
    current_end_slide = int(end_slides[0])

    # Initialize index for start position, slide, and end slide
    position_index = 1
    slide_index = 1
    end_index = 1

    while current_start_slide <= current_end_slide and slide_index <= presentation1.Slides.Count:
        if (current_position == elengil3 or current_position == elmazmor2 or current_position == elebrksis3 
            or current_position == elkatholikon3 or current_position == elbouls3):
            source_slide = presentation2.Slides(current_end_slide)
            source_slide.Copy()
            presentation1.Slides.Paste(current_position).SlideShowTransition.Hidden = False
            current_end_slide -= 1
            if(current_start_slide > current_end_slide):
                current_position += 1

        elif Bishop == True and (current_position == elaba28or8ory or current_position == elaba2basyly or 
                                 current_position == bhmot8arDes or current_position == bishopHytenDes or
                                  current_position == embiniot2 or current_position == elshokrDes):
            source_slide = presentation3.Slides(current_end_slide)
            source_slide.Copy()
            presentation1.Slides.Paste(current_position)
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

    if copticdate[2] <= 7:
        esbasmosAdam = "اسبسمس ادام كيهك الاسبوع الاول" 
    elif copticdate[2] <= 14:
        esbasmosAdam = "اسبسمس ادام كيهك الاسبوع الثاني"
    elif copticdate[2] > 21:
        esbasmosAdam = "اسبسمس ادام كيهك الاسبوع الرابع"
    else:
        esbasmosAdam = "الاسبسمس الادام"

    move_section_names = [
        "قسمة للأب في صوم و عيد الميلاد - أيها السيد الرب إلهنا",
        esbasmosAdam,
        "ختام الاسبسمس الادام",
        "اسبسمس واطس لكيهك",
        "ختام الأسبسمس الواطس"
    ]

    target_section_names = [
        "القسمة",
        "الاسبسمس الادام",
        esbasmosAdam,
        "أسبسمس واطس",
        "اسبسمس واطس لكيهك"
    ]

    # Call the function once for all moves
    move_sections(presentation1, move_section_names, target_section_names)

    presentation2.Close()
    if Bishop:
        presentation3.Close()

def odas29thOfMonth (copticdate, Bishop=False, guestBishop=0):
    from copticDate import CopticCalendar
    cd = CopticCalendar().coptic_to_gregorian(copticdate)
    prs1 = relative_path(r"قداس.pptx") 
    excel = relative_path(r"بيانات القداسات.xlsx")
    excel2 = relative_path(r"Tables.xlsx")
    des_sheet ="القداس"
    # sunday(prs1)
    if cd.weekday() == 6 :
        if copticdate[1] == 1:
            prs2 = relative_path(r"Data\القطمارس\الاحاد\القطمارس السنوي احاد (القداس).pptx")
            katamars_sheet = "قطمارس الاحاد للقداس"
            km = 1
            kd = 4
        else:
            prs2 = relative_path(r"Data\القطمارس\الايام\القطمارس السنوي ايام (القداس).pptx")
            katamars_sheet = "القطمارس السنوي القداس"
            km = 7
            kd = 29
    else: 
        prs2 = relative_path(r"Data\القطمارس\الايام\القطمارس السنوي ايام (القداس).pptx")
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

    '''
        arr = ["تكملة على حسب المناسبة", "مدائح ال29 من الشهر", "بي اويك", "بي اويك", "اك اسماروؤت", "اك اسماروؤت",
        "مزمور التوزيع", "مزمور التوزيع", "مرد توزيع البشارة", "مرد توزيع الميلاد", "مرد توزيع القيامة",
        "قسمة - أيها السيد الرب إلهنا", "قسمة - أيها السيد الرب إلهنا",
        "تذكار البشارة والميلاد والقيامة - نسبح ونمجد إله الآلهة ورب الأرباب", 
        "تذكار البشارة والميلاد والقيامة - نسبح ونمجد إله الآلهة ورب الأرباب",
        "مرد انجيل ال29 من الشهر", "مرد انجيل ال29 من الشهر", "مرد الانجيل", "مرد الانجيل", 
        "فاي اريه بي اوو", "فاي اريه بي اوو", "المزمور و الانجيل", "المزمور و الانجيل",
        "مرد مزمور ال29 من الشهر", "مرد مزمور ال29 من الشهر", "مرد المزمور", 
        "مرد ابركسيس سنوي", "مرد ابركسيس سنوي", "مرد ابركسيس البشارة", "مرد ابركسيس البشارة", 
        "مرد ابركسيس الميلاد", "مرد ابركسيس الميلاد", "مرد ابركسيس القيامة", "مرد ابركسيس القيامة",
        "الابركسيس", "الكاثوليكون", "البولس عربي", "الهيتنيات", 
        "طاي شوري", "طاي شوري", "الليلويا فاي بيبي", "الليلويا فاي بيبي"]

    '''

    twentyNine_values = find_slide_nums_arrays(excel, des_sheet, ["تكملة على حسب المناسبة", "مدائح ال29 من الشهر", "بي اويك", "بي اويك", "اك اسماروؤت", "اك اسماروؤت",
                                                                     "مزمور التوزيع", "مزمور التوزيع", "مرد توزيع البشارة", "مرد توزيع الميلاد", "مرد توزيع القيامة",
                                                                     "قسمة - أيها السيد الرب إلهنا", "قسمة - أيها السيد الرب إلهنا",
                                                                     "تذكار البشارة والميلاد والقيامة - نسبح ونمجد إله الآلهة ورب الأرباب", 
                                                                     "تذكار البشارة والميلاد والقيامة - نسبح ونمجد إله الآلهة ورب الأرباب",
                                                                     "مرد انجيل ال29 من الشهر", "مرد انجيل ال29 من الشهر", "مرد الانجيل", "مرد الانجيل", 
                                                                     "فاي اريه بي اوو", "فاي اريه بي اوو", "المزمور و الانجيل", "المزمور و الانجيل",
                                                                     "مرد مزمور ال29 من الشهر", "مرد مزمور ال29 من الشهر", "مرد المزمور", 
                                                                     "مرد ابركسيس سنوي", "مرد ابركسيس سنوي", "مرد ابركسيس البشارة", "مرد ابركسيس البشارة", 
                                                                     "مرد ابركسيس الميلاد", "مرد ابركسيس الميلاد", "مرد ابركسيس القيامة", "مرد ابركسيس القيامة",
                                                                     "الابركسيس", "الكاثوليكون", "البولس عربي", "الهيتنيات", 
                                                                     "طاي شوري", "طاي شوري", "الليلويا فاي بيبي", "الليلويا فاي بيبي"], 
                                                                    [2, 1, 1, 2, 1, 2, 1, 2, 1, 1, 1, 1, 2, 1, 2, 1, 2, 1, 2, 1, 2, 1, 2, 1, 2, 1, 1, 2, 1, 2, 1, 2, 1, 2, 2, 2, 2, 1, 1, 2, 1, 2])
    
    print(twentyNine_values)

    #الختام
    elkhetam = twentyNine_values[0]

    #التوزيع
    elmaday7 = twentyNine_values[1]
    byoyk = twentyNine_values[2]
    byoyk2 = twentyNine_values[3]
    ekesmarot = twentyNine_values[4]
    ekesmarot2 = twentyNine_values[5]
    
    #مزمور التوزيع
    fs = twentyNine_values[6] + 1 
    ls = twentyNine_values[7] - 1
    sn1 = twentyNine_values[8]
    sn2 = twentyNine_values[9]
    sn3 = twentyNine_values[10]

    #القسمة
    el2smaElsanawy = twentyNine_values[11]
    el2smaElsanawy2 = twentyNine_values[12]
    el2sma29 = twentyNine_values[13]
    el2sma292 = twentyNine_values[14]

    #الاواشي
    if copticdate == None:
        copticdate = CopticCalendar().gregorian_to_coptic()
        season = CopticCalendar().get_coptic_date_range(copticdate)
    else:
        season = CopticCalendar().get_coptic_date_range(copticdate)
    if season == "Air":
        elawashy_values = find_slide_nums_arrays_v2(excel, des_sheet, ['{BC7E3DCD-6AA8-44CC-B8AF-BC3E2BC71B5A}', '{BC7E3DCD-6AA8-44CC-B8AF-BC3E2BC71B5A}', '{A20DA654-32F7-4B4C-96CB-C76232EB96E8}', '{A20DA654-32F7-4B4C-96CB-C76232EB96E8}'], 2, [1, 2, 1, 2])
        seasonbasyly = elawashy_values[0]
        seasonbasyly2 = elawashy_values[1]
        season8r8ory = elawashy_values[2]
        season8r8ory2 = elawashy_values[3]
    elif season == 'Tree':
        elawashy_values = find_slide_nums_arrays_v2(excel, des_sheet, ['{F94B3D1F-649D-4839-BD2E-19439E173129}', '{F94B3D1F-649D-4839-BD2E-19439E173129}', '{3D4C118C-E6FF-4DF8-8E8F-B0CDF0FDBA54}', '{3D4C118C-E6FF-4DF8-8E8F-B0CDF0FDBA54}'], 2, [1, 2, 1, 2])
        seasonbasyly = elawashy_values[0]
        seasonbasyly2 = elawashy_values[1]
        season8r8ory = elawashy_values[2]
        season8r8ory2 = elawashy_values[3]
    else:
        elawashy_values = find_slide_nums_arrays_v2(excel, des_sheet, ['{C7FC170A-D45F-4D4E-BD01-F17CADBFB65C}', '{C7FC170A-D45F-4D4E-BD01-F17CADBFB65C}', '{5DD6BABA-9FE4-4D33-9F90-0C865CB95EE4}', '{5DD6BABA-9FE4-4D33-9F90-0C865CB95EE4}'], 2, [1, 2, 1, 2])
        seasonbasyly = elawashy_values[0]
        seasonbasyly2 = elawashy_values[1]
        season8r8ory = elawashy_values[2]
        season8r8ory2 = elawashy_values[3]

    #مرد الإنجيل
    mrdengil29 = twentyNine_values[15]
    mrdengil292 = twentyNine_values[16]
    mrdengil = twentyNine_values[17]
    mrdengil2 = twentyNine_values[18]
    fayerby = twentyNine_values[19]
    fayerby2 = twentyNine_values[20]

    #المزمور و الانجيل
    elengil3 = twentyNine_values[22]
    elmazmor2 = twentyNine_values[21] + 2

    #مرد المزمور
    mrdElmzmor29 = twentyNine_values[23]
    mrdElmzmor292 = twentyNine_values[24]
    mrdElmzmor = twentyNine_values[25]

    #مرد الابركسيس
    mrdelebrksis = twentyNine_values[26]
    mrdelebrksis2 = twentyNine_values[27]
    mrdelebrksisElbshara = twentyNine_values[28]
    mrdelebrksisElbshara2 = twentyNine_values[29]
    mrdelebrksisElmilad = twentyNine_values[30]
    mrdelebrksisElmilad2 = twentyNine_values[31]
    mrdelebrksisEl2yama = twentyNine_values[32]
    mrdelebrksisEl2yama2 = twentyNine_values[33]

    #القرائات
    elebrksis3 =  twentyNine_values[34]
    elkatholikon3 = twentyNine_values[35]
    elbouls3 = twentyNine_values[36]

    #ما قبل القرائات
    elhitnyat = twentyNine_values[37]
    shory = twentyNine_values[38]
    shory2 = twentyNine_values[39]
    allylouya = twentyNine_values[40]
    allylouya2 = twentyNine_values[41]
    
    if Bishop == True:
        prs3 = relative_path(r"Data\حضور الأسقف.pptx")
        sheet = "في حضور الأسقف"

        bishop_values = find_slide_nums_arrays_v2(excel, sheet, ['{D1378DB5-29D1-4800-9D96-10F2535EEB57}', '{D1378DB5-29D1-4800-9D96-10F2535EEB57}', '{7C84083F-E6D3-4669-9130-AC7E8D935A98}', '{7C84083F-E6D3-4669-9130-AC7E8D935A98}', '{97203297-EECB-4D41-B2E3-AD9A4863847E}', '{97203297-EECB-4D41-B2E3-AD9A4863847E}', '{6851F163-CBEF-4014-A853-CE100557BA6A}', '{6851F163-CBEF-4014-A853-CE100557BA6A}'], 
                                                             2, [1, 2, 1, 2, 1, 2, 1, 2])
        
        bishopDes_values = find_slide_nums_arrays_v2(excel, des_sheet, ['{A9183893-7B7E-459F-8547-F7A8F7D2D521}', '{A9183893-7B7E-459F-8547-F7A8F7D2D521}', '{474487EB-4554-4A30-B351-6EF762D2F2D6}', '{74A33555-8E08-47DF-B3CD-A1B4C7AF2B4E}', '{22F83DFC-792B-4148-8AED-E77703B6E7BB}', '{21891CBB-A1EC-4974-B0B6-F74A4B502BC2}', '{79CED7F3-DA1D-467F-AA09-4187C8DE51E8}', '{E2968C91-5339-499C-9812-DECCCF58A2CD}', '{E2968C91-5339-499C-9812-DECCCF58A2CD}', '{670DAA94-A6C9-4CCD-B4E2-958C71CD3E44}', '{670DAA94-A6C9-4CCD-B4E2-958C71CD3E44}', '{40B60A3B-1E1C-423C-A2AC-B3081CEEE693}'],
                                                                    2, [2, 1, 2, 2, 1, 2, 2, 2, 1, 2, 1, 2])

        elshokr1 = bishop_values[0]
        elshokrDes = bishopDes_values[0] - 1

        nysaviv = bishopDes_values[1]
        nysaviv2 = bishopDes_values[2]

        bhmot8ar1 = bishop_values[2]
        bhmot8arDes = bishopDes_values[3] - 1

        embiniot = bishopDes_values[4] 
        embiniot2 = bishopDes_values[5]

        bishopHyten1 = bishop_values[4]
        bishopHytenDes = bishopDes_values[6] + 1

        elaba2basyly = bishopDes_values[7] - 1
        elaba28or8ory = bishopDes_values[8] - 1

        elkhetamBishop = bishopDes_values[10]
        elkhetamBishop2 = bishopDes_values[11]

        if guestBishop > 0:
            elesbsmosDes = bishopDes_values[9]
            elesbsmos = bishop_values[6]
            if guestBishop == 1:
                elshokr2 = bishop_values[1] - 1
                bhmot8ar2 = bishop_values[3] - 1
                bishopHyten2 = bishop_values[5] - 3
                elaba2 = elshokr2
                elaba22 = elshokr2
                elesbsmos2 = bishop_values[7] - 2
            
            elif guestBishop == 2:
                elshokr2 = bishop_values[1]
                bhmot8ar2 = bishop_values[3]
                bishopHyten2 = bishop_values[5]
                elaba2 = elshokr2 - 1
                elaba22 = elshokr2
                elesbsmos2 = bishop_values[7]

            start_positions = [fs, elaba28or8ory, elaba2basyly, elesbsmosDes, elengil3, elmazmor2, elebrksis3, elkatholikon3, elbouls3, bhmot8arDes, bishopHytenDes, embiniot2, elshokrDes]
            start_slides = [sn3, elaba2, elaba2, elesbsmos, elengil, elmzmor, elebrksis1, elkatholikon1, elbouls1, bhmot8ar1, bishopHyten1, bhmot8ar1, elshokr1]
            end_slides = [sn3, elaba22, elaba22, elesbsmos2, elengil2, elmzmor, elebrksis2, elkatholikon2, elbouls2, bhmot8ar2, bishopHyten2, bhmot8ar2, elshokr2]

        else:
            elshokr2 = bishop_values[1] - 2
            start_positions = [fs, elengil3, elmazmor2, elebrksis3, elkatholikon3, elbouls3, elshokrDes]
            start_slides = [sn3, elengil, elmzmor, elebrksis1, elkatholikon1, elbouls1, elshokr1]
            end_slides = [sn3, elengil2, elmzmor, elebrksis2, elkatholikon2, elbouls2, elshokr2]
        show_array = [[elmaday7, elmaday7], [el2sma29, el2sma292], [season8r8ory, season8r8ory2], [seasonbasyly, seasonbasyly2], 
                      [mrdengil29, mrdengil292], [fayerby, fayerby2], [mrdElmzmor29, mrdElmzmor292], 
                      [mrdelebrksisElbshara, mrdelebrksisElbshara2], [mrdelebrksisElmilad, mrdelebrksisElmilad2], 
                      [mrdelebrksisEl2yama, mrdelebrksisEl2yama2], [embiniot, embiniot2], [shory, shory2], 
                      [nysaviv, nysaviv2], [allylouya, allylouya2], [elkhetamBishop, elkhetamBishop2]]

    else:
        start_positions = [fs, elengil3, elmazmor2, elebrksis3, elkatholikon3, elbouls3]
        start_slides = [sn3, elengil, elmzmor, elebrksis1, elkatholikon1, elbouls1]
        end_slides = [sn3, elengil2, elmzmor, elebrksis2, elkatholikon2, elbouls2]
        show_array = [[elmaday7, elmaday7], [el2sma29, el2sma292], [season8r8ory, season8r8ory2], [seasonbasyly, seasonbasyly2],
                      [mrdengil29, mrdengil292], [fayerby, fayerby2], [mrdElmzmor29, mrdElmzmor292], 
                      [mrdelebrksisElbshara, mrdelebrksisElbshara2], [mrdelebrksisElmilad, mrdelebrksisElmilad2], 
                      [mrdelebrksisEl2yama, mrdelebrksisEl2yama2], [shory, shory2], [allylouya, allylouya2]]

    hide_array = [[byoyk, byoyk2], [ekesmarot, ekesmarot2], [el2smaElsanawy, el2smaElsanawy2], [mrdengil, mrdengil2],
                  [mrdElmzmor, mrdElmzmor], [mrdelebrksis, mrdelebrksis2]]

    powerpoint = win32com.client.Dispatch("PowerPoint.Application")
    powerpoint.Visible = True  # Open PowerPoint application
    presentation1 = open_presentation_relative_path(prs1)
    presentation2 = open_presentation_relative_path(prs2)
    
    if Bishop == True:
        presentation3 = open_presentation_relative_path(prs3)

    khetam29 = find_slide_index_by_title(presentation1, "تذكار الاعياد السيدية", elkhetam, "up")
    khetam292 = find_slide_index_by_title(presentation1, "تذكار الاعياد السيدية 2", elkhetam, "up")

    hitniat_values = find_slide_indices_by_ordered_labels(presentation1, ["الملاك ميخائيل الخماسين", "الملاك ميخائيل الخماسين 2",
    "الملاك غبريال", "الملاك غبريال 2", "الميلاد", "الميلاد 2", "القيامة", "القيامة 2"], elhitnyat)
    hitniatElmalakMikhael2yama = hitniat_values[0]
    hitniatElmalakMikhael2yama2 = hitniat_values[1]
    hitniatElmalakGhobrial = hitniat_values[2]
    hitniatElmalakGhobiral2 = hitniat_values[3]
    hitniatElmilad = hitniat_values[4]
    hitniatElmilad2 = hitniat_values[5]
    hitniatEl2yama = hitniat_values[6]
    hitniatEl2yama2 = hitniat_values[7]
    show_array.extend([[khetam29, khetam292], [hitniatElmalakMikhael2yama, hitniatElmalakMikhael2yama2], [hitniatElmalakGhobrial, hitniatElmalakGhobiral2],
                       [hitniatElmilad, hitniatElmilad2], [hitniatEl2yama, hitniatEl2yama2]])

    show_slides(presentation1, show_array)
    hide_slides(presentation1, hide_array)
    run_vba_with_slide_id(excel, des_sheet, prs1, presentation1)
    
    # Initialize variables for current position, slide, and end index
    current_position = start_positions[0]
    current_start_slide = int(start_slides[0])
    current_end_slide = int(end_slides[0])

    # Initialize index for start position, slide, and end slide
    position_index = 1
    slide_index = 1
    end_index = 1

    while current_start_slide <= current_end_slide and slide_index <= presentation1.Slides.Count:
        if current_position == fs:
            slide_index1 = ls
            while slide_index1 >= current_position:
                source_slide = presentation1.Slides(current_start_slide)
                source_slide.Copy()
                new_slide = presentation1.Slides.Paste(slide_index1).SlideShowTransition.Hidden = False
                slide_index1 -= 1

                if current_start_slide == sn3:
                    current_start_slide = sn2
                elif current_start_slide == sn2:
                    current_start_slide = sn1
                else:
                    current_start_slide = sn3 
                
            current_start_slide += 1

        elif (current_position == elengil3 or current_position == elmazmor2 or current_position == elebrksis3 
            or current_position == elkatholikon3 or current_position == elbouls3):
            source_slide = presentation2.Slides(current_end_slide)
            source_slide.Copy()
            new_slide = presentation1.Slides.Paste(current_position).SlideShowTransition.Hidden = False
            current_end_slide -= 1
            if(current_start_slide > current_end_slide):
                current_position += 1

        elif Bishop == True and (current_position == elaba28or8ory or current_position == elaba2basyly or 
                                 current_position == bhmot8arDes or current_position == bishopHytenDes or 
                                 current_position == embiniot2 or current_position == elshokrDes):
            source_slide = presentation3.Slides(current_end_slide)
            source_slide.Copy()
            new_slide = presentation1.Slides.Paste(current_position)
            new_slide.SlideShowTransition.Hidden = False
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

    move_section_names = ['{4E0564C4-BDF3-47D3-8EAF-B0110F0233DA}']
    target_section_names = ['{01F61BFD-210C-4F99-A7C7-7308CFAA93F4}']

    # Call the function once for all moves
    move_sections_v2(presentation1, move_section_names, target_section_names)

    presentation2.Close()
    if Bishop:
        presentation3.Close()


"_____________________________________NEW_CODE_DESIGN_____________________________________"

def odasElsalyb(copticdate, Bishop=False, guestBishop=0):
    from copticDate import CopticCalendar
    cd = CopticCalendar().coptic_to_gregorian(copticdate)
    prs = relative_path(r"قداس.pptx")  # Using the relative path
    excel = relative_path(r"بيانات القداسات.xlsx")
    excel2 = relative_path(r"Tables.xlsx")
    des_sheet ="القداس"
    SlaybText = ["لأنك صُلِبتَ", "auask", "اف اشك"]
    replacefile(prs, relative_path(r"Data\CopyData\قداس.pptx"))

    prs2 = relative_path(r"Data\القطمارس\الايام\القطمارس السنوي ايام (القداس).pptx")
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

    # salyb_show_full_sections = ["الليلويا فاي بيبي", "تي شوري", "فاي ايطاف اينف", "ايطاف اني اسخاي", 
    #                             "مرد ابركسيس الصليب", "محير عيد الصليب", "تكملة مشتركة للمحير", "مرد مزمور الصليب", 
    #                             "مرد انجيل الصليب", "قسمة عيد الصليب (ايها المسيح الهنا الابن الكلمة)"]
    
    salyb_show_full_sections = ['{072F3D96-A6C8-405F-9A23-7CCA1B2F13FF}', '{43B74D4E-929B-4654-8DB7-4675EFF27370}','{44016DA1-D105-4BF5-9C30-AC5A359693BF}', '{60A6EF5E-A7D1-4E39-AAE6-A0EA92726711}', '{7C01C8C3-BA04-422E-80A4-D88848FCDFDB}', '{D1DCD63D-C047-4475-92E2-CA4D0B61C4A7}', '{DEDC0CCA-3854-4E18-8CB2-5D6FEC5BABCC}', '{0BC1F8D8-BE35-4C07-A134-EAB9CF63D177}', '{EDE09087-B069-49CA-8211-757926594D3F}', '{3DF2ADE7-6EAF-4FD9-8250-7048D4114339}']
    
    # salyb_hide_full_sections = ["اكسيا", "اجيوس الميلاد", "اجيوس الصعود", "مرد المزمور", "مرد الانجيل", "ربع للعذراء", 
    #                             "قسمة القداس الباسيلي (أيها السيد الرب إلهنا)", "اك اسماروؤت", "بي اويك"]
    
    salyb_hide_full_sections = ['{A8A52E1F-44DD-45E1-A737-4E13E15D5F1F}', '{76175201-2586-46AD-8E9F-C69E29FE2620}', '{CD4B95FF-0E0E-42D3-8DC4-224C3DD732F7}', '{F13A48F2-238D-4617-B84E-9B0A694D9A18}', '{06D592C8-05BF-4B7C-86F7-FDAB3FAB5FB1}', '{ECE652ED-1345-4C6D-B92D-5996CFA27AEE}', '{681FF6A7-4230-4171-8F41-83FD64E8C960}', '{507EFD97-98F8-4376-848B-20D72E16D2C1}', '{B9A30F5E-0C89-471B-A99A-23DBE7F58504}']
    
    salyb_show_values = []
    salyb_hide_values = []

    # salyb_values = ["تكملة على حسب المناسبة", "مزمور التوزيع", "مزمور التوزيع", "مرد توزيع الصليب",
    #                 "الانجيل", "المزمور", "اجيوس الصلب", "اجيوس الصلب", "الابركسيس", "الكاثوليكون", "البولس عربي"]
    salyb_values = find_slide_nums_arrays_v2(excel, des_sheet, 
                    ['{A18EDC94-F257-4FAC-99C7-0A8EA70F0FAF}', '{C29E5A83-A98B-4077-8194-99A6D803EF53}', '{C29E5A83-A98B-4077-8194-99A6D803EF53}', '{E4E79CE2-AD28-44FB-B8B3-6B59A5D64B62}', '{C7D4A109-F792-4661-BAD0-075FD1A1909F}', '{B74DBB8C-2B2D-46E4-9508-DA46008D19A4}', '{31685B5B-48C4-437E-858C-CF8D225C0C26}', '{31685B5B-48C4-437E-858C-CF8D225C0C26}', '{E234C6C7-3837-4CE4-A541-CDC9627AAAC2}', '{6D4B3F52-63BF-435F-BF0C-C9D41120C2A3}', '{D88055F5-EAA0-4C8E-8249-C364A572BF7B}'], 
                    2, [1, 1, 2, 1, 2, 2, 1, 2, 2, 2, 2])
    
    #الختام
    elkhetam = salyb_values[0]
    #التوزيع
    mazmorELtawzy3 = salyb_values[1] + 1
    mazmorELtawzy32 = salyb_values[2] - 1
    mrdMazmorEltawzy3 = salyb_values[3]
    #الاواشي
    AwashySeason = CopticCalendar().get_coptic_date_range(copticdate)
    match AwashySeason:
        case "Air": salyb_show_full_sections.extend(['{BC7E3DCD-6AA8-44CC-B8AF-BC3E2BC71B5A}', '{A20DA654-32F7-4B4C-96CB-C76232EB96E8}']),
        case "Tree": salyb_show_full_sections.extend(['{F94B3D1F-649D-4839-BD2E-19439E173129}', '{5DD6BABA-9FE4-4D33-9F90-0C865CB95EE4}'])
        case "Water": salyb_show_full_sections.extend(['{C7FC170A-D45F-4D4E-BD01-F17CADBFB65C}', '{3D4C118C-E6FF-4DF8-8E8F-B0CDF0FDBA54}'])
    #اجيوس
    agiosElsalb = salyb_values[6]
    agiosElsalb2 = salyb_values[7]
    #القرائات
    elengil3 = salyb_values[4]
    elmazmor3 = salyb_values[5]
    elebrksis3 = salyb_values[8]
    elkatholikon3 = salyb_values[9]
    elbouls3 = salyb_values[10]

    if Bishop == True:
        prs3 = relative_path(r"Data\حضور الأسقف.pptx")
        sheet = "في حضور الأسقف"

        # bishop_show_values = ['في حضور الاسقف', 'المزمور', 'مارو اتشاسف', 'امبين يوت اتطايوت', 
        #                       'ني سافيف تيرو', 'تكملة في حضور الاسقف', 'لحن اك اسمارؤوت']

        #bishop_hide_values = ['سوتيس امين']

        bishop_show_values = ['{2BCF4F8C-25F0-43C5-B224-6528B2EA3F2F}', '{F76B0D75-0474-45B5-B79F-7416F354543A}',
                              '{E2968C91-5339-499C-9812-DECCCF58A2CD}', '{62A12AF8-CB6D-4CC5-9DB0-B73A7C24E2AD}', 
                              '{B74DBB8C-2B2D-46E4-9508-DA46008D19A4}', '{A9183893-7B7E-459F-8547-F7A8F7D2D521}', 
                              '{670DAA94-A6C9-4CCD-B4E2-958C71CD3E44}']
        
        bishop_hide_values = ['{4D2B15D5-C978-467C-9D6C-726FE25128B8}']
        
        salyb_show_full_sections.extend(bishop_show_values)
        salyb_hide_full_sections.extend(bishop_hide_values)

        if guestBishop > 0:
            # bishop_values = ["صلاة الشكر", "صلاة الشكر", "طوبه هينا الكبيرة", "طوبه هينا الكبيرة", 
            #                  "نيم بينيوت", "نيم بينيوت", "الهيتنيات", "الهيتنيات",
            #                  "اي اغابي", "اي اغابي", "مرد الكاثوليكون", "مرد الكاثوليكون",
            #                  "الاسبسمس", "الاسبسمس"]

            # bishopDes_values = ["تكملة في حضور الاسقف", "طوبه هينا الكبيرة",
            #                     "امبين يوت اتطايوت", "تكملة الهيتنيات", "بى اهموت غار الصغيرة",
            #                     "تو ماكريو", "اي اغابي", "ابيت جيك ايفول", "مارو اتشاسف",
            #                     "الاسبسمس الادام السنوي+الختام", "ختام الاسبسمس الادام",
            #                     "اوشية الاباء (ب)", "اوشية الاباء غ"]
            
            bishop_values = find_slide_nums_arrays_v2(excel, sheet, 
                            ['{6851F163-CBEF-4014-A853-CE100557BA6A}', '{6851F163-CBEF-4014-A853-CE100557BA6A}', '{B084BC40-61E1-4477-98DA-15CFB06AEE91}', '{B084BC40-61E1-4477-98DA-15CFB06AEE91}', '{97203297-EECB-4D41-B2E3-AD9A4863847E}', '{97203297-EECB-4D41-B2E3-AD9A4863847E}', '{7C84083F-E6D3-4669-9130-AC7E8D935A98}', '{7C84083F-E6D3-4669-9130-AC7E8D935A98}', '{D0234C99-69FE-407A-82A9-D7A676919E93}', '{D0234C99-69FE-407A-82A9-D7A676919E93}', '{0B519645-F935-43AE-A9CE-6E2FC03833BB}', '{0B519645-F935-43AE-A9CE-6E2FC03833BB}', '{D1378DB5-29D1-4800-9D96-10F2535EEB57}', '{D1378DB5-29D1-4800-9D96-10F2535EEB57}'], 
                            2, [1, 2, 1, 2, 1, 2, 1, 2, 1, 2, 1, 2, 1, 2])

            bishopDes_values = find_slide_nums_arrays_v2(excel, des_sheet, 
                               ['{F76B0D75-0474-45B5-B79F-7416F354543A}', '{8DD21CDE-CB6B-4D5B-B995-D2747AB69ED1}', '{E2968C91-5339-499C-9812-DECCCF58A2CD}', '{646A8184-7F05-453A-A2F1-EB9A77D7F0EE}', '{21891CBB-A1EC-4974-B0B6-F74A4B502BC2}', '{12B7D244-BF4C-401B-A65A-D1621D7DD953}', '{F69B50D8-FB5E-4E8C-AE5C-6DCB4790AFAF}', '{38267404-7625-47AA-B0C8-31BCA5D0435D}', '{62A12AF8-CB6D-4CC5-9DB0-B73A7C24E2AD}', '{0C7A7725-643D-4F7E-A2F0-0C8A36C2A594}', '{1CDFD5FF-8DF5-48CF-A773-F5EEBD56469C}', '{74A33555-8E08-47DF-B3CD-A1B4C7AF2B4E}', '{474487EB-4554-4A30-B351-6EF762D2F2D6}'],                  
                               2, [2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2])

            elshokr = bishopDes_values[0]
            elshokr1 = bishop_values[0]
            elshokr2 = bishop_values[1]

            tobhyna = bishopDes_values[1]
            tobhyna1 = bishop_values[2]
            tobhyna2 = bishop_values[3]

            embiniot = bishopDes_values[2]
            embiniot1 = bishop_values[4]
            embiniot2 = bishop_values[5]

            hytynyat = bishopDes_values[3]
            hytynyat1 = bishop_values[6]
            hytynyat2 = bishop_values[7]

            byhmot8ar = bishopDes_values[4]-1
            byhmot8ar1 = bishop_values[4]
            byhmot8ar2 = bishop_values[5]

            tomakario = bishopDes_values[5] - 1
            tomakario1 = bishop_values[4]
            tomakario2 = bishop_values[5]

            eya8aby = bishopDes_values[6] - 1
            eya8aby1 = bishop_values[8]
            eya8aby2 = bishop_values[9]

            mrdElkatholikon = bishopDes_values[7]
            mrdElkatholikon1 = bishop_values[10]
            mrdElkatholikon2 = bishop_values[11]

            maro = bishopDes_values[8]
            maro1 = bishop_values[4]
            maro2 = bishop_values[5]

            esbasmos1 = bishopDes_values[9]
            esbasmos11 = bishop_values[12]
            esbasmos12 = bishop_values[13]

            esbasmos2 = bishopDes_values[10]
            esbasmos21 = bishop_values[12]
            esbasmos22 = bishop_values[13]

            elaba2basyly = bishopDes_values[11] - 1
            elaba2basyly1 = bishop_values[0]
            elaba2basyly2 = bishop_values[1]

            elaba28yry8ory = bishopDes_values[12] - 1
            elaba28yry8ory1 = bishop_values[0]
            elaba28yry8ory2 = bishop_values[1]

            if guestBishop < 2:
                elshokr2 = elshokr2-1
                tobhyna2 = tobhyna2-2
                embiniot2 = embiniot2-1
                hytynyat2 = hytynyat2-3
                byhmot8ar2 = byhmot8ar2-1
                tomakario2 = tomakario2-1
                eya8aby2 = eya8aby2-2
                mrdElkatholikon2 = mrdElkatholikon2-1
                maro2 = maro2-1
                esbasmos12 = esbasmos12-2
                esbasmos22 = esbasmos22-2
                elaba2basyly2 = elaba2basyly2-1
                elaba28yry8ory2 = elaba28yry8ory2-1
        
            start_positions = [elaba28yry8ory, elaba2basyly, esbasmos2, esbasmos1, elengil3, elmazmor3, maro, elebrksis3, elkatholikon3, mrdElkatholikon, elbouls3, eya8aby, tomakario, byhmot8ar, hytynyat, embiniot, tobhyna, elshokr]
            start_slides = [elaba28yry8ory1, elaba2basyly1, esbasmos21, esbasmos11, elengil, elmzmor, maro1, elebrksis1, elkatholikon1, mrdElkatholikon1, elbouls1, eya8aby1, tomakario1, byhmot8ar1, hytynyat1, embiniot1, tobhyna1, elshokr1]
            end_slides = [elaba28yry8ory2, elaba2basyly2, esbasmos22, esbasmos12, elengil2, elmzmor, maro2, elebrksis2, elkatholikon2, mrdElkatholikon2, elbouls2, eya8aby2, tomakario2, byhmot8ar2, hytynyat2, embiniot2, tobhyna2, elshokr2]

    if guestBishop == 0:
        start_positions = [mazmorELtawzy3, elengil3, elmazmor3, elebrksis3, elkatholikon3, elbouls3]
        start_slides = [mrdMazmorEltawzy3, elengil, elmzmor, elebrksis1, elkatholikon1, elbouls1]
        end_slides = [mrdMazmorEltawzy3, elengil2, elmzmor, elebrksis2, elkatholikon2, elbouls2]

    show_hide_insertImage_replaceText(prs, excel, des_sheet, salyb_show_full_sections, salyb_hide_full_sections, newText=SlaybText)

    powerpoint = win32com.client.Dispatch("PowerPoint.Application")
    powerpoint.Visible = True  # Open PowerPoint application
    presentation1 = open_presentation_relative_path(prs)
    presentation2 = open_presentation_relative_path(prs2)
    
    if guestBishop > 0:
        presentation3 = open_presentation_relative_path(prs3)

    khetamElsalyb = find_slide_index_by_title(presentation1, "الصليب", elkhetam)
    show_slides(presentation1, [[khetamElsalyb, khetamElsalyb]])

    run_vba_with_slide_id(excel, des_sheet, prs, presentation1)
    
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
                new_slide = presentation1.Slides.Paste(slide_index1).SlideShowTransition.Hidden = False
                slide_index1 -= 1
            current_start_slide += 1

        elif (current_position == elengil3 or current_position == elmazmor3 or current_position == elebrksis3 
            or current_position == elkatholikon3 or current_position == elbouls3):
            source_slide = presentation2.Slides(current_end_slide)
            source_slide.Copy()
            new_slide = presentation1.Slides.Paste(current_position)
            new_slide.SlideShowTransition.Hidden = False
            current_end_slide -= 1
            # Add slide duplication after elmzmor1 is processed
            if guestBishop == 0 and current_position == elmazmor3:
                # Perform the duplication twice
                for _ in range(2):  # Loop to duplicate the section twice
                    # Copy the first slide (agiosElsalyb) and paste it after the second slide (agiosElsalyb2)
                    agiosElmilad_slide = presentation1.Slides(agiosElsalb)
                    agiosElmilad_slide.Copy()
                    presentation1.Slides.Paste(agiosElsalb2 + 1)

                    # Copy the second slide (agiosElsalyb2) and paste it after the newly copied first slide
                    agiosElmilad2_slide = presentation1.Slides(agiosElsalb2)
                    agiosElmilad2_slide.Copy()
                    presentation1.Slides.Paste(agiosElsalb2 + 2)  # Adjust to account for new first slide
            if current_start_slide > current_end_slide:
                current_position += 1

        elif Bishop and current_position in {elaba28yry8ory, elaba2basyly, byhmot8ar, hytynyat, embiniot, 
                                             elshokr, tomakario, eya8aby, esbasmos1, esbasmos2, maro, 
                                             mrdElkatholikon, tobhyna}:
            source_slide = presentation3.Slides(current_end_slide)
            source_slide.Copy()
            if current_position in {elaba28yry8ory, elaba2basyly, byhmot8ar, hytynyat, embiniot, elshokr, maro}:
                presentation1.Slides.Paste(current_position).SlideShowTransition.Hidden = False
            else:
                presentation1.Slides.Paste(current_position).SlideShowTransition.Hidden = True
            current_end_slide -= 1
            if (current_position == maro) and (current_start_slide > current_end_slide):
                # Perform the duplication twice
                for _ in range(2):  # Loop to duplicate the section twice
                    # Copy the first slide (agiosElsalyb) and paste it after the second slide (agiosElsalyb2)
                    agiosElmilad_slide = presentation1.Slides(agiosElsalb)
                    agiosElmilad_slide.Copy()
                    presentation1.Slides.Paste(agiosElsalb2 + 1)

                    # Copy the second slide (agiosElsalyb2) and paste it after the newly copied first slide
                    agiosElmilad2_slide = presentation1.Slides(agiosElsalb2)
                    agiosElmilad2_slide.Copy()
                    presentation1.Slides.Paste(agiosElsalb2 + 2)  # Adjust to account for new first slide
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

    move_section_names = [
        "{56962C9F-AA71-4EBE-BFCC-82940ED9C771}",
        "{1CDFD5FF-8DF5-48CF-A773-F5EEBD56469C}",
        "{02523737-EC78-4811-84BB-436795EE788F}", 
        "{F0C07D6A-9B65-4AF5-AF75-7BA80C1EEEFC}"
    ]

    target_section_names = [
        "{22F83DFC-792B-4148-8AED-E77703B6E7BB}", 
        "{56962C9F-AA71-4EBE-BFCC-82940ED9C771}",
        "{2EB8E5B7-6049-402F-80E4-ED6EDFAB83F4}",
        "{02523737-EC78-4811-84BB-436795EE788F}"
    ]

    # Call the function once for all moves
    move_sections_v2(presentation1, move_section_names, target_section_names)

    presentation2.Close()
    if guestBishop > 0:
        presentation3.Close()

    presentation1.SlideShowSettings.Run()

def odasElmilad(Bishop=False, guestBishop=0):
    prs = relative_path(r"قداس.pptx")
    excel = relative_path(r"بيانات القداسات.xlsx")
    excel2 = relative_path(r"Tables.xlsx")
    des_sheet ="القداس"
    image = relative_path(r"Data\Designs\الميلاد.png")
    miladText = ["لأنك ولدت", "aumack", "اف ماسك"]
    replacefile(prs, relative_path(r"Data\CopyData\قداس.pptx"))

    # milad_show_full_sections = ['مدائح الميلاد', 'قسمة للأب في صوم و عيد الميلاد - أيها السيد الرب إلهنا', 
    #                             'فاي اريه بي اوو', 'مرد مزمور الميلاد', 'تكملة مشتركة للمحير',
    #                             'محير عيد الميلاد', 'مرد ابركسيس الميلاد', 'طاي شوري', 'ني سافيف تيرو',
    #                             'الليلويا فاي بيبي', 'اللي القربان', 'هيتينية الميلاد']
                               
    # milad_hide_full_sections = ["ابيناف شوبي", "سوتيس امين", "اجيوس الصلب", "اجيوس الصعود", "مرد المزمور", 
    #                             "مرد الانجيل", "ربع للعذراء" "قسمة - أيها السيد الرب إلهنا", "اك اسماروؤت", "بي اويك"]

    # bishop_show_full_sections = ["امبين يوت اتطايوت", "في حضور الاسقف"]

    milad_show_full_sections = ['{BBBAC16F-044D-4F33-8068-620F498B59CD}', '{072F3D96-A6C8-405F-9A23-7CCA1B2F13FF}', '{670DAA94-A6C9-4CCD-B4E2-958C71CD3E44}', '{20F525FD-C708-4DDD-8E40-FE502EFEBDDE}', '{03E2AC57-01DD-4702-A7A7-186D0E009F55}', '{8DD599A1-D7AC-4AA8-A52B-31BFD527E68E}', '{DEDC0CCA-3854-4E18-8CB2-5D6FEC5BABCC}', '{D95C2E5C-8772-445E-AE3E-2F50770CFC61}', '{B7D98377-B994-4654-B49C-DE10E0DDE4F1}', '{C2F28915-B86E-4596-8EB2-7455EF4E91BD}', '{42181297-997B-4C4C-B43B-4E9D8A23858D}', '{6A153E48-DAA6-4874-ACA3-3EB14F3DC960}', '{DD757736-F2EB-40FA-9016-1E28087A0BE5}', '{59DBF0F6-1D86-41E8-B37A-8AA2368AA8AB}', '{E6CBA825-E339-438B-84B4-326FC5C299C1}', '{973DBBAC-E645-4981-B3F5-5DB1413508D0}', '{DFBFA7AB-D078-4C2A-A184-5823F2253ED4}', '{79502253-2043-4F29-96A3-0E65F6F2C484}']
    milad_hide_full_sections = ['{ECE652ED-1345-4C6D-B92D-5996CFA27AEE}', '{E107D25B-A642-458E-A4F3-B73FDB564A7C}', '{4D2B15D5-C978-467C-9D6C-726FE25128B8}', '{31685B5B-48C4-437E-858C-CF8D225C0C26}', '{CD4B95FF-0E0E-42D3-8DC4-224C3DD732F7}', '{F13A48F2-238D-4617-B84E-9B0A694D9A18}', '{06D592C8-05BF-4B7C-86F7-FDAB3FAB5FB1}', '{681FF6A7-4230-4171-8F41-83FD64E8C960}', '{507EFD97-98F8-4376-848B-20D72E16D2C1}', '{B9A30F5E-0C89-471B-A99A-23DBE7F58504}', '{370EC778-91D3-426C-964D-7E6C28CA69DA}', '{1E7E7987-2CAA-4858-AB80-5A0AF761B6EF}', '{409B3D0A-B40A-4475-811D-72C5125134AB}', '{3CD34DC9-72C7-4E1F-A24E-3878EF0435D6}', '{FB902AC1-5581-4552-B503-116755E9D9A8}', '{A8A52E1F-44DD-45E1-A737-4E13E15D5F1F}']

    katamars = relative_path(r"Data\القطمارس\الايام\القطمارس السنوي ايام (القداس).pptx")
    katamars_sheet = "القطمارس السنوي القداس"
    katamars_values = fetch_data_arrays(excel2, katamars_sheet, 4, 29, [3, 4, 5, 6, 7, 8])
    elbouls1 = katamars_values[0]
    elkatholikon1 = katamars_values[1]
    elebrksis1 = katamars_values[2]
    elmzmor = katamars_values[3]
    elengil = katamars_values[4]
    elbouls2 = elkatholikon1 - 1
    elkatholikon2 = elebrksis1 - 1
    elebrksis2 = elmzmor - 1
    elengil2 = katamars_values[5]

    # milad_values = ["تكملة على حسب المناسبة", "مزمور التوزيع", "مزمور التوزيع", 
    #                 "مرد توزيع الميلاد", "اسبسمس ادام  لعيد الميلاد 1", "ختام الاسبسمس الادام",
    #                 "الانجيل", "المزمور", "اجيوس الميلاد", "اجيوس الميلاد",
    #                 "البولس عربي", "الكاثوليكون", "الابركسيس"]

    milad_values = find_slide_nums_arrays_v2(excel, des_sheet, 
                   ['{A18EDC94-F257-4FAC-99C7-0A8EA70F0FAF}', '{C29E5A83-A98B-4077-8194-99A6D803EF53}', '{C29E5A83-A98B-4077-8194-99A6D803EF53}', '{D10973F3-B5C6-431E-8EDA-60ABA7A98C9E}', '{5CAA4865-5192-4635-9967-094B52BFFE83}', '{1CDFD5FF-8DF5-48CF-A773-F5EEBD56469C}', '{C7D4A109-F792-4661-BAD0-075FD1A1909F}', '{B74DBB8C-2B2D-46E4-9508-DA46008D19A4}', '{76175201-2586-46AD-8E9F-C69E29FE2620}', '{76175201-2586-46AD-8E9F-C69E29FE2620}', '{E234C6C7-3837-4CE4-A541-CDC9627AAAC2}', '{6D4B3F52-63BF-435F-BF0C-C9D41120C2A3}', '{D88055F5-EAA0-4C8E-8249-C364A572BF7B}'],
                   2, [1, 1, 2, 1, 2, 1, 2, 2, 1, 2, 2, 2, 2])
    
    #الختام
    elkhetam = milad_values[0]
    #التوزيع
    mazmorELtawzy3 = milad_values[1] +1
    mazmorELtawzy32 = milad_values[2] - 1
    mrdMazmorEltawzy3 = milad_values[3] 
    #القسمة
    #الاواشي
    #الاسبسمس الواطس
    #الاسبسمس الادام
    esbasmosadam = milad_values[4]
    khetamEsbasmosAdam = milad_values[5]
    #مرد الانجيل
    #المزمور و الانجيل
    elengil3 = milad_values[6]
    elmazmor3 = milad_values[7]
    #المزمور السنجاري و مرد المزمور
    #اجيوس الميلاد
    agiosElmilad = milad_values[8]
    agiosElmilad2 = milad_values[9]
    #القرائات
    elebrksis3 = milad_values[10]
    elkatholikon3 = milad_values[11]
    elbouls3 = milad_values[12]
    #الهيتينيات، طاي شوري، ني سافسف، اللي القربان، الليلويا فاي بابي

    if Bishop == True:
        prs3 = relative_path(r"Data\حضور الأسقف.pptx")
        sheet = "في حضور الأسقف"

        bishop_show_values = ['{F76B0D75-0474-45B5-B79F-7416F354543A}', '{E2968C91-5339-499C-9812-DECCCF58A2CD}',
                              '{62A12AF8-CB6D-4CC5-9DB0-B73A7C24E2AD}', '{B74DBB8C-2B2D-46E4-9508-DA46008D19A4}',
                              '{A9183893-7B7E-459F-8547-F7A8F7D2D521}']
        milad_show_full_sections.extend(bishop_show_values)

        if guestBishop > 0:
            # bishop_values = ["صلاة الشكر", "صلاة الشكر", "طوبه هينا الكبيرة", "طوبه هينا الكبيرة", 
            #                  "نيم بينيوت", "نيم بينيوت", "الهيتنيات", "الهيتنيات",
            #                  "اي اغابي", "اي اغابي", "مرد الكاثوليكون", "مرد الكاثوليكون",
            #                  "الاسبسمس", "الاسبسمس"]

            # bishopDes_values = ["تكملة في حضور الاسقف", "طوبه هينا الكبيرة",
            #                     "امبين يوت اتطايوت", "تكملة الهيتنيات", "بى اهموت غار الصغيرة",
            #                     "تو ماكريو", "اي اغابي", "ابيت جيك ايفول", "مارو اتشاسف",
            #                     "الاسبسمس الادام السنوي+الختام", "ختام الاسبسمس الادام",
            #                     "اوشية الاباء (ب)", "اوشية الاباء غ"]
            
            bishop_values = find_slide_nums_arrays_v2(excel, sheet, 
                            ['{6851F163-CBEF-4014-A853-CE100557BA6A}', '{6851F163-CBEF-4014-A853-CE100557BA6A}', '{B084BC40-61E1-4477-98DA-15CFB06AEE91}', '{B084BC40-61E1-4477-98DA-15CFB06AEE91}', '{97203297-EECB-4D41-B2E3-AD9A4863847E}', '{97203297-EECB-4D41-B2E3-AD9A4863847E}', '{7C84083F-E6D3-4669-9130-AC7E8D935A98}', '{7C84083F-E6D3-4669-9130-AC7E8D935A98}', '{D0234C99-69FE-407A-82A9-D7A676919E93}', '{D0234C99-69FE-407A-82A9-D7A676919E93}', '{0B519645-F935-43AE-A9CE-6E2FC03833BB}', '{0B519645-F935-43AE-A9CE-6E2FC03833BB}', '{D1378DB5-29D1-4800-9D96-10F2535EEB57}', '{D1378DB5-29D1-4800-9D96-10F2535EEB57}'], 
                            2, [1, 2, 1, 2, 1, 2, 1, 2, 1, 2, 1, 2, 1, 2])

            bishopDes_values = find_slide_nums_arrays_v2(excel, des_sheet, 
                                ['{F76B0D75-0474-45B5-B79F-7416F354543A}', '{8DD21CDE-CB6B-4D5B-B995-D2747AB69ED1}', '{E2968C91-5339-499C-9812-DECCCF58A2CD}', '{646A8184-7F05-453A-A2F1-EB9A77D7F0EE}', '{21891CBB-A1EC-4974-B0B6-F74A4B502BC2}', '{12B7D244-BF4C-401B-A65A-D1621D7DD953}', '{F69B50D8-FB5E-4E8C-AE5C-6DCB4790AFAF}', '{38267404-7625-47AA-B0C8-31BCA5D0435D}', '{62A12AF8-CB6D-4CC5-9DB0-B73A7C24E2AD}', '{0C7A7725-643D-4F7E-A2F0-0C8A36C2A594}', '{1CDFD5FF-8DF5-48CF-A773-F5EEBD56469C}', '{74A33555-8E08-47DF-B3CD-A1B4C7AF2B4E}', '{474487EB-4554-4A30-B351-6EF762D2F2D6}'],                  
                                2, [2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2])

            elshokr = bishopDes_values[0]
            elshokr1 = bishop_values[0]
            elshokr2 = bishop_values[1]

            tobhyna = bishopDes_values[1]
            tobhyna1 = bishop_values[2]
            tobhyna2 = bishop_values[3]

            embiniot = bishopDes_values[2]
            embiniot1 = bishop_values[4]
            embiniot2 = bishop_values[5]

            hytynyat = bishopDes_values[3]
            hytynyat1 = bishop_values[6]
            hytynyat2 = bishop_values[7]

            byhmot8ar = bishopDes_values[4]-1
            byhmot8ar1 = bishop_values[4]
            byhmot8ar2 = bishop_values[5]

            tomakario = bishopDes_values[5] - 1
            tomakario1 = bishop_values[4]
            tomakario2 = bishop_values[5]

            eya8aby = bishopDes_values[6] - 1
            eya8aby1 = bishop_values[8]
            eya8aby2 = bishop_values[9]

            mrdElkatholikon = bishopDes_values[7]
            mrdElkatholikon1 = bishop_values[10]
            mrdElkatholikon2 = bishop_values[11]

            maro = bishopDes_values[8]
            maro1 = bishop_values[4]
            maro2 = bishop_values[5]

            esbasmos1 = bishopDes_values[9]
            esbasmos11 = bishop_values[12]
            esbasmos12 = bishop_values[13]

            esbasmos2 = bishopDes_values[10]
            esbasmos21 = bishop_values[12]
            esbasmos22 = bishop_values[13]

            elaba2basyly = bishopDes_values[11] - 1
            elaba2basyly1 = bishop_values[0]
            elaba2basyly2 = bishop_values[1]

            elaba28yry8ory = bishopDes_values[12] - 1
            elaba28yry8ory1 = bishop_values[0]
            elaba28yry8ory2 = bishop_values[1]

            if guestBishop < 2:
                elshokr2 = elshokr2-1
                tobhyna2 = tobhyna2-2
                embiniot2 = embiniot2-1
                hytynyat2 = hytynyat2-3
                byhmot8ar2 = byhmot8ar2-1
                tomakario2 = tomakario2-1
                eya8aby2 = eya8aby2-2
                mrdElkatholikon2 = mrdElkatholikon2-1
                maro2 = maro2-1
                esbasmos12 = esbasmos12-2
                esbasmos22 = esbasmos22-2
                elaba2basyly2 = elaba2basyly2-1
                elaba28yry8ory2 = elaba28yry8ory2-1
        
            start_positions = [mazmorELtawzy3, elaba28yry8ory, elaba2basyly, esbasmos2, esbasmos1, elengil3, elmazmor3, maro, elebrksis3, elkatholikon3, mrdElkatholikon, elbouls3, eya8aby, tomakario, byhmot8ar, hytynyat, embiniot, tobhyna, elshokr]
            start_slides = [mrdMazmorEltawzy3, elaba28yry8ory1, elaba2basyly1, esbasmos21, esbasmos11, elengil, elmzmor, maro1, elebrksis1, elkatholikon1, mrdElkatholikon1, elbouls1, eya8aby1, tomakario1, byhmot8ar1, hytynyat1, embiniot1, tobhyna1, elshokr1]
            end_slides = [mrdMazmorEltawzy3, elaba28yry8ory2, elaba2basyly2, esbasmos22, esbasmos12, elengil2, elmzmor, maro2, elebrksis2, elkatholikon2, mrdElkatholikon2, elbouls2, eya8aby2, tomakario2, byhmot8ar2, hytynyat2, embiniot2, tobhyna2, elshokr2]

    if guestBishop == 0:
        start_positions = [mazmorELtawzy3, elengil3, elmazmor3, elebrksis3, elkatholikon3, elbouls3]
        start_slides = [mrdMazmorEltawzy3, elengil, elmzmor, elebrksis1, elkatholikon1, elbouls1]
        end_slides = [mrdMazmorEltawzy3, elengil2, elmzmor, elebrksis2, elkatholikon2, elbouls2]

    show_hide_insertImage_replaceText(prs, excel, des_sheet, milad_show_full_sections, milad_hide_full_sections, None, None, image, miladText)
    
    powerpoint = win32com.client.Dispatch("PowerPoint.Application")
    powerpoint.Visible = True  # Open PowerPoint application
    presentation1 = open_presentation_relative_path(prs)
    presentation2 = open_presentation_relative_path(katamars)
    
    if guestBishop > 0:
        insert_image_to_slides_same_file(prs3, image)
        presentation3 = open_presentation_relative_path(prs3)

    khetamElmilad = find_slide_index_by_title(presentation1, "الميلاد", elkhetam)
    show_slides(presentation1, [[khetamElmilad, khetamElmilad]])

    vba_array = get_slide_ids_by_number_pairs(prs, [(esbasmosadam, khetamEsbasmosAdam)])
    run_vba_with_slide_id(excel, des_sheet, prs, presentation1, vba_array)
    
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
                new_slide = presentation1.Slides.Paste(slide_index1).SlideShowTransition.Hidden = False
                slide_index1 -= 1
            current_start_slide += 1

        elif (current_position == elengil3 or current_position == elmazmor3 or current_position == elebrksis3 
            or current_position == elkatholikon3 or current_position == elbouls3):
            source_slide = presentation2.Slides(current_end_slide)
            source_slide.Copy()
            new_slide = presentation1.Slides.Paste(current_position)
            pic_shape = new_slide.Shapes.AddPicture(image, LinkToFile=False, SaveWithDocument=True, Left=0, Top=144, Width=720, Height=396)
            pic_shape.ZOrder(1)
            new_slide.SlideShowTransition.Hidden = False
            current_end_slide -= 1
            # Add slide duplication after elmzmor1 is processed
            if guestBishop == 0 and current_position == elmazmor3:
                # Perform the duplication twice
                for _ in range(2):  # Loop to duplicate the section twice
                    # Copy the first slide (agiosElsalyb) and paste it after the second slide (agiosElsalyb2)
                    agiosElmilad_slide = presentation1.Slides(agiosElmilad)
                    agiosElmilad_slide.Copy()
                    presentation1.Slides.Paste(agiosElmilad2 + 1)

                    # Copy the second slide (agiosElsalyb2) and paste it after the newly copied first slide
                    agiosElmilad2_slide = presentation1.Slides(agiosElmilad2)
                    agiosElmilad2_slide.Copy()
                    presentation1.Slides.Paste(agiosElmilad2 + 2)  # Adjust to account for new first slide
            if current_start_slide > current_end_slide:
                current_position += 1

        elif Bishop and current_position in {elaba28yry8ory, elaba2basyly, byhmot8ar, hytynyat, embiniot, 
                                             elshokr, tomakario, eya8aby, esbasmos1, esbasmos2, maro, 
                                             mrdElkatholikon, tobhyna}:
            source_slide = presentation3.Slides(current_end_slide)
            source_slide.Copy()
            if current_position in {elaba28yry8ory, elaba2basyly, byhmot8ar, hytynyat, embiniot, elshokr, maro}:
                presentation1.Slides.Paste(current_position).SlideShowTransition.Hidden = False
            else:
                presentation1.Slides.Paste(current_position).SlideShowTransition.Hidden = True
            current_end_slide -= 1
            if (current_position == maro) and (current_start_slide > current_end_slide):
                    # Perform the duplication twice
                    for _ in range(2):  # Loop to duplicate the section twice
                        # Copy the first slide (agiosElmilad) and paste it after the second slide (agiosElmilad2)
                        agiosElmilad_slide = presentation1.Slides(agiosElmilad)
                        agiosElmilad_slide.Copy()
                        presentation1.Slides.Paste(agiosElmilad2 + 1)

                        # Copy the second slide (agiosElmilad2) and paste it after the newly copied first slide
                        agiosElmilad2_slide = presentation1.Slides(agiosElmilad2)
                        agiosElmilad2_slide.Copy()
                        presentation1.Slides.Paste(agiosElmilad2 + 2)  # Adjust to account for new first slide
            
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

    move_section_names = [
        "{C2F28915-B86E-4596-8EB2-7455EF4E91BD}",
        '{22F83DFC-792B-4148-8AED-E77703B6E7BB}',
        "{1CDFD5FF-8DF5-48CF-A773-F5EEBD56469C}",
        "{2445EC6C-3293-405F-A586-E32C677B7751}",
        "{F0C07D6A-9B65-4AF5-AF75-7BA80C1EEEFC}"
    ]

    target_section_names = [
        "{01F61BFD-210C-4F99-A7C7-7308CFAA93F4}",
        "{8A043029-9932-4D3C-8172-4B76E426B092}",
        '{042E4C02-9095-48E6-B619-5E95A092A467}',
        "{2EB8E5B7-6049-402F-80E4-ED6EDFAB83F4}",
        "{2445EC6C-3293-405F-A586-E32C677B7751}"
    ]

    # Call the function once for all moves
    move_sections_v2(presentation1, move_section_names, target_section_names)

    presentation2.Close()
    if Bishop:
        presentation3.Close()

    presentation1.SlideShowSettings.Run()

def odasSomNynawa(copticdate, Bishop=False, guestBishop=0):
    from copticDate import CopticCalendar
    cd = CopticCalendar().coptic_to_gregorian(copticdate)
    prs = relative_path(r"قداس.pptx")
    excel = relative_path(r"بيانات القداسات.xlsx")
    excel2 = relative_path(r"Tables.xlsx")
    des_sheet ="القداس"
    replacefile(prs, relative_path(r"Data\CopyData\قداس.pptx"))

    # milad_show_full_sections = ["الليلويا اي اي ايخون", "نيف سينتى", " انثو تي تي شوري", 
    #                             "مرد الابركسيس لايام الصوم", "قسمة صوم نينوى - الله الرحوم",
    #                             "مدائح صوم نينوى", "الختام في الصوم المقدس", "اوشية اهوية السماء غ", "اوشية اهوية السماء"]
                               
    # milad_hide_full_sections = ["سوتيس امين", "مرد ابركسيس سنوي", "اكسيا", "مرد الانجيل", "ربع للعذراء",
    #                             "قسمة - أيها السيد الرب إلهنا", "اك اسماروؤت", "بي اويك", "الختام السنوي"]

    # nynawa_hide_full_sections_ranges = [["الهيتنيات", "تين اوؤشت"]]

    nynawa_show_full_sections = ['{44ABFE06-796C-4477-8C9D-E1B568FAD2FF}', '{9810C502-7526-4D63-96A4-F676E5AF5A5F}', '{8DFA2B1F-C47F-42A1-A4F9-ED09CB4F6CB8}', '{456002DB-7C3A-44F7-87FE-507A15868231}', '{BB553AF2-8A5E-46E8-B567-4E8500E5A1C1}', '{43C9AE43-EC80-4A72-97F0-F38A712DDA09}', '{4A3AE26D-6D71-4143-8C05-7618E08EF248}', '{BC7E3DCD-6AA8-44CC-B8AF-BC3E2BC71B5A}', '{A20DA654-32F7-4B4C-96CB-C76232EB96E8}']
                               
    nynawa_hide_full_sections = ['{ECE652ED-1345-4C6D-B92D-5996CFA27AEE}', '{D5DB63D0-39EE-49CE-8855-58CE02719834}', '{AF548422-5DCE-4418-8D21-7DB43CBC4C00}', '{A8A52E1F-44DD-45E1-A737-4E13E15D5F1F}', '{06D592C8-05BF-4B7C-86F7-FDAB3FAB5FB1}', '{681FF6A7-4230-4171-8F41-83FD64E8C960}', '{507EFD97-98F8-4376-848B-20D72E16D2C1}', '{B9A30F5E-0C89-471B-A99A-23DBE7F58504}', '{0D2A50D9-F484-4E60-922B-66FF81444E2C}']

    nynawa_hide_full_sections_ranges = [['{79CED7F3-DA1D-467F-AA09-4187C8DE51E8}', '{646A8184-7F05-453A-A2F1-EB9A77D7F0EE}']]
    
    nynawa_show_full_sections_ranges = [['{82EDE42C-2707-426C-B711-39822D53D8F1}', '{B2279EA0-2880-46C0-85A4-7D2C77B00076}']]

    match cd.weekday():
        case 0: nynawa_show_full_sections.extend(["{F985FA26-9C0E-4F27-A8E4-075769969EF9}"])
        case 1: nynawa_show_full_sections.extend(["{338DD81B-993A-42AF-905A-5CE94AA09854}"])
        case 2: nynawa_show_full_sections.extend(["{C29444F8-813A-4C6E-A49E-4198160243F2}"])

    katamars = relative_path(r"Data\القطمارس\الصوم الكبير و صوم نينوى\قرائات صوم نينوى و فصح يونان.pptx")
    katamars_sheet = "صوم نينوى و فصح يونان"
    katamars_values = fetch_data_arrays(excel2, katamars_sheet, copticdate[1], copticdate[2], [10, 11, 12, 13, 14, 15])
    elbouls1 = katamars_values[0]
    elkatholikon1 = katamars_values[1]
    elebrksis1 = katamars_values[2]
    elmzmor = katamars_values[3]
    elengil = katamars_values[4]
    elbouls2 = elkatholikon1 - 1
    elkatholikon2 = elebrksis1 - 1
    elebrksis2 = elmzmor - 1
    elengil2 = katamars_values[5]

    # nynawa_values = ["تكملة على حسب المناسبة", "مزمور التوزيع", "مزمور التوزيع", "مرد توزيع صوم نينوى",
    #                  "قسمة للإبن في صوم الاربعين - أنت هو الله الرحوم مخلص", 
    #                  "قسمة للأب في الصوم الكبير - أيها السيد الرب الإله ضابط الكل", 
    #                  "الانجيل", "المزمور", "الابركسيس", "الكاثوليكون", "البولس عربي", "بدء قداس الكلمة"]
    
    nynawa_values = find_slide_nums_arrays_v2(excel, des_sheet, 
                    ['{A18EDC94-F257-4FAC-99C7-0A8EA70F0FAF}', '{C29E5A83-A98B-4077-8194-99A6D803EF53}', '{C29E5A83-A98B-4077-8194-99A6D803EF53}', '{8B7ED515-48EF-40B6-B9F2-80B8A4CC1126}', '{41DD9B19-60F1-44CE-BF9E-34D6DA71F069}', '{F076B353-F7AB-4001-959A-5D482DE256DB}', '{C7D4A109-F792-4661-BAD0-075FD1A1909F}', '{B74DBB8C-2B2D-46E4-9508-DA46008D19A4}', '{E234C6C7-3837-4CE4-A541-CDC9627AAAC2}', '{6D4B3F52-63BF-435F-BF0C-C9D41120C2A3}', '{D88055F5-EAA0-4C8E-8249-C364A572BF7B}', '{C08D8D44-E49E-47CE-8027-C8AE26B1AA9A}'],
                    2, [1, 1, 2, 1, 2, 2, 2, 2, 2, 2, 2, 2])
    
    #الختام
    elkhetam = nynawa_values[0]
    #التوزيع
    mazmorELtawzy3 = nynawa_values[1] +3
    mazmorELtawzy32 = nynawa_values[2] - 1
    mrdMazmorEltawzy3 = nynawa_values[3]
    #القسمة
    EsmaElsomElkbyr1 = nynawa_values[4]
    EsmaElsomElkbyr2 = nynawa_values[5]
    #القرائات
    elengil3 = nynawa_values[6]
    elmazmor3 = nynawa_values[7]
    elebrksis3 = nynawa_values[8]
    elkatholikon3 = nynawa_values[9]
    elbouls3 = nynawa_values[10]

    if Bishop == True:
        prs3 = relative_path(r"Data\حضور الأسقف.pptx")
        sheet = "في حضور الأسقف"

        bishop_show_values = ['{2BCF4F8C-25F0-43C5-B224-6528B2EA3F2F}', '{F76B0D75-0474-45B5-B79F-7416F354543A}',
                              '{E2968C91-5339-499C-9812-DECCCF58A2CD}', '{62A12AF8-CB6D-4CC5-9DB0-B73A7C24E2AD}', 
                              '{B74DBB8C-2B2D-46E4-9508-DA46008D19A4}', '{A9183893-7B7E-459F-8547-F7A8F7D2D521}']
        nynawa_show_full_sections.extend(bishop_show_values)

        if guestBishop > 0:
            # bishop_values = ["صلاة الشكر", "صلاة الشكر", "طوبه هينا الكبيرة", "طوبه هينا الكبيرة", 
            #                  "نيم بينيوت", "نيم بينيوت", "الهيتنيات", "الهيتنيات",
            #                  "اي اغابي", "اي اغابي", "مرد الكاثوليكون", "مرد الكاثوليكون",
            #                  "الاسبسمس", "الاسبسمس"]

            # bishopDes_values = ["تكملة في حضور الاسقف", "طوبه هينا الكبيرة",
            #                     "امبين يوت اتطايوت", "تكملة الهيتنيات", "بى اهموت غار الصغيرة",
            #                     "تو ماكريو", "اي اغابي", "ابيت جيك ايفول", "مارو اتشاسف",
            #                     "الاسبسمس الادام السنوي+الختام", "ختام الاسبسمس الادام",
            #                     "اوشية الاباء (ب)", "اوشية الاباء غ"]
            
            bishop_values = find_slide_nums_arrays_v2(excel, sheet, 
                            ['{6851F163-CBEF-4014-A853-CE100557BA6A}', '{6851F163-CBEF-4014-A853-CE100557BA6A}', '{B084BC40-61E1-4477-98DA-15CFB06AEE91}', '{B084BC40-61E1-4477-98DA-15CFB06AEE91}', '{97203297-EECB-4D41-B2E3-AD9A4863847E}', '{97203297-EECB-4D41-B2E3-AD9A4863847E}', '{7C84083F-E6D3-4669-9130-AC7E8D935A98}', '{7C84083F-E6D3-4669-9130-AC7E8D935A98}', '{D0234C99-69FE-407A-82A9-D7A676919E93}', '{D0234C99-69FE-407A-82A9-D7A676919E93}', '{0B519645-F935-43AE-A9CE-6E2FC03833BB}', '{0B519645-F935-43AE-A9CE-6E2FC03833BB}', '{D1378DB5-29D1-4800-9D96-10F2535EEB57}', '{D1378DB5-29D1-4800-9D96-10F2535EEB57}'], 
                            2, [1, 2, 1, 2, 1, 2, 1, 2, 1, 2, 1, 2, 1, 2])

            bishopDes_values = find_slide_nums_arrays_v2(excel, des_sheet, 
                               ['{F76B0D75-0474-45B5-B79F-7416F354543A}', '{8DD21CDE-CB6B-4D5B-B995-D2747AB69ED1}', '{E2968C91-5339-499C-9812-DECCCF58A2CD}', '{646A8184-7F05-453A-A2F1-EB9A77D7F0EE}', '{21891CBB-A1EC-4974-B0B6-F74A4B502BC2}', '{12B7D244-BF4C-401B-A65A-D1621D7DD953}', '{F69B50D8-FB5E-4E8C-AE5C-6DCB4790AFAF}', '{38267404-7625-47AA-B0C8-31BCA5D0435D}', '{62A12AF8-CB6D-4CC5-9DB0-B73A7C24E2AD}', '{0C7A7725-643D-4F7E-A2F0-0C8A36C2A594}', '{1CDFD5FF-8DF5-48CF-A773-F5EEBD56469C}', '{74A33555-8E08-47DF-B3CD-A1B4C7AF2B4E}', '{474487EB-4554-4A30-B351-6EF762D2F2D6}'],                  
                               2, [2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2])

            elshokr = bishopDes_values[0]
            elshokr1 = bishop_values[0]
            elshokr2 = bishop_values[1]

            tobhyna = bishopDes_values[1]
            tobhyna1 = bishop_values[2]
            tobhyna2 = bishop_values[3]

            embiniot = bishopDes_values[2]
            embiniot1 = bishop_values[4]
            embiniot2 = bishop_values[5]

            hytynyat = bishopDes_values[3]
            hytynyat1 = bishop_values[6]
            hytynyat2 = bishop_values[7]

            byhmot8ar = bishopDes_values[4]-1
            byhmot8ar1 = bishop_values[4]
            byhmot8ar2 = bishop_values[5]

            tomakario = bishopDes_values[5] - 1
            tomakario1 = bishop_values[4]
            tomakario2 = bishop_values[5]

            eya8aby = bishopDes_values[6] - 1
            eya8aby1 = bishop_values[8]
            eya8aby2 = bishop_values[9]

            mrdElkatholikon = bishopDes_values[7]
            mrdElkatholikon1 = bishop_values[10]
            mrdElkatholikon2 = bishop_values[11]

            maro = bishopDes_values[8]
            maro1 = bishop_values[4]
            maro2 = bishop_values[5]

            esbasmos1 = bishopDes_values[9]
            esbasmos11 = bishop_values[12]
            esbasmos12 = bishop_values[13]

            esbasmos2 = bishopDes_values[10]
            esbasmos21 = bishop_values[12]
            esbasmos22 = bishop_values[13]

            elaba2basyly = bishopDes_values[11] - 1
            elaba2basyly1 = bishop_values[0]
            elaba2basyly2 = bishop_values[1]

            elaba28yry8ory = bishopDes_values[12] - 1
            elaba28yry8ory1 = bishop_values[0]
            elaba28yry8ory2 = bishop_values[1]

            if guestBishop < 2:
                elshokr2 = elshokr2-1
                tobhyna2 = tobhyna2-2
                embiniot2 = embiniot2-1
                hytynyat2 = hytynyat2-3
                byhmot8ar2 = byhmot8ar2-1
                tomakario2 = tomakario2-1
                eya8aby2 = eya8aby2-2
                mrdElkatholikon2 = mrdElkatholikon2-1
                maro2 = maro2-1
                esbasmos12 = esbasmos12-2
                esbasmos22 = esbasmos22-2
                elaba2basyly2 = elaba2basyly2-1
                elaba28yry8ory2 = elaba28yry8ory2-1
        
            start_positions = [mazmorELtawzy3, elaba28yry8ory, elaba2basyly, esbasmos2, esbasmos1, elengil3, elmazmor3, maro, elebrksis3, elkatholikon3, mrdElkatholikon, elbouls3, eya8aby, tomakario, byhmot8ar, hytynyat, embiniot, tobhyna, elshokr]
            start_slides = [mrdMazmorEltawzy3, elaba28yry8ory1, elaba2basyly1, esbasmos21, esbasmos11, elengil, elmzmor, maro1, elebrksis1, elkatholikon1, mrdElkatholikon1, elbouls1, eya8aby1, tomakario1, byhmot8ar1, hytynyat1, embiniot1, tobhyna1, elshokr1]
            end_slides = [mrdMazmorEltawzy3, elaba28yry8ory2, elaba2basyly2, esbasmos22, esbasmos12, elengil2, elmzmor, maro2, elebrksis2, elkatholikon2, mrdElkatholikon2, elbouls2, eya8aby2, tomakario2, byhmot8ar2, hytynyat2, embiniot2, tobhyna2, elshokr2]

    if guestBishop == 0:
        start_positions = [mazmorELtawzy3, elengil3, elmazmor3, elebrksis3, elkatholikon3, elbouls3]
        start_slides = [mrdMazmorEltawzy3, elengil, elmzmor, elebrksis1, elkatholikon1, elbouls1]
        end_slides = [mrdMazmorEltawzy3, elengil2, elmzmor, elebrksis2, elkatholikon2, elbouls2]

    show_hide_insertImage_replaceText(prs, excel, des_sheet, nynawa_show_full_sections, nynawa_hide_full_sections, nynawa_show_full_sections_ranges, nynawa_hide_full_sections_ranges)
    
    powerpoint = win32com.client.Dispatch("PowerPoint.Application")
    powerpoint.Visible = True  # Open PowerPoint application
    presentation1 = open_presentation_relative_path(prs)
    presentation2 = open_presentation_relative_path(katamars)
    
    if Bishop == True:
        presentation3 = open_presentation_relative_path(prs3)

    khetamAhwya = find_slide_index_by_title(presentation1, "الاهوية", elkhetam)
    show_slides(presentation1, [[khetamAhwya, khetamAhwya], [EsmaElsomElkbyr1, EsmaElsomElkbyr1], [EsmaElsomElkbyr2, EsmaElsomElkbyr2]])

    run_vba_with_slide_id(excel, des_sheet, prs, presentation1)
    agbya(presentation1, nynawa_values[11], 3)
    
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

        elif (current_position == elengil3 or current_position == elmazmor3 or current_position == elebrksis3 
            or current_position == elkatholikon3 or current_position == elbouls3):
            source_slide = presentation2.Slides(current_end_slide)
            source_slide.Copy()
            presentation1.Slides.Paste(current_position).SlideShowTransition.Hidden = False
            current_end_slide -= 1
            if current_start_slide > current_end_slide:
                current_position += 1

        elif Bishop and current_position in {elaba28yry8ory, elaba2basyly, byhmot8ar, hytynyat, embiniot, 
                                             elshokr, tomakario, eya8aby, esbasmos1, esbasmos2, maro, 
                                             mrdElkatholikon, tobhyna}:
            source_slide = presentation3.Slides(current_end_slide)
            source_slide.Copy()
            if current_position in {elaba28yry8ory, elaba2basyly, byhmot8ar, hytynyat, embiniot, elshokr, maro}:
                presentation1.Slides.Paste(current_position).SlideShowTransition.Hidden = False
            else:
                presentation1.Slides.Paste(current_position).SlideShowTransition.Hidden = True
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

    move_section_names = [
        "{BB553AF2-8A5E-46E8-B567-4E8500E5A1C1}"
    ]

    target_section_names = [
        "{01F61BFD-210C-4F99-A7C7-7308CFAA93F4}"
    ]

    # Call the function once for all moves
    move_sections_v2(presentation1, move_section_names, target_section_names)

    presentation2.Close()
    if Bishop:
        presentation3.Close()

    presentation1.SlideShowSettings.Run()

def odasSanawy(copticdate, season, Bishop=False, guestBishop=0):
    from copticDate import CopticCalendar
    cd = CopticCalendar().coptic_to_gregorian(copticdate)
    prs1 = relative_path(r"قداس.pptx")  # Using the relative path
    excel = relative_path(r"بيانات القداسات.xlsx")
    excel2 = relative_path(r"Tables.xlsx")
    des_sheet ="القداس"
    replacefile(prs1, relative_path(r"Data\CopyData\قداس.pptx"))

    if season == 13:
        prs2 = relative_path(r"Data\القطمارس\الصوم الكبير و صوم نينوى\قرائات صوم نينوى و فصح يونان.pptx")
        katamars_sheet = "صوم نينوى و فصح يونان"
        km = copticdate[1]
        kd = copticdate[2]
        katamars_offsets = [10, 11, 12, 13, 14, 15]
    if season == 15.2:
        prs2 = relative_path(r"Data\القطمارس\قطمارس الصوم الكبير (القداس).pptx")
        katamars_sheet = "قطمارس الصوم الكبير"
        km = copticdate[1]
        kd = copticdate[2]
        katamars_offsets = [3, 4, 5, 6, 7, 8]
    elif cd.weekday() == 6:
        prs2 = relative_path(r"Data\القطمارس\الاحاد\القطمارس السنوي احاد (القداس).pptx")
        katamars_sheet = "قطمارس الاحاد للقداس"
        km = copticdate[1]
        kd = (copticdate[2] - 1) // 7 + 1
        katamars_offsets = [3, 4, 5, 6, 7, 8]
    else: 
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

    #sanawy_values = ["تكملة على حسب المناسبة", "الانجيل", "المزمور", "الابركسيس", "الكاثوليكون", "البولس عربي"]
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

    if Bishop == True:
        prs3 = relative_path(r"Data\حضور الأسقف.pptx")
        sheet = "في حضور الأسقف"

        # bishop_show_values = ['في حضور الاسقف', 'المزمور', 'مارو اتشاسف', 'امبين يوت اتطايوت', 
        #                       'ني سافيف تيرو', 'تكملة في حضور الاسقف', 'لحن اك اسمارؤوت']

        #bishop_hide_values = ['سوتيس امين']

        bishop_show_values = ['{2BCF4F8C-25F0-43C5-B224-6528B2EA3F2F}', '{F76B0D75-0474-45B5-B79F-7416F354543A}',
                              '{E2968C91-5339-499C-9812-DECCCF58A2CD}', '{62A12AF8-CB6D-4CC5-9DB0-B73A7C24E2AD}', 
                              '{B74DBB8C-2B2D-46E4-9508-DA46008D19A4}', '{A9183893-7B7E-459F-8547-F7A8F7D2D521}', 
                              '{670DAA94-A6C9-4CCD-B4E2-958C71CD3E44}']
        
        bishop_hide_values = ['{4D2B15D5-C978-467C-9D6C-726FE25128B8}']
        
        sanawy_show_full_sections.extend(bishop_show_values)
        sanawy_hide_full_sections.extend(bishop_hide_values)

        if guestBishop > 0:
            # bishop_values = ["صلاة الشكر", "صلاة الشكر", "طوبه هينا الكبيرة", "طوبه هينا الكبيرة", 
            #                  "نيم بينيوت", "نيم بينيوت", "الهيتنيات", "الهيتنيات",
            #                  "اي اغابي", "اي اغابي", "مرد الكاثوليكون", "مرد الكاثوليكون",
            #                  "الاسبسمس", "الاسبسمس"]

            # bishopDes_values = ["تكملة في حضور الاسقف", "طوبه هينا الكبيرة",
            #                     "امبين يوت اتطايوت", "تكملة الهيتنيات", "بى اهموت غار الصغيرة",
            #                     "تو ماكريو", "اي اغابي", "ابيت جيك ايفول", "مارو اتشاسف",
            #                     "الاسبسمس الادام السنوي+الختام", "ختام الاسبسمس الادام",
            #                     "اوشية الاباء (ب)", "اوشية الاباء غ"]
            
            bishop_values = find_slide_nums_arrays_v2(excel, sheet, 
                            ['{6851F163-CBEF-4014-A853-CE100557BA6A}', '{6851F163-CBEF-4014-A853-CE100557BA6A}', '{B084BC40-61E1-4477-98DA-15CFB06AEE91}', '{B084BC40-61E1-4477-98DA-15CFB06AEE91}', '{97203297-EECB-4D41-B2E3-AD9A4863847E}', '{97203297-EECB-4D41-B2E3-AD9A4863847E}', '{7C84083F-E6D3-4669-9130-AC7E8D935A98}', '{7C84083F-E6D3-4669-9130-AC7E8D935A98}', '{D0234C99-69FE-407A-82A9-D7A676919E93}', '{D0234C99-69FE-407A-82A9-D7A676919E93}', '{0B519645-F935-43AE-A9CE-6E2FC03833BB}', '{0B519645-F935-43AE-A9CE-6E2FC03833BB}', '{D1378DB5-29D1-4800-9D96-10F2535EEB57}', '{D1378DB5-29D1-4800-9D96-10F2535EEB57}'], 
                            2, [1, 2, 1, 2, 1, 2, 1, 2, 1, 2, 1, 2, 1, 2])

            bishopDes_values = find_slide_nums_arrays_v2(excel, des_sheet, 
                               ['{F76B0D75-0474-45B5-B79F-7416F354543A}', '{8DD21CDE-CB6B-4D5B-B995-D2747AB69ED1}', '{E2968C91-5339-499C-9812-DECCCF58A2CD}', '{646A8184-7F05-453A-A2F1-EB9A77D7F0EE}', '{21891CBB-A1EC-4974-B0B6-F74A4B502BC2}', '{12B7D244-BF4C-401B-A65A-D1621D7DD953}', '{F69B50D8-FB5E-4E8C-AE5C-6DCB4790AFAF}', '{38267404-7625-47AA-B0C8-31BCA5D0435D}', '{62A12AF8-CB6D-4CC5-9DB0-B73A7C24E2AD}', '{0C7A7725-643D-4F7E-A2F0-0C8A36C2A594}', '{1CDFD5FF-8DF5-48CF-A773-F5EEBD56469C}', '{74A33555-8E08-47DF-B3CD-A1B4C7AF2B4E}', '{474487EB-4554-4A30-B351-6EF762D2F2D6}'],                  
                               2, [2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2])

            elshokr = bishopDes_values[0]
            elshokr1 = bishop_values[0]
            elshokr2 = bishop_values[1]

            tobhyna = bishopDes_values[1]
            tobhyna1 = bishop_values[2]
            tobhyna2 = bishop_values[3]

            embiniot = bishopDes_values[2]
            embiniot1 = bishop_values[4]
            embiniot2 = bishop_values[5]

            hytynyat = bishopDes_values[3]
            hytynyat1 = bishop_values[6]
            hytynyat2 = bishop_values[7]

            byhmot8ar = bishopDes_values[4]-1
            byhmot8ar1 = bishop_values[4]
            byhmot8ar2 = bishop_values[5]

            tomakario = bishopDes_values[5] - 1
            tomakario1 = bishop_values[4]
            tomakario2 = bishop_values[5]

            eya8aby = bishopDes_values[6] - 1
            eya8aby1 = bishop_values[8]
            eya8aby2 = bishop_values[9]

            mrdElkatholikon = bishopDes_values[7]
            mrdElkatholikon1 = bishop_values[10]
            mrdElkatholikon2 = bishop_values[11]

            maro = bishopDes_values[8]
            maro1 = bishop_values[4]
            maro2 = bishop_values[5]

            esbasmos1 = bishopDes_values[9]
            esbasmos11 = bishop_values[12]
            esbasmos12 = bishop_values[13]

            esbasmos2 = bishopDes_values[10]
            esbasmos21 = bishop_values[12]
            esbasmos22 = bishop_values[13]

            elaba2basyly = bishopDes_values[11] - 1
            elaba2basyly1 = bishop_values[0]
            elaba2basyly2 = bishop_values[1]

            elaba28yry8ory = bishopDes_values[12] - 1
            elaba28yry8ory1 = bishop_values[0]
            elaba28yry8ory2 = bishop_values[1]

            if guestBishop < 2:
                elshokr2 = elshokr2-1
                tobhyna2 = tobhyna2-2
                embiniot2 = embiniot2-1
                hytynyat2 = hytynyat2-3
                byhmot8ar2 = byhmot8ar2-1
                tomakario2 = tomakario2-1
                eya8aby2 = eya8aby2-2
                mrdElkatholikon2 = mrdElkatholikon2-1
                maro2 = maro2-1
                esbasmos12 = esbasmos12-2
                esbasmos22 = esbasmos22-2
                elaba2basyly2 = elaba2basyly2-1
                elaba28yry8ory2 = elaba28yry8ory2-1
        
            start_positions = [elaba28yry8ory, elaba2basyly, esbasmos2, esbasmos1, elengil3, elmazmor3, maro, elebrksis3, elkatholikon3, mrdElkatholikon, elbouls3, eya8aby, tomakario, byhmot8ar, hytynyat, embiniot, tobhyna, elshokr]
            start_slides = [elaba28yry8ory1, elaba2basyly1, esbasmos21, esbasmos11, elengil, elmzmor, maro1, elebrksis1, elkatholikon1, mrdElkatholikon1, elbouls1, eya8aby1, tomakario1, byhmot8ar1, hytynyat1, embiniot1, tobhyna1, elshokr1]
            end_slides = [elaba28yry8ory2, elaba2basyly2, esbasmos22, esbasmos12, elengil2, elmzmor, maro2, elebrksis2, elkatholikon2, mrdElkatholikon2, elbouls2, eya8aby2, tomakario2, byhmot8ar2, hytynyat2, embiniot2, tobhyna2, elshokr2]

    if guestBishop == 0:
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
        show_hide_insertImage_replaceText(prs1, excel, des_sheet, sanawy_show_full_sections, sanawy_hide_full_sections, newText=["لأنك قمت","aktwnk", "آك طونك"])
    else:
        show_hide_insertImage_replaceText(prs1, excel, des_sheet, sanawy_show_full_sections, sanawy_hide_full_sections)

    powerpoint = win32com.client.Dispatch("PowerPoint.Application")
    powerpoint.Visible = True  # Open PowerPoint application
    presentation1 = open_presentation_relative_path(prs1)
    presentation2 = open_presentation_relative_path(prs2)

    if guestBishop>0:
        presentation3 = open_presentation_relative_path(prs3)

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

    run_vba_with_slide_id(excel, des_sheet, prs1, presentation1)

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

        elif Bishop and current_position in {elaba28yry8ory, elaba2basyly, byhmot8ar, hytynyat, embiniot, 
                                             elshokr, tomakario, eya8aby, esbasmos1, esbasmos2, maro, 
                                             mrdElkatholikon, tobhyna}:
            source_slide = presentation3.Slides(current_end_slide)
            source_slide.Copy()
            if current_position in {elaba28yry8ory, elaba2basyly, byhmot8ar, hytynyat, embiniot, elshokr, maro}:
                presentation1.Slides.Paste(current_position).SlideShowTransition.Hidden = False
            else:
                presentation1.Slides.Paste(current_position).SlideShowTransition.Hidden = True
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

    if El3adraAndAngels :
        move_sections_v2(presentation1, ['{71232865-63AA-40C5-8F02-BABBFE7297D3}'], ['{01F61BFD-210C-4F99-A7C7-7308CFAA93F4}'])
    elif season == 6:
        move_sections_v2(presentation1, ['{C2F28915-B86E-4596-8EB2-7455EF4E91BD}'], ['{01F61BFD-210C-4F99-A7C7-7308CFAA93F4}'])
    
    presentation2.Close()
    if guestBishop>0:
        presentation3.Close()

    presentation1.SlideShowSettings.Run()

def odasElSomElkbyr(copticdate, season, Bishop=False, guestBishop=0):
    from copticDate import CopticCalendar
    cd = CopticCalendar().coptic_to_gregorian(copticdate)
    prs1 = relative_path(r"قداس.pptx")  # Using the relative path
    excel = relative_path(r"بيانات القداسات.xlsx")
    excel2 = relative_path(r"Tables.xlsx")
    des_sheet ="القداس"
    replacefile(prs1, relative_path(r"Data\CopyData\قداس.pptx"))

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
    #               "الانجيل", "المزمور", "الابركسيس", "الكاثوليكون", "البولس عربي", "بدء قداس الكلمة"]
    
    som_values = find_slide_nums_arrays_v2(excel, des_sheet, 
                 ['{A18EDC94-F257-4FAC-99C7-0A8EA70F0FAF}', '{C29E5A83-A98B-4077-8194-99A6D803EF53}', '{C29E5A83-A98B-4077-8194-99A6D803EF53}', '{F9ED982E-F6FB-4E2B-8955-C5E80C70C2D6}', '{F076B353-F7AB-4001-959A-5D482DE256DB}', '{C7D4A109-F792-4661-BAD0-075FD1A1909F}', '{B74DBB8C-2B2D-46E4-9508-DA46008D19A4}', '{E234C6C7-3837-4CE4-A541-CDC9627AAAC2}', '{6D4B3F52-63BF-435F-BF0C-C9D41120C2A3}', '{D88055F5-EAA0-4C8E-8249-C364A572BF7B}', '{C08D8D44-E49E-47CE-8027-C8AE26B1AA9A}'],
                 2, [1, 1, 2, 1, 2, 2, 2, 2, 2, 2, 2])
    
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

    #في حالة احد الرفاع و الايام و (الآحاد والاثنين الاول و جمعة ختام الصوم)
    if season == 15.3: #احد الرفاع
        # som_show_full_sections.extend(["الليلويا فاي بيبي", "طاي شوري", "اونيشتي اميستيريون"])
        
        mazmorELtawzy3 -= 2
        som_show_full_sections.extend(['{DB680029-4B05-4C3F-98B3-FB5469AADBD7}', '{20F525FD-C708-4DDD-8E40-FE502EFEBDDE}', '{072F3D96-A6C8-405F-9A23-7CCA1B2F13FF}'])
    
    elif cd.weekday() == 6 or cd.weekday() == 5 or season == 15.1 or season == 15.4: #سبوت و آحاد و الاثنين الاول و جمعة ختام الصوم
        # som_show_full_sections.extend(["الليلويا جى اف ميفي", "تي شوري", "مرد ابركسيس احاد الصوم", "ميغالو",
        #                                "محير الصوم الكبير", "اونيشتي اميستيريون"])
        # som_hide_full_sections.extend(["اكسيا"])

        mazmorELtawzy3 -= 2    
        som_show_full_sections.extend(['{F8B8BD1A-9861-4FDC-A89B-06B55C0795E8}', '{43B74D4E-929B-4654-8DB7-4675EFF27370}', '{D8B0DE9D-238E-466B-827A-DEAAC424B5A5}', '{441DBF54-2A06-4556-90D4-48E9C63759E4}', '{A553FAD1-4C8F-489F-81BF-7885E6811853}', '{DB680029-4B05-4C3F-98B3-FB5469AADBD7}'])
        som_hide_full_sections.extend(['{A8A52E1F-44DD-45E1-A737-4E13E15D5F1F}'])

        if season == 15.1 or season == 15.4:
            som_show_full_sections.extend(['{9810C502-7526-4D63-96A4-F676E5AF5A5F}'])
    
    else: #ايام
        # som_hide_full_sections.extend(["سوتيس امين", "سوتيس امين 2", "مرد ابركسيس سنوي", "اكسيا", "الختام السنوي"])
        # som_show_full_sections.extend(["الليلويا اي اي ايخون", "نيف سينتى", " انثو تي تي شوري", 
        #                                "مرد الابركسيس لايام الصوم", "مرد انجيل ايام الصوم الكبير", "بي ماي رومي", 
        #                                "الختام في الصوم المقدس"])
        # som_hide_full_sections_ranges = [["الهيتنيات", "تكملة الهيتينيات"]]
        
        som_hide_full_sections.extend(['{A8A52E1F-44DD-45E1-A737-4E13E15D5F1F}', '{AF548422-5DCE-4418-8D21-7DB43CBC4C00}', '{D5DB63D0-39EE-49CE-8855-58CE02719834}', '{4D2B15D5-C978-467C-9D6C-726FE25128B8}', '{0D2A50D9-F484-4E60-922B-66FF81444E2C}'])
        som_show_full_sections.extend(['{315091E2-E367-43B7-A35E-4175DF947038}', '{229D9524-1F56-4456-A2B5-2321A4532E39}', '{456002DB-7C3A-44F7-87FE-507A15868231}', '{8DFA2B1F-C47F-42A1-A4F9-ED09CB4F6CB8}', '{9810C502-7526-4D63-96A4-F676E5AF5A5F}', '{44ABFE06-796C-4477-8C9D-E1B568FAD2FF}', '{4A3AE26D-6D71-4143-8C05-7618E08EF248}'])
        som_hide_full_sections_ranges.extend([['{79CED7F3-DA1D-467F-AA09-4187C8DE51E8}', '{646A8184-7F05-453A-A2F1-EB9A77D7F0EE}']])
    
    if Bishop == True:
        prs3 = relative_path(r"Data\حضور الأسقف.pptx")
        sheet = "في حضور الأسقف"

        bishop_show_values = ['{2BCF4F8C-25F0-43C5-B224-6528B2EA3F2F}', '{F76B0D75-0474-45B5-B79F-7416F354543A}',
                              '{E2968C91-5339-499C-9812-DECCCF58A2CD}', '{62A12AF8-CB6D-4CC5-9DB0-B73A7C24E2AD}', 
                              '{B74DBB8C-2B2D-46E4-9508-DA46008D19A4}', '{A9183893-7B7E-459F-8547-F7A8F7D2D521}']
        som_show_full_sections.extend(bishop_show_values)

        if guestBishop > 0:
            # bishop_values = ["صلاة الشكر", "صلاة الشكر", "طوبه هينا الكبيرة", "طوبه هينا الكبيرة", 
            #                  "نيم بينيوت", "نيم بينيوت", "الهيتنيات", "الهيتنيات",
            #                  "اي اغابي", "اي اغابي", "مرد الكاثوليكون", "مرد الكاثوليكون",
            #                  "الاسبسمس", "الاسبسمس"]

            # bishopDes_values = ["تكملة في حضور الاسقف", "طوبه هينا الكبيرة",
            #                     "امبين يوت اتطايوت", "تكملة الهيتنيات", "بى اهموت غار الصغيرة",
            #                     "تو ماكريو", "اي اغابي", "ابيت جيك ايفول", "مارو اتشاسف",
            #                     "الاسبسمس الادام السنوي+الختام", "ختام الاسبسمس الادام",
            #                     "اوشية الاباء (ب)", "اوشية الاباء غ"]
            
            bishop_values = find_slide_nums_arrays_v2(excel, sheet, 
                            ['{6851F163-CBEF-4014-A853-CE100557BA6A}', '{6851F163-CBEF-4014-A853-CE100557BA6A}', '{B084BC40-61E1-4477-98DA-15CFB06AEE91}', '{B084BC40-61E1-4477-98DA-15CFB06AEE91}', '{97203297-EECB-4D41-B2E3-AD9A4863847E}', '{97203297-EECB-4D41-B2E3-AD9A4863847E}', '{7C84083F-E6D3-4669-9130-AC7E8D935A98}', '{7C84083F-E6D3-4669-9130-AC7E8D935A98}', '{D0234C99-69FE-407A-82A9-D7A676919E93}', '{D0234C99-69FE-407A-82A9-D7A676919E93}', '{0B519645-F935-43AE-A9CE-6E2FC03833BB}', '{0B519645-F935-43AE-A9CE-6E2FC03833BB}', '{D1378DB5-29D1-4800-9D96-10F2535EEB57}', '{D1378DB5-29D1-4800-9D96-10F2535EEB57}'], 
                            2, [1, 2, 1, 2, 1, 2, 1, 2, 1, 2, 1, 2, 1, 2])

            bishopDes_values = find_slide_nums_arrays_v2(excel, des_sheet, 
                               ['{F76B0D75-0474-45B5-B79F-7416F354543A}', '{8DD21CDE-CB6B-4D5B-B995-D2747AB69ED1}', '{E2968C91-5339-499C-9812-DECCCF58A2CD}', '{646A8184-7F05-453A-A2F1-EB9A77D7F0EE}', '{21891CBB-A1EC-4974-B0B6-F74A4B502BC2}', '{12B7D244-BF4C-401B-A65A-D1621D7DD953}', '{F69B50D8-FB5E-4E8C-AE5C-6DCB4790AFAF}', '{38267404-7625-47AA-B0C8-31BCA5D0435D}', '{62A12AF8-CB6D-4CC5-9DB0-B73A7C24E2AD}', '{0C7A7725-643D-4F7E-A2F0-0C8A36C2A594}', '{1CDFD5FF-8DF5-48CF-A773-F5EEBD56469C}', '{74A33555-8E08-47DF-B3CD-A1B4C7AF2B4E}', '{474487EB-4554-4A30-B351-6EF762D2F2D6}'],                  
                               2, [2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2])

            elshokr = bishopDes_values[0]
            elshokr1 = bishop_values[0]
            elshokr2 = bishop_values[1]

            tobhyna = bishopDes_values[1]
            tobhyna1 = bishop_values[2]
            tobhyna2 = bishop_values[3]

            embiniot = bishopDes_values[2]
            embiniot1 = bishop_values[4]
            embiniot2 = bishop_values[5]

            hytynyat = bishopDes_values[3]
            hytynyat1 = bishop_values[6]
            hytynyat2 = bishop_values[7]

            byhmot8ar = bishopDes_values[4]-1
            byhmot8ar1 = bishop_values[4]
            byhmot8ar2 = bishop_values[5]

            tomakario = bishopDes_values[5] - 1
            tomakario1 = bishop_values[4]
            tomakario2 = bishop_values[5]

            eya8aby = bishopDes_values[6] - 1
            eya8aby1 = bishop_values[8]
            eya8aby2 = bishop_values[9]

            mrdElkatholikon = bishopDes_values[7]
            mrdElkatholikon1 = bishop_values[10]
            mrdElkatholikon2 = bishop_values[11]

            maro = bishopDes_values[8]
            maro1 = bishop_values[4]
            maro2 = bishop_values[5]

            esbasmos1 = bishopDes_values[9]
            esbasmos11 = bishop_values[12]
            esbasmos12 = bishop_values[13]

            esbasmos2 = bishopDes_values[10]
            esbasmos21 = bishop_values[12]
            esbasmos22 = bishop_values[13]

            elaba2basyly = bishopDes_values[11] - 1
            elaba2basyly1 = bishop_values[0]
            elaba2basyly2 = bishop_values[1]

            elaba28yry8ory = bishopDes_values[12] - 1
            elaba28yry8ory1 = bishop_values[0]
            elaba28yry8ory2 = bishop_values[1]

            if guestBishop < 2:
                elshokr2 = elshokr2-1
                tobhyna2 = tobhyna2-2
                embiniot2 = embiniot2-1
                hytynyat2 = hytynyat2-3
                byhmot8ar2 = byhmot8ar2-1
                tomakario2 = tomakario2-1
                eya8aby2 = eya8aby2-2
                mrdElkatholikon2 = mrdElkatholikon2-1
                maro2 = maro2-1
                esbasmos12 = esbasmos12-2
                esbasmos22 = esbasmos22-2
                elaba2basyly2 = elaba2basyly2-1
                elaba28yry8ory2 = elaba28yry8ory2-1
        
            start_positions = [mazmorELtawzy3, elaba28yry8ory, elaba2basyly, esbasmos2, esbasmos1, elengil3, elmazmor3, maro, elebrksis3, elkatholikon3, mrdElkatholikon, elbouls3, eya8aby, tomakario, byhmot8ar, hytynyat, embiniot, tobhyna, elshokr]
            start_slides = [mrdMazmorEltawzy3, elaba28yry8ory1, elaba2basyly1, esbasmos21, esbasmos11, elengil, elmzmor, maro1, elebrksis1, elkatholikon1, mrdElkatholikon1, elbouls1, eya8aby1, tomakario1, byhmot8ar1, hytynyat1, embiniot1, tobhyna1, elshokr1]
            end_slides = [mrdMazmorEltawzy3, elaba28yry8ory2, elaba2basyly2, esbasmos22, esbasmos12, elengil2, elmzmor, maro2, elebrksis2, elkatholikon2, mrdElkatholikon2, elbouls2, eya8aby2, tomakario2, byhmot8ar2, hytynyat2, embiniot2, tobhyna2, elshokr2]

    if guestBishop == 0:
        start_positions = [mazmorELtawzy3, elengil3, elmazmor3, elebrksis3, elkatholikon3, elbouls3]
        start_slides = [mrdMazmorEltawzy3, elengil, elmzmor, elebrksis1, elkatholikon1, elbouls1]
        end_slides = [mrdMazmorEltawzy3, elengil2, elmzmor, elebrksis2, elkatholikon2, elbouls2]

    show_hide_insertImage_replaceText(prs1, excel, des_sheet, som_show_full_sections, som_hide_full_sections, None, som_hide_full_sections_ranges, None, None)
    
    som_show_values.extend([[EsmaElsomElkbyr1, EsmaElsomElkbyr1]])

    powerpoint = win32com.client.Dispatch("PowerPoint.Application")
    powerpoint.Visible = True  # Open PowerPoint application
    presentation1 = open_presentation_relative_path(prs1)
    presentation2 = open_presentation_relative_path(prs2)
    
    if guestBishop>0:
        presentation3 = open_presentation_relative_path(prs3)

    if season == 15.3 or season == 15.4 or season == 15.1 or cd.weekday() == 6 or cd.weekday() == 5:
        khetamElsom = find_slide_index_by_title(presentation1, "احاد الصوم الكبير", elkhetam)
        khetamElsom2 = khetamElsom
    else:
        khetamElsom = find_slide_index_by_title(presentation1, "ايام الصوم الكبير", elkhetam)
        khetamElsom2 = find_slide_index_by_title(presentation1, "ايام الصوم الكبير 2", elkhetam)
    
    show_slides(presentation1, [[khetamElsom, khetamElsom2], [EsmaElsomElkbyr1, EsmaElsomElkbyr1]])

    run_vba_with_slide_id(excel, des_sheet, prs1, presentation1)
    agbya(presentation1, som_values[10], 3)

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

        elif Bishop and current_position in {elaba28yry8ory, elaba2basyly, byhmot8ar, hytynyat, embiniot, 
                                             elshokr, tomakario, eya8aby, esbasmos1, esbasmos2, maro, 
                                             mrdElkatholikon, tobhyna}:
            source_slide = presentation3.Slides(current_end_slide)
            source_slide.Copy()
            if not presentation1.Slides(current_position).SlideShowTransition.Hidden or current_position in {esbasmos1, esbasmos2}:
                presentation1.Slides.Paste(current_position).SlideShowTransition.Hidden = False
            else:
                presentation1.Slides.Paste(current_position).SlideShowTransition.Hidden = True
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

    move_section_names = [
        "{7C69B347-63A6-409B-9C80-6F7AA907197B}",
        "{1CDFD5FF-8DF5-48CF-A773-F5EEBD56469C}",
        "{6BD6B64E-0E1F-4BAC-A7DD-E10B1476E5E3}",
        "{F0C07D6A-9B65-4AF5-AF75-7BA80C1EEEFC}"
    ]

    target_section_names = [
        "{22F83DFC-792B-4148-8AED-E77703B6E7BB}",
        "{7C69B347-63A6-409B-9C80-6F7AA907197B}",
        "{2EB8E5B7-6049-402F-80E4-ED6EDFAB83F4}",
        "{6BD6B64E-0E1F-4BAC-A7DD-E10B1476E5E3}"
    ]
    
    move_sections_v2(presentation1, move_section_names, target_section_names)

    presentation1.SlideShowSettings.Run()
    # Call the function once for all moves

    presentation2.Close()
    if guestBishop>0:
        presentation3.Close()

def odasElbeshara(Bishop=False, guestBishop=0):
    prs1 = relative_path(r"قداس.pptx")  # Using the relative path
    excel = relative_path(r"بيانات القداسات.xlsx")
    excel2 = relative_path(r"Tables.xlsx")
    des_sheet ="القداس"
    replacefile(prs1, relative_path(r"Data\CopyData\قداس.pptx"))

    prs2 = relative_path(r"Data\القطمارس\الايام\القطمارس السنوي ايام (القداس).pptx")
    katamars_sheet = "القطمارس السنوي القداس"
    km, kd = find_Readings_Date(7, 29)
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

    # beshara_show_full_sections = ["الليلويا فاي بيبي", "طاي شوري", "هيتينية الملاك غبريال", "مرد ابركسيس البشارة",
    #                             "مرد مزمور البشارة", "فاي اريه بي اوو" , "مرد انجيل البشارة", 
    #                             "اوشية اهوية السماء", "اوشية اهوية السماء غ", "قسمة عيد البشارة المجيد (نسبح ونمجد إله الآلهة ورب الأرباب)"]
    
    beshara_show_full_sections = ['{072F3D96-A6C8-405F-9A23-7CCA1B2F13FF}', '{20F525FD-C708-4DDD-8E40-FE502EFEBDDE}', '{AE7AA37A-5543-45C2-9921-1F5B7FF26544}', '{6E13B4F2-4D01-4EE4-8763-939CEE1BDB75}', '{AED1945B-2DF8-4D90-BBFD-1D478AF28695}', '{B7D98377-B994-4654-B49C-DE10E0DDE4F1}', '{D2AF7A61-5849-46BB-9FB0-7AA977314D3C}', '{BC7E3DCD-6AA8-44CC-B8AF-BC3E2BC71B5A}', '{A20DA654-32F7-4B4C-96CB-C76232EB96E8}', '{3BCBBDE0-13A1-4825-B7A1-04CDC2E502E4}']
    
    # beshara_hide_full_sections = ["مرد المزمور", "مرد الانجيل", "قسمة القداس الباسيلي (أيها السيد الرب إلهنا)", "اك اسماروؤت", "بي اويك"]
    
    beshara_hide_full_sections = ['{F13A48F2-238D-4617-B84E-9B0A694D9A18}', '{06D592C8-05BF-4B7C-86F7-FDAB3FAB5FB1}', '{681FF6A7-4230-4171-8F41-83FD64E8C960}', '{507EFD97-98F8-4376-848B-20D72E16D2C1}', '{B9A30F5E-0C89-471B-A99A-23DBE7F58504}']
    
    beshara_show_values = []
    beshara_hide_values = []

    beshara_values = ["تكملة على حسب المناسبة", "مزمور التوزيع", "مزمور التوزيع", "مرد توزيع البشارة",
                    "الانجيل", "المزمور", "الابركسيس", "الكاثوليكون", "البولس عربي"]
    beshara_values = find_slide_nums_arrays_v2(excel, des_sheet, 
                    ['{A18EDC94-F257-4FAC-99C7-0A8EA70F0FAF}', '{C29E5A83-A98B-4077-8194-99A6D803EF53}', '{C29E5A83-A98B-4077-8194-99A6D803EF53}', '{2F035E60-4CC5-4808-A906-367BDB0FC9B0}', '{C7D4A109-F792-4661-BAD0-075FD1A1909F}', '{B74DBB8C-2B2D-46E4-9508-DA46008D19A4}', '{E234C6C7-3837-4CE4-A541-CDC9627AAAC2}', '{6D4B3F52-63BF-435F-BF0C-C9D41120C2A3}', '{D88055F5-EAA0-4C8E-8249-C364A572BF7B}'], 
                    2, [1, 1, 2, 1, 1, 2, 2, 2, 2])

    #الختام
    elkhetam = beshara_values[0]
    #التوزيع
    mazmorELtawzy3 = beshara_values[1] + 1
    mazmorELtawzy32 = beshara_values[2] - 1
    mrdMazmorEltawzy3 = beshara_values[3]
    #القرائات
    elengil3 = beshara_values[4]
    elmazmor3 = beshara_values[5]
    elebrksis3 = beshara_values[6]
    elkatholikon3 = beshara_values[7]
    elbouls3 = beshara_values[8]

    if Bishop == True:
        prs3 = relative_path(r"Data\حضور الأسقف.pptx")
        sheet = "في حضور الأسقف"

        # bishop_show_values = ['في حضور الاسقف', 'المزمور', 'مارو اتشاسف', 'امبين يوت اتطايوت', 
        #                       'ني سافيف تيرو', 'تكملة في حضور الاسقف', 'لحن اك اسمارؤوت']

        #bishop_hide_values = ['سوتيس امين']

        bishop_show_values = ['{2BCF4F8C-25F0-43C5-B224-6528B2EA3F2F}', '{F76B0D75-0474-45B5-B79F-7416F354543A}',
                              '{E2968C91-5339-499C-9812-DECCCF58A2CD}', '{62A12AF8-CB6D-4CC5-9DB0-B73A7C24E2AD}', 
                              '{B74DBB8C-2B2D-46E4-9508-DA46008D19A4}', '{A9183893-7B7E-459F-8547-F7A8F7D2D521}', 
                              '{670DAA94-A6C9-4CCD-B4E2-958C71CD3E44}']
        
        bishop_hide_values = ['{4D2B15D5-C978-467C-9D6C-726FE25128B8}']
        
        beshara_show_full_sections.extend(bishop_show_values)
        beshara_hide_full_sections.extend(bishop_hide_values)

        if guestBishop > 0:
            # bishop_values = ["صلاة الشكر", "صلاة الشكر", "طوبه هينا الكبيرة", "طوبه هينا الكبيرة", 
            #                  "نيم بينيوت", "نيم بينيوت", "الهيتنيات", "الهيتنيات",
            #                  "اي اغابي", "اي اغابي", "مرد الكاثوليكون", "مرد الكاثوليكون",
            #                  "الاسبسمس", "الاسبسمس"]

            # bishopDes_values = ["تكملة في حضور الاسقف", "طوبه هينا الكبيرة",
            #                     "امبين يوت اتطايوت", "تكملة الهيتنيات", "بى اهموت غار الصغيرة",
            #                     "تو ماكريو", "اي اغابي", "ابيت جيك ايفول", "مارو اتشاسف",
            #                     "الاسبسمس الادام السنوي+الختام", "ختام الاسبسمس الادام",
            #                     "اوشية الاباء (ب)", "اوشية الاباء غ"]
            
            bishop_values = find_slide_nums_arrays_v2(excel, sheet, 
                            ['{6851F163-CBEF-4014-A853-CE100557BA6A}', '{6851F163-CBEF-4014-A853-CE100557BA6A}', '{B084BC40-61E1-4477-98DA-15CFB06AEE91}', '{B084BC40-61E1-4477-98DA-15CFB06AEE91}', '{97203297-EECB-4D41-B2E3-AD9A4863847E}', '{97203297-EECB-4D41-B2E3-AD9A4863847E}', '{7C84083F-E6D3-4669-9130-AC7E8D935A98}', '{7C84083F-E6D3-4669-9130-AC7E8D935A98}', '{D0234C99-69FE-407A-82A9-D7A676919E93}', '{D0234C99-69FE-407A-82A9-D7A676919E93}', '{0B519645-F935-43AE-A9CE-6E2FC03833BB}', '{0B519645-F935-43AE-A9CE-6E2FC03833BB}', '{D1378DB5-29D1-4800-9D96-10F2535EEB57}', '{D1378DB5-29D1-4800-9D96-10F2535EEB57}'], 
                            2, [1, 2, 1, 2, 1, 2, 1, 2, 1, 2, 1, 2, 1, 2])

            bishopDes_values = find_slide_nums_arrays_v2(excel, des_sheet, 
                               ['{F76B0D75-0474-45B5-B79F-7416F354543A}', '{8DD21CDE-CB6B-4D5B-B995-D2747AB69ED1}', '{E2968C91-5339-499C-9812-DECCCF58A2CD}', '{646A8184-7F05-453A-A2F1-EB9A77D7F0EE}', '{21891CBB-A1EC-4974-B0B6-F74A4B502BC2}', '{12B7D244-BF4C-401B-A65A-D1621D7DD953}', '{F69B50D8-FB5E-4E8C-AE5C-6DCB4790AFAF}', '{38267404-7625-47AA-B0C8-31BCA5D0435D}', '{62A12AF8-CB6D-4CC5-9DB0-B73A7C24E2AD}', '{0C7A7725-643D-4F7E-A2F0-0C8A36C2A594}', '{1CDFD5FF-8DF5-48CF-A773-F5EEBD56469C}', '{74A33555-8E08-47DF-B3CD-A1B4C7AF2B4E}', '{474487EB-4554-4A30-B351-6EF762D2F2D6}'],                  
                               2, [2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2])

            elshokr = bishopDes_values[0]
            elshokr1 = bishop_values[0]
            elshokr2 = bishop_values[1]

            tobhyna = bishopDes_values[1]
            tobhyna1 = bishop_values[2]
            tobhyna2 = bishop_values[3]

            embiniot = bishopDes_values[2]
            embiniot1 = bishop_values[4]
            embiniot2 = bishop_values[5]

            hytynyat = bishopDes_values[3]
            hytynyat1 = bishop_values[6]
            hytynyat2 = bishop_values[7]

            byhmot8ar = bishopDes_values[4]-1
            byhmot8ar1 = bishop_values[4]
            byhmot8ar2 = bishop_values[5]

            tomakario = bishopDes_values[5] - 1
            tomakario1 = bishop_values[4]
            tomakario2 = bishop_values[5]

            eya8aby = bishopDes_values[6] - 1
            eya8aby1 = bishop_values[8]
            eya8aby2 = bishop_values[9]

            mrdElkatholikon = bishopDes_values[7]
            mrdElkatholikon1 = bishop_values[10]
            mrdElkatholikon2 = bishop_values[11]

            maro = bishopDes_values[8]
            maro1 = bishop_values[4]
            maro2 = bishop_values[5]

            esbasmos1 = bishopDes_values[9]
            esbasmos11 = bishop_values[12]
            esbasmos12 = bishop_values[13]

            esbasmos2 = bishopDes_values[10]
            esbasmos21 = bishop_values[12]
            esbasmos22 = bishop_values[13]

            elaba2basyly = bishopDes_values[11] - 1
            elaba2basyly1 = bishop_values[0]
            elaba2basyly2 = bishop_values[1]

            elaba28yry8ory = bishopDes_values[12] - 1
            elaba28yry8ory1 = bishop_values[0]
            elaba28yry8ory2 = bishop_values[1]

            if guestBishop < 2:
                elshokr2 = elshokr2-1
                tobhyna2 = tobhyna2-2
                embiniot2 = embiniot2-1
                hytynyat2 = hytynyat2-3
                byhmot8ar2 = byhmot8ar2-1
                tomakario2 = tomakario2-1
                eya8aby2 = eya8aby2-2
                mrdElkatholikon2 = mrdElkatholikon2-1
                maro2 = maro2-1
                esbasmos12 = esbasmos12-2
                esbasmos22 = esbasmos22-2
                elaba2basyly2 = elaba2basyly2-1
                elaba28yry8ory2 = elaba28yry8ory2-1
        
            start_positions = [elaba28yry8ory, elaba2basyly, esbasmos2, esbasmos1, elengil3, elmazmor3, maro, elebrksis3, elkatholikon3, mrdElkatholikon, elbouls3, eya8aby, tomakario, byhmot8ar, hytynyat, embiniot, tobhyna, elshokr]
            start_slides = [elaba28yry8ory1, elaba2basyly1, esbasmos21, esbasmos11, elengil, elmzmor, maro1, elebrksis1, elkatholikon1, mrdElkatholikon1, elbouls1, eya8aby1, tomakario1, byhmot8ar1, hytynyat1, embiniot1, tobhyna1, elshokr1]
            end_slides = [elaba28yry8ory2, elaba2basyly2, esbasmos22, esbasmos12, elengil2, elmzmor, maro2, elebrksis2, elkatholikon2, mrdElkatholikon2, elbouls2, eya8aby2, tomakario2, byhmot8ar2, hytynyat2, embiniot2, tobhyna2, elshokr2]

    if guestBishop == 0:
        start_positions = [mazmorELtawzy3, elengil3, elmazmor3, elebrksis3, elkatholikon3, elbouls3]
        start_slides = [mrdMazmorEltawzy3, elengil, elmzmor, elebrksis1, elkatholikon1, elbouls1]
        end_slides = [mrdMazmorEltawzy3, elengil2, elmzmor, elebrksis2, elkatholikon2, elbouls2]

    show_hide_insertImage_replaceText(prs1, excel, des_sheet, beshara_show_full_sections, beshara_hide_full_sections)

    powerpoint = win32com.client.Dispatch("PowerPoint.Application")
    powerpoint.Visible = True  # Open PowerPoint application
    presentation1 = open_presentation_relative_path(prs1)
    presentation2 = open_presentation_relative_path(prs2)

    if guestBishop > 0:
        presentation3 = open_presentation_relative_path(prs3)

    khetamElbeshara = find_slide_index_by_title(presentation1, "البشارة", elkhetam)
    show_slides(presentation1, [[khetamElbeshara, khetamElbeshara]])

    run_vba_with_slide_id(excel, des_sheet, prs1, presentation1)
    
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
                new_slide = presentation1.Slides.Paste(slide_index1).SlideShowTransition.Hidden = False
                slide_index1 -= 1
            current_start_slide += 1

        elif (current_position == elengil3 or current_position == elmazmor3 or current_position == elebrksis3 
            or current_position == elkatholikon3 or current_position == elbouls3):
            source_slide = presentation2.Slides(current_end_slide)
            source_slide.Copy()
            new_slide = presentation1.Slides.Paste(current_position)
            new_slide.SlideShowTransition.Hidden = False
            current_end_slide -= 1
            if current_start_slide > current_end_slide:
                current_position += 1

        elif Bishop and current_position in {elaba28yry8ory, elaba2basyly, byhmot8ar, hytynyat, embiniot, 
                                             elshokr, tomakario, eya8aby, esbasmos1, esbasmos2, maro, 
                                             mrdElkatholikon, tobhyna}:
            source_slide = presentation3.Slides(current_end_slide)
            source_slide.Copy()
            if current_position in {elaba28yry8ory, elaba2basyly, byhmot8ar, hytynyat, embiniot, elshokr, maro}:
                presentation1.Slides.Paste(current_position).SlideShowTransition.Hidden = False
            else:
                presentation1.Slides.Paste(current_position).SlideShowTransition.Hidden = True
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

    move_section_names = [
        "{0C416407-7559-4449-BC08-E161F996B9F0}",
        "{1CDFD5FF-8DF5-48CF-A773-F5EEBD56469C}",
        "{45FA414E-7385-4E3A-9163-587A152CA1F8}", 
        "{F0C07D6A-9B65-4AF5-AF75-7BA80C1EEEFC}"
    ]

    target_section_names = [
        "{22F83DFC-792B-4148-8AED-E77703B6E7BB}", 
        "{0C416407-7559-4449-BC08-E161F996B9F0}",
        "{2EB8E5B7-6049-402F-80E4-ED6EDFAB83F4}",
        "{45FA414E-7385-4E3A-9163-587A152CA1F8}"
    ]

    # Call the function once for all moves
    move_sections_v2(presentation1, move_section_names, target_section_names)

    presentation2.Close()
    if guestBishop > 0:
        presentation3.Close()

    presentation1.SlideShowSettings.Run()

def odasSbtLe3azr(copticdate, Bishop=False, guestBishop=0):
    prs1 = relative_path(r"قداس.pptx")  # Using the relative path
    excel = relative_path(r"بيانات القداسات.xlsx")
    excel2 = relative_path(r"Tables.xlsx")
    des_sheet ="القداس"
    replacefile(prs1, relative_path(r"Data\CopyData\قداس.pptx"))

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

    # le3azr_show_full_sections = ["الليلويا فاي بيبي", "طاي شوري", "هيتينية لعازر", "مرد ابركسيس سبت لعازر", 
    #                              "مرد انجيل سبت لعازر", "اوشية اهوية السماء", "اوشية اهوية السماء غ", "مدائح سبت لعازر"]
    # le3azr_hide_full_sections = ["مرد الانجيل", "اك اسماروؤت", "بي اويك"]

    le3azr_show_full_sections = ['{072F3D96-A6C8-405F-9A23-7CCA1B2F13FF}', '{20F525FD-C708-4DDD-8E40-FE502EFEBDDE}', '{A5C57AE9-87E9-4572-B4B1-1D93A609BA47}', '{45AAEE9B-78C4-425B-A0B2-7C063F95B34F}', '{7916CBAB-6268-4AAA-980D-3F75DBAB5F82}', '{BC7E3DCD-6AA8-44CC-B8AF-BC3E2BC71B5A}', '{A20DA654-32F7-4B4C-96CB-C76232EB96E8}', '{DB6B4ECA-F770-40E7-A7BA-46A8B5B46C3B}']
    le3azr_hide_full_sections = ['{06D592C8-05BF-4B7C-86F7-FDAB3FAB5FB1}', '{507EFD97-98F8-4376-848B-20D72E16D2C1}', '{B9A30F5E-0C89-471B-A99A-23DBE7F58504}']

    # le3azr_values =  ["تكملة على حسب المناسبة", "لازاروس", "الانجيل", "المزمور", "الابركسيس", "الكاثوليكون", "البولس عربي"]
    le3azr_values = find_slide_nums_arrays_v2(excel, des_sheet, 
                                      ['{A18EDC94-F257-4FAC-99C7-0A8EA70F0FAF}', '{78C64C43-B3F6-4110-8CB0-1E10086C77C2}', '{C7D4A109-F792-4661-BAD0-075FD1A1909F}', '{B74DBB8C-2B2D-46E4-9508-DA46008D19A4}', '{E234C6C7-3837-4CE4-A541-CDC9627AAAC2}', '{6D4B3F52-63BF-435F-BF0C-C9D41120C2A3}', '{D88055F5-EAA0-4C8E-8249-C364A572BF7B}'],
                                      2, [1, 2, 2, 2, 2, 2, 2])

    #الختام
    elkhetam = le3azr_values[0]
    #التوزيع
    lazaros = le3azr_values[1]
    #القرائات
    elengil3 = le3azr_values[2]
    elmazmor3 = le3azr_values[3]
    elebrksis3 = le3azr_values[4]
    elkatholikon3 = le3azr_values[5]
    elbouls3 = le3azr_values[6]

    if Bishop == True:
        prs3 = relative_path(r"Data\حضور الأسقف.pptx")
        sheet = "في حضور الأسقف"

        # bishop_show_values = ['في حضور الاسقف', 'المزمور', 'مارو اتشاسف', 'امبين يوت اتطايوت', 
        #                       'ني سافيف تيرو', 'تكملة في حضور الاسقف', 'لحن اك اسمارؤوت']

        #bishop_hide_values = ['سوتيس امين']

        bishop_show_values = ['{2BCF4F8C-25F0-43C5-B224-6528B2EA3F2F}', '{F76B0D75-0474-45B5-B79F-7416F354543A}',
                              '{E2968C91-5339-499C-9812-DECCCF58A2CD}', '{62A12AF8-CB6D-4CC5-9DB0-B73A7C24E2AD}', 
                              '{B74DBB8C-2B2D-46E4-9508-DA46008D19A4}', '{A9183893-7B7E-459F-8547-F7A8F7D2D521}', 
                              '{670DAA94-A6C9-4CCD-B4E2-958C71CD3E44}']
        
        bishop_hide_values = ['{4D2B15D5-C978-467C-9D6C-726FE25128B8}']
        
        le3azr_show_full_sections.extend(bishop_show_values)
        le3azr_hide_full_sections.extend(bishop_hide_values)

        if guestBishop > 0:
            # bishop_values = ["صلاة الشكر", "صلاة الشكر", "طوبه هينا الكبيرة", "طوبه هينا الكبيرة", 
            #                  "نيم بينيوت", "نيم بينيوت", "الهيتنيات", "الهيتنيات",
            #                  "اي اغابي", "اي اغابي", "مرد الكاثوليكون", "مرد الكاثوليكون",
            #                  "الاسبسمس", "الاسبسمس"]

            # bishopDes_values = ["تكملة في حضور الاسقف", "طوبه هينا الكبيرة",
            #                     "امبين يوت اتطايوت", "تكملة الهيتنيات", "بى اهموت غار الصغيرة",
            #                     "تو ماكريو", "اي اغابي", "ابيت جيك ايفول", "مارو اتشاسف",
            #                     "الاسبسمس الادام السنوي+الختام", "ختام الاسبسمس الادام",
            #                     "اوشية الاباء (ب)", "اوشية الاباء غ"]
            
            bishop_values = find_slide_nums_arrays_v2(excel, sheet, 
                            ['{6851F163-CBEF-4014-A853-CE100557BA6A}', '{6851F163-CBEF-4014-A853-CE100557BA6A}', '{B084BC40-61E1-4477-98DA-15CFB06AEE91}', '{B084BC40-61E1-4477-98DA-15CFB06AEE91}', '{97203297-EECB-4D41-B2E3-AD9A4863847E}', '{97203297-EECB-4D41-B2E3-AD9A4863847E}', '{7C84083F-E6D3-4669-9130-AC7E8D935A98}', '{7C84083F-E6D3-4669-9130-AC7E8D935A98}', '{D0234C99-69FE-407A-82A9-D7A676919E93}', '{D0234C99-69FE-407A-82A9-D7A676919E93}', '{0B519645-F935-43AE-A9CE-6E2FC03833BB}', '{0B519645-F935-43AE-A9CE-6E2FC03833BB}', '{D1378DB5-29D1-4800-9D96-10F2535EEB57}', '{D1378DB5-29D1-4800-9D96-10F2535EEB57}'], 
                            2, [1, 2, 1, 2, 1, 2, 1, 2, 1, 2, 1, 2, 1, 2])

            bishopDes_values = find_slide_nums_arrays_v2(excel, des_sheet, 
                               ['{F76B0D75-0474-45B5-B79F-7416F354543A}', '{8DD21CDE-CB6B-4D5B-B995-D2747AB69ED1}', '{E2968C91-5339-499C-9812-DECCCF58A2CD}', '{646A8184-7F05-453A-A2F1-EB9A77D7F0EE}', '{21891CBB-A1EC-4974-B0B6-F74A4B502BC2}', '{12B7D244-BF4C-401B-A65A-D1621D7DD953}', '{F69B50D8-FB5E-4E8C-AE5C-6DCB4790AFAF}', '{38267404-7625-47AA-B0C8-31BCA5D0435D}', '{62A12AF8-CB6D-4CC5-9DB0-B73A7C24E2AD}', '{0C7A7725-643D-4F7E-A2F0-0C8A36C2A594}', '{1CDFD5FF-8DF5-48CF-A773-F5EEBD56469C}', '{74A33555-8E08-47DF-B3CD-A1B4C7AF2B4E}', '{474487EB-4554-4A30-B351-6EF762D2F2D6}'],                  
                               2, [2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2])

            elshokr = bishopDes_values[0]
            elshokr1 = bishop_values[0]
            elshokr2 = bishop_values[1]

            tobhyna = bishopDes_values[1]
            tobhyna1 = bishop_values[2]
            tobhyna2 = bishop_values[3]

            embiniot = bishopDes_values[2]
            embiniot1 = bishop_values[4]
            embiniot2 = bishop_values[5]

            hytynyat = bishopDes_values[3]
            hytynyat1 = bishop_values[6]
            hytynyat2 = bishop_values[7]

            byhmot8ar = bishopDes_values[4]-1
            byhmot8ar1 = bishop_values[4]
            byhmot8ar2 = bishop_values[5]

            tomakario = bishopDes_values[5] - 1
            tomakario1 = bishop_values[4]
            tomakario2 = bishop_values[5]

            eya8aby = bishopDes_values[6] - 1
            eya8aby1 = bishop_values[8]
            eya8aby2 = bishop_values[9]

            mrdElkatholikon = bishopDes_values[7]
            mrdElkatholikon1 = bishop_values[10]
            mrdElkatholikon2 = bishop_values[11]

            maro = bishopDes_values[8]
            maro1 = bishop_values[4]
            maro2 = bishop_values[5]

            esbasmos1 = bishopDes_values[9]
            esbasmos11 = bishop_values[12]
            esbasmos12 = bishop_values[13]

            esbasmos2 = bishopDes_values[10]
            esbasmos21 = bishop_values[12]
            esbasmos22 = bishop_values[13]

            elaba2basyly = bishopDes_values[11] - 1
            elaba2basyly1 = bishop_values[0]
            elaba2basyly2 = bishop_values[1]

            elaba28yry8ory = bishopDes_values[12] - 1
            elaba28yry8ory1 = bishop_values[0]
            elaba28yry8ory2 = bishop_values[1]

            if guestBishop < 2:
                elshokr2 = elshokr2-1
                tobhyna2 = tobhyna2-2
                embiniot2 = embiniot2-1
                hytynyat2 = hytynyat2-3
                byhmot8ar2 = byhmot8ar2-1
                tomakario2 = tomakario2-1
                eya8aby2 = eya8aby2-2
                mrdElkatholikon2 = mrdElkatholikon2-1
                maro2 = maro2-1
                esbasmos12 = esbasmos12-2
                esbasmos22 = esbasmos22-2
                elaba2basyly2 = elaba2basyly2-1
                elaba28yry8ory2 = elaba28yry8ory2-1
        
            start_positions = [elaba28yry8ory, elaba2basyly, esbasmos2, esbasmos1, elengil3, elmazmor3, maro, elebrksis3, elkatholikon3, mrdElkatholikon, elbouls3, eya8aby, tomakario, byhmot8ar, hytynyat, embiniot, tobhyna, elshokr]
            start_slides = [elaba28yry8ory1, elaba2basyly1, esbasmos21, esbasmos11, elengil, elmzmor, maro1, elebrksis1, elkatholikon1, mrdElkatholikon1, elbouls1, eya8aby1, tomakario1, byhmot8ar1, hytynyat1, embiniot1, tobhyna1, elshokr1]
            end_slides = [elaba28yry8ory2, elaba2basyly2, esbasmos22, esbasmos12, elengil2, elmzmor, maro2, elebrksis2, elkatholikon2, mrdElkatholikon2, elbouls2, eya8aby2, tomakario2, byhmot8ar2, hytynyat2, embiniot2, tobhyna2, elshokr2]

    if guestBishop == 0:
        start_positions = [elengil3, elmazmor3, elebrksis3, elkatholikon3, elbouls3]
        start_slides = [elengil, elmzmor, elebrksis1, elkatholikon1, elbouls1]
        end_slides = [elengil2, elmzmor, elebrksis2, elkatholikon2, elbouls2]

    show_hide_insertImage_replaceText(prs1, excel, des_sheet, le3azr_show_full_sections, le3azr_hide_full_sections)

    powerpoint = win32com.client.Dispatch("PowerPoint.Application")
    powerpoint.Visible = True  # Open PowerPoint application
    presentation1 = open_presentation_relative_path(prs1)
    presentation2 = open_presentation_relative_path(prs2)

    if guestBishop > 0:
        presentation3 = open_presentation_relative_path(prs3)

    khetamle3azr = find_slide_index_by_title(presentation1, "الاهوية", elkhetam)
    show_slides(presentation1, [[khetamle3azr, khetamle3azr], [lazaros, lazaros]])

    run_vba_with_slide_id(excel, des_sheet, prs1, presentation1)

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
            new_slide = presentation1.Slides.Paste(current_position)
            new_slide.SlideShowTransition.Hidden = False
            current_end_slide -= 1
            if current_start_slide > current_end_slide:
                current_position += 1

        elif Bishop and current_position in {elaba28yry8ory, elaba2basyly, byhmot8ar, hytynyat, embiniot, 
                                             elshokr, tomakario, eya8aby, esbasmos1, esbasmos2, maro, 
                                             mrdElkatholikon, tobhyna}:
            source_slide = presentation3.Slides(current_end_slide)
            source_slide.Copy()
            if current_position in {elaba28yry8ory, elaba2basyly, byhmot8ar, hytynyat, embiniot, elshokr, maro}:
                presentation1.Slides.Paste(current_position).SlideShowTransition.Hidden = False
            else:
                presentation1.Slides.Paste(current_position).SlideShowTransition.Hidden = True
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

    move_section_names = [
        "{7D07F022-6B6E-4F94-BF58-0C259920A0B8}",
        "{1CDFD5FF-8DF5-48CF-A773-F5EEBD56469C}",
    ]

    target_section_names = [
        "{22F83DFC-792B-4148-8AED-E77703B6E7BB}", 
        "{7D07F022-6B6E-4F94-BF58-0C259920A0B8}",
    ]

    # Call the function once for all moves
    move_sections_v2(presentation1, move_section_names, target_section_names)

    presentation2.Close()
    if guestBishop > 0:
        presentation3.Close()

def odasElsh3anyn(copticdate, Bishop=False, guestBishop=0):
    prs1 = relative_path(r"قداس.pptx")  # Using the relative path
    excel = relative_path(r"بيانات القداسات.xlsx")
    excel2 = relative_path(r"Tables.xlsx")
    des_sheet ="القداس"
    image = relative_path(r"Data\Designs\الشعانين.png")
    replacefile(prs1, relative_path(r"Data\CopyData\قداس.pptx"))

    prs2 = relative_path(r"Data\القطمارس\قرائات احد الشعانين.pptx")
    katamars_sheet = "قرائات أحد الشعانين"
    km = copticdate[1]
    kd = copticdate[2]
    katamars_offsets = [9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21]
    katamars_values = fetch_data_arrays(excel2, katamars_sheet, km, kd, katamars_offsets)
    elbouls1 = katamars_values[0]
    elkatholikon1 = katamars_values[1]
    elebrksis1 = katamars_values[2]
    elmazmorElawl = katamars_values[3]
    elengilElawl = katamars_values[4]
    elengilElawl2 = katamars_values[5]
    elengilEltany = katamars_values[6]
    elengilEltany2 = katamars_values[7]
    elengilEltalt = katamars_values[8]
    elengilEltalt2 = katamars_values[9]
    elmazmorEltany = katamars_values[10]
    elengilElrab3 = katamars_values[11]
    elengilElrab32 = katamars_values[12]
    elbouls2 = elkatholikon1 - 1
    elkatholikon2 = elebrksis1 - 1
    elebrksis2 = elmazmorElawl - 1
    
    # elsh3anyn_show_full_sections = ["مدائح احد الشعانين", "قسمة تُقال في احد الشعانين (أيها الرب مثل ربنا مثل عجيب صار إسمك)", 
    #                                 "اوشية اهوية السماء غ", "اوشية اهوية السماء", "الانجيل الثالث الشعانين", "مرد الانجيل الثاني الشعانين", 
    #                                 "الانجيل الثاني الشعانين", "مرد الانجيل الاول الشعانين", "مرد مزمور الشعانين", "مزمور الشعانين الثاني قبطي",
    #                                 "مزمور الشعانين الأول قبطي", "المزمور السنجاري", "محير احد الشعانين", "إڤلوجيمينوس", 
    #                                 "مرد ابركسيس الشعانين", "طاي شوري", "ني سافيف تيرو", "الليلويا فاي بيبي"
    #                                 ]

    # elsh3anyn_hide_full_sections = [ "بي اويك", "اك اسماروؤت", "قسمة القداس الباسيلي (أيها السيد الرب إلهنا)", "ربع للعذراء", "مرد الانجيل", "مرد المزمور", "اكسيا", "السنكسار"]

    # elsh3anyn_show_full_sections_ranges = [["مرد الانجيل الثالث الشعانين", "مرد الانجيل الرابع الشعانين"],
    #                                        ["الهيتنيات المجمعة - الشهداء", "هيتينية حبيب جرجس"]]

    # elsh3anyn_hide_full_sections_ranges = [["هيتينية مارجرجس", "هيتينية مارمينا"], 
    #                                        ["مقدمة قانون الإيمان", "تقدمة الحمل"]]
    

    elsh3anyn_show_full_sections = ['{37DC4920-98DA-477A-A6F7-6D252B149A22}', '{FA39B8EF-3898-43AD-B4E9-1E0FE167AE86}', '{A20DA654-32F7-4B4C-96CB-C76232EB96E8}', '{BC7E3DCD-6AA8-44CC-B8AF-BC3E2BC71B5A}', '{97A82C93-DF39-4174-BFDD-05AEB4077FC3}', '{4AC0E349-8541-4272-8D5D-2DA7A3D0A13F}', '{F05E7BBA-1647-4E14-ABEF-90F9D32C5C14}', '{5367445D-8EA1-4DA6-9BFF-CAD85A20EB6D}', '{C53F4E58-F040-44BA-BDD7-FA62E74EE4C8}', '{0E543BA7-09D4-4CB8-8101-0BDC26CBCA40}', '{92F4141C-A127-4893-9542-743F62A24C83}', '{973DBBAC-E645-4981-B3F5-5DB1413508D0}', '{BC4B57FE-428D-453C-B07D-496033A38DB7}', '{63FA65C9-ACCF-4DDC-A42D-A971CE98C581}', '{4507D778-EA7C-439C-B8F4-B167E7780905}', '{20F525FD-C708-4DDD-8E40-FE502EFEBDDE}', '{670DAA94-A6C9-4CCD-B4E2-958C71CD3E44}', '{072F3D96-A6C8-405F-9A23-7CCA1B2F13FF}']

    elsh3anyn_hide_full_sections = ['{B9A30F5E-0C89-471B-A99A-23DBE7F58504}', '{507EFD97-98F8-4376-848B-20D72E16D2C1}', '{681FF6A7-4230-4171-8F41-83FD64E8C960}', '{ECE652ED-1345-4C6D-B92D-5996CFA27AEE}', '{06D592C8-05BF-4B7C-86F7-FDAB3FAB5FB1}', '{F13A48F2-238D-4617-B84E-9B0A694D9A18}', '{A8A52E1F-44DD-45E1-A737-4E13E15D5F1F}', '{1DA5C6AA-2FE5-461B-9E2C-40113CFC7804}']

    elsh3anyn_show_full_sections_ranges = [['{BCCE6150-258E-4CB2-B766-7A6FAE68B9EC}', '{9E224120-D518-4DE8-BFAC-4AB2F0AA5414}'],
                                           ['{C2140BA5-149F-49BF-A4F6-F0205CF4B99C}', '{457CE0B0-6F83-4DF4-8166-692625329991}']]

    elsh3anyn_hide_full_sections_ranges = [['{735C1803-8AE4-44C8-A4EA-C7B3C1493312}', '{D05DF2AC-49C5-413A-A8AC-2D17C1DDFF97}'], 
                                           ['{409B3D0A-B40A-4475-811D-72C5125134AB}', '{FB902AC1-5581-4552-B503-116755E9D9A8}']]


    # elsha3anyn_values = ["تكملة على حسب المناسبة", "مزمور التوزيع", 
    #                      "مزمور التوزيع", "مرد توزيع الشعانين",
    #                      "الانجيل الرابع الشعانين", "المزمور الثاني الشعانين",
    #                      "الانجيل الثالث الشعانين", "الانجيل الثاني الشعانين", "الانجيل", "المزمور",
    #                      "الابركسيس", "الكاثوليكون", "البولس عربي", "اللي القربان"]
    
    elsh3anyn_values = find_slide_nums_arrays_v2(excel, des_sheet,
                       ['{A18EDC94-F257-4FAC-99C7-0A8EA70F0FAF}', '{C29E5A83-A98B-4077-8194-99A6D803EF53}', '{C29E5A83-A98B-4077-8194-99A6D803EF53}', '{C95D30EC-C174-40B5-8671-08FA653E3199}', '{21F90FE7-A154-49E1-9C55-5E9AC93E43BA}', '{F9886962-A539-40A8-95D6-122DAFEC2303}', '{97A82C93-DF39-4174-BFDD-05AEB4077FC3}', '{F05E7BBA-1647-4E14-ABEF-90F9D32C5C14}', '{C7D4A109-F792-4661-BAD0-075FD1A1909F}', '{B74DBB8C-2B2D-46E4-9508-DA46008D19A4}', '{E234C6C7-3837-4CE4-A541-CDC9627AAAC2}', '{6D4B3F52-63BF-435F-BF0C-C9D41120C2A3}', '{D88055F5-EAA0-4C8E-8249-C364A572BF7B}', '{BBBAC16F-044D-4F33-8068-620F498B59CD}'],
                       2,[1, 1, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2])

    #الختام
    elkhetam = elsh3anyn_values[0]

    #التوزيع
    mazmorELtawzy3 = elsh3anyn_values[1] + 1
    mazmorELtawzy32 = elsh3anyn_values[2] - 1
    mrdMazmorEltawzy3 = elsh3anyn_values[3]

    #الانجيل الرابع الشعانين
    elengilElrab33 = elsh3anyn_values[4]

    #المزمور الثاني الشعانين
    elmazmorEltany2 = elsh3anyn_values[5]

    #الانجيل الثالث الشعانين
    elengilEltalt3 = elsh3anyn_values[6]

    #الانجيل الثاني الشعانين
    elengilEltany3 = elsh3anyn_values[7]

    #الانجيل
    elengilElawl3 = elsh3anyn_values[8]

    #المزمور
    elmazmorElawl2 = elsh3anyn_values[9]

    #الابركسيس
    elebrksis3 = elsh3anyn_values[10]

    #الكاثوليكون
    elkatholikon3 = elsh3anyn_values[11]

    #البولس عربي
    elbouls3 = elsh3anyn_values[12]

    #اللي القربان
    allyElqurban = elsh3anyn_values[13]

    if Bishop == True:
        prs3 = relative_path(r"Data\حضور الأسقف.pptx")
        sheet = "في حضور الأسقف"

        # bishop_show_values = ['في حضور الاسقف', 'فليرفعوه', 'مارو اتشاسف', 'امبين يوت اتطايوت', 
        #                       'ني سافيف تيرو', 'تكملة في حضور الاسقف']

        #bishop_hide_values = ['سوتيس امين']

        bishop_show_values = ['{A9183893-7B7E-459F-8547-F7A8F7D2D521}', '{600D2394-BB87-4926-A5A1-24D17F10DD49}', 
                              '{62A12AF8-CB6D-4CC5-9DB0-B73A7C24E2AD}', '{E2968C91-5339-499C-9812-DECCCF58A2CD}', 
                              '{670DAA94-A6C9-4CCD-B4E2-958C71CD3E44}', '{F76B0D75-0474-45B5-B79F-7416F354543A}']
        
        bishop_hide_values = ['{4D2B15D5-C978-467C-9D6C-726FE25128B8}']
        
        elsh3anyn_show_full_sections.extend(bishop_show_values)
        elsh3anyn_hide_full_sections.extend(bishop_hide_values)

        # bishopDes_values_for_one_bishop_only = ["الانجيل الرابع الشعانين", "فليرفعوه", "فليرفعوه"]

        bishopDes_values_for_one_bishop_only = find_slide_nums_arrays_v2(excel, des_sheet,
                                               ['{21F90FE7-A154-49E1-9C55-5E9AC93E43BA}', '{600D2394-BB87-4926-A5A1-24D17F10DD49}', '{600D2394-BB87-4926-A5A1-24D17F10DD49}'], 
                                               2, [1, 1, 2])

        falyrf3oh = bishopDes_values_for_one_bishop_only[0]
        falyrf3oh1 = bishopDes_values_for_one_bishop_only[1]
        falyrf3oh2 = bishopDes_values_for_one_bishop_only[2]

        if guestBishop > 0:
            # bishop_values = ["صلاة الشكر", "صلاة الشكر", "طوبه هينا الكبيرة", "طوبه هينا الكبيرة", 
            #                  "نيم بينيوت", "نيم بينيوت", "الهيتنيات", "الهيتنيات",
            #                  "اي اغابي", "اي اغابي", "مرد الكاثوليكون", "مرد الكاثوليكون",
            #                  "الاسبسمس", "الاسبسمس"]

            # bishopDes_values = ["تكملة في حضور الاسقف", "طوبه هينا الكبيرة",
            #                     "امبين يوت اتطايوت", "تكملة الهيتنيات", "بى اهموت غار الصغيرة",
            #                     "تو ماكريو", "اي اغابي", "ابيت جيك ايفول", "مارو اتشاسف",
            #                     "الاسبسمس الادام السنوي+الختام", "ختام الاسبسمس الادام",
            #                     "اوشية الاباء (ب)", "اوشية الاباء غ"]
            
            bishop_values = find_slide_nums_arrays_v2(excel, sheet, 
                            ['{6851F163-CBEF-4014-A853-CE100557BA6A}', '{6851F163-CBEF-4014-A853-CE100557BA6A}', '{B084BC40-61E1-4477-98DA-15CFB06AEE91}', '{B084BC40-61E1-4477-98DA-15CFB06AEE91}', '{97203297-EECB-4D41-B2E3-AD9A4863847E}', '{97203297-EECB-4D41-B2E3-AD9A4863847E}', '{7C84083F-E6D3-4669-9130-AC7E8D935A98}', '{7C84083F-E6D3-4669-9130-AC7E8D935A98}', '{D0234C99-69FE-407A-82A9-D7A676919E93}', '{D0234C99-69FE-407A-82A9-D7A676919E93}', '{0B519645-F935-43AE-A9CE-6E2FC03833BB}', '{0B519645-F935-43AE-A9CE-6E2FC03833BB}', '{D1378DB5-29D1-4800-9D96-10F2535EEB57}', '{D1378DB5-29D1-4800-9D96-10F2535EEB57}'], 
                            2, [1, 2, 1, 2, 1, 2, 1, 2, 1, 2, 1, 2, 1, 2])

            bishopDes_values = find_slide_nums_arrays_v2(excel, des_sheet, 
                               ['{F76B0D75-0474-45B5-B79F-7416F354543A}', '{8DD21CDE-CB6B-4D5B-B995-D2747AB69ED1}', '{E2968C91-5339-499C-9812-DECCCF58A2CD}', '{646A8184-7F05-453A-A2F1-EB9A77D7F0EE}', '{21891CBB-A1EC-4974-B0B6-F74A4B502BC2}', '{12B7D244-BF4C-401B-A65A-D1621D7DD953}', '{F69B50D8-FB5E-4E8C-AE5C-6DCB4790AFAF}', '{38267404-7625-47AA-B0C8-31BCA5D0435D}', '{62A12AF8-CB6D-4CC5-9DB0-B73A7C24E2AD}', '{0C7A7725-643D-4F7E-A2F0-0C8A36C2A594}', '{1CDFD5FF-8DF5-48CF-A773-F5EEBD56469C}', '{74A33555-8E08-47DF-B3CD-A1B4C7AF2B4E}', '{474487EB-4554-4A30-B351-6EF762D2F2D6}'],
                            2, [2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2])

            elshokr = bishopDes_values[0]
            elshokr1 = bishop_values[0]
            elshokr2 = bishop_values[1]

            tobhyna = bishopDes_values[1]
            tobhyna1 = bishop_values[2]
            tobhyna2 = bishop_values[3]

            embiniot = bishopDes_values[2]
            embiniot1 = bishop_values[4]
            embiniot2 = bishop_values[5]

            hytynyat = bishopDes_values[3]
            hytynyat1 = bishop_values[6]
            hytynyat2 = bishop_values[7]

            byhmot8ar = bishopDes_values[4]-1
            byhmot8ar1 = bishop_values[4]
            byhmot8ar2 = bishop_values[5]

            tomakario = bishopDes_values[5] - 1
            tomakario1 = bishop_values[4]
            tomakario2 = bishop_values[5]

            eya8aby = bishopDes_values[6] - 1
            eya8aby1 = bishop_values[8]
            eya8aby2 = bishop_values[9]

            mrdElkatholikon = bishopDes_values[7]
            mrdElkatholikon1 = bishop_values[10]
            mrdElkatholikon2 = bishop_values[11]

            maro = bishopDes_values[8]
            maro1 = bishop_values[4]
            maro2 = bishop_values[5]

            esbasmos1 = bishopDes_values[9]
            esbasmos11 = bishop_values[12]
            esbasmos12 = bishop_values[13]

            esbasmos2 = bishopDes_values[10]
            esbasmos21 = bishop_values[12]
            esbasmos22 = bishop_values[13]

            elaba2basyly = bishopDes_values[11] - 1
            elaba2basyly1 = bishop_values[0]
            elaba2basyly2 = bishop_values[1]

            elaba28yry8ory = bishopDes_values[12] - 1
            elaba28yry8ory1 = bishop_values[0]
            elaba28yry8ory2 = bishop_values[1]

            if guestBishop < 2:
                elshokr2 = elshokr2-1
                tobhyna2 = tobhyna2-2
                embiniot2 = embiniot2-1
                hytynyat2 = hytynyat2-3
                byhmot8ar2 = byhmot8ar2-1
                tomakario2 = tomakario2-1
                eya8aby2 = eya8aby2-2
                mrdElkatholikon2 = mrdElkatholikon2-1
                maro2 = maro2-1
                esbasmos12 = esbasmos12-2
                esbasmos22 = esbasmos22-2
                elaba2basyly2 = elaba2basyly2-1
                elaba28yry8ory2 = elaba28yry8ory2-1
        
            start_positions = [mazmorELtawzy3, elaba28yry8ory, elaba2basyly, esbasmos2, esbasmos1, elengilElrab33, falyrf3oh, elmazmorEltany2, elengilEltalt3, elengilEltany3, elengilElawl3, elmazmorElawl2, maro, elebrksis3, elkatholikon3, mrdElkatholikon, elbouls3, eya8aby, tomakario, byhmot8ar, hytynyat, embiniot, tobhyna, elshokr]
            start_slides = [mrdMazmorEltawzy3, elaba28yry8ory1, elaba2basyly1, esbasmos21, esbasmos11, elengilElrab3, falyrf3oh1, elmazmorEltany, elengilEltalt, elengilEltany, elengilElawl, elmazmorElawl, maro1, elebrksis1, elkatholikon1, mrdElkatholikon1, elbouls1, eya8aby1, tomakario1, byhmot8ar1, hytynyat1, embiniot1, tobhyna1, elshokr1]
            end_slides = [mrdMazmorEltawzy3, elaba28yry8ory2, elaba2basyly2, esbasmos22, esbasmos12, elengilElrab32, falyrf3oh2, elmazmorEltany, elengilEltalt2, elengilEltany2, elengilElawl2, elmazmorElawl, maro2, elebrksis2, elkatholikon2, mrdElkatholikon2, elbouls2, eya8aby2, tomakario2, byhmot8ar2, hytynyat2, embiniot2, tobhyna2, elshokr2]
        
        else:
            start_positions = [mazmorELtawzy3, elengilElrab33, falyrf3oh, elmazmorEltany2, elengilEltalt3, elengilEltany3, elengilElawl3, elmazmorElawl2, elebrksis3, elkatholikon3, elbouls3]
            start_slides = [mrdMazmorEltawzy3, elengilElrab3, falyrf3oh1, elmazmorEltany, elengilEltalt, elengilEltany, elengilElawl, elmazmorElawl, elebrksis1, elkatholikon1, elbouls1]
            end_slides = [mrdMazmorEltawzy3, elengilElrab32, falyrf3oh2, elmazmorEltany, elengilEltalt2, elengilEltany2, elengilElawl2, elmazmorElawl, elebrksis2, elkatholikon2, elbouls2]

    else:
        start_positions = [mazmorELtawzy3, elengilElrab33, elmazmorEltany2, elengilEltalt3, elengilEltany3, elengilElawl3, elmazmorElawl2, elebrksis3, elkatholikon3, elbouls3]
        start_slides = [mrdMazmorEltawzy3, elengilElrab3, elmazmorEltany, elengilEltalt, elengilEltany, elengilElawl, elmazmorElawl, elebrksis1, elkatholikon1, elbouls1]
        end_slides = [mrdMazmorEltawzy3, elengilElrab32, elmazmorEltany, elengilEltalt2, elengilEltany2, elengilElawl2, elmazmorElawl, elebrksis2, elkatholikon2, elbouls2]

    show_hide_insertImage_replaceText(
        prs1, excel, des_sheet, 
        elsh3anyn_show_full_sections, elsh3anyn_hide_full_sections,
        elsh3anyn_show_full_sections_ranges, elsh3anyn_hide_full_sections_ranges,
        image)
    
    powerpoint = win32com.client.Dispatch("PowerPoint.Application")
    powerpoint.Visible = True  # Open PowerPoint application
    presentation1 = open_presentation_relative_path(prs1)
    presentation2 = open_presentation_relative_path(prs2)

    if guestBishop > 0:
        presentation3 = open_presentation_relative_path(prs3)

    khetamElsha3anyn = find_slide_index_by_title(presentation1, "الشعانين", elkhetam)
    show_slides(presentation1, [[khetamElsha3anyn, khetamElsha3anyn], [allyElqurban, allyElqurban]])

    run_vba_with_slide_id(excel, des_sheet, prs1, presentation1, None, True)

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
                new_slide = presentation1.Slides.Paste(slide_index1).SlideShowTransition.Hidden = False
                slide_index1 -= 1
            current_start_slide += 1

        elif current_position in {elengilElawl3, elengilEltany3, elengilEltalt3, elengilElrab33, 
                                  elmazmorElawl2, elmazmorEltany2, elebrksis3, elkatholikon3, elbouls3}:
            source_slide = presentation2.Slides(current_end_slide)
            source_slide.Copy()
            new_slide = presentation1.Slides.Paste(current_position)
            pic_shape = new_slide.Shapes.AddPicture(image, LinkToFile=False, SaveWithDocument=True, Left=0, Top=144, Width=720, Height=396)
            pic_shape.ZOrder(1)
            new_slide.SlideShowTransition.Hidden = False
            current_end_slide -= 1
            if current_start_slide > current_end_slide:
                current_position += 1

        elif guestBishop>0 and current_position in {elaba28yry8ory, elaba2basyly, byhmot8ar, hytynyat, 
                                                      embiniot, elshokr, tomakario, eya8aby, esbasmos1,
                                                      esbasmos2, maro, mrdElkatholikon, tobhyna}:
            source_slide = presentation3.Slides(current_end_slide)
            source_slide.Copy()
            if current_position in {elaba28yry8ory, elaba2basyly, byhmot8ar, hytynyat, embiniot, elshokr, maro}:
                presentation1.Slides.Paste(current_position).SlideShowTransition.Hidden = False
            else:
                presentation1.Slides.Paste(current_position).SlideShowTransition.Hidden = True
            current_end_slide -= 1
            if(current_start_slide > current_end_slide):
                current_position += 1 

        else:
            source_slide = presentation1.Slides(current_end_slide)
            source_slide.Copy()
            is_hidden = source_slide.SlideShowTransition.Hidden
            new_slide = presentation1.Slides.Paste(current_position)
            new_slide.SlideShowTransition.Hidden = is_hidden
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
            
    move_section_names = [
        "{8F01200A-5058-41B2-8B20-179191B203ED}",
        "{1CDFD5FF-8DF5-48CF-A773-F5EEBD56469C}",
        "{12528D33-08EB-4995-AF1F-61B8AB2F51C2}", 
        "{F0C07D6A-9B65-4AF5-AF75-7BA80C1EEEFC}"
    ]

    target_section_names = [
        "{22F83DFC-792B-4148-8AED-E77703B6E7BB}", 
        "{8F01200A-5058-41B2-8B20-179191B203ED}",
        "{2EB8E5B7-6049-402F-80E4-ED6EDFAB83F4}",
        "{12528D33-08EB-4995-AF1F-61B8AB2F51C2}"
    ]

    # Call the function once for all moves
    move_sections_v2(presentation1, move_section_names, target_section_names)

    #SlideShowTransition
    presentation1.SlideShowSettings.Run()
    
    if guestBishop > 0:
        presentation3.Close()
    presentation2.Close()

def odasEl2yama(copticdate, Bishop=False, guestBishop=0):
    prs1 = relative_path(r"قداس.pptx")  # Using the relative path
    excel = relative_path(r"بيانات القداسات.xlsx")
    excel2 = relative_path(r"Tables.xlsx")
    des_sheet ="القداس"
    image = relative_path(r"Data\Designs\القيامة.png")
    el2yamaText =  ["لأنك قمت","aktwnk", "آك طونك"]
    replacefile(prs1, relative_path(r"Data\CopyData\قداس.pptx"))

    prs2 = relative_path(r"Data\القطمارس\قطمارس الخماسين (القداس).pptx")
    katamars_sheet = "قطمارس الخماسين"
    km = copticdate[1]
    kd = copticdate[2]
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

    # el2yama_show_full_sections = ["ابؤرو", "الليلويا فاي بيبي", "ني سافيف تيرو",
    #                               "طاي شوري", "هيتينية الملاك ميخائيل - للقيامة", "هيتينية القيامة", 
    #                               "مرد ابركسيس القيامة", "اجيوس القيامة", "المزمور السنجاري", "مزمور القيامة قبطي",
    #                               "مرد المزمور القيامة", "مرد انجيل القيامة", "فاي اريه بي اوو", 
    #                               "اوشية اهوية السماء", "اوشية اهوية السماء غ",
    #                               "قسمة للآب في عيد القيامة والخمسين (أيها السيد الرب الإله ضابط الكل)",
    #                               "كاطا ني خوروس", "مدائح القيامة"]
    
    # el2yama_hide_full_sections = ["سوتيس امين", "السنكسار", "اكسيا", "اجيوس الميلاد", "اجيوس الصعود",
    #                               "اجيوس الصلب", "مرد المزمور", "مرد الانجيل", "ربع للعذراء", 
    #                               "قسمة القداس الباسيلي (أيها السيد الرب إلهنا)", "اك اسماروؤت", "بي اويك"]

    # el2yama_show_full_sections_ranges = [["الهيتنيات المجمعة - الشهداء", "هيتينية حبيب جرجس"],
    #                                      ["ياكل الصفوف", "محير القيامة"]]

    # el2yama_hide_full_sections_ranges = [["هيتينية مارجرجس", "هيتينية مارمينا"],
    #                                      ["مقدمة قانون الإيمان", "تقدمة الحمل"]]


    el2yama_show_full_sections = ['{DD757736-F2EB-40FA-9016-1E28087A0BE5}', '{072F3D96-A6C8-405F-9A23-7CCA1B2F13FF}', '{670DAA94-A6C9-4CCD-B4E2-958C71CD3E44}', '{20F525FD-C708-4DDD-8E40-FE502EFEBDDE}', '{51C5460B-5CB6-46E7-A00C-E417B4008CBC}', '{E0F28C9F-5BF8-4AFC-806F-716597B95865}', '{17FBFAA9-1B87-491D-89CE-BA67D27ADCC2}', '{92AE0B66-4188-4137-A35D-D2187E4529A2}', '{973DBBAC-E645-4981-B3F5-5DB1413508D0}', '{E216042A-B60C-490D-9991-3A3DC488ACCF}', '{7C1AB872-5559-4CAE-A399-FED4376DA488}', '{B5D1AE9E-D436-4277-A6E3-20A62039B3F7}', '{B7D98377-B994-4654-B49C-DE10E0DDE4F1}', '{BC7E3DCD-6AA8-44CC-B8AF-BC3E2BC71B5A}', '{A20DA654-32F7-4B4C-96CB-C76232EB96E8}', '{3E1960C6-4E54-4563-B3A6-E8AAB1680607}', '{8E020F97-3459-40AC-828C-13346E2C46B1}', '{1D781881-AD2E-41FE-97D4-458A86F892F1}']
    el2yama_hide_full_sections = ['{4D2B15D5-C978-467C-9D6C-726FE25128B8}', '{1DA5C6AA-2FE5-461B-9E2C-40113CFC7804}', '{A8A52E1F-44DD-45E1-A737-4E13E15D5F1F}', '{76175201-2586-46AD-8E9F-C69E29FE2620}', '{CD4B95FF-0E0E-42D3-8DC4-224C3DD732F7}', '{F13A48F2-238D-4617-B84E-9B0A694D9A18}', '{06D592C8-05BF-4B7C-86F7-FDAB3FAB5FB1}', '{ECE652ED-1345-4C6D-B92D-5996CFA27AEE}', '{681FF6A7-4230-4171-8F41-83FD64E8C960}', '{507EFD97-98F8-4376-848B-20D72E16D2C1}', '{B9A30F5E-0C89-471B-A99A-23DBE7F58504}', '{31685B5B-48C4-437E-858C-CF8D225C0C26}']
    el2yama_show_full_sections_ranges = [['{C2140BA5-149F-49BF-A4F6-F0205CF4B99C}', '{457CE0B0-6F83-4DF4-8166-692625329991}'], ['{D5A75588-692D-473A-9F8E-A1598C697F76}', '{02F409A1-6F4B-4195-8303-693208D52612}']]
    el2yama_hide_full_sections_ranges = [['{735C1803-8AE4-44C8-A4EA-C7B3C1493312}', '{D05DF2AC-49C5-413A-A8AC-2D17C1DDFF97}'], ['{409B3D0A-B40A-4475-811D-72C5125134AB}', '{FB902AC1-5581-4552-B503-116755E9D9A8}']]

    # el2yama_values = ["تكملة على حسب المناسبة", "مزمور التوزيع", "مزمور التوزيع", "مرد توزيع القيامة",
    #                   "الانجيل", "المزمور", "الابركسيس", "الكاثوليكون", "البولس عربي", "اللي القربان",
    #                   "كاطا ني خوروس الحجاب", "اجيوس القيامة", "اجيوس القيامة", 
    #                   "قسمة للآب في عيد القيامة والخمسين (أيها السيد الرب الإله ضابط الكل)"]

    el2yama_values = find_slide_nums_arrays_v2(excel, des_sheet, 
                    ['{A18EDC94-F257-4FAC-99C7-0A8EA70F0FAF}', '{C29E5A83-A98B-4077-8194-99A6D803EF53}', '{C29E5A83-A98B-4077-8194-99A6D803EF53}', '{AB0FAC0A-1DAB-4ABD-BAA1-7557AA4860AA}', '{C7D4A109-F792-4661-BAD0-075FD1A1909F}', '{B74DBB8C-2B2D-46E4-9508-DA46008D19A4}', '{E234C6C7-3837-4CE4-A541-CDC9627AAAC2}', '{6D4B3F52-63BF-435F-BF0C-C9D41120C2A3}', '{D88055F5-EAA0-4C8E-8249-C364A572BF7B}', '{BBBAC16F-044D-4F33-8068-620F498B59CD}', '{20683024-614D-47D3-B2A6-1B01B420894A}', '{92AE0B66-4188-4137-A35D-D2187E4529A2}', '{92AE0B66-4188-4137-A35D-D2187E4529A2}', '{3E1960C6-4E54-4563-B3A6-E8AAB1680607}'],
                    2, [1, 1, 2, 1, 2, 2, 2, 2, 2, 2, 1, 1, 2, 2])

    #الختام
    elkhetam = el2yama_values[0]
    #التوزيع
    mazmorELtawzy3 = el2yama_values[1] + 1
    mazmorELtawzy32 = el2yama_values[2] - 1
    mrdMazmorEltawzy3 = el2yama_values[3]
    #القرائات
    elengil3 = el2yama_values[4]
    elmazmor3 = el2yama_values[5]
    elebrksis3 = el2yama_values[6]
    elkatholikon3 = el2yama_values[7]
    elbouls3 = el2yama_values[8]

    #اللي القربان و كاطا الحجاب
    allyelorban = el2yama_values[9]
    kataEl7egab = el2yama_values[10]

    #اجيوس القيامة
    agiosEl2yama = el2yama_values[11]    
    agiosEl2yama2 = el2yama_values[12]

    #القسمة
    el2esma = el2yama_values[13]

    if Bishop == True:
        prs3 = relative_path(r"Data\حضور الأسقف.pptx")
        sheet = "في حضور الأسقف"

        bishop_show_values = ['{F76B0D75-0474-45B5-B79F-7416F354543A}', '{E2968C91-5339-499C-9812-DECCCF58A2CD}',
                              '{62A12AF8-CB6D-4CC5-9DB0-B73A7C24E2AD}', '{B74DBB8C-2B2D-46E4-9508-DA46008D19A4}',
                              '{A9183893-7B7E-459F-8547-F7A8F7D2D521}']
        el2yama_show_full_sections.extend(bishop_show_values)

        if guestBishop > 0:
            # bishop_values = ["صلاة الشكر", "صلاة الشكر", "طوبه هينا الكبيرة", "طوبه هينا الكبيرة", 
            #                  "نيم بينيوت", "نيم بينيوت", "الهيتنيات", "الهيتنيات",
            #                  "اي اغابي", "اي اغابي", "مرد الكاثوليكون", "مرد الكاثوليكون",
            #                  "الاسبسمس", "الاسبسمس"]

            # bishopDes_values = ["تكملة في حضور الاسقف", "طوبه هينا الكبيرة",
            #                     "امبين يوت اتطايوت", "تكملة الهيتنيات", "بى اهموت غار الصغيرة",
            #                     "تو ماكريو", "اي اغابي", "ابيت جيك ايفول", "مارو اتشاسف",
            #                     "الاسبسمس الادام السنوي+الختام", "ختام الاسبسمس الادام",
            #                     "اوشية الاباء (ب)", "اوشية الاباء غ"]
            
            bishop_values = find_slide_nums_arrays_v2(excel, sheet, 
                            ['{6851F163-CBEF-4014-A853-CE100557BA6A}', '{6851F163-CBEF-4014-A853-CE100557BA6A}', '{B084BC40-61E1-4477-98DA-15CFB06AEE91}', '{B084BC40-61E1-4477-98DA-15CFB06AEE91}', '{97203297-EECB-4D41-B2E3-AD9A4863847E}', '{97203297-EECB-4D41-B2E3-AD9A4863847E}', '{7C84083F-E6D3-4669-9130-AC7E8D935A98}', '{7C84083F-E6D3-4669-9130-AC7E8D935A98}', '{D0234C99-69FE-407A-82A9-D7A676919E93}', '{D0234C99-69FE-407A-82A9-D7A676919E93}', '{0B519645-F935-43AE-A9CE-6E2FC03833BB}', '{0B519645-F935-43AE-A9CE-6E2FC03833BB}', '{D1378DB5-29D1-4800-9D96-10F2535EEB57}', '{D1378DB5-29D1-4800-9D96-10F2535EEB57}'], 
                            2, [1, 2, 1, 2, 1, 2, 1, 2, 1, 2, 1, 2, 1, 2])

            bishopDes_values = find_slide_nums_arrays_v2(excel, des_sheet, 
                                ['{F76B0D75-0474-45B5-B79F-7416F354543A}', '{8DD21CDE-CB6B-4D5B-B995-D2747AB69ED1}', '{E2968C91-5339-499C-9812-DECCCF58A2CD}', '{646A8184-7F05-453A-A2F1-EB9A77D7F0EE}', '{21891CBB-A1EC-4974-B0B6-F74A4B502BC2}', '{12B7D244-BF4C-401B-A65A-D1621D7DD953}', '{F69B50D8-FB5E-4E8C-AE5C-6DCB4790AFAF}', '{38267404-7625-47AA-B0C8-31BCA5D0435D}', '{62A12AF8-CB6D-4CC5-9DB0-B73A7C24E2AD}', '{0C7A7725-643D-4F7E-A2F0-0C8A36C2A594}', '{1CDFD5FF-8DF5-48CF-A773-F5EEBD56469C}', '{74A33555-8E08-47DF-B3CD-A1B4C7AF2B4E}', '{474487EB-4554-4A30-B351-6EF762D2F2D6}'],                  
                                2, [2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2])

            elshokr = bishopDes_values[0]
            elshokr1 = bishop_values[0]
            elshokr2 = bishop_values[1]

            tobhyna = bishopDes_values[1]
            tobhyna1 = bishop_values[2]
            tobhyna2 = bishop_values[3]

            embiniot = bishopDes_values[2]
            embiniot1 = bishop_values[4]
            embiniot2 = bishop_values[5]

            hytynyat = bishopDes_values[3]
            hytynyat1 = bishop_values[6]
            hytynyat2 = bishop_values[7]

            byhmot8ar = bishopDes_values[4]-1
            byhmot8ar1 = bishop_values[4]
            byhmot8ar2 = bishop_values[5]

            tomakario = bishopDes_values[5] - 1
            tomakario1 = bishop_values[4]
            tomakario2 = bishop_values[5]

            eya8aby = bishopDes_values[6] - 1
            eya8aby1 = bishop_values[8]
            eya8aby2 = bishop_values[9]

            mrdElkatholikon = bishopDes_values[7]
            mrdElkatholikon1 = bishop_values[10]
            mrdElkatholikon2 = bishop_values[11]

            maro = bishopDes_values[8]
            maro1 = bishop_values[4]
            maro2 = bishop_values[5]

            esbasmos1 = bishopDes_values[9]
            esbasmos11 = bishop_values[12]
            esbasmos12 = bishop_values[13]

            esbasmos2 = bishopDes_values[10]
            esbasmos21 = bishop_values[12]
            esbasmos22 = bishop_values[13]

            elaba2basyly = bishopDes_values[11] - 1
            elaba2basyly1 = bishop_values[0]
            elaba2basyly2 = bishop_values[1]

            elaba28yry8ory = bishopDes_values[12] - 1
            elaba28yry8ory1 = bishop_values[0]
            elaba28yry8ory2 = bishop_values[1]

            if guestBishop < 2:
                elshokr2 = elshokr2-1
                tobhyna2 = tobhyna2-2
                embiniot2 = embiniot2-1
                hytynyat2 = hytynyat2-3
                byhmot8ar2 = byhmot8ar2-1
                tomakario2 = tomakario2-1
                eya8aby2 = eya8aby2-2
                mrdElkatholikon2 = mrdElkatholikon2-1
                maro2 = maro2-1
                esbasmos12 = esbasmos12-2
                esbasmos22 = esbasmos22-2
                elaba2basyly2 = elaba2basyly2-1
                elaba28yry8ory2 = elaba28yry8ory2-1
        
            start_positions = [mazmorELtawzy3, elaba28yry8ory, elaba2basyly, esbasmos2, esbasmos1, elengil3, elmazmor3, maro, elebrksis3, elkatholikon3, mrdElkatholikon, elbouls3, eya8aby, tomakario, byhmot8ar, hytynyat, embiniot, tobhyna, elshokr]
            start_slides = [mrdMazmorEltawzy3, elaba28yry8ory1, elaba2basyly1, esbasmos21, esbasmos11, elengil, elmzmor, maro1, elebrksis1, elkatholikon1, mrdElkatholikon1, elbouls1, eya8aby1, tomakario1, byhmot8ar1, hytynyat1, embiniot1, tobhyna1, elshokr1]
            end_slides = [mrdMazmorEltawzy3, elaba28yry8ory2, elaba2basyly2, esbasmos22, esbasmos12, elengil2, elmzmor, maro2, elebrksis2, elkatholikon2, mrdElkatholikon2, elbouls2, eya8aby2, tomakario2, byhmot8ar2, hytynyat2, embiniot2, tobhyna2, elshokr2]

    if guestBishop == 0:
        start_positions = [mazmorELtawzy3, elengil3, elmazmor3, elebrksis3, elkatholikon3, elbouls3]
        start_slides = [mrdMazmorEltawzy3, elengil, elmzmor, elebrksis1, elkatholikon1, elbouls1]
        end_slides = [mrdMazmorEltawzy3, elengil2, elmzmor, elebrksis2, elkatholikon2, elbouls2]

    show_hide_insertImage_replaceText(prs1, excel, des_sheet, el2yama_show_full_sections, el2yama_hide_full_sections, el2yama_show_full_sections_ranges, el2yama_hide_full_sections_ranges, image, el2yamaText)

    powerpoint = win32com.client.Dispatch("PowerPoint.Application")
    powerpoint.Visible = True  # Open PowerPoint application
    presentation1 = open_presentation_relative_path(prs1)
    presentation2 = open_presentation_relative_path(prs2)
    
    if guestBishop > 0:
        insert_image_to_slides_same_file(prs3, image)
        presentation3 = open_presentation_relative_path(prs3)

    khetamEl2yama = find_slide_index_by_title(presentation1, "القيامة", elkhetam)
    NoSo3odEl2esma = find_slide_index_by_title(presentation1, "عيد الصعود", el2esma, "up")
    show_slides(presentation1, [[khetamEl2yama, khetamEl2yama], [allyelorban, allyelorban], [kataEl7egab, kataEl7egab]])
    hide_slides(presentation1, [[NoSo3odEl2esma, NoSo3odEl2esma]])
    run_vba_with_slide_id(excel, des_sheet, prs1, presentation1)
    
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
                new_slide = presentation1.Slides.Paste(slide_index1).SlideShowTransition.Hidden = False
                slide_index1 -= 1
            current_start_slide += 1

        elif (current_position == elengil3 or current_position == elmazmor3 or current_position == elebrksis3 
            or current_position == elkatholikon3 or current_position == elbouls3):
            source_slide = presentation2.Slides(current_end_slide)
            source_slide.Copy()
            new_slide = presentation1.Slides.Paste(current_position)
            pic_shape = new_slide.Shapes.AddPicture(image, LinkToFile=False, SaveWithDocument=True, Left=0, Top=144, Width=720, Height=396)
            pic_shape.ZOrder(1)
            new_slide.SlideShowTransition.Hidden = False
            current_end_slide -= 1
            # Add slide duplication after elmzmor1 is processed
            if guestBishop == 0 and current_position == elmazmor3:
                # Perform the duplication twice
                for _ in range(2):  # Loop to duplicate the section twice
                    # Copy the first slide (agiosElsalyb) and paste it after the second slide (agiosElsalyb2)
                    agiosElmilad_slide = presentation1.Slides(agiosEl2yama)
                    agiosElmilad_slide.Copy()
                    presentation1.Slides.Paste(agiosEl2yama2 + 1)

                    # Copy the second slide (agiosElsalyb2) and paste it after the newly copied first slide
                    agiosElmilad2_slide = presentation1.Slides(agiosEl2yama2)
                    agiosElmilad2_slide.Copy()
                    presentation1.Slides.Paste(agiosEl2yama2 + 2)  # Adjust to account for new first slide
            if current_start_slide > current_end_slide:
                current_position += 1

        elif Bishop and current_position in {elaba28yry8ory, elaba2basyly, byhmot8ar, hytynyat, embiniot, 
                                             elshokr, tomakario, eya8aby, esbasmos1, esbasmos2, maro, 
                                             mrdElkatholikon, tobhyna}:
            source_slide = presentation3.Slides(current_end_slide)
            source_slide.Copy()
            if current_position in {elaba28yry8ory, elaba2basyly, byhmot8ar, hytynyat, embiniot, elshokr, maro}:
                presentation1.Slides.Paste(current_position).SlideShowTransition.Hidden = False
            else:
                presentation1.Slides.Paste(current_position).SlideShowTransition.Hidden = True
            current_end_slide -= 1
            if (current_position == maro) and (current_start_slide > current_end_slide):
                    # Perform the duplication twice
                    for _ in range(2):  # Loop to duplicate the section twice
                        # Copy the first slide (agiosElmilad) and paste it after the second slide (agiosElmilad2)
                        agiosElmilad_slide = presentation1.Slides(agiosEl2yama)
                        agiosElmilad_slide.Copy()
                        presentation1.Slides.Paste(agiosEl2yama2 + 1)

                        # Copy the second slide (agiosElmilad2) and paste it after the newly copied first slide
                        agiosElmilad2_slide = presentation1.Slides(agiosEl2yama2)
                        agiosElmilad2_slide.Copy()
                        presentation1.Slides.Paste(agiosEl2yama2 + 2)  # Adjust to account for new first slide
            
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

    move_section_names = [
        "{2F9B712F-EC2B-474A-B445-B4A46D96A229}",
        "{1CDFD5FF-8DF5-48CF-A773-F5EEBD56469C}",
        "{9888DFF9-9A38-4DB9-80E8-E000DA07D088}", 
        "{F0C07D6A-9B65-4AF5-AF75-7BA80C1EEEFC}"
    ]

    target_section_names = [
        "{22F83DFC-792B-4148-8AED-E77703B6E7BB}", 
        "{2F9B712F-EC2B-474A-B445-B4A46D96A229}",
        "{2EB8E5B7-6049-402F-80E4-ED6EDFAB83F4}",
        "{9888DFF9-9A38-4DB9-80E8-E000DA07D088}"
    ]

    # Call the function once for all moves
    move_sections_v2(presentation1, move_section_names, target_section_names)
    
    presentation2.Close()
    if guestBishop>0:
        presentation3.Close()

    presentation1.SlideShowSettings.Run()

def odasEl5amasyn_2_39(copticdate, Bishop=False, guestBishop=0):
    prs1 = relative_path(r"قداس.pptx")  # Using the relative path
    excel = relative_path(r"بيانات القداسات.xlsx")
    excel2 = relative_path(r"Tables.xlsx")
    des_sheet ="القداس"
    image = relative_path(r"Data\Designs\القيامة.png")
    el2yamaText =  ["لأنك قمت","aktwnk", "آك طونك"]
    replacefile(prs1, relative_path(r"Data\CopyData\قداس.pptx"))

    prs2 = relative_path(r"Data\القطمارس\قطمارس الخماسين (القداس).pptx")
    katamars_sheet = "قطمارس الخماسين"
    km = copticdate[1]
    kd = copticdate[2]
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

    # el2yama_show_full_sections = ["الليلويا فاي بيبي", "طاي شوري", "هيتينية الملاك ميخائيل - للقيامة", "هيتينية القيامة", 
    #                               "مرد ابركسيس القيامة", "ياكل الصفوف", "اجيوس القيامة", "مرد المزمور القيامة",
    #                               "فاي اريه بي اوو", "اوشية اهوية السماء", "اوشية اهوية السماء غ",
    #                               "قسمة للآب في عيد القيامة والخمسين (أيها السيد الرب الإله ضابط الكل)",
    #                               "مدائح الخماسين"]
    
    # el2yama_hide_full_sections = ["السنكسار", "اكسيا", "اجيوس الميلاد", "اجيوس الصعود",
    #                               "اجيوس الصلب", "مرد المزمور", "مرد الانجيل", "ربع للعذراء", 
    #                               "قسمة القداس الباسيلي (أيها السيد الرب إلهنا)", "اك اسماروؤت", "بي اويك"]

    # el2yama_show_full_sections_ranges = [["اخرستوس انيستى", "محير القيامة"]]

    el2yama_show_full_sections = ['{072F3D96-A6C8-405F-9A23-7CCA1B2F13FF}', '{20F525FD-C708-4DDD-8E40-FE502EFEBDDE}', '{51C5460B-5CB6-46E7-A00C-E417B4008CBC}', '{E0F28C9F-5BF8-4AFC-806F-716597B95865}', '{17FBFAA9-1B87-491D-89CE-BA67D27ADCC2}', '{D5A75588-692D-473A-9F8E-A1598C697F76}', '{92AE0B66-4188-4137-A35D-D2187E4529A2}', '{7C1AB872-5559-4CAE-A399-FED4376DA488}', '{B7D98377-B994-4654-B49C-DE10E0DDE4F1}', '{BC7E3DCD-6AA8-44CC-B8AF-BC3E2BC71B5A}', '{A20DA654-32F7-4B4C-96CB-C76232EB96E8}', '{3E1960C6-4E54-4563-B3A6-E8AAB1680607}', '{8E020F97-3459-40AC-828C-13346E2C46B1}', '{F74608AC-8E8B-460A-97E7-1E6B86D50B5B}']
    el2yama_hide_full_sections = ['{1DA5C6AA-2FE5-461B-9E2C-40113CFC7804}', '{A8A52E1F-44DD-45E1-A737-4E13E15D5F1F}', '{76175201-2586-46AD-8E9F-C69E29FE2620}', '{CD4B95FF-0E0E-42D3-8DC4-224C3DD732F7}', '{31685B5B-48C4-437E-858C-CF8D225C0C26}', '{F13A48F2-238D-4617-B84E-9B0A694D9A18}', '{06D592C8-05BF-4B7C-86F7-FDAB3FAB5FB1}', '{ECE652ED-1345-4C6D-B92D-5996CFA27AEE}', '{681FF6A7-4230-4171-8F41-83FD64E8C960}', '{507EFD97-98F8-4376-848B-20D72E16D2C1}', '{B9A30F5E-0C89-471B-A99A-23DBE7F58504}']
    el2yama_show_full_sections_ranges = [['{825A877D-8247-4AF9-9C5E-F8C914939068}', '{02F409A1-6F4B-4195-8303-693208D52612}']]

    # el2yama_values = ["تكملة على حسب المناسبة", "مزمور التوزيع", "مزمور التوزيع", "مرد توزيع القيامة",
    #                   "الانجيل", "المزمور", "الابركسيس", "الكاثوليكون", "البولس عربي",
    #                   "اجيوس القيامة", "اجيوس القيامة", "مرد انجيل القيامة",
    #                   "قسمة للآب في عيد القيامة والخمسين (أيها السيد الرب الإله ضابط الكل)", "بدء قداس الكلمة"]

    el2yama_values = find_slide_nums_arrays_v2(excel, des_sheet, 
                    ['{A18EDC94-F257-4FAC-99C7-0A8EA70F0FAF}', '{C29E5A83-A98B-4077-8194-99A6D803EF53}', '{C29E5A83-A98B-4077-8194-99A6D803EF53}', '{AB0FAC0A-1DAB-4ABD-BAA1-7557AA4860AA}', '{C7D4A109-F792-4661-BAD0-075FD1A1909F}', '{B74DBB8C-2B2D-46E4-9508-DA46008D19A4}', '{E234C6C7-3837-4CE4-A541-CDC9627AAAC2}', '{6D4B3F52-63BF-435F-BF0C-C9D41120C2A3}', '{D88055F5-EAA0-4C8E-8249-C364A572BF7B}', '{92AE0B66-4188-4137-A35D-D2187E4529A2}', '{92AE0B66-4188-4137-A35D-D2187E4529A2}', '{B5D1AE9E-D436-4277-A6E3-20A62039B3F7}', '{3E1960C6-4E54-4563-B3A6-E8AAB1680607}', '{C08D8D44-E49E-47CE-8027-C8AE26B1AA9A}'],
                    2, [1, 1, 2, 1, 2, 2, 2, 2, 2, 1, 2, 2, 2, 2])

    #الختام
    elkhetam = el2yama_values[0]
    #التوزيع
    mazmorELtawzy3 = el2yama_values[1] + 1
    mazmorELtawzy32 = el2yama_values[2] - 1
    mrdMazmorEltawzy3 = el2yama_values[3]
    
    #القرائات
    elengil3 = el2yama_values[4]
    elmazmor3 = el2yama_values[5]
    elebrksis3 = el2yama_values[6]
    elkatholikon3 = el2yama_values[7]
    elbouls3 = el2yama_values[8]

    #اجيوس القيامة
    agiosEl2yama = el2yama_values[9]    
    agiosEl2yama2 = el2yama_values[10]

    #مرد انجيل القيامة
    mrdEl2yama = el2yama_values[11]

    #القسمة
    el2esma = el2yama_values[12]

    if Bishop == True:
        prs3 = relative_path(r"Data\حضور الأسقف.pptx")
        sheet = "في حضور الأسقف"

        bishop_show_values = ['{F76B0D75-0474-45B5-B79F-7416F354543A}', '{E2968C91-5339-499C-9812-DECCCF58A2CD}',
                              '{62A12AF8-CB6D-4CC5-9DB0-B73A7C24E2AD}', '{B74DBB8C-2B2D-46E4-9508-DA46008D19A4}',
                              '{A9183893-7B7E-459F-8547-F7A8F7D2D521}']
        el2yama_show_full_sections.extend(bishop_show_values)

        if guestBishop > 0:
            # bishop_values = ["صلاة الشكر", "صلاة الشكر", "طوبه هينا الكبيرة", "طوبه هينا الكبيرة", 
            #                  "نيم بينيوت", "نيم بينيوت", "الهيتنيات", "الهيتنيات",
            #                  "اي اغابي", "اي اغابي", "مرد الكاثوليكون", "مرد الكاثوليكون",
            #                  "الاسبسمس", "الاسبسمس"]

            # bishopDes_values = ["تكملة في حضور الاسقف", "طوبه هينا الكبيرة",
            #                     "امبين يوت اتطايوت", "تكملة الهيتنيات", "بى اهموت غار الصغيرة",
            #                     "تو ماكريو", "اي اغابي", "ابيت جيك ايفول", "مارو اتشاسف",
            #                     "الاسبسمس الادام السنوي+الختام", "ختام الاسبسمس الادام",
            #                     "اوشية الاباء (ب)", "اوشية الاباء غ"]
            
            bishop_values = find_slide_nums_arrays_v2(excel, sheet, 
                            ['{6851F163-CBEF-4014-A853-CE100557BA6A}', '{6851F163-CBEF-4014-A853-CE100557BA6A}', '{B084BC40-61E1-4477-98DA-15CFB06AEE91}', '{B084BC40-61E1-4477-98DA-15CFB06AEE91}', '{97203297-EECB-4D41-B2E3-AD9A4863847E}', '{97203297-EECB-4D41-B2E3-AD9A4863847E}', '{7C84083F-E6D3-4669-9130-AC7E8D935A98}', '{7C84083F-E6D3-4669-9130-AC7E8D935A98}', '{D0234C99-69FE-407A-82A9-D7A676919E93}', '{D0234C99-69FE-407A-82A9-D7A676919E93}', '{0B519645-F935-43AE-A9CE-6E2FC03833BB}', '{0B519645-F935-43AE-A9CE-6E2FC03833BB}', '{D1378DB5-29D1-4800-9D96-10F2535EEB57}', '{D1378DB5-29D1-4800-9D96-10F2535EEB57}'], 
                            2, [1, 2, 1, 2, 1, 2, 1, 2, 1, 2, 1, 2, 1, 2])

            bishopDes_values = find_slide_nums_arrays_v2(excel, des_sheet, 
                                ['{F76B0D75-0474-45B5-B79F-7416F354543A}', '{8DD21CDE-CB6B-4D5B-B995-D2747AB69ED1}', '{E2968C91-5339-499C-9812-DECCCF58A2CD}', '{646A8184-7F05-453A-A2F1-EB9A77D7F0EE}', '{21891CBB-A1EC-4974-B0B6-F74A4B502BC2}', '{12B7D244-BF4C-401B-A65A-D1621D7DD953}', '{F69B50D8-FB5E-4E8C-AE5C-6DCB4790AFAF}', '{38267404-7625-47AA-B0C8-31BCA5D0435D}', '{62A12AF8-CB6D-4CC5-9DB0-B73A7C24E2AD}', '{0C7A7725-643D-4F7E-A2F0-0C8A36C2A594}', '{1CDFD5FF-8DF5-48CF-A773-F5EEBD56469C}', '{74A33555-8E08-47DF-B3CD-A1B4C7AF2B4E}', '{474487EB-4554-4A30-B351-6EF762D2F2D6}'],                  
                                2, [2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2])

            elshokr = bishopDes_values[0]
            elshokr1 = bishop_values[0]
            elshokr2 = bishop_values[1]

            tobhyna = bishopDes_values[1]
            tobhyna1 = bishop_values[2]
            tobhyna2 = bishop_values[3]

            embiniot = bishopDes_values[2]
            embiniot1 = bishop_values[4]
            embiniot2 = bishop_values[5]

            hytynyat = bishopDes_values[3]
            hytynyat1 = bishop_values[6]
            hytynyat2 = bishop_values[7]

            byhmot8ar = bishopDes_values[4]-1
            byhmot8ar1 = bishop_values[4]
            byhmot8ar2 = bishop_values[5]

            tomakario = bishopDes_values[5] - 1
            tomakario1 = bishop_values[4]
            tomakario2 = bishop_values[5]

            eya8aby = bishopDes_values[6] - 1
            eya8aby1 = bishop_values[8]
            eya8aby2 = bishop_values[9]

            mrdElkatholikon = bishopDes_values[7]
            mrdElkatholikon1 = bishop_values[10]
            mrdElkatholikon2 = bishop_values[11]

            maro = bishopDes_values[8]
            maro1 = bishop_values[4]
            maro2 = bishop_values[5]

            esbasmos1 = bishopDes_values[9]
            esbasmos11 = bishop_values[12]
            esbasmos12 = bishop_values[13]

            esbasmos2 = bishopDes_values[10]
            esbasmos21 = bishop_values[12]
            esbasmos22 = bishop_values[13]

            elaba2basyly = bishopDes_values[11] - 1
            elaba2basyly1 = bishop_values[0]
            elaba2basyly2 = bishop_values[1]

            elaba28yry8ory = bishopDes_values[12] - 1
            elaba28yry8ory1 = bishop_values[0]
            elaba28yry8ory2 = bishop_values[1]

            if guestBishop < 2:
                elshokr2 = elshokr2-1
                tobhyna2 = tobhyna2-2
                embiniot2 = embiniot2-1
                hytynyat2 = hytynyat2-3
                byhmot8ar2 = byhmot8ar2-1
                tomakario2 = tomakario2-1
                eya8aby2 = eya8aby2-2
                mrdElkatholikon2 = mrdElkatholikon2-1
                maro2 = maro2-1
                esbasmos12 = esbasmos12-2
                esbasmos22 = esbasmos22-2
                elaba2basyly2 = elaba2basyly2-1
                elaba28yry8ory2 = elaba28yry8ory2-1
        
            start_positions = [mazmorELtawzy3, elaba28yry8ory, elaba2basyly, esbasmos2, esbasmos1, elengil3, elmazmor3, maro, elebrksis3, elkatholikon3, mrdElkatholikon, elbouls3, eya8aby, tomakario, byhmot8ar, hytynyat, embiniot, tobhyna, elshokr]
            start_slides = [mrdMazmorEltawzy3, elaba28yry8ory1, elaba2basyly1, esbasmos21, esbasmos11, elengil, elmzmor, maro1, elebrksis1, elkatholikon1, mrdElkatholikon1, elbouls1, eya8aby1, tomakario1, byhmot8ar1, hytynyat1, embiniot1, tobhyna1, elshokr1]
            end_slides = [mrdMazmorEltawzy3, elaba28yry8ory2, elaba2basyly2, esbasmos22, esbasmos12, elengil2, elmzmor, maro2, elebrksis2, elkatholikon2, mrdElkatholikon2, elbouls2, eya8aby2, tomakario2, byhmot8ar2, hytynyat2, embiniot2, tobhyna2, elshokr2]

    if guestBishop == 0:
        start_positions = [mazmorELtawzy3, elengil3, elmazmor3, elebrksis3, elkatholikon3, elbouls3]
        start_slides = [mrdMazmorEltawzy3, elengil, elmzmor, elebrksis1, elkatholikon1, elbouls1]
        end_slides = [mrdMazmorEltawzy3, elengil2, elmzmor, elebrksis2, elkatholikon2, elbouls2]

    show_hide_insertImage_replaceText(prs1, excel, des_sheet, el2yama_show_full_sections, el2yama_hide_full_sections, el2yama_show_full_sections_ranges, None, image, el2yamaText)

    powerpoint = win32com.client.Dispatch("PowerPoint.Application")
    powerpoint.Visible = True  # Open PowerPoint application
    presentation1 = open_presentation_relative_path(prs1)
    presentation2 = open_presentation_relative_path(prs2)
    
    if guestBishop > 0:
        insert_image_to_slides_same_file(prs3, image)
        presentation3 = open_presentation_relative_path(prs3)

    khetamEl2yama = find_slide_index_by_title(presentation1, "القيامة", elkhetam)
    NoSo3odEl2esma = find_slide_index_by_title(presentation1, "عيد الصعود", el2esma, "up")
    show_slides(presentation1, [[khetamEl2yama, khetamEl2yama], [mrdEl2yama, mrdEl2yama]])
    hide_slides(presentation1, [[NoSo3odEl2esma, NoSo3odEl2esma]])
    run_vba_with_slide_id(excel, des_sheet, prs1, presentation1)
    agbya(presentation1, el2yama_values[13], 1)

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
                new_slide = presentation1.Slides.Paste(slide_index1).SlideShowTransition.Hidden = False
                slide_index1 -= 1
            current_start_slide += 1

        elif (current_position == elengil3 or current_position == elmazmor3 or current_position == elebrksis3 
            or current_position == elkatholikon3 or current_position == elbouls3):
            source_slide = presentation2.Slides(current_end_slide)
            source_slide.Copy()
            new_slide = presentation1.Slides.Paste(current_position)
            pic_shape = new_slide.Shapes.AddPicture(image, LinkToFile=False, SaveWithDocument=True, Left=0, Top=144, Width=720, Height=396)
            pic_shape.ZOrder(1)
            new_slide.SlideShowTransition.Hidden = False
            current_end_slide -= 1
            # Add slide duplication after elmzmor1 is processed
            if guestBishop == 0 and current_position == elmazmor3:
                # Perform the duplication twice
                for _ in range(2):  # Loop to duplicate the section twice
                    # Copy the first slide (agiosElsalyb) and paste it after the second slide (agiosElsalyb2)
                    agiosElmilad_slide = presentation1.Slides(agiosEl2yama)
                    agiosElmilad_slide.Copy()
                    presentation1.Slides.Paste(agiosEl2yama2 + 1)

                    # Copy the second slide (agiosElsalyb2) and paste it after the newly copied first slide
                    agiosElmilad2_slide = presentation1.Slides(agiosEl2yama2)
                    agiosElmilad2_slide.Copy()
                    presentation1.Slides.Paste(agiosEl2yama2 + 2)  # Adjust to account for new first slide
            if current_start_slide > current_end_slide:
                current_position += 1

        elif Bishop and current_position in {elaba28yry8ory, elaba2basyly, byhmot8ar, hytynyat, embiniot, 
                                             elshokr, tomakario, eya8aby, esbasmos1, esbasmos2, maro, 
                                             mrdElkatholikon, tobhyna}:
            source_slide = presentation3.Slides(current_end_slide)
            source_slide.Copy()
            if current_position in {elaba28yry8ory, elaba2basyly, byhmot8ar, hytynyat, embiniot, elshokr, maro}:
                presentation1.Slides.Paste(current_position).SlideShowTransition.Hidden = False
            else:
                presentation1.Slides.Paste(current_position).SlideShowTransition.Hidden = True
            current_end_slide -= 1
            if (current_position == maro) and (current_start_slide > current_end_slide):
                    # Perform the duplication twice
                    for _ in range(2):  # Loop to duplicate the section twice
                        # Copy the first slide (agiosElmilad) and paste it after the second slide (agiosElmilad2)
                        agiosElmilad_slide = presentation1.Slides(agiosEl2yama)
                        agiosElmilad_slide.Copy()
                        presentation1.Slides.Paste(agiosEl2yama2 + 1)

                        # Copy the second slide (agiosElmilad2) and paste it after the newly copied first slide
                        agiosElmilad2_slide = presentation1.Slides(agiosEl2yama2)
                        agiosElmilad2_slide.Copy()
                        presentation1.Slides.Paste(agiosEl2yama2 + 2)  # Adjust to account for new first slide
            
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

    move_section_names = [
        "{2F9B712F-EC2B-474A-B445-B4A46D96A229}",
        "{1CDFD5FF-8DF5-48CF-A773-F5EEBD56469C}",
        "{9888DFF9-9A38-4DB9-80E8-E000DA07D088}", 
        "{F0C07D6A-9B65-4AF5-AF75-7BA80C1EEEFC}"
    ]

    target_section_names = [
        "{22F83DFC-792B-4148-8AED-E77703B6E7BB}", 
        "{2F9B712F-EC2B-474A-B445-B4A46D96A229}",
        "{2EB8E5B7-6049-402F-80E4-ED6EDFAB83F4}",
        "{9888DFF9-9A38-4DB9-80E8-E000DA07D088}"
    ]

    # Call the function once for all moves
    move_sections_v2(presentation1, move_section_names, target_section_names)
 
    presentation2.Close()
    if guestBishop>0:
        presentation3.Close()

    presentation1.SlideShowSettings.Run()

def odasElso3od(copticdate, Bishop=False, guestBishop=0, afterSo3od=False):
    prs1 = relative_path(r"قداس.pptx")  # Using the relative path
    excel = relative_path(r"بيانات القداسات.xlsx")
    excel2 = relative_path(r"Tables.xlsx")
    des_sheet ="القداس"
    image = relative_path(r"Data\Designs\القيامة.png")
    el2yamaText =  ["لأنك قمت","aktwnk", "آك طونك"]
    replacefile(prs1, relative_path(r"Data\CopyData\قداس.pptx"))

    prs2 = relative_path(r"Data\القطمارس\قطمارس الخماسين (القداس).pptx")
    katamars_sheet = "قطمارس الخماسين"
    km = copticdate[1]
    kd = copticdate[2]
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

    # elso3od_show_full_sections = ["الليلويا فاي بيبي", "طاي شوري", "هيتينية الملاك ميخائيل - للقيامة", 
    #                               "هيتينية القيامة", "مرد ابركسيس الصعود", "اجيوس الصعود", "مرد مزمور الصعود",
    #                               "فاي اريه بي اوو", "اوشية اهوية السماء", "اوشية اهوية السماء غ",
    #                               "قسمة للآب في عيد القيامة والخمسين (أيها السيد الرب الإله ضابط الكل)",
    #                               "مدائح الخماسين من 40 الى 50"]
    
    # elso3od_hide_full_sections = ["السنكسار", "اكسيا", "اجيوس الميلاد", "اجيوس الصلب", 
    #                               "مرد المزمور", "مرد الانجيل", "ربع للعذراء", 
    #                               "قسمة القداس الباسيلي (أيها السيد الرب إلهنا)", "اك اسماروؤت", "بي اويك"]

    # if afterSo3od == False:
    #     elso3od_show_full_sections_ranges = [["افريك اتفيه (الصعود)", "محير عيد الصعود"]]
    #     elso3od_show_full_sections.extend(["المزمور السنجاري", "مزمور الصعود قبطي"])
    # else:
    #     elso3od_show_full_sections.extend(["محير عيد الصعود"])

    elso3od_show_full_sections = ['{072F3D96-A6C8-405F-9A23-7CCA1B2F13FF}', '{20F525FD-C708-4DDD-8E40-FE502EFEBDDE}', '{51C5460B-5CB6-46E7-A00C-E417B4008CBC}', '{E0F28C9F-5BF8-4AFC-806F-716597B95865}', '{F7D12816-345D-4511-B2FB-BC3FEED2DFAF}', '{CD4B95FF-0E0E-42D3-8DC4-224C3DD732F7}', '{5CE9710D-A04D-4552-9FA4-F2D6982982DA}', '{B7D98377-B994-4654-B49C-DE10E0DDE4F1}', '{BC7E3DCD-6AA8-44CC-B8AF-BC3E2BC71B5A}', '{A20DA654-32F7-4B4C-96CB-C76232EB96E8}', '{3E1960C6-4E54-4563-B3A6-E8AAB1680607}', '{BAAC56F0-C9A0-4321-AB43-D79F5FCE37C9}', '{973DBBAC-E645-4981-B3F5-5DB1413508D0}', '{5689FB59-3F5F-4FDD-A02B-5FF2599B7613}']
    
    elso3od_hide_full_sections = ['{1DA5C6AA-2FE5-461B-9E2C-40113CFC7804}', '{A8A52E1F-44DD-45E1-A737-4E13E15D5F1F}', '{76175201-2586-46AD-8E9F-C69E29FE2620}', '{31685B5B-48C4-437E-858C-CF8D225C0C26}', '{F13A48F2-238D-4617-B84E-9B0A694D9A18}', '{06D592C8-05BF-4B7C-86F7-FDAB3FAB5FB1}', '{ECE652ED-1345-4C6D-B92D-5996CFA27AEE}', '{681FF6A7-4230-4171-8F41-83FD64E8C960}', '{507EFD97-98F8-4376-848B-20D72E16D2C1}', '{B9A30F5E-0C89-471B-A99A-23DBE7F58504}']

    if afterSo3od == False:
        elso3od_show_full_sections_ranges = [['{F7747767-168A-418D-A4D8-634840EBF13E}', '{2BB9A97E-3E87-49D7-9213-0C5CE26DD939}']]
        elso3od_show_full_sections.extend(['{973DBBAC-E645-4981-B3F5-5DB1413508D0}', '{5689FB59-3F5F-4FDD-A02B-5FF2599B7613}'])
    else:
        elso3od_show_full_sections.extend(['{2BB9A97E-3E87-49D7-9213-0C5CE26DD939}'])

    # elso3od_values = ["تكملة على حسب المناسبة", "مزمور التوزيع", "مزمور التوزيع", "مرد توزيع الصعود",
    #                   "الانجيل", "المزمور", "الابركسيس", "الكاثوليكون", "البولس عربي",
    #                   "اجيوس الصعود", "اجيوس الصعود", "مرد انجيل الصعود", "بدء قداس الكلمة"]
    
    elso3od_values = find_slide_nums_arrays_v2(excel, des_sheet, 
                     ['{A18EDC94-F257-4FAC-99C7-0A8EA70F0FAF}', '{C29E5A83-A98B-4077-8194-99A6D803EF53}', '{C29E5A83-A98B-4077-8194-99A6D803EF53}', '{834F8F17-CEC0-4E51-A927-8074D22B6A78}', '{C7D4A109-F792-4661-BAD0-075FD1A1909F}', '{B74DBB8C-2B2D-46E4-9508-DA46008D19A4}', '{E234C6C7-3837-4CE4-A541-CDC9627AAAC2}', '{6D4B3F52-63BF-435F-BF0C-C9D41120C2A3}', '{D88055F5-EAA0-4C8E-8249-C364A572BF7B}', '{CD4B95FF-0E0E-42D3-8DC4-224C3DD732F7}', '{CD4B95FF-0E0E-42D3-8DC4-224C3DD732F7}', '{FDACCDF0-22DF-42A8-B31E-94F2B97DAC4B}', '{C08D8D44-E49E-47CE-8027-C8AE26B1AA9A}'],
                     2, [1, 1, 2, 1, 2, 2, 2, 2, 2, 1, 2, 2, 2])

    #الختام
    elkhetam = elso3od_values[0]
    #التوزيع
    mazmorELtawzy3 = elso3od_values[1] + 1
    mazmorELtawzy32 = elso3od_values[2] - 1
    mrdMazmorEltawzy3 = elso3od_values[3]
    
    #القرائات
    elengil3 = elso3od_values[4]
    elmazmor3 = elso3od_values[5]
    elebrksis3 = elso3od_values[6]
    elkatholikon3 = elso3od_values[7]
    elbouls3 = elso3od_values[8]

    #اجيوس الصعود
    agiosElso3od = elso3od_values[9]    
    agiosElso3od2 = elso3od_values[10]

    #مرد انجيل الصعود
    mrdElso3od = elso3od_values[11]

    if Bishop == True:
        prs3 = relative_path(r"Data\حضور الأسقف.pptx")
        sheet = "في حضور الأسقف"

        bishop_show_values = ['{F76B0D75-0474-45B5-B79F-7416F354543A}', '{E2968C91-5339-499C-9812-DECCCF58A2CD}',
                              '{62A12AF8-CB6D-4CC5-9DB0-B73A7C24E2AD}', '{B74DBB8C-2B2D-46E4-9508-DA46008D19A4}',
                              '{A9183893-7B7E-459F-8547-F7A8F7D2D521}']
        elso3od_show_full_sections.extend(bishop_show_values)

        if guestBishop > 0:
            # bishop_values = ["صلاة الشكر", "صلاة الشكر", "طوبه هينا الكبيرة", "طوبه هينا الكبيرة", 
            #                  "نيم بينيوت", "نيم بينيوت", "الهيتنيات", "الهيتنيات",
            #                  "اي اغابي", "اي اغابي", "مرد الكاثوليكون", "مرد الكاثوليكون",
            #                  "الاسبسمس", "الاسبسمس"]

            # bishopDes_values = ["تكملة في حضور الاسقف", "طوبه هينا الكبيرة",
            #                     "امبين يوت اتطايوت", "تكملة الهيتنيات", "بى اهموت غار الصغيرة",
            #                     "تو ماكريو", "اي اغابي", "ابيت جيك ايفول", "مارو اتشاسف",
            #                     "الاسبسمس الادام السنوي+الختام", "ختام الاسبسمس الادام",
            #                     "اوشية الاباء (ب)", "اوشية الاباء غ"]
            
            bishop_values = find_slide_nums_arrays_v2(excel, sheet, 
                            ['{6851F163-CBEF-4014-A853-CE100557BA6A}', '{6851F163-CBEF-4014-A853-CE100557BA6A}', '{B084BC40-61E1-4477-98DA-15CFB06AEE91}', '{B084BC40-61E1-4477-98DA-15CFB06AEE91}', '{97203297-EECB-4D41-B2E3-AD9A4863847E}', '{97203297-EECB-4D41-B2E3-AD9A4863847E}', '{7C84083F-E6D3-4669-9130-AC7E8D935A98}', '{7C84083F-E6D3-4669-9130-AC7E8D935A98}', '{D0234C99-69FE-407A-82A9-D7A676919E93}', '{D0234C99-69FE-407A-82A9-D7A676919E93}', '{0B519645-F935-43AE-A9CE-6E2FC03833BB}', '{0B519645-F935-43AE-A9CE-6E2FC03833BB}', '{D1378DB5-29D1-4800-9D96-10F2535EEB57}', '{D1378DB5-29D1-4800-9D96-10F2535EEB57}'], 
                            2, [1, 2, 1, 2, 1, 2, 1, 2, 1, 2, 1, 2, 1, 2])

            bishopDes_values = find_slide_nums_arrays_v2(excel, des_sheet, 
                                ['{F76B0D75-0474-45B5-B79F-7416F354543A}', '{8DD21CDE-CB6B-4D5B-B995-D2747AB69ED1}', '{E2968C91-5339-499C-9812-DECCCF58A2CD}', '{646A8184-7F05-453A-A2F1-EB9A77D7F0EE}', '{21891CBB-A1EC-4974-B0B6-F74A4B502BC2}', '{12B7D244-BF4C-401B-A65A-D1621D7DD953}', '{F69B50D8-FB5E-4E8C-AE5C-6DCB4790AFAF}', '{38267404-7625-47AA-B0C8-31BCA5D0435D}', '{62A12AF8-CB6D-4CC5-9DB0-B73A7C24E2AD}', '{0C7A7725-643D-4F7E-A2F0-0C8A36C2A594}', '{1CDFD5FF-8DF5-48CF-A773-F5EEBD56469C}', '{74A33555-8E08-47DF-B3CD-A1B4C7AF2B4E}', '{474487EB-4554-4A30-B351-6EF762D2F2D6}'],                  
                                2, [2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2])

            elshokr = bishopDes_values[0]
            elshokr1 = bishop_values[0]
            elshokr2 = bishop_values[1]

            tobhyna = bishopDes_values[1]
            tobhyna1 = bishop_values[2]
            tobhyna2 = bishop_values[3]

            embiniot = bishopDes_values[2]
            embiniot1 = bishop_values[4]
            embiniot2 = bishop_values[5]

            hytynyat = bishopDes_values[3]
            hytynyat1 = bishop_values[6]
            hytynyat2 = bishop_values[7]

            byhmot8ar = bishopDes_values[4]-1
            byhmot8ar1 = bishop_values[4]
            byhmot8ar2 = bishop_values[5]

            tomakario = bishopDes_values[5] - 1
            tomakario1 = bishop_values[4]
            tomakario2 = bishop_values[5]

            eya8aby = bishopDes_values[6] - 1
            eya8aby1 = bishop_values[8]
            eya8aby2 = bishop_values[9]

            mrdElkatholikon = bishopDes_values[7]
            mrdElkatholikon1 = bishop_values[10]
            mrdElkatholikon2 = bishop_values[11]

            maro = bishopDes_values[8]
            maro1 = bishop_values[4]
            maro2 = bishop_values[5]

            esbasmos1 = bishopDes_values[9]
            esbasmos11 = bishop_values[12]
            esbasmos12 = bishop_values[13]

            esbasmos2 = bishopDes_values[10]
            esbasmos21 = bishop_values[12]
            esbasmos22 = bishop_values[13]

            elaba2basyly = bishopDes_values[11] - 1
            elaba2basyly1 = bishop_values[0]
            elaba2basyly2 = bishop_values[1]

            elaba28yry8ory = bishopDes_values[12] - 1
            elaba28yry8ory1 = bishop_values[0]
            elaba28yry8ory2 = bishop_values[1]

            if guestBishop < 2:
                elshokr2 = elshokr2-1
                tobhyna2 = tobhyna2-2
                embiniot2 = embiniot2-1
                hytynyat2 = hytynyat2-3
                byhmot8ar2 = byhmot8ar2-1
                tomakario2 = tomakario2-1
                eya8aby2 = eya8aby2-2
                mrdElkatholikon2 = mrdElkatholikon2-1
                maro2 = maro2-1
                esbasmos12 = esbasmos12-2
                esbasmos22 = esbasmos22-2
                elaba2basyly2 = elaba2basyly2-1
                elaba28yry8ory2 = elaba28yry8ory2-1
        
            start_positions = [mazmorELtawzy3, elaba28yry8ory, elaba2basyly, esbasmos2, esbasmos1, elengil3, elmazmor3, maro, elebrksis3, elkatholikon3, mrdElkatholikon, elbouls3, eya8aby, tomakario, byhmot8ar, hytynyat, embiniot, tobhyna, elshokr]
            start_slides = [mrdMazmorEltawzy3, elaba28yry8ory1, elaba2basyly1, esbasmos21, esbasmos11, elengil, elmzmor, maro1, elebrksis1, elkatholikon1, mrdElkatholikon1, elbouls1, eya8aby1, tomakario1, byhmot8ar1, hytynyat1, embiniot1, tobhyna1, elshokr1]
            end_slides = [mrdMazmorEltawzy3, elaba28yry8ory2, elaba2basyly2, esbasmos22, esbasmos12, elengil2, elmzmor, maro2, elebrksis2, elkatholikon2, mrdElkatholikon2, elbouls2, eya8aby2, tomakario2, byhmot8ar2, hytynyat2, embiniot2, tobhyna2, elshokr2]

    if guestBishop == 0:
        start_positions = [mazmorELtawzy3, elengil3, elmazmor3, elebrksis3, elkatholikon3, elbouls3]
        start_slides = [mrdMazmorEltawzy3, elengil, elmzmor, elebrksis1, elkatholikon1, elbouls1]
        end_slides = [mrdMazmorEltawzy3, elengil2, elmzmor, elebrksis2, elkatholikon2, elbouls2]

    show_hide_insertImage_replaceText(prs1, excel, des_sheet, elso3od_show_full_sections, elso3od_hide_full_sections, elso3od_show_full_sections_ranges, None, image, el2yamaText)

    powerpoint = win32com.client.Dispatch("PowerPoint.Application")
    powerpoint.Visible = True  # Open PowerPoint application
    presentation1 = open_presentation_relative_path(prs1)
    presentation2 = open_presentation_relative_path(prs2)
    
    if guestBishop > 0:
        insert_image_to_slides_same_file(prs3, image)
        presentation3 = open_presentation_relative_path(prs3)

    khetamElso3od = find_slide_index_by_title(presentation1, "الصعود", elkhetam)
    show_slides(presentation1, [[khetamElso3od, khetamElso3od], [mrdElso3od, mrdElso3od]])
    run_vba_with_slide_id(excel, des_sheet, prs1, presentation1)
    agbya(presentation1, elso3od_values[12], 1)

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
                new_slide = presentation1.Slides.Paste(slide_index1).SlideShowTransition.Hidden = False
                slide_index1 -= 1
            current_start_slide += 1

        elif (current_position == elengil3 or current_position == elmazmor3 or current_position == elebrksis3 
            or current_position == elkatholikon3 or current_position == elbouls3):
            source_slide = presentation2.Slides(current_end_slide)
            source_slide.Copy()
            new_slide = presentation1.Slides.Paste(current_position)
            pic_shape = new_slide.Shapes.AddPicture(image, LinkToFile=False, SaveWithDocument=True, Left=0, Top=144, Width=720, Height=396)
            pic_shape.ZOrder(1)
            new_slide.SlideShowTransition.Hidden = False
            current_end_slide -= 1
            # Add slide duplication after elmzmor1 is processed
            if guestBishop == 0 and current_position == elmazmor3:
                # Perform the duplication twice
                for _ in range(2):  # Loop to duplicate the section twice
                    # Copy the first slide (agiosElsalyb) and paste it after the second slide (agiosElsalyb2)
                    agiosElmilad_slide = presentation1.Slides(agiosElso3od)
                    agiosElmilad_slide.Copy()
                    presentation1.Slides.Paste(agiosElso3od2 + 1)

                    # Copy the second slide (agiosElsalyb2) and paste it after the newly copied first slide
                    agiosElmilad2_slide = presentation1.Slides(agiosElso3od2)
                    agiosElmilad2_slide.Copy()
                    presentation1.Slides.Paste(agiosElso3od2 + 2)  # Adjust to account for new first slide
            if current_start_slide > current_end_slide:
                current_position += 1

        elif Bishop and current_position in {elaba28yry8ory, elaba2basyly, byhmot8ar, hytynyat, embiniot, 
                                             elshokr, tomakario, eya8aby, esbasmos1, esbasmos2, maro, 
                                             mrdElkatholikon, tobhyna}:
            source_slide = presentation3.Slides(current_end_slide)
            source_slide.Copy()
            if current_position in {elaba28yry8ory, elaba2basyly, byhmot8ar, hytynyat, embiniot, elshokr, maro}:
                presentation1.Slides.Paste(current_position).SlideShowTransition.Hidden = False
            else:
                presentation1.Slides.Paste(current_position).SlideShowTransition.Hidden = True
            current_end_slide -= 1
            if (current_position == maro) and (current_start_slide > current_end_slide):
                    # Perform the duplication twice
                    for _ in range(2):  # Loop to duplicate the section twice
                        # Copy the first slide (agiosElmilad) and paste it after the second slide (agiosElmilad2)
                        agiosElmilad_slide = presentation1.Slides(agiosElso3od)
                        agiosElmilad_slide.Copy()
                        presentation1.Slides.Paste(agiosElso3od2 + 1)

                        # Copy the second slide (agiosElmilad2) and paste it after the newly copied first slide
                        agiosElmilad2_slide = presentation1.Slides(agiosElso3od2)
                        agiosElmilad2_slide.Copy()
                        presentation1.Slides.Paste(agiosElso3od2 + 2)  # Adjust to account for new first slide
            
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

    move_section_names = [
        "{A91FBB00-3AAB-4A9A-8A22-6E086F37BCF1}",
        "{1CDFD5FF-8DF5-48CF-A773-F5EEBD56469C}",
        "{BA393839-A242-4557-8097-032911C32D90}", 
        "{F0C07D6A-9B65-4AF5-AF75-7BA80C1EEEFC}"
    ]

    target_section_names = [
        "{22F83DFC-792B-4148-8AED-E77703B6E7BB}", 
        "{A91FBB00-3AAB-4A9A-8A22-6E086F37BCF1}",
        "{2EB8E5B7-6049-402F-80E4-ED6EDFAB83F4}",
        "{BA393839-A242-4557-8097-032911C32D90}"
    ]

    # Call the function once for all moves
    move_sections_v2(presentation1, move_section_names, target_section_names)
 
    presentation2.Close()
    if guestBishop>0:
        presentation3.Close()

    presentation1.SlideShowSettings.Run()

def odasEl3nsara(copticdate, Bishop=False, guestBishop=0):
    prs1 = relative_path(r"قداس.pptx")  # Using the relative path
    excel = relative_path(r"بيانات القداسات.xlsx")
    excel2 = relative_path(r"Tables.xlsx")
    des_sheet ="القداس"
    image = relative_path(r"Data\Designs\القيامة.png")
    el2yamaText =  ["لأنك قمت","aktwnk", "آك طونك"]
    replacefile(prs1, relative_path(r"Data\CopyData\قداس.pptx"))

    prs2 = relative_path(r"Data\القطمارس\قطمارس الخماسين (القداس).pptx")
    katamars_sheet = "قطمارس الخماسين"
    km = copticdate[1]
    kd = copticdate[2]
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

    # el3nsara_show_full_sections = ["الليلويا فاي بيبي", "طاي شوري", "هيتينية الملاك ميخائيل - للقيامة", 
    #                         "هيتينية القيامة", "مرد ابركسيس العنصرة", "قطع الساعة الثالثة (عيد العنصرة)",
    #                         "بي ابنفما", "محير عيد العنصرة", "تكملة مشتركة للمحير",
    #                         "المزمور السنجاري", "مزمور العنصرة قبطي", "مرد مزمور العنصرة",
    #                         "مرد انجيل العنصرة", "فاي اريه بي اوو", "اوشية اهوية السماء", 
    #                         "اوشية اهوية السماء غ", "قسمة للآب في عيد القيامة والخمسين (أيها السيد الرب الإله ضابط الكل)",
    #                         "اسومين" ,"مدائح الصعود الى العنصرة"]

    # el3nsara_hide_full_sections = ["السنكسار", "اكسيا", "اجيوس الميلاد", "اجيوس الصلب", 
    #                         "مرد المزمور", "مرد الانجيل", "ربع للعذراء", 
    #                         "قسمة القداس الباسيلي (أيها السيد الرب إلهنا)", "اك اسماروؤت", "بي اويك"]

    el3nsara_show_full_sections = ['{072F3D96-A6C8-405F-9A23-7CCA1B2F13FF}', '{20F525FD-C708-4DDD-8E40-FE502EFEBDDE}', '{51C5460B-5CB6-46E7-A00C-E417B4008CBC}', '{E0F28C9F-5BF8-4AFC-806F-716597B95865}', '{FF41886A-B825-4068-8D01-024122EEE9DA}', '{F892D3CF-2F70-4428-84F8-3D3377DF054E}', '{EC1819FB-4FF0-45EF-8E3F-112BDA630463}', '{8D5BB4B9-76DC-4905-9269-DBA484E740E4}', '{DEDC0CCA-3854-4E18-8CB2-5D6FEC5BABCC}', '{973DBBAC-E645-4981-B3F5-5DB1413508D0}', '{FEE062D9-82B3-4A9D-ACF8-F1E05816BC39}', '{6477B337-8361-45A9-80C3-7E1B17402D4D}', '{CF7B71E2-5C12-470D-A07B-80A9F0BF2670}', '{B7D98377-B994-4654-B49C-DE10E0DDE4F1}', '{BC7E3DCD-6AA8-44CC-B8AF-BC3E2BC71B5A}', '{A20DA654-32F7-4B4C-96CB-C76232EB96E8}', '{3E1960C6-4E54-4563-B3A6-E8AAB1680607}', '{3F22DD0B-D926-42B9-BABE-6C6486C9BD62}', '{BAAC56F0-C9A0-4321-AB43-D79F5FCE37C9}']

    el3nsara_hide_full_sections = ['{1DA5C6AA-2FE5-461B-9E2C-40113CFC7804}', '{A8A52E1F-44DD-45E1-A737-4E13E15D5F1F}', '{76175201-2586-46AD-8E9F-C69E29FE2620}', '{31685B5B-48C4-437E-858C-CF8D225C0C26}', '{F13A48F2-238D-4617-B84E-9B0A694D9A18}', '{06D592C8-05BF-4B7C-86F7-FDAB3FAB5FB1}', '{ECE652ED-1345-4C6D-B92D-5996CFA27AEE}', '{681FF6A7-4230-4171-8F41-83FD64E8C960}', '{507EFD97-98F8-4376-848B-20D72E16D2C1}', '{B9A30F5E-0C89-471B-A99A-23DBE7F58504}']

    el3nsara_values = ["تكملة على حسب المناسبة", "مزمور التوزيع", "مزمور التوزيع", "مرد توزيع العنصرة",
                       "الانجيل", "المزمور", "الابركسيس", "الكاثوليكون", "البولس عربي",
                       "اجيوس الصعود", "اجيوس الصعود", "بدء قداس الكلمة", "شا ني رومبي - الى منتهى الاعوام"]
    
    el3nsara_values = find_slide_nums_arrays_v2(excel, des_sheet, 
                      ['{A18EDC94-F257-4FAC-99C7-0A8EA70F0FAF}', '{C29E5A83-A98B-4077-8194-99A6D803EF53}', '{C29E5A83-A98B-4077-8194-99A6D803EF53}', '{22674B27-BBD5-49D8-98B9-ABCA6F4C5504}', '{C7D4A109-F792-4661-BAD0-075FD1A1909F}', '{B74DBB8C-2B2D-46E4-9508-DA46008D19A4}', '{E234C6C7-3837-4CE4-A541-CDC9627AAAC2}', '{6D4B3F52-63BF-435F-BF0C-C9D41120C2A3}', '{D88055F5-EAA0-4C8E-8249-C364A572BF7B}', '{CD4B95FF-0E0E-42D3-8DC4-224C3DD732F7}', '{CD4B95FF-0E0E-42D3-8DC4-224C3DD732F7}', '{C08D8D44-E49E-47CE-8027-C8AE26B1AA9A}', '{C193E31F-DBD7-4EEA-98EA-FBBA1D3F1186}'], 
                      2, [1, 1, 2, 1, 2, 2, 2, 2, 2, 1, 2, 2, 2])

    #الختام
    elkhetam = el3nsara_values[0]
    #التوزيع
    mazmorELtawzy3 = el3nsara_values[1] + 1
    mazmorELtawzy32 = el3nsara_values[2] - 1
    mrdMazmorEltawzy3 = el3nsara_values[3]
    
    #القرائات
    elengil3 = el3nsara_values[4]
    elmazmor3 = el3nsara_values[5]
    elebrksis3 = el3nsara_values[6]
    elkatholikon3 = el3nsara_values[7]
    elbouls3 = el3nsara_values[8]

    #اجيوس الصعود
    agiosElso3od = el3nsara_values[9]    
    agiosElso3od2 = el3nsara_values[10]

    #مرد انجيل الصعود
    shanirompy = el3nsara_values[12]


    if Bishop == True:
        prs3 = relative_path(r"Data\حضور الأسقف.pptx")
        sheet = "في حضور الأسقف"

        bishop_show_values = ['{F76B0D75-0474-45B5-B79F-7416F354543A}', '{E2968C91-5339-499C-9812-DECCCF58A2CD}',
                              '{62A12AF8-CB6D-4CC5-9DB0-B73A7C24E2AD}', '{B74DBB8C-2B2D-46E4-9508-DA46008D19A4}',
                              '{A9183893-7B7E-459F-8547-F7A8F7D2D521}', '{C193E31F-DBD7-4EEA-98EA-FBBA1D3F1186}']
        el3nsara_show_full_sections.extend(bishop_show_values)

        if guestBishop > 0:
            # bishop_values = ["صلاة الشكر", "صلاة الشكر", "طوبه هينا الكبيرة", "طوبه هينا الكبيرة", 
            #                  "نيم بينيوت", "نيم بينيوت", "الهيتنيات", "الهيتنيات",
            #                  "اي اغابي", "اي اغابي", "مرد الكاثوليكون", "مرد الكاثوليكون",
            #                  "الاسبسمس", "الاسبسمس"]

            # bishopDes_values = ["تكملة في حضور الاسقف", "طوبه هينا الكبيرة",
            #                     "امبين يوت اتطايوت", "تكملة الهيتنيات", "بى اهموت غار الصغيرة",
            #                     "تو ماكريو", "اي اغابي", "ابيت جيك ايفول", "مارو اتشاسف",
            #                     "الاسبسمس الادام السنوي+الختام", "ختام الاسبسمس الادام",
            #                     "اوشية الاباء (ب)", "اوشية الاباء غ"]
            
            bishop_values = find_slide_nums_arrays_v2(excel, sheet, 
                            ['{6851F163-CBEF-4014-A853-CE100557BA6A}', '{6851F163-CBEF-4014-A853-CE100557BA6A}', '{B084BC40-61E1-4477-98DA-15CFB06AEE91}', '{B084BC40-61E1-4477-98DA-15CFB06AEE91}', '{97203297-EECB-4D41-B2E3-AD9A4863847E}', '{97203297-EECB-4D41-B2E3-AD9A4863847E}', '{7C84083F-E6D3-4669-9130-AC7E8D935A98}', '{7C84083F-E6D3-4669-9130-AC7E8D935A98}', '{D0234C99-69FE-407A-82A9-D7A676919E93}', '{D0234C99-69FE-407A-82A9-D7A676919E93}', '{0B519645-F935-43AE-A9CE-6E2FC03833BB}', '{0B519645-F935-43AE-A9CE-6E2FC03833BB}', '{D1378DB5-29D1-4800-9D96-10F2535EEB57}', '{D1378DB5-29D1-4800-9D96-10F2535EEB57}'], 
                            2, [1, 2, 1, 2, 1, 2, 1, 2, 1, 2, 1, 2, 1, 2])

            bishopDes_values = find_slide_nums_arrays_v2(excel, des_sheet, 
                                ['{F76B0D75-0474-45B5-B79F-7416F354543A}', '{8DD21CDE-CB6B-4D5B-B995-D2747AB69ED1}', '{E2968C91-5339-499C-9812-DECCCF58A2CD}', '{646A8184-7F05-453A-A2F1-EB9A77D7F0EE}', '{21891CBB-A1EC-4974-B0B6-F74A4B502BC2}', '{12B7D244-BF4C-401B-A65A-D1621D7DD953}', '{F69B50D8-FB5E-4E8C-AE5C-6DCB4790AFAF}', '{38267404-7625-47AA-B0C8-31BCA5D0435D}', '{62A12AF8-CB6D-4CC5-9DB0-B73A7C24E2AD}', '{0C7A7725-643D-4F7E-A2F0-0C8A36C2A594}', '{1CDFD5FF-8DF5-48CF-A773-F5EEBD56469C}', '{74A33555-8E08-47DF-B3CD-A1B4C7AF2B4E}', '{474487EB-4554-4A30-B351-6EF762D2F2D6}'],                  
                                2, [2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2])

            elshokr = bishopDes_values[0]
            elshokr1 = bishop_values[0]
            elshokr2 = bishop_values[1]

            tobhyna = bishopDes_values[1]
            tobhyna1 = bishop_values[2]
            tobhyna2 = bishop_values[3]

            embiniot = bishopDes_values[2]
            embiniot1 = bishop_values[4]
            embiniot2 = bishop_values[5]

            hytynyat = bishopDes_values[3]
            hytynyat1 = bishop_values[6]
            hytynyat2 = bishop_values[7]

            byhmot8ar = bishopDes_values[4]-1
            byhmot8ar1 = bishop_values[4]
            byhmot8ar2 = bishop_values[5]

            tomakario = bishopDes_values[5] - 1
            tomakario1 = bishop_values[4]
            tomakario2 = bishop_values[5]

            eya8aby = bishopDes_values[6] - 1
            eya8aby1 = bishop_values[8]
            eya8aby2 = bishop_values[9]

            mrdElkatholikon = bishopDes_values[7]
            mrdElkatholikon1 = bishop_values[10]
            mrdElkatholikon2 = bishop_values[11]

            maro = bishopDes_values[8]
            maro1 = bishop_values[4]
            maro2 = bishop_values[5]

            esbasmos1 = bishopDes_values[9]
            esbasmos11 = bishop_values[12]
            esbasmos12 = bishop_values[13]

            esbasmos2 = bishopDes_values[10]
            esbasmos21 = bishop_values[12]
            esbasmos22 = bishop_values[13]

            elaba2basyly = bishopDes_values[11] - 1
            elaba2basyly1 = bishop_values[0]
            elaba2basyly2 = bishop_values[1]

            elaba28yry8ory = bishopDes_values[12] - 1
            elaba28yry8ory1 = bishop_values[0]
            elaba28yry8ory2 = bishop_values[1]

            if guestBishop < 2:
                elshokr2 = elshokr2-1
                tobhyna2 = tobhyna2-2
                embiniot2 = embiniot2-1
                hytynyat2 = hytynyat2-3
                byhmot8ar2 = byhmot8ar2-1
                tomakario2 = tomakario2-1
                eya8aby2 = eya8aby2-2
                mrdElkatholikon2 = mrdElkatholikon2-1
                maro2 = maro2-1
                esbasmos12 = esbasmos12-2
                esbasmos22 = esbasmos22-2
                elaba2basyly2 = elaba2basyly2-1
                elaba28yry8ory2 = elaba28yry8ory2-1
        
            start_positions = [mazmorELtawzy3, elaba28yry8ory, elaba2basyly, esbasmos2, esbasmos1, elengil3, elmazmor3, maro, elebrksis3, elkatholikon3, mrdElkatholikon, elbouls3, eya8aby, tomakario, byhmot8ar, hytynyat, embiniot, tobhyna, elshokr]
            start_slides = [mrdMazmorEltawzy3, elaba28yry8ory1, elaba2basyly1, esbasmos21, esbasmos11, elengil, elmzmor, maro1, elebrksis1, elkatholikon1, mrdElkatholikon1, elbouls1, eya8aby1, tomakario1, byhmot8ar1, hytynyat1, embiniot1, tobhyna1, elshokr1]
            end_slides = [mrdMazmorEltawzy3, elaba28yry8ory2, elaba2basyly2, esbasmos22, esbasmos12, elengil2, elmzmor, maro2, elebrksis2, elkatholikon2, mrdElkatholikon2, elbouls2, eya8aby2, tomakario2, byhmot8ar2, hytynyat2, embiniot2, tobhyna2, elshokr2]

    if guestBishop == 0:
        start_positions = [mazmorELtawzy3, elengil3, elmazmor3, elebrksis3, elkatholikon3, elbouls3]
        start_slides = [mrdMazmorEltawzy3, elengil, elmzmor, elebrksis1, elkatholikon1, elbouls1]
        end_slides = [mrdMazmorEltawzy3, elengil2, elmzmor, elebrksis2, elkatholikon2, elbouls2]

    show_hide_insertImage_replaceText(prs1, excel, des_sheet, el3nsara_show_full_sections, el3nsara_hide_full_sections, None, None, image, el2yamaText)

    powerpoint = win32com.client.Dispatch("PowerPoint.Application")
    powerpoint.Visible = True  # Open PowerPoint application
    presentation1 = open_presentation_relative_path(prs1)
    presentation2 = open_presentation_relative_path(prs2)
    
    if guestBishop > 0:
        insert_image_to_slides_same_file(prs3, image)
        presentation3 = open_presentation_relative_path(prs3)

    khetamEl3nsara = find_slide_index_by_title(presentation1, "العنصرة", elkhetam)
    show_slides(presentation1, [[khetamEl3nsara, khetamEl3nsara], [shanirompy, shanirompy]])
    run_vba_with_slide_id(excel, des_sheet, prs1, presentation1)
    agbya(presentation1, el3nsara_values[11], 4)
    
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
                new_slide = presentation1.Slides.Paste(slide_index1).SlideShowTransition.Hidden = False
                slide_index1 -= 1
            current_start_slide += 1

        elif (current_position == elengil3 or current_position == elmazmor3 or current_position == elebrksis3 
            or current_position == elkatholikon3 or current_position == elbouls3):
            source_slide = presentation2.Slides(current_end_slide)
            source_slide.Copy()
            new_slide = presentation1.Slides.Paste(current_position)
            pic_shape = new_slide.Shapes.AddPicture(image, LinkToFile=False, SaveWithDocument=True, Left=0, Top=144, Width=720, Height=396)
            pic_shape.ZOrder(1)
            new_slide.SlideShowTransition.Hidden = False
            current_end_slide -= 1
            # Add slide duplication after elmzmor1 is processed
            if guestBishop == 0 and current_position == elmazmor3:
                # Perform the duplication twice
                for _ in range(2):  # Loop to duplicate the section twice
                    # Copy the first slide (agiosElsalyb) and paste it after the second slide (agiosElsalyb2)
                    agiosElmilad_slide = presentation1.Slides(agiosElso3od)
                    agiosElmilad_slide.Copy()
                    presentation1.Slides.Paste(agiosElso3od2 + 1)

                    # Copy the second slide (agiosElsalyb2) and paste it after the newly copied first slide
                    agiosElmilad2_slide = presentation1.Slides(agiosElso3od2)
                    agiosElmilad2_slide.Copy()
                    presentation1.Slides.Paste(agiosElso3od2 + 2)  # Adjust to account for new first slide
            if current_start_slide > current_end_slide:
                current_position += 1

        elif Bishop and current_position in {elaba28yry8ory, elaba2basyly, byhmot8ar, hytynyat, embiniot, 
                                             elshokr, tomakario, eya8aby, esbasmos1, esbasmos2, maro, 
                                             mrdElkatholikon, tobhyna}:
            source_slide = presentation3.Slides(current_end_slide)
            source_slide.Copy()
            if current_position in {elaba28yry8ory, elaba2basyly, byhmot8ar, hytynyat, embiniot, elshokr, maro}:
                presentation1.Slides.Paste(current_position).SlideShowTransition.Hidden = False
            else:
                presentation1.Slides.Paste(current_position).SlideShowTransition.Hidden = True
            current_end_slide -= 1
            if (current_position == maro) and (current_start_slide > current_end_slide):
                    # Perform the duplication twice
                    for _ in range(2):  # Loop to duplicate the section twice
                        # Copy the first slide (agiosElmilad) and paste it after the second slide (agiosElmilad2)
                        agiosElmilad_slide = presentation1.Slides(agiosElso3od)
                        agiosElmilad_slide.Copy()
                        presentation1.Slides.Paste(agiosElso3od2 + 1)

                        # Copy the second slide (agiosElmilad2) and paste it after the newly copied first slide
                        agiosElmilad2_slide = presentation1.Slides(agiosElso3od2)
                        agiosElmilad2_slide.Copy()
                        presentation1.Slides.Paste(agiosElso3od2 + 2)  # Adjust to account for new first slide
            
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

    move_section_names = [
        "{5D584DBA-C59D-40EA-9739-C83DA37F8C12}",
        "{1CDFD5FF-8DF5-48CF-A773-F5EEBD56469C}",
        "{9C8B1B1B-D26E-4129-A376-E3FBC58BE596}", 
        "{F0C07D6A-9B65-4AF5-AF75-7BA80C1EEEFC}"
    ]

    target_section_names = [
        "{22F83DFC-792B-4148-8AED-E77703B6E7BB}", 
        "{5D584DBA-C59D-40EA-9739-C83DA37F8C12}",
        "{2EB8E5B7-6049-402F-80E4-ED6EDFAB83F4}",
        "{9C8B1B1B-D26E-4129-A376-E3FBC58BE596}"
    ]

    # Call the function once for all moves
    move_sections_v2(presentation1, move_section_names, target_section_names)
 
    presentation2.Close()
    if guestBishop>0:
        presentation3.Close()

    presentation1.SlideShowSettings.Run()

def odasDo5olElmasy7Masr(copticdate, Bishop=False, guestBishop=0):
    return
