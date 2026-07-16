from commonFunctions import *

def vba_code(excel, sheet, prs, presentation, slide_section_id='{A5B9CE2F-90E3-44D7-B22F-CAE6783C8E2F}', custom_show_name=None, custom_show_trigger_section_id=None, custom_show_return_section_id=None):
    def _to_slide_number(value, section_id):
        if isinstance(value, str):
            raise ValueError(
                f"Section ID '{section_id}' was not found in sheet '{sheet}'. "
                "Make sure the ID belongs to this sheet."
            )
        return int(value)

    def _slide_number_from_section(section_id, offset):
        return _to_slide_number(find_slide_num_v2(excel, sheet, section_id, 2, offset), section_id)
    
    slide = _slide_number_from_section(slide_section_id, 1)
    slide_id = get_slide_ids_by_number(prs, slide)

    # Access the VBA project
    vba_project = presentation.VBProject
    modules = vba_project.VBComponents

    # Add a new module to the VBA project
    new_module = modules.Add(1)  # 1 corresponds to a standard module

    custom_show_start_case = ""
    if custom_show_name is not None and custom_show_trigger_section_id is not None:
        custom_show_trigger_slide = _slide_number_from_section(custom_show_trigger_section_id, 1)
        custom_show_trigger_slide_id = get_slide_ids_by_number(prs, custom_show_trigger_slide)
        custom_show_start_case = f"""
        Case {custom_show_trigger_slide_id}
            ActivePresentation.SlideShowWindow.View.GotoNamedShow "{custom_show_name}"
"""

    custom_show_return_case = ""
    if custom_show_return_section_id is not None:
        custom_show_return_slide = _slide_number_from_section(custom_show_return_section_id, 2)
        custom_show_return_slide_id = get_slide_ids_by_number(prs, custom_show_return_slide)
        custom_show_return_case = f"""
        Case {custom_show_return_slide_id}
            ActivePresentation.SlideShowWindow.View.EndNamedShow
            StartSlideshow
            ActivePresentation.SlideShowWindow.View.GotoSlide GetSlideIndexByID({custom_show_return_slide_id})
"""

    vba_code = f"""
Dim visitedSlides As Collection ' Global collection to track visited slides

Sub OnSlideShowPageChange()
    Dim currentSlideNumber As Integer
    Dim targetShape As Shape

    ' Initialize the visitedSlides collection if it hasn't been created yet
    If visitedSlides Is Nothing Then
        Set visitedSlides = New Collection
    End If

    ' Get the current slide number in the slideshow view
    currentSlideID = ActivePresentation.SlideShowWindow.View.Slide.SlideID

    ' Check if the slide has already been visited (i.e., hyperlink followed)
    On Error Resume Next
    visitedSlides.Item currentSlideID
    If Err.Number = 0 Then
        ' Slide has already been visited; exit without doing anything
        Exit Sub
    End If
    On Error GoTo 0

    ' Use Select Case to handle actions on specific slides by slide number
    Select Case currentSlideID
        Case {slide_id} ' Replace with the slide number for the first target slide
            ' Attempt to locate and "click" the target shape
            Set targetShape = GetShapeByName("TextBox 2") ' Replace with your shape name
            
            If Not targetShape Is Nothing Then
                If targetShape.ActionSettings(ppMouseClick).Action = ppActionHyperlink Then
                    ' Follow the hyperlink
                    targetShape.ActionSettings(ppMouseClick).Hyperlink.Follow
                End If
            End If
{custom_show_start_case}{custom_show_return_case}
    End Select
End Sub

Function GetShapeByName(shapeName As String) As Shape
    On Error Resume Next
    Set GetShapeByName = ActivePresentation.SlideShowWindow.View.Slide.Shapes(shapeName)
    On Error GoTo 0
End Function

Function GetSlideIndexByID(slideID As Long) As Long
    Dim slide As slide
    For Each slide In ActivePresentation.Slides
        If slide.slideID = slideID Then
            GetSlideIndexByID = slide.SlideIndex
            Exit Function
        End If
    Next slide
    MsgBox "Slide ID " & slideID & " not found.", vbExclamation
End Function

Sub StartSlideshow()
    With ActivePresentation.SlideShowSettings
            .ShowWithNarration = True
            .ShowWithAnimation = True
            .LoopUntilStopped = False
            .AdvanceMode = ppSlideShowUseSlideTimings
            .RangeType = ppShowAll
            .Run
        End With
End Sub

    """
    # Add the generated code to the new module
    new_module.CodeModule.AddFromString(vba_code)

    # Set up the slideshow to call OnSlideShowPageChange on each slide change
    presentation.SlideShowSettings.Run()

    # Optionally run the macro immediately to initialize
    presentation.Application.Run("OnSlideShowPageChange")

    presentation.SlideShowWindow.View.Exit()

def elsagda(season, Bishop = False, guestBishop = 0):
    prs = relative_path(r"صلاة السجدة.pptx")  # Using the relative path
    excel = relative_path(r"Files Data.xlsx")
    des_sheet ="صلاة السجدة"
    replacefile(prs, relative_path(r"Data\CopyData\صلاة السجدة.pptx"))
    close_presentation_safe(prs)
    elzoksologyat(excel, season, "عشية")

    show_full_sections = []
    hide_full_sections = []

    if season == 23:
        # show_full_sections = ["السجدة الثانية – مرد انجيل العنصرة", 
        #                       "السجدة الثالثة – ارباع عيد دخول المسيح أرض مصر", 
        #                       "السجدة الثالثة – ختام ارباع الناقوس الفرايحي", 
        #                       "السجدة الثالثة – مرد مزمور دخول المسيح أرض مصر", 
        #                       "السجدة الثالثة – مرد انجيل دخول المسيح أرض مصر", 
        #                       "أوشية الإنجيل و إنجيل عشية دخول المسيح أرض مصر", 
        #                       "مزمور عشية دخول المسيح أرض مصر", "إنجيل عشية دخول المسيح أرض مصر", 
        #                       "عشية – مرد انجيل دخول المسيح أرض مصر", "جملة قانون ختام السجدة"]
        # hide_full_sections = ["تكملة مرد إنجيل السجدة الثانية", "السجدة الثالثة – مرد المزمور", 
        #                       "تكملة مرد إنجيل السجدة الثالثة"]
        
        show_full_sections = ['{A1E77CFF-8702-4752-BD83-D292974D3BDB}', '{F9A40801-A49C-4F1C-92E1-6FD857FCA84B}', '{1B1F50DF-E496-411F-B65C-0B590DB90DAD}', '{6CFFE3AA-240A-4A85-BE89-A1311636F7BA}', '{A88539B4-B27B-4D4A-9310-9C678C87AA38}', '{1DEABC00-739E-4CD6-8945-D5A256BD56F9}', '{CAE9212C-02E8-43B1-9A8A-55E0642E3057}', '{BB3B2CB3-DDE9-49AD-B21B-BFFB4ACEF1A7}', '{B9AD6CC9-3FA7-4ADA-9FED-16BB0235D2B3}', '{3EAB79FE-6F3D-4C18-B067-E25B7688B3DB}']
        hide_full_sections = ['{56B7CFF1-06D7-41D4-B626-E14ACBA15D46}', '{F13A48F2-238D-4617-B84E-9B0A694D9A18}', '{6A1A6375-A90C-4374-A779-668D1E27AB1E}']

    if Bishop:
        prs3 = relative_path(r"Data\حضور الأسقف.pptx")
        sheet = "في حضور الأسقف"

        # bishop_show_values = ["السجدة الأولى – تكملة في حضور الاسقف", "السجدة الأولى – ارباع البطريرك و الأسقف", 
        #                       "السجدة الأولى – مارو اتشاسف", "السجدة الأولى – فليرفعوه", "السجدة الثانية – تكملة في حضور الاسقف", 
        #                       "السجدة الثانية – ارباع البطريرك و الأسقف", "السجدة الثانية – مارو اتشاسف", 
        #                       "السجدة الثانية – فليرفعوه", "السجدة الثالثة – تكملة في حضور الاسقف", 
        #                       "السجدة الثالثة – ارباع البطريرك و الأسقف", "السجدة الثالثة – مارو اتشاسف", 
        #                       "السجدة الثالثة – فليرفعوه"]

        bishop_show_values = ['{F76B0D75-0474-45B5-B79F-7416F354543A}', '{10379C67-E3C7-49F6-9632-EFF77DF18C31}',
                              '{62A12AF8-CB6D-4CC5-9DB0-B73A7C24E2AD}', '{600D2394-BB87-4926-A5A1-24D17F10DD49}',
                              '{BC761F60-4905-493E-A792-F4800B18F258}', '{3CF7713D-D5DE-49BD-A9B5-04084515498E}', 
                              '{BCB45382-B78F-4FD5-A25C-E776CF2C2A42}', '{F9232667-1627-4250-B27A-8DE0725D4FA7}', 
                              '{D524086D-81EA-4CC2-B3A8-DE53B7B3A341}', '{892C7D68-138D-4033-87F6-68CD9724AFA0}', 
                              '{5417E851-A431-4753-A0AB-C575F199E5BC}', '{EDE3E188-A5CC-448F-BA89-E8EE496C531A}']
                
        show_full_sections.extend(bishop_show_values)

        if guestBishop > 0:
            # bishop_values = ["صلاة الشكر", "صلاة الشكر", "طوبه هينا الكبيرة", "طوبه هينا الكبيرة", 
            #                  "نيم بينيوت", "نيم بينيوت", "الاسبسمس", "الاسبسمس"]

            # bishopDes_values = ["السجدة الأولى – تكملة في حضور الاسقف", "السجدة الأولى – طوبه هينا الكبيرة", "السجدة الأولى – مارو اتشاسف",
            #                     "السجدة الثانية – تكملة في حضور الاسقف", "السجدة الثانية – طوبه هينا الكبيرة", "السجدة الثانية – مارو اتشاسف",
            #                     "السجدة الثالثة – تكملة في حضور الاسقف", "السجدة الثالثة – طوبه هينا الكبيرة", "السجدة الثالثة – مارو اتشاسف"]
            
            bishop_values = find_slide_nums_arrays_v2(excel, sheet, 
                            ['{6851F163-CBEF-4014-A853-CE100557BA6A}', '{6851F163-CBEF-4014-A853-CE100557BA6A}',
                             '{B084BC40-61E1-4477-98DA-15CFB06AEE91}', '{B084BC40-61E1-4477-98DA-15CFB06AEE91}',
                             '{97203297-EECB-4D41-B2E3-AD9A4863847E}', '{97203297-EECB-4D41-B2E3-AD9A4863847E}', 
                             '{D1378DB5-29D1-4800-9D96-10F2535EEB57}', '{D1378DB5-29D1-4800-9D96-10F2535EEB57}'], 
                            2, [1, 2, 1, 2, 1, 2, 1, 2])

            bishopDes_values = find_slide_nums_arrays_v2(excel, des_sheet, 
                               ['{F76B0D75-0474-45B5-B79F-7416F354543A}', '{8DD21CDE-CB6B-4D5B-B995-D2747AB69ED1}',
                                '{62A12AF8-CB6D-4CC5-9DB0-B73A7C24E2AD}', '{BC761F60-4905-493E-A792-F4800B18F258}', 
                                '{7009FA91-C485-4613-9FFA-D951043024D2}', '{BCB45382-B78F-4FD5-A25C-E776CF2C2A42}', 
                                '{D524086D-81EA-4CC2-B3A8-DE53B7B3A341}', '{24B42351-8629-4F79-902D-898CF7A72F78}', 
                                '{5417E851-A431-4753-A0AB-C575F199E5BC}'],                  
                               2, [2, 2, 2, 2, 2, 2, 2, 2, 2])

            elshokr_sagda1 = bishopDes_values[0]
            elshokr_sagda2 = bishopDes_values[3]
            elshokr_sagda3 = bishopDes_values[6]
            elshokr1 = bishop_values[0]
            elshokr2 = bishop_values[1]

            tobhyna_sagda1 = bishopDes_values[1]
            tobhyna_sagda2 = bishopDes_values[4]
            tobhyna_sagda3 = bishopDes_values[7]
            tobhyna1 = bishop_values[2]
            tobhyna2 = bishop_values[3]

            maro_sagda1 = bishopDes_values[2]
            maro_sagda2 = bishopDes_values[5]
            maro_sagda3 = bishopDes_values[8]
            maro1 = bishop_values[4]
            maro2 = bishop_values[5]

            if guestBishop < 2:
                elshokr2 = int(elshokr2) - 1
                tobhyna2 = int(tobhyna2) - 2
                maro2 = int(maro2) - 1
        
            start_positions = [maro_sagda3, tobhyna_sagda3, elshokr_sagda3, maro_sagda2, tobhyna_sagda2, elshokr_sagda2, maro_sagda1, tobhyna_sagda1, elshokr_sagda1]
            start_slides = [maro1, tobhyna1, elshokr1, maro1, tobhyna1, elshokr1, maro1, tobhyna1, elshokr1]
            end_slides = [maro2, tobhyna2, elshokr2, maro2, tobhyna2, elshokr2, maro2, tobhyna2, elshokr2]

    if season == 23 or Bishop == True:
        show_hide_insertImage_replaceText(prs, excel, des_sheet, show_full_sections, hide_full_sections)

    presentation1 = open_presentation_relative_path(prs)
    if isinstance(presentation1, tuple):
        raise RuntimeError(presentation1[0])
    
    custom_show_section_id = '{1DEABC00-739E-4CD6-8945-D5A256BD56F9}'
    vba_code(
        excel,
        des_sheet,
        prs,
        presentation1,
        slide_section_id='{A5B9CE2F-90E3-44D7-B22F-CAE6783C8E2F}',
        custom_show_trigger_section_id=custom_show_section_id,
        custom_show_name="Do5ol el masy7 ard masr",
        custom_show_return_section_id=custom_show_section_id,
    )

    if guestBishop > 0:
        presentation3 = open_presentation_relative_path(prs3)
        if isinstance(presentation3, tuple):
            raise RuntimeError(presentation3[0])
        # Initialize variables for current position, slide, and end index
        current_position = int(start_positions[0])
        current_start_slide = int(start_slides[0])
        current_end_slide = int(end_slides[0])

        # Initialize index for start position, slide, and end slide
        position_index = 1
        slide_index = 1
        end_index = 1

        while current_start_slide <= current_end_slide and slide_index <= presentation1.Slides.Count:
            try:
                presentation3.Windows(1).Activate()
            except Exception:
                pass
            source_slide = presentation3.Slides(current_end_slide)
            source_slide.Copy()
            try:
                presentation1.Windows(1).Activate()
            except Exception:
                pass
            pasted_slide = presentation1.Slides.Paste(current_position)
            if pasted_slide is None:
                raise RuntimeError("Failed to paste copied slide from the source presentation.")
            current_end_slide -= 1
            if(current_start_slide > current_end_slide):
                current_position += 1

            # Move to the next round if all slides in the current range have been processed
            if current_start_slide > current_end_slide:
                # Check if there are more rounds
                if position_index < len(start_positions):
                    # Update variables for the next round
                    current_position = int(start_positions[position_index])
                    current_start_slide = int(start_slides[slide_index])
                    current_end_slide = int(end_slides[end_index])
                    position_index += 1
                    slide_index += 1
                    end_index += 1

    if guestBishop > 0:
        close_presentation_safe(prs3)

    presentation1.SlideShowSettings.Run()

