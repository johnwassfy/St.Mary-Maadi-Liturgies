from commonFunctions import replacefile, relative_path, open_presentation_relative_path, \
                            show_hide_insertImage_replaceText


def lakanEl8etas(adam=False):
    prs = relative_path(r"صلاة اللقان.pptx")  # Using the relative path
    excel = relative_path(r"Files Data.xlsx")
    des_sheet ="اللقان"
    ghetasText = ["لأنك اعتمدت", "ak[i`wmc", "آك تشى أومس"]
    img = relative_path(r"Data\Designs\الغطاس.png")
    replacefile(prs, relative_path(r"Data\CopyData\صلاة اللقان.pptx"))

    # el8etas_show_sections = ['ارباع عيد الغطاس', 'ختام ارباع الناقوس الفرايحي', 'طاي شوري',
    #                          'عيد الغطاس - البولس', 'اجيوس الغطاس', 'اوران انشوشو', 'محير عيد الغطاس',
    #                          'مرد مزمور الغطاس', 'مزمور وانجيل عيد الغطاس', 'مرد إنجيل عيد الغطاس',
    #                          'طلبة لقان عيد الغطاس', 'مستحق وعادل - الغطاس', 'قدوس - الغطاس',
    #                          'رشومات عيد الغطاس', 'مزمور التوزيع']
    # el8etas_show_sections_ranges = [['عيد الغطاس - حبقوق (يارب سمعت صوتك)', 'عيد الغطاس - حزقيال (ثم حملنى الروح)']]

    el8etas_show_sections = ['{0778A679-4B1A-4079-AF21-5F40399D10CD}', '{B87EBA1A-E0E4-4E68-87D7-3C4A798CF278}', '{A9828217-C064-4E06-AB47-AB591B489586}', '{5AD97F36-2D1C-4076-B35B-5BBC6F54EE5F}', '{BD768037-D113-4C48-91BA-50DC7AE43AA3}', '{28701A42-B4DE-4ADB-A7EF-253A431CA3DA}', '{F446A3A7-0505-4BF3-9D3C-1F3D28D5FD9D}', '{E18DAE62-1E15-4D10-B25C-CEB09552527E}', '{0F3315A8-64B0-42A6-BBDC-67FEC445F7C8}', '{1DA3DB23-DD07-4770-9938-2955715019D1}', '{AC655B1C-A001-4050-9452-5D1EE1C97E92}', '{337C300A-9FF7-4BD1-8550-D2664CBAA5D5}', '{AFD01476-4CF1-4EEB-8AA8-A3B6CEA09670}', '{1ABCE2B1-0418-440F-B0FE-6B8C80AE76D1}', '{C29E5A83-A98B-4077-8194-99A6D803EF53}']
    el8etas_show_sections_ranges = [['{C757D167-172B-4025-ACDB-0ABC6E6104FD}', '{BBCFA3F3-05AA-48B2-99B7-B2D4B1A5F6EA}']]

    if adam:
        el8etas_show_sections.append('{D02C088A-01E0-4A8C-8D73-21E3FD3616EB}')
    else:
        el8etas_show_sections.append('{9495E38B-CE03-4E75-AED4-960DE95BA747}')

    show_hide_insertImage_replaceText(
        prs, excel, des_sheet,
        show_sections=el8etas_show_sections,
        show_sections_ranges=el8etas_show_sections_ranges, 
        image_path=img, new_Text=ghetasText,)
    
    presentation = open_presentation_relative_path(prs)
    presentation.SlideShowSettings.Run()
