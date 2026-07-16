from PyQt5.QtWidgets import (QDialog, QPushButton, QVBoxLayout, QLabel, QFrame, QHBoxLayout,
                           QScrollArea, QWidget, QGraphicsDropShadowEffect)
from PyQt5.QtGui import QFont, QPixmap, QColor
from PyQt5.QtCore import Qt, QSize
from commonFunctions import relative_path, open_presentation_relative_path, show_hide_insertImage_replaceText,\
                            replacefile, find_slide_nums_arrays_v2, get_slide_ids_by_numbers, show_slides, \
                            elzoksologyat, run_vba_with_slide_id_bakr_aashya, get_open_presentations
import qtawesome as qta
import os
import sys
from qudasDialog import Elbas5aSectionSelectionDialog

class Elbas5aDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.selected_option = None
        self.active_button_id = None
        self.active_presentation_path = None
        self.button_widgets_by_id = {}
        # Temporary bypass list: clear/remove these IDs later to restore full pre-actions.
        self.temporary_direct_open_ids = {
            "sun_gen_funeral",
            "sun_day",
            "wed_night_thursday",
            "thurs_night_friday",
            "fri_great_friday",
            "sat_night_abughalamsis",
        }

        self.default_button_style = """
            QPushButton {
                background-color: rgba(255, 255, 255, 200);
                border: none;
                border-radius: 12px;
                color: #3a0000;
                padding: 10px;
                font-size: 13px;
                font-weight: bold;
                text-align: center;
                min-height: 20px;
            }
            QPushButton:hover {
                background-color: rgba(255, 240, 240, 230);
                color: #690000;
                border: 1px solid rgba(255, 255, 255, 50);
            }
            QPushButton:pressed {
                background-color: rgba(200, 180, 180, 250);
                padding-top: 11px;
                padding-bottom: 9px;
            }
        """

        self.active_button_style = """
            QPushButton {
                background-color: rgba(235, 255, 235, 230);
                border: 2px solid rgba(120, 255, 120, 230);
                border-radius: 12px;
                color: #1c4f1c;
                padding: 10px;
                font-size: 13px;
                font-weight: bold;
                text-align: center;
                min-height: 20px;
            }
            QPushButton:hover {
                background-color: rgba(245, 255, 245, 240);
                border: 2px solid rgba(160, 255, 160, 240);
                color: #1d6b1d;
            }
            QPushButton:pressed {
                background-color: rgba(210, 245, 210, 250);
                padding-top: 11px;
                padding-bottom: 9px;
            }
        """
        
        # Button configuration: each button has file path and section visibility settings
        # Example: to show specific sections, add their section IDs (GUIDs) to show_sections
        # Example: to hide specific sections, add their section IDs (GUIDs) to hide_sections
        # Section IDs format: '{XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}'
        self.button_configs = {
            "sun_gen_funeral": {
                "old_path": r"Data\اسبوع الالام\التجنيز العام.pptx",
                "new_path": r"Data\اسبوع الالام\التجنيز العام.pptx",
                "show_sections": [],  
                "hide_sections": [],  
            },
            "sun_day": {
                "old_path": r"Data\اسبوع الالام\old\تجنيز احد الشعانين 2022.pptx",
                "new_path": r"Data\اسبوع الالام\old\تجنيز احد الشعانين 2022.pptx",
                "show_sections": [],
                "hide_sections": [],
            },
            "sun_night_monday": {
                "old_path": r"البصخة المقدسة.pptx",
                "new_path": r"Data\اسبوع الالام\البصخة المقدسة.pptx",
                "show_sections": ['{ED5E18D1-B6DA-4121-A7C2-4578F30F3EFB}', '{E13CC262-636E-4848-880F-FB621B72C394}', '{92640649-ECA2-447F-8CBD-A8361DC51F20}', '{929FC923-1774-4BB4-9606-4D2E102B6C35}', '{24C9888C-E3FB-4A2D-9CEA-A571F91E68B3}', '{E9C456C2-EBC1-4BA9-A9A0-BCBCD6818831}', '{B6D1A3B6-C89F-4611-BAC2-706FB08B8687}', '{51976FA8-BFDA-4AEC-B023-1F5B64590D9D}', '{0CEA8218-0F36-40B3-BC35-0F9EC87BE574}', '{E96A8D49-CA3C-4751-BEC3-B560E8DA658C}', '{1B566837-5E56-4387-8B57-A10AFFD2E32F}', '{7957D3F4-661D-47AA-8288-889752BDD440}', '{4A8EBB88-F53C-40BA-AC57-DBCBF6BA2AF2}', '{F57B1F73-3F49-405B-9579-63139205F24A}', '{C0E4FA3A-BCBB-4229-96B0-1B8E8CD4F9CE}', '{35012B98-6D63-46A1-B0D1-69BB7AA8B55A}', '{3793826C-C977-42F3-8D31-83E6E5C7FC38}', '{42F77A2D-9879-4D44-9CD1-9D5D248179CF}', '{37682D35-40CD-40BD-8B42-931AAAF65CC7}', '{0A09D98C-9DCB-4D31-84FA-E0516179D123}', '{34C27C36-2C1C-4C8D-A70A-CF229E95052F}', '{8449FB3E-D80E-4104-B36F-631EA62D6BC0}', '{7A920A38-FCA7-4364-92C0-5E7F4FC38B70}', '{5A56BA0D-C038-4E93-B9C4-1AC3F0FDF5EE}', '{BCA45ABD-1595-4D3A-85ED-9F409F303B35}', '{CFBBB1CE-F1E2-426F-B66D-97C7A1FD9E83}', '{1B618E58-3DCA-456E-BEB3-99F1E23CC6F0}', '{8ECE385D-379E-4752-8AB0-4B67D2446ADC}', '{7998F7AA-FE2C-4A86-995E-0296996C7F95}', '{604937D2-930D-4320-BB10-BAFA7AAA9406}', '{A59FCA74-BB6C-4E24-B8AD-C4D102EB1A48}'],
                "hide_sections": [],
            },
            "mon_day": {
                "old_path": r"البصخة المقدسة.pptx",
                "new_path": r"Data\اسبوع الالام\البصخة المقدسة.pptx",
                "show_sections": ['{DA7CB644-1BDF-46F0-8DAA-3016B5FBE36C}', '{59A6EBD8-666E-4C62-A8DC-137273C8CB10}', '{EFBFCE73-886D-4714-9E34-95323B53CC23}', '{36995511-A267-4F3C-ABB3-5383C750C2D4}', '{526DAD64-2BAC-4685-8981-238687961073}', '{8F02DEB3-A361-4206-BFAA-97D09AE34C86}', '{24B8E840-1BF1-4296-A979-564B4497A965}', '{3F5BFCC8-0C80-45EC-B9BE-D53B4DBA28C6}', '{8D6EC106-C791-4460-90C7-16D6A20F6DB8}', '{65EE8925-18E1-4080-86D0-8343B3826265}', '{37D4A68B-543B-4F9B-A88C-FF6DC46BAEA2}', '{6A2A5C29-5AE3-424A-86CA-75D67163F568}', '{F781C96B-8530-444E-855A-02A770377B8B}', '{3AE20C75-23F7-4184-9601-FFE61E1D3505}', '{716B871C-4FAD-46CE-9C26-EE2696E1C53F}', '{DF41B66E-5CCB-4C3E-92CB-FE521DE4B24B}', '{D06C3F33-FA61-411D-8360-63A3699A0EB1}', '{190D3883-5A60-46FD-B477-7F663E3B2D2D}', '{62385410-BBBC-4716-B606-DC93752849DF}', '{8487EEFE-75D0-41B9-812A-DD7249CB8B9A}', '{F695BD8E-CB63-4E58-9B78-A58DDFFA0812}', '{0C41C679-26A8-4B60-A95C-756FE2DEC537}', '{19A8FFA0-0D78-4DCF-A6C4-438BDF1CFC58}', '{3C8CCBE3-A797-4A18-BD09-8DCB4CFE516C}', '{EE86F70E-CB3C-401B-82F3-CBEF13FA0A0F}', '{14DA35EF-5027-4234-A04C-F23037E94016}', '{05F8BE6A-EFDA-4BB8-ACB4-52EDF63883DB}', '{14C94D32-0D75-4A2F-89B1-07C899248154}', '{86CDBA6E-809A-46EE-986D-4D5CE5F9CE52}', '{B2E6181B-4107-4E5C-A394-188E5C7F7AD8}', '{9CD6A7D8-47B0-4281-99EF-4238E97E5F98}', '{166EEFCE-1F45-4929-B0B8-B0C6C4EA9682}', '{4A0D3116-A6F2-4C74-A746-10BF566B7BF3}', '{89F30043-8122-4174-88E7-666316C802F9}', '{CE9E2890-9BDB-4ACD-881B-FD7F2AE5CE57}', '{AD583E12-CCD0-4BA6-A397-3509CF171C43}', '{16D098E7-E54D-4B9B-836C-C091E185AEB3}', '{06D1C8EB-6C1E-433C-A5BF-4A4503CC6F55}', '{D6466A87-CF79-4DEC-896D-D8235FFF1B6B}', '{12AFF2A6-4B68-43A8-8B4D-B8D548DDEE17}'],
                "hide_sections": [],
            },
            "mon_night_tuesday": {
                "old_path": r"البصخة المقدسة.pptx",
                "new_path": r"Data\اسبوع الالام\البصخة المقدسة.pptx",
                "show_sections": ['{28DBDFDC-808E-4610-9197-375D112E26F1}', '{F245D4C7-ED26-4336-97DB-AFA862DDAEA1}', '{2DE4734E-4604-4301-A42E-8CF7E770397D}', '{C5967DEF-206B-4355-99B7-6B960F8BADD6}', '{0FFBA638-1F42-497A-8286-5E36FD23E10C}', '{58C5CC20-CAC6-4618-8C77-3D25D30F7B38}', '{CCAFB188-38B8-4E0F-B091-8215D2B954DA}', '{BF320E20-0811-4440-81F5-01314D85B05A}', '{1D6E0D8B-BCFA-4660-877B-A45BF2DE68BA}', '{28191AC7-5314-414D-A9B8-CAABD46B6B5E}', '{73EC17C8-9C1C-4C97-8B3B-46EB998A4D8C}', '{BE2991AD-42A1-4ACC-BF00-08361287672C}', '{41751444-B14D-431F-AC74-74B7F9A14762}', '{C0655F59-48DE-413B-A777-90E1718383F7}', '{DCDD0F8D-AB86-4727-A4E8-4FDDDD32E9C9}', '{61C1F412-14E0-4B29-8472-F8C9777CA2EC}', '{43326299-58F4-4A1A-AE74-5F23F3F2FCC4}', '{27E84055-F8F7-4B2E-8563-98EAA602C652}', '{C37225B2-8575-4E5C-933E-8E87A4D7047C}', '{C05186E5-EED1-4BB3-A621-1B011B31D9C0}', '{73A7AA32-EFFA-4499-9BCF-2DAB418F997F}', '{16542A30-D88A-4600-98C3-90F158902C73}', '{217376E9-B474-4D66-8383-632BEA775DCE}', '{0D7BBB39-9DCE-4946-A5E5-CA58B56B00F2}', '{FB82C05B-C956-4464-B207-8DF79D5C9B53}', '{0D20F4C2-6402-4C1A-97DA-6207A1718947}', '{2F686399-1C4D-4CB4-B8E2-73B1E1F30B30}', '{77AB488E-1DFA-4858-9F8A-542061EA0761}', '{08CDBE54-FE33-4A22-9F88-4A226240A992}', '{3D42E547-3326-42D5-81F8-0C58198D7D34}', '{575F9C1A-8BF7-4A00-89AC-3F41AE18C1BC}'],
                "hide_sections": [],
            },
            "tue_day": {
                "old_path": r"البصخة المقدسة.pptx",
                "new_path": r"Data\اسبوع الالام\البصخة المقدسة.pptx",
                "show_sections": ['{B6C9F540-BC97-4D10-B845-FC4884F0942D}', '{6C7FAC53-AF83-40E1-BDF5-2D608C8E25A6}', '{9FDA1E8B-4D24-4675-A5E4-C756A96FF89D}', '{3153BB54-1301-4C0A-B22B-DBEB07552AEA}', '{A7B7163D-7DF4-4519-8EA2-7B4AC4349AB5}', '{4BD3186B-87A4-4778-B0D4-7EE337FEE20C}', '{AE809017-195E-46E3-A5BC-F0567334421D}', '{9A51814C-1694-409C-87BC-DBEA3C7CE642}', '{DCF7AF35-BFBA-453B-BBAC-52A57268181F}', '{4A85C392-BB20-452C-9044-DAB09C157A2A}', '{9954FB65-6829-4EAE-B4DC-08413ED46C3A}', '{17A68DF3-91E2-4A5D-9648-DE10BB4851CA}', '{89691041-3B6F-462C-A49D-53F4DF506AF3}', '{EA1A9E10-95D1-4ED7-8E49-1894BE124B13}', '{B6347C74-B80D-4A3A-95BF-2489442B9528}', '{877B1B34-26F4-45B6-B167-812B064126BE}', '{4432CE5F-9280-483C-A5F1-293D343B3248}', '{BA5D09D6-6D53-48CE-B4C8-668066A78450}', '{5B519A8B-08AE-46FB-8D0A-CA0EA6D8B304}', '{394BA323-7F82-4BF4-8AA4-B8D79A5D4994}', '{96F99B60-894E-48CF-91A9-B325DD09C490}', '{71418500-5C4A-409F-B010-F70D8197B34E}', '{5CF462D1-CF51-49DA-934A-BE392723A3ED}', '{CEDD0C66-0C53-4538-A6F6-9C0D6532CB43}', '{1A488526-A349-42E1-94B7-CA6FA4542E9D}', '{8713F022-9DA0-42D8-AA16-AA2EDFB88E7C}', '{6E43B0B8-C2B8-4B61-959B-435602354263}', '{3629A550-6B0E-4E32-8264-1401E01EEA10}', '{630516C0-3CDE-4D5C-9331-64FF18F2165C}', '{E5B91D40-8A1A-40A2-83F5-8A67253A57D8}', '{A5C238FF-B105-420F-A89B-C4FB710D8411}', '{DCC118A2-B0D3-4203-B2D5-48FF2954EED2}', '{1C6B0C95-F113-4706-A2AB-AADB7678F51D}', '{943FDD57-D947-4DA1-BE17-D1F4DF584DF1}', '{6CD11D7D-0504-4821-8164-928FC9D124D4}', '{795B2B32-08BD-4A4C-A025-2F64614F9AC5}', '{7C3A8C27-F9E2-4896-A1AE-23772687DC1A}', '{0B7F83B7-89A6-48D4-8D2C-92BC8F3F55AC}', '{BA26D33B-F934-4974-AA1E-EB9BAE08220F}', '{7724DD9A-8B67-4009-B150-7B2AD0661105}', '{853CC655-9A9D-4D51-8294-EC42748EFA7D}', '{87C7E743-4CFC-40FB-97E2-30DF757A4D00}', '{8F211F50-CE5A-41E8-8F60-A85D4107EECA}', '{F1DAED4A-C77C-403C-8A28-88C2E21F411A}', '{5ED93895-60E3-443D-9035-F5716770C187}', '{8288C766-FCAD-4627-BB6A-DEFB165E299D}'],
                "hide_sections": [],
            },            
            "tue_night_wednesday": {
                "old_path": r"البصخة المقدسة.pptx",
                "new_path": r"Data\اسبوع الالام\البصخة المقدسة.pptx",
                "show_sections": ['{26A8BF77-C454-4F5B-9DC5-6ED0672EF4FA}', '{F8BF9F41-4FDB-4E3A-B833-DAFE7A6F0442}', '{1664A562-8286-4F6F-BC04-694172E8DD6C}', '{48447719-10BC-441D-BB0D-A84254D7C063}', '{7DDCC0C1-08EF-4A88-98CE-96A975B6D447}', '{B8320126-24BD-4AC2-9BB6-452F71A88BFB}', '{AECA780B-5D2D-4AFF-8DC7-5C61CCDF4754}', '{B2B5879A-0358-47C6-931E-9AA4365CF34B}', '{15E45FC9-7215-4D77-9346-A49F674217FA}', '{517CBBCC-6CAE-4C17-B327-348D66F8175D}', '{DCAA773D-10DD-46DF-B56D-A8B03D2B56CD}', '{3339DC4D-C8C6-4081-977C-DB520FBB1A74}', '{710DD071-67BB-4E13-98F0-7FB123679F70}', '{7C5C6710-E47C-4CB1-9657-F6D38EE84D33}', '{AE05E708-3690-4209-BBC2-C9FC9769F650}', '{D1D1FBB5-FBF5-43FF-8CF4-BD0877EAAE88}', '{3028CB02-0A77-4F88-AAD6-47523D05C1AF}', '{12BBDCD8-1389-48E4-BE9E-995B378BD523}', '{1DB5C2E8-D208-43F7-B961-A6B875248545}', '{84509267-C471-4405-8469-C36AB3387EA4}', '{257EC627-6C1C-42DA-9832-65FB4424C855}', '{3730D43C-5276-44AE-9611-7A23EDC416E6}', '{645D7728-5D24-4E66-A65A-3D4FF8A8087B}', '{FB727733-DF8D-4FCC-AAA7-65EBF287E89A}', '{EEF1540A-D491-4480-8B10-0D5245B2819E}', '{DE6CED1A-BB7E-4D49-A700-1FDA8931EA98}', '{C2B8B1C4-4608-44D6-B90A-E8C6E40B3FF3}', '{B28310A8-CC91-442A-B385-960DDF5A0B1B}', '{8810D55E-A3C5-46E8-A148-C012C5585E32}', '{FB0236F6-6C53-48AA-A8D4-D9DEF99CEF0C}', '{2B59880F-40C6-4346-921B-79CAFFD38747}', '{B9887CF6-D786-4C00-8058-005757654FA2}'],
                "hide_sections": [],
            },
            "wed_day": {
                "old_path": r"البصخة المقدسة.pptx",
                "new_path": r"Data\اسبوع الالام\البصخة المقدسة.pptx",
                "show_sections": ['{FBFF4FE3-4ED8-4177-AAD3-C23A705CC099}', '{5B21492B-0C05-478F-A2F0-63F93777D8DB}', '{53CA4ED5-C74E-4CA4-A039-8AD4C11C3B1C}', '{ECDC8448-7D1F-4649-99F2-41275FDE1757}', '{D6B541CE-F633-4F7B-B3A0-BE683B9F35D5}', '{1A722178-9784-4680-A0DD-22F15163F841}', '{7EF9715F-99AE-4B48-8344-947329BA72A7}', '{1D93C885-37E8-4D47-BDB5-1CC150E45FB3}', '{B574588F-B47E-4785-A444-576812EF31C0}', '{7029B861-BBAC-4C8D-9E3A-749707669987}', '{2D5F76FC-D61E-4D27-898C-C22DDC769D4C}', '{3DFE3618-283B-4E5A-A3DE-E0B3B6D8C087}', '{EBFB5DDD-E088-4219-BA6E-3C2173265311}', '{6A0DB725-0354-4E37-868C-7C248235A9AF}', '{2E210CAB-11FE-4C41-8E64-61F65D885A48}', '{1DF49FCF-21FD-489B-8999-F0B6EDB01B26}', '{472B2BFB-2E51-4EEB-AD59-597EAF98A4D4}', '{8FBB2F86-7911-48A3-9E40-77E2C20B1B33}', '{9C93C17F-10CF-4F50-8C4D-FCBE3F4D92BA}', '{CB1EE99D-4D50-4D00-A451-C6D0586951DC}', '{85E4C0C9-F7B1-4475-9C30-7BDE31E6669F}', '{B4BD5EA2-1F01-4BF5-ABDD-3BE66CDD8980}', '{62BD7EF1-15B4-4697-9C00-9102D9380C00}', '{D3577B82-36F5-4DB3-B0F6-FCEB58C63B5D}', '{41DCB073-92A6-4B79-B3D5-6E4BE0AB4653}', '{58CED2B3-A147-4DF2-B51B-D8B287D5ABAE}', '{37738175-7039-4244-BECC-B6187A66860E}', '{811E5EE2-895C-4B2A-87BA-1A65BE4AEF46}', '{567E9395-5794-411A-A742-E245A1159D57}', '{D6605F8B-AC0D-4081-8FC6-2BA4283E024B}', '{157854D7-8765-4E29-A01E-74188B06344E}', '{7FD363DC-07B4-49A6-8F91-F5AC4AE6D102}', '{CE460A68-4CAC-4F20-A569-3622040106EA}', '{687BAAF8-ED23-4665-81DD-CE28E1FF7F71}', '{1F8307C2-4639-4D60-9545-B062964B3651}', '{AD1B0753-E7C2-4B6C-9341-2A6376AE299F}', '{55FDFA5B-5EC9-4EF6-A45B-2AED3453E911}', '{411619F4-635E-4166-865F-9F490A54188F}', '{5F84AB8B-FA96-42EB-8064-C8CC18CD77FD}', '{308DBD89-1E97-4C15-96D5-152CD0932959}', '{EFD2D44D-556E-4E79-B4B4-86A2A0A8BAA5}', '{A433467F-4BFC-47C1-B090-0EA68F985A8E}', '{BE765AD3-8EBB-452A-82B1-5DAA85252198}', '{E555183A-F078-4F33-8315-28667FF3C0D8}', '{DA7414E6-AB66-472D-9899-0B3DA732C8A4}', '{1200EF37-CD97-45B4-A517-E59D52C5B933}'],
                "hide_sections": [],
            },
            "wed_night_thursday": {
                "old_path": r"البصخة المقدسة.pptx",
                "new_path": r"Data\اسبوع الالام\البصخة المقدسة.pptx",
                "show_sections": ['{1C6E183A-F9A8-48CC-89C9-DF8FB73F11D1}', '{1CFD382D-C872-4212-90A2-E6127AC3DDA7}', '{27BEB4B2-BD10-454B-99E4-4E5D8062367F}', '{8D107DC3-A27E-4C80-8639-C73909BDB20D}', '{7D3B5E0D-D794-4EDC-885C-8A3C35765149}', '{A01158F4-6F91-4CFE-ABE5-393E2F7BD4AC}', '{192F7BF1-E972-4CCF-99A7-D497F3E9E975}', '{CA45D3EC-1707-477D-B051-75B2BE084453}', '{1DC99677-5925-46AC-903F-22E332F51134}', '{00E51318-76EB-4778-8ADF-687DF6A4B816}', '{CBA3B815-F2C2-419D-8B4A-84FC925717D7}', '{76089659-14C4-428A-9EA4-829ED42A898E}', '{31A58379-22A5-4845-A062-949338F3E8CC}', '{24DE0EE5-B6D5-416D-A8D4-B85C946D9990}', '{932F99E6-FFB8-4175-A98E-4B16B7195050}', '{6ECC0BEB-6E2C-4E75-8006-0490BD84E2B1}', '{9F790281-A1E0-4FD8-BD7F-CBF065A0AE01}', '{01A020F0-C101-4BBB-BB92-96FC504B11D7}', '{47E5DDC2-703E-4993-9D5C-B974421D0EB1}', '{2D1C166B-CEDB-46C1-8E66-4009BE5A6608}', '{1DB84FD1-2D93-4AD1-BD7B-CC0409F7109B}', '{3AA87729-C9AD-4EED-8F86-DE4E1789C539}', '{76085593-F35B-4C62-848F-284805B25D9D}', '{33F01EBC-402E-44E4-B084-21CE14FE0DDE}', '{BD82A6AF-101D-4571-A1D6-18A3F4A6DDB7}', '{CFCE01D0-5966-4D8A-8945-E61C88B64939}', '{A7DBCD8C-C91C-4878-8612-75536AB618A5}', '{576DFE7B-3F70-4E70-8DB6-98771DF69C7D}', '{86E47CB1-AC54-4D2F-834F-6727D26683D6}', '{C841676D-0060-4A2C-9BAF-80DDCB48829A}', '{FCF62C19-78FA-43DF-80EA-D94B9F03A806}', '{1F26CA13-A9DE-4E5E-A7E7-229DFAF72731}', '{656F197C-F21A-46B2-8AD6-D8B443CCF620}', '{80338EC9-122B-4DEA-831D-954617B55428}', '{21DBE32E-4B47-48B9-A63E-6100ED0E34DD}', '{E4A8B103-5FC3-4219-B017-D9A5AAA965AE}'],
                "hide_sections": [],
            },
            "thurs_maundy_thursday": {
                "old_path": r"خميس العهد.pptx",
                "new_path": r"Data\اسبوع الالام\خميس العهد.pptx",
                "show_sections": [],
                "hide_sections": [],
            },
            "thurs_night_friday": {
                "old_path": r"Data\اسبوع الالام\old\ليلة الجمعة.pptx",
                "new_path": r"Data\اسبوع الالام\old\ليلة الجمعة.pptx",
                "show_sections": [],
                "hide_sections": [],
            },
            "fri_great_friday": {
                "old_path": r"Data\اسبوع الالام\old\الجمعة العظيمة كاملة 2022.pptx",
                "new_path": r"Data\اسبوع الالام\old\الجمعة العظيمة كاملة 2022.pptx",
                "show_sections": [],
                "hide_sections": [],
            },
            "sat_night_abughalamsis": {
                "old_path": r"Data\اسبوع الالام\old\سبت النور.pptx",
                "new_path": r"Data\اسبوع الالام\old\سبت النور.pptx",
                "show_sections": [],
                "hide_sections": [],
            },
        }

        self.day_order = ["الأحد", "الإثنين", "الثلاثاء", "الأربعاء", "الخميس", "الجمعة", "السبت"]
        self.button_meta = {
            "sun_gen_funeral": {"day": "الأحد", "text": "الجناز العام"},
            "sun_day": {"day": "الأحد", "text": "يوم الأحد"},
            "sun_night_monday": {"day": "الأحد", "text": "ليلة الإثنين"},
            "mon_day": {"day": "الإثنين", "text": "يوم الإثنين"},
            "mon_night_tuesday": {"day": "الإثنين", "text": "ليلة الثلاثاء"},
            "tue_day": {"day": "الثلاثاء", "text": "يوم الثلاثاء"},
            "tue_night_wednesday": {"day": "الثلاثاء", "text": "ليلة الاربعاء"},
            "wed_day": {"day": "الأربعاء", "text": "يوم الاربعاء"},
            "wed_night_thursday": {"day": "الأربعاء", "text": "ليلة الخميس"},
            "thurs_maundy_thursday": {"day": "الخميس", "text": "خميس العهد"},
            "thurs_night_friday": {"day": "الخميس", "text": "ليلة الجمعة العطيمة"},
            "fri_great_friday": {"day": "الجمعة", "text": "الجمعة العطيمة"},
            "sat_night_abughalamsis": {"day": "السبت", "text": "ليلة أبو غلامسيس"},
        }

        # Hardcode Holy Week filter keys/keywords and constant section IDs here.
        self.holy_week_filter_map = {
            "sun_night_monday": {"mode": "by-key", "key": "(ليلة الإثنين)", "keywords": []},
            "mon_day": {"mode": "by-key", "key": "(يوم الإثنين)", "keywords": []},
            "mon_night_tuesday": {"mode": "by-key", "key": "(ليلة الثلاثاء)", "keywords": []},
            "tue_day": {"mode": "by-key", "key": "(يوم الثلاثاء)", "keywords": []},
            "tue_night_wednesday": {"mode": "by-key", "key": "(ليلة الأربعاء)", "keywords": []},
            "wed_day": {"mode": "by-key", "key": "(يوم الأربعاء)", "keywords": []},
            "thurs_maundy_thursday": {"mode": "by-keywords", "key": "", "keywords": ["باكر", "الثالثة", "السادسة", "التاسعة", "اللقان", "القداس", "الحادية عشر"]},
        }

        self.holy_week_always_include_section_ids = [
            # Add constant section IDs or section-name identifiers here.
            "{81AD693D-C29D-45B2-95D4-0C613DDDE79B}", "{96280184-2A2C-4818-90C5-678E075E9ADC}",
            "{A59756E2-838D-4C9B-B75F-0B1EBBBAEA96}", "{65A10C31-4943-4801-A0CA-50005EE82390}", 
            "{A28B469C-AD93-40D5-9639-F0D603F4A94A}", "{9BEC2A31-DA53-45BE-AF39-7BE678B18A94}",
            "{98F043A4-8610-44D9-8777-2F65D839B922}", "{03C72C0D-7CF2-4A33-BD2D-8CE569C2AA44}",
            "{C77E4783-C285-4706-B437-1BBF92A3E954}", "{BB356010-9020-47D0-8893-27CBC69CEE3C}", 
            "{47F9C68B-29CB-4CD0-8B7C-2ACFAE49A3F6}", "{59D8280C-0080-4BF9-A6E7-6448BCD95397}",
            "{FC78AF58-DD25-4F6A-9EF6-1983BA4636DB}", "{5DF0EB10-99E1-41E1-AAD9-469FAC87A801}", 
            "{9F608A75-CC79-4278-B81C-5762C746A235}", "{8779F730-B75C-4ADA-8196-E41B1DD37E8E}", 
            "{D0A19F68-9B40-46E9-8CEE-20069205C2F5}", "{FD19477E-7950-4E0D-BC82-6B8D20ABEB34}",
            "{C3A42F67-4FD7-4B5B-B950-FE518D2C219D}", "{F6CB58DD-AF61-4770-8EEB-E446C5E9C4DC}", 
            "{F87B3D77-700A-4F7B-81BF-8571F2E3164B}", "{931A7E35-642D-498D-A76F-2896828706A5}", 
            "{8F08FD7F-9103-45F8-9A4D-90499DC6C9A5}", "{4C6C9B5B-7140-413D-A572-9661512CE527}", 
            "{9D06AEAB-48F5-44D5-B54C-6166991785C3}", "{2DC0F4CA-EBCD-43BE-97CB-B89A02733713}", 
            "{C5E13859-CEAB-44EC-9F64-8E8BFDFEB6C6}", "{3078F770-85E4-4ECF-8284-F172CD38CCC8}", 
            "{0CAB87D1-863E-47CE-A15E-1CC6E87F4A09}", "{378535F0-B031-4254-99A2-726504C335DB}",
            "{68FCD4BA-4331-4DEA-8948-3AAE780BD9A8}", "{3F426EFA-C75F-4A41-90D7-D54BAE86FF0B}", 
            "{B49BB649-3054-4719-BC60-F1B97BA48DF9}", "{C2E1CCA6-25C9-4D0A-B836-CCD883297242}",
            "{26889C30-3756-4820-AAFB-40519C5C4E66}", "{4E17041D-3E27-4200-94D3-4CBF0A332257}",
            "{0B9344A5-1DF4-48D6-A2F3-3BADF07E2334}", "{37F96881-032F-4EEB-8D10-5B5C2D5966F2}",
            "{4F15C931-A18C-47AE-BD80-9EFC3E020485}", "{4684F45B-A266-433A-A91A-735CEF280595}",
            "{390F461A-F6BF-498E-A7F7-AF7650D6DA16}", "{0D51C9D7-BA81-4E8F-BDED-7F621E88C38D}",
            "{5182946E-AF6A-4F9C-B25D-2474B6EB9107}", "{934F90A5-11E9-4766-B233-DE69026BC56A}",
            "{A9183893-7B7E-459F-8547-F7A8F7D2D521}", "{A3DED752-5159-4F64-86F6-7B95F37A8327}"
        ]

        self.show_slides = []
        
        for button_id, config in self.button_configs.items():
            meta = self.button_meta.get(button_id, {"day": "", "text": button_id})
            default_vba_sheet = "خميس العهد" if button_id == "thurs_maundy_thursday" else "البصخة"

            config["id"] = button_id
            config["text"] = meta["text"]
            config["day"] = meta["day"]
            config["show_sections"] = config.get("show_sections", [])
            config["hide_sections"] = config.get("hide_sections", [])
            config["vba_sheet"] = config.get("vba_sheet", default_vba_sheet)
            config["direct_open_only"] = config.get("direct_open_only", button_id in self.temporary_direct_open_ids)
        
        # Initialize Excel file path
        self.excel = relative_path(r"Files Data.xlsx")
        
        self.setWindowTitle("أسبوع الآلام")
        self.setFixedSize(550, 480)
        
        # Make dialog modal - will be attached to parent window
        self.setWindowFlags(Qt.Dialog | Qt.FramelessWindowHint | Qt.WindowSystemMenuHint | Qt.WindowTitleHint)
        self.setModal(True)
        
        # Main layout
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)
        
        # Set gradient background for the entire dialog
        self.setStyleSheet("""
            QDialog {
                background: qlineargradient(
                    x1: 0, y1: 0, x2: 1, y2: 1,
                    stop: 0 rgba(107, 6, 6, 245),
                    stop: 0.6 rgba(140, 30, 30, 245),
                    stop: 1 rgba(180, 80, 80, 245)
                );
                border-radius: 10px;
                border: 1px solid rgba(200, 200, 200, 150);
            }
        """)
        
        # Header
        header = self.create_header()
        main_layout.addWidget(header)
        
        # Main content
        content_container = QFrame()
        content_container.setStyleSheet("background: transparent; border: none;")
        content_layout = QHBoxLayout(content_container)
        content_layout.setContentsMargins(15, 10, 15, 10)
                
        # Buttons panel on the left with scroll area
        buttons_panel = self.create_buttons_panel()
        content_layout.addWidget(buttons_panel)

        # Add a little spacing between photo and buttons
        content_layout.addSpacing(15)
        
        # Photo panel on the right
        photo_panel = self.create_photo_panel()
        content_layout.addWidget(photo_panel)

        main_layout.addWidget(content_container, 1)  # 1 = stretch factor
        
        # Set Arabic RTL layout
        self.setLayoutDirection(Qt.RightToLeft)

    def create_header(self):
        header = QFrame()
        header.setFixedHeight(50)
        header.setStyleSheet("""
            QFrame {
                background: qlineargradient(
                    x1: 0, y1: 0, x2: 1, y2: 0,
                    stop: 0 #8b0000,
                    stop: 1 #c03232
                );
                border-top-left-radius: 10px;
                border-top-right-radius: 10px;
            }
        """)
        
        header_layout = QHBoxLayout(header)
        header_layout.setContentsMargins(15, 0, 15, 0)
        
        # Title with icon
        title_layout = QHBoxLayout()
        
        # Add icon (optional)
        try:
            icon_label = QLabel()
            icon = qta.icon("fa5s.cross", color="white").pixmap(24, 24)
            icon_label.setPixmap(icon)
            icon_label.setStyleSheet("background: transparent;")
            title_layout.addWidget(icon_label)
            title_layout.addSpacing(10)
        except:
            pass
        
        # Title
        title_label = QLabel("أسبوع الآلام")
        title_font = QFont()
        title_font.setPointSize(16)
        title_font.setBold(True)
        title_label.setFont(title_font)
        title_label.setStyleSheet("color: white; background: transparent;")
        title_layout.addWidget(title_label)
        
        # Close button in header
        close_button = QPushButton()
        close_button.setFixedSize(30, 30)
        close_button.setStyleSheet("""
            QPushButton {
                background-color: transparent;
                border: none;
            }
            QPushButton:hover {
                background-color: rgba(255, 0, 0, 150);
                border-radius: 15px;
            }
        """)
        close_button.setCursor(Qt.PointingHandCursor)
        
        # Set X icon
        try:
            close_button.setIcon(qta.icon("fa5s.times", color="white"))
            close_button.setIconSize(QSize(16, 16))
        except:
            close_button.setText("×")
            close_button.setStyleSheet("""
                QPushButton {
                    color: white;
                    font-size: 16pt;
                    font-weight: bold;
                    background-color: transparent;
                    border: none;
                }
                QPushButton:hover {
                    background-color: rgba(255, 0, 0, 150);
                    border-radius: 15px;
                }
            """)
        
        close_button.clicked.connect(self.reject)
        
        header_layout.addLayout(title_layout)
        header_layout.addStretch()
        header_layout.addWidget(close_button)
        
        return header

    def create_photo_panel(self):
        # Photo panel with border
        photo_frame = QFrame()
        
        photo_layout = QVBoxLayout(photo_frame)
        photo_layout.setContentsMargins(0, 0, 0, 0)
        photo_layout.setSpacing(0)
        
        # Add the photo with larger dimensions
        try:
            photo_label = QLabel()
            pixmap = QPixmap(relative_path(r"Data\الصور\esbo3elalam.png"))
            
            # Make photo larger while maintaining aspect ratio
            pixmap = pixmap.scaled(220, 320, Qt.KeepAspectRatio, Qt.SmoothTransformation)
            
            photo_label.setPixmap(pixmap)
            photo_label.setAlignment(Qt.AlignCenter)
            photo_label.setStyleSheet("background: transparent; border: none;")
            photo_layout.addWidget(photo_label, 1, alignment=Qt.AlignCenter)
        except Exception as e:
            # Fallback text if image fails to load
            fallback = QLabel("صورة أسبوع الآلام")
            fallback.setAlignment(Qt.AlignCenter)
            fallback.setStyleSheet("color: white; font-size: 14px; background: transparent;")
            photo_layout.addWidget(fallback)
    
        return photo_frame

    def create_buttons_panel(self):
        buttons_frame = QFrame()
        buttons_frame.setMinimumWidth(250)
        
        # Main layout for the buttons frame
        buttons_frame_layout = QVBoxLayout(buttons_frame)
        buttons_frame_layout.setContentsMargins(0, 0, 0, 0)
        
        # Create scroll area
        scroll_area = QScrollArea()
        scroll_area.setStyleSheet("""
            QScrollArea {
                background-color: transparent; 
                border: none;
            }
        """)
        scroll_area.setWidgetResizable(True)
        scroll_area.setMinimumHeight(350)
        
        # Set stylesheet for scrollbar
        scroll_area.verticalScrollBar().setStyleSheet("""
            QScrollBar:vertical {
                border: none;
                background: transparent;
                width: 10px;
            }
            QScrollBar::handle:vertical {
                background: rgba(255, 255, 255, 100);
                border-radius: 5px;
            }
            QScrollBar::add-line:vertical {
                background: none;
            }
            QScrollBar::sub-line:vertical {
                background: none;
            }
            QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical {
                background: none;
            }
        """)
        
        # Create content widget for the scroll area
        scroll_content = QWidget()
        scroll_content.setStyleSheet("background: transparent;")
        
        # Create layout for the scroll content
        buttons_layout = QVBoxLayout(scroll_content)
        buttons_layout.setSpacing(15)  # Increased spacing for better separation
        
        # Add button groups directly from button configs (single source of truth)
        for day in self.day_order:
            day_buttons = [cfg for cfg in self.button_configs.values() if cfg.get("day") == day]
            if day_buttons:
                self.add_button_group(buttons_layout, day, day_buttons)
        
        # Set the scroll content to the scroll area
        scroll_area.setWidget(scroll_content)
        buttons_frame_layout.addWidget(scroll_area)
        
        return buttons_frame

    def add_button_group(self, layout, day, buttons):
        # Create a container for the day section
        day_container = QFrame()
        day_container.setStyleSheet("""
            QFrame {
                background-color: rgba(80, 10, 10, 100);
                border-radius: 10px;
                padding: 5px;
            }
        """)
        
        day_layout = QVBoxLayout(day_container)
        day_layout.setContentsMargins(10, 10, 10, 15)
        day_layout.setSpacing(8)
        
        # Create and add label for the day with icon
        day_label_container = QHBoxLayout()
        
        # Try to add icon before day label
        try:
            day_icon = QLabel()
            # Choose different icons based on day of the week
            icon_name = "fa5s.calendar-day"
            if day == "الأحد":
                icon_name = "fa5s.church"
            elif day == "الخميس":
                icon_name = "fa5s.wine-glass"
            elif day == "الجمعة":
                icon_name = "fa5s.cross"
            elif day == "السبت":
                icon_name = "fa5s.menorah"
            
            icon_pixmap = qta.icon(icon_name, color="white").pixmap(16, 16)
            day_icon.setPixmap(icon_pixmap)
            day_icon.setStyleSheet("background: transparent;")
            day_label_container.addWidget(day_icon)
            day_label_container.addSpacing(8)
        except:
            pass
        
        day_label = QLabel(day)
        day_label.setStyleSheet("""
            QLabel {
                color: white;
                font-size: 16px;
                font-weight: bold;
                background-color: transparent;
            }
        """)
        day_label_container.addWidget(day_label)
        day_label_container.addStretch()
        
        day_layout.addLayout(day_label_container)
        
        # Add a subtle separator line
        separator = QFrame()
        separator.setFrameShape(QFrame.HLine)
        separator.setFrameShadow(QFrame.Sunken)
        separator.setStyleSheet("background-color: rgba(255, 255, 255, 70); margin: 2px 0;")
        separator.setMaximumHeight(1)
        day_layout.addWidget(separator)
        
        # Add buttons for this day
        for button_config in buttons:
            button = QPushButton(button_config["text"])
            button.setProperty("buttonId", button_config["id"])
            button.setCursor(Qt.PointingHandCursor)
            button.setStyleSheet(self.default_button_style)
            
            # Add shadow effect to button
            shadow = QGraphicsDropShadowEffect()
            shadow.setBlurRadius(10)
            shadow.setColor(QColor(0, 0, 0, 80))
            shadow.setOffset(2, 2)
            button.setGraphicsEffect(shadow)
            
            self.button_widgets_by_id[button_config["id"]] = button
            button.clicked.connect(self.handle_button_click)
            day_layout.addWidget(button)
        
        layout.addWidget(day_container)

    def handle_button_click(self):
        button = self.sender()
        if not button:
            return

        button_id = button.property("buttonId")

        if self.active_presentation_path and not self.is_presentation_open(self.active_presentation_path):
            self.clear_active_button()

        if button_id == self.active_button_id and self.active_presentation_path and self.is_presentation_open(self.active_presentation_path):
            self.open_sections_dialog_for_button(button_id, self.active_presentation_path)
            return

        self.run_pre_open_actions(button_id)

    def clear_active_button(self):
        self.active_button_id = None
        self.active_presentation_path = None
        for btn in self.button_widgets_by_id.values():
            btn.setStyleSheet(self.default_button_style)

    def mark_active_button(self, button_id):
        # Active source button glow while its presentation is still open.
        for current_id, btn in self.button_widgets_by_id.items():
            btn.setStyleSheet(self.active_button_style if current_id == button_id else self.default_button_style)
        self.active_button_id = button_id

    def is_presentation_open(self, presentation_path):
        if not presentation_path:
            return False
        try:
            target = os.path.abspath(presentation_path).lower()
            open_list = get_open_presentations() or []
            for pres_path in open_list:
                if os.path.abspath(pres_path).lower() == target:
                    return True
        except Exception:
            return False
        return False

    def open_sections_dialog_for_button(self, button_id, presentation_path):
        config = self.button_configs.get(button_id, {})
        filter_cfg = self.holy_week_filter_map.get(button_id, {})
        # Reads hardcoded filter config for clicked button.
        dialog = Elbas5aSectionSelectionDialog(
            parent=self,
            title="اختيار الأقسام",
            presentation_path=presentation_path,
            filter_mode=filter_cfg.get("mode", ""),
            filter_key=filter_cfg.get("key", ""),
            filter_keywords=filter_cfg.get("keywords", []),
            always_include_section_ids=self.holy_week_always_include_section_ids,
            source_button_id=button_id,
            source_button_label=config.get("text", button_id),
        )
        dialog.exec_()
        self.raise_()
        self.activateWindow()

    def run_pre_open_actions(self, button_id):
        """Apply section show/hide settings from button config before opening presentation."""
        config = self.button_configs.get(button_id)
        if not config:
            return

        if config.get("direct_open_only"):
            open_path = config.get("new_path")
            if not open_path:
                print(f"new_path is missing for {button_id}")
                return
            try:
                open_presentation_relative_path(open_path)
            except Exception as e:
                print(f"Error opening presentation for {button_id}: {e}")
            return

        # Build slide ranges fresh for each button click.
        self.show_slides = []

        old_path = config.get("old_path")
        new_path = config.get("new_path")

        if not new_path:
            print(f"new_path is missing for {button_id}")
            return

        # Replace old path with new path before modifications
        if old_path:
            try:
                replacefile(relative_path(old_path), relative_path(new_path))
                if button_id == "thurs_maundy_thursday":
                    replacefile(relative_path(r"كتاب المدائح.pptx"), relative_path(r"Data\CopyData\كتاب المدائح.pptx"))
                    elzoksologyat(self.excel, None, "باكر")
            except Exception as e:
                print(f"Error replacing file for {button_id}: {e}")
        
        ppt_file = relative_path(old_path)
        
        show_sections = config.get("show_sections", [])
        hide_sections = config.get("hide_sections", [])
        
        # Get VBA sheet name from config (needed for show_hide function)
        vba_sheet = config.get("vba_sheet") or "البصخة"
        
        tar7_values = find_slide_nums_arrays_v2(
            self.excel,
            vba_sheet,
            ['{9BEC2A31-DA53-45BE-AF39-7BE678B18A94}', '{5DF0EB10-99E1-41E1-AAD9-469FAC87A801}', '{931A7E35-642D-498D-A76F-2896828706A5}', '{378535F0-B031-4254-99A2-726504C335DB}', '{37F96881-032F-4EEB-8D10-5B5C2D5966F2}'],
            2,
            [2, 2, 2, 2, 2],
        )

        if isinstance(tar7_values, str):
            print(f"Error reading tar7 values from Excel: {tar7_values}")
            return

        if not isinstance(tar7_values, list) or len(tar7_values) != 5:
            print(f"Error: Expected 5 tar7 values from Excel, got {tar7_values}")
            return

        for i, val in enumerate(tar7_values):
            if isinstance(val, str) and ("Error" in val or "No corresponding" in val):
                print(f"tar7 lookup error for value {i+1}: {val}")
                return

        try:
            tar7_numeric_values = [int(v) for v in tar7_values]
        except (TypeError, ValueError) as e:
            print(f"Error converting tar7 values to numbers: {e}; values={tar7_values}")
            return

        if button_id in ["mon_day", "tue_day", "wed_day", "thurs_maundy_thursday"]:
            show_sections.extend(['{4684F45B-A266-433A-A91A-735CEF280595}'])
            tar71, tar73, tar76, tar79, tar711 = [v - 1 for v in tar7_numeric_values]
        else:
            show_sections.extend(['{390F461A-F6BF-498E-A7F7-AF7650D6DA16}'])
            tar71, tar73, tar76, tar79, tar711 = tar7_numeric_values

        self.show_slides.extend([[tar71, tar71], [tar73, tar73], [tar76, tar76], [tar79, tar79], [tar711, tar711]])

        # Apply show/hide sections if any are defined
        if show_sections or hide_sections:
            try:
                show_hide_insertImage_replaceText(
                    ppt_file,
                    self.excel,  # Pass the Excel file path
                    vba_sheet,   # Pass the sheet name
                    show_sections=show_sections,  # section IDs to show
                    hide_sections=hide_sections   # section IDs to hide
                )
            except Exception as e:
                print(f"Error applying section visibility for {button_id}: {e}")
        
        presentation = open_presentation_relative_path(old_path)
        presentation_path = os.path.abspath(relative_path(old_path or new_path))

        show_slides(presentation, self.show_slides)  # Show slides based on VBA values
        if button_id == "thurs_maundy_thursday":
            run_vba_with_slide_id_bakr_aashya(self.excel, vba_sheet, old_path, presentation, '{A5B9CE2F-90E3-44D7-B22F-CAE6783C8E2F}')
        # Use self.excel instead of getting from config
        if self.excel and vba_sheet:
            if os.path.exists(self.excel):
                if hasattr(presentation, "VBProject"):
                    try:
                        self.run_vba(
                            vba_sheet,
                            relative_path(new_path),
                            presentation,
                            button_id,
                        )
                    except Exception as e:
                        print(f"Error running VBA for {button_id}: {e}")
            else:
                print(f"Excel file not found: {self.excel}")
        else:
            print(f"VBA configuration incomplete for {button_id}: excel={self.excel}, sheet={vba_sheet}")

        self.active_presentation_path = presentation_path
        self.mark_active_button(button_id)
        self.open_sections_dialog_for_button(button_id, presentation_path)

    def mousePressEvent(self, event):
        # Allow dragging the frameless window from the header area
        if event.button() == Qt.LeftButton and event.y() < 50:  # 50 is header height
            self._drag_pos = event.globalPos() - self.frameGeometry().topLeft()
            event.accept()

    def mouseMoveEvent(self, event):
        # Move the window with mouse
        if event.buttons() == Qt.LeftButton and hasattr(self, '_drag_pos'):
            self.move(event.globalPos() - self._drag_pos)
            event.accept()

    def run_vba(self, sheet, prs, presentation, button_id):
        # Validate sheet parameter
        if sheet is None or not isinstance(sheet, str) or not sheet.strip():
            print(f"VBA sheet name is invalid or None: {repr(sheet)}")
            return
        
        # Use self.excel - it's already initialized and validated in __init__
        if self.excel is None or not os.path.exists(self.excel):
            print(f"Excel file not available: {self.excel}")
            return
        
        if button_id in ["sun_night_monday", "mon_day", "mon_night_tuesday", "tue_day", "tue_night_wednesday", "wed_day", "wed_night_thursday"]:
            vba_values = find_slide_nums_arrays_v2(
                self.excel,
                sheet,
                ['{81AD693D-C29D-45B2-95D4-0C613DDDE79B}', '{C77E4783-C285-4706-B437-1BBF92A3E954}', '{D0A19F68-9B40-46E9-8CEE-20069205C2F5}', '{9D06AEAB-48F5-44D5-B54C-6166991785C3}', '{B49BB649-3054-4719-BC60-F1B97BA48DF9}', '{E4CF5EBF-7A16-40CE-8CF3-AF5E72A35258}'],
                2,
                [2, 2, 2, 2, 2, 2],
            )
        elif button_id in ["thurs_maundy_thursday"]:
            vba_values = find_slide_nums_arrays_v2(
                self.excel,
                sheet,
                ['{81AD693D-C29D-45B2-95D4-0C613DDDE79B}', '{C77E4783-C285-4706-B437-1BBF92A3E954}', '{D0A19F68-9B40-46E9-8CEE-20069205C2F5}', '{9D06AEAB-48F5-44D5-B54C-6166991785C3}', '{B49BB649-3054-4719-BC60-F1B97BA48DF9}', '{D4DC8BC9-795A-4AB0-8BF0-67ABCD5AA7BC}', '{5F8571AA-3756-4993-80D9-31D04DB137B4}', '{F0D339B8-4A92-41DD-94E0-D5AF495E0092}'],
                2,
                [2, 2, 2, 2, 2, 2, 2, 1],
            )
        else:
            return

        # Check if vba_values is a string (error message from find_slide_nums_arrays_v2)
        if isinstance(vba_values, str):
            print(f"Error reading Excel file: {vba_values}")
            return
        
        # Validate value count: 8 for thurs_maundy_thursday (with extra jump slide + destination), 6 for others
        expected_count = 8 if button_id == "thurs_maundy_thursday" else 6
        if not isinstance(vba_values, list) or len(vba_values) != expected_count:
            print(f"Error: Expected {expected_count} values from Excel, got {len(vba_values)}")
            return
        
        # Check if any values are error messages (marked as strings containing "Error" or "No corresponding")
        for i, val in enumerate(vba_values):
            if isinstance(val, str) and ("Error" in val or "No corresponding" in val):
                print(f"Excel lookup error for value {i+1}: {val}")
                return
        
        # Unpack the values
        try:
            Hour1, Hour3, Hour6, Hour9, Hour11 = [v - 1 for v in vba_values[:5]]
            returnH1, returnH3, returnH6, returnH9, returnH11 = Hour1 + 1, Hour3 + 1, Hour6 + 1, Hour9 + 1, Hour11 + 1
            returnFromTasbha = vba_values[5]
            # For thurs_maundy_thursday, extract 2 additional slide IDs
            jumpSlide = None
            jumpSlideDestination = None
            if button_id == "thurs_maundy_thursday" and len(vba_values) >= 8:
                jumpSlide = vba_values[6] # Convert to 0-indexed
                jumpSlideDestination = vba_values[7]  # Used as-is
        except (ValueError, TypeError) as e:
            print(f"Error unpacking VBA values: {e}")
            return

        # Build list of slide indices for get_slide_ids_by_numbers
        slide_indices = [Hour1, Hour3, Hour6, Hour9, Hour11, returnH1, returnH3, returnH6, returnH9, returnH11, returnFromTasbha]
        if jumpSlide is not None:
            slide_indices.extend([jumpSlide, jumpSlideDestination])
        
        vba_main_slides_ids = get_slide_ids_by_numbers(prs, slide_indices)
        
        # Unpack returned slide IDs
        if button_id == "thurs_maundy_thursday" and len(vba_main_slides_ids) == 13:
            Hour1, Hour3, Hour6, Hour9, Hour11, returnH1, returnH3, returnH6, returnH9, returnH11, returnFromTasbha, jumpSlideID, jumpSlideDestinationID = vba_main_slides_ids
        else:
            Hour1, Hour3, Hour6, Hour9, Hour11, returnH1, returnH3, returnH6, returnH9, returnH11, returnFromTasbha = vba_main_slides_ids
            jumpSlideID = None
            jumpSlideDestinationID = None  

        if button_id in ["sun_night_monday", "mon_day", "mon_night_tuesday", "thurs_maundy_thursday"]:
            hour1_show = "tasbha"
            hour3_show = "tasbha"
            hour6_show = "tasbha"
            hour9_show = "tasbha"
            hour11_show = "tasbha"
        elif button_id in ["tue_night_wednesday", "wed_day", "wed_night_thursday"]:
            hour1_show = "tasbha2"
            hour3_show = "tasbha2"
            hour6_show = "tasbha2"
            hour9_show = "tasbha2"
            hour11_show = "tasbha2"
        elif button_id == "tue_day":
            hour1_show = "tasbha"
            hour3_show = "tasbha"
            hour6_show = "tasbha"
            hour9_show = "tasbha"
            hour11_show = "tasbha2"
        else:
            return

        vba_project = presentation.VBProject
        modules = vba_project.VBComponents
        new_module = modules.Add(1)

        # Module-level variables to track trigger and deferred return target
        vba_code = "Dim triggeringHourIndex As Long  ' Stores the hour index: 1, 3, 6, 9, or 11\n"
        vba_code += "Dim pendingReturnSlideID As Long\n\n"
        
        vba_code += "Sub OnSlideShowPageChange()\n"
        vba_code += "    Dim currentSlideID As Long\n"
        vba_code += "    Dim nextSlideIndex As Long\n"
        vba_code += "    currentSlideID = ActivePresentation.SlideShowWindow.View.Slide.SlideID\n\n"
        vba_code += "    If pendingReturnSlideID <> 0 Then\n"
        vba_code += "        nextSlideIndex = GetSlideIndexByID(pendingReturnSlideID)\n"
        vba_code += "        pendingReturnSlideID = 0\n"
        vba_code += "        If nextSlideIndex > 0 Then\n"
        vba_code += "            ActivePresentation.SlideShowWindow.View.GotoSlide nextSlideIndex\n"
        vba_code += "        End If\n"
        vba_code += "        Exit Sub\n"
        vba_code += "    End If\n\n"
        vba_code += "    Select Case currentSlideID\n"
        vba_code += f"        Case {Hour1}\n"
        vba_code += "            triggeringHourIndex = 1\n"
        vba_code += f"            ActivePresentation.SlideShowWindow.View.GotoNamedShow \"{hour1_show}\"\n"
        vba_code += f"        Case {Hour3}\n"
        vba_code += "            triggeringHourIndex = 3\n"
        vba_code += f"            ActivePresentation.SlideShowWindow.View.GotoNamedShow \"{hour3_show}\"\n"
        vba_code += f"        Case {Hour6}\n"
        vba_code += "            triggeringHourIndex = 6\n"
        vba_code += f"            ActivePresentation.SlideShowWindow.View.GotoNamedShow \"{hour6_show}\"\n"
        vba_code += f"        Case {Hour9}\n"
        vba_code += "            triggeringHourIndex = 9\n"
        vba_code += f"            ActivePresentation.SlideShowWindow.View.GotoNamedShow \"{hour9_show}\"\n"
        vba_code += f"        Case {Hour11}\n"
        vba_code += "            triggeringHourIndex = 11\n"
        vba_code += f"            ActivePresentation.SlideShowWindow.View.GotoNamedShow \"{hour11_show}\"\n"
        vba_code += f"        Case {returnFromTasbha}\n"
        vba_code += "            ' Last slide of custom slideshow: close named show, restart, then jump\n"
        vba_code += "            If triggeringHourIndex <> 0 Then\n"
        vba_code += "                pendingReturnSlideID = GetReturnSlideIDByTriggeringHour()\n"
        vba_code += "                triggeringHourIndex = 0  ' Reset flag\n"
        vba_code += "                ActivePresentation.SlideShowWindow.View.EndNamedShow\n"
        vba_code += "                StartSlideshow\n"
        vba_code += "            End If\n"
        # For thurs_maundy_thursday, add case for jump slide that goes directly to destination
        if jumpSlideID is not None and jumpSlideDestinationID is not None:
            vba_code += f"        Case {jumpSlideID}\n"
            vba_code += f"            ' Direct jump slide to destination\n"
            vba_code += f"            ActivePresentation.SlideShowWindow.View.GotoSlide GetSlideIndexByID({jumpSlideDestinationID})\n"
        vba_code += "    End Select\n"
        vba_code += "End Sub\n\n"

        vba_code += "Sub StartSlideshow()\n"
        vba_code += "    Dim newShow As SlideShowWindow\n"
        vba_code += "    Dim targetSlideIndex As Long\n"
        vba_code += "    Set newShow = ActivePresentation.SlideShowSettings.Run\n"
        vba_code += "\n"
        vba_code += "    ' Jump immediately to avoid black screen until next click\n"
        vba_code += "    If pendingReturnSlideID <> 0 Then\n"
        vba_code += "        targetSlideIndex = GetSlideIndexByID(pendingReturnSlideID)\n"
        vba_code += "        pendingReturnSlideID = 0\n"
        vba_code += "        If targetSlideIndex > 0 Then\n"
        vba_code += "            newShow.View.GotoSlide targetSlideIndex\n"
        vba_code += "        End If\n"
        vba_code += "    End If\n"
        vba_code += "End Sub\n\n"
        
        # Function to get the return slide ID based on triggering hour
        vba_code += "Function GetReturnSlideIDByTriggeringHour() As Long\n"
        vba_code += "    Dim returnSlideID As Long\n\n"
        vba_code += "    Select Case triggeringHourIndex\n"
        vba_code += f"        Case 1  ' Hour1\n"
        vba_code += f"            returnSlideID = {returnH1}\n"
        vba_code += f"        Case 3  ' Hour3\n"
        vba_code += f"            returnSlideID = {returnH3}\n"
        vba_code += f"        Case 6  ' Hour6\n"
        vba_code += f"            returnSlideID = {returnH6}\n"
        vba_code += f"        Case 9  ' Hour9\n"
        vba_code += f"            returnSlideID = {returnH9}\n"
        vba_code += f"        Case 11 ' Hour11\n"
        vba_code += f"            returnSlideID = {returnH11}\n"
        vba_code += "        Case Else\n"
        vba_code += "            returnSlideID = 0\n"
        vba_code += "    End Select\n\n"
        vba_code += "    GetReturnSlideIDByTriggeringHour = returnSlideID\n"
        vba_code += "End Function\n\n"
        
        vba_code += """Function GetSlideIndexByID(slideID As Long) As Long
    Dim slide As slide
    For Each slide In ActivePresentation.Slides
        If slide.slideID = slideID Then
            GetSlideIndexByID = slide.SlideIndex
            Exit Function
        End If
    Next slide
    MsgBox "Slide ID " & slideID & " not found.", vbExclamation
    GetSlideIndexByID = 0
End Function
"""

        new_module.CodeModule.AddFromString(vba_code)

        presentation.SlideShowSettings.Run()
        presentation.Application.Run("OnSlideShowPageChange")

