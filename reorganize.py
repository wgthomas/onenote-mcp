#!/usr/bin/env python3
"""Reorganize OneNote notebooks into a single clean structure.

Moves all pages from both notebooks into organized sections in Wylvek's Notebook.
Backups at Z:/wgtho/ before running.

Run from D:/MCP_servers/onenote-mcp/
"""

import sys
import re
import time
import traceback

sys.path.insert(0, ".")
from onenote_lib import com_client

NB1_ID = "{44B98C82-9DC5-4FB0-AB34-F4C654E32241}{1}{B0}"  # Wylvek's Notebook
NB2_ID = "{5F6337CE-912C-42EC-A3BE-72B8803ADA86}{1}{B0}"  # 3734 Network_Notes

# Track results
moved = 0
failed = 0
deleted = 0
skipped = 0


def move_page(page_id, target_section_id, name="", desired_level=None):
    """Move a page by copying content (including images) to target section, then deleting original."""
    global moved, failed
    try:
        # 1. Get source page XML
        source_xml = com_client.get_page_content(page_id)

        # 2. Extract old page ID from XML
        old_id_match = re.search(r'ID="([^"]+)"', source_xml)
        if not old_id_match:
            print(f"    FAIL: Could not find page ID in XML for {name}")
            failed += 1
            return None
        old_id = old_id_match.group(1)

        # 3. Create new blank page in target section
        new_page_id = com_client.create_new_page(target_section_id)

        # 4. Replace page ID in source XML (only the first occurrence in <one:Page>)
        modified_xml = source_xml.replace(f'ID="{old_id}"', f'ID="{new_page_id}"', 1)

        # 5. Strip all objectID attributes (they reference old page's internal objects)
        modified_xml = re.sub(r'\s*objectID="[^"]+"', '', modified_xml)

        # 6. Handle images: replace CallbackID with base64 Data
        callbacks = re.findall(r'callbackID="([^"]+)"', modified_xml)
        if callbacks:
            print(f"    Fetching {len(callbacks)} image(s)...")
        for cb_id in callbacks:
            try:
                b64 = com_client.get_binary_content(page_id, cb_id).strip()
                # Replace CallbackID element with Data element
                modified_xml = re.sub(
                    rf'<one:CallbackID\s+callbackID="{re.escape(cb_id)}"\s*/>',
                    f'<one:Data>{b64}</one:Data>',
                    modified_xml,
                )
            except Exception as e:
                print(f"    Warning: Image {cb_id[:30]}... failed: {e}")

        # 7. Override page level if requested
        if desired_level is not None:
            if desired_level == 0:
                # Level 0 = top-level: REMOVE pageLevel attribute (min value is 1)
                modified_xml = re.sub(r'\s*pageLevel="\d+"', '', modified_xml, count=1)
            else:
                if 'pageLevel="' in modified_xml:
                    modified_xml = re.sub(
                        r'pageLevel="\d+"',
                        f'pageLevel="{desired_level}"',
                        modified_xml,
                        count=1,
                    )
                else:
                    # Add pageLevel attribute to Page element
                    modified_xml = modified_xml.replace(
                        '<one:Page ',
                        f'<one:Page pageLevel="{desired_level}" ',
                        1,
                    )

        # 8. Update the new page with source content
        com_client.update_page_content(modified_xml)

        # 9. Delete the source page
        com_client.delete_hierarchy(page_id)

        moved += 1
        return new_page_id

    except Exception as e:
        print(f"    FAIL: {name}: {e}")
        traceback.print_exc()
        failed += 1
        return None


def delete_page(page_id, name=""):
    """Delete a page (safe to call on already-deleted pages)."""
    global deleted
    try:
        com_client.delete_hierarchy(page_id)
        deleted += 1
        print(f"  Deleted: {name}")
    except Exception as e:
        if "0x8004200E" in str(e) or "does not exist" in str(e).lower():
            print(f"  Already gone: {name}")
        else:
            print(f"  FAIL deleting {name}: {e}")


def delete_section(section_id, name=""):
    """Delete a section (safe to call on already-deleted sections)."""
    try:
        com_client.delete_hierarchy(section_id)
        print(f"  Deleted section: {name}")
    except Exception as e:
        if "0x8004200E" in str(e) or "does not exist" in str(e).lower():
            print(f"  Already gone: {name}")
        else:
            print(f"  FAIL deleting section {name}: {e}")


# =====================================================================
# Phase 1: Create target sections
# =====================================================================

def create_target_sections():
    """Create clean target sections in Wylvek's Notebook."""
    print("\n=== PHASE 1: Creating target sections ===")
    sections = {}
    for name in ["Networking", "Programming", "Cybertron", "AI_and_ML", "Career", "Personal", "Credentials"]:
        sid = com_client.open_hierarchy(f"{name}.one", NB1_ID)
        sections[name] = sid
        print(f"  {name}: {sid}")
    return sections


# =====================================================================
# Phase 2: Delete empty/useless pages and old index pages
# =====================================================================

def delete_junk_pages():
    """Delete empty, untitled, and old index pages."""
    print("\n=== PHASE 2: Deleting junk and old index pages ===")

    # Empty/untitled pages
    junk = [
        ("{722F6090-9D24-0BA3-227F-A162FFE09B6A}{1}{E1911129322946299700841919206963003467635341}", "compusharecontact/Untitled"),
        ("{07D75124-FAD0-08CE-1D52-DF726D7CBF59}{1}{E201172912577495716800320126228906997232732331}", "Quick Notes/Untitled"),
        ("{1EE74EA4-2EE1-0D72-3387-DEEE0DBC7FD4}{1}{E1910000902351345835201934470318477145753681}", "From_thedose/Untitled 1"),
        ("{1EE74EA4-2EE1-0D72-3387-DEEE0DBC7FD4}{1}{E189794062057559030911974894908414345687271}", "From_thedose/Untitled 2"),
        ("{F31E127D-AB5D-4AE2-BBDC-2FBDC2487E73}{1}{E19569579417230897032020148934042292814101781}", "october_2025/Untitled"),
        ("{98294380-67E3-000F-115B-C2BE14C8138C}{1}{E1836122863132289441720126192946862964257901}", "CISSP/Untitled"),
        ("{F882D5EB-DBA3-0E97-111D-AA1FBFD580BC}{1}{E1837038017008423228920144383490953573768061}", "cMNL Class/Untitled"),
        ("{46775A51-35DA-07E0-3E1D-1285D8A128F9}{1}{E1847636178622340123520177345597687660541591}", "family/Untitled"),
    ]

    # Old index pages from previous organization pass
    old_indexes = [
        ("{9B939CCB-C163-4957-BC87-CAB01BFDA076}{1}{E19554730420234207745320162570782452112600601}", "_Index/Master Index"),
        ("{90E22DAF-D3AF-484E-B11E-0246DEAD9671}{1}{E195532186980318532533184975739870942545241}", "_Networking/Index"),
        ("{ED21DD7A-1B0C-4BBD-9187-33FB52A38CE8}{1}{E1953069795868320373541974653155029294603541}", "_Programming/Index"),
        ("{BEB1D1A6-8432-470F-A679-7853C53C68D1}{1}{E1956909685733593712731980479084551884184721}", "_Cybertron_Infra/Index"),
        ("{6AD896F9-37C7-421E-B4BB-1B8FE8DEBC0E}{1}{E19469212383781521538520109538782690660189561}", "_Career/Index"),
        ("{6CC1B568-8CCF-0183-3FBF-18C3141471FB}{1}{E1954722946183792180611923298079816121127841}", "AI_stuff/Index"),
        ("{07D75124-FAD0-08CE-1D52-DF726D7CBF59}{1}{E19544348731678040997720105526425923246212141}", "Quick Notes/Index"),
        ("{420B032E-77FD-48CC-B294-B8EA099DDA20}{1}{E1950436284878588668741986360131924338584711}", "NB2/_Index/Index"),
    ]

    for pid, name in junk + old_indexes:
        delete_page(pid, name)


# =====================================================================
# Phase 3: Move pages to target sections
# =====================================================================

def move_networking_pages(sections):
    """Move all networking-related pages."""
    print("\n=== Moving pages to NETWORKING ===")
    target = sections["Networking"]

    pages = [
        # -- Cisco -- (Cisco page already moved in test run)
        ("{07D75124-FAD0-08CE-1D52-DF726D7CBF59}{1}{E1911266691516780529061972731741347786788321}", 1, "Collaboration services"),
        ("{07D75124-FAD0-08CE-1D52-DF726D7CBF59}{1}{E1815318227699242178620174117899093372703121}", 0, "Vpc vs trunks"),
        ("{07D75124-FAD0-08CE-1D52-DF726D7CBF59}{1}{E1953449173380278974841970770863442792129761}", 0, "ISE WEBX info"),
        ("{07D75124-FAD0-08CE-1D52-DF726D7CBF59}{1}{E187786303654957661751924122529557048466381}", 0, "Cisco Learning Network info"),
        # -- Palo Alto / Firewall --
        ("{07D75124-FAD0-08CE-1D52-DF726D7CBF59}{1}{E188101577193432154621988642047567342538561}", 0, "Palo Alto"),
        ("{07D75124-FAD0-08CE-1D52-DF726D7CBF59}{1}{E1873045529307623551120157344110561756275481}", 1, "PaloAlto Packet Flow"),
        ("{07D75124-FAD0-08CE-1D52-DF726D7CBF59}{1}{E183912631713413636961944984480681256297001}", 0, "Bfd f5 palo alto"),
        ("{07D75124-FAD0-08CE-1D52-DF726D7CBF59}{1}{E187717704238765207111943037931339728481511}", 0, "ForcePoint"),
        ("{07D75124-FAD0-08CE-1D52-DF726D7CBF59}{1}{E184387404571052361551911568400873917117801}", 0, "Force Point linux shell"),
        ("{07D75124-FAD0-08CE-1D52-DF726D7CBF59}{1}{E19505073967075205193920165979127621433240721}", 0, "Risk determination commands"),
        # -- General Networking --
        ("{07D75124-FAD0-08CE-1D52-DF726D7CBF59}{1}{E201432214016380607977220182299479257853222751}", 0, "Web Notes"),
        ("{07D75124-FAD0-08CE-1D52-DF726D7CBF59}{1}{E182667446853408865041939586767464623181081}", 0, "Linux IP Troubleshooting"),
        ("{07D75124-FAD0-08CE-1D52-DF726D7CBF59}{1}{E1853004631723051568220184409745284664115271}", 0, "VM ware esxi vmnic vswitch"),
        # -- From PAXPATRIS --
        ("{56525385-122B-0E37-29C3-67F84037DB2B}{1}{E1849894137918472915920180627030873199600961}", 0, "NETWORK_REL"),
        # -- From cMNL Class --
        ("{F882D5EB-DBA3-0E97-111D-AA1FBFD580BC}{1}{E1849715188207639930120163433025403077156051}", 0, "DCNM discovery 2"),
        ("{F882D5EB-DBA3-0E97-111D-AA1FBFD580BC}{1}{E183587957691840144801992012294773215247541}", 1, "Discovery 2"),
        ("{F882D5EB-DBA3-0E97-111D-AA1FBFD580BC}{1}{E183044950486662679491971626086483681290691}", 1, "Discovery 3"),
        ("{F882D5EB-DBA3-0E97-111D-AA1FBFD580BC}{1}{E183431859230464683811915437411460424051401}", 1, "Discovery 4"),
        ("{F882D5EB-DBA3-0E97-111D-AA1FBFD580BC}{1}{E185297954874728338751997509767528534881681}", 0, "Discovery 6"),
        # -- From random notes --
        ("{B3EE65E4-9E27-0930-1CB8-6D89BFDD60CC}{1}{E189171892150017328091930501726057159099151}", 0, "Sec+ notes"),
        # -- From NB2 network stuff --
        ("{E97D0E54-CD21-0496-2EB5-E0A9E34F94C9}{1}{E178186253618707378420103248628521538365951}", 0, "Unifi info"),
        ("{E97D0E54-CD21-0496-2EB5-E0A9E34F94C9}{1}{E191152715346798140041185315194311628705501}", 1, "Url codes"),
        ("{E97D0E54-CD21-0496-2EB5-E0A9E34F94C9}{1}{E179460163379357698920177801655044159488321}", 1, "Udm commands"),
        ("{E97D0E54-CD21-0496-2EB5-E0A9E34F94C9}{1}{E17603379686720850591985516220349891973351}", 1, "ESXI"),
        ("{E97D0E54-CD21-0496-2EB5-E0A9E34F94C9}{1}{E1874689202314897987520137839726888632008101}", 2, "RHEL"),
        ("{E97D0E54-CD21-0496-2EB5-E0A9E34F94C9}{1}{E182838106948649423751997542190177410503761}", 1, "Commands"),
        ("{E97D0E54-CD21-0496-2EB5-E0A9E34F94C9}{1}{E1858981465655992158820183609188856865044541}", 1, "Unifi controller stop service"),
        ("{E97D0E54-CD21-0496-2EB5-E0A9E34F94C9}{1}{E188076108314512660601957311768636773836901}", 1, "Ap led flash codes"),
        ("{E97D0E54-CD21-0496-2EB5-E0A9E34F94C9}{1}{E182913618017944805391934206115114239068321}", 0, "Windows firewall Track activity"),
        ("{E97D0E54-CD21-0496-2EB5-E0A9E34F94C9}{1}{E188854829118080569491936848389844758415751}", 0, "USG firewall notes"),
        ("{E97D0E54-CD21-0496-2EB5-E0A9E34F94C9}{1}{E1872875788281676512520164017394891700854261}", 0, "Gaming configs"),
        ("{E97D0E54-CD21-0496-2EB5-E0A9E34F94C9}{1}{E182596594575083505951937020676462626406961}", 0, "Windows networking"),
        ("{E97D0E54-CD21-0496-2EB5-E0A9E34F94C9}{1}{E1947914786469783482371987353404336966734691}", 0, "NetDuma"),
        ("{E97D0E54-CD21-0496-2EB5-E0A9E34F94C9}{1}{E186787816012808353091981958986477936671191}", 0, "Pi Hole"),
    ]

    for pid, level, name in pages:
        print(f"  Moving: {name} (level {level})")
        move_page(pid, target, name, desired_level=level)


def move_programming_pages(sections):
    """Move all programming/scripting pages."""
    print("\n=== Moving pages to PROGRAMMING ===")
    target = sections["Programming"]

    pages = [
        # -- Python --
        ("{07D75124-FAD0-08CE-1D52-DF726D7CBF59}{1}{E19102208468962347888620109507544841326487181}", 0, "python"),
        ("{07D75124-FAD0-08CE-1D52-DF726D7CBF59}{1}{E182058528055222227271910614204872779891601}", 1, "Functions n stuff"),
        ("{07D75124-FAD0-08CE-1D52-DF726D7CBF59}{1}{E181673345918869900211928062621580393172801}", 0, "pip"),
        ("{07D75124-FAD0-08CE-1D52-DF726D7CBF59}{1}{E18665502471654695949187204573105131983501}", 0, "PDF issues"),
        ("{07D75124-FAD0-08CE-1D52-DF726D7CBF59}{1}{E19704805757773163670220130659334112992146371}", 0, "import subprocess"),
        ("{07D75124-FAD0-08CE-1D52-DF726D7CBF59}{1}{E1994943659569308760311915583440219722281411}", 0, "cmd_output = llm_control_cmd"),
        # -- PowerShell / CMD --
        ("{07D75124-FAD0-08CE-1D52-DF726D7CBF59}{1}{E1897349385151498189820134339730351212375591}", 0, "Powershell notes"),
        ("{07D75124-FAD0-08CE-1D52-DF726D7CBF59}{1}{E19474797630004590254720162187993781424027171}", 0, "Windows command line"),
        ("{07D75124-FAD0-08CE-1D52-DF726D7CBF59}{1}{E19111691766118514599420166947950331369666871}", 0, "CMD and batch"),
        ("{07D75124-FAD0-08CE-1D52-DF726D7CBF59}{1}{E1845129711909071695320131895560480043584251}", 0, "Windows Trace"),
        ("{07D75124-FAD0-08CE-1D52-DF726D7CBF59}{1}{E1810954117176200774720112086202908811853311}", 0, "No admin installs"),
        # -- Linux --
        ("{07D75124-FAD0-08CE-1D52-DF726D7CBF59}{1}{E181752048706630413121973527241883181583551}", 0, "sed"),
        # -- Docker --
        ("{07D75124-FAD0-08CE-1D52-DF726D7CBF59}{1}{E181419180060336290501920306819554362265611}", 0, "Docker cheat sheet"),
        ("{AD0E851F-7621-04B6-29C7-B0C1FEEE1E77}{1}{E1819599810125237372220165754603024361231931}", 0, "Docker (NB2)"),
        ("{AD0E851F-7621-04B6-29C7-B0C1FEEE1E77}{1}{E19106627404969765408920135495342840337137031}", 1, "Open webui"),
        # -- Tools --
        ("{07D75124-FAD0-08CE-1D52-DF726D7CBF59}{1}{E1950128389903434610191939236716947710055851}", 0, "VLC NOTES"),
        ("{07D75124-FAD0-08CE-1D52-DF726D7CBF59}{1}{E1897831491404860077120143890670526559465501}", 0, "VISIO visio"),
        ("{07D75124-FAD0-08CE-1D52-DF726D7CBF59}{1}{E1869482312124332184820139085061226676257841}", 0, "windows (robocopy etc)"),
        # -- From NB2 Misc_Tech_stuff --
        ("{AD0E851F-7621-04B6-29C7-B0C1FEEE1E77}{1}{E1869300511770508750920112350426171056939751}", 0, "Arcana Machina"),
        ("{AD0E851F-7621-04B6-29C7-B0C1FEEE1E77}{1}{E1818820715461985082820114613815472158630641}", 1, "Windows Sacra"),
        ("{AD0E851F-7621-04B6-29C7-B0C1FEEE1E77}{1}{E1910365793122064638841988005525452704944761}", 1, "Linux Sacra"),
        ("{AD0E851F-7621-04B6-29C7-B0C1FEEE1E77}{1}{E1862945561399689381820113862478414304371661}", 2, "Centos"),
        ("{AD0E851F-7621-04B6-29C7-B0C1FEEE1E77}{1}{E19481843268451511742120181468691884161665581}", 2, "Python Stuff"),
        ("{AD0E851F-7621-04B6-29C7-B0C1FEEE1E77}{1}{E19534471352201478866020141311267437579708201}", 2, "FAUXMO"),
        ("{AD0E851F-7621-04B6-29C7-B0C1FEEE1E77}{1}{E1898231025792317324920125226556358046618101}", 0, "Web Scraping"),
        ("{AD0E851F-7621-04B6-29C7-B0C1FEEE1E77}{1}{E1881069183226670101720112310217584231688101}", 0, "Windows genera"),
    ]

    for pid, level, name in pages:
        print(f"  Moving: {name} (level {level})")
        move_page(pid, target, name, desired_level=level)


def move_cybertron_pages(sections):
    """Move all infrastructure/Cybertron pages."""
    print("\n=== Moving pages to CYBERTRON ===")
    target = sections["Cybertron"]

    pages = [
        # -- Core Infra from NB2 infrastructure_tech --
        ("{33A07EDC-414C-015F-0CAC-1F00FA7F98FF}{1}{E183575360390355692381944233469792357788781}", 0, "Truenas"),
        ("{33A07EDC-414C-015F-0CAC-1F00FA7F98FF}{1}{E1910179028766846242601993943274332557383891}", 0, "Home Assistant"),
        ("{33A07EDC-414C-015F-0CAC-1F00FA7F98FF}{1}{E1815760895588947206120165961741908732057711}", 1, "Tasmota"),
        ("{33A07EDC-414C-015F-0CAC-1F00FA7F98FF}{1}{E186812211710492110181969427539842588098701}", 0, "Megatron"),
        ("{33A07EDC-414C-015F-0CAC-1F00FA7F98FF}{1}{E182462232840627462501932443383206915594871}", 0, "Legacy network info"),
        ("{33A07EDC-414C-015F-0CAC-1F00FA7F98FF}{1}{E1770900405690835022184122374062910049451}", 0, "Splashtop info"),
        # -- Cybertron Services from AI_stuff --
        ("{6CC1B568-8CCF-0183-3FBF-18C3141471FB}{1}{E19479829063473105229620167058921564271842841}", 0, "vault"),
        ("{6CC1B568-8CCF-0183-3FBF-18C3141471FB}{1}{E19506837813211255725620102904527635798094601}", 0, "Rustdesk"),
        ("{6CC1B568-8CCF-0183-3FBF-18C3141471FB}{1}{E19550779222332021313520144568218657620499691}", 0, "Git"),
        ("{6CC1B568-8CCF-0183-3FBF-18C3141471FB}{1}{E1957238659870990588291917006152429375891761}", 0, "Openclaw"),
        ("{6CC1B568-8CCF-0183-3FBF-18C3141471FB}{1}{E195738826567941257299181448239107549152381}", 0, "Discord"),
        ("{6CC1B568-8CCF-0183-3FBF-18C3141471FB}{1}{E1953573242609948333401992840260821178558911}", 0, "Things I want in cybertron"),
        # -- Pi projects --
        ("{829A0E69-79D8-040D-1582-1D9D43EAD3C9}{1}{E1895161366364026851120155271230596220556811}", 0, "Pi - updating"),
        ("{829A0E69-79D8-040D-1582-1D9D43EAD3C9}{1}{E18808148159647557275184548580235580221691}", 0, "Pi - webcam"),
        ("{3A7C5248-28F2-05A3-356F-F7162F2CC03B}{1}{E18195496209200974733186786609307287821161}", 0, "BeagleBone Black"),
        ("{33A07EDC-414C-015F-0CAC-1F00FA7F98FF}{1}{E186834916725875675161960056783683419986411}", 0, "OctoPrint"),
        ("{33A07EDC-414C-015F-0CAC-1F00FA7F98FF}{1}{E1876960381895166476520180151243122336279091}", 0, "Mine Craft Pi"),
        ("{33A07EDC-414C-015F-0CAC-1F00FA7F98FF}{1}{E1910580305381068636251937882606820039173451}", 0, "PullyPi"),
        ("{33A07EDC-414C-015F-0CAC-1F00FA7F98FF}{1}{E1883579834514827414320126883914571497997081}", 1, "Element14 RPi Node Red Alexa"),
        # -- Windows/GPU --
        ("{B3EE65E4-9E27-0930-1CB8-6D89BFDD60CC}{1}{E181550856004941129371937948914931601193121}", 0, "Gpu windows troubleshooting"),
        # -- Misc tech from NB2 --
        ("{AD0E851F-7621-04B6-29C7-B0C1FEEE1E77}{1}{E189422187133085305101994411747808587804451}", 0, "HD_homerun"),
        ("{AD0E851F-7621-04B6-29C7-B0C1FEEE1E77}{1}{E1835810259548718685920130793493922653161211}", 0, "USB drive notes"),
        ("{AD0E851F-7621-04B6-29C7-B0C1FEEE1E77}{1}{E19547975708728708559217672996120835639341}", 0, "Spark40"),
        ("{56AF71BA-FEFD-0322-086E-9136C3604C33}{1}{E186810120778138715861943724539574850698301}", 0, "Acronis"),
        ("{AD0E851F-7621-04B6-29C7-B0C1FEEE1E77}{1}{E1858511305043502703320122866479410720284561}", 0, "Acronis info"),
        # -- Air Quality from NB2 --
        ("{D199E833-4E8C-0069-1D75-C494B8DA9CCD}{1}{E1910381033070156002921918971890669241658921}", 0, "TVOCs"),
        ("{D199E833-4E8C-0069-1D75-C494B8DA9CCD}{1}{E173758194090633205920166548046896713587831}", 0, "CO2"),
        # -- Check out later --
        ("{B3EE65E4-9E27-0930-1CB8-6D89BFDD60CC}{1}{E185379003394629284211974909958016347814671}", 0, "Check out later"),
    ]

    for pid, level, name in pages:
        print(f"  Moving: {name} (level {level})")
        move_page(pid, target, name, desired_level=level)


def move_ai_ml_pages(sections):
    """Move all AI/ML pages."""
    print("\n=== Moving pages to AI_and_ML ===")
    target = sections["AI_and_ML"]

    pages = [
        # -- AI Tools --
        ("{6CC1B568-8CCF-0183-3FBF-18C3141471FB}{1}{E185371734241348632901945991472171052265441}", 0, "ML sauce pot"),
        ("{6CC1B568-8CCF-0183-3FBF-18C3141471FB}{1}{E183361955586143319391919521165077621532361}", 0, "ollama"),
        ("{6CC1B568-8CCF-0183-3FBF-18C3141471FB}{1}{E1868745165831421110520152279867093863260291}", 0, "tgwebui"),
        ("{6CC1B568-8CCF-0183-3FBF-18C3141471FB}{1}{E1955796753504961464821996280045903575353631}", 0, "autogenstudio"),
        ("{6CC1B568-8CCF-0183-3FBF-18C3141471FB}{1}{E19532534696598181703220149344836839660404871}", 0, "Open-webui"),
        ("{6CC1B568-8CCF-0183-3FBF-18C3141471FB}{1}{E1948350389630180652501919002165212525423561}", 0, "Crawl4ai"),
        ("{6CC1B568-8CCF-0183-3FBF-18C3141471FB}{1}{E19574705676257832700920133546881935244430411}", 0, "Qdrant rag streamlit"),
        ("{6CC1B568-8CCF-0183-3FBF-18C3141471FB}{1}{E19529666381456342547620107462012134727614831}", 0, "Claw alts"),
        # -- Research --
        ("{6CC1B568-8CCF-0183-3FBF-18C3141471FB}{1}{E181361130236051569911945863547529764480591}", 0, "Text Paraphrasing"),
        ("{6CC1B568-8CCF-0183-3FBF-18C3141471FB}{1}{E1835280563213754992720179997499719649354451}", 0, "Simulating Network Infrastructure"),
        ("{6CC1B568-8CCF-0183-3FBF-18C3141471FB}{1}{E185308225487022332791972816982992624304861}", 0, "LLM video"),
        ("{6CC1B568-8CCF-0183-3FBF-18C3141471FB}{1}{E1895542086953747538820145203407634800766051}", 1, "Conclusons"),
        # -- Agent Development --
        ("{6CC1B568-8CCF-0183-3FBF-18C3141471FB}{1}{E1949277182929545660031922696027051865606431}", 0, "Desktop agent notes SAVI"),
        ("{6CC1B568-8CCF-0183-3FBF-18C3141471FB}{1}{E1836187145577735840020165605671605850365471}", 0, "Agent_zero"),
        ("{6CC1B568-8CCF-0183-3FBF-18C3141471FB}{1}{E1948802953302709727751914950520287465327871}", 0, "Anthropic computer use"),
        ("{6CC1B568-8CCF-0183-3FBF-18C3141471FB}{1}{E1890794207619137639720120468125129546908501}", 0, "jupyter"),
        # -- From Quick Notes --
        ("{07D75124-FAD0-08CE-1D52-DF726D7CBF59}{1}{E17313632831501294891941483934509922485921}", 0, "Openai, huggingface"),
        # -- From NB2 --
        ("{0A987DF9-D646-01C1-39B2-66D69708BB9B}{1}{E1876295605427336050820107773097586420001621}", 0, "Huggingface_stories"),
    ]

    for pid, level, name in pages:
        print(f"  Moving: {name} (level {level})")
        move_page(pid, target, name, desired_level=level)


def move_career_pages(sections):
    """Move all career/professional pages."""
    print("\n=== Moving pages to CAREER ===")
    target = sections["Career"]

    pages = [
        # -- Resume --
        ("{07D75124-FAD0-08CE-1D52-DF726D7CBF59}{1}{E1911342231325720253871934044888312034445201}", 0, "Resume 2021"),
        ("{1B4CE4AB-3229-0442-3BBF-E0F77189BB93}{1}{E1911079070952311410121948515566422199787861}", 0, "Things to incorporate into new resume"),
        # -- Timecards/HR --
        ("{07D75124-FAD0-08CE-1D52-DF726D7CBF59}{1}{E189359582595965252701991198471376433339101}", 0, "Deltek info"),
        ("{F31E127D-AB5D-4AE2-BBDC-2FBDC2487E73}{1}{E19552107519435738673820159149927735190846741}", 0, "Timecards"),
        ("{6CC1B568-8CCF-0183-3FBF-18C3141471FB}{1}{E19474799180886948685220175840217088990757121}", 0, "Caci_tc automation"),
        ("{6CC1B568-8CCF-0183-3FBF-18C3141471FB}{1}{E19576017711721594953914392026658853371}", 0, "Web automations"),
        ("{07D75124-FAD0-08CE-1D52-DF726D7CBF59}{1}{E19482154303857650625920112028160474124025601}", 0, "Health plans comparison"),
        # -- Tax/Clearance --
        ("{64F1B32F-E2EE-0472-2C8A-A6DDE39479AB}{1}{E1819054136888985590820101071331231818099331}", 0, "Tax doc Printout"),
        ("{64F1B32F-E2EE-0472-2C8A-A6DDE39479AB}{1}{E1947113770798827116021945451871812405160131}", 0, "2022 e-qip"),
        # -- CISSP class notes --
        ("{98294380-67E3-000F-115B-C2BE14C8138C}{1}{E187523138103156521291942153609308518640951}", 0, "CISSP Monday class notes"),
        ("{98294380-67E3-000F-115B-C2BE14C8138C}{1}{E1837426608551900614020130081276906865536411}", 1, "OSI notes"),
        ("{98294380-67E3-000F-115B-C2BE14C8138C}{1}{E1836844160701421253120137431418098733936841}", 1, "After lunch"),
        ("{98294380-67E3-000F-115B-C2BE14C8138C}{1}{E1875764606609357422120101978221052468091421}", 2, "Dan's drawing"),
        ("{98294380-67E3-000F-115B-C2BE14C8138C}{1}{E1848653861398407451220121875496213521710681}", 0, "CISSP Tuesday notes"),
        ("{98294380-67E3-000F-115B-C2BE14C8138C}{1}{E18300891490356590765182263836698997916551}", 1, "Access Controls Review"),
        ("{98294380-67E3-000F-115B-C2BE14C8138C}{1}{E18821605549840375523182810193972555433571}", 0, "CISSP Thursday"),
        # -- School/Work stuff --
        ("{33FF0178-B334-0D50-23A6-0628448988EC}{1}{E185917676103167233301938853719960236764631}", 0, "HP info"),
        ("{33FF0178-B334-0D50-23A6-0628448988EC}{1}{E186018741221175523261969972512775070863991}", 0, "Mesh V rails.docx"),
        ("{33FF0178-B334-0D50-23A6-0628448988EC}{1}{E152292484113004141935752174958739182231}", 1, "MeshV Page 2"),
        ("{33FF0178-B334-0D50-23A6-0628448988EC}{1}{E187288617336664651911995832841788388158001}", 1, "MeshV Page 3"),
        ("{33FF0178-B334-0D50-23A6-0628448988EC}{1}{E186072157736890663401957286349390658488771}", 1, "MeshV Page 4"),
        ("{33FF0178-B334-0D50-23A6-0628448988EC}{1}{E188170910604990175521938561991896176361501}", 1, "MeshV Page 5"),
        ("{33FF0178-B334-0D50-23A6-0628448988EC}{1}{E186095485517935382641951594314293554230081}", 1, "MeshV Page 6"),
        ("{33FF0178-B334-0D50-23A6-0628448988EC}{1}{E1812130888238626327120177667757639246118591}", 1, "MeshV Page 7"),
        ("{33FF0178-B334-0D50-23A6-0628448988EC}{1}{E1861799847971106098620154933204120044385771}", 1, "MeshV Page 8"),
        ("{33FF0178-B334-0D50-23A6-0628448988EC}{1}{E181683106620037635771928859704437909760071}", 1, "MeshV Page 9"),
        ("{33FF0178-B334-0D50-23A6-0628448988EC}{1}{E181049014503652597701915366725629574794681}", 1, "MeshV Page 10"),
        ("{33FF0178-B334-0D50-23A6-0628448988EC}{1}{E18197738685421440045185473601459068109321}", 1, "MeshV Page 11"),
        # -- Tech Notes --
        ("{A2BEDEEA-711F-01A1-0087-9BA4FBB5CFC0}{1}{E1911094909294738424081962325690754209072721}", 0, "Tools"),
        ("{A2BEDEEA-711F-01A1-0087-9BA4FBB5CFC0}{1}{E19108319346159901113720138831609736457142361}", 0, "remotes"),
        ("{A2BEDEEA-711F-01A1-0087-9BA4FBB5CFC0}{1}{E1893303408590635583620104232189758099066991}", 0, "File Transfers"),
        ("{A2BEDEEA-711F-01A1-0087-9BA4FBB5CFC0}{1}{E1866445162044403563120109035099279357153571}", 0, "Cisco tweaks"),
        ("{A2BEDEEA-711F-01A1-0087-9BA4FBB5CFC0}{1}{E19105941979420500189420116016611861985564331}", 0, "Updates"),
        # -- Flexible research group --
        ("{07D75124-FAD0-08CE-1D52-DF726D7CBF59}{1}{E181718427548806171911937392186848780854361}", 0, "Flexible research group"),
        # -- POSTS Prism (work-related writing) --
        ("{07D75124-FAD0-08CE-1D52-DF726D7CBF59}{1}{E1819541488343717341217103287929782653951}", 0, "POSTS Prism"),
    ]

    for pid, level, name in pages:
        print(f"  Moving: {name} (level {level})")
        move_page(pid, target, name, desired_level=level)


def move_personal_pages(sections):
    """Move all personal/family pages."""
    print("\n=== Moving pages to PERSONAL ===")
    target = sections["Personal"]

    pages = [
        # -- Family --
        ("{46775A51-35DA-07E0-3E1D-1285D8A128F9}{1}{E189068612947468550241989172056062680735271}", 0, "Email family kids eturnal"),
        ("{46775A51-35DA-07E0-3E1D-1285D8A128F9}{1}{E1949162663931486661681942467431096302240001}", 0, "Medical"),
        ("{076C1945-4B67-0D5B-1446-49C6FF8CA597}{1}{E1884036885828877793420156235304329250045901}", 0, "Christmas 2020"),
        # -- Projects --
        ("{1BAEDAFE-5925-0F6B-1D50-2314B23FC1E7}{1}{E1645614268662041081927179503372570197011}", 0, "Deborahs B-day stuff"),
        ("{1BAEDAFE-5925-0F6B-1D50-2314B23FC1E7}{1}{E1893392998888125139420150476063698281303171}", 0, "University of Maryland"),
        ("{1BAEDAFE-5925-0F6B-1D50-2314B23FC1E7}{1}{E182204304945679809231978363023230385564251}", 0, "Gerra's book"),
        ("{1BAEDAFE-5925-0F6B-1D50-2314B23FC1E7}{1}{E18549671144810481900182773719809002012291}", 1, "Picture pages"),
        ("{1BAEDAFE-5925-0F6B-1D50-2314B23FC1E7}{1}{E1852939916311350154520108999322016481119131}", 0, "Project Printout 1"),
        ("{1BAEDAFE-5925-0F6B-1D50-2314B23FC1E7}{1}{E19513584570399515891620182179587975309098731}", 0, "Project Printout 2"),
        # -- Hobbies --
        ("{784BEF3E-9A7F-093F-2B1A-4934D80FD170}{1}{E1834246472495603531220117024848450722798711}", 0, "3d printable tracks"),
        ("{F971193E-01E5-01C9-1AA3-293E2A33628E}{1}{E1892496304786916365220123553912094828037611}", 0, "BabyGame (Unity)"),
        # -- Conlangs/Creative --
        ("{07D75124-FAD0-08CE-1D52-DF726D7CBF59}{1}{E184311174936229969901928028915909645609521}", 0, "kar'tayl"),
        ("{07D75124-FAD0-08CE-1D52-DF726D7CBF59}{1}{E18146977075319703767187157877381269458221}", 0, "Lexica Nova"),
        ("{07D75124-FAD0-08CE-1D52-DF726D7CBF59}{1}{E183980446495383661921998950711220544875831}", 1, "arguments"),
        ("{07D75124-FAD0-08CE-1D52-DF726D7CBF59}{1}{E18219200770034768266184253217298353887131}", 2, "Clusters"),
        # -- Photos from NB2 --
        ("{B41F6927-32DB-0979-1B3E-0FD63D219603}{1}{E186184811793482694451911153928927098032681}", 0, "Photo collage 1"),
        ("{B41F6927-32DB-0979-1B3E-0FD63D219603}{1}{E182262429619831357611947297243931822367151}", 0, "Photo collage 2"),
        ("{B41F6927-32DB-0979-1B3E-0FD63D219603}{1}{E1827676991936837352020123443326998026875261}", 0, "Athena's cards"),
    ]

    for pid, level, name in pages:
        print(f"  Moving: {name} (level {level})")
        move_page(pid, target, name, desired_level=level)


def move_credential_pages(sections):
    """Move all credential/password pages."""
    print("\n=== Moving pages to CREDENTIALS ===")
    target = sections["Credentials"]

    pages = [
        ("{147DAA2D-89F9-040F-1A66-C7AC559435D6}{1}{E1910859640360632002471967106744402306442421}", 0, "Seldom used logons"),
        ("{147DAA2D-89F9-040F-1A66-C7AC559435D6}{1}{E187954753243005257141970931502636604186211}", 0, "P@$$words(creds)"),
        ("{07D75124-FAD0-08CE-1D52-DF726D7CBF59}{1}{E201252419291416015094920163115278357122911011}", 0, "Meta pass"),
        ("{07D75124-FAD0-08CE-1D52-DF726D7CBF59}{1}{E176103107027039149920100086302882944886001}", 0, "Voces magicae (Docker Hub creds)"),
    ]

    for pid, level, name in pages:
        print(f"  Moving: {name} (level {level})")
        move_page(pid, target, name, desired_level=level)


# =====================================================================
# Phase 4: Clean up empty source sections
# =====================================================================

def cleanup_sections():
    """Delete emptied source sections."""
    print("\n=== PHASE 4: Cleaning up empty sections ===")

    # Old underscore index sections from NB1
    old_sections = [
        ("{9B939CCB-C163-4957-BC87-CAB01BFDA076}{1}{B0}", "_Index"),
        ("{90E22DAF-D3AF-484E-B11E-0246DEAD9671}{1}{B0}", "_Networking"),
        ("{ED21DD7A-1B0C-4BBD-9187-33FB52A38CE8}{1}{B0}", "_Programming"),
        ("{BEB1D1A6-8432-470F-A679-7853C53C68D1}{1}{B0}", "_Cybertron_Infra"),
        ("{6AD896F9-37C7-421E-B4BB-1B8FE8DEBC0E}{1}{B0}", "_Career"),
    ]

    # Emptied NB1 sections
    emptied_nb1 = [
        ("{722F6090-9D24-0BA3-227F-A162FFE09B6A}{1}{B0}", "compusharecontact"),
        ("{07D75124-FAD0-08CE-1D52-DF726D7CBF59}{1}{B0}", "Quick Notes"),
        ("{1B4CE4AB-3229-0442-3BBF-E0F77189BB93}{1}{B0}", "resume notes"),
        ("{829A0E69-79D8-040D-1582-1D9D43EAD3C9}{1}{B0}", "pi"),
        ("{147DAA2D-89F9-040F-1A66-C7AC559435D6}{1}{B0}", "Seldom used"),
        ("{1BAEDAFE-5925-0F6B-1D50-2314B23FC1E7}{1}{B0}", "projects"),
        ("{64F1B32F-E2EE-0472-2C8A-A6DDE39479AB}{1}{B0}", "tax docs"),
        ("{F971193E-01E5-01C9-1AA3-293E2A33628E}{1}{B0}", "Unity Game Dev"),
        ("{3A7C5248-28F2-05A3-356F-F7162F2CC03B}{1}{B0}", "BeagleBone"),
        ("{076C1945-4B67-0D5B-1446-49C6FF8CA597}{1}{B0}", "Holiday Gifts"),
        ("{784BEF3E-9A7F-093F-2B1A-4934D80FD170}{1}{B0}", "robots"),
        ("{F882D5EB-DBA3-0E97-111D-AA1FBFD580BC}{1}{B0}", "cMNL Class"),
        ("{56AF71BA-FEFD-0322-086E-9136C3604C33}{1}{B0}", "Software"),
        ("{46775A51-35DA-07E0-3E1D-1285D8A128F9}{1}{B0}", "family"),
        ("{6CC1B568-8CCF-0183-3FBF-18C3141471FB}{1}{B0}", "AI_stuff"),
        ("{B3EE65E4-9E27-0930-1CB8-6D89BFDD60CC}{1}{B0}", "random notes"),
        ("{A2BEDEEA-711F-01A1-0087-9BA4FBB5CFC0}{1}{B0}", "Tech_Notes"),
        ("{1EE74EA4-2EE1-0D72-3387-DEEE0DBC7FD4}{1}{B0}", "From_thedose"),
        ("{56525385-122B-0E37-29C3-67F84037DB2B}{1}{B0}", "PAXPATRIS"),
        ("{F31E127D-AB5D-4AE2-BBDC-2FBDC2487E73}{1}{B0}", "october_2025"),
        ("{33FF0178-B334-0D50-23A6-0628448988EC}{1}{B0}", "School stuff"),
    ]

    # Emptied NB2 sections
    emptied_nb2 = [
        ("{B41F6927-32DB-0979-1B3E-0FD63D219603}{1}{B0}", "pictureCollages"),
        ("{AD0E851F-7621-04B6-29C7-B0C1FEEE1E77}{1}{B0}", "Misc_Tech_stuff"),
        ("{D199E833-4E8C-0069-1D75-C494B8DA9CCD}{1}{B0}", "AirQuality"),
        ("{33A07EDC-414C-015F-0CAC-1F00FA7F98FF}{1}{B0}", "infrastructure_tech"),
        ("{E97D0E54-CD21-0496-2EB5-E0A9E34F94C9}{1}{B0}", "network stuff"),
        ("{0A987DF9-D646-01C1-39B2-66D69708BB9B}{1}{B0}", "ML_Stories"),
        ("{420B032E-77FD-48CC-B294-B8EA099DDA20}{1}{B0}", "_Index (NB2)"),
    ]

    for sid, name in old_sections + emptied_nb1 + emptied_nb2:
        delete_section(sid, name)


# =====================================================================
# Phase 5: Create new master index
# =====================================================================

def create_master_index(sections):
    """Create a clean master index as first page of a new Index section."""
    import xml.etree.ElementTree as ET

    print("\n=== PHASE 5: Creating master index ===")

    H1 = 'style="font-size:18pt;font-weight:bold"'
    H2 = 'style="font-size:14pt;font-weight:bold"'
    GRAY = 'style="color:#888888"'
    BR = "<br>"
    BR2 = "<br><br>"
    IND = "&nbsp;&nbsp;&nbsp;&nbsp;"

    idx_section = com_client.open_hierarchy("_Index.one", NB1_ID)

    b = f'<span {H1}>Wylvek\'s Notebook -- Master Index</span>{BR}'
    b += f'<span {GRAY}><i>Reorganized by Megatron2 on 2026-03-01. Two notebooks consolidated into one.</i></span>{BR2}'

    b += f'<span {H2}>Networking</span> -- Cisco, Palo Alto, ForcePoint, UniFi, firewalls, VPN, protocols{BR}'
    b += f'{IND}35 pages: Cisco commands, PA packet flow, BFD/F5, UniFi controller, firewall rules, DCNM labs{BR2}'

    b += f'<span {H2}>Programming</span> -- Python, PowerShell, CMD, Docker, Linux, scripting tools{BR}'
    b += f'{IND}26 pages: Python reference, PS diagnostics, Docker cheatsheet, sed/pip/VLC/Visio{BR2}'

    b += f'<span {H2}>Cybertron</span> -- Home lab infrastructure, services, Raspberry Pi, Windows{BR}'
    b += f'{IND}28 pages: TrueNAS, HA, Vault, RustDesk, Git, OpenClaw, Pi projects, air quality{BR2}'

    b += f'<span {H2}>AI_and_ML</span> -- AI tools, frameworks, agents, research papers{BR}'
    b += f'{IND}18 pages: Ollama, AutoGen, Crawl4ai, Qdrant RAG, LLM research, agent dev{BR2}'

    b += f'<span {H2}>Career</span> -- Resume, certs, timecards, HR, work tools, school{BR}'
    b += f'{IND}35 pages: Resume, CISSP notes, Deltek, health plans, Tech Notes, School stuff{BR2}'

    b += f'<span {H2}>Personal</span> -- Family, gifts, hobbies, creative projects, photos{BR}'
    b += f'{IND}18 pages: Family contacts, medical, Christmas, Gerra\'s book, robots, conlangs, photos{BR2}'

    b += f'<span {H2}>Credentials</span> -- Passwords, API keys, login info{BR}'
    b += f'{IND}4 pages: Seldom used logons, passwords, Meta, Docker Hub{BR2}'

    b += f'<span {H2}>CISSP</span> <span {GRAY}>(archive -- 56-page course review PDF scan)</span>{BR}'
    b += f'{IND}Left in original section. Class notes moved to Career.{BR2}'

    b += f'<span {H2}>For Cybertron Agents</span>{BR}'
    b += f'{IND}Use <b>onenote_search</b> to find any topic across the notebook.{BR}'
    b += f'{IND}Networking, Programming, and Cybertron sections are your primary references.{BR}'
    b += f'{IND}Credentials section has API keys and passwords. Handle with care.{BR}'

    # Create the page
    page_id = com_client.create_new_page(idx_section)
    page_xml = com_client.get_page_content(page_id)
    root = ET.fromstring(page_xml)
    page_id_attr = root.get("ID")

    oe_match = re.search(r'<one:OE\s+([^>]+)>', page_xml)
    oe_raw = oe_match.group(1) if oe_match else ""
    oe_raw = re.sub(r'objectID="[^"]*"', '', oe_raw).strip()

    xml = (
        f'<?xml version="1.0" encoding="utf-8"?>'
        f'<one:Page xmlns:one="http://schemas.microsoft.com/office/onenote/2013/onenote" '
        f'ID="{page_id_attr}">'
        f'<one:Title lang="en-US">'
        f'<one:OE {oe_raw}>'
        f'<one:T><![CDATA[Master Index -- Table of Contents]]></one:T>'
        f'</one:OE></one:Title>'
        f'<one:Outline><one:OEChildren><one:OE>'
        f'<one:T><![CDATA[{b}]]></one:T>'
        f'</one:OE></one:OEChildren></one:Outline>'
        f'</one:Page>'
    )
    com_client.update_page_content(xml)
    print("  Master Index created")


# =====================================================================
# Main
# =====================================================================

if __name__ == "__main__":
    start = time.time()
    print("=" * 60)
    print("OneNote Notebook Reorganizer -- Megatron2")
    print("Consolidating 2 notebooks into 1 organized structure")
    print("Backups at Z:\\wgtho\\")
    print("=" * 60)

    # Phase 1: Create sections (idempotent - open_hierarchy returns existing)
    sections = create_target_sections()

    # Phase 2: Delete junk (safe to re-run - ignores already-deleted)
    delete_junk_pages()

    # Phase 3: Move all pages
    move_networking_pages(sections)
    move_programming_pages(sections)
    move_cybertron_pages(sections)
    move_ai_ml_pages(sections)
    move_career_pages(sections)
    move_personal_pages(sections)
    move_credential_pages(sections)

    # Phase 4: Clean up empty sections
    cleanup_sections()

    # Phase 5: Master index
    create_master_index(sections)

    elapsed = time.time() - start
    print()
    print("=" * 60)
    print(f"DONE in {elapsed:.0f}s ({elapsed/60:.1f}m)")
    print(f"  Moved:   {moved}")
    print(f"  Deleted: {deleted}")
    print(f"  Failed:  {failed}")
    print(f"  Skipped: {skipped} (CISSP PDF pages left in place)")
    print("=" * 60)
