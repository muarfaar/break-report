
import streamlit as st
import pandas as pd
import numpy as np
from datetime import date, timedelta, datetime
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import re

st.set_page_config(page_title="Break Compliance Report", page_icon="📊", layout="centered")

st.markdown("# 📊 Break Compliance Report")
st.markdown("Upload attendance CSV → Get formatted report instantly!")
st.markdown("---")

# --- Department Mapping (Updated Jul 29) ---
DEPT_MAP = {
    206827222: "ICQA", 206276761: "ICQA", 206277220: "ICQA",
    205226912: "ICQA", 206199903: "ICQA", 206277223: "ICQA",
    205005081: "ICQA", 202273043: "ICQA", 113115097: "ICQA",
    206276766: "ICQA", 205221423: "ICQA", 203953061: "ICQA",
    206287369: "ICQA", 206277230: "ICQA", 204942041: "ICQA",
    206910245: "ICQA", 103788917: "ICQA", 205278825: "ICQA",
    206297797: "ICQA", 204942695: "ICQA", 203427120: "ICQA",
    106295884: "ICQA", 205281248: "ICQA", 206276765: "ICQA",
    206276768: "ICQA", 109463379: "ICQA", 205252541: "ICQA",
    205279079: "ICQA", 106268210: "ICQA", 205548278: "Inbound",
    207450006: "Inbound", 112347709: "Inbound", 206276746: "Inbound",
    205278821: "Inbound", 206118440: "Inbound", 205271349: "Inbound",
    204868847: "Inbound", 206912050: "Inbound", 205588020: "Inbound",
    205996137: "Inbound", 206503613: "Inbound", 205588038: "Inbound",
    204256643: "Inbound", 205609141: "Inbound", 205275640: "Inbound",
    206276760: "Inbound", 112102326: "Inbound", 111009303: "Inbound",
    205609142: "Inbound", 207452565: "Inbound", 206117376: "Inbound",
    205588031: "Inbound", 204880561: "Inbound", 206889002: "Inbound",
    204950245: "Inbound", 205271607: "Inbound", 206871806: "Inbound",
    203816299: "Inbound", 205227221: "Inbound", 206906470: "Inbound",
    203755273: "Inbound", 206906709: "Inbound", 206914170: "Inbound",
    112874976: "Inbound", 206064517: "Inbound", 205548592: "Inbound",
    206193608: "Inbound", 206871345: "Inbound", 205256721: "Inbound",
    206239208: "Inbound", 105445261: "Inbound", 205252350: "Inbound",
    205807261: "Inbound", 206192966: "Inbound", 206809052: "Inbound",
    205627229: "Inbound", 205928017: "Inbound", 206502928: "Inbound",
    206490657: "Inbound", 206605144: "Inbound", 206873642: "Inbound",
    206230947: "Inbound", 206889484: "Inbound", 109468051: "Inbound",
    205199521: "Inbound", 203820467: "Inbound", 206912024: "Inbound",
    112079690: "Inbound", 102718104: "Inbound", 112874979: "Inbound",
    207442202: "Inbound", 205256966: "Inbound", 205271073: "Inbound",
    206827220: "Inbound", 203287189: "Inbound", 205635077: "Inbound",
    205555756: "Inbound", 205231165: "Inbound", 111144766: "Inbound",
    106762603: "Inbound", 205231163: "Inbound", 204947098: "Inbound",
    206827384: "Inbound", 206504147: "Inbound", 206200371: "Inbound",
    205275958: "Inbound", 206826456: "Inbound", 206871360: "Inbound",
    205227219: "Inbound", 206503623: "Inbound", 206826597: "Inbound",
    205231531: "Inbound", 205995592: "Inbound", 205555739: "Inbound",
    205939859: "Inbound", 113069850: "Inbound", 200714293: "Inbound",
    204277712: "Inbound", 204869093: "Inbound", 204869085: "Inbound",
    205807009: "Inbound", 206906249: "Inbound", 204886845: "Inbound",
    205198875: "Inbound", 205628512: "Inbound", 205548599: "Inbound",
    204879184: "Inbound", 206232510: "Inbound", 206906699: "Inbound",
    206906454: "Inbound", 204887302: "Inbound", 205555489: "Inbound",
    205807010: "Inbound", 205615608: "Inbound", 206826604: "Inbound",
    206193601: "Inbound", 205231536: "Inbound", 206503615: "Inbound",
    203816397: "Inbound", 205271611: "Inbound", 207451449: "Inbound",
    109748876: "Inbound", 205628513: "Inbound", 205616246: "Inbound",
    205555757: "Inbound", 206296197: "Inbound", 203306085: "Inbound",
    206906705: "Inbound", 206230570: "Inbound", 205938172: "Inbound",
    206875498: "Inbound", 206906703: "Inbound", 203554783: "Inbound",
    206822556: "Inbound", 206871651: "Inbound", 206517043: "Inbound",
    205985787: "Inbound", 205199825: "Inbound", 206827219: "Inbound",
    206503617: "Inbound", 205226151: "Inbound", 205635734: "Inbound",
    206872074: "Inbound", 206906460: "Inbound", 105443911: "Inbound",
    207452567: "Inbound", 201771653: "Inbound", 206286062: "Inbound",
    206871352: "Inbound", 206007532: "Inbound", 206277234: "Inbound",
    204967155: "Inbound", 112318363: "Inbound", 103202992: "Inbound",
    203305869: "Inbound", 206871650: "Inbound", 205588026: "Inbound",
    206201920: "Inbound", 206201922: "Inbound", 206277236: "Inbound",
    205252534: "Inbound", 112393372: "Inbound", 203850781: "Inbound",
    205231287: "Inbound", 205627612: "Inbound", 206014830: "Inbound",
    112599352: "Inbound", 206906251: "Inbound", 205928020: "Inbound",
    112920901: "Inbound", 206827231: "Inbound", 205555752: "Inbound",
    206889040: "Inbound", 206827054: "Inbound", 206873652: "Inbound",
    206873707: "Inbound", 206193598: "Inbound", 206874050: "Inbound",
    111144765: "Inbound", 205231417: "Inbound", 206277242: "Inbound",
    206201913: "Inbound", 205938170: "Inbound", 206277245: "Inbound",
    206910488: "Inbound", 206199836: "Inbound", 206200349: "Inbound",
    205588027: "Inbound", 206277248: "Inbound", 205227532: "Inbound",
    205252368: "Inbound", 206873564: "Inbound", 203498301: "Inbound",
    205226644: "Inbound", 206504139: "Inbound", 206906248: "Inbound",
    206906244: "Inbound", 206502929: "Inbound", 206889445: "Inbound",
    206889501: "Inbound", 206871520: "Inbound", 205252364: "Inbound",
    205271605: "Inbound", 112408798: "Inbound", 206827237: "Inbound",
    206232513: "Inbound", 206888996: "Inbound", 206906252: "Inbound",
    205615603: "Inbound", 206874040: "Inbound", 206201668: "Inbound",
    203066637: "Inbound", 206504134: "Inbound", 204992326: "Inbound",
    206199902: "Inbound", 203563362: "Inbound", 102207583: "Inbound",
    204886385: "Inbound", 206827232: "Inbound", 207442606: "Inbound",
    206889486: "Inbound", 205271625: "Inbound", 205231419: "Inbound",
    206889502: "Inbound", 203011027: "Inbound", 206827236: "Inbound",
    205807260: "Inbound", 206276754: "Inbound", 205996134: "Inbound",
    206873706: "Inbound", 206296198: "Inbound", 205252753: "Inbound",
    206889495: "Inbound", 204996615: "Inbound", 203679314: "Inbound",
    205252748: "Inbound", 205635349: "Inbound", 206827055: "Inbound",
    206889490: "Inbound", 206502396: "Inbound", 206502923: "Inbound",
    206906716: "Inbound", 204887316: "Inbound", 206906465: "Inbound",
    205252366: "Inbound", 205588036: "Inbound", 206809048: "Inbound",
    206889503: "Inbound", 203859973: "Inbound", 205937333: "Inbound",
    205548602: "Inbound", 112087286: "Inbound", 206906253: "Inbound",
    206906245: "Inbound", 205257288: "Inbound", 206826451: "Inbound",
    206277252: "Inbound", 206906707: "Inbound", 200176091: "Inbound",
    204950646: "Inbound", 205199353: "Inbound", 205199831: "Inbound",
    206873709: "Inbound", 203755271: "Inbound", 206906462: "Inbound",
    206873882: "Inbound", 206906466: "Inbound", 206117375: "Inbound",
    206906698: "Inbound", 205252755: "Inbound", 205588039: "Inbound",
    206503241: "Inbound", 112874987: "Inbound", 205252369: "Inbound",
    206912369: "Inbound", 206827234: "Inbound", 205627232: "Inbound",
    205986154: "Inbound", 203285597: "Inbound", 204887314: "Inbound",
    205271613: "Outbound", 205252359: "Outbound", 205252371: "Outbound",
    206912041: "Outbound", 206912021: "Outbound", 206277256: "Outbound",
    204890376: "Outbound", 206488959: "Outbound", 105454643: "Outbound",
    205010921: "Outbound", 205937794: "Outbound", 205252367: "Outbound",
    206475834: "Outbound", 206502395: "Outbound", 206910232: "Outbound",
    102718055: "Outbound", 206276756: "Outbound", 206912372: "Outbound",
    206914177: "Outbound", 205227537: "Outbound", 205252360: "Outbound",
    205195116: "Outbound", 206117706: "Outbound", 205987383: "Outbound",
    205252358: "Outbound", 206871649: "Outbound", 205257387: "Outbound",
    205258832: "Outbound", 204950655: "Outbound", 206953457: "Outbound",
    206199926: "Outbound", 105480601: "Outbound", 205252745: "Outbound",
    204010638: "Outbound", 205226915: "Outbound", 206200799: "Outbound",
    205937793: "Outbound", 206230572: "Outbound", 205202339: "Outbound",
    205226623: "Outbound", 204886573: "Outbound", 206200350: "Outbound",
    206326900: "Outbound", 205231538: "Outbound", 102207569: "Outbound",
    205195118: "Outbound", 206277782: "Outbound", 206871356: "Outbound",
    205199355: "Outbound", 206605147: "Outbound", 206871348: "Outbound",
    206605138: "Outbound", 206605143: "Outbound", 205231427: "Outbound",
    109407241: "Outbound", 206827238: "Outbound", 204887304: "Outbound",
    204264827: "Outbound", 103751774: "Outbound", 206200373: "Outbound",
    205198880: "Outbound", 206326892: "Outbound", 206361296: "Outbound",
    204890375: "Outbound", 206934192: "Outbound", 206240779: "Outbound",
    205201893: "Outbound", 205198879: "Outbound", 112911424: "Outbound",
    205226919: "Outbound", 204882818: "Outbound", 204946694: "Outbound",
    204946692: "Outbound", 206912360: "Outbound", 206910474: "Outbound",
    206200345: "Outbound", 205252365: "Outbound", 112409312: "Outbound",
    206200812: "Outbound", 204874830: "Outbound", 206118435: "Outbound",
    206128176: "Outbound", 206063872: "Outbound", 206912025: "Outbound",
    206361291: "Outbound", 206062910: "Outbound", 206278361: "Outbound",
    206489360: "Outbound", 203660434: "Outbound", 110163302: "Outbound",
    105444861: "Outbound", 206889509: "Outbound", 206605139: "Outbound",
    206193611: "Outbound", 206277228: "Outbound", 205199354: "Outbound",
    205252532: "Outbound", 206914171: "Outbound", 204009967: "Outbound",
    206871518: "Outbound", 106295861: "Outbound", 206193602: "Outbound",
    206949880: "Outbound", 206503240: "Outbound", 106970098: "Outbound",
    206889476: "Outbound", 205252540: "Outbound", 204950644: "Outbound",
    205615612: "Outbound", 206915176: "Outbound", 206606707: "Outbound",
    204966774: "Outbound", 206233183: "Outbound", 206200343: "Outbound",
    108282049: "Outbound", 206606702: "Outbound", 205226917: "Outbound",
    205977131: "Outbound", 206327186: "Outbound", 206193139: "Outbound",
    206327624: "Outbound", 205592629: "Outbound", 206326893: "Outbound",
    206503242: "Outbound", 206889265: "Outbound", 205199828: "Outbound",
    206604513: "Outbound", 205227528: "Outbound", 206232730: "Outbound",
    102207594: "Outbound", 206242016: "Outbound", 205633860: "Outbound",
    206277227: "Outbound", 104307217: "Outbound", 206326919: "Outbound",
    206287660: "Outbound", 206606012: "Outbound", 206809049: "Outbound",
    206326922: "Outbound", 205257129: "Outbound", 203850670: "Outbound",
    206287663: "Outbound", 206874043: "Outbound", 206808897: "Outbound",
    206277226: "Outbound", 206231953: "Outbound", 206606245: "Outbound",
    206906711: "Outbound", 205252752: "Outbound", 205227527: "Outbound",
    204951298: "Outbound", 206604511: "Outbound", 206192842: "Outbound",
    205600854: "Outbound", 206277229: "Outbound", 204873067: "Outbound",
    206237382: "Outbound", 206063908: "Outbound", 206604500: "Outbound",
    205922854: "Outbound", 105444811: "Outbound", 206266031: "Outbound",
    206201506: "Outbound", 103751763: "Outbound", 204890575: "Outbound",
    205227530: "Outbound", 206193612: "Outbound", 112347684: "Outbound",
    206287373: "Outbound", 204891069: "Outbound", 205627226: "Outbound",
    205615606: "Outbound", 206912362: "Outbound", 206277217: "Outbound",
    104277451: "Outbound", 206914382: "Outbound", 206490651: "Outbound",
    205231420: "Outbound", 204950639: "Outbound", 206607046: "Outbound",
    206889491: "Outbound", 206490644: "Outbound", 206200467: "Outbound",
    206489359: "Outbound", 205278818: "Outbound", 205231166: "Outbound",
    105443956: "Outbound", 205231031: "Outbound", 205230944: "Outbound",
    205231033: "Outbound", 205231422: "Outbound", 205226924: "Outbound",
    205221561: "Outbound", 206606030: "Outbound", 205940154: "Outbound",
    205231030: "Outbound", 205278827: "Outbound", 206117700: "Outbound",
    206277246: "Outbound", 205939338: "Outbound", 206201914: "Outbound",
    206576240: "Outbound", 206889507: "Outbound", 206606246: "Outbound",
    206490655: "Outbound", 206201475: "Outbound", 205939341: "Outbound",
    205221562: "Outbound", 205231032: "Outbound", 205227538: "Outbound",
    205221565: "Outbound", 206277238: "Outbound", 205592632: "Outbound",
    205231530: "Outbound", 206504140: "Outbound", 206910494: "Outbound",
    205940163: "Outbound", 206934634: "Outbound", 206490648: "Outbound",
    110163345: "Outbound", 206118442: "Outbound", 206490646: "Outbound",
    204886247: "Outbound", 205938665: "Outbound", 102207596: "Outbound",
    204887646: "Outbound", 207337286: "Outbound", 204868855: "Outbound",
    206607045: "Outbound", 206361304: "Outbound", 205231537: "Outbound",
    206276757: "Outbound", 103989125: "Outbound", 204967050: "Outbound",
    206490653: "Outbound", 206232736: "Outbound", 205252751: "Outbound",
    102207658: "Outbound", 205586502: "Outbound", 206892313: "Outbound",
    206199999: "Outbound", 206231958: "Outbound", 204946683: "Outbound",
    205591904: "Outbound", 206327192: "Outbound", 206276764: "Outbound",
    206912048: "Outbound", 206874048: "Outbound", 205929364: "Outbound",
    206889487: "Outbound", 205199826: "Outbound", 205593691: "Outbound",
    206276762: "Outbound", 204946687: "Outbound", 207338414: "Outbound",
    205939281: "Outbound", 206809054: "Outbound", 205199829: "Outbound",
    205257283: "Outbound", 205252355: "Outbound", 206361294: "Outbound",
    206276770: "Outbound", 204886245: "Outbound", 204878040: "Outbound",
    205231292: "Outbound", 206606714: "Outbound", 206276773: "Outbound",
    206277254: "Outbound", 205256723: "Outbound", 205949501: "Outbound",
    206912065: "Outbound", 205592630: "Outbound", 205628536: "Outbound",
    204352636: "Outbound", 206237363: "Outbound", 206193613: "Outbound",
    205252348: "Warehouse Deals", 201645184: "Warehouse Deals", 206326895: "Warehouse Deals",
    205325442: "Warehouse Deals", 206827042: "Warehouse Deals", 204880757: "Warehouse Deals",
    204319795: "Warehouse Deals", 205257635: "Warehouse Deals", 206128190: "Warehouse Deals",
    205928925: "Warehouse Deals", 206912028: "Warehouse Deals", 206361289: "Warehouse Deals",
    205278824: "Warehouse Deals", 203660528: "Warehouse Deals", 206889047: "Warehouse Deals",
    206822554: "Warehouse Deals", 205257290: "Warehouse Deals", 205199113: "Warehouse Deals",
    102981764: "Warehouse Deals", 206906713: "Warehouse Deals", 206502927: "Warehouse Deals",
    204319801: "Warehouse Deals", 205937796: "Warehouse Deals", 203850805: "Warehouse Deals",
    206912036: "Warehouse Deals", 102582254: "Warehouse Deals", 102207600: "Warehouse Deals",
    206827224: "Warehouse Deals", 205206430: "Warehouse Deals", 113175683: "Warehouse Deals",
    206128186: "Warehouse Deals", 205280642: "Warehouse Deals", 205231293: "Warehouse Deals",
    205199524: "Warehouse Deals", 206327199: "Warehouse Deals", 205256726: "Warehouse Deals",
    206822072: "Warehouse Deals", 205199824: "Warehouse Deals", 204967042: "Warehouse Deals",
    206827229: "Warehouse Deals", 206910243: "Warehouse Deals", 206827391: "Warehouse Deals",
    206504127: "Warehouse Deals", 203816401: "Warehouse Deals", 206934440: "Warehouse Deals",
    204373627: "Warehouse Deals", 102205930: "Warehouse Deals", 203660393: "Warehouse Deals",
    206361301: "Warehouse Deals", 206284692: "Warehouse Deals", 206934610: "Warehouse Deals",
    206277247: "Warehouse Deals", 205271612: "Warehouse Deals", 205243458: "Warehouse Deals",
    204319465: "Warehouse Deals", 206278356: "Warehouse Deals", 205252361: "Warehouse Deals",
    106295659: "Warehouse Deals", 102205652: "Warehouse Deals", 102205991: "Warehouse Deals",
    206128178: "Warehouse Deals", 112728750: "Warehouse Deals", 201179452: "Warehouse Deals",
    206128188: "Warehouse Deals", 206327621: "Warehouse Deals", 204950874: "Warehouse Deals",
    113249574: "Warehouse Deals", 205231428: "Warehouse Deals", 205252531: "Warehouse Deals",
    113142038: "Warehouse Deals", 205252362: "Warehouse Deals", 205247342: "HR",
    107322190: "HR", 203826463: "HR", 200519423: "HR",
    205930604: "HR", 207302255: "HR", 206018669: "HR",
    205181646: "HR", 112806642: "HR", 207758631: "HR",
    207407185: "HR", 203741549: "Inbound", 205048424: "IT",
    204976985: "L&D", 205221421: "L&D", 205109148: "L&D",
    204352950: "L&D", 204821683: "L&D", 112181825: "L&D",
    205231288: "L&D", 200104984: "L&D", 205781111: "L&D",
    112235562: "L&D", 205005101: "Procurement", 204985884: "RME",
    205001748: "RME", 205004710: "RME", 204821681: "RME",
    205897595: "RME", 207245175: "RME", 102582347: "Safety",
    113135599: "Safety", 203724002: "Outbound", 203724372: "Outbound",
    206858277: "Outbound", 206858181: "Outbound", 206906084: "Outbound",
    202218639: "Outbound", 203331869: "Outbound", 203305872: "Outbound",
    203331895: "Outbound", 202178813: "Outbound", 201237965: "Outbound",
    203305868: "Outbound", 203859923: "Outbound", 205226654: "Outbound",
    203859971: "Outbound", 206908890: "RME", 203252168: "Outbound",
    203253434: "Outbound", 203305923: "Outbound", 203785491: "Outbound",
    203305996: "Outbound", 203341075: "Outbound", 203331891: "Outbound",
    203331868: "Outbound", 203285902: "Outbound", 203468105: "Outbound",
    203285645: "Outbound", 203285600: "Outbound", 203427537: "Outbound",
    203253841: "Outbound", 202973003: "Outbound", 203252084: "Outbound",
}

# --- Colors (matching desktop report) ---
SQUID_INK = '232F3E'
TEAL = '00BCD4'
CORAL = 'FF6B6B'
SUNSET = 'FFA726'
CHARCOAL = '424242'
SNOW = 'FAFAFA'
ICE_BLUE = 'E0F7FA'
LIGHT_CORAL = 'FFEBEE'
LIGHT_SUNSET = 'FFF3E0'
WHITE = 'FFFFFF'
DARK_TEXT = '212121'
REPEAT_RED = 'D32F2F'
REPEAT_BG = 'FFCDD2'

# --- Helper: Extract date from filename ---
def extract_date_from_filename(filename):
    match = re.search(r'(\d{8})\d{4}-\d{12}', filename)
    if match:
        date_str = match.group(1)
        dt = datetime.strptime(date_str, '%Y%m%d')
        return dt.strftime('%A, %d %B %Y')
    
    match = re.search(r'(\d{4})(\d{2})(\d{2})', filename)
    if match:
        year, month, day = match.groups()
        dt = datetime(int(year), int(month), int(day))
        return dt.strftime('%A, %d %B %Y')
    
    match = re.search(r'(\d{4})-(\d{2})-(\d{2})', filename)
    if match:
        year, month, day = match.groups()
        dt = datetime(int(year), int(month), int(day))
        return dt.strftime('%A, %d %B %Y')
    
    yesterday = date.today() - timedelta(days=1)
    return yesterday.strftime('%A, %d %B %Y')

# --- Helper: Extract short date for history ---
def extract_short_date(filename):
    match = re.search(r'(\d{8})\d{4}-\d{12}', filename)
    if match:
        date_str = match.group(1)
        dt = datetime.strptime(date_str, '%Y%m%d')
        return f"{dt.month}/{dt.day}/{dt.year}"
    
    match = re.search(r'(\d{4})(\d{2})(\d{2})', filename)
    if match:
        year, month, day = match.groups()
        dt = datetime(int(year), int(month), int(day))
        return f"{dt.month}/{dt.day}/{dt.year}"
    
    match = re.search(r'(\d{4})-(\d{2})-(\d{2})', filename)
    if match:
        year, month, day = match.groups()
        dt = datetime(int(year), int(month), int(day))
        return f"{dt.month}/{dt.day}/{dt.year}"
    
    yesterday = date.today() - timedelta(days=1)
    return f"{yesterday.month}/{yesterday.day}/{yesterday.year}"

# --- Helper: Auto-detect shift from filename ---
def detect_shift_from_filename(filename):
    match = re.search(r'\d{8}(\d{4})-\d{12}', filename)
    if match:
        start_time = match.group(1)
        hour = int(start_time[:2])
        if hour >= 14:
            return 1  # Night Shift
        else:
            return 0  # Day Shift
    return 0

# --- Upload History ---
st.markdown("### 📋 Upload History (optional)")
history_file = st.file_uploader("Upload previous history.csv for repeat tracking", type=['csv'], key="history", help="Max 200MB per file")

history_df = pd.DataFrame(columns=['Employee ID', 'Employee Name', 'Date', 'Shift', 'Flag'])
if history_file:
    history_df = pd.read_csv(history_file)
    st.success(f"History loaded: {len(history_df)} past records")

# --- Upload Attendance CSV ---
st.markdown("### 📊 Upload Attendance CSV")
uploaded_file = st.file_uploader("Upload your attendance CSV file", type=['csv'], key="attendance", help="Max 200MB per file")

if uploaded_file:
    df = pd.read_csv(uploaded_file)
    filename = uploaded_file.name
    
    # Extract date from filename
    report_date = extract_date_from_filename(filename)
    short_date = extract_short_date(filename)
    
    # Auto-detect shift with override option
    default_shift = detect_shift_from_filename(filename)
    shift_type = st.selectbox("Select Shift:", ["Day Shift", "Night Shift"], index=default_shift)
    
    st.info(f"📅 Date: **{report_date}** | 🔄 Shift: **{shift_type}**")
    
    # --- Detect format ---
    if 'Punch Time' in df.columns:
        st.info("Detected: **FCLM Raw Punch Data** — calculating breaks from timestamps...")
        
        df['Punch Time'] = pd.to_datetime(df['Punch Time'])
        
        results = []
        for emp_id, group in df.groupby('Employee ID'):
            punches = group.sort_values('Punch Time')
            emp_name = punches['Employee Name'].iloc[0]
            
            if len(punches) <= 2:
                continue
            
            punch_times = punches['Punch Time'].tolist()
            total_break = 0
            
            for i in range(1, len(punch_times) - 1, 2):
                break_mins = (punch_times[i+1] - punch_times[i]).total_seconds() / 60
                if 0 < break_mins < 120:
                    total_break += break_mins
            
            if total_break == 0 and len(punch_times) >= 4:
                break_mins = (punch_times[2] - punch_times[1]).total_seconds() / 60
                if 0 < break_mins < 120:
                    total_break = break_mins
            
            if total_break > 0:
                results.append({
                    'Employee ID': emp_id,
                    'Employee Name': emp_name,
                    'Break (min)': round(total_break)
                })
        
        processed_df = pd.DataFrame(results)
    
    elif 'First Half Duration' in df.columns or 'Total Duration' in df.columns:
        st.info("Detected: **Dashboard Export** — reading break totals...")
        
        if 'First Half Duration' in df.columns and 'Second Half Duration' in df.columns:
            df['Break (min)'] = df['First Half Duration'].fillna(0) + df['Second Half Duration'].fillna(0)
        elif 'Total Duration' in df.columns:
            df['Break (min)'] = df['Total Duration'].fillna(0)
        
        processed_df = df[['Employee ID', 'Employee Name', 'Break (min)']].copy()
        processed_df = processed_df[processed_df['Break (min)'] > 0]
    
    else:
        st.error("❌ Unrecognized CSV format. Please upload FCLM Raw Punch Data or Dashboard Export.")
        st.stop()
    
    if len(processed_df) == 0:
        st.warning("No break data found to process.")
        st.stop()
    
    st.success(f"Processed **{len(processed_df)}** employees!")
    
    # --- Add Department ---
    processed_df['Department'] = processed_df['Employee ID'].map(DEPT_MAP).fillna('Unknown')
    
    # --- Flag Excess and Less ---
    excess_df = processed_df[processed_df['Break (min)'] >= 65].sort_values('Break (min)', ascending=False).reset_index(drop=True)
    less_df = processed_df[processed_df['Break (min)'] <= 55].sort_values('Break (min)', ascending=True).reset_index(drop=True)
    
    # --- Count repeats from history ---
    def count_repeats(emp_id, flag):
        if history_df.empty:
            return 0
        matches = history_df[(history_df['Employee ID'] == emp_id) & (history_df['Flag'] == flag)]
        return len(matches)
    
    excess_df['Repeat'] = excess_df['Employee ID'].apply(lambda x: count_repeats(x, 'Excess'))
    less_df['Repeat'] = less_df['Employee ID'].apply(lambda x: count_repeats(x, 'Less'))
    
    # Format repeat column with ⚠️ triangle
    excess_df['Repeat'] = excess_df['Repeat'].apply(lambda x: f"⚠️ {x}x" if x > 0 else "")
    less_df['Repeat'] = less_df['Repeat'].apply(lambda x: f"⚠️ {x}x" if x > 0 else "")
    
    # --- Display Metrics ---
    total = len(excess_df) + len(less_df)
    col1, col2, col3 = st.columns(3)
    col1.metric("Total Exceptions", total)
    col2.metric("Excess (≥65 min)", len(excess_df))
    col3.metric("Less (≤55 min)", len(less_df))
    
    # --- Display Tables ---
    if len(excess_df) > 0:
        st.markdown("### 🔴 Excess Break (≥65 min)")
        st.dataframe(excess_df[['Employee ID', 'Employee Name', 'Department', 'Break (min)', 'Repeat']], use_container_width=True)
    
    if len(less_df) > 0:
        st.markdown("### 🟡 Less Break (≤55 min)")
        st.dataframe(less_df[['Employee ID', 'Employee Name', 'Department', 'Break (min)', 'Repeat']], use_container_width=True)
    
    # --- Build Updated History ---
    new_history_records = []
    for _, row in excess_df.iterrows():
        new_history_records.append({
            'Employee ID': row['Employee ID'],
            'Employee Name': row['Employee Name'],
            'Date': short_date,
            'Shift': shift_type,
            'Flag': 'Excess'
        })
    for _, row in less_df.iterrows():
        new_history_records.append({
            'Employee ID': row['Employee ID'],
            'Employee Name': row['Employee Name'],
            'Date': short_date,
            'Shift': shift_type,
            'Flag': 'Less'
        })
    
    new_history_df = pd.DataFrame(new_history_records)
    updated_history = pd.concat([history_df, new_history_df], ignore_index=True)
    
    # --- Generate Excel Report (Matching Desktop Style) ---
    def generate_excel(excess, less, report_date, shift_type):
        wb = Workbook()
        ws = wb.active
        ws.title = "Break Compliance"
        ws.sheet_view.showGridLines = False
        
        # Column widths
        ws.column_dimensions['A'].width = 14
        ws.column_dimensions['B'].width = 28
        ws.column_dimensions['C'].width = 14
        ws.column_dimensions['D'].width = 8
        ws.column_dimensions['E'].width = 10
        
        # === HEADER BANNER ===
        for c in range(1, 6):
            ws.cell(row=1, column=c).fill = PatternFill(start_color=SQUID_INK, end_color=SQUID_INK, fill_type='solid')
        ws.row_dimensions[1].height = 10
        
        ws.row_dimensions[2].height = 35
        ws.merge_cells('A2:E2')
        ws['A2'] = "BREAK COMPLIANCE REPORT"
        ws['A2'].font = Font(name='Calibri', size=16, bold=True, color=TEAL)
        ws['A2'].fill = PatternFill(start_color=SQUID_INK, end_color=SQUID_INK, fill_type='solid')
        ws['A2'].alignment = Alignment(vertical='center', horizontal='center')
        for c in range(1, 6):
            ws.cell(row=2, column=c).fill = PatternFill(start_color=SQUID_INK, end_color=SQUID_INK, fill_type='solid')
        
        # Teal accent line
        ws.row_dimensions[3].height = 4
        for c in range(1, 6):
            ws.cell(row=3, column=c).fill = PatternFill(start_color=TEAL, end_color=TEAL, fill_type='solid')
        
        # Date row
        ws.row_dimensions[4].height = 22
        ws.merge_cells('A4:E4')
        ws['A4'] = report_date
        ws['A4'].font = Font(name='Calibri', size=10, color='666666')
        ws['A4'].fill = PatternFill(start_color=SNOW, end_color=SNOW, fill_type='solid')
        ws['A4'].alignment = Alignment(vertical='center', horizontal='center')
        for c in range(1, 6):
            ws.cell(row=4, column=c).fill = PatternFill(start_color=SNOW, end_color=SNOW, fill_type='solid')
        
        ws.row_dimensions[5].height = 8
        
        # === SHIFT HEADER ===
        row = 6
        shift_time = "08:00 - 18:00" if shift_type == "Day Shift" else "20:15 - 04:15"
        ws.merge_cells(f'A{row}:E{row}')
        ws[f'A{row}'] = f"{shift_type.upper()}  |  {shift_time}"
        ws[f'A{row}'].font = Font(name='Calibri', size=12, bold=True, color=WHITE)
        ws[f'A{row}'].fill = PatternFill(start_color=SQUID_INK, end_color=SQUID_INK, fill_type='solid')
        ws[f'A{row}'].alignment = Alignment(vertical='center', horizontal='center')
        for c in range(1, 6):
            ws.cell(row=row, column=c).fill = PatternFill(start_color=SQUID_INK, end_color=SQUID_INK, fill_type='solid')
        ws.row_dimensions[row].height = 25
        row += 1
        
        # Metrics row
        ws.row_dimensions[row].height = 24
        total_exceptions = len(excess) + len(less)
        ws[f'A{row}'] = f"Exceptions: {total_exceptions}"
        ws[f'A{row}'].font = Font(name='Calibri', size=9, bold=True, color=SQUID_INK)
        ws[f'A{row}'].fill = PatternFill(start_color=ICE_BLUE, end_color=ICE_BLUE, fill_type='solid')
        ws[f'B{row}'] = f"Excess: {len(excess)}  |  Less: {len(less)}"
        ws[f'B{row}'].font = Font(name='Calibri', size=9, color='666666')
        ws[f'B{row}'].fill = PatternFill(start_color=ICE_BLUE, end_color=ICE_BLUE, fill_type='solid')
        for c in range(3, 6):
            ws.cell(row=row, column=c).fill = PatternFill(start_color=ICE_BLUE, end_color=ICE_BLUE, fill_type='solid')
        row += 1
        
        # Spacer
        ws.row_dimensions[row].height = 6
        row += 1
        
        # === EXCESS TABLE ===
        ws.merge_cells(f'A{row}:E{row}')
        ws[f'A{row}'] = "EXCESS BREAK  ≥65 min"
        ws[f'A{row}'].font = Font(name='Calibri', size=10, bold=True, color=WHITE)
        ws[f'A{row}'].fill = PatternFill(start_color=CORAL, end_color=CORAL, fill_type='solid')
        ws[f'A{row}'].alignment = Alignment(vertical='center')
        for c in range(1, 6):
            ws.cell(row=row, column=c).fill = PatternFill(start_color=CORAL, end_color=CORAL, fill_type='solid')
        ws.row_dimensions[row].height = 20
        row += 1
        
        # Column headers
        headers = ['Employee ID', 'Name', 'Department', 'Mins', 'Repeat']
        for col_idx, header in enumerate(headers, 1):
            cell = ws.cell(row=row, column=col_idx, value=header)
            cell.font = Font(name='Calibri', size=8, bold=True, color=SQUID_INK)
            cell.fill = PatternFill(start_color=LIGHT_CORAL, end_color=LIGHT_CORAL, fill_type='solid')
        ws.row_dimensions[row].height = 16
        row += 1
        
        if len(excess) == 0:
            ws[f'A{row}'] = "No exceptions"
            ws[f'A{row}'].font = Font(name='Calibri', size=9, italic=True, color='AAAAAA')
            row += 1
        else:
            for i, (_, r) in enumerate(excess.iterrows()):
                ws.row_dimensions[row].height = 17
                is_repeat = r['Repeat'] != ""
                
                if is_repeat:
                    for c in range(1, 6):
                        ws.cell(row=row, column=c).fill = PatternFill(start_color=REPEAT_BG, end_color=REPEAT_BG, fill_type='solid')
                elif i % 2 == 0:
                    for c in range(1, 6):
                        ws.cell(row=row, column=c).fill = PatternFill(start_color=LIGHT_CORAL, end_color=LIGHT_CORAL, fill_type='solid')
                
                ws.cell(row=row, column=1, value=r['Employee ID']).font = Font(name='Calibri', size=9, color=DARK_TEXT)
                ws.cell(row=row, column=2, value=r['Employee Name']).font = Font(name='Calibri', size=9, color=DARK_TEXT)
                ws.cell(row=row, column=3, value=r['Department']).font = Font(name='Calibri', size=9, color=DARK_TEXT)
                ws.cell(row=row, column=4, value=r['Break (min)']).font = Font(name='Calibri', size=9, bold=True, color=CORAL)
                
                if is_repeat:
                    ws.cell(row=row, column=5, value=r['Repeat']).font = Font(name='Calibri', size=9, bold=True, color=REPEAT_RED)
                else:
                    ws.cell(row=row, column=5, value="").font = Font(name='Calibri', size=9)
                
                row += 1
        
        # Spacer
        ws.row_dimensions[row].height = 6
        row += 1
        
        # === LESS TABLE ===
        ws.merge_cells(f'A{row}:E{row}')
        ws[f'A{row}'] = "LESS BREAK  ≤55 min"
        ws[f'A{row}'].font = Font(name='Calibri', size=10, bold=True, color=WHITE)
        ws[f'A{row}'].fill = PatternFill(start_color=SUNSET, end_color=SUNSET, fill_type='solid')
        ws[f'A{row}'].alignment = Alignment(vertical='center')
        for c in range(1, 6):
            ws.cell(row=row, column=c).fill = PatternFill(start_color=SUNSET, end_color=SUNSET, fill_type='solid')
        ws.row_dimensions[row].height = 20
        row += 1
        
        # Column headers
        for col_idx, header in enumerate(headers, 1):
            cell = ws.cell(row=row, column=col_idx, value=header)
            cell.font = Font(name='Calibri', size=8, bold=True, color=SQUID_INK)
            cell.fill = PatternFill(start_color=LIGHT_SUNSET, end_color=LIGHT_SUNSET, fill_type='solid')
        ws.row_dimensions[row].height = 16
        row += 1
        
        if len(less) == 0:
            ws[f'A{row}'] = "No exceptions"
            ws[f'A{row}'].font = Font(name='Calibri', size=9, italic=True, color='AAAAAA')
            row += 1
        else:
            for i, (_, r) in enumerate(less.iterrows()):
                ws.row_dimensions[row].height = 17
                is_repeat = r['Repeat'] != ""
                
                if is_repeat:
                    for c in range(1, 6):
                        ws.cell(row=row, column=c).fill = PatternFill(start_color=REPEAT_BG, end_color=REPEAT_BG, fill_type='solid')
                elif i % 2 == 0:
                    for c in range(1, 6):
                        ws.cell(row=row, column=c).fill = PatternFill(start_color=LIGHT_SUNSET, end_color=LIGHT_SUNSET, fill_type='solid')
                
                ws.cell(row=row, column=1, value=r['Employee ID']).font = Font(name='Calibri', size=9, color=DARK_TEXT)
                ws.cell(row=row, column=2, value=r['Employee Name']).font = Font(name='Calibri', size=9, color=DARK_TEXT)
                ws.cell(row=row, column=3, value=r['Department']).font = Font(name='Calibri', size=9, color=DARK_TEXT)
                ws.cell(row=row, column=4, value=r['Break (min)']).font = Font(name='Calibri', size=9, bold=True, color=SUNSET)
                
                if is_repeat:
                    ws.cell(row=row, column=5, value=r['Repeat']).font = Font(name='Calibri', size=9, bold=True, color=REPEAT_RED)
                else:
                    ws.cell(row=row, column=5, value="").font = Font(name='Calibri', size=9)
                
                row += 1
        
        # Footer line
        for c in range(1, 6):
            ws.cell(row=row, column=c).border = Border(top=Side(style='medium', color=TEAL))
        
        buffer = BytesIO()
        wb.save(buffer)
        buffer.seek(0)
        return buffer
    
    # --- Downloads ---
    st.markdown("---")
    st.markdown("### 📥 Downloads")
    
    col1, col2 = st.columns(2)
    
    with col1:
        excel_buffer = generate_excel(excess_df, less_df, report_date, shift_type)
        st.download_button(
            label="📊 Excel Report",
            data=excel_buffer,
            file_name=f"Break_Report_{short_date.replace('/', '-')}_{shift_type.replace(' ', '_')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.document"
        )
    
    with col2:
        history_csv = updated_history.to_csv(index=False)
        st.download_button(
            label="📋 Updated History",
            data=history_csv,
            file_name="history.csv",
            mime="text/csv"
        )
    
    st.markdown("---")
    st.markdown("💡 **Download the updated history.csv each time and upload it next session!**")

# --- Criteria ---
st.markdown("---")
st.markdown("### ℹ️ How to use:")
st.markdown("1. (Optional) Upload previous history.csv for repeat tracking")
st.markdown("2. Upload your attendance CSV (supports both Dashboard export & FCLM raw data)")
st.markdown("3. View results + download report & updated history!")
st.markdown("### 📐 Criteria:")
st.markdown("- **Excess** = ≥65 min | **Less** = ≤55 min")
st.markdown("- Repeat offenders highlighted in red with count")



