
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

# --- Department Mapping (Updated 10 Aug 2026 from DXB5 HC) ---
DEPT_MAP = {
    205024567: "ACES", 111018307: "ACES",
    106268180: "DUF7", 102718073: "DUF7", 113139328: "DUF7", 104307248: "DUF7",
    103055719: "DUF7", 207255669: "DUF7",
    109468014: "DUF8", 106970090: "DUF8", 105443916: "DUF8", 206997081: "DUF8",
    102207639: "DUF8", 113047036: "DUF8", 102577303: "DUF8", 200511225: "DUF8",
    207732118: "DUF8",
    102981763: "DXB6", 205362762: "DXB6", 203851333: "DXB6", 110984626: "DXB6",
    102207646: "DXB6", 202918221: "DXB6",
    109467993: "DXF2", 207689734: "DXF2",
    206423274: "Data Analyst",
    206659449: "Facility", 107924676: "Facility", 204671494: "Facility", 206756201: "Facility",
    207167640: "Facility", 206757798: "Facility", 113014969: "Facility",
    205671526: "HR", 205346979: "HR", 203365316: "HR", 112978301: "HR",
    207799093: "HR", 205688378: "HR", 107798716: "HR", 205730381: "HR",
    206174751: "HR", 207283256: "HR", 204604983: "HR", 204604982: "HR",
    205776070: "HR", 203305831: "HR", 203331748: "HR", 205441633: "HR",
    207275154: "HR", 205389687: "HR", 205699662: "HR", 205536271: "HR",
    105445264: "HR", 200641956: "HR",
    204228003: "ICQA", 203859926: "ICQA", 203427119: "ICQA", 203820470: "ICQA",
    204228002: "ICQA", 206953809: "ICQA", 206953793: "ICQA", 112812207: "ICQA",
    204874916: "ICQA", 204853767: "ICQA", 204286569: "ICQA", 206514868: "ICQA",
    206491324: "ICQA", 206214074: "ICQA", 113139330: "ICQA", 204276463: "ICQA",
    206957729: "ICQA", 206957742: "ICQA", 206957732: "ICQA", 206214095: "ICQA",
    102206157: "ICQA", 201635614: "ICQA", 204227753: "ICQA", 206957726: "ICQA",
    206957727: "ICQA", 206214080: "ICQA", 206214143: "ICQA", 112806229: "ICQA",
    204228379: "ICQA", 206933948: "ICQA", 202103533: "ICQA", 206214365: "ICQA",
    206214368: "ICQA", 206491342: "ICQA", 204352640: "ICQA", 206957920: "ICQA",
    205793979: "ICQA", 201227946: "ICQA", 202199019: "ICQA", 203724000: "ICQA",
    204278278: "ICQA", 206491344: "ICQA", 204285099: "ICQA", 206214356: "ICQA",
    203386625: "ICQA", 203723997: "ICQA", 206957921: "ICQA", 206957912: "ICQA",
    206957915: "ICQA", 206957916: "ICQA", 202178762: "ICQA", 203635499: "ICQA",
    203023160: "ICQA", 206214412: "ICQA", 203859948: "ICQA", 203386630: "ICQA",
    203953067: "ICQA", 203724001: "ICQA", 205032612: "ICQA", 203953115: "ICQA",
    202280315: "ICQA", 203522266: "ICQA", 112693552: "ICQA", 206958091: "ICQA",
    203319531: "ICQA", 201159413: "ICQA", 206492160: "ICQA", 206958089: "ICQA",
    203877169: "ICQA", 203953116: "ICQA", 203420630: "ICQA", 206958084: "ICQA",
    102582335: "ICQA",
    203796211: "IT", 205642507: "IT", 103156151: "IT", 205152347: "IT",
    203032901: "Inbound", 204002112: "Inbound", 204886947: "Inbound", 112677456: "Inbound",
    206901609: "Inbound", 206901610: "Inbound", 206604482: "Inbound", 204847994: "Inbound",
    206883744: "Inbound", 206857718: "Inbound", 203820466: "Inbound", 206906696: "Inbound",
    109468372: "Inbound", 206199368: "Inbound", 203940988: "Inbound", 206883740: "Inbound",
    203420371: "Inbound", 206901745: "Inbound", 206901748: "Inbound", 206184094: "Inbound",
    203066639: "Inbound", 206854140: "Inbound", 206854313: "Inbound", 204352631: "Inbound",
    203820469: "Inbound", 205988008: "Inbound", 205988012: "Inbound", 112802846: "Inbound",
    205906274: "Inbound", 206362226: "Inbound", 203066774: "Inbound", 206276772: "Inbound",
    204976536: "Inbound", 201635707: "Inbound", 204883088: "Inbound", 204975664: "Inbound",
    206933419: "Inbound", 204352959: "Inbound", 102718057: "Inbound", 206901743: "Inbound",
    202372809: "Inbound", 203391115: "Inbound", 203754940: "Inbound", 112693525: "Inbound",
    102207560: "Inbound", 113231596: "Inbound", 206277231: "Inbound", 206901749: "Inbound",
    206821916: "Inbound", 206868815: "Inbound", 201226562: "Inbound", 206857915: "Inbound",
    206871198: "Inbound", 206901746: "Inbound", 203820468: "Inbound", 203820484: "Inbound",
    202973195: "Inbound", 203395831: "Inbound", 203755274: "Inbound", 203820472: "Inbound",
    206184105: "Inbound", 206502925: "Inbound", 206901938: "Inbound", 203820482: "Inbound",
    206008093: "Inbound", 203259250: "Inbound", 203941722: "Inbound", 203305956: "Inbound",
    202351567: "Inbound", 112917727: "Inbound", 203498528: "Inbound", 203285872: "Inbound",
    206883751: "Inbound", 206117689: "Inbound", 204868852: "Inbound", 203870577: "Inbound",
    200568542: "Inbound", 205988400: "Inbound", 113004696: "Inbound", 204882797: "Inbound",
    204844634: "Inbound", 203252129: "Inbound", 205988015: "Inbound", 206858027: "Inbound",
    112917730: "Inbound", 204950246: "Inbound", 204864792: "Inbound", 206858018: "Inbound",
    206884023: "Inbound", 206884027: "Inbound", 203262180: "Inbound", 206286041: "Inbound",
    204887149: "Inbound", 203859933: "Inbound", 205198876: "Inbound", 206277779: "Inbound",
    206868642: "Inbound", 204002428: "Inbound", 204847986: "Inbound", 206901941: "Inbound",
    204872927: "Inbound", 206854143: "Inbound", 202179029: "Inbound", 113147484: "Inbound",
    206858025: "Inbound", 102207575: "Inbound", 206874038: "Inbound", 206858024: "Inbound",
    204844728: "Inbound", 113009932: "Inbound", 206184098: "Inbound", 204975663: "Inbound",
    203420353: "Inbound", 203953060: "Inbound", 206884034: "Inbound", 206901940: "Inbound",
    105545778: "Inbound", 203253124: "Inbound", 203034914: "Inbound", 204880562: "Inbound",
    206884041: "Inbound", 203498535: "Inbound", 203305958: "Inbound", 202273042: "Inbound",
    204847995: "Inbound", 206117693: "Inbound", 206858020: "Inbound", 206962034: "Inbound",
    206910465: "Inbound", 203725377: "Inbound", 206185557: "Inbound", 203859918: "Inbound",
    206278353: "Inbound", 204864930: "Inbound", 205506106: "Inbound", 205496016: "Inbound",
    106268334: "Inbound", 205988013: "Inbound", 203498279: "Inbound", 206901950: "Inbound",
    205243183: "Inbound", 205839584: "Inbound", 204874770: "Inbound", 203305873: "Inbound",
    204844723: "Inbound", 206287372: "Inbound", 206871355: "Inbound", 206886518: "Inbound",
    204352643: "Inbound", 203042958: "Inbound", 206858021: "Inbound", 203261869: "Inbound",
    206910252: "Inbound", 204865029: "Inbound", 204890071: "Inbound", 206854142: "Inbound",
    206901937: "Inbound", 206826449: "Inbound", 206871354: "Inbound", 206491333: "Inbound",
    204880569: "Inbound", 206889049: "Inbound", 203859962: "Inbound", 206901951: "Inbound",
    204865027: "Inbound", 206905459: "Inbound", 206868817: "Inbound", 206868821: "Inbound",
    113147491: "Inbound", 113004701: "Inbound", 203395808: "Inbound", 206905637: "Inbound",
    206890215: "Inbound", 206868819: "Inbound", 206287375: "Inbound", 203498534: "Inbound",
    201555283: "Inbound", 203859928: "Inbound", 206884122: "Inbound", 206854312: "Inbound",
    204844809: "Inbound", 206491337: "Inbound", 206871527: "Inbound", 203859934: "Inbound",
    203175298: "Inbound", 203252125: "Inbound", 203940997: "Inbound", 103807571: "Inbound",
    203391098: "Inbound", 206886359: "Inbound", 203427938: "Inbound", 203305878: "Inbound",
    204880758: "Inbound", 206915463: "Inbound", 206858023: "Inbound", 206017135: "Inbound",
    113231601: "Inbound", 202973194: "Inbound", 206912034: "Inbound", 206287380: "Inbound",
    204887307: "Inbound", 204278284: "Inbound", 202973250: "Inbound", 206185562: "Inbound",
    112797509: "Inbound", 206858019: "Inbound", 206858026: "Inbound", 206886353: "Inbound",
    205243456: "Inbound", 205247583: "Inbound", 203946293: "Inbound", 204886954: "Inbound",
    203391089: "Inbound", 206905640: "Inbound", 201252987: "Inbound", 206007535: "Inbound",
    112917732: "Inbound", 201228113: "Inbound", 206871916: "Inbound", 205252537: "Inbound",
    203331865: "Inbound", 206912364: "Inbound", 203252254: "Inbound", 206184609: "Inbound",
    206007513: "Inbound", 204868862: "Inbound", 206886631: "Inbound", 203391095: "Inbound",
    203032897: "Inbound", 206905642: "Inbound", 204865032: "Inbound", 206871515: "Inbound",
    204864806: "Inbound", 112782865: "Inbound", 206128891: "Inbound", 201241163: "Inbound",
    206275190: "Inbound", 206905641: "Inbound", 204865043: "Inbound", 206905644: "Inbound",
    206905643: "Inbound", 204002427: "Inbound", 206287674: "Inbound", 102207557: "Inbound",
    204865168: "Inbound", 204865162: "Inbound", 203305963: "Inbound", 203285886: "Inbound",
    204847897: "Inbound", 203252127: "Inbound", 206886637: "Inbound", 206858173: "Inbound",
    204873017: "Inbound", 200568507: "Inbound", 103775151: "Inbound", 206871310: "Inbound",
    206868476: "Inbound", 204847888: "Inbound", 112599365: "Inbound", 203820474: "Inbound",
    206128179: "Inbound", 204868354: "Inbound", 203305870: "Inbound", 203865038: "Inbound",
    204882808: "Inbound", 206905788: "Inbound", 206905942: "Inbound", 203426695: "Inbound",
    102207589: "Inbound", 206905943: "Inbound", 203306090: "Inbound", 204864933: "Inbound",
    204864802: "Inbound", 201399347: "Inbound", 203266099: "Inbound", 206887045: "Inbound",
    203011022: "Inbound", 203306084: "Inbound", 206287665: "Inbound", 206912358: "Inbound",
    204847900: "Inbound", 204887157: "Inbound", 206287664: "Inbound", 203305921: "Inbound",
    203468033: "Inbound", 203420358: "Inbound", 206007521: "Inbound", 204882796: "Inbound",
    112677526: "Inbound", 206827227: "Inbound", 206905945: "Inbound", 203859924: "Inbound",
    202351583: "Inbound", 203306093: "Inbound", 103785332: "Inbound", 104033324: "Inbound",
    206008095: "Inbound", 206007519: "Inbound", 204352632: "Inbound", 204878663: "Inbound",
    203859943: "Inbound", 205231285: "Inbound", 201562558: "Inbound", 112812216: "Inbound",
    203426682: "Inbound", 203498296: "Inbound", 206887051: "Inbound", 202304000: "Inbound",
    206008092: "Inbound", 204395799: "Inbound", 206854323: "Inbound", 206858172: "Inbound",
    203816303: "Inbound", 206287678: "Inbound", 204847899: "Inbound", 206905947: "Inbound",
    203859939: "Inbound", 205275766: "Inbound", 102582332: "Inbound", 202998383: "Inbound",
    205998216: "Inbound", 102207619: "Inbound", 106970075: "Inbound", 203042961: "Inbound",
    105443900: "Inbound", 204887305: "Inbound", 202273448: "Inbound", 206906066: "Inbound",
    206858171: "Inbound", 102718077: "Inbound", 203331931: "Inbound", 206858179: "Inbound",
    202453070: "Inbound", 204873058: "Inbound", 203319506: "Inbound", 103725224: "Inbound",
    203723992: "Inbound", 206858178: "Inbound", 203711781: "Inbound", 206906064: "Inbound",
    203066771: "Inbound", 206906062: "Inbound", 203498277: "Inbound", 205998730: "Inbound",
    113231610: "Inbound", 204847886: "Inbound", 202109274: "Inbound", 204864808: "Inbound",
    112462440: "Inbound", 203179616: "Inbound", 204873065: "Inbound", 204883086: "Inbound",
    206906063: "Inbound", 204869097: "Inbound", 204886953: "Inbound", 206199372: "Inbound",
    203306086: "Inbound", 206887048: "Inbound", 203285897: "Inbound", 202454562: "Inbound",
    203252170: "Inbound", 102207640: "Inbound", 206858176: "Inbound", 204872906: "Inbound",
    206887043: "Inbound", 203331928: "Inbound", 203753799: "Inbound", 203331927: "Inbound",
    206287668: "Inbound", 204848277: "Inbound", 206906068: "Inbound", 201226556: "Inbound",
    206287970: "Inbound", 206854320: "Inbound", 206872064: "Inbound", 206906074: "Inbound",
    204887315: "Inbound", 203182588: "Inbound", 203468031: "Inbound", 102207603: "Inbound",
    206007533: "Inbound", 203940986: "Inbound", 206182847: "Inbound", 205998232: "Inbound",
    206858174: "Inbound", 206854486: "Inbound", 206858175: "Inbound", 206185558: "Inbound",
    204887337: "Inbound", 202998976: "Inbound", 206185107: "Inbound", 201166870: "Inbound",
    112809559: "Inbound", 206007520: "Inbound", 203262045: "Inbound", 204352956: "Inbound",
    203873801: "Inbound", 205087818: "Inbound", 102207632: "Inbound", 206277253: "Inbound",
    204847988: "Inbound", 203420361: "Inbound", 206858180: "Inbound", 206488321: "Inbound",
    206906065: "Inbound", 206858177: "Inbound", 206277251: "Inbound", 206906069: "Inbound",
    206906071: "Inbound", 203305959: "Inbound", 206887052: "Inbound", 203285601: "Inbound",
    204848008: "Inbound", 103094178: "Inbound", 204287692: "Inbound", 204847998: "Inbound",
    102207636: "Inbound", 203285878: "Inbound", 203820496: "Inbound", 102207637: "Inbound",
    203253183: "Inbound", 206887056: "Inbound", 203427075: "Inbound", 112823110: "Inbound",
    206906081: "Inbound", 206858182: "Inbound", 203754935: "Inbound", 203386603: "Inbound",
    112130826: "Inbound", 206887054: "Inbound", 206884261: "Inbound", 203554751: "Inbound",
    206884258: "Inbound", 206018153: "Inbound", 206906083: "Inbound", 206827230: "Inbound",
    206906708: "Inbound", 204887318: "Inbound", 206287968: "Inbound", 206906085: "Inbound",
    112318362: "Inbound", 102207591: "Inbound", 112917725: "Inbound", 203042957: "Inbound",
    204847993: "Inbound", 204352629: "Inbound", 206858274: "Inbound", 206884252: "Inbound",
    203252256: "Inbound", 205998732: "Inbound", 204868857: "Inbound", 203262185: "Inbound",
    204865165: "Inbound", 206858272: "Inbound", 203391092: "Inbound", 203859952: "Inbound",
    206199375: "Inbound", 206182849: "Inbound", 205988010: "Inbound", 112130812: "Inbound",
    204848005: "Inbound", 204278295: "Inbound", 203754941: "Inbound", 206934205: "Inbound",
    206854487: "Inbound", 204848002: "Inbound", 204848000: "Inbound", 203498340: "Inbound",
    204848001: "Inbound", 203420352: "Inbound", 206906070: "Inbound", 206185111: "Inbound",
    205334922: "Inbound", 206868011: "Inbound", 201567441: "Inbound", 204002429: "Inbound",
    204847996: "Inbound", 206884259: "Inbound", 206906078: "Inbound", 204878665: "Inbound",
    203420369: "Inbound", 203252257: "Inbound", 206286063: "Inbound", 206872088: "Inbound",
    206286066: "Inbound", 206906710: "Inbound", 204873062: "Inbound", 204873070: "Inbound",
    108151510: "Inbound", 204873071: "Inbound", 204878666: "Inbound", 103775145: "Inbound",
    206872080: "Inbound", 202218051: "Inbound", 112130816: "Inbound", 203754942: "Inbound",
    203066773: "Inbound", 206884257: "Inbound", 203786019: "Inbound", 203498278: "Inbound",
    206128185: "Inbound", 203102684: "Inbound", 204847990: "Inbound", 206854484: "Inbound",
    206884260: "Inbound", 113004724: "Inbound", 206199451: "Inbound", 203865378: "Inbound",
    206858276: "Inbound", 204002431: "Inbound", 203753808: "Inbound", 203816302: "Inbound",
    204848136: "Inbound", 206276769: "Inbound", 203306102: "Inbound", 203305960: "Inbound",
    205998218: "Inbound", 206276453: "Inbound", 206884124: "Inbound", 203754939: "Inbound",
    203285899: "Inbound", 113090844: "Inbound", 206884123: "Inbound", 206286072: "Inbound",
    203420370: "Inbound", 203262189: "Inbound", 206884038: "Inbound", 203940991: "Inbound",
    204865042: "Inbound", 206826851: "Inbound", 203066638: "Inbound", 204872926: "Inbound",
    201219709: "Inbound", 203306100: "Inbound", 204887343: "Inbound", 206906082: "Inbound",
    206912368: "Inbound", 206884039: "Inbound", 102981725: "Inbound", 203252252: "Inbound",
    206906075: "Inbound", 206884042: "Inbound", 206007522: "Inbound", 112549655: "Inbound",
    203859986: "Inbound", 206901744: "Inbound", 203859961: "Inbound", 203391096: "Inbound",
    207634615: "L&D", 109463287: "L&D", 204892100: "L&D", 201732473: "L&D",
    203066640: "L&D", 203826457: "L&D", 205737007: "L&D", 202979606: "L&D",
    203940992: "L&D",
    203806287: "LP",
    206327189: "Outbound", 206230944: "Outbound", 206230165: "Outbound", 113167085: "Outbound",
    204864797: "Outbound", 206345578: "Outbound", 204946688: "Outbound", 206910233: "Outbound",
    203182633: "Outbound", 203498338: "Outbound", 204857357: "Outbound", 206214795: "Outbound",
    113184873: "Outbound", 206184606: "Outbound", 206913324: "Outbound", 205890140: "Outbound",
    112549603: "Outbound", 203022226: "Outbound", 206913330: "Outbound", 112318345: "Outbound",
    203800776: "Outbound", 206128577: "Outbound", 202449531: "Outbound", 206889009: "Outbound",
    206605132: "Outbound", 202238133: "Outbound", 204874697: "Outbound", 203252251: "Outbound",
    206915461: "Outbound", 205129945: "Outbound", 204483386: "Outbound", 206007537: "Outbound",
    206007525: "Outbound", 204354458: "Outbound", 205807005: "Outbound", 206858371: "Outbound",
    206006419: "Outbound", 201237964: "Outbound", 206006822: "Outbound", 206214874: "Outbound",
    206214859: "Outbound", 202444295: "Outbound", 206850849: "Outbound", 203262052: "Outbound",
    205628527: "Outbound", 203262036: "Outbound", 206214740: "Outbound", 112345881: "Outbound",
    204950249: "Outbound", 203386600: "Outbound", 203306068: "Outbound", 204880309: "Outbound",
    206913325: "Outbound", 203023161: "Outbound", 202973101: "Outbound", 112688802: "Outbound",
    112688803: "Outbound", 203109009: "Outbound", 112677538: "Outbound", 206230164: "Outbound",
    202460468: "Outbound", 202273422: "Outbound", 111249450: "Outbound", 205895595: "Outbound",
    203386577: "Outbound", 206185567: "Outbound", 206623444: "Outbound", 206230578: "Outbound",
    206909508: "Outbound", 203724287: "Outbound", 206913326: "Outbound", 203239158: "Outbound",
    205133643: "Outbound", 206215174: "Outbound", 206909834: "Outbound", 204352642: "Outbound",
    203305927: "Outbound", 206232788: "Outbound", 206183224: "Outbound", 112796297: "Outbound",
    204864794: "Outbound", 203724027: "Outbound", 205627234: "Outbound", 111231915: "Outbound",
    204395798: "Outbound", 112925665: "Outbound", 206230949: "Outbound", 102582323: "Outbound",
    206334168: "Outbound", 203468038: "Outbound", 206185553: "Outbound", 113090839: "Outbound",
    206185556: "Outbound", 206887384: "Outbound", 203395800: "Outbound", 202218619: "Outbound",
    112677485: "Outbound", 113185049: "Outbound", 204283523: "Outbound", 112087288: "Outbound",
    113090806: "Outbound", 200519173: "Outbound", 206915763: "Outbound", 204861260: "Outbound",
    204172229: "Outbound", 206230962: "Outbound", 203262049: "Outbound", 113184879: "Outbound",
    203859925: "Outbound", 202303222: "Outbound", 203788183: "Outbound", 204886951: "Outbound",
    204880308: "Outbound", 206233182: "Outbound", 112715222: "Outbound", 203785995: "Outbound",
    202973087: "Outbound", 203253579: "Outbound", 203711768: "Outbound", 206345582: "Outbound",
    204882560: "Outbound", 206335470: "Outbound", 206914993: "Outbound", 204352638: "Outbound",
    206184101: "Outbound", 204883089: "Outbound", 204872903: "Outbound", 204967154: "Outbound",
    206958457: "Outbound", 204882674: "Outbound", 203735948: "Outbound", 204873012: "Outbound",
    112677467: "Outbound", 203908839: "Outbound", 206335469: "Outbound", 204873188: "Outbound",
    113185053: "Outbound", 103194770: "Outbound", 206846112: "Outbound", 206344587: "Outbound",
    203468103: "Outbound", 204010423: "Outbound", 205270800: "Outbound", 203253177: "Outbound",
    203331890: "Outbound", 203498343: "Outbound", 203386601: "Outbound", 203785997: "Outbound",
    203724368: "Outbound", 206231609: "Outbound", 204946684: "Outbound", 203331893: "Outbound",
    203498339: "Outbound", 204872932: "Outbound", 206914995: "Outbound", 203285721: "Outbound",
    206185101: "Outbound", 203724285: "Outbound", 205088192: "Outbound", 203498566: "Outbound",
    204172554: "Outbound", 204864795: "Outbound", 206957738: "Outbound", 206914999: "Outbound",
    203859965: "Outbound", 204864804: "Outbound", 206335824: "Outbound", 202353219: "Outbound",
    206874113: "Outbound", 205987014: "Outbound", 206336226: "Outbound", 206344592: "Outbound",
    206336549: "Outbound", 201307787: "Outbound", 203020582: "Outbound", 205588021: "Outbound",
    206913331: "Outbound", 206913332: "Outbound", 203306065: "Outbound", 202178812: "Outbound",
    205270392: "Outbound", 204865026: "Outbound", 203859915: "Outbound", 206915006: "Outbound",
    112796300: "Outbound", 206231597: "Outbound", 206915774: "Outbound", 206215177: "Outbound",
    203306064: "Outbound", 206889496: "Outbound", 206915003: "Outbound", 206915481: "Outbound",
    203468099: "Outbound", 206901943: "Outbound", 206344597: "Outbound", 205256965: "Outbound",
    206886728: "Outbound", 206913329: "Outbound", 204277509: "Outbound", 203239157: "Outbound",
    206345580: "Outbound", 206345581: "Outbound", 112404657: "Outbound", 203038681: "Outbound",
    201264962: "Outbound", 206887383: "Outbound", 204060913: "Outbound", 206231598: "Outbound",
    205275764: "Outbound", 203711770: "Outbound", 206185573: "Outbound", 204883231: "Outbound",
    202218621: "Outbound", 203859920: "Outbound", 201260322: "Outbound", 204967037: "Outbound",
    204277501: "Outbound", 203861266: "Outbound", 204277502: "Outbound", 202273424: "Outbound",
    204847999: "Outbound", 203911637: "Outbound", 204277710: "Outbound", 112803410: "Outbound",
    203285803: "Outbound", 205231164: "Outbound", 203859927: "Outbound", 206909642: "Outbound",
    203306061: "Outbound", 206239759: "Outbound", 204227394: "Outbound", 206345988: "Outbound",
    203786001: "Outbound", 206887559: "Outbound", 206886629: "Outbound", 202218656: "Outbound",
    206487521: "Outbound", 204287223: "Outbound", 206345985: "Outbound", 205018781: "Outbound",
    206871915: "Outbound", 204287685: "Outbound", 204882669: "Outbound", 206905462: "Outbound",
    200003517: "Outbound", 203908862: "Outbound", 206215182: "Outbound", 203859963: "Outbound",
    206233198: "Outbound", 204864924: "Outbound", 204864801: "Outbound", 206184100: "Outbound",
    203262040: "Outbound", 206231954: "Outbound", 206345989: "Outbound", 203724460: "Outbound",
    203285901: "Outbound", 203498361: "Outbound", 206346212: "Outbound", 206886641: "Outbound",
    206128579: "Outbound", 206231956: "Outbound", 206231951: "Outbound", 205087817: "Outbound",
    206886520: "Outbound", 203468104: "Outbound", 203020584: "Outbound", 203859930: "Outbound",
    203239153: "Outbound", 205807007: "Outbound", 206871516: "Outbound", 203724371: "Outbound",
    201552881: "Outbound", 110000670: "Outbound", 203285916: "Outbound", 206887403: "Outbound",
    203285905: "Outbound", 111242954: "Outbound", 204883246: "Outbound", 204868352: "Outbound",
    203361279: "Outbound", 204882660: "Outbound", 206910231: "Outbound", 204882801: "Outbound",
    202973086: "Outbound", 206231596: "Outbound", 204872924: "Outbound", 206200802: "Outbound",
    204887138: "Outbound", 205199080: "Outbound", 112677480: "Outbound", 203262050: "Outbound",
    202973102: "Outbound", 206909658: "Outbound", 204278280: "Outbound", 204873013: "Outbound",
    204974332: "Outbound", 203182638: "Outbound", 203861268: "Outbound", 206909650: "Outbound",
    206909644: "Outbound", 204275542: "Outbound", 205300770: "Outbound", 203244441: "Outbound",
    206346214: "Outbound", 206915002: "Outbound", 111231918: "Outbound", 205888752: "Outbound",
    206184102: "Outbound", 203386584: "Outbound", 203331887: "Outbound", 202334268: "Outbound",
    201156082: "Outbound", 206215185: "Outbound", 206909643: "Outbound", 113125893: "Outbound",
    203252082: "Outbound", 201562769: "Outbound", 205592229: "Outbound", 206231957: "Outbound",
    205888753: "Outbound", 206871529: "Outbound", 205271617: "Outbound", 206346213: "Outbound",
    111231903: "Outbound", 112911438: "Outbound", 206910025: "Outbound", 206958456: "Outbound",
    206215176: "Outbound", 206871517: "Outbound", 206346224: "Outbound", 203305919: "Outbound",
    203182583: "Outbound", 113194105: "Outbound", 206183229: "Outbound", 205270805: "Outbound",
    206185559: "Outbound", 206913323: "Outbound", 206846114: "Outbound", 203285738: "Outbound",
    203305992: "Outbound", 113132665: "Outbound", 200568738: "Outbound", 105480604: "Outbound",
    206871522: "Outbound", 206913328: "Outbound", 202218626: "Outbound", 202444344: "Outbound",
    206909657: "Outbound", 204950640: "Outbound", 202444345: "Outbound", 204864812: "Outbound",
    112462416: "Outbound", 204868355: "Outbound", 103776258: "Outbound", 203725372: "Outbound",
    205270384: "Outbound", 205270388: "Outbound", 201531914: "Outbound", 203554786: "Outbound",
    202973103: "Outbound", 206909646: "Outbound", 204886950: "Outbound", 206239753: "Outbound",
    204864927: "Outbound", 206915467: "Outbound", 206913334: "Outbound", 204882819: "Outbound",
    206873872: "Outbound", 203859970: "Outbound", 205888963: "Outbound", 202266083: "Outbound",
    202353223: "Outbound", 206487523: "Outbound", 206344590: "Outbound", 206345202: "Outbound",
    205127504: "Outbound", 206136727: "Outbound", 204172228: "Outbound", 203468100: "Outbound",
    206346547: "Outbound", 203723993: "Outbound", 206873874: "Outbound", 206231949: "Outbound",
    206346554: "Outbound", 203182580: "Outbound", 206231611: "Outbound", 206345200: "Outbound",
    113004710: "Outbound", 205895638: "Outbound", 203182636: "Outbound", 205906298: "Outbound",
    206915465: "Outbound", 204887155: "Outbound", 206873875: "Outbound", 205280263: "Outbound",
    206346539: "Outbound", 206346542: "Outbound", 206346540: "Outbound", 113149155: "Outbound",
    206915468: "Outbound", 203498536: "Outbound", 102207614: "Outbound", 204882824: "Outbound",
    206909842: "Outbound", 206346546: "Outbound", 206884032: "Outbound", 206872071: "Outbound",
    206887050: "Outbound", 206887410: "Outbound", 204878046: "Outbound", 206912037: "Outbound",
    201170582: "Outbound", 113149142: "Outbound", 204882816: "Outbound", 204886086: "Outbound",
    203468098: "Outbound", 204886961: "Outbound", 206346537: "Outbound", 205897608: "Outbound",
    206846695: "Outbound", 202273049: "Outbound", 203386578: "Outbound", 204878049: "Outbound",
    206346556: "Outbound", 206872076: "Outbound", 204946129: "Outbound", 206231604: "Outbound",
    203259248: "Outbound", 206957728: "Outbound", 112318333: "Outbound", 203724454: "Outbound",
    206231963: "Outbound", 206231971: "Outbound", 203724284: "Outbound", 206231964: "Outbound",
    203331904: "Outbound", 206346558: "Outbound", 206957737: "Outbound", 203723998: "Outbound",
    206913333: "Outbound", 206346550: "Outbound", 206409550: "Outbound", 206915771: "Outbound",
    206275183: "Outbound", 205262768: "Outbound", 113154784: "Outbound", 201694540: "Outbound",
    206846691: "Outbound", 206346553: "Outbound", 206346551: "Outbound", 206346549: "Outbound",
    206346548: "Outbound", 206346555: "Outbound", 206346557: "Outbound", 206346552: "Outbound",
    206346544: "Outbound", 206346543: "Outbound", 206346541: "Outbound", 206346538: "Outbound",
    203724002: "Outbound", 203724372: "Outbound", 206858277: "Outbound", 206858181: "Outbound",
    206906084: "Outbound", 202218639: "Outbound", 203331869: "Outbound", 203305872: "Outbound",
    203331895: "Outbound", 202178813: "Outbound", 201237965: "Outbound", 203305868: "Outbound",
    203859923: "Outbound", 205226654: "Outbound", 203859971: "Outbound", 203252168: "Outbound",
    203253434: "Outbound", 203305923: "Outbound", 203785491: "Outbound", 203305996: "Outbound",
    203341075: "Outbound", 203331891: "Outbound", 203331868: "Outbound", 203285902: "Outbound",
    203468105: "Outbound", 203285645: "Outbound", 203285600: "Outbound", 203427537: "Outbound",
    203253841: "Outbound", 202973003: "Outbound", 203252084: "Outbound",
    204838576: "Procurement",
    206908890: "RME",
    205188099: "Safety", 204878673: "Safety", 112352608: "Safety", 204844228: "Safety",
    205837669: "Safety", 104725133: "Safety",
    203285915: "Warehouse Deals", 201621328: "Warehouse Deals", 203724613: "Warehouse Deals", 203724610: "Warehouse Deals",
    112812168: "Warehouse Deals", 202323675: "Warehouse Deals", 204872930: "Warehouse Deals", 203724609: "Warehouse Deals",
    203724608: "Warehouse Deals", 203426680: "Warehouse Deals", 203724607: "Warehouse Deals", 204115302: "Warehouse Deals",
    203724631: "Warehouse Deals", 206286548: "Warehouse Deals", 204125264: "Warehouse Deals", 203331892: "Warehouse Deals",
    203724611: "Warehouse Deals", 204125260: "Warehouse Deals", 203724628: "Warehouse Deals", 204125266: "Warehouse Deals",
    204872928: "Warehouse Deals", 201635281: "Warehouse Deals", 204115306: "Warehouse Deals", 204125268: "Warehouse Deals",
    206948918: "Warehouse Deals", 206873711: "Warehouse Deals", 205420337: "Warehouse Deals", 206958081: "Warehouse Deals",
    111057408: "Warehouse Deals", 201237685: "Warehouse Deals", 202444458: "Warehouse Deals", 204309066: "Warehouse Deals",
    113095004: "Warehouse Deals", 206873876: "Warehouse Deals", 206873877: "Warehouse Deals", 206346559: "Warehouse Deals",
    206346560: "Warehouse Deals", 206346561: "Warehouse Deals", 206346562: "Warehouse Deals", 206346563: "Warehouse Deals",
    206346564: "Warehouse Deals", 206346565: "Warehouse Deals", 206346566: "Warehouse Deals", 206346567: "Warehouse Deals",
    206346568: "Warehouse Deals", 206346569: "Warehouse Deals", 206346570: "Warehouse Deals", 206346571: "Warehouse Deals",
    206346572: "Warehouse Deals", 206346573: "Warehouse Deals", 206346574: "Warehouse Deals", 206346575: "Warehouse Deals",
    206346576: "Warehouse Deals", 206346577: "Warehouse Deals", 206346578: "Warehouse Deals", 206346579: "Warehouse Deals",
    206346580: "Warehouse Deals", 206346581: "Warehouse Deals", 206346582: "Warehouse Deals", 206346583: "Warehouse Deals",
    206346584: "Warehouse Deals", 206346585: "Warehouse Deals", 206346586: "Warehouse Deals", 206346587: "Warehouse Deals",
    206346588: "Warehouse Deals", 206346589: "Warehouse Deals", 206346590: "Warehouse Deals", 206346591: "Warehouse Deals",
    206346592: "Warehouse Deals", 206346593: "Warehouse Deals", 206346594: "Warehouse Deals", 206346595: "Warehouse Deals",
    206346596: "Warehouse Deals", 206346597: "Warehouse Deals", 206346598: "Warehouse Deals", 206346599: "Warehouse Deals",
    206346600: "Warehouse Deals", 206346601: "Warehouse Deals", 206346602: "Warehouse Deals", 206346603: "Warehouse Deals",
    206346604: "Warehouse Deals", 206346605: "Warehouse Deals", 206346606: "Warehouse Deals", 206346607: "Warehouse Deals",
    206346608: "Warehouse Deals", 206346609: "Warehouse Deals", 206346610: "Warehouse Deals", 206346611: "Warehouse Deals",
    206346612: "Warehouse Deals", 206346613: "Warehouse Deals", 206346614: "Warehouse Deals", 206346615: "Warehouse Deals",
    206346616: "Warehouse Deals", 206346617: "Warehouse Deals", 206346618: "Warehouse Deals", 206346619: "Warehouse Deals",
    206346620: "Warehouse Deals", 206346621: "Warehouse Deals", 206346622: "Warehouse Deals", 206346623: "Warehouse Deals",
    206346624: "Warehouse Deals", 206346625: "Warehouse Deals", 206346626: "Warehouse Deals", 206346627: "Warehouse Deals",
    206346628: "Warehouse Deals", 206346629: "Warehouse Deals", 206346630: "Warehouse Deals", 206346631: "Warehouse Deals",
    206346632: "Warehouse Deals", 206346633: "Warehouse Deals", 206346634: "Warehouse Deals", 206346635: "Warehouse Deals",
    206346636: "Warehouse Deals", 206346637: "Warehouse Deals", 206346638: "Warehouse Deals", 206346639: "Warehouse Deals",
    206346640: "Warehouse Deals", 206346641: "Warehouse Deals", 206346642: "Warehouse Deals", 206346643: "Warehouse Deals",
    206346644: "Warehouse Deals", 206346645: "Warehouse Deals", 206346646: "Warehouse Deals", 206346647: "Warehouse Deals",
    206346648: "Warehouse Deals", 206346649: "Warehouse Deals", 206346650: "Warehouse Deals", 206346651: "Warehouse Deals",
    206346652: "Warehouse Deals", 206346653: "Warehouse Deals", 206346654: "Warehouse Deals", 206346655: "Warehouse Deals",
    206346656: "Warehouse Deals", 206346657: "Warehouse Deals", 206346658: "Warehouse Deals", 206346659: "Warehouse Deals",
    206346660: "Warehouse Deals",
    102237456: "XAEC",
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
    
