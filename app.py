

import streamlit as st
import pandas as pd
import numpy as np
from datetime import date, timedelta
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

st.set_page_config(page_title="Break Compliance Report", page_icon="📊", layout="centered")

st.markdown("# 📊 Break Compliance Report")
st.markdown("Upload attendance CSV → Get formatted report instantly!")
st.markdown("---")

# === DEPARTMENT MAPPING (Employee ID -> Department) ===
# Last updated: 27 July 2026 | Total: 711 employees
# To update: edit this dictionary when people join/leave
DEPT_MAP = {
    '205247342': 'HR', '107322190': 'HR', '203826463': 'HR', '200519423': 'HR',
    '205930604': 'HR', '207302255': 'HR', '206018669': 'HR', '205181646': 'HR',
    '112806642': 'HR', '207758631': 'HR', '207407185': 'HR',
    '206827222': 'ICQA', '206276761': 'ICQA', '206277220': 'ICQA', '205226912': 'ICQA',
    '206199903': 'ICQA', '206277223': 'ICQA', '205005081': 'ICQA', '202273043': 'ICQA',
    '113115097': 'ICQA', '206276766': 'ICQA', '205221423': 'ICQA', '203953061': 'ICQA',
    '206287369': 'ICQA', '206277230': 'ICQA', '204942041': 'ICQA', '206910245': 'ICQA',
    '103788917': 'ICQA', '205278825': 'ICQA', '206297797': 'ICQA', '204942695': 'ICQA',
    '203427120': 'ICQA', '106295884': 'ICQA', '205281248': 'ICQA', '206276765': 'ICQA',
    '206276768': 'ICQA', '109463379': 'ICQA', '205252541': 'ICQA', '205279079': 'ICQA',
    '106268210': 'ICQA',
    '205048424': 'IT',
    '205548278': 'Inbound', '207450006': 'Inbound', '112347709': 'Inbound', '206276746': 'Inbound',
    '205278821': 'Inbound', '206118440': 'Inbound', '205271349': 'Inbound', '204868847': 'Inbound',
    '206912050': 'Inbound', '205588020': 'Inbound', '205996137': 'Inbound', '206503613': 'Inbound',
    '205588038': 'Inbound', '204256643': 'Inbound', '205609141': 'Inbound', '205275640': 'Inbound',
    '206276760': 'Inbound', '112102326': 'Inbound', '111009303': 'Inbound', '205609142': 'Inbound',
    '207452565': 'Inbound', '206117376': 'Inbound', '205588031': 'Inbound', '204880561': 'Inbound',
    '206889002': 'Inbound', '204950245': 'Inbound', '205271607': 'Inbound', '206871806': 'Inbound',
    '203816299': 'Inbound', '205227221': 'Inbound', '206906470': 'Inbound', '203755273': 'Inbound',
    '206906709': 'Inbound', '206914170': 'Inbound', '112874976': 'Inbound', '206064517': 'Inbound',
    '205548592': 'Inbound', '206193608': 'Inbound', '206871345': 'Inbound', '205256721': 'Inbound',
    '206239208': 'Inbound', '105445261': 'Inbound', '205252350': 'Inbound', '205807261': 'Inbound',
    '206192966': 'Inbound', '206809052': 'Inbound', '205627229': 'Inbound', '205928017': 'Inbound',
    '206502928': 'Inbound', '206490657': 'Inbound', '206605144': 'Inbound', '206873642': 'Inbound',
    '206230947': 'Inbound', '206889484': 'Inbound', '109468051': 'Inbound', '205199521': 'Inbound',
    '203820467': 'Inbound', '206912024': 'Inbound', '112079690': 'Inbound', '102718104': 'Inbound',
    '112874979': 'Inbound', '207442202': 'Inbound', '205256966': 'Inbound', '205271073': 'Inbound',
    '206827220': 'Inbound', '203287189': 'Inbound', '205635077': 'Inbound', '205555756': 'Inbound',
    '205231165': 'Inbound', '111144766': 'Inbound', '106762603': 'Inbound', '205231163': 'Inbound',
    '204947098': 'Inbound', '206827384': 'Inbound', '206504147': 'Inbound', '206200371': 'Inbound',
    '205275958': 'Inbound', '206826456': 'Inbound', '206871360': 'Inbound', '205227219': 'Inbound',
    '206503623': 'Inbound', '206826597': 'Inbound', '205231531': 'Inbound', '205995592': 'Inbound',
    '205555739': 'Inbound', '205939859': 'Inbound', '113069850': 'Inbound', '200714293': 'Inbound',
    '204277712': 'Inbound', '204869093': 'Inbound', '204869085': 'Inbound', '205807009': 'Inbound',
    '206906249': 'Inbound', '204886845': 'Inbound', '205198875': 'Inbound', '205628512': 'Inbound',
    '205548599': 'Inbound', '204879184': 'Inbound', '206232510': 'Inbound', '206906699': 'Inbound',
    '206906454': 'Inbound', '204887302': 'Inbound', '205555489': 'Inbound', '205807010': 'Inbound',
    '205615608': 'Inbound', '206826604': 'Inbound', '206193601': 'Inbound', '205231536': 'Inbound',
    '206503615': 'Inbound', '203816397': 'Inbound', '205271611': 'Inbound', '207451449': 'Inbound',
    '109748876': 'Inbound', '205628513': 'Inbound', '205616246': 'Inbound', '205555757': 'Inbound',
    '206296197': 'Inbound', '203306085': 'Inbound', '206906705': 'Inbound', '206230570': 'Inbound',
    '205938172': 'Inbound', '206875498': 'Inbound', '206906703': 'Inbound', '203554783': 'Inbound',
    '206822556': 'Inbound', '206871651': 'Inbound', '206517043': 'Inbound', '205985787': 'Inbound',
    '205199825': 'Inbound', '206827219': 'Inbound', '206503617': 'Inbound', '205226151': 'Inbound',
    '205635734': 'Inbound', '206872074': 'Inbound', '206906460': 'Inbound', '105443911': 'Inbound',
    '207452567': 'Inbound', '201771653': 'Inbound', '206286062': 'Inbound', '206871352': 'Inbound',
    '206007532': 'Inbound', '206277234': 'Inbound', '204967155': 'Inbound', '112318363': 'Inbound',
    '103202992': 'Inbound', '203305869': 'Inbound', '206871650': 'Inbound', '205588026': 'Inbound',
    '206201920': 'Inbound', '206201922': 'Inbound', '206277236': 'Inbound', '205252534': 'Inbound',
    '112393372': 'Inbound', '203850781': 'Inbound', '205231287': 'Inbound', '205627612': 'Inbound',
    '206014830': 'Inbound', '112599352': 'Inbound', '206906251': 'Inbound', '205928020': 'Inbound',
    '112920901': 'Inbound', '206827231': 'Inbound', '205555752': 'Inbound', '206889040': 'Inbound',
    '206827054': 'Inbound', '206873652': 'Inbound', '206873707': 'Inbound', '206193598': 'Inbound',
    '206874050': 'Inbound', '111144765': 'Inbound', '205231417': 'Inbound', '206277242': 'Inbound',
    '206201913': 'Inbound', '205938170': 'Inbound', '206277245': 'Inbound', '206910488': 'Inbound',
    '206199836': 'Inbound', '206200349': 'Inbound', '205588027': 'Inbound', '206277248': 'Inbound',
    '205227532': 'Inbound', '205252368': 'Inbound', '206873564': 'Inbound', '203498301': 'Inbound',
    '205226644': 'Inbound', '206504139': 'Inbound', '206906248': 'Inbound', '206906244': 'Inbound',
    '206502929': 'Inbound', '206889445': 'Inbound', '206889501': 'Inbound', '206871520': 'Inbound',
    '205252364': 'Inbound', '205271605': 'Inbound', '112408798': 'Inbound', '206827237': 'Inbound',
    '206232513': 'Inbound', '206888996': 'Inbound', '206906252': 'Inbound', '205615603': 'Inbound',
    '206874040': 'Inbound', '206201668': 'Inbound', '203066637': 'Inbound', '206504134': 'Inbound',
    '204992326': 'Inbound', '206199902': 'Inbound', '203563362': 'Inbound', '102207583': 'Inbound',
    '204886385': 'Inbound', '206827232': 'Inbound', '207442606': 'Inbound', '206889486': 'Inbound',
    '205271625': 'Inbound', '205231419': 'Inbound', '206889502': 'Inbound', '203011027': 'Inbound',
    '206827236': 'Inbound', '205807260': 'Inbound', '206276754': 'Inbound', '205996134': 'Inbound',
    '206873706': 'Inbound', '206296198': 'Inbound', '205252753': 'Inbound', '206889495': 'Inbound',
    '204996615': 'Inbound', '203679314': 'Inbound', '205252748': 'Inbound', '205635349': 'Inbound',
    '206827055': 'Inbound', '206889490': 'Inbound', '206502396': 'Inbound', '206502923': 'Inbound',
    '206906716': 'Inbound', '204887316': 'Inbound', '206906465': 'Inbound', '205252366': 'Inbound',
    '205588036': 'Inbound', '206809048': 'Inbound', '206889503': 'Inbound', '203859973': 'Inbound',
    '205937333': 'Inbound', '205548602': 'Inbound', '112087286': 'Inbound', '206906253': 'Inbound',
    '206906245': 'Inbound', '205257288': 'Inbound', '206826451': 'Inbound', '206277252': 'Inbound',
    '206906707': 'Inbound', '200176091': 'Inbound', '204950646': 'Inbound', '205199353': 'Inbound',
    '205199831': 'Inbound', '206873709': 'Inbound', '203755271': 'Inbound', '206906462': 'Inbound',
    '206873882': 'Inbound', '206906466': 'Inbound', '206117375': 'Inbound', '206906698': 'Inbound',
    '205252755': 'Inbound', '205588039': 'Inbound', '206503241': 'Inbound', '112874987': 'Inbound',
    '205252369': 'Inbound', '206912369': 'Inbound', '206827234': 'Inbound', '205627232': 'Inbound',
    '205986154': 'Inbound', '203285597': 'Inbound', '204887314': 'Inbound', '203741549': 'Inbound',
    '204976985': 'L&D', '205221421': 'L&D', '205109148': 'L&D', '204352950': 'L&D',
    '204821683': 'L&D', '112181825': 'L&D', '205231288': 'L&D', '200104984': 'L&D',
    '205781111': 'L&D', '112235562': 'L&D',
    '205271613': 'Outbound', '205252359': 'Outbound', '205252371': 'Outbound', '206912041': 'Outbound',
    '206912021': 'Outbound', '206277256': 'Outbound', '204890376': 'Outbound', '206488959': 'Outbound',
    '105454643': 'Outbound', '205010921': 'Outbound', '205937794': 'Outbound', '205252367': 'Outbound',
    '206475834': 'Outbound', '206502395': 'Outbound', '206910232': 'Outbound', '102718055': 'Outbound',
    '206276756': 'Outbound', '206912372': 'Outbound', '206914177': 'Outbound', '205227537': 'Outbound',
    '205252360': 'Outbound', '205195116': 'Outbound', '206117706': 'Outbound', '205987383': 'Outbound',
    '205252358': 'Outbound', '206871649': 'Outbound', '205257387': 'Outbound', '205258832': 'Outbound',
    '204950655': 'Outbound', '206953457': 'Outbound', '206199926': 'Outbound', '105480601': 'Outbound',
    '205252745': 'Outbound', '204010638': 'Outbound', '205226915': 'Outbound', '206200799': 'Outbound',
    '205937793': 'Outbound', '206230572': 'Outbound', '205202339': 'Outbound', '205226623': 'Outbound',
    '204886573': 'Outbound', '206200350': 'Outbound', '206326900': 'Outbound', '205231538': 'Outbound',
    '102207569': 'Outbound', '205195118': 'Outbound', '206277782': 'Outbound', '206871356': 'Outbound',
    '205199355': 'Outbound', '206605147': 'Outbound', '206871348': 'Outbound', '206605138': 'Outbound',
    '206605143': 'Outbound', '205231427': 'Outbound', '109407241': 'Outbound', '206827238': 'Outbound',
    '204887304': 'Outbound', '204264827': 'Outbound', '103751774': 'Outbound', '206200373': 'Outbound',
    '205198880': 'Outbound', '206326892': 'Outbound', '206361296': 'Outbound', '204890375': 'Outbound',
    '206934192': 'Outbound', '206240779': 'Outbound', '205201893': 'Outbound', '205198879': 'Outbound',
    '112911424': 'Outbound', '205226919': 'Outbound', '204882818': 'Outbound', '204946694': 'Outbound',
    '204946692': 'Outbound', '206912360': 'Outbound', '206910474': 'Outbound', '206200345': 'Outbound',
    '205252365': 'Outbound', '112409312': 'Outbound', '206200812': 'Outbound', '204874830': 'Outbound',
    '206118435': 'Outbound', '206128176': 'Outbound', '206063872': 'Outbound', '206912025': 'Outbound',
    '206361291': 'Outbound', '206062910': 'Outbound', '206278361': 'Outbound', '206489360': 'Outbound',
    '203660434': 'Outbound', '110163302': 'Outbound', '105444861': 'Outbound', '206889509': 'Outbound',
    '206605139': 'Outbound', '206193611': 'Outbound', '206277228': 'Outbound', '205199354': 'Outbound',
    '205252532': 'Outbound', '206914171': 'Outbound', '204009967': 'Outbound', '206871518': 'Outbound',
    '106295861': 'Outbound', '206193602': 'Outbound', '206949880': 'Outbound', '206503240': 'Outbound',
    '106970098': 'Outbound', '206889476': 'Outbound', '205252540': 'Outbound', '204950644': 'Outbound',
    '205615612': 'Outbound', '206915176': 'Outbound', '206606707': 'Outbound', '204966774': 'Outbound',
    '206233183': 'Outbound', '206200343': 'Outbound', '108282049': 'Outbound', '206606702': 'Outbound',
    '205226917': 'Outbound', '205977131': 'Outbound', '206327186': 'Outbound', '206193139': 'Outbound',
    '206327624': 'Outbound', '205592629': 'Outbound', '206326893': 'Outbound', '206503242': 'Outbound',
    '206889265': 'Outbound', '205199828': 'Outbound', '206604513': 'Outbound', '205227528': 'Outbound',
    '206232730': 'Outbound', '102207594': 'Outbound', '206242016': 'Outbound', '205633860': 'Outbound',
    '206277227': 'Outbound', '104307217': 'Outbound', '206326919': 'Outbound', '206287660': 'Outbound',
    '206606012': 'Outbound', '206809049': 'Outbound', '206326922': 'Outbound', '205257129': 'Outbound',
    '203850670': 'Outbound', '206287663': 'Outbound', '206874043': 'Outbound', '206808897': 'Outbound',
    '206277226': 'Outbound', '206231953': 'Outbound', '206606245': 'Outbound', '206906711': 'Outbound',
    '205252752': 'Outbound', '205227527': 'Outbound', '204951298': 'Outbound', '206604511': 'Outbound',
    '206192842': 'Outbound', '205600854': 'Outbound', '206277229': 'Outbound', '204873067': 'Outbound',
    '206237382': 'Outbound', '206063908': 'Outbound', '206604500': 'Outbound', '205922854': 'Outbound',
    '105444811': 'Outbound', '206266031': 'Outbound', '206201506': 'Outbound', '103751763': 'Outbound',
    '204890575': 'Outbound', '205227530': 'Outbound', '206193612': 'Outbound', '112347684': 'Outbound',
    '206287373': 'Outbound', '204891069': 'Outbound', '205627226': 'Outbound', '205615606': 'Outbound',
    '206912362': 'Outbound', '206277217': 'Outbound', '104277451': 'Outbound', '206914382': 'Outbound',
    '206490651': 'Outbound', '205231420': 'Outbound', '204950639': 'Outbound', '206607046': 'Outbound',
    '206889491': 'Outbound', '206490644': 'Outbound', '206200467': 'Outbound', '206489359': 'Outbound',
    '205278818': 'Outbound', '205231166': 'Outbound', '105443956': 'Outbound', '205231031': 'Outbound',
    '205230944': 'Outbound', '205231033': 'Outbound', '205231422': 'Outbound', '205226924': 'Outbound',
    '205221561': 'Outbound', '206606030': 'Outbound', '205940154': 'Outbound', '205231030': 'Outbound',
    '205278827': 'Outbound', '206117700': 'Outbound', '206277246': 'Outbound', '205939338': 'Outbound',
    '206201914': 'Outbound', '206576240': 'Outbound', '206889507': 'Outbound', '206606246': 'Outbound',
    '206490655': 'Outbound', '206201475': 'Outbound', '205939341': 'Outbound', '205221562': 'Outbound',
    '205231032': 'Outbound', '205227538': 'Outbound', '205221565': 'Outbound', '206277238': 'Outbound',
    '205592632': 'Outbound', '205231530': 'Outbound', '206504140': 'Outbound', '206910494': 'Outbound',
    '205940163': 'Outbound', '206934634': 'Outbound', '206490648': 'Outbound', '110163345': 'Outbound',
    '206118442': 'Outbound', '206490646': 'Outbound', '204886247': 'Outbound', '205938665': 'Outbound',
    '102207596': 'Outbound', '204887646': 'Outbound', '207337286': 'Outbound', '204868855': 'Outbound',
    '206607045': 'Outbound', '206361304': 'Outbound', '205231537': 'Outbound', '206276757': 'Outbound',
    '103989125': 'Outbound', '204967050': 'Outbound', '206490653': 'Outbound', '206232736': 'Outbound',
    '205252751': 'Outbound', '102207658': 'Outbound', '205586502': 'Outbound', '206892313': 'Outbound',
    '206199999': 'Outbound', '206231958': 'Outbound', '204946683': 'Outbound', '205591904': 'Outbound',
    '206327192': 'Outbound', '206276764': 'Outbound', '206912048': 'Outbound', '206874048': 'Outbound',
    '205929364': 'Outbound', '206889487': 'Outbound', '205199826': 'Outbound', '205593691': 'Outbound',
    '206276762': 'Outbound', '204946687': 'Outbound', '207338414': 'Outbound', '205939281': 'Outbound',
    '206809054': 'Outbound', '205199829': 'Outbound', '205257283': 'Outbound', '205252355': 'Outbound',
    '206361294': 'Outbound', '206276770': 'Outbound', '204886245': 'Outbound', '204878040': 'Outbound',
    '205231292': 'Outbound', '206606714': 'Outbound', '206276773': 'Outbound', '206277254': 'Outbound',
    '205256723': 'Outbound', '205949501': 'Outbound', '206912065': 'Outbound', '205592630': 'Outbound',
    '205628536': 'Outbound', '204352636': 'Outbound', '206237363': 'Outbound', '206193613': 'Outbound',
    '203724002': 'Outbound', '203724372': 'Outbound', '206858277': 'Outbound', '206858181': 'Outbound',
    '206906084': 'Outbound', '202218639': 'Outbound', '203331869': 'Outbound', '203305872': 'Outbound',
    '203331895': 'Outbound', '202178813': 'Outbound', '201237965': 'Outbound', '203305868': 'Outbound',
    '203859923': 'Outbound', '205226654': 'Outbound', '203859971': 'Outbound', '203252168': 'Outbound',
    '203253434': 'Outbound', '203305923': 'Outbound', '203785491': 'Outbound', '203305996': 'Outbound',
    '203341075': 'Outbound', '203331891': 'Outbound', '203331868': 'Outbound', '203285902': 'Outbound',
    '203468105': 'Outbound', '203285645': 'Outbound', '203285600': 'Outbound', '203427537': 'Outbound',
    '203253841': 'Outbound', '202973003': 'Outbound', '203252084': 'Outbound',
    '205005101': 'Procurement',
    '204985884': 'RME', '205001748': 'RME', '205004710': 'RME', '204821681': 'RME',
    '205897595': 'RME', '207245175': 'RME', '206908890': 'RME',
    '102582347': 'Safety', '113135599': 'Safety',
    '205252348': 'Warehouse Deals', '201645184': 'Warehouse Deals', '206326895': 'Warehouse Deals', '205325442': 'Warehouse Deals',
    '206827042': 'Warehouse Deals', '204880757': 'Warehouse Deals', '204319795': 'Warehouse Deals', '205257635': 'Warehouse Deals',
    '206128190': 'Warehouse Deals', '205928925': 'Warehouse Deals', '206912028': 'Warehouse Deals', '206361289': 'Warehouse Deals',
    '205278824': 'Warehouse Deals', '203660528': 'Warehouse Deals', '206889047': 'Warehouse Deals', '206822554': 'Warehouse Deals',
    '205257290': 'Warehouse Deals', '205199113': 'Warehouse Deals', '102981764': 'Warehouse Deals', '206906713': 'Warehouse Deals',
    '206502927': 'Warehouse Deals', '204319801': 'Warehouse Deals', '205937796': 'Warehouse Deals', '203850805': 'Warehouse Deals',
    '206912036': 'Warehouse Deals', '102582254': 'Warehouse Deals', '102207600': 'Warehouse Deals', '206827224': 'Warehouse Deals',
    '205206430': 'Warehouse Deals', '113175683': 'Warehouse Deals', '206128186': 'Warehouse Deals', '205280642': 'Warehouse Deals',
    '205231293': 'Warehouse Deals', '205199524': 'Warehouse Deals', '206327199': 'Warehouse Deals', '205256726': 'Warehouse Deals',
    '206822072': 'Warehouse Deals', '205199824': 'Warehouse Deals', '204967042': 'Warehouse Deals', '206827229': 'Warehouse Deals',
    '206910243': 'Warehouse Deals', '206827391': 'Warehouse Deals', '206504127': 'Warehouse Deals', '203816401': 'Warehouse Deals',
    '206934440': 'Warehouse Deals', '204373627': 'Warehouse Deals', '102205930': 'Warehouse Deals', '203660393': 'Warehouse Deals',
    '206361301': 'Warehouse Deals', '206284692': 'Warehouse Deals', '206934610': 'Warehouse Deals', '206277247': 'Warehouse Deals',
    '205271612': 'Warehouse Deals', '205243458': 'Warehouse Deals', '204319465': 'Warehouse Deals', '206278356': 'Warehouse Deals',
    '205252361': 'Warehouse Deals', '106295659': 'Warehouse Deals', '102205652': 'Warehouse Deals', '102205991': 'Warehouse Deals',
    '206128178': 'Warehouse Deals', '112728750': 'Warehouse Deals', '201179452': 'Warehouse Deals', '206128188': 'Warehouse Deals',
    '206327621': 'Warehouse Deals', '204950874': 'Warehouse Deals', '113249574': 'Warehouse Deals', '205231428': 'Warehouse Deals',
    '205252531': 'Warehouse Deals', '113142038': 'Warehouse Deals', '205252362': 'Warehouse Deals',
}

def get_department(emp_id):
    """Lookup department by Employee ID."""
    return DEPT_MAP.get(str(int(float(emp_id))) if emp_id else '', 'Unknown')

# === HISTORY UPLOAD (optional) ===
history_file = st.file_uploader("📋 Upload History (optional)", type=['csv'], help="Previous history.csv for repeat tracking")

history = pd.DataFrame(columns=['Employee ID', 'Employee Name', 'Department', 'Date', 'Flag'])
if history_file is not None:
    try:
        history = pd.read_csv(history_file)
        history['Employee ID'] = history['Employee ID'].astype(str)
        st.success(f"📋 History loaded: {len(history)} past records")
    except Exception as e:
        st.warning(f"⚠️ Could not read history: {e}")

# === MAIN UPLOAD ===
uploaded_file = st.file_uploader("📄 Upload Attendance CSV", type=['csv'], help="Required — your daily attendance export")

if uploaded_file is not None:
    df = pd.read_csv(uploaded_file)
    st.success(f"✅ Loaded {len(df)} employees!")

    today_str = date.today().strftime('%Y-%m-%d')

    def safe_int(val):
        try:
            if pd.notna(val):
                return int(float(val))
            return 0
        except:
            return 0

    # Add department from hardcoded map
    df['Employee ID'] = df['Employee ID'].astype(str)
    df['Dept'] = df['Employee ID'].apply(lambda x: DEPT_MAP.get(str(safe_int(x)), 'Unknown'))

    # Process break logic
    df['Break Type'] = np.where(
        df['1st Break Status'].str.contains('Combined', na=False), 'IB', 'OB'
    )
    df['Total Break (min)'] = np.where(
        df['Break Type'] == 'IB',
        df['1st Break (min)'].fillna(0),
        df['1st Break (min)'].fillna(0) + df['2nd Break (min)'].fillna(0)
    )

    # Separate missed punch
    missed = df[df['Total Break (min)'] == 0][['Employee ID', 'Employee Name', 'Dept', 'Total Break (min)']].sort_values('Employee Name')

    # Flag Excess/Less
    df_with_breaks = df[df['Total Break (min)'] > 0].copy()
    df_with_breaks['Break Flag'] = np.where(
        df_with_breaks['Total Break (min)'] >= 65, 'Excess Break',
        np.where(df_with_breaks['Total Break (min)'] <= 55, 'Less Break', 'OK')
    )

    excess = df_with_breaks[df_with_breaks['Break Flag']=='Excess Break'][['Employee ID','Employee Name','Dept','Total Break (min)']].sort_values('Total Break (min)', ascending=False)
    less = df_with_breaks[df_with_breaks['Break Flag']=='Less Break'][['Employee ID','Employee Name','Dept','Total Break (min)']].sort_values('Total Break (min)')

    # Repeat count function
    def get_repeat_count(employee_id, flag_type):
        if history.empty:
            return 0
        past = history[
            (history['Employee ID'] == str(employee_id)) &
            (history['Flag'] == flag_type) &
            (history['Date'] != today_str)
        ]
        return len(past)

    def add_repeat_info(data, flag_type):
        data = data.copy()
        data['Repeat'] = data['Employee ID'].apply(lambda x: get_repeat_count(x, flag_type))
        return data

    excess_display = add_repeat_info(excess, 'Excess')
    less_display = add_repeat_info(less, 'Less')
    missed_display = add_repeat_info(missed, 'Missed')

    # Show metrics
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Total Exceptions", len(excess) + len(less))
    col2.metric("Excess (≥65 min)", len(excess))
    col3.metric("Less (≤55 min)", len(less))
    col4.metric("Missed Punch", len(missed))

    st.markdown("---")

    # Display tables
    def show_table(title, emoji, data):
        st.markdown(f"### {emoji} {title}")
        if len(data) == 0:
            st.info("No exceptions found ✅")
        else:
            display = data.copy().reset_index(drop=True)
            display.columns = ['Employee ID', 'Employee Name', 'Department', 'Break (min)', 'Repeat']
            display['Repeat'] = display['Repeat'].apply(lambda x: f"⚠️ {x + 1}x" if x > 0 else "")
            st.dataframe(display, use_container_width=True)

    show_table("Excess Break (≥65 min)", "🔴", excess_display)
    show_table("Less Break (≤55 min)", "🟠", less_display)
    show_table("Missed Break Punch (0 min)", "⚫", missed_display)

    st.markdown("---")

    # Build updated history
    new_records = []
    for _, emp in excess.iterrows():
        new_records.append({'Employee ID': str(safe_int(emp['Employee ID'])), 'Employee Name': str(emp['Employee Name']), 'Department': str(emp['Dept']), 'Date': today_str, 'Flag': 'Excess'})
    for _, emp in less.iterrows():
        new_records.append({'Employee ID': str(safe_int(emp['Employee ID'])), 'Employee Name': str(emp['Employee Name']), 'Department': str(emp['Dept']), 'Date': today_str, 'Flag': 'Less'})
    for _, emp in missed.iterrows():
        new_records.append({'Employee ID': str(safe_int(emp['Employee ID'])), 'Employee Name': str(emp['Employee Name']), 'Department': str(emp['Dept']), 'Date': today_str, 'Flag': 'Missed'})

    if not history.empty:
        updated_history = history[history['Date'] != today_str].copy()
        updated_history = pd.concat([updated_history, pd.DataFrame(new_records)], ignore_index=True)
    else:
        updated_history = pd.DataFrame(new_records)

    # Generate Excel Report
    def generate_report_excel():
        wb = Workbook()
        ws = wb.active
        ws.title = "Report"
        ws.sheet_view.showGridLines = False

        squid_ink = '232F3E'
        teal = '00BCD4'
        coral = 'FF6B6B'
        sunset = 'FFA726'
        charcoal = '424242'
        snow = 'FAFAFA'
        ice_blue = 'E0F7FA'
        light_coral = 'FFEBEE'
        light_sunset = 'FFF3E0'
        light_charcoal = 'F5F5F5'
        white = 'FFFFFF'
        dark_text = '212121'
        repeat_red = 'D32F2F'
        repeat_bg = 'FFCDD2'

        ws.column_dimensions['A'].width = 14
        ws.column_dimensions['B'].width = 28
        ws.column_dimensions['C'].width = 16
        ws.column_dimensions['D'].width = 10
        ws.column_dimensions['E'].width = 12

        for r in range(1, 3):
            for c in range(1, 6):
                ws.cell(row=r, column=c).fill = PatternFill(start_color=squid_ink, end_color=squid_ink, fill_type='solid')
        ws.row_dimensions[1].height = 10
        ws.row_dimensions[2].height = 35
        ws.merge_cells('A2:E2')
        ws['A2'] = "BREAK COMPLIANCE REPORT"
        ws['A2'].font = Font(name='Calibri', size=16, bold=True, color=teal)
        ws['A2'].alignment = Alignment(vertical='center', horizontal='center')

        ws.row_dimensions[3].height = 4
        for c in range(1, 6):
            ws.cell(row=3, column=c).fill = PatternFill(start_color=teal, end_color=teal, fill_type='solid')

        ws.row_dimensions[4].height = 22
        ws.merge_cells('A4:E4')
        ws['A4'] = date.today().strftime("%A, %d %B %Y")
        ws['A4'].font = Font(name='Calibri', size=10, color='666666')
        ws['A4'].alignment = Alignment(vertical='center', horizontal='center')
        for c in range(1, 6):
            ws.cell(row=4, column=c).fill = PatternFill(start_color=snow, end_color=snow, fill_type='solid')

        ws.row_dimensions[5].height = 30
        metrics = [
            (len(excess) + len(less), squid_ink, ice_blue),
            (len(excess), coral, light_coral),
            (len(less), sunset, light_sunset),
            (len(missed), charcoal, light_charcoal),
        ]
        for i, (val, font_color, bg_color) in enumerate(metrics, 1):
            cell = ws.cell(row=5, column=i, value=val)
            cell.font = Font(name='Calibri', size=14 if i > 1 else 18, bold=True, color=font_color)
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.fill = PatternFill(start_color=bg_color, end_color=bg_color, fill_type='solid')
        ws.cell(row=5, column=5).fill = PatternFill(start_color=snow, end_color=snow, fill_type='solid')

        ws.row_dimensions[6].height = 14
        labels = [("total", '888888'), ("excess", coral), ("less", sunset), ("missed", charcoal)]
        for i, (label, color) in enumerate(labels, 1):
            cell = ws.cell(row=6, column=i, value=label)
            cell.font = Font(name='Calibri', size=8, color=color)
            cell.alignment = Alignment(horizontal='center')
        ws.row_dimensions[7].height = 8

        def write_table(ws, row, title, header_color, light_color, data_df, flag_type):
            ws.merge_cells(f'A{row}:E{row}')
            ws[f'A{row}'] = title
            ws[f'A{row}'].font = Font(name='Calibri', size=10, bold=True, color=white)
            ws[f'A{row}'].fill = PatternFill(start_color=header_color, end_color=header_color, fill_type='solid')
            ws[f'A{row}'].alignment = Alignment(vertical='center')
            for c in range(1, 6):
                ws.cell(row=row, column=c).fill = PatternFill(start_color=header_color, end_color=header_color, fill_type='solid')
            ws.row_dimensions[row].height = 20
            row += 1

            for col, header in enumerate(['Employee ID', 'Name', 'Department', 'Min', 'Repeat'], 1):
                cell = ws.cell(row=row, column=col, value=header)
                cell.font = Font(name='Calibri', size=8, bold=True, color=squid_ink)
                cell.fill = PatternFill(start_color=light_color, end_color=light_color, fill_type='solid')
            ws.row_dimensions[row].height = 16
            row += 1

            if len(data_df) == 0:
                ws[f'A{row}'] = "No exceptions"
                ws[f'A{row}'].font = Font(name='Calibri', size=9, italic=True, color='AAAAAA')
                row += 1
            else:
                for i, (_, emp) in enumerate(data_df.iterrows()):
                    emp_id = safe_int(emp['Employee ID'])
                    repeat_count = get_repeat_count(emp['Employee ID'], flag_type)
                    is_repeat = repeat_count > 0
                    ws.row_dimensions[row].height = 17

                    if is_repeat:
                        for c in range(1, 6):
                            ws.cell(row=row, column=c).fill = PatternFill(start_color=repeat_bg, end_color=repeat_bg, fill_type='solid')
                        ws.cell(row=row, column=1, value=emp_id).font = Font(name='Calibri', size=9, bold=True, color=repeat_red)
                        ws.cell(row=row, column=2, value=str(emp['Employee Name'])).font = Font(name='Calibri', size=9, bold=True, color=repeat_red)
                        ws.cell(row=row, column=3, value=str(emp['Dept'])).font = Font(name='Calibri', size=9, bold=True, color=repeat_red)
                        ws.cell(row=row, column=4, value=safe_int(emp['Total Break (min)'])).font = Font(name='Calibri', size=9, bold=True, color=repeat_red)
                        ws.cell(row=row, column=5, value=f"⚠️ {repeat_count + 1}x").font = Font(name='Calibri', size=9, bold=True, color=repeat_red)
                    else:
                        if i % 2 == 0:
                            for c in range(1, 6):
                                ws.cell(row=row, column=c).fill = PatternFill(start_color=light_color, end_color=light_color, fill_type='solid')
                        ws.cell(row=row, column=1, value=emp_id).font = Font(name='Calibri', size=9, color=dark_text)
                        ws.cell(row=row, column=2, value=str(emp['Employee Name'])).font = Font(name='Calibri', size=9, color=dark_text)
                        ws.cell(row=row, column=3, value=str(emp['Dept'])).font = Font(name='Calibri', size=9, color=dark_text)
                        ws.cell(row=row, column=4, value=safe_int(emp['Total Break (min)'])).font = Font(name='Calibri', size=9, bold=True, color=header_color)
                        ws.cell(row=row, column=5, value="").font = Font(name='Calibri', size=9)
                    row += 1

            ws.row_dimensions[row].height = 8
            row += 1
            return row

        row = 8
        row = write_table(ws, row, "EXCESS BREAK  >=65 min", coral, light_coral, excess, 'Excess')
        row = write_table(ws, row, "LESS BREAK  <=55 min", sunset, light_sunset, less, 'Less')
        row = write_table(ws, row, "MISSED BREAK PUNCH  0 min", charcoal, light_charcoal, missed, 'Missed')

        for c in range(1, 6):
            ws.cell(row=row, column=c).border = Border(top=Side(style='medium', color=teal))
        row += 1
        ws.row_dimensions[row].height = 20
        ws.merge_cells(f'A{row}:E{row}')
        ws[f'A{row}'] = "⚠️ Red highlighted rows = repeat offenders (flagged on previous days)"
        ws[f'A{row}'].font = Font(name='Calibri', size=8, italic=True, color='888888')

        output = BytesIO()
        wb.save(output)
        output.seek(0)
        return output

    # Download buttons
    st.markdown("### 📥 Downloads")
    col_dl1, col_dl2 = st.columns(2)

    with col_dl1:
        excel_file = generate_report_excel()
        st.download_button(
            label="📥 Excel Report",
            data=excel_file,
            file_name=f"Break_Compliance_{date.today().strftime('%Y-%m-%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    with col_dl2:
        history_csv = updated_history.to_csv(index=False).encode('utf-8')
        st.download_button(
            label="📋 Updated History",
            data=history_csv,
            file_name="history.csv",
            mime="text/csv",
            help="Save this and upload it next time!"
        )

    st.markdown("---")
    st.caption("💡 Download the updated history.csv each time and upload it next session!")

else:
    st.markdown("### 📋 How to use:")
    st.markdown("1. *(Optional)* Upload previous `history.csv` for repeat tracking")
    st.markdown("2. Upload your **attendance CSV**")
    st.markdown("3. View results + download report & updated history!")
    st.markdown("")
    st.markdown("---")
    st.markdown("**Criteria:**")
    st.markdown("- **Excess** = ≥65 min | **Less** = ≤55 min")
    st.markdown("- **Departments:** Inbound, Outbound, ICQA, Warehouse Deals, HR, IT, L&D, RME, Safety, Procurement")
    st.markdown("- **Repeat offenders** highlighted in red with count")
    st.markdown("- **Shift:** DXB3 | 08:00 - 18:00")

