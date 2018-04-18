import zipfile
import re
from lxml import etree

PREFIX = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
file_path = "10061167_刘牛顿_论文终稿.docx"
tempt_filename = "rules"
z = zipfile.ZipFile(file_path,"r")
doc = z.read("word/document.xml")
doct = etree.XML(doc)
sty = z.read("word/styles.xml")
styt = etree.XML(sty)
z.close()
body = doct.find(PREFIX + "body")
wT = open("wText.txt","w")
lT = open("pCat.txt","w")
tempt = {}
tf = open(tempt_filename,"r")
for line in tf:
    if line.startswith('{'):
        group=line[1:-3].split(',')
        for factor in group:
            _key = factor[:factor.index(':')]
            _val = factor[factor.index(':')+1:]
            if _key == 'key':
                rule_dkey = _val
                tempt.setdefault(_val,{})
            if _key!= 'key':
                tempt[rule_dkey].setdefault(_key,_val)
tf.close()
print(tempt)

def check_element_type(element,type):
    return element.tag == "%s%s"%(PREFIX,type)

def get_ptext(par):
    ptext = ""
    for t_ele in par.iter(tag=PREFIX+"t"):
        ptext += t_ele.text
    return ptext

def analyse(text):
    text=text.strip(' ')
    if text.isdigit():
        return 'body'
    pat1 = re.compile('[0-9]+')#以数字开头的正则表达式
    pat2 = re.compile('[0-9]+\\.[0-9]')#以X.X开头的正则表达式
    pat3 = re.compile('[0-9]+\\.[0-9]\\.[0-9]')#以X.X.X开头的正则表达式
    pat4 = re.compile('图(\s)*[0-9]+((\\.|-)[0-9])*')#图标题的正则表达式
    pat5 = re.compile('表(\s)*[0-9]+((\\.|-)[0-9])*')#表标题的正则表达式

    if pat1.match(text) and len(text)<70:
        if pat1.sub('',text)[0] == ' ':
            sort = 'firstLv'
        elif  pat1.sub('',text)[0] =='.':
            if pat2.match(text):
                if pat2.sub('',text)[0] == ' ':
                    sort = 'secondLv'
                elif pat2.sub('',text)[0]=='.':
                    if pat3.match(text):
                        if pat3.sub('',text)[0]==' ':
                            sort = 'thirdLv'
                        elif pat3.sub('',text)[0]=='.':
                            sort = 'overflow'
                            #print '    warning: 不允许出现四级标题！'
                        else:
                            sort ='thirdLv_e'
                    else:
                        sort='secondLv_e2'
                        #print '    warning: 二级标题正确的标号格式为X.X！'
                else:
                    sort = 'secondLv_e'
            else:
                sort = 'body'
        else:
            sort = 'firstLv_e'
    elif pat4.match(text) and len(text)<50:
        sort = 'objectT'
    elif pat5.match(text) and len(text)<50:
        sort = 'tableT'
    elif re.match(r'结 *论', text):
        sort = 'firstLv'
    elif re.match(r'致 *谢', text):
        sort = 'firstLv'
    elif re.match(r'绪 *论', text):
        sort = 'firstLv'
    else :
        sort ='body'
    return sort

def get_level(wp):
    for pPr in wp.iter(tag=PREFIX+"pPr"):
        for pPrc in pPr:
            if check_element_type(pPrc,"outlineLvl"):
                return pPrc.get("%s%s"%(PREFIX,"val"))
            if check_element_type(pPrc,"pStyle"):
                styleID = pPrc.get("%s%s"%(PREFIX,"val"))
                flag = True
                while flag:
                    flag = False
                    for style in styt.iter(tag=PREFIX+"style"):
                        if style.get("%s%s"%(PREFIX,"styleId"))==styleID:
                            for style_pPr in style:
                                if check_element_type(style_pPr,"pPr"):
                                    for outline_node in style_pPr.iter(tag = PREFIX+"outlineLvl"):
                                        return outline_node.get("%s%s"%(PREFIX,"val"))
                                if check_element_type(style_pPr,"baseOn"):
                                    styleID = style_pPr.get("%s%s"%(PREFIX,"val"))
                                    flag = True

def locate():
    pIndex = 0
    bigCat[1] = "cover"
    sCat[1] = "cover1"
    cur_par = 'cover'
    cur_state = "cover1"
    title = ""
    last_text = ""
    for par in body.iter(tag=PREFIX+'p'):
        pIndex += 1
        text = get_ptext(par)
        wT.write(str(pIndex))
        wT.write(" "+text+"\n")
        if text == "" or text == " ":
            continue
        if text == "论文封面书脊":
            cur_par = bigCat[pIndex] = "spine"
        elif text == "北京航空航天大学" :
            cur_par = bigCat[pIndex] = "taskbook"
        elif text == "本人声明":
            cur_par = bigCat[pIndex] = "statement"
        elif cur_par == "statement" and title in text:
            cur_par = bigCat[pIndex] = "abs"
        elif cur_par == "abs" and text.upper() == "ABSTRACT":
            cur_par = bigCat[pIndex] = "abs_en"
        elif re.match("目 *录",text):
            cur_par = bigCat[pIndex] = "menu"
        elif (cur_par == 'menu' and not text[-1].isdigit()) or (len(text)<15 and
                                                                     re.compile(r'.*绪 *论').match(text) and not text[-1].isdigit()):
            cur_par = bigCat[pIndex] = 'body'
        elif text == "参考文献":
            cur_par = bigCat[pIndex] = "refer"
        elif text.startswith("附录") and len(text)<15:
            cur_par = bigCat[pIndex] = "appdix"

        if cur_par == "cover":
            if "毕业设计" in text:
                cur_state = sCat[pIndex] = "cover2"
            elif cur_state == "cover2":
                cur_state = sCat[pIndex] = "cover3"
                title = text
            elif "院" in text and "系" in text and "名" in text and "称" in text:
                cur_state = sCat[pIndex] = "cover4"
            elif "年" in text and "月" in text:
                cur_state = sCat[pIndex] = "cover5"
        elif cur_par == "spine":
            cur_state = sCat[pIndex] = "spine"
        elif cur_par == "taskbook":
            cur_state = sCat[pIndex] = "taskbook"
        elif cur_par == "statement":
            if text == "本人声明":
                cur_state = sCat[pIndex] = "sta1"
            elif text.startswith("我声明"):
                cur_state = sCat[pIndex] = "sta2"
            elif "作者" in text:
                cur_state = sCat[pIndex] = "sta3"
        elif cur_par == "abs":
            if title in text:
                cur_state = sCat[pIndex] = "abs1"
            elif "生：" in text or "生:" in text:
                cur_state = sCat[pIndex] = "abs2"
            elif re.match("摘 *要",text):
                cur_state = sCat[pIndex] = "abs3"
                last_text = text
            elif re.match("摘 *要",last_text):
                cur_state = sCat[pIndex] = "abs4"
                last_text = ""
            elif "关键词" in text or "关键字" in text:
                cur_state = sCat[pIndex] = "abs5"
            elif cur_state == "abs5":
                cur_state = sCat[pIndex] = "abs1"
            elif "Author" in text:
                cur_state = sCat[pIndex] = "abs2"
        elif cur_par == "abs_en":
            if text.upper() == "ABSTRACT":
                cur_state = sCat[pIndex] = "abs3"
                last_text = "ABSTRACT"
            elif last_text == "ABSTRACT":
                cur_state = sCat[pIndex] = "abs4"
                last_text = ""
            elif ('KEY'in text or 'key' in text or "Key" in text) and ('WORD'in text or'word' in text or "Word" in text):
                cur_state = sCat[pIndex] = "abs5"
        elif cur_par == "menu":
            if re.match(r'目 *录', text) or re.compile(r'图 *目 *录').match(text) or re.compile(r'表 *目 *录').match(
                    text) or re.compile(r'图 *表 *目 *录').match(text):
                cur_state = sCat[pIndex] = 'menuTitle'
            elif analyse(text) in ['firstLv', 'firstLv_e']:
                cur_state = sCat[pIndex] = 'menuFirst'
            elif analyse(text) in ['secondLv', "secondLv_e", "secondLv_e2"]:
                cur_state = sCat[pIndex] = 'menuSecond'
            elif analyse(text) in ['thirdLv', "thirdLv_e"]:
                cur_state = sCat[pIndex] = 'menuThird'
            else:
                cur_state = sCat[pIndex] = 'menuFirst'  # 以汉字开头的标题都认为是一级标题
        elif cur_par == "body":
            level = get_level(par)
            analyse_result = analyse(text)
            if level == '0':
                cur_state = sCat[pIndex] = 'firstTitle'
                if analyse_result != 'firstLv' and analyse_result != 'firstLv_e':
                    # print 'warning',text
                    #warnInfo.append
                    print(pIndex,analyse_result,'warning: 标题级别和标题标号代表的级别不一致')
            elif level == '1':
                cur_state = sCat[pIndex] = 'secondTitle'
                if analyse_result != 'secondLv' and analyse_result != 'secondLv_e':
                    # print 'warning',text
                    #warnInfo.append\
                    print(pIndex,analyse_result,'warning: 标题级别和标题标号代表的级别不一致')
            elif level == '2':
                cur_state = sCat[pIndex] = 'thirdTitle'
                if analyse_result != 'thirdLv' and analyse_result != 'thirdLv_e':
                    # print 'warning',text
                    #warnInfo.append
                    print(pIndex,analyse_result,'warning: 标题级别和标题标号代表的级别不一致')
            else:
                if par.getparent().tag != '%s%s' % (PREFIX, 'body'):  # 当paragr的父节点不是body时，该para的文本不属于正文（可能是表格、图形或文本框内的文字
                    cur_state = sCat[pIndex] = 'tableText'
                elif analyse_result == 'firstLv':
                    cur_state = sCat[pIndex] = 'firstTitle'
                elif analyse_result == 'secondLv' or analyse_result == 'secondLv_e':
                    cur_state = sCat[pIndex] = 'secondTitle'
                elif analyse_result == 'thirdLv' or analyse_result == 'thirdLv_e':
                    cur_state = sCat[pIndex] = 'thirdTitle'
                elif analyse_result == 'objectT':
                    cur_state = sCat[pIndex] = 'objectTitle'
                elif analyse_result == 'tableT':
                    cur_state = sCat[pIndex] = 'tableTitle'
                elif re.match(r'结 *论', text):
                    cur_state = sCat[pIndex] = 'firstTitle'
                elif re.match(r'致 *谢', text):
                    cur_state = sCat[pIndex] = 'firstTitle'
                elif re.match(r'绪 *论', text):
                    cur_state = sCat[pIndex] = 'firstTitle'
                else:
                    cur_state = sCat[pIndex] = 'body'
        elif cur_par == 'refer':
            if text == '参考文献':
                cur_state = sCat[pIndex] = 'firstTitle'
            else:
                cur_state = sCat[pIndex] = 'reference'
        elif cur_par == 'appendix':
            if text.startswith('附') and text.endswith('录'):
                cur_state = sCat[pIndex] = 'firstTitle'
            else:
                cur_state = sCat[pIndex] = 'body'
    if not "spine" in bigCat.values():
        print("Warning:spine lost")
    if not "taskbook" in bigCat.values():
        print("Warning:taskbook lost")
    if not "statement" in bigCat.values():
        print("Warning:statement lost")
    if not "abs" in bigCat.values():
        print("Warning:abs lost")
    if not "abs_en" in bigCat.values():
        print("Warning:abs_en lost")
    if not "menu" in bigCat.values():
        print("Warning:menu lost")
    if not "body" in bigCat.values():
        print("Warning:body lost")
    if not "refer" in bigCat.values():
        print("Warning:refer lost")

def assign_fd(pPr,d):
    for node in pPr.iter(tag = etree.Element):
        if check_element_type(node,"rFonts"):
            att = node.get(PREFIX+"eastAsia")
            if att != None:
                d["fontCN"] = att
            att = node.get(PREFIX+"ascii")
            if att != None:
                d["fontEN"] = att
        elif check_element_type(node,"sz"):
            d["fontSize"] = node.get(PREFIX+"val")
        elif check_element_type(node,"b"):
            if "%s%s" % (PREFIX,"val") in node.keys():
                if node.get(PREFIX+'val') != '0' and node.get(PREFIX+ 'val') != 'false':
                    d["fontShape"] = '1'
                else:
                    d["fontShape"] = '0'
            else:
                d["fontShape"] = "1"
        elif check_element_type(node,"jc"):
            d["paraAlign"] = node.get(PREFIX+"val")
        elif check_element_type(node,"spacing"):
            if "%s%s" % (PREFIX,"line") in node.keys():
                d["paraSpace"] = node.get(PREFIX+"line")
            elif "%s%s" % (PREFIX,"before") in node.keys():
                d["paraFrontSpace"] = node.get(PREFIX + "before")
            elif "%s%s" % (PREFIX, "after") in node.keys():
                d["paraAfterSpace"] = node.get(PREFIX+"after")
        elif check_element_type(node,'ind'):#判断firstLine 和 firstLineChars的优先级
            if "%s%s"%(PREFIX,"firstLine") in node.keys():
                d['paraIsIntent']=node.get(PREFIX+'firstLine')
            if "%s%s"%(PREFIX,"firstLineChars") in node.keys():
                d['paraIsIntent']= node.get(PREFIX+'firstLineChars')
        elif check_element_type(node,'outlineLvl'):
            d['paraGrade'] = node.get(PREFIX+'val')

def get_default(d):
    for style in styt.iter(tag = PREFIX+"style"):
        if style.get(PREFIX+"default") == "1" and style.get(PREFIX+"type") == "paragraph":
            assign_fd(style,d)

def get_styleIdF(styleId,d):
    for style in styt.iter(tag=PREFIX+"style"):
        if style.get(PREFIX+"styleId") == styleId:
            for node in style.iter(tag=etree.Element):
                if check_element_type(node,"baseOn"):
                    id = node.get(PREFIX+"val")
                    get_styleIdF(id,d)
                if check_element_type(node,"pPr"):
                    assign_fd(node,d)

def get_format(p,d):
    get_default(d)
    pPr = p.find(PREFIX+"pPr")
    if pPr != None:
        pStyle = p.find(PREFIX+"pStyle")
        if pStyle != None:
            styleId = pStyle.get(PREFIX+"val")
            get_styleIdF(styleId,d)
        assign_fd(pPr,d)
    return

def check_out(sC,cur_format,pIndex,p):
    position = ['fontCN', 'fontEN', 'fontSize', 'fontShape', 'paraGrade', 'paraAlign', 'paraSpace', 'paraFrontSpace',
                'paraAfterSpace', 'paraIsIntent']
    # 这个字典的定义是为了避免对每个para都把规则字典里十个字段检查一遍，根据para的位置有选择有针对性的检查
    checkItemDct = {'cover1': ['fontCN', 'fontEN', 'fontSize', 'fontShape'],
                    'cover2': ['fontCN', 'fontSize', 'paraAlign', 'paraIsIntent'],
                    'cover3': ['fontCN', 'fontSize', 'paraAlign', 'paraIsIntent'],
                    'cover4': ['fontCN', 'fontSize', 'fontShape'],
                    'cover5': ['fontCN', 'fontSize', 'fontShape', 'paraAlign', 'paraIsIntent'],
                    'cover6': ['fontCN', 'fontSize', 'fontShape', 'paraAlign'],
                    'statm1': position,
                    'statm2': position,
                    'statm3': position,
                    'abstr1': position,
                    'abstr2': position,
                    'abstr3': position,
                    'abstr4': position,
                    'abstr5': position,
                    'abstr6': position,
                    'menuTitle': position,
                    'menuFirst': ['fontCN', 'fontSize', 'fontShape'],
                    'menuSecond': ['fontCN', 'fontSize', 'fontShape'],
                    'menuThird': ['fontCN', 'fontSize', 'fontShape'],
                    'firstTitle': position,
                    'secondTitle': position,
                    'thirdTitle': position,
                    'body': position,
                    'tableText': position,
                    'thankTitle': position,
                    'thankContent': position,
                    'extentTitle': position,
                    'extentContent': position,
                    'objectTitle': ['fontCN', 'fontEN', 'fontSize', 'fontShape', 'paraGrade', 'paraAlign',
                                    'paraIsIntent'],
                    'tableTitle': ['fontCN', 'fontEN', 'fontSize', 'fontShape', 'paraGrade', 'paraAlign',
                                   'paraIsIntent'],
                    'reference': position}
    if sC in checkItemDct.keys():
        if sC == "abstr5":
            return
    return

pIndex = 0
bigCat = {}
sCat = {}
p_format = {}.fromkeys(["fontCN","fontEN","fontSize","fontShape","paraAlign","paraSpace","paraIsIntent","paraFrontSpace","paraAfterSpace","paraGrade","leftChars","left"])
get_default(p_format)
print(p_format)
#print(doct.find("b"))
"""'fontCN':'中文字体',
'fontEN':'英文字体',
'fontSize':'字号',
'fontShape':'加粗',
'paraAlign':'对齐方式',
'paraSpace':'行间距',
'paraIsIntent':'首行缩进'
'paraFrontSpace':'段前间距',
'paraAfterSpace':'段后间距',
'paraGrade':"文本级别",
}"""
locate()

for p in body.iter(tag=PREFIX+"p"):
    pIndex += 1
    text = get_ptext(p)
    if text == " " or text=="":
        continue
    for k in p_format.keys():
        p_format[k] = None
    get_format(p,p_format)
    #print(pIndex,p_format)

for i in sCat.keys():
    lT.write(str(i)+" "+sCat[i]+"\n")
lT.close()
wT.close()
print(1)
#for ele in body.iter(tag=PREFIX+'p'):
#    print(get_ptext(ele))