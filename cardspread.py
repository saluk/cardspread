#MAGIC_SMALL_FONT="1/6in = .0625 = 4.5"
#MAGIC_MED_FONT="1/8in = .125 = 9"
#MAGIC_LARGE_FONT="3/16in = .1875 = 13.5"
import math
import os
if not os.path.exists("output"):
    os.mkdir("output")
    
global context
settings = {"cardw":63,"cardh":88,"spacing":2}

import pyexcel

all_cards = []
templates = {}

def read_card_data(filename):
    data = pyexcel.get_book(file_name=filename)
    for sheet_name in data.sheets:
        sheet = data.sheet_by_name(sheet_name)
        
        reading_mode = None
        reading_templates = None
        for line in pyexcel.get_sheet(file_name="cards.ods",sheet_name=sheet_name).to_array():
            if line[0]=="template":
                reading_mode = "templates"
                reading_templates = line[1]
                templates[reading_templates] = []
                print("found template",reading_templates)
            elif line[0]=="variables":
                reading_mode = "variables"
            elif reading_mode=="templates":
                if not line[0]:
                    reading_mode = None
                    continue
                templates[reading_templates].append(line)
            elif reading_mode=="variables":
                if not line[0]:
                    reading_mode = None
                    continue
                print(line[0],line[1])
                settings[line[0]] = eval(str(line[1]))
                print(settings)
                
        for card in pyexcel.get_records(file_name="cards.ods",sheet_name=sheet_name,name_columns_by_row=0):
            if not type(card.get("count",""))==type(1):
                break
            all_cards.append(card)
    all_cards[:] = [card for card in all_cards if card["type"] in templates]
    return all_cards,templates

import svgwrite
import textwrap
from PIL import Image

def conv_mm(num):
    num,unit = svgwrite.utils.split_coordinate(num)
    factors = {"in":25.4,"cm":10,"mm":1}
    return num*factors[unit]
def mm(num):
    return "%fmm"%num

patterns = {}

def make_img_pattern(root,img):
    with Image.open("images/%s"%img) as im:
        width,height = im.size
    pattern = root.pattern(id=img,patternUnits="userSpaceOnUse",
                    width=width,height=height)
    pattern.add(root.image("../images/%s"%img,x=0,y=0,width=width,height=height))
    patterns[(root,img)] = pattern
    root.defs.add(pattern)
style_css=None
if os.path.exists("style.css"):
    #style_css = "style.css"
    with open("style.css") as f:
        style_css=f.read()
def init_style(root):
    if style_css:
        root.defs.add(root.style(style_css))
    textshadow = root.filter(id="textshadow",x="-20%",y="-20%",width="140%",height="140%")
    textshadow.feGaussianBlur("SourceAlpha",stdDeviation="4 4",result="shadow")
    textshadow.feOffset(dx="-1",dy="-1")
    root.defs.add(textshadow)
def init_sheet(sheet):
    init_style(sheet)


def wrapped_text(text,x,y,columns,rows,vspace=5,**args):
    lines = []
    for line in text.split("\n"):
        lines.append("")
        for wd in line.split(" "):
            spc = " "
            if len(lines[-1]+" "+wd)>columns:
                lines.append("")
                spc = ""
            lines[-1] = lines[-1]+spc+wd
    text = context.text("",x=[x],y=[y],**args)
    for line in lines[:rows]:
        text.add(context.tspan(line,x=[x],y=[y]))
        y=mm(conv_mm(y)+vspace)
    return text

def read_color(s):
    if type(s) == type("") and (s.endswith(".jpg") or s.endswith(".png")):
        return "texture",s
    if type(s) == type("") and "(" in s and ")" in s:
        return "color",tuple([int(x) for x in s[s.find("(")+1:s.find(")")].split(",")])
    return s
def read_float(s):
    if(type(s)==type("")):
        return float(eval(s))
    return s
def read_x(s):
    x = read_float(s)
    if x<0:
        x+=settings['cardw']
    return x
def read_y(s):
    y = read_float(s)
    if y<0:
        y+=cardh
    return y
def read_wrap(s):
    if type(s) == type("") and "(" in s and ")" in s:
        return tuple([int(x) for x in s[s.find("(")+1:s.find(")")].split(",")])
def read_shadow(s):
    if s=="shadow":
        return True
    return False
def read_text(s):
    return str(s).replace("<br>","\n")
def addtext(text,x,y,anchor="start",text_class="desc",wrap=False,shadow=False,*a,**kwargs):
    offset = kwargs["offset"]
    text = read_text(text)
    x=read_x(x)
    y=read_y(y)
    wrap=read_wrap(wrap)
    shadow=read_shadow(shadow)
    "filter:url(#textshadow)"
    if wrap:
        element = wrapped_text(text,mm(offset[0]+x),mm(offset[1]+y),wrap[0],wrap[1],wrap[2],class_=text_class,text_anchor=anchor)
        if shadow:
            context.add(wrapped_text(text,mm(offset[0]+x),mm(offset[1]+y),wrap[0],wrap[1],wrap[2],class_=text_class,text_anchor=anchor,
            style="filter:url(#textshadow);fill:black"))
    else:
        element = context.text(text,x=[mm(offset[0]+x)],y=[mm(offset[1]+y)],class_=text_class,text_anchor=anchor)
        if shadow:
            context.add(context.text(text,x=[mm(offset[0]+x)],y=[mm(offset[1]+y)],class_=text_class,text_anchor=anchor,
            style="filter:url(#textshadow);fill:black"))
    context.add(element)
def addrect(x,y,w,h,color_or_texture=None,stroke="",round=False,*a,**kwargs):
    offset = kwargs["offset"]
    color = None
    x=read_x(x)
    y=read_y(y)
    w=read_float(w)
    h=read_float(h)
    color_mode,color=read_color(color_or_texture)
    if round=="round": round = True
    rx=ry=0
    if round:
        rx=ry=15
    opacity=1
    fill="none"
    if color_mode=="color":
        if(type(color)==type((0,0))):
            fill = "rgb(%s,%s,%s)"%color[:3]
            if(len(color)>3):
                opacity=color[3]/255.0
    if color_mode=="texture":
        fill = "url(#%s)"%color
        if (context,color) not in patterns:
            make_img_pattern(context,color)
    context.add(context.rect(
        (mm(offset[0]+x),mm(offset[1]+y)),
        (mm(w),mm(h)),
        fill=fill,
        stroke=stroke,
        rx=rx,ry=ry,
        opacity=opacity
    ))
def addimage(img,x,y,w,h,*a,**kwargs):
    if not img.strip(): return
    offset = kwargs["offset"]
    x = read_x(x)
    y = read_y(y)
    context.add(context.image("../images/"+img,x=mm(offset[0]+x),y=mm(offset[1]+y),size=(mm(w),mm(h))))
svgpieces = []
def addsvg(svgpath,*args,**kwargs):
    thecard = kwargs["thecard"]
    offset = kwargs["offset"]
    #with open("svgs/"+svgpath) as f:
    #    xml = f.read().encode("utf8")
    #xml = xml[xml.find("<svg"):xml.find("</svg>")+6]

    import lxml.etree as ET
    tree = ET.parse("svgs/"+svgpath)
    root = tree.getroot()

    def gtag(node):
        return ET.QName(node.tag).localname

    def xml_substitute(text):
        if not text: return text
        for key in thecard:
            if "$"+key+";" in text:
                text = text.replace("$"+key+";",thecard[key])
        return text

    def xml_id_sub(node):
        for key in thecard:
            if node.get("id") == "$"+key:
                if gtag(node) == "text":
                    node.text = thecard[key]
                    for child in node:
                        node.remove(child)
        return node

    def process_node(node):
        node.text = xml_substitute(node.text)
        node = xml_id_sub(node)
        for child in node:
            pass
            process_node(child)
    process_node(root)

    xml = ET.tostring(root,encoding="unicode",pretty_print=True)
    #xml = xml[xml.find("<svg"):xml.find("</svg>")+6]
    svgpieces.append((offset[0]*3.56,offset[1]*3.56,xml))
    context.add(context.g(id="svgpiece_%d"%(len(svgpieces)-1)))
        
def substitute(card,s):
    if not type(s)==type(""):
        return s
    if s.startswith("$"):
        return card[s[1:]]
    if s.startswith("!"):
        return settings.get(s[1:])
    return s
    
def draw_card(root,card,x,y):
    global context
    context = root
    template = templates.get(card["type"],[])
    for artifact in template:
        try:
            func = eval("add"+artifact[0])
        except NameError:
            continue
        vals = artifact[1:][:]
        for i in range(len(vals)):
            vals[i] = substitute(card,vals[i])
        func(offset=[x,y],thecard=card,*vals)

def save_sheet(sheet):
    sheet.save()
    return
    f = open(sheet.filename,"r",encoding="utf8")
    xml = f.read()
    f.close()
    for i in range(len(svgpieces)):
        xml = xml.replace('<g id="svgpiece_%d" />'%i,"""
        <g id="svgpiece_%d" transform="translate(%d,%d)">%s</g>
        """%(i,svgpieces[i][0],svgpieces[i][1],svgpieces[i][2]))
    f = open(sheet.filename,"w")
    f.write(xml)
    f.close()

def output_cards(filename):
    read_card_data(filename)
    count_cards = sum([int(c.get("count",1)) for c in all_cards])
    print("TOTAL COUNT:\n",count_cards)
    pagewidth = conv_mm("8.5in")
    pageheight = conv_mm("11in")
    card_per_page = 0
    x=y=0
    while y+settings['cardh']<pageheight:
        while x+settings['cardw']<pagewidth:
            card_per_page += 1
            x+=settings['cardw']+settings['spacing']        
        x=0
        y+=settings['cardh']+settings['spacing']
    num_pages = math.ceil(count_cards/card_per_page)

    print(num_pages,count_cards,card_per_page,mm(pageheight*num_pages))
    cardsheet = svgwrite.Drawing("output/all_cards.svg",size=(mm(pagewidth),mm(pageheight*(num_pages))))
    init_sheet(cardsheet)

    page=0
    drawn_sheet1 = False
    sheet1 = svgwrite.Drawing("output/cards_page%.2d.svg"%page,size=(mm(pagewidth),mm(pageheight)))
    init_sheet(sheet1)
    x,y=(0.0,0.0)
    for card in all_cards:
        for card_num in range(card.get("count",1)):
            draw_card(cardsheet,card,x,y+page*pageheight)
            draw_card(sheet1,card,x,y)
            drawn_sheet1 = True
            x+=settings['cardw']+settings['spacing']
            if x+settings['cardw']>pagewidth:
                x=0
                y+=settings['cardh']+settings['spacing']
                if y+settings['cardh']>pageheight:
                    page += 1
                    y = 0
                    save_sheet(sheet1)
                    drawn_sheet1 = False
                    sheet1 = svgwrite.Drawing("output/cards_page%.2d.svg"%page,size=(mm(pagewidth),mm(pageheight)))
                    init_sheet(sheet1)
    save_sheet(cardsheet)
    if drawn_sheet1:
        save_sheet(sheet1)

OUTPUT_SVG=0
if __name__ == "__main__":
    import sys
    cards = sys.argv[1]
    output_cards(cards)
    if OUTPUT_SVG:
        import cairosvg
        os.chdir("output")
        with open("all_cards.svg","r",encoding="utf8") as f:
            cairosvg.svg2png(file_obj=f,write_to="all_cards.png")
