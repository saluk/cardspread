#MAGIC_SMALL_FONT="1/6in = .0625 = 4.5"
#MAGIC_MED_FONT="1/8in = .125 = 9"
#MAGIC_LARGE_FONT="3/16in = .1875 = 13.5"
import math
import os
if not os.path.exists("output"):
    os.mkdir("output")
if not os.path.exists("output/cards"):
    os.mkdir("output/cards")
if not os.path.exists("output/decks"):
    os.mkdir("output/decks")
if not os.path.exists("output/tts_export"):
    os.mkdir("output/tts_export")
    
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
                props = {}
                if line[2]: props["cardw"] = line[2]
                if line[3]: props["cardh"] = line[3]
                templates[reading_templates] = {"props":props,"artifacts":[]}
                print("found template",reading_templates)
            elif line[0]=="variables":
                reading_mode = "variables"
            elif reading_mode=="templates":
                if not line[0]:
                    reading_mode = None
                    continue
                templates[reading_templates]["artifacts"].append(line)
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
    for card in all_cards:
        card["template"] = templates[card["type"]]
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
    try:
        with Image.open("images/%s"%img) as im:
            width,height = im.size
    except FileNotFoundError:
        width,height=0,0
    pattern = root.pattern(id=img,patternUnits="objectBoundingBox",
                    width=1,height=1)
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


def wrapped_text(text,x,y,w,h,**args):
    text_class = args["class_"]
    if "wrap_"+text_class not in settings:
        raise Exception("No wrap settings found for "+text_class)
    cw,ch = settings["wrap_"+text_class]
    columns = int(w/cw)
    vspace = ch+1.5
    rows = int(h/ch)
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
        text.add(context.tspan(line.strip(),x=[x],y=[y]))
        y=mm(conv_mm(y)+vspace)
    return text

def read_color(s):
    if type(s) == type("") and (s.endswith(".jpg") or s.endswith(".png")):
        return "texture",s
    if type(s) == type("") and "(" in s and ")" in s:
        return "color",tuple([int(x) for x in s[s.find("(")+1:s.find(")")].split(",")])
    if type(s) == type("") and s.startswith("#"):
        return "rawcolor",s
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
        y+=settings['cardh']
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
        element = wrapped_text(text,mm(offset[0]+x),mm(offset[1]+y),wrap[0],wrap[1],class_=text_class,text_anchor=anchor)
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
    if color_mode=="rawcolor":
        fill = color
    context.add(context.rect(
        (mm(offset[0]+x),mm(offset[1]+y)),
        (mm(w),mm(h)),
        fill=fill,
        stroke=stroke,
        rx=rx,ry=ry,
        opacity=opacity
    ))
def addimage(img,x,y,w,h,test_field,*a,**kwargs):
    if not img.strip(): return
    offset = kwargs["offset"]
    rx = read_x(x)
    ry = read_y(y)
    if test_field=="False":
        return
    settings["last_image_x"] = rx
    settings["last_image_y"] = ry
    context.add(context.image("../images/"+img,x=mm(offset[0]+rx),y=mm(offset[1]+ry),size=(mm(w),mm(h))))
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
    if s.startswith("{") and s.endswith("}"):
        result = eval(s[1:-1],card,settings)
        print ("EVAL",s[1:-1],"->",result)
        return str(result)
    if s.startswith("$"):
        return card[s[1:]]
    if s.startswith("!"):
        return card["template"]["props"].get(s[1:],settings.get(s[1:]))
    return s
    
def draw_card(root,card,x,y):
    global context
    context = root
    artifacts = card["template"]["artifacts"]
    for artifact in artifacts:
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
    
def output_tts(card,decks):
    pagewidth =card['cardw']*3
    if "[TTS]" not in card: return
    tts_type = card["[TTS]"]
    if "deck" in tts_type:
        deckname = tts_type.replace("deck_","")
        if deckname not in decks:
            drawing = svgwrite.Drawing("output/deck_%s.svg"%deckname,size=(mm(pagewidth),mm(card['cardh']*2)))
            init_sheet(drawing)
            decks[deckname] = {"face":"face","back":"back","drawing":drawing,"cards":[],"pen":[0,0],"lines":1,
                                        "width":3,"height":2}
        decks[deckname]["cards"].append(card)
        x,y = decks[deckname]["pen"]
        lines = decks[deckname]["lines"]
        draw_card(decks[deckname]["drawing"],card,x,y)
        x+=card['cardw']
        if x+card['cardw']>pagewidth:
            lines += 1
            x = 0
            y+=card['cardh']
            decks[deckname]["drawing"].attribs["height"] = mm(card['cardh']*lines)
        decks[deckname]["pen"] = [x,y]
        decks[deckname]["lines"] = lines
        decks[deckname]["height"] = max(lines,2)
        return
    with open("tts_exports/export_%s.json"%tts_type,"r") as f:
        tts_json = f.read()
    image_path = "file:///"+os.path.abspath("output/cards/%s.png"%card["name"]).replace("\\","/")
    with open("output/tts_export/%s.json"%card["name"],"w") as f:
        f.write(tts_json%{"face":image_path,"back":image_path,"name":card["name"]})
        
def output_tts_decks(decks):
    import cairosvg
    for deckname in decks:
        deck = decks[deckname]
        save_sheet(deck["drawing"])
        os.chdir("output")
        cairosvg.svg2png(bytestring=deck["drawing"].tostring().encode("utf8"),write_to="decks/%s.png"%deckname)
        os.chdir("..")
        face = "file:///"+os.path.abspath("output/decks/%s.png"%deckname).replace("\\","/")
        back = "file:///"+os.path.abspath("images/back_%s.png"%deckname).replace("\\","/")
        width = deck["width"]
        height = deck["height"]
        cards = []
        cardid = 100
        deckids = []
        with open("tts_exports/export_deck_card.json","r") as f:
            cardjson = f.read()
        for card in deck["cards"]:
            deckids.append(str(cardid))
            cards.append(cardjson%{"face":face,"back":back,"width":width,"height":height,"cardid":cardid})
            cardid += 1
        with open("tts_exports/export_deck.json","r") as f:
            json = f.read()
        with open("output/tts_export/deck_%s.json"%deckname,"w") as f:
            f.write(json%{"face":face,"back":back,"width":width,"height":height,
                "deckids":",".join(deckids),
                "cards":",".join(cards)})

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
    decks = {}
    init_sheet(sheet1)
    x,y=(0.0,0.0)
    last_type = ""
    for card in all_cards:
        for card_num in range(card.get("count",1)):

            if last_type and last_type != card["type"]:
                x,y=(0.0,0.0)
                page += 1
                save_sheet(sheet1)
                drawn_sheet1 = False
                sheet1 = svgwrite.Drawing("output/cards_page%.2d.svg"%page,size=(mm(pagewidth),mm(pageheight)))
                init_sheet(sheet1)
            last_type = card["type"]

            cardw = card['cardw'] = read_float(card["template"]["props"].get("cardw",settings["cardw"]))
            cardh = card['cardh'] = read_float(card["template"]["props"].get("cardh",settings["cardh"]))
            cardspacing = card['cardspacing'] = card["template"]["props"].get("cardspacing",settings["spacing"])

            draw_card(cardsheet,card,x,y+page*pageheight)
            draw_card(sheet1,card,x,y)
            drawn_sheet1 = True

            if OUTPUT_SVG:
                os.chdir("output")
                import cairosvg
                singlesheet = svgwrite.Drawing("output/card.svg",size=(mm(cardw),mm(cardh)))
                init_sheet(singlesheet)
                draw_card(singlesheet,card,0,0)
                cairosvg.svg2png(bytestring=singlesheet.tostring().encode("utf8"),write_to="cards/%s.png"%card["name"])
                os.chdir("..")
            output_tts(card,decks)

            x+=cardw+cardspacing
            if x+cardw>pagewidth:
                x=0
                y+=cardh+cardspacing
                if y+cardh>pageheight:
                    page += 1
                    y = 0
                    save_sheet(sheet1)
                    drawn_sheet1 = False
                    sheet1 = svgwrite.Drawing("output/cards_page%.2d.svg"%page,size=(mm(pagewidth),mm(pageheight)))
                    init_sheet(sheet1)
    save_sheet(cardsheet)
    if drawn_sheet1:
        save_sheet(sheet1)
    output_tts_decks(decks)
    

OUTPUT_SVG=1
if __name__ == "__main__":
    import sys
    cards = sys.argv[1]
    output_cards(cards)
    if OUTPUT_SVG:
        import cairosvg
        os.chdir("output")
        with open("all_cards.svg","rb") as f:
            bytes = f.read()
        cairosvg.svg2png(bytestring=bytes,write_to="all_cards.png")
