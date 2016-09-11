#MAGIC_SMALL_FONT="1/6in = .0625 = 4.5"
#MAGIC_MED_FONT="1/8in = .125 = 9"
#MAGIC_LARGE_FONT="3/16in = .1875 = 13.5"
import math
import os
if not os.path.exists("output"):
    os.mkdir("output")
    
global context

import pyexcel
data = pyexcel.get_book(file_name="cards.ods")
all_cards = []
templates = {}
for sheet_name in data.sheets:
    sheet = data.sheet_by_name(sheet_name)
    
    reading_templates = False
    for template_line in pyexcel.get_sheet(file_name="cards.ods",sheet_name=sheet_name).to_array():
        if template_line[0]=="template":
            reading_templates = template_line[1]
            templates[reading_templates] = []
            print("found template",reading_templates)
        elif reading_templates:
            if not template_line[0]:
                break
            templates[reading_templates].append(template_line)
            
    for card in pyexcel.get_records(file_name="cards.ods",sheet_name=sheet_name,name_columns_by_row=0):
        if not type(card.get("count",""))==type(1):
            break
        all_cards.append(card)
print(templates)
print(all_cards)
count_cards = sum([int(c.get("count",1)) for c in all_cards])
print(count_cards)

import svgwrite
import textwrap
from PIL import Image

def conv_mm(num):
    num,unit = svgwrite.utils.split_coordinate(num)
    factors = {"in":25.4,"cm":10,"mm":1}
    return num*factors[unit]
def mm(num):
    return "%dmm"%num

cardw = 63
cardh = 88
spacing = 2
pagewidth = conv_mm("8.5in")
pageheight = conv_mm("11in")
card_per_page = 0
x=y=0
while y+cardh<pageheight:
    while x+cardw<pagewidth:
        card_per_page += 1
        x+=cardw+spacing 
    x=0
    y+=cardh+spacing
num_pages = math.ceil(count_cards/card_per_page)

print(num_pages,count_cards,card_per_page,mm(pageheight*num_pages))
cardsheet = svgwrite.Drawing("output/all_cards.svg",size=(mm(pagewidth),mm(pageheight*(num_pages))))
patterns = {}

def make_img_pattern(root,img):
    with Image.open("images/%s"%img) as im:
        width,height = im.size
    pattern = root.pattern(id=img,patternUnits="userSpaceOnUse",
                    width=width,height=height)
    pattern.add(root.image("../images/%s"%img,x=0,y=0,width=width,height=height))
    patterns[(root,img)] = pattern
    root.add(pattern)
style_css=None
if os.path.exists("style.css"):
    with open("style.css") as f:
        style_css=f.read()
def init_style(root):
    if style_css:
        root.defs.add(root.style(style_css))
def init_sheet(sheet):
    init_style(sheet)
init_sheet(cardsheet)


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
    text = cardsheet.text("",x=[x],y=[y],**args)
    for line in lines[:rows]:
        text.add(cardsheet.tspan(line,x=[x],y=[y]))
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
        x+=cardw
    return x
def read_y(s):
    y = read_float(s)
    if y<0:
        y+=cardh
    return y
def read_wrap(s):
    if type(s) == type("") and "(" in s and ")" in s:
        return tuple([int(x) for x in s[s.find("(")+1:s.find(")")].split(",")])
def read_text(s):
    return s.replace("<br>","\n")
def addtext(text,x,y,anchor="start",text_class="desc",wrap=False,*a,**kwargs):
    offset = kwargs["offset"]
    text = read_text(text)
    x=read_x(x)
    y=read_y(y)
    wrap=read_wrap(wrap)
    if wrap:
        element = wrapped_text(text,mm(offset[0]+x),mm(offset[1]+y),wrap[0],wrap[1],wrap[2],class_=text_class,text_anchor=anchor)
    else:
        element = context.text(text,x=[mm(offset[0]+x)],y=[mm(offset[1]+y)],class_=text_class,text_anchor=anchor)
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
    offset = kwargs["offset"]
    x = read_x(x)
    y = read_y(y)
    context.add(context.image("../images/"+img,x=mm(offset[0]+x),y=mm(offset[1]+y),size=(mm(w),mm(h))))
        
def substitute(card,s):
    if not type(s)==type(""):
        return s
    if s.startswith("$"):
        return card[s[1:]]
    if s.startswith("!"):
        return eval(s[1:])
    return s
    
def draw_card(root,card,x,y):
    global context
    context = root
    template = templates[card["type"]]
    for artifact in template:
        func = eval("add"+artifact[0])
        vals = artifact[1:][:]
        for i in range(len(vals)):
            vals[i] = substitute(card,vals[i])
        func(offset=[x,y],*vals)

def output_cards():
    page=0
    drawn_sheet1 = False
    sheet1 = svgwrite.Drawing("output/cards_page%.2d.svg"%page,size=(mm(pagewidth),mm(pageheight)))
    init_sheet(sheet1)
    x,y=(0,0)
    for card in all_cards:
        for card_num in range(card.get("count",1)):
            draw_card(cardsheet,card,x,y+page*pageheight)
            draw_card(sheet1,card,x,y)
            drawn_sheet1 = True
            x+=cardw+spacing
            if x+cardw>pagewidth:
                x=0
                y+=cardh+spacing
                if y+cardh>pageheight:
                    page += 1
                    y = 0
                    sheet1.save()
                    drawn_sheet1 = False
                    sheet1 = svgwrite.Drawing("output/cards_page%.2d.svg"%page,size=(mm(pagewidth),mm(pageheight)))
                    init_sheet(sheet1)
    cardsheet.save()
    if drawn_sheet1:
        sheet1.save()

if __name__ == "__main__":
    import sys
    if len(sys.argv)==3:
        object_pos = float(sys.argv[1])
        object_height = float(sys.argv[2])
        print(conv_mm("11in")-object_pos-object_height)
    else:
        output_cards()
