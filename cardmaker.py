#MAGIC_SMALL_FONT="1/6in = .0625 = 4.5"
#MAGIC_MED_FONT="1/8in = .125 = 9"
#MAGIC_LARGE_FONT="3/16in = .1875 = 13.5"

import os
if not os.path.exists("output"):
    os.mkdir("output")

SMALL_FONT_SIZE=12
MED_FONT_SIZE=15
LARGE_FONT_SIZE=20

import pyexcel
data = pyexcel.get_book(file_name="cards.ods")
all_cards = []
for sheet_name in data.sheets:
    sheet = data.sheet_by_name(sheet_name)
    all_cards.extend(pyexcel.get_records(file_name="cards.ods",sheet_name=sheet_name,name_columns_by_row=0))
count_cards = sum([c.get("count",1) for c in all_cards])


import svgwrite
import textwrap

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
num_pages = count_cards//card_per_page

print(num_pages,len(all_cards),card_per_page,mm(pageheight*num_pages))
cardsheet = svgwrite.Drawing("output/all_cards.svg",size=(mm(pagewidth),mm(pageheight*num_pages)))

def make_img_pattern(root,img,width,height):
    pattern = root.pattern(id=img,patternUnits="userSpaceOnUse",
                    width=width,height=height)
    pattern.add(root.image("../images/%s"%img,x=0,y=0,width=width,height=height))
    root.add(pattern)
def init_patterns(root):
    make_img_pattern(root,"paper_texture.jpg",600,450)
    make_img_pattern(root,"sail_texture.jpg",253,355)
init_patterns(cardsheet)


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
def draw_card(root,card,x,y):
    if card["type"]=="pirate":
        draw_pirate(root,card,x,y)
    if card["type"]=="skill":
        draw_skill(root,card,x,y)
def draw_pirate(root,card,x,y):
    root.add(root.rect((mm(x),mm(y)),(mm(cardw),mm(cardh)),fill="url(#paper_texture.jpg)",stroke="black",rx=15,ry=15))
    root.add(root.text(card["name"],x=[mm(x+63/2)],y=[mm(y+6)],text_anchor="middle",font_size=LARGE_FONT_SIZE,font_family="'Source Sans Pro', sans-serif"))
    for i,trait in enumerate(card["traits"].split(",")):
        root.add(root.text(trait,x=[mm(x+63-25)],y=[mm(y+15+i*3)],font_size=SMALL_FONT_SIZE,font_family="'Source Sans Pro', sans-serif"))
    if "photo" in card:
        root.add(root.image("../images/"+card["photo"],x=mm(x+5),y=mm(y+7),size=(mm(28),mm(26))))
    if "description" in card:
        root.add(  root.rect((mm(x+5),mm(y+53)),
                                    (mm(63-5*2),mm(30)), 
                                    fill="rgb(255,255,255)",
                                    fill_opacity=0.4)  
        )
        root.add(wrapped_text(card["description"],x=mm(x+7),y=mm(y+57),columns=30,rows=6,font_size=MED_FONT_SIZE,font_family="'Source Sans Pro', sans-serif"))
def draw_skill(root,card,x,y):
    root.add(root.rect((mm(x),mm(y)),("63mm","88mm"),fill="url(#sail_texture.jpg)",stroke="black",rx=15,ry=15))
    root.add(root.text(card["name"],x=[mm(x+63/2)],y=[mm(y+6)],text_anchor="middle",font_size=LARGE_FONT_SIZE,font_family="'Source Sans Pro', sans-serif"))
    for i,trait in enumerate(card.get("traits","").split(",")):
        root.add(root.text(trait,x=[mm(x+63-25)],y=[mm(y+15+i*3)],font_size=SMALL_FONT_SIZE,font_family="'Source Sans Pro', sans-serif"))
    if "photo" in card:
        root.add(root.image("../images/"+card["photo"],x=mm(x+5),y=mm(y+7),size=(mm(28),mm(26))))
    if "description" in card:
        root.add(  root.rect((mm(x+5),mm(y+53)),
                                    (mm(63-5*2),mm(30)), 
                                    fill="rgb(255,255,255)",
                                    fill_opacity=0.4)  
        )
        root.add(wrapped_text(card["description"],x=mm(x+7),y=mm(y+57),columns=30,rows=6,font_size=MED_FONT_SIZE,font_family="'Source Sans Pro', sans-serif"))

def output_cards():
    page=0
    drawn_sheet1 = False
    sheet1 = svgwrite.Drawing("output/cards_page%.2d.svg"%page,size=(mm(pagewidth),mm(pageheight)))
    init_patterns(sheet1)
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
                    init_patterns(sheet1)
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
