pirates = [
        {"name":"Smackface Bartleby","type":"pirate","Speed":2,"traits":"","photo":"smackface.jpg","number":1,"description":"A lousy pirate man"},
        {"name":"Farseer Bornblind","type":"pirate","Speed":3,"traits":"blind","number":1},
        {"name":"Some Farmer dude","type":"pirate","Speed":8,"traits":"blind,weak,aquaphobia","number":1}
        ]

all_cards = pirates



import svgwrite

def conv_mm(num):
    num,unit = svgwrite.utils.split_coordinate(num)
    factors = {"in":25.4,"cm":10}
    return num*factors[unit]
def mm(num):
    return "%dmm"%num

width = conv_mm("8.5in")
height = conv_mm("11in")
cardsheet = svgwrite.Drawing("pirates.svg",size=("%dmm"%width,"%dmm"%height))
def draw_card(card,x,y):
    if card["type"]=="pirate":
        draw_pirate(card,x,y)
def draw_pirate(card,x,y):
    cardsheet.add(cardsheet.rect((mm(x),mm(y)),("63mm","88mm"),fill="white",stroke="black"))
    cardsheet.add(cardsheet.text(card["name"],x=[mm(x+63/2)],y=[mm(y+5)],text_anchor="middle"))
    for i,trait in enumerate(card["traits"].split(",")):
        cardsheet.add(cardsheet.text(trait,x=[mm(x+63-25)],y=[mm(y+15+i*3)]))
    if "photo" in card:
        cardsheet.add(cardsheet.image("images/"+card["photo"],x=mm(x+5),y=mm(y+7),size=(mm(28),mm(26))))
    if "description" in card:
        cardsheet.add(cardsheet.textArea(trait,x=[mm(x+7)],y=[mm(y+57)],size=(mm(46),mm(22))))

def output_cards():
    x,y=(0,0)
    for card in all_cards:
        for card_num in range(card.get("number",1)):
            draw_card(card,x,y)
            x+=65
            if x+63>conv_mm("8.5in"):
                x=0
                y+=90
    cardsheet.save()

if __name__ == "__main__":
    import sys
    if len(sys.argv)==3:
        object_pos = float(sys.argv[1])
        object_height = float(sys.argv[2])
        print(conv_mm("11in")-object_pos-object_height)
    else:
        output_cards()
