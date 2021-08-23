import turtle            # set up alex
wn = turtle.Screen()
alex = turtle.Turtle()

no_sides = int(input('Enter no of sides: '))
length = int(input('Enter length: '))
edge_color = input ('Enter edge color: ')
fill_color  = input ('Enter fill color: ')
alex.fillcolor(fill_color)
alex.begin_fill()
for i in range(no_sides):     # repeat four times
    alex.color(edge_color)
    alex.fillcolor(fill_color)
    alex.forward(length)
    alex.left(360/no_sides)
alex.end_fill()  

wn.exitonclick()

