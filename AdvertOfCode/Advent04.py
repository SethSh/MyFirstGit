with open ("Day4Data.txt") as f:
    drawn = [int(x) for x in f.readline().strip('/n').split(',')]
    cards =[] 
    print(drawn)

    while f.readline():
        card = []
        for i in range(5):
            card.extend([int(x) for x in f.readline().strip('/n').split(' ') if x != ''])
        print(card)
        cards.append(card)
