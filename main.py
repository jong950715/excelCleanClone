import xlwings as xw

from MyCell import MyCell


def _getRanges(_max, _once):
    for x in range(1, _max, _once):
        x1 = x
        x2 = x + _once - 1
        x2 = x2 if x2 < _max else _max
        yield x1, x2


def getRanges(maxX, maxY, maxOnce):
    xs = list(_getRanges(maxX, maxOnce))
    xy = list(_getRanges(maxY, maxOnce))
    for x in xs:
        for y in xy:
            yield (x, y)


def _cleanCopy(inRange, outRange):
    # outRange.value = inRange.value
    outRange.color = inRange.color
    outRange.font.bold = inRange.font.bold
    try:
        outRange.font.color = inRange.font.color
    except Exception as e:
        print(inRange.address, e)
    outRange.font.italic = inRange.font.italic
    # outRange.font.name = inRange.font.name
    outRange.font.size = inRange.font.size


MAX_X = 30
MAX_Y = 850


def main():
    # app = xw.App(visible=False)
    # sBook = xw.Book('samp.xlsx')
    # oBook = xw.Book()
    #
    # # r = sBook.sheets[0].range((1, 1), (1, 3)).formula
    # inSht = sBook.sheets[0]
    # oSht = oBook.sheets[0]
    #
    # oSht.range((1,1),(20,20)).formula = inSht.range((1,1),(20,20)).formula
    #
    # oBook.save()
    # app.kill()
    #
    # return
    app = xw.App(visible=False)
    book = xw.Book('input.xlsx')
    oBook = xw.Book()

    for inSh in book.sheets:
        print(inSh.name)
        names = list(map(lambda sh: sh.name, oBook.sheets))
        if inSh.name in names:
            oSh = oBook.sheets[inSh.name]
        else:
            oSh = oBook.sheets.add()
            oSh.name = inSh.name

        cellValues = inSh.range((1, 1), (MAX_Y, MAX_X)).formula
        oSh.range((1, 1), (MAX_Y, MAX_X)).formula = cellValues
        for x in range(1, MAX_X + 1):
            for y in range(1, MAX_Y + 1):
                if cellValues[y - 1][x - 1] == '':
                    continue
                try:
                    _cleanCopy(inSh.range(y, x), oSh.range(y, x))
                except Exception as e:
                    print(x, y, e)

        for x in range(1, MAX_X + 1):
            oSh.range(1, x).column_width = inSh.range(1, x).column_width
        for y in range(1, MAX_Y + 1):
            oSh.range(y, 1).row_height = inSh.range(y, 1).row_height

    oBook.save()
    app.kill()

if __name__ == '__main__':
    main()
