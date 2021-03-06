import sys

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


def cleanCopyExcel(fileName='input.xlsx', MAX_X=50, MAX_Y=1000, logF=None):
    if logF is None:
        logF = lambda _x: print(_x)

    logF(fileName + ' 의 작업을 시작합니다.')
    app = xw.App(visible=False)
    book = xw.Book(fileName)
    oBook = xw.Book()

    for inSh in book.sheets:
        logF(str(inSh.name) + ' 시트의 작업을 시작합니다.')
        names = list(map(lambda sh: sh.name, oBook.sheets))
        if inSh.name in names:
            oSh = oBook.sheets[inSh.name]
        else:
            oSh = oBook.sheets.add()
            oSh.name = inSh.name

        cellValues = inSh.range((1, 1), (MAX_Y, MAX_X)).formula
        oSh.range((1, 1), (MAX_Y, MAX_X)).formula = cellValues
        for x in range(1, MAX_X + 1):
            if x%10 == 0:
                logF(str(x) + ' 번 째 세로줄 작업중입니다.')
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


def main():
    kwargs = {}
    args = list(sys.argv)
    com = ''
    for arg in args:
        if com == '-in' or com == '-i':
            kwargs['fileName'] = arg
        elif com == '-x' or com == '-X':
            kwargs['MAX_X'] = int(arg)
        elif com == '-y' or com == '-Y':
            kwargs['MAX_Y'] = int(arg)
        com = arg

    cleanCopyExcel(**kwargs)


if __name__ == '__main__':
    main()
