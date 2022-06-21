class MyCell:
    def __init__(self, value=None, font=None):
        self.value = value
        self.font = self.setFont(font)

    def setFont(self, font):
        if font is None:
            return None
        return MyFont(bold=font.bold, color=font.color, italic=font.italic, name=font.name, size=font.size)


class MyFont:
    def __init__(self, bold=False, color=(0, 0, 0), italic=False, name='맑은 고딕', size=11.0):
        self.bold = bold
        self.color = color
        self.italic = italic
        self.name = name
        self.size = size
