import re

def getStylingAreasPosInTextFrame(texts):
  return re.finditer(r"\/\*\*(.*?)\*\*\/", ''.join(texts))

def getStylingAreaPos(texts, startPos, endPos):
  if startPos < 0:
    return () # The ara doesn't exists

  for (paragraphIndex, text) in enumerate(texts):
    length = len(text)

    if startPos >= 0:
      if startPos < length:
        startParagraphIndex = paragraphIndex

      startPos = startPos - length
    
    if endPos >= 0:
      if endPos <= length:
        endParagraphIndex = paragraphIndex
        break
      
      endPos = endPos - length
  
  return (startParagraphIndex, endParagraphIndex)

def getAllStylingAreas(texts):
  return [getStylingAreaPos(texts, area.start(), area.end()) for area in getStylingAreasPosInTextFrame(texts)]


def removeStylingAreas(paragraphs, stylingAreas):
  toRemove = []
  # Getting the paragraphs to remove
  for stylingArea in stylingAreas:
    for i in range(stylingArea[0], stylingArea[1]+1):
      toRemove.append(paragraphs[i])
  
  for paragraph in toRemove:
    paragraph._element.delete()

def parser(paragraphs, data):
  """
  Parse paragraphs to return all the style in a unique 
  """
  texts = [paragraph.text for paragraph in paragraphs]
  stylingAreas = getStylingAreasPosInTextFrame(texts)

  stylingAreas = [stylingArea.group(1).split(';') for stylingArea in stylingAreas]
  lines = [line.split(':') for sublist in stylingAreas for line in sublist]

  style = {}

  for line in lines:
    if(len(line) == 2):

      replacement = data.get(line[1])
      if not (replacement is None):
        style[line[0]] = data.get(line[1])

  return style