import re
from pptx.dml.color import RGBColor

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

def replaceTags(text, data): 
  tags = [x.group(0) for x in re.finditer(r"\$\{([A-Za-z0-9._\-]+)\}", ''.join(text))]

  for tag in tags:
    replacement = data.get(tag[2:-1])
    if not (replacement is None):
      text = text.replace(tag, replacement)
    else:
      text = text.replace(tag, '')
  
  return text

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
        style[line[0].strip()] = replaceTags(line[1], data)
  return style

def fill(placeholder, apply):
  """
  Apply background color
  """
  if 'solid' in apply:
    placeholder.fill.solid()

  color = re.search(r'#([0-9a-fA-F]{6})', apply)
  if color:
    placeholder.fill.fore_color.rgb = RGBColor.from_string(color.group(0)[1:])

def line(placeholder, apply):
  """
  Apply borders
  """
  color = re.search(r'#([0-9a-fA-F]{6})', apply)
  if color:
    placeholder.line.color.rgb = RGBColor.from_string(color.group(0)[1:])

def basicFormating(placeholder, style):
  """
  Apply style properties commun to all placeholders

  Args:
    placeholder
    style
  """
  # TODO: allow for 3 characters hex colors
  
  stylers = {
    'fill': fill,
    'line': line
  }

  for (rule, apply) in style.items():
    if rule in stylers:
      stylers.get(rule)(placeholder, apply)
  