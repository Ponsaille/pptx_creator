import copy
import re

def copyBaseTextFrame(placeholder):
  """
  Copy the texts of the original placeholder insite the SlidePlaceholder

  Args:
    placeholder: Placeholder element from python_pptx
  """

  if(not placeholder._base_placeholder.has_text_frame or not placeholder.has_text_frame):
    raise TypeError('The placeholder is lacking a text_frame')

  placeholder.text_frame.clear() # Remove all elements from the actual text_frame and adding a blank paragraph

  # Copy all xml elements
  for pel in placeholder._base_placeholder.text_frame._element.p_lst:
    placeholder.text_frame._element.append(copy.deepcopy(pel))
  
  placeholder.text_frame.paragraphs[0]._element.delete() # Remove first blank paragraph

def findTagPosition(texts, tag, startPos, endPos):
  """
  Returns the start and end positions of a tag in a list of runs

  Args:
    runs: list of runs
    tag: tag in the form of ${tagName}
    startPos: start position of the tag
    endPost: end position of the tag
  
  Returns:
    A tuple in the form of:
    (
      (
        startRunIndex: index of the run were the tag begins
        startRunPos: index of the tag in the run
      ),
      (
        endRunIndex: index of the run were the tag ends
        endRunPos: index of the end of the tag in the run
      )
    )
  """

  if startPos < 0:
    return () # The tag doesn't exists

  for (runIndex, text) in enumerate(texts):
    length = len(text)

    if startPos >= 0:
      if startPos < length:
        startRunIndex = runIndex
        startRunPos = startPos

      startPos = startPos - length
    
    if endPos >= 0:
      if endPos <= length:
        endRunIndex = runIndex
        endRunPos = endPos
        break
      
      endPos = endPos - length
  
  return ((startRunIndex, startRunPos), (endRunIndex, endRunPos))

def getTagsInParagraph(texts):
  """
  Returns all the tags and their positions in the concatenation of texts

  Args:
    texts ([strings])
  
  Returns:
    [
      tagName: name of the tag (${tagName}),
      startPos: position were the tag begins in the concatenation of texts,
      endPos: position were the tag ends in the concatenation of texts
    ]
  """
  return [(x.group(1), x.start(), x.end()) for x in re.finditer(r"\$\{([A-Za-z0-9._\-]+)\}", ''.join(texts))]

def replaceTags(text_frame, data):
  """
  Replace all tags in a text_frame according to data

  Args:
    text_frame
    data (dictionnary) : {'tagName': 'replacementValue'}
  """
  for paragraph in text_frame.paragraphs:
    runs = paragraph.runs
    texts = [run.text for run in runs]
    tagNames = getTagsInParagraph(texts)

    for (tagName, startPos, endPos) in tagNames:
      tag = '${'+tagName+'}'
      replacement = str(data.get(tagName))

      ((startRunIndex, startRunPos), (endRunIndex, endRunPos)) = findTagPosition(texts, tag, startPos, endPos)

      if startRunIndex == endRunIndex:
        # The whole tag is in the startRun
        startRun = runs[startRunIndex]
        startRun.text = startRun.text.replace(tag, replacement)
      else:
        startRun = runs[startRunIndex]
        endRun = runs[endRunIndex]

        # Put the text in the first run
        startRun.text = startRun.text[0:startRunPos] + replacement
        # Keep the rest of the endRun
        endRun.text = endRun.text[endRunPos + 1:]

        # "Removing" the other runs
        for i in range(startRunIndex+1, endRunIndex):
          runs[i].text = ''

def isTextBased(placeholder):
  """
  Returns true is the placeholder is "text-based"
  """
  return placeholder.placeholder_format.type._member_name in ('BODY', 'CENTER_TITLE', 'SUBTITLE', 'TITLE', 'VERTICAL_BODY', 'VERTICAL_TITLE')

def formateTextPlaceholder(placeholder, data):
  """
  Executes all the steps to formate a "text-based" placeholder
  """
  copyBaseTextFrame(placeholder)

  replaceTags(placeholder.text_frame, data)