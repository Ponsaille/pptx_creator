from .placeholders.text import formateTextPlaceholder, isTextBased
from .placeholders.image import isImageBased, formateImagePlaceholder

def formateSlide(slide, data):
  """
  Replace all placeholders in slide according to data

  Args:
    slide: the slide that needs to be edited
    data: a dictionnary object containing all the necessary elements

  Returns the slide element
  """

  for placeholder in slide.shapes.placeholders:

    if isTextBased(placeholder):
      formateTextPlaceholder(placeholder, data)
    
    elif isImageBased(placeholder):
      formateImagePlaceholder(placeholder, data)

  return slide

def generateSlide(presentation, config):
  """
  Creates a slide and replace all placeholders in slide according to data

  Args:
    presentation: the presentation
    config: a dictopnary object containing the config for the slide
  
  Returns the slide element
  """
  
  slideLayout = presentation.slide_layouts.get_by_name(config['slideLayout'])

  slide = presentation.slides.add_slide(slideLayout)

  return formateSlide(slide, config['data'])
