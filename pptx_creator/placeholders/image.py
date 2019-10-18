from pptx.parts.image import Image, ImagePart
from pptx.opc.constants import CONTENT_TYPE as CT, RELATIONSHIP_TYPE as RT
from pptx.oxml.shapes.picture import CT_Picture
from pptx.shapes.placeholder import PlaceholderPicture
import re
import base64

def isImageBased(placeholder):
  """
  Returns true is the placeholder is "image-based"

  Args:
    placeholder: PicturePlaceholder
  """
  return placeholder.placeholder_format.type._member_name in ('PICTURE')

def imagePartFromBlob(placeholder, blob):
  """
  Returns the image part corresponding to the image

  Args:
    placeholder: Picture Placeholder
    blob (bytes): blob image
  """

  # Check if the image doesn't already exists
  parts = placeholder.part.package._image_parts
  image = Image.from_blob(blob)

  imagePart = parts._find_by_sha1(image.sha1)

  if imagePart is None:
    # If not create a new ImagePart
    imagePart = ImagePart.new(parts._package, image)

  imagePart = ImagePart.new(parts._package, image)

  return imagePart

def imagePartAndrIdFromBlob(placeholder, blob):
  """
  Returns the imagePart and the rId corresponding to the image

  Args:
    placeholder: Picture Placeholder
    blob (bytes): blob image
  """
  imagePart = imagePartFromBlob(placeholder, blob)
  rId = placeholder.part.relate_to(imagePart, RT.IMAGE)

  return imagePart, rId

def insertBlobImage(placeholder, blob):
  """
  Inserts a blob image into a placeholder

  Args:
    placeholder: Picture Placeholder
    blob (bytes): blob image
  """
  # TODO: Multiple possibilities for resizing
  imagePart, rId = imagePartAndrIdFromBlob(placeholder, blob)
  desc, imageSize = imagePart.desc, imagePart._px_size

  shapeId, name = placeholder.shape_id, placeholder.name

  pic = CT_Picture.new_ph_pic(shapeId, name, desc, rId)

  # Equivalent to cover in css
  pic.crop_to_fit(imageSize, (placeholder.width, placeholder.height))

  placeholder._replace_placeholder_with(pic)

  return PlaceholderPicture(pic, placeholder._parent)

def formateImagePlaceholder(placeholder, data):
  """
  Place an image in the placeholder according to the first tag found

  Args:
    placeholder: Picture Placeholder
    blob (bytes): blob image
  
  Returns the PlaceholderPicture object of the image
  """
  tag = re.search(r"\$\{([A-Za-z0-9._\-]+)\}", placeholder._base_placeholder.text_frame.text).group(1)

  if(tag and data.get(tag)):
    blob = base64.decodestring(bytes(data.get(tag), 'utf-8'))
    return insertBlobImage(placeholder, blob)
