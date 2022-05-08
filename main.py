from ppt_functions import *

if __name__ == '__main__':
  """Create new presentation"""
  ppt = new_presentation()

  """Pass ppt object to add_slide method to create new slide"""
  slide = add_slide(ppt)

  """Add rectangle to slide which is filled with blue color with Green boarder"""
  add_shape(slide, 'rectangle', 0, 0, 12, 1, (0, 0, 255), 0.02, (0, 255, 0))

  """Add text which is bold and italic with font size 20 and white color over above shape"""
  add_text(slide, 'I Love Python', left=4, top=0.3, bold=True, italic=True, font_size=20, color=(255, 255, 255))

  """Add Table with 3 rows and 4 columns"""
  table = add_table(slide, 3, 4, 0, 3, 12, 1)

  """Merge first row and add heading for table"""
  edit_table(table, 0, 0, 'Table', merge=True, to_row=0, to_column=3)

  "Format text in table"
  edit_table_style(table, 16, bold=True, underline=True, text_alignment='center', color=(0, 255, 0))

  """Add diamond, star and heart with no boarders"""
  add_shape(slide, 'diamond', 0, 4.5, 4, 4, (0, 0, 255))
  add_shape(slide, 'star', 4, 4.5, 4, 4, (0, 255, 0))
  add_shape(slide, 'heart', 8, 4.5, 4, 4, (255, 0, 255))

  "Save ppt"
  save_ppt(ppt, 'my_test.pptx')
