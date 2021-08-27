from wand.image import Image
import glob
from pptx import Presentation
from pptx.util import Inches


def ppt_creation():
    j = 0
    prs = Presentation()
    for image_path_name in glob.glob("output/*.jpg"):
        j = j+1
        image_path = image_path_name
        blank_slide_layout = prs.slide_layouts[1]
        slide = prs.slides.add_slide(blank_slide_layout)

        shapes = slide.shapes

        title_shape = shapes.title
        body_shape = shapes.placeholders[1]

        title_shape.text = f'Sample title {j}'

        tf = body_shape.text_frame
        tf.text = f'Sample subtitle {j}'

        bottom = Inches(2.5)
        left = Inches(1)
        height = Inches(4)
        image_p = f"{image_path}"
        slide.shapes.add_picture(image_p, left, bottom, height=height)
    prs.save("ppt/demo.pptx")



def logo_placement(image_val, i):

    logo = Image(filename="image/nike_black.png")
    image = Image(image_val)
    logo_width = int(image.width/2)
    logo_height = int(image.height/5)
    logo.resize(logo_width,logo_height)
    image.composite_channel("all_channels", logo, "dissolve", 70, 30)
    image.save(filename=f"output/image{i}.jpg")
    ppt_creation()


def get_image():
    i = 0
    for image_name in glob.glob("image/*.jpg"):
        image_val = Image(filename=image_name)
        i = i+1
        logo_placement(image_val,i)


if __name__ == "__main__":
    get_image()




