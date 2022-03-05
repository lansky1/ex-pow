from asi import *
import asue
import ase


def GetPng(TARGET_FILE, IMAGE_FILE):
    ppt_file_path = os.getcwd() + '\\' + TARGET_FILE
    powerpoint = win32com.client.Dispatch('Powerpoint.Application')
    deck = powerpoint.Presentations.Open(ppt_file_path)

    img_file_path = os.getcwd() + '\\' + IMAGE_FILE
    powerpoint.ActivePresentation.Slides[0].Export(img_file_path, 'PNG')

    deck.Close()
    powerpoint.Quit()
    os.system('taskkill /F /IM POWERPNT.EXE')


def main(slideIndex):

    # Load the PPT and specify the slide
    root = Presentation(asue.PPT_TEMPLATE)
    slide = root.slides[0]  # here 0 is the slide number

    # Identify the components
    # for shape in slide.shapes:
    #     print(shape.shape_type)

    # Currently I have 7 AUTO_SHAPE, i.e. the colored shapes
    # and 10 TEXT_BOX in which 3 are for display - not to be edited

    shapes = slide.shapes
    text_box_list = []

    # Getting List of Text Box Shapes
    for shape_idx in range(len(shapes)):
        shape = shapes[shape_idx]
        if shape.shape_type == 17:
            text_box_list.append(shape_idx)

    # print(text_box_list)

    # Edit the Fields
    # Trial and Error
    # First Name
    shapes[7].text = ase.mName[slideIndex].split(" ", 1)[0]
    # Last Name
    shapes[8].text = ase.mName[slideIndex].split(" ", 1)[1]
    # Company
    # paragraph = shapes[9].text_frame.paragraphs[0]
    # paragraph.text = ase.mCompany[slideIndex]
    shapes[9].text = ase.mCompany[slideIndex]
    # Favorite Header
    # shapes[10].text
    # Enter_favs_here
    shapes[11].text = ase.mColor[slideIndex] + ' | ' + \
        ase.mMovie[slideIndex] + ' | ' + ase.mCurrency[slideIndex]
    # Hobbies Header
    # shapes[12].text
    # Enter_hobbies_here
    shapes[13].text = ase.mHobby[slideIndex]
    # Occupation
    shapes[14].text = ase.mOccupation[slideIndex]
    # Quote Header
    # shapes[15].text
    # Quote_here
    shapes[16].text = ase.mQuote[slideIndex]

    # Finally save the presentation
    TARGET_FILE_TITLE = asue.TARGET_FILE_NAME + \
        '_' + str(slideIndex+1) + '.pptx'
    root.save(TARGET_FILE_TITLE)


if __name__ == "__main__":
    for slideIndex in range(ase.mLength):
        main(slideIndex)

if asue.generatePngFiles:
    for file in os.listdir(os.getcwd()):
        if (file.endswith(".pptx") and file.startswith("FINAL")):
            # print(file, file.split(".")[0] + '.png')
            GetPng(file, file.split(".")[0] + '.png')
            if asue.deletePresentations:
                os.remove(os.path.join(os.getcwd(), file))
