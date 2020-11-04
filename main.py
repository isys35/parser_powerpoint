import win32com.client

PPT_PATH = 'D:\Projects\commerce\parser_powerpoint\init_pptx\Fashion.pptx'

# Types 17 - Text box
# Types 14 - Placeholder


class Slide:
    def __init__(self, slide_com_object):
        self.slide_com_object = slide_com_object
        self.id = self.slide_com_object.SlideIndex
        self.texts = self.get_texts()
        self.images = self.get_pictures()

    def get_texts(self) -> list:
        shapes = self.slide_com_object.Shapes
        text_frames = []
        for shape in shapes:
            if shape.Type == 17:
                text_frame = shape.TextFrame
                if text_frame.HasText:
                    # print(text_frame.TextRange)
                    # print(shape.Type)
                    text_frames.append(SlideText(text_frame))
        return text_frames

    def get_pictures(self) -> list:
        shapes = self.slide_com_object.Shapes
        images = []
        for shape in shapes:
            if shape.Type == 14:
                print(shape.PictureFormat.Creator)



class SlideText:
    def __init__(self, text_frame_com_object):
        self.text_frame_com_object = text_frame_com_object

    def __repr__(self):
        return self.text_frame_com_object.TextRange.Text


class PPT:
    def __init__(self, ppt_path):
        self.app = win32com.client.Dispatch("Powerpoint.Application")
        self.ppt_path = ppt_path
        self.ppt_com_object = self.app.Presentations.Open(self.ppt_path)
        self.slides = self.get_slides()

    def get_slides(self) -> list:
        return [Slide(slide_com_object) for slide_com_object in self.ppt_com_object.Slides]

    def close(self):
        self.app.Quit()


def main():
    ppt = PPT(PPT_PATH)
    ppt.close()
    # ppt_app = win32com.client.Dispatch("Powerpoint.Application")
    # ppt_presentation = ppt_app.Presentations.Open('D:\Projects\commerce\parser_powerpoint\init_pptx\Fashion.pptx')
    # ppt_slides = ppt_presentation.Slides
    # for ppt_slide in ppt_slides:
    #     ppt_slide_shapes = ppt_slide.Shapes
    #     for ppt_slide_shape in ppt_slide_shapes:
    #         if ppt_slide_shape.HasTextFrame:
    #             text_frame = ppt_slide_shape.TextFrame
    #             if text_frame:
    #                 if text_frame.HasText:
    #
    #     print('')
    # ppt_app.Quit()

if __name__ == '__main__':
    main()