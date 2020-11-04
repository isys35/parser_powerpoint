import win32com.client
import logging

PPT_PATH = 'D:\Project\commerce\parser_powerpoint\init_pptx\Fashion.pptx'

logger = logging.getLogger("parser_loger")
logger.setLevel(logging.INFO)
console_headler = logging.StreamHandler()
formatter = logging.Formatter('[%(levelname)s] %(asctime)s %(lineno)d : %(message)s')
console_headler.setFormatter(formatter)
logger.addHandler(console_headler)

# Types 17 - Text box
# Types 14 - Placeholder
# Types 16 - Media


class Slide:
    def __init__(self, slide_com_object):
        self.slide_com_object = slide_com_object
        self.id = self.slide_com_object.SlideIndex
        self.texts = self.get_texts()
        self.images = self.get_pictures()
        self.videos = self.get_videos()
        self.background = self.get_background()
        self.audio = self.get_audio()
        self.time_slide = self.get_time_slide()

    def get_texts(self) -> list:
        logger.info('Получение текстовой информации слайда №{}'.format(self.id))
        shapes = self.slide_com_object.Shapes
        text_frames = []
        for shape in shapes:
            if shape.Type == 17:
                text_frame = shape.TextFrame
                if text_frame.HasText:
                    text_frames.append(FrameText(text_frame))
        return text_frames

    def get_pictures(self) -> list:
        logger.info('Получение картинок слайда №{}'.format(self.id))
        shapes = self.slide_com_object.Shapes
        images = []
        for shape in shapes:
            if shape.Type == 14:
                images.append(ShapeImage(shape))
        return images

    def get_videos(self):
        logger.info('Получение видео слайда №{}'.format(self.id))
        shapes = self.slide_com_object.Shapes
        videos = []
        for shape in shapes:
            if shape.Type == 16:
                if shape.MediaType == 3:
                    videos.append(ShapeVideo(shape))
        return videos

    def get_audio(self):
        logger.info('Получение аудио слайда №{}'.format(self.id))
        shapes = self.slide_com_object.Shapes
        audios = []
        for shape in shapes:
            if shape.Type == 16:
                if shape.MediaType == 2:
                    audios.append(ShapeAudio(shape))
        return audios

    def get_background(self):
        logger.info('Получение фона слайда №{}'.format(self.id))
        return self.slide_com_object.Background.Fill

    def get_time_slide(self):
        logger.info('Получение продолжительности слайда №{}'.format(self.id))
        return self.slide_com_object.SlideShowTransition.Speed


class FrameText:
    def __init__(self, text_frame_com_object):
        self.text_frame_com_object = text_frame_com_object

    def __repr__(self):
        return self.text_frame_com_object.TextRange.Text


class ShapeImage:
    def __init__(self, shape_com_object):
        self.shape_com_object = shape_com_object


class ShapeVideo:
    def __init__(self, shape_com_object):
        self.shape_com_object = shape_com_object


class FillBackground:
    def __init__(self, fill_com_object):
        self.fill_com_object = fill_com_object


class ShapeAudio:
    def __init__(self, shape_com_object):
        self.shape_com_object = shape_com_object


class PPT:
    def __init__(self, ppt_path):
        logger.info('Открытие презентации {}'.format(ppt_path))
        self.app = win32com.client.Dispatch("Powerpoint.Application")
        self.ppt_path = ppt_path
        self.ppt_com_object = self.app.Presentations.Open(self.ppt_path)
        self.slides = self.get_slides()

    def get_slides(self) -> list:
        logger.info('Получение слайдов')
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
