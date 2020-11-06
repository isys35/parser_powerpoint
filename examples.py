from powerpoint import PPT, PPTColors

PPT_PATH = r'D:\Projects\commerce\parser_powerpoint\init_pptx\Fashion.pptx'  # исходник
PPT_CHANGED_PATH = r'D:\Projects\commerce\parser_powerpoint\init_pptx\new.pptx'  # изменённый файл

"""
Для изменения цветов используется другой модуль,
т.е отрыть презентацию и одновременно изменить и текст и цвет не получится,
но можно открыть перзентацию изменить текст, сохранить, потом снова открыть и изменить цвета
"""


def show_info():
    """пример который показывает информацию о презентации"""
    ppt = PPTColors(PPT_CHANGED_PATH)
    ppt.show_info()


def change_background_color():
    """пример изменения цвета фона 1 слайда на красный"""
    ppt = PPTColors(PPT_PATH)
    red = (255, 0, 0)
    ppt.slides[0].change_background_color(red)
    ppt.save_as(PPT_CHANGED_PATH)


def change_shape_color():
    """пример изменения цвета формы на 3 слайде 4 формы на зелёный"""
    ppt = PPTColors(PPT_PATH)
    green = (0, 255, 0)
    ppt.slides[2].shapes[3].change_color(green)
    ppt.save_as(PPT_CHANGED_PATH)


def duplicate_slide():
    """Пример копирование 3-го слайда и вставки его в 5-ую позицию"""
    ppt = PPT(PPT_PATH)
    ppt.duplicate_slide(3, 5)
    ppt.save_as(PPT_CHANGED_PATH)
    ppt.close()


def change_text_example():
    """пример изменения пераого текстового блока в 3 слайде"""
    ppt = PPT(PPT_PATH)
    ppt.slides[2].texts[0].change_text('ИЗМЕНЁННЫЙ ТЕКСТ')
    ppt.save_as(PPT_CHANGED_PATH)
    ppt.close()


def change_image_example():
    """пример изменения первого изображения блока в 2 слайде"""
    ppt = PPT(PPT_PATH)
    image_path = r'D:\Project\commerce\parser_powerpoint\image.jpg'
    ppt.slides[1].images[0].change_image(image_path)
    ppt.save_as(PPT_CHANGED_PATH)
    ppt.close()


def change_audio_example():
    """пример изменения первого аудио в 1 слайде"""
    ppt = PPT(PPT_PATH)
    audio_path = r'D:\Project\commerce\parser_powerpoint\audio.mp3'
    ppt.slides[0].audio[0].change_audio(audio_path)
    ppt.save_as(PPT_CHANGED_PATH)
    ppt.close()


def change_video_example():
    """пример изменения первого видео в 2 слайде"""
    ppt = PPT(PPT_PATH)
    video_path = r'D:\Project\commerce\parser_powerpoint\video.avi'
    ppt.slides[1].videos[0].change_video(video_path)
    ppt.save_as(PPT_CHANGED_PATH)
    ppt.close()


def change_speed_example():
    """пример изменения скорости 1 слайда"""
    """
    speed_id:
        1 - медленно
        2 - средне
        3 - быстро
    """
    ppt = PPT(PPT_PATH)
    speed_id = 1
    ppt.slides[0].change_speed(speed_id)
    ppt.save_as(PPT_CHANGED_PATH)
    ppt.close()


if __name__ == '__main__':
    change_shape_color()
    show_info()
