import ctypes
import datetime
import os
import platform
import random
import calendar

import subprocess
# pip install wget
import wget       

# from the pillow library - pip install pillow
from PIL import Image, ImageDraw, ImageFont

import config

def execute_set(command):
    subprocess.call(["/bin/bash", "-c", command])

def set_wallpaper(wallpaper_name='\output.jpg'):
    cwd = os.getcwd()
    if platform.system() == 'Windows':
        ctypes.windll.user32.SystemParametersInfoW(
            20, 0, cwd+wallpaper_name, 0)
    elif platform.system() == 'Linux':
        execute_set("/usr/bin/gsettings set org.gnome.desktop.background picture-uri " + cwd + wallpaper_name )
    else:
        input('Your operating system is not supported')


def create_wallpaper():
    # Load font
    fontFile = 'fonts/{}'.format(config.default_font)
    font = ImageFont.truetype(fontFile, config.font_size)
    # Load Image
    wallpaper_name = get_wallpaper()
    image = Image.open(wallpaper_name)
    draw = ImageDraw.Draw(image)
    # Gregorian Datetime
    if config.write_gregorian_datetime:
        gtoday = str(datetime.datetime.now().date())
        w, h = draw.textsize(text=gtoday, font=font)
        # draw.text((image.width-w-config.position_for_gdatetime[1], image.height-h-config.position_for_gdatetime[1]),
        #           gtoday, (255, 255, 255), font=font)
        draw.text((image.width-w-config.position_for_gdatetime[0], image.height-h-config.position_for_gdatetime[1]),
                  calendar.month(2023, 1), (255, 255, 255), font=font)
        if config.write_gregorian_next_month:
            gtoday = str(datetime.datetime.now().date())
            draw.text((image.width-w-config.position_for_gdatetime[0], image.height-h-config.position_for_gdatetime[1] + 300),
                      calendar.month(2023, 2), (255, 255, 255), font=font)
    draw = ImageDraw.Draw(image)
    image.save("output.jpg")


def get_wallpaper():
    try:
        if config.online_random_wallpaper:
            image_filename = wget.download(
                'https://picsum.photos/{}/{}'.format(config.system_resolution[0], config.system_resolution[1]), out='wallpapers')
            return image_filename
        elif config.offline_random_wallpaper:
            image_filename = random.choice(
                os.listdir(os.getcwd()+'\\wallpapers'))
            return 'wallpapers\\'+image_filename
        elif config.static_wallpaper:
            return 'wallpapers/' + config.default_wallpaper
    except:
        return 'wallpapers/default.jpg'


def start():
    create_wallpaper()
    set_wallpaper('\output.jpg')


if __name__ == '__main__':
    start()
