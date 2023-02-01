import ctypes
import datetime as dt
import os
import platform
import random
import calendar

# Windows APIs
import win32com.client

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
        execute_set(
            "/usr/bin/gsettings set org.gnome.desktop.background picture-uri " + cwd + wallpaper_name)
    else:
        input('Your operating system is not supported')


def create_wallpaper():

    # Load font
    fontFile = 'fonts/{}'.format(config.default_font)
    font = ImageFont.truetype(fontFile, config.font_size)
    heroFont = ImageFont.truetype(fontFile, config.font_size * 4)
    today = dt.date.today()

    # Load Image
    wallpaper_name = get_wallpaper()
    image = Image.open(wallpaper_name)
    draw = ImageDraw.Draw(image)

    # Show calendar for this and next month
    if config.write_gregorian_this_month:
        gtoday = str(today)
        w, h, *z = draw.textbbox((0, 0), gtoday, font)
        draw.text((image.width-w-config.position_for_calendar[0], image.height-h-config.position_for_calendar[1]),
                  calendar.month(2023, 1), (255, 255, 255), font=font)

        if config.write_gregorian_next_month:
            nextMonth = next_month_date(today)
            draw.text((image.width-w-config.position_for_calendar[0], image.height-h-config.position_for_calendar[1] + 300),
                      calendar.month(nextMonth.year, nextMonth.month), (255, 255, 255), font=font)

    # if enabled, show today's appointments from Outlook
    if config.write_todays_appts:
        appts = get_outlook_appointments(
            dt.datetime(2023, 1, 31), dt.datetime(2023, 2, 1))
        draw.text((config.position_for_appts[0], config.position_for_appts[1]),
                  appts, (255, 255, 255), font=font)

    # if enabled, show today's data (large)
    if config.write_today_big:
        draw.text((config.position_for_today[0], config.position_for_today[1]+100), str(dt.datetime.today().day),
                  (255, 255, 255), font=heroFont)
        draw.text((config.position_for_today[0], config.position_for_today[1]), calendar.day_name[dt.datetime.today().weekday()],
                  (255, 255, 255), font=heroFont)

    # create the image
    draw = ImageDraw.Draw(image)

    # output to disk
    image.save("output.jpg")


def next_month_date(d):
    year = d.year+(d.month//12)
    month = 1 if (d.month//12) else d.month + 1
    next_month_len = calendar.monthrange(year, month)[1]
    next_month = d
    if d.day > next_month_len:
        next_month = next_month.replace(day=next_month_len)
    next_month = next_month.replace(year=year, month=month)
    return next_month


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


def get_outlook_appointments(begin, end):
    # get a handle on Outlook
    outlook = win32com.client.Dispatch(
        'Outlook.Application').GetNamespace('MAPI')

    # 9 = list of all the meetings, 6 = emails
    calendar = outlook.getDefaultFolder(9).Items

    # events that repeat
    calendar.IncludeRecurrences = True

    # sort
    calendar.Sort("[Start]")

    # setup a constraint
    calendar.Sort('[Start]')
    restriction = "[Start] >= '" + begin.strftime(
        '%m/%d/%Y') + "' AND [END] <= '" + end.strftime('%m/%d/%Y') + "'"
    calendar = calendar.Restrict(restriction)

    return calendar


def groom_appointments(calendar):
    appointmentDictionary = {}

    for appointment in calendar:
        meetingDate = str(appointment.Start)
        subject = str(appointment.Subject)
        duration = str(appointment.duration)
        # date = parse(meetingDate).date()
        # time = parse(meetingDate).time()
        appointmentDictionary[subject] = {"Subject": [
            subject], "Time": [meetingDate], "Durations": [duration]}

    for subject in appointmentDictionary.keys():
        rowDict = {}
        rowDict["Subject"] = appointmentDictionary[subject]["Subject"] if appointmentDictionary[subject]["Subject"] else ""
        ApptRow(rowDict)


def ApptRow(rowDictionary):
    print(rowDictionary)


def start():
    create_wallpaper()
    set_wallpaper('\output.jpg')


if __name__ == '__main__':
    start()
