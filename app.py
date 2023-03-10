import ctypes
import datetime as dt
from os import getcwd, listdir, path
import platform
from random import choice
import calendar
import win32com.client
import subprocess
import wget
from PIL import Image, ImageDraw, ImageFont
import config


# This constant represents the uiAction parameter for setting the desktop wallpaper using the SystemParametersInfoW() function.
SPI_SETDESKWALLPAPER = 20
# This constant represents the fWinIni flag to write the new setting to the user profile.
SPIF_UPDATEINIFILE = 0x1
# This constant represents the fWinIni flag to notify all top-level windows of the change.
SPIF_SENDWININICHANGE = 0x2

# these should move to config.py
HIGHLIGHT_COLOR = (0, 255, 0)       # Today's date font color for highlighting
CALENDAR_COLOR = (255, 255, 255)    # Color for the calendar font


def execute_set(command):
    """
    Executes a shell command using subprocess.run.

    Args:
        command (str): The shell command to execute.

    Returns:
        None
    """
    try:
        subprocess.run(command.split(), check=True, shell=False)
    except subprocess.CalledProcessError as e:
        print(f"Error executing command: {e}")


def set_wallpaper(wallpaper_name='\output.jpg'):
    """
    Sets the desktop wallpaper to the specified image file.

    Args:
        wallpaper_name (str): The name of the image file to use as the wallpaper.

    Returns:
        None
    """

    cwd = getcwd()
    wallpaper_path = path.join(cwd, wallpaper_name)

    if platform.system() == 'Windows':
        print("setting wallpaper")
        success = ctypes.windll.user32.SystemParametersInfoW(
            SPI_SETDESKWALLPAPER, 0, wallpaper_path, SPIF_UPDATEINIFILE | SPIF_SENDWININICHANGE)
        if not success:
            print('Failed to set wallpaper')
        # else:
        #     print('Wallpaper set successfully')
    elif platform.system() == 'Linux':
        execute_set(
            f"/usr/bin/gsettings set org.gnome.desktop.background picture-uri {wallpaper_path}")
    else:
        input('Your operating system is not supported')


def create_wallpaper():
    """
    Generates a wallpaper image with the current date, calendar, and other information.

    Args:
        None

    Returns:
        None
    """
    # Load font
    fontFile = 'fonts/{}'.format(config.default_font)
    font = ImageFont.truetype(fontFile, config.font_size)
    heroFont = ImageFont.truetype(fontFile, config.font_size * 4)

    today = dt.date.today()

    # Load Image
    wallpaper_name = get_wallpaper()
    print(wallpaper_name)
    image = Image.open(wallpaper_name)
    draw = ImageDraw.Draw(image)

    # Show calendar for this and next month
    if config.write_gregorian_this_month:
        gtoday = str(today)
        w, h, *z = draw.textbbox((0, 0), gtoday, font)
        calendar_output = (calendar.month(today.year, today.month))

        # draw it once in today's date color
        draw.text((image.width-w-config.position_for_calendar[0], image.height-h-config.position_for_calendar[1]),
                  calendar_output, fill=HIGHLIGHT_COLOR, font=font)

        # draw it again in white but skipping today
        todays_date = " " if today.day < 10 else "  "
        # finds the first occurrence of today's date and remove it from this output
        highlighted_day = calendar_output.replace(
            str(today.day), todays_date, 1)

        draw.text((image.width-w-config.position_for_calendar[0], image.height-h-config.position_for_calendar[1]),
                  highlighted_day, fill=CALENDAR_COLOR, font=font)

        if config.write_gregorian_next_month:
            nextMonth = next_month_date(today)
            draw.text((image.width-w-config.position_for_calendar[0], image.height-h-config.position_for_calendar[1] + 300),
                      calendar.month(nextMonth.year, nextMonth.month), fill=CALENDAR_COLOR, font=font)

    # if enabled, show today's appointments from Outlook
    if config.write_todays_appts:
        appts = get_outlook_appointments(
            dt.datetime(2023, 1, 31), dt.datetime(2023, 2, 1))
        draw.text((config.position_for_appts[0], config.position_for_appts[1]),
                  appts, CALENDAR_COLOR, font=font)

    # if enabled, show today's data (large)
    if config.write_today_big:
        draw.text((config.position_for_today[0], config.position_for_today[1]+100), str(dt.datetime.today().day),
                  CALENDAR_COLOR, font=heroFont)
        draw.text((config.position_for_today[0], config.position_for_today[1]), calendar.day_name[dt.datetime.today().weekday()],
                  CALENDAR_COLOR, font=heroFont)

    # create the image
    draw = ImageDraw.Draw(image)

    # output to disk
    image.save("output.jpg")


def next_month_date(current_date):
    """
    Returns a `datetime.date` object representing the first day of the next month after the specified date.

    Args:
        current_date (datetime.date): The date to use as a reference for calculating the next month.

    Returns:
        datetime.date: A `datetime.date` object representing the first day of the next month.
    """
    year = current_date.year+(current_date.month//12)
    month = 1 if (current_date.month//12) else current_date.month + 1
    next_month_len = calendar.monthrange(year, month)[1]
    next_month = current_date
    if current_date.day > next_month_len:
        next_month = next_month.replace(day=next_month_len)
    next_month = next_month.replace(year=year, month=month)
    return next_month


def get_wallpaper():
    """
    Returns the filename of the wallpaper image to use, based on the specified configuration options.

    Args:
        None

    Returns:
        str: The filename of the wallpaper image to use.
    """
    try:
        if config.online_random_wallpaper:
            image_filename = wget.download(
                'https://picsum.photos/{}/{}'.format(config.system_resolution[0], config.system_resolution[1]), out='wallpapers')
            return image_filename
        elif config.offline_random_wallpaper:
            image_filename = choice(
                listdir(getcwd()+'\\wallpapers'))
            return 'wallpapers\\'+image_filename
        elif config.static_wallpaper:
            return 'wallpapers/' + config.default_wallpaper
    except:
        return 'wallpapers/default.jpg'


def get_outlook_appointments(begin, end):
    """
    Queries Outlook for appointment items between the specified begin and end dates.

    Args:
        begin (datetime.datetime): The start time of the interval to query, as a `datetime.datetime` object.
        end (datetime.datetime): The end time of the interval to query, as a `datetime.datetime` object.

    Returns:
        str: A string containing a formatted list of appointment items.
    """

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
    """
    Takes a list of Outlook appointment items and processes them into a dictionary format suitable for display.

    Args:
        calendar (list): A list of Outlook appointment items.

    Returns:
        None
    """
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

# each appointment row should be output to the wallpaper


def ApptRow(rowDictionary):
    print(rowDictionary)


def validate_config():
    """
    Validates the configuration file to ensure that it contains valid values.

    Args:
        None

    Returns:
        None
    """
    valid_resolutions = [(1366, 768), (1920, 1080), (2560, 1440), (3840, 2160)]
    valid_fonts = ["Kingthings Trypewriter 2.ttf"]  # add more here...

    if not isinstance(config.system_resolution, tuple) or config.system_resolution not in valid_resolutions:
        raise ValueError(
            "Invalid value for system_resolution in configuration file.")

    if not isinstance(config.default_font, str) or config.default_font not in valid_fonts:
        raise ValueError(
            "Invalid value for default_font in configuration file.")

    if not isinstance(config.font_size, int) or config.font_size <= 8:
        raise ValueError("Invalid value for font_size in configuration file.")

    if not isinstance(config.position_for_calendar, tuple) or len(config.position_for_calendar) != 2:
        raise ValueError(
            "Invalid value for position_for_calendar in configuration file.")

    if not isinstance(config.position_for_appts, tuple) or len(config.position_for_appts) != 2:
        raise ValueError(
            "Invalid value for position_for_appts in configuration file.")

    if not isinstance(config.position_for_today, tuple) or len(config.position_for_today) != 2:
        raise ValueError(
            "Invalid value for position_for_today in configuration file.")

    if not isinstance(config.write_gregorian_this_month, bool):
        raise ValueError(
            "Invalid value for write_gregorian_this_month in configuration file.")

    if not isinstance(config.write_gregorian_next_month, bool):
        raise ValueError(
            "Invalid value for write_gregorian_next_month in configuration file.")

    if not isinstance(config.write_todays_appts, bool):
        raise ValueError(
            "Invalid value for write_todays_appts in configuration file.")

    if not isinstance(config.write_today_big, bool):
        raise ValueError(
            "Invalid value for write_today_big in configuration file.")

    if not isinstance(config.online_random_wallpaper, bool):
        raise ValueError(
            "Invalid value for online_random_wallpaper in configuration file.")

    if not isinstance(config.offline_random_wallpaper, bool):
        raise ValueError(
            "Invalid value for offline_random_wallpaper in configuration file.")

    if not isinstance(config.static_wallpaper, bool):
        raise ValueError(
            "Invalid value for static_wallpaper in configuration file.")

    if not isinstance(config.default_wallpaper, str):
        raise ValueError(
            "Invalid value for default_wallpaper in configuration file.")


def start():
    try:
        validate_config()
    except ValueError as e:
        print(f"Invalid configuration file: {e}")
        return

    create_wallpaper()
    set_wallpaper('output.jpg')


if __name__ == '__main__':
    start()
