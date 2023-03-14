import ctypes
import datetime as dt
from os import getcwd, listdir, path
import platform
from random import choice
import calendar
import subprocess
import wget
from PIL import Image, ImageDraw, ImageFont
import config as settings
if platform.system() == 'Windows':
    import win32com.client


# This constant represents the uiAction parameter for setting the desktop wallpaper using the SystemParametersInfoW() function.
SPI_SETDESKWALLPAPER = 20
# This constant represents the fWinIni flag to write the new setting to the user profile.
SPIF_UPDATEINIFILE = 0x1
# This constant represents the fWinIni flag to notify all top-level windows of the change.
SPIF_SENDWININICHANGE = 0x2

# these should move to config.py
HIGHLIGHT_COLOR = (0, 255, 0)       # Today's date font color for highlighting
CALENDAR_COLOR = (255, 255, 255)    # Color for the calendar font

# date helpers
TODAYDATE = dt.date.today()
STARTOFTODAY = dt.datetime.combine(TODAYDATE, dt.time.min)
STARTOFTOMORROW = STARTOFTODAY + dt.timedelta(days=1)


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
    fontFile = 'fonts/{}'.format(settings.default_font)
    font = ImageFont.truetype(fontFile, settings.base_font_size)
    heroFont = ImageFont.truetype(fontFile, settings.base_font_size * 4)
    apptFont = ImageFont.truetype(fontFile, settings.base_font_size * 2)

    # Load Image
    wallpaper_name = get_wallpaper()
    print(wallpaper_name)
    image = Image.open(wallpaper_name)
    draw = ImageDraw.Draw(image)

    # Show calendar for this and next month
    if settings.write_gregorian_calendar_for_this_month:
        gtoday = str(TODAYDATE)
        w, h, *z = draw.textbbox((0, 0), gtoday, font)
        calendar_output = (calendar.month(TODAYDATE.year, TODAYDATE.month))

        calx = image.width-w-settings.position_for_calendar[0]
        caly = image.height-h-settings.position_for_calendar[1]

        calendar_text = calendar.month(TODAYDATE.year, TODAYDATE.month)
        _, _, text_width, text_height = draw.textbbox( (0,0), calendar_text, font)

        # do we have calendar backgrounds enabled?
        if settings.calendar_background_enabled:

            corner_radius = settings.calendar_background_corner_radius
            padding = settings.calendar_background_padding

            # Draw the rounded rectangle with 50% opacity
            fill_color = (128, 128, 128, 128)  # gray with 50% opacity
            outline_color = (255, 255, 255)  # no outline
            bgx = calx - padding
            bgy = caly - padding

            draw.rounded_rectangle((bgx, bgy, calx + text_width + padding,
                                    caly + text_height + padding), corner_radius, fill_color, outline_color)

        # draw it once in today's date color
        draw.text((calx, caly),
                  calendar_output, fill=HIGHLIGHT_COLOR, font=font)

        # draw it again in white but skipping today

        # finds the first occurrence of today's date and remove it from this output
        todays_date = " " if TODAYDATE.day < 10 else "  "
        highlighted_day = calendar_output.replace(
            str(TODAYDATE.day), todays_date, 1)

        draw.text((calx, caly),
                  highlighted_day, fill=CALENDAR_COLOR, font=font)
        

        # is next month's calendar enabled?
        if settings.write_gregorian_calendar_for_next_month:

            # space between calendars
            offset = text_height + settings.calendar_background_padding + settings.space_between_calendars

            nextMonth = start_of_next_month(TODAYDATE)
            calendar_text = calendar.month(nextMonth.year, nextMonth.month)
            _, _, text_width, text_height = draw.textbbox( (0,0), calendar_text, font)


            calx = image.width-w-settings.position_for_calendar[0]
            caly = image.height-h-settings.position_for_calendar[1]+offset

            # do we have calendar backgrounds enabled?
            if settings.calendar_background_enabled:
                corner_radius = settings.calendar_background_corner_radius
                padding = settings.calendar_background_padding

                bgx = image.width-w-settings.position_for_calendar[0] - padding
                bgy = image.height-h - \
                    settings.position_for_calendar[1]+offset - padding

                # Draw the rounded rectangle with 50% opacity
                fill_color = (128, 128, 128, 128)  # gray with 50% opacity
                outline_color = (255, 255, 255)  # no outline

                draw.rounded_rectangle((bgx, bgy, calx + text_width + padding,
                    caly + text_height + padding), corner_radius, fill_color, outline_color)

            draw.text((calx, caly),
                      calendar.month(nextMonth.year, nextMonth.month), fill=CALENDAR_COLOR, font=font)



    # if enabled, show today's appointments from Outlook on Windows
    if settings.write_todays_appts and platform.system() == 'Windows':
        appts = get_outlook_appointments(
            STARTOFTODAY, STARTOFTOMORROW)
        outputRow = 0
        for appointment in appts:
            start_time = appointment.Start.strftime("%H:%M")
            draw.text((settings.position_for_appts[0], settings.position_for_appts[1]+(
                outputRow*60)), start_time, CALENDAR_COLOR, font=apptFont)
            draw.text((settings.position_for_appts[0]+200, settings.position_for_appts[1]+(
                outputRow*60)), appointment.Subject, CALENDAR_COLOR, font=apptFont)
            outputRow += 1



    # if enabled, show today's data (large)
    if settings.write_today_big:
        draw.text((settings.position_for_today_big[0], settings.position_for_today_big[1]+100), str(dt.datetime.today().day),
                  CALENDAR_COLOR, font=heroFont)
        draw.text((settings.position_for_today_big[0], settings.position_for_today_big[1]), calendar.day_name[dt.datetime.today().weekday()],
                  CALENDAR_COLOR, font=heroFont)

    # create the image
    draw = ImageDraw.Draw(image)

    # output to disk
    image.save("output.jpg")


def start_of_next_month(date):
    """
    Returns a `datetime.date` object representing the first day of the next month after the specified date.

    Args:
        current_date (datetime.date): The date to use as a reference for calculating the next month.

    Returns:
        datetime.date: A `datetime.date` object representing the first day of the next month.
    """
    year = date.year + (date.month // 12)
    month = date.month % 12 + 1
    next_month = dt.date(year, month, 1)

    # Return the start of the next month
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
        if settings.online_random_wallpaper:
            image_filename = wget.download(
                'https://picsum.photos/{}/{}'.format(settings.image_resolution[0], settings.image_resolution[1]), out='wallpapers')
            return image_filename
        elif settings.offline_random_wallpaper:
            image_filename = choice(
                listdir(getcwd()+'\\wallpapers'))
            return 'wallpapers\\'+image_filename
        else:
            return 'wallpapers/' + settings.default_wallpaper
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

    # show what we found in Outlook
    for appointment in calendar:
        print("Subject: ", appointment.Subject)
        print("Start: ", appointment.Start)
        print("---------")

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


def start():
    try:
        settings.validate_config()
    except ValueError as e:
        print(f"Invalid configuration file: {e}")
        return

    create_wallpaper()
    set_wallpaper('output.jpg')


if __name__ == '__main__':
    start()
