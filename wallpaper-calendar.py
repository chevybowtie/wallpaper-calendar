import argparse
import calendar
import ctypes
import datetime as dt
import platform
import re
import subprocess
from os import getcwd, listdir, path
from random import choice
import wget
from PIL import Image, ImageDraw, ImageFont
import config as settings

if platform.system() == 'Windows':
    import win32com.client
else:
    import helpers.kdesetwallpaper2 as helper

from num2words import num2words

# This constant represents the uiAction parameter for setting the desktop wallpaper using the SystemParametersInfoW() function.
SPI_SETDESKWALLPAPER = 20
# This constant represents the fWinIni flag to write the new setting to the user profile.
SPIF_UPDATEINIFILE = 0x1
# This constant represents the fWinIni flag to notify all top-level windows of the change.
SPIF_SENDWININICHANGE = 0x2

# these should move to config.py
# Today's date font color for highlighting
HIGHLIGHT_COLOR = settings.today_highlight_color
# Color for the calendar font
CALENDAR_COLOR = settings.calendar_base_color

# date helpers
TODAYDATE = dt.date.today()
CALSTARTDATE = dt.datetime.combine(TODAYDATE, dt.time.min)
CALENDDATE = CALSTARTDATE + dt.timedelta(days=settings.range_in_days)

CALENDARDATE = dt.date.today()

# Initialize the ArgumentParser
parser = argparse.ArgumentParser()

# Add the arguments that you want to override
parser.add_argument('--date', type=str,
                    help='provide date to generate calendar month')
parser.add_argument('--outlook', type=bool,
                    help='when set to true, populate using Outlook integration')
# Parse the command-line arguments
args = parser.parse_args()
# Override the values in config.py with the command-line arguments
if args.date:
    date_format = "%m/%d/%Y"
    CALENDARDATE = dt.datetime.strptime(args.date, date_format)
# if we have a custom calendar, calendardate will not match todaydate
CUSTOMCALENDAR = (CALENDARDATE != TODAYDATE)
if CUSTOMCALENDAR:
    TODAYDATE = CALENDARDATE

if args.outlook:
    settings.write_todays_appts = True


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
        helper.setwallpaper(wallpaper_path, plugin='org.kde.image')
        # execute_set(
        #     f"{cwd}/helpers/kdesetwallpaper {wallpaper_path}")
#        execute_set(
#            f"/usr/bin/gsettings set org.gnome.desktop.background picture-uri {wallpaper_path}")
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
    image = Image.open(wallpaper_name).convert('RGBA')
    # image_rgba = image.convert('RGBA')
    draw = ImageDraw.Draw(image)

    # Show calendar for this and next month
    if settings.write_gregorian_calendar_for_this_month:

        gtoday = str(TODAYDATE)
        w, h, *z = draw.textbbox((0, 0), gtoday, font)

        # set Sunday as 1st day of week
        calendar.setfirstweekday(calendar.SUNDAY)

        calendar_text = (calendar.month(TODAYDATE.year, TODAYDATE.month))
        calx = image.width-w-settings.position_for_calendar[0]
        caly = image.height-h-settings.position_for_calendar[1]
        _, _, text_width, text_height = draw.textbbox(
            (0, 0), calendar_text, font)
        corner_radius = settings.calendar_background_corner_radius
        padding = settings.calendar_background_padding

        # do we have calendar backgrounds enabled?
        if settings.calendar_background_enabled:
            bgx = calx - padding
            bgy = caly - padding
            mask = Image.new("RGBA", image.size, (0, 0, 0, 0))
            mask_draw = ImageDraw.Draw(mask)
            # Draw the calendar background with 25% opacity (64)
            mask_draw.rounded_rectangle((bgx, bgy, bgx + text_width + (
                padding * 2), bgy + text_height + (padding * 2)), corner_radius, fill=(128, 128, 128, 64))
            image = Image.alpha_composite(image, mask)
        if not CUSTOMCALENDAR:
            # draw it once in today's date color
            this_month_mask = Image.new("RGBA", image.size, (0, 0, 0, 0))
            mask_draw = ImageDraw.Draw(this_month_mask)
            mask_draw.text((calx, caly), calendar_text,
                           fill=HIGHLIGHT_COLOR, font=font)
            image = Image.alpha_composite(image, this_month_mask)

        # draw it again in white but skipping today
        # finds the first occurrence of today's date and remove it from this output
        if not CUSTOMCALENDAR:
            # how many characters are we replacing
            replacement = " " * len(str(TODAYDATE.day))
            # remove today's day
            calendar_text = re.sub(rf"\b{TODAYDATE.day}\b", replacement, calendar_text)
            

        this_month_mask = Image.new("RGBA", image.size, (0, 0, 0, 0))
        mask_draw = ImageDraw.Draw(this_month_mask)
        mask_draw.text((calx, caly), calendar_text,
                       fill=CALENDAR_COLOR, font=font)
        image = Image.alpha_composite(image, this_month_mask)
        corner_radius = settings.calendar_background_corner_radius
        padding = settings.calendar_background_padding

        # is next month's calendar enabled?
        if settings.write_gregorian_calendar_for_next_month:

            # space between calendars
            offset = text_height + settings.calendar_background_padding + \
                settings.space_between_calendars

            nextMonth = start_of_next_month(TODAYDATE)
            calendar_text = calendar.month(nextMonth.year, nextMonth.month)
            _, _, text_width, text_height = draw.textbbox(
                (0, 0), calendar_text, font)

            calx = image.width-w-settings.position_for_calendar[0]
            caly = image.height-h-settings.position_for_calendar[1]+offset

            # do we have calendar backgrounds enabled?
            if settings.calendar_background_enabled:

                bgx = image.width-w-settings.position_for_calendar[0] - padding
                bgy = image.height-h - \
                    settings.position_for_calendar[1]+offset - padding

                # Draw the calendar box with 50% opacity
                mask = Image.new("RGBA", image.size, (0, 0, 0, 0))
                mask_draw = ImageDraw.Draw(mask)

                # Draw the calendar background with 25% opacity (64)
                mask_draw.rounded_rectangle((bgx, bgy, bgx + text_width + (
                    padding * 2), bgy + text_height + (padding * 2)), corner_radius, fill=(128, 128, 128, 64))
                image = Image.alpha_composite(image, mask)

            next_month_mask = Image.new("RGBA", image.size, (0, 0, 0, 0))
            mask_draw = ImageDraw.Draw(next_month_mask)

            mask_draw.text((calx, caly), calendar.month(
                nextMonth.year, nextMonth.month), fill=CALENDAR_COLOR, font=font)
            image = Image.alpha_composite(image, next_month_mask)

    # if enabled, show today's appointments from Outlook on Windows
    if settings.write_todays_appts and platform.system() == 'Windows':
        appts = get_outlook_appointments(
            CALSTARTDATE, CALENDDATE)
        
        # row of appointment text
        outputRow = 0

        appointment_mask = Image.new("RGBA", image.size, (0, 0, 0, 0))
        mask_draw = ImageDraw.Draw(appointment_mask)
        
        # day of month
        row_day = appts[0].StartInStartTimeZone.day
        prev_day = appts[0].StartInStartTimeZone.day
        
        # row we are on for a given day
        day_row = 0

        for appointment in appts:

            row_day = appts[0].StartInStartTimeZone.day
            day_row += 1

            if row_day == prev_day:

                start_time = appointment.Start.strftime("%H:%M")

                # day of the week
                if day_row == 1:
                    mask_draw.text((settings.position_for_appts[0], settings.position_for_appts[1]+(
                        outputRow*60)), appointment.Start.strftime("%a"), CALENDAR_COLOR, font=apptFont)
                
                # ordinal date
                if day_row == 2:
                    mask_draw.text((settings.position_for_appts[0], settings.position_for_appts[1]+(
                        outputRow*60)), num2words(row_day, to='orndianl_num'), CALENDAR_COLOR, font=apptFont)
                
                # appointment start time
                mask_draw.text((settings.position_for_appts[0] + 170, settings.position_for_appts[1]+(
                    outputRow*60)), start_time, CALENDAR_COLOR, font=apptFont)
                
                # appointment subject
                mask_draw.text((settings.position_for_appts[0] + 370, settings.position_for_appts[1]+(
                    outputRow*60)), appointment.Subject, CALENDAR_COLOR, font=apptFont)
                
            else:
                mask_draw.text((settings.position_for_appts[0], settings.position_for_appts[1]+(
                    outputRow*60)), " ", font=apptFont)
                prev_day = row_day
                day_row = 0

            outputRow += 1
        image = Image.alpha_composite(image, appointment_mask)

    # if enabled, show today's data (large) if this is not a custom calendar
    if settings.write_today_big and not CUSTOMCALENDAR:
        today_big_mask = Image.new("RGBA", image.size, (0, 0, 0, 0))
        mask_draw = ImageDraw.Draw(today_big_mask)
        if settings.today_big_shadow:
            mask_draw.text((settings.position_for_today_big[0] + settings.today_big_shadow_offset, settings.position_for_today_big[1] + 100 + settings.today_big_shadow_offset), str(dt.datetime.today().day),
                           settings.today_big_shadow_color, font=heroFont)

        mask_draw.text((settings.position_for_today_big[0], settings.position_for_today_big[1]+100), str(dt.datetime.today().day),
                       CALENDAR_COLOR, font=heroFont)

        if settings.today_big_shadow:
            mask_draw.text((settings.position_for_today_big[0] + settings.today_big_shadow_offset, settings.position_for_today_big[1] + settings.today_big_shadow_offset), calendar.day_name[dt.datetime.today().weekday()],
                           settings.today_big_shadow_color, font=heroFont)

        mask_draw.text((settings.position_for_today_big[0], settings.position_for_today_big[1]), calendar.day_name[dt.datetime.today().weekday()],
                       CALENDAR_COLOR, font=heroFont)

        image = Image.alpha_composite(image, today_big_mask)

    # create the image
    draw = ImageDraw.Draw(image)

    # output to disk
    image = image.convert('RGB')
    image.save("output.jpg", 'JPEG', quality=90)


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
        print("Day of month: ", appointment.StartInStartTimeZone.day)
        print("Subject: ", appointment.Subject)
        # print("Is recurring: ", appointment.IsRecurring)
        # print("Is conflicted: ", appointment.IsConflict)
        # print("Is reminder set: ", appointment.ReminderSet)
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
