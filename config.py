# Wallpaper source

# pull from https://picsum.photos
online_random_wallpaper = True
# use one already downloaded (for offline use)
offline_random_wallpaper = False
# this fallback default is used if neither random option is enabled
default_wallpaper = 'default.jpg'
image_resolution = (1920, 1080)

# text
base_font_size = 24
default_font = 'unispace.bold.otf'

# calendar background
calendar_background_enabled = True
calendar_background_padding = 30
calendar_background_corner_radius = 8


# Hero title of day of week, and today's date
write_today_big = True

# Gregorian calendar embed
write_gregorian_calendar_for_this_month = True
write_gregorian_calendar_for_next_month = True
position_for_calendar = (350, 600)
position_for_today_big = (150, 25)
space_between_calendars = 60

# Windows Outlook appointments embed
write_todays_appts = False
position_for_appts = (50, 600)


def validate_config():
    """
    Validates the configuration file to ensure that it contains valid values.

    Args:
        None

    Returns:
        None
    """
    valid_resolutions = [(1366, 768), (1920, 1080), (2560, 1440), (3840, 2160)]
    valid_fonts = ["Kingthings Trypewriter 2.ttf", "simply-mono.book.ttf", "software-tester-7.regular.ttf",
                   "unispace.bold.otf", "code-new-roman.regular.otf"]  # add more here...

    if not isinstance(image_resolution, tuple) or image_resolution not in valid_resolutions:
        raise ValueError(
            "Invalid value for system_resolution in configuration file.")

    if not isinstance(default_font, str) or default_font not in valid_fonts:
        raise ValueError(
            "Invalid value for default_font in configuration file.")

    if not isinstance(base_font_size, int) or base_font_size <= 8:
        raise ValueError("Invalid value for font_size in configuration file.")

    if not isinstance(position_for_calendar, tuple) or len(position_for_calendar) != 2:
        raise ValueError(
            "Invalid value for position_for_calendar in configuration file.")

    if not isinstance(position_for_appts, tuple) or len(position_for_appts) != 2:
        raise ValueError(
            "Invalid value for position_for_appts in configuration file.")

    if not isinstance(position_for_today_big, tuple) or len(position_for_today_big) != 2:
        raise ValueError(
            "Invalid value for position_for_today in configuration file.")

    if not isinstance(write_gregorian_calendar_for_this_month, bool):
        raise ValueError(
            "Invalid value for write_gregorian_this_month in configuration file.")

    if not isinstance(write_gregorian_calendar_for_next_month, bool):
        raise ValueError(
            "Invalid value for write_gregorian_next_month in configuration file.")

    if not isinstance(write_todays_appts, bool):
        raise ValueError(
            "Invalid value for write_todays_appts in configuration file.")

    if not isinstance(write_today_big, bool):
        raise ValueError(
            "Invalid value for write_today_big in configuration file.")

    if not isinstance(online_random_wallpaper, bool):
        raise ValueError(
            "Invalid value for online_random_wallpaper in configuration file.")

    if not isinstance(offline_random_wallpaper, bool):
        raise ValueError(
            "Invalid value for offline_random_wallpaper in configuration file.")

    if not isinstance(default_wallpaper, str):
        raise ValueError(
            "Invalid value for default_wallpaper in configuration file.")
