# Wallpaper source
online_random_wallpaper = False
online_wallpaper_source = 'unsplash' # should be a one of the keys in online_wallpaper_resources
# not yet used
online_wallpaper_resources = {
    "picsum" : "https://picsum.photos/{}/{}",
    "unsplash" : "https://source.unsplash.com/random/{}x{}"
}

# use one already downloaded (for offline use)
offline_random_wallpaper = True
# this fallback default is used if neither random option is enabled
default_wallpaper = 'default-3440x1440.jpg'
image_resolution = (1920, 1080)
# image_resolution = (3440, 1440)

# if your outlook appointments have categories, you may alter the colors
appointment_colors = {
    "Boss": (255, 0, 0),            # Red
    # "lunch": (160, 160, 160),       # dimmed (gray)
    "important": (255,0,0),         # red
    # "reminder": (0, 255, 0),        # Green
    "reminder":(128,64,0),          # brown
    # Add more categories and colors as needed

}

# List of appointment subjects to be skipped
skipped_appointment_subjects = ["lunch"]

# Wallpaper style: 'center', 'tile', 'stretch', 'fit', 'fill', 'span'
wallpaper_style = 'fit'

# text (appts = 2x, today_big = 4x)
base_font_size = 24
default_font = 'unispace.bold.otf'

# calendar background
calendar_background_enabled = True
calendar_background_padding = 30
calendar_background_corner_radius = 8

# Hero title of day of week, and today's date
write_today_big = True
today_big_shadow = True

# text_shadow_color = (204, 204, 204) # gray
text_shadow_color = (45, 45, 45) # dark gray
text_shadow_offset = 2

# Gregorian calendar embed
write_gregorian_calendar_for_this_month = True
write_gregorian_calendar_for_next_month = True
position_for_calendar = (350, 600)
position_for_today_big = (200, 25)
space_between_calendars = 60

# Today's date font color for highlighting
today_highlight_color = (187,51,255)
calendar_base_color = (255, 255, 255)    # Color for the calendar font

# appointments embed
# calendar_access = 'Outlook client'
write_todays_appts = True
position_for_appts = (200, 300)
range_in_days = 4


def validate_config():
    """
    Validates the configuration file to ensure that it contains valid values.

    Args:
        None

    Returns:
        None
    """
    valid_resolutions = [(1366, 768), (1920, 1080), (2560, 1440), (3840, 2160), (3440, 1440)]
    valid_fonts = ["Kingthings Trypewriter 2.ttf", "simply-mono.book.ttf", "software-tester-7.regular.ttf",
                   "unispace.bold.otf", "code-new-roman.regular.otf"]  # add more here...

    if not isinstance(today_highlight_color, tuple) or len(today_highlight_color) != 3:
        raise ValueError(
            "Invalid value for today_highlight_color in configuration file.")

    for value in today_highlight_color:
        if not isinstance(value, int) or value < 0 or value > 255:
            raise ValueError(
                "Invalid value for today_highlight_color in configuration file.")

    if not isinstance(calendar_base_color, tuple) or len(calendar_base_color) != 3:
        raise ValueError(
            "Invalid value for calendar_base_color in configuration file.")

    for value in calendar_base_color:
        if not isinstance(value, int) or value < 0 or value > 255:
            raise ValueError(
                "Invalid value for calendar_base_color in configuration file.")

    if not isinstance(calendar_background_enabled, bool):
        raise ValueError(
            "Invalid value for calendar_background_enabled in configuration file.")

    if not isinstance(calendar_background_padding, int) or calendar_background_padding < 0:
        raise ValueError(
            "Invalid value for calendar_background_padding in configuration file.")

    if not isinstance(calendar_background_corner_radius, int) or calendar_background_corner_radius < 0:
        raise ValueError(
            "Invalid value for calendar_background_corner_radius in configuration file.")

    if not isinstance(write_today_big, bool):
        raise ValueError(
            "Invalid value for write_today_big in configuration file.")

    if not isinstance(image_resolution, tuple) or image_resolution not in valid_resolutions:
        raise ValueError(
            "Invalid value for system_resolution in configuration file.")

    if not isinstance(default_font, str) or default_font not in valid_fonts:
        raise ValueError(
            "Invalid value for default_font in configuration file.")

    if not isinstance(text_shadow_offset, int):
        raise ValueError(
            "Invalid value for text_shadow_offset in configuration file.")

    if not isinstance(space_between_calendars, int):
        raise ValueError(
            "Invalid value for space_between_calendars in configuration file.")

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

    if not isinstance(today_big_shadow, bool):
        raise ValueError(
            "Invalid value for today_big_shadow in configuration file.")

    if not isinstance(text_shadow_color, tuple) or len(text_shadow_color) != 3:
        raise ValueError(
            "Invalid value for text_shadow_color in configuration file.")

    for value in text_shadow_color:
        if not isinstance(value, int) or value < 0 or value > 255:
            raise ValueError(
                "Invalid value for text_shadow_color in configuration file.")

    if not isinstance(online_random_wallpaper, bool):
        raise ValueError(
            "Invalid value for online_random_wallpaper in configuration file.")

    if not isinstance(offline_random_wallpaper, bool):
        raise ValueError(
            "Invalid value for offline_random_wallpaper in configuration file.")

    if not isinstance(default_wallpaper, str):
        raise ValueError(
            "Invalid value for default_wallpaper in configuration file.")
    
    if not isinstance(online_wallpaper_source, str) or online_wallpaper_source not in online_wallpaper_resources:
        raise ValueError("Invalid value for online_wallpaper_source in configuration file.")
        
    if not isinstance(online_wallpaper_resources, dict):
        raise ValueError("Invalid value for online_wallpaper_resources in configuration file.")
        
    if not isinstance(appointment_colors, dict):
        raise ValueError("Invalid value for appointment_colors in configuration file.")
        
    for category, color in appointment_colors.items():
        if not isinstance(category, str):
            raise ValueError("Invalid category value in appointment_colors in configuration file.")
        if not isinstance(color, tuple) or len(color) != 3:
            raise ValueError("Invalid color value in appointment_colors in configuration file.")
        for value in color:
            if not isinstance(value, int) or value < 0 or value > 255:
                raise ValueError("Invalid color component in appointment_colors in configuration file.")