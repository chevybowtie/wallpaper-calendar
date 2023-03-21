# Wallpaper Calendar
Wallpaper Calendar is a simple command-line utility that allows you to generate a calendar as your desktop wallpaper. It supports different styles and layouts, and can be customized to fit your preferences. By default, it use a new random image. Run it multiple times to generate different wallpapers.

## Installation
To install Wallpaper Calendar, you need to have Python 3.x installed on your system. Then, you can clone this repository and install the dependencies:

```
$ git clone https://github.com/chevybowtie/wallpaper-calendar.git
$ cd wallpaper-calendar
$ pip install -r requirements.txt
```
Once you have installed the dependencies, you can run the wallpaper-calendar.py script to generate your wallpaper calendar. By default, it will generate a calendar for the current month and year, using the "classic" style.

```
$ python wallpaper-calendar.py
```
## Customization 
Many options exist in `config.py`. You can change fonts, enable multiple calendars, and reposition most items based on values in the `config.py` file.


You may also perform one-off customizations by specifying a few command line parameters. For example, to generate a calendar for a date in the future, specify the date
```
$ python wallpaper-calendar.py --date=9/15/2023
```
This will generate a new wallpaper image with only the calendar overlaid on it.


## Outlook Integration
This feature branch adds an experimental feature that allows you to automatically generate and set your wallpaper calendar by including your Outlook appointments for the day (or multiple days).

To use this feature, you need to have Outlook installed on your system and you need to have pywin32 installed (which is not included in requirements.txt since it doesn't work on Linux):
```
$ # addtional Outlook requirement
$ pip install pywin32
```
Then, using the config value `write_todays_appts = True`, you can run the wallpaper-calendar.py script to generate and set your wallpaper calendar with Outlook's appointments

This will generate a calendar, while also embedding today's schedule based on Outlook's calendar. Note that this feature is still experimental and may not work on all systems or with all versions of Outlook. Currently it works on my Office 365 installation.