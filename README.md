# Wallpaper Calendar
Wallpaper Calendar is a simple command-line utility that allows you to generate a calendar as your desktop wallpaper. It supports different styles and layouts, and can be customized to fit your preferences.

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

You can customize the output by specifying different options. For example, to generate a calendar for a date in the future, specify the date
```
$ python wallpaper-calendar.py --date=9/15/2023
```
This will generate a new wallpaper image with only the calendar overlaid on it.


## Outlook Integration
This feature branch adds an experimental feature that allows you to automatically generate and set your wallpaper calendar by including your Outlook appointments for the day. To use this feature, you need to have Outlook installed on your system, and you need to have pywin32 installed:
```
$ pip install pywin32
```
Then, you can run the wallpaper-calendar-outlook.py script to generate and set your wallpaper calendar with today's appointments:

```
$ python wallpaper-calendar.py --outlook=true
```
This will generate a calendar, while also embedding today's schedule based on Outlook's calendar. Note that this feature is still experimental and may not work on all systems or with all versions of Outlook.