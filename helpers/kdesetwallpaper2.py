import pathlib
import string
import dbus
import random
import tempfile, shutil, os

from pathlib import Path



# HOME = str(Path.home())

# SCREEN_LOCK_CONFIG = HOME+"/.config/kscreenlockerrc"


def create_temporary_copy(path):
    # by giving the file a random name, i hope KDE will pickup the change
    file_extension = pathlib.Path(path).suffix    
    tmp = tempfile.NamedTemporaryFile(delete=True)
    shutil.copy2(path, tmp.name + file_extension)
    return tmp.name + file_extension


def setwallpaper(filepath, plugin='org.kde.image'):

    new_temp_file = create_temporary_copy(filepath)    
    print(new_temp_file)
    jscript = """
    var allDesktops = desktops();
    print (allDesktops);
    for (i=0;i<allDesktops.length;i++) {
        d = allDesktops[i];
        d.wallpaperPlugin = "%s";
        d.currentConfigGroup = Array("Wallpaper", "%s", "General");
        d.writeConfig("Image", "file://{%s}")
    }
    """
    bus = dbus.SessionBus()
    plasma = dbus.Interface(bus.get_object(
        'org.kde.plasmashell', '/PlasmaShell'), dbus_interface='org.kde.PlasmaShell')
    plasma.evaluateScript(jscript % (plugin, plugin, new_temp_file))
