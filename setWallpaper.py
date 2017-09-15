"""

Author: Naveen Bhurli

Script that fetches Bing Photo of the day, and set it as desktop wallpaper.
This script runs in the background as soon as the OS boots up and starts working once it is connected to the internet.
It uses simple urllib.request, json and winreg modules that are already part of Python base installation.
The script fetches Bing wallpaper of the day, saves it onto the local system and sets the wallpaper.

"""


# Execute this command after setting wallpaper "RUNDLL32.EXE USER32.DLL,UpdatePerUserSystemParameters 1, True"
# urlretrieve

from urllib.request import *
import json
from winreg import *
import os,datetime,re,time
import subprocess


# method to connect to the Internet and fetches photo of the day.

def get_image_from_bing():
    present_day = datetime.datetime.today()
    bingWallpaperUrl = "http://www.bing.com/HPImageArchive.aspx?format=js&idx=0&n=1&mkt=en-IN"

    with urlopen(bingWallpaperUrl) as content:
        content = urlopen(bingWallpaperUrl)

        if content.getcode() == 200:
            html = json.load(content)
            image = "https://www.bing.com" + html['images'][0]['url']
            image_path = html['images'][0]['url']
            image_name = re.search(r'/rb/([^/]+)', image_path).group(1)

            bing_wallpaper_folder = create_folder_for_images()

            current_day_image = bing_wallpaper_folder+"\\"+image_name




            urlretrieve(image, current_day_image)
            return current_day_image



        else:
            print("Connectivity problem")

def create_folder_for_images():
    user_profile = os.environ['userprofile']
    bing_wallpaper_folder = user_profile + "\\Documents\\BingPhotoOfTheDay"
    if not os.path.exists(bing_wallpaper_folder):
        os.makedirs(bing_wallpaper_folder)
    return bing_wallpaper_folder

# wallpaper path = HKEY_CURRENT_USER\Control Panel\Desktop\WallPaper

def set_wallpaper(image):
    Registry = ConnectRegistry(None, HKEY_CURRENT_USER)
    raw_key = OpenKey(Registry, "Control Panel\Desktop", 0, KEY_ALL_ACCESS)
    # query_reg_value = QueryValueEx(raw_key, "WallPaper")
    # print(query_reg_value)
    SetValueEx(raw_key, "WallPaper", 1, REG_SZ, image)
    CloseKey(Registry)



# def check_for_existing_image():
#
#     if os.path.exists()


# set_wallpaper(get_image_from_bing())

# create_folder_for_images()
bing = "https://www.bing.com"
with urlopen(bing) as conn:
    if conn.getcode() == 200:

        while True:
            current_day_image = get_image_from_bing()
            if not os.path.isfile(current_day_image):
                set_wallpaper(current_day_image)
                for refresh in range(1, 15):
                    subprocess.call(["RUNDLL32.EXE", "USER32.DLL", "UpdatePerUserSystemParameters", "1", "True"])
            time.sleep(3600)