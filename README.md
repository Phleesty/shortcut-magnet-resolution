# Shortcuts Magnet Resolution (for Windows)

I created a small GUI tool to help with applications and games that do not allow to set the desired resolution or refresh rate in their settings. I had to use the small command-line utility [NirCmd](https://www.nirsoft.net/utils/nircmd.html) to implement the main functions. Its license allows you to distribute and use it in your projects for free.

![smr_screenshot](https://i.imgur.com/z2EEJhS.png)

## How to install and use

You can simply download one of the two files `shortcuts-magnet-res-EN.exe` with the English interface or `shortcuts-magnet-res-RU.exe` with the Russian interface from [the release page](https://github.com/phleesty/shortcuts-magnet-resolution/releases). Everything will work without any additional configuration.

Or you can use python. To do this, you need to install dependencies and put nircmd.exe next to shortcuts-magnet-resolution.py

```
pip install PyQt5 
pip install pywin32
```

## How does it work

You just choose the path to the application, the desired resolution and refresh rate you want for the application. And just below that you choose which settings you want to go back to when you close the application (by default both lines are your current desktop settings). After clicking on the `Save and create shortcut` button a shortcut will appear on the desktop, which, when launched, will set all the parameters that you have set for the application, as well as launch the application itself. After closing the application, all the settings will revert to the ones you specified under `Your native resolution`.

> *Also in the settings there is such a parameter as Bits Per Pixel. If you don't understand what this is, you probably don't need this setting. Just leave the default setting. For most new monitors this value will be 32 bits.*

## When would this be useful?

For two years I struggled with a completely unstable framerate when I ran Genshin Impact at `75Hz` (my monitor's refresh rate). I had to go to windows settings and change the refresh rate every time I played the game to `60Hz` (it was perfect when I did that). But the problem was that it took a long time and every time I forgot to change the settings back. So I started looking for a solution and NirCmd turned out to be the only tool that worked perfectly.

## How the code works

It's actually very simple and it comes down to creating a bat file, which in the end will apply these three lines:

``` batch
nircmd.exe setdisplay %inGameWidth% %inGameHeight% %inGamebpp% %inGamefreq%
start /wait "" "%game%"
nircmd.exe setdisplay %myWidth% %myHeight% %myBpp% %myFreq%
```

If you don't need a GUI, you can just download NirCmd from the [official website](https://www.nirsoft.net/utils/nircmd.html) and make .bat files like this yourself:

``` batch
nircmd.exe setdisplay 1920 1080 32 60
start /wait "" " D:\Genshin Impact\GenshinImpact.exe"
nircmd.exe setdisplay 1920 1080 32 75
```

The GUI I made just automatically adds a shortcut to the desktop with the icon and name from the original application. It also offers default values so that the user is not confused by commands and designations. This is my first experiment with `python` programming as well as creating an interface in `PyQt5`. Of course I did it all for me first of all but if this little tool helps someone I will be immensely happy  <img width="25" src="https://emojis.slackmojis.com/emojis/images/1593555389/9579/blob_excited.gif?1593555389" alt="party blob" />
