# PSVideoDatabase
 PowerShell based Video Database with TMDB connection

## License
This project is deve [GNU General Public License v3.0](https://www.gnu.org/licenses/gpl-3.0.html).

## Features
- Create a SQLite based video database for your movies or series
- Analyze files using ffprobe
- Fetch information from themoviedb.org

## Requirements
- Microsoft PowerShell 6 or newer:
  https://learn.microsoft.com/en-us/powershell/scripting/install/installing-powershell-on-windows
- ffprobe from ffmpeg suite:
  https://github.com/BtbN/FFmpeg-Builds/releases
- themoviedb.org API key:
  https://www.themoviedb.org/signup

## Installation
No installation needed, just run "VDB.ps1" from an PowerShell 6 or newer.

## Usage
For the first start I recommend to begin with just a few video files. Create two base folders for movies and series, i.e.
F:\Movies
F:\Series
where F: is an USB hard drive containing your videos.
You should not change the language after you have analyzed your videos. The program will fetch the movie information
in the given language, but it will not keep text in multiple languages in the database.

In the movie folder you may create subfolders, i.e. for collections or for your favourites.
In the series folder create subfolders for each series.

The files must be named in a way that the program can search them on themoviedb.org. So a good idea is to search the
video manually and rename the video file to the name and the release year, i.e. for
"Indiana Jones and the Last Crusade"
rename the file to
"Indiana Jones and the Last Crusade (1989).mp4"
This might be a lot of work depending on the amount of movie files, but currently there is no other detection possible.
If there are multiple videos found it is possible to name a file containing the ID from the URL:
"Indiana Jones and the Last Crusade (1989) [TMDBID=89].mp4"

Also be sure to check the URL that it begins with
"https://www.themoviedb.org/movie/"
identifying the result as a movie.
If the URL begins with
"https://www.themoviedb.org/tv/"
is is a series. Be sure to put the files in the right folder structure.

For series it is recommended to also have the release year in the folder name, i.e. for
Star Trek - The Original Series
"F:\Series\SciFi\Star Trek (1966)\"
For episodes the naming convention from TMDB should be used:
Beginning with "s" for "season", followed by a number, depending on the series two digits are recommended. Than add an
"e" for "episode" and additional a number, also depending on the series, one or more digits. If an epsiode was released
as a double length, sometimes the first episode of a season, but it was broadcasted in two parts you can handle this
with an additional "p" for "part" and also a number.
To have a clear overview keep the convention over the complete series.
For the first episode of the first Star Trek season the file could be named like this:
"s01e001 The Man Trap.mp4"
If this would be split into two parts, it could be named
"s01e001p1 The Man Trap.mp4"
"s01e001p2 The Man Trap.mp4"

Also for series it is possible to name the series folder using the ID:
"F:\Series\SciFi\Star Trek (1966) [TMDBID=253]\"

File extensions / video types:
Depending on your record or conversion you may use a different file extentions as ".mp4". Currently I have no clear
information which file types are supported by ffprobe.exe from the ffmpeg project. But the following files will be
analyzed:
- .mp4
- .m4v
- .mkv
- .mpeg
- .mpg
- .avi
- .webp
- .ts
If there are other used file types you can give it a try by adding them into the array:
$fileExtensions = @( "*.mp4", "*.m4v", "*.mkv", "*.mpeg", "*.mpg", "*.avi", "*.webp", "*.ts" )

In the configuration dialog you can choose to use the drive label for movies and also for series. This have been
added due to the feature of Windows to use an other drive letter if the last one was currently used. I.e. if the
hard drive usually use letter "F:" but currently a thumb drive uses this letter "F:" and the movie drive is connected
it gets a new letter. In this case analyzing the folder wouldn't find the files and mark them as missing. To avoid
this the drive label can be used and the program will check if any drive using this leeter is found upon start and
uses the right drive letter. Of course, don't use two different drives with the same label.

For the first time you also can limit the scans to get familar with, but I recommend to create a new movie and a
new series folder structure and add only a few files. As the folder size gets filled up it is also possible to
create a structure like
F:\Movies
F:\Movies.ok
F:\Series
F:\Series.ok
where you move the files to the default "Movies" or "Series" structure. After the files have been analyzed
successfully move them to "Movies.ok" or "Series.ok". End the program, delete the "Video.db" file and start
over again. Or let the files be analyzed and delete the missing files in the program.
After all the files have been sorted, named and analyzed move them into the right folder and analyze again.

FFprobe can be downloaded using the "?" in the configuration dialog and the file has to be selected. The program
will not work without it. It might be possible to use Windows API functions which does the same analyze as
used currently from FFProbe.exe but I didn't found an easy way to implement this functionality.

Also, as written above, an API key for themoviedb.org should be available to fetch data from there. Currently
no other video database is supported. You can request an API key on the themoviedb.org home page.

The load adult content checks for adult texts and only load them if this flag has been set. But it depends more
on the movies.

The program itself has in the upper part:
Configuration button:
Open the configuration dialog

Aalyze button:
Start analyzing the existing files. If a file is not found any more it will be marked red. You can delete files
but keep the information in the database so you can see you had this video but deleted it because it was not
worth to keep it or the recorded file was damaged.

Search field and the search button:
Enter a part of a file or an actor/actress you are looking for.

Duplicates button:
Use this to check for duplicate files based on the names or the used ID from TMDB. Click the duplicates button
again to disable it and load all video files from database.

TMDB button:
If a description has not been found correctly, i.e. a similar named movies was fetched, you can enter the correct
ID from TMDB here. Also to avoid a wrong recognition in future you can use the rename button to rename the file.

Delete button:
Remove the file from the database. It is possible to remove the file also from the hard disk but this is currently
disabled. I've added a functionality to move the file to the Windows recycle bin but I also like to add a part in
the configuration to enable or disable the file deletion.

Eport button:
Export all video files into a "export.csv" file, if the duplicate check is active the currently visible data will
be exported into "exportDouble.csv". Both files will be located in the directory where the PowerShell script is
stored.

Rename button:
Rename the current video file, given a suggestion from file name, release year and TMDB ID.

Move button:
Move file to anothe directory, update the database. Not yet implemented.

Rescan button:
Rescan the selected file, i.e. after a new version in full HD have been recorded.

Fill up button:
For series, fill up all episodes, using the convention:
"sAAeBBB name.mp4"
and mark the files as missing. This might help for checking if all episodes of a series have been recorded.

In the middle part the list of video files and some basic information is diplayed. If a video is selected the
lower part will be updated with details.
