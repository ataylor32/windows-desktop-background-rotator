# Windows Desktop Background Rotator

A PowerShell 2.0+ script that sets the desktop background to a randomly chosen
image out of a pool of images. Optionally, you can use different pools of
images, dependent on the date (useful for having Christmas desktop backgrounds
appear only in December, for example). A pool will go through each of its
images one time and then start over. Pool usage history is saved in an XML file
in the same folder as this script. Intended to be run as a scheduled task.

## Instructions

1. [Download the script](https://raw.githubusercontent.com/ataylor32/windows-desktop-background-rotator/master/Rotate-Desktop-Background.ps1)

1. If your desktop backgrounds are stored in a folder named
   "Desktop Backgrounds" within your "Pictures" folder, you can go to the next
   step. Otherwise, you will need to open the script in an editor and modify
   the path used in the `Get-Default-Pool` function to point to the correct
   location.

1. If you only want to pull from one pool of images regardless of the date, you
   can go to the next step. Otherwise, follow the steps in the "Using Different
   Pools of Images Depending on the Date" section below before going to the
   next step.

1. If you want your desktop backgrounds' fit / position to be "Stretch", you
   can go to the next step. Otherwise, you will need to open the script in an
   editor and modify the last line of code. Specifically, the second parameter
   passed to `[Wallpaper.Setter]::SetWallpaper`. 0 is "Tile", 1 is "Center", 2
   is "Stretch", and 3 is "NoChange".

1. If you want this script to be run automatically on a schedule, set up a
   scheduled task to do so.

## Using Different Pools of Images Depending on the Date

There are three steps to setting up an additional pool of images to pull from
during a specific date range:

1. Define a function

1. Add an entry to the `$Dates` array

1. If necessary, update the `Get-Default-Pool` function to exclude the images
   that should only be in this additional pool

The following example assumes this file structure:

```
...
└── Pictures
    └── Desktop Backgrounds
        ├── Holidays
        │   ├── Christmas
        │   │   ├── christmas_desktop_background1.jpg
        │   │   ├── christmas_desktop_background2.jpg
        │   │   └── christmas_desktop_background3.jpg
        │   └── Thanksgiving
        │       ├── thanksgiving_desktop_background1.jpg
        │       ├── thanksgiving_desktop_background2.jpg
        │       └── thanksgiving_desktop_background3.jpg
        ├── regular_desktop_background1.jpg
        ├── regular_desktop_background2.jpg
        └── regular_desktop_background3.jpg
```

### Example Part 1: Christmas

Suppose you wanted Christmas desktop backgrounds to be used from December 1st
through December 25th each year. You would first define a function such as
`Get-Christmas-Pool`:

```powershell
Function Get-Christmas-Pool {
	Get-ChildItem "$([Environment]::GetFolderPath(`"MyPictures`"))\Desktop Backgrounds\Holidays\Christmas"
}
```

Next, you would add an entry to the `$Dates` array, like this:

```powershell
$Dates = @(
	@{
		"StartDate" = "12-01"
		"EndDate" = "12-25"
		"Pool" = "Christmas"
	}
)
```

Note that whatever you set `Pool` to must match the function name. Since our
function is named `Get-Christmas-Pool`, we set `Pool` to `Christmas`.

Finally, because of the file structure of our desktop backgrounds, we will need
to update the `Get-Default-Pool` function to exclude the holiday desktop
backgrounds, like this:

```powershell
Function Get-Default-Pool {
	Get-ChildItem "$([Environment]::GetFolderPath(`"MyPictures`"))\Desktop Backgrounds" -Recurse | Where-Object {! $_.PSIsContainer -And $_.FullName -NotMatch "Holidays"}
}
```

### Example Part 2: Thanksgiving

Now suppose you wanted Thanksgiving desktop backgrounds to be used on
Thanksgiving Day each year. You would first define a function such as
`Get-Thanksgiving-Pool`:

```powershell
Function Get-Thanksgiving-Pool {
	Get-ChildItem "$([Environment]::GetFolderPath(`"MyPictures`"))\Desktop Backgrounds\Holidays\Thanksgiving"
}
```

Next, you would add an entry to the `$Dates` array. But this time, it will need
to be done differently since the date of Thanksgiving Day varies from year to
year. Here is an example of what you might do:

```powershell
$Dates = @(
	...
)

$Today = (Get-Date)

If ($Today.Month -Eq 11 -And $Today.Day -Ge 22 -And $Today.Day -Le 28) {
	$DayOfWeekOnFirst = [Int]([DateTime]"$($Today.Year)-11-01").DayOfWeek
	$FourthThursdayDate = 22 + (11 - $DayOfWeekOnFirst) % 7
	$Dates += @{
		"StartDate" = "11-$FourthThursdayDate"
		"EndDate" = "11-$FourthThursdayDate"
		"Pool" = "Thanksgiving"
	}
}
```

Since the `Get-Default-Pool` function is already set up to exclude our holiday
desktop backgrounds, we are done.
