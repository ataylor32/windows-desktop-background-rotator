#Requires -Version 2

<#
.SYNOPSIS
	Sets the desktop background to a randomly chosen image out of a pool of
	images.
.DESCRIPTION
	Sets the desktop background to a randomly chosen image out of a pool of
	images. Optionally, you can use different pools of images, dependent on the
	date (useful for having Christmas desktop backgrounds appear only in
	December, for example). A pool will go through each of its images one time
	and then start over. Pool usage history is saved in an XML file in the same
	folder as this script. Intended to be run as a scheduled task.
.INPUTS
	None. You cannot pipe objects to this script.
.OUTPUTS
	None. This script does not generate any output.
.LINK
	https://github.com/ataylor32/windows-desktop-background-rotator
#>

[CmdletBinding()]
Param()

# Define the functions for each pool of images.

Function Get-Default-Pool {
	Get-ChildItem "$([Environment]::GetFolderPath(`"MyPictures`"))\Desktop Backgrounds" -Recurse
}

# Define which dates should use a pool other than the default pool.

$Dates = @()

# Determine which pool to use based on today's date.

$PoolToUse = "Default"
$Today = (Get-Date)
$ThisYearIsALeapYear = [System.DateTime]::IsLeapYear($Today.Year)

ForEach ($DateData In $Dates) {
	If (
		$ThisYearIsALeapYear -Eq $False -And (
			$DateData.StartDate -Eq "02-29" -Or
			$DateData.EndDate -Eq "02-29"
		)
	) {
		If ($DateData.StartDate -Eq "02-29" -And $DateData.EndDate -Eq "02-29") {
			# The user is probably intending to have this pool used only if the
			# date is exactly February 29th.

			Continue
		}

		If ($DateData.StartDate -Eq "02-29") {
			$DateData.StartDate = "02-28"
		}
		Else {
			$DateData.EndDate = "02-28"
		}
	}

	$StartDate = [DateTime]"$($Today.Year)-$($DateData.StartDate)"
	$EndDate = [DateTime]"$($Today.Year)-$($DateData.EndDate) 11:59:59 PM"

	If ($Today -Ge $StartDate -And $Today -Le $EndDate) {
		$PoolToUse = $DateData.Pool
		Break
	}
}

Write-Verbose "Pool that will be used: $PoolToUse"

# Get the pool, remove any non-images from it, and shuffle it.

$FunctionName = "Get-$($PoolToUse)-Pool"

Try {
	$Pool = Invoke-Expression $FunctionName
}
Catch [System.Management.Automation.CommandNotFoundException] {
	Throw "The `"$FunctionName`" function does not exist"
}

If ($Pool -Eq $Null) {
	Throw "The `"$FunctionName`" function returned null"
}

$Pool = $Pool | Where-Object {".bmp", ".gif", ".jpeg", ".jpg", ".png" -Eq $_.Extension}

If ($Pool -Eq $Null) {
	Throw "The `"$FunctionName`" function did not return any images"
}

Write-Verbose "Total images in the pool: $($Pool.Length)"

$Pool = $Pool | Sort-Object {Get-Random}

# Pick an unused image out of the pool. If every image in the pool has
# already been used, reset the usage history for the pool. Once an image
# has been picked, save it to the pool's usage history.

$XMLPath = (Split-Path $SCRIPT:MyInvocation.MyCommand.Path -Parent) + "\Rotate-Desktop-Background.xml"

Try {
	$PoolUsages = Import-Clixml $XMLPath
}
Catch {
	$PoolUsages = @{}
}

If ($PoolUsages.ContainsKey($PoolToUse) -Eq $False) {
	$PoolUsages.Add($PoolToUse, @())
}

ForEach ($Image In $Pool) {
	If ($PoolUsages.$PoolToUse -NotContains $Image.FullName) {
		$ImageToUse = $Image
		Break
	}
}

If ($ImageToUse -Eq $Null) {
	Write-Verbose "All of the images in the `"$PoolToUse`" pool have already been used. Starting over."
	$PoolUsages.$PoolToUse = @()
	$ImageToUse = $Pool | Get-Random
}

Write-Verbose "Image that will be used: $($ImageToUse.FullName)"

$PoolUsages.$PoolToUse += $ImageToUse.FullName
$PoolUsages | Export-Clixml $XMLPath

# If the image is in a format that Windows does not handle natively,
# convert it to a JPEG.

If (@(".bmp", ".jpeg", ".jpg") -NotContains $ImageToUse.Extension) {
	[Void][Reflection.Assembly]::LoadWithPartialName("System.Drawing")
	$JPGPath = [System.IO.Path]::GetTempPath() + [Guid]::NewGuid() + ".jpg"
	Write-Verbose "Converting the image to JPEG format, saving to `"$JPGPath`""
	$Image = [System.Drawing.Image]::FromFile($($ImageToUse.FullName))
	$Image.Save($JPGPath, [System.Drawing.Imaging.ImageFormat]::Jpeg)
	$Image.Dispose()
	$ImageToUse = Get-ChildItem $JPGPath
}

# Set the image as the desktop background using the code from
# http://stackoverflow.com/a/9440226

Add-Type @"
using System;
using System.Runtime.InteropServices;
using Microsoft.Win32;
namespace Wallpaper
{
	public enum Style : int
	{
		Tile, Center, Stretch, NoChange
	}
	public class Setter {
		public const int SetDesktopWallpaper = 20;
		public const int UpdateIniFile = 0x01;
		public const int SendWinIniChange = 0x02;
		[DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
		private static extern int SystemParametersInfo (int uAction, int uParam, string lpvParam, int fuWinIni);
		public static void SetWallpaper ( string path, Wallpaper.Style style ) {
			SystemParametersInfo( SetDesktopWallpaper, 0, path, UpdateIniFile | SendWinIniChange );
			RegistryKey key = Registry.CurrentUser.OpenSubKey("Control Panel\\Desktop", true);
			switch( style )
			{
				case Style.Stretch :
					key.SetValue(@"WallpaperStyle", "2") ;
					key.SetValue(@"TileWallpaper", "0") ;
					break;
				case Style.Center :
					key.SetValue(@"WallpaperStyle", "1") ;
					key.SetValue(@"TileWallpaper", "0") ;
					break;
				case Style.Tile :
					key.SetValue(@"WallpaperStyle", "1") ;
					key.SetValue(@"TileWallpaper", "1") ;
					break;
				case Style.NoChange :
					break;
			}
			key.Close();
		}
	}
}
"@

Write-Verbose "Setting the image as the desktop background"
[Wallpaper.Setter]::SetWallpaper($ImageToUse.FullName, 2)
