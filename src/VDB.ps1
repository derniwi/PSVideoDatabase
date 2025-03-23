# Video Database
#
# Short description:
#  Generate a local SQLite database for your video collection.
#  Show a list of movies and series, collect information from file system and
#  ffprobe.exe, retrieve data from themoviedb.org TMDB (you need an API key to
#  retrieve data).
#
# v1.0.0 (2024-10-22):
#  first version
# v1.0.1 (2024-10-25):
#  Extended with average vote from TMDB
# v1.1.0 (2024-10-31):
#  - Re-creation, because movies and series are handled different on TMDB. This causes the
#	 first table structure won't work and needs to be changed
#  - Integration of collection data like for Star Trek movies
#  - Integration of general series information and detailed information for each episode
# v1.2.0 (2024-12-03):
#  - Database structure completly new designed
#  - Added configuration dialog
#  - Handle GUI language / texts using translation and JSON file
#  - Added link for downloading ffmpeg package, where ffprobe.exe is included
#  - Added link to themoviedb.org registration web page
#  - Load and save configuration to JSON file
#  - Get results for adult videos, configurable in script
#  - Check for double files: check title, filename and TMDB ID
#  - Export button: export data to CSV file.
# v1.2.1 (2024-12-27):
#  - Database structure extended and modified
#  - Store file size additionally in bytes and audio languages in database
#  - Delete button: delete movies and according data from database.
#	- Added "Are you sure?" message box before deleting
#	- Include check for FileExists. If file exists as of last scan and path is available ask for physical deletion
#  - TMDB button: manually change TMDB ID. Remove old data in any table if they exists for the old ID.
#	Retreive new data for the new ID.
#	- Ok for movies
#	- ToDo: What about series? Should the basic series ID which is identified by the folder be received?
#	  And in this case all files within this folder has to be updated?
#  - Rename button: rename file, suggestions from TMDB for title, replace invalid filesystem characters
#  - Rescan button: rescan selected videos with ffprobe
#  - During scan if file exists also check file size. If size has changed re-scan file. If TMDB id exists, do not reload TMDB information
# v1.2.2 (2025-01-05):
#  - Added adult flag into configuration
#  - Add analyze file limit text and check box to configuration
#  - Check winsqlite3.dll-PowerShell
#	- Added path to configuration
#	- Check file existence
#	- Unblock files downloaded from the internet
#  - Added search text highlighting in actors box
#  - Added database indexes
# v1.3.0 (2025-01-24):
#  - Changed from Rene Nyffeneggers winsqlite3.dll-PowerShell tool pack to an own
#    written P/Invoke part. I had some issues with his class, often the SELECT
#    statements did not return all results. I had checked this with a complete
#    simple script, but if there is a select over some joined tables the
#    result varies each time the script has been executed.
#  - Added try-catch blocks around all database parts
#  - Added save and load windows position and state
#  - Added move file to recycle bin instead of direct deletion
#  - Added button for fill up series with missing entries (file exists = 0) and also collections
# v1.3.1 (2025-02-09):
# - Added adult flag in database
# - Added information where a video belongs to from DB (it's already there) to the detail view
# - Improved the marking of actors if more than one entry matches the text in the search box


# To Do / Ideas:
# - hide adult entries if flag is disabled
#   => extend database and add boolean adult value
# - Add comments to code
# - Add flag / variable to allow or deny deleting files
# - Path and file name for CSV export
#   => open
# - Backup video files?
#   => needs additional variables and handling for both movies and series folder
# - Backup database and configuration file?
# - Move button: move file to another folder, updating path in database
# - Add files by Drag'n'Drop, check if in sub directory of movie or series
# - Check database calls and edit exceptions in both type definition and catch part,
#   use the custom exception class
# - Add transaction for insert and update parts
# - Check deletion of (multiple) entries: DGV updated?
# - Add season and episode numbers from TMDB to detail view title
# - Add extension selection for fill up file names to configuration
# - Add supported video extension list to configuration or configuration file
# - Change TMDB ID for series: should this be done for both the series and the current episode?
#   Does it make sense to change the series ID itself since this should cause to update all current episodes


######################################################
# Clear PowerShell screen
######################################################
Clear-Host

# Set strict mode
Set-StrictMode -Version Latest

# First check if PowerShell version 6 or newer is running
if ($PSVersionTable.PSVersion.Major -lt 6) {
    Write-Warning "PowerShell Version 6 or newer required."
	Read-Host -Prompt "Press Enter to exit."
	Exit
}

# Add assemblies
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Globalization
Add-Type -AssemblyName System.Windows.Forms

# Add type and C++ class for the SQLite interface using winsqlite3.dll:
# Several sources from the internet have been checked
Add-Type -TypeDefinition @"
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

// Create helper class for SQLite
public class SQLiteHelper {
	// SQLite result codes from https://www.sqlite.org/rescode.html
	public const Int32 SQLITE_OK         =   0;
	public const Int32 SQLITE_ERROR      =   1;
	public const Int32 SQLITE_INTERNAL   =   2;
	public const Int32 SQLITE_PERM       =   3;
	public const Int32 SQLITE_ABORT      =   4;
	public const Int32 SQLITE_BUSY       =   5;
	public const Int32 SQLITE_LOCKED     =   6;
	public const Int32 SQLITE_NOMEM      =   7;
	public const Int32 SQLITE_READONLY   =   8;
	public const Int32 SQLITE_INTERRUPT  =   9;
	public const Int32 SQLITE_IOERR      =  10;
	public const Int32 SQLITE_CORRUPT    =  11;
	public const Int32 SQLITE_NOTFOUND   =  12;
	public const Int32 SQLITE_FULL       =  13;
	public const Int32 SQLITE_CANTOPEN   =  14;
	public const Int32 SQLITE_PROTOCOL   =  15;
	public const Int32 SQLITE_EMPTY      =  16;
	public const Int32 SQLITE_SCHEMA     =  17;
	public const Int32 SQLITE_TOOBIG     =  18;
	public const Int32 SQLITE_CONSTRAINT =  19;
	public const Int32 SQLITE_MISMATCH   =  20;
	public const Int32 SQLITE_MISUSE     =  21;
	public const Int32 SQLITE_NOLFS      =  22;
	public const Int32 SQLITE_AUTH       =  23;
	public const Int32 SQLITE_FORMAT     =  24;
	public const Int32 SQLITE_RANGE      =  25;
	public const Int32 SQLITE_NOTADB     =  26;
	public const Int32 SQLITE_NOTICE     =  27;
	public const Int32 SQLITE_WARNING    =  28;
	public const Int32 SQLITE_ROW        = 100;
	public const Int32 SQLITE_DONE       = 101;
	
	// Import external functions provided by "winsqlite3.dll" which are used in this project
	// Open SQLite database
	[DllImport("winsqlite3.dll", CallingConvention = CallingConvention.Cdecl)]
	public static extern int sqlite3_open(string filename, out IntPtr db);
	
	// Execute SQL statement(s)
	[DllImport("winsqlite3.dll", CallingConvention = CallingConvention.Cdecl)]
	public static extern int sqlite3_exec(IntPtr db, string sql, IntPtr callback, IntPtr arg, out IntPtr errmsg);
	
	// Prepare a SQL statement
	[DllImport("winsqlite3.dll", CallingConvention = CallingConvention.Cdecl)]
	public static extern int sqlite3_prepare_v2(IntPtr db, string sql, int nByte, out IntPtr stmt, IntPtr tail);
	
	// Bind a text to a position in the SQL statement
	[DllImport("winsqlite3.dll", CallingConvention = CallingConvention.Cdecl)]
	public static extern int sqlite3_bind_text(IntPtr stmt, int index, IntPtr text, int n, IntPtr destructor);
	
	// Bind an integer up to a 32bit value to a position in the SQL statement
	[DllImport("winsqlite3.dll", CallingConvention = CallingConvention.Cdecl)]
	public static extern int sqlite3_bind_int(IntPtr stmt, int index, int value);
	
	// Bind a 64 bit integer to a position in the SQL statement
	[DllImport("winsqlite3.dll", CallingConvention = CallingConvention.Cdecl)]
	public static extern int sqlite3_bind_int64(IntPtr stmt, int index, long value);
	
	// Bind a double value to a position in the SQL statement
	[DllImport("winsqlite3.dll", CallingConvention = CallingConvention.Cdecl)]
	public static extern int sqlite3_bind_double(IntPtr stmt, int index, double value);
	
	// Bind a blob (untestesd) to a position in the SQL statement.
	[DllImport("winsqlite3.dll", CallingConvention = CallingConvention.Cdecl)]
	public static extern int sqlite3_bind_blob(IntPtr stmt, int index, IntPtr blob, int n, IntPtr destructor);
	
	// Process one step i.e. retrieving one result line from the execution of the SQL statement
	[DllImport("winsqlite3.dll", CallingConvention = CallingConvention.Cdecl)]
	public static extern int sqlite3_step(IntPtr stmt);
	
	// Return the ID of the last inserted row
	[DllImport("winsqlite3.dll", CallingConvention = CallingConvention.Cdecl)]
	public static extern long sqlite3_last_insert_rowid(IntPtr db);
	
	// Expand the SQL statement, retuning a string with all placeholders filled
	[DllImport("winsqlite3.dll", CallingConvention = CallingConvention.Cdecl)]
	public static extern IntPtr sqlite3_expanded_sql(IntPtr stmt);
	
	// Get the name of the column
	[DllImport("winsqlite3.dll", CallingConvention = CallingConvention.Cdecl)]
	public static extern IntPtr sqlite3_column_text(IntPtr stmt, int col);
	
	// Get the name of the column UTF-8 encoded
	[DllImport("winsqlite3.dll", CallingConvention = CallingConvention.Cdecl)]
	public static extern IntPtr sqlite3_column_text16(IntPtr stmt, int col);
	
	// Get the neame of the column UTF-16 encoded
	[DllImport("winsqlite3.dll", CallingConvention = CallingConvention.Cdecl)]
	public static extern IntPtr sqlite3_column_name(IntPtr stmt, int col);
	
	// Get the column count of the last SQL result, usefull i.e. for a "SELECT * FROM" command
	[DllImport("winsqlite3.dll", CallingConvention = CallingConvention.Cdecl)]
	public static extern int sqlite3_column_count(IntPtr stmt);
	
	// Return the numeric result code for the API call
	[DllImport("winsqlite3.dll", CallingConvention = CallingConvention.Cdecl)]
	public static extern int sqlite3_errcode(IntPtr db);
	
	// Return the UTF-8 encoded result code for the API call
	[DllImport("winsqlite3.dll", CallingConvention = CallingConvention.Cdecl)]
	public static extern IntPtr sqlite3_errmsg(IntPtr db);
	
	// Return the UTF-16 encoded result code for the API call
	[DllImport("winsqlite3.dll", CallingConvention = CallingConvention.Cdecl)]
	public static extern IntPtr sqlite3_errmsg16(IntPtr db);
	
	// Returns a pointer to the sqlite3_version[] string constant
	[DllImport("winsqlite3.dll", CallingConvention = CallingConvention.Cdecl)]
	public static extern IntPtr sqlite3_libversion();
	
	// Delete the prepared SQL statement
	[DllImport("winsqlite3.dll", CallingConvention = CallingConvention.Cdecl)]
	public static extern int sqlite3_finalize(IntPtr stmt);
	
	// Destroy the database object in memory, close the database.
	[DllImport("winsqlite3.dll", CallingConvention = CallingConvention.Cdecl)]
	public static extern int sqlite3_close(IntPtr db);
	
	// Private members:
	// Pointer to database
	private IntPtr db;
	// List of pointers to allocated memory parts
	private List<IntPtr> allocatedMemory = new List<IntPtr>();
	
	// Open the database
	public void Open(string databasePath) {
		int result = sqlite3_open(databasePath, out db);
		if (result != SQLITE_OK) {
			// Throw exception
			throw new Exception($"Error while opening database. Error code: {result}");
		}
	}
	
	// Execute a given SQL statement
	public void Exec(string sql) {
		IntPtr errmsg = IntPtr.Zero;
		int result = sqlite3_exec(db, sql, IntPtr.Zero, IntPtr.Zero, out errmsg);
		if (result != SQLITE_OK) {
			string errmsgStr = Marshal.PtrToStringAnsi(errmsg);
			throw new Exception($"Error while executing SQL statement: {errmsgStr}");
		}
	}
	
	// Prepare SQL statement
	public IntPtr Prepare(string sql) {
		IntPtr stmt = IntPtr.Zero;
		int result = sqlite3_prepare_v2(db, sql, -1, out stmt, IntPtr.Zero);
		if (result != SQLITE_OK) {
			throw new Exception($"Error while preparing SQL statement. Error code: {result}");
		}
		return stmt;
	}
	
	// Bind a text to the current SQL statement
	public void BindText(IntPtr stmt, int index, string text) {
		// SQLite needs to keep the binded data on the given address until the SQL statement
		// has been executed. For this the text will be copied in an reserved memory
		// area.
		
		// Get bytes of UTF8 encoded text
		byte[] utf8Text = System.Text.Encoding.UTF8.GetBytes(text);
		
		// Allocate memory, add one byte for the C zero byte termination
		IntPtr unmanagedPointer = Marshal.AllocHGlobal(utf8Text.Length + 1);
		
		// Copy bytes
		Marshal.Copy(utf8Text, 0, unmanagedPointer, utf8Text.Length);
		
		// Write the zero byte termination
		Marshal.WriteByte(unmanagedPointer, utf8Text.Length, 0);
		
		// Add pointer to list
		allocatedMemory.Add(unmanagedPointer);
		
		// Bind the text to the given position
		int result = sqlite3_bind_text(stmt, index, unmanagedPointer, utf8Text.Length, IntPtr.Zero);
		if (result != SQLITE_OK) {
			// Free memory if bind text fails
			Marshal.FreeHGlobal(unmanagedPointer);
			// Remove pointer from list
			allocatedMemory.Remove(unmanagedPointer);
			// Throw exception
			throw new Exception($"Error while binding text. Error code: {result}");
		}
	}
	
	// Bind integer value to the current SQL statement
	public void BindInt(IntPtr stmt, int index, int value) {
		int result = sqlite3_bind_int(stmt, index, value);
		if (result != SQLITE_OK) {
			// Throw exception
			throw new Exception($"Error while binding integer. Error code: {result}");
		}
	}
	
	// Bind 64 bit integer value to the current SQL statement
	public void BindInt64(IntPtr stmt, int index, long value) {
		int result = sqlite3_bind_int64(stmt, index, value);
		if (result != SQLITE_OK) {
			// Throw exception
			throw new Exception($"Error while binding 64 bit integer. Error code: {result}");
		}
	}
	
	// Bind double value to the current SQL statement
	public void BindDouble(IntPtr stmt, int index, double value) {
		int result = sqlite3_bind_double(stmt, index, value);
		if (result != SQLITE_OK) {
			// Throw exception
			throw new Exception($"Error while binding double. Error code: {result}");
		}
	}
	
	// Bind a blob to the current SQL statemnt
	public void BindBlob(IntPtr stmt, int index, byte[] blob) {
		// Allocate memory
		IntPtr unmanagedPointer = Marshal.AllocHGlobal(blob.Length);
		
		// Copy data
		Marshal.Copy(blob, 0, unmanagedPointer, blob.Length);
		
		// Add pointer to list
		allocatedMemory.Add(unmanagedPointer);
		
		// Bind the blob to the given position
		int result = sqlite3_bind_blob(stmt, index, unmanagedPointer, blob.Length, IntPtr.Zero);
		if (result != SQLITE_OK) {
			// Free memory if bind text fails
			Marshal.FreeHGlobal(unmanagedPointer);
			// Remove pointer from list
			allocatedMemory.Remove(unmanagedPointer);
			// Throw exception
			throw new Exception($"Error while binding blob. Error code: {result}");
		}
	}
	
	// Get the ID of the last row inserted
	public long GetLastInsertRowId() {
		return sqlite3_last_insert_rowid(db);
	}
	
	// Get the expanded SQL statement with the binded values
	public string GetExpandedSql(IntPtr stmt) {
		IntPtr expandedSqlPtr = sqlite3_expanded_sql(stmt);
		return Marshal.PtrToStringAnsi(expandedSqlPtr);
	}
	
	// Get the name of the given column from the results
	public string GetColumnName(IntPtr stmt, int col) {
		IntPtr columnNamePtr = sqlite3_column_name(stmt, col);
		return Marshal.PtrToStringAnsi(columnNamePtr);
	}
	
	// Get number of columns
	public int GetColumnCount(IntPtr stmt) {
		return sqlite3_column_count(stmt);
	}
	
	// Get error code
	public int GetErrorCode() {
		return sqlite3_errcode(db);
	}
	
	// Get error message
	public string GetErrorMessage() {
		IntPtr errmsgPtr = sqlite3_errmsg(db);
		return Marshal.PtrToStringAnsi(errmsgPtr);
	}
	
	// Get error message as UTF-16
	public string GetErrorMessage16() {
		IntPtr errmsg16Ptr = sqlite3_errmsg16(db);
		return Marshal.PtrToStringUni(errmsg16Ptr);
	}
	
	// Get version of winsqlite3.dll
	public string GetLibVersion() {
		IntPtr libversionPtr = sqlite3_libversion();
		return Marshal.PtrToStringAnsi(libversionPtr);
	}
	
	// Process one step, evaluate the SQL statement
	public int Step(IntPtr stmt) {
		int result = sqlite3_step(stmt);
		if ((result != SQLITE_OK) && (result != SQLITE_DONE)) {
			// Throw exception
			throw new Exception($"Error while evaluating SQL statement. Error code: {result}");
		}
		return result;
	}
	
	// Process one step and get all values returned
	public (int, string[]) StepAndGetRow(IntPtr stmt) {
		int columnCount = sqlite3_column_count(stmt);
		return StepAndGetRow(stmt, columnCount);
	}
	
	// Process one step and get values returned
	public (int, string[]) StepAndGetRow(IntPtr stmt, int columnCount) {
		int result = sqlite3_step(stmt);
		if (result == SQLITE_ROW) {
			string[] row = new string[columnCount];
			for (int i = 0; i < columnCount; i++) {
				IntPtr textPtr = sqlite3_column_text16(stmt, i);
				row[i] = Marshal.PtrToStringUni(textPtr);
			}
			return (result, row);
		}
		
		if ((result != SQLITE_OK) && (result != SQLITE_DONE)) {
			// Throw exception
			throw new Exception($"Error while evaluating SQL statement. Error code: {result}");
		}
		
		return (result, null);
	}
	
	// Finalize, free allocated memory and clear memory pointers
	public void Finalize(IntPtr stmt) {
		int result = sqlite3_finalize(stmt);
		foreach (IntPtr ptr in allocatedMemory) {
			Marshal.FreeHGlobal(ptr);
		}
		allocatedMemory.Clear();
		if (result != SQLITE_OK) {
			// Throw exception
			throw new Exception($"Error while finalizing SQLite. Error code: {result}");
		}
	}
	
	// Close database
	public void Close() {
		sqlite3_close(db);
	}
}
"@

######################################################
# Type definiton for moving a file into the Windows recycle bin
# Currenty PowerShell does not have such a functionality yet.
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;

public class FileOperation
{
    [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
    public static extern int SHFileOperation(ref SHFILEOPSTRUCT FileOp);
	
    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    public struct SHFILEOPSTRUCT
    {
        public IntPtr hwnd;
        public uint wFunc;
        public string pFrom;
        public string pTo;
        public ushort fFlags;
        public bool fAnyOperationsAborted;
        public IntPtr hNameMappings;
        public string lpszProgressTitle;
    }
	
    public const uint FO_DELETE = 3;
    public const ushort FOF_ALLOWUNDO = 0x0040;
    public const ushort FOF_NOCONFIRMATION = 0x0010;
}
"@

# Function to move a file or folder into the Windows recycle bin
function Move-ToRecycleBin {
	param (
		[string]$Path
	)
	
	$fileOp = [FileOperation+SHFILEOPSTRUCT]::new()
	$fileOp.wFunc = [FileOperation]::FO_DELETE
	$fileOp.pFrom = $Path + "`0" + "`0"
	$fileOp.fFlags = [FileOperation]::FOF_ALLOWUNDO -bor [FileOperation]::FOF_NOCONFIRMATION
	
	# Move file to recycle bin
	$result = [FileOperation]::SHFileOperation([ref]$fileOp)
	
	return $result
}

######################################################
# Define constants
Set-Variable VDBVERSION -option Constant -value "1.3.1"

Set-Variable URLFFMPEG -option Constant -value "https://github.com/BtbN/FFmpeg-Builds/releases/download/latest/ffmpeg-master-latest-win64-gpl.zip"
Set-Variable URLTMDBSIGNUP -option Constant -value "https://www.themoviedb.org/signup"

# FFprobe errors
Set-Variable ERRORFFPOBEGUNKNOWN -option Constant -value 1
Set-Variable ERRORFFPROBERCNOTZERO -option Constant -value 2
Set-Variable ERRORFFPROBENOSTREAMS -option Constant -value 3
Set-Variable ERRORFFPROBENOVIDEOSTREAM -option Constant -value 4

# Data grid view columns
Set-Variable GRIDVIEWCOLUMNID -option Constant -value 0
Set-Variable GRIDVIEWCOLUMNTITLE -option Constant -value 1
Set-Variable GRIDVIEWCOLUMNFILENAME -option Constant -value 2
Set-Variable GRIDVIEWCOLUMNFILEPATH -option Constant -value 3
Set-Variable GRIDVIEWCOLUMNBELONGSTO -option Constant -value 4
Set-Variable GRIDVIEWCOLUMNFILESIZE -option Constant -value 5
Set-Variable GRIDVIEWCOLUMNFILESIZEMB -option Constant -value 6
Set-Variable GRIDVIEWCOLUMNRESOLUTION -option Constant -value 7
Set-Variable GRIDVIEWCOLUMNVIDEOCODEC -option Constant -value 8
Set-Variable GRIDVIEWCOLUMNAUDIOTRACKS -option Constant -value 9
Set-Variable GRIDVIEWCOLUMNAUDIOCHANNELS -option Constant -value 10
Set-Variable GRIDVIEWCOLUMNAUDIOLAYOUTS -option Constant -value 11
Set-Variable GRIDVIEWCOLUMNAUDIOLANGUAGES -option Constant -value 12
Set-Variable GRIDVIEWCOLUMNDURATION -option Constant -value 13
Set-Variable GRIDVIEWCOLUMNFILEEXISTS -option Constant -value 14
Set-Variable GRIDVIEWCOLUMNTMDBID -option Constant -value 15
Set-Variable GRIDVIEWCOLUMNVOTE -option Constant -value 16
Set-Variable GRIDVIEWCOLUMNVIDEOTYPE -option Constant -value 17
Set-Variable GRIDVIEWCOLUMNADULTCONTENT -option Constant -value 18
Set-Variable GRIDVIEWCOLUMNPATHANDFILENAME -option Constant -value 19

######################################################
# Define and declare variables
# Current script location
$scriptDirectory = Split-Path -Path $PSCommandPath

######################################################
# Supported file suffixes, since there are many file extions supported by ffprobe this list might be extended
$fileExtensions = @( "*.mp4", "*.m4v", "*.mkv", "*.mpeg", "*.mpg", "*.avi", "*.webp", "*.ts" )
# If series in the database should be filled up with missing episodes a file name has to be created and also a file extension.
# This extension will be used.
$fileExtensionForFillup = ".mp4"

######################################################
# Classes
######################################################
# Class for handling folder, driver letter, volume label
class PathInfo {
	# Class properties
	hidden [string]$Path
	hidden [string]$DriveLetter # Drive letter without colon
	hidden [string]$DriveLabel
	hidden [bool]$IsNetworkPath
	hidden [bool]$IsAvailable
	
	# Constructors
	PathInfo([string]$Path) {
		$this.Path = $Path
		$this.DriveLetter = $null
		$this.DriveLabel = $null
		$this.IsNetworkPath = $false
		$this.IsAvailable = $false
		$this.Initialize($Path)
	}
	
	PathInfo([string]$Path, [string]$VolumeLabel) {
		$this.Path = $Path
		$this.DriveLetter = $null
		$this.DriveLabel = $null
		$this.IsNetworkPath = $false
		$this.IsAvailable = $false
		$this.InitializeWithVolumeLabel($Path, $VolumeLabel)
	}
	
	# Initialize method
	[void]Initialize([string]$Path) {
		# Check for local, mounted or UNC path
		if ($Path -match "^[a-zA-Z]:\\") {
			# Path using drive letter
			$this.Path = $Path
			# Extract drive letter
			$this.DriveLetter = $this.Path.Substring(0, 1)
			$this.IsNetworkPath = $false
			try {
				# Try to get volume using drive letter
				$volume = Get-Volume -DriveLetter $this.DriveLetter -ErrorAction Stop
				if ($volume) {
					# If the volume is available, get label
					$this.DriveLabel = $volume.FileSystemLabel
				} else {
				# Try to get the label using WMI
					$volume = Get-WmiObject -Query "SELECT * FROM Win32_LogicalDisk WHERE DeviceID = '$($this.DriveLetter):'"
					if (-not($volume -eq $null)) {
						$this.DriveLabel = $volume.VolumeName
					} else {
						# Since this text is used inside this class it can't be read from the language file.
						$this.DriveLabel = "Drive not mounted"
					}
				}
			} catch {
				# Try to get the label using WMI
				$volume = Get-WmiObject -Query "SELECT * FROM Win32_LogicalDisk WHERE DeviceID = '$($this.DriveLetter):'"
				
				if (-not($volume -eq $null)) {
					$this.DriveLabel = $volume.VolumeName
				} else {
						# Since this text is used inside this class it can't be read from the language file.
					$this.DriveLabel = "Drive not mounted"
				}
			}
		} elseif ($Path -match "^\\\\") {
			# UNC (Universal Naming Convention) network path
			$this.Path = $Path
			$this.DriveLetter = $null
			$this.IsNetworkPath = $true
			$this.DriveLabel = ""
		} else {
			$this.Path = $Path
			$this.DriveLetter = $null
			$this.IsNetworkPath = $false
			# Since this text is used inside this class it can't be read from the language file.
			$this.DriveLabel = "Unknown"
		}
		$this.CheckAvailability()
	}
	
	# Initialize method with volume label
	[void]InitializeWithVolumeLabel([string]$Path, [string]$VolumeLabel) {
		try {
			# Try to get the volume using the label
			$volume = Get-Volume -FileSystemLabel $VolumeLabel -ErrorAction Stop
			if ($volume) {
				$this.DriveLetter = $volume.DriveLetter
				$this.DriveLabel = $volume.FileSystemLabel
				$this.Path = $this.Path -replace '^[a-zA-Z]:', "$($this.DriveLetter):"
				$this.IsNetworkPath = $false
			}
		} catch {
			$this.Path = $Path
			$this.DriveLetter = $this.Path.SubString(0, 1)
			$this.DriveLabel = $VolumeLabel
			$this.IsNetworkPath = $false
			$this.IsAvailable = $false
		}
		$this.CheckAvailability()
	}
	
	# Check if the path is available
	[void]CheckAvailability() {
		$this.IsAvailable = Test-Path -LiteralPath $this.Path
	}
	
	# Get current path
	[string] GetPath() {
		return $this.Path
	}
	
	# Re-intialize
	[void] SetPath([string]$path) {
		$this.Initialize($path)
	}
	
	# Change the path, but keep the current drive letter.
	[void] ChangePath([string]$path) {
		$this.Path = $path
		if ($this.Path.SubString(0, 1) -ne $this.DriveLetter) {
			$this.Path = $this.DriveLetter + $this.Path.SubString(1)
		}
	}
	
	# Get current drive letter without colon
	[string] GetDriveLetter() {
		return $this.DriveLetter
	}
	
	# Get current drive label
	[string] GetDriveLabel() {
		return $this.DriveLabel
	}
	
	# Return $True if the path is a UNC path
	[bool] GetIsNetworkPath() {
		return $this.IsNetworkPath
	}
	
	# Check and return availability of path
	[bool] GetIsAvailable() {
		$this.CheckAvailability()
		return $this.IsAvailable
	}
}


# Class for custom excpetion
class CustomException: System.Exception {
	hidden [string]$myMessage
	hidden [int]$myNumber
	
	CustomException($_Message, $_myMessage, $_myNumber) : base($_Message) {
		$this.myMessage = $_myMessage
		$this.myNumber = $_myNumber
	}
	
	CustomException($_Message, $_myMessage) : base($_Message) {
		$this.myMessage = $_myMessage
		$this.myNumber = 0
	}
	
	CustomException($_Message) : base($_Message) {
		$this.myMessage = ""
		$this.myNumber = 0
	}
}

######################################################
# Functions
######################################################
# Configuration
# Load configuration from file
function Load-Config {
	# Check if the configuration file exists
	if (Test-Path -LiteralPath $configFilePath) {
		$config = Get-Content -Raw -Path $configFilePath | ConvertFrom-Json
		# Check if stored language code is valid means does the language code
		# exists in language file?
		if (-not ($languages.PSObject.Properties.Name -Contains $config.Language)) {
			# No, change to default language
			$config.Language = $defaultLanguage
		}
		return $config
	} else {
		# Configuration file does not exists, return default values
		return $configDefaults
	}
}

# Save configuration into file
function Save-Config {
	param (
		$config
	)
	
	# Load saved configuration
	$savedConfig = Get-Content -Raw -Path $configFilePath | ConvertFrom-Json
	
	# Save current Window position and status
	$config.Width = $form.Width
	$config.Height = $form.Height
	$config.Top = $form.Top
	$config.Left = $form.Left
	# Top and left values are set to -32000 if the window is minimized. Don't store this values
	if ($config.Top -eq -32000) {
		$config.Top = 0
	}
	if ($config.Left -eq -32000) {
		$config.Left = 0
	}
	
	# Get current window state, ignore minimized when closing
	$config.WindowState = $form.WindowState
	if (($config.WindowState -ne [System.Windows.Forms.FormWindowState]::Maximized) -and ($config.WindowState -ne [System.Windows.Forms.FormWindowState]::Normal)) {
		$config.WindowState = [System.Windows.Forms.FormWindowState]::Normal
	}
	
	$config.MonitorIndex = [System.Windows.Forms.Screen]::AllScreens | ForEach-Object { $_ } | Where-Object { $_.Bounds.Contains($form.Bounds) } | Select-Object -First 1 | ForEach-Object { [Array]::IndexOf([System.Windows.Forms.Screen]::AllScreens, $_) }
	
	# Create compressed JSON array for comparing both configuration
	$currentConfigJson = $config | ConvertTo-Json -Compress
	$savedConfigJson = $savedConfig | ConvertTo-Json -Compress
	
	# Compare current configuration and saved configuration
	if ($currentConfigJson -ne $savedConfigJson) {
		# Configuration has changed, save
		$config | ConvertTo-Json -Depth 3 | Set-Content -Path $configFilePath
	}
}

# Check configuration
function Check-Configuration {
	# Loop until the configuration is valid
	do {
		# Assume configuration is fine
		$configValid = $true
		
		# Include adult content?
		if ($config.GetAdultContent) {
			$global:tmdbAdult = "include_adult=true"
		} else {
			$global:tmdbAdult = "include_adult=false"
		}
		
		# Check if language is set
		if ([string]::IsNullOrEmpty($config.Language)) {
			$configValid = $false
			# Write error in english because we do not have a language set yet
			[System.Windows.Forms.MessageBox]::Show("ERROR: No language set.")
		} else {
			# Check if movie or series folder is set
			if ([string]::IsNullOrEmpty($config.MovieFolder) -and [string]::IsNullOrEmpty($config.SeriesFolder)) {
				# Neither movie nor series folder is set.
				$errorMsg = Get-Translation -language $config.Language -key "Error.Config.MovieAndSeriePathNotSet"
				# Open error message box
				[System.Windows.Forms.MessageBox]::Show($errorMsg)
				$configValid = $false
			} else {
				# Check if movie drive label should be used for identifying
				if ($config.UseDriveLabelForMovies) {
					if (-not ([String]::IsNullOrEmpty($config.MovieFolder) -or [String]::IsNullOrEmpty($config.MovieVolumeLabel))) {
						$global:moviePath = [PathInfo]::new($config.MovieFolder, $config.MovieVolumeLabel)
					} else {
						$global:moviePath = $null
					}
				} else {
					$global:moviePath = [PathInfo]::new($config.MovieFolder)
				}
				
				# Check if series drive label should be used for identifying
				if ($config.UseDriveLabelForSeries) {
					if (-not ([string]::IsNullOrEmpty($config.SeriesFolder) -or [string]::IsNullOrEmpty($config.SeriesVolumeLabel))) {
						$global:seriesPath = [PathInfo]::new($config.SeriesFolder, $config.SeriesVolumeLabel)
					} else {
						$global:seriesPath = $null
					}
				} else {
					$global:seriesPath = [PathInfo]::new($config.SeriesFolder)
				}
			}
			
			# Check if movie path is available
			if (-not ([string]::IsNullOrEmpty($config.MovieFolder))) {
				$global:movieDriveFound = $global:moviePath.GetIsAvailable()
			}
			
			# Check if series path is available
			if (-not([string]::IsNullOrEmpty($config.SeriesFolder))) {
				$global:seriesDriveFound = $global:seriesPath.GetIsAvailable()
			}
			
			# Enable or disable analyze button
			$analyzeButton.Enabled = $global:movieDriveFound -Or $global:seriesDriveFound
			
			# Check if ffprobe path is set
			if ([string]::IsNullOrEmpty($config.FFprobePath)) {
				$errorMsg = Get-Translation -language $config.Language -key "Error.Config.FFProbePathNotSet"
				[System.Windows.Forms.MessageBox]::Show($errorMsg)
				$configValid = $false
			} else {
				# Check if ffprobe.exe exists
				If (-Not (Test-Path -LiteralPath $config.FFprobePath)) {
					$errorMsg = Get-Translation -language $config.Language -key "Error.Config.FFProbeNotFound"
					[System.Windows.Forms.MessageBox]::Show($errorMsg)
					$configValid = $false
				}
			}
			
			# Check if API key is set and valid
			if (-not([string]::IsNullOrEmpty($config.ApiKey))) {
				if (-not(Check-APIKey)) {
					$errorMsg = Get-Translation -language $config.Language -key "Error.Config.APIKeyNotValid"
					[System.Windows.Forms.MessageBox]::Show($errorMsg)
					#$configValid = $false
				}
			}
		}
		
		# Display configuration dialog on error
		if (-not($configValid)) {
			$ret = Show-ConfigDialog
			if ($ret -ne [System.Windows.Forms.DialogResult]::OK) {
				$msg = Get-Translation -language $config.Language -key "Error.Config.Canceld"
				[System.Windows.Forms.MessageBox]::Show($msg)
				#Exit
			}
		}
	} while (-not($configValid))
}

# Function for setting and updating text in configuration dialog
function Set-ConfigurationText {
	param (
		[string]$language
	)
	
	# Form name
	$configForm.Text = Get-Translation -language $language -key "Config.Form.Name"
	# Language label
	$labelLanguage.Text = Get-Translation -language $language -key "Config.Language.Label"
	
	# Movie button, label and checkbox
	$buttonMovieFolder.Text = Get-Translation -language $language -key "Config.Movie.Button"
	$labelMovieVolumeLabel.Text = Get-Translation -language $language -key "Config.Movie.VolumeLabel"
	$checkBoxMovieUseLabel.Text = Get-Translation -language $language -key "Config.Movie.UseLabel"
	
	# Series button, label and checkbox
	$buttonSeriesFolder.Text = Get-Translation -language $language -key "Config.Series.Button"
	$labelSeriesVolumeLabel.Text = Get-Translation -language $language -key "Config.Series.VolumeLabel"
	$checkBoxSeriesUseLabel.Text = Get-Translation -language $language -key "Config.Series.UseLabel"
	
	# Maximum number of files to scan for each movies and series
	$labelMaxScans.Text = Get-Translation -language $language -key "Config.MaxScans.Label"
	
	# API key label
	$labelApiKey.Text = Get-Translation -language $language -key "Config.TMDBAPIKey"
	
	# Adult content
	$checkBoxGetAdultContent.Text = Get-Translation -language $language -key "Config.GetAdultContent"
	
	# OK button
	$buttonOK.Text = Get-Translation -language $language -key "Button.OK"
}

# Function for creating and displaying configuration dialog
function Show-ConfigDialog {
	# Stop timer
	$global:timer.Stop()
	
	# Set global old folder data
	$global:movieFolderOld = ""
	$global:seriesFolderOld = ""
	
	# Load question mark icon
	$icon = [System.Drawing.SystemIcons]::Question
	
	# Create configuration form
	$configForm = New-Object System.Windows.Forms.Form
	$configForm.Size = New-Object System.Drawing.Size(400,470)
	$configForm.FormBorderStyle = "FixedDialog"
	
	# Label for language dropdown 
	$labelLanguage = New-Object System.Windows.Forms.Label
	$labelLanguage.Location = New-Object System.Drawing.Point(10, 10)
	$labelLanguage.Size = New-Object System.Drawing.Size(80, 20)
	$configForm.Controls.Add($labelLanguage)
	
	# Dropdown for language
	$comboBoxLanguage = New-Object System.Windows.Forms.ComboBox
	$comboBoxLanguage.Location = New-Object System.Drawing.Point(100,10)
	$comboBoxLanguage.Size = New-Object System.Drawing.Size(140, 20)
	# Create drop down box with available languages
	foreach ($lang in $sortedLanguages) {
		$index = $comboBoxLanguage.Items.Add($lang.Value)
	}
	# Get current used language based on language ISO code
	$cfgLangValue = $languages.PSObject.Properties[$config.Language].Value
	if($cfgLangValue -ne "") {
		# Language detected successfully, get index in language array
		$selectedLanguageIndex = [Array]::IndexOf($sortedLanguageValues, $cfgLangValue)
		if($selectedLanguageIndex -ne -1) {
			# Language found, so select the current entry
			$comboBoxLanguage.SelectedIndex = $selectedLanguageIndex
		}
	}
	# Add event handler for changes in dropdown box
	$comboBoxLanguage.Add_SelectedIndexChanged({
		# Get selected language
		$selectedLanguage = $sortedLanguages[$comboBoxLanguage.SelectedIndex].Name
		# Call the function to update all texts in configuration dialog
		Set-ConfigurationText -language $selectedLanguage
	})
	$configForm.Controls.Add($comboBoxLanguage)
	
	# Movies section
	# Folder dialog for movies
	$buttonMovieFolder = New-Object System.Windows.Forms.Button
	$buttonMovieFolder.Location = New-Object System.Drawing.Point(10,40)
	$buttonMovieFolder.Size = New-Object System.Drawing.Size(80, 20)
	$configForm.Controls.Add($buttonMovieFolder)
	# Add event handle for click on movies folder button
	$buttonMovieFolder.Add_Click({
		$folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
		If ($folderBrowser.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
			# New folder selected
			$moviePathData = [PathInfo]::new($folderBrowser.SelectedPath)
			if($moviePathData.GetIsAvailable()) {
				# Recently selected folder should be available
				# Get data and act acccordingly
				$textBoxMovieFolder.Text = $folderBrowser.SelectedPath
				$textBoxMovieVolumeLabel.Text = $moviePathData.GetDriveLabel()
				$checkBoxMovieUseLabel.Enabled = !$moviePathData.GetIsNetworkPath()
				If($moviePathData.GetIsNetworkPath()) {
					$checkBoxMovieUseLabel.Checked = $False
				}
			} else {
				# Display error messsage
				$msg = Get-Translation -language $language -key "Config.PathNotAvailable"
				$msg = $msg -replace "?", $folderBrowser.SelectedPath
				[System.Windows.Forms.MessageBox]::Show($msg)
			}
		} # no else part, folder dialog was just canceled
	})
	
	# Textbox for movies folder path
	$textBoxMovieFolder = New-Object System.Windows.Forms.TextBox
	$textBoxMovieFolder.Text = $config.MovieFolder
	$textBoxMovieFolder.Location = New-Object System.Drawing.Point(100,40)
	$textBoxMovieFolder.Size = New-Object System.Drawing.Size(240, 20)
	$configForm.Controls.Add($textBoxMovieFolder)
	# Add event handler for entering text box
	$textBoxMovieFolder.Add_Enter({
		$global:movieFolderOld = $textBoxMovieFolder.Text
	})
	# Add event handler for leaving text box
	$textBoxMovieFolder.Add_Leave({
		if($textBoxMovieFolder.Text -ne $global:movieFolderOld) {
			$moviePathData = [PathInfo]::new($textBoxMovieFolder.Text)
			if($moviePathData.GetIsAvailable()) {
				$textBoxMovieVolumeLabel.Text = $moviePathData.GetDriveLabel()
				$checkBoxMovieUseLabel.Enabled = !$moviePathData.GetIsNetworkPath()
				If($moviePathData.GetIsNetworkPath()) {
					$checkBoxMovieUseLabel.Checked = $False
				}
			} else {
				$textBoxMovieVolumeLabel.Text = $global:movieFolderOld
			}
			$global:movieFolderOld = ""
		}
	})
	
	# Label for movie volume label 
	$labelMovieVolumeLabel = New-Object System.Windows.Forms.Label
	$labelMovieVolumeLabel.Location = New-Object System.Drawing.Point(10, 70)
	$labelMovieVolumeLabel.Size = New-Object System.Drawing.Size(80, 20)
	$configForm.Controls.Add($labelMovieVolumeLabel)
	
	# Textbox for movies volume label
	$textBoxMovieVolumeLabel = New-Object System.Windows.Forms.TextBox
	$textBoxMovieVolumeLabel.Text = $config.MovieVolumeLabel
	$textBoxMovieVolumeLabel.Location = New-Object System.Drawing.Point(100,70)
	$textBoxMovieVolumeLabel.Size = New-Object System.Drawing.Size(240, 20)
	$textBoxMovieVolumeLabel.Enabled = $false
	$configForm.Controls.Add($textBoxMovieVolumeLabel)
	
	# Checkbox for using drive label
	$checkBoxMovieUseLabel = New-Object System.Windows.Forms.CheckBox
	$checkBoxMovieUseLabel.Checked = $config.UseDriveLabelForMovies
	$checkBoxMovieUseLabel.Location = New-Object System.Drawing.Point(10,100)
	$checkBoxMovieUseLabel.Size = New-Object System.Drawing.Size(340, 20)
	$configForm.Controls.Add($checkBoxMovieUseLabel)
	
	
	# Series section
	# Folder dialog for movies
	$buttonSeriesFolder = New-Object System.Windows.Forms.Button
	$buttonSeriesFolder.Location = New-Object System.Drawing.Point(10,140)
	$buttonSeriesFolder.Size = New-Object System.Drawing.Size(80, 20)
	$configForm.Controls.Add($buttonSeriesFolder)
	# Add event handle for click on series folder button
	$buttonSeriesFolder.Add_Click({
		$folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
		if ($folderBrowser.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
			# New folder selected
			$seriesPathData = [PathInfo]::new($folderBrowser.SelectedPath)
			if($seriesPathData.GetIsAvailable()) {
				# Recently selected folder should be available
				# Get data and act acccordingly
				$textBoxSeriesFolder.Text = $folderBrowser.SelectedPath
				$textBoxSeriesVolumeLabel.Text = $seriesPathData.GetDriveLabel()
				$checkBoxSeriesUseLabel.Enabled = !$seriesPathData.GetIsNetworkPath()
				If($seriesPathData.GetIsNetworkPath()) {
					$checkBoxSeriesUseLabel.Checked = $False
				}
			} else {
				# Display error messsage
				$msg = Get-Translation -language $language -key "Config.PathNotAvailable"
				$msg = $msg -replace "?", $folderBrowser.SelectedPath
				[System.Windows.Forms.MessageBox]::Show($msg)
			}
		} # no else part, folder dialog was just canceled
	})
	
	# Textbox for series folder path
	$textBoxSeriesFolder = New-Object System.Windows.Forms.TextBox
	$textBoxSeriesFolder.Text = $config.SeriesFolder
	$textBoxSeriesFolder.Location = New-Object System.Drawing.Point(100,140)
	$textBoxSeriesFolder.Size = New-Object System.Drawing.Size(240, 20)
	$configForm.Controls.Add($textBoxSeriesFolder)
	# Add event handler for entering text box
	$textBoxSeriesFolder.Add_Enter({
		$global:seriesFolderOld = $textBoxSeriesFolder.Text
	})
	# Add event handler for leaving text box
	$textBoxSeriesFolder.Add_Leave({
		if($textBoxSeriesFolder.Text -ne $global:seriesFolderOld) {
			$seriesPathData = [PathInfo]::new($textBoxSeriesFolder.Text)
			if($seriesPathData.GetIsAvailable()) {
				$textBoxSeriesVolumeLabel.Text = $seriesPathData.GetDriveLabel()
				$checkBoxSeriesUseLabel.Enabled = !$seriesPathData.GetIsNetworkPath()
				If($seriesPathData.GetIsNetworkPath()) {
					$checkBoxSeriesUseLabel.Checked = $False
				}
			} else {
				$textBoxSeriesFolder.Text = $global:seriesFolderOld
			}
			$global:seriesFolderOld = ""
		}
	})
	
	# Label for series volume label
	$labelSeriesVolumeLabel = New-Object System.Windows.Forms.Label
	$labelSeriesVolumeLabel.Location = New-Object System.Drawing.Point(10,170)
	$labelSeriesVolumeLabel.Size = New-Object System.Drawing.Size(80, 20)
	$configForm.Controls.Add($labelSeriesVolumeLabel)
	
	# Textbox for series volume label
	$textBoxSeriesVolumeLabel = New-Object System.Windows.Forms.TextBox
	$textBoxSeriesVolumeLabel.Text = $config.SeriesVolumeLabel
	$textBoxSeriesVolumeLabel.Location = New-Object System.Drawing.Point(100,170)
	$textBoxSeriesVolumeLabel.Size = New-Object System.Drawing.Size(240, 20)
	$textBoxSeriesVolumeLabel.Enabled = $false
	$configForm.Controls.Add($textBoxSeriesVolumeLabel)
	
	# Checkbox for using drive label
	$checkBoxSeriesUseLabel = New-Object System.Windows.Forms.CheckBox
	$checkBoxSeriesUseLabel.Checked = $config.UseDriveLabelForSeries
	$checkBoxSeriesUseLabel.Location = New-Object System.Drawing.Point(10,200)
	$checkBoxSeriesUseLabel.Size = New-Object System.Drawing.Size(340, 20)
	$configForm.Controls.Add($checkBoxSeriesUseLabel)
	
	# Label for maxmimum scans
	$labelMaxScans = New-Object System.Windows.Forms.Label
	$labelMaxScans.AutoSize = $true
	$labelMaxScans.Location = New-Object System.Drawing.Point(10, 230)
	$configForm.Controls.Add($labelMaxScans)
	
	# Textbox for maximum scans
	$textBoxMaxScans = New-Object System.Windows.Forms.TextBox
	$textBoxMaxScans.Location = New-Object System.Drawing.Point(10, 250)
	$textBoxMaxScans.Text = $config.MaxScans
	$configForm.Controls.Add($textBoxMaxScans)
	
	
	# File dialog button for ffprobe.exe
	$buttonFFprobePath = New-Object System.Windows.Forms.Button
	$buttonFFprobePath.Text = "ffprobe" # ToDo: translation of exe file into other language usefull?
	$buttonFFprobePath.Location = New-Object System.Drawing.Point(10,300)
	$buttonFFprobePath.Size = New-Object System.Drawing.Size(80, 20)
	$configForm.Controls.Add($buttonFFprobePath)
	$buttonFFprobePath.Add_Click({
		$fileDialog = New-Object System.Windows.Forms.OpenFileDialog
		if($textBoxFFprobePath.Text -ne "") {
			$fileDialog.InitialDirectory = $textBoxFFprobePath.Text
			$fileDialog.RestoreDirectory = $true;
		}
		$fileDialog.Filter = "ffprobe.exe|ffprobe.exe"
		if ($fileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
			$textBoxFFprobePath.Text = $fileDialog.FileName
		}
	})
	
	# Textbox for path to "ffprobe.exe"
	$textBoxFFprobePath = New-Object System.Windows.Forms.TextBox
	$textBoxFFprobePath.Text = $config.FFprobePath
	$textBoxFFprobePath.Location = New-Object System.Drawing.Point(100,300)
	$textBoxFFprobePath.Size = New-Object System.Drawing.Size(240, 20)
	$configForm.Controls.Add($textBoxFFprobePath)
	
	# Button for downloading ffmpeg suite
	$buttonFFProbeDownload = New-Object System.Windows.Forms.Button
	$buttonFFProbeDownload.Image = $icon.ToBitmap()
	$buttonFFProbeDownload.Location = New-Object System.Drawing.Point(350,300)
	$buttonFFProbeDownload.Size = New-Object System.Drawing.Size(20, 20)
	$configForm.Controls.Add($buttonFFProbeDownload)
	$buttonFFProbeDownload.Add_Click({
		Start-Process $URLFFMPEG
	})
	
	# Label for themoviedb.org API key
	$labelApiKey = New-Object System.Windows.Forms.Label
	$labelApiKey.Location = New-Object System.Drawing.Point(10,330)
	$labelApiKey.Size = New-Object System.Drawing.Size(80, 20)
	$configForm.Controls.Add($labelApiKey)
	
	# Textbox for themoviedb.org API key
	$textBoxApiKey = New-Object System.Windows.Forms.TextBox
	$textBoxApiKey.Text = $config.ApiKey
	$textBoxApiKey.Location = New-Object System.Drawing.Point(100,330)
	$textBoxApiKey.Size = New-Object System.Drawing.Size(240, 20)
	$configForm.Controls.Add($textBoxApiKey)
	
	# Button for signing up to themoviedb.org
	$buttonTMDBSignUp = New-Object System.Windows.Forms.Button
	$buttonTMDBSignUp.Image = $icon.ToBitmap()
	$buttonTMDBSignUp.Location = New-Object System.Drawing.Point(350,330)
	$buttonTMDBSignUp.Size = New-Object System.Drawing.Size(20, 20)
	$configForm.Controls.Add($buttonTMDBSignUp)
	$buttonTMDBSignUp.Add_Click({
		Start-Process $URLTMDBSIGNUP
	})
	
	# Checkbox for fetching adult content
	$checkBoxGetAdultContent = New-Object System.Windows.Forms.CheckBox
	$checkBoxGetAdultContent.Checked = $config.GetAdultContent
	$checkBoxGetAdultContent.Location = New-Object System.Drawing.Point(10,360)
	$checkBoxGetAdultContent.Size = New-Object System.Drawing.Size(340, 20)
	$configForm.Controls.Add($checkBoxGetAdultContent)
	
	
	# OK button, saving the configruation, closes configuration dialog
	$buttonOK = New-Object System.Windows.Forms.Button
	$buttonOK.Location = New-Object System.Drawing.Point(150,400)
	$buttonOK.Size = New-Object System.Drawing.Size(80, 20)
	$configForm.Controls.Add($buttonOK)
	# Event handler for OK button
	$buttonOK.Add_Click({
		# Check number of maximum files to scan
		$maxScans = $textBoxMaxScans.Text
		if ($maxScans -match '^\d+$' -and [int]$maxScans -ge 0 -and [int]$maxScans -le 50) {
			$config.Language = $sortedLanguages[$comboBoxLanguage.SelectedIndex].Name
			$config.MovieFolder = $textBoxMovieFolder.Text
			$config.MovieVolumeLabel = $textBoxMovieVolumeLabel.Text
			$config.UseDriveLabelForMovies = $checkBoxMovieUseLabel.Checked
			$config.SeriesFolder = $textBoxSeriesFolder.Text
			$config.SeriesVolumeLabel = $textBoxSeriesVolumeLabel.Text
			$config.UseDriveLabelForSeries = $checkBoxSeriesUseLabel.Checked
			$config.FFprobePath = $textBoxFFprobePath.Text
			$config.ApiKey = $textBoxApiKey.Text
			$config.GetAdultContent = $checkBoxGetAdultContent.Checked
			$config.MaxScans = $maxScans
			
			$global:moviePath = [PathInfo]::new($config.MovieFolder)
			$global:seriesPath = [PathInfo]::new($config.SeriesFolder)
			
			Save-Config -config $config
			
			$configForm.DialogResult = [System.Windows.Forms.DialogResult]::OK
			
			# Start timer
			$global:timer.Start()
			
			$configForm.Close()
		} else {
			$msg = Get-Translation -language $language -key "Config.MaxScans.Error"
			[System.Windows.Forms.MessageBox]::Show($msg)
		}
	})
	
	Set-ConfigurationText -language $config.Language
	
	# Show configuration dialog
	$configForm.ShowDialog()
}

# Function for creating and displaying themoviedb.org ID dialog.
# In this dialog the ID for the selected entry can be modified.
# This results in checking if a new ID has been entered. If it is
# changed check if the prior ID was zero. If there was an ID greater
# than 0 remove data for this video ID and TMDB ID on any table.
function Show-TMDBDialog {
	param (
		[PSObject]$row
	)
	
	# Stop timer
	$global:timer.Stop()
	
	# Get ID from VideoList
	$videoIdCurrent = $selectedRow.Cells[$GRIDVIEWCOLUMNID].Value
	# Get TMDB ID
	$tmdbIdCurrent = $selectedRow.Cells[$GRIDVIEWCOLUMNTMDBID].Value
	# Get video type
	$videoTypeCurrent = $selectedRow.Cells[$GRIDVIEWCOLUMNVIDEOTYPE].Value
	
	# Create the form
	$tmdbForm = New-Object System.Windows.Forms.Form
	$tmdbForm.Text = Get-Translation -language $config.language -key "Change.TMDBID.Form.Name"
	$tmdbForm.Size = New-Object System.Drawing.Size(300,150)
	$tmdbForm.StartPosition = "CenterParent"
	
	# Create text box
	$textBox = New-Object System.Windows.Forms.TextBox
	$textBox.Size = New-Object System.Drawing.Size(260,20)
	$textBox.Location = New-Object System.Drawing.Point(10,20)
	$textBox.Text = $tmdbIdCurrent
	# Add event handler for keydown event
	$textBox.Add_KeyDown({
		param (
			[System.Object]$sender,
			[System.Windows.Forms.KeyEventArgs]$e
		)
		
		# Search for data if the enter key has been pressed
		if ($e.KeyCode -eq [System.Windows.Forms.Keys]::Return) {
			$tmdbForm.DialogResult = [System.Windows.Forms.DialogResult]::OK
			$tmdbForm.Close()
		}
	})
	$tmdbForm.Controls.Add($textBox)
	
	# Create OK button
	$okButton = New-Object System.Windows.Forms.Button
	$okButton.Text = Get-Translation -language $config.language -key "Button.OK"
	$okButton.Location = New-Object System.Drawing.Point(50,60)
	$okButton.Add_Click({
		$tmdbForm.DialogResult = [System.Windows.Forms.DialogResult]::OK
		$tmdbForm.Close()
	})
	$tmdbForm.Controls.Add($okButton)
	
	# Create cancel button
	$cancelButton = New-Object System.Windows.Forms.Button
	$cancelButton.Text = Get-Translation -language $config.language -key "Button.Cancel"
	$cancelButton.Location = New-Object System.Drawing.Point(150,60)
	$cancelButton.Add_Click({
		$tmdbForm.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
		$tmdbForm.Close()
	})
	$tmdbForm.Controls.Add($cancelButton)
	
	# Show dialog
	$result = $tmdbForm.ShowDialog()
	
	# Check and process result
	if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
		$userInput = $textBox.Text
		if ($userInput -ne "") {
			# First remove all old data for the current ID from related tables:
			Delete-Genres -videoId $videoIdCurrent
			Delete-Actors -videoId $videoIdCurrent
			Delete-BelongsTo -videoId $videoIdCurrent
			Delete-VideoDetails -tmdbId $tmdbIdCurrent -videoType $videoTypeCurrent
			
			# Get the ID from the input field
			$tmdbId = [System.Int64]$userInput
			
			# Update DataGridView
			$selectedRow.Cells["TMDBId"].Value = $tmdbId
			
			# Check new TMDB Id
			if ($tmdbId -eq 0) {
				# The new TMDB Id is zero so don't receive
				Update-TMDBVideoInfo -videoId $videoIdCurrent -tmdbId [System.Int64]0 -vote 0.0
			} else {
				if ($videoTypeCurrent -eq "M") {
					# Update as movie
					$movieInfo = Get-MovieInfoFromTMDBbyID -tmdbId $tmdbId
					
					# If the movie was found query movie details
					if (-not ([string]::IsNullOrEmpty($movieInfo))) {
						if ($movieInfo.id -gt 0 ) {
							# Update new TMDB Id and avarage vote
							Update-TMDBVideoInfo -videoId $videoIdCurrent -tmdbId $tmdbId -vote $movieInfo.vote_average
							
							# Search for movie details in TMDB
							$movieDetails = Get-MovieDetailsFromTMDB -movieId $movieInfo.id
							
							if ($movieDetails -ne $null) {
								# Insert data into database
								$movieDetailsObject = [PSCustomObject]@{
									TMDBId		= [System.Int64]$movieDetails.id
									VideoType   = "M"
									Title		= $movieDetails.title
									Overview	= $movieDetails.overview
									ReleaseDate = $movieDetails.release_date
								}
								
								# Insert or update movie details
								Upsert-VideoDetails -videoId $videoIdCurrent -videoDetails $movieDetailsObject
								
								# Insert or update genres
								Upsert-Genres -id $videoIdCurrent -genres $movieDetails.genres
								
								# Insert or update actors
								Upsert-Actors -id $videoIdCurrent -actors $movieDetails.credits.cast
								
								# Check if the movie belongs to a collection
								$collection = $movieDetails.belongs_to_collection
								if (-not ([string]::IsNullOrEmpty($collection))) {
									# Query details about movie collection from TMDB
									$collectionDetails = Get-CollectionDetailsFromTMDB -collection $movieDetails.belongs_to_collection.id
									if (-not ( $collectionDetails -eq $null )) {
										# Insert or update collection details
										Upsert-VideoBelongsTo -videoListId $videoIdCurrent -videoType "M" -tmdbBelongsToId $collectionDetails.id -name $collectionDetails.name -overview $collectionDetails.overview
									}
								}
							} else {
								# Movie hasn't been found in TMDB so inset or update basic information only
								Update-TMDBVideoInfo -videoId $videoIdCurrent -tmdbId [System.Int64]0 -vote 0.0
							}
						} else {
							# Movie hasn't been found in TMDB so inset or update basic information only
							Update-TMDBVideoInfo -videoId $videoIdCurrent -tmdbId [System.Int64]0 -vote 0.0
						}
					} else {
						# Movie hasn't been found in TMDB so inset or update basic information only
						Update-TMDBVideoInfo -videoId $videoIdCurrent -tmdbId [System.Int64]0 -vote 0.0
					}
				} else {
					# Update as series
					# ToDo, see To do at beginning
				}
			}
		}
		
		# Trigger selection changed event on DataGridView to display new details
		$dataGridView.GetType().GetMethod("OnSelectionChanged", [System.Reflection.BindingFlags]::Instance -bor [System.Reflection.BindingFlags]::NonPublic).Invoke($dataGridView, @([System.EventArgs]::Empty))
	}
}


# Find double files
function Show-Doubles {
	cursorWait
	if (-not($global:doubleFilesView)) {
		# Keep the button pressed
		$duplicatesButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Popup
		# Currently no double files are viewed
		$global:doubleFilesView = $true
		# Reload all data, ignoring search field
		Load-Data
		
		# Find and select double entries:
		# - check for title
		# - if the type is movie, check for filename
		# - check the combination of video type (series or movie) and TMDB ID
		$duplicatesColumnTitle = @{}
		$duplicatesColumnFilename = @{}
		$duplicatesColumnTMDBID = @{}
		foreach ($row in $dataGridView.Rows) {
			if ($row.Index -lt $dataGridView.RowCount) {
				# Get values from columns
				$columnTitle = $row.Cells[$GRIDVIEWCOLUMNTITLE].Value
				$columnFilename = $row.Cells[$GRIDVIEWCOLUMNFILENAME].Value
				$columnTMDBID = $row.Cells[$GRIDVIEWCOLUMNTMDBID].Value
				$columnVideoType = $row.Cells[$GRIDVIEWCOLUMNVIDEOTYPE].Value
				
				$columnTitle = $columnTitle + "-" + $columnVideoType + "-" + $columnTMDBID
				# Check the title
				if (-not $duplicatesColumnTitle.ContainsKey($columnTitle)) {
					$duplicatesColumnTitle[$columnTitle] = 0
				}
				$duplicatesColumnTitle[$columnTitle]++
				
				# Check if the video type is movie, since several series episodes use
				# identical names.
				if ($columnVideoType -eq "M" ) {
					# Check filename
					if (-not $duplicatesColumnFilename.ContainsKey($columnFilename)) {
						$duplicatesColumnFilename[$columnFilename] = 0
					}
				} else {
					$duplicatesColumnFilename[$columnFilename] = 0
				}
				$duplicatesColumnFilename[$columnFilename]++
				
				# Check the combination of video type and TMDB ID since the IDs are not unique
				# for series and movies. And also ignore entries if the TMDB ID is 0.
				$combination = "$columnVideoType-$columnTMDBID#"
				if ($columnTMDBID -ne 0) {
					if (-not $duplicatesColumnTMDBID.ContainsKey($combination)) {
						$duplicatesColumnTMDBID[$combination] = 0
					}
				}
				$duplicatesColumnTMDBID[$combination]++
			}
		}
		
		# Remove all unique entries
		for ($i = $dataGridView.Rows.Count - 1; $i -ge 0; $i--) { # Ignore last line
			$row = $dataGridView.Rows[$i]
			$columnTitle = $row.Cells[$GRIDVIEWCOLUMNTITLE].Value
			$columnFilename = $row.Cells[$GRIDVIEWCOLUMNFILENAME].Value
			$columnTMDBID = $row.Cells[$GRIDVIEWCOLUMNTMDBID].Value
			$columnVideoType = $row.Cells[$GRIDVIEWCOLUMNVIDEOTYPE].Value
			$combination = "$columnVideoType-$columnTMDBID#"
			
			# Check that title and filename counters are 0 or TMDB ID counter is 0
			if ((($duplicatesColumnTitle[$columnTitle] -le 1) -and ($duplicatesColumnFilename[$columnFilename] -le 1)) -and (($duplicatesColumnTMDBID[$combination] -le 1)) -or ($columnTMDBID -eq 0)) {
				$dataGridView.Rows.RemoveAt($i)
			}
		}
		
		# Mark non-existing files
		Mark-Red
	} else {
		# Double files are viewed, so return to default view
		$global:doubleFilesView = $false
		Load-Data
		
		# Mark non-existing files
		Mark-Red
		$duplicatesButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Standard
	}
	
	cursorDefault
}

######################################################

######################################################
# Database functions
# Create tables
function Create-Database-Tables {
	try {
		$db.Exec('BEGIN TRANSACTION;')
		
		# "VideoList" table contains basic data like file name and path, extracted title and
		# data detected with ffprobe. Also additional data like the ID from themoviedb.org
		# and average votes.
		# The column "FileExists" will be checked. If a file is not longer found at the
		# given location it will not be deleted from the database. It can be used as an
		# indicator that te movie was watched and deleted, but it is not necessary to
		# record it again.
		$db.Exec(@"
CREATE TABLE IF NOT EXISTS VideoList (
	Id INTEGER PRIMARY KEY AUTOINCREMENT,
	Title TEXT,
	FileName TEXT,
	FilePath TEXT,
	FileSize INTEGER,
	FileSizeMB REAL,
	Resolution TEXT,
	VideoCodec TEXT,
	AudioTracks INTEGER,
	AudioChannels TEXT,
	AudioLayouts TEXT,
	AudioLanguages TEXT,
	Duration TEXT,
	FileExists INTEGER DEFAULT 1,
	TMDBId INTEGER,
	VoteAverage NUMERIC,
	VideoType CHARACTER,
	IsAdult INTEGER,
	UNIQUE(FilePath, FileName)
);
"@)
		# "VideoDetails" table contains retreived information from tmdb as their title,
		# date of release and also the overview text.
		# The table uses the video type and TMDB id as reference to TMDB and also to
		# VideoList table.
		$db.Exec(@"
CREATE TABLE IF NOT EXISTS VideoDetails (
	VideoId INTEGER,
	TMDBId INTEGER,
	VideoType TEXT,
	Title TEXT,
	Overview TEXT,
	ReleaseDate INTEGER,
	UNIQUE(TMDBId, VideoType),
	FOREIGN KEY(TMDBId, VideoType) REFERENCES VideoList(TMDBId, VideoType)
);
"@)
		# "BelongsTo" table contains an overview of a movie collection or the series where
		# a video belongs to.
		$db.Exec(@"
CREATE TABLE IF NOT EXISTS BelongsTo (
	BelongsToId INTEGER PRIMARY KEY AUTOINCREMENT,
	TMDBId INTEGER,
	VideoType TEXT,
	BelongsToName TEXT,
	OverView TEXT
);
"@)
		# "VideoBelongsTo" is the n:m relation for the video and where it belongs to.
		$db.Exec(@"
CREATE TABLE IF NOT EXISTS VideoBelongsTo (
	VideoId INTEGER,
	BelongsToId INTEGER,
	FOREIGN KEY(VideoId) REFERENCES VideoList(Id),
	FOREIGN KEY(BelongsToId) REFERENCES BelongsTo(BelongsToId),
	PRIMARY KEY(VideoId, BelongsToId)
);
"@)
		# "Genres" list with ID and genre from TMDB.
		$db.Exec(@"
CREATE TABLE IF NOT EXISTS Genres (
	TMDBId INTEGER PRIMARY KEY,
	GenreName TEXT
);
"@)
		# "Actors" list with ID and name from TMDB.
		$db.Exec(@"
CREATE TABLE IF NOT EXISTS Actors (
	TMDBId INTEGER PRIMARY KEY,
	ActorName TEXT
);
"@)
		# "VideoGenres" is the n:m relation between videos and genres.
		$db.Exec(@"
CREATE TABLE IF NOT EXISTS VideoGenres (
	VideoId INTEGER,
	GenreTMDBId INTEGER,
	FOREIGN KEY(VideoId) REFERENCES VideoList(Id),
	FOREIGN KEY(GenreTMDBId) REFERENCES Genres(TMDBId),
	PRIMARY KEY(VideoId, GenreTMDBId)
);
"@)
		# "VideoActors" is the n:m relation between video and actors.
		$db.Exec(@"
CREATE TABLE IF NOT EXISTS VideoActors (
	VideoId INTEGER,
	ActorTMDBId INTEGER,
	FOREIGN KEY(VideoId) REFERENCES VideoList(Id),
	FOREIGN KEY(ActorTMDBId) REFERENCES Actors(TMDBId),
	PRIMARY KEY(VideoId, ActorTMDBId)
);
"@)
		
		# Create indexes
		$db.Exec("CREATE INDEX idx_VideoList_Id ON VideoList(Id);")
		$db.Exec("CREATE INDEX idx_VideoList_TMDBId ON VideoList(TMDBId);")
		
		$db.Exec("CREATE INDEX idx_VideoDetails_TMDBId ON VideoDetails(TMDBId);")
		$db.Exec("CREATE INDEX idx_VideoDetails_VideoType ON VideoDetails(VideoType);")
		
		$db.Exec("CREATE INDEX idx_VideoBelongsTo_BelongsToId ON VideoBelongsTo(BelongsToId);")
		$db.Exec("CREATE INDEX idx_VideoBelongsTo_VideoID ON VideoBelongsTo(VideoID);")
		
		$db.Exec("CREATE INDEX idx_BelongsTo_BelongsToId ON BelongsTo(BelongsToId);")
		$db.Exec("CREATE INDEX idx_BelongsTo_TMDBId ON BelongsTo(TMDBId);")
		
		$db.Exec("CREATE INDEX idx_Actors_TMDBId ON Actors(TMDBId);")
		$db.Exec("CREATE INDEX idx_VideoActors_ActorTMDBId ON VideoActors(ActorTMDBId);")
		$db.Exec("CREATE INDEX idx_VideoActors_VideoId ON VideoActors(VideoId);")
		
		$db.Exec("CREATE INDEX idx_Genres_TMDBId ON Genres(TMDBId);")
		$db.Exec("CREATE INDEX idx_VideoGenres_GenreTMDBId ON VideoGenres(GenreTMDBId);")
		$db.Exec("CREATE INDEX idx_VideoGenres_VideoId ON VideoGenres(VideoId);")
		
		$db.Exec('COMMIT;')
	}
	catch {
		$db.Exec('ROLLBACK;')
		Write-Warning "function Create-Database-Tables: ERROR: Unknown error occured: $_"
		Write-Warning $Error[0].Exception.GetType().FullName
		throw $Error
	}
}

# Load video data from database
# Submit a search term for limiting results
function Load-Data {
	param (
		[String]$searchTerm = ""
	)
	
	try {
		if (([String]::IsNullOrEmpty($searchTerm)) -and ($global:lastSearchFiltered -eq $false)) {
			$stmt = $db.Prepare(@"
SELECT DISTINCT
	VideoList.Id,
	VideoList.Title,
	VideoList.FileName,
	VideoList.FilePath,
	BelongsTo.BelongsToName,
	VideoList.FileSize,
	VideoList.FileSizeMB,
	VideoList.Resolution,
	VideoList.VideoCodec,
	VideoList.AudioTracks,
	VideoList.AudioChannels,
	VideoList.AudioLayouts,
	VideoList.AudioLanguages,
	VideoList.Duration,
	VideoList.FileExists,
	VideoList.TMDBId,
	VideoList.VoteAverage,
	VideoList.VideoType,
	VideoList.IsAdult
FROM VideoList 
LEFT JOIN VideoBelongsTo ON VideoList.Id = VideoBelongsTo.VideoID
LEFT JOIN BelongsTo ON BelongsTo.BelongsToId = VideoBelongsTo.BelongsToId;
"@)
			
			$global:lastSearchFiltered = $false
		} else {
			if (([String]::IsNullOrEmpty($searchTerm)) -and ($global:lastSearchFiltered -eq $true)) {
				$searchTerm =  $searchBox.Text
			}
			
			$stmt = $db.Prepare(@"
SELECT DISTINCT
	VideoList.Id,
	VideoList.Title,
	VideoList.FileName,
	VideoList.FilePath,
	BelongsTo.BelongsToName,
	VideoList.FileSize,
	VideoList.FileSizeMB,
	VideoList.Resolution,
	VideoList.VideoCodec,
	VideoList.AudioTracks,
	VideoList.AudioChannels,
	VideoList.AudioLayouts,
	VideoList.AudioLanguages,
	VideoList.Duration,
	VideoList.FileExists,
	VideoList.TMDBId,
	VideoList.VoteAverage,
	VideoList.VideoType,
	VideoList.IsAdult
FROM VideoList 
LEFT JOIN VideoBelongsTo ON VideoList.Id = VideoBelongsTo.VideoID
LEFT JOIN BelongsTo ON BelongsTo.BelongsToId = VideoBelongsTo.BelongsToId
LEFT JOIN VideoActors ON VideoList.Id = VideoActors.VideoId
LEFT JOIN Actors ON VideoActors.ActorTMDBId = Actors.TMDBId
WHERE VideoList.FileName LIKE ?
   OR VideoList.FilePath LIKE ?
   OR BelongsTo.BelongsToName LIKE ?
   OR Actors.ActorName LIKE ?;
"@)
			$res = $db.BindText($stmt, 1, "%$searchTerm%")
			$res = $db.BindText($stmt, 2, "%$searchTerm%")
			$res = $db.BindText($stmt, 3, "%$searchTerm%")
			$res = $db.BindText($stmt, 4, "%$searchTerm%")
		}
		
		# Create a data table
		$dataTable = New-Object System.Data.DataTable
		
		# Get column count
		$columnCount = $db.GetColumnCount($stmt)
		
		# Define lists for column type
		$listInteger = @($GRIDVIEWCOLUMNID, $GRIDVIEWCOLUMNFILESIZE, $GRIDVIEWCOLUMNAUDIOTRACKS, $GRIDVIEWCOLUMNFILEEXISTS, $GRIDVIEWCOLUMNTMDBID)
		$listDouble = @($GRIDVIEWCOLUMNFILESIZEMB, $GRIDVIEWCOLUMNVOTE)
		
		# Get column names from SELECT result
		for ($i = 0; $i -lt $columnCount; $i++) {
			$colName = $db.GetColumnName($stmt, $i)
			$col = New-Object System.Data.DataColumn($colName)
			$x = $null
			if ($listInteger -contains $i) {
				$x = $dataTable.Columns.Add($col, [System.Int64])
			} elseif ($listDouble -contains $i) {
				$x = $dataTable.Columns.Add($col, [System.Double])
			} else {
				$x = $dataTable.Columns.Add($col, [System.String])
			}
		}
		
		# Get results from database query
		while ($true) {
			$result = $db.StepAndGetRow($stmt, $columnCount)
				if ($result[0] -ne $([SQLiteHelper]::SQLITE_ROW)) { # SQLITE_ROW
				break
			}
			$row = $dataTable.NewRow()
			$row.ItemArray = $result[1]
			$dataTable.Rows.Add($row)
		}
		
		# Assign the dats table to the data grid view
		$dataGridView.DataSource = $dataTable
#		$dataGridView.Columns[ 0].HeaderText = Get-Translation -language $config.language -key "DB.Video.ID"
		$dataGridView.Columns[ 1].HeaderText = Get-Translation -language $config.language -key "DB.Video.Title"
		$dataGridView.Columns[ 2].HeaderText = Get-Translation -language $config.language -key "DB.Video.FileName"
		$dataGridView.Columns[ 3].HeaderText = Get-Translation -language $config.language -key "DB.Video.FilePath"
		$dataGridView.Columns[ 4].HeaderText = Get-Translation -language $config.language -key "DB.Video.BelongsTo"
		$dataGridView.Columns[ 5].HeaderText = Get-Translation -language $config.language -key "DB.Video.FileSize"
		$dataGridView.Columns[ 6].HeaderText = Get-Translation -language $config.language -key "DB.Video.FileSizeMB"
		$dataGridView.Columns[ 7].HeaderText = Get-Translation -language $config.language -key "DB.Video.Resolution"
		$dataGridView.Columns[ 8].HeaderText = Get-Translation -language $config.language -key "DB.Video.VideoCodec"
		$dataGridView.Columns[ 9].HeaderText = Get-Translation -language $config.language -key "DB.Video.AudioTracks"
		$dataGridView.Columns[10].HeaderText = Get-Translation -language $config.language -key "DB.Video.AudioChannels"
		$dataGridView.Columns[11].HeaderText = Get-Translation -language $config.language -key "DB.Video.AudioLayouts"
		$dataGridView.Columns[12].HeaderText = Get-Translation -language $config.language -key "DB.Video.AudioLanguages"
		$dataGridView.Columns[13].HeaderText = Get-Translation -language $config.language -key "DB.Video.Duration"
#		$dataGridView.Columns[14].HeaderText = Get-Translation -language $config.language -key "DB.Video.FileExists"
#		$dataGridView.Columns[15].HeaderText = Get-Translation -language $config.language -key "DB.Video.TMDBId"
		$dataGridView.Columns[16].HeaderText = Get-Translation -language $config.language -key "DB.Video.VoteAverage"
#		$dataGridView.Columns[17].HeaderText = Get-Translation -language $config.language -key "DB.Video.VideoType"
#		$dataGridView.Columns[18].HeaderText = Get-Translation -language $config.language -key "DB.Video.AdultContent"
#		$dataGridView.Columns[$GRIDVIEWCOLUMNADULTCONTENT].HeaderText = Get-Translation -language $config.language -key "DB.Video.IsAdult"
				
		# Add a hidden column as sort key for path, setting entries to path and filename
		$res = $dataTable.Columns.Add("_PathAndFileName")
		foreach ($row in $dataGridView.Rows) {
			$row.Cells["_PathAndFileName"].Value = $row.Cells[$GRIDVIEWCOLUMNFILEPATH].Value + "\" + $row.Cells[$GRIDVIEWCOLUMNFILENAME].Value
		}
		
		# Hide columns
		$dataGridView.Columns[$GRIDVIEWCOLUMNID].Visible = $false               # ID
		$dataGridView.Columns[$GRIDVIEWCOLUMNFILESIZE].Visible = $false         # File Size
		$dataGridView.Columns[$GRIDVIEWCOLUMNFILEEXISTS].Visible = $false       # File Exists
		$dataGridView.Columns[$GRIDVIEWCOLUMNTMDBID].Visible = $false           # TMDB-ID
		$dataGridView.Columns[$GRIDVIEWCOLUMNVIDEOTYPE].Visible = $false        # Video type
		$dataGridView.Columns[$GRIDVIEWCOLUMNADULTCONTENT].Visible = $false     # Adult content
		$dataGridView.Columns[$GRIDVIEWCOLUMNPATHANDFILENAME].Visible = $false  # Path and filename
		$db.Finalize($stmt)
		
		$global:lastSearchFiltered = $true
	}
	catch {
		Write-Warning "function Load-Data: ERROR: Unknown error occured: $_"		
		Write-Warning $Error[0].Exception.GetType().FullName
		throw $Error
	}
}

# Load row from data with given ID
function Load-DataRow {
	param (
		[System.Int32]$id
	)
	
	if( $id -ge 0 ) {
		# Get data table from DataGridView
		[System.Data.DataTable]$dataTable = $dataGridView.DataSource
		
		try {
			$stmt = $db.Prepare(@"
SELECT DISTINCT
	VideoList.Id,
	VideoList.Title,
	VideoList.FileName,
	VideoList.FilePath,
	BelongsTo.BelongsToName,
	VideoList.FileSize,
	VideoList.FileSizeMB,
	VideoList.Resolution,
	VideoList.VideoCodec,
	VideoList.AudioTracks,
	VideoList.AudioChannels,
	VideoList.AudioLayouts,
	VideoList.AudioLanguages,
	VideoList.Duration,
	VideoList.FileExists,
	VideoList.TMDBId,
	VideoList.VoteAverage,
	VideoList.VideoType,
	VideoList.IsAdult
FROM VideoList 
LEFT JOIN VideoBelongsTo ON VideoList.TMDBId = VideoBelongsTo.VideoID
LEFT JOIN BelongsTo ON BelongsTo.TMDBId = VideoBelongsTo.BelongsToId
WHERE VideoList.Id = ?;
"@)
			$res = $db.BindInt($stmt, 1, $id)
			
			# Get results from database query
			$columnCount = $db.GetColumnCount($stmt)
			while ($true) {
				$result = $db.StepAndGetRow($stmt, $columnCount)
					if ($result[0] -ne $([SQLiteHelper]::SQLITE_ROW)) { # SQLITE_ROW
					break
				}
				$row = $dataTable.NewRow()
				$row.ItemArray = $result[1]
				$row["_PathAndFileName"] = $row[3] + "\" + $row[2]
				$dataTable.Rows.Add($row)
			}
			$db.Finalize($stmt)
		}
		catch {
			Write-Warning "function Load-DataRow: ERROR: Unknown error occured: $_"
			Write-Warning $Error[0].Exception.GetType().FullName
			throw $Error
		}
	}
}

# Insert or update video information in the database
function Upsert-VideoInfo {
	param (
		[PSCustomObject]$videoInfo,
		[Double]$fileSize,
		[String]$title,
		[System.Int64]$tmdbId,
		[String]$vote,
		[String]$videoType,
		[String]$isAdult
	)

	# Set ID to zero
	$id = 0
	
	# Check adult content flag
	if ($isAdult -eq "true") {
		$isAdultVal = 1
	} else {
		$isAdultVal = 0
	}

	# Check and replace decimal point with country sepcific delimeter
	$voteNumber = 0.0
	# Check if variable is a string containing a number
	if ([double]::TryParse(($vote), [ref]$null)) {
		$numberAsText = $vote.ToString()
		$voteNumber = [double]$numberAsText
	}
	
	try {
		$stmt = $db.Prepare(@"
INSERT INTO VideoList (Title, FileName, FilePath, FileSize, FileSizeMB, Resolution,
VideoCodec, AudioTracks, AudioChannels, AudioLayouts, AudioLanguages,
Duration, FileExists, TMDBId, VoteAverage, VideoType, IsAdult)
VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
ON CONFLICT(FilePath, FileName) DO UPDATE SET
	Title = excluded.Title,
	FileName = excluded.FileName,
	FilePath = excluded.FilePath,
	FileSize = excluded.FileSize,
	FileSizeMB = excluded.FileSizeMB,
	Resolution = excluded.Resolution,
	VideoCodec = excluded.VideoCodec,
	AudioTracks = excluded.AudioTracks,
	AudioChannels = excluded.AudioChannels,
	AudioLayouts = excluded.AudioLayouts,
	AudioLanguages = excluded.AudioLanguages,
	Duration = excluded.Duration,
	FileExists = excluded.FileExists,
	TMDBId = excluded.TMDBId,
	VoteAverage = excluded.VoteAverage,
	VideoType = excluded.VideoType,
	IsAdult = excluded.IsAdult;
"@)

		$res = $db.BindText($stmt,  1, $title)
		$res = $db.BindText($stmt,  2, $videoInfo.FileName)
		$res = $db.BindText($stmt,  3, $videoInfo.FilePath)
		$res = $db.BindInt64( $stmt,  4, $fileSize)
		$res = $db.BindDouble($stmt, 5, [double][math]::Round($fileSize / 1MB, 2))
		$res = $db.BindText($stmt,  6, $videoInfo.Resolution)
		$res = $db.BindText($stmt,  7, $videoInfo.VideoCodec)
		$res = $db.BindInt($stmt,  8, [System.Int32]$videoInfo.AudioTracks)
		$res = $db.BindText($stmt,  9, $videoInfo.AudioChannels)
		$res = $db.BindText($stmt, 10, $videoInfo.AudioLayouts)
		$res = $db.BindText($stmt, 11, $videoInfo.AudioLanguages)
		$res = $db.BindText($stmt, 12, $videoInfo.Duration)
		$res = $db.BindInt($stmt, 13, $videoInfo.FileExists)
		$res = $db.BindInt64($stmt, 14, $tmdbId)
		$res = $db.BindDouble($stmt, 15, $voteNumber)
		$res = $db.BindText($stmt, 16, $videoType)
		$res = $db.BindInt($stmt, 17, $isAdultVal)
		$result = $db.Step($stmt)
		
		# Get ID of last inserted row
		$id = $db.GetLastInsertRowId()
		$db.Finalize($stmt)
	}
	catch {
		Write-Warning "function Upsert-VideoInfo: ERROR: Unknown error occured: $_"
		Write-Warning $Error[0].Exception.GetType().FullName
		throw $Error
	}
	
	return $id
}

# Update TMDB information for video in the database
function Update-TMDBVideoInfo {
	param (
		[System.Int32]$videoId,
		[System.Int64]$tmdbId,
		[Double]$vote
	)
	
	try {
		$stmt = $db.Prepare("UPDATE VideoList SET TMDBId = ?, VoteAverage = ? WHERE ID = ?;")
		$db.BindInt64($stmt, 1, $tmdbId)
		$db.BindDouble($stmt, 2, $vote)
		$db.BindInt($stmt, 3, $videoId)
		$result = $db.Step($stmt)
		$db.Finalize($stmt)
	}
	catch {
		Write-Warning "function Update-TMDBVideoInfo: ERROR: Unknown error occured: $_"
		Write-Warning "SQL:" $db.GetExpandedSql($stmt)
		$db.Finalize($stmt)
		
		Write-Warning $Error[0].Exception.GetType().FullName
		throw $Error
	}
}

# Update filename for video in the database
function Update-FilenameVideoInfo {
	param (
		[System.Int32]$videoId,
		[String]$filename
	)
	
	try {
		$stmt = $db.Prepare("UPDATE VideoList SET FileName = ?, FileExists = 1 WHERE ID = ?;")
		$db.BindText($stmt, 1, $filename)
		$db.BindInt($stmt, 2, $videoId)
		$result = $db.Step($stmt)
		$db.Finalize($stmt)
	}
	catch {
		Write-Warning "function Update-FilenameVideoInfo: ERROR: Unknown error occured: $_"
		Write-Warning $Error[0].Exception.GetType().FullName
		throw $Error
	}
}

# Scroll to the last element in the data grid view
function Scroll-To-Last-Element {
	if ($dataGridView -ne $null) {
		$lastRowIndex = $dataGridView.Rows.Count - 1
		$dataGridView.FirstDisplayedScrollingRowIndex = $lastRowIndex
	}
}

# Insert or update video details in database
function Upsert-VideoDetails {
	param (
		[System.Int32]$videoId,
		[PSCustomObject]$videoDetails
	)
	
	try {
		$stmt = $db.Prepare(@"
INSERT INTO VideoDetails (VideoId, TMDBId, VideoType, Title, Overview, ReleaseDate)
VALUES (?, ?, ?, ?, ?, ?)
ON CONFLICT(TMDBId, VideoType) DO UPDATE SET
	Title = excluded.Title,
	Overview = excluded.Overview,
	ReleaseDate = excluded.ReleaseDate;
"@)
		$db.BindInt($stmt,   1, $videoId)
		$db.BindInt64($stmt, 2, $videoDetails.TMDBId)
		$db.BindText($stmt,  3, $videoDetails.VideoType)
		$db.BindText($stmt,  4, $videoDetails.Title)
		$db.BindText($stmt,  5,$videoDetails.Overview)
		$db.BindText($stmt,  6, $videoDetails.ReleaseDate)
		
		$result = $db.Step($stmt)
		$db.Finalize($stmt)
	}
	catch {
		Write-Warning "function Upsert-VideoDetails: ERROR: Unknown error occured: $_"
		Write-Warning $Error[0].Exception.GetType().FullName
		throw $Error
	}
}


# Insert or update genres
function Upsert-Genres {
	param (
		[System.Int64]$id,
		$genres
	)
	
	foreach ($genre in $genres) {
		try {
			$stmt = $db.Prepare("INSERT OR IGNORE INTO Genres (TMDBId, GenreName) VALUES (?, ?);")
			$db.BindInt64($stmt, 1, $genre.id)
			$db.BindText($stmt,  2, $genre.name)
			$result = $db.Step($stmt)
			$db.Finalize($stmt)
			
			$stmt = $db.Prepare("INSERT OR IGNORE INTO VideoGenres (VideoId, GenreTMDBId) VALUES (?, ?);")
			$db.BindInt($stmt,   1, $id)
			$db.BindInt64($stmt, 2, $genre.id)
			$result = $db.Step($stmt)
			$db.Finalize($stmt)
		}
		catch {
			Write-Warning "function Upsert-Genres: ERROR: Unknown error occured: $_"
			Write-Warning $Error[0].Exception.GetType().FullName
			throw $Error
		}
	}
}

# Insert or update actors in actors table
function Upsert-Actors {
	param (
		[System.Int32]$id,
		$actors
	)
	
	foreach ($actor in $actors) {
		try {
			$stmt = $db.Prepare("INSERT OR IGNORE INTO Actors (TMDBId, ActorName) VALUES (?, ?);")
			$db.BindInt64($stmt, 1, $actor.id)
			$db.BindText($stmt,  2, $actor.name)
			$result = $db.Step($stmt)
			$db.Finalize($stmt)
			
			$stmt = $db.Prepare("INSERT OR IGNORE INTO VideoActors (VideoId, ActorTMDBId) VALUES (?, ?);")
			$db.BindInt($stmt,   1, $id)
			$db.BindInt64($stmt, 2, $actor.id)
			$result = $db.Step($stmt)
			$db.Finalize($stmt)
		}
		catch {
			Write-Warning "function Upsert-Actors: ERROR: Unknown error occured: $_"
			Write-Warning $Error[0].Exception.GetType().FullName
			throw $Error
		}
	}
}


# Insert or update tables where video belongs to
function Upsert-VideoBelongsTo {
	param (
		[System.Int32]$videoListId,
		[String]$videoType,
		[System.Int64]$tmdbBelongsToId,
		[String]$name,
		[String]$overview
	)
	
	# ID from database where a video belongs to
	$belongsToId = $null
	
	try {
		# Query database if the entry where a video belongs to exists
		$stmt = $db.Prepare("SELECT DISTINCT BelongsToId FROM BelongsTo WHERE TMDBId = ? AND VideoType = ?;")
		$db.BindInt64($stmt, 1, $tmdbBelongsToId)
		$db.BindText($stmt,  2, $videoType)
		
		$result = $db.StepAndGetRow($stmt, 1)
		if ($result[0] -eq $([SQLiteHelper]::SQLITE_ROW)) { # SQLITE_ROW
			# The select statement returned an entry
			$belongsToId = [System.Int64]$result[1][0]
		}
		$db.Finalize($stmt)
	}
	catch {
		Write-Warning "function Upsert-VideoBelongsTo: ERROR: Unknown error occured while quering BelongsTo ID: $_"
		Write-Warning $Error[0].Exception.GetType().FullName
		throw $Error
	}
	
	if ($belongsToId -eq $null) {
		# The general information for the series or collection has not been found
		try {
			# Update the BelongsTo table with TMDB ID and type, also the name of the 
			# series or collection and the overview.
			$stmt = $db.Prepare("INSERT INTO BelongsTo (TMDBId, VideoType, BelongsToName, Overview) VALUES (?, ?, ?, ?);")
			$db.BindInt64($stmt, 1, $tmdbBelongsToId)
			$db.BindText($stmt,  2, $videoType)
			$db.BindText($stmt,  3, $name)
			$db.BindText($stmt,  4, $overview)
			$result = $db.Step($stmt)
			
			# Get ID of last inserted row
			$belongsToId = $db.GetLastInsertRowId()
			$db.Finalize($stmt)
		}
		catch {
			Write-Warning "function Upsert-VideoBelongsTo: ERROR: Unknown error occured: $_"
			Write-Warning $Error[0].Exception.GetType().FullName
			throw $Error
		}
	}
	
	try {
		# Create a n:m relation of the videos and where they belong to.
		$stmt = $db.Prepare("INSERT OR IGNORE INTO VideoBelongsTo (VideoId, BelongsToId) VALUES (?, ?);")
		$db.BindInt($stmt, 1, $videoListId)
		$db.BindInt($stmt, 2, $belongsToId)
		$result = $db.Step($stmt)
		$db.Finalize($stmt)
	}
	catch {
		Write-Warning "function Upsert-VideoBelongsTo: ERROR: Unknown error occured: $_"
		Write-Warning $Error[0].Exception.GetType().FullName
		throw $Error
	}
}

# Get series ID from DB if possible
function Get-SeriesFromDB {
	param (
		[String]$seriesPath,
		[System.Int64]$id
	)
	
	# Initialize return value
	$retVal = [PSCustomObject]@{
		ID = $null
		name = $null
		overview = $null
		adult = "false"
	}
	
	if ($id -gt 0) {
		try {
			# Query database if series path exists
			$stmt = $db.Prepare(@"
SELECT DISTINCT BelongsTo.TMDBId, BelongsTo.BelongsToName, BelongsTo.OverView, VideoList.isAdult
FROM VideoList
LEFT JOIN VideoBelongsTo ON VideoList.Id = VideoBelongsTo.VideoID
LEFT JOIN BelongsTo ON BelongsTo.BelongsToId = VideoBelongsTo.BelongsToId
WHERE BelongsTo.VideoType = 'S' AND BelongsTo.TMDBId = ?;
"@)

			$db.BindInt64($stmt, 1, $id)
			
			# Get results from database query
			while ($true) {
				$result = $db.StepAndGetRow($stmt)
				if ($result[0] -ne $([SQLiteHelper]::SQLITE_ROW)) { # SQLITE_ROW
					break
				}
				if( $result[1][0] -ne "0") {
					# Series found, use ID for series
					$retVal.ID = [System.Int64]$result[1][0]
					$retVal.name = $result[1][1]
					$retVal.Overview = $result[1][2]
					if ($result[1][3] -eq 1) {
						$retVal.adult = "true"
					}
					break
				}
			}
			$db.Finalize($stmt)
		}
		catch {
			Write-Warning "function Get-SeriesFromDB: ERROR: Unknown error occured: $_"
			Write-Warning $Error[0].Exception.GetType().FullName
			throw $Error
		}
	} else {
		try {
			# Query database if series path exists
			$stmt = $db.Prepare(@"
SELECT DISTINCT
	VideoList.Id,
	VideoList.FileName,
	VideoList.FilePath,
	VideoList.FileExists,
	VideoList.TMDBId,
	VideoList.VideoType,
	BelongsTo.BelongsToName,
	VideoList.IsAdult
FROM VideoList 
LEFT JOIN VideoBelongsTo ON VideoList.TMDBId = VideoBelongsTo.VideoID
LEFT JOIN BelongsTo ON BelongsTo.TMDBId = VideoBelongsTo.BelongsToId
WHERE VideoList.VideoType = 'S' AND
	VideoList.FileExists = '1' AND
	VideoList.TMDBId <> 0 AND
	VideoList.FilePath LIKE ?;
"@)
			$db.BindText($stmt, 1, "%$seriesPath%")
			
			# Get results from database query
			while ($true) {
				$result = $db.StepAndGetRow($stmt)
				if ($result[0] -ne $([SQLiteHelper]::SQLITE_ROW)) { # SQLITE_ROW
					break
				}
				$sId = $result[1][4]
				if( $Id -ne "0") {
					# Series found, use ID for series
					$retVal.ID = [System.Int32]$stmt.col(4)
					$retVal.name = $stmt.col(6)
					if ($stmt.col(7) -eq 1) {
						$retVal.adult = "true"
					}
					break
				}
			}
			$db.Finalize($stmt)
		}
		catch {
			Write-Warning "function Get-SeriesFromDB: ERROR: Unknown error occured: $_"
			Write-Warning $Error[0].Exception.GetType().FullName
			throw $Error
		}
	}
	
	return $retVal
}

# Get overview for the ID a video belongs to
function Get-OverviewForVideoBelongsto {
	param (
		[System.Int64]$videoId
	)
	
	# Initialize return value
	$overView = ""
	
	try {
		# Query database
		$stmt = $db.Prepare(@"
SELECT DISTINCT BelongsTo.OverView
FROM VideoList
LEFT JOIN VideoBelongsTo ON VideoList.Id = VideoBelongsTo.VideoID
LEFT JOIN BelongsTo ON BelongsTo.BelongsToId = VideoBelongsTo.BelongsToId
WHERE VideoList.Id = ?;
"@)

		$db.BindInt64($stmt, 1, $id)
			
		# Get results from database query
		while ($true) {
			$result = $db.StepAndGetRow($stmt)
			if ($result[0] -ne $([SQLiteHelper]::SQLITE_ROW)) { # SQLITE_ROW
				break
			}
			if( $result[1][0] -ne "0") {
				# Video found
				$overView = $result[1][0]
				break
			}
		}
		$db.Finalize($stmt)
	}
	catch {
		Write-Warning "function Get-OverviewForVideoBelongsto: ERROR: Unknown error occured: $_"
		Write-Warning $Error[0].Exception.GetType().FullName
		throw $Error
	}
	
	return $overView
}

# Get video details from database
function Get-VideoDetails {
	param (
		[System.Int64]$tmdbId,
		[String]$videoType
	)
	
	$details = [PSCustomObject]@{
		title = $null
		overview = $null
		releaseDate = $null
	}
	
	# Check if TMDB Id is not zero
	if ( $tmdbId -ne 0) {
		try {
			# Get title and overview
			$stmt = $db.Prepare("SELECT DISTINCT Title, Overview, ReleaseDate FROM VideoDetails WHERE TMDBId = ? AND VideoType = ?;")
			$db.BindInt64($stmt, 1, $tmdbId)
			$db.BindText($stmt,  2, $videoType)
			
			# Get results from database query
			$columnCount = $db.GetColumnCount($stmt)
			while ($true) {
				$result = $db.StepAndGetRow($stmt, $columnCount)
				if ($result[0] -ne $([SQLiteHelper]::SQLITE_ROW)) { # SQLITE_ROW
					break
				}
				# Get results from database query
				$details.title = $result[1][0]
				$details.overview = $result[1][1]
				$releaseDate = $result[1][2]
				if (($releaseDate -ne $null) -and ($releaseDate.Length -ge 4)) {
					$details.releaseDate = $releaseDate.SubString(0, 4)
				} else {
					$details.releaseDate = "-"
				}
			}
			$db.Finalize($stmt)
		}
		catch {
			Write-Warning "function Get-VideoDetails: ERROR: Unknown error occured: $_"
			Write-Warning $Error[0].Exception.GetType().FullName
			throw $Error
		}
	}
	
	return $details
}

# Get string of assigned genres for an ID
function Get-Genres {
	param (
		[System.Int32]$id
	)
	
	# Initiaize list of genres
	$genres = ""
	
	# Check if Id is not zero
	if ( $id -ne 0) {
		try {
			# Get title and overview
			$stmt = $db.Prepare("SELECT DISTINCT GenreName FROM Genres INNER JOIN VideoGenres ON Genres.TMDBId = VideoGenres.GenreTMDBId WHERE VideoGenres.VideoId = ?;")
			$db.BindInt($stmt, 1, $id)
			
			$genresList = @()
			while ($true) {
				$result = $db.StepAndGetRow($stmt, 1)
				if ($result[0] -ne $([SQLiteHelper]::SQLITE_ROW)) { # SQLITE_ROW
					break
				}
				if ($result[1][0] -ne "") {
					# Entry found, add to list
					$genresList += [String]$result[1][0]
				}
			}
			$genres = [string]::Join(", ", $genresList)
			$db.Finalize($stmt)
		}
		catch {
			Write-Warning "function Get-Genres: ERROR: Unknown error occured: $_"
			Write-Warning $Error[0].Exception.GetType().FullName
			throw $Error
		}
	}
	
	return $genres
}

# Get string of assigned genres for an ID
function Get-Actors {
	param (
		[System.Int32]$id
	)
	
	# Initiaize list of actors
	$actors = ""
	
	# Check if Id is not zero
	if ( $id -ne 0) {
		try {
			# Get title and overview
			$stmt = $db.Prepare(@"
SELECT DISTINCT ActorName FROM Actors
INNER JOIN VideoActors ON Actors.TMDBId = VideoActors.ActorTMDBId
WHERE VideoActors.VideoId = ?;
"@)
			$db.BindInt($stmt, 1, $id)
			
			$actorsList = @()
			while ($true) {
				$result = $db.StepAndGetRow($stmt, 1)
				if ($result[0] -ne $([SQLiteHelper]::SQLITE_ROW)) { # SQLITE_ROW
					break
				}
				if ($result[1][0] -ne "") {
					# Entry found, add to list
					$actorsList += [String]$result[1][0]
				}
			}
			$actors = [string]::Join(", ", $actorsList)
			$db.Finalize($stmt)
		}
		catch {
			Write-Warning "function Get-Actors: ERROR: Unknown error occured: $_"
			Write-Warning $Error[0].Exception.GetType().FullName
			throw $Error
		}
	}
	
	return $actors
}

# Mark files which only exists in database but not in file system anymore
function Mark-NonexistentFiles {
	# Array for ID if video file was found in file system
	$found = @()
	$notFound = @()
	
	$progressTextBox.Text = (Get-Translation -language $config.language -key "Main.Text.CheckingFileExistence")
	Start-Sleep -Milliseconds 5
	$form.Refresh()
	
	try {
		# Query the database
		$stmt = $db.Prepare("SELECT DISTINCT Id, FileName, FilePath, VideoType, FileExists FROM VideoList;")
		$columnCount = $db.GetColumnCount($stmt)
		while ($true) {
			$result = $db.StepAndGetRow($stmt, $columnCount)
			if ($result[0] -ne $([SQLiteHelper]::SQLITE_ROW)) { # SQLITE_ROW
				break
			}
			$filePath = ""
			
			# Extract video ID and stored state of file existing from database
			$id = [System.Int32]$result[1][0]
			$DBfileExist = [System.Int32]$result[1][4]
			
			# Check the video type
			if ($result[1][3] -eq 'M') {
				# Movie
				if ($global:moviePath.GetIsAvailable()) {
					if( -Not( [String]::IsNullOrEmpty( $global:moviePath.GetPath()))) {
						$filePath = $global:moviePath.GetPath() + "\" + $result[1][2]
					}
				} else {
					break
				}
			} elseif ($result[1][3] -eq 'S') {
				# Series
				if ($global:seriesPath.GetIsAvailable()) {
					if( -Not( [String]::IsNullOrEmpty( $global:seriesPath.GetPath()))) {
						$filePath = $global:seriesPath.GetPath() + "\" + $result[1][2]
					}
				} else {
					break
				}
			}
			
			# If the returned file path is not empty check
			if( -Not( [String]::IsNullOrEmpty( $filePath ))) {
				$filePath += "\" + $result[1][1]
				if (Test-Path -LiteralPath $filePath) {
					if ($DBfileExist -eq 0 ) {
						$found += [Int]$id
					}
				} else {
					if ($DBfileExist -eq 1 ) {
						$notFound += [Int]$id
					}
				}
			} else {
				# File path could not be tested
				$notFound += [Int]$id
			}
		}
		$db.Finalize($stmt)
		
		# Update database
		if ($found.Length -ge 1) {
			# Set FileExists to 1 for all files which has been found on the given location again
			# This method is better against SQL injection
			$tmp = ($found | ForEach-Object { "?" }) -join ","
			$stmt = $db.Prepare("UPDATE VideoList Set FileExists = 1 WHERE Id IN ($tmp);")
			
			$index = 1
			foreach ($id in $found) {
				$db.BindInt($stmt, $index, $id)
				$index++
			}
			$result = $db.Step($stmt)
			$db.Finalize($stmt)
		}
		
		if ($notFound.Length -ge 1 ) {
			# Set FileExists to 0 for all files which are not found on the given location
			# This method is better against SQL injection
			$tmp = ($notFound | ForEach-Object { "?" }) -join ","
			$stmt = $db.Prepare("UPDATE VideoList Set FileExists = 0 WHERE Id IN ($tmp);")
			
			$index = 1
			foreach ($id in $notFound) {
				$db.BindInt($stmt, $index, $id)
				$index++
			}
			$result = $db.Step($stmt)
			$db.Finalize($stmt)
		}
	}
	catch {
		Write-Warning "function Mark-NonexistentFiles: ERROR: Unknown error occured: $_"
		Write-Warning $Error[0].Exception.GetType().FullName
		throw $Error
	}
}

# Delete genres association for given video id.
# Also delete genres if they are no longer used.
function Delete-Genres {
	param (
		[System.Int32]$videoId
	)
	
	try {
		# Delete all genres which are only associated to the given video
		$stmt = $db.Prepare(@"
DELETE FROM Genres
WHERE TMDBId IN ( 
SELECT GenreTMDBId
FROM VideoGenres
GROUP BY GenreTMDBId
HAVING COUNT(VideoId) = 1
AND MAX(VideoId) = ?
);
"@)
		$db.BindInt($stmt, 1, $videoId)
		$result = $db.Step($stmt)
		$db.Finalize($stmt)
		
		# Delete all n:m relations for the given video
		$stmt = $db.Prepare("DELETE FROM VideoGenres WHERE VideoId = ?;")
		$db.BindInt($stmt, 1, $videoId)
		$result = $db.Step($stmt)
		$db.Finalize($stmt)
	}
	catch {
		Write-Warning "function Delete-Genres: ERROR: Unknown error occured: $_"
		Write-Warning $Error[0].Exception.GetType().FullName
		throw $Error
	}
}

# Delete actors association for given video id.
# Also delete actors if they are no longer used.
function Delete-Actors {
	param (
		[System.Int32]$videoId
	)
	
	try {
		# Delete all actors which are only associated to the given video
		$stmt = $db.Prepare(@"
DELETE FROM Actors
WHERE TMDBId IN ( 
SELECT ActorTMDBId
FROM VideoActors
GROUP BY ActorTMDBId
HAVING COUNT(VideoId) = 1
AND MAX(VideoId) = ?
);
"@)
		$db.BindInt($stmt, 1, $videoId)
		$result = $db.Step($stmt)
		$db.Finalize($stmt)
		
		# Delete all n:m relations for the given video
		$stmt = $db.Prepare("DELETE FROM VideoActors WHERE VideoId = ?;")
		$db.BindInt($stmt, 1, $videoId)
		$result = $db.Step($stmt)
		$db.Finalize($stmt)
	}
	catch {
		Write-Warning "function Delete-Actors: ERROR: Unknown error occured: $_"
		Write-Warning $Error[0].Exception.GetType().FullName
		throw $Error
	}
}

# Delete data where a video belongs to
function Delete-BelongsTo {
	param (
		[System.Int32]$videoId
	)
	
	try {
		# Delete all series or collection information which are only associated to the given video
		$stmt = $db.Prepare(@"
DELETE FROM BelongsTo
WHERE BelongsToId IN ( 
SELECT BelongsToId
FROM VideoBelongsTo
GROUP BY BelongsToId
HAVING COUNT(VideoId) = 1
AND MAX(VideoId) = ?
);
"@)
		$db.BindInt($stmt, 1, $videoId)
		$result = $db.Step($stmt)
		$db.Finalize($stmt)
		
		# Delete all n:m relations for the given video
		$stmt = $db.Prepare("DELETE FROM VideoBelongsTo WHERE VideoId = ?;")
		$db.BindInt($stmt, 1, $videoId)
		$result = $db.Step($stmt)
		$db.Finalize($stmt)
	}
	catch {
		Write-Warning "function Delete-BelongsTo: ERROR: Unknown error occured: $_"
		Write-Warning $Error[0].Exception.GetType().FullName
		throw $Error
	}
}

# Delete video details for given TMDB id and video type
function Delete-VideoDetails {
	param (
		[System.Int64]$tmdbId,
		[String]$videoType
	)
	
	try {
		# Delete all n:m relations for the given video
		$stmt = $db.Prepare("DELETE FROM VideoDetails WHERE TMDBId = ? AND VideoType = ?;")
		$db.BindInt64($stmt, 1, $tmdbId)
		$db.BindText($stmt, 1, $videoType)
		$result = $db.Step($stmt)
		$db.Finalize($stmt)
	}
	catch {
		Write-Warning "function Delete-VideoDetails: ERROR: Unknown error occured: $_"
		Write-Warning $Error[0].Exception.GetType().FullName
		throw $Error
	}
}

# Delete the video from the video list
function Delete-Video {
	param (
		[System.Int32]$videoId
	)
	
	try {
		$stmt = $db.Prepare("DELETE FROM VideoList WHERE Id = ?;")
		$db.BindInt($stmt, 1, $videoId)
		$result = $db.Step($stmt)
		$db.Finalize($stmt)
	}
	catch {
		Write-Warning "function Delete-Video: ERROR: Unknown error occured: $_"
		Write-Warning $Error[0].Exception.GetType().FullName
		throw $Error
	}
}

# Get the TMDB id where a given video belongs to
function Get-BelongsToId {
	param (
		[System.Int32]$videoId
	)
	
	# Initialize return value
	$tmdbId = $null
	try {
		# Delete all series or collection information which are only associated to the given video
		$stmt = $db.Prepare("SELECT BelongsToId FROM VideoBelongsTo WHERE VideoId = ?;")
		$db.BindInt($stmt, 1, $videoId)
		$result = $db.StepAndGetRow($stmt, 1)
		
		if ($result[0] -eq $([SQLiteHelper]::SQLITE_ROW)) { # SQLITE_ROW
			if ($result[1][0] -ne "") {
				# Entry found
				$tmdbId = [System.Int64]$result[1][0]
			}
		}
		$db.Finalize($stmt)
	}
	catch {
		Write-Warning "function Get-BelongsToId: ERROR: Unknown error occured: $_"
		Write-Warning $Error[0].Exception.GetType().FullName
		throw $Error
	}
	
	return $tmdbId
}

# Get all entries for a given series ID
function Get-AllExistingEpisodesForID {
	param (
		[System.Int32]$belongsToId
	)
	
	# Create list for all existing episodes of the series
	$existingEpisodes = @()
	
	try {
		$stmt = $db.Prepare(@"
SELECT DISTINCT
VideoList.FileName
FROM VideoList 
LEFT JOIN VideoBelongsTo ON VideoList.Id = VideoBelongsTo.VideoID
WHERE VideoBelongsTo.BelongsToId = ? AND VideoList.VideoType = 'S';
"@)
		$res = $db.BindInt($stmt, 1, $belongsToId)
		
		# Get results from database query
		while ($true) {
			$result = $db.StepAndGetRow($stmt)
			if ($result[0] -ne $([SQLiteHelper]::SQLITE_ROW)) { # SQLITE_ROW
				break
			}
			
			$details = Get-SeriesDetailsFromFilename -fileName $result[1][0]
			$existingEpisodes += @(@{ season = $details.season; episode = $details.episode })
		}
		
		$db.Finalize($stmt)
	}
	catch {
		Write-Warning "function Get-AllExistingEpisodesForID: ERROR: Unknown error occured: $_"
		Write-Warning $Error[0].Exception.GetType().FullName
		throw $Error
	}
	
	return $existingEpisodes
}


# Get all entries for a given collection ID
function Get-AllExistingPartsForID {
	param (
		[System.Int32]$belongsToId
	)
	
	# Create list for all existing parts of the collection
	$existingParts = @()
	
	try {
		$stmt = $db.Prepare(@"
SELECT DISTINCT
VideoList.TMDBId
FROM VideoList 
LEFT JOIN VideoBelongsTo ON VideoList.Id = VideoBelongsTo.VideoID
WHERE VideoBelongsTo.BelongsToId = ? AND VideoList.VideoType = 'M';
"@)
		$res = $db.BindInt($stmt, 1, $belongsToId)
		
		# Get results from database query
		while ($true) {
			$result = $db.StepAndGetRow($stmt)
			if ($result[0] -ne $([SQLiteHelper]::SQLITE_ROW)) { # SQLITE_ROW
				break
			}
			$existingParts += @(@{ tmdbId = $result[1][0] })
		}
		
		$db.Finalize($stmt)
	}
	catch {
		Write-Warning "function Get-AllExistingPartsForID: ERROR: Unknown error occured: $_"
		Write-Warning $Error[0].Exception.GetType().FullName
		throw $Error
	}
	
	return $existingParts
}


# Get the TMDB id where a given video belongs to
function Get-TMDBBelongsTo {
	param (
		[System.Int32]$belongsToId
	)
	
	# Prepare return object
	$tmdbBelongsTo = [PSCustomObject]@{
		TMDBId        = $null
		BelongsToName = $null
		Type          = $null
		Overview      = $null
	}
	
	try {
		# Get the TMDB id and video type where the given ID belongs to
		$stmt = $db.Prepare("SELECT TMDBId, BelongsToName, VideoType, OverView FROM BelongsTo WHERE BelongsToId = ?;")
		$db.BindInt($stmt, 1, $belongsToId)
		$result = $db.StepAndGetRow($stmt)
		if ($result[0] -eq $([SQLiteHelper]::SQLITE_ROW)) { # SQLITE_ROW
			if ($result[1][0] -ne "") {
				# Entry found
				$tmdbBelongsTo.TMDBId        = [System.Int64]$result[1][0]
				$tmdbBelongsTo.BelongsToName = [String]$result[1][1]
				$tmdbBelongsTo.Type          = [String]$result[1][2]
				$tmdbBelongsTo.Overview      = [String]$result[1][3]
			}
		}
		$db.Finalize($stmt)
	}
	catch {
		Write-Warning "function Get-TMDBBelongsTo: ERROR: Unknown error occured: $_"
		Write-Warning $Error[0].Exception.GetType().FullName
		throw $Error
	}
	
	return $tmdbBelongsTo
}


# Mark rows in the datagrid red if the FileExists value is zero
function Mark-Red {
	foreach ($row in $dataGridView.Rows) {
		if ($row.Cells["FileExists"].Value -eq 0) {
			$row.DefaultCellStyle.BackColor = [System.Drawing.Color]::Red
		}
	}
}


######################################################
# Change mouse cursor to waiting symbol
function cursorWait {
	$form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
}

# Change mouse cursor back to standard
function cursorDefault {
	$form.Cursor = [System.Windows.Forms.Cursors]::Default
}


######################################################
# Language functions
# Get translation based on used langauge
function Get-Translation {
	param (
		[String]$language,
		[String]$key
	)
	
	# Check if requested translation exist
	if (!($translations.PSObject.Properties.Name -contains $language)) {
		# Requested language does not exist, try default language
		$language = $defaultLanguage
	}
	
	# Check if key exists in requested translation
	if ($translations.$language.PSObject.Properties.Name -contains $key) {
		# Yes, key exists, get translation
		$translation = $translations.$language.$key
	} else {
		# No, key is not translated yet, check for default langauge
		$language = $defaultLanguage
		if ($translations.$language.PSObject.Properties.Name -contains $key) {
			$translation = $translations.$language.$key
		} else {
			# Requested language does not exist
			$translation = $key
		}
	}
	
	return $translation
}

######################################################
# Translation
######################################################
# Function for setting and updating text in main dialog
function Set-MainText {
	param (
		[String]$language
	)
	$form.Text = (Get-Translation -language $language -key "Main.Form.Name") + " " + $VDBVERSION
	
	$buttonOpenConfig.Text = Get-Translation -language $language -key "Main.Button.Configuration"
	Adjust-ButtonFontSize -button $buttonOpenConfig
	
	$analyzeButton.Text = Get-Translation -language $language -key "Main.Button.Analyze"
	Adjust-ButtonFontSize -button $analyzeButton
	
	$searchButton.Text = Get-Translation -language $language -key "Main.Button.Search"
	Adjust-ButtonFontSize -button $searchButton
	
	$duplicatesButton.Text = Get-Translation -language $language -key "Main.Button.Duplicates"
	Adjust-ButtonFontSize -button $duplicatesButton
	
	$deleteButton.Text = Get-Translation -language $language -key "Main.Button.Delete"
	Adjust-ButtonFontSize -button $deleteButton
	
	$exportButton.Text = Get-Translation -language $language -key "Main.Button.Export"
	Adjust-ButtonFontSize -button $exportButton
	
	$renameButton.Text = Get-Translation -language $language -key "Main.Button.Rename"
	Adjust-ButtonFontSize -button $renameButton
	
	$moveButton.Text = Get-Translation -language $language -key "Main.Button.Move"
	Adjust-ButtonFontSize -button $moveButton
	
	$rescanButton.Text = Get-Translation -language $language -key "Main.Button.Rescan"
	Adjust-ButtonFontSize -button $rescanButton

	$fillUpButton.Text = Get-Translation -language $language -key "Main.Button.FillUp"
	Adjust-ButtonFontSize -button $rescanButton
	
	$titleLabel.Text = Get-Translation -language $language -key "Main.Label.Title"
	$yearLabel.Text = Get-Translation -language $language -key "Main.Label.Year"
	$overviewLabel.Text = Get-Translation -language $language -key "Main.Label.Overview"
	$genresLabel.Text = Get-Translation -language $language -key "Main.Label.Genres"
	$actorsLabel.Text = Get-Translation -language $language -key "Main.Label.Actors"
}
######################################################


######################################################
# File functions
# Extract title and release year from filename
function Get-TitleAndYearFromFilename {
	param (
		[String]$fileName
	)
	
	# Initialize variables for title and release year
	$returnValue = [PSCustomObject]@{
		title = $null
		year = $null
		tmdbID = 0
	}
	
	# Regex pattern for extracting filename, year and/or TMDB id
	$pattern = '^(?<TITLE>[^\[\]]*?)\s*(?:\((?<YEAR>\d{4})\))?\s*(?:\[TMDBID=(?<TMDBID>\d+)\])?\s*\.(?<SUFFIX>[^.]+)$|^(?<TITLE>[^\(\)]*?)\s*(?:\[TMDBID=(?<TMDBID>\d+)\])?\s*(?:\((?<YEAR>\d{4})\))?\s*\.(?<SUFFIX>[^.]+)$'
	
	# Check if filename matches the pattern and extract data
	if ($fileName -match $pattern) {
		$returnValue.title = $matches['TITLE'].Trim()
		$returnValue.year = $matches['YEAR']
		$returnValue.tmdbID = $matches['TMDBID']
		$suffix = $matches['SUFFIX']
	} else {
		# No release year was submitted
		$returnValue.title = $fileName.Trim()
		$returnValue.year = ""
	}
	
	return [PSCustomObject]$returnValue
}

# Extract series title and release year from path
function Get-TitleAndYearFromPath {
	param (
		[String]$folderName
	)
	
	# Extract the name of the folder
	$folder = Split-Path -Path $folderName -Leaf
	$folder = $folder.TrimEnd("\")
	
	# Initialize variables for title and release year
	$returnValue = [PSCustomObject]@{
		title = $null
		year = 0
		tmdbId = [System.Int64]0
	}
	
	# Regex pattern for extracting pathname, year and/or TMDB id
	$pattern = '^(?<TITLE>[^\[\]]*?)\s*(?:\((?<YEAR>\d{4})\))?\s*(?:\[TMDBID=(?<TMDBID>\d+)\])?$|^(?<TITLE>[^\(\)]*?)\s*(?:\[TMDBID=(?<TMDBID>\d+)\])?\s*(?:\((?<YEAR>\d{4})\))?$'
	
	# Check if folderName matches the pattern and extract data
	if ($folder -match $pattern) {
		$returnValue.title = $matches['TITLE'].Trim()
		$returnValue.year = [System.Int32]$matches['YEAR']
		$returnValue.tmdbID = [System.Int64]$matches['TMDBID']
	} else {
		# No release year was submitted
		$returnValue.title = $folderName.Trim()
	}
	
	return [PSCustomObject]$returnValue
}

# Extract series details from filename, if given:
# - season
# - episode
# - part
# - title
function Get-SeriesDetailsFromFilename {
	param (
		[string]$fileName
	)
	
	# Initialize return object
	$result = [PSCustomObject]@{
		season = $null
		episode = $null
		part = $null
		title = $null
	}
	
	# Remove file extension
	$fileName = [System.IO.Path]::GetFileNameWithoutExtension($fileName)
	
	if ($fileName -match '^[Ss](\d+)[Ee](\d+)[Pp](\d+)\s*(.*)$') {
		# Extract season, episode, part and title
		$result.season = [System.Int32]$matches[1].Trim()
		$result.episode = [System.Int32]$matches[2].Trim()
		$result.part = [System.Int32]$matches[3].Trim()
		$result.title = $matches[4].Trim()
	} elseif ($fileName -match '^[Ss](\d+)[Ee](\d+)\s*(.*)$') {
		# Extract season, episode and title
		$result.season = [System.Int32]$matches[1].Trim()
		$result.episode = [System.Int32]$matches[2].Trim()
		$result.title = $matches[3].Trim()
	} else {
		# Get only title
		$result.title = $fileName.Trim()
	}

	return $result
}


# Scan movie folders recursivly and analyze video files
function Analyze-Movies {
	if (-not ([string]::IsNullOrEmpty($global:moviePath))) {
		if ($global:moviePath.GetIsAvailable()) {
			# Get current folder path length including dive letter
			$curMoviePath = $global:moviePath.GetPath()
			$folderPathLength = $curMoviePath.Length
			# Check for trailing backslash
			if ($curMoviePath.SubString($curMoviePath.Length - 1, 1) -eq "\" ) {
				$folderPathLength--
			}
			
			# Set progress text box
			$progressTextBox.Text = (Get-Translation -language $config.language -key "Main.Text.LoadDirEntries") + $curMoviePath
			Start-Sleep -Milliseconds 5
			$form.Refresh()
			
			# Get all video in folder structure
			$videoFiles = Get-ChildItem -Path $curMoviePath -Recurse -File -Include $fileExtensions
			$videoFiles = @($videoFiles | Sort-Object -Property @{Expression = {$_.DirectoryName}}, @{Expression = {$_.Name}})
			
			# Check if folder is not empty
			if (-not ([string]::IsNullOrEmpty($videoFiles))) {
				# Initialize some variables
				$fileFound = $false
				
				# Loop over all files
				$fileCounter = 0
				$fileTotal = $videoFiles.Count
				foreach ($file in $videoFiles) {
					# Check if a maximum number of files to csan is set and is reached
					if (($config.MaxScans -gt 0) -and ($fileCounter -ge $config.MaxScans)) {
						break
					}
					
					# Extract file path of video
					$filePathOnly = [System.IO.Path]::GetDirectoryName($file.FullName)
					$filePathForDB = $filePathOnly.SubString($folderPathLength)
					$filePathForDB = $filePathForDB.Trim("\")
					# Extract filename
					$fileName = [System.IO.Path]::GetFileName($file.FullName)
					
					# Update GUI status folder field
					$progressTextBox.Text = (Get-Translation -language $config.language -key "Config.Movie.Button") + ": [$((++$fileCounter))] " + (Get-Translation -language $config.language -key "Main.Text.Of") + " [$fileTotal]: $($file.FullName)"
					Start-Sleep -Milliseconds 1
					$form.Refresh()
					
					# Check if video is already in displayed video list
					$fileFound = $false
					foreach ($row in $dataGridView.Rows) {
						# Check all rows, get filename and path from DataGridView
						$dgvFileName = $row.Cells[$GRIDVIEWCOLUMNFILENAME].Value
						$dgvFilePath = $row.Cells[$GRIDVIEWCOLUMNFILEPATH].Value
						$dgvFileSize = [System.Int64]$row.Cells[$GRIDVIEWCOLUMNFILESIZE].Value

						# Check if path and filename exists
						if (($filePathForDB -eq $dgvFilePath) -and ($fileName -eq $dgvFileName) -and ($file.Length -eq $dgvFileSize)) {
							# Found
							$fileFound = $true
							break
						}
					}
					
					# Insert the movie into the database if it was not found
					if (-not ($fileFound)) {
						# Extract title and year from filename
						$ty = Get-TitleAndYearFromFilename -fileName $filename
						
						# Analyze file
						$analyzed = $true
						try {
							$videoInfo = Get-VideoInfoFromFile -filePath $file.FullName -baseFolder $curMoviePath
						}
						catch [CustomException] {
							$analyzed = $false		
							# Error number should only be one of:
							# ERRORFFPOBEGUNKNOWN
							# ERRORFFPROBERCNOTZERO
							# ERRORFFPROBENOSTREAMS
							# ERRORFFPROBENOVIDEOSTREAM
							$e = $_.Exception
							
							$outString = Get-Translation -language $config.Language -key "File"
							$outString += ": '" + $file.FullName + "'"
							write-host $outString
							
							$outString = Get-Translation -language $config.Language -key "Error"
							$outString += ": " + $e.myMessage
							write-host $outString
						} catch {
							$analyzed = $false
							write-host $Error[0].Exception.GetType().FullName
						}
						
						if ($analyzed) {
							$videoId = 0
							if (-not ([string]::IsNullOrEmpty($config.ApiKey))) {
								# Query information from themoviedb.org
								$movieInfo = Get-MovieInfoFromTMDB -fileName $fileName
								
								# If the movie was found query movie details
								if (-not ([string]::IsNullOrEmpty($movieInfo))) {
									if ($movieInfo.id -gt 0 ) {
										# Insert or update in database
										$videoId = Upsert-VideoInfo -videoInfo $videoInfo -fileSize $file.Length -tmdbId $movieInfo.id -title $ty.title -vote $movieInfo.vote_average -videoType 'M' -isAdult $movieInfo.adult
										
										# Search for movie details in TMDB
										$movieDetails = Get-MovieDetailsFromTMDB -movieId $movieInfo.id
										
										if ($movieDetails -ne $null) {
											# Insert data into database
											$movieDetailsObject = [PSCustomObject]@{
												TMDBId		= [System.Int64]$movieDetails.id
												VideoType   = "M"
												Title		= $movieDetails.title
												Overview	= $movieDetails.overview
												ReleaseDate = $movieDetails.release_date
											}
											
											# Insert or update movie details
											Upsert-VideoDetails -videoId $videoId -videoDetails $movieDetailsObject
											
											# Insert or update genres
											Upsert-Genres -id $videoId -genres $movieDetails.genres
											
											# Insert or update actors
											Upsert-Actors -id $videoId -actors $movieDetails.credits.cast
											
											# Check if the movie belongs to a collection
											$collection = $movieDetails.belongs_to_collection
											if (-not ([string]::IsNullOrEmpty($collection))) {
												# Query details about movie collection from TMDB
												$collectionDetails = Get-CollectionDetailsFromTMDB -collection $movieDetails.belongs_to_collection.id
												if (-not ( $collectionDetails -eq $null )) {
													# Insert or update collection details
													Upsert-VideoBelongsTo -videoListId $videoId -videoType "M" -tmdbBelongsToId $collectionDetails.id -name $collectionDetails.name -overview $collectionDetails.overview
												}
											}
										} else {
											# Movie hasn't been found in TMDB so insert or update basic information only
											$videoId = Upsert-VideoInfo -videoInfo $videoInfo -fileSize $file.Length -title $ty.title -tmdbId 0 -vote "0.0" -videoType 'M' -isAdult "false"
										}
										
									} else {
										# Movie hasn't been found in TMDB so insert or update basic information only
										$videoId = Upsert-VideoInfo -videoInfo $videoInfo -fileSize $file.Length -title $ty.title -tmdbId 0 -vote "0.0" -videoType 'M' -isAdult "false"
									}
								} else {
									# Movie hasn't been found in TMDB so insert or update basic information only
									$videoId = Upsert-VideoInfo -videoInfo $videoInfo -fileSize $file.Length -title $ty.title -tmdbId 0 -vote "0.0" -videoType 'M' -isAdult "false"
								}
							} else {
								# No TMDB API key present so insert or update basic information only
								$videoId = Upsert-VideoInfo -videoInfo $videoInfo -fileSize $file.Length -title $ty.title -tmdbId 0 -vote "0.0" -videoType 'M' -isAdult "false"
							}
							
							# Load row
#							if ($videoId -ne 0) {
#								Load-DataRow $videoId
#								Scroll-To-Last-Element
#							}
						} else {
							# ToDo: Add warning to String and display after all files have been scanned
							Write-Warning "ERROR: File $($file.FullName) couldn't be analyzed!"
							Write-Warning ""
						}
					}
				}
			}
		}
	}
}

# Scan series folders recursivly and analyze video files
# 1. check if folder is accessible
# 2. get files ordered by folder structure and files
# 3. 
function Analyze-Series {
	if (-not ([string]::IsNullOrEmpty($global:seriesPath))) {
		if ($global:seriesPath.GetIsAvailable()) {
			# Get current folder path length including drive letter
			$curSeriesPath = $global:seriesPath.GetPath()
			$folderPathLength = $curSeriesPath.Length
			# Check for trailing backslash
			if ($curSeriesPath.SubString($curSeriesPath.Length - 1, 1) -eq "\" ) {
				$folderPathLength--
			}
			
			# Set progress text box
			$progressTextBox.Text = (Get-Translation -language $config.language -key "Main.Text.LoadDirEntries") + $curSeriesPath
			Start-Sleep -Milliseconds 1
			$form.Refresh()
			
			# Get all video in folder structure
			$videoFiles = Get-ChildItem -Path $curSeriesPath -Recurse -File -Include $fileExtensions
			$videoFiles = @($videoFiles | Sort-Object -Property @{Expression = {$_.DirectoryName}}, @{Expression = {$_.Name}})
			
			# Check if folder is not empty
			if (-not ([string]::IsNullOrEmpty($videoFiles))) {
				# Initialize some variables
				$fileFound = $false	            # Current file found
				$lastSeriesFolder = ""          # Folder of the series from last loop
				$lastSeriesID = [System.Int64]0	# TMDB of the series from the last loop, not from the episode
				$seriesID = [System.Int64]0		# TMDB-ID of the series found either in local database or received from TMDB
				$seriesInfo = $null
				
				# Loop over all files
				$fileCounter = 0
				$fileTotal = $videoFiles.Count
				foreach ($file in $videoFiles) {
					# Check if a maximum number of files to csan is set and is reached
					if (($config.MaxScans -gt 0) -and ($fileCounter -ge $config.MaxScans)) {
						break
					}
					
					# Extract name of series folder, do not use seasons sub structure, just use
					# \Series\Star Trek (1966)\s1e01 The Man Trap.mp4
					# \Series\Star Trek (1966)\s1e02 Charlie X.mp4
					$filePathOnly = [System.IO.Path]::GetDirectoryName($file.FullName)
					$filePathForDB = $filePathOnly.SubString($folderPathLength)
					$filePathForDB = $filePathForDB.Trim("\")
					
					# Check if series video is in a sub folder
					if ($filePathForDB.toLower() -eq $global:seriesPath.GetPath().toLower()) {
						Write-Warning (Get-Translation -language $config.language -key "Error.Series.NotInSubfolder")
						Write-Warning $file.FullName
					} else {
						# Extract filename
						$fileName = [System.IO.Path]::GetFileName($file.FullName)
						
						# Update GUI status folder field
						$progressTextBox.Text = (Get-Translation -language $config.language -key "Config.Series.Button") + ": [$((++$fileCounter))] " + (Get-Translation -language $config.language -key "Main.Text.Of") + " [$fileTotal]: $($file.FullName)"
						Start-Sleep -Milliseconds 5
						$form.Refresh()
						
						# Check if video is already in displayed video list
						$fileFound = $false
						foreach ($row in $dataGridView.Rows) {
							# Check all rows, get filename and path from DataGridView
							$dgvFileName = $row.Cells[$GRIDVIEWCOLUMNFILENAME].Value
							$dgvFilePath = $row.Cells[$GRIDVIEWCOLUMNFILEPATH].Value
							$dgvFileSize = [System.Int64]$row.Cells[$GRIDVIEWCOLUMNFILESIZE].Value
							
							# Check if path and filename exists
							if (($filePathForDB -eq $dgvFilePath) -and ($fileName -eq $dgvFileName) -and ($file.Length -eq $dgvFileSize)) {
								# Found
								$fileFound = $true
								break
							}
						}
						
						# Insert the series into the database if it wasn't found
						if (-not ($fileFound)) {
							# Extract series title and year from path
							$ty = Get-TitleAndYearFromPath -folderName $filePathForDB
							$videoInfo = $null
							
							# Analyze file
							$analyzed = $true
							try {
								$videoInfo = Get-VideoInfoFromFile -filePath $file.FullName -baseFolder $curSeriesPath
							}
							catch [CustomException] {
								$e = $_.Exception
								if($e.myMessage -eq 'Something terrible happened') {
									Write-Warning "$($e.myNumber) terrible things happened"
								}
								$analyzed = $false
							}
							
							if ($videoInfo -eq $null) {
								$analyzed = $false
							}
							
							if ($analyzed) {
								# Check if the of of the last file is the same as the current
								$seriesInfo = $null
								$seriesID = [System.Int64]0
								$seriesFoundinDB = $false
								
								# No, check if series is already in database
								$seriesPath = $filePathOnly.Trim("\")
								$seriesInfo = Get-SeriesFromDB -seriesPath $seriesPath -id $ty.tmdbId
								if( $seriesInfo.ID -eq $null) {
									$seriesInfo.ID = [System.Int64]0
									$seriesID = [System.Int64]0
								} else {
									# Series already in database
									$seriesFoundinDB = $true
									$seriesID = [System.Int64]$seriesInfo.id
								}
								if( $seriesID -eq 0) {
									# Query information from themoviedb.org
									if (-not ([string]::IsNullOrEmpty($config.ApiKey))) {
										# No, get series information from TMDB
										$seriesInfo = Get-SeriesInfoFromTMDB -title $ty.title -year $ty.year -tmdbId $ty.tmdbId
										if ($seriesInfo -ne $null) {
											$seriesID = $seriesInfo.id
										}
									}
								}
								
								if ((-not ([string]::IsNullOrEmpty($config.ApiKey))) -and ($seriesID -ne 0)) {
									$videoId = 0
									
									# Query information from themoviedb.org
									# Extract series title, season and episode from filename
									# Expected format:
									# sxxxeyyy title.suffix
									$details = Get-SeriesDetailsFromFilename -fileName $fileName
									
									if (($details.season -lt 0) -or ($details.season -eq $null) -or ($details.episode -eq 0)-or ($details.episode -eq $null)) {
										Write-Warning "ERROR: Season or episode could not be extracted"
									} else {
										# Get information for given episode from TMDB
										$episodeInfo = Get-EpisodesInfoFromTMDB -seriesID $seriesID -season $details.season -episode $details.episode
										
										# If the movie was found query movie details
										if ((![string]::IsNullOrEmpty($episodeInfo)) -and ($episodeInfo.id -gt 0)) {
											# Check if part information exists and append part number to season and episode
											$part = ""
											if( $details.part -ne $null) {
												$part = "p" + ($details.part).ToString('00')
											}
											$title = $seriesInfo.name + ": s" + ($episodeInfo.season_number).ToString('00') + "e" + ($episodeInfo.episode_number).ToString('000') + $part + " " + $episodeInfo.name
											
											# Insert or update in database
											$videoId = Upsert-VideoInfo -videoInfo $videoInfo -fileSize $file.Length -tmdbId $episodeInfo.id -title $title -vote $episodeInfo.vote_average -videoType 'S' -isAdult $seriesInfo.adult

											# Insert data into database
											$seriesDetailsObject = [PSCustomObject]@{
												TMDBId		= [System.Int64]$episodeInfo.id
												VideoType   = "S"
												Title		= $episodeInfo.name
												Overview	= $episodeInfo.overview
												ReleaseDate = $episodeInfo.air_date
											}
											
											# Insert or update video details
											Upsert-VideoDetails -videoId $videoId -videoDetails $seriesDetailsObject
											
											if (-not( $seriesInfo -eq $null)) {
												# Insert or update series details
												Upsert-VideoBelongsTo -videoListId $videoId -videoType "S" -tmdbBelongsToId $seriesID -name $seriesInfo.name -overview $seriesInfo.overview
												
												if (-not($seriesFoundinDB)) {
													Upsert-Genres -id $videoId -genres $seriesInfo.genres
												}
											}
											Upsert-Actors -id $videoId -actors $episodeInfo.credits.cast
											Upsert-Actors -id $videoId -actors $episodeInfo.credits.guest_stars
										} else {
											# No TMDB API key present so inset or update basic information only
											$title = $ty.title + ": " + [io.path]::GetFileNameWithoutExtension($fileName)
											$videoId = Upsert-VideoInfo -videoInfo $videoInfo -fileSize $file.Length -title $title -tmdbId 0 -vote "0.0" -videoType 'S' -isAdult "false"
										}
									}
								} else {
									# No TMDB API key present so insert or update basic information only
									$title = $ty.title + ": " + $fileName
									$videoId = Upsert-VideoInfo -videoInfo $videoInfo -fileSize $file.Length -title $title -tmdbId 0 -vote "0.0" -videoType 'S' -isAdult "false"
								}
								
								# Load row
#								if ($videoId -ne 0) {
#									Load-DataRow $videoId
#									Scroll-To-Last-Element
#								}
								
								$lastSeriesFolder = $filePathOnly
								$lastSeriesID = $seriesID
							}
						}
					}
				}
			}
		}
	}
}

# Scan folders recursivly and analyze video files
function Analyze-Videos {
	$global:doubleFilesView = $false
	
	# Mesure time of execution
	$elapsedTime = Measure-Command {
		# Save status of search
		$filteredSearch = $global:lastSearchFiltered
		$global:lastSearchFiltered = $false
		
		# Load database entries
		Load-Data
		
		# First check for current database entries if the corresponding files still exists
		Mark-NonexistentFiles
		
		# Scan folders and anlayze videos
		Analyze-Movies
		Analyze-Series
		
		# Restore status of search
		$global:lastSearchFiltered = $filteredSearch
		
		# Re-load the database entries
		Load-Data
		
		# Mark non-existing files
		Mark-Red
	}
	
	# Format used time hh:mm:ss
	$hours = [math]::Floor($elapsedTime.TotalSeconds / 3600)
	$minutes = [math]::Floor(($elapsedTime.TotalSeconds % 3600) / 60)
	$seconds = [math]::Floor($elapsedTime.TotalSeconds % 60)
	$formattedTime = "{0:D2}:{1:D2}:{2:D2}" -f [System.Int32]$hours, [System.Int32]$minutes, [System.Int32]$seconds
	
	# Set progress text to done and time needed
	$progressTextBox.Text = (Get-Translation -language $config.language -key "Main.Text.Done") + " [" +$formattedTime + "]"
	Start-Sleep -Milliseconds 5
	$form.Refresh()
}


# Analyze video file with ffprobe
function Get-VideoInfoFromFile {
	param (
		[String]$filePath,
		[String]$baseFolder
	)
	cursorWait
	
	# Get information from file system
	# Get video folder path length
	$folderPathLength = $baseFolder.Length
	# Check for trailing backslash
	if ($baseFolder.SubString($baseFolder.Length - 1, 1) -eq "\" ) {
		$folderPathLength--
	}
	
	# Extract file name from current file path
	$fileName = [System.IO.Path]::GetFileName($filePath)
	# Extract file path only
	$filePathOnly = [System.IO.Path]::GetDirectoryName($filePath)
	# Remove video path
	$filePathOnly = $filePathOnly.SubString($folderPathLength)
	$filePathOnly = $filePathOnly.Trim("\")
	
	# Prepare return object
	$details = [PSCustomObject]@{
		FileName	  = $fileName
		FilePath	  = $filePathOnly
		Resolution	= "-"
		VideoCodec	= "-"
		AudioTracks   = 0
		AudioChannels = "-"
		AudioLayouts  = "-"
		AudioLanguages = "-"
		Duration	  = "0:00:00"
		FileExists	= 1
	}	
	
	# Get information from ffprobe
	try {
		# Call ffprobe
		$ffprobeOutput = $null
		$ffprobeOutput = & $config.FFprobePath -v quiet -print_format json -show_format -show_streams $filePath | ConvertFrom-Json
		$ffprobeExitCode = $LASTEXITCODE
		
		# Check return code of ffprobe execution
		if ($ffprobeExitCode -ne 0) {
			$errorMsg = Get-Translation -language $config.Language -key "Error.FFProbe.ReturnCodeNotZero"
			throw [CustomException]::new("FFPROBE-ERROR", $errorMsg + "$ffprobeExitCode", $ERRORFFPROBERCNOTZERO)
		}

		# Check if ffprobe has detected any stream
		if ($ffprobeOutput.PSObject.Properties.Match("streams") -eq $null) {
			$errorMsg = Get-Translation -language $config.Language -key "Error.FFProbe.NoSreams"
			throw [CustomException]::new("FFPROBE-ERROR", $errorMsg, $ERRORFFPROBENOSTREAMS)
		}
		
		# Get information from ffprobe
		# Check if a video stream was detected
		$videoStream = $ffprobeOutput.streams | Where-Object { $_.codec_type -eq "video" }
		if( $videoStream -eq $null ) {
			$errorMsg = Get-Translation -language $config.Language -key "Error.FFProbe.NoVideoSream"
			throw [CustomException]::new("FFPROBE-ERROR", $errorMsg, $ERRORFFPROBENOVIDEOSTREAM)
		}
		$details.Resolution	= "$($videoStream.width)x$($videoStream.height)"
		$details.VideoCodec = [string]::Join(", ", $videoStream.codec_name)
		
		# Check if an audio stream was detected
		$audioStreams = $ffprobeOutput.streams | Where-Object { $_.codec_type -eq "audio" }
		if( $audioStreams -ne $null ) {
			$details.AudioTracks   = $audioStreams.Count
			$details.AudioChannels = ($audioStreams | ForEach-Object { $_.channels }) -join ", "
			$details.AudioLayouts  = ($audioStreams | ForEach-Object { $_.channel_layout }) -join ", "
			# Get anguage information if available
			try {
				$details.AudioLanguages = ($audioStreams | ForEach-Object { $_.tags.language }) -join ", "
			} catch {
				# No language information found
				$errorMsg = Get-Translation -language $config.Language -key "Error.FFProbe.NoLanguage"
				$pathMsg = Get-Translation -language $config.Language -key "Error.FFProbe.Path"
				$fileMsg = Get-Translation -language $config.Language -key "Error.FFProbe.File"
				Write-Host $errorMsg
				Write-Host $pathMsg $filePath
				Write-Host $fileMsg $fileName
			}
		} else {
			# No audio stream found, stop analyzing
			$errorMsg = Get-Translation -language $config.Language -key "Error.FFProbe.NoAudioStream"
			$pathMsg = Get-Translation -language $config.Language -key "Error.FFProbe.Path"
			$fileMsg = Get-Translation -language $config.Language -key "Error.FFProbe.File"
			Write-Host $errorMsg
			Write-Host $pathMsg $filePath
			Write-Host $fileMsg $fileName
			throw [CustomException]::new("FFPROBE-ERROR", "Error no audio stream", 4)
		}
		
		$durationSeconds = [math]::Round($ffprobeOutput.format.duration, 0)
		$details.Duration = [TimeSpan]::FromSeconds($durationSeconds).ToString("hh\:mm\:ss")
	}
	catch {
		# Unknown error occured
		$errorMsg = Get-Translation -language $config.Language -key "Error.FFProbe.UnknownError"
		$pathMsg = Get-Translation -language $config.Language -key "Error.FFProbe.Path"
		$fileMsg = Get-Translation -language $config.Language -key "Error.FFProbe.File"
		Write-Host $errorMsg
		Write-Host $pathMsg $filePath
		Write-Host $fileMsg $fileName
		throw [CustomException]::new("FFPROBE-ERROR", $errorMsg, $ERRORFFPOBEGUNKNOWN)
	}
	
	cursorDefault
	
	return $details
}


######################################################
# Web query functions for themovedb.org
# Get data from web using Invoke-RestMethod
function Get-URL {
	param (
		[String]$Url
	)
	
	# Initialize return value
	$response = $null
	
	try {
		# Get data from web
		$response = Invoke-RestMethod -Uri ($Url + "&" + $global:tmdbAdult) -Method Get
	}
	catch {
		# If an error occured just output error message
		# ToDo Translation
		Write-Warning "URL:" $Url
		Write-Warning $_.Exception.GetType().FullName
	}
	
	return $response
}

# Check if the API key is valid
function Check-APIKey {
	# Beispiel-URI für die Validierung des API-Keys
	$uri = "https://api.themoviedb.org/3/authentication/token/new?api_key="+ $config.ApiKey
	
	# Assume the key is not valid
	$global:tmdbAPIKeyValid = $false
	
	# Check the API key
	$response = Get-URL -Url $uri
	
	# Check response if key is valid
	if (-not($response -eq $null)) {
		if ($response.success -eq $true) {
			$global:tmdbAPIKeyValid = $true
		}
	}
	
	$openTMDBButton.Enabled = $global:tmdbAPIKeyValid
	
	return $global:tmdbAPIKeyValid
}

# Get information about the movie
function Get-MovieInfoFromTMDB {
	param (
		[String]$fileName
	)
	
	$videoData = Get-TitleAndYearFromFilename -fileName $fileName
	
	if (-not ([string]::IsNullOrEmpty($videoData.title))) {
		# Make title URL compatible
		$titleEncoded = [System.Web.HttpUtility]::UrlEncode($videoData.title)
		
		if ($videoData.tmdbID -gt 0) {
			# The filename contains a TMDB id, so try to get data directly
			$searchUrl = "https://api.themoviedb.org/3/movie/" + $videoData.tmdbID + "?api_key=" + $config.ApiKey + "&language=" + $config.language
			
			# Lookup the video at TMDB
			$response = Get-URL -Url $searchUrl
			
			return $response
		} elseif ($videoData.year -eq "") {
			$searchUrl = "https://api.themoviedb.org/3/search/movie?api_key=" + $config.ApiKey + "&query=" + $titleEncoded + "&language=" + $config.language
		} else {
			$searchUrl = "https://api.themoviedb.org/3/search/movie?api_key=" + $config.ApiKey + "&query=" +$titleEncoded + "&year=" + $videoData.year + "&language=" + $config.language
		}
		
		# Lookup the video at TMDB
		$response = Get-URL -Url $searchUrl
		
		# Check if something was found
		if ($response -ne $null) {
			if ($response.total_results -eq 0 ) {
				return $null
			}
			
			# If more than one entry was found check if the right one can be determined
			if ($response.results -is [Array]) {
				$res = @()
				# Remove all special characters from title
				$titleToCheck = $videoData.title -replace "[^a-zA-Z0-9.-]"
				# Lowercase the string
				$titleToCheck = $titleToCheck.ToLower()
				foreach($result in $response.results) {
					$curTitle = $result.title -replace "[^a-zA-Z0-9.-]"
					$curTitle = $curTitle.ToLower()
					# Compare the two down striped strings
					if ($curTitle -eq $curTitle) {
						$res += $result
					}
				}
				
				if ($res -is [Array]) {
					return $response.results[0]
				} else {
					return $res
				}
			} else {
				return $response.results
			}
		} else {
			return $null
		}
	} else {
		# No title could be extracted from filename
		# ToDo write warning to console
		return $null
	}
}

# Get information about the movie by given TMDB Id
function Get-MovieInfoFromTMDBbyID {
	param (
		[System.Int64]$tmdbId
	)
	
	# A TMDB Id is given
	if ($tmdbId -gt 0) {
		$searchUrl = "https://api.themoviedb.org/3/movie/" + $tmdbID + "?append_to_response=credits&api_key=" + $config.ApiKey + "&language=" + $config.language
		
		# Lookup the video at TMDB
		$response = Get-URL -Url $searchUrl
		
		return $response
	}
	return $null
}

# Get information about the series
function Get-SeriesInfoFromTMDB {
	param (
		[String]$title,
		[System.Int32]$year,
		[System.Int64]$tmdbId
	)
	
	if ($tmdbId -gt 0) {
		# Make title URL compatible
		$titleEncoded = [System.Web.HttpUtility]::UrlEncode($title)
		$searchUrl = "https://api.themoviedb.org/3/tv/" + $tmdbId + "?api_key=" + $config.ApiKey + "&query=" + $titleEncoded + "&language=" + $config.language
		
		# Lookup the video at TMDB
		$response = Get-URL -Url $searchUrl
		
		return $response
	} else {
		if (-not ([string]::IsNullOrEmpty($title))) {
			# Make title URL compatible
			$titleEncoded = [System.Web.HttpUtility]::UrlEncode($title)
			
			if($year -eq 0) {
				$searchUrl = "https://api.themoviedb.org/3/search/tv?api_key=" + $config.ApiKey + "&query=" + $titleEncoded + "&language=" + $config.language
			} else {
				$searchUrl = "https://api.themoviedb.org/3/search/tv?api_key=" + $config.ApiKey + "&query=" +$titleEncoded + "&year=" + $year + "&language=" + $config.language
			}
			
			# Lookup the video at TMDB
			$response = Get-URL -Url $searchUrl
			
			$id = 0
			# If more than one entry was found check if the right one can be determined
			if ($response.results -is [Array]) {
				$res = @()
				# Remove all special characters from title
				$titleToCheck = $title -replace "[^a-zA-Z0-9.]"
				# Lowercase the string
				$titleToCheck = $titleToCheck.ToLower()
				foreach($result in $response.results) {
					$curTitle = $result.name -replace "[^a-zA-Z0-9.]"
					$curTitle = $curTitle.ToLower()
					$resYear = $result.first_air_date
					if ($resYear -ne $null) {
						if ($resYear.Length -ge 4) {
							$resYear = $resYear.SubString(0, 4)
						} else{
							$resYear = "0000"
						}
					} else {
						$resYear = "0000"
					}
					
					# Compare the two down striped strings
					if (($titleToCheck -eq $curTitle) -and ($result.first_air_date.SubString(0, 4) -eq $year)) {
						$res += $result
					}
				}
				
				if ($res -is [Array]) {
					if (($res.Count) -gt 1) {
						# ToDo write warning that more than one entry has been found
#write-host "More than one entry found!"
#write-host $searchUrl
#write-host $title
#write-host "res:"
#write-host $res
#write-host "repsonse:"
#write-host $response
#$resFile = ".\" + $title + ".json"
#$result | ConvertTo-Json -Depth 3 | Set-Content -Path $resFile
					
					}
					if ($res -ne $null) {
						$id = [System.Int64]$response.results[0].id
					} else {
						# $res is an empty array
						$id = 0
					}
				} else {
					$id = [System.Int64]$res.id
				}
			} else {
				$id = [System.Int64]$response.results.id
			}
			
			if ($id -ne 0) {
				$searchUrl = "https://api.themoviedb.org/3/tv/" + $id + "?api_key=" + $config.ApiKey + "&query=" + $titleEncoded + "&language=" + $config.language
				
				# Lookup the video at TMDB
				$response = Get-URL -Url $searchUrl
				
				return $response
			} else {
				return $null
			}
		} else {
			# No title could be extracted from filename
			# ToDo write warning
			return $null
		}
	}
}

# Get information about the episode from TMDB
# ToDo check return code
function Get-EpisodesInfoFromTMDB {
	param (
		[System.Int64]$seriesIDinTMDB,
		[System.Int32]$season,
		[System.Int32]$episode
	)
	
	$searchUrl = "https://api.themoviedb.org/3/tv/" + $seriesIDinTMDB + "/season/" + $season + "/episode/" + $episode + "?append_to_response=credits&api_key=" + $config.ApiKey + "&language=" + $config.language
	
	try {
		# Lookup the video at TMDB
		$response = Get-URL -Url $searchUrl
		
		return $response
	}
	catch [Microsoft.PowerShell.Commands.HttpResponseException] {
		$StatusCode = [Int]$_.Exception.Response.StatusCode
		
		if ($_.Exception.Response.StatusCode -eq 404) {
Write-Host "Error: Page not found"
		} else {
Write-Host "An error occured: $_"
		}
		
		if (-not($Error[0].success)) {
			if ($Error[0].status_code -eq 34) {
				# "The resource you requested could not be found."
				return $null
			} else {
Write-Host "Unknown error: "
Write-Host $Error[0]
Write-Host $Error[0].Exception.GetType().FullName
			}
		} else {
Write-Host "Unknown error: "
Write-Host $Error[0]
Write-Host $Error[0].Exception.GetType().FullName
		}
		return $null
	}
	catch {
		Write-Host "function Get-EpisodesInfoFromTMDB: ERROR: Unknown error occured:"
		Write-Host "URL:" $searchUrl
		Write-Host $Error[0].Exception.GetType().FullName
#		throw $Error
		return $null
	}
}

# Query movie details from TMDB for given movie ID
function Get-MovieDetailsFromTMDB {
	param (
		[System.Int64]$movieId
	)
	
	$response = $null
	if ($movieId -ne 0) {
		$searchUrl = "https://api.themoviedb.org/3/movie/" + $movieId + "?api_key=" + $config.ApiKey + "&append_to_response=credits&language=" + $config.language
		try {
			$response = Get-URL -Url $searchUrl
		} catch {
			$Error[0].Exception.GetType().FullName
		}
		
		return $response
	} else {
		return $null
	}
}

# Query collection details from TMDB for given collection ID
function Get-CollectionDetailsFromTMDB {
	param (
		[System.Int64]$collectionId
	)
	
	if ($collectionId -ne 0) {
		$searchUrl = "https://api.themoviedb.org/3/collection/" + $collectionId + "?api_key=" + $config.ApiKey + "&append_to_response=credits&language=" + $config.language
		try {
			$response = Get-URL -Url $searchUrl
		} catch {
			$Error[0].Exception.GetType().FullName
		}
		
		return $response
	} else {
		return $Null
	}
}

# Get all episode numbers for the given series from TMDB
function Get-AllSeriesEpisodesFromTMDB {
	param (
		[System.Int64]$seriesID
	)
	
	# Create list for all parts of the series
	$episodes = @()
	
	# Get information about the series
	$seasonsUrl = "https://api.themoviedb.org/3/tv/" + $seriesID + "?language=" + $config.language + "&api_key=" + $config.ApiKey
	$tvSeries = Get-URL -Url $seasonsUrl
	
	# Check if series has been found
	if (-not ([string]::IsNullOrEmpty($tvSeries))) {
		# Check if seasons exist
		if (-not ([string]::IsNullOrEmpty($tvSeries.seasons))) {
			# Loop through all seasons
			foreach ($season in $tvSeries.seasons) {
				# Get current season
				$seasonNumber = $season.season_number
				# Season 0 is used for extras, background information etc.
				# For the regular episodes this will be ignored
				if ($seasonNumber -gt 0) {
					# Create episode entry for all episodes in the current season
					for ($i = 1; $i -le $season.episode_count; $i++) {
						$episodes += @(@{ season = $seasonNumber; episode = $i })
					}
					# Sometimes double length episodes exists for a series. And sometimes they are
					# labeled as usual, sometimes they were sent as part 1 and part 2, having a
					# gap in the enumeration, like:
					# s01e01 Episode 1
					# s01e02 Episode 2, Part 1 and Part 2
					# s01e04 Episode 3
					# For this reason the given number from the TMDB tv series list is required and detected
					# in this way.
				}
			}
		}
	}
	
	return $episodes
}

# Get all episode numbers for the given series from TMDB
function Get-AllMoviesForCollectionFromTMDB {
	param (
		[System.Int64]$collectionID
	)
	
	# Create list for all parts of the series
	$collectionMovie = @()
	
	# Get information about the series
	$seasonsUrl = "https://api.themoviedb.org/3/collection/" + $collectionID + "?language=" + $config.language + "&api_key=" + $config.ApiKey
	$collection = Get-URL -Url $seasonsUrl
	
	# Check if parts exist
	if (-not ([string]::IsNullOrEmpty($collection))) {
		if (-not ([string]::IsNullOrEmpty($collection.parts))) {
			foreach ($part in $collection.parts) {
				# Get current season
				$collectionMovie += @(@{ id = [System.Int64]$part.id})
			}
		}
	}
	
	return $collectionMovie
}

######################################################

# Get the name of the calling function
# This will be used to check which is the calling function for debugging purposes
function Get-CallerName {
    # Get current PowerSgell stack
    $callStack = Get-PSCallStack
	
    # Check if the stack have enough entries
    if ($callStack.Count -gt 2) {
		# The second erntry is the calling function
        return $callStack[2].Command
    } else {
        return "No calling function."
    }
}

######################################################


# Adjusting the buttons font size depending on the button text
function Adjust-ButtonFontSize {
	param (
		[System.Windows.Forms.Button]$button
	)
	
	# Maximum width, substracting 6 pixels buffer for left and right border
	$maxWidth = $button.Width - 6
	# Start with font size 11
	$fontSize = 11
	$font = New-Object System.Drawing.Font($button.Font.Name, $fontSize)
	$textSize = [System.Windows.Forms.TextRenderer]::MeasureText($button.Text, $font)
	
	# Loop while real text width in pixels is larger than button width
	while ($textSize.Width -gt $maxWidth -and $fontSize -gt 5) {

		$fontSize--
		$font = New-Object System.Drawing.Font($button.Font.Name, $fontSize)
		$textSize = [System.Windows.Forms.TextRenderer]::MeasureText($button.Text, $font)
	}
	
	# Assign new generated font to button
	$button.Font = $font
}


######################################################
# Global variables
$global:movieDriveFound = $false
$global:seriesDriveFound = $false

$global:moviePath = $Null
$global:seriesPath = $Null

# Sort direction for file path
$global:sortDirection = [System.ComponentModel.ListSortDirection]::Ascending

# Was the last search with a filter?
$global:lastSearchFiltered = $false

# Are currently doube files views?
$global:doubleFilesView = $false

# Is API key set and valid?
$global:tmdbAPIKeyValid = $false

# Initialize event variables
$global:usbConnectEvent = $null
$global:usbDisconnectEvent = $null


######################################################
# Languages
# Path to language file
$languageFilePath = "$scriptDirectory\VDB-lang.json"
# Check if language file exists
if (-not (Test-Path $languageFilePath)) {
	Write-Host 'ERROR: language file "$languageFilePath" not found!'
	Exit
}
# Set en-US as default language
$defaultLanguage = "en-US"
# Load language file
$jsonData = Get-Content -Encoding UTF8 -Path $languageFilePath | ConvertFrom-Json
# Retreive translations from language file
$translations = $jsonData.translations
# Retrieve languages form language file
$languages = $jsonData.languages
# Order language into array by value which is the common real name
$sortedLanguages = $languages.PSObject.Properties | Sort-Object -Property Value
# Create array with sorted languages
$sortedLanguageValues = $sortedLanguages | ForEach-Object { $_.Value }


######################################################
# Configuration file
# Path to configuration file
$configFilePath = "$scriptDirectory\VDB-config.json"
# Configuration default values
$configDefaults = @{
	Language = $defaultLanguage
	MovieFolder = ""
	MovieVolumeLabel = ""
	UseDriveLabelForMovies = $false
	MovieFolderIsNetwork = $False
	SeriesFolder = ""
	SeriesVolumeLabel = ""
	UseDriveLabelForSeries = $false
	SeriesFolderIsNetwork = $false
	FFprobePath = ""
	ApiKey = ""
	GetAdultContent = $false
	MaxScans = 0
	Width = 1040
	Height = 600
	Top = 0
	Left = 0
	WindowState = [System.Windows.Forms.FormWindowState]::Normal
}

# Load configuration from file
$config = Load-Config

# Database
# Path to SQLite database
$dbPath = "$($pwd)\Video.db"

# Export files
$exportFileCSV  = "$($pwd)\export.csv"
$exportFileDoubleCSV = "$($pwd)\exportDouble.csv"


######################################################
# Create GUI
# Create form
$form = New-Object System.Windows.Forms.Form
$form.Size = New-Object System.Drawing.Size(1120, 600)
$form.MinimumSize = New-Object System.Drawing.Size(1120, 600)

# Create top panel
# used for buttons, search field, status bar
$topPanel = New-Object System.Windows.Forms.Panel
$topPanel.Dock = [System.Windows.Forms.DockStyle]::Top
$topPanel.Height = 80

# Button for opening configuration dialog
$buttonOpenConfig = New-Object System.Windows.Forms.Button
$buttonOpenConfig.Location = New-Object System.Drawing.Point(10,10)
$buttonOpenConfig.Size = New-Object System.Drawing.Size(70, 20)
$buttonOpenConfig.Add_Click({
	$ret = Show-ConfigDialog
	if($ret -eq [System.Windows.Forms.DialogResult]::OK) {
		Check-Configuration
		Set-MainText $config.Language
	}
})

$analyzeButton = New-Object System.Windows.Forms.Button
$analyzeButton.Location = New-Object System.Drawing.Point(90, 10)
$analyzeButton.Size = New-Object System.Drawing.Size(70, 20)
$analyzeButton.Enabled = $global:movieDriveFound -or $global:seriesDriveFound
$analyzeButton.Add_Click({
	cursorWait
	Analyze-Videos
	cursorDefault
})

$searchBox = New-Object System.Windows.Forms.TextBox
$searchBox.Location = New-Object System.Drawing.Point(170, 10)
$searchBox.Size = New-Object System.Drawing.Size(200, 20)
# Add event handler for keydown event
$searchBox.Add_KeyDown({
	param (
		[System.Object]$sender,
		[System.Windows.Forms.KeyEventArgs]$e
	)
	
	# Search for data if the enter key has been pressed
	if ($e.KeyCode -eq [System.Windows.Forms.Keys]::Return) {
		Load-Data $searchBox.Text
		
		# Mark non-existing files
		Mark-Red
	}
})

$searchButton = New-Object System.Windows.Forms.Button
$searchButton.Location = New-Object System.Drawing.Point(380, 10)
$searchButton.Size = New-Object System.Drawing.Size(70, 20)
$searchButton.Add_Click({
	# Search data
	Load-Data $searchBox.Text
	
	# Mark non-existing files
	Mark-Red
})

$duplicatesButton = New-Object System.Windows.Forms.Button
$duplicatesButton.Location = New-Object System.Drawing.Point(460, 10)
$duplicatesButton.Size = New-Object System.Drawing.Size(70, 20)
$duplicatesButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Standard
$duplicatesButton.Add_Click({
	Show-Doubles
})

# Button for opening a dialog to enter the ID from themoviedb.org directly
$openTMDBButton = New-Object System.Windows.Forms.Button
$openTMDBButton.Text = "TMDB"
$openTMDBButton.Location = New-Object System.Drawing.Point(540,10)
$openTMDBButton.Size = New-Object System.Drawing.Size(70, 20)
$openTMDBButton.Add_Click({
	# Only one row is allowed to be selected
	if ($dataGridView.SelectedRows.Count -eq 1) {
		# Get row number
		$selectedRow = $dataGridView.SelectedRows[0]
		Show-TMDBDialog -row $selectedRow
	}
})

# Button for removing an entry and, if possible, also the file
$deleteButton = New-Object System.Windows.Forms.Button
$deleteButton.Location = New-Object System.Drawing.Point(620, 10)
$deleteButton.Size = New-Object System.Drawing.Size(70, 20)
$deleteButton.Add_Click({
	# Get number of selected rows
	$count = $dataGridView.SelectedRows.Count
	# Check if at least one row is selected
	if ($count -gt 0) {
		# Messagebox header
		$headerText = Get-Translation -language $config.Language -key "Main.Text.AreYouSure"
		
		# Check if only one row is selected
		if ($count -eq 1) {
			# Get text for one row deletion confirmation
			$messageText = Get-Translation -language $config.Language -key "Main.Text.Delete1"
		} else {
			# Get text for mulitple enries deletion confirmation
			$messageText = Get-Translation -language $config.Language -key "Main.Text.Delete2a"
			$messageText += [string]$count
			$messageText += Get-Translation -language $config.Language -key "Main.Text.Delete2b"
		}
		
		# Safety query
		$result = [System.Windows.MessageBox]::Show($messageText, $headerText, [System.Windows.MessageBoxButton]::OKCancel)
		if ($result -eq [System.Windows.MessageBoxResult]::OK) {
			$deleteFiles = $false
			if ($global:movieDriveFound -or $global:seriesDriveFound) {
				# Get text for file deletion confirmation
				$messageText = Get-Translation -language $config.Language -key "Main.Text.DeleteFiles"
				$result = [System.Windows.MessageBox]::Show($messageText, $headerText, [System.Windows.MessageBoxButton]::YesNo)
				if ($result -eq [System.Windows.MessageBoxResult]::Yes) {
					$deleteFiles = $true
				}
			}
			
			for ($i = 0; $i -lt $count; $i++) {
				# Get row number
				$selectedRow = $dataGridView.SelectedRows[$i]
				# Get current video ID
				$videoIdCurrent = $selectedRow.Cells["ID"].Value
				# Get TMDB ID
				$tmdbIdCurrent = [System.Int64]$selectedRow.Cells[$GRIDVIEWCOLUMNTMDBID].Value
				# Get video type
				$videoTypeCurrent = $selectedRow.Cells[$GRIDVIEWCOLUMNVIDEOTYPE].Value
				# Get file existence
				$fileExistsCurrent = $selectedRow.Cells["FileExists"].Value
				
				# First remove all old data for the current ID from related tables:
				Delete-Genres -videoId $videoIdCurrent
				Delete-Actors -videoId $videoIdCurrent
				Delete-BelongsTo -videoId $videoIdCurrent
				Delete-VideoDetails -tmdbId $tmdbIdCurrent -videoType $videoTypeCurrent
				# Now delete the video
				Delete-Video -videoId $videoIdCurrent
				
				# Delete file if confirmed
				if ($deleteFiles) {
					# Check if file exists after last analyze
					if( $fileExistsCurrent -eq "1") {
						# Create full path and filename
						$fileName = ""
						if ($videoTypeCurrent -eq "M") {
							$fileName = $global:moviePath.GetPath()
						} else {
							$fileName = $global:seriesPath.GetPath()
						}
						$fileName += $selectedRow.Cells[$GRIDVIEWCOLUMNPATHANDFILENAME].Value
						if (Test-Path -LiteralPath $fileName) {
							# Move file to recycle bin
#							$res = Move-ToRecycleBin -Path $fileName
						}
					}
				}
			}
			
			Load-Data
			
			# Mark non-existing files
			Mark-Red
		}
	}
})

# Button for exporting the video list into a CSV file
$exportButton = New-Object System.Windows.Forms.Button
$exportButton.Location = New-Object System.Drawing.Point(700, 10)
$exportButton.Size = New-Object System.Drawing.Size(70, 20)
$exportButton.Add_Click({
	cursorWait
	
	# Columns to export
	$columnsToExport = @( $GRIDVIEWCOLUMNTITLE, $GRIDVIEWCOLUMNFILENAME, $GRIDVIEWCOLUMNFILEPATH,
		$GRIDVIEWCOLUMNBELONGSTO, $GRIDVIEWCOLUMNFILESIZE, $GRIDVIEWCOLUMNFILESIZEMB,
		$GRIDVIEWCOLUMNRESOLUTION, $GRIDVIEWCOLUMNVIDEOCODEC, $GRIDVIEWCOLUMNAUDIOTRACKS,
		$GRIDVIEWCOLUMNAUDIOCHANNELS, $GRIDVIEWCOLUMNAUDIOLAYOUTS, $GRIDVIEWCOLUMNAUDIOLANGUAGES,
		$GRIDVIEWCOLUMNDURATION, $GRIDVIEWCOLUMNVOTE, $GRIDVIEWCOLUMNVIDEOTYPE)
	
	# Create CSV data list
	$csvData = @()
	
	# Add header
	$header = ($columnsToExport | ForEach-Object { '"' + $dataGridView.Columns[$_].HeaderText + '"' }) -join ";"
	$csvData += $header
	
	# Add the lines from the data gride view to the CSV data
	foreach ($row in $dataGridView.Rows) {
		if($row.Cells[$GRIDVIEWCOLUMNFILEEXISTS].Value -eq "1") {
			$rowData = @()
			foreach ($column in $columnsToExport) {
				$rowData += '"' + $($row.Cells[$column].Value) + '"'
			}
			$csvData += ($rowData -join ";")
		}
	}
	
	# Write data into UTF8 encoded CSV file
	if($global:doubleFilesView -eq $false) {
		$csvData | Out-File -FilePath $exportFileCSV -Encoding UTF8BOM
	} else {
		$csvData | Out-File -FilePath $exportFileDoubleCSV -Encoding UTF8BOM
	}
	
	cursorDefault
})

# Button for renaming a file
$renameButton = New-Object System.Windows.Forms.Button
$renameButton.Location = New-Object System.Drawing.Point(780, 10)
$renameButton.Size = New-Object System.Drawing.Size(70, 20)
$renameButton.Enabled = $global:movieDriveFound -or $global:seriesDriveFound
$renameButton.Add_Click({
	# Check if only one row is selected
	if ($dataGridView.SelectedRows.Count -eq 1) {
		# Get invalid characters for a filename
		$invalidChars = [System.IO.Path]::GetInvalidFileNameChars()
		
		# Get row number
		$selectedRow = $dataGridView.SelectedRows[0]
		
		# Get current video ID
		$videoIdCurrent = $selectedRow.Cells["ID"].Value
		# Get filename and extract extension
		$fileNameOnly = $selectedRow.Cells[$GRIDVIEWCOLUMNFILENAME].value
		$ext = [System.IO.Path]::GetExtension($fileNameOnly)
		
		# Get TMDB Id
		$tmdbIdCurrent = $selectedRow.Cells[$GRIDVIEWCOLUMNTMDBID].Value
		# Get video type
		$videoTypeCurrent = $selectedRow.Cells[$GRIDVIEWCOLUMNVIDEOTYPE].Value
		
		# Create full path and new filename
		$fullPath = ""
		$fileNameNew = ""
		if ($videoTypeCurrent -eq "M") {
			# Video file is a movie
			$fullPath = $global:moviePath.GetPath()
			$fileNameNew = $titleTextBox.Text
			if ($yearTextBox.Text -ne "") {
				$fileNameNew += " (" + $yearTextBox.Text + ")"
			}
			$fileNameNew += " [TMDBID=" + $tmdbIdCurrent + "]" + $ext
		} else {
			# Video file is a series part
			$fullPath = $global:seriesPath.GetPath()
			$details = Get-SeriesDetailsFromFilename -fileName $fileNameOnly
			if (($details.season -gt 0) -and ($details.episode -gt 0)) {
				# Create new filename
				$fileNameNew = "s" + ($details.season).ToString('00') + "e" + ($details.episode).ToString('000') + " " + $titleTextBox.Text + $ext
			}
		}
		
		# Replace invalid characters in new filename with underscore
		$fileNameNew = $fileNameNew -replace "[$([regex]::Escape([string]::Join('', $invalidChars)))]", '_'
		
		$fullPath += "\" + $selectedRow.Cells[$GRIDVIEWCOLUMNPATHANDFILENAME].Value
		if (Test-Path -LiteralPath $fullPath) {
			# Create new form
			$renameForm = New-Object System.Windows.Forms.Form
			$renameForm.Text = Get-Translation -language $config.Language -key "Main.Button.Rename"
			$renameForm.Size = New-Object System.Drawing.Size(600, 180)
			
			# Create label for current filename
			$labelForCurrentFileName = New-Object System.Windows.Forms.Label
			$labelForCurrentFileName.Location = New-Object System.Drawing.Point(10, 10)
			$labelForCurrentFileName.Size = New-Object System.Drawing.Size(560, 20)
			$labelForCurrentFileName.Text = Get-Translation -language $config.Language -key "Rename.Current.FileName"
			$renameForm.Controls.Add($labelForCurrentFileName)
			
			# Create text box with current filename
			$textBoxCurrentFileName = New-Object System.Windows.Forms.TextBox
			$textBoxCurrentFileName.Location = New-Object System.Drawing.Point(10, 30)
			$textBoxCurrentFileName.Size = New-Object System.Drawing.Size(560, 20)
			$textBoxCurrentFileName.ReadOnly = $true
			$textBoxCurrentFileName.text = $fileNameOnly
			$renameForm.Controls.Add($textBoxCurrentFileName)
			
			# Create label for new filename
			$labelForCurrentFileName = New-Object System.Windows.Forms.Label
			$labelForCurrentFileName.Location = New-Object System.Drawing.Point(10, 50)
			$labelForCurrentFileName.Size = New-Object System.Drawing.Size(560, 20)
			$labelForCurrentFileName.Text = Get-Translation -language $config.Language -key "Rename.New.FileName"
			$renameForm.Controls.Add($labelForCurrentFileName)
			
			# Create text box for new filename
			$textBoxNewFileName = New-Object System.Windows.Forms.TextBox
			$textBoxNewFileName.Location = New-Object System.Drawing.Point(10, 70)
			$textBoxNewFileName.Size = New-Object System.Drawing.Size(560, 20)
			$textBoxNewFileName.Text = $fileNameNew
			$renameForm.Controls.Add($textBoxNewFileName)
			
			# Create OK button
			$okButton = New-Object System.Windows.Forms.Button
			$okButton.Location = New-Object System.Drawing.Point(10, 100)
			$okButton.Text = Get-Translation -language $config.Language -key "Button.OK"
			$okButton.Add_Click({
				# Replace invalid characters with underscore
				$validString = $textBoxNewFileName.Text -replace "[$([regex]::Escape([string]::Join('', $invalidChars)))]", '_'
				
				if ($textBoxNewFileName.Text -eq $validString) {
					$renameForm.DialogResult = [System.Windows.Forms.DialogResult]::OK
					$renameForm.Close()
				} else {
					# Update text box with fixed characters
					$validString = $validString -replace "_{2,}", "_"
					$validString = $validString -replace " {2,}", " "
					$textBoxNewFileName.Text = $validString
				}
			})
			$renameForm.Controls.Add($okButton)
			
			# Create cancel button
			$cancelButton = New-Object System.Windows.Forms.Button
			$cancelButton.Text = Get-Translation -language $config.Language -key "Button.Cancel"
			$cancelButton.Location = New-Object System.Drawing.Point(100, 100)
			$cancelButton.Add_Click({
				$renameForm.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
				$renameForm.Close()
			})
			$renameForm.Controls.Add($cancelButton)
			
			# Show the form
			$result = $renameForm.ShowDialog()
			
			# Check the result
			if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
				if (($textBoxNewFileName.Text -ne "") -and ($textBoxNewFileName.Text -ne $fileNameOnly)) {
					# Rename file
					Rename-Item -LiteralPath $fullPath -NewName $textBoxNewFileName.Text
					
					# Update filename and FileExists in database
					Update-FilenameVideoInfo -videoId $videoIdCurrent -filename $textBoxNewFileName.Text
					
					# Update grid view
					$selectedRow.Cells[$GRIDVIEWCOLUMNFILENAME].value = $textBoxNewFileName.Text
					$fpNew = $selectedRow.Cells[$GRIDVIEWCOLUMNPATHANDFILENAME].Value
					$fpNew = Split-Path -Path $fpNew -Parent
					$fpNew = $fpNew.TrimEnd("\") + "\" + $textBoxNewFileName.Text
					$selectedRow.Cells[$GRIDVIEWCOLUMNPATHANDFILENAME].Value = $fpNew
				}
			}
		}
	}
})

# Button for moving a file
$moveButton = New-Object System.Windows.Forms.Button
$moveButton.Location = New-Object System.Drawing.Point(860, 10)
$moveButton.Size = New-Object System.Drawing.Size(70, 20)
$moveButton.Enabled = $global:movieDriveFound -or $global:seriesDriveFound

# Button for re-scanning a file
$rescanButton = New-Object System.Windows.Forms.Button
$rescanButton.Location = New-Object System.Drawing.Point(940, 10)
$rescanButton.Size = New-Object System.Drawing.Size(70, 20)
$rescanButton.Enabled = $global:movieDriveFound -or $global:seriesDriveFound
$rescanButton.Add_Click({
	# Create culture instance fpr the current culture
	$culture = [System.Globalization.CultureInfo]::GetCultureInfo($config.Language)
	
	$count = $dataGridView.SelectedRows.Count
	for ($i = 0; $i -lt $count; $i++) {
		# Get row number
		$selectedRow = $dataGridView.SelectedRows[$i]
		
		# Get current video ID
		$videoId = $selectedRow.Cells["ID"].Value
		# Get filename and extract extension
		$fileNameOnly = $selectedRow.Cells[$GRIDVIEWCOLUMNFILENAME].value
		$filePathOnly = $selectedRow.Cells[$GRIDVIEWCOLUMNFILEPATH].value
		$filePath = $filePathOnly + "\" + $fileNameOnly
		
		# Get title
		$title = $selectedRow.Cells[$GRIDVIEWCOLUMNTITLE].Value
		# Get TMDB Id
		$tmdbId = [System.Int64]$selectedRow.Cells[$GRIDVIEWCOLUMNTMDBID].Value
		# Get vote as string and convert it into a number
		$voteStr = [string]$selectedRow.Cells[$GRIDVIEWCOLUMNVOTE].Value
		$vote = [double]::Parse($voteStr, $culture)
		
		# Get video type
		$videoType = $selectedRow.Cells[$GRIDVIEWCOLUMNVIDEOTYPE].Value
		
		# Create full path and new filename
		$baseFolder = ""
		if ($videoType -eq "M") {
			# Video file is a movie
			$baseFolder = $global:moviePath.GetPath()
		} else {
			# Video file is a series part
			$baseFolder = $global:seriesPath.GetPath()
		}
		
		# Get file
		$fileFullPath = $baseFolder + "\" + $filePath
		$file = Get-Item -LiteralPath $fileFullPath
		
		if (Test-Path -LiteralPath $fileFullPath) {
			# Analyze file
			$analyzed = $true
			$videoInfo = $null
			try {
				$videoInfo = Get-VideoInfoFromFile -filePath $fileFullPath -baseFolder $baseFolder
			}
			catch [CustomException] {
				$analyzed = $false		
				# Error number should only be one of:
				# ERRORFFPOBEGUNKNOWN
				# ERRORFFPROBERCNOTZERO
				# ERRORFFPROBENOSTREAMS
				# ERRORFFPROBENOVIDEOSTREAM
				$e = $_.Exception
				
				$outString = Get-Translation -language $config.Language -key "File"
				$outString += ": '" + $file.FullName + "'"
				write-host $outString
				
				$outString = Get-Translation -language $config.Language -key "Error"
				$outString += ": " + $e.myMessage
				write-host $outString
			} catch {
				$analyzed = $false						
				write-host $Error[0].Exception.GetType().FullName
			}
			
			if ($analyzed) {
				# Insert or update in database
				$videoId = Upsert-VideoInfo -videoInfo $videoInfo -fileSize $file.Length -tmdbId $tmdbId -title $title -vote $vote -videoType $videoType -isAdult "false"
				
				# Update grid view
				$selectedRow.Cells[$GRIDVIEWCOLUMNFILESIZE].value = $file.Length
				$selectedRow.Cells[$GRIDVIEWCOLUMNFILESIZEMB].value = [double][math]::Round($file.Length / 1MB, 2)
				$selectedRow.Cells[$GRIDVIEWCOLUMNRESOLUTION].value = $videoInfo.Resolution
				$selectedRow.Cells[$GRIDVIEWCOLUMNVIDEOCODEC].value = $videoInfo.VideoCodec
				$selectedRow.Cells[$GRIDVIEWCOLUMNAUDIOTRACKS].value = $videoInfo.AudioTracks
				$selectedRow.Cells[$GRIDVIEWCOLUMNAUDIOCHANNELS].value = $videoInfo.AudioChannels
				$selectedRow.Cells[$GRIDVIEWCOLUMNAUDIOLAYOUTS].value = $videoInfo.AudioLayouts
				$selectedRow.Cells[$GRIDVIEWCOLUMNAUDIOLANGUAGES].value = $videoInfo.AudioLanguages
				$selectedRow.Cells[$GRIDVIEWCOLUMNDURATION].value = $videoInfo.Duration
			}
		}
	}
})


# Button for filling up missing series or movie parts
$fillUpButton = New-Object System.Windows.Forms.Button
$fillUpButton.Location = New-Object System.Drawing.Point(1020, 10)
$fillUpButton.Size = New-Object System.Drawing.Size(70, 20)
#$fillUpButton.Enabled = $true # ToDo check for valid TMDB API key
$fillUpButton.Add_Click({
	cursorWait
	
	# Get invalid characters for a filename
	$invalidChars = [System.IO.Path]::GetInvalidFileNameChars()
	
	$count = $dataGridView.SelectedRows.Count
	for ($i = 0; $i -lt $count; $i++) {
		# Get number of selected row from DataGridView
		$selectedRow = $dataGridView.SelectedRows[$i]
		
		# Get current video ID
		$videoId = $selectedRow.Cells[$GRIDVIEWCOLUMNID].Value
		# Get file path
		$filePathOnly = $selectedRow.Cells[$GRIDVIEWCOLUMNFILEPATH].Value
		# Get where the current video belongs to
		$belongsTo = $selectedRow.Cells[$GRIDVIEWCOLUMNBELONGSTO].Value
		# Get video type
		$videoType = $selectedRow.Cells[$GRIDVIEWCOLUMNVIDEOTYPE].Value
		
		# Get the series or collection ID
		$belongsToId = Get-BelongsToId -videoId $videoId
		if ($belongsToId -ne $null) {
			if ($belongsToId -gt 0) {
				# Get the TMDB id of the series or collection
				$tmdbBelongsTo = Get-TMDBBelongsTo -belongsToId $belongsToId
				if ($tmdbBelongsTo.TMDBId -ne $null) {
					if ($videoType -eq "S") {
						# Process all series episodes
						$allEpisodes = Get-AllSeriesEpisodesFromTMDB -seriesID $tmdbBelongsTo.TMDBId
						$existingEpisodes = Get-AllExistingEpisodesForID -belongsToId $belongsToId
						
						# Remove existing episodes from all episodes
						foreach ($existingEpisode in $existingEpisodes) {
							$allEpisodes = $allEpisodes | Where-Object { 
								!($_.season -eq $existingEpisode.season -and $_.episode -eq $existingEpisode.episode) 
							}
						}
						
						$seriesInfo = Get-SeriesInfoFromTMDB -title "" -year 0 -tmdbId $tmdbBelongsTo.TMDBId
						
						# Add all missing episodes
						$allEpisodes | ForEach-Object {
							# Get information for given episode from TMDB
							$episodeInfo = Get-EpisodesInfoFromTMDB -seriesID $tmdbBelongsTo.TMDBId -season $_.season -episode $_.episode
							
							# Create filename
							$fileName = "s" + ($_.season).ToString('00') + "e" + ($_.episode).ToString('000') + " " + $episodeInfo.name + ".mp4"
							# Replace invalid characters in new filename with underscore
							$fileName = $fileName -replace "[$([regex]::Escape([string]::Join('', $invalidChars)))]", '_'
							
							# Create title
							$title = $tmdbBelongsTo.BelongsToName + ": s" + ($_.season).ToString('00') + "e" + ($_.episode).ToString('000') + " " + $episodeInfo.name
							
							$duration = "0:00:00"
							if ($episodeInfo.runtime -ne $null) {
								# Create dummy duration time
								$timeSpan = [TimeSpan]::FromMinutes($episodeInfo.runtime)
								# Format timespan into string of format "h:mm:ss"
								$duration = $timeSpan.ToString("h\:mm\:ss")
							}
							
							# Create details for inserting into database
							$details = [PSCustomObject]@{
								FileName       = $fileName
								FilePath       = $filePathOnly
								Resolution     = "-"
								VideoCodec     = "-"
								AudioTracks    = 0
								AudioChannels  = "-"
								AudioLayouts   = "-"
								AudioLanguages = "-"
								Duration       = $duration
								FileExists     = 0
							}
							
							# Create entry in database
							$videoId = Upsert-VideoInfo -videoInfo $details -fileSize 0 -tmdbId $episodeInfo.id -title $title -vote $episodeInfo.vote_average -videoType 'S' -isAdult $seriesInfo.adult
							
							# Insert data into database
							$seriesDetailsObject = [PSCustomObject]@{
								TMDBId		= $episodeInfo.id
								VideoType   = "S"
								Title		= $episodeInfo.name
								Overview	= $episodeInfo.overview
								ReleaseDate = $episodeInfo.air_date
							}
							
							# Insert or update video details
							Upsert-VideoDetails -videoId $videoId -videoDetails $seriesDetailsObject
							
							# Insert or update series details
							Upsert-VideoBelongsTo -videoListId $videoId -videoType "S" -tmdbBelongsToId $tmdbBelongsTo.TMDBId -name $tmdbBelongsTo.BelongsToName -overview $tmdbBelongsTo.overview
							
							Upsert-Genres -id $videoId -genres $seriesInfo.genres
							
							Upsert-Actors -id $videoId -actors $episodeInfo.credits.cast
							Upsert-Actors -id $videoId -actors $episodeInfo.credits.guest_stars
						}
					} elseif ($videoType -eq "M" ) {
						# Process all collection parts
						$allParts = Get-AllMoviesForCollectionFromTMDB -collectionID $tmdbBelongsTo.TMDBId
						$existingParts = Get-AllExistingPartsForID -belongsToId $belongsToId
						
						# Remove existing parts from all parts
						foreach ($existingPart in $existingParts) {
							$allParts = $allParts | Where-Object {
								!($_.id -eq $existingPart.tmdbId) 
							}
						}
						
						# Add all missing episodes
						$allParts | ForEach-Object {
							# Get information for given episode from TMDB
							$partInfo = Get-MovieInfoFromTMDBByID -tmdbId $_.id
							
							# Create filename
							$fileName = $partInfo.title + ".mp4"
							# Replace invalid characters in new filename with underscore
							$fileName = $fileName -replace "[$([regex]::Escape([string]::Join('', $invalidChars)))]", '_'
							
							$duration = "0:00:00"
							if ($partInfo.runtime -ne $null) {
								# Create dummy duration time
								$timeSpan = [TimeSpan]::FromMinutes($partInfo.runtime)
								# Format timespan into string of format "h:mm:ss"
								$duration = $timeSpan.ToString("h\:mm\:ss")
							}
							
							# Create details for inserting into database
							$details = [PSCustomObject]@{
								FileName       = $fileName
								FilePath       = $filePathOnly
								Resolution	   = "-"
								VideoCodec	   = "-"
								AudioTracks    = 0
								AudioChannels  = "-"
								AudioLayouts   = "-"
								AudioLanguages = "-"
								Duration       = $duration
								FileExists     = 0
							}
							
							# Create entry in database
							$videoId = Upsert-VideoInfo -videoInfo $details -fileSize 0 -tmdbId $partInfo.id -title $partInfo.title -vote $partInfo.vote_average -videoType 'M' -isAdult $partInfo.adult
							
							# Insert data into database
							$partDetailsObject = [PSCustomObject]@{
								TMDBId      = $partInfo.id
								VideoType   = "M"
								Title       = $partInfo.title
								Overview    = $partInfo.overview
								ReleaseDate = $partInfo.release_date
							}
							
							# Insert or update video details
							Upsert-VideoDetails -videoId $videoId -videoDetails $partDetailsObject
							
							# Insert or update series details
							Upsert-VideoBelongsTo -videoListId $videoId -videoType "M" -tmdbBelongsToId $tmdbBelongsTo.TMDBId -name $tmdbBelongsTo.BelongsToName -overview $tmdbBelongsTo.overview
							
							Upsert-Genres -id $videoId -genres $partInfo.genres
							
							Upsert-Actors -id $videoId -actors $partInfo.credits.cast
						}
					}
				}
			}
		}
	}
	
	Load-Data

	# Mark non-existing files
	Mark-Red

	cursorDefault
})


# Textbox for displaying the current file path while analyzine video files
$progressTextBox = New-Object System.Windows.Forms.TextBox
$progressTextBox.Location = New-Object System.Drawing.Point(10, 40)
$progressTextBox.Width = $topPanel.Width - 20
$progressTextBox.Height = 20
$progressTextBox.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$progressTextBox.ReadOnly = $true

$topPanel.Controls.AddRange(
	@($buttonOpenConfig, $analyzeButton, $searchBox, $searchButton, $duplicatesButton,
	$openTMDBButton, $deleteButton, $exportButton, $renameButton, $moveButton, $rescanButton, $fillUpButton, $progressTextBox)
)


# Create mid panel
# Used for video files table
$videoPanel = New-Object System.Windows.Forms.Panel
$videoPanel.Dock = [System.Windows.Forms.DockStyle]::Fill

# Data grid view for displaying video information
$dataGridView = New-Object System.Windows.Forms.DataGridView
$dataGridView.Location = New-Object System.Drawing.Point(10, 80)
$dataGridView.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
$dataGridView.ReadOnly = $true
$dataGridView.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
$dataGridView.AllowUserToAddRows = $false
$dataGridView.AllowUserToDeleteRows = $false

# Add event handler if one row is selected
$dataGridView.add_SelectionChanged({
	param($sender, $e)
	# Is at least one row selected?
	if ($dataGridView.SelectedRows.Count -gt 0) {
		# Get number of first selected row
		$selectedRow = $dataGridView.SelectedRows[0]
		
		# Get data from table
		$id = $selectedRow.Cells["Id"].Value
		$tmdbId = $selectedRow.Cells["TMDBId"].Value
		$videoType = $selectedRow.Cells["Videotype"].Value
		if (($tmdbId -ne "") -And ($tmdbId -ne "0")) {
			# TMDB is available so get video details
			$details = Get-VideoDetails -tmdbId $tmdbId -videoType $videoType

			# Get overview text if the video belongs to a series or collection
			$belongsToOverview = Get-OverviewForVideoBelongsto -videoId $id

			$titleTextBox.Text = $details.title
			$yearTextBox.Text = $details.releaseDate

			# Overview text
			if ([string]::IsNullOrEmpty($belongsToOverview)) {
				$overviewTextBox.Text = $details.overview
			} else {
				$overviewTextBox.Text = $details.overview
				$nl = "`r`n"
				$overviewTextBox.AppendText($nl)
				$overviewTextBox.AppendText("==========")
				$overviewTextBox.AppendText($nl)
				$overviewTextBox.AppendText($belongsToOverview)
			}
			$overviewTextBox.SelectionStart = 0
			$overviewTextBox.SelectionLength = 0
			$overviewTextBox.ScrollToCaret()
			
			$genresTextBox.Text = Get-Genres -id $id
			
			$actorsTextBox.Text = Get-Actors -id $id
			$actorsTextBox.SelectionStart = 0
			$actorsTextBox.SelectionLength = 0
			$actorsTextBox.ScrollToCaret()
			# Check for entry in search textbox
			if ($searchBox.Text -ne "") {
				# Textbox not empty, search foll appearances in actors
				$position = 0
				$startPos = $position
				while ($position -ne -1) {
					$position = $actorsTextBox.Text.IndexOf($searchBox.Text, $startPos, [System.StringComparison]::OrdinalIgnoreCase)
					if ($position -ne -1) {
						$actorsTextBox.SelectionStart = $position
						$actorsTextBox.SelectionLength = $searchBox.Text.Length
						$actorsTextBox.SelectionColor = [System.Drawing.Color]::Red
						$actorsTextBox.ScrollToCaret()
						$startPos = $position + $searchBox.Text.Length
					}
				}
			}
		} else {
			# No ID present
			$titleTextBox.Text = ""
			$overviewTextBox.Text = ""
			$genresTextBox.Text = ""
			$actorsTextBox.Text = ""
			$yearTextBox.Text = ""
		}
	}
})

# Add event handler for clicking on grid view header
$dataGridView.add_ColumnHeaderMouseClick({
	param($sender, $e)
	
	# Get selected column number
	$columnNumber = $e.ColumnIndex
	
	# Special handling for file path
	if ($dataGridView.Columns[$columnNumber].Index -eq $GRIDVIEWCOLUMNFILEPATH) {
		$dataGridView.Sort($dataGridView.Columns[$GRIDVIEWCOLUMNVIDEOTYPE], $global:sortDirection)
		# Change sort direction
		if ($global:sortDirection -eq [System.ComponentModel.ListSortDirection]::Ascending) {
			$global:sortDirection = [System.ComponentModel.ListSortDirection]::Descending
			$dataGridView.Columns[$GRIDVIEWCOLUMNFILEPATH].HeaderCell.SortGlyphDirection = [System.Windows.Forms.SortOrder]::Ascending
		} else {
			$global:sortDirection = [System.ComponentModel.ListSortDirection]::Ascending
			$dataGridView.Columns[$GRIDVIEWCOLUMNFILEPATH].HeaderCell.SortGlyphDirection = [System.Windows.Forms.SortOrder]::Descending
		}
	} else {
		# Set sort option for file path column to none
		$dataGridView.Columns[$GRIDVIEWCOLUMNFILEPATH].HeaderCell.SortGlyphDirection = [System.Windows.Forms.SortOrder]::None
		$global:sortDirection = [System.ComponentModel.ListSortDirection]::Ascending
	}
})

# Add event if data grid view has been sorted
$dataGridView.Add_Sorted({
	Mark-Red
})


# Add event handler for DragEnter
$dataGridView.Add_DragEnter({
	param($sender, $e)
	if ($e.Data.GetDataPresent([System.Windows.Forms.DataFormats]::FileDrop)) {
		$e.Effect = [System.Windows.Forms.DragDropEffects]::Copy
	}
})

# Add event handler for Drag'n'Drop
# ToDo implement drag'n'drop
$dataGridView.Add_DragDrop({
	param($sender, $e)
	$files = $e.Data.GetData([System.Windows.Forms.DataFormats]::FileDrop)
	foreach ($file in $files) {
		Write-Host "dataGridView.Add_DragDrop:file=" $file
		# To Do:
		# - check if path is from series or movie
		# - analyze
	}
})

$videoPanel.Controls.Add($dataGridView)


# Create bottom panel
# used for details from themoviedb.org
$detailsPanel = New-Object System.Windows.Forms.Panel
$detailsPanel.Dock = [System.Windows.Forms.DockStyle]::Bottom
$detailsPanel.Height = 230


# First column used for labels
$titleLabel = New-Object System.Windows.Forms.Label
$titleLabel.Location = New-Object System.Drawing.Point(10, 10)
$titleLabel.Size = New-Object System.Drawing.Size(80, 30)

$yearLabel = New-Object System.Windows.Forms.Label
$yearLabel.Location = New-Object System.Drawing.Point(10, 10)
$yearLabel.Size = New-Object System.Drawing.Size(40, 30)

$overviewLabel = New-Object System.Windows.Forms.Label
$overviewLabel.Location = New-Object System.Drawing.Point(10, 50)
$overviewLabel.Size = New-Object System.Drawing.Size(80, 30)

$genresLabel = New-Object System.Windows.Forms.Label
$genresLabel.Location = New-Object System.Drawing.Point(10, 120)
$genresLabel.Size = New-Object System.Drawing.Size(80, 30)

$actorsLabel = New-Object System.Windows.Forms.Label
$actorsLabel.Location = New-Object System.Drawing.Point(10, 160)
$actorsLabel.Size = New-Object System.Drawing.Size(80, 30)

# Second column used for data
$titleTextBox = New-Object System.Windows.Forms.TextBox
$titleTextBox.Location = New-Object System.Drawing.Point((10 + $titleLabel.Width + 10), 10)
$titleTextBox.Width = $detailsPanel.Width - $titleLabel.Width - 30 - 300
$titleTextBox.ReadOnly = $true

$yearTextBox = New-Object System.Windows.Forms.TextBox
$yearTextBox.Location = New-Object System.Drawing.Point(($detailsPanel.Width - 100 + 10 + $titleLabel.Width + 10), 10)
$yearTextBox.Width = 60
$yearTextBox.ReadOnly = $true

$overviewTextBox = New-Object System.Windows.Forms.TextBox
$overviewTextBox.Location = New-Object System.Drawing.Point((10 + $overviewLabel.Width + 10), 50)
$overviewTextBox.Width = $detailsPanel.Width - $overviewLabel.Width - 30
$overviewTextBox.Height = 60
$overviewTextBox.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$overviewTextBox.Multiline = $true
$overviewTextBox.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical + [System.Windows.Forms.ScrollBars]::Horizontal
$overviewTextBox.WordWrap = $true
$overviewTextBox.ReadOnly = $true

$genresTextBox = New-Object System.Windows.Forms.TextBox
$genresTextBox.Location = New-Object System.Drawing.Point((10 + $genresLabel.Width + 10), 120)
$genresTextBox.Width = $detailsPanel.Width - $genresLabel.Width - 30
$genresTextBox.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$genresTextBox.ReadOnly = $true

$actorsTextBox = New-Object System.Windows.Forms.RichTextBox
$actorsTextBox.Location = New-Object System.Drawing.Point((10 + $actorsLabel.Width + 10), 160)
$actorsTextBox.Width = $detailsPanel.Width - $actorsLabel.Width - 30
$actorsTextBox.Height = 60
$actorsTextBox.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$actorsTextBox.Margin = New-Object System.Windows.Forms.Padding(0, 10, 5, 0)
$actorsTextBox.Multiline = $true
$actorsTextBox.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
$actorsTextBox.WordWrap = $true
$actorsTextBox.ReadOnly = $true

$detailsPanel.Controls.AddRange(@($titleLabel, $yearLabel, $overviewLabel, $genresLabel, $actorsLabel, $titleTextBox, $yearTextBox, $overviewTextBox, $genresTextBox, $actorsTextBox ))

# Add panels to form
$form.Controls.AddRange(@($topPanel, $videoPanel, $detailsPanel))
$dataGridView.Width = $videoPanel.Width - 20
$dataGridView.Height = $videoPanel.Height - 20 - 80

# Add resize event for form
$form.Add_Resize({
	$progressTextBox.Width = $topPanel.Width - 20
	$dataGridView.Width = $videoPanel.Width - 20
	$dataGridView.Height = $videoPanel.Height - 20 - 80

	$yearTextBox.Location = New-Object System.Drawing.Point(($detailsPanel.Width - $yearTextBox.Width - 10), $yearTextBox.Location.Y)
	$yearLabel.Location = New-Object System.Drawing.Point(($detailsPanel.Width- $yearTextBox.Width - $yearLabel.Width - 10), $yearLabel.Location.Y)
	$titleTextBox.Width = $yearLabel.Location.X - $titleTextBox.Location.X - 20
})

# Add the shown-event to the form.
$form.Add_Shown({ Mark-Red })

# Add event handler for closing windows
$form.Add_FormClosing({
    if ($form.WindowState -eq [System.Windows.Forms.FormWindowState]::Minimized) {
        $form.WindowState = [System.Windows.Forms.FormWindowState]::Normal
    }
})

# Re-calculate positions after the form is arranged
$yearTextBox.Location = New-Object System.Drawing.Point(($detailsPanel.Width - $yearTextBox.Width - 10), $yearTextBox.Location.Y)
$yearLabel.Location = New-Object System.Drawing.Point(($detailsPanel.Width- $yearTextBox.Width - $yearLabel.Width - 10), $yearLabel.Location.Y)
$titleTextBox.Width = $yearLabel.Location.X - $titleTextBox.Location.X - 20

# Check the loaded configuration
Check-Configuration

# Create SQLite Helper instance
$db = [SQLiteHelper]::new()

# Try opening or creating database
try {
	if (Test-Path -LiteralPath $dbPath) {
		# Just open the video database if it exists
		$db.Open("$dbPath")
	} else {
		# Create a new video database if it does not exists
		$db.Open("$dbPath")
		
		Create-Database-Tables
	}
}
catch {
	# Catch error, if the database couldn't be created or opened
	# write error message and exit
	Write-Warning "ERROR: $_"
	Write-Warning $Error[0].Exception.GetType().FullName
	Exit
}


# Load video data from database
Load-Data

# Sort data grid view by title
$dataGridView.Sort($dataGridView.Columns[$GRIDVIEWCOLUMNTITLE], $global:sortDirection)

# Mark non-existing files
Mark-Red

# Create a timer for regular checking availability of the drives
$global:timer = New-Object System.Windows.Forms.Timer
$global:timer.Interval = 5000 # Intervall in milli seconds

# Timer tick event
$global:timer.Add_Tick({
	$global:moviePath.CheckAvailability()
	$global:seriesPath.CheckAvailability()
	
	# Check if movies path is available
	if (-not ([string]::IsNullOrEmpty($config.MovieFolder))) {
		$global:movieDriveFound = $global:moviePath.GetIsAvailable()
	}
	# Check if series path is available
	if (-not ([string]::IsNullOrEmpty($config.SeriesFolder))) {
		$global:seriesDriveFound = $global:seriesPath.GetIsAvailable()
	}
	
	# Enable the buttons only if movie or series path is available
	$analyzeButton.Enabled = $global:movieDriveFound -or $global:seriesDriveFound
	$renameButton.Enabled = $global:movieDriveFound -or $global:seriesDriveFound
	$moveButton.Enabled = $global:movieDriveFound -or $global:seriesDriveFound
	$rescanButton.Enabled = $global:movieDriveFound -or $global:seriesDriveFound
})

# Start timer
$global:timer.Start()
# Trigger timer manually 
$global:timer.GetType().GetMethod("OnTick", [System.Reflection.BindingFlags]::Instance -bor [System.Reflection.BindingFlags]::NonPublic).Invoke($global:timer, @($null))

######################################################
# Display GUI
Set-MainText $config.Language

# Restore Window position if possible
$form.StartPosition = [System.Windows.Forms.FormStartPosition]::Manual
$form.Size = New-Object System.Drawing.Size($config.Width, $config.Height)
$form.Top = $config.Top
$form.Left = $config.Left
$form.WindowState = $config.WindowState

$closure = $form.ShowDialog()
$form.WindowState = [System.Windows.Forms.FormWindowState]::Normal

# Stop timer
$global:timer.Stop()

# Save configuration if something has changed
Save-Config -config $config

# Close database
$db.Close()

# Force garbage collection
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()

Exit
