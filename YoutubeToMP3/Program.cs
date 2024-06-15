using System;
using IWshRuntimeLibrary;
using System.IO;
using System.Threading.Tasks;
using VideoLibrary;
using YoutubeExplode;
using YoutubeExplode.Videos.Streams;
using YoutubeExplode.Common;
using YoutubeExplode.Playlists;
using YoutubeExplode.Videos;
using NAudio.Wave;
using NAudio.Lame;
using AngleSharp.Dom;
class Program
{
    public static Linkreader LR = new Linkreader();
    static async Task Main(string[] args)
    {
        //Check if all Data and Folder are Created
        LR.DataCheck();
        //Start the Programm
        await Menu();
        Console.WriteLine();
        Console.WriteLine("Press enter to close");
        Console.ReadKey();
    }

    static async Task Menu()
    {

        Console.WriteLine("YoutubeToMP3 by RayCraft Studios");
        Console.WriteLine("--------------------------------");
        Console.WriteLine("Select download mode: ");
        Console.WriteLine("1:   Download as MP4");
        Console.WriteLine("2:   Download as MP3");
        string format = Console.ReadLine();

        var youtube = new YoutubeClient();

        //Create list to load txt file line for line in it
        List<string> ytlinks = new List<string>();
        using (StreamReader sr = new StreamReader(LR.GetFile()))
        {
            string zeile;
            while ((zeile = sr.ReadLine()) != null)
            {
                if (zeile.Trim().StartsWith("http"))
                {
                    ytlinks.Add(zeile);
                }
            }
        }

        /**
         * Download Videos
         */

        //check if list is empty and if not start Downloader for every entry in list
        if (ytlinks != null && ytlinks.Count > 0)
        {
            Console.WriteLine("Start Downloading...");
            try
            {
                foreach (var videoUrl in ytlinks)
                {
                    //Check if URL is from a Playlist
                    if (videoUrl.Contains("playlist"))
                    {
                        await DownloadPlaylist(youtube, videoUrl, format);
                    }
                    else
                    {
                        await DownloadVideo(youtube, videoUrl, format);
                    }

                    Console.WriteLine("Download abgeschlossen.");
                    LR.RemoveLink(videoUrl);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while downloading the videos: " + ex.Message);
            }
        }
        else
        {
            Console.WriteLine("No links in grablist");
        }

        
    }

    //Download Video method
    static async Task DownloadVideo(YoutubeClient youtube, string videoUrl, string format)
    {

        var videoId = VideoId.Parse(videoUrl);
        var video = await youtube.Videos.GetAsync(videoId);
        Console.WriteLine($"Downloading {video.Title}...");

        var streamManifest = await youtube.Videos.Streams.GetManifestAsync(videoId);
        var streamInfo = streamManifest.GetMuxedStreams().GetWithHighestVideoQuality();

        //$"{LR.GetVideoFolder}/{video.Title}.mp4"
        var mp4FilePath = Path.Combine(LR.GetVideoFolder(), video.Title + ".mp4");
        var mp3FilePath = Path.Combine(LR.GetMusicFolder(), video.Title + ".mp3");

        //If you want to install it as MP4
        if (format == "1")
        {
            await youtube.Videos.Streams.DownloadAsync(streamInfo, mp4FilePath);
        }
        else //If you want to install it as MP3
        {
            await youtube.Videos.Streams.DownloadAsync(streamInfo, mp4FilePath);

            ConvertMp4ToMp3(mp4FilePath, mp3FilePath);

            // Original MP4-Datei löschen
            if (System.IO.File.Exists(mp3FilePath))
            {
                System.IO.File.Delete(mp4FilePath);
                Console.WriteLine("Original MP4-Datei gelöscht.");
            }
        }
    }

    //Download whole Playlist
    static async Task DownloadPlaylist(YoutubeClient youtube, string playlistUrl, string format)
    {
        var playlistId = PlaylistId.Parse(playlistUrl);
        var playlist = await youtube.Playlists.GetAsync(playlistId);

        await foreach (var video in youtube.Playlists.GetVideosAsync(playlistId))
        {
            Console.WriteLine($"Downloading {video.Title}...");
            await DownloadVideo(youtube, video.Url, format);
        }
    }

    //Convert the File
    static void ConvertMp4ToMp3(string inputFilePath, string outputFilePath)
    {
        try
        {
            // Extrahiere das Audio aus der MP4-Datei
            using (var mediaFile = System.IO.File.OpenRead(inputFilePath))
            {
                var outputWaveFile = Path.ChangeExtension(outputFilePath, ".wav");
                using (var reader = new MediaFoundationReader(inputFilePath))
                {
                    WaveFileWriter.CreateWaveFile(outputWaveFile, reader);
                }

                // Konvertiere die WAV-Datei in eine MP3-Datei
                using (var reader = new AudioFileReader(outputWaveFile))
                using (var writer = new LameMP3FileWriter(outputFilePath, reader.WaveFormat, LAMEPreset.VBR_90))
                {
                    reader.CopyTo(writer);
                }

                // Lösche die temporäre WAV-Datei
                System.IO.File.Delete(outputWaveFile);
            }

            Console.WriteLine("Umwandlung abgeschlossen.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Fehler bei der Umwandlung: {ex.Message}");
        }
    }

    internal class Linkreader
    {
        /**
        * Set filepath and folderpath
        */
        private static string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
        private static string rootPath = @"C:\YoutubeToMP3";
        private static string filePath = Path.Combine(rootPath, "grablist.txt");
        public static string videoFolderPath = Path.Combine(rootPath, "Downloaded_Videos");
        public static string musicFolderPath = Path.Combine(rootPath, "Downloaded_Music");
        private static bool IsChanged = false;

        //Get methods
        public String GetFile() { return filePath; }
        public String GetVideoFolder() { return videoFolderPath; }
        public String GetMusicFolder() { return musicFolderPath; }

        /*
         Check if data exist
         */
        public void DataCheck()
        {
            //check if root folder for app exist
            if (!Directory.Exists(rootPath))
            {
                Console.WriteLine(rootPath + " not found. Start create folder");
                Console.WriteLine();
                //Create Dir
                Directory.CreateDirectory(rootPath);

                //Create Shortcut for Dir
                CreateShortcut(desktopPath, "YoutubeToMP3.lnk", rootPath);
                Console.WriteLine();
                Console.WriteLine(rootPath + " and shortcut created");
                IsChanged = true;
            }

            //check if folder for Downloaded videos exist
            if (!Directory.Exists(videoFolderPath))
            {
                Console.WriteLine(videoFolderPath + " not found. Start create folder");
                Console.WriteLine();
                //Create Dir
                Directory.CreateDirectory(videoFolderPath);
                Console.WriteLine();
                Console.WriteLine(videoFolderPath + " created");
                IsChanged = true;
            }

            //check if folder for Downloaded videos exist
            if (!Directory.Exists(musicFolderPath))
            {
                Console.WriteLine(musicFolderPath + " not found. Start create folder");
                Console.WriteLine();
                //Create Dir
                Directory.CreateDirectory(musicFolderPath);
                Console.WriteLine();
                Console.WriteLine(musicFolderPath + " created");
                IsChanged = true;
            }

            //check if txt file for video scraping exists
            if (!System.IO.File.Exists(filePath))
            {
                Console.WriteLine(filePath + " not found. Start create file");
                Console.WriteLine();
                //Create txt file
                System.IO.File.WriteAllText(filePath, "//Paste Links here! One line per Link. Don't make empty lines between");
                Console.WriteLine();
                Console.WriteLine(filePath + " created");
                Console.WriteLine();
                Console.WriteLine("Fill your links in grablist and restart this application");
                Console.WriteLine();
                IsChanged = true;
            }

            if (IsChanged)
            {
                Console.WriteLine("Press enter to close");
                Console.ReadKey();
                //End programm
                Environment.Exit(0);
            }
        }

        //Create shortcut to Desktop
        private void CreateShortcut(string targetDirectory, string shortcutName, string targetPath)
        {
            string shortcutPath = Path.Combine(targetDirectory, shortcutName);

            // Create Object for Shortcut
            WshShell shell = new WshShell();
            IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(shortcutPath);

            //set shortcut attributes and save
            shortcut.TargetPath = targetPath;
            shortcut.Save();
        }

        public void RemoveLink(string link)
        {
            // Temporäre Datei erstellen
            string tempDatei = @"C:\Users\Public\Videos\tmp.txt";
            string deleted = null;

            // StreamReader zum Lesen der aktuellen Datei
            using (StreamReader sr = new StreamReader(filePath))
            {
                // StreamWriter zum Schreiben auf die temporäre Datei
                using (StreamWriter sw = new StreamWriter(tempDatei))
                {
                    string zeile;
                    while ((zeile = sr.ReadLine()) != null)
                    {
                        if (!zeile.Contains(link))
                        {
                            sw.WriteLine(zeile);
                        }
                        else
                        {
                            deleted = zeile; // Speichere die gelöschte Zeile, falls gewünscht
                        }
                    }
                }
            }

            // Kopiere die temporäre Datei zurück in die ursprüngliche Datei
            System.IO.File.Copy(tempDatei, filePath, true);

            // Lösche die temporäre Datei
            System.IO.File.Delete(tempDatei);

            Console.WriteLine(deleted + " wurde aus der Liste entfernt.");
        }
    }
}
