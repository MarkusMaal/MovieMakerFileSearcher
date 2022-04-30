using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace MovieMakerFileSearcher
{
    class Program
    {

        // we define main as int, so that we can return error codes
        //
        // this can be convinient for scripting or programming
        // purposes
        //
        // the error codes are the following:
        // 0 - Success, no errors detected
        // 1 - Incorrect syntax
        // 2 - Not enough arguments
        // 3 - Too many arguments
        // 4 - File/folder does not exist
        // 5 - XML data not found
        // 6 - No thumbnails found
        // 7 - Footer found before header
        // 8 - Couldn't find any MSWMM headers
        static int Main(string[] args)
        {
            if (args.Length == 0)
            {
                while (true)
                {
                    Console.Clear();
                    Console.WriteLine("Welcome to Windows Movie Maker forensics!");
                    Console.WriteLine("\nPlease choose an activity or press Ctrl+C to quit:\n");
                    Console.WriteLine("1. Find Movie Maker files");
                    Console.WriteLine("\tIn case you have lost a Movie Maker file somewhere, changed the file extension or used\n\ta data recovery software such as PhotoRec, you can use this to find Movie Maker headers.");
                    Console.WriteLine("2. Extract thumbnail images");
                    Console.WriteLine("\tDid you know that old Movie Maker files have thumbnail images cached right inside the\n\tproject file? Well, this option allows you to recover those thumbnail images, even\n\tif the project itself is corrupted!");
                    Console.WriteLine("3. Extract Movie Maker XML file");
                    Console.WriteLine("\tAttempt to find and extract the XML data from the Movie Maker file. This is useful when\n\tyou can't open the file in any way, but still want to see how the project was put\n\ttogether.");
                    switch (Console.ReadKey().Key)
                    {
                        case ConsoleKey.D1:
                            Console.Write("\nDRAG and DROP the folder you'd like to search from to this window and press Enter. This folder can contain subfolders.\n\n");
                            FindMovieMk(Console.ReadLine());
                            break;
                        case ConsoleKey.D2:
                            Console.Write("\nDRAG and DROP the .MSWMM file to this window and press Enter. Corrupted files are supported.\n\n");
                            ExtractJFIF(Console.ReadLine());
                            break;
                        case ConsoleKey.D3:
                            Console.Write("\nDRAG and DROP the .MSWMM file to this window and press Enter. Semi-corrupted files are supported.\n\n");
                            ExtractXML(Console.ReadLine());
                            break;
                    }
                }
                return 0;
            }
            switch (args[0])
            {
                case "/?":
                    if (args.Length > 2)
                    {
                        Console.Error.WriteLine("Too many arguments specified.");
                        return 3;
                    }
                    string me = AppDomain.CurrentDomain.FriendlyName;
                    Console.WriteLine("Windows Movie Maker forensics");
                    Console.WriteLine(string.Format("\nSyntax: {0} [/f /i /x /?] [mswmm file/folder path]", me));
                    Console.WriteLine("\n/f - This allows you to find MSWMM files inside a specific folder.");
                    Console.WriteLine("/i - This allows you to extract thumbnails from the MSWMM file.");
                    Console.WriteLine("/x - This allows you to extract and deobfuscate the XML data from the MSWMM file.");
                    Console.WriteLine("/? - Display quick help.");
                    Console.WriteLine("\nExamples:\n\n");
                    Console.WriteLine(string.Format("{0} /f \"C:\\random\\folder, which\\might\\contain\\mswmm\\files\"", me));
                    Console.WriteLine(string.Format("{0} /i awesome.mswmm", me));
                    Console.WriteLine(string.Format("{0} /x \"C:\\random folder\\awesome.mswmm\"", me));
                    return 0;
                case "/f":
                    if (args.Length == 1)
                    {
                        Console.Error.WriteLine("Not enough arguments specified.");
                        return 2;
                    }
                    else if (args.Length > 2)
                    {
                        Console.Error.WriteLine("Too many arguments specified.");
                        return 3;
                    }
                    return FindMovieMk(args[1]);
                case "/i":
                    if (args.Length == 1)
                    {
                        Console.Error.WriteLine("Not enough arguments specified.");
                        return 2;
                    }
                    else if (args.Length > 2)
                    {
                        Console.Error.WriteLine("Too many arguments specified.");
                        return 3;
                    }
                    return ExtractJFIF(args[1]);
                case "/x":
                    if (args.Length == 1)
                    {
                        Console.Error.WriteLine("Not enough arguments specified.");
                        return 2;
                    }
                    else if (args.Length > 2)
                    {
                        Console.Error.WriteLine("Too many arguments specified.");
                        return 3;
                    }
                    return ExtractXML(args[1]);
                default:
                    Console.Error.WriteLine("The syntax of the command is incorrect.");
                    return 1;
            }
        }

        static int ExtractXML(string mswmm)
        {
            // remove " from filename
            mswmm = mswmm.Replace("\"", "");
            if (!File.Exists(mswmm))
            {
                Console.Error.WriteLine("The specified file does not exist.");
                return 4;
            }
            // get source directory
            string source = new FileInfo(mswmm).Directory.FullName;
            // set destination filename
            string dest = string.Join(".", new FileInfo(mswmm).Name.Split('.')[..^1]) + ".XML";
            Console.WriteLine("Searching for XML data...");

            string fullname = source + "\\" + dest;
            // specify how many bytes to read at a time
            // do not modify, unless you know what you're doing
            int precision = 32;
            // this list stores raw bytes of the XML data
            List<byte> unicode = new List<byte>();
            using (Stream src = File.OpenRead(mswmm))
            {
                // create buffer
                byte[] buffer = new byte[precision];
                int offset = 0;
                // if true, allow appending bytes to unicode list
                bool write = false;
                // create last buffer memory
                byte[] last = buffer;

                while ((offset = src.Read(buffer, 0, buffer.Length)) > 0)
                {
                    // specify XML header and footer
                    // do not modify, unless you know exactly what you're doing
                    string start = "3C 00 4D 00 6F 00 76 00 69 00 65 00 4D 00 61 00 6B 00 65 00 72";
                    string end = "69 00 65 00 4D 00 61 00 6B 00 65 00 72 00 3E 00";
                    // copy buffer to separate array
                    // this ensures, only 1 read is being done to the
                    // buffer at a time
                    byte[] current = buffer;
                    if (!write)
                    {
                        if (BitConverter.ToString(current).Replace("-", " ").Contains(start))
                        {
                            Console.WriteLine("XML data found. Processing...");
                            write = true;
                        }
                    }
                    if (write)
                    {
                        unicode.AddRange(current);
                        if (BitConverter.ToString(current).Replace("-", " ").Contains(end))
                        {
                            Console.WriteLine("Finished processing XML data");
                            break;
                        }
                    }
                }
                // here, Unicode = UTF-16 LE
                // mswmm files use UTF-16, because it basically allows to save any character
                // from most languages at the cost of 2x space usage
                if (unicode.Count == 0)
                {
                    Console.Error.WriteLine("Couldn't find any XML data. This could mean that the file is either severly corrupted or not a valid MSWMM file.");
                    return 5;
                }
                Console.WriteLine("Converting byte array to string using UTF-16 LE encoding...");
                string uni = Encoding.Unicode.GetString(unicode.ToArray());
                // Normally, MSWMM contains the XML data in a single line
                //
                // This is easy for computers to read, but might be difficult for
                // humans to understand. This is why we need to deobfuscate the
                // XML.
                Console.WriteLine("Deobfuscating XML...");
                string finalxml = PrintXML(uni);

                // Deletes existing XML file, if it exists
                if (File.Exists(fullname))
                {
                    Console.WriteLine("Deleting existing file...");
                    File.Delete(fullname);
                }

                // Write XML to a file
                // To open the file, your text editor must support UTF-16 LE
                // encoding (which is fine, because the Windows notepad does
                // support it).
                Console.WriteLine("Writing new file...");
                File.WriteAllText(fullname, finalxml, Encoding.Unicode);
            }
            Console.WriteLine("Finished!");
            return 0;
        }

        // deobfuscates the XML data
        public static string PrintXML(string xml)
        {
            // make sure the XML is properly aligned
            xml = string.Join(">", xml.Split('>')[..^1]) + ">";
            string final = "";
            string tab = "";
            // split xml by "<" tokens
            // reconstruct a human-readable version
            foreach (string line in xml.Split("<"))
            {
                if (!line.Contains(">")) { continue; }
                // de-indent if neccessary
                if (line.Contains("/") && line.EndsWith(">") && !line.Contains("=") && !line.EndsWith("/>"))
                {
                    if (tab.Length > 0)
                    {
                        tab = tab.Substring(1);
                    }
                }
                // we add \r\n, so that even older versions of Windows can display
                // the XML file properly using notepad and it's a more Windows way
                // of doing things
                final += tab + "<" + line + "\r\n";
                // indent
                if ((!line.EndsWith("/>")) && line.EndsWith(">"))
                {
                    tab = tab + "\t";
                }
                // de-indent if neccessary
                if (line.Contains("/") && line.EndsWith(">") && !line.Contains("=") && !line.EndsWith("/>"))
                {
                    if (tab.Length > 0)
                    {
                        tab = tab.Substring(1);
                    }
                }
            }
            // remove the last carriage return and newline from the final string
            return final.Substring(0, final.Length - 2);
        }

        // for extracting thumbnails
        static int ExtractJFIF(string mswmm)
        {
            // remove " from the filename
            mswmm = mswmm.Replace("\"", "");
            if (!File.Exists(mswmm))
            {
                Console.Error.WriteLine("The file specified does not exist.");
                return 4;
            }
            // get file location
            string source = new FileInfo(mswmm).Directory.FullName;
            // gets the filename without the file extension
            // we split and rejoin by ".", because the filename can contain
            // multiple dots
            string dest = string.Join(".", new FileInfo(mswmm).Name.Split('.')[..^1]);
            // create a new directory for storing extracted images
            if (!Directory.Exists(source + "\\" + dest))
            {
                Directory.CreateDirectory(source + "\\" + dest);
            }
            // this int is incremented each time a new thumbnail is found
            int filename = 1;
            // setup by how many bytes should the file bytes be walked by
            // do not mess with this, unless you know what you're doing
            int precision = 16;
            bool warning = false;
            Console.WriteLine(string.Format("Extracted files will be saved as {0}\\[X].JFIF", dest));
            using (Stream src = File.OpenRead(mswmm))
            {
                // setup the buffer
                byte[] buffer = new byte[precision];
                int offset = 0;
                long loc = 0;
                // if true, the thumbnail is being written to
                bool write = false;
                // full path name where to save the thumbnail
                string fullname = "";
                // last buffer data
                byte[] last = buffer;
                // if true, stops writing the current thumbnail file
                bool endwrite = false;
                while ((offset = src.Read(buffer, 0, buffer.Length)) > 0)
                {
                    // read some data from buffer, copy to a separate array
                    byte[] current = buffer;
                    // specify JFIF header and footer
                    string header = "FF D8 FF E0 00 10 4A 46 49 46";
                    string footer = "FF D9 00 00";
                    if (!write)
                    {
                        // check if current buffer contains JFIF header
                        if (BitConverter.ToString(current).Replace("-", " ").Contains(header))
                        {
                            Console.WriteLine(string.Format("Found header at {0} - Filename: {1}.JFIF", loc, filename));
                            fullname = source + "\\" + dest + "\\" + filename.ToString() + ".JFIF";
                            write = true;
                            if (last != current)
                            {
                                using (Stream dst = File.OpenWrite(fullname))
                                {
                                    dst.Write(last, 0, last.Length);
                                }
                            }
                        } else
                        {
                            // JFIF header not found, continue searching
                            loc += precision;
                            last = current;

                            if (BitConverter.ToString(current).Replace("-", " ").Contains(footer))
                            {
                                Console.Error.WriteLine("Warning: JFIF footer found before header - file may be severly corrupted!!!");
                                warning = true;
                            }
                            continue;
                        }
                    }
                    // search for footer
                    if (BitConverter.ToString(current).Replace("-", " ").Contains(footer))
                    {
                        endwrite = true;
                        filename++;
                        //Console.WriteLine(string.Format("Found footer at {0}", loc));
                    }
                    if (write)
                    {
                        // append current buffer to destination file
                        using (var dst = new FileStream(fullname, FileMode.Append))
                        {
                            dst.Write(current, 0, current.Length);
                        }
                        // if endwrite requested, stop writing
                        if (endwrite)
                        {
                            write = false;
                            endwrite = false;
                        }
                    }
                    // increment position
                    loc += precision;
                    last = current;
                }
            }
            if (filename == 1)
            {
                Console.Error.WriteLine("Couldn't find any thumbnails!!!");
                return 6;
            }
            else
            {
                Console.WriteLine("Done!");
                if (warning)
                {
                    return 7;
                }
                else
                {
                    return 0;
                }
            }
        }

        // searches for MSWMM files
        static int FindMovieMk(string folder, bool root = true)
        {
            if (root)
            {
                if (File.Exists("search.log"))
                {
                    Console.WriteLine("Deleting existing log file");
                    File.Delete("search.log");
                }
                Console.WriteLine("Search has been initiated. Please wait...");
            }
            // remove " from filename
            folder = folder.Replace("\"", "");
            if (!Directory.Exists(folder))
            {
                Console.Error.WriteLine("The folder specified doesn't exist.");
                return 4;
            }
            foreach (DirectoryInfo di in new DirectoryInfo(folder).EnumerateDirectories())
            {
                // uncomment the if statement, if you want to search PhotoRec directories specifically
                //if (di.Name.StartsWith("recup_dir."))
                //{

                // the if statement here makes sure that the folder isn't actually a junction or symlink
                // (can help avoid infinite loops)
                if ((di.Attributes & FileAttributes.ReparsePoint) != FileAttributes.ReparsePoint) {
                    try
                    {
                        FindMovieMk(di.FullName, false);
                    } catch (Exception ex)
                    {
                        Console.Error.WriteLine(string.Format("Unable to scan \"{0}\" - {1}", di.FullName, ex.Message));
                    }
                }

                //}
            }
            foreach (FileInfo fn in new DirectoryInfo(folder).EnumerateFiles())
            {
                bool exitfile = false;
                using (Stream src = File.OpenRead(fn.FullName))
                {
                    byte[] buffer = new byte[2048];
                    int offset = 0;
                    long loc = 0;
                    byte[] previous = new byte[2048];
                    while ((offset = src.Read(buffer, 0, buffer.Length)) > 0)
                    {
                        byte[] current = buffer;
                        //
                        // search for the "<MovieMaker" XML header
                        // this ensures that the file is indeed actually a MSWMM file
                        //
                        // the commented search strings may allow you to find more files, but
                        // are suspect to false positives and false negatives
                        //
                        string searchable = "3C 00 4D 00 6F 00 76 00 69 00 65 00 4D 00 61 00 6B 00 65 00 72 00";
                        //string searchable = "3C 50 72 6F 6A 65 63 74";
                        //string searchable = "50 00 72 00 6F 00 64 00 75 00 63 00 65 00 72 00";
                        if (BitConverter.ToString(current.Concat(previous).ToArray()).Replace("-", " ").Contains(searchable))
                        {
                            File.AppendAllText("search.log", string.Format("{0}\\{1} - MSWMM XML header found!", new DirectoryInfo(folder).FullName, fn.Name) + "\r\n");
                            Console.WriteLine(string.Format("{0}\\{1} - MSWMM XML header found!", new DirectoryInfo(folder).FullName, fn.Name));
                            exitfile = true;
                        }
                        previous = current;
                        loc += 2048;
                        if (exitfile) { break; }
                    }
                }

            }
            if (root)
            {
                if (!File.Exists("search.log"))
                {
                    Console.Error.WriteLine("No MSWMM headers found!!!");
                    return 8;
                }
                Console.WriteLine("Press any key to continue . . . ");
                Console.ReadKey();
            }
            return 0;
        }
    }
}
