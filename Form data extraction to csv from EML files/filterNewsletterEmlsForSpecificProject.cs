/*
 * User: BananaAcid
 * Date: 09.02.2011
 * Type: Commandline
 * Note: Form data extraction to csv from email (*.eml) files
 * Licence: MIT
 */
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Diagnostics;
using System.Text.RegularExpressions;

namespace FilterWebsiteFormEmlsForSpecificProject
{
    class Program
    {
        static void Main(string[] args)
        {
            /* idea:
             * 
             * read each eml file
             * - extract body ( rest of mail after header's empty line
             * - strip tags (html mail)
             * - concat lines on line-ending equals sign "="
             * - break lines on dashes "-" (remove -)
             * 
             * - check for newsletter: ja
             * - extract emails
             */

            /*
             * second / simple way:
             * 
             * read each eml file
             * - strip tags
             * - strip spaces
             * - find lowercase "newsletter:ja"
             *    '-> regex find email
             */

            var currentPath = Path.GetDirectoryName(Process.GetCurrentProcess().MainModule.FileName);

            var outFile = currentPath + @"\emails.txt";

            //Console.WriteLine(currentPath);

            Regex EmailRegex = new Regex("(?:(?#local-part)(?#quoted)" + (char)34 + "[^\\" + (char)34 + "]*" + (char)34 + "|(?#non-quoted)[a-z0-9&+_-](?:\\.?[a-z0-9&+_-]+)*)@(?:(?#domain)(?#domain-name)[a-z0-9](?:[a-z0-9-]*[a-z0-9])*(?:\\.[a-z0-9](?:[a-z0-9-]*[a-z0-9])*)*" + "|(?#ip)(\\[)?(?:[01]?\\d?\\d|2[0-4]\\d|25[0-5])(?:\\.(?:[01]?\\d?\\d|2[0-4]\\d|25[0-5])){3}(?(1)\\]|))", RegexOptions.IgnoreCase);

            List<Regex> PartsRegexList = new List<Regex>()
            {
                new Regex(@"(?<find>[\w|\-|\.]+@([\w|\-|\.]+\.)+[\w|\-|\.]+)"),
                new Regex(@"(Vorname|Vorame):(\s+)(?<find>.*)\n", RegexOptions.IgnoreCase),
                new Regex(@"Nachname:(\s+)(?<find>.*)\n", RegexOptions.IgnoreCase)
            };


            var totalEmails = 0;
            var mailcounts = 0;

            var cc = Console.ForegroundColor;

            Console.Title = "Extracting email from multiple eml bodies if 'newsletter:ja'";
            Console.SetWindowSize(111, Console.LargestWindowHeight);
            Console.WindowTop = 0;

            if (File.Exists(outFile))
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("outputfile exists allready, remove it !\n\nif you press a key, it will append to it!");
                Console.ReadKey();
                Console.Clear();
                Console.ForegroundColor = cc;
            }

            using (StreamWriter writer = File.AppendText(outFile))
            {

                DirectoryInfo di = new DirectoryInfo(currentPath + @"\files");
                FileInfo[] rgFiles = di.GetFiles("*.eml");

                totalEmails = rgFiles.Length;

                Console.WriteLine("TOTAL FILES: " + totalEmails.ToString() + "\n");
                writer.WriteLine("TOTAL FILES: " + totalEmails.ToString() + "\n");

                var i = 0;
                foreach (FileInfo fi in rgFiles)
                {
                    i++;

                    string contents = File.ReadAllText(fi.FullName).Replace("\r", "").Replace("\0", "");


                    // strip html tags
                    contents = StripTagsCharArray(contents);

                    contents = contents.Replace("=\n", "");
                    contents = contents.Replace(" - ", "\n");

                    if (contents.ToLower().Contains("quoted-printable"))
                    {
                        contents = DecodeQuotedPrintable(contents);

                        Console.ForegroundColor = ConsoleColor.Green;
                    }

                    // has newsletter param?
                    if (contents.Replace(" ", "").ToLower().Contains("newsletter:ja"))
                    {
                        try
                        {
                            // get body part
                            contents = contents.Substring(contents.IndexOf("\n\n", 1));

                            // do find rest
                            var userdata = "";
                            foreach (Regex rx in PartsRegexList)
                            {
                                string mc = rx.Match(contents).Groups["find"].Value;
                                userdata += mc.Trim() + "\t";
                            }


                            // do emails ..... argh... there are some cases killing this....
                            /*
                            MatchCollection totalMatches = EmailRegex.Matches(contents);
                            foreach (Match emailMatch in totalMatches)
                            {
                                writer.WriteLine(emailMatch.ToString() + userdata);
                            }
                            
                            mailcounts += totalMatches.Count;
                            Console.Write((totalMatches.Count == 1 ? "+" : totalMatches.Count.ToString()));
                            */

                            writer.WriteLine(userdata + "\t\t\t\t\t\t" + fi.Name);
                            Console.Write("+"); mailcounts++;
                        }
                        catch (Exception e)
                        {
                            Console.ForegroundColor = ConsoleColor.Red;
                            Console.Write("E");
                        }
                    }
                    else
                        Console.Write(".");


                    //writer.WriteLine(fi.Name);


                    Console.ForegroundColor = cc;
                    if (i % 80 == 0) Console.Write(" " + i.ToString() + "/" + rgFiles.Length.ToString() + "\n");
                }

                Console.ForegroundColor = cc;
                Console.Write(" " + i.ToString() + "/" + rgFiles.Length.ToString() + "\n");
            }


            Console.WriteLine("\n\nTOTAL FILES: " + totalEmails.ToString());
            Console.WriteLine("TOTAL EMAILS FOUND: " + mailcounts.ToString());
            
            Console.WriteLine("\n\npress any key to exit and to open the outputfile");
            Console.ReadKey();

            Process.Start(outFile);
        }



        /// <summary>
        /// Remove HTML tags from string using char array.
        /// </summary>
        public static string StripTagsCharArray(string source)
        {
            char[] array = new char[source.Length];
            int arrayIndex = 0;
            bool inside = false;

            for (int i = 0; i < source.Length; i++)
            {
                char let = source[i];
                if (let == '<')
                {
                    inside = true;
                    continue;
                }
                if (let == '>')
                {
                    inside = false;
                    continue;
                }
                if (!inside)
                {
                    array[arrayIndex] = let;
                    arrayIndex++;
                }
            }
            return new string(array, 0, arrayIndex);
        }

        public static String DecodeQuotedPrintable(String strQuoted)
        {
            strQuoted = Regex.Replace(strQuoted,
                                       @"=\x0d\x0a", "",
                                       RegexOptions.Compiled);
            strQuoted = Regex.Replace(strQuoted,
                                       @"=(?<hexcode>[0-9a-f][0-9a-f])", new MatchEvaluator(DecodeQuotedPrintableMatchEvaluator),
                                       RegexOptions.Compiled | RegexOptions.IgnoreCase);
            return strQuoted;
        }

        protected static String DecodeQuotedPrintableMatchEvaluator(Match m)
        {
            return new String((char)(Int32.Parse(m.Groups["hexcode"].Value, System.Globalization.NumberStyles.HexNumber)), 1);
        }
    }
}