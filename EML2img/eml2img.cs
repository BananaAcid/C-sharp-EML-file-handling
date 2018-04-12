/*
 * User: BananaAcid
 * Date: 30.07.2012
 * Type: Commandline
 * Licence: MIT
 */
using System;
using System.IO;
using System.Linq;
using System.Text;

public class eml2img
{
	private static int imageCount;
	
	private const string dirDone = @".\__done";

	
	
	static void Main(string[] args)
	{
		if (args.Length == 0)
		{
			args = (Directory.GetFiles(@".\", "*.*")
				.Where(s => s.EndsWith(".txt", StringComparison.OrdinalIgnoreCase) || s.EndsWith(".eml", StringComparison.OrdinalIgnoreCase))
				).Cast<string>().ToArray();
		}
		
		// proccess files
		if (args.Length > 0)
		{
			StreamReader reader = null;
			
			if (!Directory.Exists(dirDone))
				Directory.CreateDirectory(dirDone);

			foreach (string arg in args)
				try
				{
					Console.WriteLine(arg);
					reader = new StreamReader(arg);
					string line;
	
					while ((line = reader.ReadLine()) != null)
					{
						if (line.ToLower().StartsWith("content-disposition: attachment;") || line.ToLower().StartsWith("content-disposition: inline;")) // found attachment
						{
							ExtractContent(reader, GetAttachment(reader, line));
						}
						if (line.ToLower().StartsWith("content-type: image/")) // found embedded image
						{
							ExtractContent(reader, GetImage(reader, line));
						}
					}
				}
				catch (IOException)
				{
					Console.WriteLine("Unable to open file!");
				}
				finally
				{
					if (reader != null) reader.Close();
					
					File.Move(arg, dirDone + @"\" + Path.GetFileName(arg));
				}
		}
		
		Console.WriteLine("[done]");
		Console.ReadKey();
	}

	private static string GetAttachment(TextReader reader, string line)
	{
		if (!line.Contains("filename"))
		{
			line = reader.ReadLine(); // Thunderbird: filename start at second line
		}

		return GetFilename(reader, line);
	}

	private static string GetImage(TextReader reader, string line)
	{
		if (!line.Contains("name"))
		{
			line = reader.ReadLine(); // Thunderbird: filename start at second line
		}

		if (!line.Contains("name")) // embedded image does not have name
		{
			AdvanceToEmptyLine(reader);

			return "image" + imageCount++ + ".jpg"; // default to jpeg
		}

		return GetFilename(reader, line);
	}

	private static string GetFilename(TextReader reader, string line)
	{
		string filename;
		int filenameStart = line.IndexOf('"') + 1;

		if (filenameStart > 0)
		{
			filename = line.Substring(filenameStart, line.Length -
			                          filenameStart - 1);
		}
		else // filename does not have quote
		{
			filenameStart = line.IndexOf('=') + 1;
			filename = line.Substring(filenameStart, line.Length -
			                          filenameStart);
		}
		Console.WriteLine(" - found name: " + filename);
		AdvanceToEmptyLine(reader);

		return filename;
	}

	private static void AdvanceToEmptyLine(TextReader reader)
	{
		string line;

		while ((line = reader.ReadLine()) != null)
		{
			if (String.IsNullOrEmpty(line)) break;
		}
	}

	private static void ExtractContent(TextReader reader, string filename)
	{
		string line;
		var content = new StringBuilder();

		while ((line = reader.ReadLine()) != null)
		{
			if (String.IsNullOrEmpty(line) || line.StartsWith("--------------")) break;
			
			content.Append(line);
		}

		if (content.Length > 0)
		{
			string str = content.ToString().Trim().Replace("\n", "").Replace("\r", "").Replace(" ", "");
			Console.WriteLine(" - Str len: " + str.Length);
			
			//using (StreamWriter f = File.CreateText("test.txt"))
			//       f.Write(str);
			
			
			byte[] buffer = Convert.FromBase64String(str);

			Console.WriteLine("creating " + filename);
			
			using (Stream writer = new FileStream(filename, FileMode.Create))
			{
				writer.Write(buffer, 0, buffer.Length);
			}
		}
	}
}