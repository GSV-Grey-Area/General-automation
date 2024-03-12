using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.IO.Compression;
using System.Xml;
using System.Xml.Linq;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace ZIPTest
{
	class Program
	{
		[DllImport("user32.dll", SetLastError = true)]
		[return: MarshalAs(UnmanagedType.Bool)]
		private static extern bool SetWindowPos
		(
			IntPtr hWnd,
			IntPtr hWndInsertAfter,
			int x,
			int y,
			int cx,
			int cy,
			int uFlags
		);

		private const int HWND_TOPMOST = -1;
		private const int SWP_NOMOVE = 0x0002;
		private const int SWP_NOSIZE = 0x0001;

		static void Main(string[] args)
		{
			IntPtr hWnd = Process.GetCurrentProcess().MainWindowHandle;
			SetWindowPos(hWnd, new IntPtr(HWND_TOPMOST), 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE);

			string[] zipPathGroup =
			{
				@"..\..\..\..\XGSDA - 00 - XXX.docx",
				@"..\..\..\..\XGSDA - 01 - PRO.docx",
				@"..\..\..\..\XGSDA - 02 - BAS.docx",
				@"..\..\..\..\XGSDA - 03 - SIS.docx",
				@"..\..\..\..\XGSDA - 04 - LEN.docx",
				@"..\..\..\..\XGSDA - 05 - ENT.docx",
				@"..\..\..\..\XGSDA - 06 - FOL.docx",
				@"..\..\..\..\XGSDA - 07 - SER.docx",
				@"..\..\..\..\XGSDA - 08 - CLI.docx",
				@"..\..\..\..\XGSDA - 09 - DIS.docx",
				@"..\..\..\..\XGSDA - 10 - DES.docx",
				@"..\..\..\..\XGSDA - 11 - EIE.docx",
				@"..\..\..\..\XGSDA - 12 - ING.docx",
				@"..\..\..\..\XGSDA - 13 - PDR.docx",
				@"..\..\..\..\XGSDA - 14 - FDR.docx"
			};

			int totWords = 0;
			int totPages = 0;

			foreach (string path in zipPathGroup)
			{
				Console.Write(path + ": ");
				int[] data = new int[2];
				int words = 0;
				int pages = 0;
				using (ZipArchive archive = ZipFile.OpenRead(path))
				{
					foreach (ZipArchiveEntry entry in archive.Entries)
					{
						int wordCount = 0;

						if (entry.Name == "document.xml")
						{
							XmlDocument xmlDoc = new XmlDocument();
							xmlDoc.Load(entry.Open());
							XmlNamespaceManager namespaceManager = new XmlNamespaceManager(xmlDoc.NameTable);
							namespaceManager.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

							XmlNode bodyNode = xmlDoc.SelectSingleNode("//w:body", namespaceManager);
							if (bodyNode != null)
							{
								XmlNodeList paragraphNodes = bodyNode.SelectNodes(".//w:p", namespaceManager);
                
								foreach (XmlNode paragraphNode in paragraphNodes)
								{
									string paragraphText = paragraphNode.InnerText.Trim();
									string[] words1 = paragraphText.Split(new char[] { ' ', '\t', '\n' }, StringSplitOptions.RemoveEmptyEntries);
									wordCount += words1.Length;
								}
							}
              
							data[0] = wordCount;
							totWords += wordCount;
						}

						string input = entry.Name;
						if (entry.Name == "app.xml")
						{
							XElement xelement = XElement.Load(entry.Open());
              
							foreach (XElement e in xelement.DescendantsAndSelf())
							{
								if
								(
									e.Name == @"{http://schemas.openxmlformats.org/officeDocument/2006/extended-properties}Pages"
								)
								{
									pages = (int)e;
									data[1] = pages; 
									totPages += pages;
								}
							}
						}
					}
          Console.Write("\tWords: " + data[0] + "\tPages: " + data[1]);
        }
				Console.WriteLine();
			}

			Console.Write("\t\t\t Total: ");
			Console.Write("\tWords: " + totWords);
			Console.WriteLine("\tPages: " + totPages);

			Console.ReadLine();
		}
	}
}
