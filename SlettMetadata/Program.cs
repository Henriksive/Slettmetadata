using System;
using DocumentFormat.OpenXml.Packaging;

namespace SlettMetadata
{
	class Program
	{
		static void Main(string[] args)
		{
			try
			{
				var filename = args[0];
				using (var package = WordprocessingDocument.Open(filename, true))
				{
					// modify properties
					package.PackageProperties.Creator = null;
					package.PackageProperties.LastModifiedBy = null;
					package.PackageProperties.Revision = null;
				}
				Console.WriteLine("Slettet metadata for filen: " + filename);
			}
			catch (Exception e)
			{
				Console.WriteLine("En feil oppstod: " + e.Message);
			}
		}
	}
}
