using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml;
using System.Xml.Serialization;

namespace ExcelImport
{
    public static class WorkbookFactory
    {
        public static SharedStrings SharedStrings { get; set; }

        public static Workbook LoadFile(string nomeDoArquivo)
        {
            var workbook = new Workbook();

            if (!File.Exists(nomeDoArquivo))
                throw new FileNotFoundException();

            ZipArchive zipArchive = ZipFile.Open(nomeDoArquivo, ZipArchiveMode.Read);

            SharedStrings = DeserializedZipEntry<SharedStrings>(GetZipArchiveEntry(zipArchive, @"xl/sharedStrings.xml"));
            foreach (var worksheetEntry in (WorkSheetFileNames(zipArchive)).OrderBy(x => x.FullName))
            {
                workbook.Worksheets.Add(DeserializedZipEntry<Worksheet>(worksheetEntry));
            }

            zipArchive.Dispose();

            return workbook;
        }
        
        private static ZipArchiveEntry GetZipArchiveEntry(ZipArchive ZipArchive, string ZipEntryName)
        {
            return ZipArchive.Entries.First<ZipArchiveEntry>(n => n.FullName.Equals(ZipEntryName));
        }
        private static IEnumerable<ZipArchiveEntry> WorkSheetFileNames(ZipArchive ZipArchive)
        {
            foreach (var zipEntry in ZipArchive.Entries)
                if (zipEntry.FullName.StartsWith("xl/worksheets/sheet"))
                    yield return zipEntry;
        }
        private static T DeserializedZipEntry<T>(ZipArchiveEntry ZipArchiveEntry)
        {
            using (Stream stream = ZipArchiveEntry.Open())
                return (T)new XmlSerializer(typeof(T)).Deserialize(XmlReader.Create(stream));
        }
    }
}
