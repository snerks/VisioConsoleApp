using System;
using System.IO;
using System.IO.Packaging;
using System.Runtime.CompilerServices;
using Xunit;

namespace VisioConsoleApp.Unit.Test
{
    public class Tests : FileCleanupTestBase
    {
        private const string Mime_MediaTypeNames_Text_Xml = "text/xml";
        private const string Mime_MediaTypeNames_Image_Jpeg = "image/jpeg"; // System.Net.Mime.MediaTypeNames.Image.Jpeg
        private const string s_DocumentXml = @"<Hello>Test</Hello>";
        private const string s_ResourceXml = @"<Resource>Test</Resource>";

        private FileInfo GetTempFileInfoFromExistingFile(string existingFileName, [CallerMemberName] string memberName = null, [CallerLineNumber] int lineNumber = 0)
        {
            FileInfo existingDoc = new FileInfo(existingFileName);
            byte[] content = File.ReadAllBytes(existingDoc.FullName);
            FileInfo newFile = new FileInfo($"{GetTestFilePath(null, memberName, lineNumber)}.{existingDoc.Extension}");
            File.WriteAllBytes(newFile.FullName, content);
            return newFile;
        }

        public FileInfo GetTempFileInfoWithExtension(string extension, [CallerMemberName] string memberName = null, [CallerLineNumber] int lineNumber = 0)
        {
            return new FileInfo($"{GetTestFilePath(null, memberName, lineNumber)}.{extension}");
        }

        [Fact]
        public void WriteRelationsTwice()
        {
            //FileInfo tempGuidFile = GetTempFileInfoWithExtension(".zip");
            FileInfo tempGuidFile = GetTempFileInfoWithExtension("zip");

            using (Package package = Package.Open(tempGuidFile.FullName, FileMode.OpenOrCreate, FileAccess.ReadWrite))
            {
                //first part
                PackagePart packagePart = 
                    package.CreatePart(
                        PackUriHelper.CreatePartUri(new Uri("MyFile1.xml", UriKind.Relative)),
                        System.Net.Mime.MediaTypeNames.Application.Octet);

                using (packagePart.GetStream(FileMode.Create))
                {
                    //do stuff with stream - not necessary to reproduce bug
                }

                package.CreateRelationship(
                    PackUriHelper.CreatePartUri(new Uri("MyFile1.xml", UriKind.Relative)),
                    TargetMode.Internal, "http://my-fancy-relationship.com");

                package.Flush();

                //create second part after flush
                packagePart = 
                    package.CreatePart(
                        PackUriHelper.CreatePartUri(new Uri("MyFile2.xml", UriKind.Relative)),
                        System.Net.Mime.MediaTypeNames.Application.Octet);

                using (packagePart.GetStream(FileMode.Create))
                {
                    //do stuff with stream - not necessary to reproduce bug
                }

                package.CreateRelationship(
                    PackUriHelper.CreatePartUri(new Uri("MyFile2.xml", UriKind.Relative)),
                    TargetMode.Internal, "http://my-fancy-relationship.com");
            }
        }
    }
}
