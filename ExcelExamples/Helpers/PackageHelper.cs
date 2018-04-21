using System;
using System.IO;
using System.Text;

namespace ExcelExamples.Helpers {
    public static class PackageHelper {

        public static void CopyStream(Stream source, Stream target) {
            const int bufSize = 16384;
            byte[] buf = new byte[bufSize];
            int bytesRead = 0;
            while ((bytesRead = source.Read(buf, 0, bufSize)) > 0)
                target.Write(buf, 0, bytesRead);
        }

        public static Uri GetRelativeUri(string currentFile) {
            string relPath = currentFile.Substring(currentFile
            .IndexOf('\\')).Replace('\\', '/').Replace(' ', '_');
            return new Uri(RemoveAccents(relPath), UriKind.Relative);
        }

        public static string RemoveAccents(string input) {
            string normalized = input.Normalize(NormalizationForm.FormKD);
            Encoding removal = Encoding.GetEncoding(Encoding.ASCII.CodePage, new EncoderReplacementFallback(""), new DecoderReplacementFallback(""));
            byte[] bytes = removal.GetBytes(normalized);
            return Encoding.ASCII.GetString(bytes);
        }
    }
}
