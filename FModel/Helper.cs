using System.IO;
using System.Xml.Serialization;

namespace F
{
    public static class Helper
    {
        private static readonly XmlSerializer Xs = new XmlSerializer(typeof(FootballDto));

        public static void WriteToFile(string fileName, FootballDto data)
        {
            using (var fileStream = File.Create(fileName))
            {
                Xs.Serialize(fileStream, data);
            }
        }

        public static FootballDto LoadFromFile (string fileName)
        {
            using (var fileStream = File.OpenRead(fileName))
            {
                return (FootballDto)Xs.Deserialize(fileStream);
            }
        }

        public static FootballDto LoadFromStream(Stream file)
        {
            return (FootballDto)Xs.Deserialize(file);
        }
    }
}
