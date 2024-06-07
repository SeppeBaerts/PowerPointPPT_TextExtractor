using Extractor_Engine_Service.Extractors;

namespace PowerPointTextExtractor
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Hello, World!");
            var link = Console.ReadLine().Trim('\"');
            Console.WriteLine($"You entered: {link}");
            PptExtractor extractor = new();
            var text = extractor.Extract(link);
        }
    }
}
