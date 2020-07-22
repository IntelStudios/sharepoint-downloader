using System;
using System.Diagnostics;
using System.Threading.Tasks;

using SharepointDownloader.Sharepoint;

namespace SharepointDownloader
{
    class Program
    {
        static async Task Main(string[] args)
        {
            var context = new Sharepoint365Context(new Sharepoint365Options
            {
                Login = "",
                Password = "",
                RootFolder = "",
                Url = ""
            });

            var stopwatch = Stopwatch.StartNew();
            var files = await context.GetFilesRecursivelyWithoutData();
            Console.WriteLine($"List files in {stopwatch.ElapsedMilliseconds}ms");


            stopwatch = Stopwatch.StartNew();
            Parallel.ForEach(files, file =>
            {
                Debug.WriteLine($"FileName: {file.FileName}, created: {file.DateCreated}, modified: {file.DateModified}");
                file.FileData = context.LoadFile(file.FileStorageIdentifier).GetAwaiter().GetResult();
            });
            Console.WriteLine($"Downloading files data in {stopwatch.ElapsedMilliseconds}ms");

            Console.WriteLine("Press any key to exit");
            Console.ReadKey();
        }   
    }       
}           
            