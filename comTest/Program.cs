using System;

namespace comTest
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                Console.WriteLine("Hello World!");

                var type = Type.GetTypeFromProgID("Loader.Application");
                Console.WriteLine(type.GUID.ToString());

                using (dynamic instance = COMObject.CreateObject("Loader.Application"))
                {
                    Console.WriteLine("import");
                    instance.Import(@"C:\OKB\CDSFiles\14\08072021221455\TST_01739_20210707_02_001.cds", @"C:\OKB\LoaderLight");
                    using (dynamic doc = new COMObject(instance.ActiveDocument))
                    {
                        string nm = doc.Name;
                        Console.WriteLine("Imported CDS to " + nm);

                        if (doc.Verify())
                        {
                            Console.WriteLine("CDS Verified");

                            Console.WriteLine("SummaryReportPath: " + doc.SummaryReportPath);
                            Console.WriteLine("CDS Verified: " + doc.StatisticalReportPath);
                            Console.WriteLine("CDS Verified: " + doc.DetailedReportPath);
                            Console.WriteLine("NumberOfCriticalErrors: " + doc.NumberOfCriticalErrors);
                        }
                    }

                    Console.WriteLine("imported");
                }
            }
            catch (Exception e)
            { Console.WriteLine(e.Message); }
            Console.ReadKey();
        }
    }
}