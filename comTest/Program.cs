using Dynamitey;
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

                Console.WriteLine(type.ToString());
                var Methods = type.GetMethods();
                Console.WriteLine("Methods: " + Methods.Length);
                for (int i = 0; i < Methods.Length; i++)
                { Console.WriteLine(Methods.GetValue(i)); }
                #region types
                var types = type.GetNestedTypes();
                foreach (var t in types)
                {
                    Console.WriteLine(t.Name);
                    Console.WriteLine(t.GetType());
                }
                #endregion


                //dynamic instance = Activator.CreateInstance(type, true);
                //instance.Import(@"C:\OKB\CDSFiles\10\06072021161639\TST_01878_20210705_02_001.cds", @"C:\OKB\LoaderLight");
                //var it = instance.GetType();
                //Console.WriteLine("instance: " + it);

                //using (dynamic instance = Activator.CreateInstance(type, true))
                //{
                //    instance.Import(@"C:\OKB\CDSFiles\10\06072021161639\TST_01878_20210705_02_001.cds", @"C:\OKB\LoaderLight");

                //};

                using (dynamic instance = COMObject.CreateObject("Loader.Application"))
                {
                    Console.WriteLine("import");
                    instance.Import(@"C:\OKB\CDSFiles\14\08072021221455\TST_01739_20210707_02_001.cds", @"C:\OKB\LoaderLight");
                    using (dynamic doc = new COMObject(instance.ActiveDocument))
                    {
                        string nm = doc.Name;
                        Console.WriteLine("Imported CDS to " + nm);

                        //dynamic verifyValue = doc.Verify();
                        //Console.WriteLine(verifyValue);
                        //Console.WriteLine("SummaryReportPath: " + doc.SummaryReportPath);
                        //Console.WriteLine("CDS Verified: " + doc.StatisticalReportPath);
                        //Console.WriteLine("CDS Verified: " + doc.DetailedReportPath);
                        //Console.WriteLine("NumberOfCriticalErrors: " + doc.NumberOfCriticalErrors);

                        if (doc.Verify())
                        {
                            Console.WriteLine("CDS Verified");

                            Console.WriteLine("SummaryReportPath: " + doc.SummaryReportPath);
                            Console.WriteLine("CDS Verified: " + doc.StatisticalReportPath);
                            Console.WriteLine("CDS Verified: " + doc.DetailedReportPath);
                            Console.WriteLine("NumberOfCriticalErrors: " + doc.NumberOfCriticalErrors);
                        }
                    }

                    //if (loaderL.ActiveDocument.Verify())
                    //{
                    //    _logger.LogInformation("CDS Verified");
                    //    var reportList = new List<ValidationReport>()
                    //            {
                    //                new ValidationReport { ExportFileId = cdsFile.ExportFileId, TypeId = ValidationReportTypeEnum.Summary, FilePath = (string)loaderL.ActiveDocument.SummaryReportPath},
                    //                new ValidationReport { ExportFileId = cdsFile.ExportFileId, TypeId = ValidationReportTypeEnum.Statistic, FilePath = (string)loaderL.ActiveDocument.StatisticalReportPath},
                    //                new ValidationReport { ExportFileId = cdsFile.ExportFileId, TypeId = ValidationReportTypeEnum.Detail, FilePath = (string)loaderL.ActiveDocument.DetailedReportPath}
                    //            };


                    ////dynamic inv = ((object)instance.ActiveDocument).GetType().GUID.ToString();
                    //dynamic inv = instance.ActiveDocument.Name;
                    ////dynamic nm = Dynamic.InvokeGetChain(inv, "Name");
                    //Console.WriteLine(inv);

                    //loaderL.ActiveDocument.Name;
                    Console.WriteLine("imported");
                    //var tps = ((COMObject)instance).GetTypes();
                    //foreach(var tp in tps)
                    //{ Console.WriteLine(tp); }
                    //GC.SuppressFinalize(instance);
                }
                //var it = instance.GetType();
                // Console.WriteLine("instance: " + it.ToString());
            }
            catch (Exception e)
            { Console.WriteLine(e.Message); }
            Console.ReadKey();
        }
    }
}