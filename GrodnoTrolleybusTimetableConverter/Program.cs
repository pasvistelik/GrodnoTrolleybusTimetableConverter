using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace GrodnoTrolleybusTimetableConverter
{
    static class Program
    {
        /// <summary>
        /// Главная точка входа для приложения.
        /// </summary>
        [STAThread]
        static void Main()
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Title = "Выберите файлы с расписанием";
            ofd.Filter =  "Расписание в формате xls|*.xls";
            ofd.InitialDirectory = @"D:\Files\Other\Projects\PublicTransport\Converters\GrodnoTrolleybusTimetable";//AppDomain.CurrentDomain.BaseDirectory;
            ofd.Multiselect = true;
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                List<Thread> threads = new List<Thread>();
                for (int i = 0, n = ofd.FileNames.Length, processorCount = Environment.ProcessorCount; i < processorCount; i++)
                {
                    Thread tr = new Thread(delegate ()
                    {
                        for (int j = i; j < n; j += processorCount)
                        {
                            Converter.Convert(ofd.FileNames[j]);
                        }
                    });
                    
                    threads.Add(tr);
                    tr.Start();
                    tr.Join();
                }
                //foreach (Thread tr in threads) tr.Join();
                //List<string> timetablesJSON = new List<string>();
                /*foreach (string filepath in ofd.FileNames)
                {
                    //MessageBox.Show(filepath);
                    Converter.Convert(filepath);
                    //timetablesJSON.Add(Converter.Convert(filepath));
                }*/
            }
        }
    }
}
