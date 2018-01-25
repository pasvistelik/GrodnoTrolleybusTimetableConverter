using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Collections.ObjectModel;
using Newtonsoft.Json;
using System.Runtime.InteropServices;
using TransportClasses;
using System.Dynamic;

namespace GrodnoTrolleybusTimetableConverter
{
    static class TrolleybusesTimetableOfOperarorConverter
    {

        private static int ParseTime(SimpleTime item)
        {
            int tmp = (item.hour > 4) ? 0 : 86400;
            return tmp + item.hour * 3600 + item.minute * 60;
        }

        public static dynamic Convert(string filepath)
        {
            List<Timetable>[] timetables = null;

            dynamic convertation_result = new ExpandoObject();
            convertation_result.transport_company_name = "Гродненское троллейбусное управление";
            convertation_result.area_name = "Гродно";
            convertation_result.routes = new List<dynamic>();
            
            
            
            Excel.Application ObjWorkExcel = new Excel.Application(); //открыть эксель
            Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(filepath, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing); //открыть файл


            for (int k = 1, n = ObjWorkBook.Sheets.Count; k <= n; k++)
            {
                convertation_result.routes.Add(ParseWorkSheet(ObjWorkBook.Sheets[k]));
            }

            
            ObjWorkBook.Close(false, Type.Missing, Type.Missing); //закрыть не сохраняя
            ObjWorkExcel.Quit(); // выйти из экселя
            Marshal.ReleaseComObject(ObjWorkBook);
            Marshal.ReleaseComObject(ObjWorkExcel);
            ObjWorkBook = null;
            ObjWorkExcel = null;
            GC.Collect(); // убрать за собой

            return convertation_result;
        }









        public static dynamic ParseWorkSheet(Excel.Worksheet ObjWorkSheet)
        {
            Excel.Range cell = ObjWorkSheet.Cells;
            

            var lastCell = ObjWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);//1 ячейку
            int allRows = lastCell.Row, allCollumns = lastCell.Column;
            int next_fragment = 0;
            for (int i = 2, n = allRows; i < n; i++)
            {
                if (cell[i, 1].Text != string.Empty) {
                    next_fragment = i + 1;
                    break;
                }
            }
            if (next_fragment == 0) throw new Exception();

            dynamic convertation_result_route = new ExpandoObject();
            convertation_result_route.route_type = "trolleybus";
            convertation_result_route.route_number = ((string)(cell[6, 2].Text)).Trim();
            convertation_result_route.route_name = ((string)(cell[5, 2].Text)).Trim();
            convertation_result_route.ways = new List<dynamic>();




            List<Timetable>[] fullTable = new List<Timetable>[2];
            List<Timetable>[] fullTable_depo = new List<Timetable>[2];


            List<DayOfWeek> workingDays = new List<DayOfWeek>();
            workingDays.AddRange(new DayOfWeek[] { DayOfWeek.Monday, DayOfWeek.Tuesday, DayOfWeek.Wednesday, DayOfWeek.Thursday, DayOfWeek.Friday });
            List<DayOfWeek> weekDays = new List<DayOfWeek>();
            weekDays.AddRange(new DayOfWeek[] { DayOfWeek.Saturday, DayOfWeek.Sunday });

            Regex timePattern = new Regex("[0-9]{2}");
            List<Timetable>[] timetables = new List<Timetable>[2];

            for (int k = 1; k <= 2; k++)
            {
                fullTable[k - 1] = new List<Timetable>();
                fullTable_depo[k - 1] = new List<Timetable>();

                timetables[k - 1] = new List<Timetable>();


                int firstRow = (k == 1) ? 2 : next_fragment;
                int lastRow = (k == 1) ? next_fragment - 2 : allRows;

                

                dynamic main_way = new ExpandoObject();
                dynamic way_to_depo = new ExpandoObject();
                dynamic way_from_depo = new ExpandoObject();

                string first_station_name = (string)(cell[firstRow, 3].Text);
                string last_station_name = (string)(cell[lastRow - 8, 3].Text);

                main_way.way_name = first_station_name + " - " + last_station_name;
                way_to_depo.way_name = first_station_name + " - Троллейбусное депо";
                way_from_depo.way_name = "Троллейбусное депо - " + last_station_name;

                main_way.stations_names = new List<string>();
                way_to_depo.stations_names = new List<string>();
                way_from_depo.stations_names = new List<string>();

                main_way.trips_by_days = new dynamic[2];
                way_to_depo.trips_by_days = new dynamic[2];
                way_from_depo.trips_by_days = new dynamic[2];

                main_way.trips_by_days[0] = new ExpandoObject();
                main_way.trips_by_days[1] = new ExpandoObject();
                main_way.trips_by_days[0].days_of_week = workingDays;
                main_way.trips_by_days[1].days_of_week = weekDays;
                main_way.trips_by_days[0].arrives = new List<List<int>>();
                main_way.trips_by_days[1].arrives = new List<List<int>>();

                way_to_depo.trips_by_days[0] = new ExpandoObject();
                way_to_depo.trips_by_days[1] = new ExpandoObject();
                way_to_depo.trips_by_days[0].days_of_week = workingDays;
                way_to_depo.trips_by_days[1].days_of_week = weekDays;
                way_to_depo.trips_by_days[0].arrives = new List<List<int>>();
                way_to_depo.trips_by_days[1].arrives = new List<List<int>>();

                way_from_depo.trips_by_days[0] = new ExpandoObject();
                way_from_depo.trips_by_days[1] = new ExpandoObject();
                way_from_depo.trips_by_days[0].days_of_week = workingDays;
                way_from_depo.trips_by_days[1].days_of_week = weekDays;
                way_from_depo.trips_by_days[0].arrives = new List<List<int>>();
                way_from_depo.trips_by_days[1].arrives = new List<List<int>>();





                for (int startString = firstRow; startString < lastRow; startString += 9) // по всем строкам
                {
                    bool is_to_depo_finded = false;



                    ObservableCollection<SimpleTime> workingTimes = new ObservableCollection<SimpleTime>();
                    ObservableCollection<SimpleTime> workingTimesToDepo = new ObservableCollection<SimpleTime>();
                    if (Regex.IsMatch(cell[startString + 4, 3].Text, "^5"))
                    {
                        //workingTimes.Add(new SimpleTime(5, int.Parse((timePattern.Match(cell[startString + 4, 3].Text)).Value)));
                        //MessageBox.Show(workingTimes[workingTimes.Count-1].ToString());
                        for (int i = startString + 4; i <= startString + 6; i++)
                        {
                            foreach (Match s in timePattern.Matches(cell[i, 3].Text))
                            {
                                workingTimes.Add(new SimpleTime(5, int.Parse(s.Value)));
                                //MessageBox.Show(workingTimes[workingTimes.Count - 1].ToString());
                            }
                        }
                    }
                    for (int j = 4, hour = (j + 2) % 24; j <= 22; j++, hour = (j + 2) % 24)
                    {
                        for (int i = startString + 4; i <= startString + 6; i++)
                        {
                            if (!timePattern.IsMatch(cell[i, j].Text)) continue;

                            //cell[i, j]//.NumberFormat = "Text";
                            Excel.Range tmp_cell = (cell[i, j] as Excel.Range);

                            //var tmp00 = tmp_cell.Characters;

                            //try
                            //{
                            for (int q = 1; q < tmp_cell.Value.ToString().Length; q += 3)
                            {
                                string tm0000 = null;
                                try
                                {
                                    tm0000 = tmp_cell.Characters[q, 2].Text;
                                }
                                catch
                                {
                                    tm0000 = tmp_cell.Value.ToString();
                                }

                                if (tmp_cell.Characters[q, 2].Font.Color == 0)
                                {
                                    //MessageBox.Show("В депо: " + (cell.Cells[i, j] as Excel.Range).Characters[q, 2].Text);
                                    workingTimesToDepo.Add(new SimpleTime(hour, int.Parse(tm0000)));
                                    //MessageBox.Show(workingTimesToDepo[workingTimesToDepo.Count - 1].ToString());
                                    is_to_depo_finded = true;
                                }
                                else
                                {
                                    //MessageBox.Show((cell.Cells[i, j] as Excel.Range).Characters[q, 2].Font.Color);
                                    //MessageBox.Show("Не в депо: " + (cell.Cells[i, j] as Excel.Range).Characters[q, 2].Text);
                                    workingTimes.Add(new SimpleTime(hour, int.Parse(tm0000)));
                                    //MessageBox.Show(workingTimes[workingTimes.Count - 1].ToString());
                                }
                            }
                            /*}
                            catch(Exception ex)
                            {
                                foreach (Match s in timePattern.Matches(cell[i, j].Text))
                                {
                                    workingTimes.Add(new SimpleTime(hour, int.Parse(s.Value)));
                                    //MessageBox.Show(workingTimes[workingTimes.Count - 1].ToString());
                                }
                            }*/
                        }
                    }
                    Table workingTable = new Table(workingDays, workingTimes);
                    Table workingTableToDepo = new Table(workingDays, workingTimesToDepo);

                    ObservableCollection<SimpleTime> weekTimes = new ObservableCollection<SimpleTime>();
                    ObservableCollection<SimpleTime> weekTimesToDepo = new ObservableCollection<SimpleTime>();
                    if (Regex.IsMatch(cell[startString + 7, 3].Text, "^5"))
                    {
                        //weekTimes.Add(new SimpleTime(5, int.Parse((timePattern.Match(cell[startString + 7, 3].Text)).Value)));
                        //MessageBox.Show(weekTimes[weekTimes.Count - 1].ToString());
                        for (int i = startString + 7; i <= startString + 8; i++)
                        {
                            foreach (Match s in timePattern.Matches(cell[i, 3].Text))
                            {
                                weekTimes.Add(new SimpleTime(5, int.Parse(s.Value)));
                                //MessageBox.Show(weekTimes[weekTimes.Count - 1].ToString());
                            }
                        }
                    }
                    for (int j = 4, hour = (j + 2) % 24; j <= 22; j++, hour = (j + 2) % 24)
                    {
                        for (int i = startString + 7; i <= startString + 8; i++)
                        {
                            if (!timePattern.IsMatch(cell[i, j].Text)) continue;

                            Excel.Range tmp_cell = (cell[i, j] as Excel.Range);
                            //try
                            //{
                            for (int q = 1; q < (cell.Cells[i, j] as Excel.Range).Characters.Count; q += 3)
                            {
                                string tm0000 = null;
                                try
                                {
                                    tm0000 = tmp_cell.Characters[q, 2].Text;
                                }
                                catch
                                {
                                    tm0000 = tmp_cell.Value.ToString();
                                }

                                if ((cell.Cells[i, j] as Excel.Range).Characters[q, 2].Font.Color == 0)
                                {
                                    //MessageBox.Show("В депо: " + (cell.Cells[i, j] as Excel.Range).Characters[q, 2].Text);
                                    weekTimesToDepo.Add(new SimpleTime(hour, int.Parse(tm0000)));
                                    //MessageBox.Show(weekTimesToDepo[weekTimesToDepo.Count - 1].ToString());
                                    is_to_depo_finded = true;
                                }
                                else
                                {
                                    //MessageBox.Show("Не в депо: " + (cell.Cells[i, j] as Excel.Range).Characters[q, 2].Text);
                                    weekTimes.Add(new SimpleTime(hour, int.Parse(tm0000)));
                                    //MessageBox.Show(weekTimes[weekTimes.Count - 1].ToString());
                                }
                            }
                            /*}
                            catch
                            {
                                foreach (Match s in timePattern.Matches(cell[i, j].Text))
                                {
                                    weekTimes.Add(new SimpleTime(hour, int.Parse(s.Value)));
                                    //MessageBox.Show(weekTimes[weekTimes.Count - 1].ToString());
                                }
                            }*/
                        }
                    }
                    Table weekTable = new Table(weekDays, weekTimes);
                    Table weekTableToDepo = new Table(weekDays, weekTimesToDepo);

                    ObservableCollection<Table> t = new ObservableCollection<Table>();
                    t.Add(workingTable);
                    t.Add(weekTable);
                    Timetable tbl = new Timetable(null, null, TableType.table, t);
                    fullTable[k - 1].Add(tbl);

                    timetables[k - 1].Add(tbl);

                    //SW.Write(JsonConvert.SerializeObject(tbl));
                    /*SW.Write(tbl.Serialize());
                    SW.Close();*/

                    ObservableCollection<Table> t2 = new ObservableCollection<Table>();
                    t2.Add(workingTableToDepo);
                    t2.Add(weekTableToDepo);
                    Timetable tbl2 = new Timetable(null, null, TableType.table, t2);
                    fullTable_depo[k - 1].Add(tbl2);

                    timetables[k - 1].Add(tbl2);

                    //SW_depo.Write(JsonConvert.SerializeObject(tbl2));
                    /*SW_depo.Write(tbl2.Serialize());
                    SW_depo.Close();*/


                    main_way.stations_names.Add((string)(cell[startString, 3].Text));
                    if (is_to_depo_finded) way_to_depo.stations_names.Add((string)(cell[startString, 3].Text));
                    else way_from_depo.stations_names.Add((string)(cell[startString, 3].Text));
                }




                for (int i = 0, n = fullTable_depo[k - 1].Count; i < n; i++)
                {
                    if (fullTable_depo[k - 1][i].table[0].times.Count > 0)
                    {
                        List<int> tmp0 = new List<int>();
                        foreach (var item in fullTable_depo[k - 1][i].table[0].times)
                        {
                            tmp0.Add(ParseTime(item));
                        }
                        way_to_depo.trips_by_days[0].arrives.Add(tmp0);
                    }
                    if (fullTable_depo[k - 1][i].table[1].times.Count > 0)
                    {
                        List<int> tmp1 = new List<int>();
                        foreach (var item in fullTable_depo[k - 1][i].table[1].times)
                        {
                            tmp1.Add(ParseTime(item));
                        }
                        way_to_depo.trips_by_days[1].arrives.Add(tmp1);
                    }
                }

                List<int> tmp0_from_depo_indexes = new List<int>();
                List<int> tmp1_from_depo_indexes = new List<int>();
                for (int i = 0, n = fullTable[k - 1].Count; i < n; i++)
                {
                    if (fullTable[k - 1][i].table[0].times.Count > 0)
                    {
                        List<int> tmp0 = new List<int>();
                        List<int> tmp0_from_depo = new List<int>();
                        if (i == 0)
                        {
                            foreach (var item in fullTable[k - 1][i].table[0].times)
                            {
                                tmp0.Add(ParseTime(item));
                            }
                        }
                        else
                        {
                            for (int j = 0, m = fullTable[k - 1][i].table[0].times.Count, t = 0; j < m; j++)
                            {
                                //for (int r = 0; r < ) ;
                                int tmp_int_time = ParseTime(fullTable[k - 1][i].table[0].times[j]);
                                int max_length = main_way.trips_by_days[0].arrives[i - 1].Count;
                                //ParseTime(fullTable[k - 1][i - 1].table[0].times[j - t0])
                                if (j - t < max_length && main_way.trips_by_days[0].arrives[i - 1][j - t] <= tmp_int_time + 60 /*!!!!!!!*/ && !tmp0_from_depo_indexes.Contains(j))
                                {
                                    if (main_way.trips_by_days[0].arrives[i - 1][j - t] < tmp_int_time) tmp0.Add(tmp_int_time);
                                    else tmp0.Add(tmp_int_time + 60);//!!!!!
                                }
                                else
                                {
                                    tmp0_from_depo.Add(tmp_int_time);
                                    t++;
                                    tmp0_from_depo_indexes.Add(j);
                                }
                            }
                        }
                        main_way.trips_by_days[0].arrives.Add(tmp0);
                        if (tmp0_from_depo.Count > 0) way_from_depo.trips_by_days[0].arrives.Add(tmp0_from_depo);
                    }

                    if (fullTable[k - 1][i].table[1].times.Count > 0)
                    {
                        List<int> tmp1 = new List<int>();
                        List<int> tmp1_from_depo = new List<int>();
                        if (i == 0)
                        {
                            foreach (var item in fullTable[k - 1][i].table[1].times)
                            {
                                tmp1.Add(ParseTime(item));
                            }
                        }
                        else
                        {
                            for (int j = 0, m = fullTable[k - 1][i].table[1].times.Count, t = 0; j < m; j++)
                            {
                                //for (int r = 0; r < ) ;
                                int tmp_int_time = ParseTime(fullTable[k - 1][i].table[1].times[j]);
                                int max_length = main_way.trips_by_days[1].arrives[i - 1].Count;
                                //ParseTime(fullTable[k - 1][i - 1].table[0].times[j - t0])
                                if (j - t < max_length && main_way.trips_by_days[1].arrives[i - 1][j - t] <= tmp_int_time + 60 /*!!!!!!!*/ && !tmp1_from_depo_indexes.Contains(j))
                                {
                                    if (main_way.trips_by_days[1].arrives[i - 1][j - t] < tmp_int_time) tmp1.Add(tmp_int_time);
                                    else tmp1.Add(tmp_int_time + 60);//!!!!!
                                }
                                else
                                {
                                    tmp1_from_depo.Add(tmp_int_time);
                                    t++;
                                    tmp1_from_depo_indexes.Add(j);
                                }
                            }
                        }
                        main_way.trips_by_days[1].arrives.Add(tmp1);
                        if (tmp1_from_depo.Count > 0) way_from_depo.trips_by_days[1].arrives.Add(tmp1_from_depo);
                    }
                }


                (convertation_result_route.ways as List<dynamic>).AddRange(new dynamic[] { main_way, way_from_depo, way_to_depo });
            }




            return convertation_result_route;
        }

    }
}
