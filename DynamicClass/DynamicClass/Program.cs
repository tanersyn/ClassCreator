using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Xml.Linq;
using IronXL;
using NPOI.SS.UserModel;
using System.Linq;
using NPOI.HPSF;
using SixLabors.ImageSharp.Drawing;

namespace DynamicClass
{
    public class Program: SystemException
    {
        public static void Main(string[] args)
        {

            var path = "your_path";


            String filePath = $"{@path}";
            WorkBook workBook = WorkBook.Load(filePath);
            WorkSheet sheet = workBook.GetWorkSheet("Sheet1");

            var datatable = new DataTable("tblData");
            datatable.Columns.AddRange(new[]
            {
                new DataColumn("nameSpace", typeof (string)), new DataColumn("ClassType", typeof (string)), new DataColumn("ClassName", typeof (string)),
                 new DataColumn("Type", typeof (string)), new DataColumn("VariableType", typeof (string)),new DataColumn("variable_name", typeof (string))
            });

            

            int row1 = 0;
            var row = datatable.NewRow();
            foreach (var cell in sheet) //["A2:A8"]
            {

                if (row1 > 5)
                {
                    row1 = 0;
                }
                row[row1] = cell.Text;

                if (row1 == 5)
                { 
                datatable.Rows.Add(row);
                    row = datatable.NewRow();
                }
                row1 = row1 + 1;
            }


            List<Entitytable> excellist = new List<Entitytable>();

            for (int k =0; k < datatable.Rows.Count; k++)
            {

                Entitytable entity = new Entitytable();
                entity.nameSpace = datatable.Rows[k]["nameSpace"].ToString();
                entity.ClassType = datatable.Rows[k]["ClassType"].ToString();
                entity.ClassName = datatable.Rows[k]["ClassName"].ToString();
                entity.Type = datatable.Rows[k]["Type"].ToString();
                entity.VariableType = datatable.Rows[k]["VariableType"].ToString();
                entity.variable_name = datatable.Rows[k]["variable_name"].ToString();
                
               excellist.Add(entity);   


            }

            foreach(var item in excellist)
            {
              Console.WriteLine(item.nameSpace);
            }

            var gruppedlist = excellist.GroupBy(u => u.ClassName).ToDictionary(x=>x);


            foreach (var group in gruppedlist)
            {
                Console.WriteLine(group);
                int ss = 0;
                string open_brackets = "{";
                string close_brackets = "}";
                string get_set = "{ get; set; }";


                int counter = 0;
                int count = group.Value.Count();

                foreach (var user in group.Value)
                {
                    
                    if (ss==0)
                    {
                        using (StreamWriter sw = File.AppendText($"{path}/{user.ClassName}.cs"))
                        {
                            sw.WriteLine($"using System;\r\nusing System.IO;\n\nnamespace {user.nameSpace}\r\n{open_brackets} \r\n\t public class {user.ClassName} \r\n\t {open_brackets}");
                        }
                        ss = 1;
                    }

                    using (StreamWriter sw = File.AppendText($"{path}/{user.ClassName}.cs"))
                    {
                        sw.WriteLine($"  \t\t {user.ClassType} {user.VariableType} {user.variable_name} {get_set}\r\n \t");
                    }

                    counter++;

                    if(counter == count)
                    {
                        using (StreamWriter sw = File.AppendText($"{path}/{user.ClassName}.cs"))
                        {
                            sw.WriteLine($"\t {close_brackets}");
                            sw.WriteLine($"{close_brackets}");
                        }
                    }

                }

                ss = 1;
                //tw.Close();
            }
            //DataRow[] groupTable1 = datatable.Select("ClassName = 'User'");

        }

    }
}
