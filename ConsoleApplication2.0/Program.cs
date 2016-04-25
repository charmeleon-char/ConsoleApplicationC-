using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.IO;
using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.text.html.simpleparser;
using System.Data;
using System.Drawing;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;
using System.Threading;
using MongoDB.Driver;
using MongoDB.Bson;
using Mandrill;
using ConsoleApplication2._0;
namespace ConsoleApplication1
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("---------------------------CreateAndSendingReports-------------------------");
            Console.WriteLine("starting...................................................................");
            Program anInstanceofClass = new Program();
            anInstanceofClass.core();


        }

        public class test
        {
        
            public string Nombre { get; set; }
            public DateTime? FechaEntradaMonday { get; set; }
            public DateTime? FechaEntradaTuesday { get; set; }
            public DateTime? FechaEntradaWendnsday { get; set; }
            public DateTime? FechaEntradaThursday { get; set; }
            public DateTime? FechaEntradaFriday { get; set; }
            public DateTime? FechaEntradaSaturday { get; set; }
            public DateTime? FechaEntradaSunday { get; set; }
            public DateTime? FechaSalidaMonday { get; set; }
            public DateTime? FechaSalidaTuesday { get; set; }
            public DateTime? FechaSalidaWendnsday { get; set; }
            public DateTime? FechaSalidaThursday { get; set; }
            public DateTime? FechaSalidaFriday { get; set; }
            public DateTime? FechaSalidaSaturday { get; set; }
            public DateTime? FechaSalidaSunday { get; set; }
            public Decimal? HorasMonday { get; set; }
            public Decimal? HorasTuesday { get; set; }
            public Decimal? HorasWendnsday { get; set; }
            public Decimal? HorasThursday { get; set; }
            public Decimal? HorasFriday { get; set; }
            public Decimal? HorasSaturday { get; set; }
            public Decimal? HorasSunday { get; set; }
        }


        public class WorkingPerDay
        {
            public string Name { get; set; }
            public DateTime? HourInput { get; set; }
            public DateTime? HourOutput { get; set; }
            //    public Decimal? workingHours { get; set; }
        }

        public void sendEmail(string monday, string name, string dept_name, string to, string path)
        {

            if (to != null)
            {
                //  to = "rodolfo.vesely@laureate.net"; //To address    
              // to = "rayelmejor1@gmail.com"; //From address    \
                // to = "Jose.Posas@laureate.net";
                string subject = "Sending " + name + " email with a asistance report  " + dept_name;
                string mailbody = "Hello " + name + " this is your assistance report the past week " + dept_name;

                try
                {

                    //   Attachment att = new Attachment(@"C:\\Pdf\\Asistance_Report_"+ monday + path + dept_name + "_" + monday + ".pdf");
                    Attachment att = new Attachment(@"C:\\Pdf\\Asistance_Report_" + monday + path + "Asistance_Report_" + dept_name + "_" + monday + ".pdf");

                    //   Attachment att = new Attachment(@"C:\\Pdf\\Asistance_Report_" + monday + "\\SubDepts\\" + dept_name + "_" + monday + ".pdf");
                    MailMessage mail = new MailMessage("no-reply@laureate.net", to, subject, mailbody);
                    mail.IsBodyHtml = true;
                    mail.SubjectEncoding = Encoding.Default;
                    mail.BodyEncoding = Encoding.UTF8;
                    mail.Attachments.Add(att);

                    SmtpClient client = new SmtpClient();

                    client.Host = "smtp.mandrillapp.com";
                    client.Port = int.Parse("587");
                    client.DeliveryMethod = SmtpDeliveryMethod.Network;
                    client.UseDefaultCredentials = false;
                
                    try
                    {
                        Console.WriteLine("SendingEmailis " + dept_name + " for " + name + " ..................................");
                        client.Send(mail);
                    }
                    catch (Exception ex)
                    {

                        Console.WriteLine("SendingEmailisError " + ex + "!!! with " + to + "..............................");
                        throw ex;
                    }
                }
                catch
                {
                }
            }

        }
        public class boos
        {
            public string name { get; set; }
            public string correo { get; set; }
            public string departamento { get; set; }
            public string subdept { get; set; }
            public string deptID { get; set; }
        }
        public void core()
        {
            DateTime now = DateTime.Now;
            var oneweekAffter = now.AddDays(7);
            DateTime day = DateTime.Today;
            int offset = day.DayOfWeek - DayOfWeek.Monday;
            DateTime lastmonday = Convert.ToDateTime(day.AddDays(-offset - 7));
            
            
            string monday = lastmonday.ToString("MM-dd-yyyy");
            string path = "";
            //  monday = ("12-14-2015");
            Anviz_Data_BaseEntities2 dc = new Anviz_Data_BaseEntities2();
 
            var client = new MongoClient();
            var db = client.GetDatabase("admin");
            
            var collection = db.GetCollection<infoDept>("managers");
            var document = db.GetCollection<Getreport>("managers");
            var leads = from l in document.AsQueryable<Getreport>()
                        select l.deptID;

            var test = from d in collection.AsQueryable<infoDept>()
                       select d.Id_Dept;
            var all = from e in dc.Userinfoes select new GetPeople { userid = e.Userid, name = e.Name, email = e.Address };
            string folder = @"C:\\Pdf\\Asistance_Report_" + monday;

            Console.WriteLine("Conecting with Mongodb................................");
            Thread.Sleep(2300);
            try
            {
                if (test.Count() <= 1)
                {
                    Console.WriteLine("Error!!!: cannot conect with MongoDb  .................");
                    Console.ReadLine();
                    return;


                }

            }
            catch (Exception e)
            {
                Console.WriteLine("Error!!!: cannot conect with MongoDb  "+e+".................");
                Console.ReadLine();
                return;

            }

            Console.WriteLine("status conection::: success with MongoDb................................");
            Console.WriteLine("Conecting with Mssql................................");
            Thread.Sleep(2300);
            try{
            var whynot = all.Count() <= 1;
            }
            catch (Exception e)
            {
                Console.WriteLine("Error!!!: cannot conect with Mssql!!!!  ................................");
                Console.WriteLine(e);
                Console.ReadLine();
                return;

            }
            Console.WriteLine("status conection::: success with Mssql................................");
            path = @"C:\\Pdf\\Asistance_Report_" + monday + "\\Depts\\";
            bool folderExists = Directory.Exists((folder));
            if (!folderExists)
                Directory.CreateDirectory((folder));
            folderExists = Directory.Exists((folder));
            folderExists = Directory.Exists((path));
            if (!folderExists)
                Directory.CreateDirectory((path));
            folderExists = Directory.Exists((path));
            CreatePdf(test, collection, monday, true, path);
            foreach (var items in test)
            {
                var managers = collection
                      .Find(a => a.Id_Dept == items)
                      .ToListAsync()
                       .Result;
                foreach (var i in managers)
                {
                    string[] sub_managers = i.Name;
                    List<Getreport> get = new List<Getreport>(i.GetReport);
                    path = @"C:\\Pdf\\Asistance_Report_" + monday + "\\SubDepts";
                    foreach (var sub in get)
                    {

                        //      allPDF(monday, sub.deptID, sub.subdept, monday);
                        Console.WriteLine("CreateReportforSubDept" + sub.subdept + "...........................................");
                        CreatePdf(sub.subdept, sub.deptID, monday, true, path);
                       sendEmail(monday, sub.subName, sub.subdept, sub.subEmail, "\\SubDepts\\");


                    }

                    for (int y = 0; y < sub_managers.Length; y++)
                    {





                        //    allPDF(monday, sub.deptID, sub.subdept, monday);
                        //  sendEmail(monday, i.Name[y], sub.subdept,i.correo[y]);
                        // sendEmail(monday, sub.subName, sub.subdept,sub.subEmail);



                        // allPDF(monday, i.Id_Dept, i.Nombre_Dept, monday);

                        //   sendEmail(monday, i.Name[y], i.Nombre_Dept,i.correo[y]);
                         sendEmail(monday, i.Name[y], "All" + i.Nombre_Dept, i.correo[y], "\\Depts\\");
                        //if (oneweekAffter.Month != now.Month)
                        //{

                        //    Console.WriteLine("ultimo lunes ................................");
                        //}
                        //else
                        //{

                        //    Console.WriteLine(" no es el ultimo lunes b................................");
                        //}

                    }

                }

            }

           sendEmail(monday, "Flor de Liz Reyes", "LNO-HN", "flor.reyes@laureate.net", "\\Depts\\");
          sendEmail(monday, "Flor de Liz Reyes", "LNO-HN", "gina.almendares@laureate.net", "\\Depts\\");
            path = @"C:\\Pdf\\Asistance_Report_" + monday + "\\Employes\\";
            foreach (var items in all)
            {
               CreatePdf(items.name, items.userid, monday, false, path);
                sendEmail(monday, items.name, items.name, items.email, "\\Employes\\");
            }
    
        }


        public class infoDept
        {
            public ObjectId Id { get; set; }
            public string Nombre_Dept { get; set; }
            public string[] Name { get; set; }
            public string[] correo { get; set; }
            public string Id_Dept { get; set; }
            public Getreport[] GetReport { get; set; }
        }

        public class Getreport
        {
            public ObjectId Id { get; set; }
            public string subdept { get; set; }
            public string deptID { get; set; }
            public string subEmail { get; set; }
            public string subName { get; set; }
        }

        public class GetPeople
        {
            public string name { get; set; }
            public string userid { get; set; }
            public string email { get; set; }


        }

        public class GetIds
        {
            public string Id { get; set; }
        }


        public PdfPTable archive(PdfPTable table, int deptID, string subdept, string dates, bool type)
        {
            List<test> all = new List<test>();
            List<test> listWeek1 = new List<test>();
            List<test> listWeek2 = new List<test>();
            List<test> listWeek3 = new List<test>();
            List<test> listWeek4 = new List<test>();
            List<test> potato = new List<test>();
            PdfPCell cell = new PdfPCell();
            PdfPCell cell2 = new PdfPCell();
            PdfPCell cellName = new PdfPCell();
            PdfPCell cellMonday = new PdfPCell();
            PdfPCell cellTuesday = new PdfPCell();
            PdfPCell cellWednesday = new PdfPCell();
            PdfPCell cellThursday = new PdfPCell();
            PdfPCell cellfriday = new PdfPCell();
            PdfPCell cellsaturday = new PdfPCell();
            PdfPCell cellsunday = new PdfPCell();
            Array[] list = { };
            iTextSharp.text.Font georgia = FontFactory.GetFont("georgia", 14f);
            georgia.Color = iTextSharp.text.BaseColor.BLUE;
            iTextSharp.text.Font labels = FontFactory.GetFont("georgia", 14f);
            labels.Color = iTextSharp.text.BaseColor.WHITE;
            iTextSharp.text.Font Headers = FontFactory.GetFont("georgia", 18f);
            Headers.Color = iTextSharp.text.BaseColor.BLACK;

            PdfPTable table2 = new PdfPTable(24);
            table.TotalWidth = 5000f;

            float[] widths = new float[] { 1300f, 1300, 1300f, 1300f, 1300f, 1300f, 1300f, 1300f, 1300f, 1300f, 1300f, 1300f, 1300f, 1300f, 1300f, 1300f, 1300f, 1300f, 1300f, 1300f, 1300f, 1300f, 1300f, 1300f };
            table.SetWidths(widths);
            DateTime week2 = Convert.ToDateTime(dates).AddDays(-7);
            DateTime week3 = Convert.ToDateTime(dates).AddDays(-14);
            DateTime week4 = Convert.ToDateTime(dates).AddDays(-21);
       

            //  var init = Convert.ToInt32(deptID.ToString());
            using (Anviz_Data_BaseEntities2 dc = new Anviz_Data_BaseEntities2())
            {
                //   all = (from e in dc).ToList();
          
 
                if (type == true)
                {
                    all = (from e in dc.sp_matrix_weekdays(deptID, dates) select new test { Nombre = e.name, FechaEntradaMonday = e.Hora_de_LLegada_Lunes, FechaSalidaMonday = e.Hora_de_salida_Lunes, HorasMonday = e.Monday, FechaEntradaTuesday = e.Hora_de_LLegada_Martes, FechaSalidaTuesday = e.Hora_de_salida_Martes, HorasTuesday = e.Tuesday, FechaEntradaWendnsday = e.Hora_de_LLegada_Miercoles, FechaSalidaWendnsday = e.Hora_de_salida_Miercoles, HorasWendnsday = e.Wednsday, FechaEntradaThursday = e.Hora_de_LLegada_Jueves, FechaSalidaThursday = e.Hora_de_salida_Jueves, HorasThursday = e.Thursday, FechaEntradaFriday = e.Hora_de_LLegada_Viernes, FechaSalidaFriday = e.Hora_de_salida_Viernes, HorasFriday = e.Friday, FechaEntradaSaturday = e.Hora_de_LLegada_Sabado, FechaSalidaSaturday = e.Hora_de_salida_Sabado, HorasSaturday = e.Saturday, FechaEntradaSunday = e.Hora_de_LLegada_Domingo, FechaSalidaSunday = e.Hora_de_salida_Domingo, HorasSunday = e.Sunday }).ToList();
                  
                }
                else
                {
                    all = (from e in dc.GetEmployes(deptID, dates) select new test { Nombre = e.name, FechaEntradaMonday = e.Hora_de_LLegada_Lunes, FechaSalidaMonday = e.Hora_de_salida_Lunes, HorasMonday = e.Monday, FechaEntradaTuesday = e.Hora_de_LLegada_Martes, FechaSalidaTuesday = e.Hora_de_salida_Martes, HorasTuesday = e.Tuesday, FechaEntradaWendnsday = e.Hora_de_LLegada_Miercoles, FechaSalidaWendnsday = e.Hora_de_salida_Miercoles, HorasWendnsday = e.Wednsday, FechaEntradaThursday = e.Hora_de_LLegada_Jueves, FechaSalidaThursday = e.Hora_de_salida_Jueves, HorasThursday = e.Thursday, FechaEntradaFriday = e.Hora_de_LLegada_Viernes, FechaSalidaFriday = e.Hora_de_salida_Viernes, HorasFriday = e.Friday, FechaEntradaSaturday = e.Hora_de_LLegada_Sabado, FechaSalidaSaturday = e.Hora_de_salida_Sabado, HorasSaturday = e.Saturday, FechaEntradaSunday = e.Hora_de_LLegada_Domingo, FechaSalidaSunday = e.Hora_de_salida_Domingo, HorasSunday = e.Sunday }).ToList();
                  
                }
              
            }

            DateTime date = Convert.ToDateTime(dates);

            //PdfPCell june = new PdfPCell(new Phrase("Cell 2",georgia));


            cell2 = new PdfPCell(new Phrase(subdept, Headers));
            cell2.HorizontalAlignment = Element.ALIGN_CENTER;
            cell2.BackgroundColor = new iTextSharp.text.BaseColor(245, 92, 24);
            cell2.BorderColor = new iTextSharp.text.BaseColor(0, 0, 0);
            cell2.Colspan = 24;
            table.AddCell(cell2);




            PdfPCell nombre = new PdfPCell(new Phrase("Name", labels));
            nombre.BackgroundColor = new iTextSharp.text.BaseColor(89, 89, 89);
            nombre.BorderColor = new iTextSharp.text.BaseColor(0, 0, 0);
            nombre.BorderWidthRight = 1;
            nombre.Rowspan = 2;
            nombre.Colspan = 2;
            nombre.HorizontalAlignment = 1;
            nombre.VerticalAlignment = 1;
            table.AddCell(nombre);

            PdfPCell june = new PdfPCell(new Phrase("Monday  " + "\n" + Convert.ToDateTime(date).ToShortDateString().ToString(), Headers));
            june.BorderWidthRight = 1;
            june.HorizontalAlignment = 1;
            june.BackgroundColor = new iTextSharp.text.BaseColor(247, 150, 70);
            june.BorderColor = new iTextSharp.text.BaseColor(0, 0, 0);
            june.Colspan = 3;
            table.AddCell(june);
            june = new PdfPCell(new Phrase("Tuesday  " + "\n" + Convert.ToDateTime(date).AddDays(1).ToShortDateString(), Headers));
            june.BorderWidthRight = 1;
            june.BackgroundColor = new iTextSharp.text.BaseColor(247, 150, 70);
            june.BorderColor = new iTextSharp.text.BaseColor(0, 0, 0);
            june.HorizontalAlignment = 1;
            june.Colspan = 3;
            table.AddCell(june);
            june = new PdfPCell(new Phrase("Wednesday  " + "\n" + Convert.ToDateTime(date).AddDays(2).ToShortDateString(), Headers));
            june.BorderWidthRight = 1;
            june.BackgroundColor = new iTextSharp.text.BaseColor(247, 150, 70);
            june.BorderColor = new iTextSharp.text.BaseColor(0, 0, 0);
            june.HorizontalAlignment = 1;
            june.Colspan = 3;
            table.AddCell(june);
            june = new PdfPCell(new Phrase("Thursday  " + "\n" + Convert.ToDateTime(date).AddDays(3).ToShortDateString(), Headers));
            june.BorderWidthRight = 1;
            june.BackgroundColor = new iTextSharp.text.BaseColor(247, 150, 70);
            june.BorderColor = new iTextSharp.text.BaseColor(0, 0, 0);
            june.HorizontalAlignment = 1;
            june.Colspan = 3;
            table.AddCell(june);

            june = new PdfPCell(new Phrase("Friday   " + "\n" + Convert.ToDateTime(date).AddDays(4).ToShortDateString(), Headers));
            june.BorderWidthRight = 1;
            june.BackgroundColor = new iTextSharp.text.BaseColor(247, 150, 70);
            june.BorderColor = new iTextSharp.text.BaseColor(0, 0, 0);
            june.HorizontalAlignment = 1;
            june.Colspan = 3;
            table.AddCell(june);


            june = new PdfPCell(new Phrase("Saturday  " + "\n" + Convert.ToDateTime(date).AddDays(5).ToShortDateString(), Headers));
            june.BorderWidthRight = 1;
            june.BackgroundColor = new iTextSharp.text.BaseColor(247, 150, 70);
            june.BorderColor = new iTextSharp.text.BaseColor(0, 0, 0);
            june.HorizontalAlignment = 1;
            june.Colspan = 3;
            table.AddCell(june);
            june = new PdfPCell(new Phrase("Sunday   " + "\n" + Convert.ToDateTime(date).AddDays(6).ToShortDateString(), Headers));
            june.BorderWidthRight = 1;
            june.BackgroundColor = new iTextSharp.text.BaseColor(247, 150, 70);
            june.BorderColor = new iTextSharp.text.BaseColor(0, 0, 0);
            june.HorizontalAlignment = 1;
            june.Colspan = 3;
            table.AddCell(june);

            june = new PdfPCell(new Phrase("Weekly Hours Worked", labels));
            june.BorderWidthRight = 1;
            june.Rowspan = 2;
            june.HorizontalAlignment = 1;
            june.BackgroundColor = new iTextSharp.text.BaseColor(89, 89, 89);
            june.BorderColor = new iTextSharp.text.BaseColor(0, 0, 0);
            table.AddCell(june);




            PdfPCell In = new PdfPCell();
            PdfPCell OUT = new PdfPCell();
            PdfPCell second = new PdfPCell();

            for (var y = 0; y <= 6; y++)
            {
                In = new PdfPCell(new Phrase("In", labels));

                In.BackgroundColor = new iTextSharp.text.BaseColor(89, 89, 89);
                In.BorderColor = new iTextSharp.text.BaseColor(0, 0, 0);
                In.HorizontalAlignment = 1;
                table.AddCell(In);

                OUT = new PdfPCell(new Phrase("Out", labels));
                OUT.BackgroundColor = new iTextSharp.text.BaseColor(89, 89, 89);
                OUT.BorderColor = new iTextSharp.text.BaseColor(0, 0, 0);
                OUT.HorizontalAlignment = 1;
                table.AddCell(OUT);

                second = new PdfPCell(new Phrase("Worked Hours", labels));
                second.BackgroundColor = new iTextSharp.text.BaseColor(89, 89, 89);
                second.BorderColor = new iTextSharp.text.BaseColor(0, 0, 0);
                second.BorderWidthRight = 1;
                second.HorizontalAlignment = 1;
                table.AddCell(second);


            }

            foreach (var item in all)
            {

                cellName.Phrase = new Phrase(item.Nombre);
                cellName.HorizontalAlignment = 1;
                cellName.Colspan = 2;
                cellName.BorderWidthRight = 1;
                table.AddCell(cellName);

                cell.Phrase = new Phrase(Convert.ToDateTime(item.FechaEntradaMonday).ToShortTimeString().ToString());
                cell.HorizontalAlignment = 1;
                table.AddCell(cell);


                cell.Phrase = new Phrase(Convert.ToDateTime(item.FechaSalidaMonday).ToShortTimeString().ToString());
                cell.HorizontalAlignment = 1;
                table.AddCell(cell);

                cellMonday = new PdfPCell(new Phrase((item.HorasMonday).ToString()));
                cellMonday.BackgroundColor = new iTextSharp.text.BaseColor(242, 242, 242);
                cellMonday.BorderWidthRight = 1;
                cellMonday.HorizontalAlignment = 1;
                table.AddCell(cellMonday);

                cell.Phrase = new Phrase(Convert.ToDateTime(item.FechaEntradaTuesday).ToShortTimeString().ToString());
                cell.HorizontalAlignment = 1;
                table.AddCell(cell);


                cell.Phrase = new Phrase(Convert.ToDateTime(item.FechaSalidaTuesday).ToShortTimeString().ToString());
                cell.HorizontalAlignment = 1;
                table.AddCell(cell);

                cellTuesday = new PdfPCell(new Phrase((item.HorasTuesday).ToString()));
                cellTuesday.BackgroundColor = new iTextSharp.text.BaseColor(242, 242, 242);
                cellTuesday.BorderWidthRight = 1;
                cellTuesday.HorizontalAlignment = 1;
                table.AddCell(cellTuesday);

                cell.Phrase = new Phrase(Convert.ToDateTime(item.FechaEntradaWendnsday).ToShortTimeString().ToString());
                cell.HorizontalAlignment = 1;
                table.AddCell(cell);


                cell.Phrase = new Phrase(Convert.ToDateTime(item.FechaSalidaWendnsday).ToShortTimeString().ToString());
                cell.HorizontalAlignment = 1;
                table.AddCell(cell);

                cellWednesday = new PdfPCell(new Phrase((item.HorasWendnsday).ToString()));
                cellWednesday.BackgroundColor = new iTextSharp.text.BaseColor(242, 242, 242);
                cellWednesday.BorderWidthRight = 1;
                cellWednesday.HorizontalAlignment = 1;
                table.AddCell(cellWednesday);


                cell.Phrase = new Phrase(Convert.ToDateTime(item.FechaEntradaThursday).ToShortTimeString().ToString());
                cell.HorizontalAlignment = 1;
                table.AddCell(cell);


                cell.Phrase = new Phrase(Convert.ToDateTime(item.FechaSalidaThursday).ToShortTimeString().ToString());
                cell.HorizontalAlignment = 1;
                table.AddCell(cell);

                cellThursday = new PdfPCell(new Phrase((item.HorasThursday).ToString()));
                cellThursday.BackgroundColor = new iTextSharp.text.BaseColor(242, 242, 242);
                cellThursday.BorderWidthRight = 1;
                cellThursday.HorizontalAlignment = 1;
                table.AddCell(cellThursday);

                cell.Phrase = new Phrase(Convert.ToDateTime(item.FechaEntradaFriday).ToShortTimeString().ToString());
                cell.HorizontalAlignment = 1;
                table.AddCell(cell);


                cell.Phrase = new Phrase(Convert.ToDateTime(item.FechaSalidaFriday).ToShortTimeString().ToString());
                cell.HorizontalAlignment = 1;
                table.AddCell(cell);

                cellfriday = new PdfPCell(new Phrase((item.HorasFriday).ToString()));
                cellfriday.BackgroundColor = new iTextSharp.text.BaseColor(242, 242, 242);
                cellfriday.BorderWidthRight = 1;
                cellfriday.HorizontalAlignment = 1;
                table.AddCell(cellfriday);

                cell.Phrase = new Phrase(Convert.ToDateTime(item.FechaEntradaSaturday).ToShortTimeString().ToString());
                cell.HorizontalAlignment = 1;
                table.AddCell(cell);


                cell.Phrase = new Phrase(Convert.ToDateTime(item.FechaSalidaSaturday).ToShortTimeString().ToString());
                cell.HorizontalAlignment = 1;
                table.AddCell(cell);

                cellsaturday = new PdfPCell(new Phrase((item.HorasSaturday).ToString()));
                cellsaturday.BackgroundColor = new iTextSharp.text.BaseColor(242, 242, 242);
                cellsaturday.BorderWidthRight = 1;
                cellsaturday.HorizontalAlignment = 1;
                table.AddCell(cellsaturday);

                cell.Phrase = new Phrase(Convert.ToDateTime(item.FechaEntradaSunday).ToShortTimeString().ToString());
                cell.HorizontalAlignment = 1;
                table.AddCell(cell);


                cell.Phrase = new Phrase(Convert.ToDateTime(item.FechaSalidaSunday).ToShortTimeString().ToString());
                cell.HorizontalAlignment = 1;
                table.AddCell(cell);

                cellsunday = new PdfPCell(new Phrase((item.HorasSunday).ToString()));
                cellsunday.BackgroundColor = new iTextSharp.text.BaseColor(242, 242, 242);
                cellsunday.BorderWidthRight = 1;
                cellsunday.HorizontalAlignment = 1;
                table.AddCell(cellsunday);

                cell.Phrase = new Phrase(Convert.ToString(Convert.ToDecimal(item.HorasMonday) + Convert.ToDecimal(item.HorasTuesday) + Convert.ToDecimal(item.HorasWendnsday) + Convert.ToDecimal(item.HorasThursday) + Convert.ToDecimal(item.HorasFriday) + Convert.ToDecimal(item.HorasSaturday) + Convert.ToDecimal(item.HorasSunday)));
                cell.HorizontalAlignment = 1;

                table.AddCell(cell);

                //table.Rows.Add(cell);
                //                 

            }



            return table;
        }

        public void CreatePdf(string name, string deptId, string date, bool type, string path)
        {
            Document document = new Document(PageSize.A1, 5, 5, 155, 15);
            var Id = Convert.ToInt32(deptId);
            name = name.Replace("ñ", "n");
            name = name.Replace("Ñ", "n");
            name = name.Replace("\u0001", " ");
            string folder = @"C:\\Pdf\\Asistance_Report_" + date;
            bool folderExists = Directory.Exists((folder));
            if (!folderExists)
                Directory.CreateDirectory((folder));
            folderExists = Directory.Exists((path));
            if (!folderExists)
                Directory.CreateDirectory((path));
            folderExists = Directory.Exists((path));
            //      string filepath = @"C:\\Pdf\\Asistance_Report_"+date+"\\"+path+"\\Asistance_Report_" + name + "_" + date + ".pdf";    
            path = path + "\\Asistance_Report_" + name + "_" + date + ".pdf";
            if ((!System.IO.File.Exists(path)))
            {
                using (FileStream output = new FileStream((path), FileMode.Create))
                {
                    using (PdfWriter writer = PdfWriter.GetInstance(document, output))
                    {


                        document.Open();
                        PdfPTable table = new PdfPTable(24);
                        table.TotalWidth = 5000f;
                        document.Add(archive(table, Id, name, date, type));
                        document.Close();
                        output.Close();
                    }


                }

            }
        }

        public void CreatePdf(IQueryable<string> test, IMongoCollection<infoDept> collection, string dates, bool type, string path)
        {
            DateTime now = DateTime.Now;
            var oneweekAffter = now.AddDays(7);

            Document document = new Document(PageSize.A1, 5, 5, 155, 15);
            int Id = 0;


            string filepath = path + "\\" + "Asistance_Report_LNO-HN" + "_" + dates + ".pdf";

            if ((!System.IO.File.Exists(filepath)))
            {
                using (FileStream output = new FileStream((filepath), FileMode.Create))
                {

                    using (PdfWriter writer = PdfWriter.GetInstance(document, output))
                    {
                        document.Open();


                        foreach (var items in test)
                        {
                            var managers = collection
                            .Find(a => a.Id_Dept == items)
                            .ToListAsync()
                             .Result;

                            foreach (var i in managers)
                            {
                                List<Getreport> get = new List<Getreport>(i.GetReport);
                                foreach (var sub in get)
                                {
                                    Id = Convert.ToInt32(sub.deptID);
                                    PdfPTable table = new PdfPTable(24);
                                    table.TotalWidth = 5000f;
                                    if (oneweekAffter.Month != now.Month)
                                    {
                                        document.Add(ArchiveSpin4(table, Id, sub.subdept, dates, type));
                               
                                    }
                                    else
                                    {
                                        document.Add(archive(table, Id, sub.subdept, dates, type));
                                   
                                      }
                                     }
                            }
                        }
                        document.Close();
                        output.Close();

                    }
                }
            }
            Id = 0;
            foreach (var items in test)
            {
                var managers = collection
                .Find(a => a.Id_Dept == items)
                .ToListAsync()
                 .Result;

                foreach (var i in managers)
                {
                    for (var y = 0; y < i.Name.Length; y++)
                    {

                        if (oneweekAffter.Month != now.Month)
                        {

                            Console.WriteLine("ultimo lunes ................................");
                            filepath = path + "\\Asistance_Report_All_" + i.Nombre_Dept + "_" + dates + ".pdf";
                            if ((!System.IO.File.Exists(filepath)))
                            {
                                using (FileStream output = new FileStream((filepath), FileMode.Create))
                                {
                                    Document Alldocument = new Document(PageSize.A1, 5, 5, 155, 15);

                                    using (PdfWriter writer = PdfWriter.GetInstance(Alldocument, output))
                                    {


                                        Alldocument.Open();
                                        List<Getreport> get = new List<Getreport>(i.GetReport);
                                        foreach (var sub in get)
                                        {
                                            Id = Convert.ToInt32(sub.deptID);
                                            PdfPTable table = new PdfPTable(24);
                                            table.TotalWidth = 5000f;
                                            Alldocument.Add(ArchiveSpin4(table, Id, sub.subdept, dates, type));
                                        }
                                        Alldocument.Close();
                                        output.Close();
                                    }

                                }


                            }
                        }
                        else
                        {

                            Console.WriteLine(" no es el ultimo lunes del mes................................");
                            filepath = path + "\\Asistance_Report_All" + i.Nombre_Dept + "_" + dates + ".pdf";
                            if ((!System.IO.File.Exists(filepath)))
                            {
                                using (FileStream output = new FileStream((filepath), FileMode.Create))
                                {
                                    Document Alldocument = new Document(PageSize.A1, 5, 5, 155, 15);

                                    using (PdfWriter writer = PdfWriter.GetInstance(Alldocument, output))
                                    {


                                        Alldocument.Open();
                                        List<Getreport> get = new List<Getreport>(i.GetReport);
                                        foreach (var sub in get)
                                        {
                                            Id = Convert.ToInt32(sub.deptID);
                                            PdfPTable table = new PdfPTable(24);
                                            table.TotalWidth = 5000f;
                                            //archive
                                            Alldocument.Add(archive(table, Id, sub.subdept, dates, type));
                                        }
                                        Alldocument.Close();
                                        output.Close();
                                    }

                                }


                            }
                        }

                        /*
                        filepath = path + "\\Asistance_Report_All" + i.Nombre_Dept + "_" + dates + ".pdf";
                        if ((!System.IO.File.Exists(filepath)))
                        {
                            using (FileStream output = new FileStream((filepath), FileMode.Create))
                            {
                                Document Alldocument = new Document(PageSize.A1, 5, 5, 155, 15);

                                using (PdfWriter writer = PdfWriter.GetInstance(Alldocument, output))
                                {


                                    Alldocument.Open();
                                    List<Getreport> get = new List<Getreport>(i.GetReport);
                                    foreach (var sub in get)
                                    {
                                        Id = Convert.ToInt32(sub.deptID);
                                        PdfPTable table = new PdfPTable(24);
                                        table.TotalWidth = 5000f;
                                        Alldocument.Add(archive(table, Id, sub.subdept, dates, type));
                                    }
                                    Alldocument.Close();
                                    output.Close();
                                }

                            }


                        }
                        */
                    }
                }
            }
        }


        public PdfPTable ArchiveSpin4(PdfPTable table, int deptID, string subdept, string dates, bool type)
        {
            List<test> all = new List<test>();
            PdfPCell cell = new PdfPCell();
            PdfPCell cell2 = new PdfPCell();
            PdfPCell cellName = new PdfPCell();
            PdfPCell cellMonday = new PdfPCell();
            PdfPCell cellTuesday = new PdfPCell();
            PdfPCell cellWednesday = new PdfPCell();
            PdfPCell cellThursday = new PdfPCell();
            PdfPCell cellfriday = new PdfPCell();
            PdfPCell cellsaturday = new PdfPCell();
            PdfPCell cellsunday = new PdfPCell();
            Array[] list = { };

            iTextSharp.text.Font georgia = FontFactory.GetFont("georgia", 14f);
            georgia.Color = iTextSharp.text.BaseColor.BLUE;
            iTextSharp.text.Font labels = FontFactory.GetFont("georgia", 14f);
            labels.Color = iTextSharp.text.BaseColor.WHITE;
            iTextSharp.text.Font Headers = FontFactory.GetFont("georgia", 18f);
            Headers.Color = iTextSharp.text.BaseColor.BLACK;
            List<test> concat = new List<test>();

            PdfPTable table2 = new PdfPTable(24);
            table.TotalWidth = 5000f;

            float[] widths = new float[] { 1300f, 1300, 1300f, 1300f, 1300f, 1300f, 1300f, 1300f, 1300f, 1300f, 1300f, 1300f, 1300f, 1300f, 1300f, 1300f, 1300f, 1300f, 1300f, 1300f, 1300f, 1300f, 1300f, 1300f };
            table.SetWidths(widths);
            DateTime week2 = Convert.ToDateTime(dates).AddDays(-7);
            DateTime week3 = Convert.ToDateTime(dates).AddDays(-14);
            DateTime week4 = Convert.ToDateTime(dates).AddDays(-21);
            DateTime week1 = Convert.ToDateTime(dates);
            List<test> listWeek1 = new List<test>();
            List<test> listWeek2 = new List<test>();
            List<test> listWeek3 = new List<test>();
            List<test> listWeek4 = new List<test>();
             List<test> potato = new List<test>();
            //  var init = Convert.ToInt32(deptID.ToString());
            int count = -7;
            
            for (int i = 0; i < 3; i++)
            {

                using (Anviz_Data_BaseEntities2 dc = new Anviz_Data_BaseEntities2())
                {
                    //   all = (from e in dc).ToList();
                    if (type == true)
                    {
                        listWeek1 = (from e in dc.sp_matrix_weekdays(deptID, (week1.ToString("MM-dd-yyyy"))) select new test { Nombre = e.name, FechaEntradaMonday = e.Hora_de_LLegada_Lunes, FechaSalidaMonday = e.Hora_de_salida_Lunes, HorasMonday = e.Monday, FechaEntradaTuesday = e.Hora_de_LLegada_Martes, FechaSalidaTuesday = e.Hora_de_salida_Martes, HorasTuesday = e.Tuesday, FechaEntradaWendnsday = e.Hora_de_LLegada_Miercoles, FechaSalidaWendnsday = e.Hora_de_salida_Miercoles, HorasWendnsday = e.Wednsday, FechaEntradaThursday = e.Hora_de_LLegada_Jueves, FechaSalidaThursday = e.Hora_de_salida_Jueves, HorasThursday = e.Thursday, FechaEntradaFriday = e.Hora_de_LLegada_Viernes, FechaSalidaFriday = e.Hora_de_salida_Viernes, HorasFriday = e.Friday, FechaEntradaSaturday = e.Hora_de_LLegada_Sabado, FechaSalidaSaturday = e.Hora_de_salida_Sabado, HorasSaturday = e.Saturday, FechaEntradaSunday = e.Hora_de_LLegada_Domingo, FechaSalidaSunday = e.Hora_de_salida_Domingo, HorasSunday = e.Sunday }).ToList();
                        listWeek2 = (from e in dc.sp_matrix_weekdays(deptID, (week2.ToString("MM-dd-yyyy"))) select new test { Nombre = e.name, FechaEntradaMonday = e.Hora_de_LLegada_Lunes, FechaSalidaMonday = e.Hora_de_salida_Lunes, HorasMonday = e.Monday, FechaEntradaTuesday = e.Hora_de_LLegada_Martes, FechaSalidaTuesday = e.Hora_de_salida_Martes, HorasTuesday = e.Tuesday, FechaEntradaWendnsday = e.Hora_de_LLegada_Miercoles, FechaSalidaWendnsday = e.Hora_de_salida_Miercoles, HorasWendnsday = e.Wednsday, FechaEntradaThursday = e.Hora_de_LLegada_Jueves, FechaSalidaThursday = e.Hora_de_salida_Jueves, HorasThursday = e.Thursday, FechaEntradaFriday = e.Hora_de_LLegada_Viernes, FechaSalidaFriday = e.Hora_de_salida_Viernes, HorasFriday = e.Friday, FechaEntradaSaturday = e.Hora_de_LLegada_Sabado, FechaSalidaSaturday = e.Hora_de_salida_Sabado, HorasSaturday = e.Saturday, FechaEntradaSunday = e.Hora_de_LLegada_Domingo, FechaSalidaSunday = e.Hora_de_salida_Domingo, HorasSunday = e.Sunday }).ToList();
                        listWeek3 = (from e in dc.sp_matrix_weekdays(deptID, (week3.ToString("MM-dd-yyyy"))) select new test { Nombre = e.name, FechaEntradaMonday = e.Hora_de_LLegada_Lunes, FechaSalidaMonday = e.Hora_de_salida_Lunes, HorasMonday = e.Monday, FechaEntradaTuesday = e.Hora_de_LLegada_Martes, FechaSalidaTuesday = e.Hora_de_salida_Martes, HorasTuesday = e.Tuesday, FechaEntradaWendnsday = e.Hora_de_LLegada_Miercoles, FechaSalidaWendnsday = e.Hora_de_salida_Miercoles, HorasWendnsday = e.Wednsday, FechaEntradaThursday = e.Hora_de_LLegada_Jueves, FechaSalidaThursday = e.Hora_de_salida_Jueves, HorasThursday = e.Thursday, FechaEntradaFriday = e.Hora_de_LLegada_Viernes, FechaSalidaFriday = e.Hora_de_salida_Viernes, HorasFriday = e.Friday, FechaEntradaSaturday = e.Hora_de_LLegada_Sabado, FechaSalidaSaturday = e.Hora_de_salida_Sabado, HorasSaturday = e.Saturday, FechaEntradaSunday = e.Hora_de_LLegada_Domingo, FechaSalidaSunday = e.Hora_de_salida_Domingo, HorasSunday = e.Sunday }).ToList();
                        listWeek4 = (from e in dc.sp_matrix_weekdays(deptID, (week4.ToString("MM-dd-yyyy"))) select new test { Nombre = e.name, FechaEntradaMonday = e.Hora_de_LLegada_Lunes, FechaSalidaMonday = e.Hora_de_salida_Lunes, HorasMonday = e.Monday, FechaEntradaTuesday = e.Hora_de_LLegada_Martes, FechaSalidaTuesday = e.Hora_de_salida_Martes, HorasTuesday = e.Tuesday, FechaEntradaWendnsday = e.Hora_de_LLegada_Miercoles, FechaSalidaWendnsday = e.Hora_de_salida_Miercoles, HorasWendnsday = e.Wednsday, FechaEntradaThursday = e.Hora_de_LLegada_Jueves, FechaSalidaThursday = e.Hora_de_salida_Jueves, HorasThursday = e.Thursday, FechaEntradaFriday = e.Hora_de_LLegada_Viernes, FechaSalidaFriday = e.Hora_de_salida_Viernes, HorasFriday = e.Friday, FechaEntradaSaturday = e.Hora_de_LLegada_Sabado, FechaSalidaSaturday = e.Hora_de_salida_Sabado, HorasSaturday = e.Saturday, FechaEntradaSunday = e.Hora_de_LLegada_Domingo, FechaSalidaSunday = e.Hora_de_salida_Domingo, HorasSunday = e.Sunday }).ToList();
                        var y = week2.Year;
                        var Lr3 = listWeek1.Concat(listWeek2).Concat(listWeek3).Concat(listWeek4).ToList();

                        if (y % 4 == 0 && (y % 100 != 0 || y % 400 == 0) && DateTime.Today.Month==3)
                        {
                            DateTime week5 = Convert.ToDateTime(dates).AddDays(-28);
                            List<test> listWeek5 = (from e in dc.sp_matrix_weekdays(deptID, (week5.ToString("MM-dd-yyyy"))) select new test { Nombre = e.name, FechaEntradaMonday = e.Hora_de_LLegada_Lunes, FechaSalidaMonday = e.Hora_de_salida_Lunes, HorasMonday = e.Monday, FechaEntradaTuesday = e.Hora_de_LLegada_Martes, FechaSalidaTuesday = e.Hora_de_salida_Martes, HorasTuesday = e.Tuesday, FechaEntradaWendnsday = e.Hora_de_LLegada_Miercoles, FechaSalidaWendnsday = e.Hora_de_salida_Miercoles, HorasWendnsday = e.Wednsday, FechaEntradaThursday = e.Hora_de_LLegada_Jueves, FechaSalidaThursday = e.Hora_de_salida_Jueves, HorasThursday = e.Thursday, FechaEntradaFriday = e.Hora_de_LLegada_Viernes, FechaSalidaFriday = e.Hora_de_salida_Viernes, HorasFriday = e.Friday, FechaEntradaSaturday = e.Hora_de_LLegada_Sabado, FechaSalidaSaturday = e.Hora_de_salida_Sabado, HorasSaturday = e.Saturday, FechaEntradaSunday = e.Hora_de_LLegada_Domingo, FechaSalidaSunday = e.Hora_de_salida_Domingo, HorasSunday = e.Sunday }).ToList();
                            Lr3 = Lr3.Concat(listWeek5).ToList();
                        }


                                 foreach (var row in Lr3)
                                 {
                                     int index = potato.FindIndex(a => a.Nombre == row.Nombre);
                                     if (!(index >= 0))
                                     {
                                         potato.Add(row);

                                     }
                                     else
                                     {
                                         var tomato = potato[index];
                                         decimal? tempHoursi = (row.HorasMonday.Equals(null) ? 0:row.HorasMonday);
                                         decimal? tempHoursi2 =(row.HorasTuesday.Equals(null) ? 0:row.HorasTuesday) ;
                                         decimal? tempHoursi3 = (row.HorasWendnsday.Equals(null) ? 0 : row.HorasWendnsday);
                                         decimal? tempHoursi4 = (row.HorasThursday.Equals(null) ? 0:row.HorasThursday);
                                         decimal? tempHoursi5 = (row.HorasFriday.Equals(null) ? 0:row.HorasFriday);
                                         decimal? tempHoursi6 = (row.HorasSaturday.Equals(null) ? 0:row.HorasSaturday);
                                         decimal? tempHoursi7 = (row.HorasSunday.Equals(null) ? 0:row.HorasSunday);


                                         decimal? tempHours2 = (tomato.HorasMonday.Equals(null) ?0:tomato.HorasMonday) + tempHoursi;
                                         decimal? tempHours3 = (tomato.HorasTuesday.Equals(null) ?0:tomato.HorasTuesday)+ tempHoursi2;
                                         decimal? tempHours4 = (tomato.HorasWendnsday.Equals(null) ?0:tomato.HorasWendnsday) + tempHoursi3;
                                         decimal? tempHours5 = (tomato.HorasThursday.Equals(null) ?0:tomato.HorasThursday) + tempHoursi4;
                                         decimal? tempHours6 = (tomato.HorasFriday.Equals(null) ?0:tomato.HorasFriday) + tempHoursi5;
                                         decimal? tempHours7 = (tomato.HorasSaturday.Equals(null) ?0:tomato.HorasSaturday) + tempHoursi6;
                                         decimal? tempHours8 = (tomato.HorasSunday.Equals(null) ?0:tomato.HorasSunday) + tempHoursi7;


                                         tomato.HorasMonday = tempHours2;
                                         tomato.HorasTuesday = tempHours3;
                                          tomato.HorasWendnsday = tempHours4;
                                         tomato.HorasThursday = tempHours5;
                                         tomato.HorasFriday = tempHours6;
                                         tomato.HorasSaturday = tempHours7;
                                         tomato.HorasSunday = tempHours8;

                                         potato[index] = tomato;
                                     }


                                     all = potato;

                                 }
                
                        //    WorkingHoursPerDay = (from i in dc.sp_GetWorkingHoursPerDay(1, Convert.ToDateTime("08-13-2015").AddDays(w)) select new WorkingPerDay { Name = i.Name, HourInput = i.horaEntrada, HourOutput = i.horasalida, workingHours = Convert.ToInt32(i.horaTrabajada) }).ToList();
                        //     list[w] = WorkingHoursPerDay.ToList().ToArray();
                    }/*
                    else
                    {
                        all = (from e in dc.GetEmployes(deptID, dates) select new test { Nombre = e.name, FechaEntradaMonday = e.Hora_de_LLegada_Lunes, FechaSalidaMonday = e.Hora_de_salida_Lunes, HorasMonday = e.Monday, FechaEntradaTuesday = e.Hora_de_LLegada_Martes, FechaSalidaTuesday = e.Hora_de_salida_Martes, HorasTuesday = e.Tuesday, FechaEntradaWendnsday = e.Hora_de_LLegada_Miercoles, FechaSalidaWendnsday = e.Hora_de_salida_Miercoles, HorasWendnsday = e.Wednsday, FechaEntradaThursday = e.Hora_de_LLegada_Jueves, FechaSalidaThursday = e.Hora_de_salida_Jueves, HorasThursday = e.Thursday, FechaEntradaFriday = e.Hora_de_LLegada_Viernes, FechaSalidaFriday = e.Hora_de_salida_Viernes, HorasFriday = e.Friday, FechaEntradaSaturday = e.Hora_de_LLegada_Sabado, FechaSalidaSaturday = e.Hora_de_salida_Sabado, HorasSaturday = e.Saturday, FechaEntradaSunday = e.Hora_de_LLegada_Domingo, FechaSalidaSunday = e.Hora_de_salida_Domingo, HorasSunday = e.Sunday }).ToList();

                    }*/
            
                }
               
                DateTime datee = Convert.ToDateTime(dates);
                DateTime back = Convert.ToDateTime(datee.AddDays(-count));
                count = count - 7;
                dates = back.ToString();
           
            }

            DateTime date = Convert.ToDateTime(dates);

            //PdfPCell june = new PdfPCell(new Phrase("Cell 2",georgia));


            cell2 = new PdfPCell(new Phrase(subdept, Headers));
            cell2.HorizontalAlignment = Element.ALIGN_CENTER;
            cell2.BackgroundColor = new iTextSharp.text.BaseColor(245, 92, 24);
            cell2.BorderColor = new iTextSharp.text.BaseColor(0, 0, 0);
            cell2.Colspan = 24;
            table.AddCell(cell2);




            PdfPCell nombre = new PdfPCell(new Phrase("Name", labels));
            nombre.BackgroundColor = new iTextSharp.text.BaseColor(89, 89, 89);
            nombre.BorderColor = new iTextSharp.text.BaseColor(0, 0, 0);
            nombre.BorderWidthRight = 1;
            nombre.Rowspan = 2;
            nombre.Colspan = 2;
            nombre.HorizontalAlignment = 1;
            nombre.VerticalAlignment = 1;
            table.AddCell(nombre);

            PdfPCell june = new PdfPCell(new Phrase("Monday  " + "\n"));
            june.BorderWidthRight = 1;
            june.HorizontalAlignment = 1;
            june.BackgroundColor = new iTextSharp.text.BaseColor(247, 150, 70);
            june.BorderColor = new iTextSharp.text.BaseColor(0, 0, 0);
            june.Colspan = 3;
            table.AddCell(june);
            june = new PdfPCell(new Phrase("Tuesday  " + "\n" ));
            june.BorderWidthRight = 1;
            june.BackgroundColor = new iTextSharp.text.BaseColor(247, 150, 70);
            june.BorderColor = new iTextSharp.text.BaseColor(0, 0, 0);
            june.HorizontalAlignment = 1;
            june.Colspan = 3;
            table.AddCell(june);
            june = new PdfPCell(new Phrase("Wednesday  " + "\n"  ));
            june.BorderWidthRight = 1;
            june.BackgroundColor = new iTextSharp.text.BaseColor(247, 150, 70);
            june.BorderColor = new iTextSharp.text.BaseColor(0, 0, 0);
            june.HorizontalAlignment = 1;
            june.Colspan = 3;
            table.AddCell(june);
            june = new PdfPCell(new Phrase("Thursday  " + "\n" ));
            june.BorderWidthRight = 1;
            june.BackgroundColor = new iTextSharp.text.BaseColor(247, 150, 70);
            june.BorderColor = new iTextSharp.text.BaseColor(0, 0, 0);
            june.HorizontalAlignment = 1;
            june.Colspan = 3;
            table.AddCell(june);

            june = new PdfPCell(new Phrase("Friday   " + "\n" ));
            june.BorderWidthRight = 1;
            june.BackgroundColor = new iTextSharp.text.BaseColor(247, 150, 70);
            june.BorderColor = new iTextSharp.text.BaseColor(0, 0, 0);
            june.HorizontalAlignment = 1;
            june.Colspan = 3;
            table.AddCell(june);


            june = new PdfPCell(new Phrase("Saturday  " + "\n" ));
            june.BorderWidthRight = 1;
            june.BackgroundColor = new iTextSharp.text.BaseColor(247, 150, 70);
            june.BorderColor = new iTextSharp.text.BaseColor(0, 0, 0);
            june.HorizontalAlignment = 1;
            june.Colspan = 3;
            table.AddCell(june);
            june = new PdfPCell(new Phrase("Sunday   " + "\n" ));
            june.BorderWidthRight = 1;
            june.BackgroundColor = new iTextSharp.text.BaseColor(247, 150, 70);
            june.BorderColor = new iTextSharp.text.BaseColor(0, 0, 0);
            june.HorizontalAlignment = 1;
            june.Colspan = 3;
            table.AddCell(june);

            june = new PdfPCell(new Phrase("Monthly Hours Worked", labels));
            june.BorderWidthRight = 1;
            june.Rowspan = 2;
            june.HorizontalAlignment = 1;
            june.BackgroundColor = new iTextSharp.text.BaseColor(89, 89, 89);
            june.BorderColor = new iTextSharp.text.BaseColor(0, 0, 0);
            table.AddCell(june);




            PdfPCell In = new PdfPCell();
            PdfPCell OUT = new PdfPCell();
            PdfPCell second = new PdfPCell();

            for (var y = 0; y <= 6; y++)
            {
               
                second = new PdfPCell(new Phrase("Worked Hours", labels));
                second.BackgroundColor = new iTextSharp.text.BaseColor(89, 89, 89);
                second.BorderColor = new iTextSharp.text.BaseColor(0, 0, 0);
                second.BorderWidthRight = 1;
                second.HorizontalAlignment = 1;
                second.Colspan = 3;
                table.AddCell(second);


            }
        
            foreach (var item in all)
            {

                
                cellName.Phrase = new Phrase(item.Nombre);
                cellName.HorizontalAlignment = 1;
                cellName.Colspan = 2;
                cellName.BorderWidthRight = 1;
                table.AddCell(cellName);

                cellMonday = new PdfPCell(new Phrase((item.HorasMonday).ToString()));
                cellMonday.BackgroundColor = new iTextSharp.text.BaseColor(242, 242, 242);
                cellMonday.BorderWidthRight = 1;
                cellMonday.HorizontalAlignment = 1;
                cellMonday.Colspan = 3;
                table.AddCell(cellMonday);


                cellTuesday = new PdfPCell(new Phrase((item.HorasTuesday).ToString()));
                cellTuesday.BackgroundColor = new iTextSharp.text.BaseColor(242, 242, 242);
                cellTuesday.BorderWidthRight = 1;
                cellTuesday.HorizontalAlignment = 1;
                cellTuesday.Colspan = 3;
                table.AddCell(cellTuesday);

                cellWednesday = new PdfPCell(new Phrase((item.HorasWendnsday).ToString()));
                cellWednesday.BackgroundColor = new iTextSharp.text.BaseColor(242, 242, 242);
                cellWednesday.BorderWidthRight = 1;
                cellWednesday.HorizontalAlignment = 1;
                cellWednesday.Colspan = 3;
                table.AddCell(cellWednesday);


                cellThursday = new PdfPCell(new Phrase((item.HorasThursday).ToString()));
                cellThursday.BackgroundColor = new iTextSharp.text.BaseColor(242, 242, 242);
                cellThursday.BorderWidthRight = 1;
                cellThursday.HorizontalAlignment = 1;
                cellThursday.Colspan = 3;
                table.AddCell(cellThursday);

                cellfriday = new PdfPCell(new Phrase((item.HorasFriday).ToString()));
                cellfriday.BackgroundColor = new iTextSharp.text.BaseColor(242, 242, 242);
                cellfriday.BorderWidthRight = 1;
                cellfriday.HorizontalAlignment = 1;
                cellfriday.Colspan = 3;
                table.AddCell(cellfriday);

                
                cellsaturday = new PdfPCell(new Phrase((item.HorasSaturday).ToString()));
                cellsaturday.BackgroundColor = new iTextSharp.text.BaseColor(242, 242, 242);
                cellsaturday.BorderWidthRight = 1;
                cellsaturday.HorizontalAlignment = 1;
                cellsaturday.Colspan = 3;
                table.AddCell(cellsaturday);

                cellsunday = new PdfPCell(new Phrase((item.HorasSunday).ToString()));
                cellsunday.BackgroundColor = new iTextSharp.text.BaseColor(242, 242, 242);
                cellsunday.BorderWidthRight = 1;
                cellsunday.HorizontalAlignment = 1;
                cellsunday.Colspan = 3;
                table.AddCell(cellsunday);

                cell.Phrase = new Phrase(Convert.ToString(Convert.ToDecimal(item.HorasMonday) + Convert.ToDecimal(item.HorasTuesday) + Convert.ToDecimal(item.HorasWendnsday) + Convert.ToDecimal(item.HorasThursday) + Convert.ToDecimal(item.HorasFriday) + Convert.ToDecimal(item.HorasSaturday) + Convert.ToDecimal(item.HorasSunday)));
                cell.HorizontalAlignment = 1;

                table.AddCell(cell);

                //table.Rows.Add(cell);
                //                 

            }



            return table;
        }


        public IEnumerable<object> query { get; set; }
    }
}

