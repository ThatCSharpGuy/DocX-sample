using System;
using System.Drawing;
using System.Linq;
using Novacode;

namespace DocX_Sample
{
    internal class Program
    {
        private static FakeDatabase _database = new FakeDatabase();


        private static void Main(string[] args)
        {
            var teacher = GetMostActiveTeacher();
            var lectures = GetLecturesForTeacher(teacher.Id);

            using (var document = DocX.Create("Prueba.docx"))
            {
				var titleParagraph = document.InsertParagraph();
				titleParagraph.Append("Reporte " + teacher.LastName).Heading(HeadingType.Heading1);

				var reportParagraph = document.InsertParagraph();
				reportParagraph.Append ("Este es un reporte perteneciente a las clases que imparte ");
				reportParagraph.Append (teacher.GivenName + " " + teacher.LastName).Bold().Append(", generado en ");
				reportParagraph.Append (DateTime.Now.ToShortDateString ()).Italic().Append (". ")
					.Append ("El profesor/ra ").Append(teacher.LastName).Font(new FontFamily("Arial Black"))
					.Append(" imparte actualmente ")
					.Append(lectures.Length + " asignaturas.").Color(Color.Blue).Italic().Bold();

				document.AddHeaders ();
				document.AddFooters ();

				var header = document.Headers.odd.InsertParagraph ();
				header.Append ("Reporte - That C# Guy").Font(new FontFamily("Courier New"));

				document.Footers.odd.PageNumbers = true;

				var table = document.InsertTable (1, 3);

				table.AutoFit = AutoFit.Window;
				var border = new Border (BorderStyle.Tcbs_single, BorderSize.one, 0, Color.Black);
				table.SetBorder (TableBorderType.InsideH, border);
				table.SetBorder (TableBorderType.InsideV, border);
				table.SetBorder (TableBorderType.Top, border);
				table.SetBorder (TableBorderType.Right, border);
				table.SetBorder (TableBorderType.Bottom, border);
				table.SetBorder (TableBorderType.Left, border);

				table.Design = TableDesign.ColorfulGrid;

				var tableHeaders = table.Rows[0];
				tableHeaders.Cells[0].InsertParagraph().Append("ID").Bold();
				tableHeaders.Cells[1].InsertParagraph().Append("Clase").Bold();
				tableHeaders.Cells[2].InsertParagraph().Append("Nivel").Bold();

				foreach (var lecture in lectures) 
				{
					var tableRow = table.InsertRow();
					tableRow.Cells[0].InsertParagraph().Append (lecture.Id.ToString());
					tableRow.Cells[1].InsertParagraph().Append(lecture.Name);
					tableRow.Cells[2].InsertParagraph().Append(lecture.Level);
				}

                document.Save();
            }


			using (var template = DocX.Load ("template.docx")) 
			{
				template.ReplaceText("esta entrada", "este post sobre DocX");
				template.ReplaceText("querido", "querido y respetable");
				template.ReplaceText("Facebook", "Twitter");
				template.ReplaceText("correo electrónico", "feregrino@thatcsharpguy.com");

				template.SaveAs("out.docx");
			}
        }

        static Teacher GetMostActiveTeacher()
        {
            var mostActiveTeacherId = (from lecture in _database.Lectures
                                       group lecture by lecture.TeacherId into group1
                                       orderby group1.Count() descending
                                       select group1.Key).FirstOrDefault();

            return _database.Teachers.Single(t => t.Id == mostActiveTeacherId);
        }

        static Lecture[] GetLecturesForTeacher(int teacherId)
        {
            return _database.Lectures.Where(lecture => lecture.TeacherId == teacherId).ToArray();
        }
    }
}