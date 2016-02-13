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

				var reportParagraph = document.InsertParagraph();
				reportParagraph.Append ("Este es un reporte perteneciente a las clases que imparte ");
				reportParagraph.Append (teacher.GivenName + " " + teacher.LastName).Bold().Append(", generado en ");
				reportParagraph.Append (DateTime.Now.ToShortDateString ()).Italic().Append (". ")
					.Append ("El profesor/ra ").Append(teacher.LastName).Font(new FontFamily("Arial Black"))
					.Append(" imparte actualmente ")
					.Append(lectures.Length + " asignaturas.").Color(Color.Blue).Italic().Bold();

                document.Save();
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