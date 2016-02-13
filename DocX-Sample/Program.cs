using System;
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

                var reportText = document.InsertParagraph();
                reportText.Append("Este es un reporte perteneciente a las clases que imparte " + teacher.GivenName + " " +
                                    teacher.LastName + ", generado en " + DateTime.Now.ToShortDateString() + ". ")
                            .Append("El profesor/ra " + teacher.LastName + " imparte actualmente " + lectures.Length + " asignaturas.");

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