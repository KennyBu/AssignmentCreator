using System;
using System.Collections.Generic;
using System.Linq;
using LinqToExcel;

namespace AssignmentCreatorConsoleApplication
{
    class Program
    {
        static void Main(string[] args)
        {
            const string excelPath = "c:/temp/excel/Assignments.xlsx";

            var excel = new ExcelQueryFactory(excelPath);

            var assignments = from e in excel.Worksheet<Assignment>("Assignments")
                            select e;

            var db = new PetaPoco.Database("AssignifyItDatabase");
            var calendarReminder = new CalendarReminder.SqlCalendarReminder(db);


            foreach (var assignment in assignments)
            {
                var calendarEvent = MapFrom(assignment);
                calendarReminder.Create(calendarEvent);
            }

        }

        private static void WriteInfo(IEnumerable<Assignment> assignments)
        {
            foreach (var assignment in assignments)
            {
                Console.WriteLine("Date:{0} Title:{1} Name{2} Email:{3} ",
                    assignment.AssignmentDate, assignment.Title, assignment.Name, assignment.Email);
            }
        }

        private static CalendarEvent MapFrom(Assignment assignment)
        {
            var calendarEvent = new CalendarEvent
            {
                Id = Guid.NewGuid(),
                AssignmnetDate = assignment.AssignmentDate,
                Title = assignment.Title,
                Name = assignment.Name,
                Assignment = assignment.Title,
                Email = assignment.Email
            };

            return calendarEvent;
        }
    }
}
