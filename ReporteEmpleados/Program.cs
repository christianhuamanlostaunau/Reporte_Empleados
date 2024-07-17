using System.Data.SqlClient;
using OfficeOpenXml;
using System.Net;
using System.Net.Mail;
using ReporteEmpleados;

class Program
{
    static void Main()
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        Console.WriteLine("Inicializando");
        // Configuración de las fechas de inicio y fin desde variables de entorno
        DateTime fechaInicio = DateTime.Parse(Environment.GetEnvironmentVariable("FechaInicio") ?? "2020-01-01");
        DateTime fechaFin = DateTime.Parse(Environment.GetEnvironmentVariable("FechaFin") ?? "2022-12-31");

        // Conexión a la base de datos SQL Server
        string connectionString = "Server=GKNLPROYECTO020\\SQL2022;Database=EmployeeDB;User Id=sa;Password=Chuaman;Encrypt=False;";
        List<Employee> employees = new List<Employee>();

        using (SqlConnection connection = new SqlConnection(connectionString))
        {
            Console.WriteLine("Consultando Employees");
            string sqlQuery = "SELECT * FROM Employees WHERE Admission_Date >= @FechaInicio AND Admission_Date <= @FechaFin";
            SqlCommand command = new SqlCommand(sqlQuery, connection);
            command.Parameters.AddWithValue("@FechaInicio", fechaInicio);
            command.Parameters.AddWithValue("@FechaFin", fechaFin);

            connection.Open();
            SqlDataReader reader = command.ExecuteReader();

            while (reader.Read())
            {
                Employee emp = new Employee
                {
                    Id = Convert.ToInt32(reader["Id"]),
                    Name = Convert.ToString(reader["Name"]),
                    Document_Number = Convert.ToString(reader["Document_Number"]),
                    Salary = Convert.ToDecimal(reader["Salary"]),
                    Age = Convert.ToInt32(reader["Age"]),
                    Profile = Convert.ToString(reader["Profile"]),
                    Admission_Date = Convert.ToDateTime(reader["Admission_Date"])
                };
                employees.Add(emp);
            }

            reader.Close();
            connection.Close();
        }

        Console.WriteLine("Generando Excel");

        // Verificar si el archivo existe y eliminarlo si es necesario
        string filePath = "ReporteEmpleados.xlsx";
        if (File.Exists(filePath))
        {
            File.Delete(filePath);
        }

        // Generar reporte Excel con EPPlus
        var file = new FileInfo(filePath);
        using (ExcelPackage package = new ExcelPackage(file))
        {
            ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Reporte Empleados");

            worksheet.Cells.LoadFromCollection(employees, true);
            worksheet.Cells["G2:G" + (employees.Count + 1)].Style.Numberformat.Format = "dd-mm-yyyy";
            package.Save();
        }

        Console.WriteLine("Excel Generado");

        try
        {
            Console.WriteLine("Iniciando envío de correo");
            // Enviar correo electrónico con el reporte adjunto
            SendEmailWithAttachment("franco.paredes@oechsle.pe", "Reporte Empleados - Examen Técnico Oechsle", "Adjunto se encuentra el reporte solicitado.", file.FullName);
            Console.WriteLine("Correo enviado");
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.Message);
        }
    }

    static void SendEmailWithAttachment(string toEmail, string subject, string body, string attachmentPath)
    {
        string fromEmail = "examenreporteempleados@gmail.com";
        string password = "zshqsnfcjogzdnvw"; // Utiliza la contraseña de aplicación generada aquí

        using (MailMessage mail = new MailMessage(fromEmail, toEmail))
        {
            mail.Subject = subject;
            mail.Body = body;
            mail.IsBodyHtml = false;

            Attachment attachment = new Attachment(attachmentPath);
            mail.Attachments.Add(attachment);

            using (SmtpClient smtp = new SmtpClient("smtp.gmail.com", 587))
            {
                smtp.UseDefaultCredentials = false;
                smtp.Credentials = new NetworkCredential(fromEmail, password);
                smtp.EnableSsl = true;
                smtp.Send(mail);
            }
        }
    }
}


