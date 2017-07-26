using Slack.Webhooks;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Net;
using System.Net.Mail;

namespace Noti_Turnos
{
    class Program
    {
        static void Main(string[] args)
        {
            /*
            // Inicio :3
            string greet = "\n" +
                            "\t##  ## ##   ##  ##  ####  ###### ##### " + "\n" +
                            "\t##  ## ##   ##  ## ###    ##     ##  ##" + "\n" +
                            "\t###### ##   ##  ##  ####  ####   ##### " + "\n" +
                            "\t##  ## ##   ##  ##    ### ##     ## ## " + "\n" +
                            "\t##  ## ##    ####   ####  ###### ##  ##" + "\n";

            // Console.WriteLine(greet);
            */


            // Seteo del idioma español para representacion textual de fechas
            System.Globalization.CultureInfo esp = new System.Globalization.CultureInfo("es-HN");

            // Creacion del objeto que establece conexion a la base de datos
            string conn = "Data Source=ADALBERTO\\BERTZ;Initial Catalog=Control Turnos;Integrated Security=True;";
            SqlConnection sql = new SqlConnection(conn);

            // Inicia ciclo principal
            try
            {
                // Apertura de la conexion
                sql.Open();
                Console.WriteLine("Connected");

                // Query que consigue la lista de turnos registrados en la base de datos
                SqlCommand command = new SqlCommand("Select * from Turno inner join Empleado on Turno.empleado = Empleado.idEmpleado inner join Sucursal on Turno.sucursal = Sucursal.idSucursal", sql);

                // Lectura de los resultados del query
                using (SqlDataReader rd = command.ExecuteReader())
                {
                    if (rd.Read())
                    {
                        // Inicializacion de variables que seran reutilizadas por cada registro leido
                        string empleado, sucursal, message, storeDate = "";
                        List<Turno> turnos = new List<Turno>();
                        TimeSpan horaTurno;
                        DateTime fechaTurno, now = DateTime.Now;
                        Turno turno;



                        // Introduccion del correo
                        message = "Buen dia.<br><br>Este fin de semana en soporte los turnos son los siguientes:<br>";

                        // Creacion del objeto que genera un correo
                        MailMessage mail = new MailMessage()
                        {
                            From        = new MailAddress("ahernandezr@diunsa.hn"),
                            Subject     = "Control Turnos DIUNSA",
                            IsBodyHtml  = true
                        };

                        // Creacion del objeto que conecta al cliente SMTP
                        SmtpClient client = new SmtpClient()
                        {
                            // Parametros de conexion SMTP
                            Port = 587,
                            DeliveryMethod = SmtpDeliveryMethod.Network,
                            Host = "diunsaex.diunsa.hn",
                            UseDefaultCredentials = false,
                            Credentials = new NetworkCredential("ahernandezr", "Ah83701996", "diunsa.hn")
                        };

                        do
                        {
                            // Se arma el nombre del empleado del turno
                            empleado = rd["primerNombre"] + " " + rd["primerApellido"];
                            
                            // Se guarda la fecha y hora de inicio del turno en un solo valor
                            // y las horas que durara el turno en otro valor. 
                            fechaTurno = (DateTime)rd["fecha"];
                            horaTurno = TimeSpan.FromHours((int)rd["horas"]);
                            Console.WriteLine("Fecha Turno " + rd["primerNombre"] + ": " + fechaTurno);
                            Console.WriteLine("Horas: " + horaTurno.ToString("hh"));

                            // Se guarda el nombre de la sucursal
                            sucursal = rd["nombre"].ToString();

                            // Comparacion de fechas
                            // A la fecha de turno se le restan tres dias correspondiendo al fin de semana
                            if (DateTime.Compare(now, fechaTurno.AddDays(-3)) >= 0 && DateTime.Compare(now, fechaTurno) < 0)
                            {
                                try
                                {
                                    Console.WriteLine("Sending email");

                                    // Se arma un objeto tipo Turno para guardar los datos anteriormente sacados
                                    turno = new Turno(empleado, fechaTurno, horaTurno, sucursal, esp);

                                    // Se agrega un correo a la lista de recipientes
                                    mail.To.Add(new MailAddress((String)rd["correo"]));


                                    // Se arma el mensaje a enviar por correo
                                    /* 
                                    FORMATO
                                    
                                    Dia/Mes
                                        (nombre de empleado), Sucursal (nombre de sucursal)
                                        (hora inicio de turno) - (hora fin de turno)
                                        (nombre de empleado), Sucursal (nombre de sucursal)
                                        (hora inicio de turno) - (hora fin de turno)
                                    
                                    */
                                    if (!storeDate.Equals(turno.getFecha()))
                                    {
                                        storeDate = turno.getFecha();
                                        message += "<br><b>" + storeDate + "</b><br>";
                                    }
                                    message += "&emsp;&emsp;" + empleado + ", Sucursal "+sucursal+"<br>&emsp;&emsp;" + turno.getHoraInicio() + " - " + turno.getHoraFin() + "<br>";

                                    // Se guarda el objeto Turno en una lista para la integracion con Slack
                                    turnos.Add(turno);

                                    
                                }
                                catch (Exception e)
                                {
                                    Console.WriteLine("Error: " + e.Message);
                                    if (e.InnerException != null) Console.WriteLine("Details: " + e.InnerException);
                                }
                            }
                            else Console.WriteLine("No emails to send");
                        } while (rd.Read());

                        // Se pega el mensaje al correo
                        mail.Body = message;

                        // Envio del correo
                        client.Send(mail);
                        Console.WriteLine("Email sent successfully");




                        //---                               ---
                        //---   SLACK WEBHOOK INTEGRATION   ---
                        //---                               ---


                        // Se inicializa el cliente Slack con un link "webhook"
                        Console.WriteLine("Integrating Webhook");
                        string webhookUrl = "https://hooks.slack.com/services/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";
                        SlackClient slack = new SlackClient(webhookUrl);

                        // Creacion de elementos del mensaje Slack
                        List<SlackAttachment> attachments = new List<SlackAttachment>();
                        List<SlackField> fields = new List<SlackField>();

                        
                        string attachDate = "";
                        foreach(Turno t in turnos)
                        {
                            if(!t.getFecha().Equals(attachDate))
                            {
                                attachDate = t.getFecha();

                                // Espacio vacio para formato
                                fields.Add(new SlackField { Value = " " });

                                // (dia/mes)
                                fields.Add(new SlackField { Title = attachDate });
                            }

                            // (nombre de empleado), Sucursal (nombre de sucursal)
                            // (hora inicio de turno) - (hora fin de turno)
                            fields.Add(new SlackField { Value = t.getEmpleado() + ", Sucursal " + t.getSucursal() });
                            fields.Add(new SlackField { Value = t.getHoraInicio() + " - " + t.getHoraFin() });

                            // Espacio vacio para formato
                            fields.Add(new SlackField { Value = " " });
                        }

                        // Se arma el mensaje de notificacion
                        attachments.Add(new SlackAttachment
                        {
                            AuthorName  = "DIUNSA",
                            Fallback    = "TURNOS FIN DE SEMANA",
                            Title       = "TURNOS FIN DE SEMANA",
                            Color       = "#07f",
                            Fields      = fields
                        });


                        // Se inscribe la notificacion dentro de un objeto con parametros de envio
                        SlackMessage slackMessage = new SlackMessage
                        {
                            Attachments = attachments,
                            Channel = "#test",
                            Username = "Control Turnos",
                            Mrkdwn = false
                        };

                        // Envio de notificacion a Slack
                        slack.Post(slackMessage);
                        Console.WriteLine("Slack message sent successfully");
                    }
                    else
                    {
                        Console.WriteLine("No emails to send");
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("Error: " + e.Message);
            }
            finally
            {
                // Cierre de conexion a la base de datos
                sql.Close();
            }


            /*
            // Final :3
            string bye = "\n" +
                            "\t#####  ##  ## ######   ##  ##  ####  ###### ##### " + "\n" +
                            "\t##  ## ##  ## ##       ##  ## ###    ##     ##  ##" + "\n" +
                            "\t#####   ####  ####     ##  ##  ####  ####   ##### " + "\n" +
                            "\t##  ##   ##   ##       ##  ##    ### ##     ## ## " + "\n" +
                            "\t#####    ##   ######    ####   ####  ###### ##  ##" + "\n";

            // Console.WriteLine(bye);

            // Thread.Sleep(1500);
            */

        }
    }

    public class Turno
    {
        string empleado, fecha, hora1, hora2, sucursal;
        public Turno(string empleado, DateTime fecha, TimeSpan horas, string sucursal, System.Globalization.CultureInfo esp)
        {
            this.empleado   = empleado;
            this.fecha      = fecha.ToString("dd/MMMM", esp);
            this.hora1      = fecha.ToString("hh:mm tt");
            this.hora2      = fecha.AddHours(horas.Hours).ToString("hh:mm tt");
            this.sucursal   = sucursal;
        }

        public string getEmpleado() { return this.empleado; }

        public string getFecha() { return this.fecha; }

        public string getHoraInicio() { return this.hora1; }

        public string getHoraFin() { return this.hora2; }

        public string getSucursal() { return this.sucursal; }
    }
}
