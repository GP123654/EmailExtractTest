using System;
using System.IO;
using System.Data;
using System.Data.SqlClient;

namespace EmailExtractTest
{
    /// <summary>
    /// Class used to insert data into the database
    /// </summary>
    public class InsertDataJobLog
    {
        /// <summary>
        /// 
        /// </summary>
        /// <param name="directoryPath"></param>
        /// <param name="connectionString"></param>
        public void InsertDataIntoTables(string directoryPath, string connectionString)
        {
            DirectoryInfo dir = new DirectoryInfo(directoryPath);

            //Foreach file that is a text file in the directory
            foreach (var file in dir.GetFiles("*.txt"))
            {
                //Reads all the lines of the file
                string[] lines = File.ReadAllLines(file.FullName);

                //For each line in the file
                for (int i = 2; i < lines.Length; i++) // Start from the third line
                {
                    //getting and storing the value of the 3rd row
                    string line = lines[i];
                    //Splitting the values in each row by a pipe ( | )
                    string[] values = line.Split('|');

                    if (values.Length >= 3)
                    {
                        string serverName = values[0];
                        string categoryName = values[1];
                        string name = values[2];
                        string ownerID = values[3];
                        string enabled = values[4];
                        string scheduled = values[5];
                        string description = values[6];
                        string reportServerJob = values[7];
                        string frequencyType = values[8];
                        string occurence = values[9];
                        string frequency = values[10];
                        string averageDurationSeconds = values[11];
                        string averageDuration = values[12];
                        string nextScheduledRunDate = values[13];
                        string lastRunDate = values[14];
                        string runStatusDesc = values[15];
                        string runTimeInSeconds = values[16];
                        string runTime = values[17];
                        string message = values[18];

                        //Inserts the data into the data base with the values from the file
                        InsertIntoTemp_DataFromEmailExtractionTable(connectionString, serverName, categoryName, name, ownerID, enabled, scheduled, description, reportServerJob, frequencyType, occurence, frequency, averageDurationSeconds, averageDuration, nextScheduledRunDate, lastRunDate, runStatusDesc, runTimeInSeconds, runTime, message);

                    }
                    else
                    {
                        //If there is a problem it will print this message
                        Console.WriteLine($"Skipped: Invalid data format in file {file.Name}");
                    }
                }
            }
        }

        /// <summary>
        /// 
        /// </summary>
        private void InsertIntoTemp_DataFromEmailExtractionTable(string connectionString, string serverName, string categoryName, string name, string ownerID, string enabled, string scheduled, string description, string reportServerJob, string frequencyType, string occurence, string frequency, string averageDurationSeconds, string averageDuration, string nextScheduledRunDate, string lastRunDate, string runStatusDesc, string runTimeInSeconds, string runTime, string message)
        {
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();

                using (SqlCommand command = connection.CreateCommand())
                {
                    command.CommandText = "INSERT INTO Temp_DataFromEmailExtraction (ServerName, CategoryName, Name, OwnerID, Enabled, Scheduled, Description, ReportServerJob, FrequencyType, Occurence, Frequency, AverageDurationSeconds, AverageDuration, NextScheduledRunDate, LastRunDate, RunStatusDesc, RunTimeInSeconds, RunTime, Message)" +
                                                                          " VALUES (@ServerName, @CategoryName, @Name, @OwnerID, @Enabled, @Scheduled, @Description, @ReportServerJob, @FrequencyType, @Occurence, @Frequency, @AverageDurationSeconds, @AverageDuration, @NextScheduledRunDate, @LastRunDate, @RunStatusDesc, @RunTimeInSeconds, @RunTime, @Message)";
                    command.Parameters.AddWithValue("@ServerName", serverName);
                    command.Parameters.AddWithValue("@CategoryName", categoryName);
                    command.Parameters.AddWithValue("@Name", name);
                    command.Parameters.AddWithValue("@OwnerID", ownerID);
                    command.Parameters.AddWithValue("@Enabled", enabled);
                    command.Parameters.AddWithValue("@Scheduled", scheduled);
                    command.Parameters.AddWithValue("@Description", description);
                    command.Parameters.AddWithValue("@ReportServerJob", reportServerJob);
                    command.Parameters.AddWithValue("@FrequencyType", frequencyType);
                    command.Parameters.AddWithValue("@Occurence", occurence);
                    command.Parameters.AddWithValue("@Frequency", frequency);
                    command.Parameters.AddWithValue("@AverageDurationSeconds", averageDurationSeconds);
                    command.Parameters.AddWithValue("@AverageDuration", averageDuration);
                    command.Parameters.AddWithValue("@NextScheduledRunDate", nextScheduledRunDate);
                    command.Parameters.AddWithValue("@LastRunDate", lastRunDate);
                    command.Parameters.AddWithValue("@RunStatusDesc", runStatusDesc);
                    command.Parameters.AddWithValue("@RunTimeInSeconds", runTimeInSeconds);
                    command.Parameters.AddWithValue("@RunTime", runTime);
                    command.Parameters.AddWithValue("@Message", message);

                    command.ExecuteNonQuery();
                }
            }
        }

    }
}