
using EmailExtractTest;


//I think I will use this one Email Extract 2
//EmailExtract2 emailExtract2 = new EmailExtract2();

//emailExtract2.ExtractDataFromEmail();



InsertDataJobLog insertDataJobLog = new InsertDataJobLog();

insertDataJobLog.InsertDataIntoTables("D:\\Other Things\\Not for school\\RevRed\\EmailExtractTest\\files", "Data Source=DESKTOP-89NJE0M\\SQLEXPRESS;Initial Catalog=SystemStatus;Integrated Security=True");