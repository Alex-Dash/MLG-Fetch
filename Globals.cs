using System;
using System.Collections.Generic;
using System.Windows.Forms;


namespace MLG_Fetch
{
    public static class Globals
    {
        public const Int32 BUFFER_SIZE = 512; // Unmodifiable
        public static String FILE_NAME = "Output.txt"; // Modifiable
        public static readonly String CODE_PREFIX = "US-"; // Unmodifiable
        
        public static OpenFileDialog tableFileDialog = new OpenFileDialog();
        public static Int32 REGIONS_COUNT = 38;
        public static Int32 SheetId = 1;
        public static string LAST_DATA_MESSAGES = "";
        public static String HISTORY_FILE_NAME = "_"; // Modifiable
        //public static Excel EXCEL_SHEET = new Excel(FILE_NAME, SheetId);
        public static IDictionary<int, string> RegionsID = new Dictionary<int, string>();
        public static IDictionary<int, string> RegionsIDF = new Dictionary<int, string>();
        public static IDictionary<string, int[]> CumSum = new Dictionary<string, int[]>();
        public static bool CSumOk = true;
        public static Microsoft.Office.Interop.Word.Application app;
        public static Microsoft.Office.Interop.Word.Document doc;

        public static IDictionary<int, string> Deps2 = new Dictionary<int, string>();
        public static IDictionary<int, string> DepsSel = new Dictionary<int, string>();

        public static IDictionary<string, string[]> MediaDB = new Dictionary<string, string[]>();
        public static string MediaMode = "OR";

        public static int SecretCounter = 0;

        //public static IDictionary<int, string[]> DepsReg = new Dictionary<int, string[]>();
        //public static IDictionary<int, string[]> DepsRegSel = new Dictionary<int, string[]>();



    }
}
