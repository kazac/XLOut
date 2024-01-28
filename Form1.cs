using System.Data;

namespace XLOut
{
    /// <summary>The aim has to have full fidelity copy from the source to Excel (not rounded or transformed).</summary>
    
    public partial class Form1 : Form
    {
        DataTable tbl = new();

        public Form1()
        {
            InitializeComponent();
            // create a simple data table with a varity of data types as a source.
            tbl.Columns.Add("A", typeof(int));
            tbl.Columns.Add("B", typeof(string));
            tbl.Columns.Add("C", typeof(DateTime));
            tbl.Columns.Add("D", typeof(decimal));
            tbl.Rows.Add(1, "x", DateTime.Now.AddMinutes(1), 1.1m);
            tbl.Rows.Add(2, "y", DateTime.Now.AddMinutes(5), 2.2m);
            tbl.Rows.Add(3, "z", DateTime.Now.AddMinutes(66), 3.3m);

        }

        private void button1_Click(object sender, EventArgs e)
        {
            var sb = new System.Text.StringBuilder();
            sb.Append(DocHeader);
            var rowOne = true;
            foreach (DataRow row in tbl.Rows)
                rowOne = DocRow(row, sb, rowOne);

            sb.Append(DocFooter);
            PasteDoc(sb);
        }

        private static string DocHeader { get; } =
        """
            <?xml version="1.0" encoding="UTF-8"?><?mso-application progid="Excel.Sheet"?>
            <Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet" xmlns:html="http://www.w3.org/TR/REC-html40">
            <Styles>
                <Style ss:ID="s21"><NumberFormat ss:Format="Fixed"/></Style>
                <Style ss:ID="s62"><NumberFormat ss:Format="Short Date"/></Style>
                <Style ss:ID="s65"><NumberFormat ss:Format="Standard"/></Style>
                <Style ss:ID="s66"><NumberFormat ss:Format="0"/></Style>
                <Style ss:ID="s67"><NumberFormat ss:Format="@"/></Style>
                <Style ss:ID="s68"><NumberFormat ss:Format="_-&quot;$&quot;* #,##0.00_-;\-&quot;$&quot;* #,##0.00_-;_-&quot;$&quot;* &quot;-&quot;??_-;_-@_-"/></Style>
                <Style ss:ID="s72"><NumberFormat ss:Format="[$-C09]dd\-mmm\-yy; @"/></Style>
            </Styles>
            <Worksheet ss:Name="Max2 Data"> 
            <Table>
            """;

        private static string DocFooter { get; } = "</Table></Worksheet></Workbook>";

        private static bool DocRow(DataRow row, System.Text.StringBuilder sb, bool rowOne)
        {
            if (rowOne)
            {
                sb.Append("<Row>");
                foreach (DataColumn column in row.Table.Columns)
                {
                    sb.Append($"""<Cell ss:StyleID="s67"><Data ss:Type="String">{column.ColumnName}</Data></Cell>""");
                }
                sb.Append("</Row>");
                sb.AppendLine("");
            }
            sb.Append("<Row>");
            sb.AppendLine("");
            foreach (var cell in row.ItemArray)
            {
                var s = """<Cell ss:StyleID="s67"><Data ss:Type="String"></Data></Cell>""";
                s = cell switch
                {
                    string x => $"""<Cell ss:StyleID="s67"><Data ss:Type="String">{x}</Data></Cell>""",
                    int x => $"""<Cell ss:StyleID="s66"><Data ss:Type="Number">{x.ToString("D")}</Data></Cell>""",
                    double x => $"""<Cell ss:StyleID="s21"><Data ss:Type="Number">{x.ToString("R")}</Data></Cell>""",
                    decimal x => $"""<Cell ss:StyleID="s21"><Data ss:Type="Number">{x.ToString("G")}</Data></Cell>""",
                    DateTime x => $"""<Cell ss:StyleID="s62"><Data ss:Type="DateTime">{x.ToString("s")}</Data></Cell>""",
                    _ => throw new NotImplementedException()
                };
                sb.Append(s);
            }
            sb.AppendLine("");
            sb.Append("</Row>");
            sb.AppendLine("");
            return false;
        }

        public static void PasteDoc(System.Text.StringBuilder sb)
        {
            var dataObject = new System.Windows.Forms.DataObject();
            var bytes = System.Text.Encoding.UTF8.GetBytes(sb.ToString());
            var stream = new System.IO.MemoryStream(bytes);
            dataObject.SetData("XML Spreadsheet", stream);
            System.Windows.Forms.Clipboard.SetDataObject(dataObject, true);
        }
    }
}
