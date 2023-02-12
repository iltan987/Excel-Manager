using System.Xml;

namespace ExcelManager
{
    public class Worksheet
    {
        public string Name { get; set; }
        internal XmlDocument? xmlDocument { get; set; }
        public Dictionary<string, Dictionary<string, Cell>> Rows { get; internal set; }

        public Worksheet(string Name)
        {
            this.Name = Name;
            Rows = new Dictionary<string, Dictionary<string, Cell>>();
        }

        public Cell this[string row, string col]
        {
            get
            {
                if (!Rows.ContainsKey(row))
                {
                    Rows[row] = new Dictionary<string, Cell>
                    {
                        [col] = new Cell()
                    };
                }
                else if (!Rows[row].ContainsKey(col))
                    Rows[row][col] = new Cell();
                return Rows[row][col];
            }

            set
            {
                if (value != null)
                    Rows[row][col] = value;
            }
        }
    }
}