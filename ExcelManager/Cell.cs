namespace ExcelManager
{
    public class Cell
    {
        public int Int32Value
        {
            get => (int)DecimalValue;
            set { DecimalValue = value; }
        }

        public long Int64Value
        {
            get => (long)DecimalValue;
            set { DecimalValue = value; }
        }

        public decimal DecimalValue
        {
            get => Value == null ? -1M : (decimal)Value;
            set { Value = value; type = null; }
        }

        public string? StringValue
        {
            get => Value == null ? default : Value.ToString();
            set { Value = value; type = "s"; }
        }


        public object? Value { get; private set; }

        internal string? type = null;
    }
}