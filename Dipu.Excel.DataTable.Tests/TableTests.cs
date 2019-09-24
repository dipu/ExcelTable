using Dipu.Excel.DataTable;

namespace Dipu.Excel.DataTable.Tests
{
    using Microsoft.Office.Interop.Excel;
    using Xunit;

    public class TableTests : Test
    {
        public class Model
        {
            public string Name { get; set; }

            public string[] Values { get; set; }
        }

        [Fact]
        public void BindTest()
        {
            var allorsWorksheet = new AllorsWorksheet((Worksheet)this.Workbook.Sheets[1]);

            var columns = new[]
            {
                new Column<Model> {ToExcel = v => v.Name},
                new Column<Model> {ToExcel = v => v.Values[0]},
                new Column<Model> {ToExcel = v => v.Values[1]},
            };

            var table = new Table<Model>(allorsWorksheet, columns, 1, 1);

            var values = new[]
            {
                new Model{ Name = "Walter", Values = new []{"a1", "a2"}},
                new Model{ Name = "Koen", Values = new []{"b1", "b2"}},
            };

            table.Bind(values);

            Assert.Equal("Walter", table.Rows[0].Cells[0].Value);
            Assert.Equal("a1", table.Rows[0].Cells[1].Value);
            Assert.Equal("a2", table.Rows[0].Cells[2].Value);

            Assert.Equal("Koen", table.Rows[1].Cells[0].Value);
            Assert.Equal("b1", table.Rows[1].Cells[1].Value);
            Assert.Equal("b2", table.Rows[1].Cells[2].Value);
        }

        [Fact]
        public void RangesTest()
        {
            var allorsWorksheet = new AllorsWorksheet((Worksheet)this.Workbook.Sheets[1]);

            var columns = new[]
            {
                new Column<Model> {ToExcel = v => v.Name, ToDomain = (model,value) => model.Name = value.ToString().ToUpper()},
                new Column<Model> {ToExcel = v => v.Values[0]},
                new Column<Model> {ToExcel = v => v.Values[1]},
            };

            var table = new Table<Model>(allorsWorksheet, columns, 1, 1);

            var values = new[]
            {
                new Model{ Name = "Walter", Values = new []{"a1", "a2"}},
                new Model{ Name = "Koen", Values = new []{"b1", "b2"}},
                new Model{ Name = "Martien", Values = new []{"c1", "c2"}},
            };

            table.Bind(values);

            var ranges = table.Flush();

            Assert.Single(ranges);
            Assert.Equal(0, ranges[0][0]);
            Assert.Equal(2, ranges[0][1]);

            table.Bind(values);

            ranges = table.Flush();

            Assert.Empty(ranges);

            values[1].Values[1] = "b2_1";
            table.Bind(values);

            ranges = table.Flush();

            Assert.Single(ranges);
            Assert.Equal(1, ranges[0][0]);
            Assert.Equal(1, ranges[0][1]);
        }

    }
}
