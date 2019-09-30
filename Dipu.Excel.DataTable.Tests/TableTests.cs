using System;
using System.Linq;

namespace Dipu.Excel.DataTable.Tests
{
    using Microsoft.Office.Interop.Excel;
    using Xunit;

    public class TableTests : Test
    {
        [Fact]
        public void BindTest()
        {
            var allorsWorksheet = new AllorsWorksheet((Worksheet)this.Workbook.Sheets[1]);
            
            var table = new Table<CompanyModel>(allorsWorksheet, this.Population.DefaultCompanyColumns, 1, 1);
            
            table.Bind(this.Population.Companies);

            Assert.Equal("Pizza Enzo", table.Rows[0].Cells[0].Value);
            Assert.Equal("enzo", table.Rows[0].Cells[1].Value);
            Assert.Equal("napoli", table.Rows[0].Cells[2].Value);

            Assert.Equal("Di Piu", table.Rows[1].Cells[0].Value);
            Assert.Equal("pizza", table.Rows[1].Cells[1].Value);
            Assert.Equal("calabria", table.Rows[1].Cells[2].Value);
        }

        [Fact]
        public void RangesTest()
        {
            var allorsWorksheet = new AllorsWorksheet((Worksheet)this.Workbook.Sheets[1]);
            
            var table = new Table<CompanyModel>(allorsWorksheet, this.Population.DefaultCompanyColumns, 1, 1);
            
            table.Bind(this.Population.Companies);

            var ranges = table.Flush();

            Assert.Single(ranges);
            Assert.Equal(0, ranges[0][0]);
            Assert.Equal(2, ranges[0][1]);

            // When we bind the same objects, nothing has changed
            table.Bind(this.Population.Companies);
            ranges = table.Flush();
            Assert.Empty(ranges);

            // When we change a single property, then that is the only range we need to update
            this.Population.Companies[1].KeyWords[1] = "10% discount on all pizza's";
            table.Bind(this.Population.Companies);
            ranges = table.Flush();

            Assert.Single(ranges);
            Assert.Equal(1, ranges[0][0]);
            Assert.Equal(1, ranges[0][1]);
        }


        [Fact]
        public void RangesCanStartAtA2Test()
        {
            var allorsWorksheet = new AllorsWorksheet((Worksheet)this.Workbook.Sheets[1]);
            
            var table = new Table<CompanyModel>(allorsWorksheet, this.Population.DefaultCompanyColumns, 2, 1);
            table.Bind(this.Population.Companies);

            var ranges = table.Flush();
            
            Assert.Single(ranges);
            Assert.Equal(0, ranges[0][0]);
            Assert.Equal(2, ranges[0][1]);

            table.Bind(this.Population.Companies);

            ranges = table.Flush();

            Assert.Empty(ranges);

            this.Population.Companies[1].KeyWords[1] = "sicilia";
            table.Bind(this.Population.Companies);

            ranges = table.Flush();

            Assert.Single(ranges);
            Assert.Equal(1, ranges[0][0]);
            Assert.Equal(1, ranges[0][1]);
        }

        [Fact]
        public void RangesCanStartAtB2_Test()
        {
            var allorsWorksheet = new AllorsWorksheet((Worksheet)this.Workbook.Sheets[1]);
            
            var table = new Table<CompanyModel>(allorsWorksheet, this.Population.DefaultCompanyColumns, 2, 2);
            
            table.Bind(this.Population.Companies);

            var ranges = table.Flush();
            
            Assert.Single(ranges);
            Assert.Equal(0, ranges[0][0]);
            Assert.Equal(2, ranges[0][1]);

            table.Bind(this.Population.Companies);

            ranges = table.Flush();

            Assert.Empty(ranges);

            this.Population.Companies[1].KeyWords[1] = "new Keyword";
            table.Bind(this.Population.Companies);

            ranges = table.Flush();

            Assert.Single(ranges);
            Assert.Equal(1, ranges[0][0]);
            Assert.Equal(1, ranges[0][1]);
        }

        [Fact]
        public void TableCanHaveManyRanges()
        {
            var allorsWorksheet = new AllorsWorksheet((Worksheet)this.Workbook.Sheets[1]);
            
            var companyTable = new Table<CompanyModel>(allorsWorksheet, this.Population.DefaultCompanyColumns, 2, 2);
            
            companyTable.Bind(this.Population.Companies);
            companyTable.Flush();
            
            var productTable = new Table<ProductModel>(allorsWorksheet, this.Population.DefaultProductColumns, 8, 5);
            
            productTable.Bind(this.Population.Products);
            productTable.Flush();

            // Changing the Manufacturer will only update that cell!
            this.Population.Products[1].Manufacturer = this.Population.Companies[1];
            productTable.Bind(this.Population.Products);

            var ranges = productTable.Flush();

            Assert.Single(ranges);
            Assert.Equal(1, ranges[0][0]);
            Assert.Equal(1, ranges[0][1]);
        }

        [Fact]
        public void TableCanHaveManyRangesOnSameRow()
        {
            var allorsWorksheet = new AllorsWorksheet((Worksheet)this.Workbook.Sheets[1]);
            
            var companyTable = new Table<CompanyModel>(allorsWorksheet, this.Population.DefaultCompanyColumns, 1, 1);
            
            companyTable.Bind(this.Population.Companies);
            companyTable.Flush();
            
            var productTable = new Table<ProductModel>(allorsWorksheet, this.Population.DefaultProductColumns, 8, 5);
            
            productTable.Bind(this.Population.Products);
            productTable.Flush();

            // Changing the Manufacturer will only update that cell!
            this.Population.Products[1].Manufacturer = this.Population.Companies[1];
            productTable.Bind(this.Population.Products);

            var ranges = productTable.Flush();

            Assert.Single(ranges);
            Assert.Equal(1, ranges[0][0]);
            Assert.Equal(1, ranges[0][1]);
        }

        [Fact]
        public void TableCanNotHaveOverlappingRanges()
        {
            var allorsWorksheet = new AllorsWorksheet((Worksheet)this.Workbook.Sheets[1]);
            
            var companyTable = new Table<CompanyModel>(allorsWorksheet, this.Population.DefaultCompanyColumns, 1, 1);
            companyTable.Bind(this.Population.Companies);
            
            var productTable = new Table<ProductModel>(allorsWorksheet, this.Population.DefaultProductColumns, 1, 2);
            //TODO: this should throw an exception?
        }
    }
}
