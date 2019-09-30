using System;
using System.Linq;

namespace Dipu.Excel.DataTable.Tests
{
    public class Population
    {
        private CompanyModel[] _companies;
        private ProductModel[] _products;

        public Column<CompanyModel>[] DefaultCompanyColumns
        {
            get
            {
                var companyColumns = new[]
                {
                    new Column<CompanyModel> {ToExcel = v => v.Name, ToDomain = (model,value) => model.Name = Convert.ToString(value)},
                    new Column<CompanyModel> {ToExcel = v => v.KeyWords[0], ToDomain = (model,value) => model.KeyWords[0] = Convert.ToString(value)},
                    new Column<CompanyModel> {ToExcel = v => v.KeyWords[1], ToDomain = (model,value) => model.KeyWords[1] = Convert.ToString(value)},
                };

                return companyColumns;
            }
        }

        public Column<ProductModel>[] DefaultProductColumns
        {
            get
            {
                var productColumns = new[]
                {
                    new Column<ProductModel> {ToExcel = v => v.Name, ToDomain = (model, name) => model.Name = name.ToString()},
                    new Column<ProductModel> {ToExcel = v => v.Description, ToDomain = (model, description) => model.Description = description.ToString()},
                    new Column<ProductModel> {ToExcel = v => v.Price, ToDomain =  (product, price) => product.Price = Convert.ToDecimal(price)},
                    new Column<ProductModel> {ToExcel = v => v.Manufacturer?.Name, ToDomain =  (product, name) => product.Manufacturer = this.Companies.FirstOrDefault(v => string.Equals(v.Name, (string) name, StringComparison.OrdinalIgnoreCase))},
                };

                return productColumns;
            }
        }

        public CompanyModel[] Companies
        {
            get
            {
                return _companies ?? (_companies = new[]
                {
                    new CompanyModel {Name = "Pizza Enzo", KeyWords = new[] {"enzo", "napoli", "pizza"}},
                    new CompanyModel {Name = "Di Piu", KeyWords = new[] {"pizza", "calabria", "maasmechelen"}},
                    new CompanyModel {Name = "Pizza Napoli", KeyWords = new[] {"pizza", "napoli", "maasmechelen"}},
                });
            }
        }

        public ProductModel[] Products
        {
            get
            {
                return _products ?? (_products = new[]
                {
                    new ProductModel{ Name = "Pizza Margherita", Description = "Best pizza in town", Manufacturer = _companies[0], Price = 6.5M},
                    new ProductModel{ Name = "Pizza Napoli", Description = "Original Italian pizza", Manufacturer = _companies[0], Price = 6.8M},
                });
            }
        }
    }

    public class CompanyModel
    {
        public string Name { get; set; }

        public string[] KeyWords { get; set; }
    }

    public class ProductModel
    {
        public string Name { get; set; }

        public string Description { get; set; }

        public decimal Price { get; set; }

        public CompanyModel Manufacturer { get; set; }
    }
}
