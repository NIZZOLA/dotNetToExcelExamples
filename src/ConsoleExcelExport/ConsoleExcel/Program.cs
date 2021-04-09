using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;

namespace ConsoleExcel
{
    class Program
    {
        static void Main(string[] args)
        {
            var lista = GenerateData();
            GenerateFile("Despesas", "Despesas.xlsx", lista);
        }


        static void GenerateHeader( IXLWorksheet worksheet)
        {
            worksheet.Cell("A1").Value = "Código";
            worksheet.Cell("B1").Value = "Fornecedor";
            worksheet.Cell("C1").Value = "Valor R$";
            worksheet.Cell("D1").Value = "Vencimento";
            worksheet.Cell("E1").Value = "Pagamento";
            worksheet.Cell("F1").Value = "Valor Pago";
            worksheet.Cell("G1").Value = "Descrição";
        }

        static void GenerateFile(string tabName, string fileName, ICollection<DespesasModel> lista)
        {
            string filePathName = System.IO.Directory.GetCurrentDirectory() + "\\"+ fileName;
            
            if (File.Exists(filePathName))
                File.Delete(filePathName);

            using (var workbook = new XLWorkbook())
            {
                var planilha = workbook.Worksheets.Add(tabName);

                int line = 1;
                GenerateHeader(planilha);
                line++;

                foreach (var item in lista)
                {
                    planilha.Cell("A" + line).Value = item.Id;
                    planilha.Cell("B" + line).Value = item.Fornecedor;
                    planilha.Cell("C" + line).Value = item.ValorDevido;
                    planilha.Cell("D" + line).Value = item.Vencimento;
                    planilha.Cell("E" + line).Value = item.Pagamento;
                    planilha.Cell("F" + line).Value = item.ValorPago;
                    planilha.Cell("G" + line).Value = item.Descricao;
                    line++;
                }
                workbook.SaveAs(filePathName);
            }
        }


        static List<DespesasModel> GenerateData()
        {
            List<DespesasModel> retorno = new List<DespesasModel>();
            retorno.Add(new DespesasModel() { Id = 1, Fornecedor = "FABRICA A", ValorDevido = 500, Vencimento = DateTime.Today.AddDays(30), Descricao = "Conta usada para teste", Pagamento = DateTime.Today.AddDays(30), ValorPago = 500 });
            retorno.Add(new DespesasModel() { Id = 1, Fornecedor = "FABRICA B", ValorDevido = 600, Vencimento = DateTime.Today.AddDays(30), Descricao = "Conta usada para teste", Pagamento = DateTime.Today.AddDays(30), ValorPago = 600 });
            retorno.Add(new DespesasModel() { Id = 1, Fornecedor = "FABRICA C", ValorDevido = 700, Vencimento = DateTime.Today.AddDays(30), Descricao = "Conta usada para teste", Pagamento = DateTime.Today.AddDays(30), ValorPago = 700 });
            retorno.Add(new DespesasModel() { Id = 1, Fornecedor = "FABRICA D", ValorDevido = 800, Vencimento = DateTime.Today.AddDays(30), Descricao = "Conta usada para teste", Pagamento = DateTime.Today.AddDays(30), ValorPago = 800 });
            return retorno;
        }

        public static string[] GetFieldNames(Type t)
        {
            FieldInfo[] fieldInfos = t.GetFields(BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);
            return fieldInfos.Select(x => x.Name).ToArray();
        }

    }
}
