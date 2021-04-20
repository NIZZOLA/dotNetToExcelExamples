using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;

namespace ConsoleExcelImport
{
    class Program
    {
        static void Main(string[] args)
        {
            string filePathName = System.IO.Directory.GetCurrentDirectory() + "\\Resources\\Despesas.xlsx";
            if (File.Exists(filePathName))
            {
                List<DespesasModel> lista = new List<DespesasModel>();              //criamos uma lista vazia para receber cada linha
                var workbook = new XLWorkbook(filePathName);                        // abrimos o objeto do tipo XLWorkbook 
                var nonEmptyDataRows = workbook.Worksheet(1).RowsUsed();            // obtem apenas as linhas que foram utilizadas da planilha

                foreach (var dataRow in nonEmptyDataRows)                           // percorremos linha a linha da planilha
                {
                    if (dataRow.RowNumber() > 1)                                    //obteremos apenas após a linha 1 para não carregar o cabeçalho
                    {
                        var despesa = new DespesasModel();                          // criamos um objeto para popular com os valores obtidos da linha
                        despesa.Id = Convert.ToInt32(dataRow.Cell(1).Value);        // obtemos o valor de cada célula pelo seu nº de coluna
                        despesa.Fornecedor = dataRow.Cell(2).Value.ToString();
                        despesa.ValorDevido = Convert.ToDecimal(dataRow.Cell(3).Value);
                        despesa.Descricao = dataRow.Cell(7).Value.ToString();

                        DateTime.TryParse(dataRow.Cell(4).Value.ToString(), out DateTime dataVencto);
                        despesa.Vencimento = dataVencto;

                        if (!string.IsNullOrEmpty(dataRow.Cell(5).Value.ToString()))
                        {
                            DateTime.TryParse(dataRow.Cell(5).Value.ToString(), out DateTime dataPagto);
                            despesa.Pagamento = dataPagto;
                        }

                        if ( !string.IsNullOrEmpty( dataRow.Cell(6).Value.ToString()))
                                despesa.ValorPago = Convert.ToDecimal(dataRow.Cell(6).Value);
                        
                        lista.Add(despesa);                                         // adicionamos o objeto criado à lista
                    }
                }
                Console.WriteLine(JsonSerializer.Serialize(lista));                 // pronto, exibimos a lista em formato Json
            }
            else
                Console.WriteLine("File not found:" + filePathName);
        }
    }
}
