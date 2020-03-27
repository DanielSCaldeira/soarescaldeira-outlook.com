using OfficeOpenXml;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Reflection;

namespace GerarHtmlExcel
{
    class GerarHtmlExcel
    {
        List<string> listaDeNumerosPercorridos = new List<string>();

        private const string ALPHABET = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";

        public void GerarExel()
        {
            string path = "C:\\Users\\danielcaldeira\\Desktop\\Excel.xlsx";
            FileInfo theFile = new FileInfo(path);
            using (ExcelPackage xlPackage = new ExcelPackage(theFile))
            {
                string htmll = $@"<!DOCTYPE html>
                                <html>
                                    <head>
                                        <meta charset='utf-8'/> 
                                        <meta http - equiv ='Content-Language' content='pt-br'>      
                                        <meta name ='description' content='RH WEB'/>         
                                        <meta name ='viewport' content='width=device-width'/>            
                                        <meta http-equiv='X-UA-Compatible' content ='IE=Edge'>  
                                        <link href='excel.css' rel='stylesheet' />
                                        <title>Excel Gerado</title>
                                    </head>
                                <body><table>";

                ExcelWorksheet worksheet = xlPackage.Workbook.Worksheets.FirstOrDefault();
                int rows = worksheet.Dimension.Rows;
                int columns = worksheet.Dimension.Columns;

                for (int r = 1; r < rows; r++)
                {
                    var linha = "<tr>";
                    var coluna = "";
                    for (int c = 1; c < columns; c++)
                    {
                        var existe = $"{c}:{r}";
                        if (!listaDeNumerosPercorridos.Any(x => x == existe))
                        {
                            //está selula esta mesclada
                            if (worksheet.MergedCells[r, c] != null)
                            {
                                var mergedadress = worksheet.MergedCells[r, c];
                                var array = mergedadress.Split(':');

                                var primeiraCelula = ToNumericCoordinates(array[0]).Split(','); //coluna -- linha 
                                var segundaCelula = ToNumericCoordinates(array[1]).Split(','); //coluna -- linha

                                var resultado = AdicionaListaDeNumerosPercorridos(primeiraCelula, segundaCelula);
                                var qtdColunas = resultado.coluna;
                                var qtdLinhas = resultado.linha;

                                var valorDaCelula = worksheet.Cells[mergedadress].Value;
                                string[] textos = PegaOstextosDaCelulaMerjada(valorDaCelula);

                                var texto = "";
                                foreach (var item in textos)
                                {
                                    texto += item;
                                }
                                coluna += $"<td {(qtdColunas == 0 ? "" : $"colspan='{qtdColunas}'")} {(qtdLinhas == 0 ? "" : $"rowspan='{qtdLinhas}'")}>{texto}</td>";
                            }
                            else
                            {
                                coluna += $"<td>{worksheet.Cells[r, c].Value}</td>";
                            }
                        }

                    }
                    linha += coluna;
                    linha += "</tr>";
                    htmll += linha;
                }
                htmll += $@"</table></body></html>";
                var pronto = htmll.Replace("'", "\"");
            }
        }

        public (int coluna, int linha) AdicionaListaDeNumerosPercorridos(string[] primeiraCelula, string[] segundaCelula)
        {

            var colCelula1 = Convert.ToInt32(primeiraCelula[0]);
            var rowCelula1 = Convert.ToInt32(primeiraCelula[1]);

            var colCelula2 = Convert.ToInt32(segundaCelula[0]);
            var rowCelula2 = Convert.ToInt32(segundaCelula[1]);

            List<string> row = new List<string>();
            if (colCelula1 == colCelula2)
            {
                for (int linha = rowCelula1; linha <= rowCelula2; linha++)
                {
                    var celula = $"{colCelula1}:{linha}";
                    listaDeNumerosPercorridos.Add(celula);
                    row.Add(celula);
                }
            }
            List<string> coll = new List<string>();
            if (rowCelula1 == rowCelula2)
            {
                for (int coluna = colCelula1; coluna <= colCelula2; coluna++)
                {
                    var celula = $"{coluna}:{rowCelula1}";
                    listaDeNumerosPercorridos.Add(celula);
                    coll.Add(celula);
                }
            }

            return (coll.Count(), row.Count());
        }

        public string[] PegaOstextosDaCelulaMerjada(object valorDaCelula)
        {

            string[] textos = ((IEnumerable)valorDaCelula).Cast<object>().Where(y => y != null)
              .Select(x => x.ToString())
              .ToArray();

            return textos;
        }

        public static string ToExcelCoordinates(string coordinates)
        {
            string first = coordinates.Substring(0, coordinates.IndexOf(','));
            int i = int.Parse(first);
            string second = coordinates.Substring(first.Length + 1);

            string str = string.Empty;
            while (i > 0)
            {
                str = ALPHABET[(i - 1) % 26] + str;
                i /= 26;
            }

            return str + second;
        }

        public static string ToNumericCoordinates(string coordinates)
        {
            string first = string.Empty;
            string second = string.Empty;

            CharEnumerator ce = coordinates.GetEnumerator();
            while (ce.MoveNext())
                if (char.IsLetter(ce.Current))
                    first += ce.Current;
                else
                    second += ce.Current;

            int i = 0;
            ce = first.GetEnumerator();
            while (ce.MoveNext())
                i = (26 * i) + ALPHABET.IndexOf(ce.Current) + 1;

            string str = i.ToString();
            return str + "," + second;
        }
    }

}
