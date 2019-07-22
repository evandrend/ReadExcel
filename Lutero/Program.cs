using System;
using System.IO;
using System.Text;
using ClosedXML.Excel;

namespace Lutero
{
    class Program
    {
        static void Main(string[] args)
        {


            StreamWriter wrProvincia = new StreamWriter(File.Open(@"C: \Users\utilizador\Documents\tipos-insert-values-provincia.txt", FileMode.OpenOrCreate), Encoding.UTF8);
            StreamWriter wrMunicipio = new StreamWriter(File.Open(@"C: \Users\utilizador\Documents\tipos-insert-values-municipio.txt", FileMode.OpenOrCreate), Encoding.UTF8);
            StreamWriter wrComunas = new StreamWriter(File.Open(@"C: \Users\utilizador\Documents\tipos-insert-values-comunas.txt", FileMode.OpenOrCreate), Encoding.UTF8);
            //wr.WriteLine();



            //Abrir Arquivo XL
            XLWorkbook livro = new XLWorkbook(@"C:\Users\utilizador\Downloads\lista_de_municpios_e_provncias_de_angola_2.xlsx");
            var folha = livro.Worksheet(1);
            //For provincia
            int idPro = 0;
            int idMunicipio = 0;
            for (int i = 7; i <=188; i++)
            {
                string provindia = folha.Cell(i, 1).Value.ToString();
                if (string.IsNullOrWhiteSpace(provindia))
                {
                    string municipio=folha.Cell(i, 2).Value.ToString();
                    var comunas=folha.Cell(i, 3).Value.ToString().Split(new string[] { "," }, StringSplitOptions.None);
                    idMunicipio = i;
                    wrMunicipio.WriteLine($"({idMunicipio},'{municipio}',{idPro})\n");
                    Console.WriteLine($"--->{municipio} \n");



                    for (int k = 0; k < comunas.Length; k++)
                    {
                        Console.WriteLine("::::::::>"+comunas[k]);
                        wrComunas.WriteLine($"('default','{comunas[k]}',{idMunicipio})\n");

                    }
                    
                    
                }
                else
                {
                    idPro = i + 100;
                    wrProvincia.WriteLine($"({idPro},'{provindia}'),\n");

                    Console.WriteLine($"\n {provindia} \n");

                }



            }
            wrMunicipio.Close();
            wrComunas.Close();
            wrProvincia.Close();

            





         
            //wr.Close();
            //Console.WriteLine(result);
            Console.ReadKey();
        }

        //static void Main(string[] args)
        //{
        //    string result="";

        //    StreamReader rd = new StreamReader(@"C:\Users\utilizador\Documents\tipos.txt");

        //    while (!rd.EndOfStream)
        //    {
        //        var linha = rd.ReadLine();
        //        Console.WriteLine(linha);
        //        //string linhaFormatada = l+linha.ToLower();
        //        result += $"('{linha}'),\n";
        //    }
        //    rd.Close();
           

        //    StreamWriter wr = new StreamWriter(File.Open(@"C: \Users\utilizador\Documents\tipos - insert - values.txt",FileMode.Open),Encoding.UTF8);
           
        //    wr.WriteLine(result);
        //    wr.Close();
        //    Console.WriteLine(result);
        //    Console.ReadKey();
        //}
    }
}
