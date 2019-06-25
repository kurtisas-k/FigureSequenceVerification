using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;


namespace FigureSequenceVerification
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            string outPath = @"C:\Users\kstaples\Documents\Projects\Update ILMS\outreport.csv";
            List<string> outData = new List<string>();
            Dictionary<string, int[]> dictionaryOfModuleFigureIndexNumbers = new Dictionary<string, int[]>();
            Console.Write(DateTime.Now);
            Microsoft.Office.Interop.Word.Application wrdApp = new Microsoft.Office.Interop.Word.Application();
            wrdApp.Visible = false;
            // for all docs
            var files = Directory.GetFiles(@"C:\Users\kstaples\Documents\Projects\Update ILMS\Modules", "*.docx", SearchOption.AllDirectories);
            
            // for all files try and find associative module
            for (var i = 0; i < files.Length; i++)
            {
                Console.WriteLine(DateTime.Now+", "+i);
                if (files[i].IndexOf("~") > -1) { continue; }
                var doc = wrdApp.Documents.Open(files[i],false,true);
                getArrayOfFigureIndexNumbers(doc, wrdApp, dictionaryOfModuleFigureIndexNumbers);
                doc.Close(SaveChanges: false);
            }

            foreach(KeyValuePair<string, int[]> entry in dictionaryOfModuleFigureIndexNumbers)
            {
                var arrFigureInts = entry.Value;
                if (figureIdsAreSpacedByOne(arrFigureInts) == false)
                {
                    outData.Add(entry.Key);
                };
            }

            File.WriteAllLines(outPath,outData);
        }

        private bool figureIdsAreSpacedByOne(int[] fieldIds)
        {
            for(var i = 0; i < fieldIds.Length-1; i++)
            {
                //Console.WriteLine(fieldIds[i + 1] - fieldIds[i]);
                if((fieldIds[i + 1] - fieldIds[i]) == 0)
                {
                    return false;
                }
            }
            return true;
        }

        private void getArrayOfFigureIndexNumbers(Document doc, Microsoft.Office.Interop.Word.Application wrdApp, Dictionary<string, int[]> moduleFigureArray)
        {
            int[] figureNumberArray = findAllInstancesOfWordFigure(doc, wrdApp);
            moduleFigureArray[doc.Name] = figureNumberArray;
        }

        private int[] findAllInstancesOfWordFigure(Document doc, Microsoft.Office.Interop.Word.Application wrdApp)
        {
            var start = 0;
            var end = 0;
            Console.WriteLine(doc.Name);
            List<int> ints = new List<int>();
            var intFound = 0;
            var rng = doc.Content;
            var missing = System.Type.Missing;
            rng.Find.ClearFormatting();
            rng.Find.Forward = true;
            rng.Find.Text = "Figure";

            rng.Find.Execute(
                ref missing, ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing, ref missing);

            while (rng.Find.Found)
            {
                intFound++;
                rng.Find.Execute(
                    ref missing, ref missing, ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing, ref missing, ref missing);
                Range extendedInstanceOfFigure = doc.Range(rng.Start, rng.End + 4);
                if(extendedInstanceOfFigure.ParagraphStyle != null && extendedInstanceOfFigure.ParagraphStyle.NameLocal != null)
                {
                    if (extendedInstanceOfFigure.ParagraphStyle.NameLocal == "Comment" || extendedInstanceOfFigure.ParagraphStyle.NameLocal == "Caption")
                    {
                        // if the range is the same 
                        if (rng.Start == start && rng.End == end)
                        {

                        }
                        else
                        {
                            start = rng.Start;
                            end = rng.End;
                            string[] numbers = Regex.Split(extendedInstanceOfFigure.Text, @"\D+");
                            foreach (var value in numbers)
                            {
                                if (!string.IsNullOrEmpty(value))
                                {
                                    int i = int.Parse(value);
                                    ints.Add(i);
                                    //Console.WriteLine(i);
                                }
                            }
                        }

                    }
                }                
            }
            return ints.ToArray();
        }
    }
}
