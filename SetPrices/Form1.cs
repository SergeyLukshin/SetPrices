using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Globalization;
using EXCEL = Microsoft.Office.Interop.Excel;

namespace SetPrices
{
    public partial class SetPriceForm : Form
    {
        public class PriceInfo
        {
            public PriceInfo()
            {
                bFind = false;
                Cnt = 0;
            }
            public string strPrice;
            public string strPriceEuro;
            public bool bFind;
            public int Cnt;
        }

        public class TypeInfo
        {
            public string strTypeForDB;
            public string strType;
        }

        public Dictionary<string, PriceInfo> dictTovars = new Dictionary<string, PriceInfo>();
        public Dictionary<string, string> dictColors = new Dictionary<string, string>();
        public Dictionary<string, string> dictColorsNoTranslate = new Dictionary<string, string>();
        public Dictionary<string, string> dictSizes = new Dictionary<string, string>();
        public Dictionary<string, Dictionary<int, List< string>>> dictPictures = new Dictionary<string, Dictionary<int, List<string>>>();
        //public Dictionary<string, TypeInfo> dictTypes = new Dictionary<string, TypeInfo>();

        public Dictionary<string, int> dictErrorColors = new Dictionary<string, int>();
        public Dictionary<string, int> dictErrorSizes = new Dictionary<string, int>();
        public Dictionary<string, int> dictErrorPictures = new Dictionary<string, int>();

        public Dictionary<string, KeyValuePair<int, int>> dictErrorDublicates = new Dictionary<string, KeyValuePair<int, int>>();

        public SetPriceForm()
        {
            InitializeComponent();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.CheckFileExists = true;
            dlg.CheckPathExists = true;
            dlg.Filter = "Excel files (*.xls)|*.xls";
            dlg.Multiselect = false;
            if (dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                tbTextFile.Text = dlg.FileName;
            }
        }

        
        public static string FirstUpper(string str)
        {
            if (str == "") return "";
            return str.Substring(0, 1).ToUpper() + (str.Length > 1 ? str.Substring(1) : "");
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (tbTextFile.Text == "")
            {
                MessageBox.Show("Необходимо выбрать файл с товарами.");
                return;
            }

            if (!cbOnlyPrices.Checked)
            {
                // вытаскиваем переводы цвета
                try
                {
                    StreamReader fileColor = File.OpenText(Path.GetDirectoryName(Application.ExecutablePath) + "\\colors.txt");
                    char[] delimitersColor = { '\t' };
                    while (true)
                    {
                        string st = fileColor.ReadLine();
                        if (st == null)
                            break;

                        string[] values = st.Split(delimitersColor);
                        if (values.Count() == 2)
                        {
                            dictColors[values[0].ToLower()] = values[1];
                            dictColorsNoTranslate[values[1].ToLower()] = values[1];
                        }
                    }

                    fileColor.Close();
                }
                catch (Exception)
                {
                }

                // вытаскиваем размеры
                try
                {
                    StreamReader fileColor = File.OpenText(Path.GetDirectoryName(Application.ExecutablePath) + "\\sizes.txt");
                    //char[] delimitersColor = { '\t' };
                    while (true)
                    {
                        string st = fileColor.ReadLine();
                        if (st == null)
                            break;

                        dictSizes[st.ToLower()] = "";
                    }

                    fileColor.Close();
                }
                catch (Exception)
                {
                }


                // получаем список изображений
                string[] files = Directory.GetFiles(Path.GetDirectoryName(tbTextFile.Text), "*.jpg");
                List<string> result = files.ToList();
                for (int i = 0; i < result.Count(); i++)
                {
                    result[i] = result[i].Replace(".jpg", "");
                }
                result.Sort();
                for (int i = 0; i < result.Count(); i++)
                {
                    result[i] = result[i] + ".jpg";
                }
                //var result = files.OrderBy(a => new string(a.ToCharArray().Reverse().ToArray()));
                char[] delimitersFiles = { ',' };
                //char[] delimitersFiles2 = { '#' };
                foreach (string item in result)
                {
                    string strFile = Path.GetFileName(item);
                    strFile = strFile.Replace(".jpg", "");
                    //strFile = strFile.Replace("(", "#");
                    //strFile = strFile.Replace(")", "");
                    string[] values = strFile.Split(delimitersFiles);
                    for (int j = 0; j < values.Count(); j++)
                    {
                        string strTemp = values[j];
                        string strArticul = values[j].ToLower().Replace("#2", "").Replace("#3", "").Replace("#4", "").Replace("#5", "").Replace("#6", "").Replace("#7", "").Trim();
                        int pos = 1;
                        if (!dictPictures.ContainsKey(strArticul))
                        {
                            dictPictures[strArticul] = new Dictionary<int,List<string>>();
                        }
                        if (strTemp.IndexOf("#2") >= 0) pos = 2;
                        if (strTemp.IndexOf("#3") >= 0) pos = 3;
                        if (strTemp.IndexOf("#4") >= 0) pos = 4;
                        if (strTemp.IndexOf("#5") >= 0) pos = 5;
                        if (strTemp.IndexOf("#6") >= 0) pos = 6;
                        if (strTemp.IndexOf("#7") >= 0) pos = 7;

                        while (true)
                        {
                            if (!dictPictures[strArticul].ContainsKey(pos))
                            {
                                dictPictures[strArticul][pos] = new List<string>();
                                break;
                            }

                            pos ++;
                        }
                        dictPictures[strArticul][pos].Add(item);
                    }

                    try
                    {
                        GflAx.GflAx fl = new GflAx.GflAx();

                        fl.LoadBitmap(item);
                        int width2 = fl.width;
                        int height2 = fl.height;

                        fl.SaveFormat = GflAx.AX_SaveFormats.AX_JPEG;
                        if (width2 > height2)
                        {
                            if (width2 < 1000)
                            {
                                height2 = height2 * 1000 / width2;
                                fl.Resize(1000, height2);
                                fl.SaveBitmap(item);
                            }

                            if (width2 > 1600)
                            {
                                height2 = height2 * 1600 / width2;
                                fl.Resize(1600, height2);
                                fl.SaveBitmap(item);
                            }
                        }
                        else
                        {
                            if (height2 < 1000)
                            {
                                width2 = width2 * 1000 / height2;
                                fl.Resize(width2, 1000);
                                fl.SaveBitmap(item);
                            }

                            if (height2 > 1600)
                            {
                                width2 = width2 * 1600 / height2;
                                fl.Resize(width2, 1600);
                                fl.SaveBitmap(item);
                            }
                        };
                        /*if (width2 > 1600)
                        {
                            height2 = height2 * 1600 / width2;
                            fl.Resize(1600, height2);
                            fl.SaveBitmap(files[i]);
                        }*/
                    }
                    catch (Exception)
                    {
                    }
                }

                // вытаскиваем виды товаров
                /*try
                {
                    StreamReader fileType = File.OpenText(Path.GetDirectoryName(Application.ExecutablePath) + "\\types.txt");
                    char[] delimitersType = { '\t' };
                    while (true)
                    {
                        string st = fileType.ReadLine();
                        if (st == null)
                            break;

                        string[] values = st.Split(delimitersType);
                        if (values.Count() == 3)
                        {
                            dictTypes[values[0].ToLower()] = new TypeInfo();
                            dictTypes[values[0].ToLower()].strType = values[1];
                            dictTypes[values[0].ToLower()].strTypeForDB = values[2];
                        }
                    }

                    fileType.Close();
                }
                catch (Exception)
                {
                }*/

                EXCEL.Application app = new EXCEL.Application();
                EXCEL.Workbook wb = app.Workbooks.Open(tbTextFile.Text);
                EXCEL.Worksheet sheet = wb.Sheets[1];

                app.DisplayAlerts = false;

                int columnsCount = sheet.UsedRange.Columns.Count;
                int rowsCount = sheet.UsedRange.Rows.Count;

                dictTovars.Clear();

                char[] delimiters = { ';' };
                char[] delimitersColors = { ',' };
                char[] delimitersColors2 = { '/' };

                StreamWriter fileMain = new StreamWriter(Path.GetDirectoryName(tbTextFile.Text) + "\\main.txt", false, Encoding.GetEncoding(1251));

                for (int i = 2; i <= rowsCount; i++)
                {
                    //string strType = (sheet.Cells[i, 5] as EXCEL.Range).Text;
                    //string strBrand = (sheet.Cells[i, 1] as EXCEL.Range).Text;
                    //string strSeason = (sheet.Cells[i, 2] as EXCEL.Range).Text;
                    //string strArticul = (sheet.Cells[i, 3] as EXCEL.Range).Text;
                    string strArticul = (sheet.Cells[i, 1] as EXCEL.Range).Text;
                    string strSiteName = (sheet.Cells[i, 2] as EXCEL.Range).Text;
                    string strColors = (sheet.Cells[i, 3] as EXCEL.Range).Text;
                    string strSizes = (sheet.Cells[i, 4] as EXCEL.Range).Text;
                    string strPrice = (sheet.Cells[i, 5] as EXCEL.Range).Text;
                    string strPriceInEuro = (sheet.Cells[i, 6] as EXCEL.Range).Text;
                    string strNote = (sheet.Cells[i, 11] as EXCEL.Range).Text;

                    //strBrand = strBrand.Trim();
                    //strSeason = strSeason.Trim();
                    strArticul = strArticul.Trim();
                    //strFullName = strFullName.Trim();
                    strSiteName = strSiteName.Trim();
                    strColors = strColors.Trim();
                    strSizes = strSizes.Trim();
                    strPrice = strPrice.Trim();
                    strPriceInEuro = strPriceInEuro.Trim();
                    strNote = strNote.Trim();

                    string strSeason = cbSeason.SelectedItem.ToString();
                    string strBrand = cbBrand.SelectedItem.ToString();
                    string strFullName = strBrand + " " + strArticul;
                    strSiteName = strSiteName.ToLower() + " " + strArticul;


                    /*string[] valColors = strColors.Split(delimitersColors);
                    for (int j = 0; j < valColors.Count(); j++)
                    {
                        string color = valColors[j];
                        string[] valColors2 = color.Split(delimitersColors2);

                        string translate_color = "";

                        for (int k = 0; k < valColors2.Count(); k++)
                        {
                            string strColor = valColors2[k].Trim().ToLower();
                            if (strColor.IndexOf("combo") >= 0)
                            {
                                strColor = strColor.Replace("combo", "").Trim();
                            }
                            if (!dictColors.TryGetValue(strColor, out translate_color))
                            {
                                //MessageBox.Show(strColor);
                                dictErrorColors[strColor] = i;
                            }
                        }
                    }*/

                    string[] valColors = strColors.Split(delimitersColors);
                    string translate_colors = "";
                    for (int j = 0; j < valColors.Count(); j++)
                    {
                        string color = valColors[j];
                        if (!cbNoTranslate.Checked)
                            if (color == "цвет каталога") color = "Unique";

                        string[] valColors2 = color.Split(delimitersColors2);

                        string new_color = "";
                        string new_translate_color = "";
                        string translate_color = "";

                        for (int k = 0; k < valColors2.Count(); k++)
                        {
                            bool bCombo = false;
                            string strColor = valColors2[k].Trim().ToLower();
                            if (!cbNoTranslate.Checked)
                                if (strColor == "цвет каталога") strColor = "Unique";

                            if (strColor.IndexOf("combo") >= 0 || strColor.IndexOf("комбо") >= 0)
                            {
                                bCombo = true;
                                strColor = strColor.Replace("combo", "").Trim();
                                strColor = strColor.Replace("комбо", "").Trim();
                            }

                            if (!cbNoTranslate.Checked)
                            {
                                if (!dictColors.TryGetValue(strColor, out translate_color))
                                {
                                    //MessageBox.Show(strColor);
                                    dictErrorColors[strColor] = i;
                                    translate_color = "";
                                    //return;
                                }
                                else
                                {
                                    if (bCombo) translate_color = translate_color + " комбо";
                                }
                                if (k == 0)
                                {
                                    new_color += FirstUpper(valColors2[k]);
                                    new_translate_color += translate_color;
                                }
                                else
                                {
                                    new_color += "/" + FirstUpper(valColors2[k]);
                                    new_translate_color += "/" + translate_color;
                                }
                            }
                            else
                            {
                                if (!dictColorsNoTranslate.TryGetValue(strColor, out translate_color))
                                {
                                    //MessageBox.Show(strColor);
                                    dictErrorColors[strColor] = i;
                                    translate_color = "";
                                    //return;
                                }
                                else
                                {
                                    if (bCombo) translate_color = translate_color + " комбо";
                                }
                                if (k == 0)
                                {
                                    new_color += FirstUpper(valColors2[k]);
                                    new_translate_color += translate_color;
                                }
                                else
                                {
                                    new_color += "/" + FirstUpper(valColors2[k]);
                                    new_translate_color += "/" + translate_color;
                                }
                            }
                        }

                        if (!cbNoTranslate.Checked)
                        {
                            if (j == 0)
                                translate_colors += new_color + " (" + new_translate_color + ")";
                            else
                                translate_colors += "," + new_color + " (" + new_translate_color + ")";
                        }
                        else
                        {
                            if (j == 0)
                                translate_colors += new_translate_color;
                            else
                                translate_colors += "," + new_translate_color;
                        }
                    }

                    string[] valSizes = strSizes.Split(delimitersColors);
                    for (int j = 0; j < valSizes.Count(); j++)
                    {
                        string strSize = valSizes[j].Trim().ToLower();
                        string translate_size = "";
                        if (!dictSizes.TryGetValue(strSize, out translate_size))
                        {
                            //MessageBox.Show(strColor);
                            dictErrorSizes[strSize] = i;
                        }
                    }

                    string strType_ = "";
                    if (strSiteName.IndexOf("платье") == 0) strType_ = "платья";
                    if (strSiteName.IndexOf("блуза") == 0) strType_ = "блузки";
                    if (strSiteName.IndexOf("блузка") == 0) strType_ = "блузки";
                    if (strSiteName.IndexOf("блузон") == 0) strType_ = "блузки";
                    if (strSiteName.IndexOf("боди") == 0) strType_ = "блузки";
                    if (strSiteName.IndexOf("туника") == 0) strType_ = "туники";
                    if (strSiteName.IndexOf("юбка") == 0) strType_ = "юбки";
                    if (strSiteName.IndexOf("брюки") == 0) strType_ = "брюки и легинсы";
                    if (strSiteName.IndexOf("джинсы") == 0) strType_ = "брюки и легинсы";
                    if (strSiteName.IndexOf("болеро") == 0) strType_ = "пиджаки";
                    if (strSiteName.IndexOf("топ") == 0) strType_ = "блузки";
                    if (strSiteName.IndexOf("футболка") == 0) strType_ = "блузки";
                    if (strSiteName.IndexOf("безрукавка") == 0) strType_ = "блузки";
                    if (strSiteName.IndexOf("тренч") == 0) strType_ = "верхняя одежда";
                    if (strSiteName.IndexOf("жакет") == 0) strType_ = "жакеты и кардиганы";
                    if (strSiteName.IndexOf("пиджак") == 0) strType_ = "пиджаки";
                    if (strSiteName.IndexOf("шорты") == 0) strType_ = "шорты";
                    if (strSiteName.IndexOf("сумка") == 0) strType_ = "аксессуары";
                    if (strSiteName.IndexOf("воротник") == 0) strType_ = "аксессуары";
                    if (strSiteName.IndexOf("спортивная куртка") == 0) strType_ = "спортивная одежда";
                    if (strSiteName.IndexOf("спортивные брюки") == 0) strType_ = "спортивная одежда";
                    if (strSiteName.IndexOf("спортивный костюм") == 0) strType_ = "спортивная одежда";
                    if (strSiteName.IndexOf("пальто") == 0) strType_ = "верхняя одежда";
                    if (strSiteName.IndexOf("куртка") == 0) strType_ = "верхняя одежда";
                    if (strSiteName.IndexOf("комплект(топ+жакет)") >= 0) strType_ = "пиджаки";
                    if (strSiteName.IndexOf("комплект(топ+туника)") >= 0) strType_ = "пиджаки";
                    if (strSiteName.IndexOf("комплект") >= 0) strType_ = "пиджаки";
                    if (strSiteName.IndexOf("комбинезон") >= 0) strType_ = "пиджаки";
                    if (strSiteName.IndexOf("жакет+топ") >= 0) strType_ = "пиджаки";
                    if (strSiteName.IndexOf("кардиган+топ") >= 0) strType_ = "пиджаки";
                    if (strSiteName.IndexOf("платье+жакет") >= 0) strType_ = "пиджаки";
                    if (strSiteName.IndexOf("платье+болеро") >= 0) strType_ = "пиджаки";
                    if (strSiteName.IndexOf("пончо") == 0) strType_ = "верхняя одежда";
                    if (strSiteName.IndexOf("пуховик") == 0) strType_ = "верхняя одежда";
                    if (strSiteName.IndexOf("накидка") == 0) strType_ = "верхняя одежда";
                    if (strSiteName.IndexOf("кофта") == 0) strType_ = "пиджаки";
                    if (strSiteName.IndexOf("свитер") == 0) strType_ = "пиджаки";
                    if (strSiteName.IndexOf("капри") >= 0) strType_ = "брюки и легинсы";
                    if (strSiteName.IndexOf("шарф") == 0) strType_ = "аксессуары";
                    if (strSiteName.IndexOf("платок") == 0) strType_ = "аксессуары";
                    if (strSiteName.IndexOf("шапочка вязаная") == 0) strType_ = "аксессуары";
                    if (strSiteName.IndexOf("шапка") == 0) strType_ = "аксессуары";
                    if (strSiteName.IndexOf("меховой воротник") == 0) strType_ = "аксессуары";
                    if (strSiteName.IndexOf("легинсы") == 0) strType_ = "брюки и легинсы";
                    if (strSiteName.IndexOf("берет") == 0) strType_ = "аксессуары";
                    if (strSiteName.IndexOf("палантин") == 0) strType_ = "аксессуары";
                    if (strSiteName.IndexOf("кардиган") == 0) strType_ = "жакеты и кардиганы";
                    if (strSiteName.IndexOf("жилет") == 0) strType_ = "пиджаки";
                    if (strSiteName.IndexOf("пуловер") == 0) strType_ = "блузки";
                    if (strSiteName.IndexOf("батник") == 0) strType_ = "блузки";
                    if (strSiteName.IndexOf("рубашка") == 0) strType_ = "блузки";
                    if (strSiteName.IndexOf("водолазка") == 0) strType_ = "блузки";
                    if (strSiteName.IndexOf("сорочка") == 0) strType_ = "блузки";
                    if (strSiteName.IndexOf("плащ") == 0) strType_ = "верхняя одежда";
                    if (strSiteName.IndexOf("лосины") == 0) strType_ = "брюки и легинсы";
                    if (strSiteName.IndexOf("двойка") == 0) strType_ = "пиджаки";

                    if (strSiteName.IndexOf("трикотажная блуза") == 0) strType_ = "блузки";
                    if (strSiteName.IndexOf("трикотажная туника") == 0) strType_ = "туники";
                    if (strSiteName.IndexOf("трикотажная юбка") == 0) strType_ = "юбки";
                    if (strSiteName.IndexOf("трикотажный топ") == 0) strType_ = "блузки";
                    if (strSiteName.IndexOf("трикотажный джемпер") == 0) strType_ = "пиджаки";

                    if (strSiteName.IndexOf("гетры") == 0) strType_ = "аксессуары";
                    if (strSiteName.IndexOf("пояс") == 0) strType_ = "аксессуары";
                    if (strSiteName.IndexOf("перчатки") >= 0) strType_ = "аксессуары";

                    if (strType_ == "")
                        MessageBox.Show(i.ToString() + ": " + strSiteName);

                    /*if (strSeason.ToLower().IndexOf("winter") != 0 && strSeason.ToLower().IndexOf("autumn") != 0 &&
                        strSeason.ToLower().IndexOf("summer") != 0 && strSeason.ToLower().IndexOf("spring") != 0)
                    {
                        MessageBox.Show(i.ToString() + ": " + strSeason);
                    }

                    strSeason = strSeason.ToLower().Replace("winter", "зима ");
                    strSeason = strSeason.ToLower().Replace("autumn", "осень ");
                    strSeason = strSeason.ToLower().Replace("spring", "весна ");
                    strSeason = strSeason.ToLower().Replace("summer", "лето ");

                    if (strSeason != cbSeason.SelectedItem.ToString())
                    {
                        MessageBox.Show(i.ToString() + ": " + strSeason);
                    }

                    if (strBrand != cbBrand.SelectedItem.ToString())
                    {
                        MessageBox.Show(i.ToString() + ": " + strBrand);
                    }*/

                    string strImagePath1 = "";
                    string strImagePath2 = "";
                    string strImagePath3 = "";
                    string strImagePath4 = "";
                    string strImagePath5 = "";

                    //KeyValuePair<int, string> kvp = new KeyValuePair<int, string>();
                    if (!dictPictures.ContainsKey(strArticul.ToLower()))
                    {
                        dictErrorPictures[strArticul] = i;
                    }
                    else
                    {
                        foreach (KeyValuePair<int, List<string>> pr in dictPictures[strArticul.ToLower()])
                        {
                            for (int ii = 0; ii < pr.Value.Count; ii++)
                            {
                            	if (pr.Key == 1) strImagePath1 = pr.Value[ii];
                                if (pr.Key == 2) strImagePath2 = pr.Value[ii];
                                if (pr.Key == 3) strImagePath3 = pr.Value[ii];
                                if (pr.Key == 4) strImagePath4 = pr.Value[ii];
                                if (pr.Key == 5) strImagePath5 = pr.Value[ii];
                            }
                        }

                        while (true)
                        {
                            bool bStop = true;
                            if (strImagePath4 == "" && strImagePath5 != "")
                            {
                                strImagePath4 = strImagePath5;
                                strImagePath5 = "";
                                bStop = false;
                            }
                            if (strImagePath3 == "" && strImagePath4 != "")
                            {
                                strImagePath3 = strImagePath4;
                                strImagePath4 = "";
                                bStop = false;
                            }
                            if (strImagePath2 == "" && strImagePath3 != "")
                            {
                                strImagePath2 = strImagePath3;
                                strImagePath3 = "";
                                bStop = false;
                            }
                            if (strImagePath1 == "" && strImagePath2 != "")
                            {
                                strImagePath1 = strImagePath2;
                                strImagePath2 = "";
                                bStop = false;
                            }
                            if (bStop) break;
                        }


                        /*for (int ii = 0; ii < dictPictures[strArticul.ToLower()].Count(); ii++)
                        {
                            //dictPictures[strArticul.ToLower()][i] = new KeyValuePair<int, string>(1, kvp.Value);
                            if (ii == 0) strImagePath1 = dictPictures[strArticul.ToLower()][ii].Value;
                            if (ii == 1) strImagePath2 = dictPictures[strArticul.ToLower()][ii].Value;
                            if (ii == 2) strImagePath3 = dictPictures[strArticul.ToLower()][ii].Value;
                            if (ii == 3) strImagePath4 = dictPictures[strArticul.ToLower()][ii].Value;
                            if (ii == 4) strImagePath5 = dictPictures[strArticul.ToLower()][ii].Value;
                        }*/
                    }

                    if (dictErrorDublicates.ContainsKey(strArticul.ToLower()))
                    {
                        dictErrorDublicates[strArticul.ToLower()] = new KeyValuePair<int, int>(dictErrorDublicates[strArticul.ToLower()].Key, i);
                    }
                    else
                    {
                        dictErrorDublicates[strArticul.ToLower()] = new KeyValuePair<int, int>(i, 0);
                    }

                    /*if (strFullName.IndexOf(strArticul) < 0)
                    {
                        MessageBox.Show(i.ToString() + ": " + strFullName + " <> " + strArticul);
                    }

                    if (strSiteName.IndexOf(strArticul) < 0)
                    {
                        MessageBox.Show(i.ToString() + ": " + strSiteName + " <> " + strArticul);
                    }*/

                    //strBrand
                    //strSeason
                    //strFullName
                    //strSiteName
                    //strArticul
                    //strType_
                    //translate_colors
                    //strSizes
                    //strPrice
                    //strPriceInEuro
                    //strImagePath

                    string strForWrite = strBrand + "\t" + strSeason + "\t" + strFullName + "\t" + strSiteName + "\t"
                        + strArticul + "\t" + strType_ + "\t" + translate_colors + "\t" + strSizes + "\t" + strPrice + "\t"
                        + strPriceInEuro + "\t" + strImagePath1 + "\t" + strImagePath2 + "\t" + strImagePath3 + "\t" + strImagePath4 + "\t" + strImagePath5 + "\t" + strNote;
                    fileMain.WriteLine(strForWrite);
                }

                fileMain.Close();

                wb.Close();
                app.Quit();

                wb = null;
                app = null;
                sheet = null;
                GC.Collect();

                //if (dictErrorColors.Count > 0)
                {
                    StreamWriter file3 = new StreamWriter(Path.GetDirectoryName(tbTextFile.Text) + "\\not_found_color.txt", false, Encoding.GetEncoding(1251));
                    foreach (KeyValuePair<string, int> val in dictErrorColors)
                    {
                        string strForWrite = val.Key + "\t" + val.Value;
                        file3.WriteLine(strForWrite);
                    }
                    file3.Close();
                }

                //if (dictErrorSizes.Count > 0)
                {
                    StreamWriter file3 = new StreamWriter(Path.GetDirectoryName(tbTextFile.Text) + "\\not_found_size.txt", false, Encoding.GetEncoding(1251));
                    foreach (KeyValuePair<string, int> val in dictErrorSizes)
                    {
                        string strForWrite = val.Key + "\t" + val.Value;
                        file3.WriteLine(strForWrite);
                    }
                    file3.Close();
                }

                //if (dictErrorPictures.Count > 0)
                {
                    StreamWriter file3 = new StreamWriter(Path.GetDirectoryName(tbTextFile.Text) + "\\not_found_pic.txt", false, Encoding.GetEncoding(1251));
                    foreach (KeyValuePair<string, int> val in dictErrorPictures)
                    {
                        string strForWrite = val.Key + "\t" + val.Value;
                        file3.WriteLine(strForWrite);
                    }
                    file3.Close();
                }

                {
                    StreamWriter file4 = new StreamWriter(Path.GetDirectoryName(tbTextFile.Text) + "\\dublicate.txt", false, Encoding.GetEncoding(1251));
                    foreach (KeyValuePair<string, KeyValuePair<int, int>> val in dictErrorDublicates)
                    {
                        if (val.Value.Value != 0)
                        {
                            string strForWrite = val.Key + "\t" + val.Value.Key + "\t" + val.Value.Value;
                            file4.WriteLine(strForWrite);
                        }
                    }
                    file4.Close();
                }

                /*StreamWriter file4 = new StreamWriter(Path.GetDirectoryName(tbTextFile.Text) + "\\not_found_pic2.txt", false, Encoding.GetEncoding(1251));
                foreach (KeyValuePair<string, KeyValuePair<int, string>> val in dictPictures)
                {
                    if (val.Value.Key == 0)
                    {
                        string strForWrite = val.Value.Value;
                        file4.WriteLine(strForWrite);
                    }
                }
                file4.Close();*/

                MessageBox.Show("Файл создан.");
            }
            else
            {
                EXCEL.Application app = new EXCEL.Application();
                EXCEL.Workbook wb = app.Workbooks.Open(tbTextFile.Text);
                EXCEL.Worksheet sheet = wb.Sheets[1];

                app.DisplayAlerts = false;

                int columnsCount = sheet.UsedRange.Columns.Count;
                int rowsCount = sheet.UsedRange.Rows.Count;

                dictTovars.Clear();

                char[] delimiters = { ';' };
                char[] delimitersColors = { ',' };
                char[] delimitersColors2 = { '/' };

                StreamWriter fileMain = new StreamWriter(Path.GetDirectoryName(tbTextFile.Text) + "\\price.txt", false, Encoding.GetEncoding(1251));

                for (int i = 2; i <= rowsCount; i++)
                {
                    //string strType = (sheet.Cells[i, 5] as EXCEL.Range).Text;
                    string strBrand = (sheet.Cells[i, 1] as EXCEL.Range).Text;
                    string strSeason = (sheet.Cells[i, 2] as EXCEL.Range).Text;
                    string strArticul = (sheet.Cells[i, 3] as EXCEL.Range).Text;
                    string strPrice = (sheet.Cells[i, 4] as EXCEL.Range).Text;

                    strBrand = strBrand.Trim();
                    strSeason = strSeason.Trim();
                    strArticul = strArticul.Trim();
                    strPrice = strPrice.Trim();

                    if (strSeason.ToLower().IndexOf("winter") != 0 && strSeason.ToLower().IndexOf("autumn") != 0 &&
                        strSeason.ToLower().IndexOf("summer") != 0 && strSeason.ToLower().IndexOf("spring") != 0)
                    {
                        MessageBox.Show(strSeason);
                    }

                    strSeason = strSeason.ToLower().Replace("winter", "зима ");
                    strSeason = strSeason.ToLower().Replace("autumn", "осень ");
                    strSeason = strSeason.ToLower().Replace("spring", "весна ");
                    strSeason = strSeason.ToLower().Replace("summer", "лето ");

                    if (dictErrorDublicates.ContainsKey(strArticul.ToLower()))
                    {
                        dictErrorDublicates[strArticul.ToLower()] = new KeyValuePair<int, int>(dictErrorDublicates[strArticul.ToLower()].Key, i);
                    }
                    else
                    {
                        dictErrorDublicates[strArticul.ToLower()] = new KeyValuePair<int, int>(i, 0);
                    }

                    //strBrand
                    //strSeason
                    //strArticul
                    //strPrice

                    string strForWrite = strBrand + "\t" + strSeason + "\t" + strArticul + "\t" + strPrice;
                    fileMain.WriteLine(strForWrite);
                }

                fileMain.Close();

                wb.Close();
                app.Quit();

                wb = null;
                app = null;
                sheet = null;
                GC.Collect();

                {
                    StreamWriter file4 = new StreamWriter(Path.GetDirectoryName(tbTextFile.Text) + "\\dublicate.txt", false, Encoding.GetEncoding(1251));
                    foreach (KeyValuePair<string, KeyValuePair<int, int>> val in dictErrorDublicates)
                    {
                        if (val.Value.Value != 0)
                        {
                            string strForWrite = val.Key + "\t" + val.Value.Key + "\t" + val.Value.Value;
                            file4.WriteLine(strForWrite);
                        }
                    }
                    file4.Close();
                }

                /*StreamWriter file4 = new StreamWriter(Path.GetDirectoryName(tbTextFile.Text) + "\\not_found_pic2.txt", false, Encoding.GetEncoding(1251));
                foreach (KeyValuePair<string, KeyValuePair<int, string>> val in dictPictures)
                {
                    if (val.Value.Key == 0)
                    {
                        string strForWrite = val.Value.Value;
                        file4.WriteLine(strForWrite);
                    }
                }
                file4.Close();*/

                MessageBox.Show("Файл создан.");
            }
        }

        private void cbOnlyPrices_CheckedChanged(object sender, EventArgs e)
        {

        }
    }
}
