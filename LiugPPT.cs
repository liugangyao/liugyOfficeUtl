using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using PPT = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using System.Diagnostics;
using LiugyCommon;

namespace liugyOfficeUtl
{

    public class LiugyPPT
    {

        /// <summary>
        /// Read the all pptx file to txt from one fold and sub fold
        /// </summary>
        /// <param name="_rootPath">the root of folder</param>
        /// <param name="_outputPath">the fold to output</param>
        /// <returns>void</returns>
        public static void readPPT2txt(string _rootPath,string _outputPath)
        {
            IEnumerable<string> files =  
                Directory.GetFiles(_rootPath, "*.pptx", SearchOption.AllDirectories);

            foreach (string f in files)
            {
                List<string> notes = new List<string>();

                Debug.WriteLine(f.ToString() + " Start!!!!");

                notes = PPT2txt(f);

                string fn = Path.GetFileNameWithoutExtension(f) + ".txt";
                FileHelper.writeFile(notes, _outputPath + "\\text\\"+fn);

                Debug.WriteLine(f.ToString() + " End!!!!");

            }
        }

        /// <summary>
        /// Read the file of PPT and output to List of string
        /// </summary>
        /// <param name="file">full path of ppt</param>
        /// <returns>list of string</returns>
        public static List<string> PPT2txt(string file)
        {

            PPT.Application app = null;
            PPT.Presentation pptx = null;
            List<string> notes = new List<string>();
            try
            {
                // PPTのインスタンス作成
                app = new PPT.Application();

                // ファイルオープン
                pptx = app.Presentations.Open(
                    file,
                    Office.MsoTriState.msoTrue,    // 読み取り専用
                    Office.MsoTriState.msoTrue,
                    Office.MsoTriState.msoFalse
               );

                foreach (PPT.Slide sld in pptx.Slides)
                { 
                    foreach (PPT.Shape shp in sld.Shapes)
                {
                    if (shp.HasTextFrame == Office.MsoTriState.msoTrue)
                    {
                        if (shp.TextFrame.HasText == Office.MsoTriState.msoTrue)
                        {
                            notes.Add(shp.TextFrame.TextRange.Text);
                        }
                    }

                    else if (shp.HasTable == Office.MsoTriState.msoTrue)
                    {
                        foreach (PPT.Row rw in shp.Table.Rows)
                        {
                            string strRW = "";
                            foreach (PPT.Cell cl in rw.Cells)
                            {
                                strRW = strRW + cl.Shape.TextFrame.TextRange.Text == "" ? "△" : cl.Shape.TextFrame.TextRange.Text;
                            }
                            notes.Add(strRW);
                        }
                    }

                    else if (shp.HasChart == Office.MsoTriState.msoTrue)
                    {
                        if (shp.Chart.HasTitle)
                        {
                                notes.Add(shp.Chart.Title);
                        }
                    }

                    else if (shp.HasSmartArt == Office.MsoTriState.msoTrue)
                    {
                        foreach (Office.SmartArtNode art in shp.SmartArt.Nodes)
                        {
                                notes.Add(art.TextFrame2.TextRange.Text);
                        }
                    }
                }

                }


            }
            finally
            {
                // ファイルを閉じる
                if (pptx != null)
                {
                    pptx.Close();
                    pptx = null;
                }

                // PPTを閉じる
                if (app != null)
                {
                    app.Quit();
                    app = null;
                }
            }

            return notes;
        }

    }
}
