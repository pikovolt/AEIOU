using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;
using System.Windows.Forms;
using System.IO;
using System.Xml;

namespace AEIOU
{
    class Settings
    {
        public Settings()
        {
        }

        // 設定用 各種変数
        public FormStartPosition StartPosition = FormStartPosition.Manual;  //フォーム初期位置設定モード
        public Point Location = new Point(0,0);     //フォーム座標
        public Size Size = new Size(300,600);       //フォームサイズ
        public bool TopMost = true;                 //フォームのStayOnTop設定
        public int ColLength = 4;                   //列数
        public int RowLength = 1800;                //行数
        public int CellCountLimit = 100;            //編集可能セル枚数のリミット
        //public int EtcLength = 1;                   //その他の情報数: 色情報[1]
        //public int ColMaxLength = 26;               //最大列数
        //public int RowMaxLength = 9999;             //最大行数
        public int SheetSec = 6;                    //シートの秒数
        public int Fps = 24;                        //フレームレート
        public int FirstFrame = 1;                  //開始コマ数
        public int SheetDivide = 6;                 //24fps時の分割コマ数[4,6,12]
        public String KaraCell = "100";             //カラセル番号
        public int UndoCount = 0;                   //アンドゥ整理番号（１操作毎に＋１）
        public int UndoLimitCount = 65535;          //アンドゥ履歴登録個数（アンドゥ回数ではなく UndoOperation登録個数）
        public String CurrentDir;                   //カレントパス
        public bool IsAutoadjust = false;           //自動サイズ調整
        public bool IsKaraNoMove = false;           //空セル入力時の移動動作
        public bool IsDisplayFrameNumber = false;   //空セル入力時の移動動作
        public bool IsAlwaysAppend = false;         //入力動作：常時追加入力にするか否か
        public String AfterPath = "C:\\Program Files\\Adobe\\Adobe After Effects CS5\\Support Files\\AfterFX.exe";
                                                    //afterfx.exe パス
        public String AfterOption = "-r C:\\Program Files\\Adobe\\Adobe After Effects CS5\\Support Files\\Scripts\\setRemap.jsx";
                                                    //afterfx オプション
        public RemapOutputMode AfterOutputMode = RemapOutputMode.AERemap;
                                                    //afterfx 出力モード
        public String AfterRemapVersion = "9.0";    //コピー文字列に含まれるバージョン番号

        // ショートカット用定義
        public KeyBinds keys = new KeyBinds();


        // ノードの文字列要素を取得
        private String getInnerText(String name, XmlNodeList node)
        {
            String val = "";
            if (node.Count == 0)
            {
                MessageBox.Show("ノード:" + name + " を取得出来ませんでした.");
            }
            else
            {
                try
                {
                    val = node[0].ChildNodes[0].InnerText;
                }
                catch (Exception e)
                {
                    MessageBox.Show("ノード:" + node[0].Name + "から文字列を取り出すことが出来ませんでした.");
                }
            }
            return val;
        }

        // 初期設定ファイルの読み込み
        public bool loadSettingFile(String path)
        {
            XmlDocument doc = new XmlDocument();
            try
            {
                doc.Load(path);
            }
            catch (Exception e)
            {
                MessageBox.Show("初期設定ファイルを読み込めませんでした.");
                return false;
            }

            // formatType&Version
            string fileType = getInnerText("fileFormatType",doc.SelectNodes("/document/head/fileFormatType"));
            string fileVersion = getInnerText("fileFormatVersion", doc.SelectNodes("/document/head/fileFormatVersion"));
            if (fileType != "AEIOU Preference Format") return false;
            if (fileVersion != "1.0") return false;

            // window pos, size, stayOnTop, autoAdjust
            XmlNodeList windowPos = doc.SelectNodes("/document/pref/window/position");
            XmlNodeList windowSize = doc.SelectNodes("/document/pref/window/size");
            string windowStayOnTop = getInnerText("stayOnTop", doc.SelectNodes("/document/pref/window/stayOnTop"));
            string windowAutoadjust = getInnerText("autoAdjust", doc.SelectNodes("/document/pref/window/autoAdjust"));

            int x, y, w, h;
            try
            {
                x = int.Parse(windowPos[0].Attributes["x"].Value);
                y = int.Parse(windowPos[0].Attributes["y"].Value);
                w = int.Parse(windowSize[0].Attributes["w"].Value);
                h = int.Parse(windowSize[0].Attributes["h"].Value);
                // ウィンドウ位置の初期設定
                this.StartPosition = FormStartPosition.Manual;
                this.Location = new Point(x, y);
                this.Size = new Size(w, h);
            }
            catch (Exception e)
            {
                MessageBox.Show("windows position,size に指定される文字列を数値に変換できませんでした.");
                // 初期設定ファイルがない場合のウィンドウ位置の初期設定
                this.StartPosition = FormStartPosition.Manual;
                this.Location = new Point(0, 0);
            }
            if (windowStayOnTop == "1")
                this.TopMost = true;
            else
                this.TopMost = false;
            if (windowAutoadjust == "1")
                IsAutoadjust = true;
            else
                IsAutoadjust = false;

            // setting: fps, sheetSec, firstFrame, sheetDivide
            string fps = getInnerText("fps", doc.SelectNodes("/document/pref/setting/fps"));
            string sheetSec = getInnerText("sheetSec", doc.SelectNodes("/document/pref/setting/sheetSec"));
            string firstFrm = getInnerText("firstFrame", doc.SelectNodes("/document/pref/setting/firstFrame"));
            string sheetDiv = getInnerText("sheetDivide", doc.SelectNodes("/document/pref/setting/sheetDivide"));
            try
            {
                Fps = int.Parse(fps);
            }
            catch (Exception e)
            {
                MessageBox.Show("fps に指定される文字列を数値に変換できませんでした.");
            }
            try
            {
                SheetSec = int.Parse(sheetSec);
            }
            catch (Exception e)
            {
                MessageBox.Show("sheetSec に指定される文字列を数値に変換できませんでした.");
            }
            try
            {
                FirstFrame = int.Parse(firstFrm);
            }
            catch (Exception e)
            {
                MessageBox.Show("firstFrame に指定される文字列を数値に変換できませんでした.");
            }
            try
            {
                SheetDivide = int.Parse(sheetDiv);
            }
            catch (Exception e)
            {
                MessageBox.Show("sheetDivide に指定される文字列を数値に変換できませんでした.");
            }

            // setting: colLength, rowLength, cellCountLimit
            string colLen = getInnerText("colLength", doc.SelectNodes("/document/pref/setting/colLength"));
            string rowLen = getInnerText("rowLength", doc.SelectNodes("/document/pref/setting/rowLength"));
            string cellLimit = getInnerText("cellCountLimit", doc.SelectNodes("/document/pref/setting/cellCountLimit"));
            try
            {
                ColLength = int.Parse(colLen);
            }
            catch (Exception e)
            {
                MessageBox.Show("colLength に指定される文字列を数値に変換できませんでした.");
            }
            try
            {
                RowLength = int.Parse(rowLen);
            }
            catch (Exception e)
            {
                MessageBox.Show("rowLength に指定される文字列を数値に変換できませんでした.");
            }
            try
            {
                CellCountLimit = int.Parse(cellLimit);
            }
            catch (Exception e)
            {
                MessageBox.Show("cellCountLimit に指定される文字列を数値に変換できませんでした.");
            }

            // setting: karaCell(value, noMove)
            string karaValue = getInnerText("value", doc.SelectNodes("/document/pref/setting/karaCell/value"));
            string karaMove = getInnerText("noMove", doc.SelectNodes("/document/pref/setting/karaCell/noMove"));
            KaraCell = karaValue;
            if (karaMove == "1")
                IsKaraNoMove = true;
            else
                IsKaraNoMove = false;

            // setting: undoLimit, IsDisplayFrameNumber
            string undoLimit = getInnerText("undoLimit", doc.SelectNodes("/document/pref/setting/undoLimit"));
            string isDisplayFrameNumber = getInnerText("isDisplayFrameNum", doc.SelectNodes("/document/pref/setting/isDisplayFrameNum"));
            try
            {
                UndoLimitCount = int.Parse(undoLimit);
            }
            catch (Exception e)
            {
                MessageBox.Show("undoLimit に指定される文字列を数値に変換できませんでした.");
            }
            if (isDisplayFrameNumber == "1")
                IsDisplayFrameNumber = true;
            else
                IsDisplayFrameNumber = false;

            // setting: AfterFX(path, option, outputMode)
            string Path = getInnerText("path", doc.SelectNodes("/document/pref/setting/afterFX/path"));
            string Option = getInnerText("option", doc.SelectNodes("/document/pref/setting/afterFX/option"));
            string Mode = getInnerText("outputMode", doc.SelectNodes("/document/pref/setting/afterFX/outputMode"));
            string RemapVersion = getInnerText("remapVersion", doc.SelectNodes("/document/pref/setting/afterFX/remapVersion"));
            AfterPath = Path;
            AfterOption = Option;
            AfterRemapVersion = RemapVersion;
            try
            {
                switch(Mode)
                {
                    case "None":
                        AfterOutputMode = RemapOutputMode.None;
                        break;
                    case "AERemap":
                        AfterOutputMode = RemapOutputMode.AERemap;
                        break;
                    case "DirectRemap":
                        AfterOutputMode = RemapOutputMode.DirectRemap;
                        break;
                    case "JSDirectRemap":
                        AfterOutputMode = RemapOutputMode.JSDirectRemap;
                        break;
                    case "JSRemap":
                        AfterOutputMode = RemapOutputMode.JSRemap;
                        break;
                }
            }
            catch (Exception e)
            {
                MessageBox.Show("afterFX/outputMode に指定される文字列を数値に変換できませんでした.");
            }

            // setting: IsAlwaysAppend
            string isAlwaysAppend = getInnerText("isAlwaysAppend", doc.SelectNodes("/document/pref/setting/isAlwaysAppend"));
            if (isAlwaysAppend == "1")
                IsAlwaysAppend = true;
            else
                IsAlwaysAppend = false;

            
            //ショートカットキー定義の読み込み
            keys.loadShortcutsSetting(doc);

            //キーバインド定義の読み込み
            keys.loadKeybindsSetting(doc);

            return true;
        }

        // 初期設定ファイルの書き出し
        public bool saveSettingFile(String path)
        {
            XmlDocument doc = new XmlDocument();
            XmlDeclaration declaration = doc.CreateXmlDeclaration("1.0", "UTF-8", null);
            XmlElement root = doc.CreateElement("document");
            doc.AppendChild(declaration);
            doc.AppendChild(root);

            {
                XmlElement head = doc.CreateElement("head");
                XmlElement type = doc.CreateElement("fileFormatType");
                XmlElement ver = doc.CreateElement("fileFormatVersion");
                root.AppendChild(head);
                head.AppendChild(type);
                head.AppendChild(ver);
                type.InnerText = "AEIOU Preference Format";
                ver.InnerText = "1.0";
            }

            {
                XmlElement pref = doc.CreateElement("pref");
                XmlElement win = doc.CreateElement("window");
                XmlElement set = doc.CreateElement("setting");
                XmlElement shortcuts = doc.CreateElement("shortcuts");
                XmlElement binds = doc.CreateElement("keybinds");
                root.AppendChild(pref);
                pref.AppendChild(win);
                pref.AppendChild(set);
                pref.AppendChild(shortcuts);
                pref.AppendChild(binds);
                {
                    XmlElement pos = doc.CreateElement("position");
                    XmlElement siz = doc.CreateElement("size");
                    XmlElement top = doc.CreateElement("stayOnTop");
                    XmlElement auto = doc.CreateElement("autoAdjust");
                    win.AppendChild(pos);
                    win.AppendChild(siz);
                    win.AppendChild(top);
                    win.AppendChild(auto);
                    pos.SetAttribute("x", Location.X.ToString());
                    pos.SetAttribute("y", Location.Y.ToString());
                    siz.SetAttribute("w", Size.Width.ToString());
                    siz.SetAttribute("h", Size.Height.ToString());
                    top.InnerText = (TopMost) ? "1" : "0";
                    auto.InnerText = (IsAutoadjust) ? "1" : "0";
                }
                {
                    XmlElement fps = doc.CreateElement("fps");
                    XmlElement sec = doc.CreateElement("sheetSec");
                    XmlElement frm = doc.CreateElement("firstFrame");
                    XmlElement col = doc.CreateElement("colLength");
                    XmlElement row = doc.CreateElement("rowLength");
                    XmlElement cellLimit = doc.CreateElement("cellCountLimit");
                    XmlElement div = doc.CreateElement("sheetDivide");
                    XmlElement kara = doc.CreateElement("karaCell");
                    set.AppendChild(fps);
                    set.AppendChild(sec);
                    set.AppendChild(frm);
                    set.AppendChild(col);
                    set.AppendChild(cellLimit);
                    set.AppendChild(row);
                    set.AppendChild(div);
                    set.AppendChild(kara);
                    fps.InnerText = Fps.ToString();
                    sec.InnerText = SheetSec.ToString();
                    frm.InnerText = FirstFrame.ToString();
                    col.InnerText = ColLength.ToString();
                    row.InnerText = RowLength.ToString();
                    cellLimit.InnerText = CellCountLimit.ToString();
                    div.InnerText = SheetDivide.ToString();
                    {
                        XmlElement value = doc.CreateElement("value");
                        XmlElement move = doc.CreateElement("noMove");
                        kara.AppendChild(value);
                        kara.AppendChild(move);
                        value.InnerText = KaraCell.ToString();
                        move.InnerText = (IsKaraNoMove) ? "1" : "0";
                    }
                    XmlElement limit = doc.CreateElement("undoLimit");
                    XmlElement isDisplayFrameNumber = doc.CreateElement("isDisplayFrameNum");
                    XmlElement after = doc.CreateElement("afterFX");
                    set.AppendChild(limit);
                    set.AppendChild(isDisplayFrameNumber);
                    set.AppendChild(after);
                    limit.InnerText = UndoLimitCount.ToString();
                    isDisplayFrameNumber.InnerText = (IsDisplayFrameNumber) ? "1" : "0";
                    {
                        XmlElement afterPath = doc.CreateElement("path");
                        XmlElement option = doc.CreateElement("option");
                        XmlElement mode = doc.CreateElement("outputMode");
                        XmlElement remapVersion = doc.CreateElement("remapVersion");
                        after.AppendChild(afterPath);
                        after.AppendChild(option);
                        after.AppendChild(mode);
                        after.AppendChild(remapVersion);
                        afterPath.InnerText = AfterPath;
                        option.InnerText = AfterOption;
                        mode.InnerText = AfterOutputMode.ToString();
                        remapVersion.InnerText = AfterRemapVersion;
                    }
                    XmlElement isAlwaysAppend = doc.CreateElement("isAlwaysAppend");
                    set.AppendChild(isAlwaysAppend);
                    isAlwaysAppend.InnerText = (IsAlwaysAppend) ? "1" : "0";
                }
            }

            // ショートカット定義の書き出し
            keys.saveShortcutsSetting(doc);

            //キーバインド定義の書き出し
            keys.saveKeybindsSetting(doc);

            // XML定義の書き出し
            doc.Save(path);

            return true;
        }


    }
}
