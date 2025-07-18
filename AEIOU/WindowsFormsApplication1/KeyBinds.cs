using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Xml;

namespace AEIOU
{
    class KeyBinds
    {
        public KeyBinds()
        {
        }
        
        // 各種ショートカットキー
        public String Undo;
        public String AECopy;
        public String AECopy2;
        public String AECopy3;
        public String AEPaste;
        public String Nuki;
        public String NukiCancel;
        public String Hari;
        public String HariCancel;
        public String Copy;
        public String Cut;
        public String Paste;
        public String Set30FPS;             //FPS変更: 30
        public String Set24FPS;             //FPS変更: 24
        public String InputCellCount;       //セル枚数の指定
        public String InputFirstFrame;      //開始フレームの指定
        public String InputKaraCellValue;   //カラセル文字列の指定
        public String KaraCellNoMove;       //カラセル入力時の移動抑止
        public String InputSheetSec;        //シートの秒数の指定
        public String SetSheetDiv4;         //基準線:4
        public String SetSheetDiv6;         //基準線:6
        public String SetSheetDiv12;        //基準線:12
        public String StayOnTop;            //常に一番上に表示
        public String AutoAdjust;           //セル変更時にウィンドウ調整する
        public String DisplayFrameNumber;   //フレーム数表示に
        public String AlwaysAppend;         //セル入力を"常に追加"に
        public String AfterFXPath;          //AfterFX.exeパスの指定
        public String AfterFXOption;        //スクリプト jsxパスの指定
        public String AllInitialize;        //クリア
        public String AuxInputRepeat;                   //(入力補助)繰り返し
        public String AuxInputSequentialNumber;         //(入力補助)連番
        public String AuxInputReplace;                  //(入力補助)置換
        public String AuxInputReverse;                  //(入力補助)反転
        public String AuxInputDuplicate;                //(入力補助)複製
        public String AuxInputFourArithmeticOperation;  //(入力補助)四則演算

        // キーコード変換設定
        public Dictionary<int, int> Keymap = new Dictionary<int, int>();
        public Dictionary<int, int> KeymapRev = new Dictionary<int, int>();     //逆引き用


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

        // 初期設定ファイル(Shortcuts)の読み込み
        public bool loadShortcutsSetting(XmlDocument doc)
        {
            Undo = getInnerText("undo",doc.SelectNodes("/document/pref/shortcuts/undo"));
            AECopy = getInnerText("remapToAE",doc.SelectNodes("/document/pref/shortcuts/remapToAE"));
            AECopy2 = getInnerText("remapToAE2", doc.SelectNodes("/document/pref/shortcuts/remapToAE2"));
            AECopy3 = getInnerText("remapToAE3", doc.SelectNodes("/document/pref/shortcuts/remapToAE3"));
            AEPaste = getInnerText("remapFromAE", doc.SelectNodes("/document/pref/shortcuts/remapFromAE"));
            Nuki = getInnerText("komanuki", doc.SelectNodes("/document/pref/shortcuts/komanuki"));
            NukiCancel = getInnerText("komanukiCancel", doc.SelectNodes("/document/pref/shortcuts/komanukiCancel"));
            Hari = getInnerText("harikomi", doc.SelectNodes("/document/pref/shortcuts/harikomi"));
            HariCancel = getInnerText("harikomiCancel", doc.SelectNodes("/document/pref/shortcuts/harikomiCancel"));
            Copy = getInnerText("copy", doc.SelectNodes("/document/pref/shortcuts/copy"));
            Cut = getInnerText("cut", doc.SelectNodes("/document/pref/shortcuts/cut"));
            Paste = getInnerText("paste", doc.SelectNodes("/document/pref/shortcuts/paste"));
            Set30FPS = getInnerText("set30fps", doc.SelectNodes("/document/pref/shortcuts/set30fps"));
            Set24FPS = getInnerText("set24fps", doc.SelectNodes("/document/pref/shortcuts/set24fps"));
            InputCellCount = getInnerText("inputCellCount", doc.SelectNodes("/document/pref/shortcuts/inputCellCount"));
            InputFirstFrame = getInnerText("inputFirstFrame", doc.SelectNodes("/document/pref/shortcuts/inputFirstFrame"));
            InputKaraCellValue = getInnerText("inputKaraCellValue", doc.SelectNodes("/document/pref/shortcuts/inputKaraCellValue"));
            KaraCellNoMove = getInnerText("karaCellNoMove", doc.SelectNodes("/document/pref/shortcuts/karaCellNoMove"));
            InputSheetSec = getInnerText("inputSheetSec", doc.SelectNodes("/document/pref/shortcuts/inputSheetSec"));
            SetSheetDiv4 = getInnerText("setSheetDiv4", doc.SelectNodes("/document/pref/shortcuts/setSheetDiv4"));
            SetSheetDiv6 = getInnerText("setSheetDiv6", doc.SelectNodes("/document/pref/shortcuts/setSheetDiv6"));
            SetSheetDiv12 = getInnerText("setSheetDiv12", doc.SelectNodes("/document/pref/shortcuts/setSheetDiv12"));
            StayOnTop = getInnerText("stayOnTop", doc.SelectNodes("/document/pref/shortcuts/stayOnTop"));
            AutoAdjust = getInnerText("autoAdjust", doc.SelectNodes("/document/pref/shortcuts/autoAdjust"));
            DisplayFrameNumber = getInnerText("displayFrameNumber", doc.SelectNodes("/document/pref/shortcuts/displayFrameNumber"));
            AlwaysAppend = getInnerText("alwaysAppend", doc.SelectNodes("/document/pref/shortcuts/alwaysAppend"));
            AfterFXPath = getInnerText("afterFxPath", doc.SelectNodes("/document/pref/shortcuts/afterFxPath"));
            AfterFXOption = getInnerText("afterFxOption", doc.SelectNodes("/document/pref/shortcuts/afterFxOption"));
            AllInitialize = getInnerText("allInitialize", doc.SelectNodes("/document/pref/shortcuts/allInitialize"));
            AuxInputRepeat = getInnerText("auxInputRepeat", doc.SelectNodes("/document/pref/shortcuts/auxInputRepeat"));
            AuxInputSequentialNumber = getInnerText("auxInputSequentialNumber", doc.SelectNodes("/document/pref/shortcuts/auxInputSequentialNumber"));
            AuxInputReplace = getInnerText("auxInputReplace", doc.SelectNodes("/document/pref/shortcuts/auxInputReplace"));
            AuxInputReverse = getInnerText("auxInputReverse", doc.SelectNodes("/document/pref/shortcuts/auxInputReverse"));
            AuxInputDuplicate = getInnerText("auxInputDuplicate", doc.SelectNodes("/document/pref/shortcuts/auxInputDuplicate"));
            AuxInputFourArithmeticOperation = getInnerText("auxInputFourArithmeticOperation", doc.SelectNodes("/document/pref/shortcuts/auxInputFourArithmeticOperation"));

            return true;
        }

        // 初期設定ファイル(Shortcuts)の書き出し
        public bool saveShortcutsSetting(XmlDocument doc)
        {
            XmlNodeList list = doc.SelectNodes("/document/pref/shortcuts");
            XmlElement keys = list[0] as XmlElement;
            XmlElement undo = doc.CreateElement("undo");
            XmlElement aecopy = doc.CreateElement("remapToAE");
            XmlElement aecopy2 = doc.CreateElement("remapToAE2");
            XmlElement aecopy3 = doc.CreateElement("remapToAE3");
            XmlElement aepaste = doc.CreateElement("remapFromAE");
            XmlElement nuki = doc.CreateElement("komanuki");
            XmlElement nukicancel = doc.CreateElement("komanukiCancel");
            XmlElement hari = doc.CreateElement("harikomi");
            XmlElement haricancel = doc.CreateElement("harikomiCancel");
            XmlElement copy = doc.CreateElement("copy");
            XmlElement cut = doc.CreateElement("cut");
            XmlElement paste = doc.CreateElement("paste");
            XmlElement set30fps = doc.CreateElement("set30fps");
            XmlElement set24fps = doc.CreateElement("set24fps");
            XmlElement inputcellcount = doc.CreateElement("inputCellCount");
            XmlElement inputfirstframe = doc.CreateElement("inputFirstFrame");
            XmlElement inputkaracellvalue = doc.CreateElement("inputKaraCellValue");
            XmlElement karacellnomove = doc.CreateElement("karaCellNoMove");
            XmlElement inputsheetsec = doc.CreateElement("inputSheetSec");
            XmlElement setsheetdiv4 = doc.CreateElement("setSheetDiv4");
            XmlElement setsheetdiv6 = doc.CreateElement("setSheetDiv6");
            XmlElement setsheetdiv12 = doc.CreateElement("setSheetDiv12");
            XmlElement stayontop = doc.CreateElement("stayOnTop");
            XmlElement autoadjust = doc.CreateElement("autoAdjust");
            XmlElement displayframenumber = doc.CreateElement("displayFrameNumber");
            XmlElement alwaysAppend = doc.CreateElement("alwaysAppend");
            XmlElement afterfxpath = doc.CreateElement("afterFxPath");
            XmlElement afterfxoption = doc.CreateElement("afterFxOption");
            XmlElement allinitialize = doc.CreateElement("allInitialize");
            XmlElement auxInputRepeat = doc.CreateElement("auxInputRepeat");
            XmlElement auxInputSequentialNumber = doc.CreateElement("auxInputSequentialNumber");
            XmlElement auxInputReplace = doc.CreateElement("auxInputReplace");
            XmlElement auxInputReverse = doc.CreateElement("auxInputReverse");
            XmlElement auxInputDuplicate = doc.CreateElement("auxInputDuplicate");
            XmlElement auxInputFourArithmeticOperation = doc.CreateElement("auxInputFourArithmeticOperation");
            keys.AppendChild(undo);
            keys.AppendChild(aecopy);
            keys.AppendChild(aecopy2);
            keys.AppendChild(aecopy3);
            keys.AppendChild(aepaste);
            keys.AppendChild(nuki);
            keys.AppendChild(nukicancel);
            keys.AppendChild(hari);
            keys.AppendChild(haricancel);
            keys.AppendChild(copy);
            keys.AppendChild(cut);
            keys.AppendChild(paste);
            keys.AppendChild(set30fps);
            keys.AppendChild(set24fps);
            keys.AppendChild(inputcellcount);
            keys.AppendChild(inputfirstframe);
            keys.AppendChild(inputkaracellvalue);
            keys.AppendChild(karacellnomove);
            keys.AppendChild(inputsheetsec);
            keys.AppendChild(setsheetdiv4);
            keys.AppendChild(setsheetdiv6);
            keys.AppendChild(setsheetdiv12);
            keys.AppendChild(stayontop);
            keys.AppendChild(autoadjust);
            keys.AppendChild(displayframenumber);
            keys.AppendChild(alwaysAppend);
            keys.AppendChild(afterfxpath);
            keys.AppendChild(afterfxoption);
            keys.AppendChild(allinitialize);
            keys.AppendChild(auxInputRepeat);
            keys.AppendChild(auxInputSequentialNumber);
            keys.AppendChild(auxInputReplace);
            keys.AppendChild(auxInputReverse);
            keys.AppendChild(auxInputDuplicate);
            keys.AppendChild(auxInputFourArithmeticOperation);
            undo.InnerText = Undo;
            aecopy.InnerText = AECopy;
            aecopy2.InnerText = AECopy2;
            aecopy3.InnerText = AECopy3;
            aepaste.InnerText = AEPaste;
            nuki.InnerText = Nuki;
            nukicancel.InnerText = NukiCancel;
            hari.InnerText = Hari;
            haricancel.InnerText = HariCancel;
            copy.InnerText = Copy;
            cut.InnerText = Cut;
            paste.InnerText = Paste;
            set30fps.InnerText = Set30FPS;
            set24fps.InnerText = Set24FPS;
            inputcellcount.InnerText = InputCellCount;
            inputfirstframe.InnerText = InputFirstFrame;
            inputkaracellvalue.InnerText = InputKaraCellValue;
            karacellnomove.InnerText = KaraCellNoMove;
            inputsheetsec.InnerText = InputSheetSec;
            setsheetdiv4.InnerText = SetSheetDiv4;
            setsheetdiv6.InnerText = SetSheetDiv6;
            setsheetdiv12.InnerText = SetSheetDiv12;
            stayontop.InnerText = StayOnTop;
            autoadjust.InnerText = AutoAdjust;
            displayframenumber.InnerText = DisplayFrameNumber;
            alwaysAppend.InnerText = AlwaysAppend;
            afterfxpath.InnerText = AfterFXPath;
            afterfxoption.InnerText = AfterFXOption;
            allinitialize.InnerText = AllInitialize;
            auxInputRepeat.InnerText = AuxInputRepeat;
            auxInputSequentialNumber.InnerText = AuxInputSequentialNumber;
            auxInputReplace.InnerText = AuxInputReplace;
            auxInputReverse.InnerText = AuxInputReverse;
            auxInputDuplicate.InnerText = AuxInputDuplicate;
            auxInputFourArithmeticOperation.InnerText = AuxInputFourArithmeticOperation;
            return true;
        }

        // 初期設定ファイル(keybinds)の読み込み
        public bool loadKeybindsSetting(XmlDocument doc)
        {
            XmlNodeList binds = doc.SelectNodes("/document/pref/keybinds");
            foreach (XmlElement node in binds[0].ChildNodes)
            {
                // 16進数文字列⇒数値 変換
                int key = Convert.ToInt32(node.Attributes["original"].Value,16);
                int val = Convert.ToInt32(node.Attributes["custom"].Value,16);

                // キーペアを登録
                Keymap.Add(val, key);
                KeymapRev.Add(key, val);
            }
            return true;
        }

        // 初期設定ファイル(keybinds)の書き出し
        public bool saveKeybindsSetting(XmlDocument doc)
        {
            XmlNodeList list = doc.SelectNodes("/document/pref/keybinds");
            XmlElement binds = list[0] as XmlElement;
            foreach (KeyValuePair<int, int> keyPair in Keymap) {
                // 数値⇒16進数文字列 変換
                XmlElement pairTag = doc.CreateElement("pair");
                pairTag.SetAttribute("original", Convert.ToString(keyPair.Key,16).PadLeft(3, '0'));
                pairTag.SetAttribute("custom", Convert.ToString(keyPair.Value,16).PadLeft(3, '0'));

                // キーペアを登録
                binds.AppendChild(pairTag);
            }
            return true;
        }

        //コンビネーションキー状態の変換
        private int convShift(bool alt, bool ctrl, bool shift)
        {
            int code = 0;
            code |= alt ? (int)CombinationKeyState.ALTKey : 0;
            code |= ctrl ? (int)CombinationKeyState.CTRLKey : 0;
            code |= shift ? (int)CombinationKeyState.SHIFTKey : 0;
            return code;
        }

        //キーコードの変換
        public int convKey(int code, bool alt, bool ctrl, bool shift)
        {
            //変換した値を取得
            int value = code + convShift(alt, ctrl, shift);
            try
            {
                value = Keymap[value];
            }
            catch (Exception e)
            {
                // (対応する値がない場合はここに来る)
            }
            return value;
        }

        //変換前のコンビネーションキー状態の確認
        public bool checkShiftBeforeConvertion(int code, CombinationKeyState cmbkey)
        {
            // 変換前の値を取得
            int value = code;
            try
            {
                value = KeymapRev[value];
            }
            catch (Exception e)
            {
                // (対応する値がない場合はここに来る)
            }

            // コンビネーションキー情報を取り出す
            CombinationKeyState state = (CombinationKeyState)(value & 0x700);
            if (cmbkey == state)
            {
                // 該当する場合は trueを返す
                return true;
            }

            return false;
        }

        //変換前のコンビネーションキー状態の取得
        public CombinationKeyState getShiftBeforeConvertion(int code)
        {
            // 変換前の値を取得
            int value = code;
            try
            {
                value = KeymapRev[value];
            }
            catch (Exception e)
            {
                // (対応する値がない場合はここに来る)
            }

            // コンビネーションキー情報を取り出す
            return (CombinationKeyState)(value & 0x700);
        }

    }
}
