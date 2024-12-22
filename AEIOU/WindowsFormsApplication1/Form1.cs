// AfterEffects Input/Output Utility v3.x  Programed by kanbara.
//
//  ■v3.0
//  (24-12-06 19:35) v3.0.0.4
//      ・設定ファイルにバージョン文字列を持つように変更
//      ・起動時に、dataGridViewにフォーカスを当てる（起動時に、クリックしないとキー入力が効かないため）
//
//  (13-04-20 13:00) v3.0.0.3
//　　　・設定ファイル読み込みのエラー処理が足りなかった部分を修正
//　　　　→Settingsの読み込み時にエラー処理追加。
//
//  (13-04-16 20:30) v3.0.0.2
//　　　・twitterにて、（セル入力が常に追加編集なので）deleteキー押すのに慣れなくて..というようなツイートが..
//　　　　→v2同様、初回編集ステートを追加。
//　　　　　状態移行忘れ箇所も出るだろうから、他の機能を使った後のセル入力動作が怪しい部分があるかもしれない。
//　　　　→今までの動作を残す意味で、オプションに『セル入力を"常に追加"に設定』項目を追加。
//
//  (11-09-26 10:00) v3.0.0.1
//      ・バグ出し、挙動の変化している部分の調査（個々の機能を一通り使ってみる）
//      　・セル挿入/削除が動かなくなっている
//      　　→初期化実装の修正でエンバグしていた。領域の再確保を忘れていたので追加。
//		　・AEからペースト時に開始位置が１つ下にずれている
//      　　→違っていたのは "AEへコピー"の方だった。"AEへコピー"のフレーム計算を 0から数える計算に修正。
//		　・行削除の初期キーバインドが違う（現状:Ctrl+BS 以前:Shift+Del）
//      　　→前の実装に合わせた。
//		　・内部変数から初回入力フラグが消えたので、BackSpace時の削除動作が、常に１桁削除になっている
//		　　※他にもあるかもしれない
//      　　→ひとまずBackSpaceで巻き戻し削除時はセル内容の消去処理に実装を修正。
//          　新規入力～編集状態に入った場合ではない時も１桁削除のままなので、新規入力フラグがないと同じにできそうにない..
//		　・入力位置が画面2/3になると画面送りする動作がない
//      　　→実装を追加(Enter, ↓, ピリオド, +, -, K)
//      　　→キーフレーム頭出しの画面送り動作が特別処理だったのも普通に送るように修正
//      　・HOMEキー押し下げ時に選択領域の内部変数が初期化されていなかったのを修正
//		　・"+", "-"押しでカラセル番号に到達した場合に達した場合の実装がない
//      　　→同じ動作をするように実装していたが、繰り返しカラセル入力されるのもおかしいので、反応しないように変更。
//		　・AEからペーストでFPS設定が変わる場合に、メニューのチェック内容が追従していない
//      　　→変更する実装を追加
//		　・複数セルに跨った動作の実装を忘れている
//		　　・複数セル同時入力ができない
//      　　　→対応済：キー頭出し, 範囲削除(Delete, BS), 数値入力, カラセル入力
//      　　　→未対応：自動入力(+,-)
//		　・セル枚数指定実行時のアラート表示実装がない
//      　　→そもそも上限値を設定しないので必要がなかった（64bit化もあったので、今回特に縛りを設けていない）
//		　・セル名称編集実装がない
//      　　→実装した
//      　・セル名称が毎回初期化されている（挿入/削除でも初期化される）
//      　　→挿入/削除時の動作を修正（挿入セルは空欄。他は処理前のものを採用）
//      　　→セル枚数入力時の動作も別途修正（起動時は全セルに名称を付け、以降追加時は空欄にする）
//      　　→クリア選択時は、起動時と同じくセル名称を初期化するようにした
//		　・入力補助実装に対応するショートカットキーの設定項目がない
//      　　→実装した
//
//  (11-09-25 17:00) v3.0.0.0
//      →ほぼV2.1.0.3相当の実装を完了。
//      　バグや挙動の変わった点の洗い出し開始。
//
//  (11-09-08 15:30) v3.0.0.0
//      →C# & .NET Frameworkで書き直し中
//      ・フレーム数の表示
//      ・セルの選択単位毎のカーソル移動
//      ・セルの色分け（奇数偶数列（入力あり/なし）, アクティブ列, 選択範囲）
//      ・セルの基準線描画（24fps時 12k, 1秒, 6秒毎）
//      ・セル内容の表示（タイミング、カラセル、継続記号）
//      ・セルへの数値入力(アクティブセルのみ)、桁削除、選択領域のセル消去
//      ・AEへのコピー（アクティブセル）
//      ・列毎の入力状態(入力済/未入力)の処理
//      ・カラセル入力、加減算入力(+,-)、選択範囲の拡縮(*,/)
//      ・中抜き＆切り貼り範囲設定、範囲の描画
//      ・フレーム数表示：中抜き＆切り貼り、及び30fps他対応
//      ・基準線描画：中抜き＆切り貼り、及び30fps他対応
//      ・AEへのコピー（アクティブセル）：中抜き＆切り貼り対応
//      ・列の順序変更対応：セルの色分け意味を成さなくなる問題への対応
//        →CellPaintingイベントハンドラにて、DataGridView.Columns[e.ColumnIndex].DisplayIndex基準で色分けを行うようにした
//      ・アンドゥ機能（とりあえず実装）
//      　→削除時に使用状況が正しく反映されるように対応（空白⇒入力済, 入力済⇒空白 の状態変化時に対応）
//        →履歴の保存数を指定個数に制限するようにした（古いグループから順に削除, １回分の手順が制限を超えた場合も履歴に残らない）
//      ・切り貼り処理に自動コマ挿入＆削除動作を追加
//      　→アンドゥ対応すると登録個数が嵩む, 切り貼り動作はアンドゥ対象外にし、履歴は初期化する
//
//      ・確認事項
//      　→コマンドライン文字列の取得
//      　　→String[] cmds = System.Environment.GetCommandLineArgs();
//      　→ホイール動作
//      　　→実装不要だった。
//      　→別アプリの起動
//          →// using System.Diagnostics 
//      　　　// パラメータを指定して実行
//            Process.Start("notepad.exe", @"C:\boot.ini");
//      　→ショートカットの動的変更
//      　　→menuFileNew.ShortcutKeys = Keys.Control | Keys.N; とかそんな感じ
//          →Enumから文字列への変換は ここを参照
//          　http://devlabo.blogspot.com/2009/03/c.html
//          →試した結果は以下のような感じ
//                MessageBox.Show(this.undoToolStripMenuItem.ShortcutKeys.ToString());
//                Keys keys = (Keys)Enum.Parse(typeof(Keys),"A, Control");
//                this.undoToolStripMenuItem.ShortcutKeys = keys;
//      　　→メニューを使わない場合には仕掛けが必要な模様
//            http://youryella.wankuma.com/Library/Extensions/Button/ShortcutKey.aspx
//      　→テキストファイルの入出力方法
//          →読み込みは以下
//            StreamReader sr = new StreamReader(
//                "readme.txt", Encoding.GetEncoding("UTF-8"));
//            string text = sr.ReadToEnd();
//            sr.Close();
//
//          →書き込みは以下
//            String text = "サンプルテキスト";
//            System.IO.StreamWriter sw = new System.IO.StreamWriter(
//                @"c:\test.txt",
//                false,
//                System.Text.Encoding.GetEncoding("UTF-8"));
//            sw.Write(text);
//            sw.Close();
//
//      　→XMLファイルのパース処理（XPathが使えるか否かも重要）
//          →とりあえず以下を参照
//            http://msdn.microsoft.com/ja-jp/academic/cc987569
//
//      ・初期設定ファイル読込処理の実装
//      　→以前の初期設定ファイル"aeiou.dat"の要素を実装（ちょっと後半の要素は忘れていて嘘くさいけど..）
//      　→保存動作の実装
//      　→ショートカット設定読込動作の実装
//      　→キーコード設定読込動作の実装
//      ・AEへコピー(Script仲介)の実装
//      　→別アプリ起動動作の実装
//      ・棚田さんからのバグ報告で起動直後に"+"キーを押した時に"0"が入力される動作を修正
//      　→後でバグ取り予定だったけど、さすがにこれはないので忘れないうちに修正
//      　→ついでなので、マウスクリック時にselectRange更新動作をするように修正
//
//      ・DataGridViewColumn & DataGridViewCellを継承したクラスの定義、及びセル描画の移譲
//      ・キーコード変換動作の実装
//      　→とりあえず、読み書き処理のみ実装
//      ・AEへコピー(TimeRemap以外)の実装
//      ・AEからペースト 動作の実装
//      　→アンドゥ履歴はフラッシュする
//      ・選択範囲のコピー、カット＆ペースト 動作の実装
//      　→コピー、カット＆ペースト時のアンドゥ処理を考えないといけない..
//      　　現状何も考えてないので、単にペースト時にアンドゥをフラッシュしているが、これはアンドゥの意味がない感じ..
//      　　→アンドゥ実装に機能追加（操作毎のリスト作成、UndoOperationプロパティの見直し）
//      　　→コピー、カット＆ペースト時のアンドゥ対応
//      ・選択範囲ドラッグによる範囲移動 動作の実装
//      ・ページUP/DOWN 動作の実装
//      ・キーの頭出し 動作の実装
//      ・行 挿入 動作の実装
//      　→アンドゥ履歴はフラッシュする
//      ・行 削除 動作の実装
//      　→アンドゥ履歴はフラッシュする
//      ・fps切り替え 動作の実装
//      　→30fps時の基準線表示実装を間違っていたので修正
//      ・常に一番上に表示設定 動作の実装
//      ・カラセル入力時のカーソル移動設定 動作の実装
//      ・フレーム表示切替(シート/コマ⇔フレーム) 動作の実装
//      ・開始フレーム設定 動作の実装
//      　→なんだか美しくない実装だけど別ダイアログを用意して、それを呼び出すようにした
//      ・カラセル設定 動作の実装
//      ・シートの秒数 動作の実装
//      ・基準線の秒数 動作の実装
//      ・AfterFXパス設定 動作の実装
//      　→やり方が良く判ってないのでファイルダイアログではなく文字列入力式として実装
//      ・リマップ用jsx パス設定 動作の実装
//      　→やり方が良く判ってないのでファイルダイアログではなく文字列入力式として実装
//      ・作業内容の初期化 動作の実装
//      ・自動入力機能に、切り張り/中抜きフレームを読み飛ばす実装の追加
//      ・セル情報とアンドゥ履歴だけを初期化する 実装の追加
//      　→指定要素の初期化を行う初期化関数として実装（※現状は未使用：選択的に初期化が必要な用途向け）
//      ・セル枚数の設定 動作の実装
//      　→編集内容は消滅する実装
//      ・自動サイズ調整 動作の実装
//
//      ・ショートカット関連設定の項目追加分の実装
//      ・Ctrlキー同時押しのマルチセレクト未対応につき、その場合の抑止処理を追加
//      ・キーコード変換動作の実装
//      ・ファイルダイアログによる指定の実装（パス設定関連）
//      ・セル枚数設定時の実装を前実装と同じ動作に
//      　→アンドゥ履歴やコピーバッファはフラッシュする
//    　・セル挿入 動作の実装
//      　→アンドゥ履歴やコピーバッファはフラッシュする
//      　⇒元の実装は複数セル処理するようになっていたが、とりあえず１セル追加処理を実装
//    　・セル削除 動作の実装
//      　→アンドゥ履歴やコピーバッファはフラッシュする
//      　⇒元の実装は複数セル処理するようになっていたが、とりあえず１セル追加処理を実装
//      ・アンドゥ非対応機能について、アンドゥ実装の改修方法の検討
//      　→考えてはみたものの、やっぱり行や列が減った状況では、
//      　　履歴やコピーバッファの内容を正しく取得反映できない場面が出てくる。体系的に対応するにはノウハウが足りない..
//      ・フレーム数描画処理をCellクラスへの移譲を検討
//		　→やっぱり見送り..(--;
//      ・リドゥ動作についての検討
//		　→やっぱり見送り..(--;
//      ・Plugin機能実装の検討
//      　(参考サイト)　http://dobon.net/vb/dotnet/programing/plugin.html
//      　→入力補助用の実装をユーザー拡張可能に出来るか検討してみる
//		　　・入力ダイアログを開きパラメータを入力させ、パラメータに沿ってセルに結果を入力
//		　　・pluginに入力ダイアログ作成、結果パラメータによるセル入力を任せる
//      　→入出力の実装をユーザー拡張可能に出来るかも検討してみる
//		　　・load/saveダイアログを開きファイルパスを入力させ、ファイル⇔セルの読み書き
//		　　・pluginにはファイル⇔セルの読み書きを任せる
//		　　・オプション設定可能な場合を考えると、オプションダイアログ作成、設定の読み書きが必要
//			　→設定読み書きの方法はどうするか..
//			　　→settingクラスのload/save時にPluginのload/saveメソッドを呼べるようにしないといけない
//      　→その他ユーザー拡張可能な動作の洗い出し
//		　→テスト実装をしてみようと簡単なサンプルを作ってみる最中に、意外と事前準備が手間なことが判った
//		　　慣れてないのもあると思うが、もうちょっとこの手の実装に習熟してから再度検討する
//		　　実際問題お手軽に使えないのでは、誰もわざわざ使ってはくれない
//
//      （入力補助関連）
//      ・連番作成 動作の実装
//      ・繰り返し 動作の実装
//      ・選択範囲の並びを反転 動作の実装
//      ・置き換え 動作の実装
//      ・四則演算 動作の実装
//      ・選択範囲を２回繰り返す(複製ボタン) 動作の実装
//
//      （STS入出力関連）
//      ・STS入出力 動作の実装
//      ・STS用カット番号入力フォーム の実装
//      　⇒必要度が判らなくなってきたので実装保留
//      ・指定ファイル読み込み 動作の実装
//
//      ⇒バグ出し、挙動の変化している部分の調査（個々の機能を一通り使ってみる）
//      　・セル挿入/削除が動かなくなっている
//      　　→初期化実装の修正でエンバグしていた。領域の再確保を忘れていたので追加。
//		　・AEからペースト時に開始位置が１つ下にずれている
//      　　→違っていたのは "AEへコピー"の方だった。"AEへコピー"のフレーム計算を 0から数える計算に修正。
//		　・行削除の初期キーバインドが違う（現状:Ctrl+BS 以前:Shift+Del）
//      　　→前の実装に合わせた。
//		　・内部変数から初回入力フラグが消えたので、BackSpace時の削除動作が、常に１桁削除になっている
//		　　※他にもあるかもしれない
//      　　→ひとまずBackSpaceで巻き戻し削除時はセル内容の消去処理に実装を修正。
//          　新規入力～編集状態に入った場合ではない時も１桁削除のままなので、新規入力フラグがないと同じにできそうにない..
//		　・入力位置が画面2/3になると画面送りする動作がない
//      　　→実装を追加(Enter, ↓, ピリオド, +, -, K)
//      　　→キーフレーム頭出しの画面送り動作が特別処理だったのも普通に送るように修正
//      　・HOMEキー押し下げ時に選択領域の内部変数が初期化されていなかったのを修正
//		　・"+", "-"押しでカラセル番号に到達した場合に達した場合の実装がない
//      　　→同じ動作をするように実装していたが、繰り返しカラセル入力されるのもおかしいので、反応しないように変更。
//		　・AEからペーストでFPS設定が変わる場合に、メニューのチェック内容が追従していない
//      　　→変更する実装を追加
//		　・複数セルに跨った動作の実装を忘れている
//		　　・複数セル同時入力ができない
//      　　　→対応済：キー頭出し, 範囲削除(Delete, BS), 数値入力, カラセル入力
//      　　　→未対応：自動入力(+,-)
//		　・セル枚数指定実行時のアラート表示実装がない
//      　　→そもそも上限値を設定しないので必要がなかった（64bit化もあったので、今回特に縛りを設けていない）
//		　・セル名称編集実装がない
//      　　→実装した
//      　・セル名称が毎回初期化されている（挿入/削除でも初期化される）
//      　　→挿入/削除時の動作を修正（挿入セルは空欄。他は処理前のものを採用）
//      　　→セル枚数入力時の動作も別途修正（起動時は全セルに名称を付け、以降追加時は空欄にする）
//		　・入力補助実装に対応するショートカットキーの設定項目がない
//      　　→実装した
//
//      （バージョン管理関連）
//      ⇒バージョンの色変更 動作の実装
//      ⇒バージョン管理(入出力) 動作の実装
//      ⇒バージョン変更 動作の実装
//      ⇒指定バージョンの読み込み 動作の実装
//
//

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Xml;
using System.Diagnostics;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;

namespace AEIOU
{

    // 基準線の状態 を表す列挙体
    // (パクリ元)⇒http://smdn.jp/programming/netfx/enum/1_flags/
    [Flags]
    enum SheetBorder : int
    {
        EveryNFrames = 0x00000001,  //ｎコマ毎
        EverySec = 0x00000002,      //１秒毎
        EverySheet = 0x00000004,    //シート毎
        None = 0x00000000,          //なし
    }

    // リマップ出力モード を表す列挙体
    [Flags]
    enum RemapOutputMode : int
    {
        AERemap = 0x00000001,       //AEへコピー(通常のリマップ計算)
        DirectRemap = 0x00000002,   //AEへコピー(ダイレクトリマップ)
        JSRemap = 0x00000004,       //JavaScript経由のコピー(通常のリマップ計算)
        JSDirectRemap = 0x00000008, //JavaScript経由のコピー(ダイレクトリマップ)
        None = 0x00000000,          //なし
    }

    // アンドゥ反映対象 を表す列挙体
    [Flags]
    enum UndoOperationTarget : int
    {
        ToGrid = 0x00000001,            //グリッドへ
        ToCopyBuffer = 0x00000002,      //コピーバッファへ
        MoveToCopyBuffer = 0x00000003,  //グリッドからコピーバッファへ移動
        FromCopyBuffer = 0x00000004,    //コピーバッファからグリッドへ
        None = 0x00000000,              //なし
    }

    // 初期化対象 を表す列挙体
    [Flags]
    enum InitializeTarget : int
    {
        EditTemp = 0x00000001,          //カーソル位置, 選択範囲などの作業用変数
        Timing = 0x00000002,            //タイミング情報
        CopyBuffer = 0x00000004,        //コピーバッファ
        UndoHistory = 0x00000008,       //アンドゥ履歴
        RangeSetting = 0x00000010,      //範囲設定
        None = 0x00000000,              //なし
    }

    // コンビネーションキー を表す列挙体
    [Flags]
    enum CombinationKeyState : int
    {
        ALTKey = 0x00000100,            //ALT
        CTRLKey = 0x00000200,           //CTRL
        SHIFTKey = 0x00000400,          //SHIFT
        None = 0x00000000,              //なし
    }

    // 四則演算のモード を表す列挙体
    enum CalcMode : int
    {
        Plus = 0x00000001,              //加算
        Minus,                          //減算
        Multiple,                       //乗算
        Divide,                         //除算
        None = 0x00000000,              //なし
    }


    public partial class Form1 : Form
    {
        //----------------------------------------------------------------------------------------
        // コンストラクタ
        public Form1()
        {
            InitializeComponent();

            // 自分のウィンドウハンドルを取得しておく（tooltip表示中に、違うウィンドウハンドルを取得してしまうので）
            this.owner = Control.FromHandle(this.Handle);

            //datagridviewのダブルバッファを有効にする
            //参考:http://raluck.exblog.jp/14873007/
            typeof(DataGridView).
                GetProperty("DoubleBuffered",
                    BindingFlags.Instance | BindingFlags.NonPublic).
                SetValue(this.dataGridView1, true, null);

            //コマンドラインを配列で取得する
            // →起動スイッチのコマ数指定などは削除（⇒初期設定ファイルに移譲）
            string[] cmds;
            cmds = System.Environment.GetCommandLineArgs();

            // 実行ファイルのパスを取得（※設定ファイル保存などの用途）
            setting.CurrentDir = Path.GetDirectoryName(cmds[0]);

            // 初期設定ファイルの読み込み
            if (setting.loadSettingFile(setting.CurrentDir + @"\aeiou.ini") == false)
            {
                // 初期設定ファイルがない場合
                this.StartPosition = FormStartPosition.Manual;
                this.Location = new Point(0, 0);
                this.Size = new Size(300, 450);
                this.TopMost = false;
                
                // 状態をメニューに反映
                this.fPS30ToolStripMenuItem.Checked = (setting.Fps == 30) ? true : false;
                this.fPS24ToolStripMenuItem.Checked = (setting.Fps == 24) ? true : false;
                this.stayOnTopToolStripMenuItem.Checked = setting.TopMost;
                this.karacellNoMoveToolStripMenuItem.Checked = setting.IsKaraNoMove;
                this.displayFrameNumberToolStripMenuItem.Checked = setting.IsDisplayFrameNumber;
                this.div4ToolStripMenuItem.Checked = (setting.SheetDivide == 4) ? true : false;
                this.div6ToolStripMenuItem.Checked = (setting.SheetDivide == 6) ? true : false;
                this.div12ToolStripMenuItem.Checked = (setting.SheetDivide == 12) ? true : false;
                this.autoAdjustToolStripMenuItem.Checked = setting.IsAutoadjust;
                this.alwaysAppendToolStripMenuItem.Checked = setting.IsAlwaysAppend;
            }
            else
            {
                // 初期設定ファイルがあった場合
                this.StartPosition = setting.StartPosition;
                this.Location = setting.Location;
                this.Size = setting.Size;
                this.TopMost = setting.TopMost;

                // 状態をメニューに反映
                this.fPS30ToolStripMenuItem.Checked = (setting.Fps == 30) ? true : false;
                this.fPS24ToolStripMenuItem.Checked = (setting.Fps == 24) ? true : false;
                this.stayOnTopToolStripMenuItem.Checked = setting.TopMost;
                this.karacellNoMoveToolStripMenuItem.Checked = setting.IsKaraNoMove;
                this.displayFrameNumberToolStripMenuItem.Checked = setting.IsDisplayFrameNumber;
                this.div4ToolStripMenuItem.Checked = (setting.SheetDivide == 4) ? true : false;
                this.div6ToolStripMenuItem.Checked = (setting.SheetDivide == 6) ? true : false;
                this.div12ToolStripMenuItem.Checked = (setting.SheetDivide == 12) ? true : false;
                this.autoAdjustToolStripMenuItem.Checked = setting.IsAutoadjust;
                this.alwaysAppendToolStripMenuItem.Checked = setting.IsAlwaysAppend;

                // ショートカットの反映
                this.undoToolStripMenuItem.ShortcutKeys = (Keys)Enum.Parse(typeof(Keys), setting.keys.Undo);
                this.aECopyToolStripMenuItem.ShortcutKeys = (Keys)Enum.Parse(typeof(Keys), setting.keys.AECopy);
                this.directRemapToolStripMenuItem.ShortcutKeys = (Keys)Enum.Parse(typeof(Keys), setting.keys.AECopy2);
                this.jSRemapToolStripMenuItem.ShortcutKeys = (Keys)Enum.Parse(typeof(Keys), setting.keys.AECopy3);
                this.pasteFromAEToolStripMenuItem.ShortcutKeys = (Keys)Enum.Parse(typeof(Keys), setting.keys.AEPaste);
                this.setNakanukiToolStripMenuItem.ShortcutKeys = (Keys)Enum.Parse(typeof(Keys), setting.keys.Nuki);
                this.cancelNakanukiToolStripMenuItem.ShortcutKeys = (Keys)Enum.Parse(typeof(Keys), setting.keys.NukiCancel);
                this.setKiribariToolStripMenuItem.ShortcutKeys = (Keys)Enum.Parse(typeof(Keys), setting.keys.Hari);
                this.cancelKiribariToolStripMenuItem.ShortcutKeys = (Keys)Enum.Parse(typeof(Keys), setting.keys.HariCancel);
                this.copyToolStripMenuItem.ShortcutKeys = (Keys)Enum.Parse(typeof(Keys), setting.keys.Copy);
                this.cutToolStripMenuItem.ShortcutKeys = (Keys)Enum.Parse(typeof(Keys), setting.keys.Cut);
                this.pasteToolStripMenuItem.ShortcutKeys = (Keys)Enum.Parse(typeof(Keys), setting.keys.Paste);
                this.fPS30ToolStripMenuItem.ShortcutKeys = (Keys)Enum.Parse(typeof(Keys), setting.keys.Set30FPS);
                this.fPS30ToolStripMenuItem.ShortcutKeys = (Keys)Enum.Parse(typeof(Keys), setting.keys.Set24FPS);
                this.inputCellCountToolStripMenuItem.ShortcutKeys = (Keys)Enum.Parse(typeof(Keys), setting.keys.InputCellCount);
                this.firstFrameToolStripMenuItem.ShortcutKeys = (Keys)Enum.Parse(typeof(Keys), setting.keys.InputFirstFrame);
                this.karacellValueToolStripMenuItem.ShortcutKeys = (Keys)Enum.Parse(typeof(Keys), setting.keys.InputKaraCellValue);
                this.karacellNoMoveToolStripMenuItem.ShortcutKeys = (Keys)Enum.Parse(typeof(Keys), setting.keys.KaraCellNoMove);
                this.secondsPerSheetToolStripMenuItem.ShortcutKeys = (Keys)Enum.Parse(typeof(Keys), setting.keys.InputSheetSec);
                this.div4ToolStripMenuItem.ShortcutKeys = (Keys)Enum.Parse(typeof(Keys), setting.keys.SetSheetDiv4);
                this.div6ToolStripMenuItem.ShortcutKeys = (Keys)Enum.Parse(typeof(Keys), setting.keys.SetSheetDiv6);
                this.div12ToolStripMenuItem.ShortcutKeys = (Keys)Enum.Parse(typeof(Keys), setting.keys.SetSheetDiv12);
                this.stayOnTopToolStripMenuItem.ShortcutKeys = (Keys)Enum.Parse(typeof(Keys), setting.keys.StayOnTop);
                this.autoAdjustToolStripMenuItem.ShortcutKeys = (Keys)Enum.Parse(typeof(Keys), setting.keys.AutoAdjust);
                this.displayFrameNumberToolStripMenuItem.ShortcutKeys = (Keys)Enum.Parse(typeof(Keys), setting.keys.DisplayFrameNumber);
                this.afterFXPathToolStripMenuItem.ShortcutKeys = (Keys)Enum.Parse(typeof(Keys), setting.keys.AfterFXPath);
                this.afterFXOptionToolStripMenuItem.ShortcutKeys = (Keys)Enum.Parse(typeof(Keys), setting.keys.AfterFXOption);
                this.allInitializeToolStripMenuItem.ShortcutKeys = (Keys)Enum.Parse(typeof(Keys), setting.keys.AllInitialize);
                this.repeatNumberToolStripMenuItem.ShortcutKeys = (Keys)Enum.Parse(typeof(Keys), setting.keys.AuxInputRepeat);
                this.sequentialNumberToolStripMenuItem.ShortcutKeys = (Keys)Enum.Parse(typeof(Keys), setting.keys.AuxInputSequentialNumber);
                this.replaceToolStripMenuItem.ShortcutKeys = (Keys)Enum.Parse(typeof(Keys), setting.keys.AuxInputReplace);
                this.reverseToolStripMenuItem.ShortcutKeys = (Keys)Enum.Parse(typeof(Keys), setting.keys.AuxInputReverse);
                this.duplicateToolStripMenuItem.ShortcutKeys = (Keys)Enum.Parse(typeof(Keys), setting.keys.AuxInputDuplicate);
                this.fourArithmeticOperationToolStripMenuItem.ShortcutKeys = (Keys)Enum.Parse(typeof(Keys), setting.keys.AuxInputFourArithmeticOperation);

            }

            // 作業データの初期化
            InitializeWork(true);

            // 読み込みファイル指定がある場合 ファイル読込を行う
            if (cmds.Length > 1 && File.Exists(cmds[1]))
            {
                loadSTS(cmds[1]);
            }

            // グリッドビューにフォーカスを設定
            dataGridView1.Focus();
        }

        //----------------------------------------------------------------------------------------
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            // 終了処理

            // ウィンドウの状態を settingに反映
            setting.StartPosition = this.StartPosition;
            setting.Location = this.Location;
            setting.Size = this.Size;
            setting.TopMost = this.TopMost;

            // ショートカットの状態を setting.keysに反映
            setting.keys.Undo = this.undoToolStripMenuItem.ShortcutKeys.ToString();
            setting.keys.AECopy = this.aECopyToolStripMenuItem.ShortcutKeys.ToString();
            setting.keys.AECopy2 = this.directRemapToolStripMenuItem.ShortcutKeys.ToString();
            setting.keys.AECopy3 = this.jSRemapToolStripMenuItem.ShortcutKeys.ToString();
            setting.keys.AEPaste = this.pasteFromAEToolStripMenuItem.ShortcutKeys.ToString();
            setting.keys.Nuki = this.setNakanukiToolStripMenuItem.ShortcutKeys.ToString();
            setting.keys.NukiCancel = this.cancelNakanukiToolStripMenuItem.ShortcutKeys.ToString();
            setting.keys.Hari = this.setKiribariToolStripMenuItem.ShortcutKeys.ToString();
            setting.keys.HariCancel = this.cancelKiribariToolStripMenuItem.ShortcutKeys.ToString();
            setting.keys.Copy = this.copyToolStripMenuItem.ShortcutKeys.ToString();
            setting.keys.Cut = this.cutToolStripMenuItem.ShortcutKeys.ToString();
            setting.keys.Paste = this.pasteToolStripMenuItem.ShortcutKeys.ToString();
            setting.keys.Set30FPS = this.fPS30ToolStripMenuItem.ShortcutKeys.ToString();
            setting.keys.Set24FPS = this.fPS24ToolStripMenuItem.ShortcutKeys.ToString();
            setting.keys.InputCellCount = this.inputCellCountToolStripMenuItem.ShortcutKeys.ToString();
            setting.keys.InputFirstFrame = this.firstFrameToolStripMenuItem.ShortcutKeys.ToString();
            setting.keys.InputKaraCellValue = this.karacellValueToolStripMenuItem.ShortcutKeys.ToString();
            setting.keys.KaraCellNoMove = this.karacellNoMoveToolStripMenuItem.ShortcutKeys.ToString();
            setting.keys.InputSheetSec = this.secondsPerSheetToolStripMenuItem.ShortcutKeys.ToString();
            setting.keys.SetSheetDiv4 = this.div4ToolStripMenuItem.ShortcutKeys.ToString();
            setting.keys.SetSheetDiv6 = this.div6ToolStripMenuItem.ShortcutKeys.ToString();
            setting.keys.SetSheetDiv12 = this.div12ToolStripMenuItem.ShortcutKeys.ToString();
            setting.keys.StayOnTop = this.stayOnTopToolStripMenuItem.ShortcutKeys.ToString();
            setting.keys.AutoAdjust = this.autoAdjustToolStripMenuItem.ShortcutKeys.ToString();
            setting.keys.DisplayFrameNumber = this.displayFrameNumberToolStripMenuItem.ShortcutKeys.ToString();
            setting.keys.AfterFXPath = this.afterFXPathToolStripMenuItem.ShortcutKeys.ToString();
            setting.keys.AfterFXOption = this.afterFXOptionToolStripMenuItem.ShortcutKeys.ToString();
            setting.keys.AllInitialize = this.allInitializeToolStripMenuItem.ShortcutKeys.ToString();
            setting.keys.AuxInputRepeat = this.repeatNumberToolStripMenuItem.ShortcutKeys.ToString();
            setting.keys.AuxInputSequentialNumber = this.sequentialNumberToolStripMenuItem.ShortcutKeys.ToString();
            setting.keys.AuxInputReplace = this.replaceToolStripMenuItem.ShortcutKeys.ToString();
            setting.keys.AuxInputReverse = this.reverseToolStripMenuItem.ShortcutKeys.ToString();
            setting.keys.AuxInputDuplicate = this.duplicateToolStripMenuItem.ShortcutKeys.ToString();
            setting.keys.AuxInputFourArithmeticOperation = this.fourArithmeticOperationToolStripMenuItem.ShortcutKeys.ToString();
            setting.keys.AlwaysAppend = this.alwaysAppendToolStripMenuItem.ShortcutKeys.ToString();

            // settingの状態を書き出す
            setting.saveSettingFile(setting.CurrentDir + @"\aeiou.ini");

        }

        //----------------------------------------------------------------------------------------
        Control owner;                              //子ウィンドウ（ダイアログ）に与える自分のウィンドウハンドル
        
        // 配色
        ColorDefinitions gridPalette = new ColorDefinitions();        //グリッド描画用 色設定

        // 設定用 各種変数
        Settings setting = new Settings();

        // 各種変数
        Rect selectRange;                           //選択範囲
        Rect copyRect;                              //コピー範囲
        Point mouseDownPoint;                       //マウス押下位置
        bool isRectDrag;                            //範囲移動状態
        bool isFirstEdit;                           //初回編集状態（セル入力を中断するような操作の度に trueにされるべき）

        // 配列
        int[] aryCellUsedCount;                     //セル使用状況
        String[,] aryCopyBuf;                       //コピーバッファ
        //int[, ,] gAryBuff;                          //バッファ

        // リスト
        List<Range> delRange;                       // 中抜き範囲
        List<Range> addRange;                       // 切り貼り範囲
        List<UndoGroup> undoList;                   // アンドゥリスト


        //----------------------------------------------------------------------------------------
        // 作業内容の初期化
        private void InitializeWork(bool isBoot)
        {
            // ヘッダ、各カラム及び行数の設定
            dataGridInitialize(setting.ColLength, setting.RowLength, 50, isBoot);    // 列, 行, 列幅

            // 配列作成
            aryCellUsedCount = new int[setting.ColLength];
            aryCopyBuf = new String[setting.ColLength, setting.RowLength];
            //gAryBuff = new int[setting.ColLength, setting.RowLength, gEtcLength];

            for (int i = 0; i < setting.RowLength; i++)
                for (int j = 0; j < setting.ColLength; j++)
                {
                    aryCopyBuf[j, i] = "";
                }

            // リスト作成
            delRange = new List<Range>();
            addRange = new List<Range>();
            undoList = new List<UndoGroup>();

            // アンドゥメニュを初期化
            setUndoText("");

            // 各種変数の初期化
            selectRange = new Rect(0, 0, 1, 1);     //選択範囲
            copyRect = new Rect(-1, -1, 0, 0);      //コピー範囲
            mouseDownPoint = new Point(-1, -1);     //マウス押下位置
            isRectDrag = false;                     //範囲移動状態
            isFirstEdit = true;
        }

        //----------------------------------------------------------------------------------------
        // 作業内容の初期化(※項目毎に初期化可否を指定)
        private void InitializeWork(InitializeTarget target)
        {
            if ((target & InitializeTarget.EditTemp) != 0)
            {
                // カーソル位置、選択範囲等 作業用変数の初期化
                dataGridView1.CurrentCell = dataGridView1[0,0]; //カーソル位置
                mouseDownPoint = new Point(-1, -1);             //マウス押下位置
                selectRange = new Rect(0, 0, 1, 1);             //選択範囲
                isRectDrag = false;                             //範囲移動状態
                isFirstEdit = true;
            }
            if ((target & InitializeTarget.Timing) != 0)
            {
                // シートの入力情報（タイミング）の初期化
                // ※バージョンの扱いをどうするのかは未定
                dataGridInitialize(setting.ColLength, setting.RowLength, 50, false);    // 列, 行, 列幅
                aryCellUsedCount = new int[setting.ColLength];
                for (int i = 0; i < setting.RowLength; i++)
                    for (int j = 0; j < setting.ColLength; j++)
                    {
                        dataGridView1[j, i].Value = "";
                    }
                isFirstEdit = true;
            }
            if ((target & InitializeTarget.CopyBuffer) != 0)
            {
                // コピーバッファの初期化
                copyRect = new Rect(-1, -1, 0, 0);     //コピー範囲
                aryCopyBuf = new String[setting.ColLength, setting.RowLength];
                for (int i = 0; i < setting.RowLength; i++)
                    for (int j = 0; j < setting.ColLength; j++)
                    {
                        aryCopyBuf[j, i] = "";
                    }
            }
            if ((target & InitializeTarget.UndoHistory) != 0)
            {
                // アンドゥ履歴の初期化
                undoList = new List<UndoGroup>();
                setUndoText("");                    // アンドゥメニュを初期化
            }
            if ((target & InitializeTarget.RangeSetting) != 0)
            {
                // 中抜き・切り貼り設定範囲の初期化
                delRange = new List<Range>();
                addRange = new List<Range>();
            }

        }

        //----------------------------------------------------------------------------------------
        // DataGridViewのヘッダ、各カラム及び行数の設定を初期化
        private void dataGridInitialize(int columnCount, int rowCount, int columnWidth, bool isBoot)
        {
            dataGridView1.Rows.Clear();
            dataGridView1.Columns.Clear();
            dataGridView1.RowTemplate.Height = 16;  // セルの高さ(新規追加分)


            // "各セル" 列の設定
            // ※セル名称の設定は起動時のみ行う
            for (int i = 0; i < columnCount; i++)
            {
                String cellName;
                if (i < 26 && isBoot)
                {
                    // 'A'-'Z'なら、セル名称を設定
                    byte[] data = { (byte)i };
                    data[0] += (byte)'A';
                    cellName = System.Text.Encoding.GetEncoding(932).GetString(data);
                }
                else
                {
                    // 'Z'以降は セル名称を空欄に
                    cellName = "";
                }

                // 列情報を TimingColumnで登録
                TimingColumn col = new TimingColumn();
                col.HeaderText = cellName;
                col.Width = columnWidth;
                col.SortMode = DataGridViewColumnSortMode.NotSortable; //ヘッダークリックによるソート動作を禁止
                this.dataGridView1.Columns.Add(col);
            }

            // 初期化 (行の生成 : 中身は空)
            dataGridView1.RowCount = setting.RowLength;
            for (int i = 0; i < setting.RowLength; i++)
                for (int j = 0; j < setting.ColLength; j++)
                    dataGridView1[j, i].Value = "";

        }

        //----------------------------------------------------------------------------------------
        private void resizeDataGridView1(int newCol, int newRow)
        {
            // dataGridView1のグリッドサイズを変更する

            // 現在の状態をテンポラリにコピー
            // コピーするセル数は、新しいセル数が少なければ新しいセル数に、そうでなければ古いセル数の分だけ行う
            int col = (newCol < setting.ColLength) ? newCol : setting.ColLength;
            int row = (newRow < setting.RowLength) ? newRow : setting.RowLength;
            String[,] temp = new String[col, row];
            String[] headerName = new String[col];
            int[] tempCellUsed = new int[col];
            for (int i = 0; i < col; i++)
            {
                tempCellUsed[i] = aryCellUsedCount[i];
                headerName[i] = dataGridView1.Columns[i].HeaderText;
                for (int j = 0; j < row; j++)
                {
                    temp[i, j] = dataGridView1[i, j].Value.ToString();
                }
            }

            // グリッド（及びコピーバッファ）を初期化
            setting.ColLength = newCol;
            setting.RowLength = newRow;
            InitializeWork(InitializeTarget.Timing);

            // テンポラリ内容を書き戻す
            for (int i = 0; i < col; i++)
            {
                aryCellUsedCount[i] = tempCellUsed[i];
                dataGridView1.Columns[i].HeaderText = headerName[i];
                for (int j = 0; j < row; j++)
                {
                    dataGridView1[i, j].Value = temp[i, j];
                }
            }
        }

        //---------------------------------------------------------------------------
        void adjustWindowSize()
        {
            //ウィンドウサイズ・位置を調整
            if (setting.IsAutoadjust)
            {
                //ウィンドウサイズの計算
                int columnWidth = dataGridView1.Columns[0].Width;
                int dividerWidth = dataGridView1.Columns[0].DividerWidth;
                int width = (columnWidth + dividerWidth) * (setting.ColLength + 1) + 36;
                int screenWidth = Screen.PrimaryScreen.Bounds.Width;

                //新しいサイズがウィンドウに収まらない場合は調整する
                width = (screenWidth < width) ? screenWidth : width;

                //新しいサイズをフォームに反映
                this.Width = width;

                //位置のチェック
                // ※フォーム左右の端のどちらが、画面端に近いのか
                if (this.Left > (screenWidth - this.Right))
                {
                    //画面右寄せ
                    this.Left = screenWidth - width;
                    //Form4->Left = Form1->Left;
                }
                else
                {
                    //画面左寄せ
                    this.Left = 0;
                    //Form4->Left = Form1->Left;
                }
            }
        }

        //----------------------------------------------------------------------------------------
        private void undoFunction()
        {
            if (undoList.Count == 0)
            {
                //リストが空の場合はなにもしない
                return;
            }

            // 操作を１つ取り出す（※リストから省く）
            UndoGroup group = undoList[undoList.Count - 1];
            undoList.Remove(group);

            // 操作を逆順に復元
            for (int i = group.Operations.Count - 1; i >= 0; i--)
            {
                // 操作を取り出す
                UndoOperation undo = group.Operations[i];
                group.Operations.Remove(undo);

                // グリッドへの編集を元に戻す
                if (group.Target == UndoOperationTarget.ToGrid ||
                    group.Target == UndoOperationTarget.MoveToCopyBuffer ||
                    group.Target == UndoOperationTarget.FromCopyBuffer)
                {
                    bool isBlank = dataGridView1[undo.Column, undo.Row].Value.ToString() == "" ? true : false;
                    String val = "";
                    if (group.Target == UndoOperationTarget.MoveToCopyBuffer)
                    {
                        // グリッドからバッファに移動した場合
                        val = undo.NewValue;
                    }
                    else
                    {
                        // グリッドを編集した場合
                        // グリッドからバッファにコピーした場合
                        val = undo.OldValue;
                    }

                    // グリッドの値を復元
                    dataGridView1[undo.Column, undo.Row].Value = val;

                    // セルの中身が"空白 or 入力済み"に変わった場合は、使用状況を修正
                    if (val.Length == 0 && !isBlank)
                    {
                        // 入力済み ⇒ 空白
                        aryCellUsedCount[undo.Column]--;
                    }
                    if (val.Length > 0 && isBlank)
                    {
                        // 空白 ⇒ 入力済み
                        aryCellUsedCount[undo.Column]++;
                    }
                    isFirstEdit = true;
                }

                // バッファへの編集を元に戻す
                if (group.Target == UndoOperationTarget.ToCopyBuffer ||
                    group.Target == UndoOperationTarget.MoveToCopyBuffer)
                {
                    aryCopyBuf[undo.Column, undo.Row] = undo.OldValue;
                }

            }
            // バッファ⇒グリッドの編集を元に戻した場合は 範囲設定も復元
            if (group.Target == UndoOperationTarget.FromCopyBuffer)
            {
                copyRect = group.Rect;
            }

            // アンドゥメニュ更新 (※空の場合もあり)
            if (undoList.Count > 0)
                setUndoText(undoList[undoList.Count - 1].Title);
            else
                setUndoText("");

            // 描画更新(継続記号の更新の為)
            dataGridView1.Invalidate();
        }

        //---------------------------------------------------------------------------
        private void copyToCell(int col, int row, Rect rect)
        {
            // アンドゥグループの作成
            UndoGroup undoGroup = new UndoGroup(
                setting.UndoCount++, "ペースト", UndoOperationTarget.FromCopyBuffer);

            // 範囲を記録
            undoGroup.Rect = rect;

            //選択範囲の内容(バッファ)をセルに複写
            for (int r = rect.Height - 1; r >= 0; r--)
                for (int c = rect.Width - 1; c >= 0; c--)
                {
                    int gx = col + c;
                    int gy = row + r;
                    int bx = rect.Left + c;
                    int by = rect.Top + r;

                    if (by < 0 || bx < 0) continue;
                    if (gy < 0 || gx < 0) continue;
                    if (dataGridView1.Rows.Count <= by ||
                       dataGridView1.Columns.Count <= bx) continue;
                    if (dataGridView1.Rows.Count <= gy ||
                       dataGridView1.Columns.Count <= gx) continue;

                    DataGridViewCell cell = dataGridView1[gx, gy];
                    String val = aryCopyBuf[bx, by];
                    if (cell.Value.ToString() == "" && val != "")
                    {
                        // 空のセルに入力する場合はカウント＋１
                        aryCellUsedCount[gx]++;

                        // ⇒修正時はバージョン変更
                    }
                    else if (cell.Value.ToString() != "" && val == "")
                    {
                        // 中身のあるセルを空にする場合はカウント－１
                        aryCellUsedCount[gx]--;

                        // ⇒修正時はバージョン変更
                    }

                    // アンドゥ情報の作成
                    UndoOperation undo = new UndoOperation(gx, gy, cell.Value.ToString(), val);

                    // セルにペースト
                    cell.Value = val;

                    // アンドゥ情報の記録
                    undoGroup.Operations.Add(undo);
                }

            // アンドゥ情報の記録
            undoList.Add(undoGroup);
            // アンドゥメニュ更新
            setUndoText(undoGroup.Title);
            // アンドゥ履歴の調整
            processUndoLimit();
        }

        //----------------------------------------------------------------------------------------
        private void copyToBuf(Rect rect)
        {
            // アンドゥグループの作成
            UndoGroup undoGroup = new UndoGroup(
                setting.UndoCount++, "コピー", UndoOperationTarget.ToCopyBuffer);

            // 範囲の記録
            undoGroup.Rect = rect;

            //選択範囲をバッファに複写
            for (int r = rect.Height - 1; r >= 0; r--)
                for (int c = rect.Width - 1; c >= 0; c--)
                {
                    int x = rect.Left + c;
                    int y = rect.Top + r;

                    if (y < 0 || x < 0) continue;
                    if (dataGridView1.Rows.Count <= y ||
                       dataGridView1.Columns.Count <= x) continue;

                    String _old = aryCopyBuf[x, y];
                    String _new = dataGridView1[x, y].Value.ToString();

                    // アンドゥ情報の作成
                    UndoOperation undo = new UndoOperation(x, y, _old, _new);

                    // バッファにコピー
                    aryCopyBuf[x, y] = _new;

                    // アンドゥ情報の記録
                    undoGroup.Operations.Add(undo);
                }

            // アンドゥ情報の記録
            undoList.Add(undoGroup);
            // アンドゥメニュ更新
            setUndoText(undoGroup.Title);
            // アンドゥ履歴の調整
            processUndoLimit();
        }

        //---------------------------------------------------------------------------
        private void cutToBuf(Rect rect)
        {
            // アンドゥグループの作成
            UndoGroup undoGroup = new UndoGroup(
                setting.UndoCount++, "切り取り", UndoOperationTarget.MoveToCopyBuffer);

            // 範囲の記録
            undoGroup.Rect = rect;

            //選択範囲をバッファに切り取り
            for (int r = rect.Height - 1; r >= 0; r--)
                for (int c = rect.Width - 1; c >= 0; c--)
                {
                    int x = rect.Left + c;
                    int y = rect.Top + r;
                    if (y < 0 || x < 0) continue;

                    if (dataGridView1.Rows.Count <= y ||
                       dataGridView1.Columns.Count <= x) continue;

                    String _old = aryCopyBuf[x, y];
                    String _new = dataGridView1[x, y].Value.ToString();

                    // アンドゥ情報の作成
                    UndoOperation undo = new UndoOperation(x, y, _old, _new);

                    // 消去前に情報が入っているか確認
                    if (checkCellValue(x, y))
                    {
                        //入っている場合は使用状況を修正
                        aryCellUsedCount[x]--;
                    }

                    // バッファにコピー（※セル内容はクリア）
                    aryCopyBuf[x, y] = _new;
                    dataGridView1[x, y].Value = "";
                    //pIsUsedCell[c+rect.Left-1] = false;

                    // アンドゥ情報の記録
                    undoGroup.Operations.Add(undo);
                }

            // アンドゥ情報の記録
            undoList.Add(undoGroup);
            // アンドゥメニュ更新
            setUndoText(undoGroup.Title);
            // アンドゥ履歴の調整
            processUndoLimit();
        }

        //----------------------------------------------------------------------------------------
        private void deleteRect(Rect rect)
        {
            // アンドゥグループの作成
            UndoGroup undoGroup = new UndoGroup(
                setting.UndoCount++, "消去", UndoOperationTarget.ToGrid);

            // 選択範囲のセル内容を消去
            for (int i = 0; i < rect.Height; i++)
                for (int j = 0; j < rect.Width; j++)
                {
                    // 消去前に情報が入っているか確認
                    if (checkCellValue(rect.X + j, rect.Y + i))
                    {
                        //入っている場合は使用状況を修正
                        aryCellUsedCount[rect.X]--;
                    }

                    // アンドゥ情報の作成
                    UndoOperation undo = new UndoOperation(
                        rect.X + j, rect.Y + i, 
                        dataGridView1[rect.X + j, rect.Y + i].Value.ToString(), "");

                    // 消去
                    dataGridView1[rect.X + j, rect.Y + i].Value = "";

                    // アンドゥ情報の記録
                    undoGroup.Operations.Add(undo);
                }

            // アンドゥ情報の記録
            undoList.Add(undoGroup);
            // アンドゥメニュ更新
            setUndoText(undoGroup.Title);
            // アンドゥ履歴の調整
            processUndoLimit();

            isFirstEdit = true;
        }

        //----------------------------------------------------------------------------------------
        void insertToAllCell(int Row, int Count)
        {
            // 指定セルの指定位置以降を、指定数だけ後ろに送る
            for (int i = 0; i < setting.ColLength; i++)
            {
                for (int rw = (setting.RowLength - 1); rw >= Row + Count; rw--)
                {
                    // セル入力値をコピー
                    dataGridView1[i, rw].Value =
                        dataGridView1[i, rw - Count].Value;

                    // セル色情報のコピー
                    //if(DataGridView1[i, rw].Value.ToString() != "")
                    //{
                    //    gAryBuff[i, rw, gColorBufferNumber] = versionNumber;
                    //}
                }
            }

            // 指定範囲に被る領域を削除（空白にする）
            Rect r = new Rect(0, Row, setting.ColLength, Count);
            deleteRect(r);
        }

        //----------------------------------------------------------------------------------------
        void cutToAllCell(int Row, int Count)
        {
            //全てのセルの指定位置以降を、指定数だけ前に戻す
            for (int i = 0; i < setting.ColLength; i++)
            {
                for (int rw = Row; (rw + Count) < setting.RowLength; rw++)
                {
                    // セル入力値をコピー
                    dataGridView1[i, rw].Value =
                        dataGridView1[i, rw + Count].Value;

                    // セル色情報のコピー
                    //if(dataGridView1[i, rw].Value.ToString() != "")
                    //{
                    //    (*pColorBuf)[i][rw] = versionNumber;
                    //}
                }
            }

            // 範囲末尾の不要領域を削除（空白にする）
            Rect r = new Rect(0, setting.RowLength - Count - 1, setting.ColLength, Count);
            deleteRect(r);

            isFirstEdit = true;

        }

        //----------------------------------------------------------------------------------------
        // @param isInsert = true  : 挿入計算
        // @param isInsert = false : 削除計算
        // @param top              : 対象範囲の先頭
        // @param length           : 対象範囲の長さ
        void calcNakanukiRange(bool isInsert, int top, int length)
        {
            //挿入・削除に合わせて中抜き領域を再計算する
            if (isInsert)
            {
                //挿入動作
                foreach (Range range in delRange)
                {
                    //(処理対象は後ろの範囲)
                    // 各中抜き領域が指定範囲以降か否かのチェック
                    if (range.Top >= top)
                    {
                        range.Top += length;
                        range.Bottom += length;
                    }
                }
            }
            else
            {
                //削除動作
                foreach (Range range in delRange)
                {
                    //(処理対象は後ろの範囲)
                    // 各中抜き領域が指定範囲以降か否かのチェック
                    if (range.Top >= top)
                    {
                        range.Top -= length;
                        range.Bottom -= length;
                    }
                }
            }
        }

        //----------------------------------------------------------------------------------------
        // @param isInsert = true  : 挿入計算
        // @param isInsert = false : 削除計算
        // @param top              : 対象範囲の先頭
        // @param length           : 対象範囲の長さ
        void calcKiribariRange(bool isInsert, int top, int length)
        {
            //挿入・削除に合わせて切り貼り領域を再計算する
            if (isInsert)
            {
                //挿入動作
                foreach (Range range in addRange)
                {
                    //(処理対象は後ろの範囲)
                    // 各切り貼り領域が指定範囲以降か否かのチェック
                    if (range.Top >= top)
                    {
                        range.Top += length;
                        range.Bottom += length;
                    }
                }

            }
            else
            {
                //削除動作
                foreach (Range range in addRange)
                {
                    //(処理対象は後ろの範囲)
                    // 各切り貼り領域が指定範囲以降か否かのチェック
                    if (range.Top >= top)
                    {
                        range.Top -= length;
                        range.Bottom -= length;
                    }
                }
            }
        }

        //---------------------------------------------------------------------------
        private int cursorMoveWithNakaNuki()
        {
            int  cur;           // カーソル位置
            int  range;         // チェック長  : カーソルの次のコマ～最終的にカーソルが移動するコマまで
            int  line;          // 実際の移動量
            int  length;        // 基本の移動量

            //(中抜き範囲補正付き)移動量計算
            cur    = selectRange.Top;
            range  = selectRange.Height;
            line   = 1;
            for(length = 0; length < range; )
            {
              bool bInRange = false;
              if(line > setting.RowLength) return -1;   //範囲を超える場合は終了

              for (int i = 0; i < delRange.Count; i++)
              {
                  Range r = delRange[i];
                  //各中抜き領域内か否かのチェック
                  if (r.Top <= (cur + line) && r.Bottom >= (cur + line))
                  {
                      bInRange = true;   // 領域内
                      break;
                  }
              }
              line++;           //実際の移動量は毎回+1
              if(!bInRange)
              {
                //領域内でない場合のみ、基本の移動量を+1
                length++;
              }
            }
            return (line - 1);  //最終的な移動量を上位に返す
        }

        //----------------------------------------------------------------------------------------
        // フレーム表示文字列の生成
        private String frmToSheet(int frm)
        {
            String Time = "";
            int addFrame = 0;
            int localFrm = 0;
            bool isAddRange = false;

            //切り貼りフレームのカウント
            foreach (Range r in addRange)
            {
                //切り貼りコマ数の積算
                if (r.Top <= frm)
                {
                    addFrame += (r.Bottom - r.Top + 1);
                }

                //切り貼り範囲のチェック
                if (r.Top <= frm && r.Bottom >= frm)
                {
                    //範囲内だったら、切り貼り内ローカルのコマ数を計算
                    // ※また、フレーム表示は１～ なので更に＋１
                    isAddRange = true;
                    localFrm = (frm - r.Top) + 1;
                }
            }

            // カレントフレームに、切り貼り分を加味
            // ※また、フレーム表示は１～ なので更に＋１
            frm = (frm - addFrame) + 1;

            // 24fps
            if (setting.Fps == 24)
            {
                if (isAddRange)
                {
                    // 付けたしフレーム表示処理
                    // フレーム計算(Page + f)
                    int frm_p = 0;  //ページ数は0
                    int frm_f = (localFrm - 1) % (setting.SheetSec * 24);

                    if (((localFrm - 1) % 12) == 0)
                    {
                        if (frm_p < 10) Time = "0" + frm_p.ToString() + '/';
                        else Time = frm_p.ToString() + '/';
                    }
                    else
                    {
                        Time = "     ";
                    }

                    if ((frm_f + 1) < 10) Time += "0" + (frm_f + 1).ToString();
                    else Time += (frm_f + 1).ToString();

                }
                else
                {
                    //通常フレーム表示処理
                    // フレーム計算(Page + f)
                    int frm_p = (frm - 1) / (setting.SheetSec * setting.Fps) + 1;    //ページ数は１から
                    int frm_f = (frm - 1) % (setting.SheetSec * setting.Fps);

                    if (((frm - 1) % 12) == 0)
                    {
                        if (frm_p < 10) Time = "0" + frm_p.ToString() + '/';
                        else Time = frm_p.ToString() + '/';
                    }
                    else
                    {
                        Time = "     ";
                    }

                    if ((frm_f + 1) < 10) Time += "0" + (frm_f + 1).ToString();
                    else Time += (frm_f + 1).ToString();
                }
            }

            // 1 & 30fps
            if ((setting.Fps == 1) || (setting.Fps == 30))
            {
                if (isAddRange)
                {
                    // 付けたしフレーム表示処理
                    // フレーム計算(Page + f)
                    int frm_p = 0;  //ページ数は0
                    int frm_f = (localFrm - 1) % (setting.SheetSec * setting.Fps);

                    if (((localFrm - 1) % 15) == 0)
                    {
                        if (frm_p < 10) Time = "0" + frm_p.ToString() + '/';
                        else Time = frm_p.ToString() + '/';
                    }
                    else
                    {
                        Time = "     ";
                    }

                    if ((frm_f + 1) < 10) Time += "0" + (frm_f + 1).ToString();
                    else Time += (frm_f + 1).ToString();

                }
                else
                {
                    //通常フレーム表示処理

                    // フレーム計算(Page + f)
                    int frm_p = (frm - 1) / (setting.SheetSec * setting.Fps) + 1;    //ページ数は１から
                    int frm_f = (frm - 1) % (setting.SheetSec * setting.Fps);

                    if (((frm - 1) % 15) == 0)
                    {
                        if (frm_p < 10) Time = "0" + frm_p.ToString() + '/';
                        else Time = frm_p.ToString() + '/';
                    }
                    else
                    {
                        Time = "     ";
                    }

                    if ((frm_f + 1) < 10) Time += "0" + (frm_f + 1).ToString();
                    else Time += (frm_f + 1).ToString();
                }
            }

            return Time;
        }

        //----------------------------------------------------------------------------------------
        // 基準線描画位置の計算
        private SheetBorder calcBorderState(DataGridViewCellPaintingEventArgs e)
        {
            SheetBorder retValue = SheetBorder.None;
            bool bHariFirst = false;
            int addFrame = 0;

            // 切り貼りフレーム数のカウント
            {
                // 先頭フレーム前の切り貼り有無
                if(addRange.Count > 0)
                {
                    if(addRange[0].Top == 0)
                    {
                        //有り
                        bHariFirst = true;
                    }
                }

                // 追加のコマ数を数える
                foreach(Range range in addRange)
                {
                    if(range.Top <= e.RowIndex)
                    {
                        addFrame += (range.Bottom - range.Top + 1);
                    }
                }
            }

            // [12],24フレーム毎に基準線を引く
            // (１シート毎に基準線を引く：ページ線)
            int r = e.RowIndex - addFrame + 1;
            if (r < 1) r = e.RowIndex + 1;
            if (setting.Fps == 24 && e.ColumnIndex >= 0)
            {
                int sheetline = setting.SheetDivide;
                if ((((r % sheetline) == 0) && (r != 0)) || ((r == addFrame) && ((e.RowIndex + 1) - addFrame == 0) && (addFrame != 0) && (r != 0)))
                {
                    if ((r % 24) == 0)
                    {
                        //１秒毎の基準線を描画
                        retValue = SheetBorder.EverySec;
                    }
                    else
                    {
                        //ｎコマ毎の基準線を描画
                        retValue = SheetBorder.EveryNFrames;
                    }

                    if ((r % (setting.SheetSec * setting.Fps)) == 0 || ((e.RowIndex + 1) == addFrame) && bHariFirst == true)
                    {
                        //シート毎の基準線を描画
                        retValue = SheetBorder.EverySheet;
                    }
                }
            }

            // 1 & 30fps
            if (((setting.Fps == 1) || (setting.Fps == 30)) && e.ColumnIndex >= 0)
            {
                // 15,30フレーム毎に基準線を引く
                // (１シート毎に基準線を引く：ページ線)
                if ((((r % 15) == 0) && (r != 0)) || ((r == addFrame) && ((e.RowIndex + 1) - addFrame == 0) && (addFrame != 0) && (r != 0)))
                {
                    if ((r % 30) == 0)
                    {
                        //１秒毎の基準線を描画
                        retValue = SheetBorder.EverySec;
                    }
                    else
                    {
                        //ｎコマ毎の基準線を描画
                        retValue = SheetBorder.EveryNFrames;
                    }

                    if ((r % (setting.SheetSec * setting.Fps)) == 0 || ((e.RowIndex + 1) == addFrame) && bHariFirst == true)
                    {
                        //シート毎の基準線を描画
                        retValue = SheetBorder.EverySheet;
                    }
                }
            }
            return retValue;
        }

        //----------------------------------------------------------------------------------------
        // フレーム数 描画
        private void drawFrameNumber(DataGridViewCellPaintingEventArgs e)
        {
            // 背景塗り
            using (Brush backColorBrush = new SolidBrush(gridPalette.Header))
            {
                e.Graphics.FillRectangle(backColorBrush, e.CellBounds);
            }

            //範囲を取得
            Rectangle _rect = e.CellBounds;
            _rect.Inflate(-2, -2);
            //文字列を描画
            String value = "";
            if(setting.IsDisplayFrameNumber)
            {
                // フレーム数 表示
                value = (e.RowIndex + setting.FirstFrame).ToString();
            }
            else
            {
                // シート/コマ数 表示
                value = frmToSheet(e.RowIndex);
            }
            TextRenderer.DrawText(
                e.Graphics,
                value,
                e.CellStyle.Font,
                _rect,
                e.CellStyle.ForeColor,
                TextFormatFlags.Right | TextFormatFlags.VerticalCenter);

            //背景(及びヘッダー部の三角カーソル)以外を描画して貰う
            DataGridViewPaintParts _paintParts =
                e.PaintParts & ~DataGridViewPaintParts.Background;
            //_paintParts &= ~DataGridViewPaintParts.ContentBackground;
            //残りの描画を要求
            e.Paint(e.ClipBounds, _paintParts);

            //描画完了の通知
            e.Handled = true;
        }

        //----------------------------------------------------------------------------------------
        // タイミング継続か否かの確認
        private bool checkContinuty(int X, int Y)
        {
            // 遡って状態を確認
            for (int i = Y; i >= 0; i--)
            {
                // 空白は無視、カラセルを見付けたら falseを返す
                if (dataGridView1[X, i].Value.ToString() == "") continue;
                if (dataGridView1[X, i].Value.ToString() == setting.KaraCell) return false;

                // タイミングの入力を見付けたら trueを返す
                return true;
            }

            return false;
        }

        //----------------------------------------------------------------------------------------
        // 各セルの描画
        private void dataGridView1_CellPainting(object sender, DataGridViewCellPaintingEventArgs e)
        {
            // (仮)フレーム数の表示 [ここから]------------------------------------------------------------
            if (e.ColumnIndex == -1)
            {
                if (e.RowIndex < 0) return; // フレーム数表示列の最上段の場合には、描画処理を一切行わない

                //フレーム数の描画
                // ※フレーム数を表示するセルに対応するセルが存在しない為、TimingCellの描画はできない
                // ※また、setting情報を丸ごとTimingCellの描画に投げるのも微妙なのでとりあえずここで描画
                drawFrameNumber(e);
                return;
            }
            // (仮)フレーム数の表示 [ここまで]------------------------------------------------------------

            // セルの使用状態(未入力/入力済) 評価
            bool bUsed = (aryCellUsedCount[e.ColumnIndex] > 0) ? true : false;

            // アクティブセル(カーソル列か否か) 評価
            bool bActive = (e.ColumnIndex == dataGridView1.CurrentCell.ColumnIndex) ? true : false;

            // 継続記号 評価
            bool bLine = checkContinuty(e.ColumnIndex, e.RowIndex);

            // 切り貼りフレーム数のカウント＆先頭フレームへの切り貼り有無 評価

            // 中抜き 評価
            bool bNuki = false;
            foreach (Range nuki in delRange)
                if (nuki.Top <= e.RowIndex && nuki.Bottom >= e.RowIndex)
                    bNuki = true;

            // 切り貼り 評価
            bool bKiribari = false;
            foreach (Range kiribari in addRange)
                if (kiribari.Top <= e.RowIndex && kiribari.Bottom >= e.RowIndex)
                    bKiribari = true;

            // 背景色の設定
            Color bgColor = new Color();
            if (bNuki)
            {
                // "中抜き"の色設定
                bgColor = gridPalette.Nakanuki;
            }
            else if (bKiribari)
            {
                // "切り貼り"の色設定
                bgColor = gridPalette.Harikomi;
            }
            else if (e.ColumnIndex == -1)
            {
                // "Frames"列の色設定
                bgColor = gridPalette.Header;
            }
            else
            {
                // 各セルの色設定
                switch (dataGridView1.Columns[e.ColumnIndex].DisplayIndex % 2)
                {
                    case 0:  //偶数列
                        // カーソル行の場合はアクティブ色
                        if (bActive) bgColor = gridPalette.BgCell1A;
                        else
                        {
                            // 入力済か否かで、色分け
                            bgColor = (bUsed) ? gridPalette.BgCell1R : gridPalette.BgCell1;
                        }
                        break;
                    case 1:  //奇数列
                        // カーソル行の場合はアクティブ色
                        if (bActive) bgColor = gridPalette.BgCell2A;
                        else
                        {
                            // 入力済か否かで、色分け
                            bgColor = (bUsed) ? gridPalette.BgCell2R : gridPalette.BgCell2;
                        }
                        break;
                }
            }

            // 選択チェック
            bool bSelected = false;
            if ((e.PaintParts & DataGridViewPaintParts.SelectionBackground) ==
                    DataGridViewPaintParts.SelectionBackground &&
                (e.State & DataGridViewElementStates.Selected) ==
                    DataGridViewElementStates.Selected)
            {
                // 選択セル
                bSelected = true;
            }

            // ドラッグ中の選択元 領域チェック
            bool bDragSource = false;
            {
                Rect rect = selectRange;
                if (isRectDrag &&
                    (rect.Top <= e.RowIndex && rect.Bottom >= e.RowIndex &&
                    rect.Left <= e.ColumnIndex && rect.Right >= e.ColumnIndex))
                {
                    // ドラッグ中の選択元セル
                    bDragSource = true;
                }
            }

            // ドラッグ中の移動先 領域チェック
            bool bMovingRange = false;
            if (isRectDrag)
            {
                // 移動先範囲のチェック
                Point offset = new Point(mouseDownPoint.X - selectRange.X, mouseDownPoint.Y - selectRange.Y);
                int col = dataGridView1.CurrentCell.ColumnIndex;
                int row = dataGridView1.CurrentCell.RowIndex;
                int w = selectRange.Width - 1;
                int h = selectRange.Height - 1;
                if (((col - offset.X) > e.ColumnIndex) ||
                    ((col + w - offset.X) < e.ColumnIndex) ||
                    ((row - offset.Y) > e.RowIndex) ||
                    ((row + h - offset.Y) < e.RowIndex))
                {
                }
                else
                {
                    // ドラッグ中の移動先セル
                    bMovingRange = true;
                }
            }

            // ドラッグ中か否かで、選択表示処理を分ける
            if (isRectDrag)
            {
                // ドラッグ時
                if (bMovingRange)
                {
                    // 選択色に設定
                    bgColor = gridPalette.Selected;
                }
                else if (bDragSource)
                {
                    // ドラッグ中選択元色に設定
                    bgColor = Color.DarkGray;
                }
            }
            else
            {
                // 非ドラッグ時
                if (bSelected)
                {
                    // 選択色に設定
                    bgColor = gridPalette.Selected;
                }
            }

            //背景色を設定
            e.CellStyle.BackColor = bgColor;

            //値の取得範囲を制限
            //※ヘッダー部で値取得すると、中身がnullの為に例外が発生する
            if (e.RowIndex >= 0 && e.ColumnIndex >= 0)
            {   // セルのタイミング入力部分
                String str = dataGridView1[e.ColumnIndex, e.RowIndex].Value.ToString();

                //基準線描画位置の計算
                (dataGridView1[e.ColumnIndex, e.RowIndex] as TimingCell).BorderState = calcBorderState(e);

                // カラセル(×印)
                if (str == setting.KaraCell)
                {
                    (dataGridView1[e.ColumnIndex, e.RowIndex] as TimingCell).IsKaraCell = true;
                }
                else
                {
                    (dataGridView1[e.ColumnIndex, e.RowIndex] as TimingCell).IsKaraCell = false;
                }
                // 継続記号
                if (str == "" && bLine)
                {
                    (dataGridView1[e.ColumnIndex, e.RowIndex] as TimingCell).IsContinuousLine = true;
                }
                else
                {
                    (dataGridView1[e.ColumnIndex, e.RowIndex] as TimingCell).IsContinuousLine = false;
                }

            }
            else
            {
                //ヘッダー部分 : 特になにもしない
            }

            //描画を要求
            e.Paint(e.ClipBounds, e.PaintParts);

            //描画完了の通知
            e.Handled = true;

        }

        //----------------------------------------------------------------------------------------
        // 選択範囲の取得
        private Rect getSelectedRect()
        {
            Rect r = new Rect(setting.ColLength,setting.RowLength,0,0);
            int Count = dataGridView1.SelectedCells.Count;

            // 選択範囲を取得
            for (int i = 0; i < Count; i++)
            {
                int rowIndex = dataGridView1.SelectedCells[i].RowIndex;
                int colIndex = dataGridView1.SelectedCells[i].ColumnIndex;
                if (rowIndex < r.Y)
                {
                    r.Y = rowIndex;
                }
                if (colIndex < r.X)
                {
                    r.X = colIndex;
                }
                if (rowIndex > r.Height)
                {
                    r.Height = rowIndex;
                }
                if (colIndex > r.Width)
                {
                    r.Width = colIndex;
                }
            }

            // 高さと幅を計算
            r.Height = r.Height - r.Y + 1;
            r.Width = r.Width - r.X + 1;

            return r;
        }

        //----------------------------------------------------------------------------------------
        // 指定セルの入力有無をチェック
        private bool checkCellValue(int X, int Y)
        {
            bool val = false;

            // チェック範囲は X,Y共に 0以上
            if(X >= 0 && Y >= 0)
            {
                // 値が入っていたら trueを返す
                String str = dataGridView1[X, Y].Value.ToString();
                if (str.Length > 0)
                    val = true;
            }
            return val;
        }

        //----------------------------------------------------------------------------------------
        // 画面2/3より下に移動した場合の画面送り
        private void scrollingForward()
        {
            Rectangle client = dataGridView1.ClientRectangle;               //dataGridViewの表示範囲
            int height = dataGridView1.Rows[0].Height;                      //１行分の高さ
            int rowTop = dataGridView1.FirstDisplayedScrollingRowIndex;     //表示最上段の行数
            int currentRow = dataGridView1.CurrentCell.RowIndex - rowTop;   //表示領域内での上からのインデックス（行）
            int rowCount = (client.Height - (height - 1)) / height;         //表示行数（完全に表示されているもの）
            int border = (int)(rowCount * (2.0 / 3.0));                     //閾値となる行数
            if (border < currentRow)
            {
                //閾値に食い込んだ分、先に進める
                int forwardCount = currentRow - border;
                int top = selectRange.Top + forwardCount;
                int btm = setting.RowLength - selectRange.Height;
                // 先頭が 終端-選択幅(垂直方向) を超える場合は 補正する
                if (top > btm)
                {
                    //はみ出す場合
                    dataGridView1.FirstDisplayedScrollingRowIndex = btm;
                }
                else
                {
                    //はみ出さない場合
                    dataGridView1.FirstDisplayedScrollingRowIndex += forwardCount;
                }

            }
        }

        //----------------------------------------------------------------------------------------
        // KeyDownイベントハンドラ
        private void dataGridView1_KeyDown(object sender, KeyEventArgs e)
        {
            bool isCellEdit = false;
            int keyValue = setting.keys.convKey(e.KeyValue, e.Alt, e.Control, e.Shift);
            switch (keyValue & 0x0ff)
            {
                case 8:     // BackSpace
                    {
                        // 選択範囲を取得
                        Rect rect = getSelectedRect();
                        bool isBackward = false;        // 巻き戻し削除になるかどうかのフラグ(default: 巻き戻しなし)

                        // カレントセルの内容を確認
                        bool isNoBlank = false;
                        for (int i = 0; i < rect.Width; i++)
                        {
                            if (dataGridView1[rect.X + i, rect.Y].Value.ToString() != "")
                            {
                                isNoBlank = true;
                                break;
                            }
                        }
                        if(isNoBlank)
                        {
                            // 現在セルに値が入っている場合
                            // （なにもしない）

                        }
                        else
                        {
                            // 現在セルに値が入っていない場合
                            for (int i = rect.Y; i >= 0 && !isBackward; i--)
                                for (int l = rect.Left; l <= rect.Right; l++)
                                {
                                    // 空のセルは無視
                                    String str = dataGridView1[l, i].Value.ToString();
                                    if (str == "") continue;

                                    // 選択範囲をクリア
                                    dataGridView1.ClearSelection();

                                    // カーソルのカレント位置を設定
                                    dataGridView1.CurrentCell = dataGridView1[dataGridView1.CurrentCell.ColumnIndex, i];

                                    // 新しい選択範囲を設定
                                    for (int j = 0; j < rect.Height; j++)
                                        for (int k = 0; k < rect.Width; k++)
                                            dataGridView1[rect.X + k, i + j].Selected = true;

                                    // 範囲を保存
                                    rect.Y = i;
                                    selectRange = rect;

                                    // フラグを立てる
                                    isBackward = true;
                                    isNoBlank = true;

                                    // 探索を終わる
                                    break;
                                }
                        }

                        if (isNoBlank)
                        {
                            // 現在セルに値が入っている場合

                            // アンドゥグループの作成
                            UndoGroup undoGroup = new UndoGroup(
                                setting.UndoCount++, "削除", UndoOperationTarget.ToGrid);
                            undoList.Add(undoGroup);

                            for (int i = rect.Left; i <= rect.Right; i++)
                            {
                                // 空白セルは無視する
                                String str = dataGridView1[i, rect.Y].Value.ToString();
                                if (str.Length == 0) continue;

                                // アンドゥ情報の作成
                                UndoOperation undo = new UndoOperation(i, rect.Y, str, "");

                                if (isBackward)
                                {
                                    //セル内容の消去
                                    dataGridView1[i, rect.Y].Value = "";
                                    //使用状況を修正
                                    aryCellUsedCount[i]--;
                                }
                                else
                                {
                                    //１桁削る
                                    str = str.Substring(0, str.Length - 1);
                                    dataGridView1[i, rect.Y].Value = str;
                                    isCellEdit = true;

                                    // セルの中身が空白になった場合は、使用状況を修正
                                    if (str.Length == 0)
                                    {
                                        aryCellUsedCount[i]--;
                                    }
                                }

                                // アンドゥ情報の記録
                                undo.NewValue = str;
                                undoGroup.Operations.Add(undo);

                            }

                            // アンドゥメニュ更新
                            setUndoText(undoGroup.Title);
                            // アンドゥ履歴の調整
                            processUndoLimit();
                        }

                        // 描画更新(継続記号の更新の為)
                        dataGridView1.Invalidate();
                        
                        // 処理済みフラグを立てる
                        e.Handled = true;
                    }
                    break;

                case 13:    // Enter
                    {
                        // 選択範囲を取得
                        Rect rect = getSelectedRect();

                        // 選択範囲をクリア
                        dataGridView1.ClearSelection();

                        // 先頭位置の計算(※下端のはみ出しチェック)
                        int len = cursorMoveWithNakaNuki();
                        if (len > 0)
                        {
                            int top = rect.Y + len;
                            int btm = setting.RowLength - rect.Height;
                            top = (top > btm) ? btm : top;  // 先頭が 終端-選択幅(垂直方向) を超える場合は 補正する

                            // カーソルのカレント位置を設定
                            dataGridView1.CurrentCell = dataGridView1[rect.X, top];

                            // 新しい選択範囲を設定
                            for (int i = 0; i < rect.Height; i++)
                                for (int j = 0; j < rect.Width; j++)
                                    dataGridView1[rect.X + j, top + i].Selected = true;

                            // 範囲を保存
                            rect.Y = top;
                            selectRange = rect;
                        }

                        isFirstEdit = true;

                        // 画面2/3より下に移動した場合の画面送り
                        scrollingForward();

                        // 処理済みフラグを立てる
                        e.Handled = true;
                    }
                    break;

                case 33:    // PageUp
                    {
                        // 基準秒数を足し合わせて新しい行数を求める
                        int moveSize;
                        if (setting.keys.checkShiftBeforeConvertion(keyValue, CombinationKeyState.CTRLKey))
                        {
                            // +Ctrl
                            moveSize = setting.Fps * setting.SheetSec;
                        }
                        else
                        {
                            // その他
                            moveSize = setting.Fps;
                        }

                        int row = dataGridView1.FirstDisplayedScrollingRowIndex;
                        int newRow = row - moveSize;

                        // 先頭行を基準フレームにしたいので余りを差し引く
                        if ((newRow % moveSize) != 0)
                        {
                            newRow += moveSize - (newRow % moveSize);
                        }

                        // 移動可能な範囲ならページ移動
                        // ※カーソル位置はページ先頭
                        if (newRow >= 0)
                        {
                            dataGridView1.FirstDisplayedScrollingRowIndex = newRow;
                            dataGridView1.CurrentCell = dataGridView1[dataGridView1.CurrentCell.ColumnIndex, newRow];
                        }

                        // 処理済みフラグを立てる
                        e.Handled = true;
                    }
                    break;

                case 34:    // PageDown
                    {
                        // 基準秒数を足し合わせて新しい行数を求める
                        int moveSize;
                        if (setting.keys.checkShiftBeforeConvertion(keyValue, CombinationKeyState.CTRLKey))
                        {
                            // +Ctrl
                            moveSize = setting.Fps * setting.SheetSec;
                        }
                        else
                        {
                            // その他
                            moveSize = setting.Fps;
                        }

                        int row = dataGridView1.FirstDisplayedScrollingRowIndex;
                        int newRow = row + moveSize;

                        // 先頭行を基準フレームにしたいので余りを差し引く
                        newRow -= newRow % moveSize;

                        // 移動可能な範囲ならページ移動
                        // ※カーソル位置はページ先頭
                        if (newRow < setting.RowLength)
                        {
                            dataGridView1.FirstDisplayedScrollingRowIndex = newRow;
                            dataGridView1.CurrentCell = dataGridView1[dataGridView1.CurrentCell.ColumnIndex, newRow];
                        }

                        // 処理済みフラグを立てる
                        e.Handled = true;
                    }
                    break;

                case 36:    // Home
                    {
                        // カーソルを先頭フレームに戻す
                        dataGridView1.CurrentCell = dataGridView1[dataGridView1.CurrentCell.ColumnIndex, 0];
                        selectRange.X = dataGridView1.CurrentCell.ColumnIndex;
                        selectRange.Y = 0;
                        selectRange.Width = 1;
                        selectRange.Height = 1;

                        // 処理済みフラグを立てる
                        e.Handled = true;
                    }
                    break;

                case 37:    // ←
                    // 選択範囲を縮小
                    if (setting.keys.checkShiftBeforeConvertion(keyValue, CombinationKeyState.SHIFTKey))
                    {
                        // +Shiftの場合

                        // 選択範囲を取得
                        Rect rect = getSelectedRect();

                        // 領域縮小が可能な場合のみ処理する
                        if (rect.Width > 1)
                        {
                            // 範囲を縮小
                            for (int i = 0; i < rect.Height; i++)
                                dataGridView1[rect.Right, rect.Y + i].Selected = false;

                            // 範囲を保存
                            rect.Width--;
                            selectRange = rect;
                        }

                    }

                    // 選択範囲単位の移動
                    if (setting.keys.checkShiftBeforeConvertion(keyValue, CombinationKeyState.None))
                    {
                        // 単に ← の場合

                        // 選択範囲を取得
                        Rect rect = getSelectedRect();

                        // 選択範囲をクリア
                        dataGridView1.ClearSelection();

                        // 先頭位置の計算(※上端のはみ出しチェック)
                        int top = rect.X - 1;
                        top = (top < 0) ? 0 : top;      // 先頭が 0を下回る場合は 補正する

                        // カーソルのカレント位置を設定
                        dataGridView1.CurrentCell = dataGridView1[top, rect.Y];

                        // 新しい選択範囲を設定
                        for (int i = 0; i < rect.Height; i++)
                            for (int j = 0; j < rect.Width; j++)
                                dataGridView1[top + j, rect.Y + i].Selected = true;

                        // 範囲を保存
                        rect.X = top;
                        selectRange = rect;

                        // 描画更新(アクティブセルの色分けの為)
                        dataGridView1.Invalidate();
                    }
                    // 処理済みフラグを立てる
                    e.Handled = true;
                    break;

                case 38:    //↑
                    // 選択範囲を縮小
                    if (setting.keys.checkShiftBeforeConvertion(keyValue, CombinationKeyState.SHIFTKey))
                    {
                        // +SHIFT の場合

                        // 選択範囲を取得
                        Rect rect = getSelectedRect();

                        // 領域縮小が可能な場合のみ処理する
                        if (rect.Height > 1)
                        {
                            // 範囲を縮小
                            for (int j = 0; j < rect.Width; j++)
                                dataGridView1[rect.X + j, rect.Bottom].Selected = false;

                            // 範囲を保存
                            rect.Height--;
                            selectRange = rect;
                        }

                    }

                    // 選択範囲単位の移動
                    {
                        CombinationKeyState state = setting.keys.getShiftBeforeConvertion(keyValue);
                        if ((state & CombinationKeyState.ALTKey) == 0 &&
                            (state & CombinationKeyState.SHIFTKey) == 0)
                        {
                            // ↑ 又は ALT+↑ の場合

                            // 選択範囲を取得
                            Rect rect = getSelectedRect();

                            // 選択範囲をクリア
                            dataGridView1.ClearSelection();

                            // 先頭位置の計算(※上端のはみ出しチェック)
                            int top;
                            if((state & CombinationKeyState.CTRLKey) != 0)
                            {
                                // +Ctrl時
                                top = rect.Y - 1;
                            }
                            else
                            {
                                // ↑のみ
                                top = rect.Y - rect.Height;
                            }
                            top = (top < 0) ? 0 : top;      // 先頭が 0を下回る場合は 補正する

                            // カーソルのカレント位置を設定
                            dataGridView1.CurrentCell = dataGridView1[rect.X, top];

                            // 新しい選択範囲を設定
                            for (int i = 0; i < rect.Height; i++)
                                for (int j = 0; j < rect.Width; j++)
                                    dataGridView1[rect.X + j, top + i].Selected = true;

                            // 範囲を保存
                            rect.Y = top;
                            selectRange = rect;

                        }
                    }

                    // 処理済みフラグを立てる
                    e.Handled = true;
                    break;

                case 39:    // →
                    // 選択範囲を拡大
                    if (setting.keys.checkShiftBeforeConvertion(keyValue, CombinationKeyState.SHIFTKey))
                    {
                        // + SHIFT の場合

                        // 選択範囲を取得
                        Rect rect = getSelectedRect();

                        // 領域拡大が可能な場合のみ処理する
                        if ((rect.X + rect.Width) < setting.ColLength)
                        {
                            // 範囲を拡大
                            rect.Width++;
                            for (int i = 0; i < rect.Height; i++)
                                dataGridView1[rect.Right, rect.Y + i].Selected = true;

                            // 範囲を保存
                            selectRange = rect;
                        }

                    }

                    // 選択範囲単位の移動
                    if (setting.keys.checkShiftBeforeConvertion(keyValue, CombinationKeyState.None))
                    {
                        // 単に → の場合

                        // 選択範囲を取得
                        Rect rect = getSelectedRect();

                        // 選択範囲をクリア
                        dataGridView1.ClearSelection();

                        // 先頭位置の計算(※上端のはみ出しチェック)
                        int top = rect.X + 1;
                        int limit = setting.ColLength - rect.Width;
                        top = (top > limit) ? limit : top;      // 先頭が 終端-選択幅(水平方向) を超える場合は 補正する

                        // カーソルのカレント位置を設定
                        dataGridView1.CurrentCell = dataGridView1[top, rect.Y];

                        // 新しい選択範囲を設定
                        for (int i = 0; i < rect.Height; i++)
                            for (int j = 0; j < rect.Width; j++)
                                dataGridView1[top + j, rect.Y + i].Selected = true;

                        // 範囲を保存
                        rect.X = top;
                        selectRange = rect;

                        // 描画更新(アクティブセルの色分けの為)
                        dataGridView1.Invalidate();
                    }
                    // 処理済みフラグを立てる
                    e.Handled = true;
                    break;

                case 40:    // ↓
                    // 選択範囲を拡大
                    if (setting.keys.checkShiftBeforeConvertion(keyValue, CombinationKeyState.SHIFTKey))
                    {
                        // +SHIFT の場合

                        // 選択範囲を取得
                        Rect rect = getSelectedRect();

                        // 領域拡大が可能な場合のみ処理する
                        if ((rect.Y + rect.Height) < setting.RowLength)
                        {
                            // 範囲を拡大
                            rect.Height++;
                            for (int j = 0; j < rect.Width; j++)
                                dataGridView1[rect.X + j, rect.Bottom].Selected = true;

                            // 範囲を保存
                            selectRange = rect;
                        }

                    }

                    // 選択範囲単位の移動
                    {
                        CombinationKeyState state = setting.keys.getShiftBeforeConvertion(keyValue);
                        if ((state & CombinationKeyState.ALTKey) == 0 &&
                            (state & CombinationKeyState.SHIFTKey) == 0)
                        {
                            // ↓ 又は CTRL+↓ の場合

                            // 選択範囲を取得
                            Rect rect = getSelectedRect();

                            // 選択範囲をクリア
                            dataGridView1.ClearSelection();

                            // 先頭位置の計算(※下端のはみ出しチェック)
                            int top;
                            if ((state & CombinationKeyState.CTRLKey) != 0)
                            {
                                // +Ctrl時
                                top = rect.Y + 1;
                            }
                            else
                            {
                                // ↓のみ
                                top = rect.Y + rect.Height;
                            }
                            int limit = setting.RowLength - rect.Height;
                            top = (top > limit) ? limit : top;  // 先頭が 終端-選択幅(垂直方向) を超える場合は 補正する

                            // カーソルのカレント位置を設定
                            dataGridView1.CurrentCell = dataGridView1[rect.X, top];

                            // 新しい選択範囲を設定
                            for (int i = 0; i < rect.Height; i++)
                                for (int j = 0; j < rect.Width; j++)
                                    dataGridView1[rect.X + j, top + i].Selected = true;

                            // 範囲を保存
                            rect.Y = top;
                            selectRange = rect;

                            // 画面2/3より下に移動した場合の画面送り
                            scrollingForward();
                        }
                    }
                    // 処理済みフラグを立てる
                    e.Handled = true;
                    break;

                case 45:    // Insert
                    {
                        // 範囲の追加
                        insertToAllCell(selectRange.Top, selectRange.Height);
                        calcNakanukiRange(true, selectRange.Top, selectRange.Height);
                        calcKiribariRange(true, selectRange.Top, selectRange.Height);

                        // アンドゥ履歴をフラッシュ
                        flushUndoList();

                        isFirstEdit = true;

                        // 描画更新(切り貼り範囲反映の為)
                        dataGridView1.Invalidate();

                        // 処理済みフラグを立てる
                        e.Handled = true;
                    }
                    break;

                case 46:    // Delete
                    {
                        if (setting.keys.checkShiftBeforeConvertion(keyValue, CombinationKeyState.SHIFTKey))
                        {
                            // +Shiftの場合

                            // 範囲を削除
                            cutToAllCell(selectRange.Top, selectRange.Height);
                            calcNakanukiRange(false, selectRange.Top, selectRange.Height);
                            calcKiribariRange(false, selectRange.Top, selectRange.Height);

                            // アンドゥ履歴をフラッシュ
                            flushUndoList();

                            isFirstEdit = true;

                            // 描画更新(切り貼り範囲反映の為)
                            dataGridView1.Invalidate();

                            // 処理済みフラグを立てる
                            e.Handled = true;
                        }
                        else if (setting.keys.checkShiftBeforeConvertion(keyValue, CombinationKeyState.None))
                        {
                            // 選択範囲のセル内容を消去
                            deleteRect(getSelectedRect());

                            isFirstEdit = true;

                            // 処理済みフラグを立てる
                            e.Handled = true;

                            // 描画更新(継続記号の更新の為)
                            dataGridView1.Invalidate();
                        }
                    }
                    break;

                case 48:    // 0-9(Full-key)
                case 49:
                case 50:
                case 51:
                case 52:
                case 53:
                case 54:
                case 55:
                case 56:
                case 57:

                case 96:    // 0-9(10key)
                case 97:
                case 98:
                case 99:
                case 100:
                case 101:
                case 102:
                case 103:
                case 104:
                case 105:
                    {
                        // フルキーのコードはテンキーコードに補正
                        int key = (e.KeyValue < 96) ? e.KeyValue + 48 : e.KeyValue;

                        // 入力値を計算
                        byte[] ch = { (byte)key };
                        ch[0] += (byte)'0';
                        ch[0] -= 96;

                        // アンドゥグループの作成
                        UndoGroup undoGroup = new UndoGroup(
                            setting.UndoCount++, "入力", UndoOperationTarget.ToGrid);
                        undoList.Add(undoGroup);

                        Rect rect = selectRange;
                        for(int i = 0; i < rect.Width; i++)
                        {
                            // アンドゥ情報の作成
                            UndoOperation undo = new UndoOperation(
                                rect.X + i, rect.Y, dataGridView1[rect.X + i, rect.Y].Value.ToString(), "");

                            // 入力前に情報が入っているか確認
                            if (!checkCellValue(rect.X + i, rect.Y))
                            {
                                // 空白の場合は使用状況を修正
                                aryCellUsedCount[rect.X + i]++;
                            }

                            // 初回フラグが立ち, 尚且つ "常に追加"が未チェックの場合のみ 初回編集を上書きにする
                            if (isFirstEdit && !this.alwaysAppendToolStripMenuItem.Checked)
                            {
                                // セルに値を設定(初回編集)
                                dataGridView1[rect.X + i, rect.Y].Value = System.Text.Encoding.GetEncoding(932).GetString(ch);
                            }
                            else
                            {
                                // セルに値を設定(継続編集)
                                dataGridView1[rect.X + i, rect.Y].Value += System.Text.Encoding.GetEncoding(932).GetString(ch);
                            }
                            isCellEdit = true;

                            // アンドゥ情報の記録
                            undo.NewValue = dataGridView1[rect.X + i, rect.Y].Value.ToString();
                            undoGroup.Operations.Add(undo);
                        }
                        // アンドゥメニュ更新
                        setUndoText(undoGroup.Title);
                        // アンドゥ履歴の調整
                        processUndoLimit();

                        // 処理済みフラグを立てる
                        e.Handled = true;

                        // 描画更新(継続記号の更新の為)
                        dataGridView1.Invalidate();
                    }
                    break;

                case 74:    // J
                case 75:    // K
                    {
                        // キーの頭出し
                        int value = selectRange.Top;
                        int col, row;
                        if (e.KeyValue == 74)
                        {
                            // 後方検索
                            for (row = selectRange.Top - 1; row >= 0 && value == selectRange.Top; row--)
                            {
                                for (col = selectRange.Left; col <= selectRange.Right; col++)
                                {
                                    if (dataGridView1[col, row].Value.ToString() == "") continue;
                                    value = row;
                                    break;
                                }
                            }
                        }
                        if (e.KeyValue == 75)
                        {
                            // 前方検索
                            for (row = selectRange.Bottom + 1; row < setting.RowLength && value == selectRange.Top; row++)
                            {
                                for (col = selectRange.Left; col <= selectRange.Right; col++)
                                {
                                    if (dataGridView1[col, row].Value.ToString() == "") continue;
                                    value = row;
                                    break;
                                }
                            }
                        }

                        //選択範囲を修正
                        selectRange.Y = value;
                        dataGridView1.CurrentCell = dataGridView1[dataGridView1.CurrentCell.ColumnIndex, selectRange.Y];

                        //選択範囲を解除
                        dataGridView1.ClearSelection();

                        //新しい範囲に設定
                        for(int i = 0; i < selectRange.Height; i++)
                            for (int j = 0; j < selectRange.Width; j++)
                            {
                                dataGridView1[selectRange.X + j, selectRange.Y + i].Selected = true;
                            }

                        //表示位置の調整
                        //dataGridView1.FirstDisplayedScrollingRowIndex = selectRange.Y  - (selectRange.Y % setting.Fps);

                        // 画面2/3より下に移動した場合の画面送り
                        scrollingForward();

                        // 処理済みフラグを立てる
                        e.Handled = true;
                    }
                    break;

                case 106:   // '*'(10key)
                    // 選択範囲を拡大
                    {
                        // 選択範囲を取得
                        Rect rect = getSelectedRect();

                        // 領域拡大が可能な場合のみ処理する
                        if ((rect.Y + rect.Height) < setting.RowLength)
                        {
                            // 範囲を拡大
                            rect.Height++;
                            for (int j = 0; j < rect.Width; j++)
                                dataGridView1[rect.X + j, rect.Bottom].Selected = true;

                            // 範囲を保存
                            selectRange = rect;
                        }

                        // 処理済みフラグを立てる
                        e.Handled = true;
                    }
                    break;

                case 107:   // '+'(10key)
                    {
                        // 入力前に情報が入っているか確認
                        Rect rect = selectRange;
                        if (!checkCellValue(rect.X, rect.Y))
                        {
                            // 空白の場合は使用状況を修正
                            aryCellUsedCount[rect.X]++;
                        }

                        // 手前の入力を検索
                        int key = 1;
                        for (int i = rect.Y - 1; i >= 0; i--)
                        {
                            // 空白と空セルは無視
                            String str = dataGridView1[rect.X, i].Value.ToString();
                            //if (str == "" || str == setting.KaraCell) continue;
                            if (str == "" ) continue;
                            if (str == setting.KaraCell) return;    // カラセルが入力されていたら、処理を中断

                            // 値を見付けたら、取得後＋１して検索を終わる
                            key = int.Parse(str);
                            key++;
                            break;
                        }

                        // アンドゥグループの作成
                        UndoGroup undoGroup = new UndoGroup(
                            setting.UndoCount++, "入力", UndoOperationTarget.ToGrid);

                        // アンドゥ情報の作成
                        UndoOperation undo = new UndoOperation(
                            rect.X, rect.Y, dataGridView1[rect.X, rect.Y].Value.ToString(), "");

                        // セルに値を設定
                        dataGridView1[rect.X, rect.Y].Value = key.ToString();

                        // アンドゥ情報の記録
                        undo.NewValue = dataGridView1[rect.X, rect.Y].Value.ToString();
                        undoGroup.Operations.Add(undo);
                        undoList.Add(undoGroup);
                        // アンドゥメニュ更新
                        setUndoText(undoGroup.Title);
                        // アンドゥ履歴の調整
                        processUndoLimit();

                        // 選択範囲をクリア
                        dataGridView1.ClearSelection();

                        // 先頭位置の計算(※下端のはみ出しチェック)
                        int len = cursorMoveWithNakaNuki();
                        if (len > 0)
                        {
                            int top = rect.Y + len;
                            int limit = setting.RowLength - rect.Height;
                            top = (top > limit) ? limit : top;  // 先頭が 終端-選択幅(垂直方向) を超える場合は 補正する

                            // カーソルのカレント位置を設定
                            dataGridView1.CurrentCell = dataGridView1[rect.X, top];

                            // 新しい選択範囲を設定
                            for (int i = 0; i < rect.Height; i++)
                                for (int j = 0; j < rect.Width; j++)
                                    dataGridView1[rect.X + j, top + i].Selected = true;

                            // 範囲を保存
                            rect.Y = top;
                            selectRange = rect;
                        }

                        isFirstEdit = true;

                        // 画面2/3より下に移動した場合の画面送り
                        scrollingForward();

                        // 処理済みフラグを立てる
                        e.Handled = true;

                        // 描画更新(継続記号の更新の為)
                        dataGridView1.Invalidate();
                    }
                    break;

                case 109:   // '-'(10key)
                    {
                        // 入力前に情報が入っているか確認
                        Rect rect = selectRange;
                        if (!checkCellValue(rect.X, rect.Y))
                        {
                            // 空白の場合は使用状況を修正
                            aryCellUsedCount[rect.X]++;
                        }

                        // 手前の入力を検索
                        int key = 0;
                        for (int i = rect.Y - 1; i >= 0; i--)
                        {
                            // 空白と空セルは無視
                            String str = dataGridView1[rect.X, i].Value.ToString();
                            //if (str == "" || str == setting.KaraCell) continue;
                            if (str == "") continue;
                            if (str == setting.KaraCell) return;    // カラセルが入力されていたら、処理を中断

                            // 値を見付けたら、取得後－１して検索を終わる
                            key = int.Parse(str);
                            key = (key - 1) > 0 ? key-1 : key;  // １以下にならないようにしておく
                            break;
                        }

                        // アンドゥグループの作成
                        UndoGroup undoGroup = new UndoGroup(
                            setting.UndoCount++, "入力", UndoOperationTarget.ToGrid);

                        // アンドゥ情報の作成
                        UndoOperation undo = new UndoOperation(
                            rect.X, rect.Y, dataGridView1[rect.X, rect.Y].Value.ToString(), "");

                        // セルに値を設定
                        dataGridView1[rect.X, rect.Y].Value = key.ToString();

                        // アンドゥ情報の記録
                        undo.NewValue = dataGridView1[rect.X, rect.Y].Value.ToString();
                        undoGroup.Operations.Add(undo);
                        undoList.Add(undoGroup);
                        // アンドゥメニュ更新
                        setUndoText(undoGroup.Title);
                        // アンドゥ履歴の調整
                        processUndoLimit();

                        // 選択範囲をクリア
                        dataGridView1.ClearSelection();

                        // 先頭位置の計算(※下端のはみ出しチェック)
                        int len = cursorMoveWithNakaNuki();
                        if (len > 0)
                        {
                            int top = rect.Y + len;
                            int limit = setting.RowLength - rect.Height;
                            top = (top > limit) ? limit : top;  // 先頭が 終端-選択幅(垂直方向) を超える場合は 補正する

                            // カーソルのカレント位置を設定
                            dataGridView1.CurrentCell = dataGridView1[rect.X, top];

                            // 新しい選択範囲を設定
                            for (int i = 0; i < rect.Height; i++)
                                for (int j = 0; j < rect.Width; j++)
                                    dataGridView1[rect.X + j, top + i].Selected = true;

                            // 範囲を保存
                            rect.Y = top;
                            selectRange = rect;
                        }

                        isFirstEdit = true;

                        // 画面2/3より下に移動した場合の画面送り
                        scrollingForward();

                        // 処理済みフラグを立てる
                        e.Handled = true;

                        // 描画更新(継続記号の更新の為)
                        dataGridView1.Invalidate();
                    }
                    break;

                case 111:   // '/'(10key)
                    // 選択範囲を縮小
                    {
                        // 選択範囲を取得
                        Rect rect = getSelectedRect();

                        // 領域縮小が可能な場合のみ処理する
                        if (rect.Height > 1)
                        {
                            // 範囲を縮小
                            for (int j = 0; j < rect.Width; j++)
                                dataGridView1[rect.X + j, rect.Bottom].Selected = false;

                            // 範囲を保存
                            rect.Height--;
                            selectRange = rect;
                        }

                        // 処理済みフラグを立てる
                        e.Handled = true;
                    }
                    break;

                case 110:   // '.'(10key)
                case 190:   // '.'(full-key)
                    {
                        // アンドゥグループの作成
                        UndoGroup undoGroup = new UndoGroup(
                            setting.UndoCount++, "入力", UndoOperationTarget.ToGrid);
                        undoList.Add(undoGroup);

                        // 値の入力
                        Rect rect = selectRange;
                        for (int i = rect.Left; i <= rect.Right; i++)
                        {
                            // アンドゥ情報の作成
                            UndoOperation undo = new UndoOperation(
                                rect.X, rect.Y, dataGridView1[i, rect.Y].Value.ToString(), "");

                            // 入力前に情報が入っているか確認
                            if (!checkCellValue(i, rect.Y))
                            {
                                // 空白の場合は使用状況を修正
                                aryCellUsedCount[i]++;
                            }

                            // セルに "カラセルの値"を設定
                            dataGridView1[i, rect.Y].Value = setting.KaraCell;

                            // アンドゥ情報の記録
                            undo.NewValue = dataGridView1[i, rect.Y].Value.ToString();
                            undoGroup.Operations.Add(undo);
                        }
                        // アンドゥメニュ更新
                        setUndoText(undoGroup.Title);
                        // アンドゥ履歴の調整
                        processUndoLimit();

                        // 設定によりカーソル移動
                        int len = cursorMoveWithNakaNuki();
                        if (!setting.IsKaraNoMove && len > 0)
                        {
                            // ※基本的に Enterキー動作と同じ

                            // 選択範囲を取得
                            rect = selectRange;

                            // 選択範囲をクリア
                            dataGridView1.ClearSelection();

                            // 先頭位置の計算(※下端のはみ出しチェック)
                            int top = rect.Y + len;
                            int btm = setting.RowLength - rect.Height;
                            top = (top > btm) ? btm : top;  // 先頭が 終端-選択幅(垂直方向) を超える場合は 補正する

                            // カーソルのカレント位置を設定
                            dataGridView1.CurrentCell = dataGridView1[rect.X, top];

                            // 新しい選択範囲を設定
                            for (int i = 0; i < rect.Height; i++)
                                for (int j = 0; j < rect.Width; j++)
                                    dataGridView1[rect.X + j, top + i].Selected = true;

                            // 範囲を保存
                            rect.Y = top;
                            selectRange = rect;

                            // 画面2/3より下に移動した場合の画面送り
                            scrollingForward();
                        }

                        isFirstEdit = true;

                        // 処理済みフラグを立てる
                        e.Handled = true;

                        // 描画更新(継続記号の更新の為)
                        dataGridView1.Invalidate();
                    }
                    break;
            }

            //初期編集状態の設定/解除
            isFirstEdit = isCellEdit ? false : true;

            return;
        }

        //----------------------------------------------------------------------------------------
        // KeyPressイベントハンドラ
        private void dataGridView1_KeyPress(object sender, KeyPressEventArgs e)
        {
            //
            return;
        }

        //----------------------------------------------------------------------------------------
        // KeyUpイベントハンドラ
        private void dataGridView1_KeyUp(object sender, KeyEventArgs e)
        {
            //
            return;
        }

        //----------------------------------------------------------------------------------------
        // MouseDownイベントハンドラ
        private void dataGridView1_CellMouseDown(object sender, DataGridViewCellMouseEventArgs e)
        {
            // 左クリックかチェック
            if ((e.Button & System.Windows.Forms.MouseButtons.Left) != 0)
            {
                // 選択範囲内かチェック
                Rect rect = selectRange;
                int col = e.ColumnIndex;
                int row = e.RowIndex;
                if (rect.Top > row || rect.Bottom < row || rect.Left > col || rect.Right < col)
                {
                    // 初期状態に戻す
                    mouseDownPoint = new Point(-1, -1);
                    isRectDrag = false;
                }
                else if (rect.Top <= row && rect.Bottom >= row && rect.Left <= col && rect.Right >= col)
                {
                    // カレントの位置を取得
                    mouseDownPoint = new Point(col, row);
                    isRectDrag = true;
                }
            }

            // 描画更新(アクティブセルの色分けの為)
            dataGridView1.Invalidate();

            return;
        }

        //----------------------------------------------------------------------------------------
        // MouseMoveイベントハンドラ
        private void dataGridView1_CellMouseMove(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (isRectDrag)
            {
                // 描画更新(範囲描画のため)
                dataGridView1.Invalidate();
            }
        }

        //----------------------------------------------------------------------------------------
        // MouseUpイベントハンドラ
        private void dataGridView1_CellMouseUp(object sender, DataGridViewCellMouseEventArgs e)
        {
            // ドラッグ中だった場合は 選択範囲を修正し移動（又はコピー）処理を行う
            if (isRectDrag)
            {
                //選択範囲を解除
                dataGridView1.ClearSelection();

                //カレントセルは選択しておく
                dataGridView1.CurrentCell.Selected = true;

                if ((Control.ModifierKeys & Keys.Control) != 0)
                {
                    //選択元をコピー＆ペースト
                    int col = dataGridView1.CurrentCell.ColumnIndex - (mouseDownPoint.X - selectRange.X);
                    int row = dataGridView1.CurrentCell.RowIndex - (mouseDownPoint.Y - selectRange.Y);
                    copyToBuf(selectRange);
                    copyToCell(col, row, selectRange);
                }
                else
                {
                    //選択元をカット＆ペースト
                    int col = dataGridView1.CurrentCell.ColumnIndex - (mouseDownPoint.X - selectRange.X);
                    int row = dataGridView1.CurrentCell.RowIndex - (mouseDownPoint.Y - selectRange.Y);
                    cutToBuf(selectRange);
                    copyToCell(col, row, selectRange);
                }
            }
            else
            {
                // Ctrlキー同時押しの抑止
                if ((Control.ModifierKeys & Keys.Control) != 0)
                {
                    //MessageBox.Show("Ctrlキー同時押しによるマルチセレクト機能には未対応です。");

                    // 選択領域を解除 (カレントフレームのみ選択する)
                    dataGridView1.ClearSelection();
                    dataGridView1.CurrentCell.Selected = true;
                }
            }

            // 初期状態に戻す
            mouseDownPoint = new Point(-1, -1);
            isRectDrag = false;
            isFirstEdit = true;

            // 選択範囲を保存
            selectRange = getSelectedRect();

            // 描画更新(範囲描画のため)
            dataGridView1.Invalidate();

            return;
        }

        //----------------------------------------------------------------------------------------
        int textBoxColumn = 0;
        private void TextBox_Terminate()
        {
            // 自分自身をdataGridViewから外す
            textBox1.Visible = false;
        }

        //----------------------------------------------------------------------------------------
        private void TextBox_KeyDown(object sender, KeyEventArgs e)
        {
            //キー入力があったら、内容をチェック
            switch(e.KeyCode)
            {
                case Keys.Enter:
                    // 入力確定
                    dataGridView1.Columns[textBoxColumn].HeaderText = textBox1.Text;
                    TextBox_Terminate();
                    e.Handled = true;
                    break;
                case Keys.Escape:
                    // 入力キャンセル
                    TextBox_Terminate();
                    e.Handled = true;
                    break;
            }
        }

        //----------------------------------------------------------------------------------------
        private void TextBox_Leave(object sender, EventArgs e)
        {
            // テキストボックスからフォーカスが外れた場合、自分自身をdataGridViewから外す
            TextBox_Terminate();

        }

        //----------------------------------------------------------------------------------------
        private void dataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            // セルがダブルクリックされた
            if (e.ColumnIndex >= 0 && e.RowIndex == -1)
            {
                // ヘッダの場合
                textBox1.Text = dataGridView1.Columns[e.ColumnIndex].HeaderText;
                textBoxColumn = e.ColumnIndex;
                textBox1.Left = (e.ColumnIndex + 1) * 50 + 1;
                textBox1.Top = 3;
                textBox1.Width = 50;
                textBox1.Height = 24;
                textBox1.Visible = true;
                textBox1.Focus();

                isFirstEdit = true;
            }
            
        }

        //----------------------------------------------------------------------------------------
        private void SetClipboardTextWithRetry(string text, int maxRetries = 5, int delayMs = 100)
        {
            for (int i = 0; i < maxRetries; i++)
            {
                try
                {
                    Clipboard.SetText(text);
                    return;
                }
                catch (ExternalException)
                {
                    if (i == maxRetries - 1) throw; // 最後のリトライで失敗した場合、例外を再スロー
                    Thread.Sleep(delayMs);
                }
            }
        }

        //----------------------------------------------------------------------------------------
        private bool IsClipboardAvailable()
        {
            try
            {
                Clipboard.GetDataObject();
                return true;
            }
            catch (ExternalException)
            {
                return false;
            }
        }

        //----------------------------------------------------------------------------------------
        private void AECopy(bool isDirect)
        {
            // AEへコピー
            String Copytext = "";
            int rlen = 0;

            // カレントの列を取得
            int col = dataGridView1.CurrentCell.ColumnIndex;

            //ヘッダ書き出し(定型文字列)
            Copytext = "Adobe After Effects " + setting.AfterRemapVersion + " Keyframe Data\r\n";
            Copytext += "\r\n";
            Copytext += "\tUnits Per Second\t" + setting.Fps.ToString() + "\r\n";
            Copytext += "\tSource Width\t640\r\n";
            Copytext += "\tSource Height\t480\r\n";
            Copytext += "\tSource Pixel Aspect Ratio\t1\r\n";
            Copytext += "\tComp Pixel Aspect Ratio\t1\r\n";
            Copytext += "\r\n";
            Copytext += "Time Remap\r\n";
            Copytext += "\tFrame\tseconds\r\n";

            // タイミングの書き出し
            for (int i = 0; i < setting.RowLength; i++)
            {
                // 中抜き範囲は無視
                bool bNuki = false;
                foreach (Range r in delRange)
                {
                    if ((r.Top <= i) && (r.Bottom >= i))
                    {
                        //中抜き範囲時は、範囲長に積算する
                        rlen++;
                        bNuki = true;
                        break;
                    }
                }
                //中抜き範囲内は無視
                if (bNuki) continue;

                // カラセルは無視
                String str = dataGridView1[col, i].Value.ToString();
                if (str == "") continue;

                // 中ヌキ範囲長でフレーム値を補正
                int fcnt = i - rlen;

                // 秒数計算
                double t = 0;
                if (!isDirect)
                {
                    // 通常のリマップ計算
                    t = (double)int.Parse(str) - setting.FirstFrame;
                    t /= (double)setting.Fps;
                }
                else
                {
                    // 入力値のまま渡す
                    t = (double)int.Parse(str);
                }

                // 計算結果を文字列に収める(タイミングの秒数は、有効桁数を６桁に丸める)
                Copytext += "\t" + fcnt.ToString() + "\t" + t.ToString("g6") + "\r\n";

            }

            //フッタ書き出し(定型文字列)
            Copytext += "\r\n";
            Copytext += "\r\n";
            Copytext += "End of Keyframe Data\r\n";

            //クリップボードに反映
            SetClipboardTextWithRetry(Copytext);
            if (!IsClipboardAvailable())
            {
                MessageBox.Show("クリップボードにアクセスできません。");
            }
        }

        //----------------------------------------------------------------------------------------
        private void AECopyToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // AEへコピー
            AECopy(false);

        }

        //----------------------------------------------------------------------------------------
        private void directRemapToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // AEへコピー(TimeRemap以外)
            AECopy(true);

        }

        //----------------------------------------------------------------------------------------
        private void AECopyWithScript(bool isDirect)
        {
            // AEへコピー(Script仲介)
            // ※処理的には、AEへコピーと同じ。
            // ※違うのは４点 : 書式, fcntを 0から数える、タイミングの秒数計算時のfpsで割る処理の省略, AfterEffects呼び出し
            String Copytext = "";
            int rlen = 0;

            // カレントの列を取得
            int col = dataGridView1.CurrentCell.ColumnIndex;

            //ヘッダ書き出し(定型文字列)
            Copytext = "{property:'Time Remap',scale:" + setting.Fps.ToString("f") + ",keys:[";

            //タイミングの書き出し
            for (int i = 0; i < setting.RowLength; i++)
            {
                // 中抜き範囲は無視
                bool bNuki = false;
                foreach (Range r in delRange)
                {
                    if ((r.Top <= i) && (r.Bottom >= i))
                    {
                        //中抜き範囲時は、範囲長に積算する
                        rlen++;
                        bNuki = true;
                        break;
                    }
                }
                //中抜き範囲内は無視
                if (bNuki) continue;

                // カラセルは無視
                String str = dataGridView1[col, i].Value.ToString();
                if (str == "") continue;

                // 中ヌキ範囲長でフレーム値を補正
                int fcnt = i - rlen;

                // 秒数計算
                double t = 0;
                if (!isDirect)
                {
                    // 通常のリマップ計算(※スクリプトに渡す場合は fps計算せず スクリプト側に一任)
                    t = (double)int.Parse(str) - setting.FirstFrame;
                }
                else
                {
                    // 入力値のまま渡す
                    t = (double)int.Parse(str);
                }

                // 計算結果を文字列に収める
                Copytext += "{'t':"+fcnt.ToString()+",'v':["+t.ToString()+"]},";

            }

            //フッタ書き出し(定型文字列)
            Copytext += "]}";

            //クリップボードに反映
            SetClipboardTextWithRetry(Copytext);
            if (!IsClipboardAvailable())
            {
                MessageBox.Show("クリップボードにアクセスできません。");
            }

            //AfterEffects側の呼び出し
            Process.Start(setting.AfterPath, "-r " + setting.AfterOption);
        }

        //----------------------------------------------------------------------------------------
        private void jSRemapToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // AEへコピー(Script仲介)
            AECopyWithScript(false);

        }

        //----------------------------------------------------------------------------------------
        private void pasteFromAEToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // AEからペースト
            String[] clip = System.Text.RegularExpressions.Regex.Split(Clipboard.GetText(), "\r\n");
            int i = 0;

            // カレントの列を取得
            int col = dataGridView1.CurrentCell.ColumnIndex;

            // ヘッダチェック
            if (clip[i].IndexOf("Adobe After Effects ") < 0 ||
                clip[i].IndexOf("Keyframe Data") < 0)
            {
                MessageBox.Show("ヘッダ : 未対応のヘッダ.");
                return;
            }

            // fpsのチェック
            i += 2;
            {
                String[] buf = clip[i].Split('\t');
                if (clip[i].IndexOf("\tUnits Per Second") < 0 ||
                    buf.Length < 3)
                {
                    MessageBox.Show("Units Per Second : 目的の情報が見つからない.");
                    return;
                }
                setting.Fps = int.Parse(buf[2]);

                // fps変更をメニューに反映
                switch (setting.Fps)
                {
                    case 24:
                        this.fPS30ToolStripMenuItem.Checked = false;
                        this.fPS24ToolStripMenuItem.Checked = true;
                        break;
                    case 30:
                        this.fPS30ToolStripMenuItem.Checked = true;
                        this.fPS24ToolStripMenuItem.Checked = false;
                        break;
                    default:
                        this.fPS30ToolStripMenuItem.Checked = false;
                        this.fPS24ToolStripMenuItem.Checked = false;
                        break;
                }

            }

            // "TimeRemap"で始まる行を探す
            for (; i < clip.Length; i++)
            {
                if (clip[i].IndexOf("Time Remap") != -1) break;
            }

            // リマップ情報を読む
            i+=2;
            for (; i < clip.Length; i++)
            {
                String[] buf = clip[i].Split('\t');
                if (buf.Length <= 1) break;

                int frm = int.Parse(buf[1]);
                double val = double.Parse(buf[2]);
                //int t = (int)((double)setting.Fps * val);
                //if ((((double)setting.Fps * val) - ((double)t)) >= 0.5) t += 1; // コマ数の四捨五入
                int t = (int)Math.Round(setting.Fps * val);

                // タイミング情報をセルに書き込む
                // ※書き込むセルが空欄の場合は、使用カウントを＋１
                DataGridViewCell cell = dataGridView1[col, frm];
                if (cell.Value.ToString().Length == 0)
                {
                    aryCellUsedCount[col]++;
                }
                cell.Value = (t + setting.FirstFrame).ToString();
            }

            isFirstEdit = true;

            // 描画更新(継続記号の更新の為)
            dataGridView1.Invalidate();

            // アンドゥ履歴をフラッシュ
            flushUndoList();
        }

        //----------------------------------------------------------------------------------------
        private void setNakanukiToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //中抜き範囲を設定

            // 範囲の追加
            Rect rect = getSelectedRect();
            Range range = new Range(rect.Y, rect.Bottom);
            delRange.Add(range);

            isFirstEdit = true;

            // 描画更新(中抜き範囲反映の為)
            dataGridView1.Invalidate();
        }

        //----------------------------------------------------------------------------------------
        private void setKiribariToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //切り貼り範囲を設定

            // 範囲が重なっている場合は 登録中の領域を削除
            Rect rect = getSelectedRect();
            int t = rect.Top;
            int b = rect.Bottom;
            for (int i = 0; i < addRange.Count; i++ )
            {
                Range r = addRange[i];
                if (t <= r.Top && r.Top <= b && r.Bottom >= t)
                {
                    // 切り貼り範囲を削除
                    cutToAllCell(r.Top, r.Bottom - r.Top + 1);
                    calcNakanukiRange(false, r.Top, r.Bottom - r.Top + 1);
                    calcKiribariRange(false, r.Top, r.Bottom - r.Top + 1);

                    addRange.RemoveAt(i--);
                }
                else if (t >= r.Top && t <= r.Bottom && b >= r.Top)
                {
                    // 切り貼り範囲を削除
                    cutToAllCell(r.Top, r.Length);
                    calcNakanukiRange(false, r.Top, r.Length);
                    calcKiribariRange(false, r.Top, r.Length);

                    addRange.RemoveAt(i--);
                }
            }

            // 切り貼り範囲に空白を挿入
            insertToAllCell(rect.Top, rect.Height);
            calcNakanukiRange(true, rect.Top, rect.Height);
            calcKiribariRange(true, rect.Top, rect.Height);

            // 範囲の追加
            {
                Range range = new Range(rect.Top, rect.Bottom);
                addRange.Add(range);
            }

            // アンドゥ履歴をフラッシュ
            flushUndoList();

            isFirstEdit = true;

            // 描画更新(切り貼り範囲反映の為)
            dataGridView1.Invalidate();
        }

        //----------------------------------------------------------------------------------------
        private void cancelNakanukiToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //中抜き範囲の解除

            // 登録範囲の検索
            Rect rect = getSelectedRect();
            foreach(Range r in delRange)
            {
                if (r.Top <= rect.Y && r.Bottom >= rect.Y)
                {
                    delRange.Remove(r);
                    break;
                }
            }

            isFirstEdit = true;

            // 描画更新(中抜き範囲反映の為)
            dataGridView1.Invalidate();
        }

        //----------------------------------------------------------------------------------------
        private void cancelKiribariToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //切り貼り範囲の解除

            // 登録範囲の検索
            Rect rect = getSelectedRect();
            foreach (Range r in addRange)
            {
                if (r.Top <= rect.Y && r.Bottom >= rect.Y)
                {
                    // 切り貼り範囲を削除
                    cutToAllCell(r.Top, r.Length);
                    calcNakanukiRange(false, r.Top, r.Length);
                    calcKiribariRange(false, r.Top, r.Length);

                    addRange.Remove(r);
                    break;
                }
            }

            // アンドゥ履歴をフラッシュ
            flushUndoList();

            isFirstEdit = true;

            // 描画更新(切り貼り範囲反映の為)
            dataGridView1.Invalidate();
        }

        //----------------------------------------------------------------------------------------
        private void undoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // 操作を元に戻す
            undoFunction();

            isFirstEdit = true;
        }

        //----------------------------------------------------------------------------------------
        private void setUndoText(String text)
        {
            // アンドゥ メニュー表示の更新
            if (text.Length > 0)
            {
                undoToolStripMenuItem.Text = text + "を元に戻す";
                undoToolStripMenuItem.Enabled = true;
            }
            else
            {
                undoToolStripMenuItem.Text = "元に戻す";
                undoToolStripMenuItem.Enabled = false;
            }
        }

        //----------------------------------------------------------------------------------------
        private void processUndoLimit()
        {
            int total = 0;
            for (int i = undoList.Count - 1; i >= 0; i--)
            {
                // 登録操作の合計を数える（新しい方から順に）
                total += undoList[i].Operations.Count;

                // limitを超えたら、古い順に履歴の削除を行う
                if (total > setting.UndoLimitCount)
                {
                    // 削除するのは先頭から現在のインデックスまで
                    int last = i;
                    for (int j = 0; j <= last; j++)
                    {
                        undoList.RemoveAt(0);
                    }
                    // 削除が終わったら終了
                    break;
                }
            }
        }

        //----------------------------------------------------------------------------------------
        private void flushUndoList()
        {
            // アンドゥ非対応機能を使用した場合などに アンドゥ履歴をフラッシュする
            undoList.Clear();
            // メニュの表示も初期状態に戻す
            setUndoText("");
        }

        //----------------------------------------------------------------------------------------
        private void copyToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // コピー
            copyRect = getSelectedRect();
            copyToBuf(copyRect);
            isFirstEdit = true;
        }

        //----------------------------------------------------------------------------------------
        private void cutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // 切り取り
            copyRect = getSelectedRect();
            cutToBuf(copyRect);
            isFirstEdit = true;

            // 描画更新(継続記号の更新の為)
            dataGridView1.Invalidate();
        }

        //----------------------------------------------------------------------------------------
        private void pasteToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // 貼り付け
            int col = dataGridView1.CurrentCell.ColumnIndex;
            int row = dataGridView1.CurrentCell.RowIndex;
            copyToCell(col, row, copyRect);
            isFirstEdit = true;

            // 描画更新(継続記号の更新の為)
            dataGridView1.Invalidate();
        }

        //----------------------------------------------------------------------------------------
        private void fPS30ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // 30FPSに変更
            setting.Fps = 30;
            this.fPS30ToolStripMenuItem.Checked = true;
            this.fPS24ToolStripMenuItem.Checked = false;

            // 描画更新(フレーム表示、基準線更新の為)
            dataGridView1.Invalidate();
        }

        //----------------------------------------------------------------------------------------
        private void fPS24ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // 24FPSに変更
            setting.Fps = 24;
            this.fPS30ToolStripMenuItem.Checked = false;
            this.fPS24ToolStripMenuItem.Checked = true;

            // 描画更新(フレーム表示、基準線更新の為)
            dataGridView1.Invalidate();
        }

        //----------------------------------------------------------------------------------------
        private void stayOnTopToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // TopMost設定の切り替え
            this.TopMost = !this.TopMost;
            setting.TopMost = this.TopMost;
            stayOnTopToolStripMenuItem.Checked = this.TopMost;
        }

        //----------------------------------------------------------------------------------------
        private void karacellNoMoveToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // カラセル入力時の移動抑止設定の切り替え
            setting.IsKaraNoMove = !setting.IsKaraNoMove;
            karacellNoMoveToolStripMenuItem.Checked = setting.IsKaraNoMove;
        }

        //----------------------------------------------------------------------------------------
        private void displayFrameNumberToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // フレーム数表示⇔シート/コマ数表示 切り替え
            setting.IsDisplayFrameNumber = !setting.IsDisplayFrameNumber;
            this.displayFrameNumberToolStripMenuItem.Checked = setting.IsDisplayFrameNumber;

            // 描画更新(フレーム表示更新の為)
            dataGridView1.Invalidate();
        }

        //----------------------------------------------------------------------------------------
        private void firstFrameToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // 開始フレーム設定
            InputBox1 dialog = new InputBox1("開始フレーム数");
            dialog.LabelName1 = "開始フレーム数";
            dialog.Value1 = setting.FirstFrame.ToString();
            if (dialog.ShowDialog(this.owner) == System.Windows.Forms.DialogResult.OK)
            {
                try
                {
                    setting.FirstFrame = int.Parse(dialog.Value1);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("入力された値を数値に変換できませんでした.");
                }
            }
        }

        //----------------------------------------------------------------------------------------
        private void karacellValueToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // カラセル文字列の設定
            InputBox1 dialog = new InputBox1("カラセル文字列");
            dialog.LabelName1 = "カラセル文字列";
            dialog.Value1 = setting.KaraCell;
            if (dialog.ShowDialog(this.owner) == System.Windows.Forms.DialogResult.OK)
            {
                setting.KaraCell = dialog.Value1;
            }
        }

        //----------------------------------------------------------------------------------------
        private void secondsPerSheetToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // シートの秒数
            InputBox1 dialog = new InputBox1("シートの秒数");
            dialog.LabelName1 = "シートの秒数";
            dialog.Value1 = setting.SheetSec.ToString();
            if (dialog.ShowDialog(this.owner) == System.Windows.Forms.DialogResult.OK)
            {
                try
                {
                    setting.SheetSec = int.Parse(dialog.Value1);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("入力された値を数値に変換できませんでした.");
                }
            }
        }

        //----------------------------------------------------------------------------------------
        private void div4ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // シートの基準線指定: 4コマ毎
            setting.SheetDivide = 4;
            this.div4ToolStripMenuItem.Checked = true;
            this.div6ToolStripMenuItem.Checked = false;
            this.div12ToolStripMenuItem.Checked = false;

        }

        //----------------------------------------------------------------------------------------
        private void div6ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // シートの基準線指定: 6コマ毎
            setting.SheetDivide = 6;
            this.div4ToolStripMenuItem.Checked = false;
            this.div6ToolStripMenuItem.Checked = true;
            this.div12ToolStripMenuItem.Checked = false;

        }

        //----------------------------------------------------------------------------------------
        private void div12ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // シートの基準線指定: 12コマ毎
            setting.SheetDivide = 12;
            this.div4ToolStripMenuItem.Checked = false;
            this.div6ToolStripMenuItem.Checked = false;
            this.div12ToolStripMenuItem.Checked = true;

        }

        //----------------------------------------------------------------------------------------
        private void afterFXPathToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // AfterFX.exe パス設定
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.FileName = @"AfterFX.exe";
            dialog.InitialDirectory = @"C:\Program Files\Adobe\Adobe After Effects CS5\Support Files\";
            dialog.Filter = "実行ファイル(*.exe)|*.exe";
            dialog.FilterIndex = 0;
            dialog.Title = "起動する AfterFX.exeファイルを選択してください";
            dialog.RestoreDirectory = false;
            dialog.CheckFileExists = true;
            dialog.CheckPathExists = true;
            dialog.Multiselect = false;

            //ダイアログを表示する
            if (dialog.ShowDialog(this.owner) == DialogResult.OK)
            {
                //OKボタンがクリックされた場合は、設定更新
                setting.AfterPath = dialog.FileName;
            }

        }

        //----------------------------------------------------------------------------------------
        private void afterFXOptionToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // リマップ用 jsx パス設定
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.FileName = @"setRemap.jsx";
            dialog.InitialDirectory = @"C:\Program Files\Adobe\Adobe After Effects CS5\Support Files\Scripts\";
            dialog.Filter = "スクリプトファイル(*.jsx)|*.jsx";
            dialog.FilterIndex = 0;
            dialog.Title = "呼び出すスクリプトファイルを選択してください";
            dialog.RestoreDirectory = false;
            dialog.CheckFileExists = true;
            dialog.CheckPathExists = true;
            dialog.Multiselect = false;

            //ダイアログを表示する
            if (dialog.ShowDialog(this.owner) == DialogResult.OK)
            {
                //OKボタンがクリックされた場合は、設定更新
                setting.AfterOption = dialog.FileName;
            }
        }

        //----------------------------------------------------------------------------------------
        private void autoAdjustToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // セル変更時にウィンドウ位置調整する 設定
            setting.IsAutoadjust = !setting.IsAutoadjust;
            autoAdjustToolStripMenuItem.Checked = setting.IsAutoadjust;
        }

        //----------------------------------------------------------------------------------------
        private void allInitializeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // 作業情報の全初期化
            InitializeWork(true);
        }

        //----------------------------------------------------------------------------------------
        private void inputCellCountToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // セル枚数の指定
            InputBox1 dialog = new InputBox1("セル枚数の指定");
            int newCount = 0;
            dialog.LabelName1 = "セル枚数";
            dialog.Value1 = setting.ColLength.ToString();
            if (dialog.ShowDialog(this.owner) == System.Windows.Forms.DialogResult.OK)
            {
                try
                {
                    newCount = int.Parse(dialog.Value1);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("入力された値を数値に変換できませんでした.");
                }
            }
            
            // 指定の枚数に調整
            // ※入力されているセル情報は保持（削減されたセルの情報は消滅）
            if (newCount > 0 && newCount <= setting.CellCountLimit && setting.ColLength != newCount)
            {
                // グリッドサイズを変更
                resizeDataGridView1(newCount,setting.RowLength);

                // ウィンドウ位置調整
                adjustWindowSize();
            }

            // アンドゥ履歴をフラッシュ
            flushUndoList();

            // コピーバッファをフラッシュ
            InitializeWork(InitializeTarget.CopyBuffer);
        }

        //----------------------------------------------------------------------------------------
        private void insertCellToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // セルの挿入

            // カレントセルの位置を保存
            int col = dataGridView1.CurrentCell.ColumnIndex;

            // グリッドサイズを変更
            resizeDataGridView1(setting.ColLength + 1, setting.RowLength);

            // ウィンドウ位置調整
            adjustWindowSize();

            // カレントセルの位置を空ける
            {
                // カレントセル位置を空けるように位置をずらす
                int firstColIndex = (setting.ColLength - 1) - 1;
                for (int i = firstColIndex; i >= col; i--)
                {
                    aryCellUsedCount[i + 1] = aryCellUsedCount[i];
                    dataGridView1.Columns[i + 1].HeaderText = dataGridView1.Columns[i].HeaderText;
                    for (int j = 0; j < setting.RowLength; j++)
                        dataGridView1[i + 1, j].Value = dataGridView1[i, j].Value.ToString();
                }
                // 開いた場所を空欄にする
                aryCellUsedCount[col] = 0;
                dataGridView1.Columns[col].HeaderText = "";
                for (int i = 0; i < setting.RowLength; i++)
                    dataGridView1[col, i].Value = "";
            }

            // アンドゥ履歴をフラッシュ
            flushUndoList();

            // コピーバッファをフラッシュ
            InitializeWork(InitializeTarget.CopyBuffer);
        }

        //----------------------------------------------------------------------------------------
        private void deleteCellToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // セルの削除

            // カレントセルの位置を詰める
            int col = dataGridView1.CurrentCell.ColumnIndex;
            {
                // カレントセル位置を埋めるように位置をずらす
                for (int i = col; i < setting.ColLength - 1; i++)
                {
                    aryCellUsedCount[i] = aryCellUsedCount[i + 1];
                    dataGridView1.Columns[i].HeaderText = dataGridView1.Columns[i + 1].HeaderText;
                    for (int j = 0; j < setting.RowLength; j++)
                        dataGridView1[i, j].Value = dataGridView1[i + 1, j].Value.ToString();
                }
            }

            // グリッドサイズを変更
            resizeDataGridView1(setting.ColLength - 1, setting.RowLength);

            // ウィンドウ位置調整
            adjustWindowSize();

            // アンドゥ履歴をフラッシュ
            flushUndoList();

            // コピーバッファをフラッシュ
            InitializeWork(InitializeTarget.CopyBuffer);
        }

        //----------------------------------------------------------------------------------------
        private void sequentialNumberToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // 連番作成
            InputBox1 dialog = new InputBox1("連番作成");
            dialog.LabelName1 = "開始番号";
            dialog.LabelName2 = "ステップ数";
            dialog.CheckName1 = "番号を飛ばす";
            dialog.Value1 = "1";
            dialog.Value2 = "1";
            if (dialog.ShowDialog(this.owner) == System.Windows.Forms.DialogResult.OK)
            {
                int start = 1;
                int step = 1;
                bool skip = false;
                try
                {
                    start = int.Parse(dialog.Value1);
                    step = int.Parse(dialog.Value2);
                    skip = dialog.CheckValue1;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("入力された値を数値に変換できませんでした.");
                    return;
                }

                int col = selectRange.Left;
                int row = selectRange.Top;
                int frm = start;
                int frmStep = (skip == true) ? step : step / Math.Abs(step);
                int cnt = selectRange.Height;

                // アンドゥグループの作成
                UndoGroup undoGroup = new UndoGroup(
                    setting.UndoCount++, "連番作成", UndoOperationTarget.ToGrid);

                // 入力
                for (int i = 0; i < cnt; i += Math.Abs(step))
                {
                    // アンドゥ情報の作成
                    UndoOperation undo = new UndoOperation(
                        col, row + i, dataGridView1[col, row + i].Value.ToString(), frm.ToString());

                    if (dataGridView1[col, row + i].Value.ToString() == "")
                    {
                        aryCellUsedCount[col]++;
                    }
                    dataGridView1[col, row + i].Value = frm.ToString();

                    // アンドゥ情報の記録
                    undoGroup.Operations.Add(undo);

                    //(*pColorBuf)[Col][Row + 1] = versionNumber;
                    frm += frmStep;
                }

                // アンドゥ情報の記録
                undoList.Add(undoGroup);
                // アンドゥメニュ更新
                setUndoText(undoGroup.Title);
                // アンドゥ履歴の調整
                processUndoLimit();

                isFirstEdit = true;

                // 描画更新(継続記号の更新の為)
                dataGridView1.Invalidate();

            }
        }

        //----------------------------------------------------------------------------------------
        private void repeatNumberToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // 繰り返し
            RepeatInputBox dialog = new RepeatInputBox();
            if (dialog.ShowDialog(this.owner) == System.Windows.Forms.DialogResult.OK)
            {
                // 入力チェック
                // (空白時は中止)
                if (dialog.Value1 == "")
                { MessageBox.Show("開始番号がない."); return; }
                if (dialog.Value2 == "")
                { MessageBox.Show("終了番号がない."); return; }
                if (dialog.Value3 == "")
                { MessageBox.Show("コマ打ち数がない."); return; }
                if (dialog.Value4 == "")
                { MessageBox.Show("繰り返し数がない."); return; }
                if (dialog.Value5 == "")
                { MessageBox.Show("スキップ数がない."); return; }
                try
                {
                    if (dialog.Value6.Length > 0)
                    {
                        int num = int.Parse(dialog.Value6);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("入力された値を数値に変換できませんでした."); return;
                }

                int count, start, end, step, loop, skip, insert;
                try
                {
                    start = int.Parse(dialog.Value1);
                    end = int.Parse(dialog.Value2);
                    step = int.Parse(dialog.Value3);
                    loop = int.Parse(dialog.Value4);
                    skip = int.Parse(dialog.Value5) + 1;
                    insert = 1;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("入力された値を数値に変換できませんでした."); return;
                }

                // 入力数: 開始#～終了#
                // ※挿入番号がある場合、開始～終了番号と交互に入るので更に２倍
                // ※スキップ数が有る場合、１／スキップ数に回数を減らす
                // ※ループ回数は１回分の長さが決定したところで計算
                count = end - start + 1;
                if (dialog.Value6 != "")
                {
                    insert = 2;
                    count *= 2;
                }
                if (skip > 1)
                {
                    count /= skip;
                    count++;
                    //挿入番号があり、カウントが奇数の場合は偶数に補正
                    if (insert > 1 && (count % 2) == 1)
                    {
                        count++;
                    }
                }
                count *= loop;

                //範囲の消去
                {
                    deleteRect(selectRange);
                }

                // アンドゥグループの作成
                UndoGroup undoGroup = new UndoGroup(
                    setting.UndoCount++, "繰り返し", UndoOperationTarget.ToGrid);

                //番号入力
                int row = selectRange.Top;
                for (int num = start, col = selectRange.Left; col <= selectRange.Right; col++)
                {
                    for (int i = 0; i < count; i++)
                    {
                        // アンドゥ情報の作成
                        UndoOperation undo = new UndoOperation(
                            col, row + (i * step),
                            dataGridView1[col, row + (i * step)].Value.ToString(), num.ToString());

                        if (insert == 1)
                        {
                            //挿入番号なし
                            dataGridView1[col, row + (i * step)].Value = num.ToString();
                            aryCellUsedCount[col]++;
                            //(*pColorBuf)[Col][Row + (i * step)] = versionNumber;
                            num += skip;
                            if (num > end) num = start;

                            // アンドゥ情報の記録
                            undoGroup.Operations.Add(undo);
                        }
                        else if ((i % 2) == 0)
                        {
                            //挿入番号あり（開始＃～終了＃）
                            //※連番と挿入番号が交互なのでカウンタを1/2にして番号計算
                            dataGridView1[col, row + (i * step)].Value = num.ToString();
                            aryCellUsedCount[col]++;
                            //(*pColorBuf)[Col][Row + (i * step)] = versionNumber;
                            num += skip;
                            if (num > end) num = start;

                            // アンドゥ情報の記録
                            undoGroup.Operations.Add(undo);
                        }
                        else
                        {
                            //挿入番号あり（挿入＃）
                            dataGridView1[col, row + (i * step)].Value = dialog.Value6;
                            aryCellUsedCount[col]++;
                            //(*pColorBuf)[Col][Row + (i * step)] = versionNumber;

                            // アンドゥ情報の記録
                            undoGroup.Operations.Add(undo);
                        }
                    }
                }

                // アンドゥ情報の記録
                undoList.Add(undoGroup);
                // アンドゥメニュ更新
                setUndoText(undoGroup.Title);
                // アンドゥ履歴の調整
                processUndoLimit();

                isFirstEdit = true;

                // 描画更新(継続記号の更新の為)
                dataGridView1.Invalidate();
            }
        }

        //----------------------------------------------------------------------------------------
        private void replaceToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //置換
            InputBox1 dialog = new InputBox1("置換");
            dialog.LabelName1 = "置換前";
            dialog.LabelName2 = "置換後";
            if (dialog.ShowDialog(this.owner) == System.Windows.Forms.DialogResult.OK)
            {
                // 入力チェック
                // (空白時は中止)
                if (dialog.Value1 == "")
                { MessageBox.Show("変換前指定がない."); return; }
                if (dialog.Value2 == "")
                { MessageBox.Show("変換後指定がない."); return; }

                int Col, Row, Cnt;
                String A = dialog.Value1;
                String B = dialog.Value2;
                Col = selectRange.Left;
                Row = selectRange.Top;
                Cnt = selectRange.Height;

                // アンドゥグループの作成
                UndoGroup undoGroup = new UndoGroup(
                    setting.UndoCount++, "置換", UndoOperationTarget.ToGrid);

                // 置き換え
                for (int i = 0; i < Cnt; i++)
                {
                    if (dataGridView1[Col, Row + i].Value.ToString() == "") continue;

                    // フレーム毎に値を調べて、変換前を見つけたら、変換後に書き換え
                    if (dataGridView1[Col, Row + i].Value.ToString() == A)
                    {
                        // アンドゥ情報の作成
                        UndoOperation undo = new UndoOperation(
                            Col, Row + i,
                            dataGridView1[Col, Row + i].Value.ToString(), B);

                        if (dataGridView1[Col, Row + i].Value.ToString() == "")
                        {
                            aryCellUsedCount[Col]++;
                        }
                        dataGridView1[Col, Row + i].Value = B;
                        //(*pColorBuf)[Col][Row + i] = versionNumber;

                        // アンドゥ情報の記録
                        undoGroup.Operations.Add(undo);
                    }

                    // アンドゥ情報の記録
                    undoList.Add(undoGroup);
                    // アンドゥメニュ更新
                    setUndoText(undoGroup.Title);
                    // アンドゥ履歴の調整
                    processUndoLimit();

                    isFirstEdit = true;

                    // 描画更新(継続記号の更新の為)
                    dataGridView1.Invalidate();
                }
            }
        }

        //----------------------------------------------------------------------------------------
        private void reverseToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //反転
            int i, c, l, Col, Row, Cnt;
            Col = selectRange.Left;
            Row = selectRange.Top;
            Cnt = selectRange.Height;

            //選択範囲内の記述を取得
            int targetCount = 0;
            String[] temp = new String[Cnt];
            for(i = 0; i < Cnt; i++)
            {
                String val = dataGridView1[Col, Row + i].Value.ToString();
                if(val == "") continue;
                temp[targetCount++] = val;
            }

            // アンドゥグループの作成
            UndoGroup undoGroup = new UndoGroup(
                setting.UndoCount++, "反転", UndoOperationTarget.ToGrid);

            //選択範囲内の記述を逆順に適応
            for (i = 0; i < Cnt; i++)
            {
                if (dataGridView1[Col, Row + i].Value.ToString() == "") continue;

                // アンドゥ情報の作成
                UndoOperation undo = new UndoOperation(
                    Col, Row + i,
                    dataGridView1[Col, Row + i].Value.ToString(), "");

                dataGridView1[Col, Row + i].Value = temp[--targetCount];
                undo.NewValue = temp[targetCount];
                //(*pColorBuf)[Col][Row + i] = versionNumber;

                // アンドゥ情報の記録
                undoGroup.Operations.Add(undo);
            }

            // アンドゥ情報の記録
            undoList.Add(undoGroup);
            // アンドゥメニュ更新
            setUndoText(undoGroup.Title);
            // アンドゥ履歴の調整
            processUndoLimit();

            isFirstEdit = true;

            // 描画更新(継続記号の更新の為)
            dataGridView1.Invalidate();
        }

        //----------------------------------------------------------------------------------------
        private void fourArithmeticOperationToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //四則演算

            InputBox1 dialog = new InputBox1("四則演算");
            dialog.LabelName1 = "四則演算";
            if (dialog.ShowDialog(this.owner) == System.Windows.Forms.DialogResult.OK)
            {
                String val = dialog.Value1.Trim();
                CalcMode mode = CalcMode.None;

                int len = val.Length;
                char mark = (char)(val[0]);
                switch (mark)
                {
                    case '+':
                        mode = CalcMode.Plus;
                        break;
                    case '-':
                        mode = CalcMode.Minus;
                        break;
                    case '*':
                        mode = CalcMode.Multiple;
                        break;
                    case '/':
                        mode = CalcMode.Divide;
                        break;
                    default:
                        //エラー
                        //１文字目は加減乗除記号が必要
                        MessageBox.Show("１文字目には\"+-*/\"記号のいずれか１文字の入力が必要");
                        return;
                        //break;
                }

                int num = 0;
                try
                {
                    num = int.Parse(val.Substring(1, len - 1));
                }
                catch (Exception ex)
                {
                    MessageBox.Show("入力された値を数値に変換できませんでした."); return;
                }

                // アンドゥグループの作成
                UndoGroup undoGroup = new UndoGroup(
                    setting.UndoCount++, "四則演算", UndoOperationTarget.ToGrid);

                // アンドゥ情報の記録
                undoList.Add(undoGroup);

                Rect r = selectRange;
                for (int c = r.Left; c <= r.Right; c++)
                {
                    for (int i = r.Top; i <= r.Bottom; i++)
                    {
                        if (dataGridView1[c,i].Value.ToString() == "" ||
                           dataGridView1[c,i].Value.ToString() == setting.KaraCell) continue;

                        int celNum = 0;
                        try
                        {
                            celNum = int.Parse(dataGridView1[c, i].Value.ToString());
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("セルの値を数値に変換できませんでした.");

                            // アンドゥを行い、途中までの入力結果を取り消す
                            undoFunction();

                            return;
                        }

                        // アンドゥ情報の作成
                        UndoOperation undo = new UndoOperation(
                            c, i, dataGridView1[c, i].Value.ToString(), "");

                        switch (mode)
                        {
                            case CalcMode.Plus:
                                dataGridView1[c,i].Value = (celNum + num).ToString();
                                undo.NewValue = (celNum + num).ToString();
                                //(*pColorBuf)[c,i] = versionNumber;
                                // アンドゥ情報の記録
                                undoGroup.Operations.Add(undo);
                                break;
                            case CalcMode.Minus:
                                dataGridView1[c,i].Value = (celNum - num).ToString();
                                undo.NewValue = (celNum - num).ToString();
                                //(*pColorBuf)[c][i] = versionNumber;
                                break;
                            case CalcMode.Multiple:
                                if (celNum != 0)
                                {
                                    dataGridView1[c, i].Value = (celNum * num).ToString();
                                    undo.NewValue = (celNum * num).ToString();
                                    //(*pColorBuf)[c][i] = versionNumber;
                                    // アンドゥ情報の記録
                                    undoGroup.Operations.Add(undo);
                                }
                                break;
                            case CalcMode.Divide:
                                if (celNum != 0)
                                {
                                    dataGridView1[c, i].Value = (celNum / num).ToString();
                                    undo.NewValue = (celNum / num).ToString();
                                    //(*pColorBuf)[c][i] = versionNumber;
                                    // アンドゥ情報の記録
                                    undoGroup.Operations.Add(undo);
                                }
                                break;
                        }
                    }
                }

                // アンドゥ情報の記録
                undoList.Add(undoGroup);
                // アンドゥメニュ更新
                setUndoText(undoGroup.Title);
                // アンドゥ履歴の調整
                processUndoLimit();

                isFirstEdit = true;

                // 描画更新(継続記号の更新の為)
                dataGridView1.Invalidate();
            }
        }

        //----------------------------------------------------------------------------------------
        private void duplicateToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // 複製

            int Col, Row, Cnt, Len;
            Col = selectRange.Left;
            Row = selectRange.Top;
            Len = selectRange.Height;

            // （回数は２回固定）
            Cnt = 2;

            // アンドゥグループの作成
            UndoGroup undoGroup = new UndoGroup(
                setting.UndoCount++, "複製", UndoOperationTarget.ToGrid);

            //複製操作
            for (int i = 0; i < Len; i++)
            {
                String val = dataGridView1[Col, Row + i].Value.ToString();

                // アンドゥ情報の作成
                UndoOperation undo = new UndoOperation(
                    Col, Row + Len + i,
                    dataGridView1[Col, Row + Len + i].Value.ToString(), val);

                dataGridView1[Col, Row + Len + i].Value = val;
                //(*pColorBuf)[Col][Row + Len + i] = versionNumber;

                // アンドゥ情報の記録
                undoGroup.Operations.Add(undo);
            }

            // 新しい選択範囲を設定
            Rect rect = selectRange;
            dataGridView1.ClearSelection();
            rect.Y += Len;
            for (int i = 0; i < rect.Height; i++)
                for (int j = 0; j < rect.Width; j++)
                    dataGridView1[rect.X + j, rect.Y + i].Selected = true;

            // 範囲を保存
            selectRange = rect;

            // アンドゥ情報の記録
            undoList.Add(undoGroup);
            // アンドゥメニュ更新
            setUndoText(undoGroup.Title);
            // アンドゥ履歴の調整
            processUndoLimit();

            isFirstEdit = true;

            // 描画更新(継続記号の更新の為)
            dataGridView1.Invalidate();
        }

        //----------------------------------------------------------------------------------------
        private void saveSTSToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // STS 保存
            SaveFileDialog dialog = new SaveFileDialog();
            dialog.FileName = @"*.sts";
            //dialog.InitialDirectory = @".\";
            dialog.Filter = "STSファイル(*.sts)|*.sts";
            dialog.FilterIndex = 0;
            dialog.Title = "保存するファイル名を指定してください";
            dialog.RestoreDirectory = false;
            dialog.CheckFileExists = false;
            dialog.CheckPathExists = true;

            //ダイアログを表示する
            if (dialog.ShowDialog(this.owner) == DialogResult.OK)
            {
                //OKボタンがクリックされた場合は、保存実行

                //------------------------
                // 各セルのキーをなめて、コマ数とセルの値を書き出す
                // ※fpsなどの設定は無視
                FileStream outfs;
                char[] header = { (char)0x11, 'S', 'h', 'i', 'r', 'a', 'h', 'e', 'i', 'T', 'i', 'm', 'e', 'S', 'h', 'e', 'e', 't' };
                int col = setting.ColLength;
                long row = setting.RowLength;
                try
                {
                    outfs = new FileStream(dialog.FileName, FileMode.Create);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("指定ファイルを開けませんでした.");
                    return;
                }

                {
                    // ヘッダ
                    byte[] temp = new byte[header.Length];
                    for (int i = 0; i < header.Length; i++)
                        temp[i] = (byte)header[i];
                    outfs.Write(temp,0,temp.Length);
                }

                {
                    // 列数
                    byte colCount = (Byte)col;
                    outfs.WriteByte(colCount);

                    // 行数
                    UInt32 rowCount = (UInt32)row;
                    byte[] temp = BitConverter.GetBytes(rowCount);
                    outfs.Write(temp, 0, temp.Length);
                }

                // セル
                for (int i = 0; i < col; i++)
                {
                    UInt16 current = 0;    // 初期状態は空セル( = 0)
                    for (int j = 0; j < row; j++)
                    {
                        if (dataGridView1[i, j].Value.ToString() == "")
                        {
                            //(特に何もしない)
                        }
                        else
                        {
                            //セルの値を変換してカレント値を更新
                            UInt16 val = 0;
                            try
                            {
                                val = UInt16.Parse(dataGridView1[i, j].Value.ToString());
                            }
                            catch (Exception ex)
                            {
                            }
                            current = val;
                        }
                        // カレントの値を書き出す
                        byte[] temp = BitConverter.GetBytes(current);
                        outfs.Write(temp, 0, temp.Length);
                    }
                }

                // 各セルの名称を出力
                for (int i = 0; i < col; i++)
                {
                    byte[] name = Encoding.GetEncoding("Shift_JIS").GetBytes(dataGridView1.Columns[i].HeaderText);
                    outfs.WriteByte((byte)name.Length);
                    outfs.Write(name, 0, name.Length);
                }

                // ストリームを閉じる
                outfs.Close();

            }

        }

        //----------------------------------------------------------------------------------------
        private void loadSTS(String path)
        {
            FileStream inpfs;
            char[] header = { (char)0x11, 'S', 'h', 'i', 'r', 'a', 'h', 'e', 'i', 'T', 'i', 'm', 'e', 'S', 'h', 'e', 'e', 't' };
            int col = setting.ColLength;
            int row = setting.RowLength;
            try
            {
                inpfs = new FileStream(path, FileMode.Open);
            }
            catch (Exception ex)
            {
                MessageBox.Show("指定ファイルを開けませんでした.");
                return;
            }

            {
                //ヘッダ
                byte[] temp = new byte[header.Length];
                inpfs.Read(temp, 0, header.Length);
                for (int i = 0; i < header.Length; i++)
                {
                    if ((byte)header[i] != temp[i])
                    {
                        MessageBox.Show("未対応のファイルの為、開けませんでした.");
                        return;
                    }
                }
            }

            {
                // 列・行
                byte[] rowLen = new byte[sizeof(UInt32)];
                col = inpfs.ReadByte();
                inpfs.Read(rowLen, 0, rowLen.Length);
                row = (int)BitConverter.ToUInt32(rowLen, 0);
            }

            // グリッドサイズをデータに合わせる
            setting.ColLength = col;
            setting.RowLength = row;
            InitializeWork(false);

            // セル
            byte[] cell = new byte[sizeof(UInt16)];
            for (int i = 0; i < col; i++)
            {
                UInt16 current = 0;
                for (int j = 0; j < row; j++)
                {
                    inpfs.Read(cell, 0, cell.Length);
                    UInt16 val = BitConverter.ToUInt16(cell, 0);

                    // 読み込む際は値の変更部分だけをセルに反映
                    if (current != val)
                    {
                        current = val;
                        dataGridView1[i, j].Value = val.ToString();
                    }
                }
            }

            // 各セルの名称を入力
            for (int i = 0; i < col; i++)
            {
                int nameLen = inpfs.ReadByte();
                byte[] temp = new byte[nameLen];
                inpfs.Read(temp, 0, nameLen);
                dataGridView1.Columns[i].HeaderText =
                    Encoding.GetEncoding("Shift_JIS").GetString(temp);
            }

            // ストリームを閉じる
            inpfs.Close();
        }

        //----------------------------------------------------------------------------------------
        private void loadSTSToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // STS 読み込み
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.FileName = @"*.sts";
            dialog.InitialDirectory = @".\";
            dialog.Filter = "STSファイル(*.sts)|*.sts";
            dialog.FilterIndex = 0;
            dialog.Title = "読み込むファイル名を指定してください";
            dialog.RestoreDirectory = false;
            dialog.CheckFileExists = true;
            dialog.CheckPathExists = true;
            dialog.Multiselect = false;

            //ダイアログを表示する
            if (dialog.ShowDialog(this.owner) == DialogResult.OK)
            {
                //OKボタンがクリックされた場合は、読み込み実行
                loadSTS(dialog.FileName);

            }

            // 描画更新(継続記号の更新の為)
            dataGridView1.Invalidate();
        }

        //----------------------------------------------------------------------------------------
        private void Form1_Shown(object sender, EventArgs e)
        {
            //コントロールにフォーカスを設定
            dataGridView1.Focus();
        }
    }

}
