// AfterEffects Input/Output Utility v3.x  Programed by kanbara.
//
//  ■v3.2
//  (26-10-03) v3.2.0.0
//      ・データグリッドをVirtualMode化し、描画・入力・Undo・STS/AEデータ処理をForm1から分離
//      ・キー入力をコマンド化して、移動・範囲選択・値入力・行削除の処理経路を整理
//      ・自動処理を共通契約とホストへ移管し、外部DLLの検出・実行とランダム入力／AEコピーのサンプルを追加
//
//  ■v3.1
//  (24-11-19 6:00) v3.1.0.0
//　　　・アンドゥ実装の書き換え
//　　　　→リドゥ動作の追加
//      ・タイミングコピー時のバージョン文字列を iniファイルに記録する動作の追加（.iniファイルは削除して、新しく作る必要あり）
//      ・起動時に、dataGridViewにフォーカスを当てる（起動時に、クリックしないとキー入力が効かないため）
//      ・セルの列ヘッダクリックで、１列全選択する動作を追加（何故か付けてなかった..）
//      ・繰り返し操作のダイアログをモードレスに変更して、Okボタンを押す事に、Undoして繰り返し情報の入力。という動作に
//        （単にUndoするもので、繰り返しの入力だけを狙ってUndoするわけではない点に注意）
//      ・初回コピーで例外が上がる場合があるのを修正（リトライ処理、リトライ失敗時の通知）
//
//  ■v3.0
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
using global::AEIOU.Automation;

namespace AEIOU
{
    public partial class Form1 : Form, IGridShortcutHandler, IAutomationChangeTarget, IAutomationSessionTarget
    {
	    //----------------------------------------------------------------------------------------
	    // 配色
	    ColorDefinitions gridPalette = new ColorDefinitions();        //グリッド描画用 色設定

	    // 設定用 各種変数
	    Settings setting = new Settings();

	    // 各種変数
	    Control owner;                              //子ウィンドウ（ダイアログ）に与える自分のウィンドウハンドル
	    Rect selectRange;                           //選択範囲
	    Rect copyRect;                              //コピー範囲
	    Point mouseDownPoint;                       //マウス押下位置
	    bool isRectDrag;                            //範囲移動状態
        bool isCtrlDragCopy;                        //Ctrlドラッグでコピー動作を行うか
	    bool isFirstEdit;                           //初回編集状態（セル入力を中断するような操作の度に trueにされるべき）
        bool isCellEdit;                            //セルが編集されたかどうかのフラグ

	    // 配列
	    int[] aryCellUsedCount;                     //セル使用状況
	                                                //int[, ,] gAryBuff;                          //バッファ

	    // リスト
	    List<Range> delRange;                       // 中抜き範囲
	    List<Range> addRange;                       // 切り貼り範囲

        // アンドゥ処理
        private GridViewManager gridViewManager = new GridViewManager();
        private GridFrameLabelFormatter gridFrameLabelFormatter;
        private GridFrameHeaderPainter gridFrameHeaderPainter;
        private GridBorderStateCalculator gridBorderStateCalculator;
        private GridShortcutRouter gridShortcutRouter;
        private GridKeyCommandDispatcher gridKeyCommandDispatcher;
        private GridMoveSelectionCommand gridMoveSelectionCommand;
        private GridValueInputCommand gridValueInputCommand;
        private GridCellValueService gridCellValueService;
        private GridMouseEventHandler gridMouseEventHandler;
        private readonly StsFileService stsFileService = new StsFileService();
        private readonly AfterEffectsDataService afterEffectsDataService = new AfterEffectsDataService();
        private readonly AutomationRegistry automationRegistry;
        private readonly AutomationHost automationHost;
        private long automationGeneration;

        // 先行分離したサービス
        GridSelectionService gridSelectionService;
        GridScrollService gridScrollService;
        GridCellStyleResolver gridCellStyleResolver;
        GridCellRenderer gridCellRenderer;
        GridInputInterpreter gridInputInterpreter;
        TimingSheetModel timingSheetModel;
        ContinuityStateService continuityStateService;
        bool isCellValuePushedBound;

        // 繰り返しダイアログ
        private RepeatInputBox _repeatInputDialog;
        private AutomationSession _repeatAutomationSession;

        //----------------------------------------------------------------------------------------
        // コンストラクタ
        public Form1()
        {
            automationRegistry = BuiltInAutomationRegistry.Create();
            automationHost = new AutomationHost(1000000, automationRegistry);
            InitializeComponent();
            LoadAutomationExtensions();
            gridViewManager.View = dataGridView1;

            // 自分のウィンドウハンドルを取得しておく
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
                this.fPS24ToolStripMenuItem.ShortcutKeys = (Keys)Enum.Parse(typeof(Keys), setting.keys.Set24FPS);
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

            // 先行分離サービスの初期化
            gridSelectionService = new GridSelectionService(dataGridView1, setting);
            gridScrollService = new GridScrollService(dataGridView1, setting);
            gridCellStyleResolver = new GridCellStyleResolver();
            gridCellRenderer = new GridCellRenderer(dataGridView1, setting, GetCellValue);
            gridFrameLabelFormatter = new GridFrameLabelFormatter(setting.Fps, setting.SheetSec);
            gridFrameHeaderPainter = new GridFrameHeaderPainter(gridFrameLabelFormatter);
            gridBorderStateCalculator = new GridBorderStateCalculator(setting.Fps, setting.SheetSec, setting.SheetDivide);
            gridInputInterpreter = new GridInputInterpreter(setting.keys);
            gridShortcutRouter = new GridShortcutRouter(this); // Initialize GridShortcutRouter here
            gridKeyCommandDispatcher = new GridKeyCommandDispatcher(gridInputInterpreter, gridShortcutRouter, tryExecuteShortcut);
            gridCellValueService = new GridCellValueService(gridViewManager, GetCellValue, checkCellValue, (col) => aryCellUsedCount[col]++);
            gridMoveSelectionCommand = new GridMoveSelectionCommand(dataGridView1, setting, gridSelectionService, gridScrollService);
            gridValueInputCommand = new GridValueInputCommand(dataGridView1, setting, gridCellValueService, gridSelectionService, gridScrollService);
            gridMouseEventHandler = new GridMouseEventHandler(gridViewManager, copyToBuf, cutToBuf, copyToCell, getSelectedRect);
            continuityStateService = new ContinuityStateService(setting, GetCellValue);
            continuityStateService.Reinitialize(GetSheetColumnCount(), GetSheetRowCount());

            // 読み込みファイル指定がある場合 ファイル読込を行う
            if (cmds.Length > 1 && File.Exists(cmds[1]))
            {
                loadSTS(cmds[1]);
            }
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
        // 作業内容の初期化
        public void InitializeWork(bool isBoot)
        {
            SetCellValuePushedBinding(false);

            timingSheetModel = new TimingSheetModel(setting.ColLength, setting.RowLength);

            // ヘッダ、各カラム及び行数の設定
            dataGridInitialize(setting.ColLength, setting.RowLength, 50, isBoot);    // 列, 行, 列幅

            // 配列作成
            aryCellUsedCount = new int[setting.ColLength];
            //gAryBuff = new int[setting.ColLength, setting.RowLength, gEtcLength];

            // リスト作成
            delRange = new List<Range>();
            addRange = new List<Range>();

            // アンドゥメニュを初期化
            gridViewManager.InitializeWork(dataGridView1, timingSheetModel);
            gridViewManager.CellValueChanged -= OnGridViewManagerCellValueChanged;
            gridViewManager.CellValueChanged += OnGridViewManagerCellValueChanged;

            if (continuityStateService != null)
            {
                continuityStateService.Reinitialize(GetSheetColumnCount(), GetSheetRowCount());
            }

            // 各種変数の初期化
            selectRange = new Rect(0, 0, 1, 1);     //選択範囲
            copyRect = new Rect(-1, -1, 0, 0);      //コピー範囲
            mouseDownPoint = new Point(-1, -1);     //マウス押下位置
            isRectDrag = false;                     //範囲移動状態
            isCtrlDragCopy = false;
            isFirstEdit = true;
            isCellEdit = false;

            SetCellValuePushedBinding(true);
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
                isCtrlDragCopy = false;
                isFirstEdit = true;
                isCellEdit = false;
            }
            if ((target & InitializeTarget.Timing) != 0)
            {
                SetCellValuePushedBinding(false);

                // シートの入力情報（タイミング）の初期化
                // ※バージョンの扱いをどうするのかは未定
                timingSheetModel = new TimingSheetModel(setting.ColLength, setting.RowLength);
                gridViewManager.Model = timingSheetModel;
                dataGridInitialize(setting.ColLength, setting.RowLength, 50, false);    // 列, 行, 列幅
                aryCellUsedCount = new int[setting.ColLength];
                for (int i = 0; i < setting.RowLength; i++)
                    for (int j = 0; j < setting.ColLength; j++)
                    {
                        SetCellValue(j, i, "");
                    }

                if (continuityStateService != null)
                {
                    continuityStateService.Reinitialize(GetSheetColumnCount(), GetSheetRowCount());
                }

                isFirstEdit = true;
                isCellEdit = false;

                SetCellValuePushedBinding(true);
            }
            if ((target & InitializeTarget.CopyBuffer) != 0)
            {
                // コピーバッファの初期化
                copyRect = new Rect(-1, -1, 0, 0);     //コピー範囲
            }
            if ((target & InitializeTarget.UndoHistory) != 0)
            {
                // アンドゥ処理の初期化
                gridViewManager.InitializeWork(dataGridView1, timingSheetModel);

            }
            if ((target & InitializeTarget.RangeSetting) != 0)
            {
                // 中抜き・切り貼り設定範囲の初期化
                delRange = new List<Range>();
                addRange = new List<Range>();
            }

        }

        //----------------------------------------------------------------------------------------
        // CellValuePushedイベントのバインド設定
        private void SetCellValuePushedBinding(bool enabled)
        {
            if (enabled)
            {
                if (!isCellValuePushedBound)
                {
                    dataGridView1.CellValuePushed += new DataGridViewCellValueEventHandler(this.dataGridView1_CellValuePushed);
                    isCellValuePushedBound = true;
                }

                return;
            }

            if (isCellValuePushedBound)
            {
                dataGridView1.CellValuePushed -= new DataGridViewCellValueEventHandler(this.dataGridView1_CellValuePushed);
                isCellValuePushedBound = false;
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
                SetHeaderValue(i, cellName);
            }

            // 初期化 (行の生成 : 中身は空)
            gridViewManager.SyncGridShape(columnCount, rowCount);
            for (int i = 0; i < rowCount; i++)
                for (int j = 0; j < columnCount; j++)
                    SetCellValue(j, i, "");

        }

        //----------------------------------------------------------------------------------------
        // dataGridView1のグリッドサイズを変更する
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
                headerName[i] = GetHeaderValue(i);
                for (int j = 0; j < row; j++)
                {
                    temp[i, j] = GetCellValue(i, j);
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
                SetHeaderValue(i, headerName[i]);
                for (int j = 0; j < row; j++)
                {
                    SetCellValue(i, j, temp[i, j]);
                }
            }
        }

        //---------------------------------------------------------------------------
        // ウィンドウサイズ・位置の調整
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
        // リドゥ処理
        private void redoFunction()
        {
            gridViewManager.Redo();

            // 描画更新(継続記号の更新の為)
            dataGridView1.Invalidate();
        }

        //----------------------------------------------------------------------------------------
        // アンドゥ処理
        private void undoFunction()
        {
            gridViewManager.Undo();

            // 描画更新(継続記号の更新の為)
            dataGridView1.Invalidate();

        }

        //---------------------------------------------------------------------------
        // コピー先セルへのペースト処理
        private void copyToCell(int col, int row, bool shouldInvalidate)
        {
            // PasteOperationのインスタンスを作成
            PasteOperation pasteOperation = new PasteOperation(row, col);
            
            // GridViewManagerを使用して操作を実行
            gridViewManager.ExecuteOperation(pasteOperation);
            
            // 描画更新(継続記号の更新の為)
            if (shouldInvalidate)
            {
                dataGridView1.Invalidate();
            }
        }

        //----------------------------------------------------------------------------------------
        // コピー範囲をコピーバッファにコピーする
        private void copyToBuf(Rect rect)
        {
            // CopyOperationのインスタンスを作成
            CopyOperation copyOperation = new CopyOperation(rect);
            
            // GridViewManagerを使用して操作を実行
            gridViewManager.ExecuteOperation(copyOperation);

        }

        //---------------------------------------------------------------------------
        // 切り取り範囲をコピーバッファにコピーし、シートからは削除する
        private void cutToBuf(Rect rect, bool shouldInvalidate)
        {
            // CutOperationのインスタンスを作成
            CutOperation cutOperation = new CutOperation(rect);
            
            // GridViewManagerを使用して操作を実行
            gridViewManager.ExecuteOperation(cutOperation);

            // 描画更新(継続記号の更新の為)
            if (shouldInvalidate)
            {
                dataGridView1.Invalidate();
            }

        }

        //----------------------------------------------------------------------------------------
        // 指定範囲をシートから削除する
        private void deleteRect(Rect rect)
        {
            deleteRect(rect, true);
        }

        //----------------------------------------------------------------------------------------
        // 指定範囲をシートから削除する
        private void deleteRect(Rect rect, bool shouldInvalidate)
        {
            // DeleteOperationのインスタンスを作成
            DeleteOperation deleteOperation = new DeleteOperation(rect);
            
            // GridViewManagerを使用して操作を実行
            gridViewManager.ExecuteOperation(deleteOperation);
            
            // 描画更新(継続記号の更新の為)
            if (shouldInvalidate)
            {
                dataGridView1.Invalidate();
            }

            isFirstEdit = true;
        }

        //----------------------------------------------------------------------------------------
        // 指定セルに列を挿入する
        void insertToAllCell(int Row, int Count)
        {
            IList<CellWriteEntry> writes = SheetRowEditCalculator.CreateInsertRows(
                GetSheetRowCount(), GetSheetColumnCount(), Row, Count,
                delegate(int row, int column) { return GetCellValue(column, row); });
            ApplyRowEdit("行の挿入", writes);

        }

        //----------------------------------------------------------------------------------------
        // 指定セルから列を削除する
        void cutToAllCell(int Row, int Count)
        {
            IList<CellWriteEntry> writes = SheetRowEditCalculator.CreateDeleteRows(
                GetSheetRowCount(), GetSheetColumnCount(), Row, Count,
                delegate(int row, int column) { return GetCellValue(column, row); });
            ApplyRowEdit("行の削除", writes);

        }

        //----------------------------------------------------------------------------------------
        // 中抜き・切り貼り領域の再計算
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
        // 切り貼り領域の再計算
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
        // カーソル移動量の計算 (中抜き範囲補正付き)
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
        // 基準線描画位置の計算
        private SheetBorder calcBorderState(DataGridViewCellPaintingEventArgs e)
        {
            return gridBorderStateCalculator.CalcBorderState(e, addRange);
        }

        //----------------------------------------------------------------------------------------
        // フレーム数 描画
        private void drawFrameNumber(DataGridViewCellPaintingEventArgs e)
        {
            gridFrameHeaderPainter.DrawHeader(
                e,
                gridPalette.Header,
                setting.IsDisplayFrameNumber,
                setting.FirstFrame,
                addRange);
        }

        //----------------------------------------------------------------------------------------
        // タイミング継続か否かの確認（事前計算済み結果の参照のみ）
        private bool checkContinuty(int X, int Y)
        {
            return continuityStateService != null &&
                   continuityStateService.GetContinuityFlag(X, Y);
        }

        //----------------------------------------------------------------------------------------
        // セルの値が変化した際の処理
        private void OnGridViewManagerCellValueChanged(int col, int row, string value)
        {
            if (continuityStateService == null)
            {
                return;
            }

            if (col < 0 || col >= GetSheetColumnCount())
            {
                return;
            }

            continuityStateService.RecalculateColumn(col, GetSheetRowCount());
        }

        //----------------------------------------------------------------------------------------
        // 各セルの描画
        private void dataGridView1_CellPainting(object sender, DataGridViewCellPaintingEventArgs e)
        {
            // フレーム数の表示 ------------------------------------------------------------
            if (gridCellRenderer.TryPaintRowHeader(e, drawFrameNumber))
            {
                return;
            }

            // 継続記号 評価
            bool bLine = checkContinuty(e.ColumnIndex, e.RowIndex);
            Color bgColor = gridCellStyleResolver.ResolveBackColor(
                dataGridView1,
                e,
                gridPalette,
                aryCellUsedCount,
                delRange,
                addRange,
                isRectDrag,
                selectRange,
                mouseDownPoint);

            //背景色を設定
            gridCellRenderer.ApplyBackColor(e, bgColor);

            //値の取得範囲を制限
            //※ヘッダー部で値取得すると、中身がnullの為に例外が発生する
            if (e.RowIndex >= 0 && e.ColumnIndex >= 0)
            {   // セルのタイミング入力部分
                //基準線描画位置の計算と TimingCell 状態更新
                SheetBorder borderState = calcBorderState(e);
                gridCellRenderer.ApplyTimingCellState(e, borderState, bLine);

            }
            else
            {
                //ヘッダー部分 : 特になにもしない
            }

            //描画を要求
            gridCellRenderer.PaintCell(e);

        }

        //----------------------------------------------------------------------------------------
        // セルの値の取得
        private void dataGridView1_CellValueNeeded(object sender, DataGridViewCellValueEventArgs e)
        {
            if (!IsValidCellIndex(e.ColumnIndex, e.RowIndex))
            {
                e.Value = string.Empty;
                return;
            }

            string value;
            e.Value = gridViewManager.TryHandleCellValueNeeded(e.ColumnIndex, e.RowIndex, out value)
                ? value ?? string.Empty
                : string.Empty;
        }

        //----------------------------------------------------------------------------------------
        // セルの値の設定
        private void dataGridView1_CellValuePushed(object sender, DataGridViewCellValueEventArgs e)
        {
            if (!IsValidCellIndex(e.ColumnIndex, e.RowIndex))
            {
                return;
            }

            gridViewManager.PushCellValue(e.ColumnIndex, e.RowIndex, e.Value);

            if (continuityStateService != null)
            {
                continuityStateService.RecalculateColumn(e.ColumnIndex, GetSheetRowCount());
            }
        }

        //----------------------------------------------------------------------------------------
        // 選択範囲の取得
        private Rect getSelectedRect()
        {
            return gridSelectionService.GetSelectedRect();
        }

        private string GetCellValue(int col, int row)
        {
            string value;
            if (!TryGetCellValue(col, row, out value))
            {
                return string.Empty;
            }

            return value ?? string.Empty;
        }

        //----------------------------------------------------------------------------------------
        // セルの値の取得 (失敗理由も返す版)
        private bool TryGetCellValue(int col, int row, out string value, out string failureReason)
        {
            if (IsGridViewManagerBoundToCurrentModel())
            {
                return gridViewManager.TryGetCellValue(col, row, out value, out failureReason);
            }

            if (timingSheetModel == null)
            {
                value = string.Empty;
                failureReason = "ModelNotInitialized";
                return false;
            }

            return timingSheetModel.TryGetCell(col, row, out value, out failureReason);
        }

        //----------------------------------------------------------------------------------------
        // セルの値の取得 (失敗理由は不要な簡易版)
        private bool TryGetCellValue(int col, int row, out string value)
        {
            string failureReason;
            return TryGetCellValue(col, row, out value, out failureReason);
        }

        //----------------------------------------------------------------------------------------
        // セルの値の設定
        private bool IsValidCellIndex(int col, int row)
        {
            return col >= 0 &&
                   row >= 0 &&
                   col < GetSheetColumnCount() &&
                   row < GetSheetRowCount();
        }

        //----------------------------------------------------------------------------------------
        // セルの値の設定
        private void SetCellValue(int col, int row, string value)
        {
            if (!IsValidCellIndex(col, row))
            {
                return;
            }

            if (IsGridViewManagerBoundToCurrentModel())
            {
                gridViewManager.SetCellValue(col, row, value);
            }
            else
            {
                string normalizedValue = value ?? "";
                timingSheetModel.SetCell(col, row, normalizedValue);
                if (gridViewManager != null)
                {
                    gridViewManager.SetCellDisplayValue(col, row, normalizedValue);
                }
            }

            if (continuityStateService != null)
            {
                continuityStateService.RecalculateColumn(col, GetSheetRowCount());
            }
        }

        //----------------------------------------------------------------------------------------
        // セルの値の設定（セル値が変化した場合のみ）
        private bool IsGridViewManagerBoundToCurrentModel()
        {
            return gridViewManager != null &&
                   timingSheetModel != null &&
                   gridViewManager.Model == timingSheetModel;
        }

        //----------------------------------------------------------------------------------------
        // シートの列数の取得
        private int GetSheetColumnCount()
        {
            if (IsGridViewManagerBoundToCurrentModel())
            {
                return gridViewManager.ColumnCount;
            }

            return setting.ColLength;
        }

        //----------------------------------------------------------------------------------------
        // シートの行数の取得
        private int GetSheetRowCount()
        {
            if (IsGridViewManagerBoundToCurrentModel())
            {
                return gridViewManager.RowCount;
            }

            return setting.RowLength;
        }

        //----------------------------------------------------------------------------------------
        // ヘッダの値の取得・設定
        private string GetHeaderValue(int col)
        {
            if (IsGridViewManagerBoundToCurrentModel())
            {
                return gridViewManager.GetHeaderValue(col);
            }

            return timingSheetModel.GetHeader(col);
        }

        //----------------------------------------------------------------------------------------
        // ヘッダの値の設定
        private void SetHeaderValue(int col, string value)
        {
            if (IsGridViewManagerBoundToCurrentModel())
            {
                gridViewManager.SetHeaderValue(col, value);
            }
            else
            {
                string normalizedValue = value ?? "";
                timingSheetModel.SetHeader(col, normalizedValue);
                if (gridViewManager != null)
                {
                    gridViewManager.SetHeaderDisplayValue(col, normalizedValue);
                }
            }
        }

        //----------------------------------------------------------------------------------------
        // 列編集 calculator が作成した snapshot を順番どおり反映する
        private void ApplyColumnWrites(IList<ColumnWriteEntry> writes, int rowCount)
        {
            foreach (ColumnWriteEntry write in writes)
            {
                aryCellUsedCount[write.Column] = write.UsedCount;
                SetHeaderValue(write.Column, write.Header);
                for (int row = 0; row < rowCount; row++)
                {
                    SetCellValue(write.Column, row, write.GetValue(row));
                }
            }
        }

        //----------------------------------------------------------------------------------------
        // 複数セルへの書き込みをグループ化して実行する
        private void ExecuteWriteGroup(string groupName, Action action)
        {
            gridViewManager.BeginBatchUpdate();
            gridViewManager.BeginGroup(groupName);
            try
            {
                action();
            }
            finally
            {
                try
                {
                    gridViewManager.EndGroup();
                }
                finally
                {
                    gridViewManager.EndBatchUpdate();
                }
            }
        }

        //----------------------------------------------------------------------------------------
        // SetValueOperation の引数順（row, col, value）を明示して、
        // 既存コードの Col/Row 変数名との取り違えを防ぐ。
        private void QueueCellWrite(int row, int col, string value)
        {
            var operation = new SetValueOperation(row, col, value);
            gridViewManager.ExecuteOperation(operation);
        }

        //----------------------------------------------------------------------------------------
        // セルの値が変化した場合のみ、値を設定する
        private bool SetCellValueIfChanged(int col, int row, string value)
        {
            string normalizedValue = value ?? "";
            if (GetCellValue(col, row) == normalizedValue)
            {
                return false;
            }

            SetCellValue(col, row, normalizedValue);
            return true;
        }

        //----------------------------------------------------------------------------------------
        // 複数セルへの書き込みをグループ化して実行する
        private void ApplyCellWrites(string groupName, IList<CellWriteEntry> writes)
        {
            if (writes == null)
            {
                throw new ArgumentNullException("writes");
            }
            if (writes.Count == 0)
            {
                return;
            }

            // Undo group を開く前に全件を検証し、不正な一覧の部分適用を防ぐ。
            CellWriteBatch.Validate(writes, GetSheetRowCount(), GetSheetColumnCount());

            ExecuteWriteGroup(groupName, delegate
            {
                QueueCellWrites(writes);
            });
        }

        //----------------------------------------------------------------------------------------
        // 行編集の計算後は、検証済みの一覧を1 Undo group で適用してから
        // isFirstEdit と継続記号の表示をまとめて更新する。
        private void ApplyRowEdit(string groupName, IList<CellWriteEntry> writes)
        {
            ApplyCellWrites(groupName, writes);
            FinishWriteOperation(true);
        }

        //----------------------------------------------------------------------------------------
        // 複数セルへの書き込みをキューに追加する
        private void QueueCellWrites(IList<CellWriteEntry> writes)
        {
            if (writes == null || writes.Count == 0)
            {
                return;
            }

            foreach (CellWriteEntry write in writes)
            {
                QueueCellWrite(row: write.Row, col: write.Col, value: write.Value);
            }
        }

        //----------------------------------------------------------------------------------------
        // 書き込み操作の終了処理
        private void FinishWriteOperation(bool shouldInvalidate)
        {
            isFirstEdit = true;

            if (shouldInvalidate)
            {
                // 描画更新(継続記号の更新の為)
                dataGridView1.Invalidate();
            }
        }

        private AutomationRequest CreateAutomationRequest(IDictionary<string, string> parameters)
        {
            List<AutomationCell> cells = new List<AutomationCell>();
            string failureReason;
            for (int column = selectRange.Left; column <= selectRange.Right; column++)
                for (int row = selectRange.Top; row <= selectRange.Bottom; row++)
                {
                    string value;
                    if (!TryGetCellValue(column, row, out value, out failureReason)) value = String.Empty;
                    cells.Add(new AutomationCell(row, column, value));
                }

            return new AutomationRequest(GetSheetRowCount(), GetSheetColumnCount(), automationGeneration,
                new AutomationSelection(selectRange.Top, selectRange.Left, selectRange.Height, selectRange.Width),
                cells, parameters, setting.KaraCell);
        }

        private void ExecuteAutomationCommand(string commandId, IDictionary<string, string> parameters)
        {
            AutomationHostResult result = automationHost.Execute(commandId, CreateAutomationRequest(parameters), this);
            if (!result.Succeeded)
            {
                MessageBox.Show(result.Error);
                return;
            }
            FinishWriteOperation(true);
        }

        private void LoadAutomationExtensions()
        {
            string applicationDirectory = AppDomain.CurrentDomain.BaseDirectory;
            AutomationExtensionLoadResult loadResult = new AutomationExtensionLoader().Load(
                Path.Combine(applicationDirectory, "Extensions"), automationRegistry);
            AutomationExtensionLoader.AppendDiagnostics(
                Path.Combine(applicationDirectory, "automation-extensions.log"), loadResult.Diagnostics);
            if (loadResult.Commands.Count == 0) return;

            ToolStripMenuItem extensionsMenu = new ToolStripMenuItem("拡張自動処理");
            foreach (AutomationCommandDescriptor descriptor in loadResult.Descriptors)
            {
                ToolStripMenuItem item = new ToolStripMenuItem(descriptor.DisplayName);
                item.Tag = descriptor.Id;
                item.Click += externalAutomationToolStripMenuItem_Click;
                extensionsMenu.DropDownItems.Add(item);
            }
            int insertionIndex = contextMenuStrip1.Items.IndexOf(fourArithmeticOperationToolStripMenuItem) + 1;
            contextMenuStrip1.Items.Insert(insertionIndex, extensionsMenu);
        }

        private void externalAutomationToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ToolStripMenuItem item = sender as ToolStripMenuItem;
            AutomationCommandDescriptor descriptor;
            if (item == null || !automationRegistry.TryGetDescriptor((string)item.Tag, out descriptor)) return;
            IDictionary<string, string> parameters;
            if (!TryCollectAutomationParameters(descriptor, out parameters)) return;
            ExecuteAutomationCommand(descriptor.Id, parameters);
        }

        private bool TryCollectAutomationParameters(AutomationCommandDescriptor descriptor,
            out IDictionary<string, string> parameters)
        {
            parameters = null;
            Dictionary<string, Control> inputs = new Dictionary<string, Control>(StringComparer.Ordinal);
            using (Form dialog = new Form())
            using (TableLayoutPanel layout = new TableLayoutPanel())
            {
                dialog.Text = descriptor.DisplayName;
                dialog.FormBorderStyle = FormBorderStyle.FixedDialog;
                dialog.StartPosition = FormStartPosition.CenterParent;
                dialog.MinimizeBox = false;
                dialog.MaximizeBox = false;
                dialog.AutoSize = true;
                dialog.AutoSizeMode = AutoSizeMode.GrowAndShrink;
                layout.AutoSize = true;
                layout.ColumnCount = 2;
                layout.Padding = new Padding(8);
                dialog.Controls.Add(layout);

                int row = 0;
                foreach (AutomationParameterDefinition definition in descriptor.Parameters)
                {
                    Label label = new Label();
                    label.Text = definition.DisplayName;
                    label.AutoSize = true;
                    label.Anchor = AnchorStyles.Left;
                    Control input;
                    if (definition.Type == AutomationParameterType.Boolean)
                    {
                        CheckBox checkBox = new CheckBox();
                        bool checkedValue;
                        Boolean.TryParse(definition.DefaultValue, out checkedValue);
                        checkBox.Checked = checkedValue;
                        input = checkBox;
                    }
                    else if (definition.Type == AutomationParameterType.Choice)
                    {
                        ComboBox comboBox = new ComboBox();
                        comboBox.DropDownStyle = ComboBoxStyle.DropDownList;
                        foreach (string choice in definition.Choices) comboBox.Items.Add(choice);
                        int selected = comboBox.Items.IndexOf(definition.DefaultValue);
                        comboBox.SelectedIndex = selected >= 0 ? selected : 0;
                        input = comboBox;
                    }
                    else
                    {
                        TextBox textBox = new TextBox();
                        textBox.Text = definition.DefaultValue ?? String.Empty;
                        textBox.Width = 180;
                        input = textBox;
                    }
                    layout.Controls.Add(label, 0, row);
                    layout.Controls.Add(input, 1, row++);
                    inputs.Add(definition.Id, input);
                }

                FlowLayoutPanel buttons = new FlowLayoutPanel();
                buttons.AutoSize = true;
                buttons.FlowDirection = FlowDirection.RightToLeft;
                Button ok = new Button();
                ok.Text = "OK";
                ok.DialogResult = DialogResult.OK;
                Button cancel = new Button();
                cancel.Text = "キャンセル";
                cancel.DialogResult = DialogResult.Cancel;
                buttons.Controls.Add(ok);
                buttons.Controls.Add(cancel);
                layout.Controls.Add(buttons, 0, row);
                layout.SetColumnSpan(buttons, 2);
                dialog.AcceptButton = ok;
                dialog.CancelButton = cancel;
                if (dialog.ShowDialog(owner) != DialogResult.OK) return false;

                Dictionary<string, string> collected = new Dictionary<string, string>(StringComparer.Ordinal);
                foreach (AutomationParameterDefinition definition in descriptor.Parameters)
                {
                    Control input = inputs[definition.Id];
                    CheckBox checkBox = input as CheckBox;
                    ComboBox comboBox = input as ComboBox;
                    collected.Add(definition.Id, checkBox != null ? checkBox.Checked.ToString() :
                        comboBox != null ? (string)comboBox.SelectedItem : input.Text);
                }
                parameters = collected;
                return true;
            }
        }

        public bool TryApply(long expectedGeneration, IList<AutomationChange> changes, string operationName)
        {
            if (expectedGeneration != automationGeneration) return false;
            List<CellWriteEntry> writes = new List<CellWriteEntry>();
            foreach (AutomationChange change in changes)
                writes.Add(new CellWriteEntry(change.Row, change.Column, change.Value));
            ApplyCellWrites(operationName, writes);
            automationGeneration++;
            return true;
        }

        public bool TryReplace(long expectedGeneration, object previousApplication,
            IList<AutomationChange> changes, string operationName, out object application)
        {
            application = null;
            if (expectedGeneration != automationGeneration) return false;
            if (previousApplication != null)
            {
                GridViewOperation previousOperation = previousApplication as GridViewOperation;
                if (previousOperation == null || !gridViewManager.TryUndo(previousOperation)) return false;
            }

            OperationGroup group = new OperationGroup(operationName);
            foreach (AutomationChange change in changes)
                group.AddOperation(new SetValueOperation(change.Row, change.Column, change.Value));
            gridViewManager.BeginBatchUpdate();
            try
            {
                gridViewManager.ExecuteOperation(group);
            }
            finally
            {
                gridViewManager.EndBatchUpdate();
            }
            application = group;
            automationGeneration++;
            return true;
        }

        //----------------------------------------------------------------------------------------
        // 指定セルの入力有無をチェック
        private bool checkCellValue(int X, int Y)
        {
            bool val = false;

            // チェック範囲は X,Y共に 0以上
            int columnCount = GetSheetColumnCount();
            int rowCount = GetSheetRowCount();
            if ((X >= 0 && X < columnCount) &&
                (Y >= 0 && Y < rowCount))
            {
                // 値が入っていたら trueを返す
                String str = GetCellValue(X, Y);
                if (str.Length > 0)
                    val = true;
            }
            return val;
        }

        //----------------------------------------------------------------------------------------
        // バックスペースキーによる削除処理
        private bool deleteRect_with_backspace(bool isCellEdit)
        {
            // 選択範囲を取得
            Rect rect = getSelectedRect();
            bool isBackward = false;        // 巻き戻し削除になるかどうかのフラグ(default: 巻き戻しなし)

            // カレントセルの内容を確認
            bool isNoBlank = false;
            for (int row = rect.Top; row <= rect.Bottom && !isNoBlank; row++)
            {
                for (int col = rect.Left; col <= rect.Right; col++)
                {
                    if (GetCellValue(col, row) != "")
                    {
                        isNoBlank = true;
                        break;
                    }
                }
            }
            if (isNoBlank)
            {
                // 現在セルに値が入っている場合
                // （なにもしない）

            }
            else
            {
                // 現在セルに値が入っていない場合
                // (選択範囲を変更する)
                for (int i = rect.Y; i >= 0 && !isBackward; i--)
                    for (int l = rect.Left; l <= rect.Right; l++)
                    {
                        // 空白のセルは無視
                        String str = GetCellValue(l, i);
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
                // (選択範囲を変更, 入力値はズラす)

                gridViewManager.BeginGroup("削除");

                // rect.Left～rect.Right で回しているため、ループ変数は「選択範囲内の相対位置」ではなく
                // DataGridView 全体に対する「絶対列インデックス」。
                // そのため SetValueOperation の列引数は col をそのまま渡す（rect.X + col にはしない）。
                // 複数列選択でも Left～Right の各列を1回ずつ処理するため、列ずれは発生しない。
                for (int row = rect.Top; row <= rect.Bottom; row++)
                {
                    for (int col = rect.Left; col <= rect.Right; col++)
                    {
                        // 空白セルは無視する
                        String new_value = GetCellValue(col, row);
                        if (new_value.Length == 0) continue;

                        if (isBackward)
                        {
                            //セル内容の消去
                            new_value = "";
                            //使用状況を修正
                            aryCellUsedCount[col]--;
                        }
                        else
                        {
                            //１桁削る
                            new_value = new_value.Substring(0, new_value.Length - 1);
                            isCellEdit = true;

                            // セルの中身が空白になった場合は、使用状況を修正
                            if (new_value.Length == 0)
                            {
                                aryCellUsedCount[col]--;
                            }
                        }

                        // アンドゥ情報の記録
                        var operation = new SetValueOperation(row, col, new_value);
                        gridViewManager.ExecuteOperation(operation);
                    }

                }

                gridViewManager.EndGroup();

            }

            // 描画更新(継続記号の更新の為)
            dataGridView1.Invalidate();

            return isCellEdit;

        }

        //----------------------------------------------------------------------------------------
        // ショートカットキーの実行
        private bool tryExecuteShortcut(ToolStripItemCollection items, Keys keyData)
        {
            foreach (ToolStripItem item in items)
            {
                ToolStripMenuItem menuItem = item as ToolStripMenuItem;
                if (menuItem == null)
                {
                    continue;
                }

                if (menuItem.Enabled &&
                    menuItem.ShortcutKeys != Keys.None &&
                    menuItem.ShortcutKeys == keyData)
                {
                    menuItem.PerformClick();
                    return true;
                }

                if (menuItem.HasDropDownItems && tryExecuteShortcut(menuItem.DropDownItems, keyData))
                {
                    return true;
                }
            }

            return false;
        }

        //----------------------------------------------------------------------------------------
        // アクティブ列の遷移を無効化
        private void InvalidateActiveColumnTransition(int previousCol)
        {
            if (dataGridView1.CurrentCell == null)
            {
                dataGridView1.Invalidate();
                return;
            }

            int currentCol = dataGridView1.CurrentCell.ColumnIndex;
            int maxCol = GetSheetColumnCount() - 1;

            bool previousValid = previousCol >= 0 && previousCol <= maxCol;
            bool currentValid = currentCol >= 0 && currentCol <= maxCol;

            if (!currentValid)
            {
                dataGridView1.Invalidate();
                return;
            }

            dataGridView1.InvalidateColumn(currentCol);

            if (previousValid && previousCol != currentCol)
            {
                dataGridView1.InvalidateColumn(previousCol);
            }
        }

        //----------------------------------------------------------------------------------------
        // KeyDownイベントハンドラ
        private void dataGridView1_KeyDown(object sender, KeyEventArgs e)
        {
            int previousCol = (dataGridView1.CurrentCell != null) ? dataGridView1.CurrentCell.ColumnIndex : -1;

            isCellEdit = false; // Reset for this key press

            int keyValue;
            GridDispatchResult dispatchResult = gridKeyCommandDispatcher.Dispatch(e, contextMenuStrip1.Items, out keyValue);
            if (dispatchResult == GridDispatchResult.MenuShortcut)
            {
                return;
            }

            if (dispatchResult == GridDispatchResult.GridCommand)
            {
                e.Handled = true;
                InvalidateActiveColumnTransition(previousCol);
            }

            //初期編集状態の設定/解除
            isFirstEdit = (!isCellEdit);

            return;
        }

        //----------------------------------------------------------------------------------------
        // KeyPressイベントハンドラ
        private void dataGridView1_KeyPress(object sender, KeyPressEventArgs e)
        {
            // (なにもしない)
            return;
        }

        //----------------------------------------------------------------------------------------
        // KeyUpイベントハンドラ
        private void dataGridView1_KeyUp(object sender, KeyEventArgs e)
        {
            // (なにもしない)
            return;
        }

        //----------------------------------------------------------------------------------------
        // MouseDownイベントハンドラ
        private void dataGridView1_CellMouseDown(object sender, DataGridViewCellMouseEventArgs e)
        {
            gridMouseEventHandler.HandleMouseDown(e, selectRange, ref mouseDownPoint, ref isRectDrag, ref isCtrlDragCopy);

            // 描画更新(アクティブセルの色分けの為)
            dataGridView1.Invalidate();

            return;
        }

        //----------------------------------------------------------------------------------------
        // MouseMoveイベントハンドラ
        private void dataGridView1_CellMouseMove(object sender, DataGridViewCellMouseEventArgs e)
        {
            gridMouseEventHandler.HandleMouseMove(e, isRectDrag, dataGridView1, ref isCtrlDragCopy);
        }

        //----------------------------------------------------------------------------------------
        // MouseUpイベントハンドラ
        private void dataGridView1_CellMouseUp(object sender, DataGridViewCellMouseEventArgs e)
        {
            gridMouseEventHandler.HandleMouseUp(e, dataGridView1, ref selectRange, ref mouseDownPoint, ref isRectDrag, ref isCtrlDragCopy, ref isFirstEdit);
            return;
        }

        //----------------------------------------------------------------------------------------
        // ヘッダ編集用テキストボックスの表示と操作
        int textBoxColumn = 0;
        private void TextBox_Terminate()
        {
            // 自分自身をdataGridViewから外す
            textBox1.Visible = false;
        }

        //----------------------------------------------------------------------------------------
        // ヘッダ編集用テキストボックスのKeyDownイベントハンドラ
        private void TextBox_KeyDown(object sender, KeyEventArgs e)
        {
            //キー入力があったら、内容をチェック
            switch(e.KeyCode)
            {
                case Keys.Enter:
                    // 入力確定
                    SetHeaderValue(textBoxColumn, textBox1.Text);
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
        // ヘッダ編集用テキストボックスのLeaveイベントハンドラ
        private void TextBox_Leave(object sender, EventArgs e)
        {
            // テキストボックスからフォーカスが外れた場合、自分自身をdataGridViewから外す
            TextBox_Terminate();

        }

        //----------------------------------------------------------------------------------------
        // セルのMouseDoubleClickイベントハンドラ
        private void dataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {

            // セルがダブルクリックされた
            if (e.ColumnIndex > -1 && e.RowIndex == -1)
            {
                // ヘッダの場合 (テキストボックスを出す)
                textBox1.Text = GetHeaderValue(e.ColumnIndex);
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
        // クリップボードへのテキスト設定をリトライ付きで行う
        private void SetClipboardTextWithRetry(string text, int maxRetries = 5, int delayMs = 100)
        {
            for (int i = 0; i < maxRetries; i++)
            {
                try
                {
                    Clipboard.Clear();              // 事前に初期化したほうがよいらしい
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
        // クリップボードの利用可能状態の確認
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
        // AEへコピー
        private void AECopy(bool isDirect)
        {
            int col = dataGridView1.CurrentCell.ColumnIndex;
            string copyText = afterEffectsDataService.CreateKeyframeData(
                setting.AfterRemapVersion, setting.Fps, setting.FirstFrame, isDirect,
                GetSheetRowCount(), delegate(int row) { return GetCellValue(col, row); }, IsExcludedRow);

            //クリップボードに反映
            SetClipboardTextWithRetry(copyText);
            if (!IsClipboardAvailable())
            {
                MessageBox.Show("クリップボードにアクセスできません。");
            }
        }

        //----------------------------------------------------------------------------------------
        // AEへコピー
        private void AECopyToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // AEへコピー
            AECopy(false);

        }

        //----------------------------------------------------------------------------------------
        // AEへコピー(TimeRemap以外)
        private void directRemapToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // AEへコピー(TimeRemap以外)
            AECopy(true);

        }

        //----------------------------------------------------------------------------------------
        // AEへコピー(Script仲介)
        private void AECopyWithScript(bool isDirect)
        {
            int col = dataGridView1.CurrentCell.ColumnIndex;
            string copyText = afterEffectsDataService.CreateScriptData(
                setting.Fps, setting.FirstFrame, isDirect, GetSheetRowCount(),
                delegate(int row) { return GetCellValue(col, row); }, IsExcludedRow);

            //クリップボードに反映
            SetClipboardTextWithRetry(copyText);
            if (!IsClipboardAvailable())
            {
                MessageBox.Show("クリップボードにアクセスできません。");
            }

            //AfterEffects側の呼び出し
            Process.Start(setting.AfterPath, "-r " + setting.AfterOption);
        }

        private bool IsExcludedRow(int row)
        {
            foreach (Range range in delRange)
            {
                if (range.Top <= row && range.Bottom >= row)
                {
                    return true;
                }
            }
            return false;
        }

        //----------------------------------------------------------------------------------------
        // AEへコピー(Script仲介)
        private void jSRemapToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // AEへコピー(Script仲介)
            AECopyWithScript(false);

        }

        //----------------------------------------------------------------------------------------
        // AEへコピー(TimeRemap以外)(Script仲介)
        private void pasteFromAEToolStripMenuItem_Click(object sender, EventArgs e)
        {
            AfterEffectsPasteData pasteData;
            string errorMessage;
            if (!afterEffectsDataService.TryParseKeyframeData(Clipboard.GetText(), out pasteData, out errorMessage))
            {
                MessageBox.Show(errorMessage);
                return;
            }

            int col = dataGridView1.CurrentCell.ColumnIndex;
            setting.Fps = pasteData.Fps;
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

            List<CellWriteEntry> writes = new List<CellWriteEntry>();
            foreach (AfterEffectsKeyframe keyframe in pasteData.Keyframes)
            {
                int t = (int)Math.Round(setting.Fps * keyframe.Value);
                string writeValue = (t + setting.FirstFrame).ToString();
                string currentValue;
                string failureReason;
                if (!TryGetCellValue(col, keyframe.Frame, out currentValue, out failureReason))
                {
                    continue;
                }

                if (currentValue == writeValue)
                {
                    continue;
                }

                if (currentValue.Length == 0)
                {
                    aryCellUsedCount[col]++;
                }

                writes.Add(new CellWriteEntry(keyframe.Frame, col, writeValue));
            }

            ApplyCellWrites("AEペースト", writes);

            if (writes.Count > 0)
            {
                isFirstEdit = true;
            }

            // 描画更新(継続記号の更新の為)
            dataGridView1.Invalidate();

            // アンドゥ履歴をフラッシュ
            flushUndoHistory();
        }

        //----------------------------------------------------------------------------------------
        // 中抜き範囲の設定
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
        // 切り貼り範囲の設定
        private void setKiribariToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //切り貼り範囲を設定

            // 範囲が重なっている場合は 登録中の領域を削除
            Rect rect = getSelectedRect();
            int t = rect.Top;
            int b = rect.Bottom;
            for (int i = 0; i < addRange.Count; i++)
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
            flushUndoHistory();

            isFirstEdit = true;

            // 描画更新(切り貼り範囲反映の為)
            dataGridView1.Invalidate();
        }

        //----------------------------------------------------------------------------------------
        // 中抜き範囲の解除
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
        // 切り貼り範囲の解除
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
            flushUndoHistory();

            isFirstEdit = true;

            // 描画更新(切り貼り範囲反映の為)
            dataGridView1.Invalidate();
        }

        //----------------------------------------------------------------------------------------
        // 操作のやり直し
        private void redoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // 操作をやり直す
            redoFunction();

            isFirstEdit = true;

        }

        //----------------------------------------------------------------------------------------
        // 操作の元に戻す
        private void undoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // 操作を元に戻す
            undoFunction();

            isFirstEdit = true;
        }

        //----------------------------------------------------------------------------------------
        // アンドゥ履歴のフラッシュ
        private void flushUndoHistory()
        {
            // アンドゥ非対応機能を使用した場合などに アンドゥ履歴をフラッシュする
            gridViewManager.InitializeWork(dataGridView1, timingSheetModel);
        }

        //----------------------------------------------------------------------------------------
        // コピー
        private void copyToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // コピー
            copyRect = getSelectedRect();
            copyToBuf(copyRect);
            isFirstEdit = true;
        }

        //----------------------------------------------------------------------------------------
        // 切り取り
        private void cutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // 切り取り
            copyRect = getSelectedRect();
            cutToBuf(copyRect, false);
            isFirstEdit = true;

            // 描画更新(継続記号の更新の為)
            dataGridView1.Invalidate();
        }

        //----------------------------------------------------------------------------------------
        // 貼り付け
        private void pasteToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // 貼り付け
            int col = dataGridView1.CurrentCell.ColumnIndex;
            int row = dataGridView1.CurrentCell.RowIndex;
            copyToCell(col, row, false);
            isFirstEdit = true;

            // 描画更新(継続記号の更新の為)
            dataGridView1.Invalidate();
        }

        //----------------------------------------------------------------------------------------
        // フレームレート 30FPS
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
        // フレームレート 24FPS
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
        // 常に手前に表示
        private void stayOnTopToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // TopMost設定の切り替え
            this.TopMost = !this.TopMost;
            setting.TopMost = this.TopMost;
            stayOnTopToolStripMenuItem.Checked = this.TopMost;
        }

        //----------------------------------------------------------------------------------------
        // カラセル入力時の移動抑止
        private void karacellNoMoveToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // カラセル入力時の移動抑止設定の切り替え
            setting.IsKaraNoMove = !setting.IsKaraNoMove;
            karacellNoMoveToolStripMenuItem.Checked = setting.IsKaraNoMove;
        }

        //----------------------------------------------------------------------------------------
        // フレーム数表示⇔シート/コマ数表示 切り替え
        private void displayFrameNumberToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // フレーム数表示⇔シート/コマ数表示 切り替え
            setting.IsDisplayFrameNumber = !setting.IsDisplayFrameNumber;
            this.displayFrameNumberToolStripMenuItem.Checked = setting.IsDisplayFrameNumber;

            // 描画更新(フレーム表示更新の為)
            dataGridView1.Invalidate();
        }

        //----------------------------------------------------------------------------------------
        // 開始フレーム数の設定
        private void firstFrameToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // 開始フレーム設定
            InputBox1 dialog = new InputBox1("開始フレーム数");
            dialog.LabelName1 = "開始フレーム数";
            dialog.Value1 = setting.FirstFrame.ToString();
            if (dialog.ShowDialog(this.owner) == System.Windows.Forms.DialogResult.OK)
            {
                if (int.TryParse(dialog.Value1, out int frame))
                {
                    setting.FirstFrame = frame;
                } else {
                    MessageBox.Show("入力された値を数値に変換できませんでした.");
                }
            }
        }

        //----------------------------------------------------------------------------------------
        // カラセル文字列の設定
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
        // シートの秒数の設定
        private void secondsPerSheetToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // シートの秒数
            InputBox1 dialog = new InputBox1("シートの秒数");
            dialog.LabelName1 = "シートの秒数";
            dialog.Value1 = setting.SheetSec.ToString();
            if (dialog.ShowDialog(this.owner) == System.Windows.Forms.DialogResult.OK)
            {
                if (int.TryParse(dialog.Value1, out int sec))
                {
                    setting.SheetSec = sec;
                } else {
                    MessageBox.Show("入力された値を数値に変換できませんでした.");
                }
            }
        }

        //----------------------------------------------------------------------------------------
        // シートの基準線指定: 4コマ毎
        private void div4ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // シートの基準線指定: 4コマ毎
            setting.SheetDivide = 4;
            this.div4ToolStripMenuItem.Checked = true;
            this.div6ToolStripMenuItem.Checked = false;
            this.div12ToolStripMenuItem.Checked = false;

        }

        //----------------------------------------------------------------------------------------
        // シートの基準線指定: 6コマ毎
        private void div6ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // シートの基準線指定: 6コマ毎
            setting.SheetDivide = 6;
            this.div4ToolStripMenuItem.Checked = false;
            this.div6ToolStripMenuItem.Checked = true;
            this.div12ToolStripMenuItem.Checked = false;

        }

        //----------------------------------------------------------------------------------------
        // シートの基準線指定: 12コマ毎
        private void div12ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // シートの基準線指定: 12コマ毎
            setting.SheetDivide = 12;
            this.div4ToolStripMenuItem.Checked = false;
            this.div6ToolStripMenuItem.Checked = false;
            this.div12ToolStripMenuItem.Checked = true;

        }

        //----------------------------------------------------------------------------------------
        // AfterFX.exe パス設定
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
        // リマップ用 jsx パス設定
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
        // セル変更時にウィンドウ位置調整する 設定
        private void autoAdjustToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // セル変更時にウィンドウ位置調整する 設定
            setting.IsAutoadjust = !setting.IsAutoadjust;
            autoAdjustToolStripMenuItem.Checked = setting.IsAutoadjust;
        }

        //----------------------------------------------------------------------------------------
        // 作業情報の全初期化
        private void allInitializeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // 作業情報の全初期化
            InitializeWork(true);
        }

        //----------------------------------------------------------------------------------------
        // セル枚数の指定
        private void inputCellCountToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // セル枚数の指定
            InputBox1 dialog = new InputBox1("セル枚数の指定");
            int newCount = 0;
            dialog.LabelName1 = "セル枚数";
            dialog.Value1 = setting.ColLength.ToString();
            if (dialog.ShowDialog(this.owner) == System.Windows.Forms.DialogResult.OK)
            {
                if (int.TryParse(dialog.Value1, out int count))
                {
                    newCount = count;
                } else {
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
            flushUndoHistory();

            // コピーバッファをフラッシュ
            InitializeWork(InitializeTarget.CopyBuffer);
        }

        //----------------------------------------------------------------------------------------
        // セルの挿入
        private void insertCellToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // セルの挿入

            // カレントセルの位置を保存
            int col = dataGridView1.CurrentCell.ColumnIndex;
            int originalColumnCount = GetSheetColumnCount();

            // グリッドサイズを変更
            resizeDataGridView1(setting.ColLength + 1, setting.RowLength);

            // ウィンドウ位置調整
            adjustWindowSize();

            // resize が旧列を同じ index に復元した後で snapshot を作り、反映する。
            int rowCount = GetSheetRowCount();
            IList<ColumnWriteEntry> writes = SheetColumnEditCalculator.CreateInsertColumn(
                rowCount, originalColumnCount, col,
                delegate(int row, int column) { return GetCellValue(column, row); },
                delegate(int column) { return GetHeaderValue(column); },
                delegate(int column) { return aryCellUsedCount[column]; });
            ApplyColumnWrites(writes, rowCount);

            // アンドゥ履歴をフラッシュ
            flushUndoHistory();

            // コピーバッファをフラッシュ
            InitializeWork(InitializeTarget.CopyBuffer);
        }

        //----------------------------------------------------------------------------------------
        // セルの削除
        private void deleteCellToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // セルの削除

            // カレントセルの位置を詰める
            int col = dataGridView1.CurrentCell.ColumnIndex;
            int rowCount = GetSheetRowCount();
            int columnCount = GetSheetColumnCount();

            // 削除対象を resize で失う前に snapshot を作り、左詰めを反映する。
            IList<ColumnWriteEntry> writes = SheetColumnEditCalculator.CreateDeleteColumn(
                rowCount, columnCount, col,
                delegate(int row, int column) { return GetCellValue(column, row); },
                delegate(int column) { return GetHeaderValue(column); },
                delegate(int column) { return aryCellUsedCount[column]; });
            ApplyColumnWrites(writes, rowCount);

            // グリッドサイズを変更
            resizeDataGridView1(setting.ColLength - 1, setting.RowLength);

            // ウィンドウ位置調整
            adjustWindowSize();

            // アンドゥ履歴をフラッシュ
            flushUndoHistory();

            // コピーバッファをフラッシュ
            InitializeWork(InitializeTarget.CopyBuffer);
        }

        //----------------------------------------------------------------------------------------
        // 連番作成
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
                Dictionary<string, string> parameters = new Dictionary<string, string>();
                parameters.Add(SequentialNumberCommand.StartParameter, dialog.Value1);
                parameters.Add(SequentialNumberCommand.StepParameter, dialog.Value2);
                parameters.Add(SequentialNumberCommand.SkipParameter, dialog.CheckValue1.ToString());
                ExecuteAutomationCommand(SequentialNumberCommand.CommandId, parameters);
            }
        }

        //----------------------------------------------------------------------------------------
        // 連番作成(複数列、挿入番号、スキップ、ループ対応)
        private void HandleRepeatInput(RepeatInputValues values)
        {
            Dictionary<string, string> parameters = new Dictionary<string, string>();
            parameters.Add(RepeatNumberCommand.StartParameter, values.Start);
            parameters.Add(RepeatNumberCommand.EndParameter, values.End);
            parameters.Add(RepeatNumberCommand.RowIntervalParameter, values.RowInterval);
            parameters.Add(RepeatNumberCommand.LoopParameter, values.Loop);
            parameters.Add(RepeatNumberCommand.SkipParameter, values.Skip);
            parameters.Add(RepeatNumberCommand.InsertParameter, values.Insert);
            AutomationHostResult result = _repeatAutomationSession.Execute(RepeatNumberCommand.CommandId,
                CreateAutomationRequest(parameters), this);
            if (!result.Succeeded) MessageBox.Show(result.Error);
            else FinishWriteOperation(true);
        }

        private void HandleRepeatInputClosed(object sender, FormClosedEventArgs e)
        {
            RepeatInputBox dialog = sender as RepeatInputBox;
            if (dialog != null)
            {
                dialog.OnRepeatInput -= HandleRepeatInput;
                dialog.FormClosed -= HandleRepeatInputClosed;
            }
            if (Object.ReferenceEquals(_repeatInputDialog, dialog)) _repeatInputDialog = null;
            if (_repeatAutomationSession != null) _repeatAutomationSession.Close();
            _repeatAutomationSession = null;
        }

        //----------------------------------------------------------------------------------------
        // 連番作成(複数列、挿入番号、スキップ、ループ対応)
        private void repeatNumberToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (_repeatInputDialog != null && !_repeatInputDialog.IsDisposed)
            {
                _repeatInputDialog.Activate();
                return;
            }

            // 繰り返しダイアログの表示
            _repeatAutomationSession = new AutomationSession(automationHost);
            _repeatInputDialog = new RepeatInputBox();
            _repeatInputDialog.OnRepeatInput += HandleRepeatInput;
            _repeatInputDialog.FormClosed += HandleRepeatInputClosed;
            _repeatInputDialog.Show();
        }

        //----------------------------------------------------------------------------------------
        // 置換
        private void replaceToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //置換
            InputBox1 dialog = new InputBox1("置換");
            dialog.LabelName1 = "置換前";
            dialog.LabelName2 = "置換後";
            if (dialog.ShowDialog(this.owner) == System.Windows.Forms.DialogResult.OK)
            {
                Dictionary<string, string> parameters = new Dictionary<string, string>();
                parameters.Add(ReplaceCommand.BeforeParameter, dialog.Value1);
                parameters.Add(ReplaceCommand.AfterParameter, dialog.Value2);
                ExecuteAutomationCommand(ReplaceCommand.CommandId, parameters);
            }
        }

        //----------------------------------------------------------------------------------------
        // 反転
        private void reverseToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ExecuteAutomationCommand(ReverseCommand.CommandId, new Dictionary<string, string>());
        }

        //----------------------------------------------------------------------------------------
        // 四則演算
        private void fourArithmeticOperationToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //四則演算

            InputBox1 dialog = new InputBox1("四則演算");
            dialog.LabelName1 = "四則演算";
            if (dialog.ShowDialog(this.owner) == System.Windows.Forms.DialogResult.OK)
            {
                String value = dialog.Value1.Trim();
                string operation = value.Length == 0 ? String.Empty : value.Substring(0, 1);
                string operand = value.Length < 2 ? String.Empty : value.Substring(1);
                Dictionary<string, string> parameters = new Dictionary<string, string>();
                parameters.Add(ArithmeticCommand.OperatorParameter, operation);
                parameters.Add(ArithmeticCommand.OperandParameter, operand);
                ExecuteAutomationCommand(ArithmeticCommand.CommandId, parameters);
            }
        }

        //----------------------------------------------------------------------------------------
        // 複製
        private void duplicateToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // 複製

            int Col, Row, Len;
            Col = selectRange.Left;
            Row = selectRange.Top;
            Len = selectRange.Height;

            // 貼り付け先の末尾がシート行数を超える場合は、範囲内だけ複製する
            int maxCopyLength = GetSheetRowCount() - (Row + Len);
            if (maxCopyLength <= 0)
            {
                return;
            }

            int copyLength = (Len < maxCopyLength) ? Len : maxCopyLength;

            List<CellWriteEntry> writes = new List<CellWriteEntry>();
            //複製操作
            for (int i = 0; i < copyLength; i++)
            {
                String val = GetCellValue(Col, Row + i);
                writes.Add(new CellWriteEntry(Row + Len + i, Col, val));
                //(*pColorBuf)[Col][Row + Len + i] = versionNumber;
            }

            ApplyCellWrites("複製", writes);

            // 新しい選択範囲を設定
            Rect rect = selectRange;
            dataGridView1.ClearSelection();
            rect.Y += Len;
            rect.Height = copyLength;
            for (int i = 0; i < rect.Height; i++)
                for (int j = 0; j < rect.Width; j++)
                    dataGridView1[rect.X + j, rect.Y + i].Selected = true;

            // 範囲を保存
            selectRange = rect;

            FinishWriteOperation(true);
        }

        //----------------------------------------------------------------------------------------
        // STS 保存
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
                try
                {
                    stsFileService.Save(dialog.FileName, GetSheetColumnCount(), GetSheetRowCount(),
                        GetCellValue, GetHeaderValue);
                }
                catch (Exception)
                {
                    MessageBox.Show("指定ファイルを開けませんでした.");
                }
            }

        }

        //----------------------------------------------------------------------------------------
        // STS 読み込み
        private void loadSTS(String path)
        {
            StsDocument document;
            try
            {
                document = stsFileService.Load(path);
            }
            catch (InvalidDataException)
            {
                MessageBox.Show("未対応のファイルの為、開けませんでした.");
                return;
            }
            catch (Exception)
            {
                MessageBox.Show("指定ファイルを開けませんでした.");
                return;
            }

            // グリッドサイズをデータに合わせる
            setting.ColLength = document.ColumnCount;
            setting.RowLength = document.RowCount;
            InitializeWork(false);

            foreach (StsCellChange change in document.CellChanges)
            {
                SetCellValueIfChanged(change.Column, change.Row, change.Value);
            }

            for (int i = 0; i < document.Headers.Count; i++)
            {
                SetHeaderValue(i, document.Headers[i]);
            }

            // 読み込み結果は確定状態とし、Undo履歴をクリアする（旧実装互換）。
            flushUndoHistory();
        }

        //----------------------------------------------------------------------------------------
        // STS 読み込み(メニューから)
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
        // セル列のヘッダがクリックされた場合に、１列選択する（ダブルクリックが名称編集なので、クリックで全選択に）
        private void dataGridView1_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            // セル列のヘッダがクリックされた場合に、１列選択する（ダブルクリックが名称編集なので、クリックで全選択に）

            // 選択をクリア
            dataGridView1.ClearSelection();

            // 選択した列の最初のセルに移動
            dataGridView1.CurrentCell = dataGridView1[e.ColumnIndex, 0];

            // 列全体を選択
            for (int row = 0; row < dataGridView1.RowCount; row++)
            {
                dataGridView1[e.ColumnIndex, row].Selected = true;
            }
        }

        //----------------------------------------------------------------------------------------
        // フォーム表示後の処理
        private void Form1_Shown(object sender, EventArgs e)
        {
            // コントロールにフォーカスを設定
            dataGridView1.Focus();

        }
    }

}
