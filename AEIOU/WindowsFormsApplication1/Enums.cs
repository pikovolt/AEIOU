using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace AEIOU
{
    // 基準線の状態 を表す列挙体
    // (パクリ元)⇒http://smdn.jp/programming/netfx/enum/1_flags/
    [Flags]
    enum SheetBorder : int
    {
        EveryNFrames = 0x00000001,  //ｎコマ毎
        HarfSec = 0x00000002,      //0.5秒毎
        EverySec = 0x00000004,      //１秒毎
        EverySheet = 0x00000008,    //シート毎
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

}
