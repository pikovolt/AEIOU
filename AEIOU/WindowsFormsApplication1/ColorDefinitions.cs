using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;

namespace AEIOU
{
    class ColorDefinitions
    {
        public ColorDefinitions()
        {
        }

        // グリッドの色設定
        public Color BgCell1 = Color.FromArgb(255, 255, 255, 255);  // 縞々BG1(  未入力  )
        public Color BgCell2 = Color.FromArgb(255, 250, 250, 250);  //       2(    〃    )
        public Color BgCell1R = Color.FromArgb(255, 245, 245, 255); // 縞々BG1(  入力済  )
        public Color BgCell2R = Color.FromArgb(255, 235, 235, 255); //       2(    〃    )
        public Color BgCell1A = Color.FromArgb(255, 235, 255, 235); //       1(アクティブ)
        public Color BgCell2A = Color.FromArgb(255, 225, 250, 225); //       2(　　〃　　)
        public Color Selected = Color.FromArgb(255, 170, 255, 214); // 選択範囲色
        public Color Nakanuki = Color.FromArgb(255, 220, 220, 220); // 中ヌキ色
        public Color Harikomi = Color.FromArgb(255, 255, 230, 230); // 張り込み色
        public Color Header = Color.FromArgb(255, 244, 244, 255);   // ヘッダ色
    }
}
