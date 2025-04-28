from PySide6.QtWidgets import QApplication, QTableWidget, QTableWidgetItem, QTableWidgetSelectionRange, QAbstractItemView
from PySide6.QtCore import Qt
import sys
import pyperclip  # クリップボード操作用ライブラリ

class GridView(QTableWidget):
    def __init__(self, rows, columns):
        super().__init__(rows, columns)
        self.init_ui()
        self.is_editing = False  # フラグ: セルで数値入力中かどうか
        self.shift_base_row = None  # Shiftキー押下時の基準行

    def init_ui(self):
        self.setWindowTitle("Excel-like Grid View")
        self.setEditTriggers(QTableWidget.NoEditTriggers)  # 編集モードを無効化
        self.setFocusPolicy(Qt.StrongFocus)
        self.setSelectionMode(QTableWidget.ContiguousSelection)  # 一般的な範囲選択方式に変更

        # 選択範囲の色を変更
        self.setStyleSheet("""
            QTableWidget::item:selected {
                background-color: #77CCAA;  /* 緑色 */
                color: black;              /* 文字色を黒に */
            }
        """)
        self.setCurrentCell(0, 0)

        # キーとハンドラ関数のマッピングを定義
        self.key_handlers = {
            (Qt.Key_0, Qt.NoModifier): self.handle_numeric_key,
            (Qt.Key_1, Qt.NoModifier): self.handle_numeric_key,
            (Qt.Key_2, Qt.NoModifier): self.handle_numeric_key,
            (Qt.Key_3, Qt.NoModifier): self.handle_numeric_key,
            (Qt.Key_4, Qt.NoModifier): self.handle_numeric_key,
            (Qt.Key_5, Qt.NoModifier): self.handle_numeric_key,
            (Qt.Key_6, Qt.NoModifier): self.handle_numeric_key,
            (Qt.Key_7, Qt.NoModifier): self.handle_numeric_key,
            (Qt.Key_8, Qt.NoModifier): self.handle_numeric_key,
            (Qt.Key_9, Qt.NoModifier): self.handle_numeric_key,
            (Qt.Key_Return, Qt.NoModifier): self.handle_enter_key,
            (Qt.Key_Enter, Qt.NoModifier): self.handle_enter_key,
            (Qt.Key_Up, Qt.NoModifier): self.handle_arrow_key,
            (Qt.Key_Down, Qt.NoModifier): self.handle_arrow_key,
            (Qt.Key_Backspace, Qt.NoModifier): self.handle_backspace_key,
            (Qt.Key_Asterisk, Qt.NoModifier): self.handle_asterisk_key,
            (Qt.Key_Slash, Qt.NoModifier): self.handle_slash_key,
            (Qt.Key_O, Qt.ControlModifier): self.handle_ctrl_o_key,  # Ctrl+O
        }

    def keyPressEvent(self, event):
        """
        キーボード入力イベントを処理する関数。
        ユーザーがキーボードで特定のキーを押した際に呼び出され、押されたキーに応じて適切な処理を実行する。

        処理の流れ:
        1. 現在選択されているセルの範囲を取得する。
           - 選択範囲がない場合は、親クラスの keyPressEvent を呼び出して終了。
        2. 選択範囲の行・列情報を取得し、選択範囲の行数を計算する。
        3. 押されたキーに応じて、対応する処理関数（ハンドラ）を呼び出す。
           - ハンドラが存在しない場合は、親クラスの keyPressEvent を呼び出す。

        補足:
        - `is_editing` フラグ:
          セルが現在編集中かどうかを管理するフラグ。
          編集中の場合、数値入力や Backspace キーの動作が異なる。
        - `shift_base_row`:
          Shiftキーを押しながら選択範囲を拡張する際の基準行を保持する。
        """
        selected_ranges = self.selectedRanges()
        if not selected_ranges:
            super().keyPressEvent(event)
            return

        # 選択範囲の情報を取得
        selection = selected_ranges[0]
        top_row, bottom_row = selection.topRow(), selection.bottomRow()
        left_column, right_column = selection.leftColumn(), selection.rightColumn()
        selected_row_count = bottom_row - top_row + 1

        # キーと修飾キーの組み合わせを取得
        key = event.key()
        modifiers = event.modifiers()

        handler = self.key_handlers.get((key, modifiers))
        if handler:
            handler(event, top_row, bottom_row, left_column, right_column, selected_row_count)
        else:
            super().keyPressEvent(event)

    def handle_numeric_key(self, event, top_row, bottom_row, left_column, right_column, selected_row_count):
        """
        数値キーが押されたときの処理。
        - 非編集状態の場合、セルの内容をクリアして編集状態にする。
        - 押されたキーの数値をセルのテキストに追加する。

        ポイント:
        - `is_editing` フラグを利用して編集状態を管理。
        """
        current_item = self.item(top_row, left_column)
        if not current_item:
            current_item = QTableWidgetItem("")
            self.setItem(top_row, left_column, current_item)

        if not self.is_editing:  # 非編集状態の場合
            current_item.setText("")  # セルの内容をクリア
            self.is_editing = True

        current_text = current_item.text()
        new_text = current_text + event.text()  # 入力を追加
        current_item.setText(new_text)

    def handle_enter_key(self, event, top_row, bottom_row, left_column, right_column, selected_row_count):
        """
        Enterキーが押されたときの処理。
        - 編集状態を解除し、選択範囲を次の行に移動。
        - 必要に応じて新しい行を追加。

        ポイント:
        - 選択範囲の行数を維持しながら移動。
        """
        self.is_editing = False  # 編集状態を解除
        new_top_row = bottom_row + 1
        new_bottom_row = new_top_row + selected_row_count - 1
        if new_bottom_row >= self.rowCount():
            for _ in range(new_bottom_row - self.rowCount() + 1):
                self.insertRow(self.rowCount())  # 必要に応じて行を追加
        self.clearSelection()
        # 選択範囲を設定 (表示範囲に収めるようにする)
        self.setCurrentCell(new_bottom_row, left_column)
        self.setCurrentCell(new_top_row, left_column)
        self.setRangeSelected(
            QTableWidgetSelectionRange(new_top_row, left_column, new_bottom_row, right_column), True
        )

    def handle_arrow_key(self, event, top_row, bottom_row, left_column, right_column, selected_row_count):
        """
        上下の矢印キーが押されたときの処理。
        - 編集状態を解除し、選択範囲を上下に移動。
        - Shiftキーが押されている場合、選択範囲を拡張。

        ポイント:
        - `shift_base_row` を利用してShiftキーによる範囲選択を管理。
        """
        self.is_editing = False  # 選択範囲を移動した際に編集状態をリセット
        if event.modifiers() & Qt.ShiftModifier:  # Shiftキーが押されている場合
            if self.shift_base_row is None:
                self.shift_base_row = bottom_row

            new_bottom_row = (
                max(top_row, bottom_row - 1) if event.key() == Qt.Key_Up else
                min(self.rowCount() - 1, bottom_row + 1)
            )
            self.clearSelection()
            # 選択範囲を設定 (表示範囲に収めるようにする)
            self.setCurrentCell(new_bottom_row, left_column)
            self.setCurrentCell(top_row, left_column)
            self.setRangeSelected(
                QTableWidgetSelectionRange(top_row, left_column, new_bottom_row, right_column), True
            )
            self.shift_base_row = new_bottom_row
        else:  # Shiftキーが押されていない場合
            self.shift_base_row = None
            new_top_row = (
                max(0, top_row - selected_row_count) if event.key() == Qt.Key_Up else
                min(self.rowCount() - selected_row_count, top_row + selected_row_count)
            )
            new_bottom_row = new_top_row + selected_row_count - 1
            self.clearSelection()
            # 選択範囲を設定 (表示範囲に収めるようにする)
            self.setCurrentCell(new_bottom_row, left_column)
            self.setCurrentCell(new_top_row, left_column)
            self.setRangeSelected(
                QTableWidgetSelectionRange(new_top_row, left_column, new_bottom_row, right_column), True
            )

    def handle_backspace_key(self, event, top_row, bottom_row, left_column, right_column, selected_row_count):
        """
        Backspaceキーが押されたときの処理。
        - 編集状態の場合、セルのテキストを1文字削除。
        - 非編集状態の場合、選択範囲内のセルの内容をすべて削除。

        ポイント:
        - 編集状態と非編集状態で動作が異なる。
        """
        if self.is_editing:  # 数値入力中の場合
            current_item = self.item(top_row, left_column)
            if current_item:
                current_text = current_item.text()
                if current_text:
                    current_item.setText(current_text[:-1])  # 下1桁を削除
                if not current_item.text():  # 全て削除された場合、編集状態を解除
                    self.is_editing = False
        else:  # 非数値入力状態の場合
            # 選択範囲内のセルの内容を削除
            for row in range(top_row, bottom_row + 1):
                for col in range(left_column, right_column + 1):
                    current_item = self.item(row, col)
                    if current_item:
                        current_item.setText("")  # セルの内容を削除

            # 遡って入力があるセルを探す
            for row in range(top_row - 1, -1, -1):
                for col in range(left_column, right_column + 1):
                    current_item = self.item(row, col)
                    if current_item and current_item.text():  # 入力があるセルを発見
                        self.clearSelection()
                        # 選択範囲を設定 (表示範囲に収めるようにする)
                        bottom_row = row + selected_row_count - 1
                        self.setCurrentCell(bottom_row, left_column)
                        self.setCurrentCell(row, left_column)
                        self.setRangeSelected(
                            QTableWidgetSelectionRange(row, left_column, bottom_row, right_column), True
                        )
                        return

            # 入力がない場合、選択範囲を単に戻す
            new_top_row = max(0, top_row - selected_row_count)
            new_bottom_row = new_top_row + selected_row_count - 1
            self.clearSelection()
            # 選択範囲を設定 (表示範囲に収めるようにする)
            self.setCurrentCell(new_bottom_row, left_column)
            self.setCurrentCell(new_top_row, left_column)
            self.setRangeSelected(
                QTableWidgetSelectionRange(new_top_row, left_column, new_bottom_row, right_column), True
            )

    def handle_asterisk_key(self, event, top_row, bottom_row, left_column, right_column, selected_row_count):
        """
        アスタリスク（*）キーが押されたときの処理。
        - 選択範囲を1行下に拡張。

        ポイント:
        - 行数を超えないように制限。
        """
        new_bottom_row = min(self.rowCount() - 1, bottom_row + 1)
        self.clearSelection()
        self.setRangeSelected(
            QTableWidgetSelectionRange(top_row, left_column, new_bottom_row, right_column), True
        )

    def handle_slash_key(self, event, top_row, bottom_row, left_column, right_column, selected_row_count):
        """
        スラッシュ（/）キーが押されたときの処理。
        - 選択範囲を1行上に縮小。

        ポイント:
        - 選択範囲の上限を超えないように制限。
        """
        new_bottom_row = max(top_row, bottom_row - 1)
        self.clearSelection()
        self.setRangeSelected(
            QTableWidgetSelectionRange(top_row, left_column, new_bottom_row, right_column), True
        )

    def handle_ctrl_o_key(self, event, top_row, bottom_row, left_column, right_column, selected_row_count):
        """
        Ctrl+Oキーが押されたときの処理。
        - タイムリマップ文字列の生成を行う
        - カレントの列のセルの行を frame, セルの入力値を valueとして、辞書を生成
        - calc_timingで辞書を生成, copy_timingでタイムリマップ情報を生成して、クリップボードにコピーする
        """
        keys = {}
        for row in range(self.rowCount()):  # 全行を対象に処理
            item = self.item(row, left_column)
            if item and item.text().isdigit():
                frame = row     # 行番号をフレーム番号に変換 (0始まりとする)
                value = int(item.text())
                keys[frame] = value

        if not keys:
            print("選択範囲に有効なデータがありません。")
            return

        try:
            # タイムリマップ情報を計算
            calced_keys = calc_timing(keys)
            # タイムリマップ情報をクリップボードにコピー
            copy_timing(calced_keys)
        except ValueError as e:
            print(f"エラー: {e}")


def calc_timing(keys, fps=24, start_frame=1):
    """
    キー情報を計算する
    - keys = {frame: value, ..}, frame = フレーム数(0~), value = セル番号(start_frame~)
    - start_frame = 0 or 1
    - 上記の値から、{frame: (value - start_frame) / fps} の辞書を生成
    """
    if not isinstance(keys, dict):
        raise ValueError("keys must be a dictionary with frame: value pairs")
    if fps <= 0:
        raise ValueError("fps must be a positive number")

    # 辞書内包表記を使って計算
    return {frame: (value - start_frame) / fps for frame, value in keys.items()}


def copy_timing(calced_keys, version="8.0"):
    """
    AfterEffects向けのタイムリマップ情報を生成して、クリップボードにコピーする
    - calced_keys = {frame: 秒数}
    - 上記の値から、 AfterEffects向けのタイムリマップ情報を生成
    - 生成したタイムリマップ情報をクリップボードにコピー
    - version: ヘッダー文字列のバージョンを指定 (デフォルトは "8.0")
    """
    if not isinstance(calced_keys, dict):
        raise ValueError("calced_keys must be a dictionary with frame: seconds pairs")

    # ヘッダー部分
    keyframe_data = [
        f"Adobe After Effects {version} Keyframe Data",  # バージョンを指定
        "",
        "\tUnits Per Second\t24",
        "\tSource Width\t640",
        "\tSource Height\t480",
        "\tSource Pixel Aspect Ratio\t1",
        "\tComp Pixel Aspect Ratio\t1",
        "",
        "Time Remap",
        "\tFrame\tseconds"
    ]

    # キーフレームデータ部分
    for frame, seconds in sorted(calced_keys.items()):
        keyframe_data.append(f"\t{frame}\t{seconds:.6f}")

    # フッター部分
    keyframe_data.append("")
    keyframe_data.append("End of Keyframe Data")

    # リストを文字列に結合
    result = "\n".join(keyframe_data)

    # クリップボードにコピー
    pyperclip.copy(result)
    print("タイムリマップ情報をクリップボードにコピーしました。")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    grid_view = GridView(100, 3)    # 100行 x 3列 グリッドビュー
    grid_view.resize(350, 400)      # サイズを設定
    grid_view.show()
    sys.exit(app.exec())
