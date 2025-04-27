'''
Python3系 + PySide2モジュール利用
セル入力の基礎的な実装サンプル（'25.04.27 by kanbara with GitHub Copilot.）

Pythonを使って、ツールを自作する際の参考にどうぞ。
VSCodeの GitHub Copilotに出して貰ったので、Copilotに質問をすれば、実装の動作内容や修正案など、いろいろ教えてくれると思います。
'''
from PySide2.QtWidgets import QApplication, QTableWidget, QTableWidgetItem, QTableWidgetSelectionRange
from PySide2.QtCore import Qt
import sys

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

        # キーごとの処理をディスパッチ (キーごとの動作は、別関数に分離)
        key_handlers = {
            Qt.Key_0: self.handle_numeric_key,
            Qt.Key_1: self.handle_numeric_key,
            Qt.Key_2: self.handle_numeric_key,
            Qt.Key_3: self.handle_numeric_key,
            Qt.Key_4: self.handle_numeric_key,
            Qt.Key_5: self.handle_numeric_key,
            Qt.Key_6: self.handle_numeric_key,
            Qt.Key_7: self.handle_numeric_key,
            Qt.Key_8: self.handle_numeric_key,
            Qt.Key_9: self.handle_numeric_key,
            Qt.Key_Return: self.handle_enter_key,
            Qt.Key_Enter: self.handle_enter_key,
            Qt.Key_Up: self.handle_arrow_key,
            Qt.Key_Down: self.handle_arrow_key,
            Qt.Key_Backspace: self.handle_backspace_key,
            Qt.Key_Asterisk: self.handle_asterisk_key,
            Qt.Key_Slash: self.handle_slash_key,
        }

        handler = key_handlers.get(event.key(), None)
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
                        self.setRangeSelected(
                            QTableWidgetSelectionRange(row, left_column, row + selected_row_count - 1, right_column), True
                        )
                        return

            # 入力がない場合、選択範囲を単に戻す
            new_top_row = max(0, top_row - selected_row_count)
            new_bottom_row = new_top_row + selected_row_count - 1
            self.clearSelection()
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

if __name__ == "__main__":
    app = QApplication(sys.argv)
    grid_view = GridView(100, 3)    # 100行 x 3列 グリッドビュー
    grid_view.resize(350, 400)      # サイズを設定
    grid_view.show()
    sys.exit(app.exec_())
