"""Google Sheetsへのエクスポート処理（統合版）"""
import gspread
from google.oauth2.service_account import Credentials
from typing import List, Dict
from data_structures import RoundData, PlayerData
from config import CREDENTIAL_FILE, SPREADSHEET_ID, load_fans, DORA_FANS, ORIGIN_POINT_3, ORIGIN_POINT_4
import time

class SheetsExporter:
    def __init__(self):
        # Google Sheets APIの認証
        scope = [
            'https://spreadsheets.google.com/feeds',
            'https://www.googleapis.com/auth/spreadsheets',
            'https://www.googleapis.com/auth/drive.file',
            'https://www.googleapis.com/auth/drive'
        ]

        creds = Credentials.from_service_account_file(str(CREDENTIAL_FILE), scopes=scope)
        self.client = gspread.authorize(creds)
        self.spreadsheet = self.client.open_by_key(SPREADSHEET_ID)
        self.fan_names = load_fans()

    def clean_all_sheets(self):
        """すべての既存シートを削除（デフォルトシート以外）"""
        try:
            # 現在のすべてのシートを取得
            worksheets = self.spreadsheet.worksheets()

            # 削除対象のパターン
            patterns_to_delete = [
                "【四麻】", "【三麻】",
                "プレイヤーデータ", "総合結果",
                "Match", "_4P", "_3P"
            ]

            sheets_to_delete = []
            for ws in worksheets:
                # 削除対象のパターンにマッチするシートを探す
                if any(pattern in ws.title for pattern in patterns_to_delete):
                    sheets_to_delete.append(ws.title)

            # シートを削除
            for sheet_name in sheets_to_delete:
                try:
                    worksheet = self.spreadsheet.worksheet(sheet_name)
                    self.spreadsheet.del_worksheet(worksheet)
                    print(f"  Deleted sheet: {sheet_name}")
                except Exception as e:
                    print(f"  Could not delete sheet {sheet_name}: {e}")
                time.sleep(2)  # API制限対策

            if sheets_to_delete:
                print(f"  Total {len(sheets_to_delete)} sheets deleted")
            else:
                print("  No sheets to delete")

        except Exception as e:
            print(f"Warning: Could not clean sheets: {e}")

    def clean_mahjong_sheets(self, player_n: int):
        """指定された人数の既存シートを削除"""
        try:
            # 現在のすべてのシートを取得
            worksheets = self.spreadsheet.worksheets()

            # 削除対象のパターン
            if player_n == 4:
                patterns_to_delete = ["【四麻】", "Match", "_4P"]
            else:
                patterns_to_delete = ["【三麻】", "_3P"]

            sheets_to_delete = []
            for ws in worksheets:
                # 削除対象のパターンにマッチするシートを探す
                if any(pattern in ws.title for pattern in patterns_to_delete):
                    sheets_to_delete.append(ws.title)

            # シートを削除
            for sheet_name in sheets_to_delete:
                try:
                    worksheet = self.spreadsheet.worksheet(sheet_name)
                    self.spreadsheet.del_worksheet(worksheet)
                    print(f"  Deleted sheet: {sheet_name}")
                except Exception as e:
                    print(f"  Could not delete sheet {sheet_name}: {e}")
                time.sleep(2)  # API制限対策

            if sheets_to_delete:
                print(f"  Total {len(sheets_to_delete)} {'四麻' if player_n == 4 else '三麻'} sheets deleted")
            else:
                print(f"  No {'四麻' if player_n == 4 else '三麻'} sheets to delete")

        except Exception as e:
            print(f"Warning: Could not clean {'四麻' if player_n == 4 else '三麻'} sheets: {e}")

    def export_round_sheet(self, round_data: RoundData, sheet_name: str, player_n: int, player_data_dict: Dict[str, PlayerData]):
        """半荘のデータをシートに出力"""
        try:
            worksheet = self.spreadsheet.worksheet(sheet_name)
        except gspread.exceptions.WorksheetNotFound:
            worksheet = self.spreadsheet.add_worksheet(title=sheet_name, rows=200, cols=26)

        # データを準備
        all_values = []

        # 牌譜リンク
        paifu_url = f'https://game.mahjongsoul.com/?paipu={round_data.uuid}'
        row1 = [f'=HYPERLINK("{paifu_url}", "牌譜")'] + [''] * (player_n + 4)
        all_values.append(row1)

        # 方角行
        directions = ["東", "南", "西", "北"][:player_n]
        row2 = ["", "", ""] + directions + ["(供託)", "(和了詳細)"]
        all_values.append(row2)

        # 空行
        row3 = [""] * (player_n + 5)
        all_values.append(row3)

        # HN行とプレイヤー名行
        row4 = ["", "", "HN"] + round_data.names + ["", ""]
        all_values.append(row4)

        # 各局のデータ
        origin_point = ORIGIN_POINT_3 if player_n == 3 else ORIGIN_POINT_4
        scores = [origin_point] * player_n + [0]

        for hand in round_data.hands:
            # スコア表示行
            score_row = ["", "", ""] + [scores[i] for i in range(player_n)] + [scores[player_n], ""]
            all_values.append(score_row)

            # 和了詳細テキスト作成
            hule_text = ""
            for hule in hand.huleData:
                if len(hule_text) > 0:
                    hule_text += "\n"
                if hule.isNagashi:
                    hule_text += f"{round_data.names[hule.seat]} 流し満貫 8000"
                else:
                    hule_text += f"{round_data.names[hule.seat]} "
                    hule_text += "ツモ" if hule.rongPlayer == -1 else "ロン"
                    hule_text += f" {hule.dadian}\n"
                    hule_text += hule.get_fans_text(self.fan_names, DORA_FANS)

            # 局情報と点数変動行（局情報をB列、C列は空）
            delta_main_row = ["", hand.roundStr, ""] + [hand.deltaMain[i] if hand.deltaMain[i] != 0 else "" for i in range(player_n)] + ["", hule_text]
            all_values.append(delta_main_row)

            # 供託変動行
            delta_sub_other = -sum(hand.deltaSub)
            delta_sub_row = ["", "", ""] + [hand.deltaSub[i] if hand.deltaSub[i] != 0 else "" for i in range(player_n)] + [delta_sub_other if delta_sub_other != 0 else "", ""]
            all_values.append(delta_sub_row)

            # スコア更新
            for i in range(player_n):
                scores[i] += hand.deltaMain[i] + hand.deltaSub[i]
            scores[player_n] += delta_sub_other

        # 最終スコア行
        final_score_row = ["", "", ""] + [scores[i] for i in range(player_n)] + [scores[player_n], ""]
        all_values.append(final_score_row)

        # 最終得点行（JSONのtotal_pointを1000で割った値）
        final_points_row = ["", "", ""]
        for i in range(player_n):
            final_score = round_data.scores[i] / 1000
            final_points_row.append(final_score)
        final_points_row.extend(["", ""])
        all_values.append(final_points_row)

        # 一括更新
        worksheet.clear()
        if all_values:
            worksheet.update('A1', all_values, value_input_option='USER_ENTERED')

        # チーム色の設定
        self._apply_team_colors(worksheet, round_data, player_data_dict, player_n)

        # 少し待機（レート制限対策）
        time.sleep(10)

    def export_player_sheet(self, player_data_dict: Dict[str, PlayerData], player_n: int):
        """プレイヤーデータをシートに出力"""
        sheet_name = f"【{'四麻' if player_n == 4 else '三麻'}】プレイヤーデータ"
        try:
            worksheet = self.spreadsheet.worksheet(sheet_name)
        except gspread.exceptions.WorksheetNotFound:
            worksheet = self.spreadsheet.add_worksheet(title=sheet_name, rows=200, cols=10)

        # データ準備
        all_values = []

        # 空行
        all_values.append([])

        # ヘッダー
        header = ["HN", "スコア", "最大和了", "支払合計", "ドラ合計", "レア役"]
        all_values.append(header)

        # データ
        for name, player_data in player_data_dict.items():
            for player_single_data in player_data.dataList:
                rare_fans_text = ",".join([
                    self.fan_names[str(fan_id)]
                    for fan_id in player_single_data.rareFans
                ])

                row_data = [
                    name,
                    player_single_data.score,
                    player_single_data.maxHule,
                    player_single_data.paySum,
                    player_single_data.doraCount,
                    rare_fans_text,
                ]
                all_values.append(row_data)

        # 一括更新
        worksheet.clear()
        if all_values:
            worksheet.update('B1', all_values, value_input_option='USER_ENTERED')

        # 少し待機
        time.sleep(10)

    def export_total_result_sheet(self, round_data_list: List[RoundData], player_data_dict: Dict[str, PlayerData], player_n: int):
        """総合結果をシートに出力"""
        sheet_name = f"【{'四麻' if player_n == 4 else '三麻'}】総合結果"
        try:
            worksheet = self.spreadsheet.worksheet(sheet_name)
        except gspread.exceptions.WorksheetNotFound:
            worksheet = self.spreadsheet.add_worksheet(title=sheet_name, rows=200, cols=10)

        # データ準備
        all_values = []

        # 空行
        all_values.append([])

        # チーム名とマッピング
        if player_n == 4:
            teams = ["青チーム", "赤チーム", "白チーム", "黒チーム"]
        else:
            teams = ["チームA", "チームB", "チームC"]

        # 空行
        all_values.append([])

        # ヘッダー（B列から開始）
        header_row = [""] + teams[:player_n]
        all_values.append(header_row)

        # データ
        team_totals = [0] * player_n
        for round_data in round_data_list:
            row_data = [""] + [""] * player_n  # A列は空、B列から開始
            for i in range(player_n):
                player_name = round_data.names[i]
                if player_name in player_data_dict:
                    team = player_data_dict[player_name].team
                    score = round_data.scores[i] / 1000

                    # チームに応じてカラムを決定
                    if player_n == 4:
                        team_to_col = {"青チーム": 0, "赤チーム": 1, "白チーム": 2, "黒チーム": 3}
                    else:
                        team_to_col = {"チームA": 0, "チームB": 1, "チームC": 2}

                    col = team_to_col.get(team, 0)
                    if col < player_n:
                        row_data[col + 1] = score  # +1でB列から開始
                        team_totals[col] += score

            all_values.append(row_data)

        # 合計行を追加
        total_row = ["合計"] + team_totals
        all_values.append(total_row)

        # 一括更新
        worksheet.clear()
        if all_values:
            worksheet.update('A1', all_values, value_input_option='USER_ENTERED')

            # ヘッダー行に色を適用
            self._apply_total_header_colors(worksheet, player_n)

            # 合計行の上に罫線を追加
            total_row_index = len(all_values)
            range_start = f"A{total_row_index}"
            range_end = chr(ord('A') + player_n) + str(total_row_index)
            worksheet.format(f"{range_start}:{range_end}", {
                "borders": {
                    "top": {"style": "SOLID", "width": 1}
                }
            })

        # 少し待機
        time.sleep(10)

    def _apply_team_colors(self, worksheet, round_data: RoundData, player_data_dict: Dict[str, PlayerData], player_n: int):
        """チーム色を適用"""
        try:
            # チーム色の定義
            if player_n == 4:
                team_colors = {
                    "青チーム": {"red": 0.788, "green": 0.855, "blue": 0.972},
                    "赤チーム": {"red": 0.957, "green": 0.8, "blue": 0.8},
                    "白チーム": {"red": 1, "green": 1, "blue": 1},
                    "黒チーム": {"red": 0.851, "green": 0.851, "blue": 0.851},
                }
            else:
                team_colors = {
                    "チームA": {"red": 0.788, "green": 0.855, "blue": 0.972},
                    "チームB": {"red": 0.957, "green": 0.8, "blue": 0.8},
                    "チームC": {"red": 1, "green": 1, "blue": 1},
                }

            # プレイヤー名の背景色設定
            for i, name in enumerate(round_data.names):
                if name in player_data_dict:
                    team = player_data_dict[name].team
                    if team in team_colors:
                        # プレイヤー名のセル（D4からの位置）
                        cell = chr(ord('D') + i) + '4'
                        worksheet.format(cell, {
                            "backgroundColor": team_colors[team]
                        })

                        # 方角のセル（D2からの位置）
                        direction_cell = chr(ord('D') + i) + '2'
                        worksheet.format(direction_cell, {
                            "backgroundColor": team_colors[team]
                        })

        except Exception as e:
            print(f"Warning: Could not apply team colors: {e}")

    def _apply_total_header_colors(self, worksheet, player_n: int):
        """総合結果のヘッダーにチーム色を適用"""
        try:
            # チーム色の定義
            if player_n == 4:
                team_colors = [
                    {"red": 0.788, "green": 0.855, "blue": 0.972},  # 青チーム
                    {"red": 0.957, "green": 0.8, "blue": 0.8},      # 赤チーム
                    {"red": 1, "green": 1, "blue": 1},              # 白チーム
                    {"red": 0.851, "green": 0.851, "blue": 0.851},  # 黒チーム
                ]
            else:
                team_colors = [
                    {"red": 0.788, "green": 0.855, "blue": 0.972},  # チームA
                    {"red": 0.957, "green": 0.8, "blue": 0.8},      # チームB
                    {"red": 1, "green": 1, "blue": 1},              # チームC
                ]

            # ヘッダー行（2行目）の各チーム列に色を適用
            for i in range(player_n):
                cell = chr(ord('B') + i) + '2'
                worksheet.format(cell, {
                    "backgroundColor": team_colors[i]
                })

        except Exception as e:
            print(f"Warning: Could not apply total header colors: {e}")