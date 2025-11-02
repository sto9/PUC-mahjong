#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import argparse
import glob
from pathlib import Path
from typing import Dict, List
from icecream import ic
import time

from config import (
    load_members, load_fans, PAIFU_DIR,
    DORA_FANS, RARE_FANS
)
from data_structures import (
    PlayerData, PlayerHalfRoundData, RoundData
)
from parser import PaifuParser
from sheets_exporter import SheetsExporter

def create_members_map(members):
    members_map = {}
    for member in members:
        game_name = member.get('game_name', '')
        if game_name:
            members_map[game_name] = member
    return members_map

def calc_player_data_by_round(round_data, player_data_dict, player_n, team_field):
    records = [PlayerHalfRoundData() for _ in range(player_n)]

    for i in range(player_n):
        records[i].score = round_data.scores[i] / 1000

    for hand in round_data.hands:
        for i in range(player_n):
            records[i].maxHule = max(records[i].maxHule, hand.deltaMain[i])
            records[i].paySum += max(0, -hand.deltaMain[i])
        for hule in hand.huleData:
            records[hule.seat].reflect_fans(hule.fans, DORA_FANS, RARE_FANS)

    for i in range(player_n):
        name = round_data.names[i]
        if name not in player_data_dict:
            player_data_dict[name] = PlayerData()
        player_data_dict[name].dataList.append(records[i])

def process_files(player_n, members_map):
    parser = PaifuParser(player_n, members_map)
    round_data_list = []
    player_data_dict = {}

    paifu_dir = PAIFU_DIR / str(player_n)
    json_files = sorted(paifu_dir.glob("*.json"))

    if not json_files:
        print(f"No JSON files found in {paifu_dir}")
        return [], {}

    for json_file in json_files:
        ic(f"Processing: {json_file}")
        round_data = parser.parse_round(json_file)
        round_data_list.append(round_data)

        team_field = f"team{player_n}"
        calc_player_data_by_round(round_data, player_data_dict, player_n, team_field)

    members = load_members()
    for name, player_data in player_data_dict.items():
        for member in members:
            if member.get('name') == name:
                team_field = f"team{player_n}"
                player_data.team = member.get(team_field, '')
                break

    return round_data_list, player_data_dict

def process_4player_games(exporter, members_map):
    """四麻の処理"""
    print("\nCleaning existing 4-player sheets...")
    exporter.clean_mahjong_sheets(4)

    print("\nProcessing 4-player games...")
    round_data_list_4, player_data_dict_4 = process_files(4, members_map)
    if round_data_list_4:
        print(f"4-player: {len(round_data_list_4)} games processed")

        for i, round_data in enumerate(round_data_list_4, 1):
            sheet_name = f"【四麻】第{i}試合"
            print(f"  Exporting {sheet_name}...")
            exporter.export_round_sheet(round_data, sheet_name, 4, player_data_dict_4)
            time.sleep(20)  # API制限対策

        time.sleep(20)
        if round_data_list_4:
            print("  Exporting total results (4-player)...")
            exporter.export_total_result_sheet(round_data_list_4, player_data_dict_4, 4)
            time.sleep(20)  # API制限対策

            print("  Exporting player data (4-player)...")
            exporter.export_player_sheet(player_data_dict_4, 4)
            time.sleep(20)  # API制限対策

    print("\n4-player processing complete!")

def process_3player_games(exporter, members_map):
    """三麻の処理"""
    print("\nCleaning existing 3-player sheets...")
    exporter.clean_mahjong_sheets(3)

    print("\nProcessing 3-player games...")
    round_data_list_3, player_data_dict_3 = process_files(3, members_map)
    if round_data_list_3:
        print(f"3-player: {len(round_data_list_3)} games processed")

        for i, round_data in enumerate(round_data_list_3, 1):
            sheet_name = f"【三麻】第{i}試合"
            print(f"  Exporting {sheet_name}...")
            exporter.export_round_sheet(round_data, sheet_name, 3, player_data_dict_3)
            time.sleep(20)  # API制限対策

        if round_data_list_3:
            print("  Exporting total results (3-player)...")
            exporter.export_total_result_sheet(round_data_list_3, player_data_dict_3, 3)
            time.sleep(20)  # API制限対策

            print("  Exporting player data (3-player)...")
            exporter.export_player_sheet(player_data_dict_3, 3)
            time.sleep(20)  # API制限対策

    print("\n3-player processing complete!")

def process_summary_only(exporter, members_map):
    """総合結果のみを処理"""
    print("\nCleaning existing summary sheets...")
    # 総合結果シートのみを削除
    try:
        worksheets = exporter.spreadsheet.worksheets()
        patterns_to_delete = ["総合結果"]

        for ws in worksheets:
            if any(pattern in ws.title for pattern in patterns_to_delete):
                try:
                    exporter.spreadsheet.del_worksheet(ws)
                    print(f"  Deleted sheet: {ws.title}")
                except Exception as e:
                    print(f"  Could not delete sheet {ws.title}: {e}")
                time.sleep(0.5)
    except Exception as e:
        print(f"Warning: Could not clean summary sheets: {e}")

    print("\nProcessing summary data only...")

    # 四麻の処理
    round_data_list_4, player_data_dict_4 = process_files(4, members_map)
    if round_data_list_4:
        print(f"4-player: {len(round_data_list_4)} games processed")
        print("  Exporting total results (4-player)...")
        exporter.export_total_result_sheet(round_data_list_4, player_data_dict_4, 4)
        time.sleep(5)

    # 三麻の処理
    round_data_list_3, player_data_dict_3 = process_files(3, members_map)
    if round_data_list_3:
        print(f"3-player: {len(round_data_list_3)} games processed")
        print("  Exporting total results (3-player)...")
        exporter.export_total_result_sheet(round_data_list_3, player_data_dict_3, 3)
        time.sleep(5)

    print("\nSummary processing complete!")

def main():
    parser = argparse.ArgumentParser(description='麻雀大会結果集計プログラム')
    parser.add_argument('mode', choices=['4', '3', 'all', 'summary'],
                       help='処理モード: 4=四麻のみ, 3=三麻のみ, all=両方, summary=総合結果のみ')

    args = parser.parse_args()

    print("Starting mahjong tournament result aggregation...")
    mode_descriptions = {
        '4': '四麻のみ',
        '3': '三麻のみ',
        'all': '四麻と三麻',
        'summary': '総合結果のみ'
    }
    print(f"Mode: {mode_descriptions.get(args.mode, args.mode)}")

    members = load_members()
    members_map = create_members_map(members)

    exporter = SheetsExporter()

    if args.mode == '4':
        process_4player_games(exporter, members_map)
    elif args.mode == '3':
        process_3player_games(exporter, members_map)
    elif args.mode == 'all':
        # 既存のシートをクリーンアップ
        print("\nCleaning existing sheets...")
        exporter.clean_all_sheets()

        process_4player_games(exporter, members_map)
        process_3player_games(exporter, members_map)
    elif args.mode == 'summary':
        process_summary_only(exporter, members_map)

    print("\nAll processing complete!")

if __name__ == "__main__":
    main()
