"""牌譜JSONパーサー"""
import json
from pathlib import Path
from typing import Dict, List
from icecream import ic

from data_structures import (
    HuleSingleData, HandData, RoundData,
    PlayerHalfRoundData, PlayerData
)
from config import DORA_FANS, RARE_FANS, ORIGIN_POINT_3, ORIGIN_POINT_4

class PaifuParser:
    def __init__(self, player_n: int, members_map: Dict[str, Dict]):
        self.player_n = player_n
        self.members_map = members_map
        self.origin_point = ORIGIN_POINT_3 if player_n == 3 else ORIGIN_POINT_4

    def parse_round(self, filename: Path) -> RoundData:
        """半荘のデータをパース"""
        with open(filename, 'r', encoding='utf-8') as f:
            data = json.load(f)

        round_data = RoundData(self.player_n)
        round_data.uuid = data["head"]["uuid"]

        # head
        for player in data["head"]["result"]["players"]:
            round_data.scores[player["seat"]] = player["total_point"]

        for account in data["head"]["accounts"]:
            seat = account.get("seat", 0)
            game_name = account["nickname"]
            # game_nameからメンバー情報を取得
            member_info = self.members_map.get(game_name, {})
            display_name = member_info.get("name", game_name)
            round_data.names[seat] = display_name

        ic(round_data.names)

        # action
        current_scores = [0] * self.player_n
        current_parent = 0
        current_round_hand_data = None

        for ai, action in enumerate(data["data"]["data"]["actions"]):
            if action["type"] == 1:
                # 局開始
                if action["result"]["name"] == ".lq.RecordNewRound":
                    current_parent = action["result"]["data"]["ju"]
                    current_round_hand_data = HandData(self.player_n, action["result"]["data"])
                    current_scores = action["result"]["data"]["scores"]
                    ic(current_scores)

                # 和了
                elif action["result"]["name"] == ".lq.RecordHule":
                    for hule in action["result"]["data"]["hules"]:
                        if hule["zimo"]:  # ツモ
                            current_round_hand_data.huleData.append(
                                HuleSingleData(
                                    seat=hule["seat"],
                                    isNagashi=False,
                                    rongPlayer=-1,
                                    dadian=hule["dadian"],
                                    han=hule["count"],
                                    fu=hule["fu"],
                                    fans=hule["fans"],
                                )
                            )
                            # それぞれの収支を計算
                            for seat in range(self.player_n):
                                if seat == hule["seat"]:
                                    current_round_hand_data.deltaMain[seat] += hule["dadian"]
                                elif seat == current_parent:
                                    current_round_hand_data.deltaMain[seat] -= hule["point_zimo_qin"]
                                else:
                                    current_round_hand_data.deltaMain[seat] -= hule["point_zimo_xian"]
                        else:  # ロン
                            # 放銃者の情報を取得
                            for back in range(2, 7):
                                if ai - back < 0:
                                    continue
                                prev_action = data["data"]["data"]["actions"][ai - back]
                                if (prev_action["type"] == 1 and
                                    prev_action["result"]["name"] == ".lq.RecordDiscardTile"):
                                    rong_player = prev_action["result"]["data"]["seat"]
                                    current_round_hand_data.huleData.append(
                                        HuleSingleData(
                                            seat=hule["seat"],
                                            isNagashi=False,
                                            rongPlayer=rong_player,
                                            dadian=hule["dadian"],
                                            han=hule["count"],
                                            fu=hule["fu"],
                                            fans=hule["fans"],
                                        )
                                    )
                                    current_round_hand_data.deltaMain[hule["seat"]] += hule["dadian"]
                                    current_round_hand_data.deltaMain[rong_player] -= hule["dadian"]
                                    break

                    round_data.hands.append(current_round_hand_data)

                    for seat in range(self.player_n):
                        current_round_hand_data.deltaSub[seat] = (
                            (action["result"]["data"]["old_scores"][seat] - current_scores[seat])
                            + action["result"]["data"]["delta_scores"][seat]
                            - current_round_hand_data.deltaMain[seat]
                        )

                # 流局
                elif action["result"]["name"] == ".lq.RecordNoTile":
                    if "scores" in action["result"]["data"]:
                        scores_data = action["result"]["data"]["scores"]
                        if scores_data:
                            score = scores_data[0]
                            for seat in range(self.player_n):
                                old_score = score["old_scores"][seat] if "old_scores" in score else current_scores[seat]
                                delta_score = score["delta_scores"][seat] if "delta_scores" in score else 0
                                current_round_hand_data.deltaSub[seat] = (
                                    old_score - current_scores[seat] + delta_score
                                )

                            # 流し満貫
                            if "seat" in score:
                                hule_single_data = HuleSingleData(
                                    seat=score["seat"],
                                    isNagashi=True,
                                    rongPlayer=-1,
                                    dadian=0,
                                    han=0,
                                    fu=0,
                                    fans=[{"id": -1, "val": 0}],
                                )
                                current_round_hand_data.huleData.append(hule_single_data)

                    round_data.hands.append(current_round_hand_data)

            # 思考時間とスタンプは集計対象外のため削除

        return round_data