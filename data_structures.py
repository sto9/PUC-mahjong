"""データ構造の定義"""
from typing import List, Dict, Optional
from dataclasses import dataclass, field

@dataclass
class HuleSingleData:
    """和了単体のデータ"""
    seat: int
    isNagashi: bool
    rongPlayer: int  # ツモなら-1、ロンならseat
    dadian: int
    han: int
    fu: int
    fans: List[Dict[str, int]]

    def get_fans_text(self, fan_names: Dict[str, str], dora_fans: List[int]) -> str:
        """役のテキストを取得"""
        fan_list = [(fan["id"], fan["val"]) for fan in self.fans]
        fan_list.sort()
        return ",".join([
            fan_names[str(fan[0])] + (str(fan[1]) if fan[0] in dora_fans else "")
            for fan in fan_list
        ])

@dataclass
class HandData:
    """局のデータ"""
    roundStr: str = ""
    deltaMain: List[int] = field(default_factory=list)
    deltaSub: List[int] = field(default_factory=list)
    huleData: List[HuleSingleData] = field(default_factory=list)

    def __init__(self, player_n: int, result_data_json: Optional[Dict] = None):
        self.deltaMain = [0] * player_n
        self.deltaSub = [0] * player_n
        self.huleData = []

        if result_data_json:
            CHANG = ["東", "南", "西"]
            JU = ["一", "二", "三", "四", "五", "六", "七", "八", "九", "十"]
            self.roundStr = (
                CHANG[result_data_json["chang"]]
                + JU[result_data_json["ju"]]
                + "局\n"
                + str(result_data_json["ben"])
                + "本場"
            )

@dataclass
class RoundData:
    """半荘のデータ"""
    uuid: str = ""
    names: List[str] = field(default_factory=list)
    scores: List[int] = field(default_factory=list)
    hands: List[HandData] = field(default_factory=list)

    def __init__(self, player_n: int):
        self.names = [""] * player_n
        self.scores = [0] * player_n
        self.hands = []

@dataclass
class PlayerHalfRoundData:
    """プレイヤーの半荘データ"""
    score: float = 0
    maxHule: int = 0
    paySum: int = 0
    doraCount: int = 0
    rareFans: List[int] = field(default_factory=list)

    def reflect_fans(self, fans: List[Dict], dora_fans: List[int], rare_fans: List[int]):
        """役を反映"""
        for fan in fans:
            if fan["id"] in dora_fans:
                self.doraCount += fan["val"]
            if fan["id"] in rare_fans:
                self.rareFans.append(fan["id"])

@dataclass
class PlayerData:
    """プレイヤーデータ"""
    team: str = ""
    dataList: List[PlayerHalfRoundData] = field(default_factory=list)