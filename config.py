THEATER_DETAIL_URL = "https://www.vscinemas.com.tw/theater/detail.aspx?id=1"
RECORDS_FILE = "records.xlsx"

# 新增要監控的電影：
#   name    — 網頁上顯示的英文名稱（用於搜尋）
#   title   — 記錄到 Excel 的顯示名稱
#   versions — 要監控的版本清單，留空 [] 代表監控所有版本
MOVIES = [
    {
        "name": "HOPPERS",
        "title": "狸想世界",
        "versions": ["數位(英)"],
    },
    {
        "name": "PROJECT HAIL MARY",
        "title": "極限返航",
        "versions": ["數位(英)"],
    },
]
