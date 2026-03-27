from datetime import datetime

from config import MOVIES
from click_hoppers_seat_icon import (
    init_driver, click_today_date, get_all_sessions,
    click_seat_icon, parse_seat_info, write_record
)


# 主程式邏輯：依 config 抓取所有電影今天的場次售座資訊
def main():
    print(f'[{datetime.now().strftime("%Y-%m-%d %H:%M")}] 開始爬取...')
    driver = init_driver(headless=False)
    try:
        for movie in MOVIES:
            movie_name = movie["name"]
            movie_title = movie["title"]
            versions = movie.get("versions") or None

            print(f'\n【{movie_title}】')
            click_today_date(driver)
            sessions = get_all_sessions(driver, movie_name, versions)
            print(f'找到 {len(sessions)} 個場次')

            main_tab = driver.window_handles[0]

            for time_str, lang, session_id in sessions:
                click_seat_icon(driver, session_id)
                sold, total = parse_seat_info(driver)
                rate = round(sold / total * 100, 1) if total > 0 else 0
                print(f'  [{lang}] {time_str} — 已售 {sold}/{total} ({rate}%)')

                # 關閉座位分頁，切回戲院頁面
                driver.close()
                driver.switch_to.window(main_tab)

                write_record(time_str, lang, session_id, sold, total, movie_title)

        print('\n寫入完成')
    finally:
        driver.quit()


if __name__ == '__main__':
    main()
