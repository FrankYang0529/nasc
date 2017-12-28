import chardet
import json
from bs4 import BeautifulSoup


MISSION_TYPE = {
    '第一值(備)勤': 'SP1',
    '第二值(備)勤': 'SP2',
    '第三值(備)勤': 'SP3',
    '訓練飛行': 'TF',
    '共勤訓練': 'SM',
    '海洋(岸)空偵巡護': 'CG',
    '試車': 'MGR',
    '維護飛行': 'MTF',
    '救護轉診': 'EMS'
}


def parse_mission(htm):
    missions = list()

    soup = BeautifulSoup(htm, 'html.parser')
    mission_soup_list = soup.find_all(
        lambda tag: tag.name == 'tr' and (tag.get('class') == ['GridRow_Office2007'] or tag.get('class') == ['GridAltRow_Office2007']))
    div_soup_list = soup.find_all(
        lambda tag: tag.name == 'div' and tag.get('style') == 'background-color: #c0c0c0; padding: 5px 5px 5px 5px;')

    for idx, mission_soup in enumerate(mission_soup_list):
        plane_num = mission_soup.find('a').string.strip()
        zh_type = mission_soup.find_all('span', {'class': 'DropDownList'})[0].string.strip()
        en_type = MISSION_TYPE.get(zh_type, '')
        time = mission_soup.find_all('span', {'class': 'DropDownList'})[1].string.strip().replace(':', '')
        note = mission_soup.find_all('td')[-1].string.strip()

        people = str()
        tr_soup_list = div_soup_list[idx].find_all(lambda tag: tag.name == 'tr' and tag.get('style') is None)
        for tr_soup in tr_soup_list:
            title = tr_soup.find('td', {'align': 'center'}).string.strip()
            name = tr_soup.find('td', {'align': 'left'}).string.strip()
            people += f'{title}:{name}'

        missions.append({
            'plane-num': plane_num,
            'zh_type': zh_type,
            'en_type': en_type,
            'people': people.strip(),
            'time': time,
            'note': note
        })

    return missions


def parse_time(htm, past_missions):
    soup = BeautifulSoup(htm, 'html.parser')
    tr_soup_list = soup.find_all(
        lambda tag: tag.name == 'tr' and (tag.get('class') == ['GridRow_Office2007'] or tag.get('class') == ['GridAltRow_Office2007']))

    for tr_soup in tr_soup_list:
        plane_num = tr_soup.find_all('td')[4].string.strip()
        zh_type = tr_soup.find_all('td')[6].string.strip()
        origin_time = tr_soup.find_all('span', {'class': 'DropDownList'})[1].string.strip().replace(':', '')
        start_time = tr_soup.find_all('span', {'class': 'DropDownList'})[2].string.strip().replace(':', '')
        end_time = tr_soup.find_all('span', {'class': 'DropDownList'})[5].string.strip().replace(':', '')

        mission_idx = next(
            idx for (idx, m) in enumerate(past_missions) if m['plane-num'] == plane_num and m['zh_type'] == zh_type and m['time'] == origin_time)
        past_missions[mission_idx]['time'] = f'{start_time} - {end_time}'

    return past_missions


def parse_plane_status(htm):
    plane_status_list = list()

    soup = BeautifulSoup(htm, 'html.parser')
    time_soup_list = soup.find_all(
        lambda tag: tag.name == 'tr' and (tag.get('class') == ['GridRow_Office2007'] or tag.get('class') == ['GridAltRow_Office2007']))
    device_soup_list = soup.find_all('table', {'border': '1', 'cellpadding': '3', 'cellspacing': '5'})

    for idx in range(0, len(time_soup_list)):
        time_soup = time_soup_list[idx]
        device_soup = device_soup_list[idx]

        plane_num = time_soup.find_all('td')[2].string.strip()
        status = time_soup.find_all('td')[3].find('span').string.strip()
        check_date = time_soup.find_all('span', {'class': 'DropDownList'})[0].string.strip()
        position = time_soup.find_all('span', {'class': 'DropDownList'})[1].string.strip()
        yesterday_time = time_soup.find_all('td')[6].string.strip()
        plane_time = time_soup.find_all('td')[7].string.strip()
        engine_time = time_soup.find_all('td')[8].string.strip()
        distance_check_time = time_soup.find_all('td')[9].string.strip()

        person_hang = device_soup.find_all('td')[1].string.strip()
        emergency_buoy = device_soup.find_all('td')[6].string.strip()

        plane_status_list.append({
            'plane-num': plane_num,
            'status': status,
            'check_date': check_date,
            'position': position,
            'yesterday_time': yesterday_time,
            'plane_time': plane_time,
            'engine_time': engine_time,
            'distance_check_time': distance_check_time,
            'person_hang': person_hang,
            'emergency_buoy': emergency_buoy
        })

    return plane_status_list


if __name__ == '__main__':
    with open('past_mission.txt', 'rb') as my_file:
        past_mission_bytes = my_file.read()
        past_mission_htm = past_mission_bytes.decode(chardet.detect(past_mission_bytes)['encoding'])

    with open('today_mission.txt', 'rb') as my_file:
        today_mission_bytes = my_file.read()
        today_mission_htm = today_mission_bytes.decode(chardet.detect(today_mission_bytes)['encoding'])

    with open('plane_status.txt', 'rb') as my_file:
        plane_status_bytes = my_file.read()
        plane_status_htm = plane_status_bytes.decode(chardet.detect(plane_status_bytes)['encoding'])

    with open('time.txt', 'rb') as my_file:
        time_bytes = my_file.read()
        time_htm = time_bytes.decode(chardet.detect(time_bytes)['encoding'])

    past_missions = parse_mission(past_mission_htm)
    past_missions = [mission for mission in past_missions if (mission['en_type'] not in ('SP1', 'SP2', 'SP3'))]

    today_missions = parse_mission(today_mission_htm)
    past_missions = parse_time(time_htm, past_missions)
    plane_status_list = parse_plane_status(plane_status_htm)

    output_json = {
        'past_missions': past_missions,
        'today_missions': today_missions,
        'plane_status_list': plane_status_list
    }
    with open('pptx_input.json', 'w') as out_file:
        json.dump(output_json, out_file, ensure_ascii=False, indent=4)

    output_str = '一、前日任務檢討：\n'
    for idx, mission in enumerate(past_missions):
        output_str += f'{idx+1}. {mission["en_type"]}: {mission["plane-num"]} {mission["people"]} {mission["time"]} {mission["note"]}\n'

    output_str += '二、本日任務重點：(人員編組及飛機情況詳如任務派遣單)\n'
    for idx, mission in enumerate(today_missions):
        output_str += f'{idx+1}. {mission["en_type"]}: {mission["plane-num"]} {mission["people"]} {mission["time"]} {mission["note"]}\n'

    with open('任務提示.txt', 'w') as my_file:
        my_file.write(output_str)
