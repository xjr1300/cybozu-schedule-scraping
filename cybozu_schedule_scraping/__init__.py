from dataclasses import dataclass
from datetime import date, time
from getpass import getpass
from typing import Any, Dict, List, NewType, Optional, Tuple

import requests
from bs4 import BeautifulSoup, Tag

# サイボウズルートURI
CB_ROOT_URI = "http://192.168.220.14/scripts/cbag/ag.exe?"

# 組織コード
DivisionCode = NewType("DivisionCode", str)


class CBScrapingException(Exception):
    """サイボウズスクレイピング例外"""


@dataclass
class LoginInfo:
    """ログイン情報"""

    # サイボウズにログインするユーザーを選択するときに指定する組織
    division_name: str
    # サイボウズにログインするユーザーの名前
    name: str
    # 上記ユーザーのパスワード
    password: str


class YearMonth:
    """年月"""

    # 年
    year: int
    # 月
    month: int

    def __init__(self, year: int, month: int) -> None:
        """イニシャライザ

        引数:
            year: 年。
            month: 月。
        """
        if year < 1900 or 2100 < year:
            raise CBScrapingException("年は1900以上2100以下を指定してください。")
        if month < 1 or 12 < month:
            raise CBScrapingException("月は1以上12以下を指定してください。")
        self.year = year
        self.month = month

    def __str__(self) -> str:
        return f"{self.year:04}/{self.month:02}"

    def __repr__(self) -> str:
        return f"YearMonth(year={self.year}, month={self.month})"

    @property
    def text_jp(self) -> str:
        """年月を日本語で表現した文字列を返却する。

        戻り値:
            年月を日本語で表現した文字列。
        """
        return f"{self.year:04}年{self.month:02}月"


def _str_to_time(text: str) -> Optional[time]:
    """文字列を時刻に変換する。

    引数:
        text: 時刻に変換する文字列。
    戻り値:
        文字列から変換した時刻。
        時刻に変換できない場合は`None`。
    """
    text = text.strip()
    if not text:
        return None
    splitted = text.split(":", 1)
    if len(splitted) != 2:
        return None
    return time(hour=int(splitted[0]), minute=int(splitted[1]))


def _time_to_str(time: Optional[time]) -> str:
    """時刻を文字列に変換する。

    引数:
        time: 文字列に変換する時刻。
    戻り値:
        時刻を変換した文字列。
        時刻が`None`の場合は`""`。
    """
    return time.strftime("%H:%M") if time else ""


def prompt_user_for_login_info() -> LoginInfo:
    """ユーザーにサイボウズにログインするときに必要な情報の入力を求める。

    戻り値:
        ログイン情報。
    """

    division_name = input("サイボウズでユーザーを選択するときの組織名: ")
    name = input("サイボウズにログインするユーザーの名前: ")
    password = getpass("パスワード: ")
    return LoginInfo(division_name, name, password)


def prompt_user_for_year_month() -> YearMonth:
    """ユーザーにスケジュールを取得する年と月の入力を求める。

    戻り値:
        年月。
    """
    today = date.today()
    separator = "/"
    default_value = f"{today.year:04}{separator}{today.month:02}"
    text = input(f"スケジュールを取得する年月 [default: {default_value}]: ")
    if not text:
        text = default_value
    splitted = text.split(separator)
    if len(splitted) != 2:
        raise CBScrapingException(f"年と月を{separator}で区切って入力してください。")
    try:
        year = int(splitted[0])
    except ValueError:
        raise CBScrapingException("年を数値として認識できません。")
    try:
        month = int(splitted[1])
    except ValueError:
        raise CBScrapingException("月を数値として認識できません。")
    return YearMonth(year, month)


def call_http_method(
    session: requests.Session,
    uri_part: str,
    method: str = "get",
    data: Optional[Dict[str, Any]] = None,
) -> requests.Response:
    """サイボウズにGETリクエストを送信する。

    引数:
        session: HTTPセッション。
        uri_part: サイボウズのルートURL以下のURI。
        method: メソッド。デフォルトは`"get"`。
        data: 送信するデータ。デフォルトは`None`。
    戻り値:
        レスポンス。
    例外:
        CBScrapingException
    """
    handler = getattr(session, method)
    uri = f"{CB_ROOT_URI}{uri_part}"
    try:
        response = handler(url=uri, data=data)
        response.raise_for_status()
    except requests.HTTPError:
        raise CBScrapingException(f"`{uri}`への{method.upper()}リクエストに失敗しました。")
    return response


def retrieve_division_code(
    session: requests.Session, division_name: str
) -> DivisionCode:
    """サイボウズの組織選択ページから組織コードを取得して返却する。

    サイボウズの組織選択ページから仮引数で渡された組織名と一致する組織をselect > option要素から
    検索して、その組織コード(optionのvalue属性の値)を返却する。

    引数:
        session: HTTPセッション。
        division_name: 組織名。
    戻り値:
        組織コード。
    例外:
        CBScrapingException
    """
    # 組織選択ページを取得
    response = call_http_method(session, "page=LoginGroup")
    # 組織選択ページのHTMLコンテンツをDOMに展開
    soup = BeautifulSoup(response.content, "html.parser")
    # 組織を選択するselect要素のoption要素を取得
    options = soup.select("select.select-gid[name='Group'] option")
    options = [option for option in options if option.text.strip() == division_name]
    if not options:
        raise CBScrapingException(f"入力された組織({division_name})が組織選択ページで見つかりませんでした。")
    return DivisionCode(options[0]["value"])


def login(
    session: requests.Session, division_code: DivisionCode, login_info: LoginInfo
) -> str:
    """ログインページに遷移して、ログインする。

    引数:
        session: HTTPセッション。
        division_code: 組織コード。
        login_info: ログイン情報。
    戻り値:
        ユーザーID。
    例外:
        CBScrapingException
    """
    # ログインページに遷移
    uri_part = f"gid={division_code}&&Group={division_code}"
    response = call_http_method(session, uri_part)
    # ログインページのHTMLコンテンツをDOMに展開
    soup = BeautifulSoup(response.content, "html.parser")
    # ログインページのコンテンツからユーザーのIDを取得
    options = soup.select("td.loginmain select.vr_loginForm[name='_ID'] option")
    options = [option for option in options if option.text.strip() == login_info.name]
    if not options:
        raise CBScrapingException(f"入力された名前({login_info.name})がログインページで見つかりませんでした。")
    user_id = options[0]["value"]
    # ユーザークレデンシャルをログインページにPOSTしてログイン
    data = {
        "csrf_ticket": "",
        "_System": "login",
        "_Login": "1",
        "LoginMethod": "1",
        "_ID": user_id,
        "Password": login_info.password,
    }
    response = call_http_method(session, uri_part, method="post", data=data)
    return user_id


@dataclass
class Schedule:
    # 日
    day: int
    # 開始時刻
    begin: Optional[time]
    # 終了時刻
    end: Optional[time]
    # スケジュールのタイトル
    title: Optional[str]

    def __str__(self) -> str:
        if not self.begin and not self.end:
            return f"{self.day}日 {self.title}"
        else:
            begin_str = _time_to_str(self.begin)
            end_str = _time_to_str(self.end)
            time_range = f"{begin_str}-{end_str}" if end_str else begin_str
            return f"{self.day}日 {time_range} {self.title}"


def _is_event_cell_at_month(event_cell: Tag, month: int) -> bool:
    """td.eventcellがスケジュールを抽出したい月であるか確認する。

    引数:
        event_cell: td.eventcell要素。
        month: スケジュールを抽出したい月。

    戻り値:
        td.eventcellがスケジュールを抽出したい月の場合はTrue。それ以外はFalse。
    """
    date_span = event_cell.select_one("span.date")
    m, _ = date_span.text.split("/", 1)
    return int(m) == month


def _retrieve_event_links_in_event_cell(event_cell: Tag) -> Tuple[Tag, List[Tag]]:
    """td.eventcell要素に含まれる、すべてのdiv.eventLink要素を抽出する。

    引数:
        event_cell: td.eventcell要素。
    戻り値:
        仮引数で渡されたtd.eventcell要素と、その要素内に含まれるすべてのdiv.eventLink要素を格納したリスト。
    """
    return (
        event_cell,
        event_cell.select("div.eventLink"),
    )


def retrieve_monthly_schedules(
    session: requests.Session, user_id: str, ym: YearMonth
) -> List[Schedule]:
    """仮引数で指定された年月のスケジュールをすべて取得して返却する。

    引数:
        session: HTTPセッション。
        ym: スケジュールを取得する年月。
    Returns:
        スケジュールを格納したリスト。
    """
    # サイボウズの個人月表示ページに遷移
    date_str = f"da.{ym.year:04}.{ym.month:02}.01"
    uri_part = f"page=ScheduleUserMonth&UID={user_id}&Date={date_str}"
    response = call_http_method(session, uri_part)

    # 組織選択ページのHTMLコンテンツをDOMに展開
    soup = BeautifulSoup(response.content, "html.parser")

    # td.eventcell要素を取得
    event_cells = soup.select("td.eventcell")
    # 指定された月のtd.eventcell要素のみを抽出
    event_cells = [
        event_cell
        for event_cell in event_cells
        if _is_event_cell_at_month(event_cell, ym.month)
    ]
    # div.eventLink要素を取得
    event_cell_links = [
        _retrieve_event_links_in_event_cell(event_cell) for event_cell in event_cells
    ]
    # スケジュールを抽出
    schedules: List[Schedule] = []
    for event_cell, event_links in event_cell_links:
        # td.eventcell span.date要素から日を取得
        day = int(event_cell.select_one("span.date").text.split("/", 1)[1])
        # div.eventLink要素から時間帯及びタイトルを取得
        for event_link in event_links:
            begin: Optional[time] = None
            end: Optional[time] = None
            # div.eventLink div.eventInner span.eventDateTime要素から時間帯を取得
            time_range_elem = event_link.select_one("div.eventInner span.eventDateTime")
            if time_range_elem:
                time_range_text = time_range_elem.text.removesuffix("&nbsp;")
                splitted = time_range_text.split("-", 1)
                begin = _str_to_time(splitted[0])
                if len(splitted) == 2:
                    end = _str_to_time(splitted[1])
            # div.eventLink div.eventInner span.eventDetail a.event要素からタイトルを取得
            title_elem = event_link.select_one("div.eventInner a.event")
            if hasattr(title_elem, "title"):
                schedules.append(Schedule(day, begin, end, title_elem["title"]))
    return schedules
