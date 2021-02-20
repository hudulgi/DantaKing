import telegram
from dt_config import *


class Telegram:
    def __init__(self, _token):
        self.bot = telegram.Bot(_token)

    def send_msg(self, _id, _msg):
        try:
            self.bot.sendMessage(_id, _msg)
            return True
        except telegram.exception.TelegramError:
            print("텔레그램 전송 실패: 아이디 오류")
            return False


if __name__ == "__main__":
    telg = Telegram(token)
    telg.send_msg(telegram_id, "테스트 메시지 입니다.")
