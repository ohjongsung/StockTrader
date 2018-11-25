import requests
import json


# 슬랙에 메시지를 보내는 클래스
class Slack(object):
    def __init__(self):
        # 파일에서 웹훅 url 로드
        with open('../.ignores/webhook.txt') as f:
            items = list(f.readlines())
            self.webhook_url = items[1].strip()

    def push(self, message):
        slack_message = {'text': message}
        requests.post(
            self.webhook_url, json=slack_message
        )
