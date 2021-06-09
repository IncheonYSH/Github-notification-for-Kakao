from flask import Flask
from flask import request
import win32con
import win32gui
import win32api
import time
import json


class EventHandler:
    """
    github 이벤트를 처리하는 클래스

    Attributes:
        event_list: 처리할 github 이벤트 리스트와 메소드를 mapping 한 딕셔너리, dict 객체
    """
    def __init__(self):
        self.event_list = {
            'push': self._push,
            'issues': self._issues,
            'issue_comment': self._issue_comment,
            'commit_comment': self._commit_comment,
            'create': self._create,
            'delete': self._delete,
            'pull_request': self._pull_request,
            'pull_request_review': self._pull_request_review,
            'pull_request_review_comment': self._pull_request_review_comment,
            'fork': self._fork
        }

    def _push(self, context):
        commit_info = []
        hlink = (context['compare'])
        if len(context['commits']) == 0:
            return '0'
        elif len(context['commits']) == 1:
            commit_info.append(str(len(context['commits'])) + ' new commit')
        else:
            commit_info.append(str(len(context['commits'])) + ' new commits')
        for commit in context['commits']:
            commit_info.append(
                commit["id"][:6] + ' - "' + commit['message'] + '" by ' + commit['committer']['username']
            )
        result = '\n'.join(commit_info)
        return hlink, result

    def _issues(self, context):
        commit_info = []
        hlink = (context['issue']['html_url'])
        commit_info.append(
            'Issue ' + context['action'] + ' # ' + str(context['issue']['number']) + ': ' + context['issue']['title']
        )
        result = '\n'.join(commit_info)
        return hlink, result

    def _issue_comment(self, context):
        commit_info = []
        hlink = (context['issue']['html_url'])
        commit_info.append(
            'Comment ' + context['action'] + ' on issue # ' + str(context['issue']['number'])
        )
        result = '\n'.join(commit_info)
        return hlink, result

    def _commit_comment(self, context):
        commit_info = []
        hlink = (context['comment']['html_url'])
        commit_info.append(
            'Comment ' + context['action'] + ' on commit ' + str(context['comment']['commit_id'][:6])
        )
        result = '\n'.join(commit_info)
        return hlink, result

    def _create(self, context):
        commit_info = []
        hlink = (context['repository']['html_url'])
        commit_info.append(
            'Created ' + context['ref_type'] + ': ' + context['ref']
        )
        result = '\n'.join(commit_info)
        return hlink, result

    def _delete(self, context):
        commit_info = []
        hlink = (context['repository']['html_url'])
        commit_info.append(
            'Deleted ' + context['ref_type'] + ': ' + context['ref']
        )
        result = '\n'.join(commit_info)
        return hlink, result

    def _pull_request(self, context):
        commit_info = []
        hlink = (context['pull_request']['html_url'])
        if context['action'] == 'closed' and context['pull_request']['merged'] == False:
            action = 'closed and not merged'
        elif context['action'] == 'closed' and context['pull_request']['merged'] == True:
            action = 'merged'
        else:
            action = context['action']
        commit_info.append(
            'Pull request ' + action + ' # ' + str(context['number']) + ': ' + context['pull_request']['title']
        )
        if len(context['pull_request']['body']) == 0:
            pass
        else:
            commit_info.append(context['pull_request']['body'])
        result = '\n'.join(commit_info)
        return hlink, result

    def _pull_request_review(self, context):
        commit_info = []
        hlink = (context['review']['html_url'])
        commit_info.append(
            'Pull request review ' + context['action'] + ' on pull request # ' + str(context['pull_request']['number'])
        )
        result = '\n'.join(commit_info)
        return hlink, result

    def _pull_request_review_comment(self, context):
        commit_info = []
        hlink = (context['comment']['html_url'])
        commit_info.append(
            'Pull request review comment ' + context['action'] + ' on pull request # ' + str(context['pull_request']['number'])
        )
        result = '\n'.join(commit_info)
        return hlink, result

    def _fork(self, context):
        commit_info = []
        hlink = (context['forkee']['html_url'])
        commit_info.append(
            'Fork created: ' + context['forkee']['full_name']
        )
        result = '\n'.join(commit_info)
        return hlink, result


class MetaSingleton(type):
    """
    싱글톤 구현을 위한 메타클래스
    """
    _instances = {}
    def __call__(cls, *args, **kwargs):
        if cls not in cls._instances:
            cls._instances[cls] = super(MetaSingleton, cls).__call__(*args, **kwargs)
        return cls._instances[cls]

class Kakao(metaclass=MetaSingleton):
    """
    윈도우 api를 통해 카카오톡으로 메세지를 발송하기 위한 클래스

    Attributes:
        handle_textbox: 카카오톡
    """
    def __init__(self, chatroom_name):
        handle_kakao = win32gui.FindWindow(None, "카카오톡")
        handle_main = win32gui.FindWindowEx(handle_kakao, None, "EVA_ChildWindow", None)
        handle_friend = win32gui.FindWindowEx(handle_main, None, "EVA_Window", None)
        handle_chatlist = win32gui.FindWindowEx(handle_main, handle_friend, "EVA_Window", None)
        handle_roomsearch = win32gui.FindWindowEx(handle_chatlist, None, "Edit", None)
        win32api.SendMessage(handle_roomsearch, win32con.WM_SETTEXT, 0, chatroom_name)
        time.sleep(1)
        self._press_enter(handle_roomsearch)
        time.sleep(1)
        handle_chatroom = win32gui.FindWindow(None, chatroom_name)

        self.handle_textbox = win32gui.FindWindowEx(handle_chatroom, None, "RICHEDIT50W", None)

    def message_send(self, message):
        win32api.SendMessage(self.handle_textbox, win32con.WM_SETTEXT, 0, message)
        self._press_enter(self.handle_textbox)

    def _press_enter(self, handle):
        win32api.PostMessage(handle, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)
        time.sleep(0.01)
        win32api.PostMessage(handle, win32con.WM_KEYUP, win32con.VK_RETURN, 0)



# Flask app
app = Flask(__name__)

@app.route("/", methods=['POST'])
def webhook_receiver():
    context = request.get_json()  # dict object
    event = request.headers.get('X-GitHub-Event')
    api_info = EventHandler()

    # prompt input 으로 받은 채팅방 이름
    global CHATROOM_NAME
    kakao = Kakao(CHATROOM_NAME)

    # github 에서 발생한 이벤트가 지원하는 이벤트 리스트에 있으면 카카오톡 메세지 발송
    if event in api_info.event_list.keys():
        return_message = 'Kakao send'
        result = api_info.event_list[event](context)
        if len(result) != 2:
            return_message = 'Kakao do not send'
            pass
        else:
            for message in result:
                kakao.message_send(message)
        return return_message
    else:
        result = 'success'
    return result

if __name__ == "__main__":
    # prompt input 으로 타겟으로 할 채팅방의 이름 등록
    CHATROOM_NAME = input("채팅방 이름: ")
    port = input("포트: ")
    app.run(host="0.0.0.0", port=port)