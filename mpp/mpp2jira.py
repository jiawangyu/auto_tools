# -*- coding: utf-8 -*-
# **************************************************************
# *  Filename:    mpp2jira.py
# *  Copyright:   SMEE Co., Ltd.
# *  @author:     jun.zhou
# *  将mpp里的数据同步到jira
# *  @date     2018/7/17  @Reviser  Initial Version
# **************************************************************
import os, sys, datetime
import win32com.client
import traceback
import pythoncom

import logging
import argparse

from jira import JIRA
from jira.client import GreenHopper

reload(sys)
sys.setdefaultencoding('utf-8')

jira_url = "http://127.0.0.1:8080/"
usrer    = "yujiawang"
password = "198317"
mpp_path = "D:\\june\smee\jira.mpp"

class JiraTool:
    def __init__(self):
        self.server = jira_url
        self.basic_auth = (usrer, password)
        self.jiraClinet = None
 
    def login(self):
        self.jiraClinet = JIRA(server=self.server, basic_auth=self.basic_auth)
        if self.jiraClinet != None:
            return True
        else:
            return False
 
    def findIssueById(self, issueId):
        if issueId:
            if self.jiraClinet == None:
                self.login()
            return self.jiraClinet.issue(issueId)
        else:
            return 'Please input your issueId'
 
    def createIssue(self, description, assignee, project, parent, issuetype):
#        new_issue = self.jiraClinet.create_issue(project='SCRUM', summary='New issue from jira-python',
#                              description='Look into this one', issuetype={'id': '10001'})
#        return
#        issue_dict = {
#            'project': {'key': 'SCRUM'},
#            'issuetype': {'id': 10001},
#            #'parent': {'key': parent},
#            #'summary': '[xxx]Auto sync file %s from storage-cnn to storage',
#            #'description': description,
#            #'assignee': {'name': assignee},
#            'customfield_10000': {'value': 'xxx'},
#            #'components': [{'name': 'SCM'}],
#            #'priority': {'id': 3},
#        }
        issue_dict = {
            'project': {'key': 'SCRUM'},
            'summary': 'New issue from jira-python 1',
            'description': 'Look into this one',
            'issuetype': {'id': '10001', 'name': 'Task'},
            #'priority': {'id': 3},
        }
        if self.jiraClinet == None:
            self.login()

        return self.jiraClinet.create_issue(issue_dict)

def initLogger(projectName):
    logPath = os.path.join(os.getcwd()+'/out/')
    if (False == os.path.exists(logPath)):
        os.makedirs(logPath)

    logFile = logPath+projectName+'.log'

    logger = logging.getLogger()
    logger.setLevel(logging.INFO)

    fh = logging.FileHandler(logFile, mode='w')  
    fh.setLevel(logging.INFO)
 
    ch = logging.StreamHandler()  
    ch.setLevel(logging.INFO)

    formatter = logging.Formatter("[%(funcName)s %(filename)s:%(lineno)d - %(levelname)s] %(message)s")
    fh.setFormatter(formatter)  
    ch.setFormatter(formatter)  

    logger.addHandler(fh)  
    logger.addHandler(ch)

def main():
    initLogger('mpp2jira')
    pythoncom.CoInitialize() #防止出现重复打开异常
    jiraTool = JiraTool()
    logging.info("jira login ... ")
    jiraTool.login()

    description = 'abcdefaaaaaaaaaaaaa'
    assignee = 'yujiawang'
    watchers = None
    project = 'scrum_project'
    parent = 'xxxx'
    openIssuekey = 'mppIssue1'
    cnnStoragePath = 'xxxxx'
    logging.info("create a issue.")
    issue = jiraTool.createIssue(description, assignee, project, parent, 'sub-task')
    issueKey = issue.key
    logging.info("add comment.")
    jiraTool.jiraClinet.add_comment(issue=issueKey, body='user does not exis')

    #logging.info("start read ... ")
    #reaadMpp(mpp_path)

if __name__ == '__main__':
    main()