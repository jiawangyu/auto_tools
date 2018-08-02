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
mpp_path = "D:\\june\\smee\\src\\auto_tools\\test\\jira.mpp"

issuetypes = {
    'Sprint':"10001",
    'Epic':"10000",
    'Task':"10002",
    'Sub-Task':"10003",
}

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

    def createIssue(self, project, issuetype, assignee, summary, description):
        logging.info("create issue ...")
        issue_dict = {
            'project': {'key': project},
            'issuetype': {'id': issuetype},
            'summary': summary,
            'description': description,
            'assignee': {'name': assignee},
            #'duedate': '2018-8-3',
        }

        if self.jiraClinet == None:
            self.login()

        return self.jiraClinet.create_issue(issue_dict)

def dump(Project):
    """Dump file contents, for debugging purposes."""
    try:
        print "This project has ", str(Project.Tasks.Count), " Tasks"
        for i in range(1,Project.Tasks.Count+1):
            print i,
            task = Project.Tasks.Item(i)
            if (1 == task.OutlineLevel):
                space=""
            elif (2 == task.OutlineLevel):
                space="  |->"
            elif (3 == task.OutlineLevel):
                space="    |->"
            elif (4 == task.OutlineLevel):
                space="    |->"
            try:
                print space + task.Name[:100].decode("utf-8").encode("gbk"),
                print task.OutlineLevel,
                print task.Text1.decode("utf-8").encode("gbk"),   # 自定义列1                
                print task.Text2.decode("utf-8").encode("gbk"),   # 自定义列2
                print task.ResourceNames.decode("utf-8").encode("gbk"),
                print task.Start,
                print task.Finish,
                print task.PercentWorkComplete,
                print '%'
            except:
                print 'Empty'
        return True
    except Exception, e:
        print "Error:", e
        return False

def sync_from_mpp(infile, jiraTool):
    mpp     = win32com.client.Dispatch("MSProject.Application")
    mpp.Visible = False
    mpp.FileOpen(infile)
    proj = mpp.ActiveProject

    #dump(proj)

    for i in range(1, proj.Tasks.Count+1):
        task = proj.Tasks.Item(i)
        if (1 == task.OutlineLevel):
            issuetype = issuetypes['Sprint']
        elif (2 == task.OutlineLevel):
            issuetype = issuetypes['Epic']
        elif (3 == task.OutlineLevel):
            issuetype = issuetypes['Task']
        elif (4 == task.OutlineLevel):
            issuetype = issuetypes['Sub-Task']
        sumary = task.Name[:100].decode("utf-8").encode("gbk")
        description = '[deliverables]'+ task.Text1.decode("utf-8").encode("gbk")
        description = '[risk]'+ task.Text2.decode("utf-8").encode("gbk")

        issue = jiraTool.createIssue('SCRUM', issuetype, 'yujiawang', sumary, description)
        break
        #issue = jiraTool.createIssue('SCRUM', 'Task', 'sumary', 'description')
        #issueKey = issue.key
        #logging.info("add comment.")
        #jiraTool.jiraClinet.add_comment(issue=issueKey, body='user does not exis')

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

    logging.info("sync_from_mpp ... ")
    sync_from_mpp(mpp_path, jiraTool)

if __name__ == '__main__':
    main()