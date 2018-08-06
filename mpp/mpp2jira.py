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

reload(sys)
sys.setdefaultencoding('utf-8')

jira_url = "http://127.0.0.1:8080/"
usrer    = "yujiawang"
password = "198317"
mpp_path = "D:\\june\\smee\\src\\auto_tools\\test\\jira.mpp"

issuetypes = {
    'Epic':"10000",
    'Task':"10002",
    'Sub-Task':"10003",
}

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

def dump(Project):
    """Dump file contents, for debugging purposes."""
    try:
        logging.info("This project has %s Tasks" %(str(Project.Tasks.Count)))
        for i in range(1,Project.Tasks.Count+1):
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
                if task.ResourceNames!=None and str(task.ResourceNames) != '':
                    print task.ResourceNames
                print '%'
            except:
                print 'Empty'
        return True
    except Exception, e:
        print "Error:", e
        return False

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

    def getProject(self, name):
        projects = self.jiraClinet.projects()
        logging.info(projects)
        for project in projects:
            if(name == project.name):
                return project
        
        return None
        #return self.jiraClinet.create_project(key='SCRUM', name=name, assignee='yujiawang', type="Software", template_name='Scrum')

    def getBoard(self, project, name):
        boards = self.jiraClinet.boards()
        print "###########boards#############"
        for board in boards:
            print '----------------'
            print '  name:' + board.name
            print '  id:' + str(board.id)
            sprints = self.jiraClinet.sprints(board.id)
            if(name == board.name):
                return board
            #for sprint in sprints:
            #    print '----------------'
            #    print '    name:' + sprint.name
            #    print '    id:' + str(sprint.id)
        return jira_create_board(name, '')

    def createSprint(self, name, board_id, startDate=None, endDate=None):
        return self.jiraClinet.create_sprint(name, board_id, startDate, endDate)

    def createEpicTask(self, sprint, project, summary, description, assignee):
        logging.info(project)
        issue_dict = {
            'project': {'key': project.key},
            'issuetype': {'id': issuetypes['Epic']},
            'customfield_10002' : 'EpicTest',
            'summary': summary,
            'description': description,
            'assignee': {'name': assignee},
        }
        issue = self.jiraClinet.create_issue(issue_dict)
        logging.info("add issue to sprint.")
        #self.jiraClinet.add_issues_to_sprint(sprint.id, issue.key)
        return issue

    def createTask(self, project, epic, summary, description, assignee):
        issue_dict = {
            'project': {'key': project.key},
            'issuetype': {'id': issuetypes['Task']},
            'summary': summary,
            'description': description,
            'assignee': {'name': assignee},
        }

        issue = self.jiraClinet.create_issue(issue_dict)
        #self.jiraClinet.add_issues_to_epic(epic.id, issue.key)
        return issue

    def createSubTask(self, project, parent, summary, description, assignee):
        issue_dict = {
            'project': {'key': project.key},
            'parent': {'key': parent.key},
            'issuetype': {'id': issuetypes['Sub-Task']},
            'summary': summary,
            'description': description,
            'assignee': {'name': assignee},
        }
        issue = self.jiraClinet.create_issue(issue_dict)
        return issue

    def createIssue(self, project):
        logging.info("create issue ...")
        issue_dict = {
            "project": {
                "key": project.key
            },
            "summary": "something's wrong for create issue debug.",
            "issuetype": {
                "id": "10000"
            },
            "assignee": {
                "name": "yujiawang"
            },
            "reporter": {
                "name": "yujiawang"
            },
            "priority": {
                "id": "3"
            },
            "labels": [
                "bugfix",
                "blitz_test"
            ],
            #"timetracking": {
            #    "originalEstimate": "10",
            #    "remainingEstimate": "5"
            #},
            #"security": {
            #    "id": "10000"
            #},
            #"versions": [
            #    {
            #        "id": "10000"
            #    }
            #],
            #"environment": "environment",
            "description": "description",
            #"duedate": "2011-03-11",
            #"fixVersions": [
            #    {
            #        "id": "10001"
            #    }
            #],
            "components": [
                {
                    "id": "10000"
                }
            ],
            #"customfield_60000": "jira-developers",
            #"customfield_20000": "06/Jul/11 3:25 PM",
            #"customfield_80000": {
            #    "value": "red"
            #},
            ###### Epic name #########
            "customfield_10002": "mpp",
            ###### Sprint name #########
            "customfield_10004": 1
            #"customfield_40000": "this is a text field",
            #"customfield_30000": [
            #    "10000",
            #    "10002"
            #],
            #"customfield_70000": [
            #    "jira-administrators",
            #    "jira-users"
            #],
            #"customfield_50000": "this is a text area. big text.",
            #"customfield_10000": "09/Jun/81"
        }
        if self.jiraClinet == None:
            self.login()

        return self.jiraClinet.create_issue(issue_dict)

def sync(jira_tool, mpp_file, jira_prj):
    mpp         = win32com.client.Dispatch("MSProject.Application")
    mpp.Visible = False
    mpp.FileOpen(mpp_file)
    mpp_prj     = mpp.ActiveProject

    #dump(mpp_prj)

    sprint    = None
    epic_task = None
    task      = None
    for i in range(1, mpp_prj.Tasks.Count+1):
        task        = mpp_prj.Tasks.Item(i)
        #summary      = task.Name[:100].decode("utf-8").encode("gbk")
        summary = task.Name
        description = '[deliverables]'+ task.Text1
        description += '[risk]'+ task.Text2
        assignee = 'yujiawang'
        logging.info("%s; %s" %(task.Name, description))
        if (1 == task.OutlineLevel):
            board     = jira_tool.getBoard(jira_prj, 'EC Kanban')
            sprint    = jira_tool.createSprint(summary, board.id)
            logging.info("%s" %(summary))
            # next sprint need reset its epic
            epic_task = None
            task      = None
        elif (2 == task.OutlineLevel):
            logging.info("  %s" %(summary))
            if sprint is None:
                logging.error("task not in sprint, check pls you mpp format!!")
                break
            epic_task = jira_tool.createEpicTask(sprint, jira_prj, summary, description, assignee)
            # next epic task, need reset its subtasks 
            task = None
        elif (3 == task.OutlineLevel):
            logging.info("    %s" %(summary))
            if epic_task is None:
                logging.error("task not in sprint, check pls you mpp format!!")
                break
            task = jira_tool.createTask(jira_prj, epic_task, summary, description, assignee)
        elif (4 == task.OutlineLevel):
            logging.info("      %s" %(summary))
            if task is None:
                logging.error("parent task is none, check pls you mpp format!!")
                break
            jira_tool.createSubTask(jira_prj, task, summary, description, assignee)

    mpp.Quit()

def main():
    initLogger('mpp2jira')

    options = {
        'server': jira_url}

    pythoncom.CoInitialize() #防止出现重复打开异常
    jira_tool = JiraTool()
    logging.info("jira login ... ")
    jira_tool.login()

    project = jira_tool.getProject('echromium')
    if (project is None):
        logging.error("%s project does not exist.")
        return
    #jira_tool.createIssue(project)
    logging.info("start sync project: %s" %(project.name))
    sync(jira_tool, mpp_path, project)

if __name__ == '__main__':
    main()

