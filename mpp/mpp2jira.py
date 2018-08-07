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
from optparse import OptionParser

from jira import JIRA

reload(sys)
sys.setdefaultencoding('utf-8')

jira_url = "http://127.0.0.1:8080/"
usrer    = "yujiawang"
password = "198317"

issuetypes = {
    'Epic':"10000",
    'Task':"10002",
    'Sub-Task':"10003",
}

resouceTable = {
    '周俊':"zhoujun",
    '俞晓慧':'yuxiaohui',
    '张三':"zhangsan",
    '李四':"lisi",
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

def dumpIssue(issue):
    logging.info('--------------- dump issue [%s] key:%s id:%s beging --------------- '%(issue.raw['fields']['summary'], issue.raw['key'], issue.raw['id']))
    for field_name in issue.raw['fields']:
        logging.info("  %s:%s" %(field_name, issue.raw['fields'][field_name]))
    logging.info('--------------- dump issue end --------------- ')

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

def assigneeAndParticipant(resourceNames):
    if(resourceNames == None or str(resourceNames) == ''):
        logging.info("resource names is empty.")
        return None, None

    resources = resourceNames.split(',')
    # chinese to english
    participant = ""
    for i in range(0, len(resources)):
        for chinese, english in resouceTable.items():              
            if(chinese == resources[i]):
                if(0 == i):
                    assignee = english
                else:
                    participant += " "
                    participant += english

    return assignee, participant

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

    def deleteAllIssue(self, project):
        project_issues = self.jiraClinet.search_issues('project='+project.name)
        for issue in project_issues:
            logging.info('delete issue %s' %(issue))
            issue.delete()

    def deleteAllSprint(self, board):
        sprints = self.jiraClinet.sprints(board.id)
        for sprint in sprints:
            logging.info('delete sprint %s' %(sprint.name))
            sprint.delete()
        #    print '----------------'
        #    print '    name:' + sprint.name
        #    print '    id:' + str(sprint.id)

    def getProject(self, name):
        projects = self.jiraClinet.projects()
        #logging.info("get project %s" %(name))
        for project in projects:
            #logging.info("project %s" %(project.name))
            if(name == project.name):
                return project

        return None
        #return self.jiraClinet.create_project(key='SCRUM', name=name, assignee='yujiawang', type="Software", template_name='Scrum')

    def getBoard(self, project, name):
        boards = self.jiraClinet.boards()
        for board in boards:
            #print '  name:' + board.name
            #print '  id:' + str(board.id)
            if(name == board.name):
                logging.info("board：%s id:%d" %(board.name, board.id))
                return board

            sprints = self.jiraClinet.sprints(board.id)
            #for sprint in sprints:
            #    print '----------------'
            #    print '    name:' + sprint.name
            #    print '    id:' + str(sprint.id)
        return self.jiraClinet.create_board(name, [project.id])

    def createSprint(self, board, name, startDate=None, endDate=None):
        logging.info("==== create sprint[%s] in board[%s] ====>" % (name, board.name))
        sprint = self.jiraClinet.create_sprint(name, board.id, startDate, endDate)
        return sprint

    def getSprint(self, board_id, sprint_name):
        sprints = self.jiraClinet.sprints(board_id)
        for sprint in sprints:
            if sprint.name == sprint_name:
                return sprint
        return None

    def createEpicTask(self, sprint, project, summary, description, assignee, participant=''):
        issue_dict = {
            'project': {'key': project.key},
            'issuetype': {'id': issuetypes['Epic']},
            'customfield_10002': summary, #epic 名称
            'summary': summary,
            'description': description,
            "customfield_10004": sprint.id, #sprint
            'assignee': {'name': assignee},
            'customfield_10301' : participant, #参与人
        }

        logging.info(issue_dict) #juse for debug
        issue = self.jiraClinet.create_issue(issue_dict)
        self.jiraClinet.add_issues_to_sprint(sprint.id, [issue.raw['key']])
        logging.info("===> add epic task[%s key:%s] to sprint [%s]" %(issue.raw['fields']['summary'], issue.raw['key'], sprint.name))
        #dumpIssue(issue) #juse for debug
        return issue

    def createTask(self, project, epic, summary, description, assignee, participant=''):
        issue_dict = {
            'project': {'key': project.key},
            'issuetype': {'id': issuetypes['Task']},
            'summary': summary,
            #'customfield_10002': epic.raw['fields']['summary'], #epic
            'description': description,
            'assignee': {'name': assignee},
            'customfield_10301' : participant, #参与人
        }

        issue = self.jiraClinet.create_issue(issue_dict)
        logging.info("==> add task[%s key:%s] link epic [%s key: %s]" %(issue.raw['fields']['summary'], issue.raw['key'], epic.raw['fields']['summary'], epic.raw['key']))
        self.jiraClinet.add_issues_to_epic(epic.id, [issue.raw['key']])
        return issue

    def createSubTask(self, project, parent, summary, description, assignee, participant=''):
        issue_dict = {
            'project': {'key': project.key},
            'parent': {'key': parent.raw['key']},
            'issuetype': {'id': issuetypes['Sub-Task']},
            'summary': summary,
            'description': description,
            'assignee': {'name': assignee},
            'customfield_10301' : participant, #参与人
        }
        issue = self.jiraClinet.create_issue(issue_dict)
        logging.info("=> add sub task[%s key:%s] to task [%s key: %s]" %(issue.raw['fields']['summary'], issue.raw['key'], parent.raw['fields']['summary'], parent.raw['key']))
        return issue

def sync(jira_tool, mpp_file, jira_prj, board):
    mpp         = win32com.client.Dispatch("MSProject.Application")
    mpp.Visible = False
    try:
        mpp.FileOpen(mpp_file)
        mppPrj      = mpp.ActiveProject
    except pythoncom.com_error as error:
        print(error.strerror)
    #dump(mppPrj)

    sprint    = None
    epic_task = None
    task      = None
    for i in range(1, mppPrj.Tasks.Count+1):
        mppTask     = mppPrj.Tasks.Item(i)
        summary     = mppTask.Name
        description = '[deliverables]'+ mppTask.Text1
        description += '[risk]'+ mppTask.Text2

        assignee, participant = assigneeAndParticipant(mppTask.ResourceNames)
        logging.info('assignee :%s' %(assignee))
        #logging.info('participant :%s' %(" ".join(str(i) for i in participant)))
        logging.info('participant :%s' %(participant))
        continue

        if (1 == mppTask.OutlineLevel):
            sprint = jira_tool.getSprint(board.id, summary)
            if sprint is None:
                sprint = jira_tool.createSprint(board, summary)

            # next sprint need reset its epic
            epic_task = None
            task      = None
        elif (2 == mppTask.OutlineLevel):
            if sprint is None:
                logging.error("task not in sprint, check pls you mpp format!!")
                break
            epic_task = jira_tool.createEpicTask(sprint, jira_prj, summary, description, assignee, participant)
            # next epic task, need reset its subtasks 
            task = None
        elif (3 == mppTask.OutlineLevel):
            if epic_task is None:
                logging.error("task not in sprint, check pls you mpp format!!")
                break
            task = jira_tool.createTask(jira_prj, epic_task, summary, description, assignee, participant)
        elif (4 == mppTask.OutlineLevel):
            if task is None:
                logging.error("parent task is none, check pls you mpp format!!")
                break
            jira_tool.createSubTask(jira_prj, task, summary, description, assignee, participant)

    mpp.Quit()

def main(argv):
    parser = OptionParser()
    parser.add_option("--mpp_file",
                     help="output director for jira export to mpp file.",
                     default="D:\\june\\smee\\auto_tools.mpp")
    parser.add_option("--project",
                     help="the name of the project to be exported in jira.",
                     default="sync_from_mpp")

    (options, _) = parser.parse_args(args=argv)
    mpp_file  = options.mpp_file
    project   = options.project

    initLogger('mpp2jira')

    pythoncom.CoInitialize() #防止出现重复打开异常
    jira_tool = JiraTool()
    logging.info("jira login ... ")
    jira_tool.login()

    #for debug
    users = jira_tool.jiraClinet.search_users('z')
    logging.info(users)
    return

    project = jira_tool.getProject(project)
    if (project is None):
        logging.error("%s project does not exist.")
        return

    logging.info('clear ...')
    jira_tool.deleteAllIssue(project)
    board  = jira_tool.getBoard(project, project.name)
    if board:
        jira_tool.deleteAllSprint(board)

    logging.info("start sync project: %s ..." %(project))
    sync(jira_tool, mpp_file, project, board)

if __name__ == '__main__':
    sys.exit(main(sys.argv[1:]))

