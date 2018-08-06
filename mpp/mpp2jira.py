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

    def deleteAllIssue(self, projectName):
        project_issues = self.jiraClinet.search_issues('project='+projectName)
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
        logging.info(projects)
        for project in projects:
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
        return self.jiraClinet.create_board(name, project.id)

    def createSprint(self, board, name, startDate=None, endDate=None):
        logging.info("==== create sprint[%s] in board[%s] ====" % (name, board.name))
        sprint = self.jiraClinet.create_sprint(name, board.id, startDate, endDate)
        return sprint

    def getSprint(self, board_id, sprint_name):
        sprints = self.jiraClinet.sprints(board_id)
        for sprint in sprints:
            if sprint.name == sprint_name:
                return sprint
        return None

    def createEpicTask(self, sprint, project, summary, description, assignee):
        issue_dict = {
            'project': {'key': project.key},
            'issuetype': {'id': issuetypes['Epic']},
            'customfield_10002': summary, #epic 名称
            'summary': summary,
            'description': description,
            "customfield_10004": sprint.id, #sprint
            'assignee': {'name': assignee},
        }

        #logging.info(issue_dict) #juse for debug
        issue = self.jiraClinet.create_issue(issue_dict)
        self.jiraClinet.add_issues_to_sprint(sprint.id, [issue.raw['key']])
        logging.info("  add epic task[%s key:%s] to sprint [%s]" %(issue.raw['fields']['summary'], issue.raw['key'], sprint.name))
        #dumpIssue(issue) #juse for debug
        return issue

    def createTask(self, project, epic, summary, description, assignee):
        issue_dict = {
            'project': {'key': project.key},
            'issuetype': {'id': issuetypes['Task']},
            'summary': summary,
            #'customfield_10002': epic.raw['fields']['summary'], #epic
            'description': description,
            'assignee': {'name': assignee},
        }

        issue = self.jiraClinet.create_issue(issue_dict)
        logging.info("    add task[%s key:%s] link epic [%s key: %s]" %(issue.raw['fields']['summary'], issue.raw['key'], epic.raw['fields']['summary'], epic.raw['key']))
        self.jiraClinet.add_issues_to_epic(epic.id, [issue.raw['key']])
        return issue

    def createSubTask(self, project, parent, summary, description, assignee):
        issue_dict = {
            'project': {'key': project.key},
            'parent': {'key': parent.raw['key']},
            'issuetype': {'id': issuetypes['Sub-Task']},
            'summary': summary,
            'description': description,
            'assignee': {'name': assignee},
        }
        issue = self.jiraClinet.create_issue(issue_dict)
        logging.info("      add sub task[%s key:%s] to task [%s key: %s]" %(issue.raw['fields']['summary'], issue.raw['key'], parent.raw['fields']['summary'], parent.raw['key']))
        return issue

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
        mpp_task = mpp_prj.Tasks.Item(i)
        summary  = mpp_task.Name
        description = '[deliverables]'+ mpp_task.Text1
        description += '[risk]'+ mpp_task.Text2
        assignee    = 'yujiawang'

        if (1 == mpp_task.OutlineLevel):
            board  = jira_tool.getBoard(jira_prj, 'EC Kanban')
            sprint = jira_tool.getSprint(board.id, summary)
            if sprint is None:
                sprint    = jira_tool.createSprint(board, summary)

            # next sprint need reset its epic
            epic_task = None
            task      = None
        elif (2 == mpp_task.OutlineLevel):
            if sprint is None:
                logging.error("task not in sprint, check pls you mpp format!!")
                break
            epic_task = jira_tool.createEpicTask(sprint, jira_prj, summary, description, assignee)
            # next epic task, need reset its subtasks 
            task = None
        elif (3 == mpp_task.OutlineLevel):
            if epic_task is None:
                logging.error("task not in sprint, check pls you mpp format!!")
                break
            task = jira_tool.createTask(jira_prj, epic_task, summary, description, assignee)
        elif (4 == mpp_task.OutlineLevel):
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

    logging.info('clear ...')
    jira_tool.deleteAllIssue('echromium')
    board  = jira_tool.getBoard('echromium', 'EC Kanban')
    if board:
        jira_tool.deleteAllSprint(board)

    logging.info("start sync project: %s ..." %(project.name))
    sync(jira_tool, mpp_path, project)

if __name__ == '__main__':
    main()

