# -*- coding: utf-8 -*-
# **************************************************************
# *  Filename:    jira2mpp.py
# *  Copyright:   SMEE Co., Ltd.
# *  @author:     jun.zhou
# *  将jira里的数据导出为mpp文件格式
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

jira = JIRA('http://127.0.0.1:8080/',basic_auth=('yujiawang','198317'))

pythoncom.CoInitialize() #防止出现重复打开异常

gSprints = {}

class Sprint:
    def __init__(self, id='', state='', name='', startDate='', endDate='', completeDate='', sequence=''):
        self.id           = id
        self.state        = state
        self.name         = name
        self.startDate    = startDate
        self.endDate      = endDate
        self.completeDate = completeDate
        self.sequence     = sequence
        self.epics        = []
        self.tasks        = []

    def getEpic(self, summary):
        for epic in self.epics:
            if (epic.summary == summary):
                return epic
        return None

    def getTask(self, summary):
        for task in self.tasks:
            if (task.summary == summary):
                return task
        return None

class Epic:
   def __init__(self, summary, created):
       self.summary = summary
       self.created = created
       self.tasks = []

class Task:
    def __init__(self, summary, created, assignee=''):
        self.summary = summary
        self.created = created
        self.subtasks = []
        self.assignee = assignee

def txt_wrap_by(str, start_str, end):
    start = str.find(start_str)
    if start >= 0:
        start += len(start_str)
        end = str.find(end, start)
        if end >= 0:
            return str[start:end].strip()

def get_sprint_name(sprint_field):
    value = txt_wrap_by(sprint_field, '[',']')
    value_list = value.split(',')
    for item in value_list:
        fields=item.split('=')
        if('name' == fields[0]):
            return fields[1]

def getSprint(name):
    for sprint_name, sprint in gSprints.items():
        if(name == sprint_name):
            return sprint

    return None

def writeMpp(outFile):
    proj    = ''
    mpp     = None

    mpp         = win32com.client.Dispatch("MSProject.Application")
    mpp.Visible = True
    mpp.FileNew(None,None,None,False)
    mpp.WBSCodeMaskEdit('',1,0)                  #导入顺序不一致添加
    mpp.WBSCodeRenumber(All=True) 
    proj = mpp.ActiveProject

    line = 1
    for sprint in gSprints.values():
        sprintTask = proj.Tasks.Add(sprint.name,line)       # 参数:任务名称、任务在第几行
        line += 1
        sprintTask.OutlineLevel  = 1;
        sprintTask.ResourceNames = '' # owner
        sprintTask.ActualStart   = sprint.startDate      # 开始时间
        sprintTask.ActualFinish  = sprint.endDate        # 结束时间
        sprintTask.Predecessors  = ''                    # 前置任务id  注:前置任务id应该在导出完成后保存Task对象，重新循环添加前置任务。不然会出现任务3在第三行，而他的前置任务在第4行，那么会出现导出空的行
        sprintTask.Milestone=False                       # 是否是milestone
        sprintTask.ConstraintType = 5                    # 任务限制类型:越早越好、不得早于等等.  5:设置为不得晚于...开始，不会出现ms-project自动修改时间
        sprintTask.ConstraintDate = ''                   # 任务限制日期
        sprintTask.PercentComplete = '0'                 # 完成百分比
        logging.info("sprint %s "%(sprint.name))
        for epic in sprint.epics:
            epicTask = proj.Tasks.Add(epic.summary, line)
            line += 1
            epicTask.OutlineLevel = 2;
#            epicTask.ResourceNames= epic.assignee
            epicTask.ActualStart= epic.created
            logging.info("|--epic %s "%(epic.summary))
            for task in epic.tasks:
                logging.info("   |--task %s "%(task.summary))
                mTask = proj.Tasks.Add(task.summary, line)
                line += 1
                mTask.OutlineLevel  = 3;
                mTask.ResourceNames = task.assignee
                mTask.ActualStart   = task.created
                for sub_task in task.subtasks:
                    logging.info("      |--subtask %s "%(sub_task.summary))
                    subTask = proj.Tasks.Add(sub_task.summary, line)
                    line += 1
                    subTask.OutlineLevel  = 4;
                    subTask.ResourceNames = sub_task.assignee
                    subTask.ActualStart   = sub_task.created
        for task in sprint.tasks:
            logging.info("|--task %s "%(task.summary))
            mTask = proj.Tasks.Add(task.summary, line)
            line += 1
            mTask.OutlineLevel = 2;
#            mTask.ResourceNames= task.assignee
            mTask.ActualStart= task.created
            for sub_task in task.subtasks:
                logging.info("    |--subtask %s "%(sub_task.summary))
                subTask = proj.Tasks.Add(sub_task.summary, line)
                line += 1
                subTask.OutlineLevel = 3;
                subTask.ResourceNames= sub_task.assignee
                subTask.ActualStart= sub_task.created

    mpp.FileSaveAs(outFile);
    mpp.Quit(); 
            
    mpp = None

def dumpIssue(issue):
    logging.info('--------------- dump issue beging --------------- ')
    for field_name in issue.raw['fields']:
        logging.info("  %s:%s" %(field_name, issue.raw['fields'][field_name]))
    logging.info('--------------- dump issue end --------------- ')

def timeFormat(time):
    return time[::-1].split('T', 1)[-1][::-1].replace('-','/')

def export(projectName):
    project_issues = jira.search_issues(projectName)
    for issue in project_issues:
        created_time = timeFormat(issue.fields.created)
        logging.info('<--------------------------')
        sprint_name = ''
        sprint = None
        if issue.fields.customfield_10004:
            issue_sprint = issue.fields.customfield_10004
            sprint_name = get_sprint_name(issue_sprint[0])
        if '' == sprint_name:
            logging.error("issue[%s] not in any sprint." %(epic_issue.fields.summary))
            continue
            logging.info('-------------------------->')

        sprint = getSprint(sprint_name)
        if sprint is None:
            sprint = Sprint(name=sprint_name)
            gSprints[sprint_name] = sprint
        else:
            sprint = gSprints[sprint_name]

        epic = None
        if issue.fields.customfield_10000:
            epic_issue = jira.issue(issue.fields.customfield_10000)
            epic = sprint.getEpic(epic_issue.fields.summary)
            if epic is None:
                epic = Epic(epic_issue.fields.summary, created_time)
                sprint.epics.append(epic)
                logging.info("sprint[%s] add epic_issue: %s" %(sprint.name, epic.summary))

            # epic sub tasks
            for sub in issue.raw['fields']['subtasks']:
                sub_issue = jira.issue(sub['key'])
                logging.info("      sub issue: %s" %(sub_issue.fields.summary))
                created_time = timeFormat(sub_issue.fields.created)
                sub_task = Task(sub_issue.fields.summary, created_time)
                epic.tasks.append(sub_task)

        issue_type = issue.fields.issuetype.self
        issue_type = issue_type[issue_type.rfind('/')+1:]
        # 对于某些Epic中没有关联的问题在这里处理
        if '10000' == issue_type:   # Epic
            logging.info("  epic issue: %s" %(issue.fields.summary))
            epic = sprint.getEpic(issue.fields.summary)
            if epic is None:
                epic = Epic(issue.fields.summary, created_time)
                sprint.epics.append(epic)
                # sub tasks
                for sub in issue.raw['fields']['subtasks']:
                    sub_issue = jira.issue(sub['key'])
                    logging.info("      sub issue: %s" %(sub_issue.fields.summary))
                    created_time = timeFormat(sub_issue.fields.created)
                    sub_task = Task(sub_issue.fields.summary, created_time)
                    epic.tasks.append(sub_task)
        elif '10001' == issue_type: # Story
            task = Task(issue.fields.summary, created_time)
            if epic is None:
                sprint.tasks.append(task)
                logging.info("sprint[%s] add story_issue: %s" %(sprint.name, task.summary))
            else:
                epic.tasks.append(task)
                logging.info("  epic[%s] add story_issue: %s" %(epic.summary, task.summary))
            # sub tasks
            for sub in issue.raw['fields']['subtasks']:
                sub_issue = jira.issue(sub['key'])
                created_time = timeFormat(sub_issue.fields.created)
                sub_task = Task(sub_issue.fields.summary, created_time)
                task.subtasks.append(sub_task)
                logging.info("    story_issue[%s] add sub_task: %s" %(task.summary, sub_task.summary))
        elif '10002' == issue_type: # Task
            logging.info("  task issue: %s" %(issue.fields.summary))
            task = Task(issue.fields.summary, created_time)
            if epic is None:
                logging.info("sprint[%s] add task_issue: %s" %(sprint.name, task.summary))
                sprint.tasks.append(task)
            else:
                epic.tasks.append(task)
                logging.info("  epic[%s] add task_issue: %s" %(epic.summary, task.summary))
            # sub tasks
            for sub in issue.raw['fields']['subtasks']:
                sub_issue = jira.issue(sub['key'])
                created_time = timeFormat(sub_issue.fields.created)
                sub_task = Task(sub_issue.fields.summary, created_time)
                task.subtasks.append(sub_task)
                logging.info("    task_issue[%s] add sub_task: %s" %(task.summary, sub_task.summary))
#        elif '10003' == issue_type: # SubTask
        logging.info('-------------------------->')

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
    initLogger('jir2mpp')

    logging.info("start export ... ")
    export('project=echromium')
    logging.info("start write ... ")
    writeMpp('D:\\june\smee\jira.mpp')

if __name__ == '__main__':
    main()