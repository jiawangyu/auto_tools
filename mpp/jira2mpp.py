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
import pywintypes

import logging
import argparse
from optparse import OptionParser

from jira import JIRA
from jira.client import GreenHopper

reload(sys)
sys.setdefaultencoding('utf-8')

jira_url = "http://127.0.0.1:8080/"
usrer    = "yujiawang"
password = "198317"

ISSUE_EPIC_TYPE  = '10000'
ISSUE_STORY_TYPE = '10001'
ISSUE_TASK_TYPE  = '10002'

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

def exportSprint(name):
    sprint = None
    for sprint_name, sprint in gSprints.items():
        if(name == sprint_name):
            return sprint

    sprint = Sprint(name)
    gSprints[name] = sprint
    return sprint

def writeSprintTask(mppPrj, sprint, line):
    mppTask = mppPrj.Tasks.Add(sprint.name,line)       # 参数:任务名称、任务在第几行
    line += 1
    mppTask.OutlineLevel  = 1;
    mppTask.ResourceNames = '' # owner
    mppTask.ActualStart   = sprint.startDate      # 开始时间
    mppTask.ActualFinish  = sprint.endDate        # 结束时间
    mppTask.Predecessors  = ''                    # 前置任务id  注:前置任务id应该在导出完成后保存Task对象，重新循环添加前置任务。不然会出现任务3在第三行，而他的前置任务在第4行，那么会出现导出空的行
    mppTask.Milestone=False                       # 是否是milestone
    mppTask.ConstraintType = 5                    # 任务限制类型:越早越好、不得早于等等.  5:设置为不得晚于...开始，不会出现ms-project自动修改时间
    mppTask.ConstraintDate = ''                   # 任务限制日期
    mppTask.PercentComplete = '0'                 # 完成百分比

    logging.info("sprint %s "%(sprint.name))
    for epic in sprint.epics:
        line = writeEpciTask(mppPrj, epic, line)
    for task in sprint.tasks:
        line = writeTask(mppPrj, task, line)
    
    return line

def writeEpciTask(mppPrj, epic, line):
    epicTask = mppPrj.Tasks.Add(epic.summary, line)
    line += 1
    epicTask.OutlineLevel = 2;
    #epicTask.ResourceNames= epic.assignee
    epicTask.ActualStart= epic.created
    logging.info("|--epic %s "%(epic.summary))
    for task in epic.tasks:
        line = writeTask(mppPrj, task, line)
    
    return line

def writeTask(mppPrj, task, line):
    logging.info("|--task %s "%(task.summary))
    mTask = mppPrj.Tasks.Add(task.summary, line)
    line += 1
    mTask.OutlineLevel = 2;
    #mTask.ResourceNames= task.assignee
    mTask.ActualStart= task.created
    for sub_task in task.subtasks:
        logging.info("    |--subtask %s "%(sub_task.summary))
        subTask = mppPrj.Tasks.Add(sub_task.summary, line)
        line += 1
        subTask.OutlineLevel = 3;
        subTask.ResourceNames= sub_task.assignee
        subTask.ActualStart= sub_task.created
    
    return line

def writeMpp(outFile):
    proj        = ''
    mpp         = None

    mpp         = win32com.client.Dispatch("MSProject.Application")
    mpp.Visible = False
    mpp.FileNew(None,None,None,False)
    #mpp.FileOpen("D:\\june\\smee\\123.mpp")

    mpp.WBSCodeMaskEdit('',1,0)                  #导入顺序不一致添加
    mpp.WBSCodeRenumber(All=True)

    # import vba macro
    with open('init.bas') as f:
        macro = f.read()

    project = mpp.ActiveProject
    vbCode = project.VBProject.VBComponents("ThisProject").CodeModule
    vbCode.AddFromString(macro)

    # run vba macro
    mpp.Run("AddNewColum")

    line = 1
    for sprint in gSprints.values():
        line = writeSprintTask(project, sprint, line)

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

def exportSubTaskIssue(jira, task, issue):
    for sub in issue.raw['fields']['subtasks']:
        sub_issue = jira.issue(sub['key'])
        created_time = timeFormat(sub_issue.fields.created)
        sub_task = Task(sub_issue.fields.summary, created_time)
        task.subtasks.append(sub_task)
        logging.info("    issue[%s] add sub_task: %s" %(task.summary, sub_task.summary))

def exportEpicIssue(jira, sprint, issue):
    epic_issue = issue
    if issue.fields.customfield_10000: # issue 的epic字段
        epic_issue = jira.issue(issue.fields.customfield_10000)

    logging.info("  epic issue: %s" %(epic_issue.fields.summary))
    epic = sprint.getEpic(epic_issue.fields.summary)
    if epic is None:
        epic = Epic(epic_issue.fields.summary, timeFormat(epic_issue.fields.created))
        sprint.epics.append(epic)
    
    exportSubTaskIssue(jira, epic, epic_issue)

    return epic

def exportStoryIssue(jira, sprint, epic, issue):
    task = Task(issue.fields.summary, timeFormat(issue.fields.created))
    if epic is None:
        sprint.tasks.append(task)
        logging.info("sprint[%s] add story_issue: %s" %(sprint.name, task.summary))
    else:
        epic.tasks.append(task)
        logging.info("  epic[%s] add story_issue: %s" %(epic.summary, task.summary))

    exportSubTaskIssue(jira, epic, epic_issue)

    return task

def exportTaskIssue(jira, sprint, epic, issue):
    task = Task(issue.fields.summary, timeFormat(issue.fields.created))
    if epic is None:
        logging.info("sprint[%s] add task_issue: %s" %(sprint.name, task.summary))
        sprint.tasks.append(task)
    else:
        epic.tasks.append(task)
        logging.info("  epic[%s] add task_issue: %s" %(epic.summary, task.summary))

    exportSubTaskIssue(jira, task, issue)
    return task

def export(jira, projectName):
    project_issues = jira.search_issues("project="+projectName)
    for issue in project_issues:
        dumpIssue(issue)

        logging.info('<--------------------------')
        if issue.fields.customfield_10004:     # issue的sprint字段
            issue_sprint = issue.fields.customfield_10004
            sprint_name = get_sprint_name(issue_sprint[0])

        if sprint_name is None or '' == sprint_name:
            logging.error("issue[%s] not in any sprint." %(epic_issue.fields.summary))
            continue

        sprint = exportSprint(sprint_name)      # 对应sprint不存在则创建一个sprint对象 
        epic   = exportEpicIssue(jira, sprint, issue) # 如果某个issue关联了Epic，但对应的Epic issue还未导出，则先导出关联的Epic issue

        issue_type = issue.fields.issuetype.self
        issue_type = issue_type[issue_type.rfind('/')+1:]
        # 对于某些issue本身是Epic issue在这里导出，如果已经导出过的在exportEpicIssue会忽略掉
        if ISSUE_EPIC_TYPE == issue_type:    # Epic
            epic = exportEpicIssue(jira, sprint, issue)
        elif ISSUE_STORY_TYPE == issue_type: # Story
            task = exportStoryIssue(jira, sprint, epic, issue)
        elif ISSUE_TASK_TYPE == issue_type: # Task
            task = exportTaskIssue(jira, sprint, epic, issue)

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

def main(argv):
    parser = OptionParser()
    parser.add_option("--mpp_file",
                     help="output director for jira export to mpp file.",
                     default="D:\\june\\smee\\auto_tools_1.mpp")
    parser.add_option("--project",
                     help="the name of the project to be exported in jira.",
                     default="auto_tools")

    (options, _) = parser.parse_args(args=argv)
    output  = options.mpp_file
    project = options.project

    initLogger('jir2mpp')

    jira = JIRA(jira_url,basic_auth=(usrer, password))
    pythoncom.CoInitialize()

    logging.info("start export %s ... " %(project))
    export(jira, project)

    logging.info("start write to %s ... " %(output))
    writeMpp(output)

if __name__ == '__main__':
    sys.exit(main(sys.argv[1:]))