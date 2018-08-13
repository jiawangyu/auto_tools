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

#--------------------- custom area start -----------------------
jira_url = "http://127.0.0.1:8080/"
usrer    = "yujiawang"
password = "198317"
gSprints = {}

ISSUE_EPIC_TYPE  = '10000'
ISSUE_STORY_TYPE = '10001'
ISSUE_TASK_TYPE  = '10002'

def issueSprintField(issue):
    return issue.fields.customfield_10004

def issueEpicField(issue):
    return issue.fields.customfield_10000
#--------------------- custom area end -----------------------

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

    def getEpic(self, name):
        for epic in self.epics:
            if (epic.name == name):
                return epic
        return None

    def getTask(self, name):
        for task in self.tasks:
            if (task.name == name):
                return task
        return None

class Epic:
   def __init__(self, name, created):
       self.name    = name
       self.created = created
       self.tasks = []

class Task:
    def __init__(self, name, created, assignee=''):
        self.name    = name
        self.created = created
        self.subtasks = []
        self.assignee = assignee

def writeSprintTask(mppPrj, sprint):
    mppTask = mppPrj.Tasks.Add(sprint.name, mppPrj.Tasks.Count+1)       # 参数:任务名称、任务在第几行
    mppTask.OutlineLevel    = 1;
    mppTask.ResourceNames   = ''                  # owner
    mppTask.ActualStart     = sprint.startDate    # 开始时间
    mppTask.ActualFinish    = sprint.endDate      # 结束时间
    mppTask.Predecessors    = ''                  # 前置任务id  注:前置任务id应该在导出完成后保存Task对象，重新循环添加前置任务。不然会出现任务3在第三行，而他的前置任务在第4行，那么会出现导出空的行
    mppTask.Milestone       = False               # 是否是milestone
    mppTask.ConstraintType  = 5                   # 任务限制类型:越早越好、不得早于等等.  5:设置为不得晚于...开始，不会出现ms-project自动修改时间
    mppTask.ConstraintDate  = ''                  # 任务限制日期
    mppTask.PercentComplete = '0'                 # 完成百分比

    writeLog(sprint, 1)
    for epic in sprint.epics:
        writeEpciTask(mppPrj, epic)
    for task in sprint.tasks:
        writeTask(mppPrj, task, 2)

def writeEpciTask(mppPrj, epic):
    epicTask = mppPrj.Tasks.Add(epic.name, mppPrj.Tasks.Count+1)
    epicTask.OutlineLevel = 2;
    #epicTask.ResourceNames= epic.assignee
    epicTask.ActualStart= epic.created

    writeLog(epic, 2)
    for task in epic.tasks:
        writeTask(mppPrj, task, 3)

# 直接挂在sprint下的task为二级任务
# 挂在epic下的task为三级任务
def writeTask(mppPrj, task, level):
    writeLog(task, level)
    mTask = mppPrj.Tasks.Add(task.name, mppPrj.Tasks.Count+1)

    mTask.OutlineLevel = level;
    #mTask.ResourceNames= task.assignee
    mTask.ActualStart= task.created
    for sub_task in task.subtasks:
        writeLog(sub_task, level + 1)
        subTask = mppPrj.Tasks.Add(sub_task.name, mppPrj.Tasks.Count+1)

        subTask.OutlineLevel  = level + 1;
        subTask.ResourceNames = sub_task.assignee
        subTask.ActualStart   = sub_task.created

def dumpExport():
    logging.info("start dump export data ...")
    for sprint in gSprints.values():
        logging.info("%s" %(sprint.name))
        for epic in sprint.epics:
            logging.info("  |--> %s" %(epic.name))
            for task in epic.tasks:
                logging.info("      |--> %s" %(task.name))
                for sub_task in task.subtasks:
                    logging.info("          |--> %s" %(sub_task.name))
        for task in sprint.tasks:
            logging.info("    |--> %s" %(task.name))
            for sub_task in task.subtasks:
                logging.info("        |--> %s" %(sub_task.name))

def writeLog(task, level):
    if (1 == level):
        logging.info("一级任务 %s" %(task.name))
    elif(2 == level):
        logging.info("  |-> 二级任务 %s" %(task.name))
    elif(3 == level):
        logging.info("     |-> 三级任务 %s" %(task.name))
    elif(4 == level):
        logging.info("        |-> 四级任务 %s" %(task.name))

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

    for sprint in gSprints.values():
        writeSprintTask(project, sprint)

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

def exportStoryIssue(jira, sprint, epic, issue):
    task = Task(issue.fields.summary, timeFormat(issue.fields.created))
    if epic is None:
        sprint.tasks.append(task)
    else:
        epic.tasks.append(task)

    exportSubTaskIssue(jira, epic, epic_issue)
    return task

def exportTaskIssue(jira, sprint, epic, issue):
    task = Task(issue.fields.summary, timeFormat(issue.fields.created))
    if epic is None:
        sprint.tasks.append(task)
    else:
        epic.tasks.append(task)

    exportSubTaskIssue(jira, task, issue)
    return task

def exportEpicIssue(jira, sprint, issue, isEpicIssue):
    if (False == isEpicIssue):
        if(issueEpicField(issue)):
            issue = jira.issue(issueEpicField(issue))
        else:
            return None

    epic = sprint.getEpic(issue.fields.summary)
    if epic is None:
        epic = Epic(issue.fields.summary, timeFormat(issue.fields.created))
        sprint.epics.append(epic)
    
    exportSubTaskIssue(jira, epic, issue)
    return epic

def txt_wrap_by(str, start_str, end):
    start = str.find(start_str)
    if start >= 0:
        start += len(start_str)
        end = str.find(end, start)
        if end >= 0:
            return str[start:end].strip()

def getSprintName(sprint_field):
    value = txt_wrap_by(sprint_field, '[',']')
    value_list = value.split(',')
    for item in value_list:
        fields=item.split('=')
        if('name' == fields[0]):
            return fields[1]

def exportSprint(issue):
    sprint      = None
    sprint_name = ''
    issue_sprint = issueSprintField(issue)
    if issue_sprint:     #### issue的sprint字段
        sprint_name = getSprintName(issue_sprint[0])

    if '' == sprint_name:
        logging.error("issue[%s][%s] not in any sprint." %(issue.fields.summary, issue.key))
        return None

    #如果存在同名的sprint直接返回
    for name, sprint in gSprints.items():
        if(name == sprint_name):
            return sprint

    sprint = Sprint(name=sprint_name)
    gSprints[sprint_name] = sprint
    return sprint

def export(jira, projectName):
    issues = jira.search_issues("project="+projectName)
    for issue in issues:
        #dumpIssue(issue)
        sprint = exportSprint(issue)      # 对应sprint不存在则创建一个sprint对象 
        if(sprint is None):
            continue

        epic   = exportEpicIssue(jira, sprint, issue, False) # 如果某个issue关联了Epic，但对应的Epic issue还未导出，则先导出关联的Epic issue

        issue_type = issue.fields.issuetype.self
        issue_type = issue_type[issue_type.rfind('/')+1:]
        # 对于某些issue本身是Epic issue在这里导出，如果已经导出过的在exportEpicIssue会忽略掉
        if ISSUE_EPIC_TYPE == issue_type:    # Epic
            epic = exportEpicIssue(jira, sprint, issue, True)
        elif ISSUE_STORY_TYPE == issue_type: # Story
            task = exportStoryIssue(jira, sprint, epic, issue)
        elif ISSUE_TASK_TYPE == issue_type:  # Task
            task = exportTaskIssue(jira, sprint, epic, issue)

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
                     default="D:\\june\\smee\\auto_tools.mpp")
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
    dumpExport()

    logging.info("start write to %s ... " %(output))
    writeMpp(output)

if __name__ == '__main__':
    sys.exit(main(sys.argv[1:]))