# -*- coding: utf-8 -*-
# **************************************************************
# *  Filename:    jira2mpp.py
# *  Copyright:   SMEE Co., Ltd.
# *  @author:     jun.zhou
# *  ��jira������ݵ���Ϊmpp�ļ���ʽ
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

pythoncom.CoInitialize() #��ֹ�����ظ����쳣

