# mpp与jira同步
概述：将MS project的任务同步到JIRA中
## 需求：
1. 导入JIRA时，4个等级任务要形成对应关系:
一级任务 = sprint
二级任务 = epic
三级任务 = task
四级任务 = subtask
2. mpp的计划开始时间 = jira的创建日期
   mpp的计划完成时间 = jira的到期日
3. mpp里程碑导入JIRA时需要忽略
4. mmp资源导入jira时，第一个资源作为经办人，所有资源作为参与人
5. mpp中的交付列，导出到JIRA的描述字段中，如果mpp的交付件列不为空，那么将此列的值导到JIRA描述字段，显示为【交付件】***，
如果存在多个交付件，那么在描述字段中多行显示。
6. mpp中的风险列，导出到JIRA描述字段中，如果mpp的交付件列不为空，那么将此列的值导到JIRA描述符字段，显示为【风险】***，如果存在多个风险，那么在描述符字段中按多行显示。

## 关键指标：
1. 代码需要提交到bitbucket上归档
2. 代码经过走读
3. 使用实际项目的mpp文件作为测试数据
44. 需要进行架构设计，考虑使用ms project更新JIRA现有任务数据的可能性


## 交付件要求
报告类（调研、概念、创意、方案等）
编码类（源码、演示、程序、测试等）