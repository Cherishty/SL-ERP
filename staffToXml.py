from xml.dom.minidom import Document

def GenerateXml(result):
    root = Document()
    staffList = root.createElement('StaffList');
    staffList.setAttribute('name','ShuangLong')
    root.appendChild(staffList)
    departmentList=['前纺','细沙甲','细沙乙','筒摇','辅助','保全','行政']
    for department in departmentList:
        departNode=root.createElement('Department')
        departNode.setAttribute('name',department)
        staffList.appendChild(departNode)
    for each in result['lastTem']:
        staffNode=root.createElement('staff')
        staffNode.setAttribute('Name',each[0])
        staffNode.setAttribute('No.','1')
        staffNode.setAttribute('bankNo.',each[1])
        departNode.appendChild(staffNode)
    f=open('StaffList.xml','w')
    root.writexml(f,'',' ','\n','utf-8')
    f.close()

if __name__ == '__main__':
    result={}
    GenerateXml(result)